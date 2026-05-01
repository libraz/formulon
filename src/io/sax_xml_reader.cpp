// Copyright 2026 libraz. Licensed under the MIT License.
//
// SAX-style scanner for `xl/worksheets/sheet*.xml`. See the header for
// the public contract.
//
// Implementation strategy:
//   * Single forward pass over the byte buffer, never random-access.
//   * Scratch buffers are reused across rows/cells: the scanner only
//     allocates when a text node needs entity decoding (`&amp;` etc.) or
//     when the inline-string body contains nested `<r>` runs whose
//     concatenation cannot be expressed as a single byte slice. Plain
//     `<v>123</v>` and `<t>hello</t>` payloads are returned as
//     zero-copy `string_view`s into the input.
//   * The element vocabulary recognised by the parser is intentionally
//     narrow (sheetData, row, c, f, v, is, t, r). Everything else is
//     skipped via a depth-counted scan that walks past nested elements
//     in O(content) time.

#include "io/sax_xml_reader.h"

// On WASM the OOXML reader's `kSaxThresholdBytes` is `SIZE_MAX`, so
// `read_sheet_data_sax` is statically unreachable and the linker
// drops both it and `scan_sheet_data`. `#if`-guarding the entire TU
// off in that build configuration removes the *compile-time* cost of
// the streaming scanner as well, keeping `formulon_core` lean.
//
// Native builds and any WASM build that opts in via
// `-DFORMULON_WASM_ENABLE_SAX=1` keep the implementation. Tests live
// on the native build path so they always see the full
// implementation.
#if !defined(FORMULON_WASM) || defined(FORMULON_WASM_ENABLE_SAX)

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>

#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace {

constexpr std::size_t kMaxColumnLetters = 3U;  // Excel max column = XFD.

/// Returns true for ASCII whitespace per the XML spec (S production).
bool IsXmlSpace(char c) noexcept {
  return c == ' ' || c == '\t' || c == '\r' || c == '\n';
}

/// Skips ASCII whitespace at `*p`, advancing the pointer.
void SkipSpace(const char* end, const char** p) noexcept {
  while (*p < end && IsXmlSpace(**p)) {
    ++(*p);
  }
}

/// Builds a `kIoXmlParse` error. We intentionally drop the per-call-
/// site message and offset hint: keeping a single static error keeps
/// the linker from emitting per-callsite rodata strings + `to_string`
/// instantiations, which together cost noticeable WASM bytes. The
/// scanner is malformed-input-resistant; callers get a structured
/// error code (`kIoXmlParse`) and a generic message.
Error MakeXmlParseError(std::size_t /*offset*/, const char* /*what*/) {
  return make_error(FormulonErrorCode::kIoXmlParse, "sax: malformed sheet xml", "context=sax_xml_reader");
}

/// Decodes an Excel A1 column reference at `text[*i]` into a 1-based
/// column index. Returns 0 when no letter is present or the column
/// exceeds Excel's XFD limit (3 letters). Advances `*i` over the
/// consumed letters.
std::uint32_t DecodeColumnLetters(std::string_view text, std::size_t* i) noexcept {
  std::uint32_t col = 0;
  std::size_t consumed = 0;
  while (*i < text.size()) {
    const char c = text[*i];
    if (c < 'A' || c > 'Z') {
      break;
    }
    if (consumed >= kMaxColumnLetters) {
      return 0U;  // Past XFD.
    }
    col = col * 26U + static_cast<std::uint32_t>(c - 'A' + 1);
    ++(*i);
    ++consumed;
  }
  if (consumed == 0U) {
    return 0U;
  }
  return col;
}

/// Decodes an unsigned decimal integer at `text[*i]`. Advances `*i`
/// past the digits. Returns 0 when no digit was consumed (callers treat
/// 0 as a sentinel because Excel rows are 1-based).
std::uint32_t DecodeUnsigned(std::string_view text, std::size_t* i) noexcept {
  std::uint64_t v = 0;
  bool any = false;
  while (*i < text.size()) {
    const char c = text[*i];
    if (c < '0' || c > '9') {
      break;
    }
    v = v * 10U + static_cast<std::uint64_t>(c - '0');
    if (v > 0xFFFFFFFFULL) {
      return 0U;  // overflow guard
    }
    ++(*i);
    any = true;
  }
  return any ? static_cast<std::uint32_t>(v) : 0U;
}

/// Decodes an A1-shaped cell reference (e.g. `"AB12"`) into 0-based row
/// + 0-based column. Returns false on malformed input.
bool DecodeA1(std::string_view ref, std::uint32_t* row_out, std::uint32_t* col_out) noexcept {
  if (ref.empty()) {
    return false;
  }
  std::size_t i = 0;
  const std::uint32_t col_1based = DecodeColumnLetters(ref, &i);
  if (col_1based == 0U) {
    return false;
  }
  const std::uint32_t row_1based = DecodeUnsigned(ref, &i);
  if (row_1based == 0U) {
    return false;
  }
  if (i != ref.size()) {
    return false;
  }
  *row_out = row_1based - 1U;
  *col_out = col_1based - 1U;
  return true;
}

// ---------------------------------------------------------------------------
// Tag scanner.
// ---------------------------------------------------------------------------

/// Outcome of scanning one element header (between `<` and the matching
/// `>` or `/>`).
struct TagHeader {
  std::string_view name;       // local name, no namespace prefix stripped
  std::string_view raw_attrs;  // raw bytes of the attribute region (post-name, pre-`>`/`/>`)
  bool self_closing = false;   // true if the tag ended with `/>`
  bool is_end_tag = false;     // true if this was a `</name>` close tag
  std::size_t end_offset = 0;  // byte offset just past the closing `>` / `/>`
};

/// Skips the rest of an XML processing instruction (`<? ... ?>`),
/// comment (`<!-- ... -->`), or CDATA (`<![CDATA[ ... ]]>`). On entry
/// `*p` points just past the leading `<`. Returns true on success.
bool SkipMarkupNonElement(const char* end, const char** p, std::size_t base_offset, Error* err) {
  if (*p >= end) {
    *err = MakeXmlParseError(base_offset, "unterminated markup at '<'");
    return false;
  }
  const char first = **p;
  if (first == '?') {
    // Processing instruction: scan to "?>".
    ++(*p);
    while (*p + 1 < end) {
      if ((*p)[0] == '?' && (*p)[1] == '>') {
        *p += 2;
        return true;
      }
      ++(*p);
    }
    *err = MakeXmlParseError(base_offset, "unterminated <? processing instruction");
    return false;
  }
  if (first == '!') {
    // Comment "<!--" or CDATA "<![CDATA[" or DOCTYPE.
    ++(*p);
    if (*p + 1 < end && (*p)[0] == '-' && (*p)[1] == '-') {
      *p += 2;  // past "--"
      while (*p + 2 < end) {
        if ((*p)[0] == '-' && (*p)[1] == '-' && (*p)[2] == '>') {
          *p += 3;
          return true;
        }
        ++(*p);
      }
      *err = MakeXmlParseError(base_offset, "unterminated <!-- comment");
      return false;
    }
    if (*p + 7 < end && std::memcmp(*p, "[CDATA[", 7) == 0) {
      *p += 7;
      while (*p + 2 < end) {
        if ((*p)[0] == ']' && (*p)[1] == ']' && (*p)[2] == '>') {
          *p += 3;
          return true;
        }
        ++(*p);
      }
      *err = MakeXmlParseError(base_offset, "unterminated <![CDATA[ section");
      return false;
    }
    // Generic <!...>: scan to matching '>'.
    while (*p < end) {
      if (**p == '>') {
        ++(*p);
        return true;
      }
      ++(*p);
    }
    *err = MakeXmlParseError(base_offset, "unterminated <! markup");
    return false;
  }
  return false;  // not a non-element markup; caller handles as a tag.
}

/// Reads one element header starting at `*p` (which must point just
/// past `<`). On success advances `*p` to byte after `>` / `/>` and
/// fills `out`. Returns false on malformed input (`*err` set).
bool ParseTagHeader(const char* begin, const char* end, const char** p, TagHeader* out, Error* err) {
  const std::size_t base = static_cast<std::size_t>(*p - begin);
  if (*p >= end) {
    *err = MakeXmlParseError(base, "unexpected end of input inside tag");
    return false;
  }
  out->is_end_tag = (**p == '/');
  if (out->is_end_tag) {
    ++(*p);
  }
  // Element name = chars while not whitespace, '/', '>'.
  const char* name_begin = *p;
  while (*p < end) {
    const char c = **p;
    if (c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '/' || c == '>') {
      break;
    }
    ++(*p);
  }
  if (*p == name_begin) {
    *err = MakeXmlParseError(base, "empty element name");
    return false;
  }
  out->name = std::string_view(name_begin, static_cast<std::size_t>(*p - name_begin));

  // Attribute span (or whitespace-then-close) up to the closing > / />.
  const char* attrs_begin = *p;
  while (*p < end) {
    const char c = **p;
    if (c == '>') {
      out->raw_attrs = std::string_view(attrs_begin, static_cast<std::size_t>(*p - attrs_begin));
      out->self_closing = false;
      ++(*p);
      out->end_offset = static_cast<std::size_t>(*p - begin);
      return true;
    }
    if (c == '/' && *p + 1 < end && (*p)[1] == '>') {
      out->raw_attrs = std::string_view(attrs_begin, static_cast<std::size_t>(*p - attrs_begin));
      out->self_closing = true;
      *p += 2;
      out->end_offset = static_cast<std::size_t>(*p - begin);
      return true;
    }
    ++(*p);
  }
  *err = MakeXmlParseError(base, "unterminated tag (no closing '>' / '/>')");
  return false;
}

/// Reads attribute "key=\"value\"" pairs out of `raw_attrs`. Quotes may
/// be `"` or `'`. Stops on first malformed input and returns the
/// captured pair count so far (callers tolerate partial sets — the
/// vocabulary we care about is small and well-formed in practice).
struct AttrIterator {
  std::string_view rest;

  /// Pulls the next attribute. Returns false when no more remain.
  bool next(std::string_view* key, std::string_view* value) noexcept {
    // Skip leading whitespace.
    while (!rest.empty() && IsXmlSpace(rest.front())) {
      rest.remove_prefix(1);
    }
    if (rest.empty()) {
      return false;
    }
    // Key = chars up to '=' (no whitespace permitted in well-formed XML
    // but be lenient).
    std::size_t i = 0;
    while (i < rest.size() && rest[i] != '=' && !IsXmlSpace(rest[i])) {
      ++i;
    }
    if (i == 0 || i == rest.size()) {
      return false;
    }
    *key = rest.substr(0, i);
    rest.remove_prefix(i);
    while (!rest.empty() && IsXmlSpace(rest.front())) {
      rest.remove_prefix(1);
    }
    if (rest.empty() || rest.front() != '=') {
      return false;
    }
    rest.remove_prefix(1);  // '='
    while (!rest.empty() && IsXmlSpace(rest.front())) {
      rest.remove_prefix(1);
    }
    if (rest.empty()) {
      return false;
    }
    const char quote = rest.front();
    if (quote != '"' && quote != '\'') {
      return false;
    }
    rest.remove_prefix(1);
    std::size_t j = 0;
    while (j < rest.size() && rest[j] != quote) {
      ++j;
    }
    if (j == rest.size()) {
      return false;
    }
    *value = rest.substr(0, j);
    rest.remove_prefix(j + 1);  // past closing quote
    return true;
  }
};

/// Looks up `name` among the attributes of `header` and returns its
/// value, or an empty `string_view` when absent.
std::string_view AttrOf(const TagHeader& header, std::string_view name) noexcept {
  AttrIterator it{header.raw_attrs};
  std::string_view k;
  std::string_view v;
  while (it.next(&k, &v)) {
    if (k == name) {
      return v;
    }
  }
  return {};
}

// ---------------------------------------------------------------------------
// Text decoding.
// ---------------------------------------------------------------------------

/// Extracts the text content of an element whose opening tag has just
/// been parsed (`*p` points at the first byte after `>`). Returns the
/// raw byte slice of the content (possibly with entities), advances
/// `*p` to the byte after `</name>`, and sets `had_entities` when the
/// slice contains at least one `&...;` reference.
///
/// Caller decides whether to decode the entities (cheap path: if
/// `had_entities` is false, the slice can be returned verbatim as a
/// `string_view` into the input).
bool ScanTextContent(const char* begin, const char* end, const char** p, std::string_view tag_name,
                     std::string_view* raw_content, bool* had_entities, Error* err) {
  const char* content_begin = *p;
  *had_entities = false;
  while (*p < end) {
    if (**p == '<') {
      *raw_content = std::string_view(content_begin, static_cast<std::size_t>(*p - content_begin));
      // Expect `</tag_name>`.
      ++(*p);  // past '<'
      if (*p < end && **p == '/') {
        ++(*p);  // past '/'
        const char* close_name_begin = *p;
        while (*p < end && **p != '>' && !IsXmlSpace(**p)) {
          ++(*p);
        }
        const std::string_view close_name(close_name_begin, static_cast<std::size_t>(*p - close_name_begin));
        if (close_name != tag_name) {
          *err = MakeXmlParseError(static_cast<std::size_t>(close_name_begin - begin), "unexpected close tag");
          return false;
        }
        SkipSpace(end, p);
        if (*p >= end || **p != '>') {
          *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unterminated close tag");
          return false;
        }
        ++(*p);  // past '>'
        return true;
      }
      // Anything other than `</...>` inside a `<v>` / `<t>` body is
      // unsupported in our recognised vocabulary. Surface as parse error.
      *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unexpected child element inside text content");
      return false;
    }
    if (**p == '&') {
      *had_entities = true;
    }
    ++(*p);
  }
  *err = MakeXmlParseError(static_cast<std::size_t>(content_begin - begin), "unterminated text content");
  return false;
}

/// Decodes the five predefined XML entity references (`&amp; &lt;
/// &gt; &quot; &apos;`) into `out`. Numeric character references and
/// any other `&...;` form pass through verbatim — Excel does not emit
/// numeric refs into worksheet payloads, so a strict five-name table
/// keeps the SAX reader compact.
void DecodeEntitiesInto(std::string_view src, std::string* out) {
  out->clear();
  out->reserve(src.size());
  for (std::size_t i = 0; i < src.size();) {
    const char c = src[i];
    if (c != '&') {
      out->push_back(c);
      ++i;
      continue;
    }
    const std::size_t semi = src.find(';', i + 1);
    if (semi == std::string_view::npos || semi > i + 8) {
      out->push_back('&');
      ++i;
      continue;
    }
    const std::string_view ent = src.substr(i + 1, semi - i - 1);
    char repl = 0;
    if (ent == "amp") {
      repl = '&';
    } else if (ent == "lt") {
      repl = '<';
    } else if (ent == "gt") {
      repl = '>';
    } else if (ent == "quot") {
      repl = '"';
    } else if (ent == "apos") {
      repl = '\'';
    }
    if (repl != 0) {
      out->push_back(repl);
    } else {
      out->append(src.data() + i, semi - i + 1);
    }
    i = semi + 1;
  }
}

// ---------------------------------------------------------------------------
// Element skip / cell / row / sheetData walkers.
// ---------------------------------------------------------------------------

/// Skips an element whose opening header has just been consumed (and is
/// not self-closing). `*p` is at the byte after `>`. The skip is
/// depth-counted so nested elements are walked in O(content) time
/// without decoding their children.
bool SkipUntilClose(const char* begin, const char* end, const char** p, std::string_view name, Error* err) {
  std::size_t depth = 1;
  while (*p < end && depth > 0) {
    // Scan to next '<'.
    while (*p < end && **p != '<') {
      ++(*p);
    }
    if (*p >= end) {
      break;
    }
    const std::size_t lt_off = static_cast<std::size_t>(*p - begin);
    ++(*p);  // past '<'
    if (*p < end && (**p == '?' || **p == '!')) {
      Error sub{};
      if (!SkipMarkupNonElement(end, p, lt_off, &sub)) {
        *err = sub;
        return false;
      }
      continue;
    }
    TagHeader header;
    if (!ParseTagHeader(begin, end, p, &header, err)) {
      return false;
    }
    if (header.is_end_tag) {
      // We do not validate name nesting strictly; the well-formed
      // assumption is good enough for Excel-emitted output. We still
      // stop when we see a close at depth 1 with the matching name.
      --depth;
      if (depth == 0 && header.name != name) {
        // Tolerated: pugixml does the same for foreign nesting. We
        // continue rather than failing because the outer scanner only
        // cares about returning to the surrounding context.
      }
      continue;
    }
    if (!header.self_closing) {
      ++depth;
    }
  }
  if (depth != 0) {
    *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unterminated element");
    return false;
  }
  return true;
}

/// State carried across the per-cell scan so allocations amortise:
///   * `decoded_value`  - holds entity-decoded `<v>` body.
///   * `inline_string`  - holds the concatenation of `<is><t>` /
///                        `<is><r><t>` payloads.
///   * `decoded_formula` - holds the entity-decoded `<f>` body when
///                         decoding is required. Kept distinct from
///                         the others so `record->formula` and
///                         `record->value` can coexist for the rare
///                         well-formed input that has both (e.g. a
///                         hand-rolled `t="inlineStr"` cell with an
///                         `<f>` child).
struct CellScratch {
  std::string decoded_value;
  std::string inline_string;
  std::string decoded_formula;
};

/// Reads `<is> ... </is>` content, concatenating every `<t>` body
/// encountered along the way (whether at the top level or inside an
/// `<r>` rich-text run). Foreign elements are tolerated and their
/// text content is dropped.
///
/// The opening `<is>` header has already been consumed; on entry
/// `*p` points just past `>`.
///
/// Implementation: rather than tracking nested-element structure
/// strictly, we stream-find every `<t>...</t>` pair within the `<is>`
/// body and bail when we see the matching `</is>`. This is correct
/// for Excel-emitted output (which only ever nests `<t>` under `<is>`
/// or `<r>`) and shaves substantial code size compared to a strict
/// recursive walk.
bool ScanInlineString(const char* begin, const char* end, const char** p, CellScratch* scratch, Error* err) {
  scratch->inline_string.clear();
  // Tracks whether we are currently inside an <rPh> (phonetic guide)
  // subtree. <rPh> wraps a <t> element carrying kana, which must NOT
  // be concatenated into the surface text — otherwise PHONETIC's
  // surface-text fallback for unannotated cells would silently include
  // the kana too. The DOM path's `cell_parser.cpp` extracts <rPh>
  // separately into ParsedCell::phonetic_text; the SAX path does not
  // yet capture phonetic, but at minimum it must avoid mixing kana
  // bytes into the inline-string body.
  bool in_rph = false;
  while (*p < end) {
    while (*p < end && **p != '<') {
      ++(*p);
    }
    if (*p >= end) {
      break;
    }
    const std::size_t lt_off = static_cast<std::size_t>(*p - begin);
    ++(*p);  // past '<'
    if (*p < end && (**p == '?' || **p == '!')) {
      Error sub{};
      if (!SkipMarkupNonElement(end, p, lt_off, &sub)) {
        *err = sub;
        return false;
      }
      continue;
    }
    TagHeader header;
    if (!ParseTagHeader(begin, end, p, &header, err)) {
      return false;
    }
    if (header.is_end_tag && header.name == "is") {
      return true;
    }
    if (header.is_end_tag && header.name == "rPh") {
      in_rph = false;
      continue;
    }
    if (header.is_end_tag) {
      // Close of an inner element (e.g. </r>): keep streaming.
      continue;
    }
    if (header.self_closing) {
      continue;
    }
    if (header.name == "rPh") {
      in_rph = true;
      continue;
    }
    if (header.name == "t") {
      std::string_view raw;
      bool had_entities = false;
      if (!ScanTextContent(begin, end, p, "t", &raw, &had_entities, err)) {
        return false;
      }
      if (in_rph) {
        // Inside <rPh>: drop the kana; we still had to consume the
        // <t>...</t> body to advance the cursor past the matching
        // close tag.
        continue;
      }
      if (had_entities) {
        std::string tmp;
        DecodeEntitiesInto(raw, &tmp);
        scratch->inline_string.append(tmp);
      } else {
        scratch->inline_string.append(raw.data(), raw.size());
      }
    }
    // Other open elements (e.g. <r>, <rPr>) are descended into by
    // simply continuing the loop; their <t> children will be picked
    // up on subsequent iterations and their close tags hit the
    // is_end_tag branch above. <rPr> contents have no <t> so the
    // walk passes through them harmlessly.
  }
  *err = MakeXmlParseError(0, "unterminated <is>");
  return false;
}

/// Reads a `<c ...>` element to completion, populating `record`. The
/// opening tag has already been parsed into `header`. On a self-closing
/// `<c .../>` we just emit the record with empty `formula` / `value`.
bool ScanCell(const char* begin, const char* end, const char** p, const TagHeader& cell_header, CellRecord* record,
              CellScratch* scratch, Error* err) {
  // Decode r=, t=, s=.
  const std::string_view r_attr = AttrOf(cell_header, "r");
  if (r_attr.empty()) {
    *err = MakeXmlParseError(cell_header.end_offset, "<c> missing r= attribute");
    return false;
  }
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  if (!DecodeA1(r_attr, &row, &col)) {
    std::string ctx("context=sax_xml_reader r=");
    ctx.append(r_attr);
    *err = make_error(FormulonErrorCode::kIoXmlParse, "sax: malformed <c r=...>", std::move(ctx));
    return false;
  }
  record->row = row;
  record->col = col;
  record->t = AttrOf(cell_header, "t");
  record->s = AttrOf(cell_header, "s");
  record->formula = std::string_view{};
  record->value = std::string_view{};
  record->is_inline_string = false;

  if (cell_header.self_closing) {
    return true;
  }

  while (*p < end) {
    // Skip any whitespace between children.
    while (*p < end && IsXmlSpace(**p)) {
      ++(*p);
    }
    if (*p >= end) {
      *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unterminated <c>");
      return false;
    }
    if (**p != '<') {
      // Excel never emits text directly inside <c>; if some producer
      // does, treat as parse error rather than silently dropping it.
      *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unexpected text inside <c>");
      return false;
    }
    const std::size_t lt_off = static_cast<std::size_t>(*p - begin);
    ++(*p);  // past '<'
    if (*p < end && (**p == '?' || **p == '!')) {
      Error sub{};
      if (!SkipMarkupNonElement(end, p, lt_off, &sub)) {
        *err = sub;
        return false;
      }
      continue;
    }
    TagHeader child;
    if (!ParseTagHeader(begin, end, p, &child, err)) {
      return false;
    }
    if (child.is_end_tag) {
      if (child.name != "c") {
        *err = MakeXmlParseError(lt_off, "unexpected close tag inside <c>");
        return false;
      }
      return true;
    }
    if (child.name == "f") {
      if (child.self_closing) {
        continue;
      }
      std::string_view raw;
      bool had_entities = false;
      if (!ScanTextContent(begin, end, p, "f", &raw, &had_entities, err)) {
        return false;
      }
      // Strip leading '=' if present.
      if (!raw.empty() && raw.front() == '=') {
        raw.remove_prefix(1);
      }
      if (had_entities) {
        DecodeEntitiesInto(raw, &scratch->decoded_formula);
        record->formula = std::string_view(scratch->decoded_formula);
      } else {
        record->formula = raw;
      }
    } else if (child.name == "v") {
      if (child.self_closing) {
        continue;
      }
      std::string_view raw;
      bool had_entities = false;
      if (!ScanTextContent(begin, end, p, "v", &raw, &had_entities, err)) {
        return false;
      }
      if (had_entities) {
        DecodeEntitiesInto(raw, &scratch->decoded_value);
        record->value = std::string_view(scratch->decoded_value);
      } else {
        record->value = raw;
      }
    } else if (child.name == "is") {
      if (child.self_closing) {
        record->is_inline_string = true;
        record->value = std::string_view{};
        continue;
      }
      if (!ScanInlineString(begin, end, p, scratch, err)) {
        return false;
      }
      record->is_inline_string = true;
      record->value = std::string_view(scratch->inline_string);
    } else if (!child.self_closing) {
      // Unrecognised child of <c>: skip it.
      if (!SkipUntilClose(begin, end, p, child.name, err)) {
        return false;
      }
    }
  }
  *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unterminated <c>");
  return false;
}

/// Reads a `<row>` element to completion, dispatching `<c>` children
/// through `cb.on_cell`. The opening tag has already been parsed into
/// `header`.
bool ScanRow(const char* begin, const char* end, const char** p, const TagHeader& row_header,
             const SheetSaxCallbacks& cb, CellScratch* scratch, Error* err) {
  // Decode r= (1-based).
  std::uint32_t row_1based = 0;
  const std::string_view r_attr = AttrOf(row_header, "r");
  if (!r_attr.empty()) {
    std::size_t i = 0;
    row_1based = DecodeUnsigned(r_attr, &i);
  }

  if (cb.on_row_start != nullptr) {
    auto rs = cb.on_row_start(cb.user_data, row_1based);
    if (!rs) {
      *err = rs.error();
      return false;
    }
  }
  if (row_header.self_closing) {
    if (cb.on_row_end != nullptr) {
      auto re = cb.on_row_end(cb.user_data, row_1based);
      if (!re) {
        *err = re.error();
        return false;
      }
    }
    return true;
  }

  while (*p < end) {
    while (*p < end && IsXmlSpace(**p)) {
      ++(*p);
    }
    if (*p >= end) {
      *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unterminated <row>");
      return false;
    }
    if (**p != '<') {
      *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unexpected text inside <row>");
      return false;
    }
    const std::size_t lt_off = static_cast<std::size_t>(*p - begin);
    ++(*p);
    if (*p < end && (**p == '?' || **p == '!')) {
      Error sub{};
      if (!SkipMarkupNonElement(end, p, lt_off, &sub)) {
        *err = sub;
        return false;
      }
      continue;
    }
    TagHeader child;
    if (!ParseTagHeader(begin, end, p, &child, err)) {
      return false;
    }
    if (child.is_end_tag) {
      if (child.name != "row") {
        *err = MakeXmlParseError(lt_off, "unexpected close tag inside <row>");
        return false;
      }
      if (cb.on_row_end != nullptr) {
        auto re = cb.on_row_end(cb.user_data, row_1based);
        if (!re) {
          *err = re.error();
          return false;
        }
      }
      return true;
    }
    if (child.name == "c") {
      CellRecord rec;
      if (!ScanCell(begin, end, p, child, &rec, scratch, err)) {
        return false;
      }
      if (cb.on_cell != nullptr) {
        auto cr = cb.on_cell(cb.user_data, rec);
        if (!cr) {
          *err = cr.error();
          return false;
        }
      }
    } else if (!child.self_closing) {
      // Unrecognised <row> child: skip.
      if (!SkipUntilClose(begin, end, p, child.name, err)) {
        return false;
      }
    }
  }
  *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unterminated <row>");
  return false;
}

/// Walks `<sheetData>` to completion, dispatching its `<row>` children
/// through the row scanner. The opening `<sheetData>` header has just
/// been parsed.
bool ScanSheetData(const char* begin, const char* end, const char** p, const TagHeader& sd_header,
                   const SheetSaxCallbacks& cb, CellScratch* scratch, Error* err) {
  if (sd_header.self_closing) {
    return true;
  }
  while (*p < end) {
    while (*p < end && IsXmlSpace(**p)) {
      ++(*p);
    }
    if (*p >= end) {
      *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unterminated <sheetData>");
      return false;
    }
    if (**p != '<') {
      *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unexpected text inside <sheetData>");
      return false;
    }
    const std::size_t lt_off = static_cast<std::size_t>(*p - begin);
    ++(*p);
    if (*p < end && (**p == '?' || **p == '!')) {
      Error sub{};
      if (!SkipMarkupNonElement(end, p, lt_off, &sub)) {
        *err = sub;
        return false;
      }
      continue;
    }
    TagHeader child;
    if (!ParseTagHeader(begin, end, p, &child, err)) {
      return false;
    }
    if (child.is_end_tag) {
      if (child.name != "sheetData") {
        *err = MakeXmlParseError(lt_off, "unexpected close tag inside <sheetData>");
        return false;
      }
      return true;
    }
    if (child.name == "row") {
      if (!ScanRow(begin, end, p, child, cb, scratch, err)) {
        return false;
      }
    } else if (!child.self_closing) {
      if (!SkipUntilClose(begin, end, p, child.name, err)) {
        return false;
      }
    }
  }
  *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unterminated <sheetData>");
  return false;
}

}  // namespace

Expected<void, Error> scan_sheet_data(ByteSpan xml, const SheetSaxCallbacks& callbacks) {
  if (xml.data == nullptr || xml.size == 0) {
    return Expected<void, Error>::Ok();
  }
  const char* begin = reinterpret_cast<const char*>(xml.data);
  const char* end = begin + xml.size;
  const char* p = begin;

  CellScratch scratch;

  // Walk the document until we find <sheetData>, descending into the
  // wrapper element (<worksheet> in well-formed Excel output) rather
  // than skipping it. Non-<sheetData> siblings (dimension, sheetViews,
  // cols, mergeCells, hyperlinks, conditionalFormatting, ...) are
  // skipped opaquely with `SkipUntilClose`, which never invokes the
  // callbacks.
  bool descended_into_wrapper = false;
  while (p < end) {
    while (p < end && p[0] != '<') {
      ++p;
    }
    if (p >= end) {
      break;
    }
    const std::size_t lt_off = static_cast<std::size_t>(p - begin);
    ++p;  // past '<'
    if (p < end && (p[0] == '?' || p[0] == '!')) {
      Error sub{};
      if (!SkipMarkupNonElement(end, &p, lt_off, &sub)) {
        return sub;
      }
      continue;
    }
    TagHeader header;
    Error parse_err{};
    if (!ParseTagHeader(begin, end, &p, &header, &parse_err)) {
      return parse_err;
    }
    if (header.is_end_tag) {
      // Close of the wrapper element (e.g. </worksheet>): done.
      return Expected<void, Error>::Ok();
    }
    if (header.name == "sheetData") {
      Error sub{};
      if (!ScanSheetData(begin, end, &p, header, callbacks, &scratch, &sub)) {
        return sub;
      }
      // Continue: there may be siblings (mergeCells, etc.) we just
      // skip past; eventually the wrapper close tag ends the scan.
      continue;
    }
    // Either the document root wrapper (descend) or a non-sheetData
    // sibling (skip). We descend exactly once: the outermost open
    // element is the wrapper.
    if (header.self_closing) {
      continue;
    }
    if (!descended_into_wrapper) {
      descended_into_wrapper = true;
      continue;  // walk into the wrapper's children
    }
    // A non-sheetData sibling of <sheetData>: opaquely skip past it.
    Error sub{};
    if (!SkipUntilClose(begin, end, &p, header.name, &sub)) {
      return sub;
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace io
}  // namespace formulon

#endif  // !FORMULON_WASM || FORMULON_WASM_ENABLE_SAX
