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

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>

#include "io/a1_ref.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace {

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

/// Decodes an unsigned decimal integer at `text[*i]`. Wraps the shared
/// `parse_uint` helper, returning 0 (a sentinel because Excel rows are
/// 1-based) on no-digit or overflow input.
std::uint32_t DecodeUnsigned(std::string_view text, std::size_t* i) noexcept {
  std::uint32_t v = 0;
  if (!parse_uint(text, i, &v)) {
    return 0U;
  }
  return v;
}

/// Wraps the shared `parse_a1_ref` helper. Preserved as a local thin
/// shim so the rest of this TU keeps reading naturally.
bool DecodeA1(std::string_view ref, std::uint32_t* row_out, std::uint32_t* col_out) noexcept {
  return parse_a1_ref(ref, row_out, col_out);
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

/// Forward declaration: decodes the five predefined XML entity references
/// (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`) plus numeric character
/// references (`&#N;` / `&#xN;`) into `out`. The full implementation lives
/// further down with the `<v>` / `<t>` text-node helpers.
void DecodeEntitiesInto(std::string_view src, std::string* out);

/// True when `src` contains at least one `&` character — i.e. the byte
/// span needs entity decoding before it can be exposed to callers as
/// canonical text. Cheap O(N) scan; used as a fast-path gate before the
/// (relatively expensive) full decoder kicks in.
bool ContainsEntities(std::string_view src) noexcept {
  return src.find('&') != std::string_view::npos;
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

/// Looks up `name` among the attributes of `header` and returns its raw
/// value, or an empty `string_view` when absent. The returned view aliases
/// into the input buffer and may carry XML entity references such as
/// `&amp;` or `&#xNN;` verbatim — callers that need canonical text must
/// route the result through `AttrOfDecoded` instead.
std::string_view AttrOfRaw(const TagHeader& header, std::string_view name) noexcept {
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

/// Same as `AttrOfRaw` but transparently decodes XML entity references
/// (`&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;`, `&#N;`, `&#xN;`) into
/// canonical UTF-8 text. When decoding is required the result is
/// materialised into `*scratch` and the returned view aliases into that
/// buffer; the caller must keep `*scratch` alive for the lifetime of the
/// returned view. When the value contains no `&`, the function returns a
/// zero-copy view into the input buffer (and `*scratch` is not touched).
///
/// Excel-emitted worksheet output rarely needs decoding here — `r="A1"`,
/// `t="n"` and friends are pure ASCII — but third-party producers and
/// hand-rolled fixtures do exercise the entity path (e.g. `r="A&amp;amp;
/// B1"`-shaped pathological cases that, while ill-typed, are still
/// well-formed XML). Skipping decoding produces a `kIoSheetCorrupt` parse
/// failure that masquerades as a bad reference; decoding here keeps the
/// SAX path symmetric with the text-node path (`<v>` / `<t>`) which has
/// always run through `DecodeEntitiesInto`.
std::string_view AttrOfDecoded(const TagHeader& header, std::string_view name, std::string* scratch) {
  const std::string_view raw = AttrOfRaw(header, name);
  if (!ContainsEntities(raw)) {
    return raw;
  }
  DecodeEntitiesInto(raw, scratch);
  return std::string_view(*scratch);
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

/// Encodes a Unicode code point into UTF-8 bytes and appends them to
/// `out`. Returns false when `cp` is outside the legal Unicode range
/// (> 0x10FFFF) or falls inside the surrogate band — the caller leaves
/// the original `&...;` text in place when this happens. The branchy
/// shape mirrors the canonical UTF-8 encoder; an encoding-table fallback
/// would not be cheaper at this size and would cost rodata.
bool AppendUtf8(std::uint32_t cp, std::string* out) {
  if (cp <= 0x7FU) {
    out->push_back(static_cast<char>(cp));
    return true;
  }
  if (cp <= 0x7FFU) {
    out->push_back(static_cast<char>(0xC0U | (cp >> 6U)));
    out->push_back(static_cast<char>(0x80U | (cp & 0x3FU)));
    return true;
  }
  if (cp <= 0xFFFFU) {
    if (cp >= 0xD800U && cp <= 0xDFFFU) {
      return false;  // Surrogate halves are not legal in UTF-8.
    }
    out->push_back(static_cast<char>(0xE0U | (cp >> 12U)));
    out->push_back(static_cast<char>(0x80U | ((cp >> 6U) & 0x3FU)));
    out->push_back(static_cast<char>(0x80U | (cp & 0x3FU)));
    return true;
  }
  if (cp <= 0x10FFFFU) {
    out->push_back(static_cast<char>(0xF0U | (cp >> 18U)));
    out->push_back(static_cast<char>(0x80U | ((cp >> 12U) & 0x3FU)));
    out->push_back(static_cast<char>(0x80U | ((cp >> 6U) & 0x3FU)));
    out->push_back(static_cast<char>(0x80U | (cp & 0x3FU)));
    return true;
  }
  return false;
}

/// Decodes the five predefined XML entity references (`&amp; &lt; &gt;
/// &quot; &apos;`) plus decimal (`&#N;`) and hex (`&#xN;`) numeric
/// character references into `out`. Unknown / malformed `&...;` forms
/// are emitted verbatim — Excel does not exercise this path in
/// well-formed output, but third-party producers occasionally emit
/// numeric refs, and some hand-rolled fixtures use entity-encoded
/// attribute values that the attribute iterator now plumbs through this
/// helper.
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
    // Numeric character references can be longer than the longest named
    // entity; bound the search so a stray `&` followed by hostile junk
    // does not scan to end-of-buffer. 12 chars covers `&#x10FFFF;`.
    constexpr std::size_t kMaxEntityScan = 12U;
    const std::size_t scan_end = std::min(src.size(), i + 1U + kMaxEntityScan);
    std::size_t semi = std::string_view::npos;
    for (std::size_t j = i + 1; j < scan_end; ++j) {
      if (src[j] == ';') {
        semi = j;
        break;
      }
    }
    if (semi == std::string_view::npos) {
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
      i = semi + 1;
      continue;
    }
    if (!ent.empty() && ent.front() == '#') {
      std::uint32_t cp = 0;
      bool ok = false;
      if (ent.size() >= 2 && (ent[1] == 'x' || ent[1] == 'X')) {
        // Hex numeric reference: &#xNN;.
        ok = ent.size() > 2;
        for (std::size_t j = 2; j < ent.size() && ok; ++j) {
          const char d = ent[j];
          std::uint32_t v = 0;
          if (d >= '0' && d <= '9') {
            v = static_cast<std::uint32_t>(d - '0');
          } else if (d >= 'a' && d <= 'f') {
            v = static_cast<std::uint32_t>(d - 'a' + 10);
          } else if (d >= 'A' && d <= 'F') {
            v = static_cast<std::uint32_t>(d - 'A' + 10);
          } else {
            ok = false;
            break;
          }
          cp = (cp << 4U) | v;
          if (cp > 0x10FFFFU) {
            ok = false;
            break;
          }
        }
      } else {
        // Decimal numeric reference: &#N;.
        ok = ent.size() > 1;
        for (std::size_t j = 1; j < ent.size() && ok; ++j) {
          const char d = ent[j];
          if (d < '0' || d > '9') {
            ok = false;
            break;
          }
          cp = cp * 10U + static_cast<std::uint32_t>(d - '0');
          if (cp > 0x10FFFFU) {
            ok = false;
            break;
          }
        }
      }
      if (ok && AppendUtf8(cp, out)) {
        i = semi + 1;
        continue;
      }
    }
    // Unknown entity: pass through verbatim.
    out->append(src.data() + i, semi - i + 1);
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

/// Reads the next child element header inside an already-open parent.
/// Whitespace and non-element markup are consumed transparently; text
/// directly under the parent is treated as malformed input, matching the
/// previous per-parent scanner behavior.
bool ReadChildHeader(const char* begin, const char* end, const char** p, Error* err, TagHeader* out,
                     std::size_t* lt_off) {
  while (true) {
    while (*p < end && IsXmlSpace(**p)) {
      ++(*p);
    }
    if (*p >= end) {
      *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unterminated parent element");
      return false;
    }
    if (**p != '<') {
      *err = MakeXmlParseError(static_cast<std::size_t>(*p - begin), "unexpected text inside parent element");
      return false;
    }
    *lt_off = static_cast<std::size_t>(*p - begin);
    ++(*p);  // past '<'
    if (*p < end && (**p == '?' || **p == '!')) {
      Error sub{};
      if (!SkipMarkupNonElement(end, p, *lt_off, &sub)) {
        *err = sub;
        return false;
      }
      continue;
    }
    return ParseTagHeader(begin, end, p, out, err);
  }
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
  /// Concatenation of every `<rPh><t>` payload encountered while
  /// scanning an `<is>` body. Cleared at the start of each
  /// `ScanInlineString` call.
  std::string inline_string_phonetic;
  /// Scratch buffers for entity-decoded `<c>` / `<row>` attribute values.
  /// Each buffer is owned by the surrounding cell / row scan, so the
  /// `string_view`s returned by `AttrOfDecoded` remain valid for the
  /// duration of the dispatching callback. Three slots cover every
  /// attribute the SAX vocabulary inspects (`r=`, `t=`, `s=` on `<c>`,
  /// plus `r=` on `<row>`); the row attribute reuses the cell's `r`
  /// slot because the two never overlap in time.
  std::string decoded_attr_r;
  std::string decoded_attr_t;
  std::string decoded_attr_s;
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
  scratch->inline_string_phonetic.clear();
  // Tracks whether we are currently inside an <rPh> (phonetic guide)
  // subtree. <rPh> wraps a <t> element carrying kana, which must NOT
  // be concatenated into the surface text — otherwise PHONETIC's
  // surface-text fallback for unannotated cells would silently include
  // the kana too. The kana is captured separately into
  // `inline_string_phonetic` so PHONETIC can return it for inline-
  // string cells the SAX path materialises.
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
      std::string* dest = in_rph ? &scratch->inline_string_phonetic : &scratch->inline_string;
      if (had_entities) {
        std::string tmp;
        DecodeEntitiesInto(raw, &tmp);
        dest->append(tmp);
      } else {
        dest->append(raw.data(), raw.size());
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
  // Decode r=, t=, s=. Attributes are routed through `AttrOfDecoded` so
  // entity-encoded values (e.g. `r="A&amp;1"`) decode to canonical text
  // before the A1 / type-token validation runs.
  const std::string_view r_attr = AttrOfDecoded(cell_header, "r", &scratch->decoded_attr_r);
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
  record->t = AttrOfDecoded(cell_header, "t", &scratch->decoded_attr_t);
  record->s = AttrOfDecoded(cell_header, "s", &scratch->decoded_attr_s);
  record->formula = std::string_view{};
  record->f_t = std::string_view{};
  record->f_si = std::string_view{};
  record->f_ref = std::string_view{};
  record->value = std::string_view{};
  record->is_inline_string = false;
  record->phonetic = std::string_view{};

  if (cell_header.self_closing) {
    return true;
  }

  while (*p < end) {
    TagHeader child;
    std::size_t lt_off = 0;
    if (!ReadChildHeader(begin, end, p, err, &child, &lt_off)) {
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
      // Capture the shared/array-formula attributes before handling the
      // body. `t`/`si`/`ref` never carry XML entities in practice, so the
      // raw views (aliasing the source buffer) are safe to surface. A
      // shared-formula follower is self-closing with no body, but its
      // `si` still has to reach the consumer so the group master can be
      // shifted into place.
      record->f_t = AttrOfRaw(child, "t");
      record->f_si = AttrOfRaw(child, "si");
      record->f_ref = AttrOfRaw(child, "ref");
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
        record->phonetic = std::string_view{};
        continue;
      }
      if (!ScanInlineString(begin, end, p, scratch, err)) {
        return false;
      }
      record->is_inline_string = true;
      record->value = std::string_view(scratch->inline_string);
      record->phonetic = std::string_view(scratch->inline_string_phonetic);
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
  // Decode r= (1-based). The row's `r` attribute reuses the cell-level
  // `decoded_attr_r` scratch because the row scan is single-shot before
  // any cell scan begins.
  std::uint32_t row_1based = 0;
  const std::string_view r_attr = AttrOfDecoded(row_header, "r", &scratch->decoded_attr_r);
  if (!r_attr.empty()) {
    std::size_t i = 0;
    row_1based = DecodeUnsigned(r_attr, &i);
  }

  if (cb.on_row_start != nullptr) {
    // Surface the row's override attributes. These are numeric / boolean
    // and never entity-encoded, so raw views (aliasing the buffer) are
    // safe for the callback's duration.
    RowRecord rec;
    rec.row_1based = row_1based;
    rec.ht = AttrOfRaw(row_header, "ht");
    rec.hidden = AttrOfRaw(row_header, "hidden");
    rec.custom_height = AttrOfRaw(row_header, "customHeight");
    rec.outline_level = AttrOfRaw(row_header, "outlineLevel");
    auto rs = cb.on_row_start(cb.user_data, rec);
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
    TagHeader child;
    std::size_t lt_off = 0;
    if (!ReadChildHeader(begin, end, p, err, &child, &lt_off)) {
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
    TagHeader child;
    std::size_t lt_off = 0;
    if (!ReadChildHeader(begin, end, p, err, &child, &lt_off)) {
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
