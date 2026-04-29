// Copyright 2026 libraz. Licensed under the MIT License.
//
// `<c>` element parser. See cell_parser.h for the public contract.
//
// Implementation notes:
//   * We do NOT use `strtod` directly here — instead we hand off to the
//     shared `parse_a1`-free `std::strtod` path used elsewhere; the OOXML
//     `<v>` payload is always C-locale ASCII (Excel emits `%.17g`-style
//     printouts, never localised), so a plain `std::strtod` is sufficient.
//   * Inline-string handling walks the `<is>` element's text nodes in
//     document order, concatenating `<t>` children. Rich-text shape (a
//     sequence of `<r><t>...</t></r>`) is preserved as plain text since
//     this slice has no rich-text storage.
//   * `t="str"` is the legacy formula-result-as-string shape; we treat it
//     identically to `t="inlineStr"` because the cached value is the only
//     thing we need.
//   * `t="d"` (ISO date string) is accepted but routed through a placeholder:
//     we surface `Text` rather than parsing into a serial number, because
//     date conversion is the styles-aware path (Bundle 2.3+). Keeps the
//     parser usable on date-bearing fixtures without lying about semantics.

#include "io/cell_parser.h"

#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <deque>
#include <string>
#include <string_view>
#include <utility>

#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace io {
namespace {

/// Tries to map an Excel error display name (e.g. `"#DIV/0!"`) to its
/// `ErrorCode`. Returns `false` on unknown spellings; the caller surfaces
/// `kIoSheetCorrupt`.
bool ParseErrorDisplay(std::string_view text, ErrorCode* out) {
  if (text == "#NULL!") {
    *out = ErrorCode::Null;
    return true;
  }
  if (text == "#DIV/0!") {
    *out = ErrorCode::Div0;
    return true;
  }
  if (text == "#VALUE!") {
    *out = ErrorCode::Value;
    return true;
  }
  if (text == "#REF!") {
    *out = ErrorCode::Ref;
    return true;
  }
  if (text == "#NAME?") {
    *out = ErrorCode::Name;
    return true;
  }
  if (text == "#NUM!") {
    *out = ErrorCode::Num;
    return true;
  }
  if (text == "#N/A") {
    *out = ErrorCode::NA;
    return true;
  }
  if (text == "#GETTING_DATA") {
    *out = ErrorCode::GettingData;
    return true;
  }
  if (text == "#SPILL!") {
    *out = ErrorCode::Spill;
    return true;
  }
  if (text == "#CALC!") {
    *out = ErrorCode::Calc;
    return true;
  }
  if (text == "#FIELD!") {
    *out = ErrorCode::Field;
    return true;
  }
  if (text == "#BLOCKED!") {
    *out = ErrorCode::Blocked;
    return true;
  }
  if (text == "#CONNECT!") {
    *out = ErrorCode::Connect;
    return true;
  }
  if (text == "#EXTERNAL!") {
    *out = ErrorCode::External;
    return true;
  }
  if (text == "#BUSY!") {
    *out = ErrorCode::Busy;
    return true;
  }
  if (text == "#PYTHON!") {
    *out = ErrorCode::Python;
    return true;
  }
  if (text == "#UNKNOWN!") {
    *out = ErrorCode::Unknown;
    return true;
  }
  return false;
}

/// Parses an unsigned decimal integer at `text[*i]`. Advances `*i` past
/// the digits. Returns the parsed value, or `0` when no digit was
/// consumed (a row index of 0 is invalid in Excel — rows are 1-based —
/// so callers can use 0 as a sentinel). Caps at 32-bit unsigned.
std::uint32_t ParseUintAdvance(std::string_view text, std::size_t* i) {
  std::uint64_t v = 0;
  bool any = false;
  while (*i < text.size()) {
    const char c = text[*i];
    if (c < '0' || c > '9') {
      break;
    }
    v = v * 10U + static_cast<std::uint64_t>(c - '0');
    if (v > 0xFFFFFFFFULL) {
      // Overflow guard. Caller treats 0 as "invalid" already.
      return 0;
    }
    ++(*i);
    any = true;
  }
  return any ? static_cast<std::uint32_t>(v) : 0U;
}

/// Parses Excel column letters (A-Z, AA-XFD) starting at `text[*i]` into
/// a 1-based column index. Returns 0 on error (no letter consumed) or on
/// overflow past `XFD`. Advances `*i`.
std::uint32_t ParseColumnLetters(std::string_view text, std::size_t* i) {
  std::uint32_t col = 0;
  std::size_t consumed = 0;
  while (*i < text.size()) {
    const char c = text[*i];
    if (c < 'A' || c > 'Z') {
      break;
    }
    if (consumed >= 3U) {
      // Excel max is XFD (3 letters); reject 4+ letter columns.
      return 0U;
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

/// Walks `is_node`'s descendants and concatenates every `<t>` text node
/// payload into `out`, in document order. `<r><t>...</t></r>` rich-text
/// runs are flattened: this slice does not preserve formatting because
/// the workbook model has no rich-text storage yet.
void ConcatInlineStringText(const pugi::xml_node& is_node, std::string& out) {
  // Direct `<t>` (the simple inlineStr shape).
  for (pugi::xml_node t = is_node.child("t"); t; t = t.next_sibling("t")) {
    out.append(t.text().get());
  }
  // `<r><t>...</t></r>` rich-text runs. Walked in document order so the
  // catenation matches the original visual order.
  for (pugi::xml_node r = is_node.child("r"); r; r = r.next_sibling("r")) {
    for (pugi::xml_node rt = r.child("t"); rt; rt = rt.next_sibling("t")) {
      out.append(rt.text().get());
    }
  }
}

/// Parses a `<v>` text node as a double. Returns false on non-empty
/// trailing characters or empty input; `out` is unchanged in that case.
bool ParseDouble(std::string_view text, double* out) {
  if (text.empty()) {
    return false;
  }
  // strtod requires a NUL-terminated string. The pugixml `text().get()`
  // we receive already is, but the slice we hand in may have been built
  // from a `string_view` substring (we currently never do that); copy
  // into a small stack buffer when oversized to stay safe.
  char small_buf[64];
  const char* nstr = nullptr;
  std::string heap;
  if (text.size() < sizeof(small_buf)) {
    std::memcpy(small_buf, text.data(), text.size());
    small_buf[text.size()] = '\0';
    nstr = small_buf;
  } else {
    heap.assign(text.data(), text.size());
    nstr = heap.c_str();
  }
  char* end = nullptr;
  const double v = std::strtod(nstr, &end);
  if (end == nstr) {
    return false;
  }
  // Trailing garbage is not allowed.
  while (end != nullptr && *end != '\0') {
    if (*end != ' ' && *end != '\t' && *end != '\r' && *end != '\n') {
      return false;
    }
    ++end;
  }
  *out = v;
  return true;
}

}  // namespace

Expected<std::pair<std::uint32_t, std::uint32_t>, Error> parse_a1(std::string_view ref) {
  if (ref.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell ref: empty", "context=cell_parser parse_a1=<empty>");
  }
  // Reject sheet-qualified or range shapes.
  if (ref.find('!') != std::string_view::npos || ref.find(':') != std::string_view::npos) {
    std::string ctx("context=cell_parser parse_a1=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell ref: must not be qualified or a range", std::move(ctx));
  }
  // Reject `$` markers: OOXML cell refs in `<c r=...>` are always plain.
  if (ref.find('$') != std::string_view::npos) {
    std::string ctx("context=cell_parser parse_a1=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell ref: must not contain '$'", std::move(ctx));
  }

  std::size_t i = 0;
  const std::uint32_t col_1based = ParseColumnLetters(ref, &i);
  if (col_1based == 0U) {
    std::string ctx("context=cell_parser parse_a1=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell ref: missing or oversized column letters",
                      std::move(ctx));
  }
  const std::uint32_t row_1based = ParseUintAdvance(ref, &i);
  if (row_1based == 0U) {
    std::string ctx("context=cell_parser parse_a1=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell ref: missing or invalid row number", std::move(ctx));
  }
  if (i != ref.size()) {
    std::string ctx("context=cell_parser parse_a1=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell ref: trailing characters", std::move(ctx));
  }
  // Bounds check: Excel 365 limits.
  if (col_1based > Sheet::kMaxCols || row_1based > Sheet::kMaxRows) {
    std::string ctx("context=cell_parser parse_a1=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell ref: out of Excel range", std::move(ctx));
  }
  return std::pair<std::uint32_t, std::uint32_t>{row_1based - 1U, col_1based - 1U};
}

Expected<ParsedCell, Error> parse_cell_element(const pugi::xml_node& node, std::deque<std::string>& text_storage) {
  ParsedCell out;

  // r=A1 — required.
  pugi::xml_attribute r_attr = node.attribute("r");
  if (!r_attr) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: <c> missing 'r' attribute",
                      "context=cell_parser");
  }
  std::string_view ref = r_attr.value();
  auto rc_or = parse_a1(ref);
  if (!rc_or) {
    return rc_or.error();
  }
  out.row = rc_or.value().first;
  out.col = rc_or.value().second;

  // t= — type. Default is "n" (number) when absent.
  std::string_view t = node.attribute("t").value();
  if (t.empty()) {
    t = "n";
  }
  // Recognised set: n, b, e, s, str, inlineStr, d. Anything else is an
  // explicit error: silently mapping unknown types to a default would
  // lose data (and round-trip would fail).
  const bool t_ok = (t == "n" || t == "b" || t == "e" || t == "s" || t == "str" || t == "inlineStr" || t == "d");
  if (!t_ok) {
    std::string ctx("context=cell_parser ref=");
    ctx.append(ref);
    ctx.append(" t=");
    ctx.append(t);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: unsupported t= value", std::move(ctx));
  }

  // <f>...</f> — formula. The OOXML cached value is held in the sibling
  // <v>, which we read below regardless of formula presence.
  pugi::xml_node f_node = node.child("f");
  if (f_node) {
    // Strip leading '=' if present (Excel's writer never emits one, but
    // be defensive for other producers).
    std::string formula = f_node.text().get();
    if (!formula.empty() && formula.front() == '=') {
      formula.erase(0, 1);
    }
    out.formula = std::move(formula);
  }

  // <v> / <is> — value.
  pugi::xml_node v_node = node.child("v");
  pugi::xml_node is_node = node.child("is");

  if (t == "inlineStr" || t == "str") {
    // inlineStr: child is `<is>`; flatten rich-text.
    // str (legacy formula-as-string result): cached value lives in `<v>`
    //   as plain text rather than a number.
    if (t == "inlineStr") {
      if (!is_node) {
        // inlineStr without <is>: blank text rather than crash. Excel
        // accepts this shape on read.
        out.value = Value::text(std::string_view{});
        return out;
      }
      // `Value::text` is non-owning; append the decoded string to the
      // caller's `text_storage` (a pointer-stable container) so the
      // resulting `string_view` outlives this function. The deque's
      // back() reference stays valid until the caller drops the
      // container.
      text_storage.emplace_back();
      ConcatInlineStringText(is_node, text_storage.back());
      out.value = Value::text(text_storage.back());
      return out;
    }
    // t == "str"
    if (v_node) {
      text_storage.emplace_back(v_node.text().get());
      out.value = Value::text(text_storage.back());
    } else {
      out.value = Value::text(std::string_view{});
    }
    return out;
  }

  if (t == "s") {
    // Shared string: surface index, write a placeholder. Bundle 2.3
    // resolves this against the workbook's SST.
    if (!v_node) {
      std::string ctx("context=cell_parser ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='s' without <v> child", std::move(ctx));
    }
    double idx = 0.0;
    if (!ParseDouble(v_node.text().get(), &idx) || idx < 0.0 || idx > 4294967295.0) {
      std::string ctx("context=cell_parser ref=");
      ctx.append(ref);
      ctx.append(" v=");
      ctx.append(v_node.text().get());
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='s' has unparseable index", std::move(ctx));
    }
    out.is_sst_index = true;
    out.sst_index = static_cast<std::uint32_t>(idx);
    // Empty placeholder; writer will replace once SST is loaded.
    out.value = Value::text(std::string_view{});
    return out;
  }

  if (t == "b") {
    if (!v_node) {
      std::string ctx("context=cell_parser ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='b' without <v> child", std::move(ctx));
    }
    const std::string_view body = v_node.text().get();
    if (body == "1") {
      out.value = Value::boolean(true);
    } else if (body == "0") {
      out.value = Value::boolean(false);
    } else {
      std::string ctx("context=cell_parser ref=");
      ctx.append(ref);
      ctx.append(" v=");
      ctx.append(body);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='b' must be 0 or 1", std::move(ctx));
    }
    return out;
  }

  if (t == "e") {
    if (!v_node) {
      std::string ctx("context=cell_parser ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='e' without <v> child", std::move(ctx));
    }
    ErrorCode code{};
    const std::string_view body = v_node.text().get();
    if (!ParseErrorDisplay(body, &code)) {
      std::string ctx("context=cell_parser ref=");
      ctx.append(ref);
      ctx.append(" v=");
      ctx.append(body);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: unrecognised error display", std::move(ctx));
    }
    out.value = Value::error(code);
    return out;
  }

  if (t == "d") {
    // ISO-8601 date string. Excel rarely emits this and resolving it to
    // a serial number is the styles layer's job; surface as Text for
    // round-trip preservation.
    if (v_node) {
      text_storage.emplace_back(v_node.text().get());
      out.value = Value::text(text_storage.back());
    } else {
      out.value = Value::text(std::string_view{});
    }
    return out;
  }

  // Default / t == "n": number, or blank if no <v>.
  if (!v_node) {
    // No cached value. For formula cells this is legal (Excel will
    // recalc on load); for literal cells this means a truly blank
    // numeric cell, which we map to Blank.
    out.value = Value::blank();
    return out;
  }
  double num = 0.0;
  if (!ParseDouble(v_node.text().get(), &num)) {
    std::string ctx("context=cell_parser ref=");
    ctx.append(ref);
    ctx.append(" v=");
    ctx.append(v_node.text().get());
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='n' has unparseable <v>", std::move(ctx));
  }
  out.value = Value::number(num);
  return out;
}

}  // namespace io
}  // namespace formulon
