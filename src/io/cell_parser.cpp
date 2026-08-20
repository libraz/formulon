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
//   * `t="d"` (ISO 8601 date string, strict OOXML) is parsed into an Excel
//     serial number so date arithmetic (`A1+1`, `YEAR(A1)`) matches Excel.
//     A body that does not parse as a strict ISO date falls back to `Text`
//     rather than failing the cell, keeping the parser tolerant of
//     non-conforming producers.

#include "io/cell_parser.h"

#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <deque>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/a1_ref.h"
#include "io/iso_date.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_double.h"
#include "io/xsd_int.h"
#include "phonetic.h"
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

/// Thin wrappers over the shared A1 helpers in `io/a1_ref.h`, preserving
/// the legacy "0 = error" sentinel that the original local helpers used.
/// Excel rows are 1-based so a 0 result is unambiguously an error.
std::uint32_t ParseUintAdvance(std::string_view text, std::size_t* i) {
  std::uint32_t v = 0;
  if (!parse_uint(text, i, &v)) {
    return 0U;
  }
  return v;
}

std::uint32_t ParseColumnLetters(std::string_view text, std::size_t* i) {
  std::uint32_t col = 0;
  if (!parse_column_letters(text, i, &col)) {
    return 0U;
  }
  return col;
}

/// Walks `is_node`'s descendants and concatenates every `<t>` text node
/// payload into `out`, in document order. Thin wrapper over the shared
/// `append_rich_text` helper (which skips `<rPh>` so phonetic guides do
/// not leak into the surface string; `CollectInlinePhoneticRuns` walks
/// them separately).
void ConcatInlineStringText(const pugi::xml_node& is_node, std::string& out) {
  (void)append_rich_text(is_node, out);
}

/// Collects every `<rPh>` direct child of `is_node` into `out`, one run
/// per block, in document order. The `sb`/`eb` attributes delimit the
/// surface-text span each block's kana covers, in UTF-16 code units, and
/// are preserved: PHONETIC replaces only the annotated spans and passes
/// the rest of the string through, so the boundaries are observable.
/// Mirrors `CollectPhoneticRuns` on the shared-strings side.
void CollectInlinePhoneticRuns(const pugi::xml_node& is_node, std::vector<PhoneticRun>& out) {
  for (pugi::xml_node rph = is_node.child("rPh"); rph; rph = rph.next_sibling("rPh")) {
    PhoneticRun run;
    run.sb = static_cast<std::uint32_t>(rph.attribute("sb").as_uint(0U));
    run.eb = static_cast<std::uint32_t>(rph.attribute("eb").as_uint(0U));
    for (pugi::xml_node t = rph.child("t"); t; t = t.next_sibling("t")) {
      AppendOoxmlTextUnescaped(run.text, t.text().get());
    }
    out.push_back(std::move(run));
  }
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

Expected<ParsedCell, Error> decode_cell_payload(std::string_view t, std::string_view v_text, bool value_present,
                                                bool is_inline_string, std::deque<std::string>& text_storage) {
  ParsedCell out;
  if (t.empty()) {
    t = std::string_view("n");
  }
  const bool t_ok = (t == "n" || t == "b" || t == "e" || t == "s" || t == "str" || t == "inlineStr" || t == "d");
  if (!t_ok) {
    std::string ctx("context=cell_parser t=");
    ctx.append(t);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: unsupported t= value", std::move(ctx));
  }

  if (t == "inlineStr" || is_inline_string) {
    if (!value_present) {
      out.value = Value::text(std::string_view{});
      return out;
    }
    text_storage.emplace_back(std::string(v_text));
    out.value = Value::text(text_storage.back());
    return out;
  }
  if (t == "str") {
    if (value_present) {
      text_storage.emplace_back(std::string(v_text));
      out.value = Value::text(text_storage.back());
    } else {
      out.value = Value::text(std::string_view{});
    }
    return out;
  }
  if (t == "s") {
    if (!value_present) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='s' without <v> child",
                        "context=cell_parser");
    }
    // Parse the SST index directly as a 32-bit unsigned integer. Going
    // through `double` here would silently round indices >= 2^24+1 onto
    // even neighbours and surface the wrong shared string. The pugixml
    // text() payload is NUL-terminated so strtoul is safe.
    std::size_t i = 0;
    const std::uint32_t parsed_idx = ParseUintAdvance(v_text, &i);
    // ParseUintAdvance returns 0 both for "no digits consumed" and for
    // overflow past 2^32-1. Distinguish: empty / non-digit / overflow
    // leaves `i == 0` or `i < v_text.size()`.
    if (i == 0U || i != v_text.size()) {
      std::string ctx("context=cell_parser v=");
      ctx.append(v_text);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='s' has unparseable index", std::move(ctx));
    }
    out.is_sst_index = true;
    out.sst_index = parsed_idx;
    out.value = Value::text(std::string_view{});
    return out;
  }
  if (t == "b") {
    if (!value_present) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='b' without <v> child",
                        "context=cell_parser");
    }
    if (v_text == "1") {
      out.value = Value::boolean(true);
    } else if (v_text == "0") {
      out.value = Value::boolean(false);
    } else {
      std::string ctx("context=cell_parser v=");
      ctx.append(v_text);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='b' must be 0 or 1", std::move(ctx));
    }
    return out;
  }
  if (t == "e") {
    if (!value_present) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='e' without <v> child",
                        "context=cell_parser");
    }
    ErrorCode code{};
    if (!ParseErrorDisplay(v_text, &code)) {
      std::string ctx("context=cell_parser v=");
      ctx.append(v_text);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: unrecognised error display", std::move(ctx));
    }
    out.value = Value::error(code);
    return out;
  }
  if (t == "d") {
    if (!value_present) {
      out.value = Value::text(std::string_view{});
      return out;
    }
    double serial = 0.0;
    if (parse_iso_date_serial(v_text, &serial)) {
      out.value = Value::number(serial);
    } else {
      // Non-conforming producer: keep the raw text rather than failing.
      text_storage.emplace_back(std::string(v_text));
      out.value = Value::text(text_storage.back());
    }
    return out;
  }
  // Default / t == "n".
  // Excel and several third-party producers use an empty <v/> for a
  // cached blank numeric cell. Treat it exactly like an absent <v>; this
  // keeps the DOM reader aligned with the streaming reader, which has no
  // payload to distinguish in either form.
  if (!value_present || v_text.empty()) {
    out.value = Value::blank();
    return out;
  }
  double num = 0.0;
  if (!parse_xsd_double(v_text, &num)) {
    std::string ctx("context=cell_parser v=");
    ctx.append(v_text);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='n' has unparseable <v>", std::move(ctx));
  }
  out.value = Value::number(num);
  return out;
}

Expected<ParsedCell, Error> parse_cell_element(const pugi::xml_node& node, std::deque<std::string>& text_storage) {
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
  std::uint32_t row = rc_or.value().first;
  std::uint32_t col = rc_or.value().second;

  // t= — type.
  std::string_view t = node.attribute("t").value();

  // s= — style index (xf row in xl/styles.xml). Absent or outside the
  // shared non-negative-integer lexical space degrades to the schema
  // default 0: a style index is cosmetic, so rejecting the whole workbook
  // over one malformed value would trade a formatting loss for a load
  // failure. The streaming reader applies the same lexer and the same
  // disposition to the same attribute.
  std::uint32_t xf_index = 0;
  if (!parse_xsd_nonneg_int(attr_str(node, "s"), &xf_index)) {
    xf_index = 0;
  }

  // <f>...</f> — formula text (leading '=' stripped).
  std::string formula;
  pugi::xml_node f_node = node.child("f");
  if (f_node) {
    formula = f_node.text().get();
    if (!formula.empty() && formula.front() == '=') {
      formula.erase(0, 1);
    }
  }

  // <v> / <is> — value. Inline-string concatenation is done eagerly
  // into a scratch buffer so the shared `decode_cell_payload` only
  // sees a flat `string_view`.
  pugi::xml_node v_node = node.child("v");
  pugi::xml_node is_node = node.child("is");
  std::string inline_concat;
  std::string_view v_text;
  bool value_present = false;
  bool is_inline = false;
  if (is_node) {
    is_inline = true;
    value_present = true;
    ConcatInlineStringText(is_node, inline_concat);
    v_text = inline_concat;
  } else if (v_node) {
    value_present = true;
    v_text = v_node.text().get();
  }

  auto payload_or = decode_cell_payload(t, v_text, value_present, is_inline, text_storage);
  if (!payload_or) {
    // Re-decorate the error context with the cell ref for actionable
    // diagnostics.
    Error e = payload_or.error();
    if (!ref.empty()) {
      e.context.append(" ref=").append(ref);
    }
    return e;
  }
  ParsedCell out = payload_or.value();
  out.row = row;
  out.col = col;
  out.formula = std::move(formula);
  out.xf_index = xf_index;

  // Capture <rPh> phonetic runs from the inline-string block (when
  // present). SST-referenced cells (t="s") carry their phonetic on the
  // matching <si> in xl/sharedStrings.xml; that path is plumbed through
  // `SharedStringTable::phonetic_for_entries` instead. The runs own
  // their kana, so unlike the cell's surface text they do not borrow
  // from `text_storage`.
  if (is_node) {
    CollectInlinePhoneticRuns(is_node, out.phonetic_runs);
  }
  return out;
}

}  // namespace io
}  // namespace formulon
