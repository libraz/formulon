#pragma once

#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <iterator>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "tests/oracle/json_reader.h"
#include "tests/oracle/oracle_anchor.h"
#include "tests/oracle/oracle_runner.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon::tests::oracle::support {

struct SetupFormulaCell {
  std::size_t sheet_index = 0;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string formula;
};

// Maps the golden JSON's `"#DIV/0!"` style display name to the matching
// ErrorCode enum. Returns `false` on an unknown code.
inline bool display_name_to_code(std::string_view name, ErrorCode* out) {
  struct Entry {
    std::string_view display;
    ErrorCode code;
  };
  static constexpr Entry kTable[] = {
      {"#NULL!", ErrorCode::Null},
      {"#DIV/0!", ErrorCode::Div0},
      {"#VALUE!", ErrorCode::Value},
      {"#REF!", ErrorCode::Ref},
      {"#NAME?", ErrorCode::Name},
      {"#NUM!", ErrorCode::Num},
      {"#N/A", ErrorCode::NA},
      {"#SPILL!", ErrorCode::Spill},
      {"#CALC!", ErrorCode::Calc},
      {"#FIELD!", ErrorCode::Field},
      {"#BLOCKED!", ErrorCode::Blocked},
      {"#CONNECT!", ErrorCode::Connect},
      {"#EXTERNAL!", ErrorCode::External},
      {"#BUSY!", ErrorCode::Busy},
      {"#PYTHON!", ErrorCode::Python},
      {"#UNKNOWN!", ErrorCode::Unknown},
  };
  for (const auto& e : kTable) {
    if (e.display == name) {
      *out = e.code;
      return true;
    }
  }
  return false;
}

// Applies a single {kind, value} JSON record to a cell. Returns nullptr on
// success or a short error string on failure (unknown kind, missing value,
// etc.) that the caller folds into the test failure message.
inline const char* apply_cell_value(const JsonValue& spec, Sheet& sheet, std::uint32_t row, std::uint32_t col,
                                    Arena& text_arena) {
  if (!spec.is_object())
    return "setup cell is not an object";
  const JsonValue* kind_v = spec.find("kind");
  if (kind_v == nullptr || !kind_v->is_string())
    return "missing 'kind'";
  const std::string& kind = kind_v->as_string();

  if (kind == "formula") {
    const JsonValue* f = spec.find("formula");
    if (f == nullptr || !f->is_string())
      return "formula cell missing 'formula'";
    sheet.set_cell_formula(row, col, f->as_string());
    return nullptr;
  }

  const JsonValue* val_v = spec.find("value");
  if (kind == "blank") {
    // Nothing to do: absence from storage already means blank.
    return nullptr;
  }
  if (kind == "number") {
    if (val_v == nullptr || !val_v->is_number())
      return "number missing 'value'";
    sheet.set_cell_value(row, col, Value::number(val_v->as_number()));
    return nullptr;
  }
  if (kind == "bool") {
    if (val_v == nullptr || !val_v->is_bool())
      return "bool missing 'value'";
    sheet.set_cell_value(row, col, Value::boolean(val_v->as_bool()));
    return nullptr;
  }
  if (kind == "text") {
    if (val_v == nullptr || !val_v->is_string())
      return "text missing 'value'";
    // Intern the string into the arena so the Value's string_view remains
    // valid for the lifetime of the test.
    const std::string& payload = val_v->as_string();
    std::string_view view = text_arena.intern(payload);
    sheet.set_cell_value(row, col, Value::text(view));
    return nullptr;
  }
  if (kind == "error") {
    const JsonValue* code_v = spec.find("code");
    if (code_v == nullptr || !code_v->is_string())
      return "error missing 'code'";
    ErrorCode code;
    if (!display_name_to_code(code_v->as_string(), &code)) {
      return "error has unknown 'code'";
    }
    sheet.set_cell_value(row, col, Value::error(code));
    return nullptr;
  }
  return "setup cell has unknown 'kind'";
}

// Applies the optional case-level `merges` list emitted by oracle_gen. The
// formula drivers use Excel's inclusive A1 range syntax; keep the native
// verifier on the same representation and normalise transposed corners just
// as the C API merge surface does.
inline const char* apply_merge_ranges(const JsonValue& spec, Sheet& sheet) {
  if (!spec.is_array()) {
    return "case 'merges' is not an array";
  }
  for (const JsonValue& item : spec.as_array()) {
    if (!item.is_string() || item.as_string().empty()) {
      return "merge range is not a non-empty A1 string";
    }
    const std::string& ref = item.as_string();
    const std::size_t separator = ref.find(':');
    if (separator != std::string::npos && ref.find(':', separator + 1U) != std::string::npos) {
      return "merge range contains more than one ':'";
    }
    const std::string first_ref = ref.substr(0, separator);
    const std::string last_ref = separator == std::string::npos ? first_ref : ref.substr(separator + 1U);
    std::uint32_t first_row = 0;
    std::uint32_t first_col = 0;
    std::uint32_t last_row = 0;
    std::uint32_t last_col = 0;
    if (!a1_to_row_col(first_ref, &first_row, &first_col) || !a1_to_row_col(last_ref, &last_row, &last_col)) {
      return "merge range contains an invalid A1 address";
    }
    if (first_row > last_row) {
      const std::uint32_t tmp = first_row;
      first_row = last_row;
      last_row = tmp;
    }
    if (first_col > last_col) {
      const std::uint32_t tmp = first_col;
      first_col = last_col;
      last_col = tmp;
    }
    sheet.mutable_merges().push_back(MergeRange{first_row, first_col, last_row, last_col});
  }
  return nullptr;
}

// Renders a Formulon Value as a short display string for assertion failure
// messages. Mirrors the debug_to_string helper but keeps error / text
// formatting consistent with the JSON schema.
inline std::string format_value(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Blank:
      return "blank";
    case ValueKind::Number:
      return "number(" + std::to_string(v.as_number()) + ")";
    case ValueKind::Bool:
      return v.as_boolean() ? "bool(TRUE)" : "bool(FALSE)";
    case ValueKind::Text:
      return std::string("text(\"") + std::string(v.as_text()) + "\")";
    case ValueKind::Error:
      return std::string("error(") + display_name(v.as_error()) + ")";
    case ValueKind::Array:
      return "array(" + std::to_string(v.as_array_rows()) + "x" + std::to_string(v.as_array_cols()) + ")";
    default:
      return "<unsupported>";
  }
}

/// Result of parsing an Excel complex-number text representation.
///
/// Excel complex numbers are formatted as `<real><sign><imag><suffix>` (or
/// `<real>` alone, or `<sign><imag><suffix>` alone), where `<suffix>` is
/// `i` or `j`. See `parse_excel_complex` for the precise grammar.
struct ParsedComplex {
  double real = 0.0;
  double imag = 0.0;
  bool ok = false;
};

/// Parses an Excel complex-number text into its real and imaginary parts.
///
/// Recognised shapes:
///   `"3"`                                 -> {3, 0}
///   `"4i"`                                -> {0, 4}
///   `"i"`                                 -> {0, 1}
///   `"-i"`                                -> {0, -1}
///   `"-2.5+i"`                            -> {-2.5, 1}
///   `"0.62-0.30i"`                        -> {0.62, -0.30}
///   `"1.0e-3+2.5e+10i"`                   -> {1.0e-3, 2.5e10}
///
/// The split between real and imaginary parts is the LAST `+` or `-` in the
/// trimmed string that is neither at position 0 (a leading sign on the real
/// part) nor immediately after `e` / `E` (a scientific-notation exponent
/// sign).
inline ParsedComplex parse_excel_complex(std::string_view s) {
  ParsedComplex out;
  if (s.empty())
    return out;

  // Strip the imaginary suffix, if any. Excel uses `i`; Formulon currently
  // only emits `i`, but accept `j` for symmetry with the SUFFIX argument
  // family of IM* functions.
  bool has_imag = false;
  std::string_view body = s;
  const char last = body.back();
  if (last == 'i' || last == 'j') {
    has_imag = true;
    body.remove_suffix(1);
  }

  // Locate the split between real and imaginary parts.
  std::size_t split = std::string_view::npos;
  for (std::size_t i = body.size(); i-- > 0;) {
    const char c = body[i];
    if (c != '+' && c != '-')
      continue;
    if (i == 0)
      break;  // leading sign on the real part
    const char prev = body[i - 1];
    if (prev == 'e' || prev == 'E')
      continue;  // exponent sign
    split = i;
    break;
  }

  const auto parse_double = [](std::string_view sv, double* out_d) {
    if (sv.empty())
      return false;
    // strtod requires a NUL-terminated buffer; copy into a small std::string.
    const std::string buf(sv);
    const char* begin = buf.c_str();
    char* end = nullptr;
    const double v = std::strtod(begin, &end);
    if (end != begin + buf.size())
      return false;
    *out_d = v;
    return true;
  };

  if (split == std::string_view::npos) {
    if (has_imag) {
      // Whole trimmed body is the imaginary coefficient.
      if (body.empty() || body == "+") {
        out.imag = 1.0;
      } else if (body == "-") {
        out.imag = -1.0;
      } else if (!parse_double(body, &out.imag)) {
        return out;
      }
      out.real = 0.0;
    } else {
      if (!parse_double(body, &out.real))
        return out;
      out.imag = 0.0;
    }
    out.ok = true;
    return out;
  }

  // Split present: there must be an imaginary suffix (otherwise we have a
  // trailing sign with no `i`, which is malformed).
  if (!has_imag)
    return out;

  const std::string_view real_part = body.substr(0, split);
  const std::string_view imag_part = body.substr(split);  // includes sign
  if (!parse_double(real_part, &out.real))
    return out;
  if (imag_part == "+") {
    out.imag = 1.0;
  } else if (imag_part == "-") {
    out.imag = -1.0;
  } else if (!parse_double(imag_part, &out.imag)) {
    return out;
  }
  out.ok = true;
  return out;
}

// Component-wise comparator for Excel complex-number text. Returns an empty
// string on match (real and imag both within tolerance); otherwise a short
// diagnostic that quotes both sides verbatim.
inline std::string compare_complex_text(const std::string& want, const std::string& got, double tol_abs,
                                        double tol_rel) {
  const auto wc = parse_excel_complex(want);
  const auto gc = parse_excel_complex(got);
  if (!wc.ok || !gc.ok) {
    return "text mismatch (complex parse failed): expected \"" + want + "\", got \"" + got + "\"";
  }
  const auto component_ok = [&](double a, double b) {
    const double diff = std::abs(a - b);
    if (diff == 0.0)
      return true;
    if (tol_abs > 0.0 && diff <= tol_abs)
      return true;
    const double scale = std::max(std::abs(a), std::abs(b));
    if (tol_rel > 0.0 && scale > 0.0 && diff / scale <= tol_rel)
      return true;
    return false;
  };
  if (component_ok(wc.real, gc.real) && component_ok(wc.imag, gc.imag)) {
    return {};
  }
  return "complex mismatch (within text but components diverge): expected \"" + want + "\", got \"" + got + "\"";
}

inline bool parse_numeric_text(std::string_view text, double* out) {
  std::string owned(text);
  char* end = nullptr;
  double value = std::strtod(owned.c_str(), &end);
  if (end == owned.c_str()) {
    return false;
  }
  while (*end != '\0') {
    if (!std::isspace(static_cast<unsigned char>(*end))) {
      return false;
    }
    ++end;
  }
  *out = value;
  return true;
}

inline bool numbers_match(double want, double got, double tol_abs, double tol_rel) {
  if (std::isnan(want) && std::isnan(got)) {
    return true;
  }
  const double diff = std::abs(want - got);
  if (diff == 0.0) {
    return true;
  }
  if (tol_abs > 0.0 && diff <= tol_abs) {
    return true;
  }
  const double scale = std::max(std::abs(want), std::abs(got));
  return tol_rel > 0.0 && scale > 0.0 && diff / scale <= tol_rel;
}

inline std::string compare_value(const JsonValue& expect, const Value& raw_actual, double tol_abs, double tol_rel,
                                 std::string_view compare_mode);

inline std::string compare_json_scalar(const JsonValue& expect, const Value& actual, double tol_abs, double tol_rel,
                                       std::string_view compare_mode) {
  if (expect.is_null()) {
    return actual.is_blank() ? std::string{} : "expected blank, got " + format_value(actual);
  }
  if (expect.is_number()) {
    if (!actual.is_number())
      return "expected number, got " + format_value(actual);
    return numbers_match(expect.as_number(), actual.as_number(), tol_abs, tol_rel)
               ? std::string{}
               : "number mismatch: expected " + std::to_string(expect.as_number()) + ", got " +
                     std::to_string(actual.as_number());
  }
  if (expect.is_bool()) {
    if (!actual.is_boolean())
      return "expected bool, got " + format_value(actual);
    return actual.as_boolean() == expect.as_bool()
               ? std::string{}
               : std::string("bool mismatch: expected ") + (expect.as_bool() ? "TRUE" : "FALSE") + ", got " +
                     (actual.as_boolean() ? "TRUE" : "FALSE");
  }
  if (expect.is_string()) {
    if (!actual.is_text())
      return "expected text, got " + format_value(actual);
    if (expect.as_string() == actual.as_text())
      return {};
    if (compare_mode == "complex_text") {
      return compare_complex_text(expect.as_string(), std::string(actual.as_text()), tol_abs, tol_rel);
    }
    return "text mismatch: expected \"" + expect.as_string() + "\", got \"" + std::string(actual.as_text()) + "\"";
  }
  if (expect.is_object()) {
    return compare_value(expect, actual, tol_abs, tol_rel, compare_mode);
  }
  return "array golden cell is not scalar";
}

// Compares `actual` to the golden `expect` JSON record under the given
// tolerance. Returns an empty string on match; otherwise a human-readable
// diff message.
//
// `compare_mode` is "" or "exact" for the historical strict path, or a
// structured comparator key (e.g. "complex_text") that selects an
// alternative routine when the strict byte-equality check fails.
inline std::string compare_value(const JsonValue& expect, const Value& raw_actual, double tol_abs, double tol_rel,
                                 std::string_view compare_mode) {
  if (!expect.is_object())
    return "golden 'expect' is not an object";
  const JsonValue* kind_v = expect.find("kind");
  if (kind_v == nullptr || !kind_v->is_string()) {
    return "golden 'expect' missing 'kind'";
  }
  const std::string& kind = kind_v->as_string();

  if (kind == "array") {
    if (!raw_actual.is_array())
      return "expected array, got " + format_value(raw_actual);
    const JsonValue* shape_v = expect.find("shape");
    if (shape_v == nullptr || !shape_v->is_array())
      return "array golden missing 'shape'";
    const auto& shape = shape_v->as_array();
    if (shape.size() != 2 || !shape[0].is_number() || !shape[1].is_number())
      return "array golden has invalid 'shape'";
    const auto want_rows = static_cast<std::uint32_t>(shape[0].as_number());
    const auto want_cols = static_cast<std::uint32_t>(shape[1].as_number());
    if (raw_actual.as_array_rows() != want_rows || raw_actual.as_array_cols() != want_cols) {
      return "array shape mismatch: expected " + std::to_string(want_rows) + "x" + std::to_string(want_cols) +
             ", got " + std::to_string(raw_actual.as_array_rows()) + "x" + std::to_string(raw_actual.as_array_cols());
    }
    const JsonValue* value_v = expect.find("value");
    if (value_v == nullptr) {
      return {};
    }
    if (!value_v->is_array())
      return "array golden 'value' is not an array";
    const auto& expected_cells = value_v->as_array();
    const std::size_t n = static_cast<std::size_t>(want_rows) * static_cast<std::size_t>(want_cols);
    if (expected_cells.size() != n)
      return "array golden value length mismatch: expected " + std::to_string(n) + ", got " +
             std::to_string(expected_cells.size());
    const Value* actual_cells = raw_actual.as_array_cells();
    for (std::size_t i = 0; i < n; ++i) {
      std::string diff = compare_json_scalar(expected_cells[i], actual_cells[i], tol_abs, tol_rel, compare_mode);
      if (!diff.empty()) {
        return "array cell " + std::to_string(i) + ": " + diff;
      }
    }
    return {};
  }

  if (kind == "array_shape") {
    return "golden 'expect' kind 'array_shape' must be compared with compare_array_shape";
  }

  // Anchor projection: Excel reports the top-left scalar of any spill
  // region while xlwings reads back only that cell. Formulon's eval
  // surfaces the full Array; project it down so scalar expectations see
  // the same value Excel reports. Full array expectations above bypass
  // this projection and compare shape + cells.
  const Value& actual = anchor_or_self(raw_actual);

  if (kind == "blank") {
    if (actual.is_blank())
      return {};
    return "expected blank, got " + format_value(actual);
  }
  if (kind == "number") {
    if (!actual.is_number())
      return "expected number, got " + format_value(actual);
    const JsonValue* val_v = expect.find("value");
    if (val_v == nullptr || !val_v->is_number())
      return "golden missing 'value'";
    double want = val_v->as_number();
    double got = actual.as_number();
    if (compare_mode == "datevalue_roundtrip_readback" && want == -1.0 && got == 0.0)
      return {};
    // Exact equality is the strict path. When tolerances are non-zero, we
    // accept a match if either absolute or relative diff fits. NaN matches
    // NaN (Excel treats NaN as #NUM!, so this usually doesn't arise).
    if (numbers_match(want, got, tol_abs, tol_rel))
      return {};
    return "number mismatch: expected " + std::to_string(want) + ", got " + std::to_string(got);
  }
  if (kind == "bool") {
    if (!actual.is_boolean())
      return "expected bool, got " + format_value(actual);
    const JsonValue* val_v = expect.find("value");
    if (val_v == nullptr || !val_v->is_bool())
      return "golden missing 'value'";
    if (actual.as_boolean() == val_v->as_bool())
      return {};
    return std::string("bool mismatch: expected ") + (val_v->as_bool() ? "TRUE" : "FALSE") + ", got " +
           (actual.as_boolean() ? "TRUE" : "FALSE");
  }
  if (kind == "text") {
    const JsonValue* val_v = expect.find("value");
    if (val_v == nullptr || !val_v->is_string())
      return "golden missing 'value'";
    const std::string& want_text = val_v->as_string();

    if (compare_mode == "numeric_text" && actual.is_number()) {
      double want_number = 0.0;
      if (!parse_numeric_text(want_text, &want_number)) {
        return "expected numeric text but golden value is not numeric: \"" + want_text + "\"";
      }
      const double got = actual.as_number();
      if (numbers_match(want_number, got, tol_abs, tol_rel)) {
        return {};
      }
      return "numeric text mismatch: expected \"" + want_text + "\", got " + std::to_string(got);
    }

    if (!actual.is_text())
      return "expected text, got " + format_value(actual);
    const std::string actual_text(actual.as_text());

    // Strict byte compare wins fast on the common path; structured
    // comparators only run when the bytes already disagree.
    if (want_text == actual_text)
      return {};

    if (compare_mode == "complex_text") {
      return compare_complex_text(want_text, actual_text, tol_abs, tol_rel);
    }

    return "text mismatch: expected \"" + want_text + "\", got \"" + actual_text + "\"";
  }
  if (kind == "error") {
    if (!actual.is_error())
      return "expected error, got " + format_value(actual);
    const JsonValue* code_v = expect.find("code");
    if (code_v == nullptr || !code_v->is_string())
      return "golden missing 'code'";
    ErrorCode want_code;
    if (!display_name_to_code(code_v->as_string(), &want_code)) {
      return "golden has unknown error 'code': " + code_v->as_string();
    }
    if (actual.as_error() == want_code)
      return {};
    return std::string("error mismatch: expected ") + code_v->as_string() + ", got " + display_name(actual.as_error());
  }
  return "unknown expect kind: " + kind;
}

// Verifies a spill the oracle could not record cell by cell. The golden
// carries the dynamic-array shape plus the cells the case asked to sample,
// keyed by their absolute A1 address on the sheet Excel spilled into. Both
// sides anchor the spill at the formula cell, so a sample resolves to the
// array element at its offset from that cell.
//
// The shape is the load-bearing half: an implementation that trims a spill
// to its populated extent, or collapses it to a scalar, fails here before
// any sample is read.
inline std::string compare_array_shape(const JsonValue& expect, const Value& actual, std::uint32_t case_row,
                                       std::uint32_t case_col, double tol_abs, double tol_rel,
                                       std::string_view compare_mode) {
  if (!actual.is_array())
    return "expected array, got " + format_value(actual);
  const JsonValue* shape_v = expect.find("shape");
  if (shape_v == nullptr || !shape_v->is_array())
    return "array_shape golden missing 'shape'";
  const auto& shape = shape_v->as_array();
  if (shape.size() != 2 || !shape[0].is_number() || !shape[1].is_number())
    return "array_shape golden has invalid 'shape'";
  const auto want_rows = static_cast<std::uint32_t>(shape[0].as_number());
  const auto want_cols = static_cast<std::uint32_t>(shape[1].as_number());
  if (actual.as_array_rows() != want_rows || actual.as_array_cols() != want_cols) {
    return "array shape mismatch: expected " + std::to_string(want_rows) + "x" + std::to_string(want_cols) + ", got " +
           std::to_string(actual.as_array_rows()) + "x" + std::to_string(actual.as_array_cols());
  }

  const JsonValue* samples_v = expect.find("samples");
  if (samples_v == nullptr || !samples_v->is_object())
    return "array_shape golden missing 'samples'";
  const auto& samples = samples_v->as_object();
  if (samples.empty())
    return "array_shape golden has an empty 'samples'";

  const Value* cells = actual.as_array_cells();
  for (const auto& entry : samples) {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    if (!a1_to_row_col(entry.first, &row, &col))
      return "array_shape sample '" + entry.first + "' is not an A1 address";
    if (row < case_row || col < case_col) {
      return "array_shape sample '" + entry.first + "' is above or left of the formula cell";
    }
    const std::uint32_t offset_row = row - case_row;
    const std::uint32_t offset_col = col - case_col;
    if (offset_row >= want_rows || offset_col >= want_cols) {
      return "array_shape sample '" + entry.first + "' is outside the " + std::to_string(want_rows) + "x" +
             std::to_string(want_cols) + " spill";
    }
    const std::size_t index = static_cast<std::size_t>(offset_row) * static_cast<std::size_t>(want_cols) +
                              static_cast<std::size_t>(offset_col);
    std::string diff = compare_json_scalar(entry.second, cells[index], tol_abs, tol_rel, compare_mode);
    if (!diff.empty()) {
      return "array_shape sample " + entry.first + ": " + diff;
    }
  }
  return {};
}

// ---------------------------------------------------------------------------
// Parameter provider
// ---------------------------------------------------------------------------

inline const std::vector<OracleCase>& oracle_cases() {
  // Loaded once at first call; `load_oracle_cases` is safe to call multiple
  // times but we cache to keep test discovery deterministic even if the
  // directory is mutated mid-run (it isn't, but it's cheap insurance).
  //
  // Variant goldens are appended after the primary set. Cases inherit the
  // variant tag from the load call; the parameter-name printer uses it to
  // suffix `__<tag>` so primary and variant entries never collide. When
  // no variants are configured (the default build) the appended sequence
  // is empty and the parameter list is bit-for-bit identical to the
  // primary-only flow.
  static const std::vector<OracleCase> cached = []() {
    std::vector<OracleCase> all = load_oracle_cases(configured_golden_dir(), "");
    for (const auto& [tag, dir] : configured_variant_dirs()) {
      auto vc = load_oracle_cases(dir, tag);
      all.insert(all.end(), std::make_move_iterator(vc.begin()), std::make_move_iterator(vc.end()));
    }
    return all;
  }();
  return cached;
}

}  // namespace formulon::tests::oracle::support
