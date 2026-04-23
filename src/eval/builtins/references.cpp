// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `ADDRESS`. The function is a pure scalar text
// builder: given row/column coordinates plus optional absoluteness
// modifier, style flag (A1 vs R1C1), and sheet qualifier, it emits the
// canonical Excel spelling without touching any cell. No `EvalContext`
// is consulted, so ADDRESS rides the eager dispatch path in the registry.

#include "eval/builtins/references.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "eval/reference_lazy.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Helper: coerce an arg to a finite integer. Empty / non-finite / NaN
// surfaces the matching Excel error. Caller checks the returned int against
// its domain-specific bounds.
Expected<int, ErrorCode> read_int(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return static_cast<int>(std::trunc(d));
}

// True iff `ch` is an ASCII letter or digit.
bool is_alnum_ascii(char ch) noexcept {
  return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9');
}

// Returns true iff the sheet name needs single-quote wrapping per Excel
// convention: any byte outside `[A-Za-z0-9_]` (including leading digits
// and high-bit bytes) forces quoting. Empty names are rejected by the
// caller before we get here.
bool needs_sheet_quotes(std::string_view name) noexcept {
  if (name.empty()) {
    return false;
  }
  // Leading digit forces quoting (Excel treats e.g. `2020` as a year-like
  // identifier that must be quoted to be a sheet name).
  if (name[0] >= '0' && name[0] <= '9') {
    return true;
  }
  for (char ch : name) {
    if (!is_alnum_ascii(ch) && ch != '_') {
      return true;
    }
  }
  return false;
}

// Appends `name` to `out`, wrapping in single quotes and escaping
// embedded apostrophes iff required. The result always carries the
// trailing `!` separator.
void append_quoted_sheet(std::string* out, std::string_view name) {
  if (!needs_sheet_quotes(name)) {
    out->append(name);
    out->push_back('!');
    return;
  }
  out->push_back('\'');
  for (char ch : name) {
    if (ch == '\'') {
      // Excel escapes an embedded apostrophe by doubling it.
      out->append("''");
    } else {
      out->push_back(ch);
    }
  }
  out->push_back('\'');
  out->push_back('!');
}

// ADDRESS(row, col, [abs_num], [a1], [sheet_text])
//
// `abs_num` encodes which endpoints are absolute:
//   1 (default) -> $A$1
//   2           -> A$1  (column relative, row absolute)
//   3           -> $A1  (column absolute, row relative)
//   4           -> A1   (both relative)
// Values outside {1,2,3,4} are `#VALUE!`.
//
// `a1` decides the style: truthy / omitted -> A1; falsy -> R1C1.
// `sheet_text` is prepended as `Sheet!` (or `'Sheet'!` when it contains
// non-identifier bytes) when present and non-empty.
Value Address(const Value* args, std::uint32_t arity, Arena& arena) {
  // Row and col parameters are required.
  auto row_exp = read_int(args[0]);
  if (!row_exp) {
    return Value::error(row_exp.error());
  }
  auto col_exp = read_int(args[1]);
  if (!col_exp) {
    return Value::error(col_exp.error());
  }
  const int row = row_exp.value();
  const int col = col_exp.value();
  if (row < 1 || static_cast<std::uint32_t>(row) > Sheet::kMaxRows) {
    return Value::error(ErrorCode::Value);
  }
  if (col < 1 || static_cast<std::uint32_t>(col) > Sheet::kMaxCols) {
    return Value::error(ErrorCode::Value);
  }

  // abs_num defaults to 1 ($A$1).
  int abs_num = 1;
  if (arity >= 3) {
    auto abs_exp = read_int(args[2]);
    if (!abs_exp) {
      return Value::error(abs_exp.error());
    }
    abs_num = abs_exp.value();
    if (abs_num < 1 || abs_num > 4) {
      return Value::error(ErrorCode::Value);
    }
  }

  // a1 defaults to TRUE (A1 style). Accept any truthy value through the
  // standard bool coercion; an unparseable text is already `#VALUE!`.
  bool a1_style = true;
  if (arity >= 4) {
    auto bool_exp = coerce_to_bool(args[3]);
    if (!bool_exp) {
      return Value::error(bool_exp.error());
    }
    a1_style = bool_exp.value();
  }

  // sheet_text is optional. When the argument is supplied at all (even
  // as empty text), Excel emits the trailing `!` separator — so
  // `=ADDRESS(1,1,4,TRUE,"")` returns `"!A1"`, not `"A1"`. `has_sheet`
  // tracks whether we need that separator independently of whether the
  // name itself is empty. Error Values already short-circuited in the
  // dispatcher.
  std::string sheet_text;
  bool has_sheet = false;
  if (arity >= 5) {
    auto text_exp = coerce_to_text(args[4]);
    if (!text_exp) {
      return Value::error(text_exp.error());
    }
    sheet_text = std::move(text_exp.value());
    has_sheet = true;
  }

  // Absolute-flag bitmap decoded from abs_num. `col_abs` governs the
  // column endpoint, `row_abs` the row endpoint.
  const bool col_abs = (abs_num == 1 || abs_num == 3);
  const bool row_abs = (abs_num == 1 || abs_num == 2);

  std::string out;
  // Worst case: 5-char column letters + 7-digit row + `$` markers + sheet.
  out.reserve(sheet_text.size() + 16);
  if (has_sheet) {
    // Emit the sheet prefix whenever the caller supplied the argument;
    // an empty name still produces `"!A1"` (matches Excel 365). Bare
    // names skip quoting; names with spaces / punctuation / leading
    // digits are quoted with `''` apostrophe escaping.
    append_quoted_sheet(&out, sheet_text);
  }

  if (a1_style) {
    // A1 style: `$`-prefix each absolute endpoint, then letters + digits.
    if (col_abs) {
      out.push_back('$');
    }
    char letters[4] = {0, 0, 0, 0};
    const std::size_t n_letters = refs_internal::column_letters(static_cast<std::uint32_t>(col), letters);
    out.append(letters, n_letters);
    if (row_abs) {
      out.push_back('$');
    }
    out.append(std::to_string(row));
  } else {
    // R1C1 style. Absolute axes are written with the absolute index;
    // relative axes are written with `[offset]` brackets. ADDRESS has no
    // context to compute a real offset, so a relative axis emits `R[0]`
    // / `C[0]` when the numeric value is read as an absolute coordinate
    // and the caller still asked for the relative form — this matches
    // Excel's "R[0]C[0] relative to the current cell" semantics.
    out.push_back('R');
    if (row_abs) {
      out.append(std::to_string(row));
    } else {
      out.push_back('[');
      out.append(std::to_string(row));
      out.push_back(']');
    }
    out.push_back('C');
    if (col_abs) {
      out.append(std::to_string(col));
    } else {
      out.push_back('[');
      out.append(std::to_string(col));
      out.push_back(']');
    }
  }

  return Value::text(arena.intern(out));
}

}  // namespace

void register_reference_builtins(FunctionRegistry& registry) {
  // ADDRESS is the only reference builtin that fits the eager dispatch
  // path; INDIRECT and OFFSET are routed through the lazy table in
  // `tree_walker.cpp`.
  registry.register_function(FunctionDef{"ADDRESS", 2u, 5u, &Address});
}

}  // namespace eval
}  // namespace formulon
