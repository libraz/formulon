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

#include "eval/a1_parse.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/eval_profile_scope.h"
#include "eval/function_registry.h"
#include "excel_locale.h"
#include "parser/ast_format.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ADDRESS's row / column snap to a near integer. A raw value below 1 is
// `#VALUE!` before the snap, and above `max` after it.
Expected<int, ErrorCode> read_snapped_coordinate(const Value& v, std::uint32_t max) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return std::move(coerced.error());
  }
  const double d = coerced.value();
  if (d < 1.0) {
    return ErrorCode::Value;
  }
  const double coordinate = std::trunc(snap_near_integer(d));
  if (coordinate > static_cast<double>(max)) {
    return ErrorCode::Value;
  }
  return static_cast<int>(coordinate);
}

Expected<int, ErrorCode> read_optional_int(const Value* args, std::uint32_t arity, std::uint32_t index,
                                           int default_value) {
  if (arity <= index) {
    return default_value;
  }
  return coerce_to_truncated_int(args[index]);
}

Expected<bool, ErrorCode> read_optional_bool(const Value* args, std::uint32_t arity, std::uint32_t index,
                                             bool default_value) {
  if (arity <= index) {
    return default_value;
  }
  return coerce_to_bool(args[index]);
}

// Appends `name` and the `!` separator. An empty sheet_text stays bare
// (`!$A$1`); any other name is quoted by the A1 sheet-qualifier rule in both
// reference styles.
void append_quoted_sheet(std::string* out, std::string_view name) {
  if (!name.empty()) {
    parser::append_sheet_name(name, /*force_quote=*/false, *out);
  }
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
  auto row_exp = read_snapped_coordinate(args[0], Sheet::kMaxRows);
  if (!row_exp) {
    return Value::error(row_exp.error());
  }
  auto col_exp = read_snapped_coordinate(args[1], Sheet::kMaxCols);
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

  auto abs_exp = read_optional_int(args, arity, 2, 1);
  if (!abs_exp) {
    return Value::error(abs_exp.error());
  }
  const int abs_num = abs_exp.value();
  if (abs_num < 1 || abs_num > 4) {
    return Value::error(ErrorCode::Value);
  }

  auto bool_exp = read_optional_bool(args, arity, 3, true);
  if (!bool_exp) {
    return Value::error(bool_exp.error());
  }
  const bool a1_style = bool_exp.value();

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
    // an empty name still produces `"!A1"` (matches Excel 365).
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
    const LocaleFacts& facts = locale_facts(current_eval_profile());
    out.push_back(facts.r1c1_row);
    if (row_abs) {
      out.append(std::to_string(row));
    } else {
      out.push_back(facts.r1c1_open);
      out.append(std::to_string(row));
      out.push_back(facts.r1c1_close);
    }
    out.push_back(facts.r1c1_col);
    if (col_abs) {
      out.append(std::to_string(col));
    } else {
      out.push_back(facts.r1c1_open);
      out.append(std::to_string(col));
      out.push_back(facts.r1c1_close);
    }
  }

  return Value::text(arena.intern(out));
}

}  // namespace

void register_reference_builtins(FunctionRegistry& registry) {
  // ADDRESS is the only reference builtin that fits the eager dispatch
  // path; INDIRECT and OFFSET are routed through the lazy table in
  // `tree_walker.cpp`.
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"ADDRESS", 2u, 5u, &Address},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
