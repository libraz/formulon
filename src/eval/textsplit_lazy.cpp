
#include "eval/textsplit_lazy.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/shape_ops_lazy.h"
#include "eval/text_ops.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Reads a delimiter argument (col or row) into a list of non-empty
// strings. Empty entries inside an array are dropped silently (Mac Excel
// observed behaviour: `{",",""}` behaves identically to `{","}`); a
// scalar empty string yields an empty list which means "no splitting at
// this level". 2D delimiter arrays surface `#VALUE!` because the
// orientation is ambiguous. Errors propagate via *out_err.
std::vector<std::string> extract_delimiters(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                            const EvalContext& ctx, Value* out_err) {
  std::vector<std::string> result;
  const Value v = eval_node_as_array(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return result;
  }
  if (!v.is_array()) {
    // Scalar fallback: rare path because eval_node_as_array always wraps
    // into a 1x1 array. Keep it as a defensive branch in case the array
    // context seam ever returns a raw scalar for a particular AST shape.
    auto t = coerce_to_text(v);
    if (!t) {
      *out_err = Value::error(t.error());
      return result;
    }
    if (!t.value().empty()) {
      result.push_back(std::move(t.value()));
    }
    return result;
  }
  const ArrayValue* arr = v.as_array();
  // Reject 2D delimiter arrays. A degenerate 1x1 is fine (it's how a
  // scalar reaches us through `eval_node_as_array`); 1xN or Mx1 are also
  // fine and represent "any of these strings".
  if (arr->rows > 1U && arr->cols > 1U) {
    *out_err = Value::error(ErrorCode::Value);
    return result;
  }
  const std::size_t n = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = arr->cells[i];
    if (cell.is_error()) {
      *out_err = cell;
      return result;
    }
    auto t = coerce_to_text(cell);
    if (!t) {
      *out_err = Value::error(t.error());
      return result;
    }
    if (!t.value().empty()) {
      result.push_back(std::move(t.value()));
    }
  }
  return result;
}

// Splits `text` on any occurrence of any string in `needles`. Empty
// needles are assumed already filtered out by the caller. When `needles`
// is empty the function returns a single token containing all of `text`
// (matches Excel: an effective-empty delimiter list at a level is a
// no-op). Match precedence at a tied position is longest-needle-first;
// case folding is ASCII-only via `to_lower_ascii` (matches the
// TEXTBEFORE / TEXTAFTER MVP policy).
std::vector<std::string> split_by_any(std::string_view text, const std::vector<std::string>& needles,
                                      bool case_insensitive) {
  std::vector<std::string> tokens;
  if (needles.empty()) {
    tokens.emplace_back(text);
    return tokens;
  }
  std::string lowered;
  std::vector<std::string> lowered_needles;
  std::string_view hay = text;
  if (case_insensitive) {
    lowered = to_lower_ascii(text);
    hay = lowered;
    lowered_needles.reserve(needles.size());
    for (const auto& n : needles) {
      lowered_needles.push_back(to_lower_ascii(n));
    }
  }
  const auto& effective_needles = case_insensitive ? lowered_needles : needles;

  std::string current;
  std::size_t i = 0;
  while (i < text.size()) {
    std::size_t best_match_len = 0;
    for (const auto& needle : effective_needles) {
      if (needle.empty()) {
        continue;
      }
      if (i + needle.size() > hay.size()) {
        continue;
      }
      if (hay.compare(i, needle.size(), needle) == 0) {
        if (needle.size() > best_match_len) {
          best_match_len = needle.size();
        }
      }
    }
    if (best_match_len > 0) {
      tokens.push_back(std::move(current));
      current.clear();
      i += best_match_len;
      continue;
    }
    current.push_back(text[i]);
    ++i;
  }
  tokens.push_back(std::move(current));
  return tokens;
}

// Reads an int-valued optional arg with a default. Validates against the
// allowed value set. Errors propagate via *out_err and the function
// returns the default value (caller must check *out_err first).
int read_int_opt(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                 int default_value, Value* out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return default_value;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return default_value;
  }
  return static_cast<int>(coerced.value());
}

bool read_bool_opt(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                   bool default_value, Value* out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return default_value;
  }
  auto coerced = coerce_to_bool(v);
  if (!coerced) {
    *out_err = Value::error(coerced.error());
    return default_value;
  }
  return coerced.value();
}

}  // namespace

Value eval_textsplit_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 6U) {
    return Value::error(ErrorCode::Value);
  }

  // text (scalar). Errors propagate; arrays surface the first cell's
  // value (matches scalar-coercion idiom).
  const Value text_v = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (text_v.is_error()) {
    return text_v;
  }
  auto text_c = coerce_to_text(text_v);
  if (!text_c) {
    return Value::error(text_c.error());
  }
  const std::string text = std::move(text_c.value());

  Value err = Value::blank();
  std::vector<std::string> col_delims = extract_delimiters(call.as_call_arg(1), arena, registry, ctx, &err);
  if (err.is_error()) {
    return err;
  }

  std::vector<std::string> row_delims;
  if (arity >= 3U) {
    row_delims = extract_delimiters(call.as_call_arg(2), arena, registry, ctx, &err);
    if (err.is_error()) {
      return err;
    }
  }

  bool ignore_empty = false;
  if (arity >= 4U) {
    ignore_empty = read_bool_opt(call.as_call_arg(3), arena, registry, ctx, false, &err);
    if (err.is_error()) {
      return err;
    }
  }

  int match_mode = 0;
  if (arity >= 5U) {
    match_mode = read_int_opt(call.as_call_arg(4), arena, registry, ctx, 0, &err);
    if (err.is_error()) {
      return err;
    }
    if (match_mode != 0 && match_mode != 1) {
      return Value::error(ErrorCode::Value);
    }
  }
  const bool case_insensitive = match_mode == 1;

  Value pad = Value::error(ErrorCode::NA);
  if (arity == 6U) {
    const Value pad_v = eval_node(call.as_call_arg(5), arena, registry, ctx);
    if (pad_v.is_error()) {
      return pad_v;
    }
    if (pad_v.is_array()) {
      const ArrayValue* pa = pad_v.as_array();
      if (pa->rows == 0U || pa->cols == 0U) {
        return Value::error(ErrorCode::Value);
      }
      pad = pa->cells[0];
    } else {
      pad = pad_v;
    }
  }

  // Split into rows first. When `row_delims` is empty (omitted or all
  // empty entries), we have a single row containing the whole text.
  const std::vector<std::string> raw_rows = split_by_any(text, row_delims, case_insensitive);

  // Then split each row into columns. Track each row's column tokens; the
  // output column count is the widest row.
  std::vector<std::vector<std::string>> grid;
  grid.reserve(raw_rows.size());
  for (const auto& row : raw_rows) {
    std::vector<std::string> cols = split_by_any(row, col_delims, case_insensitive);
    if (ignore_empty) {
      std::vector<std::string> kept;
      kept.reserve(cols.size());
      for (auto& c : cols) {
        if (!c.empty()) {
          kept.push_back(std::move(c));
        }
      }
      cols = std::move(kept);
    }
    grid.push_back(std::move(cols));
  }

  if (ignore_empty) {
    // Drop rows whose every column is empty (ie. nothing left after the
    // per-row filter).
    std::vector<std::vector<std::string>> kept_rows;
    kept_rows.reserve(grid.size());
    for (auto& row : grid) {
      if (!row.empty()) {
        kept_rows.push_back(std::move(row));
      }
    }
    grid = std::move(kept_rows);
  }

  // Mac Excel surfaces #CALC! when ignore_empty wipes the grid clean.
  if (grid.empty()) {
    return Value::error(ErrorCode::Calc);
  }

  std::size_t out_cols = 0;
  for (const auto& row : grid) {
    if (row.size() > out_cols) {
      out_cols = row.size();
    }
  }
  if (out_cols == 0U) {
    out_cols = 1U;
  }
  const std::size_t out_rows = grid.size();

  Value* buffer = nullptr;
  ArrayValue* out = allocate_array_value(static_cast<std::uint32_t>(out_rows), static_cast<std::uint32_t>(out_cols),
                                         arena, buffer, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t r = 0; r < out_rows; ++r) {
    for (std::size_t c = 0; c < out_cols; ++c) {
      const std::size_t dst = r * out_cols + c;
      if (c < grid[r].size()) {
        buffer[dst] = Value::text(arena.intern(grid[r][c]));
      } else {
        buffer[dst] = pad;
      }
    }
  }

  return Value::array(out);
}

}  // namespace eval
}  // namespace formulon
