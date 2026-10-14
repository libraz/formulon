
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
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
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
// no-op). Match precedence at a tied position is longest-needle-first. The
// caller folds the needles once when match_mode=1; this helper only folds
// each haystack when requested, and always slices the original text so the
// output keeps its original casing and UTF-8 bytes.
std::vector<std::string> split_by_any(std::string_view text, const std::vector<std::string>& needles,
                                      bool fold_haystack) {
  std::vector<std::string> tokens;
  if (needles.empty()) {
    tokens.emplace_back(text);
    return tokens;
  }

  // The common scalar-delimiter path can use the string library's find loop;
  // unlike the general path it does not inspect every byte against every
  // delimiter. ASCII folding is byte-preserving, so offsets in `hay` are
  // also valid offsets in the original `text`.
  if (needles.size() == 1U) {
    const std::string& needle = needles.front();
    if (needle.empty()) {
      tokens.emplace_back(text);
      return tokens;
    }
    std::string folded;
    std::string_view hay = text;
    if (fold_haystack) {
      folded = strings::to_ascii_lower_owned(std::string(text));
      hay = folded;
    }
    std::size_t token_start = 0;
    std::size_t search_start = 0;
    while (true) {
      const std::size_t match = hay.find(needle, search_start);
      if (match == std::string_view::npos) {
        break;
      }
      tokens.emplace_back(text.substr(token_start, match - token_start));
      token_start = match + needle.size();
      search_start = token_start;
    }
    tokens.emplace_back(text.substr(token_start));
    return tokens;
  }

  std::string lowered;
  std::string_view hay = text;
  if (fold_haystack) {
    lowered = strings::to_ascii_lower_owned(std::string(text));
    hay = lowered;
  }

  std::size_t token_start = 0;
  std::size_t i = 0;
  while (i < text.size()) {
    std::size_t best_match_len = 0;
    for (const auto& needle : needles) {
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
      tokens.emplace_back(text.substr(token_start, i - token_start));
      i += best_match_len;
      token_start = i;
      continue;
    }
    ++i;
  }
  tokens.emplace_back(text.substr(token_start));
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
  // Saturated before the caller's whitelist check, so a huge argument narrows defined.
  return truncate_saturated_int(coerced.value());
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

  // Fold effective delimiters once after every argument has been evaluated.
  // The folded strings keep their existing allocations; split_by_any can then
  // reuse them for every row instead of rebuilding a folded needle list per
  // invocation.
  if (case_insensitive) {
    for (auto& delimiter : col_delims) {
      delimiter = strings::to_ascii_lower_owned(std::move(delimiter));
    }
    for (auto& delimiter : row_delims) {
      delimiter = strings::to_ascii_lower_owned(std::move(delimiter));
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
      auto write = cols.begin();
      for (auto read = cols.begin(); read != cols.end(); ++read) {
        if (!read->empty()) {
          if (write != read) {
            *write = std::move(*read);
          }
          ++write;
        }
      }
      cols.erase(write, cols.end());
    }
    grid.push_back(std::move(cols));
  }

  if (ignore_empty) {
    // Drop rows whose every column is empty (ie. nothing left after the
    // per-row filter).
    auto write = grid.begin();
    for (auto read = grid.begin(); read != grid.end(); ++read) {
      if (!read->empty()) {
        if (write != read) {
          *write = std::move(*read);
        }
        ++write;
      }
    }
    grid.erase(write, grid.end());
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
