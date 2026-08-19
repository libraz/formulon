//
// Paired numeric-sample collection shared by the regression family
// (CORREL / COVARIANCE / SLOPE / INTERCEPT / RSQ / PEARSON / STEYX and the
// LINEST-style drivers) and the hypothesis family (T.TEST's paired mode,
// PROB).
//
// The two families agree on everything the pairing rule does and disagree on
// exactly one point — whether a row vector may pair with a column vector — so
// that is the only knob.

#ifndef FORMULON_EVAL_NUMERIC_PAIRS_H_
#define FORMULON_EVAL_NUMERIC_PAIRS_H_

#include <variant>
#include <vector>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon::eval {

/// Numeric samples distilled from two resolved array arguments, named by
/// argument position rather than by role: the callers disagree on which
/// argument carries the independent variable, so `first`/`second` is the only
/// naming that reads correctly at every call site. Both vectors always have
/// the same length.
struct NumericPairs {
  std::vector<double> first;   ///< Numeric cells of the first argument.
  std::vector<double> second;  ///< Numeric cells of the second argument.
};

/// Resolves both array arguments, enforces shape agreement, propagates errors
/// in row-major scan order (first argument first, so the leftmost error wins
/// as Excel requires), and collects every pair whose *both* cells are numeric.
///
/// A pair with a non-numeric cell on either side is dropped whole: Blank, Bool
/// and Text are not numeric samples, and dropping one side alone would
/// misalign the two sequences.
///
/// `allow_transpose` selects the shape rule. When true a row vector pairs with
/// a column vector as long as the total cell counts agree — the arrays are
/// 1-D sequences read in row-major order and the transpose is implicit, which
/// is what the regression family accepts. When false the two arrays must agree
/// on both dimensions. A shape mismatch is `#N/A` either way.
///
/// Returns the error `Value` to propagate on the left of the variant, or the
/// collected pairs on the right.
std::variant<Value, NumericPairs> collect_numeric_pairs(const parser::AstNode& first_arg,
                                                        const parser::AstNode& second_arg, Arena& arena,
                                                        const FunctionRegistry& registry, const EvalContext& ctx,
                                                        bool allow_transpose);

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_NUMERIC_PAIRS_H_
