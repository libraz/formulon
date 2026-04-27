// Copyright 2026 libraz. Licensed under the MIT License.
//
// TEXTSPLIT — splits a text scalar into a 1D or 2D array of substrings.
//
// `TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty]=FALSE,
//            [match_mode]=0, [pad_with]=#N/A)`
//
// `col_delimiter` and `row_delimiter` may each be either a scalar text or
// a 1D array of texts. When an array is supplied, any of its non-empty
// entries is treated as a separator at that level. When both are
// scalars / arrays of texts, the input text is first split into rows by
// `row_delimiter`, then each row is split into columns by
// `col_delimiter`. Rows that yield fewer columns than the widest row are
// padded with `pad_with` (default `#N/A`) to keep the output rectangular.
//
// Output shapes:
//   * only `col_delimiter` -> 1 x N row
//   * only `row_delimiter` -> M x 1 column
//   * both -> M x N matrix
//
// Match mode is ASCII case folding (Excel's documented behaviour for the
// modern text family — Unicode case folding is intentionally deferred
// until the locale layer lands; matches the existing TEXTBEFORE /
// TEXTAFTER policy in `eval/builtins/text_modern.cpp`).
//
// Empty-delimiter entries inside a delimiter array are filtered out
// (matches Mac Excel: an `""` entry has no observable effect on the
// split). When the *effective* delimiter list at a level is empty (e.g.
// the user passed `""` as the only column delimiter), no splitting
// happens at that level — the full text is kept as a single token.

#ifndef FORMULON_EVAL_TEXTSPLIT_LAZY_H_
#define FORMULON_EVAL_TEXTSPLIT_LAZY_H_

#include "utils/arena.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty]=FALSE,
///            [match_mode]=0, [pad_with]=#N/A)` — splits `text` into a
/// rectangular array using `col_delimiter` as the column separator and
/// `row_delimiter` (when supplied) as the row separator.
///
/// Arguments:
///   * `text` — scalar text. Number / Bool coerce; Error propagates.
///   * `col_delimiter` — text or 1D array of texts. Empty entries inside
///     a delimiter array are dropped. If the resulting list is empty,
///     no column-level splitting happens.
///   * `row_delimiter` — text or 1D array of texts (optional). Same
///     empty-entry filter. When omitted, the output is one row.
///   * `ignore_empty` — bool (optional, default FALSE). When TRUE, empty
///     tokens are dropped from each axis before padding.
///   * `match_mode` — `0` case-sensitive (default) or `1` case-insensitive
///     ASCII (matches TEXTBEFORE / TEXTAFTER). Anything else -> `#VALUE!`.
///   * `pad_with` — any value (optional, default `#N/A`). Used to fill
///     short rows when the column counts differ.
///
/// Errors:
///   * propagation: any error in any argument surfaces as that error;
///   * `match_mode` out of `{0, 1}` -> `#VALUE!`;
///   * `col_delimiter` or `row_delimiter` shaped as 2D -> `#VALUE!`;
///   * `col_delimiter` argument missing entirely (Excel's required arg)
///     is rejected by the dispatcher's arity check (min = 2).
Value eval_textsplit_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXTSPLIT_LAZY_H_
