//
// Intra-subdirectory header shared by `reference/indirect.cpp`,
// `reference/offset.cpp`, and `reference/intersection.cpp`. It is NOT part
// of the public eval surface; sibling TUs include it directly to share
// the small carrier structs (`OffsetBase`, `IndirectReference`) and the
// rectangle-construction helpers (`resolve_indirect_reference`,
// `resolve_offset_base`, `compute_offset_rect`, `apply_offset`,
// `read_int`) that the three reference-family pipelines all touch.
//
// The public A1-text parser (`parse_a1_ref`, `column_letters`, `A1Parse`)
// continues to live in `eval/a1_parse.h`; the bodies are defined in
// `reference/common.cpp` and are reachable from outside the subdirectory
// through that header. Public lazy impl declarations
// (`eval_indirect_lazy`, `eval_offset_lazy`) live in `reference/indirect.h`
// and `reference/offset.h`; the range-expander / range-resolver externs
// live in `eval/range_expanders.h` and `eval/range_resolvers.h`.
//
// TODO(D-07): the A1 endpoint parser duplicates work that the OOXML-side
// `src/io/a1_ref.h` already does for a more restricted dialect. Hoisting
// a shared parser into a refs layer depends on cross-layer design
// decisions (eval -> io dependency direction); defer the unification
// until that lands. Until then, the parser stays local here.

#ifndef FORMULON_EVAL_REFERENCE_COMMON_H_
#define FORMULON_EVAL_REFERENCE_COMMON_H_

#include <cmath>
#include <cstdint>
#include <string_view>

#include "eval/coerce.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

namespace refs_internal {

// Shape of the base `reference` argument to OFFSET. Populated by
// `resolve_offset_base` from either a bare Ref (1x1) or a literal
// `RangeOp` rectangle. Stored as 0-based indices.
struct OffsetBase {
  std::string_view sheet;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t rows = 1;
  std::uint32_t cols = 1;
};

// Result of decoding an INDIRECT call's textual reference into a
// rectangle. `sheet` is empty when the source text omits the qualifier;
// `range_syntax` reflects whether the parsed source contained a `:`
// (independent of whether the resulting rectangle collapses to 1x1).
struct IndirectReference {
  std::string_view sheet;
  std::uint32_t top_row = 0;
  std::uint32_t left_col = 0;
  std::uint32_t bottom_row = 0;
  std::uint32_t right_col = 0;
  bool is_range = false;
  bool range_syntax = false;
  // Whole-column (`D:D`) / whole-row (`5:5`) shapes. The rectangle fields
  // are still populated (spanning `0..kMax-1` along the unbounded axis),
  // but callers that materialise the rectangle into a `Value::Array`
  // should reject these because the spill would be unbounded.
  bool is_full_col = false;
  bool is_full_row = false;
};

// Converts a signed offset plus a non-negative base (both measured in
// row / column units) into a 0-based grid coordinate, returning `false`
// if the result falls outside [0, max). `max` is `Sheet::kMaxRows` or
// `Sheet::kMaxCols`.
bool apply_offset(std::uint32_t base, int offset, std::uint32_t max, std::uint32_t* out);

// Reads an integer arg via truncation. `#VALUE!` on coercion failure,
// `#NUM!` on NaN/Inf.
Expected<int, ErrorCode> read_int(const Value& v);

// Evaluates the INDIRECT call AST, decoding its text argument into a
// rectangle. Returns `false` on any error and writes the Excel-visible
// code into `*out_err`; on success populates `*out` and returns `true`.
bool resolve_indirect_reference(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx, IndirectReference* out, ErrorCode* out_err);

// Normalises the base reference from OFFSET's first argument. Literal
// Ref and `Ref:Ref` RangeOp shapes are accepted directly; INDIRECT /
// OFFSET nested calls go through `resolve_reference_call` so
// `OFFSET(INDIRECT("A1"), …)` and `OFFSET(OFFSET(A1,0,0,2,2), …)`
// resolve without dereferencing the base.
bool resolve_offset_base(const parser::AstNode& arg, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, OffsetBase* out, ErrorCode* out_err);

// Computes the (0-based) top-left corner and the (height, width) of the
// rectangle produced by `OFFSET(base, rows, cols, [height], [width])`.
// Returns `false` on any error, writing the Excel-visible code into
// `*out_err`.
bool compute_offset_rect(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, OffsetBase* out_base, std::uint32_t* out_top_row,
                         std::uint32_t* out_left_col, std::uint32_t* out_height, std::uint32_t* out_width,
                         ErrorCode* out_err);

}  // namespace refs_internal
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_COMMON_H_
