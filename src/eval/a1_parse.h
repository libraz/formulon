//
// Internal A1-reference parser used by the formula-side evaluator
// (`INDIRECT`, `CELL`, `INFO`, ...).
//
// Distinct from `src/io/a1_ref.h`, which decodes the OOXML cell-address
// flavour (no sheet qualifier, no range, no `$`). The eval-side variant
// here understands sheet-qualified references with single-quote escaping
// (`'O''Brien'!A1`), `$`-anchored absolute markers (silently dropped --
// INDIRECT does not preserve abs/rel distinction), `:`-separated ranges,
// and full-column / full-row shapes (`D:D`, `5:5`).
//
// The two variants are not unified because the OOXML helper rejects every
// form except the bare `XX99` one and is exercised on the hot xlsx-load
// path; collapsing them would either pessimise that path or pull more
// surface area into the io layer than warranted.
//
// Symbols live under `formulon::eval::refs_internal::` for back-compat
// with existing tests that exercise the parser directly. Treat the
// `refs_internal` namespace as the package-private implementation
// surface: the public lazy impls (INDIRECT / OFFSET / ...) consume these
// types through their own headers.

#ifndef FORMULON_EVAL_A1_PARSE_H_
#define FORMULON_EVAL_A1_PARSE_H_

#include <cstddef>
#include <cstdint>
#include <string_view>

namespace formulon {
namespace eval {
namespace refs_internal {

/// Output of `parse_a1_ref`: sheet qualifier (empty if unqualified),
/// 0-based row/col, and a `valid` flag. When `valid` is false the other
/// fields are meaningless. `is_range` is true when the source text
/// contained a `:` separator; the second endpoint populates
/// `row2` / `col2`.
///
/// Full-column (`D:D`, `$FF:FG`) and full-row (`5:5`, `$12:$23`) shapes
/// set `is_full_col` / `is_full_row` respectively; in those cases
/// `is_range` is also true and `row`/`row2`/`col`/`col2` are populated
/// with the resulting rectangle (full-column: rows span
/// `0..Sheet::kMaxRows-1`; full-row: cols span `0..Sheet::kMaxCols-1`).
struct A1Parse {
  bool valid = false;
  bool is_range = false;
  bool is_full_col = false;
  bool is_full_row = false;
  std::string_view sheet;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t row2 = 0;
  std::uint32_t col2 = 0;
};

/// Parses `text` as an A1-style reference (with optional sheet qualifier
/// and optional `:` range). Dollar signs in `$A$1` are accepted and
/// ignored for evaluation purposes (INDIRECT does not preserve
/// absolute/relative distinction because the return path does not need
/// them). Supports single-quoted sheet names (`'Sheet 1'!A1`) with
/// doubled-quote escaping (`'O''Brien'!A1`). Also recognises full-column
/// (`D:D`, `$FF:FG`) and full-row (`5:5`, `$12:$23`) shapes, setting
/// `is_full_col` / `is_full_row` and expanding to the implied rectangle.
/// Returns an `A1Parse` with `valid = false` for any malformed input.
A1Parse parse_a1_ref(std::string_view text);

/// Writes the uppercase A1 column letters for the 1-based column `col`
/// to `out`. Returns the number of letters written (1..3); `out` must
/// have room for at least 3 chars. `col` must satisfy
/// `1 <= col <= 16384`; callers are expected to range-check first.
std::size_t column_letters(std::uint32_t col, char* out);

}  // namespace refs_internal
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_A1_PARSE_H_
