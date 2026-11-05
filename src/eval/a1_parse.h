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

/// The cell an R1C1 relative axis is measured from.
///
/// `present` is false when the formula has no cell of its own — the
/// ad-hoc "evaluate this text" entry points have none — which leaves a
/// relative axis with nothing to resolve against. Absolute axes do not
/// consult the base and keep working either way.
struct R1C1Base {
  bool present = false;
  std::uint32_t row = 0;  ///< 0-based.
  std::uint32_t col = 0;  ///< 0-based.
};

/// The characters an R1C1 endpoint is spelled with: the row and column
/// letters (matched ASCII case-insensitively, given in upper case) and the
/// relative brackets. The defaults are the invariant spelling.
struct R1C1Letters {
  char row = 'R';
  char col = 'C';
  char open = '[';
  char close = ']';
};

/// One axis of an R1C1 endpoint as written.
struct R1C1EndpointAxis {
  bool present = false;
  bool relative = false;  ///< `R[-2]` or a bare `R`; false for `R5`.
  long long value = 0;    ///< 1-based index when absolute (saturated), signed offset when relative.
};

/// What an R1C1 endpoint reads as.
enum class R1C1EndpointKind : std::uint8_t {
  /// Every absolute axis inside the grid and every bracketed offset at most
  /// `Sheet::kMaxRows` in magnitude.
  kReference,
  /// No bracket, and an absolute axis outside the grid (`R0`, `C16385`,
  /// `R4294967296`): the text is a name, not a reference.
  kName,
  /// No axis, a bracket that does not close on a signed digit run, a
  /// bracketed offset past `Sheet::kMaxRows`, or a bracketed endpoint with an
  /// absolute axis outside the grid.
  kMalformed,
};

/// One R1C1 endpoint read by `parse_r1c1_endpoint_axes`.
struct R1C1Endpoint {
  R1C1EndpointKind kind = R1C1EndpointKind::kMalformed;
  R1C1EndpointAxis row;
  R1C1EndpointAxis col;
  bool bracketed = false;  ///< An axis is written with brackets, read or not.
  std::size_t end = 0;     ///< Offset just past the bytes read.
};

/// Reads one endpoint (`R1C1`, `RC[-1]`, `R[2]`, `C`) at `text[start]`: an
/// optional row axis, then an optional column axis, each the letter alone
/// (offset 0), a digit run (absolute, any length) or a bracketed signed
/// offset. Reading stops at the first byte that continues neither axis;
/// the caller judges what follows `end`.
R1C1Endpoint parse_r1c1_endpoint_axes(std::string_view text, std::size_t start, const R1C1Letters& letters);

/// Resolves one axis to a 0-based coordinate against `base`. An absolute
/// axis outside `[1, max]` fails; a relative axis wraps around the sheet
/// edge as Excel does, and fails when there is no base to measure from.
bool resolve_r1c1_axis(const R1C1EndpointAxis& axis, bool base_present, std::uint32_t base, std::uint32_t max,
                       std::uint32_t* out);

/// Parses `text` as an R1C1-style reference. Accepts the same optional
/// sheet qualifier as `parse_a1_ref` and the same `:` range separator,
/// and reports its result in the same `A1Parse` shape so both styles
/// reach one consumer.
///
/// The row/column letters and the relative brackets follow the current
/// eval profile (`R`/`C` and `[ ]` for ja/en, `Z`/`S` and `( )` for de, ...);
/// another locale's spelling is malformed.
///
/// An axis is written either absolutely (`R5`, 1-based) or relative to
/// `base` (`R[-2]`, or a bare `R` meaning offset 0). An endpoint that
/// names only one axis is unbounded along the other — `R5` is the whole
/// of row 5 and `C2` the whole of column B — which sets `is_full_row` /
/// `is_full_col` exactly as `5:5` and `B:B` do on the A1 side.
///
/// Returns `valid = false` for malformed input, for an axis that resolves
/// outside the grid, for a relative axis with no base to measure from,
/// and for a range whose endpoints name different axes (`R2:R3C4`),
/// which describes no rectangle.
A1Parse parse_r1c1_ref(std::string_view text, const R1C1Base& base);

}  // namespace refs_internal
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_A1_PARSE_H_
