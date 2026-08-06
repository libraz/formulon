//
// Paper geometry for the print/pagination engine.
//
// Maps OOXML `<pageSetup paperSize>` codes to physical sheet dimensions in
// PostScript points (1 inch = 72 points) and derives the printable body
// area after subtracting page margins. The body area is what the
// pagination engine fills with cells when computing page breaks.

#ifndef FORMULON_PRINT_PAGE_SETUP_H_
#define FORMULON_PRINT_PAGE_SETUP_H_

#include "sheet.h"

namespace formulon {
namespace print {

/// PostScript points per inch (the unit physical paper dimensions and
/// margins are reduced to before any layout arithmetic).
inline constexpr double kPointsPerInch = 72.0;

/// Millimetres per inch, used to convert ISO (A-series) paper sizes.
inline constexpr double kMillimetresPerInch = 25.4;

/// The printable region of a single physical page, in points.
///
/// `width_pt` / `height_pt` are already orientation-adjusted (landscape
/// swaps the paper's long and short edges) and margin-reduced. They never
/// go negative: oversized margins clamp the body area to zero.
struct PrintableArea {
  double width_pt = 0.0;   ///< Body width available for cell columns, in points.
  double height_pt = 0.0;  ///< Body height available for cell rows, in points.
};

/// The physical dimensions of one paper size, in points, in portrait
/// orientation (short edge as width, long edge as height).
struct PaperDimensions {
  double width_pt = 0.0;   ///< Portrait width (short edge), in points.
  double height_pt = 0.0;  ///< Portrait height (long edge), in points.
};

/// Resolves an OOXML `paperSize` code to physical portrait dimensions.
///
/// Recognised codes: 1 (Letter), 5 (Legal), 8 (A3), 9 (A4), 11 (A5),
/// 70 (A6). Any unrecognised code falls back to A4, matching Excel's
/// behaviour of treating an unknown printer paper as the locale default.
PaperDimensions resolve_paper_dimensions(std::uint32_t paper_size) noexcept;

/// Computes the printable body area for a page.
///
/// Takes the paper dimensions for `setup.paper_size`, swaps width and
/// height when `setup.orientation == Orientation::kLandscape`, then
/// subtracts the left+right margins from the width and the top+bottom
/// margins plus the header/footer band heights from the height (margins
/// are inches; converted via `kPointsPerInch`).
///
/// Excel paginates the cell body strictly between the header and footer
/// bands: a row only fits on a page when it sits below `top + header`
/// and above `bottom + footer`, so the body height shrinks by both
/// side-margin pairs.
PrintableArea compute_printable_area(const PageSetup& setup, const PageMargins& margins) noexcept;

}  // namespace print
}  // namespace formulon

#endif  // FORMULON_PRINT_PAGE_SETUP_H_
