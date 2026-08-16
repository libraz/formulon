
#include "print/page_setup.h"

#include <algorithm>

namespace formulon {
namespace print {
namespace {

// --- OOXML paperSize codes (ECMA-376 §18.3.1.63, `ST_PaperSize`). ---
constexpr std::uint32_t kPaperLetter = 1U;
constexpr std::uint32_t kPaperLegal = 5U;
constexpr std::uint32_t kPaperA3 = 8U;
constexpr std::uint32_t kPaperA4 = 9U;
constexpr std::uint32_t kPaperA5 = 11U;
constexpr std::uint32_t kPaperA6 = 70U;

// --- Imperial paper sizes, in inches (portrait: width x height). ---
constexpr double kLetterWidthInches = 8.5;
constexpr double kLetterHeightInches = 11.0;
constexpr double kLegalWidthInches = 8.5;
constexpr double kLegalHeightInches = 14.0;

// --- ISO A-series paper sizes, in millimetres (portrait: width x height). ---
constexpr double kA3WidthMm = 297.0;
constexpr double kA3HeightMm = 420.0;
constexpr double kA4WidthMm = 210.0;
constexpr double kA4HeightMm = 297.0;
constexpr double kA5WidthMm = 148.0;
constexpr double kA5HeightMm = 210.0;
constexpr double kA6WidthMm = 105.0;
constexpr double kA6HeightMm = 148.0;

constexpr double InchesToPoints(double inches) noexcept {
  return inches * kPointsPerInch;
}

constexpr double MmToPoints(double mm) noexcept {
  return mm * kPointsPerInch / kMillimetresPerInch;
}

}  // namespace

PaperDimensions resolve_paper_dimensions(std::uint32_t paper_size) noexcept {
  switch (paper_size) {
    case kPaperLetter:
      return PaperDimensions{InchesToPoints(kLetterWidthInches), InchesToPoints(kLetterHeightInches)};
    case kPaperLegal:
      return PaperDimensions{InchesToPoints(kLegalWidthInches), InchesToPoints(kLegalHeightInches)};
    case kPaperA3:
      return PaperDimensions{MmToPoints(kA3WidthMm), MmToPoints(kA3HeightMm)};
    case kPaperA4:
      return PaperDimensions{MmToPoints(kA4WidthMm), MmToPoints(kA4HeightMm)};
    case kPaperA5:
      return PaperDimensions{MmToPoints(kA5WidthMm), MmToPoints(kA5HeightMm)};
    case kPaperA6:
      return PaperDimensions{MmToPoints(kA6WidthMm), MmToPoints(kA6HeightMm)};
    default:
      // Unknown printer paper: fall back to A4, the locale default.
      return PaperDimensions{MmToPoints(kA4WidthMm), MmToPoints(kA4HeightMm)};
  }
}

PrintableArea compute_printable_area(const PageSetup& setup, const PageMargins& margins) noexcept {
  const PaperDimensions paper = resolve_paper_dimensions(setup.paper_size);

  double page_width = paper.width_pt;
  double page_height = paper.height_pt;
  if (setup.orientation == Orientation::kLandscape) {
    std::swap(page_width, page_height);
  }

  const double horizontal_margin = (margins.left + margins.right) * kPointsPerInch;
  const double vertical_margin = (margins.top + margins.bottom) * kPointsPerInch;
  // Excel's `<pageMargins>` stores header / footer as the distance from
  // the page edge to that band, and the top / bottom margins as the
  // distance to the cell body. The bands therefore sit *inside* the
  // margins and steal no body height at the standard preset (header 0.3"
  // inside a 0.75" top margin). Only when a band plus its one-line text
  // strip reaches past its margin does the body give way, which is what
  // setting a header margin larger than the top margin does.
  //
  // Reserving both bands unconditionally cost ~71 pt of body height and
  // broke pages Excel keeps whole: the workbook goldens' measured row
  // geometry bounds the total reserve at 36.18 pt (print_pagination.
  // tall_table_horizontal_breaks) and no captured case requires any.
  //
  // The strip is sized for the default 11 pt header/footer font at 1.2x
  // line height, ~13.2 pt; 14 pt keeps the body in lockstep with Excel's
  // `Pages.Count`.
  constexpr double kHeaderFooterTextBandPt = 14.0;
  const double header_band = margins.header * kPointsPerInch;
  const double footer_band = margins.footer * kPointsPerInch;
  const double top_margin = margins.top * kPointsPerInch;
  const double bottom_margin = margins.bottom * kPointsPerInch;
  const double header_encroach =
      header_band > 0.0 ? std::max(0.0, header_band + kHeaderFooterTextBandPt - top_margin) : 0.0;
  const double footer_encroach =
      footer_band > 0.0 ? std::max(0.0, footer_band + kHeaderFooterTextBandPt - bottom_margin) : 0.0;

  PrintableArea area;
  area.width_pt = std::max(0.0, page_width - horizontal_margin);
  area.height_pt = std::max(0.0, page_height - vertical_margin - header_encroach - footer_encroach);
  return area;
}

}  // namespace print
}  // namespace formulon
