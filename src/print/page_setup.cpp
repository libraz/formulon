// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

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

  PrintableArea area;
  area.width_pt = std::max(0.0, page_width - horizontal_margin);
  area.height_pt = std::max(0.0, page_height - vertical_margin);
  return area;
}

}  // namespace print
}  // namespace formulon
