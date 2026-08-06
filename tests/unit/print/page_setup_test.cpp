//
// Unit tests for the print paper-geometry layer (`src/print/page_setup`).

#include "print/page_setup.h"

#include "gtest/gtest.h"
#include "sheet.h"

namespace formulon {
namespace print {
namespace {

// Tolerance for point comparisons: the mm->pt conversions are not
// integral, so an exact equality would be brittle.
constexpr double kTol = 0.05;

TEST(PageSetupTest, A4PortraitDimensions) {
  const PaperDimensions a4 = resolve_paper_dimensions(/*A4=*/9);
  EXPECT_NEAR(a4.width_pt, 595.27, kTol);
  EXPECT_NEAR(a4.height_pt, 841.89, kTol);
}

TEST(PageSetupTest, LetterDimensions) {
  const PaperDimensions letter = resolve_paper_dimensions(/*Letter=*/1);
  EXPECT_NEAR(letter.width_pt, 612.0, kTol);
  EXPECT_NEAR(letter.height_pt, 792.0, kTol);
}

TEST(PageSetupTest, LegalDimensions) {
  const PaperDimensions legal = resolve_paper_dimensions(/*Legal=*/5);
  EXPECT_NEAR(legal.width_pt, 612.0, kTol);
  EXPECT_NEAR(legal.height_pt, 1008.0, kTol);
}

TEST(PageSetupTest, UnknownPaperCodeFallsBackToA4) {
  const PaperDimensions unknown = resolve_paper_dimensions(/*nonexistent=*/9999);
  const PaperDimensions a4 = resolve_paper_dimensions(9);
  EXPECT_DOUBLE_EQ(unknown.width_pt, a4.width_pt);
  EXPECT_DOUBLE_EQ(unknown.height_pt, a4.height_pt);
}

TEST(PageSetupTest, LandscapeSwapsWidthAndHeight) {
  PageSetup setup;  // Defaults: A4, no margins set yet.
  setup.paper_size = 9;
  PageMargins margins;
  margins.left = margins.right = margins.top = margins.bottom = 0.0;
  margins.header = margins.footer = 0.0;

  setup.orientation = Orientation::kPortrait;
  const PrintableArea portrait = compute_printable_area(setup, margins);

  setup.orientation = Orientation::kLandscape;
  const PrintableArea landscape = compute_printable_area(setup, margins);

  EXPECT_NEAR(landscape.width_pt, portrait.height_pt, kTol);
  EXPECT_NEAR(landscape.height_pt, portrait.width_pt, kTol);
}

TEST(PageSetupTest, PrintableAreaSubtractsDefaultMargins) {
  PageSetup setup;
  setup.paper_size = 9;  // A4.
  setup.orientation = Orientation::kPortrait;
  PageMargins margins;  // OOXML defaults: 0.7in sides, 0.75in top/bottom,
                        // 0.3in header/footer.

  const PrintableArea area = compute_printable_area(setup, margins);
  // A4 portrait is ~595.27 x 841.89 pt; 1.4in side margin = 100.8 pt,
  // 1.5in top/bottom = 108 pt, 0.6in header+footer band = 43.2 pt,
  // plus a 28 pt header/footer text reservation (one line each).
  EXPECT_NEAR(area.width_pt, 595.27 - 100.8, kTol);
  EXPECT_NEAR(area.height_pt, 841.89 - 108.0 - 43.2 - 28.0, kTol);
}

TEST(PageSetupTest, PrintableAreaSubtractsCustomMargins) {
  PageSetup setup;
  setup.paper_size = 1;  // Letter, 612 x 792 pt.
  setup.orientation = Orientation::kPortrait;
  PageMargins margins;
  margins.left = 1.0;
  margins.right = 1.0;
  margins.top = 0.5;
  margins.bottom = 0.5;
  margins.header = 0.25;
  margins.footer = 0.25;

  const PrintableArea area = compute_printable_area(setup, margins);
  EXPECT_NEAR(area.width_pt, 612.0 - 144.0, kTol);                // 2in horizontal.
  EXPECT_NEAR(area.height_pt, 792.0 - 72.0 - 36.0 - 28.0, kTol);  // 1in vertical + 0.5in H/F band
                                                                  // + 28pt text reservation.
}

TEST(PageSetupTest, OversizedMarginsClampBodyToZero) {
  PageSetup setup;
  setup.paper_size = 70;  // A6: small page.
  setup.orientation = Orientation::kPortrait;
  PageMargins margins;
  margins.left = 100.0;
  margins.right = 100.0;
  margins.top = 100.0;
  margins.bottom = 100.0;

  const PrintableArea area = compute_printable_area(setup, margins);
  EXPECT_DOUBLE_EQ(area.width_pt, 0.0);
  EXPECT_DOUBLE_EQ(area.height_pt, 0.0);
}

}  // namespace
}  // namespace print
}  // namespace formulon
