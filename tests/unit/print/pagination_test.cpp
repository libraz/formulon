//
// Unit tests for the page-break engine (`src/print/pagination`).
//
// The tests assert structural correctness — page count, break ordering,
// and which track each break precedes — rather than exact 1-bit parity
// with Excel's font-metric-dependent rounding.

#include "print/pagination.h"

#include <algorithm>
#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "print/page_setup.h"
#include "print/print_area.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

io::DefinedName PrintArea(std::string formula, std::int32_t sheet_id) {
  io::DefinedName dn;
  dn.name = "_xlnm.Print_Area";
  dn.formula = std::move(formula);
  dn.local_sheet_id = sheet_id;
  return dn;
}

// A sheet-scoped built-in name (`_xlnm.Print_Titles` and friends).
io::DefinedName SheetScopedName(std::string name, std::string formula, std::int32_t sheet_id) {
  io::DefinedName dn;
  dn.name = std::move(name);
  dn.formula = std::move(formula);
  dn.local_sheet_id = sheet_id;
  return dn;
}

// Sets a uniform column width (in OOXML character units) for [first, last].
void SetColumnWidth(Sheet* sheet, std::uint32_t first, std::uint32_t last, double width) {
  ColumnLayout span;
  span.first = first;
  span.last = last;
  span.width = width;
  sheet->mutable_layout().columns.push_back(span);
}

// Returns the furthest track index any break precedes, or 0 when there
// are no breaks.
std::uint32_t MaxBreak(const std::vector<std::uint32_t>& breaks) {
  return breaks.empty() ? 0U : *std::max_element(breaks.begin(), breaks.end());
}

// Sets the height (in points) of a single row.
void SetRowHeight(Sheet* sheet, std::uint32_t row, double height) {
  RowLayout layout;
  layout.row = row;
  layout.height = height;
  sheet->mutable_layout().row_overrides.push_back(layout);
}

TEST(PaginationTest, EmptySheetWithNoPrintAreaProducesNoPages) {
  Workbook wb = Workbook::create();
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 0U);
  EXPECT_TRUE(result.value().print_area.empty());
}

TEST(PaginationTest, OutOfRangeSheetIndexIsRejected) {
  Workbook wb = Workbook::create();
  auto result = paginate(wb, 99);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(PaginationTest, SingleSmallPrintAreaIsOnePage) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$C$3", 0)});
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 1U);
  EXPECT_TRUE(result.value().h_breaks.empty());
  EXPECT_TRUE(result.value().v_breaks.empty());
}

TEST(PaginationTest, WideTableWrapsOntoFurtherPageColumns) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // 20 very wide columns (60 char units ~= 315 pt each) overrun the ~494 pt
  // A4 body width several times, so the overflow wraps onto further
  // page-columns. The column axis breaks on the body width exactly as the
  // row axis breaks on the body height; it does not clip at the right
  // margin. Breaks are ascending, strictly inside the area, and each one
  // starts a page.
  SetColumnWidth(&sheet, 0, 19, 60.0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$T$5", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const std::vector<std::uint32_t>& breaks = result.value().v_breaks;
  ASSERT_FALSE(breaks.empty());
  for (std::size_t i = 1; i < breaks.size(); ++i) {
    EXPECT_LT(breaks[i - 1], breaks[i]);
  }
  EXPECT_GT(breaks.front(), 0U);
  EXPECT_LE(breaks.back(), 19U);
  EXPECT_EQ(result.value().page_count, breaks.size() + 1U);
}

TEST(PaginationTest, TallTableForcesHorizontalBreak) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // Inflate the first 30 rows to 100 pt each so a 30-row print area
  // (3000 pt) far exceeds the ~734 pt A4 body height.
  for (std::uint32_t row = 0; row < 30; ++row) {
    SetRowHeight(&sheet, row, 100.0);
  }
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$B$30", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_FALSE(result.value().h_breaks.empty());
  for (std::size_t i = 1; i < result.value().h_breaks.size(); ++i) {
    EXPECT_LT(result.value().h_breaks[i - 1], result.value().h_breaks[i]);
  }
  EXPECT_GE(result.value().page_count, 2U);
}

TEST(PaginationTest, EmptySheetWithFullColumnPrintAreaProducesNoPages) {
  // A whole-column print area on an empty sheet must short-circuit rather
  // than walk all 1,048,576 rows. The result mirrors the empty-sheet /
  // no-content spec: zero pages, returned immediately.
  Workbook wb = Workbook::create();
  wb.set_defined_names({PrintArea("Sheet1!$A:$A", 0)});
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 0U);
  EXPECT_TRUE(result.value().h_breaks.empty());
  EXPECT_TRUE(result.value().v_breaks.empty());
}

TEST(PaginationTest, EmptySheetWithFullRowPrintAreaProducesNoPages) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({PrintArea("Sheet1!$1:$1", 0)});
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 0U);
}

// A whole-column print area on a sheet that *does* carry content is the
// ordinary case, and it must be bounded by the content just as the
// empty-sheet case is. The row track vector and the per-area walk are
// both sized from the effective rectangle, so a rectangle that still
// claims all 1,048,576 rows costs megabytes of transient allocation and
// a walk to match -- and the page-count limit cannot intervene, because
// it is only consulted after that walk. Every break index landing
// inside the populated extent is the observable form of that bound: a
// walk that ran past the content would report breaks past it.
TEST(PaginationTest, FullColumnPrintAreaOnPopulatedSheetWalksOnlyThePopulatedRows) {
  const auto build = [](const char* area) {
    Workbook wb = Workbook::create();
    Sheet& sheet = wb.sheet(0);
    for (std::uint32_t row = 0; row < 10U; ++row) {
      SetRowHeight(&sheet, row, 100.0);  // ~7 rows to an A4 body
      for (std::uint32_t col = 0; col < 4U; ++col) {
        sheet.set_cell_value(row, col, Value::number(1.0));
      }
    }
    wb.set_defined_names({PrintArea(area, 0)});
    return wb;
  };
  Workbook unbounded = build("Sheet1!$A:$D");
  Workbook bounded = build("Sheet1!$A$1:$D$10");

  auto result = paginate(unbounded, 0);
  auto control = paginate(bounded, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_TRUE(static_cast<bool>(control)) << control.error().message;

  // No track outside the populated 10 rows was walked. Asserted on the
  // furthest break rather than per break: a walk over the whole grid
  // produces breaks in the six figures, and one bound states it as well
  // as a hundred thousand identical failures would.
  EXPECT_LT(MaxBreak(result.value().h_breaks), 10U) << "break reported past the populated extent";
  EXPECT_LE(result.value().h_breaks.size(), 10U);
  EXPECT_LE(result.value().page_count, 10U);
  // Clipping the unbounded edge to the content is exactly the bounded
  // declaration, breaks included.
  EXPECT_EQ(result.value().page_count, control.value().page_count);
  EXPECT_EQ(result.value().h_breaks, control.value().h_breaks);
  EXPECT_EQ(result.value().v_breaks, control.value().v_breaks);
  // The reported print area still mirrors what the workbook declared:
  // the clip drives pagination only, it is not a reporting change.
  ASSERT_EQ(result.value().print_area.size(), 1U);
  EXPECT_EQ(result.value().print_area.front().last_row, Sheet::kMaxRows - 1U);
}

TEST(PaginationTest, FullRowPrintAreaOnPopulatedSheetWalksOnlyThePopulatedColumns) {
  const auto build = [](const char* area) {
    Workbook wb = Workbook::create();
    Sheet& sheet = wb.sheet(0);
    SetColumnWidth(&sheet, 0U, 3U, 60.0);  // ~315 pt each, ~1 column to a body
    for (std::uint32_t row = 0; row < 2U; ++row) {
      for (std::uint32_t col = 0; col < 4U; ++col) {
        sheet.set_cell_value(row, col, Value::number(1.0));
      }
    }
    wb.set_defined_names({PrintArea(area, 0)});
    return wb;
  };
  Workbook unbounded = build("Sheet1!$1:$2");
  Workbook bounded = build("Sheet1!$A$1:$D$2");

  auto result = paginate(unbounded, 0);
  auto control = paginate(bounded, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_TRUE(static_cast<bool>(control)) << control.error().message;

  EXPECT_LT(MaxBreak(result.value().v_breaks), 4U) << "break reported past the populated extent";
  EXPECT_LE(result.value().v_breaks.size(), 4U);
  EXPECT_EQ(result.value().page_count, control.value().page_count);
  EXPECT_EQ(result.value().v_breaks, control.value().v_breaks);
}

TEST(PaginationTest, FullColumnPrintAreaStartingPastTheContentProducesNoPages) {
  // The clip leaves this rectangle empty (content ends at row 10, the
  // area starts at row 100), so it contributes no pages rather than
  // paginating a million blank rows.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t row = 0; row < 10U; ++row) {
    sheet.set_cell_value(row, 0U, Value::number(1.0));
  }
  wb.set_defined_names({PrintArea("Sheet1!$A$100:$D$1048576", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 0U);
  EXPECT_TRUE(result.value().h_breaks.empty());
}

// Repeated print titles are printed content: `<pageSetup scale>` shrinks
// them exactly as it shrinks the data. The body they are deducted from
// is physical page space, so the deduction has to be their scaled size.
// Deducting the raw model size reserves `1 / scale` times too much band
// and pushes data onto later pages.
//
// The expected break is derived here rather than hardcoded: the body
// comes from the same `compute_printable_area` the engine uses, and the
// arithmetic below *is* the invariant -- reserve = title height x scale,
// with the data rows scaled by the same factor.
TEST(PaginationTest, PrintTitleReservationIsScaledLikeTheDataItCompetesWith) {
  constexpr double kRowHeightPt = 30.0;
  constexpr std::uint32_t kScalePercent = 50U;
  constexpr double kScale = kScalePercent / 100.0;
  // Five title rows, which is also the engine's minimum reserve band, so
  // the floor and the actual title height agree and neither can mask the
  // other.
  constexpr double kTitleRows = 5.0;

  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.mutable_format_defaults().has_default_row_height = true;
  sheet.mutable_format_defaults().default_row_height = kRowHeightPt;
  for (std::uint32_t row = 0; row < 200U; ++row) {
    sheet.set_cell_value(row, 0U, Value::number(1.0));
  }
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$A$200", 0), SheetScopedName("_xlnm.Print_Titles", "Sheet1!$1:$5", 0)});
  sheet.mutable_print_settings().page_setup.scale = kScalePercent;

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_FALSE(result.value().h_breaks.empty());

  const PrintableArea body =
      compute_printable_area(sheet.print_settings().page_setup, sheet.print_settings().page_margins);
  const double reserved_pt = kTitleRows * kRowHeightPt * kScale;
  const double printed_row_pt = kRowHeightPt * kScale;
  const auto rows_per_page = static_cast<std::uint32_t>((body.height_pt - reserved_pt) / printed_row_pt);
  EXPECT_EQ(result.value().h_breaks.front(), rows_per_page);
}

// The other half of the scale contract. `fitToPage` derives the factor
// instead of reading it, and repeated titles are part of what has to
// fit: every one of the N pages reprints them, so N pages carry N title
// bands. A factor that charged for the titles only once would come out
// too large and leave the last page spilling onto another.
//
// Asserted against the same request without titles rather than against
// the requested number outright: an exact fit lands mid-track, and the
// walk only breaks on whole rows, so a request can round to one page
// more. That rounding is a property of the fit itself and applies with
// or without titles -- comparing the two isolates the title accounting
// from it. `fitToHeight=1` has no such ambiguity and is pinned exactly.
TEST(PaginationTest, FitToPageAccommodatesRepeatedTitlesOnEveryPage) {
  const auto build = [](std::uint32_t pages_tall, bool with_titles) {
    Workbook wb = Workbook::create();
    Sheet& sheet = wb.sheet(0);
    sheet.mutable_format_defaults().has_default_row_height = true;
    sheet.mutable_format_defaults().default_row_height = 30.0;
    for (std::uint32_t row = 0; row < 200U; ++row) {
      sheet.set_cell_value(row, 0U, Value::number(1.0));
    }
    std::vector<io::DefinedName> names{PrintArea("Sheet1!$A$1:$A$200", 0)};
    if (with_titles) {
      names.push_back(SheetScopedName("_xlnm.Print_Titles", "Sheet1!$1:$5", 0));
    }
    wb.set_defined_names(std::move(names));
    sheet.mutable_print_settings().page_setup.fit_to_page = true;
    sheet.mutable_print_settings().page_setup.fit_to_width = 1;
    sheet.mutable_print_settings().page_setup.fit_to_height = pages_tall;
    return wb;
  };

  for (const std::uint32_t pages_tall : {1U, 2U, 3U}) {
    Workbook titled = build(pages_tall, true);
    Workbook plain = build(pages_tall, false);
    auto titled_result = paginate(titled, 0);
    auto plain_result = paginate(plain, 0);
    ASSERT_TRUE(static_cast<bool>(titled_result)) << titled_result.error().message;
    ASSERT_TRUE(static_cast<bool>(plain_result)) << plain_result.error().message;
    EXPECT_EQ(titled_result.value().page_count, plain_result.value().page_count)
        << "repeated titles cost a page at fit_to_height=" << pages_tall;
  }

  Workbook single = build(1U, true);
  auto single_result = paginate(single, 0);
  ASSERT_TRUE(static_cast<bool>(single_result)) << single_result.error().message;
  EXPECT_EQ(single_result.value().page_count, 1U);
  EXPECT_TRUE(single_result.value().h_breaks.empty());
}

TEST(PaginationTest, BoundedPrintAreaIsNotClippedToTheContent) {
  // The clip applies only to an edge declared at the grid limit. A
  // declared rectangle that names real geometry drives the page grid
  // whether or not the cells inside it are populated, so a sparse sheet
  // and a dense one break identically under the same declaration.
  const auto build = [](bool dense) {
    Workbook wb = Workbook::create();
    Sheet& sheet = wb.sheet(0);
    SetColumnWidth(&sheet, 0U, 7U, 60.0);
    for (std::uint32_t row = 0; row < 30U; ++row) {
      for (std::uint32_t col = 0; col < (dense ? 8U : 1U); ++col) {
        sheet.set_cell_value(row, col, Value::number(1.0));
      }
    }
    wb.set_defined_names({PrintArea("Sheet1!$A$1:$H$30", 0)});
    return wb;
  };
  Workbook sparse = build(false);
  Workbook dense = build(true);

  auto sparse_result = paginate(sparse, 0);
  auto dense_result = paginate(dense, 0);
  ASSERT_TRUE(static_cast<bool>(sparse_result)) << sparse_result.error().message;
  ASSERT_TRUE(static_cast<bool>(dense_result)) << dense_result.error().message;

  EXPECT_FALSE(sparse_result.value().v_breaks.empty());
  EXPECT_EQ(sparse_result.value().v_breaks, dense_result.value().v_breaks);
  EXPECT_EQ(sparse_result.value().h_breaks, dense_result.value().h_breaks);
  EXPECT_EQ(sparse_result.value().page_count, dense_result.value().page_count);
}

TEST(PaginationTest, HiddenRowsAreExcludedFromPaginationExtent) {
  // 30 rows at 100 pt each force several horizontal breaks; hiding all but
  // the first five collapses the printed height to a single page.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t row = 0; row < 30; ++row) {
    RowLayout layout;
    layout.row = row;
    layout.height = 100.0;
    layout.hidden = row >= 5;  // rows 6..30 hidden.
    sheet.mutable_layout().row_overrides.push_back(layout);
  }
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$B$30", 0)});
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  // Five visible rows (500 pt) fit in one A4 body; hidden rows add nothing.
  EXPECT_TRUE(result.value().h_breaks.empty());
  EXPECT_EQ(result.value().page_count, 1U);
}

TEST(PaginationTest, HiddenWidthlessColumnDoesNotConsumeFitToPageWidth) {
  // fit_to_width=1 makes the total column width set the scale factor, and
  // the scale then decides how many rows fit per page. A hidden column
  // whose width leaked into that total would shrink the sheet further and
  // show up as a smaller page count -- observable without reaching into an
  // internal geometry helper.
  const auto make_workbook = [](bool hidden, bool explicit_width) {
    Workbook wb = Workbook::create();
    Sheet& sheet = wb.sheet(0);
    sheet.set_cell_value(0U, 0U, Value::number(1.0));
    sheet.set_cell_value(4999U, 1U, Value::number(2.0));

    ColumnLayout first;
    first.first = 0U;
    first.last = 0U;
    first.has_width = true;
    first.width = 100.0;
    sheet.mutable_layout().columns.push_back(first);

    ColumnLayout second;
    second.first = 1U;
    second.last = 1U;
    second.hidden = hidden;
    if (explicit_width) {
      second.has_width = true;
      second.width = hidden ? 100.0 : 0.0;
    }
    sheet.mutable_layout().columns.push_back(second);

    wb.set_defined_names({PrintArea("Sheet1!$A$1:$B$5000", 0)});
    sheet.mutable_print_settings().page_setup.fit_to_page = true;
    sheet.mutable_print_settings().page_setup.fit_to_width = 1;
    sheet.mutable_print_settings().page_setup.fit_to_height = 0;
    return wb;
  };

  Workbook hidden_widthless = make_workbook(true, false);
  Workbook hidden_explicit = make_workbook(true, true);
  Workbook visible_absent = make_workbook(false, false);
  Workbook visible_zero = make_workbook(false, true);

  auto hidden_widthless_result = paginate(hidden_widthless, 0);
  auto hidden_explicit_result = paginate(hidden_explicit, 0);
  auto visible_absent_result = paginate(visible_absent, 0);
  auto visible_zero_result = paginate(visible_zero, 0);
  ASSERT_TRUE(static_cast<bool>(hidden_widthless_result)) << hidden_widthless_result.error().message;
  ASSERT_TRUE(static_cast<bool>(hidden_explicit_result)) << hidden_explicit_result.error().message;
  ASSERT_TRUE(static_cast<bool>(visible_absent_result)) << visible_absent_result.error().message;
  ASSERT_TRUE(static_cast<bool>(visible_zero_result)) << visible_zero_result.error().message;

  // Hidden always wins over width, and an explicit width of zero is not the
  // same as an absent width (which falls back to the standard default).
  EXPECT_EQ(hidden_widthless_result.value().page_count, hidden_explicit_result.value().page_count);
  EXPECT_EQ(hidden_widthless_result.value().page_count, visible_zero_result.value().page_count);
  EXPECT_GT(hidden_widthless_result.value().page_count, visible_absent_result.value().page_count);
}

TEST(PaginationTest, OutlineOnlyRowUsesDefaultHeight) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(199, 0, Value::number(2.0));
  // Outline metadata only: no `ht`, so the row keeps the sheet default.
  sheet.mutable_layout().row_overrides.push_back(RowLayout{24U, 0.0, false, 1U, false});
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$A$200", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;

  // Compared against the same sheet with no override at all. Asserting a
  // page count here would only say what the current body height is; what
  // has to hold is that carrying outline metadata changes nothing.
  Workbook control = Workbook::create();
  control.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  control.sheet(0).set_cell_value(199, 0, Value::number(2.0));
  control.set_defined_names({PrintArea("Sheet1!$A$1:$A$200", 0)});
  auto control_result = paginate(control, 0);
  ASSERT_TRUE(static_cast<bool>(control_result)) << control_result.error().message;

  EXPECT_GT(result.value().page_count, 1U);
  EXPECT_EQ(result.value().page_count, control_result.value().page_count);
  EXPECT_EQ(result.value().h_breaks, control_result.value().h_breaks);
}

TEST(PaginationTest, ExplicitVisibleRowWithoutHeightUsesDefaultHeight) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(199, 0, Value::number(2.0));
  // This models `<row hidden="0">`: an explicit row override with no `ht`,
  // so it must not collapse during pagination.
  sheet.mutable_layout().row_overrides.push_back(RowLayout{24U, 0.0, false, 0U, false});
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$A$200", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;

  // Same control as the outline-only case: the override must be inert.
  Workbook control = Workbook::create();
  control.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  control.sheet(0).set_cell_value(199, 0, Value::number(2.0));
  control.set_defined_names({PrintArea("Sheet1!$A$1:$A$200", 0)});
  auto control_result = paginate(control, 0);
  ASSERT_TRUE(static_cast<bool>(control_result)) << control_result.error().message;

  EXPECT_GT(result.value().page_count, 1U);
  EXPECT_EQ(result.value().page_count, control_result.value().page_count);
  EXPECT_EQ(result.value().h_breaks, control_result.value().h_breaks);
}

TEST(PaginationTest, FiftyThousandMetadataOnlyRowsRemainPractical) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  constexpr std::uint32_t kRows = 50'000U;
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(kRows - 1U, 0, Value::number(2.0));
  sheet.mutable_layout().row_overrides.reserve(kRows);
  for (std::uint32_t row = 0; row < kRows; ++row) {
    sheet.mutable_layout().row_overrides.push_back(RowLayout{row, 0.0, false, 1U, false});
  }
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$A$50000", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_GT(result.value().page_count, 1U);
}

TEST(PaginationTest, ManualColumnBreakIsHonored) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$F$3", 0)});
  // Force a vertical break before column index 3 (column D).
  ManualBreak brk;
  brk.id = 3;
  brk.min = 0;
  brk.max = 0;
  brk.manual = true;
  sheet.mutable_print_settings().manual_col_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_FALSE(result.value().v_breaks.empty());
  EXPECT_EQ(result.value().v_breaks.front(), 3U);
  EXPECT_EQ(result.value().page_count, 2U);
}

TEST(PaginationTest, ManualRowBreakIsHonored) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$B$10", 0)});
  ManualBreak brk;
  brk.id = 4;         // Break before row index 4 (row 5).
  brk.manual = true;  // `man` defaults to false; a honored break is manual.
  sheet.mutable_print_settings().manual_row_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_FALSE(result.value().h_breaks.empty());
  EXPECT_EQ(result.value().h_breaks.front(), 4U);
  EXPECT_EQ(result.value().page_count, 2U);
}

TEST(PaginationTest, FitToWidthCollapsesToSingleColumnPage) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // Without fit-to-page this wide print area would need multiple
  // column-pages; fitToWidth=1 must shrink it to exactly one.
  SetColumnWidth(&sheet, 0, 19, 60.0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$T$5", 0)});
  sheet.mutable_print_settings().page_setup.fit_to_page = true;
  sheet.mutable_print_settings().page_setup.fit_to_width = 1;
  sheet.mutable_print_settings().page_setup.fit_to_height = 0;  // Unconstrained.

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().v_breaks.empty());

  // And the same area without fit-to-page does need several column pages,
  // so the collapse above is the fit doing work rather than the geometry
  // happening to fit.
  Workbook unfitted = Workbook::create();
  SetColumnWidth(&unfitted.sheet(0), 0, 19, 60.0);
  unfitted.set_defined_names({PrintArea("Sheet1!$A$1:$T$5", 0)});
  auto unfitted_result = paginate(unfitted, 0);
  ASSERT_TRUE(static_cast<bool>(unfitted_result)) << unfitted_result.error().message;
  EXPECT_FALSE(unfitted_result.value().v_breaks.empty());
}

TEST(PaginationTest, ScalePercentChangesPageCount) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // A tall single-column print area that fits on one page at 100% scale.
  // One column never breaks horizontally, so this exercises the ROW axis:
  // at 400% scale each row is four times as tall, overflowing the body and
  // forcing breaks.
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$A$40", 0)});

  sheet.mutable_print_settings().page_setup.scale = 100;
  auto at_full = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(at_full)) << at_full.error().message;

  sheet.mutable_print_settings().page_setup.scale = 400;
  auto enlarged = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(enlarged)) << enlarged.error().message;
  EXPECT_GT(enlarged.value().page_count, at_full.value().page_count);
}

TEST(PaginationTest, UsedRangeIsPaginatedWhenPrintAreaAbsent) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(2, 4, Value::number(2.0));

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  // `result.print_area` mirrors Excel's `PageSetup.PrintArea` exactly:
  // empty when no `_xlnm.Print_Area` defined name is set, even if the
  // sheet has populated cells. The used range is only consulted
  // internally to size the page grid.
  EXPECT_TRUE(result.value().print_area.empty());
  EXPECT_EQ(result.value().page_count, 1U);
}

TEST(PaginationTest, SpillPhantomsExtendUsedRangeForPagination) {
  // With no print area, pagination falls back to the used range, which must
  // include a dynamic-array spill's phantoms, not just its anchor. A tall
  // single-column spill A1:A60 (60 rows * 15 pt = 900 pt) exceeds the A4 body
  // height and forces at least one horizontal break; counting only the anchor
  // A1 would leave a single page.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells;
  cells.reserve(60);
  for (std::uint32_t i = 0; i < 60; ++i) {
    cells.push_back(Value::number(static_cast<double>(i + 1)));
  }
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 60U, 1U, std::move(cells)));

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_GE(result.value().page_count, 2U);
  EXPECT_FALSE(result.value().h_breaks.empty());
}

TEST(PaginationTest, AutomaticColumnBreakDoesNotForceAnExtraPage) {
  // Excel persists automatic breaks (man="0") once a sheet is previewed.
  // An auto column break must not be treated as a forced page boundary.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$F$3", 0)});
  ManualBreak brk;
  brk.id = 3;
  brk.manual = false;  // Automatic break.
  sheet.mutable_print_settings().manual_col_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().v_breaks.empty());
  EXPECT_EQ(result.value().page_count, 1U);
}

TEST(PaginationTest, AutomaticRowBreakDoesNotForceAnExtraPage) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$B$10", 0)});
  ManualBreak brk;
  brk.id = 4;
  brk.manual = false;  // Automatic break.
  sheet.mutable_print_settings().manual_row_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().h_breaks.empty());
  EXPECT_EQ(result.value().page_count, 1U);
}

TEST(PaginationTest, MultiAreaColumnBreakCountedPerArea) {
  // Two print rectangles, each spanning columns A..F, share a manual column
  // break at column index 3. Each print area is an independent page grid
  // (area boundaries break the page), so the break splits BOTH areas' column
  // span — symmetric with how a shared manual ROW break splits both of two
  // side-by-side areas. Each area is a single row-band, so
  // page_count = 2 + 2 = 4. The break still appears once in the
  // de-duplicated v_breaks position collection.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$F$3,Sheet1!$A$5:$F$7", 0)});
  // Populate both rectangles so the used-range intersection keeps them.
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(0, 5, Value::number(1.0));
  sheet.set_cell_value(2, 5, Value::number(1.0));
  sheet.set_cell_value(4, 0, Value::number(1.0));
  sheet.set_cell_value(6, 5, Value::number(1.0));
  ManualBreak brk;
  brk.id = 3;  // Column D, inside both rectangles' column span.
  brk.manual = true;
  sheet.mutable_print_settings().manual_col_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 4U);
  // The break appears once in the de-duplicated v_breaks collection.
  ASSERT_EQ(result.value().v_breaks.size(), 1U);
  EXPECT_EQ(result.value().v_breaks.front(), 3U);
}

TEST(PaginationTest, MultiAreaRowAndColumnBreaksAreSymmetric) {
  // Symmetry check: two side-by-side areas (columns A..C and E..G) sharing a
  // manual ROW break, versus two stacked areas (rows 1..3 and 5..7) sharing
  // a manual COLUMN break, must produce the same per-area page count.
  //
  // Side-by-side areas + shared row break: each area splits into 2 row-bands
  // over 1 column-page -> 2 + 2 = 4.
  Workbook rows_wb = Workbook::create();
  Sheet& rows_sheet = rows_wb.sheet(0);
  rows_wb.set_defined_names({PrintArea("Sheet1!$A$1:$C$6,Sheet1!$E$1:$G$6", 0)});
  rows_sheet.set_cell_value(0, 0, Value::number(1.0));
  rows_sheet.set_cell_value(5, 0, Value::number(1.0));
  rows_sheet.set_cell_value(0, 4, Value::number(1.0));
  rows_sheet.set_cell_value(5, 6, Value::number(1.0));
  ManualBreak row_brk;
  row_brk.id = 3;  // Row 4, inside both areas' row span.
  row_brk.manual = true;
  rows_sheet.mutable_print_settings().manual_row_breaks.push_back(row_brk);
  auto rows_result = paginate(rows_wb, 0);
  ASSERT_TRUE(static_cast<bool>(rows_result)) << rows_result.error().message;
  EXPECT_EQ(rows_result.value().page_count, 4U);

  // Stacked areas + shared column break: mirror shape -> also 4.
  Workbook cols_wb = Workbook::create();
  Sheet& cols_sheet = cols_wb.sheet(0);
  cols_wb.set_defined_names({PrintArea("Sheet1!$A$1:$F$3,Sheet1!$A$5:$F$7", 0)});
  cols_sheet.set_cell_value(0, 0, Value::number(1.0));
  cols_sheet.set_cell_value(0, 5, Value::number(1.0));
  cols_sheet.set_cell_value(4, 0, Value::number(1.0));
  cols_sheet.set_cell_value(6, 5, Value::number(1.0));
  ManualBreak col_brk;
  col_brk.id = 3;  // Column D, inside both areas' column span.
  col_brk.manual = true;
  cols_sheet.mutable_print_settings().manual_col_breaks.push_back(col_brk);
  auto cols_result = paginate(cols_wb, 0);
  ASSERT_TRUE(static_cast<bool>(cols_result)) << cols_result.error().message;
  EXPECT_EQ(cols_result.value().page_count, rows_result.value().page_count);
}

// Gives rows [0, count) a height that no page body can hold two of, so each
// row occupies a page of its own and the row-page count equals `count`.
// 409.5 pt is the tallest height OOXML can express for a row.
void SetOnePageEachRowHeights(Sheet* sheet, std::uint32_t count) {
  constexpr double kMaxRowHeightPt = 409.5;
  sheet->mutable_layout().row_overrides.reserve(count);
  for (std::uint32_t row = 0; row < count; ++row) {
    sheet->mutable_layout().row_overrides.push_back(RowLayout{row, kMaxRowHeightPt, false, 0U, true});
  }
}

// Adds a manual column break before every column in [1, count), giving the
// column axis `count` pages.
void SetManualBreakBeforeEveryColumn(Sheet* sheet, std::uint32_t count) {
  auto& breaks = sheet->mutable_print_settings().manual_col_breaks;
  breaks.reserve(count);
  for (std::uint32_t col = 1; col < count; ++col) {
    ManualBreak brk;
    brk.id = col;
    brk.manual = true;
    breaks.push_back(brk);
  }
}

TEST(PaginationTest, PageGridBeyondTheCeilingIsRejected) {
  // A break before every track of a 524,288-row by 8,192-column grid is a
  // legal thing for a file to declare: neither `<row ht>` nor `<colBreaks>`
  // caps how many entries it may carry. Its true page count is 2^32, which a
  // 32-bit product wraps to 0. Pagination must refuse the request instead of
  // reporting a count it cannot represent.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  constexpr std::uint32_t kRows = 524288U;  // 2^19
  constexpr std::uint32_t kCols = 8192U;    // 2^13
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(kRows - 1U, kCols - 1U, Value::number(1.0));
  SetOnePageEachRowHeights(&sheet, kRows);
  SetManualBreakBeforeEveryColumn(&sheet, kCols);

  auto result = paginate(wb, 0);
  ASSERT_FALSE(static_cast<bool>(result)) << "page_count=" << result.value().page_count;
  EXPECT_EQ(result.error().code, FormulonErrorCode::kPrintPageCountOverflow);
}

TEST(PaginationTest, LargePageCountBelowTheCeilingIsExact) {
  // 100,000 row-pages across 101 column-pages is 10,100,000 pages: far past
  // what a 32-bit multiply of the two axes handles comfortably, yet inside
  // `kMaxPaginationPages`, so it must be counted rather than rejected.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  constexpr std::uint32_t kRows = 100000U;
  constexpr std::uint32_t kCols = 101U;
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(kRows - 1U, kCols - 1U, Value::number(1.0));
  SetOnePageEachRowHeights(&sheet, kRows);
  SetManualBreakBeforeEveryColumn(&sheet, kCols);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, kRows * kCols);
  EXPECT_EQ(result.value().h_breaks.size(), kRows - 1U);
  EXPECT_EQ(result.value().v_breaks.size(), kCols - 1U);
}

TEST(PaginationTest, MultiAreaPageCountIsExact) {
  // Two stacked areas, each split by a manual row break and the shared manual
  // column break: 2 row-pages * 2 column-pages per area, summed over both.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$F$6,Sheet1!$A$8:$F$13", 0)});
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(5, 5, Value::number(1.0));
  sheet.set_cell_value(7, 0, Value::number(1.0));
  sheet.set_cell_value(12, 5, Value::number(1.0));
  ManualBreak col_brk;
  col_brk.id = 3;  // Column D, inside both areas.
  col_brk.manual = true;
  sheet.mutable_print_settings().manual_col_breaks.push_back(col_brk);
  for (std::uint32_t row : {3U, 10U}) {  // One row break inside each area.
    ManualBreak row_brk;
    row_brk.id = row;
    row_brk.manual = true;
    sheet.mutable_print_settings().manual_row_breaks.push_back(row_brk);
  }

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 8U);
  EXPECT_EQ(result.value().h_breaks, (std::vector<std::uint32_t>{3U, 10U}));
  EXPECT_EQ(result.value().v_breaks, (std::vector<std::uint32_t>{3U}));
}

}  // namespace
}  // namespace print
}  // namespace formulon
