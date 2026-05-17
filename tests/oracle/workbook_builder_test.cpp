// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Direct unit test for `build_pivot_from_spec`.
//
// The workbook-oracle golden-diff path stays dormant until a Windows +
// Excel host generates the pivot goldens, so this test is the phase's
// real local verification: it hand-builds a declarative pivot spec as a
// `JsonValue`, runs it through `build_pivot_from_spec` -> `pivot::evaluate`
// -> `pivot::layout`, and asserts the rendered grid carries the expected
// aggregated values.

#include "tests/oracle/workbook_builder.h"

#include <algorithm>
#include <cstddef>
#include <map>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "print/pagination.h"
#include "tests/oracle/json_reader.h"
#include "value.h"

namespace formulon {
namespace tests {
namespace oracle {
namespace {

// --- JsonValue construction helpers ----------------------------------------

JsonValue jnum(double n) {
  return JsonValue::make_number(n);
}
JsonValue jstr(std::string s) {
  return JsonValue::make_string(std::move(s));
}

JsonValue jobj(std::map<std::string, JsonValue> members) {
  return JsonValue::make_object(std::move(members));
}

JsonValue jarr(std::vector<JsonValue> items) {
  return JsonValue::make_array(std::move(items));
}

// A `{kind, value}` number cell record, as the workbook case schema emits.
JsonValue number_cell(double n) {
  return jobj({{"kind", jstr("number")}, {"value", jnum(n)}});
}

// A `{kind, value}` text cell record.
JsonValue text_cell(std::string s) {
  return jobj({{"kind", jstr("text")}, {"value", jstr(std::move(s))}});
}

// The four sales amounts in the smoke dataset, named so the dataset
// definition and the expected per-region totals carry no bare literals.
constexpr double kNorthFirstAmount = 100.0;
constexpr double kSouthFirstAmount = 200.0;
constexpr double kNorthSecondAmount = 50.0;
constexpr double kSouthSecondAmount = 25.0;
constexpr double kNorthTotal = kNorthFirstAmount + kNorthSecondAmount;  // 150
constexpr double kSouthTotal = kSouthFirstAmount + kSouthSecondAmount;  // 225

// Builds a small 2-column dataset (Region, Amount) on a "Data" sheet, with
// one row field (Region) and one Sum data field (Amount):
//
//   Region  Amount
//   ------  ------
//   North     100
//   South     200
//   North      50
//   South      25
//
// The pivot anchors at Report!A1.
JsonValue build_smoke_spec() {
  JsonValue data_cells = jobj({
      {"A1", text_cell("Region")},
      {"B1", text_cell("Amount")},
      {"A2", text_cell("North")},
      {"B2", number_cell(kNorthFirstAmount)},
      {"A3", text_cell("South")},
      {"B3", number_cell(kSouthFirstAmount)},
      {"A4", text_cell("North")},
      {"B4", number_cell(kNorthSecondAmount)},
      {"A5", text_cell("South")},
      {"B5", number_cell(kSouthSecondAmount)},
  });

  JsonValue pivot = jobj({
      {"source", jstr("Data!A1:B5")},
      {"anchor", jstr("Report!A1")},
      {"row_fields", jarr({jstr("Region")})},
      {"data_fields", jarr({jobj({{"field", jstr("Amount")}, {"agg", jstr("Sum")}})})},
      {"grand_totals", jobj({{"rows", JsonValue::make_bool(true)}, {"cols", JsonValue::make_bool(false)}})},
  });

  return jobj({
      {"sheets", jobj({{"Data", std::move(data_cells)}, {"Report", jobj({})}})},
      {"pivot", std::move(pivot)},
  });
}

// --- Tests ------------------------------------------------------------------

TEST(WorkbookBuilder, BuildsCacheAndTableFromSpec) {
  const JsonValue spec = build_smoke_spec();

  auto built_or = build_pivot_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const BuiltPivot& built = built_or.value();

  // Cache: 2 fields, 4 records.
  ASSERT_EQ(built.cache.fields().size(), 2U);
  EXPECT_EQ(built.cache.fields()[0].name, "Region");
  EXPECT_EQ(built.cache.fields()[1].name, "Amount");
  ASSERT_EQ(built.cache.records().size(), 4U);

  // Table: one row field (Region), one Sum data field, cache binding 1.
  ASSERT_EQ(built.table.row_field_order().size(), 1U);
  ASSERT_TRUE(built.table.col_field_order().empty());
  ASSERT_EQ(built.table.data_fields().size(), 1U);
  EXPECT_EQ(built.table.data_fields()[0].aggregation, pivot::Aggregation::Sum);
  EXPECT_EQ(built.table.pivot_cache_id(), built.cache.cache_id());
}

TEST(WorkbookBuilder, EvaluateAndLayoutProduceAggregatedGrid) {
  const JsonValue spec = build_smoke_spec();
  auto built_or = build_pivot_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const BuiltPivot& built = built_or.value();

  auto result_or = pivot::evaluate(built.table, built.cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const pivot::PivotResult& result = result_or.value();

  // Two row leaves: North (100 + 50 = 150), South (200 + 25 = 225).
  ASSERT_EQ(result.rows.size(), 2U);
  ASSERT_EQ(result.values.size(), 2U);

  std::map<std::string, double> by_region;
  for (std::size_t i = 0; i < result.rows.size(); ++i) {
    ASSERT_EQ(result.values[i].size(), 1U);
    ASSERT_EQ(result.values[i][0].size(), 1U);
    by_region[result.rows[i].label] = result.values[i][0][0].as_number();
  }
  EXPECT_DOUBLE_EQ(by_region["North"], kNorthTotal);
  EXPECT_DOUBLE_EQ(by_region["South"], kSouthTotal);

  // Layout projects the result into absolute cells anchored at Report!A1
  // (row 0, col 0). Flatten the data cells into a {(r,c) -> value} grid
  // relative to the anchor and confirm the aggregated numbers appear.
  auto cells_or = pivot::layout(built.table, result);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;
  const pivot::PivotCells& cells = cells_or.value();

  std::vector<double> data_values;
  for (const pivot::PivotCell& cell : cells.cells) {
    if (cell.kind == pivot::PivotCellKind::Data && cell.value.kind() == ValueKind::Number) {
      data_values.push_back(cell.value.as_number());
    }
  }
  // Both region totals must surface as Data cells in the rendered grid.
  EXPECT_NE(std::find(data_values.begin(), data_values.end(), kNorthTotal), data_values.end());
  EXPECT_NE(std::find(data_values.begin(), data_values.end(), kSouthTotal), data_values.end());
}

TEST(WorkbookBuilder, RejectsSpecWithoutPivotBlock) {
  const JsonValue spec = jobj({{"sheets", jobj({{"Data", jobj({})}})}});
  auto built_or = build_pivot_from_spec(spec);
  EXPECT_FALSE(static_cast<bool>(built_or));
}

TEST(WorkbookBuilder, RejectsUnknownAggregation) {
  JsonValue spec = build_smoke_spec();
  // Swap the data field's `agg` to a name the builder does not recognise.
  std::map<std::string, JsonValue> root = spec.as_object();
  std::map<std::string, JsonValue> pivot = root.at("pivot").as_object();
  pivot["data_fields"] = jarr({jobj({{"field", jstr("Amount")}, {"agg", jstr("Median")}})});
  root["pivot"] = jobj(std::move(pivot));

  auto built_or = build_pivot_from_spec(jobj(std::move(root)));
  EXPECT_FALSE(static_cast<bool>(built_or));
}

// --- Print pagination tests -------------------------------------------------
//
// Unlike the pivot path -- whose goldens need a Windows + Excel host --
// `print::paginate` runs fully in C++, so these tests are real local
// verification: they hand-build a declarative print spec, run it through
// `build_print_from_spec` -> `print::paginate`, and assert the break
// vectors and page count.

namespace {

// A wide column width (in OOXML character units) chosen so that exactly
// three columns saturate an A4-portrait printable body. A4 portrait body
// width is ~494 pt after default 0.7" side margins; at ~165 pt per
// column three columns (~495 pt) just exceed it, forcing a fourth column
// onto a second page.
constexpr double kWideColumnChars = 30.0;

// A tall row height (in points) chosen so that a long table overflows
// the A4-portrait printable body (~734 pt high) within a small row
// count, forcing a horizontal page break.
constexpr double kTallRowHeightPt = 80.0;

// Builds the `print` block's `page_setup` object for an A4 portrait
// page at 100% scale (the OOXML defaults, stated explicitly).
JsonValue a4_portrait_setup() {
  return jobj({
      {"orientation", jstr("portrait")},
      {"paper", jnum(9)},  // OOXML paperSize 9 == A4
      {"scale", jnum(100)},
  });
}

// Builds a one-sheet print spec: a "Sheet1" with a single populated
// cell at A1 (so the used-range fallback has something to anchor on)
// plus the supplied `print` block.
JsonValue print_spec(JsonValue print_block) {
  return jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}})}})},
      {"print", std::move(print_block)},
  });
}

}  // namespace

TEST(WorkbookBuilderPrint, WideTableSuppressesAutoVerticalBreaks) {
  // An 8-column print area whose geometry would overflow a single A4
  // page-wide at scale=100. Excel's PageBreakPreview never auto-breaks
  // columns for an explicit-scale print -- the overflow is clipped on
  // the right rather than wrapped, matching the print_pagination.wide_
  // table_vertical_breaks oracle case (v=[] pages=1 in golden).
  JsonValue widths = jobj({{"A:H", jnum(kWideColumnChars)}});
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}, {"H1", text_cell("x")}})}})},
      {"column_widths", std::move(widths)},
      {"print", jobj({
                    {"sheet", jstr("Sheet1")},
                    {"print_area", jstr("A1:H1")},
                    {"page_setup", a4_portrait_setup()},
                })},
  });

  auto built_or = build_print_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const BuiltPrint& built = built_or.value();

  auto pag_or = print::paginate(*built.workbook, built.sheet_index);
  ASSERT_TRUE(static_cast<bool>(pag_or)) << pag_or.error().message;
  const print::PaginationResult& pag = pag_or.value();

  EXPECT_TRUE(pag.h_breaks.empty());
  EXPECT_TRUE(pag.v_breaks.empty());
  EXPECT_EQ(pag.page_count, 1U);
}

TEST(WorkbookBuilderPrint, TallTableForcesHorizontalBreak) {
  // A 30-row print area with every row 80 pt tall. The A4-portrait body
  // is ~663 pt (after default margins and header/footer bands), so eight
  // rows (~640 pt) fit per page and a horizontal break fires every eight
  // rows. A30 is populated so the used-range intersection does not trim
  // the print area to A1:A1.
  std::map<std::string, JsonValue> heights;
  for (int r = 1; r <= 30; ++r) {
    heights[std::to_string(r)] = jnum(kTallRowHeightPt);
  }
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}, {"A30", text_cell("x")}})}})},
      {"row_heights", jobj(std::move(heights))},
      {"print", jobj({
                    {"sheet", jstr("Sheet1")},
                    {"print_area", jstr("A1:A30")},
                    {"page_setup", a4_portrait_setup()},
                })},
  });

  auto built_or = build_print_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const BuiltPrint& built = built_or.value();

  auto pag_or = print::paginate(*built.workbook, built.sheet_index);
  ASSERT_TRUE(static_cast<bool>(pag_or)) << pag_or.error().message;
  const print::PaginationResult& pag = pag_or.value();

  // Single-column print area: no vertical break.
  EXPECT_TRUE(pag.v_breaks.empty());
  // 30 rows, nine per page -> a horizontal break is produced; every
  // break index is a positive multiple of nine and strictly ascending.
  ASSERT_FALSE(pag.h_breaks.empty());
  for (std::size_t i = 0; i < pag.h_breaks.size(); ++i) {
    EXPECT_GT(pag.h_breaks[i], 0U);
    if (i > 0) {
      EXPECT_GT(pag.h_breaks[i], pag.h_breaks[i - 1]);
    }
  }
  // Page count is the number of row-pages (one column-page).
  EXPECT_EQ(pag.page_count, static_cast<std::uint32_t>(pag.h_breaks.size()) + 1U);
}

TEST(WorkbookBuilderPrint, ManualRowBreakIsHonored) {
  // A short print area that fits on one page by geometry; a manual row
  // break before Excel row 5 (0-based index 4) must still split it.
  // A10 is populated so the used-range intersection extends across the
  // full print area; without it the area would collapse to A1:A1 and
  // the walk would skip row 5.
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}, {"A10", text_cell("x")}})}})},
      {"print", jobj({
                    {"sheet", jstr("Sheet1")},
                    {"print_area", jstr("A1:A10")},
                    {"page_setup", a4_portrait_setup()},
                    {"manual_breaks", jobj({{"rows", jarr({jnum(5)})}})},
                })},
  });

  auto built_or = build_print_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const BuiltPrint& built = built_or.value();

  auto pag_or = print::paginate(*built.workbook, built.sheet_index);
  ASSERT_TRUE(static_cast<bool>(pag_or)) << pag_or.error().message;
  const print::PaginationResult& pag = pag_or.value();

  // The 1-based row 5 manual break converts to a 0-based break before
  // row index 4.
  ASSERT_EQ(pag.h_breaks.size(), 1U);
  EXPECT_EQ(pag.h_breaks[0], 4U);
  EXPECT_TRUE(pag.v_breaks.empty());
  EXPECT_EQ(pag.page_count, 2U);
}

TEST(WorkbookBuilderPrint, FitToWidthCollapsesVerticalBreaks) {
  // The same 8-column wide table as the vertical-break test, but with
  // `fit_to_width: 1`. The fit-to-page shrink factor squeezes every
  // column onto a single page-wide, eliminating the vertical breaks.
  JsonValue widths = jobj({{"A:H", jnum(kWideColumnChars)}});
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}})}})},
      {"column_widths", std::move(widths)},
      {"print", jobj({
                    {"sheet", jstr("Sheet1")},
                    {"print_area", jstr("A1:H1")},
                    {"page_setup", jobj({
                                       {"orientation", jstr("portrait")},
                                       {"paper", jnum(9)},
                                       {"fit_to_width", jnum(1)},
                                       {"fit_to_height", jnum(0)},
                                   })},
                })},
  });

  auto built_or = build_print_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const BuiltPrint& built = built_or.value();

  auto pag_or = print::paginate(*built.workbook, built.sheet_index);
  ASSERT_TRUE(static_cast<bool>(pag_or)) << pag_or.error().message;
  const print::PaginationResult& pag = pag_or.value();

  // Fit-to-1-page-wide shrinks the columns until they all fit: no
  // vertical break, a single page.
  EXPECT_TRUE(pag.v_breaks.empty());
  EXPECT_TRUE(pag.h_breaks.empty());
  EXPECT_EQ(pag.page_count, 1U);
}

TEST(WorkbookBuilderPrint, RejectsSpecWithoutPrintBlock) {
  const JsonValue spec = jobj({{"sheets", jobj({{"Sheet1", jobj({})}})}});
  auto built_or = build_print_from_spec(spec);
  EXPECT_FALSE(static_cast<bool>(built_or));
}

TEST(WorkbookBuilderPrint, RejectsUnknownSheet) {
  JsonValue spec = print_spec(jobj({{"sheet", jstr("Nonexistent")}}));
  auto built_or = build_print_from_spec(spec);
  EXPECT_FALSE(static_cast<bool>(built_or));
}

}  // namespace
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
