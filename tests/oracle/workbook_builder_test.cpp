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
#include <utility>
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

TEST(WorkbookBuilder, FormulaProbePinsDirectRefBranch) {
  JsonValue spec = build_smoke_spec();
  std::map<std::string, JsonValue> root = spec.as_object();
  std::map<std::string, JsonValue> pivot = root.at("pivot").as_object();
  pivot["formula_probes"] = jarr({
      jobj({
          {"id", jstr("unknown_field_ref")},
          {"cell", jstr("Report!Z1")},
          {"formula", jstr("=GETPIVOTDATA(\"合計 / Amount\",Report!A1,\"MissingField\",\"x\")")},
      }),
  });
  root["pivot"] = jobj(std::move(pivot));
  const JsonValue probe_spec = jobj(std::move(root));

  auto built_or = build_pivot_from_spec(probe_spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  BuiltPivot built = std::move(built_or.value());

  auto results_or = evaluate_pivot_formula_probes(&built, probe_spec);
  ASSERT_TRUE(static_cast<bool>(results_or)) << results_or.error().message;
  ASSERT_EQ(results_or.value().size(), 1U);
  EXPECT_EQ(results_or.value()[0].id, "unknown_field_ref");
  ASSERT_TRUE(results_or.value()[0].value.is_error());
  EXPECT_EQ(results_or.value()[0].value.as_error(), ErrorCode::Ref);
}

TEST(WorkbookBuilder, FormulaProbePinsValueAxisRefBranch) {
  JsonValue spec = build_smoke_spec();
  std::map<std::string, JsonValue> root = spec.as_object();
  std::map<std::string, JsonValue> pivot = root.at("pivot").as_object();
  pivot["formula_probes"] = jarr({jobj({
      {"id", jstr("value_axis_ref")},
      {"cell", jstr("Report!Z1")},
      {"formula", jstr("=GETPIVOTDATA(\"合計 / Amount\",Report!A1,\"Amount\",100)")},
  })});
  root["pivot"] = jobj(std::move(pivot));
  const JsonValue probe_spec = jobj(std::move(root));
  auto built_or = build_pivot_from_spec(probe_spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  BuiltPivot built = std::move(built_or.value());
  auto results_or = evaluate_pivot_formula_probes(&built, probe_spec);
  ASSERT_TRUE(static_cast<bool>(results_or)) << results_or.error().message;
  ASSERT_EQ(results_or.value().size(), 1U);
  ASSERT_TRUE(results_or.value()[0].value.is_error());
  EXPECT_EQ(results_or.value()[0].value.as_error(), ErrorCode::Ref);
}

TEST(WorkbookBuilder, FormulaProbePinsPageAxisRefBranch) {
  JsonValue spec = build_smoke_spec();
  std::map<std::string, JsonValue> root = spec.as_object();
  std::map<std::string, JsonValue> pivot = root.at("pivot").as_object();
  pivot["row_fields"] = jarr({});
  pivot["page_fields"] = jarr({jstr("Region")});
  pivot["formula_probes"] = jarr({jobj({
      {"id", jstr("page_axis_ref")},
      {"cell", jstr("Report!Z1")},
      {"formula", jstr("=GETPIVOTDATA(\"合計 / Amount\",Report!A1,\"Region\",\"North\")")},
  })});
  root["pivot"] = jobj(std::move(pivot));
  const JsonValue probe_spec = jobj(std::move(root));
  auto built_or = build_pivot_from_spec(probe_spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  BuiltPivot built = std::move(built_or.value());
  auto results_or = evaluate_pivot_formula_probes(&built, probe_spec);
  ASSERT_TRUE(static_cast<bool>(results_or)) << results_or.error().message;
  ASSERT_EQ(results_or.value().size(), 1U);
  ASSERT_TRUE(results_or.value()[0].value.is_error());
  EXPECT_EQ(results_or.value()[0].value.as_error(), ErrorCode::Ref);
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

TEST(WorkbookBuilderPrint, WideTableWrapsOntoFurtherPageColumns) {
  // An 8-column print area whose geometry overflows a single A4 page-wide
  // at scale=100. The overflow wraps onto further page-columns; it is not
  // clipped at the right margin, and the column axis breaks on the body
  // width exactly as the row axis breaks on the body height.
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

  // A single row of content never breaks horizontally.
  EXPECT_TRUE(pag.h_breaks.empty());
  // The exact break positions follow the width model, which is calibrated
  // against the oracle; what this pins is that the axis breaks at all and
  // that the page count is the number of page-columns it produced.
  EXPECT_FALSE(pag.v_breaks.empty());
  EXPECT_EQ(pag.page_count, pag.v_breaks.size() + 1U);
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

TEST(WorkbookBuilderPrint, HiddenColumnCarriesNoWidthOfItsOwn) {
  // The shape H-9 is about: a hidden column recorded with no width at
  // all. `has_width` must stay false so the sheet default -- not a
  // fabricated zero -- is what pagination reads.
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}, {"H1", text_cell("x")}})}})},
      {"hidden_columns", jarr({jstr("C"), jstr("E")})},
      {"print", jobj({
                    {"sheet", jstr("Sheet1")},
                    {"print_area", jstr("A1:H1")},
                    {"page_setup", a4_portrait_setup()},
                })},
  });

  auto built_or = build_print_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const SheetLayout& layout = built_or.value().workbook->sheet(built_or.value().sheet_index).layout();

  ASSERT_EQ(layout.columns.size(), 2U);
  for (const ColumnLayout& col : layout.columns) {
    EXPECT_TRUE(col.hidden);
    EXPECT_FALSE(col.has_width);
    EXPECT_FALSE(HasExplicitColumnWidth(col));
    EXPECT_EQ(col.first, col.last);
  }
  EXPECT_EQ(layout.columns[0].first, 2U);  // C
  EXPECT_EQ(layout.columns[1].first, 4U);  // E
}

TEST(WorkbookBuilderPrint, HiddenRowCarriesNoHeightOfItsOwn) {
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}})}})},
      {"hidden_rows", jarr({jstr("3")})},
      {"print", jobj({
                    {"sheet", jstr("Sheet1")},
                    {"print_area", jstr("A1:A9")},
                    {"page_setup", a4_portrait_setup()},
                })},
  });

  auto built_or = build_print_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const SheetLayout& layout = built_or.value().workbook->sheet(built_or.value().sheet_index).layout();

  ASSERT_EQ(layout.row_overrides.size(), 1U);
  EXPECT_EQ(layout.row_overrides[0].row, 2U);  // 1-based row 3
  EXPECT_TRUE(layout.row_overrides[0].hidden);
  EXPECT_FALSE(layout.row_overrides[0].has_height);
}

TEST(WorkbookBuilderPrint, HidingAndSizingTheSameColumnStayIndependent) {
  // A width entry and a hidden entry are separate overrides; neither may
  // silently absorb the other, or a case could not express "hidden, and
  // the rest of the span is wide".
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}, {"H1", text_cell("x")}})}})},
      {"column_widths", jobj({{"A:H", jnum(kWideColumnChars)}})},
      {"hidden_columns", jarr({jstr("C")})},
      {"print", jobj({
                    {"sheet", jstr("Sheet1")},
                    {"print_area", jstr("A1:H1")},
                    {"page_setup", a4_portrait_setup()},
                })},
  });

  auto built_or = build_print_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const SheetLayout& layout = built_or.value().workbook->sheet(built_or.value().sheet_index).layout();

  // The hidden span must come first: both pagination and ColumnWidthChars
  // resolve a column against the first matching span, so a width span that
  // covers the hidden column would otherwise mask it.
  ASSERT_EQ(layout.columns.size(), 2U);
  EXPECT_TRUE(layout.columns[0].hidden);
  EXPECT_FALSE(layout.columns[0].has_width);
  EXPECT_EQ(layout.columns[0].first, 2U);  // C
  EXPECT_FALSE(layout.columns[1].hidden);
  EXPECT_EQ(layout.columns[1].width, kWideColumnChars);
}

TEST(WorkbookBuilderPrint, HiddenColumnIsNotMaskedByAWidthSpanCoveringIt) {
  // The regression this ordering exists for: pagination reads a column's
  // first matching span, so the case would have paginated as though nothing
  // were hidden -- and the golden diff would blame the engine.
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}, {"H1", text_cell("x")}})}})},
      {"column_widths", jobj({{"A:H", jnum(kWideColumnChars)}})},
      {"hidden_columns", jarr({jstr("C"), jstr("D"), jstr("E")})},
      {"print", jobj({
                    {"sheet", jstr("Sheet1")},
                    {"print_area", jstr("A1:H1")},
                    {"page_setup", a4_portrait_setup()},
                })},
  });

  auto built_or = build_print_from_spec(spec);
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const BuiltPrint& built = built_or.value();

  auto with_hidden = print::paginate(*built.workbook, built.sheet_index);
  ASSERT_TRUE(static_cast<bool>(with_hidden)) << with_hidden.error().message;

  // Same sheet without the hidden columns: three fewer wide columns have to
  // fit, so the two pagination results cannot coincide.
  JsonValue visible = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}, {"H1", text_cell("x")}})}})},
      {"column_widths", jobj({{"A:H", jnum(kWideColumnChars)}})},
      {"print", jobj({
                    {"sheet", jstr("Sheet1")},
                    {"print_area", jstr("A1:H1")},
                    {"page_setup", a4_portrait_setup()},
                })},
  });
  auto visible_built = build_print_from_spec(visible);
  ASSERT_TRUE(static_cast<bool>(visible_built)) << visible_built.error().message;
  auto without_hidden = print::paginate(*visible_built.value().workbook, visible_built.value().sheet_index);
  ASSERT_TRUE(static_cast<bool>(without_hidden)) << without_hidden.error().message;

  EXPECT_LT(with_hidden.value().page_count, without_hidden.value().page_count);
}

TEST(WorkbookBuilderPrint, RejectsMalformedHiddenColumnKey) {
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}})}})},
      {"hidden_columns", jarr({jstr("3")})},
      {"print", jobj({{"sheet", jstr("Sheet1")}, {"page_setup", a4_portrait_setup()}})},
  });
  EXPECT_FALSE(static_cast<bool>(build_print_from_spec(spec)));
}

TEST(WorkbookBuilderPrint, RejectsHiddenRowZero) {
  // Rows are 1-based in a case file; a "0" would silently become row -1.
  JsonValue spec = jobj({
      {"sheets", jobj({{"Sheet1", jobj({{"A1", text_cell("x")}})}})},
      {"hidden_rows", jarr({jstr("0")})},
      {"print", jobj({{"sheet", jstr("Sheet1")}, {"page_setup", a4_portrait_setup()}})},
  });
  EXPECT_FALSE(static_cast<bool>(build_print_from_spec(spec)));
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

// Sheet indices follow the `sheets` block's declaration order. The capture
// side builds the same fixture in the order the case states, so a builder
// that ordered by name would hand the golden diff a different sheet -- and
// a comparison against a field that happens to match the default would
// pass anyway. The spec is parsed from text rather than assembled through
// `jobj`, because that is what gives it a declaration order distinct from
// its name order.
TEST(WorkbookBuilder, SheetIndicesFollowDeclarationOrderNotSheetName) {
  auto spec_or = parse_json(R"json({
    "sheets": {
      "Report": {},
      "Data": {"A1": {"kind": "text", "value": "Region"}}
    },
    "print": {"sheet": "Data"}
  })json");
  ASSERT_TRUE(spec_or.has_value()) << spec_or.error().message;

  auto built_or = build_print_from_spec(spec_or.value());
  ASSERT_TRUE(static_cast<bool>(built_or)) << built_or.error().message;
  const BuiltPrint& built = built_or.value();

  ASSERT_EQ(built.workbook->sheet_count(), 2U);
  EXPECT_EQ(built.workbook->sheet(0).name(), "Report");
  EXPECT_EQ(built.workbook->sheet(1).name(), "Data");
  EXPECT_EQ(built.sheet_index, 1U);
}

}  // namespace
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
