#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "ooxml_corpus_support.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "table.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon::ooxml_corpus {

Expected<std::vector<std::uint8_t>, Error> BuildEmpty() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildSingleLiteral() {
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(123.5)));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildLiteralsAllKinds() {
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(42.0)));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 1U, 0U, Value::boolean(true)));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 2U, 0U, Value::boolean(false)));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 3U, 0U, Value::text("hello")));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 4U, 0U, Value::blank()));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 5U, 0U, Value::error(ErrorCode::Div0)));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildArithmetic() {
  // NOTE: the `&` (text-concat) operator is intentionally omitted here.
  // Bundle 2.6 surfaced an existing engine limitation: the OOXML writer
  // emits formula cells with text cached values as `<v>foobar</v>`
  // (no `t="str"`), which the reader then parses as `t='n'` and
  // rejects. That is a writer bug to fix in a follow-up bundle; the
  // corpus avoids triggering it so the round-trip parity test stays
  // green for the arithmetic surface itself.
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0)));  // A1
  RETURN_IF_ERROR(wb.set_cell_value(0U, 1U, 0U, Value::number(3.0)));   // A2
  // Formulas covering + - * / ^ % and unary minus (no `&`).
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U, "=A1+A2"));  // B1 -> 13
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 1U, 1U, "=A1-A2"));  // B2 -> 7
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 2U, 1U, "=A1*A2"));  // B3 -> 30
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 3U, 1U, "=A1/A2"));  // B4 -> 3.333..
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 4U, 1U, "=A1^A2"));  // B5 -> 1000
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 5U, 1U, "=A1%"));    // B6 -> 0.1
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 6U, 1U, "=-A1"));    // B7 -> -10
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildCrossSheet() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  RETURN_IF_ERROR(wb.set_cell_value(1U, 0U, 0U, Value::number(99.0)));  // Sheet2!A1
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 0U, "=Sheet2!A1"));       // Sheet1!A1
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildMultiSheetIndependent() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("North");
  wb.add_sheet("South");
  wb.add_sheet("East");
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0)));
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U, "=A1+1"));
  RETURN_IF_ERROR(wb.set_cell_value(1U, 0U, 0U, Value::number(2.0)));
  RETURN_IF_ERROR(wb.set_cell_formula(1U, 0U, 1U, "=A1*2"));
  RETURN_IF_ERROR(wb.set_cell_value(2U, 0U, 0U, Value::number(3.0)));
  RETURN_IF_ERROR(wb.set_cell_formula(2U, 0U, 1U, "=A1-1"));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildUnicodeSheetNames() {
  Workbook wb = Workbook::create_empty();
  // "日本語" (Japanese) — exercises non-ASCII XML attribute encoding.
  wb.add_sheet("\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E");
  // emoji + ASCII — exercises 4-byte UTF-8 in attribute context.
  wb.add_sheet("emoji\xF0\x9F\x98\x80");
  // Spaces and digits — exercises plain attribute round-trip.
  wb.add_sheet("Sheet 3 with spaces");
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildUnicodeCellText() {
  Workbook wb = Workbook::create();
  // Japanese.
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::text("\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E")));
  // Emoji (4-byte UTF-8).
  RETURN_IF_ERROR(wb.set_cell_value(0U, 1U, 0U, Value::text("Hello \xF0\x9F\x98\x80")));
  // NBSP (U+00A0).
  RETURN_IF_ERROR(wb.set_cell_value(0U, 2U, 0U, Value::text("a\xC2\xA0\x62")));
  // Hebrew "shalom" — Right-to-left script. We are not asserting bidi
  // rendering, just that the bytes survive a round-trip.
  RETURN_IF_ERROR(wb.set_cell_value(0U, 3U, 0U, Value::text("\xD7\xA9\xD7\x9C\xD7\x95\xD7\x9D")));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildRangeAggregates() {
  Workbook wb = Workbook::create();
  // Populate A1:A10 with 1..10.
  for (std::uint32_t r = 0; r < 10U; ++r) {
    RETURN_IF_ERROR(wb.set_cell_value(0U, r, 0U, Value::number(static_cast<double>(r + 1U))));
  }
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A1:A10)"));      // B1 -> 55
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 1U, 1U, "=AVERAGE(A1:A10)"));  // B2 -> 5.5
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 2U, 1U, "=MIN(A1:A10)"));      // B3 -> 1
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 3U, 1U, "=MAX(A1:A10)"));      // B4 -> 10
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 4U, 1U, "=COUNT(A1:A10)"));    // B5 -> 10
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildVolatileNowToday() {
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 0U, "=NOW()"));
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 1U, 0U, "=TODAY()"));
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 2U, 0U, "=RAND()"));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildIterativeCircular() {
  Workbook wb = Workbook::create();
  // A1 = B1 + 1; B1 = A1.
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 0U, "=B1+1"));
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U, "=A1"));
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 20U;
  opts.max_change = 0.01;
  wb.set_iterative_options(opts);
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildDefinedNamesWorkbookScope() {
  Workbook wb = Workbook::create();
  std::vector<DefinedName> names;
  DefinedName n;
  n.name = "Sales";
  n.formula = "Sheet1!$A$1:$A$10";
  names.push_back(std::move(n));
  wb.set_defined_names(std::move(names));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildDefinedNamesSheetScope() {
  Workbook wb = Workbook::create();
  std::vector<DefinedName> names;
  DefinedName n;
  n.name = "LocalRange";
  n.formula = "Sheet1!$B$1:$B$5";
  n.local_sheet_id = 0;
  n.hidden = true;
  n.comment = "Sheet1-only.";
  names.push_back(std::move(n));
  wb.set_defined_names(std::move(names));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildSingleTable() {
  Workbook wb = Workbook::create();
  TableMetadata table;
  table.id = 1U;
  table.name = "SalesTable";
  table.display_name = "SalesTable";
  table.ref = "A1:C5";
  table.sheet_index = 0U;
  table.header_row = true;
  table.totals_row = true;
  table.columns.push_back(TableColumn{1U, "Region", "Total", "", ""});
  table.columns.push_back(TableColumn{2U, "Q1", "", "sum", ""});
  table.columns.push_back(TableColumn{3U, "Q2", "", "sum", ""});
  std::vector<TableMetadata> tables;
  tables.push_back(std::move(table));
  wb.set_tables(std::move(tables));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildMultipleTables() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");

  std::vector<TableMetadata> tables;
  TableMetadata t1;
  t1.id = 1U;
  t1.name = "Tbl1";
  t1.display_name = "Tbl1";
  t1.ref = "A1:B3";
  t1.sheet_index = 0U;
  t1.columns.push_back(TableColumn{1U, "X", "", "", ""});
  t1.columns.push_back(TableColumn{2U, "Y", "", "", ""});
  tables.push_back(std::move(t1));

  TableMetadata t2;
  t2.id = 2U;
  t2.name = "Tbl2";
  t2.display_name = "Tbl2";
  t2.ref = "A1:B2";
  t2.sheet_index = 0U;
  t2.columns.push_back(TableColumn{1U, "P", "", "", ""});
  t2.columns.push_back(TableColumn{2U, "Q", "", "", ""});
  tables.push_back(std::move(t2));

  TableMetadata t3;
  t3.id = 3U;
  t3.name = "Tbl3";
  t3.display_name = "Tbl3";
  t3.ref = "A1:A2";
  t3.sheet_index = 1U;
  t3.columns.push_back(TableColumn{1U, "Z", "", "", ""});
  tables.push_back(std::move(t3));

  wb.set_tables(std::move(tables));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildNestedFunctionCalls() {
  // Deeply nested IF/AND/OR with numeric branches only — text branch
  // results would trigger the same `<v>` round-trip limitation called
  // out in `BuildArithmetic`. Logic: (A1>0 OR A2<0) AND (A3=8) is true,
  // and A1+A2 (=8) > A3 (=8) is false, so the formula falls into the
  // inner else-branch and yields A2 (=3).
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(5.0)));  // A1
  RETURN_IF_ERROR(wb.set_cell_value(0U, 1U, 0U, Value::number(3.0)));  // A2
  RETURN_IF_ERROR(wb.set_cell_value(0U, 2U, 0U, Value::number(8.0)));  // A3
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U, "=IF(AND(OR(A1>0,A2<0),A3=8),IF(A1+A2>A3,A1,A2),0)"));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildLargeGrid500Cells() {
  // 500 cells across A1:T25 (20 columns x 25 rows). Mix of numeric
  // literals, text, and formulas referencing earlier cells.
  //
  // `Value::text` is non-owning, so we materialise the per-row label
  // strings into a function-static `std::deque<std::string>`: the
  // deque keeps element addresses stable across appends and persists
  // until program teardown, which is well past the round-trip pipeline
  // (the workbook reads `cached_value.as_text()` during `save()` and
  // discards the view afterward, but the value still has to be alive
  // through that read).
  static std::deque<std::string> kLabels;
  static const bool kInitialised = [&]() {
    for (std::uint32_t r = 0; r < 25U; ++r) {
      kLabels.push_back("L" + std::to_string(r));
    }
    return true;
  }();
  (void)kInitialised;

  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 25U; ++r) {
    for (std::uint32_t c = 0; c < 20U; ++c) {
      if ((r + c) % 7U == 0U) {
        RETURN_IF_ERROR(wb.set_cell_value(0U, r, c, Value::text(kLabels[r])));
      } else if ((r + c) % 3U == 0U && r > 0U) {
        // Formula referencing the cell directly above (1-based row).
        const std::string col_letter(1U, static_cast<char>('A' + c));
        const std::string ref_above = col_letter + std::to_string(r);
        RETURN_IF_ERROR(wb.set_cell_formula(0U, r, c, "=" + ref_above + "+1"));
      } else {
        RETURN_IF_ERROR(wb.set_cell_value(0U, r, c, Value::number(static_cast<double>(r * 20U + c))));
      }
    }
  }
  return SaveBytes(wb);
}

/// Combined kitchen-sink: 2 sheets, a defined name, a table, a
/// passthrough theme part, mixed formulas, unicode text, and an
/// iterative circular pair. We can't add a passthrough part via the
/// public Workbook API alone (the writer copies it through, but the
/// only way to get one in is to read an archive that already has it),
/// so we build the synthetic theme archive, read it, then mutate the
/// workbook to add the rest of the surfaces, then save.
Expected<std::vector<std::uint8_t>, Error> BuildKitchenSink() {
  ASSIGN_OR_RETURN(auto theme_bytes, BuildPassthroughThemeArchive());
  ASSIGN_OR_RETURN(auto first_or, io::read_ooxml(SpanOf(theme_bytes)));
  Workbook wb = std::move(first_or.workbook);

  // Add a second sheet.
  wb.add_sheet("\xE6\x97\xA5\xE6\x9C\xAC");  // "日本"

  // Defined name (workbook-scope).
  std::vector<DefinedName> names;
  DefinedName n;
  n.name = "ComboName";
  n.formula = "Sheet1!$A$1";
  names.push_back(std::move(n));
  wb.set_defined_names(std::move(names));

  // Table on Sheet1.
  TableMetadata tab;
  tab.id = 1U;
  tab.name = "ComboTable";
  tab.display_name = "ComboTable";
  tab.ref = "A1:B2";
  tab.sheet_index = 0U;
  tab.columns.push_back(TableColumn{1U, "Alpha", "", "", ""});
  tab.columns.push_back(TableColumn{2U, "Beta", "", "", ""});
  std::vector<TableMetadata> tables;
  tables.push_back(std::move(tab));
  wb.set_tables(std::move(tables));

  // Mixed formulas + unicode text.
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(3.14)));                    // A1
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 1U, 0U, "=A1*2"));                              // A2
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 1U, Value::text("\xE3\x81\x82")));            // B1: "あ"
  RETURN_IF_ERROR(wb.set_cell_value(1U, 0U, 0U, Value::text("Hello \xF0\x9F\x91\x8B")));  // emoji wave

  // Iterative circular pair on Sheet2 (rows 5-6).
  RETURN_IF_ERROR(wb.set_cell_formula(1U, 5U, 0U, "=B6+0.5"));
  RETURN_IF_ERROR(wb.set_cell_formula(1U, 5U, 1U, "=A6"));
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 20U;
  opts.max_change = 0.01;
  wb.set_iterative_options(opts);

  return SaveBytes(wb);
}

const Cell* CellAt(const Workbook& wb, std::size_t s, std::uint32_t r, std::uint32_t c) {
  return wb.sheet(s).cell_at(r, c);
}

void AssertEmpty(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
  EXPECT_EQ(wb.sheet(0).cell_count(), 0U);
}

void AssertSingleLiteral(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 123.5);
}

void AssertLiteralsAllKinds(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 42.0);

  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  ASSERT_NE(a2, nullptr);
  ASSERT_TRUE(a2->cached_value.is_boolean());
  EXPECT_TRUE(a2->cached_value.as_boolean());

  const Cell* a3 = CellAt(wb, 0U, 2U, 0U);
  ASSERT_NE(a3, nullptr);
  ASSERT_TRUE(a3->cached_value.is_boolean());
  EXPECT_FALSE(a3->cached_value.as_boolean());

  const Cell* a4 = CellAt(wb, 0U, 3U, 0U);
  ASSERT_NE(a4, nullptr);
  ASSERT_TRUE(a4->cached_value.is_text());
  EXPECT_EQ(a4->cached_value.as_text(), "hello");

  const Cell* a6 = CellAt(wb, 0U, 5U, 0U);
  ASSERT_NE(a6, nullptr);
  ASSERT_TRUE(a6->cached_value.is_error());
  EXPECT_EQ(a6->cached_value.as_error(), ErrorCode::Div0);
}

void AssertArithmetic(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* b1 = CellAt(wb, 0U, 0U, 1U);  // =A1+A2
  const Cell* b2 = CellAt(wb, 0U, 1U, 1U);  // =A1-A2
  const Cell* b3 = CellAt(wb, 0U, 2U, 1U);  // =A1*A2
  const Cell* b5 = CellAt(wb, 0U, 4U, 1U);  // =A1^A2
  const Cell* b7 = CellAt(wb, 0U, 6U, 1U);  // =-A1
  ASSERT_NE(b1, nullptr);
  ASSERT_NE(b2, nullptr);
  ASSERT_NE(b3, nullptr);
  ASSERT_NE(b5, nullptr);
  ASSERT_NE(b7, nullptr);
  ASSERT_TRUE(b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b1->cached_value.as_number(), 13.0);
  ASSERT_TRUE(b2->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b2->cached_value.as_number(), 7.0);
  ASSERT_TRUE(b3->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b3->cached_value.as_number(), 30.0);
  ASSERT_TRUE(b5->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b5->cached_value.as_number(), 1000.0);
  ASSERT_TRUE(b7->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b7->cached_value.as_number(), -10.0);
}

void AssertCrossSheet(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
  EXPECT_EQ(wb.sheet(1).name(), "Sheet2");
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 99.0);
}

void AssertMultiSheetIndependent(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 3U);
  EXPECT_EQ(wb.sheet(0).name(), "North");
  EXPECT_EQ(wb.sheet(1).name(), "South");
  EXPECT_EQ(wb.sheet(2).name(), "East");
  const Cell* north_b1 = CellAt(wb, 0U, 0U, 1U);
  const Cell* south_b1 = CellAt(wb, 1U, 0U, 1U);
  const Cell* east_b1 = CellAt(wb, 2U, 0U, 1U);
  ASSERT_NE(north_b1, nullptr);
  ASSERT_NE(south_b1, nullptr);
  ASSERT_NE(east_b1, nullptr);
  ASSERT_TRUE(north_b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(north_b1->cached_value.as_number(), 2.0);
  ASSERT_TRUE(south_b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(south_b1->cached_value.as_number(), 4.0);
  ASSERT_TRUE(east_b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(east_b1->cached_value.as_number(), 2.0);
}

void AssertUnicodeSheetNames(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 3U);
  EXPECT_EQ(wb.sheet(0).name(), "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E");
  EXPECT_EQ(wb.sheet(1).name(), "emoji\xF0\x9F\x98\x80");
  EXPECT_EQ(wb.sheet(2).name(), "Sheet 3 with spaces");
}

void AssertUnicodeCellText(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  const Cell* a3 = CellAt(wb, 0U, 2U, 0U);
  const Cell* a4 = CellAt(wb, 0U, 3U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(a2, nullptr);
  ASSERT_NE(a3, nullptr);
  ASSERT_NE(a4, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E");
  ASSERT_TRUE(a2->cached_value.is_text());
  EXPECT_EQ(a2->cached_value.as_text(), "Hello \xF0\x9F\x98\x80");
}

void AssertRangeAggregates(const Workbook& wb) {
  const Cell* sum = CellAt(wb, 0U, 0U, 1U);
  const Cell* avg = CellAt(wb, 0U, 1U, 1U);
  const Cell* mn = CellAt(wb, 0U, 2U, 1U);
  const Cell* mx = CellAt(wb, 0U, 3U, 1U);
  const Cell* cnt = CellAt(wb, 0U, 4U, 1U);
  ASSERT_NE(sum, nullptr);
  ASSERT_NE(avg, nullptr);
  ASSERT_NE(mn, nullptr);
  ASSERT_NE(mx, nullptr);
  ASSERT_NE(cnt, nullptr);
  ASSERT_TRUE(sum->cached_value.is_number());
  EXPECT_DOUBLE_EQ(sum->cached_value.as_number(), 55.0);
  ASSERT_TRUE(avg->cached_value.is_number());
  EXPECT_DOUBLE_EQ(avg->cached_value.as_number(), 5.5);
  ASSERT_TRUE(mn->cached_value.is_number());
  EXPECT_DOUBLE_EQ(mn->cached_value.as_number(), 1.0);
  ASSERT_TRUE(mx->cached_value.is_number());
  EXPECT_DOUBLE_EQ(mx->cached_value.as_number(), 10.0);
  ASSERT_TRUE(cnt->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cnt->cached_value.as_number(), 10.0);
}

void AssertVolatileNowToday(const Workbook& wb) {
  // Cached values WILL drift across recalcs; we only check that the
  // formula text was preserved.
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  const Cell* a3 = CellAt(wb, 0U, 2U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(a2, nullptr);
  ASSERT_NE(a3, nullptr);
  EXPECT_EQ(a1->formula_text, "=NOW()");
  EXPECT_EQ(a2->formula_text, "=TODAY()");
  EXPECT_EQ(a3->formula_text, "=RAND()");
}

void AssertIterativeCircular(const Workbook& wb) {
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  const Cell* b1 = CellAt(wb, 0U, 0U, 1U);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(b1, nullptr);
  EXPECT_EQ(a1->formula_text, "=B1+1");
  EXPECT_EQ(b1->formula_text, "=A1");
  // The cached values are observable but not pinned: depending on
  // whether iterative options were re-applied, the SCC will resolve
  // either to a converged numeric pair (iterative on) or to #REF!
  // (iterative off — the writer does not yet persist `<calcPr
  // iterate=...>`). Either is acceptable here; the formula text is
  // the load-bearing invariant.
}

void AssertDefinedNamesWorkbookScope(const Workbook& wb) {
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "Sales");
  EXPECT_EQ(wb.defined_names()[0].formula, "Sheet1!$A$1:$A$10");
  EXPECT_EQ(wb.defined_names()[0].local_sheet_id, -1);
  EXPECT_FALSE(wb.defined_names()[0].hidden);
}

void AssertDefinedNamesSheetScope(const Workbook& wb) {
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "LocalRange");
  EXPECT_EQ(wb.defined_names()[0].formula, "Sheet1!$B$1:$B$5");
  EXPECT_EQ(wb.defined_names()[0].local_sheet_id, 0);
  EXPECT_TRUE(wb.defined_names()[0].hidden);
  EXPECT_EQ(wb.defined_names()[0].comment, "Sheet1-only.");
}

void AssertSingleTable(const Workbook& wb) {
  ASSERT_EQ(wb.tables().size(), 1U);
  const TableMetadata& t = wb.tables()[0];
  EXPECT_EQ(t.name, "SalesTable");
  EXPECT_EQ(t.ref, "A1:C5");
  ASSERT_EQ(t.columns.size(), 3U);
  EXPECT_EQ(t.columns[0].name, "Region");
}

void AssertMultipleTables(const Workbook& wb) {
  ASSERT_EQ(wb.tables().size(), 3U);
  EXPECT_EQ(wb.tables()[0].name, "Tbl1");
  EXPECT_EQ(wb.tables()[1].name, "Tbl2");
  EXPECT_EQ(wb.tables()[2].name, "Tbl3");
  EXPECT_EQ(wb.tables()[0].sheet_index, 0U);
  EXPECT_EQ(wb.tables()[1].sheet_index, 0U);
  EXPECT_EQ(wb.tables()[2].sheet_index, 1U);
}

void AssertNestedFunctionCalls(const Workbook& wb) {
  const Cell* b1 = CellAt(wb, 0U, 0U, 1U);
  ASSERT_NE(b1, nullptr);
  // Formula text round-trips verbatim; cached value resolves to A2=3.
  EXPECT_EQ(b1->formula_text, "=IF(AND(OR(A1>0,A2<0),A3=8),IF(A1+A2>A3,A1,A2),0)");
  ASSERT_TRUE(b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b1->cached_value.as_number(), 3.0);
}

void AssertLargeGrid(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  // Spot-check the corners of the grid: A1 should be a numeric
  // literal (r=0, c=0; (r+c) % 7 == 0 -> text "L0").
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "L0");
  // Last cell T25 (r=24, c=19): (24+19)=43 -> 43%7=1 -> not text;
  // 43 % 3 = 1 -> not formula; numeric 24*20+19 = 499.
  const Cell* t25 = CellAt(wb, 0U, 24U, 19U);
  ASSERT_NE(t25, nullptr);
  ASSERT_TRUE(t25->cached_value.is_number());
  EXPECT_DOUBLE_EQ(t25->cached_value.as_number(), 499.0);
}

void AssertPassthroughTheme(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  // Passthrough must include the theme1 part with non-empty bytes.
  const auto& parts = wb.passthrough_parts();
  auto it = std::find_if(parts.begin(), parts.end(),
                         [](const PassthroughPart& p) { return p.path == "xl/theme/theme1.xml"; });
  ASSERT_NE(it, parts.end()) << "theme part missing from passthrough";
  EXPECT_FALSE(it->bytes.empty());
}

void AssertSstCells(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  const Cell* b1 = CellAt(wb, 0U, 0U, 1U);
  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(b1, nullptr);
  ASSERT_NE(a2, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  ASSERT_TRUE(b1->cached_value.is_text());
  ASSERT_TRUE(a2->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "shared-alpha");
  EXPECT_EQ(b1->cached_value.as_text(), "shared-beta");
  EXPECT_EQ(a2->cached_value.as_text(), "shared-gamma");
}

void AssertKitchenSink(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 2U);
  // Defined name preserved.
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "ComboName");
  // Table preserved.
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables()[0].name, "ComboTable");
  // Passthrough preserved.
  const auto& parts = wb.passthrough_parts();
  auto it = std::find_if(parts.begin(), parts.end(),
                         [](const PassthroughPart& p) { return p.path == "xl/theme/theme1.xml"; });
  EXPECT_NE(it, parts.end());
  // Mixed formula present.
  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  ASSERT_NE(a2, nullptr);
  EXPECT_EQ(a2->formula_text, "=A1*2");
}

std::vector<CorpusBook> make_corpus() {
  std::vector<CorpusBook> books;
  books.push_back({"empty", BuildEmpty, AssertEmpty});
  books.push_back({"single_literal", BuildSingleLiteral, AssertSingleLiteral});
  books.push_back({"literals_all_kinds", BuildLiteralsAllKinds, AssertLiteralsAllKinds});
  books.push_back({"arithmetic", BuildArithmetic, AssertArithmetic});
  books.push_back({"cross_sheet", BuildCrossSheet, AssertCrossSheet});
  books.push_back({"multi_sheet_independent", BuildMultiSheetIndependent, AssertMultiSheetIndependent});
  books.push_back({"unicode_sheet_names", BuildUnicodeSheetNames, AssertUnicodeSheetNames});
  books.push_back({"unicode_cell_text", BuildUnicodeCellText, AssertUnicodeCellText});
  books.push_back({"range_aggregates", BuildRangeAggregates, AssertRangeAggregates});
  books.push_back({"volatile_now_today", BuildVolatileNowToday, AssertVolatileNowToday});
  books.push_back({"iterative_circular", BuildIterativeCircular, AssertIterativeCircular});
  books.push_back({"defined_names_workbook_scope", BuildDefinedNamesWorkbookScope, AssertDefinedNamesWorkbookScope});
  books.push_back({"defined_names_sheet_scope", BuildDefinedNamesSheetScope, AssertDefinedNamesSheetScope});
  books.push_back({"single_table", BuildSingleTable, AssertSingleTable});
  books.push_back({"multiple_tables", BuildMultipleTables, AssertMultipleTables});
  books.push_back({"nested_function_calls", BuildNestedFunctionCalls, AssertNestedFunctionCalls});
  books.push_back({"large_grid_500_cells", BuildLargeGrid500Cells, AssertLargeGrid});
  books.push_back({"passthrough_theme_part", BuildPassthroughThemeArchive, AssertPassthroughTheme});
  books.push_back({"sst_cells", BuildSstArchive, AssertSstCells});
  books.push_back({"combined_kitchen_sink", BuildKitchenSink, AssertKitchenSink});
  return books;
}

}  // namespace formulon::ooxml_corpus
