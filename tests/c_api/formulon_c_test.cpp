// Stable C ABI (`src/c_api/formulon_c.h`) end-to-end tests.
//
// The test driver is C++ for gtest convenience but everything it
// touches across the boundary is the pure-C surface declared in
// `formulon_c.h`.

#include <cstddef>
#include <string>

#include "c_api/parts/common.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "io/auto_filter_xml.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "workbook.h"

static_assert(offsetof(fm_parallel_recalc_stats, cells_evaluated) == 0U);
static_assert(offsetof(fm_parallel_recalc_stats, sccs_processed) == 8U);
static_assert(offsetof(fm_parallel_recalc_stats, parallel_steps) == 16U);
static_assert(offsetof(fm_parallel_recalc_stats, serial_fallback_steps) == 24U);
static_assert(offsetof(fm_parallel_recalc_stats, cycle_recoveries) == 32U);
static_assert(offsetof(fm_parallel_recalc_stats, worker_threads_started) == 40U);
static_assert(offsetof(fm_parallel_recalc_stats, worker_threads_used) == 44U);

// The two counter structs must have the same layout on native and wasm32,
// which is the whole reason they use `uint32_t` rather than `size_t`. A
// binding's hand-written offsets are only safe while this holds.
static_assert(sizeof(fm_read_diagnostics_t) == 20U);
static_assert(offsetof(fm_read_diagnostics_t, undecoded_formula_count) == 0U);
static_assert(offsetof(fm_read_diagnostics_t, undecoded_defined_name_count) == 4U);
static_assert(offsetof(fm_read_diagnostics_t, undecoded_part_count) == 8U);
static_assert(offsetof(fm_read_diagnostics_t, skipped_feature_count) == 12U);
static_assert(offsetof(fm_read_diagnostics_t, unknown_content_type_count) == 16U);
static_assert(sizeof(fm_save_diagnostics_t) == 20U);
static_assert(offsetof(fm_save_diagnostics_t, downgraded_formula_count) == 0U);
static_assert(offsetof(fm_save_diagnostics_t, deferred_feature_count) == 4U);
static_assert(offsetof(fm_save_diagnostics_t, dropped_part_count) == 8U);
static_assert(offsetof(fm_save_diagnostics_t, dropped_relationship_count) == 12U);
static_assert(offsetof(fm_save_diagnostics_t, renumbered_part_count) == 16U);
static_assert(sizeof(fm_parallel_recalc_stats) == 48U);

// `fm_value_t` is the most widely passed record on the boundary: every cell
// read, every ad-hoc evaluation and every pivot cell writes one through a
// caller-supplied block. The union's `double` fixes the alignment at 8, so the
// discriminator's four bytes of tail padding are part of the layout rather than
// an implementation detail, and both host bindings decode the payload from the
// resulting offset 8. Identical on native and wasm32 because the widest union
// member is the `double` on both.
static_assert(sizeof(fm_value_t) == 16U, "fm_value_t ABI layout changed");
static_assert(alignof(fm_value_t) == 8U, "fm_value_t ABI alignment changed");
static_assert(offsetof(fm_value_t, kind) == 0U, "fm_value_t.kind offset changed");
static_assert(offsetof(fm_value_t, u) == 8U, "fm_value_t.u offset changed");
static_assert(sizeof(decltype(fm_value_t::u)) == 8U, "fm_value_t.u payload width changed");

// `fm_print_range_t` is written through a caller-supplied block by
// `fm_pagination_print_area_at`. Four `uint32_t` with no padding, so it is
// identical on native and wasm32 and a binding may decode it as four
// little-endian words.
static_assert(sizeof(fm_print_range_t) == 16U, "fm_print_range_t ABI layout changed");
static_assert(alignof(fm_print_range_t) == 4U, "fm_print_range_t ABI alignment changed");
static_assert(offsetof(fm_print_range_t, first_row) == 0U, "fm_print_range_t.first_row offset changed");
static_assert(offsetof(fm_print_range_t, first_col) == 4U, "fm_print_range_t.first_col offset changed");
static_assert(offsetof(fm_print_range_t, last_row) == 8U, "fm_print_range_t.last_row offset changed");
static_assert(offsetof(fm_print_range_t, last_col) == 12U, "fm_print_range_t.last_col offset changed");

TEST(FormulonCApi, TableCreateUpdateRemoveRoundTripsThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* columns[] = {"Product", "Amount"};
  size_t index = 99;
  ASSERT_EQ(
      fm_workbook_table_create(wb.handle, 0, "A1:B3", "Sales", "Sales", columns, 2, "TableStyleMedium2", 1, 0, &index),
      0);
  EXPECT_EQ(index, 0U);
  EXPECT_EQ(fm_workbook_table_count(wb.handle), 1U);

  const char* name = nullptr;
  const char* display_name = nullptr;
  const char* ref = nullptr;
  size_t sheet = 99;
  ASSERT_EQ(fm_workbook_table_at(wb.handle, index, &name, &display_name, &ref, &sheet), 0);
  EXPECT_STREQ(name, "Sales");
  EXPECT_STREQ(display_name, "Sales");
  EXPECT_STREQ(ref, "A1:B3");
  EXPECT_EQ(sheet, 0U);

  ASSERT_EQ(fm_workbook_table_update(wb.handle, index, "A1:B4", "TableStyleLight9", 1, 1), 0);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  ASSERT_EQ(fm_workbook_table_at(loaded.handle, 0, &name, &display_name, &ref, &sheet), 0);
  EXPECT_STREQ(ref, "A1:B4");

  ASSERT_EQ(fm_workbook_table_remove(loaded.handle, 0), 0);
  EXPECT_EQ(fm_workbook_table_count(loaded.handle), 0U);
}
TEST(FormulonCApi, TableMutationsReindexStructuredRefDependents) {
  // A `StructuredRef` resolves to a static rectangle once, at formula
  // registration time (`eval/dep_extractor.cpp`'s `walk_structured_ref`),
  // and does not re-resolve on its own; table create/update/remove must
  // reindex the formulas that name the table or their dep-graph edges and
  // cached values go stale.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "Product"), 0);                 // A1
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 1, "Amount"), 0);                  // B1
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 1, 0, "Widget"), 0);                  // A2
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 1, 10.0), 0);                    // B2
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 2, 0, "Gadget"), 0);                  // A3
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 2, 1, 20.0), 0);                    // B3
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 3, "=SUM(Sales[Amount])"), 0);  // D1

  const char* columns[] = {"Product", "Amount"};
  size_t index = 99;
  ASSERT_EQ(fm_workbook_table_create(wb.handle, 0, "A1:B3", "Sales", "Sales", columns, 2, "", 1, 0, &index), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 3, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 30.0);

  // Widen the table to include a new data row and update it; the formula's
  // pinned rectangle must be re-derived from the new `ref`, not left
  // pointing at the pre-update B2:B3.
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 3, 1, 5.0), 0);  // B4
  ASSERT_EQ(fm_workbook_table_update(wb.handle, index, "A1:B4", "", 1, 0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 3, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 35.0) << "table_update must reindex the formula onto the widened range";

  // Removing the table invalidates the reference outright; the stale 35.0
  // must not survive as a silently-uncomputed cached value.
  ASSERT_EQ(fm_workbook_table_remove(wb.handle, index), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 3, &v), 0);
  EXPECT_NE(v.kind, FM_VAL_NUMBER)
      << "table_remove must reindex the formula so it re-evaluates against a missing table";
}
TEST(FormulonCApi, TableRangeMustMatchTheColumnList) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* columns[] = {"Product", "Amount"};
  size_t index = 99;

  // Two headers cannot describe a three-column range, and a range that is
  // not a plain A1 area has no width to check against at all.
  EXPECT_NE(fm_workbook_table_create(wb.handle, 0, "A1:C3", "Sales", "Sales", columns, 2, "", 1, 0, &index), 0);
  EXPECT_NE(fm_workbook_table_create(wb.handle, 0, "Sheet1!A1:B3", "Sales", "Sales", columns, 2, "", 1, 0, &index), 0);
  EXPECT_NE(fm_workbook_table_create(wb.handle, 0, "$A$1:$B$3", "Sales", "Sales", columns, 2, "", 1, 0, &index), 0);
  EXPECT_EQ(fm_workbook_table_count(wb.handle), 0U);

  ASSERT_EQ(fm_workbook_table_create(wb.handle, 0, "A1:B3", "Sales", "Sales", columns, 2, "", 1, 0, &index), 0);
  // Growing rows is fine; growing columns would orphan the column list.
  EXPECT_EQ(fm_workbook_table_update(wb.handle, index, "A1:B9", "", 1, 0), 0);
  EXPECT_NE(fm_workbook_table_update(wb.handle, index, "A1:C9", "", 1, 0), 0);

  const char* duplicate_columns[] = {"Product", "product"};
  EXPECT_NE(fm_workbook_table_create(wb.handle, 0, "D1:E3", "Other", "Other", duplicate_columns, 2, "", 1, 0, &index),
            0);
}
TEST(FormulonCApi, TableUpdatePreservesRawMetadataAcrossSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* columns[] = {"Product", "Amount"};
  size_t index = 99;
  ASSERT_EQ(fm_workbook_table_create(wb.handle, 0, "A1:B3", "Sales", "Sales", columns, 2, "CustomStyle", 1, 1, &index),
            0);

  // Seed payload that the C ABI deliberately does not model. The update must
  // rewrite only the opening autoFilter ref and retain every raw payload.
  auto& table = wb.handle->wb->mutable_tables()[index];
  table.auto_filter_xml.set(formulon::io::auto_filter_from_xml(
      "<autoFilter ref=\"A1:B3\"><filterColumn colId=\"0\"><filters><filter val=\"West\"/></filters></filterColumn>"
      "</autoFilter>"));
  table.sort_state_xml = "<sortState ref=\"A1:B3\"><sortCondition ref=\"B2:B3\" descending=\"1\"/></sortState>";
  table.table_style_info_xml = "<tableStyleInfo name=\"CustomStyle\" showRowStripes=\"0\"/>";
  table.ext_lst_xml = "<extLst><ext uri=\"urn:formulon:test\"><futureTableData value=\"kept\"/></ext></extLst>";

  ASSERT_EQ(fm_workbook_table_update(wb.handle, index, "A1:B4", nullptr, -1, -1), 0);
  EXPECT_EQ(table.ref, "A1:B4");
  EXPECT_TRUE(table.header_row);
  EXPECT_TRUE(table.totals_row);
  EXPECT_NE(formulon::io::auto_filter_xml(table.auto_filter_xml.get()).find("ref=\"A1:B4\""), std::string::npos);
  EXPECT_NE(formulon::io::auto_filter_xml(table.auto_filter_xml.get()).find("filterColumn"), std::string::npos);
  EXPECT_NE(table.sort_state_xml.find("ref=\"A1:B3\""), std::string::npos);
  EXPECT_NE(table.sort_state_xml.find("sortCondition"), std::string::npos);
  EXPECT_EQ(table.table_style_info_xml, "<tableStyleInfo name=\"CustomStyle\" showRowStripes=\"0\"/>");
  EXPECT_EQ(table.ext_lst_xml,
            "<extLst><ext uri=\"urn:formulon:test\"><futureTableData value=\"kept\"/></ext></extLst>");

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  const auto& reloaded = loaded.handle->wb->tables()[0];
  EXPECT_EQ(reloaded.ref, "A1:B4");
  EXPECT_TRUE(reloaded.header_row);
  EXPECT_TRUE(reloaded.totals_row);
  EXPECT_NE(formulon::io::auto_filter_xml(reloaded.auto_filter_xml.get()).find("ref=\"A1:B4\""), std::string::npos);
  EXPECT_NE(formulon::io::auto_filter_xml(reloaded.auto_filter_xml.get()).find("filterColumn"), std::string::npos);
  EXPECT_NE(formulon::io::auto_filter_xml(reloaded.auto_filter_xml.get()).find("West"), std::string::npos);
  EXPECT_EQ(reloaded.sort_state_xml,
            "<sortState ref=\"A1:B3\"><sortCondition ref=\"B2:B3\" descending=\"1\"/></sortState>");
  EXPECT_EQ(reloaded.table_style_info_xml, "<tableStyleInfo name=\"CustomStyle\" showRowStripes=\"0\"/>");
  EXPECT_EQ(reloaded.ext_lst_xml,
            "<extLst><ext uri=\"urn:formulon:test\"><futureTableData value=\"kept\"/></ext></extLst>");

  formulon::io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(formulon::io::ByteSpan{saved.data, saved.len})));
  auto table_part_or = zip.read_entry("xl/tables/table1.xml");
  ASSERT_TRUE(static_cast<bool>(table_part_or)) << table_part_or.error().message;
  const std::string table_xml(table_part_or.value().begin(), table_part_or.value().end());
  const std::size_t auto_filter_pos = table_xml.find("<autoFilter");
  const std::size_t sort_state_pos = table_xml.find("<sortState");
  const std::size_t table_columns_pos = table_xml.find("<tableColumns");
  const std::size_t style_info_pos = table_xml.find("<tableStyleInfo");
  const std::size_t ext_lst_pos = table_xml.find("<extLst");
  ASSERT_NE(auto_filter_pos, std::string::npos);
  ASSERT_NE(sort_state_pos, std::string::npos);
  ASSERT_NE(table_columns_pos, std::string::npos);
  ASSERT_NE(style_info_pos, std::string::npos);
  ASSERT_NE(ext_lst_pos, std::string::npos);
  EXPECT_LT(auto_filter_pos, sort_state_pos);
  EXPECT_LT(sort_state_pos, table_columns_pos);
  EXPECT_LT(table_columns_pos, style_info_pos);
  EXPECT_LT(style_info_pos, ext_lst_pos);
  EXPECT_NE(table_xml.find("ref=\"A1:B4\""), std::string::npos);
  EXPECT_NE(table_xml.find("filterColumn"), std::string::npos);
  EXPECT_NE(table_xml.find("sortCondition ref=\"B2:B3\""), std::string::npos);
  EXPECT_NE(table_xml.find("futureTableData value=\"kept\""), std::string::npos);
}
TEST(FormulonCApi, PaginationSnapshotExposesBreaksAndUsedRangeFallback) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 199, 0, 2.0), 0);

  fm_pagination_t* pagination = nullptr;
  ASSERT_EQ(fm_workbook_paginate(wb.handle, 0, &pagination), 0);
  ASSERT_NE(pagination, nullptr);
  // 200 default-height rows in one column: the row axis breaks, the column
  // axis does not. The exact break row follows the geometry model and is
  // pinned by the workbook oracle, so this asserts the ABI's shape -- a
  // reported count that matches what the indexed accessor will hand back,
  // and an out-of-range index that fails rather than reading past the end.
  const auto breaks = static_cast<std::uint32_t>(fm_pagination_horizontal_break_count(pagination));
  ASSERT_GE(breaks, 1U);
  EXPECT_EQ(fm_pagination_page_count(pagination), breaks + 1U);
  // No explicit _xlnm.Print_Area is reported even though pagination falls
  // back to the used range internally.
  EXPECT_EQ(fm_pagination_print_area_count(pagination), 0U);
  std::uint32_t row = 0;
  std::uint32_t previous = 0;
  for (std::uint32_t i = 0; i < breaks; ++i) {
    ASSERT_EQ(fm_pagination_horizontal_break_at(pagination, i, &row), 0);
    EXPECT_GT(row, previous);
    EXPECT_LE(row, 199U);
    previous = row;
  }
  EXPECT_EQ(fm_pagination_vertical_break_count(pagination), 0U);
  EXPECT_NE(fm_pagination_horizontal_break_at(pagination, breaks, &row), 0);
  fm_pagination_destroy(pagination);

  EXPECT_NE(fm_workbook_paginate(nullptr, 0, &pagination), 0);
  EXPECT_NE(fm_workbook_paginate(wb.handle, 0, nullptr), 0);
}
TEST(FormulonCApi, CellPhoneticCanBeReadClearedAndRejectsInvalidArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "漢字"), 0);
  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "かんじ"), 0);

  const char* phonetic = nullptr;
  ASSERT_EQ(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, &phonetic), 0);
  ASSERT_NE(phonetic, nullptr);
  EXPECT_STREQ(phonetic, "かんじ");

  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, ""), 0);
  ASSERT_EQ(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, &phonetic), 0);
  ASSERT_NE(phonetic, nullptr);
  EXPECT_STREQ(phonetic, "");

  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "かんじ"), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "文字列"), 0);
  ASSERT_EQ(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, &phonetic), 0);
  ASSERT_NE(phonetic, nullptr);
  EXPECT_STREQ(phonetic, "");

  EXPECT_NE(fm_workbook_set_cell_phonetic(nullptr, 0, 0, 0, "x"), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, nullptr), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic(wb.handle, 0, formulon::Sheet::kMaxRows, 0, "x"), 0);
  EXPECT_NE(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, nullptr), 0);
}
TEST(FormulonCApi, CellPhoneticSettersDirtyPhoneticFormulaDependents) {
  // `PHONETIC(A1)` depends on A1 through the same reference edge any other
  // function argument creates, so changing what A1's reading is (without
  // changing A1's own text) must still re-dirty that dependent -- recalc
  // alone is dirty-only and cannot discover the change on its own.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "win-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "漢字"), 0);              // A1
  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "かんじ"), 0);   // A1 reading
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 1, "=PHONETIC(A1)"), 0);  // B1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 1, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_TEXT);
  EXPECT_STREQ(v.u.text, "かんじ");

  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "べつのよみ"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 1, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_TEXT);
  EXPECT_STREQ(v.u.text, "べつのよみ") << "changing A1's phonetic reading must dirty PHONETIC(A1)";

  // The runs and properties setters are the same seam and must dirty too.
  const fm_phonetic_run_t run{0U, 3U, "らん"};
  ASSERT_EQ(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, &run, 1U), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 1, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_TEXT);
  EXPECT_STREQ(v.u.text, "らん") << "changing A1's phonetic runs must dirty PHONETIC(A1)";
}
TEST(FormulonCApi, CellPhoneticRunsPreserveTheirSpans) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "東京都"), 0);

  const fm_phonetic_run_t runs[] = {{0U, 2U, "トウキョウ"}, {2U, 3U, "ト"}};
  ASSERT_EQ(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, runs, 2U), 0);

  uint32_t count = 0;
  ASSERT_EQ(fm_workbook_get_cell_phonetic_run_count(wb.handle, 0, 0, 0, &count), 0);
  EXPECT_EQ(count, 2U);

  fm_phonetic_run_t read{};
  ASSERT_EQ(fm_workbook_get_cell_phonetic_run(wb.handle, 0, 0, 0, 0U, &read), 0);
  EXPECT_EQ(read.sb, 0U);
  EXPECT_EQ(read.eb, 2U);
  EXPECT_STREQ(read.text, "トウキョウ");
  ASSERT_EQ(fm_workbook_get_cell_phonetic_run(wb.handle, 0, 0, 0, 1U, &read), 0);
  EXPECT_EQ(read.sb, 2U);
  EXPECT_EQ(read.eb, 3U);
  EXPECT_STREQ(read.text, "ト");

  // The flattening getter still reports the concatenation, and feeding that
  // back through the whole-cell setter is exactly the collapse the run API
  // exists to avoid.
  const char* flattened = nullptr;
  ASSERT_EQ(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, &flattened), 0);
  EXPECT_STREQ(flattened, "トウキョウト");
  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "トウキョウト"), 0);
  ASSERT_EQ(fm_workbook_get_cell_phonetic_run_count(wb.handle, 0, 0, 0, &count), 0);
  EXPECT_EQ(count, 1U);

  // An empty batch clears; a cell that was never annotated reports zero runs.
  ASSERT_EQ(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, nullptr, 0U), 0);
  ASSERT_EQ(fm_workbook_get_cell_phonetic_run_count(wb.handle, 0, 0, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  ASSERT_EQ(fm_workbook_get_cell_phonetic_run_count(wb.handle, 0, 5, 5, &count), 0);
  EXPECT_EQ(count, 0U);
}
TEST(FormulonCApi, CellPhoneticRunsRejectMalformedInput) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "東京都"), 0);

  const fm_phonetic_run_t backwards[] = {{2U, 1U, "ト"}};
  EXPECT_NE(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, backwards, 1U), 0);
  const fm_phonetic_run_t overlapping[] = {{0U, 2U, "トウキョウ"}, {1U, 3U, "ト"}};
  EXPECT_NE(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, overlapping, 2U), 0);
  const fm_phonetic_run_t null_text[] = {{0U, 2U, nullptr}};
  EXPECT_NE(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, null_text, 1U), 0);

  const fm_phonetic_run_t ok[] = {{0U, 3U, "トウキョウト"}};
  EXPECT_NE(fm_workbook_set_cell_phonetic_runs(nullptr, 0, 0, 0, ok, 1U), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, nullptr, 1U), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, formulon::Sheet::kMaxRows, 0, ok, 1U), 0);

  // A rejected batch leaves the previous annotation intact.
  ASSERT_EQ(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, ok, 1U), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, overlapping, 2U), 0);
  uint32_t count = 0;
  ASSERT_EQ(fm_workbook_get_cell_phonetic_run_count(wb.handle, 0, 0, 0, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_phonetic_run_t read{};
  EXPECT_NE(fm_workbook_get_cell_phonetic_run(wb.handle, 0, 0, 0, 1U, &read), 0);
  EXPECT_NE(fm_workbook_get_cell_phonetic_run(wb.handle, 0, 0, 0, 0U, nullptr), 0);
  EXPECT_NE(fm_workbook_get_cell_phonetic_run_count(wb.handle, 0, 0, 0, nullptr), 0);
}
TEST(FormulonCApi, CellPhoneticPropertiesAreIndependentOfTheRuns) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "大阪"), 0);

  // An unannotated cell reads the state Excel infers for a guide written
  // with no `<phoneticPr>` at all.
  uint32_t font_id = 9;
  uint32_t type = 9;
  uint32_t alignment = 9;
  ASSERT_EQ(fm_workbook_get_cell_phonetic_properties(wb.handle, 0, 0, 0, &font_id, &type, &alignment), 0);
  EXPECT_EQ(font_id, 0U);
  EXPECT_EQ(type, FM_PHONETIC_TYPE_HALFWIDTH_KATAKANA);
  EXPECT_EQ(alignment, FM_PHONETIC_ALIGNMENT_NO_CONTROL);

  ASSERT_EQ(fm_workbook_set_cell_phonetic_properties(wb.handle, 0, 0, 0, 3U, FM_PHONETIC_TYPE_HIRAGANA,
                                                     FM_PHONETIC_ALIGNMENT_CENTER),
            0);
  // Writing the readings must not reset the rendering, which is the whole
  // reason the two entry points are separate.
  const fm_phonetic_run_t runs[] = {{0U, 2U, "おおさか"}};
  ASSERT_EQ(fm_workbook_set_cell_phonetic_runs(wb.handle, 0, 0, 0, runs, 1U), 0);
  ASSERT_EQ(fm_workbook_get_cell_phonetic_properties(wb.handle, 0, 0, 0, &font_id, &type, &alignment), 0);
  EXPECT_EQ(font_id, 3U);
  EXPECT_EQ(type, FM_PHONETIC_TYPE_HIRAGANA);
  EXPECT_EQ(alignment, FM_PHONETIC_ALIGNMENT_CENTER);

  // Each out-parameter is optional.
  uint32_t only_type = 0;
  ASSERT_EQ(fm_workbook_get_cell_phonetic_properties(wb.handle, 0, 0, 0, nullptr, &only_type, nullptr), 0);
  EXPECT_EQ(only_type, FM_PHONETIC_TYPE_HIRAGANA);

  // Overwriting the value discards the guide and its rendering together.
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "京都"), 0);
  ASSERT_EQ(fm_workbook_get_cell_phonetic_properties(wb.handle, 0, 0, 0, &font_id, &type, &alignment), 0);
  EXPECT_EQ(font_id, 0U);
  EXPECT_EQ(type, 0U);
  EXPECT_EQ(alignment, 0U);
}
TEST(FormulonCApi, CellPhoneticPropertiesRejectValuesTheContainerCannotHold) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // type and alignment are two bits each in the xlsb trailer, and font_id a
  // u16; a wider value would be truncated into a different, valid-looking
  // one on save.
  EXPECT_NE(fm_workbook_set_cell_phonetic_properties(wb.handle, 0, 0, 0, 0U, 4U, 0U), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic_properties(wb.handle, 0, 0, 0, 0U, 0U, 4U), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic_properties(wb.handle, 0, 0, 0, 0x10000U, 0U, 0U), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic_properties(wb.handle, 0, formulon::Sheet::kMaxRows, 0, 0U, 0U, 0U), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic_properties(wb.handle, 1U, 0, 0, 0U, 0U, 0U), 0);
  EXPECT_NE(fm_workbook_set_cell_phonetic_properties(nullptr, 0, 0, 0, 0U, 0U, 0U), 0);
  EXPECT_NE(fm_workbook_get_cell_phonetic_properties(nullptr, 0, 0, 0, nullptr, nullptr, nullptr), 0);

  // A rejected call leaves the stored rendering alone.
  ASSERT_EQ(fm_workbook_set_cell_phonetic_properties(wb.handle, 0, 0, 0, 1U, FM_PHONETIC_TYPE_NO_CONVERSION,
                                                     FM_PHONETIC_ALIGNMENT_DISTRIBUTED),
            0);
  EXPECT_NE(fm_workbook_set_cell_phonetic_properties(wb.handle, 0, 0, 0, 1U, 7U, 0U), 0);
  uint32_t type = 0;
  uint32_t alignment = 0;
  ASSERT_EQ(fm_workbook_get_cell_phonetic_properties(wb.handle, 0, 0, 0, nullptr, &type, &alignment), 0);
  EXPECT_EQ(type, FM_PHONETIC_TYPE_NO_CONVERSION);
  EXPECT_EQ(alignment, FM_PHONETIC_ALIGNMENT_DISTRIBUTED);
}
