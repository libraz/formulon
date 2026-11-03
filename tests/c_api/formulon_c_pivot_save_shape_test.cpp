// A pivot assembled through the C ABI (cache fields, inline record values,
// fields placed on axes without items) must save as a package Excel opens: every
// axis field lists its items and every discrete cache field carries shared items.

#include <string>
#include <vector>

#include "formulon_c_pivot_test_helpers.h"

namespace {

struct SourceRow {
  const char* region;  // nullptr is a blank cell.
  const char* product;
  double qty;
  double amt;
};

constexpr SourceRow kRows[] = {{"East", "Pen", 3, 30},  {"East", "Ink", 5, 100}, {"West", "Pen", 4, 40},
                               {"West", "Pen", 6, 60},  {"West", "Ink", 2, 50},  {"North", "Pen", 1, 10},
                               {nullptr, "Ink", 7, 70}, {"East", "Pen", 2, 20}};

void SetText(fm_workbook_t* wb, std::uint32_t cache_id, std::size_t record, std::size_t field, const char* text) {
  if (text == nullptr) {
    ASSERT_EQ(fm_workbook_pivot_cache_record_set_blank(wb, cache_id, record, field), 0) << fm_last_error_message();
  } else {
    ASSERT_EQ(fm_workbook_pivot_cache_record_set_text(wb, cache_id, record, field, text), 0) << fm_last_error_message();
  }
}

// The sequence a host drives: cache fields and inline records, one pivot field
// per cache field, row / column orders, then two data fields.
void BuildHostPivot(fm_workbook_t* wb) {
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_create(wb, 0, &cache_id), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb, cache_id, 1, "A1:D9", "Sheet1", nullptr), 0);
  const char* const headers[] = {"Region", "Product", "Qty", "Amt"};
  for (const char* header : headers) {
    std::size_t idx = 99;
    ASSERT_EQ(fm_workbook_pivot_cache_field_add(wb, cache_id, header, &idx), 0) << fm_last_error_message();
  }
  for (const SourceRow& row : kRows) {
    std::size_t record = 99;
    ASSERT_EQ(fm_workbook_pivot_cache_record_add(wb, cache_id, &record), 0);
    SetText(wb, cache_id, record, 0, row.region);
    SetText(wb, cache_id, record, 1, row.product);
    ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb, cache_id, record, 2, row.qty), 0);
    ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb, cache_id, record, 3, row.amt), 0);
  }
  std::size_t pivot = 99;
  ASSERT_EQ(fm_workbook_pivot_create(wb, 0, "PT1", cache_id, 2U, 5U, &pivot), 0) << fm_last_error_message();
  const fm_pivot_axis_t axes[] = {FM_PIVOT_AXIS_ROW, FM_PIVOT_AXIS_COL, FM_PIVOT_AXIS_VALUE, FM_PIVOT_AXIS_VALUE};
  for (std::size_t i = 0; i < 4; ++i) {
    fm_pivot_field_spec_t spec{};
    spec.source_name = headers[i];
    spec.custom_name = "";
    spec.axis = axes[i];
    spec.subtotal_top = 1;
    spec.number_format = "";
    std::size_t field = 99;
    ASSERT_EQ(fm_workbook_pivot_field_add(wb, 0, pivot, &spec, &field), 0) << fm_last_error_message();
    ASSERT_EQ(field, i);
  }
  const std::uint32_t rows[] = {0U};
  const std::uint32_t cols[] = {1U};
  ASSERT_EQ(fm_workbook_pivot_set_row_field_order(wb, 0, pivot, rows, 1U), 0);
  ASSERT_EQ(fm_workbook_pivot_set_col_field_order(wb, 0, pivot, cols, 1U), 0);
  const struct {
    const char* name;
    std::uint32_t field;
  } values[] = {{"Amt sum", 3U}, {"Qty sum", 2U}};
  for (const auto& value : values) {
    fm_pivot_data_field_spec_t df{};
    df.name = value.name;
    df.field_index = value.field;
    df.aggregation = FM_PIVOT_AGG_SUM;
    df.number_format = "";
    df.show_as = FM_PIVOT_SHOW_AS_NORMAL;
    df.show_as_base_field = -1;
    df.show_as_base_item = -1;
    std::size_t idx = 99;
    ASSERT_EQ(fm_workbook_pivot_data_field_add(wb, 0, pivot, &df, &idx), 0) << fm_last_error_message();
  }
  ASSERT_EQ(fm_workbook_pivot_set_layout(wb, 0, pivot, FM_PIVOT_LAYOUT_COMPACT), 0);
}

struct Cell {
  std::uint32_t row;
  std::uint32_t col;
  fm_pivot_cell_kind_t kind;
  fm_value_kind_t value_kind;
  std::string text;
  double number;
  bool operator==(const Cell& o) const {
    return row == o.row && col == o.col && kind == o.kind && value_kind == o.value_kind && text == o.text &&
           number == o.number;
  }
};

std::vector<Cell> ProjectedCells(fm_workbook_t* wb) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_workbook_pivot_layout(wb, 0, 0, &projected.handle), 0) << fm_last_error_message();
  std::vector<Cell> out;
  for (const fm_pivot_cell_t& c : CollectCells(projected.handle)) {
    Cell cell{c.row, c.col, c.kind, c.value.kind, {}, 0.0};
    if (c.value.kind == FM_VAL_TEXT && c.value.u.text != nullptr) {
      cell.text = c.value.u.text;
    } else if (c.value.kind == FM_VAL_NUMBER) {
      cell.number = c.value.u.number;
    }
    out.push_back(std::move(cell));
  }
  return out;
}

}  // namespace

TEST(FormulonCApiPivotSaveShape, HostBuiltPivotSavesWithItemsAndSharedItems) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  BuildHostPivot(wb.handle);
  const std::vector<Cell> before = ProjectedCells(wb.handle);
  ASSERT_FALSE(before.empty());

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string table = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  const std::string def = ExtractZipEntry(package, "xl/pivotCache/pivotCacheDefinition1.xml");
  const std::string rec = ExtractZipEntry(package, "xl/pivotCache/pivotCacheRecords1.xml");

  // Region and Product sit on axes: each lists one item per shared item plus the default subtotal.
  // Region has East, West, North and blank; Product has Pen and Ink.
  EXPECT_EQ(CountOccurrences(table, "<items count=\"5\">"), 1U) << table;
  EXPECT_EQ(CountOccurrences(table, "<items count=\"3\">"), 1U) << table;
  EXPECT_EQ(CountOccurrences(table, "<items"), 2U) << table;
  // East, North, West ascending, blank last; Ink before Pen.
  EXPECT_NE(table.find("<item x=\"0\"/><item x=\"2\"/><item x=\"1\"/><item x=\"3\"/><item t=\"default\"/>"),
            std::string::npos)
      << table;
  EXPECT_NE(table.find("<item x=\"1\"/><item x=\"0\"/><item t=\"default\"/>"), std::string::npos) << table;

  // Discrete fields carry shared items, never the numeric placeholder.
  EXPECT_NE(def.find("<cacheField name=\"Region\"><sharedItems containsBlank=\"1\" count=\"4\"><s v=\"East\"/>"
                     "<s v=\"West\"/><s v=\"North\"/><m/></sharedItems>"),
            std::string::npos)
      << def;
  EXPECT_NE(def.find("<cacheField name=\"Product\"><sharedItems count=\"2\"><s v=\"Pen\"/><s v=\"Ink\"/>"),
            std::string::npos)
      << def;
  EXPECT_EQ(CountOccurrences(def, "containsNumber=\"1\"/>"), 2U) << def;  // Qty and Amt keep the range form.

  // Records reference the derived fields by index and hold no inline text for them.
  EXPECT_EQ(rec.find("<s v="), std::string::npos) << rec;
  EXPECT_EQ(rec.find("<m/>"), std::string::npos) << rec;
  EXPECT_NE(rec.find("<r><x v=\"0\"/><x v=\"0\"/><n v=\"3\"/><n v=\"30\"/></r>"), std::string::npos) << rec;
  EXPECT_NE(rec.find("<r><x v=\"3\"/><x v=\"1\"/><n v=\"7\"/><n v=\"70\"/></r>"), std::string::npos) << rec;

  // Reloading evaluates to the report the host saw.
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0) << fm_last_error_message();
  EXPECT_EQ(ProjectedCells(loaded.handle), before);
}
