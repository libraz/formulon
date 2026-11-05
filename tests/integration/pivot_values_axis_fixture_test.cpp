//
// Workbooks Mac Excel 365 saved with the Values level among the row fields
// (two data fields, "Values on rows"). Through the C API the row field order
// must list `FM_PIVOT_VALUES_FIELD` where Excel placed it, moving it to the
// columns must give the default layout, and a save must spell the placement
// as `dataOnRows` and `<field x="-2"/>`.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "miniz.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace {

struct WorkbookHandle {
  fm_workbook_t* handle = nullptr;
  WorkbookHandle() = default;
  WorkbookHandle(const WorkbookHandle&) = delete;
  WorkbookHandle& operator=(const WorkbookHandle&) = delete;
  ~WorkbookHandle() { fm_workbook_destroy(handle); }
};

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

std::string ExtractEntry(const std::uint8_t* bytes, std::size_t len, const char* name) {
  mz_zip_archive zip{};
  if (mz_zip_reader_init_mem(&zip, bytes, len, 0) == MZ_FALSE) {
    return {};
  }
  std::size_t size = 0;
  void* data = mz_zip_reader_extract_file_to_heap(&zip, name, &size, 0);
  std::string out = data != nullptr ? std::string(static_cast<const char*>(data), size) : std::string();
  mz_free(data);
  mz_zip_reader_end(&zip);
  return out;
}

// The text between `<tag ...>` and `</tag>`, or empty.
std::string ElementBody(const std::string& xml, const std::string& tag) {
  const std::size_t open = xml.find("<" + tag + " ");
  const std::size_t close = xml.find("</" + tag + ">");
  return open == std::string::npos || close == std::string::npos ? std::string() : xml.substr(open, close - open);
}

std::vector<std::uint32_t> ReadOrder(const fm_workbook_t* wb, bool row) {
  const std::uint32_t* indices = nullptr;
  std::size_t count = 0;
  const fm_status_t status = row ? fm_sheet_get_pivot_row_field_order(wb, 0U, 0U, &indices, &count)
                                 : fm_sheet_get_pivot_col_field_order(wb, 0U, 0U, &indices, &count);
  EXPECT_EQ(status, 0) << fm_get_last_error_message();
  if (status != 0 || count == 0U) {
    return {};
  }
  return {indices, indices + count};
}

std::string SavedPivotXml(fm_workbook_t* wb) {
  std::uint8_t* data = nullptr;
  std::size_t len = 0;
  EXPECT_EQ(fm_workbook_save(wb, &data, &len), 0) << fm_get_last_error_message();
  const std::string xml = ExtractEntry(data, len, "xl/pivotTables/pivotTable1.xml");
  fm_buffer_destroy(data);
  return xml;
}

struct Case {
  const char* file;
  std::vector<std::uint32_t> rows;
};

class PivotValuesAxisFixture : public ::testing::TestWithParam<Case> {
 protected:
  void SetUp() override {
    const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_options/" + GetParam().file;
    const std::vector<std::uint8_t> bytes = ReadFileBytes(path);
    ASSERT_FALSE(bytes.empty()) << "could not read " << path;
    ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb_.handle), 0) << fm_get_last_error_message();
  }

  WorkbookHandle wb_;
};

}  // namespace

TEST_P(PivotValuesAxisFixture, RowOrderListsTheValuesFieldWhereExcelPutIt) {
  EXPECT_EQ(ReadOrder(wb_.handle, true), GetParam().rows);
  EXPECT_TRUE(ReadOrder(wb_.handle, false).empty());
}

TEST_P(PivotValuesAxisFixture, SaveKeepsTheValuesFieldOnTheRows) {
  const std::string xml = SavedPivotXml(wb_.handle);
  EXPECT_NE(xml.find("dataOnRows=\"1\""), std::string::npos);
  EXPECT_NE(ElementBody(xml, "rowFields").find("<field x=\"-2\"/>"), std::string::npos);
  EXPECT_EQ(xml.find("<colFields"), std::string::npos);
}

TEST_P(PivotValuesAxisFixture, MovingTheValuesFieldToTheColumnsGivesTheDefaultLayout) {
  uint32_t top = 0;
  uint32_t left = 0;
  uint32_t rows_before = 0;
  uint32_t cols_before = 0;
  fm_pivot_cells_t* layout = nullptr;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb_.handle, 0U, 0U, &layout), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_pivot_cells_get_bounds(layout, &top, &left, &rows_before, &cols_before), 0);
  fm_pivot_cells_destroy(layout);

  std::vector<std::uint32_t> columns = {FM_PIVOT_VALUES_FIELD};
  ASSERT_EQ(fm_sheet_set_pivot_col_field_order(wb_.handle, 0U, 0U, columns.data(), columns.size()), 0)
      << fm_get_last_error_message();
  std::vector<std::uint32_t> rows(GetParam().rows.begin(), GetParam().rows.end() - 1);
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb_.handle, 0U, 0U, rows.data(), rows.size()), 0)
      << fm_get_last_error_message();

  EXPECT_EQ(ReadOrder(wb_.handle, true), rows);
  EXPECT_EQ(ReadOrder(wb_.handle, false), columns);

  // One row per label instead of one per label and data field, one column per data field.
  uint32_t rows_after = 0;
  uint32_t cols_after = 0;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb_.handle, 0U, 0U, &layout), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_pivot_cells_get_bounds(layout, &top, &left, &rows_after, &cols_after), 0);
  fm_pivot_cells_destroy(layout);
  EXPECT_LT(rows_after, rows_before);
  EXPECT_GT(cols_after, cols_before);

  const std::string xml = SavedPivotXml(wb_.handle);
  EXPECT_EQ(xml.find("dataOnRows=\"1\""), std::string::npos);
  EXPECT_EQ(ElementBody(xml, "rowFields").find("x=\"-2\""), std::string::npos);
  EXPECT_NE(ElementBody(xml, "colFields").find("<field x=\"-2\"/>"), std::string::npos);
}

TEST_P(PivotValuesAxisFixture, MovingItBackToTheRowsRestoresTheExcelPlacement) {
  std::vector<std::uint32_t> columns = {FM_PIVOT_VALUES_FIELD};
  ASSERT_EQ(fm_sheet_set_pivot_col_field_order(wb_.handle, 0U, 0U, columns.data(), columns.size()), 0);
  std::vector<std::uint32_t> rows = GetParam().rows;
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb_.handle, 0U, 0U, rows.data(), rows.size()), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(ReadOrder(wb_.handle, true), GetParam().rows);
  EXPECT_TRUE(ReadOrder(wb_.handle, false).empty());

  const std::string xml = SavedPivotXml(wb_.handle);
  EXPECT_NE(xml.find("dataOnRows=\"1\""), std::string::npos);
  EXPECT_NE(ElementBody(xml, "rowFields").find("<field x=\"-2\"/>"), std::string::npos);
}

INSTANTIATE_TEST_SUITE_P(Fixtures, PivotValuesAxisFixture,
                         ::testing::Values(Case{"values_on_rows_one_row_field.xlsx", {0U, FM_PIVOT_VALUES_FIELD}},
                                           Case{"values_on_rows_two_row_fields.xlsx",
                                                {0U, 1U, FM_PIVOT_VALUES_FIELD}}));
