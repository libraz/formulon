//
// Stable C ABI tests for table slicers: layout pin, add / list / update /
// remove, argument validation, and removal with the table or the sheet.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_pivot_test_helpers.h"
#include "gtest/gtest.h"
#include "miniz.h"
#include "utils/error.h"

namespace {

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_slicer) == (kWasm32 ? 104U : 128U), "fm_slicer ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

std::vector<std::uint8_t> Save(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0) << fm_last_error_message();
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

void Load(const std::vector<std::uint8_t>& bytes, WorkbookGuard& out) {
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out.handle), 0) << fm_last_error_message();
}

bool HasEntry(const std::vector<std::uint8_t>& archive, const char* path) {
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, archive.data(), archive.size(), 0) == MZ_FALSE) {
    return false;
  }
  const bool found = mz_zip_reader_locate_file(&reader, path, nullptr, 0) >= 0;
  mz_zip_reader_end(&reader);
  return found;
}

bool Has(const std::string& text, const std::string& needle) {
  return text.find(needle) != std::string::npos;
}

/// Sheet1 with Product/Amount headers and table `Sales` over A1:B3.
void CreateSalesTable(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "Product"), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 1, "Amount"), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 1, 0, "Widget"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 1, 10.0), 0);
  const char* columns[] = {"Product", "Amount"};
  std::size_t index = 99U;
  ASSERT_EQ(
      fm_workbook_table_create(wb.handle, 0, "A1:B2", "Sales", "Sales", columns, 2, "TableStyleMedium2", 1, 0, &index),
      0);
}

fm_slicer Anchored(std::uint32_t fields) {
  fm_slicer s{};
  s.fields = fields | FM_SLICER_FIELD_ANCHOR;
  s.from_row = 4;
  s.from_col = 3;
  s.to_row = 14;
  s.to_col = 5;
  return s;
}

TEST(FormulonCApiSlicers, EnumOrdinalsAndFieldBits) {
  EXPECT_EQ(FM_SLICER_SORT_DESCENDING, 1);
  EXPECT_EQ(FM_SLICER_CROSS_SHOW_ITEMS_WITH_NO_DATA, 1);
  EXPECT_EQ(FM_SLICER_CROSS_NONE, 2);
  EXPECT_EQ(FM_SLICER_FIELD_CROSS_FILTER, 0x80u);
}

TEST(FormulonCApiSlicers, AddListUpdateRemoveAndReload) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  size_t count = 99;
  ASSERT_EQ(fm_sheet_slicer_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  fm_slicer spec = Anchored(FM_SLICER_FIELD_CAPTION | FM_SLICER_FIELD_STYLE | FM_SLICER_FIELD_COLUMN_COUNT |
                            FM_SLICER_FIELD_SORT_ORDER | FM_SLICER_FIELD_CROSS_FILTER);
  spec.caption = "Pick a product";
  spec.style = "SlicerStyleLight3";
  spec.column_count = 2;
  spec.sort_order = FM_SLICER_SORT_DESCENDING;
  spec.cross_filter = FM_SLICER_CROSS_NONE;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &spec, &id), 0) << fm_last_error_message();
  EXPECT_EQ(id, 2U);

  WorkbookGuard again;
  Load(Save(wb.handle), again);
  ASSERT_EQ(fm_sheet_slicer_count(again.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_slicer out{};
  ASSERT_EQ(fm_sheet_slicer_at(again.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.fields, 0U);
  EXPECT_EQ(out.object_id, 2U);
  EXPECT_STREQ(out.name, "Product");
  EXPECT_STREQ(out.caption, "Pick a product");
  EXPECT_STREQ(out.cache_name, "Slicer_Product");
  EXPECT_STREQ(out.table_name, "Sales");
  EXPECT_STREQ(out.column_name, "Product");
  EXPECT_STREQ(out.style, "SlicerStyleLight3");
  EXPECT_EQ(out.column_count, 2U);
  EXPECT_EQ(out.row_height, 241300);
  EXPECT_EQ(out.sort_order, FM_SLICER_SORT_DESCENDING);
  EXPECT_EQ(out.cross_filter, FM_SLICER_CROSS_NONE);
  EXPECT_EQ(out.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(out.from_row, 4U);
  EXPECT_EQ(out.from_col, 3U);
  EXPECT_EQ(out.to_row, 14U);
  EXPECT_EQ(out.to_col, 5U);

  // The slicer's frame is a drawing object of the sheet.
  size_t objects = 0;
  ASSERT_EQ(fm_sheet_drawing_object_count(again.handle, 0, &objects), 0);
  EXPECT_EQ(objects, 1U);

  fm_slicer update = Anchored(FM_SLICER_FIELD_CAPTION | FM_SLICER_FIELD_STYLE | FM_SLICER_FIELD_COLUMN_COUNT);
  update.caption = "Renamed";
  update.style = "";
  update.column_count = 1;
  update.from_row = 0;
  update.to_row = 8;
  ASSERT_EQ(fm_sheet_update_slicer(again.handle, 0, 0, &update), 0) << fm_last_error_message();
  ASSERT_EQ(fm_sheet_slicer_at(again.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.caption, "Renamed");
  EXPECT_STREQ(out.style, "");
  EXPECT_EQ(out.column_count, 1U);
  EXPECT_EQ(out.sort_order, FM_SLICER_SORT_DESCENDING);
  EXPECT_EQ(out.from_row, 0U);
  EXPECT_EQ(out.to_row, 8U);

  const std::vector<std::uint8_t> saved = Save(again.handle);
  EXPECT_TRUE(Has(ExtractZipEntry(saved, "xl/slicers/slicer1.xml"), "caption=\"Renamed\""));
  EXPECT_TRUE(Has(ExtractZipEntry(saved, "[Content_Types].xml"), "/xl/slicerCaches/slicerCache1.xml"));

  ASSERT_EQ(fm_sheet_remove_slicer(again.handle, 0, 0), 0);
  ASSERT_EQ(fm_sheet_slicer_count(again.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  const std::vector<std::uint8_t> emptied = Save(again.handle);
  EXPECT_FALSE(HasEntry(emptied, "xl/slicers/slicer1.xml"));
  EXPECT_FALSE(HasEntry(emptied, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(Has(ExtractZipEntry(emptied, "xl/workbook.xml"), "Slicer_Product"));
}

TEST(FormulonCApiSlicers, RemovingTheTableRemovesItsSlicers) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer spec = Anchored(0);
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &spec, &id), 0);
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 1, &spec, &id), 0);
  size_t count = 0;
  ASSERT_EQ(fm_sheet_slicer_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);

  ASSERT_EQ(fm_workbook_table_remove(wb.handle, 0), 0);
  ASSERT_EQ(fm_sheet_slicer_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_FALSE(HasEntry(saved, "xl/slicers/slicer1.xml"));
  EXPECT_FALSE(HasEntry(saved, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(HasEntry(saved, "xl/slicerCaches/slicerCache2.xml"));
  const std::string book = ExtractZipEntry(saved, "xl/workbook.xml");
  EXPECT_FALSE(Has(book, "Slicer_")) << book;
  EXPECT_FALSE(Has(book, "extLst")) << book;
  const std::string types = ExtractZipEntry(saved, "[Content_Types].xml");
  EXPECT_FALSE(Has(types, "slicer")) << types;
}

TEST(FormulonCApiSlicers, RemovingTheSheetRemovesTheSlicersItHostsAndTheirCaches) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Dash"), 0);
  const size_t sheet = 1;
  fm_slicer spec = Anchored(0);
  uint32_t id = 0;
  // The slicer lives on the second sheet, over the first sheet's table.
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, sheet, 0, 0, &spec, &id), 0) << fm_last_error_message();
  size_t count = 0;
  ASSERT_EQ(fm_sheet_slicer_count(wb.handle, sheet, &count), 0);
  ASSERT_EQ(count, 1U);

  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, sheet), 0);
  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_FALSE(HasEntry(saved, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(Has(ExtractZipEntry(saved, "xl/workbook.xml"), "Slicer_"));

  // Removing the table's own sheet takes the slicers bound to its table.
  WorkbookGuard other;
  CreateSalesTable(other);
  ASSERT_EQ(fm_workbook_add_sheet(other.handle, "Dash"), 0);
  ASSERT_EQ(fm_sheet_add_slicer(other.handle, sheet, 0, 0, &spec, &id), 0);
  ASSERT_EQ(fm_workbook_remove_sheet(other.handle, 0), 0);
  const std::vector<std::uint8_t> after = Save(other.handle);
  EXPECT_FALSE(HasEntry(after, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(HasEntry(after, "xl/slicers/slicer1.xml"));
}

TEST(FormulonCApiSlicers, RejectsBadArguments) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer spec = Anchored(0);
  uint32_t id = 0;
  size_t count = 0;
  fm_slicer out{};
  EXPECT_EQ(fm_sheet_slicer_count(wb.handle, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_slicer_count(wb.handle, 7, &count), kInvalidArgument);
  EXPECT_EQ(fm_sheet_slicer_at(wb.handle, 0, 0, &out), kInvalidArgument);
  EXPECT_EQ(fm_sheet_slicer_at(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &spec, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 7, 0, 0, &spec, &id), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 1, 0, &spec, &id), kInvalidArgument);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 2, &spec, &id), kInvalidArgument);
  fm_slicer no_anchor{};
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &no_anchor, &id), kInvalidArgument);
  fm_slicer bad_bit = Anchored(0x4000u);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &bad_bit, &id), kInvalidArgument);
  fm_slicer bad_sort = Anchored(FM_SLICER_FIELD_SORT_ORDER);
  bad_sort.sort_order = 2;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &bad_sort, &id), kInvalidArgument);
  fm_slicer bad_cross = Anchored(FM_SLICER_FIELD_CROSS_FILTER);
  bad_cross.cross_filter = 3;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &bad_cross, &id), kInvalidArgument);
  fm_slicer bad_count = Anchored(FM_SLICER_FIELD_COLUMN_COUNT);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &bad_count, &id), kInvalidArgument);
  fm_slicer backwards = Anchored(0);
  backwards.to_row = 1;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &backwards, &id), kInvalidArgument);
  ASSERT_EQ(fm_sheet_slicer_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, &spec, &id), 0);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, 1, &spec), kInvalidArgument);
  fm_slicer rename = Anchored(FM_SLICER_FIELD_NAME);
  rename.name = "Other";
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, 0, &rename), kInvalidArgument);
  EXPECT_EQ(fm_sheet_remove_slicer(wb.handle, 0, 1), kInvalidArgument);
  EXPECT_EQ(fm_sheet_remove_slicer(wb.handle, 7, 0), kInvalidArgument);
}

}  // namespace
