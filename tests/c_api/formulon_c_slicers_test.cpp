//
// Stable C ABI tests for table slicers: layout pin, add / list / update /
// remove, argument validation, and removal with the table or the sheet.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "gtest/gtest.h"
#include "miniz.h"
#include "utils/error.h"

namespace {

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_slicer_info_t) == (kWasm32 ? 104U : 128U), "fm_slicer_info_t ABI layout changed");
static_assert(sizeof(fm_slicer_patch_t) == (kWasm32 ? 160U : 184U), "fm_slicer_patch_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

std::vector<std::uint8_t> Save(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

void Load(const std::vector<std::uint8_t>& bytes, WorkbookGuard& out) {
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out.handle), 0) << fm_get_last_error_message();
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
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "Amount"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 1, 0, "Widget"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 10.0), 0);
  const char* columns[] = {"Product", "Amount"};
  std::size_t index = 99U;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 1, 1}), "Sales", columns, 2, "TableStyleMedium2", 1,
                     0, &index),
            0);
}

template <typename T, typename V>
void Set(fm_patch_op_t& op, T& field, V value) {
  op = FM_PATCH_OP_SET;
  field = value;
}

/// A patch with the four anchor cells SET to D5:F15 and every other field KEEP.
fm_slicer_patch_t Anchored() {
  fm_slicer_patch_t s = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(s.from_row_op, s.from_row, 4U);
  Set(s.from_col_op, s.from_col, 3U);
  Set(s.to_row_op, s.to_row, 14U);
  Set(s.to_col_op, s.to_col, 5U);
  return s;
}

fm_slicer_info_t Read(const fm_workbook_t* wb, std::uint32_t id) {
  fm_slicer_info_t out = FM_STRUCT_INIT(fm_slicer_info_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_slicer(wb, 0, id, &present, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return out;
}

TEST(FormulonCApiSlicers, EnumOrdinals) {
  // Zero is the OOXML default `ascending`.
  EXPECT_EQ(FM_SLICER_SORT_ORDER_ASCENDING, 0);
  EXPECT_EQ(FM_SLICER_SORT_ORDER_DESCENDING, 1);
  EXPECT_EQ(FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_NO_DATA, 1);
  EXPECT_EQ(FM_SLICER_CROSS_FILTER_NONE, 2);
}

TEST(FormulonCApiSlicers, AddListUpdateRemoveAndReload) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  size_t count = 99;
  ASSERT_EQ(fm_sheet_get_slicer_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  fm_slicer_patch_t spec = Anchored();
  Set(spec.caption_op, spec.caption, "Pick a product");
  Set(spec.style_op, spec.style, "SlicerStyleLight3");
  Set(spec.column_count_op, spec.column_count, 2U);
  Set(spec.sort_order_op, spec.sort_order, FM_SLICER_SORT_ORDER_DESCENDING);
  Set(spec.cross_filter_op, spec.cross_filter, FM_SLICER_CROSS_FILTER_NONE);
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(id, 2U);

  WorkbookGuard again;
  Load(Save(wb.handle), again);
  ASSERT_EQ(fm_sheet_get_slicer_count(again.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_slicer_info_t out = FM_STRUCT_INIT(fm_slicer_info_t);
  ASSERT_EQ(fm_sheet_get_slicer_at(again.handle, 0, 0, &out), 0);
  EXPECT_EQ(out.object_id, 2U);
  EXPECT_STREQ(out.name, "Product");
  EXPECT_STREQ(out.caption, "Pick a product");
  EXPECT_STREQ(out.cache_name, "Slicer_Product");
  EXPECT_STREQ(out.table_name, "Sales");
  EXPECT_STREQ(out.column_name, "Product");
  EXPECT_STREQ(out.style, "SlicerStyleLight3");
  EXPECT_EQ(out.column_count, 2U);
  EXPECT_EQ(out.row_height_emu, 241300);
  EXPECT_EQ(out.sort_order, FM_SLICER_SORT_ORDER_DESCENDING);
  EXPECT_TRUE(Has(ExtractZipEntry(Save(wb.handle), "xl/slicerCaches/slicerCache1.xml"), "sortOrder=\"descending\""));
  EXPECT_EQ(out.cross_filter, FM_SLICER_CROSS_FILTER_NONE);
  EXPECT_EQ(out.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(out.from_row, 4U);
  EXPECT_EQ(out.from_col, 3U);
  EXPECT_EQ(out.to_row, 14U);
  EXPECT_EQ(out.to_col, 5U);

  // The slicer's frame is a drawing object of the sheet.
  size_t objects = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(again.handle, 0, &objects), 0);
  EXPECT_EQ(objects, 1U);

  fm_slicer_patch_t update = Anchored();
  Set(update.caption_op, update.caption, "Renamed");
  Set(update.style_op, update.style, "");
  Set(update.column_count_op, update.column_count, 1U);
  update.from_row = 0;
  update.to_row = 8;
  ASSERT_EQ(fm_sheet_update_slicer(again.handle, 0, id, &update), 0) << fm_get_last_error_message();
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_slicer(again.handle, 0, id, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(out.object_id, id);
  EXPECT_STREQ(out.caption, "Renamed");
  EXPECT_EQ(out.style, nullptr);  // a SET empty style removes it
  EXPECT_EQ(out.column_count, 1U);
  EXPECT_EQ(out.sort_order, FM_SLICER_SORT_ORDER_DESCENDING);
  EXPECT_EQ(out.from_row, 0U);
  EXPECT_EQ(out.to_row, 8U);

  // Ascending, the schema default, is accepted as well.
  fm_slicer_patch_t ascending = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(ascending.sort_order_op, ascending.sort_order, FM_SLICER_SORT_ORDER_ASCENDING);
  ASSERT_EQ(fm_sheet_update_slicer(again.handle, 0, id, &ascending), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_slicer(again.handle, 0, id, &present, &out), 0);
  EXPECT_EQ(out.sort_order, FM_SLICER_SORT_ORDER_ASCENDING);

  const std::vector<std::uint8_t> saved = Save(again.handle);
  EXPECT_TRUE(Has(ExtractZipEntry(saved, "xl/slicers/slicer1.xml"), "caption=\"Renamed\""));
  EXPECT_TRUE(Has(ExtractZipEntry(saved, "[Content_Types].xml"), "/xl/slicerCaches/slicerCache1.xml"));

  ASSERT_EQ(fm_sheet_remove_slicer(again.handle, 0, id), 0);
  ASSERT_EQ(fm_sheet_get_slicer_count(again.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  ASSERT_EQ(fm_sheet_get_slicer(again.handle, 0, id, &present, &out), 0);
  EXPECT_EQ(present, 0);
  const std::vector<std::uint8_t> emptied = Save(again.handle);
  EXPECT_FALSE(HasEntry(emptied, "xl/slicers/slicer1.xml"));
  EXPECT_FALSE(HasEntry(emptied, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(Has(ExtractZipEntry(emptied, "xl/workbook.xml"), "Slicer_Product"));
}

TEST(FormulonCApiSlicers, RemovingTheTableRemovesItsSlicers) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_patch_t spec = Anchored();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0);
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 1, &spec, &id), 0);
  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_slicer_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);

  ASSERT_EQ(fm_sheet_remove_table(wb.handle, 0, 0), 0);
  ASSERT_EQ(fm_sheet_get_slicer_count(wb.handle, 0, &count), 0);
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
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Dash", ignored_index()), 0);
  const size_t sheet = 1;
  fm_slicer_patch_t spec = Anchored();
  uint32_t id = 0;
  // The slicer lives on the second sheet, over the first sheet's table.
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, sheet, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_slicer_count(wb.handle, sheet, &count), 0);
  ASSERT_EQ(count, 1U);

  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, sheet), 0);
  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_FALSE(HasEntry(saved, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(Has(ExtractZipEntry(saved, "xl/workbook.xml"), "Slicer_"));

  // Removing the table's own sheet takes the slicers bound to its table.
  WorkbookGuard other;
  CreateSalesTable(other);
  ASSERT_EQ(fm_workbook_add_sheet(other.handle, "Dash", ignored_index()), 0);
  ASSERT_EQ(fm_sheet_add_slicer(other.handle, sheet, 0, 0, 0, &spec, &id), 0);
  ASSERT_EQ(fm_workbook_remove_sheet(other.handle, 0), 0);
  const std::vector<std::uint8_t> after = Save(other.handle);
  EXPECT_FALSE(HasEntry(after, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(HasEntry(after, "xl/slicers/slicer1.xml"));
}

TEST(FormulonCApiSlicers, TableIsAddressedBySheetAndSheetLocalIndex) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Dash", ignored_index()), 0);
  fm_slicer_patch_t spec = Anchored();
  uint32_t id = 0;
  // The table lives on sheet 0; sheet 1 has none, so its index 0 is past the end.
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 1, 1, 0, 0, &spec, &id), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 1, 5, 0, 0, &spec, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 1, 0, 1, 0, &spec, &id), FM_STATUS_INDEX_OUT_OF_RANGE);
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 1, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiSlicers, RejectsBadArguments) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_patch_t spec = Anchored();
  uint32_t id = 0;
  size_t count = 0;
  fm_slicer_info_t out = FM_STRUCT_INIT(fm_slicer_info_t);
  EXPECT_EQ(fm_sheet_get_slicer_count(wb.handle, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_slicer_count(wb.handle, 7, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_slicer_at(wb.handle, 0, 0, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_slicer_at(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 7, 0, 0, 0, &spec, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 1, 0, &spec, &id), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 7, 0, 0, &spec, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 2, &spec, &id), kInvalidArgument);
  fm_slicer_patch_t no_anchor = FM_STRUCT_INIT(fm_slicer_patch_t);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &no_anchor, &id), kInvalidArgument);
  fm_slicer_patch_t three_cells = Anchored();
  three_cells.to_col_op = FM_PATCH_OP_KEEP;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &three_cells, &id), kInvalidArgument);
  fm_slicer_patch_t bad_op = Anchored();
  bad_op.caption_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &bad_op, &id), kInvalidArgument);
  fm_slicer_patch_t null_caption = Anchored();
  Set(null_caption.caption_op, null_caption.caption, nullptr);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &null_caption, &id), kInvalidArgument);
  fm_slicer_patch_t bad_utf8 = Anchored();
  Set(bad_utf8.caption_op, bad_utf8.caption, "\xC3");
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &bad_utf8, &id), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  fm_slicer_patch_t bad_sort = Anchored();
  Set(bad_sort.sort_order_op, bad_sort.sort_order, 2);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &bad_sort, &id), kInvalidArgument);
  fm_slicer_patch_t bad_cross = Anchored();
  Set(bad_cross.cross_filter_op, bad_cross.cross_filter, 3);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &bad_cross, &id), kInvalidArgument);
  fm_slicer_patch_t bad_count = Anchored();
  Set(bad_count.column_count_op, bad_count.column_count, 0U);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &bad_count, &id), kInvalidArgument);
  fm_slicer_patch_t backwards = Anchored();
  backwards.to_row = 1;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &backwards, &id), kInvalidArgument);
  ASSERT_EQ(fm_sheet_get_slicer_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id + 1U, &spec), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, 0U, &spec), FM_STATUS_NOT_FOUND);
  fm_slicer_patch_t rename = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(rename.name_op, rename.name, "Other");
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &rename), kInvalidArgument);
  EXPECT_EQ(fm_sheet_remove_slicer(wb.handle, 0, id + 1U), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_slicer(wb.handle, 7, id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  int32_t present = 1;
  EXPECT_EQ(fm_sheet_get_slicer(wb.handle, 0, id, nullptr, &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_slicer(wb.handle, 7, id, &present, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  ASSERT_EQ(fm_sheet_get_slicer(wb.handle, 0, id + 1U, &present, &out), 0);
  EXPECT_EQ(present, 0);
}

// A slicer patch writes only its SET fields: an offset alone moves within the
// current cells, CLEAR removes the style, and a refused field (CLEAR outside
// `style`, a partial set of anchor cells) leaves the whole slicer unchanged.
TEST(FormulonCApiSlicers, PatchWritesOnlyItsSetFieldsAndRefusesAtomically) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_patch_t spec = Anchored();
  Set(spec.style_op, spec.style, "SlicerStyleLight3");
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
  const fm_slicer_info_t added = Read(wb.handle, id);
  EXPECT_STREQ(added.name, "Product");
  EXPECT_STREQ(added.caption, "Product");
  EXPECT_EQ(added.column_count, 1U);
  EXPECT_EQ(added.row_height_emu, 241300);
  EXPECT_EQ(added.sort_order, FM_SLICER_SORT_ORDER_ASCENDING);
  EXPECT_EQ(added.cross_filter, FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_DATA_AT_TOP);
  EXPECT_EQ(added.from_row_offset_emu, 0);

  fm_slicer_patch_t caption = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(caption.caption_op, caption.caption, "Pick one");
  ASSERT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &caption), 0) << fm_get_last_error_message();
  fm_slicer_info_t out = Read(wb.handle, id);
  EXPECT_STREQ(out.caption, "Pick one");
  EXPECT_STREQ(out.style, "SlicerStyleLight3");
  EXPECT_EQ(out.from_row, 4U);
  EXPECT_EQ(out.to_col, 5U);

  fm_slicer_patch_t offset = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(offset.from_col_offset_emu_op, offset.from_col_offset_emu, 12700);
  ASSERT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &offset), 0) << fm_get_last_error_message();
  out = Read(wb.handle, id);
  EXPECT_EQ(out.from_col_offset_emu, 12700);
  EXPECT_EQ(out.from_row, 4U);
  EXPECT_EQ(out.from_col, 3U);
  EXPECT_EQ(out.to_row, 14U);

  fm_slicer_patch_t clear = FM_STRUCT_INIT(fm_slicer_patch_t);
  clear.style_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &clear), 0) << fm_get_last_error_message();
  EXPECT_EQ(Read(wb.handle, id).style, nullptr);

  fm_slicer_patch_t bad = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(bad.column_count_op, bad.column_count, 3U);
  bad.caption_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &bad), kInvalidArgument);
  bad.caption_op = FM_PATCH_OP_KEEP;
  Set(bad.from_row_op, bad.from_row, 0U);  // one anchor cell without the other three
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &bad), kInvalidArgument);
  bad.from_row_op = FM_PATCH_OP_KEEP;
  bad.sort_order_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &bad), kInvalidArgument);
  out = Read(wb.handle, id);
  EXPECT_EQ(out.column_count, 1U);
  EXPECT_STREQ(out.caption, "Pick one");
  EXPECT_EQ(out.from_row, 4U);

  const fm_slicer_patch_t keep = FM_STRUCT_INIT(fm_slicer_patch_t);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &keep), 0) << fm_get_last_error_message();
  fm_slicer_patch_t short_size = FM_STRUCT_INIT(fm_slicer_patch_t);
  short_size.struct_size = 0;
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &short_size), kInvalidArgument);
}

}  // namespace
