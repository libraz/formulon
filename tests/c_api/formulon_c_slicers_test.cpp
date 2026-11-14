//
// Stable C ABI tests for table slicers: layout pin, add / list / update /
// remove, argument validation, and removal with the table or the sheet.

#include <cstddef>
#include <cstdint>
#include <limits>
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
static_assert(sizeof(fm_slicer_info_t) == (kWasm32 ? 56U : 72U), "fm_slicer_info_t ABI layout changed");
static_assert(sizeof(fm_slicer_spec_t) == (kWasm32 ? 128U : 152U), "fm_slicer_spec_t ABI layout changed");
static_assert(sizeof(fm_slicer_patch_t) == (kWasm32 ? 56U : 72U), "fm_slicer_patch_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);

std::vector<std::uint8_t> Save(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0) << fm_get_last_error_message();
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

/// A spec placing the slicer at D5:F15, every other field left at its default.
fm_slicer_spec_t Anchored() {
  fm_slicer_spec_t s = FM_STRUCT_INIT(fm_slicer_spec_t);
  s.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  s.placement.anchor.from_row = 4;
  s.placement.anchor.from_col = 3;
  s.placement.anchor.to_row = 14;
  s.placement.anchor.to_col = 5;
  return s;
}

fm_slicer_info_t Read(const fm_workbook_t* wb, std::uint32_t id) {
  fm_slicer_info_t out = FM_STRUCT_INIT(fm_slicer_info_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_slicer(wb, 0, id, &present, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return out;
}

/// The slicer's frame in the sheet's drawing.
fm_drawing_object_info_t ReadFrame(const fm_workbook_t* wb, size_t sheet, std::uint32_t id) {
  fm_drawing_object_info_t out = FM_STRUCT_INIT(fm_drawing_object_info_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_drawing_object(wb, sheet, id, &present, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return out;
}

/// How many drawing objects of `sheet` are table slicers.
size_t SlicerCount(const fm_workbook_t* wb, size_t sheet) {
  size_t objects = 0;
  EXPECT_EQ(fm_sheet_get_drawing_object_count(wb, sheet, &objects), 0);
  size_t slicers = 0;
  for (size_t i = 0; i < objects; ++i) {
    fm_drawing_object_info_t obj = FM_STRUCT_INIT(fm_drawing_object_info_t);
    EXPECT_EQ(fm_sheet_get_drawing_object_at(wb, sheet, i, &obj), 0);
    fm_slicer_info_t info = FM_STRUCT_INIT(fm_slicer_info_t);
    int32_t present = 0;
    EXPECT_EQ(fm_sheet_get_slicer(wb, sheet, obj.object_id, &present, &info), 0);
    slicers += present == 1 ? 1U : 0U;
  }
  return slicers;
}

TEST(FormulonCApiSlicers, EnumOrdinals) {
  // Zero is the OOXML default `ascending`.
  EXPECT_EQ(FM_SLICER_SORT_ORDER_ASCENDING, 0);
  EXPECT_EQ(FM_SLICER_SORT_ORDER_DESCENDING, 1);
  EXPECT_EQ(FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_NO_DATA, 1);
  EXPECT_EQ(FM_SLICER_CROSS_FILTER_NONE, 2);
}

TEST(FormulonCApiSlicers, AddReadUpdateRemoveAndReload) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  EXPECT_EQ(SlicerCount(wb.handle, 0), 0U);

  fm_slicer_spec_t spec = Anchored();
  spec.caption = "Pick a product";
  spec.style_name = "SlicerStyleLight3";
  spec.column_count = 2U;
  spec.sort_order = FM_SLICER_SORT_ORDER_DESCENDING;
  spec.cross_filter = FM_SLICER_CROSS_FILTER_NONE;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(id, 2U);

  WorkbookGuard again;
  Load(Save(wb.handle), again);
  EXPECT_EQ(SlicerCount(again.handle, 0), 1U);
  fm_slicer_info_t out = Read(again.handle, id);
  EXPECT_EQ(out.object_id, 2U);
  EXPECT_STREQ(out.caption, "Pick a product");
  EXPECT_STREQ(out.cache_name, "Slicer_Product");
  EXPECT_STREQ(out.table_name, "Sales");
  EXPECT_STREQ(out.column_name, "Product");
  EXPECT_STREQ(out.style_name, "SlicerStyleLight3");
  EXPECT_EQ(out.column_count, 2U);
  EXPECT_DOUBLE_EQ(out.row_height_pt, 19.0);
  EXPECT_EQ(out.sort_order, FM_SLICER_SORT_ORDER_DESCENDING);
  EXPECT_TRUE(Has(ExtractZipEntry(Save(wb.handle), "xl/slicerCaches/slicerCache1.xml"), "sortOrder=\"descending\""));
  EXPECT_EQ(out.cross_filter, FM_SLICER_CROSS_FILTER_NONE);
  // The name and the anchor belong to the slicer's frame.
  fm_drawing_object_info_t frame = ReadFrame(again.handle, 0, id);
  EXPECT_STREQ(frame.name, "Product");
  EXPECT_EQ(frame.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(frame.anchor.from_row, 4U);
  EXPECT_EQ(frame.anchor.from_col, 3U);
  EXPECT_EQ(frame.anchor.to_row, 14U);
  EXPECT_EQ(frame.anchor.to_col, 5U);

  fm_slicer_patch_t update = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(update.caption_op, update.caption, "Renamed");
  update.style_name_op = FM_PATCH_OP_CLEAR;
  Set(update.column_count_op, update.column_count, 1U);
  ASSERT_EQ(fm_sheet_update_slicer(again.handle, 0, id, nullptr, &update), 0) << fm_get_last_error_message();
  fm_drawing_patch_t move = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(move.anchor.from_row_op, move.anchor.from_row, 0U);
  Set(move.anchor.to_row_op, move.anchor.to_row, 8U);
  ASSERT_EQ(fm_sheet_update_drawing(again.handle, 0, id, &move), 0) << fm_get_last_error_message();
  out = Read(again.handle, id);
  EXPECT_EQ(out.object_id, id);
  EXPECT_STREQ(out.caption, "Renamed");
  EXPECT_EQ(out.style_name, nullptr);  // CLEAR removes the style
  EXPECT_EQ(out.column_count, 1U);
  EXPECT_EQ(out.sort_order, FM_SLICER_SORT_ORDER_DESCENDING);
  frame = ReadFrame(again.handle, 0, id);
  EXPECT_EQ(frame.anchor.from_row, 0U);
  EXPECT_EQ(frame.anchor.to_row, 8U);
  EXPECT_EQ(frame.anchor.to_col, 5U);

  // Ascending, the schema default, is accepted as well.
  fm_slicer_patch_t ascending = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(ascending.sort_order_op, ascending.sort_order, FM_SLICER_SORT_ORDER_ASCENDING);
  ASSERT_EQ(fm_sheet_update_slicer(again.handle, 0, id, nullptr, &ascending), 0) << fm_get_last_error_message();
  EXPECT_EQ(Read(again.handle, id).sort_order, FM_SLICER_SORT_ORDER_ASCENDING);

  const std::vector<std::uint8_t> saved = Save(again.handle);
  EXPECT_TRUE(Has(ExtractZipEntry(saved, "xl/slicers/slicer1.xml"), "caption=\"Renamed\""));
  EXPECT_TRUE(Has(ExtractZipEntry(saved, "[Content_Types].xml"), "/xl/slicerCaches/slicerCache1.xml"));

  ASSERT_EQ(fm_sheet_remove_drawing(again.handle, 0, id), 0) << fm_get_last_error_message();
  EXPECT_EQ(SlicerCount(again.handle, 0), 0U);
  int32_t present = 1;
  ASSERT_EQ(fm_sheet_get_slicer(again.handle, 0, id, &present, &out), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_sheet_remove_drawing(again.handle, 0, id), FM_STATUS_NOT_FOUND);
  const std::vector<std::uint8_t> emptied = Save(again.handle);
  EXPECT_FALSE(HasEntry(emptied, "xl/slicers/slicer1.xml"));
  EXPECT_FALSE(HasEntry(emptied, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(Has(ExtractZipEntry(emptied, "xl/workbook.xml"), "Slicer_Product"));
}

TEST(FormulonCApiSlicers, PlacementNameDescriptionAndHiddenReachTheFrame) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_spec_t spec = Anchored();
  spec.placement.name = "Region picker";
  spec.placement.description = "filters sales";
  spec.placement.hidden = FM_TRISTATE_TRUE;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t frame = ReadFrame(wb.handle, 0, id);
  EXPECT_STREQ(frame.name, "Region picker");
  EXPECT_STREQ(frame.description, "filters sales");
  EXPECT_EQ(frame.hidden, 1);

  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(patch.description_op, patch.description, "new text");
  Set(patch.hidden_op, patch.hidden, 0);
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t after = ReadFrame(wb.handle, 0, id);
  EXPECT_STREQ(after.description, "new text");
  EXPECT_EQ(after.hidden, 0);
}

TEST(FormulonCApiSlicers, FrameNameCannotBeSetAndItsAnchorStaysTwoCell) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_spec_t spec = Anchored();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();

  fm_drawing_patch_t rename = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(rename.name_op, rename.name, "Other");
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &rename), kPreconditionFailed);
  fm_drawing_patch_t one_cell = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(one_cell.anchor.kind_op, one_cell.anchor.kind, FM_ANCHOR_KIND_ONE_CELL);
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &one_cell), kInvalidArgument);
  const fm_drawing_object_info_t frame = ReadFrame(wb.handle, 0, id);
  EXPECT_STREQ(frame.name, "Product");
  EXPECT_EQ(frame.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
}

TEST(FormulonCApiSlicers, RemovingTheTableRemovesItsSlicers) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_spec_t spec = Anchored();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0);
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 1, &spec, &id), 0);
  ASSERT_EQ(SlicerCount(wb.handle, 0), 2U);

  ASSERT_EQ(fm_sheet_remove_table(wb.handle, 0, 0), 0);
  EXPECT_EQ(SlicerCount(wb.handle, 0), 0U);
  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_FALSE(HasEntry(saved, "xl/slicers/slicer1.xml"));
  EXPECT_FALSE(HasEntry(saved, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(HasEntry(saved, "xl/slicerCaches/slicerCache2.xml"));
  const std::string book = ExtractZipEntry(saved, "xl/workbook.xml");
  EXPECT_FALSE(Has(book, "Slicer_")) << book;
  EXPECT_FALSE(Has(book, "slicerCache")) << book;
  const std::string types = ExtractZipEntry(saved, "[Content_Types].xml");
  EXPECT_FALSE(Has(types, "slicer")) << types;
}

TEST(FormulonCApiSlicers, RemovingTheSheetRemovesTheSlicersItHostsAndTheirCaches) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Dash", append_position(wb.handle), ignored_index()), 0);
  const size_t sheet = 1;
  fm_slicer_spec_t spec = Anchored();
  uint32_t id = 0;
  // The slicer lives on the second sheet, over the first sheet's table.
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, sheet, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
  ASSERT_EQ(SlicerCount(wb.handle, sheet), 1U);

  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, sheet), 0);
  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_FALSE(HasEntry(saved, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(Has(ExtractZipEntry(saved, "xl/workbook.xml"), "Slicer_"));

  // Removing the table's own sheet takes the slicers bound to its table.
  WorkbookGuard other;
  CreateSalesTable(other);
  ASSERT_EQ(fm_workbook_add_sheet(other.handle, "Dash", append_position(other.handle), ignored_index()), 0);
  ASSERT_EQ(fm_sheet_add_slicer(other.handle, sheet, 0, 0, 0, &spec, &id), 0);
  ASSERT_EQ(fm_workbook_remove_sheet(other.handle, 0), 0);
  const std::vector<std::uint8_t> after = Save(other.handle);
  EXPECT_FALSE(HasEntry(after, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(HasEntry(after, "xl/slicers/slicer1.xml"));
}

TEST(FormulonCApiSlicers, TableIsAddressedBySheetAndSheetLocalIndex) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Dash", append_position(wb.handle), ignored_index()), 0);
  fm_slicer_spec_t spec = Anchored();
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
  fm_slicer_spec_t spec = Anchored();
  uint32_t id = 0;
  fm_slicer_info_t out = FM_STRUCT_INIT(fm_slicer_info_t);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 7, 0, 0, 0, &spec, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 1, 0, &spec, &id), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 7, 0, 0, &spec, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  // A column position past the table's columns is a position out of range, like a table past the end.
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 2, &spec, &id), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_slicer_spec_t no_anchor = FM_STRUCT_INIT(fm_slicer_spec_t);
  no_anchor.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &no_anchor, &id), kInvalidArgument);
  fm_slicer_spec_t short_placement = Anchored();
  short_placement.placement.struct_size = 0;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &short_placement, &id), kInvalidArgument);
  fm_slicer_spec_t one_cell = Anchored();
  one_cell.placement.anchor.kind = FM_ANCHOR_KIND_ONE_CELL;
  one_cell.placement.anchor.width_emu = 100000;
  one_cell.placement.anchor.height_emu = 100000;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &one_cell, &id), kInvalidArgument);
  fm_slicer_spec_t absolute = Anchored();
  absolute.placement.anchor.kind = FM_ANCHOR_KIND_ABSOLUTE;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &absolute, &id), kInvalidArgument);
  fm_slicer_spec_t bad_hidden = Anchored();
  bad_hidden.placement.hidden = 3;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &bad_hidden, &id), kInvalidArgument);
  fm_slicer_spec_t bad_utf8 = Anchored();
  bad_utf8.caption = "\xC3";
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &bad_utf8, &id), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  fm_slicer_spec_t bad_sort = Anchored();
  bad_sort.sort_order = 2;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &bad_sort, &id), kInvalidArgument);
  fm_slicer_spec_t bad_cross = Anchored();
  bad_cross.cross_filter = 3;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &bad_cross, &id), kInvalidArgument);
  fm_slicer_spec_t backwards = Anchored();
  backwards.placement.anchor.to_row = 1;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &backwards, &id), kInvalidArgument);
  EXPECT_EQ(SlicerCount(wb.handle, 0), 0U);

  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0);
  fm_slicer_patch_t keep = FM_STRUCT_INIT(fm_slicer_patch_t);
  EXPECT_EQ(fm_sheet_update_slicer(nullptr, 0, id, nullptr, &keep), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, nullptr), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id + 1U, nullptr, &keep), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, 0U, nullptr, &keep), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 7, id, nullptr, &keep), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, id + 1U), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 7, id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  int32_t present = 1;
  EXPECT_EQ(fm_sheet_get_slicer(wb.handle, 0, id, nullptr, &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_slicer(wb.handle, 0, id, &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_slicer(wb.handle, 7, id, &present, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  ASSERT_EQ(fm_sheet_get_slicer(wb.handle, 0, id + 1U, &present, &out), 0);
  EXPECT_EQ(present, 0);
}

// A slicer patch writes only its SET fields, CLEAR removes the style, and a
// refused field (CLEAR outside `style_name`, a bad ordinal) leaves the whole slicer unchanged.
TEST(FormulonCApiSlicers, PatchWritesOnlyItsSetFieldsAndRefusesAtomically) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_spec_t spec = Anchored();
  spec.style_name = "SlicerStyleLight3";
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
  const fm_slicer_info_t added = Read(wb.handle, id);
  EXPECT_STREQ(added.caption, "Product");
  EXPECT_EQ(added.column_count, 1U);
  EXPECT_DOUBLE_EQ(added.row_height_pt, 19.0);
  EXPECT_EQ(added.sort_order, FM_SLICER_SORT_ORDER_ASCENDING);
  EXPECT_EQ(added.cross_filter, FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_DATA_AT_TOP);
  EXPECT_STREQ(ReadFrame(wb.handle, 0, id).name, "Product");
  EXPECT_EQ(ReadFrame(wb.handle, 0, id).anchor.from_row_offset_emu, 0);

  fm_slicer_patch_t caption = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(caption.caption_op, caption.caption, "Pick one");
  ASSERT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &caption), 0) << fm_get_last_error_message();
  fm_slicer_info_t out = Read(wb.handle, id);
  EXPECT_STREQ(out.caption, "Pick one");
  EXPECT_STREQ(out.style_name, "SlicerStyleLight3");
  EXPECT_EQ(ReadFrame(wb.handle, 0, id).anchor.from_row, 4U);

  fm_drawing_patch_t offset = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(offset.anchor.from_col_offset_emu_op, offset.anchor.from_col_offset_emu, 12700);
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &offset), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t frame = ReadFrame(wb.handle, 0, id);
  EXPECT_EQ(frame.anchor.from_col_offset_emu, 12700);
  EXPECT_EQ(frame.anchor.from_row, 4U);
  EXPECT_EQ(frame.anchor.from_col, 3U);
  EXPECT_EQ(frame.anchor.to_row, 14U);

  fm_slicer_patch_t clear = FM_STRUCT_INIT(fm_slicer_patch_t);
  clear.style_name_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &clear), 0) << fm_get_last_error_message();
  EXPECT_EQ(Read(wb.handle, id).style_name, nullptr);

  fm_slicer_patch_t bad = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(bad.column_count_op, bad.column_count, 3U);
  bad.caption_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &bad), kInvalidArgument);
  bad.caption_op = FM_PATCH_OP_KEEP;
  bad.sort_order_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &bad), kInvalidArgument);
  out = Read(wb.handle, id);
  EXPECT_EQ(out.column_count, 1U);
  EXPECT_STREQ(out.caption, "Pick one");

  const fm_slicer_patch_t keep = FM_STRUCT_INIT(fm_slicer_patch_t);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &keep), 0) << fm_get_last_error_message();
  fm_slicer_patch_t short_size = FM_STRUCT_INIT(fm_slicer_patch_t);
  short_size.struct_size = 0;
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &short_size), kInvalidArgument);
}

TEST(FormulonCApiSlicerUpdateAtomic, FailureChangesNothing) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_spec_t spec = Anchored();
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();

  // A valid frame half with a refused slicer half writes neither.
  fm_drawing_patch_t common = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(common.hidden_op, common.hidden, 1);
  Set(common.description_op, common.description, "Filter");
  fm_slicer_patch_t bad = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(bad.caption_op, bad.caption, "Changed");
  Set(bad.row_height_pt_op, bad.row_height_pt, -1.0);
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &common, &bad), kInvalidArgument);
  fm_drawing_object_info_t frame = ReadFrame(wb.handle, 0, id);
  EXPECT_EQ(frame.hidden, 0);
  EXPECT_STREQ(frame.description, "");
  EXPECT_STREQ(Read(wb.handle, id).caption, "Product");

  // A refused frame half (a slicer's name is fixed) keeps the slicer half unapplied.
  fm_drawing_patch_t rename = FM_STRUCT_INIT(fm_drawing_patch_t);
  Set(rename.name_op, rename.name, "Other");
  fm_slicer_patch_t caption = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(caption.caption_op, caption.caption, "Changed");
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &rename, &caption), kPreconditionFailed);
  EXPECT_STREQ(Read(wb.handle, id).caption, "Product");
  EXPECT_STREQ(ReadFrame(wb.handle, 0, id).name, "Product");

  // Both halves valid: one call applies both.
  ASSERT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, &common, &caption), 0) << fm_get_last_error_message();
  frame = ReadFrame(wb.handle, 0, id);
  EXPECT_EQ(frame.hidden, 1);
  EXPECT_STREQ(frame.description, "Filter");
  EXPECT_STREQ(Read(wb.handle, id).caption, "Changed");
}

TEST(FormulonCApiSlicerRowHeightAndStyleName, RowHeightInPoints) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_spec_t spec = Anchored();
  spec.row_height_pt = 25.5;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &spec, &id), 0) << fm_get_last_error_message();
  EXPECT_DOUBLE_EQ(Read(wb.handle, id).row_height_pt, 25.5);

  fm_slicer_patch_t taller = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(taller.row_height_pt_op, taller.row_height_pt, 30.0);
  ASSERT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &taller), 0) << fm_get_last_error_message();
  WorkbookGuard again;
  Load(Save(wb.handle), again);
  EXPECT_DOUBLE_EQ(Read(again.handle, id).row_height_pt, 30.0);

  fm_slicer_patch_t negative = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(negative.row_height_pt_op, negative.row_height_pt, -2.0);
  EXPECT_EQ(fm_sheet_update_slicer(again.handle, 0, id, nullptr, &negative), kInvalidArgument);
  fm_slicer_patch_t not_finite = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(not_finite.row_height_pt_op, not_finite.row_height_pt, std::numeric_limits<double>::quiet_NaN());
  EXPECT_EQ(fm_sheet_update_slicer(again.handle, 0, id, nullptr, &not_finite), kInvalidArgument);
  EXPECT_DOUBLE_EQ(Read(again.handle, id).row_height_pt, 30.0);

  fm_slicer_spec_t bad = Anchored();
  bad.row_height_pt = -1.0;
  uint32_t unused = 0;
  EXPECT_EQ(fm_sheet_add_slicer(again.handle, 0, 0, 0, 1, &bad, &unused), kInvalidArgument);
  bad.row_height_pt = std::numeric_limits<double>::infinity();
  EXPECT_EQ(fm_sheet_add_slicer(again.handle, 0, 0, 0, 1, &bad, &unused), kInvalidArgument);
  EXPECT_EQ(SlicerCount(again.handle, 0), 1U);
}

TEST(FormulonCApiSlicerRowHeightAndStyleName, StyleName) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_slicer_spec_t empty = Anchored();
  empty.style_name = "";
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &empty, &id), kInvalidArgument);
  EXPECT_EQ(SlicerCount(wb.handle, 0), 0U);

  fm_slicer_spec_t unstyled = Anchored();
  ASSERT_EQ(fm_sheet_add_slicer(wb.handle, 0, 0, 0, 0, &unstyled, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(Read(wb.handle, id).style_name, nullptr);

  fm_slicer_patch_t styled = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(styled.style_name_op, styled.style_name, "SlicerStyleDark1");
  ASSERT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &styled), 0) << fm_get_last_error_message();
  EXPECT_STREQ(Read(wb.handle, id).style_name, "SlicerStyleDark1");

  fm_slicer_patch_t set_empty = FM_STRUCT_INIT(fm_slicer_patch_t);
  Set(set_empty.style_name_op, set_empty.style_name, "");
  EXPECT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &set_empty), kInvalidArgument);
  EXPECT_STREQ(Read(wb.handle, id).style_name, "SlicerStyleDark1");

  fm_slicer_patch_t clear = FM_STRUCT_INIT(fm_slicer_patch_t);
  clear.style_name_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_slicer(wb.handle, 0, id, nullptr, &clear), 0) << fm_get_last_error_message();
  EXPECT_EQ(Read(wb.handle, id).style_name, nullptr);
}

}  // namespace
