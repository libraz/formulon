//
// C ABI regression tests for worksheet table metadata (style toggles, column
// totals, rename) and custom table / PivotTable styles.

#include <cstddef>
#include <cstdint>
#include <functional>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_setters.h"
#include "gtest/gtest.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "miniz.h"
#include "table.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kAlreadyExists = static_cast<fm_status_t>(formulon::FormulonErrorCode::kAlreadyExists);

std::vector<std::uint8_t> Save(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

/// Rebuilds `archive` with `path`'s body replaced by `edit(body)`.
std::vector<std::uint8_t> RewriteEntry(const std::vector<std::uint8_t>& archive, std::string_view path,
                                       const std::function<std::string(std::string)>& edit) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, archive.data(), archive.size(), 0), MZ_FALSE);
  std::vector<std::string> names;
  std::vector<std::string> bodies;
  for (mz_uint i = 0; i < mz_zip_reader_get_num_files(&reader); ++i) {
    char name[512];
    mz_zip_reader_get_filename(&reader, i, name, sizeof(name));
    names.emplace_back(name);
    bodies.push_back(ExtractZipEntry(archive, name));
    if (names.back() == path) {
      bodies.back() = edit(bodies.back());
    }
  }
  mz_zip_reader_end(&reader);
  std::vector<PartFile> parts;
  for (std::size_t i = 0; i < names.size(); ++i) {
    parts.push_back(PartFile{names[i].c_str(), bodies[i]});
  }
  return BuildZip(parts);
}

std::string ReplaceOnce(std::string text, std::string_view from, std::string_view to) {
  const std::size_t at = text.find(from);
  EXPECT_NE(at, std::string::npos) << from;
  if (at != std::string::npos) {
    text.replace(at, from.size(), to);
  }
  return text;
}

void Load(const std::vector<std::uint8_t>& bytes, WorkbookGuard& out) {
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out.handle), 0);
}

/// Sheet1 with Product/Amount headers, two data rows and table `Sales` over
/// A1:B3.
void CreateSalesTable(WorkbookGuard& wb, const char* style = "TableStyleMedium2", int32_t totals = 0,
                      fm_cell_rect_t ref = {0, 0, 2, 1}) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "Amount"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 1, 0, "Widget"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 10.0), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 0, "Gadget"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2, 1, 20.0), 0);
  const char* columns[] = {"Product", "Amount"};
  std::size_t index = 99U;
  ASSERT_EQ(AddTable(wb.handle, 0, &ref, "Sales", columns, 2, style, 1, totals, &index), 0);
  ASSERT_EQ(index, 0U);
}

fm_table_info_t Info(fm_workbook_t* wb, std::size_t index = 0U, std::size_t sheet = 0U) {
  fm_table_info_t info = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_sheet_get_table_at(wb, sheet, index, &info), 0);
  return info;
}

/// Patch with every op at KEEP.
fm_table_patch_t TablePatch() {
  return FM_STRUCT_INIT(fm_table_patch_t);
}

/// Sets the table's style name, leaving its toggles.
fm_status_t SetTableStyleName(fm_workbook_t* wb, std::size_t index, const char* style_name) {
  fm_table_style_options_patch_t patch = FM_STRUCT_INIT(fm_table_style_options_patch_t);
  patch.style_name_op = FM_PATCH_OP_SET;
  patch.style_name = style_name;
  return fm_sheet_update_table_style_options(wb, 0, index, &patch);
}

/// Patches the table / PivotTable style defaults: `nullptr` keeps a name, "" clears it, anything else sets it.
fm_status_t UpdateStyleDefaults(fm_workbook_t* wb, const char* table_style, const char* pivot_style) {
  fm_table_style_defaults_patch_t patch = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  if (table_style != nullptr) {
    patch.table_style_name_op = table_style[0] == '\0' ? FM_PATCH_OP_CLEAR : FM_PATCH_OP_SET;
    patch.table_style_name = table_style[0] == '\0' ? nullptr : table_style;
  }
  if (pivot_style != nullptr) {
    patch.pivot_style_name_op = pivot_style[0] == '\0' ? FM_PATCH_OP_CLEAR : FM_PATCH_OP_SET;
    patch.pivot_style_name = pivot_style[0] == '\0' ? nullptr : pivot_style;
  }
  return fm_workbook_update_table_style_defaults(wb, &patch);
}

/// The style record together with the `out_present` flag of the getter.
struct TableStyleRead : fm_table_style_options_t {
  std::int32_t present = 0;
};

TableStyleRead StyleOptions(const fm_workbook_t* wb, std::size_t index = 0U, std::size_t sheet = 0U) {
  TableStyleRead read;
  static_cast<fm_table_style_options_t&>(read) = FM_STRUCT_INIT(fm_table_style_options_t);
  EXPECT_EQ(fm_sheet_get_table_style_options(wb, sheet, index, &read.present, &read), 0);
  return read;
}

std::string Formula(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  const char* text = nullptr;
  std::int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_formula(wb, 0, row, col, &present, &text), 0);
  return present != 0 && text != nullptr ? text : "";
}

/// `pivot` / `table` are 0 / 1 here and land as the tristate's FALSE / TRUE.
fm_table_style_t MakeTableStyle(const char* name, int32_t pivot, int32_t table,
                                const fm_table_style_element_t* elements, std::uint32_t element_count) {
  fm_table_style_t style = FM_STRUCT_INIT(fm_table_style_t);
  style.name = name;
  style.pivot = pivot != 0 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  style.table = table != 0 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  style.elements = elements;
  style.element_count = element_count;
  return style;
}

fm_value_t Value(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row, col, &v), 0);
  return v;
}

TEST(FormulonCApiTable, CreateDefaultsAndReadBack) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  const fm_table_info_t info = Info(wb.handle);
  EXPECT_STREQ(info.name, "Sales");
  EXPECT_EQ(info.range.first_row, 0U);
  EXPECT_EQ(info.range.first_col, 0U);
  EXPECT_EQ(info.range.last_row, 2U);
  EXPECT_EQ(info.range.last_col, 1U);
  EXPECT_STREQ(info.style_name, "TableStyleMedium2");
  EXPECT_EQ(info.show_first_column, 0);
  EXPECT_EQ(info.show_last_column, 0);
  EXPECT_EQ(info.show_row_stripes, 1);
  EXPECT_EQ(info.show_column_stripes, 0);
  EXPECT_EQ(info.header_row_count, 1U);
  EXPECT_EQ(info.totals_row_count, 0U);
  EXPECT_EQ(info.column_count, 2U);

  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  ASSERT_EQ(fm_sheet_get_table_column_at(wb.handle, 0, 0, 1, &column), 0);
  EXPECT_STREQ(column.name, "Amount");
  EXPECT_EQ(column.totals_row_function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  EXPECT_EQ(column.totals_row_label, nullptr) << "no label reads NULL";
  EXPECT_EQ(column.totals_row_formula, nullptr) << "no custom formula reads NULL";
  EXPECT_EQ(fm_sheet_get_table_column_at(wb.handle, 0, 0, 2, &column), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_at(wb.handle, 0, 1, nullptr), kBindingNullPointer);
  fm_table_info_t ignored = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_sheet_get_table_at(wb.handle, 0, 1, &ignored), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_at(wb.handle, 7, 0, &ignored), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  ignored.struct_size = 0U;
  EXPECT_EQ(fm_sheet_get_table_at(wb.handle, 0, 0, &ignored), kInvalidArgument) << "struct_size is checked";

  WorkbookGuard plain;
  CreateSalesTable(plain, "");
  const fm_table_info_t unstyled = Info(plain.handle);
  EXPECT_EQ(unstyled.show_row_stripes, 0);
  EXPECT_EQ(unstyled.style_name, nullptr) << "a table without a style reads NULL";
  EXPECT_EQ(StyleOptions(plain.handle).present, 0);
}

TEST(FormulonCApiTable, LoadedStyleTogglesSurviveAStyleNameChange) {
  WorkbookGuard source;
  CreateSalesTable(source, "TableStyleLight9");
  const std::vector<std::uint8_t> patched =
      RewriteEntry(Save(source.handle), "xl/tables/table1.xml", [](std::string xml) {
        return ReplaceOnce(std::move(xml),
                           "<tableStyleInfo name=\"TableStyleLight9\" showFirstColumn=\"0\" showLastColumn=\"0\" "
                           "showRowStripes=\"1\" showColumnStripes=\"0\"/>",
                           "<tableStyleInfo name=\"TableStyleLight9\" showFirstColumn=\"1\" showLastColumn=\"1\" "
                           "showRowStripes=\"0\" showColumnStripes=\"1\"/>");
      });
  WorkbookGuard wb;
  Load(patched, wb);
  fm_table_info_t info = Info(wb.handle);
  EXPECT_STREQ(info.style_name, "TableStyleLight9");
  EXPECT_EQ(info.show_first_column, 1);
  EXPECT_EQ(info.show_last_column, 1);
  EXPECT_EQ(info.show_row_stripes, 0);
  EXPECT_EQ(info.show_column_stripes, 1);

  ASSERT_EQ(SetTableStyleName(wb.handle, 0, "TableStyleMedium4"), 0);
  info = Info(wb.handle);
  EXPECT_STREQ(info.style_name, "TableStyleMedium4");
  EXPECT_EQ(info.show_first_column, 1);
  EXPECT_EQ(info.show_last_column, 1);
  EXPECT_EQ(info.show_row_stripes, 0);
  EXPECT_EQ(info.show_column_stripes, 1);

  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_NE(ExtractZipEntry(saved, "xl/tables/table1.xml")
                .find("<tableStyleInfo name=\"TableStyleMedium4\" showFirstColumn=\"1\" showLastColumn=\"1\" "
                      "showRowStripes=\"0\" showColumnStripes=\"1\"/>"),
            std::string::npos);
  WorkbookGuard reloaded;
  Load(saved, reloaded);
  info = Info(reloaded.handle);
  EXPECT_STREQ(info.style_name, "TableStyleMedium4");
  EXPECT_EQ(info.show_first_column, 1);
  EXPECT_EQ(info.show_column_stripes, 1);
}

fm_table_style_options_patch_t StylePatch() {
  return FM_STRUCT_INIT(fm_table_style_options_patch_t);
}

TEST(FormulonCApiTable, UpdateStyleOptionsWritesTheElementInSchemaPosition) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "");
  fm_table_style_options_patch_t patch = StylePatch();
  patch.style_name_op = FM_PATCH_OP_SET;
  patch.style_name = "TableStyleDark1";
  patch.show_first_column_op = FM_PATCH_OP_SET;
  patch.show_first_column = 1;
  patch.show_last_column_op = FM_PATCH_OP_SET;
  patch.show_last_column = 0;
  patch.show_row_stripes_op = FM_PATCH_OP_SET;
  patch.show_row_stripes = 0;
  patch.show_column_stripes_op = FM_PATCH_OP_SET;
  patch.show_column_stripes = 1;
  ASSERT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &patch), 0);
  fm_table_info_t info = Info(wb.handle);
  EXPECT_STREQ(info.style_name, "TableStyleDark1");
  EXPECT_EQ(info.show_first_column, 1);
  EXPECT_EQ(info.show_column_stripes, 1);

  // Clearing the style name keeps the element and its toggles, as the "None" style does.
  patch = StylePatch();
  patch.style_name_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &patch), 0);
  info = Info(wb.handle);
  EXPECT_EQ(StyleOptions(wb.handle).present, 1);
  EXPECT_EQ(info.style_name, nullptr);
  EXPECT_EQ(info.show_first_column, 1);
  EXPECT_EQ(info.show_column_stripes, 1);
  const std::string xml = ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml");
  const std::size_t columns_end = xml.find("</tableColumns>");
  const std::size_t style = xml.find(
      "<tableStyleInfo showFirstColumn=\"1\" showLastColumn=\"0\" "
      "showRowStripes=\"0\" showColumnStripes=\"1\"/>");
  ASSERT_NE(columns_end, std::string::npos);
  ASSERT_NE(style, std::string::npos);
  EXPECT_LT(columns_end, style);

  patch = StylePatch();
  patch.style_name_op = FM_PATCH_OP_SET;
  patch.style_name = "X";
  EXPECT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 3, &patch), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_table_style_options(wb.handle, 4, 0, &patch), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiTable, GetStyleOptionsReadsTheAppliedStyle) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleLight9");
  auto info = StyleOptions(wb.handle);
  EXPECT_EQ(info.present, 1);
  EXPECT_STREQ(info.style_name, "TableStyleLight9");
  EXPECT_EQ(info.show_first_column, 0);
  EXPECT_EQ(info.show_last_column, 0);
  EXPECT_EQ(info.show_row_stripes, 1);
  EXPECT_EQ(info.show_column_stripes, 0);

  WorkbookGuard plain;
  CreateSalesTable(plain, "");
  info = StyleOptions(plain.handle);
  EXPECT_EQ(info.present, 0);
  EXPECT_EQ(info.style_name, nullptr);
  EXPECT_EQ(info.show_row_stripes, 0);

  fm_table_style_options_t ignored = FM_STRUCT_INIT(fm_table_style_options_t);
  std::int32_t ignored_present = 0;
  EXPECT_EQ(fm_sheet_get_table_style_options(wb.handle, 0, 0, &ignored_present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_table_style_options(wb.handle, 0, 0, nullptr, &ignored), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_table_style_options(nullptr, 0, 0, &ignored_present, &ignored), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_table_style_options(wb.handle, 0, 1, &ignored_present, &ignored),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_style_options(wb.handle, 9, 0, &ignored_present, &ignored),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiTable, KeepOnlyStylePatchChangesNothing) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleLight9");
  const std::string before = ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml");
  fm_table_style_options_patch_t patch = StylePatch();
  ASSERT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &patch), 0);
  EXPECT_EQ(ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml"), before);
}

TEST(FormulonCApiTable, UpdateStyleOptionsMergesOnlyTheSetFields) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleLight9");
  fm_table_style_options_patch_t patch = StylePatch();
  patch.show_last_column_op = FM_PATCH_OP_SET;
  patch.show_last_column = 1;
  ASSERT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &patch), 0);
  const auto info = StyleOptions(wb.handle);
  EXPECT_STREQ(info.style_name, "TableStyleLight9");
  EXPECT_EQ(info.show_last_column, 1);
  EXPECT_EQ(info.show_row_stripes, 1);
  EXPECT_EQ(info.show_first_column, 0);
}

TEST(FormulonCApiTable, UpdateStyleOptionsOnAnUnstyledTableCreatesTheElementWithDefaults) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "");
  fm_table_style_options_patch_t patch = StylePatch();
  patch.show_first_column_op = FM_PATCH_OP_SET;
  patch.show_first_column = 1;
  ASSERT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &patch), 0);
  const auto info = StyleOptions(wb.handle);
  EXPECT_EQ(info.present, 1);
  EXPECT_EQ(info.show_first_column, 1);
  EXPECT_EQ(info.show_row_stripes, 1) << "the creation default of a styled table";
  EXPECT_EQ(info.show_last_column, 0);
  EXPECT_EQ(info.show_column_stripes, 0);
}

TEST(FormulonCApiTable, InvalidStylePatchIsRejectedBeforeAnythingChanges) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleLight9");
  const std::string before = ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml");

  fm_table_style_options_patch_t bad_op = StylePatch();
  bad_op.show_first_column_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &bad_op), kInvalidArgument);

  // CLEAR is only meaningful on the style name.
  fm_table_style_options_patch_t clear_toggle = StylePatch();
  clear_toggle.style_name_op = FM_PATCH_OP_SET;
  clear_toggle.style_name = "TableStyleDark1";
  clear_toggle.show_row_stripes_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &clear_toggle), kInvalidArgument);

  fm_table_style_options_patch_t null_name = StylePatch();
  null_name.style_name_op = FM_PATCH_OP_SET;
  null_name.show_last_column_op = FM_PATCH_OP_SET;
  null_name.show_last_column = 1;
  EXPECT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &null_name), FM_STATUS_INVALID_ARGUMENT);

  fm_table_style_options_patch_t bad_toggle = StylePatch();
  bad_toggle.show_first_column_op = FM_PATCH_OP_SET;
  bad_toggle.show_first_column = 2;
  EXPECT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &bad_toggle), kInvalidArgument);

  EXPECT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml"), before);
}

TEST(FormulonCApiTable, AddWithoutAStyleNameTakesTheWorkbookDefault) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* columns[] = {"A", "B"};
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "A"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "B"), 0);
  std::size_t index = 99U;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 2, 1}), "Plain", columns, 2, nullptr, 1, 0, &index),
            0);
  EXPECT_EQ(index, 0U);
  auto info = StyleOptions(wb.handle, 0);
  EXPECT_STREQ(info.style_name, "TableStyleMedium2");
  EXPECT_EQ(info.show_row_stripes, 1);
  EXPECT_EQ(info.show_first_column, 0);
  EXPECT_EQ(info.show_last_column, 0);
  EXPECT_EQ(info.show_column_stripes, 0);

  ASSERT_EQ(UpdateStyleDefaults(wb.handle, "TableStyleLight9", nullptr), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 5, 0, "A"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 5, 1, "B"), 0);
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{5, 0, 7, 1}), "Later", columns, 2, nullptr, 1, 0, &index),
            0);
  EXPECT_EQ(index, 1U);
  info = StyleOptions(wb.handle, 1);
  EXPECT_STREQ(info.style_name, "TableStyleLight9");
  EXPECT_EQ(info.show_row_stripes, 1);
  EXPECT_STREQ(StyleOptions(wb.handle, 0).style_name, "TableStyleMedium2") << "an existing table is not restyled";

  const std::string xml = ExtractZipEntry(Save(wb.handle), "xl/tables/table2.xml");
  EXPECT_NE(xml.find("<tableStyleInfo name=\"TableStyleLight9\""), std::string::npos) << xml;
}

TEST(FormulonCApiTable, AddWithAnEmptyStyleNameCreatesNoStyleElement) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(UpdateStyleDefaults(wb.handle, "TableStyleLight9", nullptr), 0);
  CreateSalesTable(wb, "");
  const auto info = StyleOptions(wb.handle);
  EXPECT_EQ(info.present, 0);
  EXPECT_EQ(ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml").find("<tableStyleInfo"), std::string::npos);
}

TEST(FormulonCApiTable, ColumnTotalsWriteTheTotalsRowCells) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleMedium2", 1, {0, 0, 3, 1});
  ASSERT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 0, FM_TABLE_TOTALS_ROW_FUNCTION_NONE, "Total", nullptr), 0);
  ASSERT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr), 0);
  EXPECT_EQ(Formula(wb.handle, 3, 1), "=SUBTOTAL(109,[Amount])");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t total = Value(wb.handle, 3, 1);
  ASSERT_EQ(total.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(total.u.number, 30.0);
  fm_value_t label = Value(wb.handle, 3, 0);
  ASSERT_EQ(label.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(label.u.text, "Total");

  const std::vector<std::uint8_t> saved = Save(wb.handle);
  const std::string table_xml = ExtractZipEntry(saved, "xl/tables/table1.xml");
  EXPECT_NE(table_xml.find("totalsRowCount=\"1\""), std::string::npos);
  EXPECT_NE(table_xml.find("<tableColumn id=\"1\" name=\"Product\" totalsRowLabel=\"Total\"/>"), std::string::npos);
  EXPECT_NE(table_xml.find("<tableColumn id=\"2\" name=\"Amount\" totalsRowFunction=\"sum\"/>"), std::string::npos);
  EXPECT_NE(ExtractZipEntry(saved, "xl/worksheets/sheet1.xml").find("<f>SUBTOTAL(109,Sales[Amount])</f>"),
            std::string::npos);

  WorkbookGuard reloaded;
  Load(saved, reloaded);
  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  ASSERT_EQ(fm_sheet_get_table_column_at(reloaded.handle, 0, 0, 0, &column), 0);
  EXPECT_STREQ(column.totals_row_label, "Total");
  EXPECT_EQ(column.totals_row_function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  ASSERT_EQ(fm_sheet_get_table_column_at(reloaded.handle, 0, 0, 1, &column), 0);
  EXPECT_EQ(column.totals_row_function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
  EXPECT_EQ(Info(reloaded.handle).totals_row_count, 1U);
  EXPECT_EQ(Formula(reloaded.handle, 3, 1), "=SUBTOTAL(109,[Amount])");

  // Each built-in function maps to its SUBTOTAL number; `none` clears the cell.
  const struct {
    int32_t function;
    const char* formula;
  } kCases[] = {{FM_TABLE_TOTALS_ROW_FUNCTION_AVERAGE, "=SUBTOTAL(101,[Amount])"},
                {FM_TABLE_TOTALS_ROW_FUNCTION_COUNT_NUMBERS, "=SUBTOTAL(102,[Amount])"},
                {FM_TABLE_TOTALS_ROW_FUNCTION_COUNT, "=SUBTOTAL(103,[Amount])"},
                {FM_TABLE_TOTALS_ROW_FUNCTION_MAX, "=SUBTOTAL(104,[Amount])"},
                {FM_TABLE_TOTALS_ROW_FUNCTION_MIN, "=SUBTOTAL(105,[Amount])"},
                {FM_TABLE_TOTALS_ROW_FUNCTION_STD_DEV, "=SUBTOTAL(107,[Amount])"},
                {FM_TABLE_TOTALS_ROW_FUNCTION_VAR, "=SUBTOTAL(110,[Amount])"}};
  for (const auto& c : kCases) {
    ASSERT_EQ(SetTableColumnTotals(reloaded.handle, 0, 0, 1, c.function, nullptr, nullptr), 0) << c.function;
    EXPECT_EQ(Formula(reloaded.handle, 3, 1), c.formula);
  }
  ASSERT_EQ(SetTableColumnTotals(reloaded.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_NONE, nullptr, nullptr), 0);
  EXPECT_EQ(Formula(reloaded.handle, 3, 1), "");
  EXPECT_EQ(Value(reloaded.handle, 3, 1).kind, FM_VALUE_KIND_BLANK);
}

TEST(FormulonCApiTable, TotalsRowFunctionOrdinalsFollowTheSchemaOrder) {
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_NONE, 0);
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_SUM, 1);
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_MIN, 2);
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_MAX, 3);
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_AVERAGE, 4);
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_COUNT, 5);
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_COUNT_NUMBERS, 6);
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_STD_DEV, 7);
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_VAR, 8);
  EXPECT_EQ(FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, 9);
}

TEST(FormulonCApiTable, ColumnTotalsRejectAFunctionOutsideTheEnumeration) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleMedium2", 1, {0, 0, 3, 1});
  ASSERT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr), 0);
  EXPECT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, -1, nullptr, nullptr), kInvalidArgument);
  EXPECT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM + 1, nullptr, nullptr),
            kInvalidArgument);
  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  ASSERT_EQ(fm_sheet_get_table_column_at(wb.handle, 0, 0, 1, &column), 0);
  EXPECT_EQ(column.totals_row_function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
}

TEST(FormulonCApiTable, ColumnTotalsRefuseAnOutOfRangeTableSheetOrColumn) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleMedium2", 1, {0, 0, 3, 1});
  EXPECT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 2, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetTableColumnTotals(wb.handle, 0, 1, 0, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(SetTableColumnTotals(wb.handle, 6, 0, 0, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  EXPECT_EQ(fm_sheet_get_table_column_at(wb.handle, 6, 0, 0, &column), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_column_at(wb.handle, 0, 1, 0, &column), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiTable, CustomTotalsFormulaRoundTrips) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleMedium2", 1, {0, 0, 3, 1});
  EXPECT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr, nullptr),
            kInvalidArgument);
  EXPECT_NE(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr, "=SUM(("), 0);
  ASSERT_EQ(
      SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr, "=SUM(Sales[Amount])*2"),
      0);
  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  ASSERT_EQ(fm_sheet_get_table_column_at(wb.handle, 0, 0, 1, &column), 0);
  EXPECT_EQ(column.totals_row_function, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM);
  EXPECT_STREQ(column.totals_row_formula, "=SUM([Amount])*2");
  EXPECT_EQ(Formula(wb.handle, 3, 1), "=SUM([Amount])*2");

  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_NE(ExtractZipEntry(saved, "xl/tables/table1.xml")
                .find("<tableColumn id=\"2\" name=\"Amount\" totalsRowFunction=\"custom\">\n"
                      "      <totalsRowFormula>SUM(Sales[Amount])*2</totalsRowFormula>\n"
                      "    </tableColumn>"),
            std::string::npos);
  WorkbookGuard reloaded;
  Load(saved, reloaded);
  ASSERT_EQ(fm_sheet_get_table_column_at(reloaded.handle, 0, 0, 1, &column), 0);
  EXPECT_EQ(column.totals_row_function, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM);
  EXPECT_STREQ(column.totals_row_formula, "=SUM([Amount])*2");
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  const fm_value_t v = Value(reloaded.handle, 3, 1);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 60.0);
}

TEST(FormulonCApiTable, TurningTheTotalsRowOnRestoresItsCells) {
  // A file whose totals row was shown, then hidden, keeps the column
  // settings; showing it again, with the range it then has, writes the
  // cells back.
  WorkbookGuard source;
  CreateSalesTable(source);
  fm_table_patch_t toggle = TablePatch();
  toggle.totals_row_op = FM_PATCH_OP_SET;
  toggle.totals_row = 1;
  ASSERT_EQ(fm_sheet_update_table(source.handle, 0, 0, &toggle), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetTableColumnTotals(source.handle, 0, 0, 0, FM_TABLE_TOTALS_ROW_FUNCTION_NONE, "Total", nullptr), 0);
  ASSERT_EQ(SetTableColumnTotals(source.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_MAX, nullptr, nullptr), 0);
  toggle.totals_row = 0;
  ASSERT_EQ(fm_sheet_update_table(source.handle, 0, 0, &toggle), 0) << fm_get_last_error_message();
  EXPECT_EQ(Value(source.handle, 3, 1).kind, FM_VALUE_KIND_BLANK) << "no cell while the totals row is off";
  WorkbookGuard wb;
  Load(Save(source.handle), wb);
  fm_table_patch_t show_totals = TablePatch();
  show_totals.range_op = FM_PATCH_OP_SET;
  show_totals.range = fm_cell_rect_t{0, 0, 3, 1};
  show_totals.totals_row_op = FM_PATCH_OP_SET;
  show_totals.totals_row = 1;
  ASSERT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &show_totals), 0) << fm_get_last_error_message();
  EXPECT_EQ(Info(wb.handle).range.last_row, 3U) << "the range is the one the toggle gives";
  EXPECT_EQ(Formula(wb.handle, 3, 1), "=SUBTOTAL(104,[Amount])");
  const fm_value_t label = Value(wb.handle, 3, 0);
  ASSERT_EQ(label.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(label.u.text, "Total");
  const std::string xml = ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml");
  EXPECT_EQ(xml.find("totalsRowShown"), std::string::npos);
  EXPECT_NE(xml.find("totalsRowCount=\"1\""), std::string::npos);
}

TEST(FormulonCApiTable, LoadedTotalsShownMarkerRoundTrips) {
  WorkbookGuard source;
  CreateSalesTable(source);
  // A table created without a totals row is saved as never having shown one.
  const std::vector<std::uint8_t> saved = Save(source.handle);
  ASSERT_NE(ExtractZipEntry(saved, "xl/tables/table1.xml").find("totalsRowShown=\"0\""), std::string::npos);
  WorkbookGuard wb;
  Load(saved, wb);
  EXPECT_NE(ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml").find("totalsRowShown=\"0\""), std::string::npos);
}

TEST(FormulonCApiTable, TotalsFormulaEscapesSpecialColumnNames) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Q1 #est"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "a,b"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 0, 4.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 5.0), 0);
  const char* columns[] = {"Q1 #est", "a,b"};
  std::size_t index = 0;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 2, 1}), "T", columns, 2, "", 1, 1, &index), 0);
  ASSERT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 0, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr), 0);
  ASSERT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr), 0);
  EXPECT_EQ(Formula(wb.handle, 2, 0), "=SUBTOTAL(109,[Q1 '#est])");
  EXPECT_EQ(Formula(wb.handle, 2, 1), "=SUBTOTAL(109,[a,b])");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v = Value(wb.handle, 2, 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 4.0);
  v = Value(wb.handle, 2, 1);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 5.0);
}

TEST(FormulonCApiTable, RenameRewritesStructuredReferences) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  const char* columns[] = {"Key"};
  std::size_t other = 0;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 5, 1, 5}), "Other", columns, 1, "", 1, 0, &other), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 3, "=SUM(Sales[Amount])+COUNTA(sales[Product])"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 3, "=Sales"), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Doubled", "=SUM(Sales[Amount])*2"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 3, "=Doubled"), 0);

  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, "other"), kAlreadyExists);
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, ""), kInvalidArgument);
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, 5, "Revenue"), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 5, 0, "Revenue"), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  // A case-only rename of the table itself is not a collision.
  ASSERT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, "SALES"), 0);
  ASSERT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, "Revenue"), 0);

  EXPECT_STREQ(Info(wb.handle).name, "Revenue");
  EXPECT_EQ(Formula(wb.handle, 0, 3), "=SUM(Revenue[Amount])+COUNTA(Revenue[Product])");
  EXPECT_EQ(Formula(wb.handle, 1, 3), "=Revenue") << "an entered bare table name binds to the whole table";
  fm_defined_name_t dn = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.formula, "=SUM(Revenue[Amount])*2");

  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v = Value(wb.handle, 0, 3);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 32.0);
  v = Value(wb.handle, 2, 3);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 60.0);

  ASSERT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, "RevenueShown"), 0);
  EXPECT_EQ(Formula(wb.handle, 0, 3), "=SUM(RevenueShown[Amount])+COUNTA(RevenueShown[Product])");

  // The file's `name` attribute follows the one public name.
  const std::vector<std::uint8_t> saved = Save(wb.handle);
  const std::string table_xml = ExtractZipEntry(saved, "xl/tables/table1.xml");
  EXPECT_NE(table_xml.find(" name=\"RevenueShown\""), std::string::npos) << table_xml;
  EXPECT_NE(table_xml.find(" displayName=\"RevenueShown\""), std::string::npos) << table_xml;
  WorkbookGuard reloaded;
  Load(saved, reloaded);
  EXPECT_STREQ(Info(reloaded.handle).name, "RevenueShown");
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  v = Value(reloaded.handle, 0, 3);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 32.0);
}

fm_dxf_t FillDxf() {
  fm_dxf_t dxf{};
  dxf.has_fill = 1;
  dxf.fill.pattern = 1;
  dxf.fill.foreground.kind = FM_COLOR_KIND_RGB;
  dxf.fill.foreground.argb = 0xFFDDEBF7U;
  return dxf;
}

TEST(FormulonCApiTableStyles, CustomStyleRoundTripsAndATableUsesIt) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  std::uint32_t dxf = 99;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, record_arg(FillDxf()), &dxf), 0);

  const fm_table_style_element_t elements[] = {
      {FM_TABLE_STYLE_ELEMENT_TYPE_WHOLE_TABLE, 1U, 1, static_cast<uint32_t>(dxf)},
      {FM_TABLE_STYLE_ELEMENT_TYPE_HEADER_ROW, 1U, 1, static_cast<uint32_t>(dxf)},
      {FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_ROW_STRIPE, 2U, 1, static_cast<uint32_t>(dxf)},
      {FM_TABLE_STYLE_ELEMENT_TYPE_TOTAL_ROW, 1U, 0, 0U},
  };
  fm_table_style_t style = MakeTableStyle("Corporate", 0, 1, elements, 4U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &style), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_table_style_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);
  fm_table_style_defaults_t defaults = FM_STRUCT_INIT(fm_table_style_defaults_t);
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &defaults), 0);
  EXPECT_STREQ(defaults.table_style_name, "TableStyleMedium2");
  EXPECT_STREQ(defaults.pivot_style_name, "PivotStyleLight16");
  fm_table_style_options_patch_t corporate = FM_STRUCT_INIT(fm_table_style_options_patch_t);
  corporate.style_name_op = FM_PATCH_OP_SET;
  corporate.style_name = "Corporate";
  corporate.show_first_column_op = FM_PATCH_OP_SET;
  corporate.show_first_column = 0;
  corporate.show_last_column_op = FM_PATCH_OP_SET;
  corporate.show_last_column = 0;
  corporate.show_row_stripes_op = FM_PATCH_OP_SET;
  corporate.show_row_stripes = 1;
  corporate.show_column_stripes_op = FM_PATCH_OP_SET;
  corporate.show_column_stripes = 0;
  ASSERT_EQ(fm_sheet_update_table_style_options(wb.handle, 0, 0, &corporate), 0);

  const std::vector<std::uint8_t> saved = Save(wb.handle);
  const std::string styles_xml = ExtractZipEntry(saved, "xl/styles.xml");
  const std::size_t dxfs_end = styles_xml.find("</dxfs>");
  const std::size_t table_styles = styles_xml.find(
      "<tableStyles count=\"1\" defaultTableStyle=\"TableStyleMedium2\" "
      "defaultPivotStyle=\"PivotStyleLight16\">");
  ASSERT_NE(dxfs_end, std::string::npos);
  ASSERT_NE(table_styles, std::string::npos);
  EXPECT_LT(dxfs_end, table_styles);
  for (const char* later : {"<colors", "<extLst"}) {
    const std::size_t at = styles_xml.find(later);
    if (at != std::string::npos) {
      EXPECT_LT(table_styles, at) << later;
    }
  }
  EXPECT_NE(styles_xml.find("<tableStyle name=\"Corporate\" pivot=\"0\" count=\"4\">"), std::string::npos);
  EXPECT_NE(
      styles_xml.find("<tableStyleElement type=\"firstRowStripe\" size=\"2\" dxfId=\"" + std::to_string(dxf) + "\"/>"),
      std::string::npos);
  EXPECT_NE(styles_xml.find("<tableStyleElement type=\"totalRow\"/>"), std::string::npos);

  WorkbookGuard reloaded;
  Load(saved, reloaded);
  EXPECT_STREQ(Info(reloaded.handle).style_name, "Corporate");
  ASSERT_EQ(fm_workbook_get_table_style_count(reloaded.handle, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_table_style_t got = FM_STRUCT_INIT(fm_table_style_t);
  ASSERT_EQ(fm_workbook_get_table_style_at(reloaded.handle, 0, &got), 0);
  EXPECT_STREQ(got.name, "Corporate");
  EXPECT_EQ(got.pivot, FM_TRISTATE_FALSE);
  EXPECT_EQ(got.table, FM_TRISTATE_TRUE);
  ASSERT_EQ(got.element_count, 4U);
  EXPECT_EQ(got.elements[2].type, FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_ROW_STRIPE);
  EXPECT_EQ(got.elements[2].size, 2U);
  EXPECT_EQ(got.elements[2].has_dxf_index, 1);
  EXPECT_EQ(got.elements[2].dxf_index, dxf);
  EXPECT_EQ(got.elements[3].has_dxf_index, 0);
  fm_dxf_t reloaded_dxf{};
  ASSERT_EQ(fm_workbook_get_dxf_at(reloaded.handle, got.elements[0].dxf_index, &reloaded_dxf), 0);
  EXPECT_EQ(reloaded_dxf.has_fill, 1);

  // Setting the record a get returned, views and all, leaves the part as it was.
  const std::string before = ExtractZipEntry(Save(reloaded.handle), "xl/styles.xml");
  ASSERT_EQ(fm_workbook_get_table_style_at(reloaded.handle, 0, &got), 0);
  ASSERT_EQ(fm_workbook_set_table_style(reloaded.handle, &got), 0);
  EXPECT_EQ(ExtractZipEntry(Save(reloaded.handle), "xl/styles.xml"), before);
}

TEST(FormulonCApiTableStyles, ElementTypeOrdinalsFollowTheSchemaOrder) {
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_WHOLE_TABLE, 0);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_HEADER_ROW, 1);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_TOTAL_ROW, 2);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_COLUMN, 3);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_LAST_COLUMN, 4);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_ROW_STRIPE, 5);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_ROW_STRIPE, 6);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_COLUMN_STRIPE, 7);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_COLUMN_STRIPE, 8);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_HEADER_CELL, 9);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_LAST_HEADER_CELL, 10);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_TOTAL_CELL, 11);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_LAST_TOTAL_CELL, 12);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_SUBTOTAL_COLUMN, 13);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_SUBTOTAL_COLUMN, 14);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_THIRD_SUBTOTAL_COLUMN, 15);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_SUBTOTAL_ROW, 16);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_SUBTOTAL_ROW, 17);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_THIRD_SUBTOTAL_ROW, 18);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_BLANK_ROW, 19);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_COLUMN_SUBHEADING, 20);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_COLUMN_SUBHEADING, 21);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_THIRD_COLUMN_SUBHEADING, 22);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_ROW_SUBHEADING, 23);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_ROW_SUBHEADING, 24);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_THIRD_ROW_SUBHEADING, 25);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_PAGE_FIELD_LABELS, 26);
  EXPECT_EQ(FM_TABLE_STYLE_ELEMENT_TYPE_PAGE_FIELD_VALUES, 27);
}

TEST(FormulonCApiTableStyles, SetReplacesByNameAndValidates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t dxf = 0;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, record_arg(FillDxf()), &dxf), 0);
  const fm_table_style_element_t one[] = {{FM_TABLE_STYLE_ELEMENT_TYPE_WHOLE_TABLE, 1U, 1, static_cast<uint32_t>(dxf)}};
  fm_table_style_t style = MakeTableStyle("Mine", 1, 1, one, 1U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &style), 0);
  const fm_table_style_element_t two[] = {{FM_TABLE_STYLE_ELEMENT_TYPE_PAGE_FIELD_LABELS, 1U, 0, 0U},
                                          {FM_TABLE_STYLE_ELEMENT_TYPE_BLANK_ROW, 1U, 1, static_cast<uint32_t>(dxf)}};
  fm_table_style_t replaced = MakeTableStyle("MINE", 1, 0, two, 2U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &replaced), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_table_style_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);
  fm_table_style_t got = FM_STRUCT_INIT(fm_table_style_t);
  ASSERT_EQ(fm_workbook_get_table_style_at(wb.handle, 0, &got), 0);
  EXPECT_STREQ(got.name, "MINE");
  EXPECT_EQ(got.table, FM_TRISTATE_FALSE);
  ASSERT_EQ(got.element_count, 2U);
  EXPECT_EQ(got.elements[0].type, FM_TABLE_STYLE_ELEMENT_TYPE_PAGE_FIELD_LABELS);

  const fm_table_style_element_t past_last[] = {{FM_TABLE_STYLE_ELEMENT_TYPE_PAGE_FIELD_VALUES + 1, 1U, 0, 0U}};
  fm_table_style_t bad = MakeTableStyle("Other", 1, 1, past_last, 1U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  const fm_table_style_element_t negative[] = {{-1, 1U, 0, 0U}};
  bad = MakeTableStyle("Other", 1, 1, negative, 1U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  const fm_table_style_element_t repeated[] = {{FM_TABLE_STYLE_ELEMENT_TYPE_WHOLE_TABLE, 1U, 0, 0U},
                                               {FM_TABLE_STYLE_ELEMENT_TYPE_WHOLE_TABLE, 1U, 0, 0U}};
  bad = MakeTableStyle("Other", 1, 1, repeated, 2U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  const fm_table_style_element_t zero_size[] = {{FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_ROW_STRIPE, 0U, 0, 0U}};
  bad = MakeTableStyle("Other", 1, 1, zero_size, 1U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  const fm_table_style_element_t dangling[] = {
      {FM_TABLE_STYLE_ELEMENT_TYPE_WHOLE_TABLE, 1U, 1, static_cast<uint32_t>(dxf + 1U)}};
  bad = MakeTableStyle("Other", 1, 1, dangling, 1U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  bad = MakeTableStyle("TableStyleMedium2", 1, 1, nullptr, 0U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  bad = MakeTableStyle("pivotstylelight16", 1, 1, nullptr, 0U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  bad = MakeTableStyle("", 1, 1, nullptr, 0U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  bad = MakeTableStyle("Other", 1, 1, nullptr, 1U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kBindingNullPointer);
  // `TableStyleMedium29` is past the built-in range, so it is a legal name.
  bad = MakeTableStyle("TableStyleMedium29", 1, 1, nullptr, 0U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), 0);
  ASSERT_EQ(fm_workbook_get_table_style_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 2U);
}

TEST(FormulonCApiTableStyles, DefaultsAndRemove) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_table_style_t style = MakeTableStyle("Mine", 1, 1, nullptr, 0U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &style), 0);
  ASSERT_EQ(UpdateStyleDefaults(wb.handle, "Mine", nullptr), 0);
  fm_table_style_defaults_t defaults = FM_STRUCT_INIT(fm_table_style_defaults_t);
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &defaults), 0);
  EXPECT_STREQ(defaults.table_style_name, "Mine");
  EXPECT_STREQ(defaults.pivot_style_name, "PivotStyleLight16");

  EXPECT_EQ(fm_workbook_remove_table_style(wb.handle, "Missing"), FM_STATUS_NOT_FOUND);
  ASSERT_EQ(fm_workbook_remove_table_style(wb.handle, "mine"), 0);
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &defaults), 0);
  EXPECT_STREQ(defaults.table_style_name, "TableStyleMedium2");

  ASSERT_EQ(UpdateStyleDefaults(wb.handle, nullptr, ""), 0);
  const std::string styles_xml = ExtractZipEntry(Save(wb.handle), "xl/styles.xml");
  EXPECT_NE(styles_xml.find("<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\"/>"), std::string::npos);
}

namespace {

/// Loads into `wb` the sales table (A1:B3) and the scratch pivot (D1) on one sheet, both naming `style`. The style
/// is created applying to both kinds and the flags are then rewritten in the saved package, as the reader takes a
/// file whose flags exclude a kind its users have.
void TableAndPivotUsing(WorkbookGuard& wb, const char* style, int32_t pivot, int32_t table) {
  WorkbookGuard source;
  CreateSalesTable(source);
  ASSERT_EQ(BuildScratchPivot(source.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  fm_table_style_t custom = MakeTableStyle(style, 1, 1, nullptr, 0U);
  ASSERT_EQ(fm_workbook_set_table_style(source.handle, &custom), 0);
  ASSERT_EQ(SetTableStyleName(source.handle, 0U, style), 0);
  fm_pivot_style_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  patch.style_name_op = FM_PATCH_OP_SET;
  patch.style_name = style;
  ASSERT_EQ(fm_sheet_update_pivot_style_options(source.handle, 0, 0, &patch), 0) << fm_get_last_error_message();
  const std::string element = std::string("<tableStyle name=\"") + style + "\"";
  const std::vector<std::uint8_t> flagged = RewriteEntry(Save(source.handle), "xl/styles.xml", [&](std::string xml) {
    return ReplaceOnce(std::move(xml), element,
                       element + (pivot != 0 ? "" : " pivot=\"0\"") + (table != 0 ? "" : " table=\"0\""));
  });
  Load(flagged, wb);
}

std::string TableStyleOf(fm_workbook_t* wb) {
  const fm_table_info_t info = Info(wb);
  return info.style_name != nullptr ? info.style_name : "";
}

std::string PivotStyleOf(const fm_workbook_t* wb) {
  fm_pivot_style_options_t style = FM_STRUCT_INIT(fm_pivot_style_options_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_pivot_style_options(wb, 0, 0, &present, &style), 0);
  return style.style_name != nullptr ? style.style_name : "";
}

}  // namespace

// Users of a removed style take their kind's default, as the desktop application re-points them.
TEST(FormulonCApiTableStyles, RemoveRepointsTablesAndPivotsToTheirDefaults) {
  WorkbookGuard wb;
  TableAndPivotUsing(wb, "Mine", 1, 1);
  ASSERT_EQ(UpdateStyleDefaults(wb.handle, "TableStyleLight9", "PivotStyleMedium4"), 0);
  ASSERT_EQ(fm_workbook_remove_table_style(wb.handle, "MINE"), 0);
  EXPECT_EQ(TableStyleOf(wb.handle), "TableStyleLight9");
  EXPECT_EQ(PivotStyleOf(wb.handle), "PivotStyleMedium4");
}

// A loaded style that excludes a kind its users have is removed like any other: every user takes its own default.
TEST(FormulonCApiTableStyles, RemoveRepointsEveryUserToItsOwnKindsDefault) {
  WorkbookGuard pivot_only;
  TableAndPivotUsing(pivot_only, "PivotOnly", 1, 0);
  ASSERT_EQ(TableStyleOf(pivot_only.handle), "PivotOnly");
  ASSERT_EQ(fm_workbook_remove_table_style(pivot_only.handle, "PivotOnly"), 0);
  EXPECT_EQ(TableStyleOf(pivot_only.handle), "TableStyleMedium2");
  EXPECT_EQ(PivotStyleOf(pivot_only.handle), "PivotStyleLight16");

  WorkbookGuard table_only;
  TableAndPivotUsing(table_only, "TableOnly", 0, 1);
  ASSERT_EQ(PivotStyleOf(table_only.handle), "TableOnly");
  ASSERT_EQ(fm_workbook_remove_table_style(table_only.handle, "TableOnly"), 0);
  EXPECT_EQ(TableStyleOf(table_only.handle), "TableStyleMedium2");
  EXPECT_EQ(PivotStyleOf(table_only.handle), "PivotStyleLight16");
}

// Removing the style that is the default of both kinds resets both defaults, and users take them.
TEST(FormulonCApiTableStyles, RemovingTheDefaultResetsItAndUsersTakeTheBuiltIn) {
  WorkbookGuard wb;
  TableAndPivotUsing(wb, "Both", 1, 1);
  ASSERT_EQ(UpdateStyleDefaults(wb.handle, "Both", "Both"), 0);
  ASSERT_EQ(fm_workbook_remove_table_style(wb.handle, "Both"), 0);
  fm_table_style_defaults_t defaults = FM_STRUCT_INIT(fm_table_style_defaults_t);
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &defaults), 0);
  EXPECT_STREQ(defaults.table_style_name, "TableStyleMedium2");
  EXPECT_STREQ(defaults.pivot_style_name, "PivotStyleLight16");
  EXPECT_EQ(TableStyleOf(wb.handle), "TableStyleMedium2");
  EXPECT_EQ(PivotStyleOf(wb.handle), "PivotStyleLight16");
  EXPECT_EQ(fm_workbook_remove_table_style(wb.handle, "Both"), FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiTableStyles, DefaultsPatchRefusesABadOpAndChangesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(UpdateStyleDefaults(wb.handle, "TableStyleLight1", nullptr), 0);
  fm_table_style_defaults_patch_t patch = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  patch.table_style_name_op = FM_PATCH_OP_SET;
  patch.table_style_name = "TableStyleLight2";
  patch.pivot_style_name_op = static_cast<fm_patch_op_t>(3);
  patch.pivot_style_name = "PivotStyleLight2";
  EXPECT_EQ(fm_workbook_update_table_style_defaults(wb.handle, &patch), kInvalidArgument);
  patch.pivot_style_name_op = FM_PATCH_OP_SET;
  patch.pivot_style_name = "\xff";
  EXPECT_EQ(fm_workbook_update_table_style_defaults(wb.handle, &patch),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingUtf8EncodingError));
  patch.struct_size = 0U;
  EXPECT_EQ(fm_workbook_update_table_style_defaults(wb.handle, &patch), kInvalidArgument);
  EXPECT_EQ(fm_workbook_update_table_style_defaults(wb.handle, nullptr), kBindingNullPointer);
  fm_table_style_defaults_t defaults = FM_STRUCT_INIT(fm_table_style_defaults_t);
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &defaults), 0);
  EXPECT_STREQ(defaults.table_style_name, "TableStyleLight1") << "a refused patch writes nothing";
}

namespace {

/// A workbook holding the sales table and the scratch pivot plus the custom styles `PivotOnly` (not for tables) and
/// `TableOnly` (not for PivotTables).
void KindStyles(WorkbookGuard& wb) {
  CreateSalesTable(wb);
  ASSERT_EQ(BuildScratchPivot(wb.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  fm_table_style_t pivot_only = MakeTableStyle("PivotOnly", 1, 0, nullptr, 0U);
  fm_table_style_t table_only = MakeTableStyle("TableOnly", 0, 1, nullptr, 0U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &pivot_only), 0);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &table_only), 0);
}

fm_status_t SetPivotStyleName(fm_workbook_t* wb, const char* style) {
  fm_pivot_style_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  patch.style_name_op = FM_PATCH_OP_SET;
  patch.style_name = style;
  return fm_sheet_update_pivot_style_options(wb, 0, 0, &patch);
}

std::size_t TableCount(fm_workbook_t* wb) {
  std::size_t count = 99U;
  EXPECT_EQ(fm_sheet_get_table_count(wb, 0, &count), 0);
  return count;
}

}  // namespace

TEST(FormulonCApiTableStyleKinds, AddTableRefusesAStyleThatExcludesTables) {
  WorkbookGuard wb;
  KindStyles(wb);
  const char* columns[] = {"A", "B"};
  const fm_cell_rect_t range = {5, 5, 7, 6};
  std::size_t index = 99U;
  EXPECT_EQ(AddTable(wb.handle, 0, &range, "Other", columns, 2, "PivotOnly", 1, 0, &index), kInvalidArgument);
  EXPECT_STREQ(fm_get_last_error_message(), "style_name must name a style that applies to tables");
  EXPECT_STREQ(fm_get_last_error_context(), "style_name=PivotOnly kind=table");
  EXPECT_EQ(TableCount(wb.handle), 1U) << "a refused add creates nothing";
  ASSERT_EQ(AddTable(wb.handle, 0, &range, "Other", columns, 2, "TableOnly", 1, 0, &index), 0);
}

TEST(FormulonCApiTableStyleKinds, UpdateStyleOptionsRefusesAStyleThatExcludesTheKind) {
  WorkbookGuard wb;
  KindStyles(wb);
  EXPECT_EQ(SetTableStyleName(wb.handle, 0, "PivotOnly"), kInvalidArgument);
  EXPECT_STREQ(fm_get_last_error_context(), "style_name=PivotOnly kind=table");
  EXPECT_EQ(TableStyleOf(wb.handle), "TableStyleMedium2");
  EXPECT_EQ(SetPivotStyleName(wb.handle, "TableOnly"), kInvalidArgument);
  EXPECT_STREQ(fm_get_last_error_message(), "style_name must name a style that applies to PivotTables");
  EXPECT_STREQ(fm_get_last_error_context(), "style_name=TableOnly kind=pivot");
  EXPECT_EQ(PivotStyleOf(wb.handle), "PivotStyleLight16");
  EXPECT_EQ(SetTableStyleName(wb.handle, 0, "TableOnly"), 0);
  EXPECT_EQ(SetPivotStyleName(wb.handle, "PivotOnly"), 0);
}

TEST(FormulonCApiTableStyleKinds, BuiltInSlicerStylesAreRefusedOnTablesAndPivots) {
  WorkbookGuard wb;
  KindStyles(wb);
  const char* columns[] = {"A", "B"};
  const fm_cell_rect_t range = {5, 5, 7, 6};
  std::size_t index = 99U;
  for (const char* name : {"SlicerStyleLight1", "slicerstyledark6", "SlicerStyleOther2"}) {
    EXPECT_EQ(AddTable(wb.handle, 0, &range, "Other", columns, 2, name, 1, 0, &index), kInvalidArgument) << name;
    EXPECT_EQ(SetTableStyleName(wb.handle, 0, name), kInvalidArgument) << name;
    EXPECT_EQ(SetPivotStyleName(wb.handle, name), kInvalidArgument) << name;
    EXPECT_EQ(UpdateStyleDefaults(wb.handle, name, nullptr), kInvalidArgument) << name;
    EXPECT_EQ(UpdateStyleDefaults(wb.handle, nullptr, name), kInvalidArgument) << name;
  }
  EXPECT_STREQ(fm_get_last_error_context(), "pivot_style_name=SlicerStyleOther2 kind=pivot");
  EXPECT_EQ(TableCount(wb.handle), 1U);
  EXPECT_EQ(TableStyleOf(wb.handle), "TableStyleMedium2");
  EXPECT_EQ(PivotStyleOf(wb.handle), "PivotStyleLight16");
  // A name outside every family is unknown, not a slicer style.
  EXPECT_EQ(SetTableStyleName(wb.handle, 0, "SlicerStyleLight7"), 0);
}

TEST(FormulonCApiTableStyleKinds, UnknownAndCrossKindBuiltInNamesAreAccepted) {
  WorkbookGuard wb;
  KindStyles(wb);
  EXPECT_EQ(SetTableStyleName(wb.handle, 0, "PivotStyleLight16"), 0);
  EXPECT_EQ(SetTableStyleName(wb.handle, 0, "NoSuchStyle"), 0);
  EXPECT_EQ(SetPivotStyleName(wb.handle, "TableStyleMedium2"), 0);
  EXPECT_EQ(SetPivotStyleName(wb.handle, "NoSuchStyle"), 0);
  EXPECT_EQ(PivotStyleOf(wb.handle), "NoSuchStyle");
}

TEST(FormulonCApiTableStyleKinds, DefaultsPatchRefusesAStyleThatExcludesTheKind) {
  WorkbookGuard wb;
  KindStyles(wb);
  EXPECT_EQ(UpdateStyleDefaults(wb.handle, "PivotOnly", nullptr), kInvalidArgument);
  EXPECT_STREQ(fm_get_last_error_message(), "table_style_name must name a style that applies to tables");
  EXPECT_EQ(UpdateStyleDefaults(wb.handle, nullptr, "TableOnly"), kInvalidArgument);
  EXPECT_STREQ(fm_get_last_error_message(), "pivot_style_name must name a style that applies to PivotTables");
  EXPECT_EQ(UpdateStyleDefaults(wb.handle, "TableOnly", "PivotOnly"), 0);
}

TEST(FormulonCApiTableStyleKinds, SetTableStyleRefusesToStrandAUser) {
  WorkbookGuard wb;
  KindStyles(wb);
  ASSERT_EQ(SetTableStyleName(wb.handle, 0, "TableOnly"), 0);
  fm_table_style_t no_tables = MakeTableStyle("TableOnly", 1, 0, nullptr, 0U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &no_tables), FM_STATUS_PRECONDITION_FAILED);
  EXPECT_STREQ(fm_get_last_error_message(), "table style is in use by a table; the new style must apply to tables");
  EXPECT_STREQ(fm_get_last_error_context(), "name=TableOnly kind=table users=1");
  fm_table_style_t still_table = MakeTableStyle("TableOnly", 0, 1, nullptr, 0U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &still_table), 0) << "flags that keep the users are accepted";

  ASSERT_EQ(SetPivotStyleName(wb.handle, "PivotOnly"), 0);
  fm_table_style_t no_pivots = MakeTableStyle("PivotOnly", 0, 1, nullptr, 0U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &no_pivots), FM_STATUS_PRECONDITION_FAILED);
  EXPECT_STREQ(fm_get_last_error_message(),
               "table style is in use by a PivotTable; the new style must apply to PivotTables");
  EXPECT_STREQ(fm_get_last_error_context(), "name=PivotOnly kind=pivot users=1");

  ASSERT_EQ(UpdateStyleDefaults(wb.handle, "TableOnly", nullptr), 0);
  ASSERT_EQ(SetTableStyleName(wb.handle, 0, "TableStyleLight1"), 0);
  fm_table_style_t no_default = MakeTableStyle("TableOnly", 1, 0, nullptr, 0U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &no_default), FM_STATUS_PRECONDITION_FAILED);
  EXPECT_STREQ(fm_get_last_error_message(),
               "table style is the default table style; the new style must apply to tables");
  EXPECT_STREQ(fm_get_last_error_context(), "name=TableOnly kind=table users=1");
  fm_table_style_t read = FM_STRUCT_INIT(fm_table_style_t);
  int32_t present = 0;
  ASSERT_EQ(fm_workbook_get_table_style(wb.handle, "TableOnly", &present, &read), 0);
  EXPECT_EQ(read.table, FM_TRISTATE_TRUE) << "a refused replacement changes nothing";
}

TEST(FormulonCApiTableStyleKinds, ALoadedDefaultThatExcludesItsKindRefusesNewObjects) {
  WorkbookGuard source;
  CreateSalesTable(source);
  ASSERT_EQ(BuildScratchPivot(source.handle, nullptr, nullptr), 0) << fm_get_last_error_message();
  fm_table_style_t both = MakeTableStyle("Both", 1, 1, nullptr, 0U);
  ASSERT_EQ(fm_workbook_set_table_style(source.handle, &both), 0);
  ASSERT_EQ(UpdateStyleDefaults(source.handle, "Both", "Both"), 0);
  const std::vector<std::uint8_t> saved = Save(source.handle);
  const auto flagged = [&](const char* flag) {
    return RewriteEntry(saved, "xl/styles.xml", [&](std::string xml) {
      return ReplaceOnce(std::move(xml), "<tableStyle name=\"Both\"", std::string("<tableStyle name=\"Both\" ") + flag);
    });
  };

  WorkbookGuard no_tables;
  Load(flagged("table=\"0\""), no_tables);
  const char* columns[] = {"A", "B"};
  const fm_cell_rect_t range = {5, 5, 7, 6};
  std::size_t index = 99U;
  EXPECT_EQ(AddTable(no_tables.handle, 0, &range, "Other", columns, 2, nullptr, 1, 0, &index),
            FM_STATUS_PRECONDITION_FAILED);
  EXPECT_STREQ(fm_get_last_error_message(), "the workbook's default table style does not apply to tables");
  EXPECT_STREQ(fm_get_last_error_context(), "default=Both kind=table");
  EXPECT_EQ(TableCount(no_tables.handle), 1U);
  EXPECT_EQ(AddTable(no_tables.handle, 0, &range, "Other", columns, 2, "TableStyleLight1", 1, 0, &index), 0);
  ASSERT_EQ(UpdateStyleDefaults(no_tables.handle, "TableStyleLight1", nullptr), 0)
      << "recoverable through the defaults";

  WorkbookGuard no_pivots;
  Load(flagged("pivot=\"0\""), no_pivots);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(no_pivots.handle, 0, &cache_id), 0);
  fm_pivot_spec_t spec = FM_STRUCT_INIT(fm_pivot_spec_t);
  spec.name = "Second";
  spec.cache_id = cache_id;
  spec.anchor_row = 20;
  spec.anchor_col = 0;
  std::size_t pivot_index = 99U;
  EXPECT_EQ(fm_sheet_add_pivot(no_pivots.handle, 0, &spec, &pivot_index), FM_STATUS_PRECONDITION_FAILED);
  EXPECT_STREQ(fm_get_last_error_message(), "the workbook's default PivotTable style does not apply to PivotTables");
  EXPECT_STREQ(fm_get_last_error_context(), "default=Both kind=pivot");
}

TEST(FormulonCApiTableStyleKinds, ALoadedSlicerStyleNameIsKeptThroughASave) {
  WorkbookGuard source;
  CreateSalesTable(source, "TableStyleLight9");
  const std::vector<std::uint8_t> patched =
      RewriteEntry(Save(source.handle), "xl/tables/table1.xml", [](std::string xml) {
        return ReplaceOnce(std::move(xml), "name=\"TableStyleLight9\"", "name=\"SlicerStyleLight1\"");
      });
  WorkbookGuard wb;
  Load(patched, wb);
  EXPECT_STREQ(Info(wb.handle).style_name, "SlicerStyleLight1");
  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_NE(ExtractZipEntry(saved, "xl/tables/table1.xml").find("<tableStyleInfo name=\"SlicerStyleLight1\""),
            std::string::npos);
}

TEST(FormulonCApiTableStyleKinds, SetTableStyleRefusesEveryBuiltInFamilyName) {
  WorkbookGuard wb;
  KindStyles(wb);
  for (const char* name :
       {"SlicerStyleLight1", "slicerstyleother2", "SlicerStyleDark6", "TableStyleMedium9", "PivotStyleLight16"}) {
    fm_table_style_t style = MakeTableStyle(name, 1, 1, nullptr, 0U);
    EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &style), kInvalidArgument) << name;
  }
  std::size_t count = 99U;
  ASSERT_EQ(fm_workbook_get_table_style_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 2U);
}

namespace {

/// Loads a package whose custom style `Mine` and the sales table naming it are renamed to `builtin`.
void LoadCustomStyleNamedLike(const char* builtin, WorkbookGuard& wb) {
  WorkbookGuard source;
  CreateSalesTable(source);
  fm_table_style_t mine = MakeTableStyle("Mine", 1, 1, nullptr, 0U);
  ASSERT_EQ(fm_workbook_set_table_style(source.handle, &mine), 0);
  ASSERT_EQ(SetTableStyleName(source.handle, 0, "Mine"), 0);
  const std::string quoted = std::string("name=\"") + builtin + "\"";
  std::vector<std::uint8_t> bytes = RewriteEntry(Save(source.handle), "xl/styles.xml", [&](std::string xml) {
    return ReplaceOnce(std::move(xml), "<tableStyle name=\"Mine\"", "<tableStyle " + quoted);
  });
  bytes = RewriteEntry(bytes, "xl/tables/table1.xml",
                       [&](std::string xml) { return ReplaceOnce(std::move(xml), "name=\"Mine\"", quoted); });
  Load(bytes, wb);
}

}  // namespace

TEST(FormulonCApiTableStyleKinds, ALoadedCustomStyleNamedLikeABuiltInIsDroppedAndTheUsersKeepTheName) {
  for (const char* builtin : {"SlicerStyleLight1", "slicerstyledark3", "TableStyleMedium9"}) {
    WorkbookGuard wb;
    LoadCustomStyleNamedLike(builtin, wb);
    std::size_t count = 99U;
    ASSERT_EQ(fm_workbook_get_table_style_count(wb.handle, &count), 0) << builtin;
    EXPECT_EQ(count, 0U) << builtin;
    EXPECT_EQ(TableStyleOf(wb.handle), builtin);
    const std::vector<std::uint8_t> saved = Save(wb.handle);
    EXPECT_NE(ExtractZipEntry(saved, "xl/tables/table1.xml").find(std::string("<tableStyleInfo name=\"") + builtin),
              std::string::npos)
        << builtin;
  }
}

/// An `.xlsb` package whose only table `Sales` (id `id`) stays a verbatim
/// part: a record the model has no field for sits before BrtEndList.
std::vector<std::uint8_t> RetainedSalesPackage(std::uint32_t id) {
  constexpr std::uint16_t kBrtEndList = 344;
  formulon::Workbook book = formulon::Workbook::create_empty();
  book.add_sheet("Sheet1");
  formulon::TableMetadata table;
  table.id = id;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:B3";
  table.header_row = true;
  table.columns = {formulon::TableColumn(1, "Product", "", "", ""), formulon::TableColumn(2, "Amount", "", "", "")};
  table.style_info.present = true;
  table.style_info.name = "TableStyleLight1";
  book.mutable_tables().push_back(table);
  auto saved = formulon::io::xlsb::write_xlsb_with_result(book);
  EXPECT_TRUE(static_cast<bool>(saved));
  if (!saved) {
    return {};
  }
  return RewriteEntry(saved.value().bytes, "xl/tables/table" + std::to_string(id) + ".bin", [](std::string body) {
    const std::vector<std::uint8_t> part(body.begin(), body.end());
    std::vector<std::uint8_t> patched;
    for (const formulon::io::xlsb::FramedRecord& rec : formulon::io::xlsb::split_records(part)) {
      if (rec.type == kBrtEndList) {
        formulon::io::xlsb::emit_record(patched, 35, std::vector<std::uint8_t>{0x02, 0x0E, 0x00, 0x00});
        formulon::io::xlsb::emit_record(patched, 5000, std::vector<std::uint8_t>{0x01, 0x01, 0x00});
        formulon::io::xlsb::emit_record(patched, 36, formulon::io::ByteSpan{});
      }
      formulon::io::xlsb::append_record(patched, part, rec);
    }
    return std::string(patched.begin(), patched.end());
  });
}

TEST(FormulonCApiTable, CreateRejectsTheNameOfAVerbatimXlsbTable) {
  WorkbookGuard wb;
  Load(RetainedSalesPackage(3), wb);
  ASSERT_EQ(count_of(fm_sheet_get_table_count, wb.handle, std::size_t{0}), 0U);
  const char* columns[] = {"A", "B"};
  std::size_t index = 0;
  EXPECT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "sales", columns, 2, nullptr, 1, 0, &index),
            kAlreadyExists);
  EXPECT_EQ(count_of(fm_sheet_get_table_count, wb.handle, std::size_t{0}), 0U);
  EXPECT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "Fresh", columns, 2, nullptr, 1, 0, &index),
            0);
}

TEST(FormulonCApiTable, RenameRejectsTheNameOfAVerbatimXlsbTable) {
  WorkbookGuard wb;
  Load(RetainedSalesPackage(3), wb);
  const char* columns[] = {"A", "B"};
  std::size_t index = 0;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "Fresh", columns, 2, nullptr, 1, 0, &index),
            0);
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, index, "Sales"), kAlreadyExists);
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, index, "sALES"), kAlreadyExists);
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, index, "Fresh2"), 0);
}

TEST(FormulonCApiTable, ACellReferenceShapedNameIsRefusedOnAddAndRename) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  const char* columns[] = {"A", "B"};
  // Excel's Table Name box refuses a name that reads as a cell reference in either notation.
  for (const char* name : {"T1", "A1", "R1C1", "XFD1", "xfd1048576", "r1"}) {
    SCOPED_TRACE(name);
    std::size_t index = 0;
    EXPECT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), name, columns, 2, nullptr, 1, 0, &index),
              kInvalidArgument);
    EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, name), kInvalidArgument);
  }
  EXPECT_EQ(count_of(fm_sheet_get_table_count, wb.handle, std::size_t{0}), 1U);
  // Past the grid the same letters are an ordinary name.
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, 0, "XFE1"), 0);
}

TEST(FormulonCApiTable, AddUpdateAndRenameRefuseWhatExcelRefuses) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  const char* columns[] = {"A", "B"};
  std::size_t index = 0;
  EXPECT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "T1", columns, 2, nullptr, 1, 0, &index),
            kInvalidArgument);
  EXPECT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{2, 1, 4, 2}), "Other", nullptr, 0, nullptr, 1, 0, &index),
            FM_STATUS_PRECONDITION_FAILED)
      << "overlaps Sales";
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "Other", columns, 2, nullptr, 1, 0, &index),
            0);
  EXPECT_EQ(UpdateTableRange(wb.handle, 0, index, fm_cell_rect_t{0, 1, 2, 2}), FM_STATUS_PRECONDITION_FAILED)
      << "overlaps Sales";
  // A resize keeps the first row and, under a header row, one data row.
  EXPECT_EQ(UpdateTableRange(wb.handle, 0, index, fm_cell_rect_t{1, 4, 3, 5}), kInvalidArgument);
  EXPECT_EQ(UpdateTableRange(wb.handle, 0, index, fm_cell_rect_t{0, 4, 0, 5}), kInvalidArgument);
  EXPECT_EQ(UpdateTableRange(wb.handle, 0, index, fm_cell_rect_t{0, 4, 3, 5}), 0);
  EXPECT_EQ(fm_sheet_rename_table(wb.handle, 0, index, "R2D2"), kInvalidArgument);
  EXPECT_EQ(count_of(fm_sheet_get_table_count, wb.handle, std::size_t{0}), 2U);
}

TEST(FormulonCApiTable, UpdateWritesOnlyTheFieldsItSets) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_table_patch_t patch = TablePatch();
  ASSERT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), 0) << "an all-KEEP patch is a no-op";
  fm_table_info_t info = Info(wb.handle);
  EXPECT_EQ(info.range.last_row, 2U);
  EXPECT_EQ(info.header_row_count, 1U);
  EXPECT_EQ(info.totals_row_count, 0U);

  patch.totals_row_op = FM_PATCH_OP_SET;
  patch.totals_row = 1;
  ASSERT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), 0) << fm_get_last_error_message();
  info = Info(wb.handle);
  EXPECT_EQ(info.totals_row_count, 1U);
  EXPECT_EQ(info.header_row_count, 1U) << "KEEP leaves the header row";
  EXPECT_EQ(info.range.last_row, 3U) << "the totals row joins the range";
  EXPECT_STREQ(info.style_name, "TableStyleMedium2");
}

TEST(FormulonCApiTable, TotalsRowOnGrowsTheRangeAndKeepsTheDataRows) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  ASSERT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr), 0);
  fm_table_patch_t patch = TablePatch();
  patch.totals_row_op = FM_PATCH_OP_SET;
  patch.totals_row = 1;
  ASSERT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), 0) << fm_get_last_error_message();
  const fm_table_info_t info = Info(wb.handle);
  EXPECT_EQ(info.range.first_row, 0U);
  EXPECT_EQ(info.range.last_row, 3U);
  EXPECT_EQ(info.totals_row_count, 1U);
  const fm_value_t last_data = Value(wb.handle, 2, 1);
  ASSERT_EQ(last_data.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(last_data.u.number, 20.0);
  EXPECT_EQ(Formula(wb.handle, 3, 1), "=SUBTOTAL(109,[Amount])");

  // A patch that changes nothing is a no-op.
  ASSERT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), 0);
  EXPECT_EQ(Info(wb.handle).range.last_row, 3U);
}

TEST(FormulonCApiTable, TotalsRowOffShrinksTheRangeAndClearsTheRow) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  ASSERT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr), 0);
  fm_table_patch_t patch = TablePatch();
  patch.totals_row_op = FM_PATCH_OP_SET;
  patch.totals_row = 1;
  ASSERT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), 0);
  patch.totals_row = 0;
  ASSERT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), 0) << fm_get_last_error_message();
  const fm_table_info_t info = Info(wb.handle);
  EXPECT_EQ(info.range.last_row, 2U);
  EXPECT_EQ(info.totals_row_count, 0U);
  EXPECT_EQ(Value(wb.handle, 3, 1).kind, FM_VALUE_KIND_BLANK);
}

TEST(FormulonCApiTable, TotalsRowOnShiftsAnOccupiedRowBelowAndAppliesThePatch) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 3, 0, "x"), 0);
  fm_table_patch_t patch = TablePatch();
  patch.header_row_op = FM_PATCH_OP_SET;
  patch.header_row = 0;
  patch.totals_row_op = FM_PATCH_OP_SET;
  patch.totals_row = 1;
  ASSERT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), 0) << fm_get_last_error_message();
  const fm_table_info_t info = Info(wb.handle);
  EXPECT_EQ(info.range.first_row, 1U) << "hiding the header moves the first row down";
  EXPECT_EQ(info.range.last_row, 3U);
  EXPECT_EQ(info.header_row_count, 0U);
  EXPECT_EQ(info.totals_row_count, 1U);
  EXPECT_EQ(Value(wb.handle, 0, 0).kind, FM_VALUE_KIND_BLANK) << "the header cells are deleted";
  const fm_value_t occupant = Value(wb.handle, 4, 0);
  ASSERT_EQ(occupant.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(occupant.u.text, "x");
}

TEST(FormulonCApiTable, TotalsRowOnRefusesAnOccupiedRowAboveAStraddlingMergeAndChangesNothing) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 3, 0, "x"), 0);
  const fm_cell_rect_t merge{4, 1, 4, 2};  // B5:C5 straddles the table's columns A:B.
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, &merge, nullptr), 0);
  fm_table_patch_t patch = TablePatch();
  patch.header_row_op = FM_PATCH_OP_SET;
  patch.header_row = 0;
  patch.totals_row_op = FM_PATCH_OP_SET;
  patch.totals_row = 1;
  EXPECT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), FM_STATUS_PRECONDITION_FAILED);
  const fm_table_info_t info = Info(wb.handle);
  EXPECT_EQ(info.range.last_row, 2U);
  EXPECT_EQ(info.header_row_count, 1U) << "the header toggle was not applied either";
  EXPECT_EQ(info.totals_row_count, 0U);
  const fm_value_t header = Value(wb.handle, 0, 0);
  ASSERT_EQ(header.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(header.u.text, "Product");
  const fm_value_t occupant = Value(wb.handle, 3, 0);
  ASSERT_EQ(occupant.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(occupant.u.text, "x");
}

TEST(FormulonCApiTable, TotalsRowOnRefusesATableOnTheLastGridRow) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleMedium2", 0, fm_cell_rect_t{1048574U, 0, 1048575U, 1});
  fm_table_patch_t patch = TablePatch();
  patch.totals_row_op = FM_PATCH_OP_SET;
  patch.totals_row = 1;
  EXPECT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), FM_STATUS_PRECONDITION_FAILED);
  const fm_table_info_t info = Info(wb.handle);
  EXPECT_EQ(info.range.last_row, 1048575U);
  EXPECT_EQ(info.totals_row_count, 0U);
}

TEST(FormulonCApiTable, ColumnRenameFoldsNonAsciiCase) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  const char* name = nullptr;
  ASSERT_EQ(fm_sheet_rename_table_column(wb.handle, 0, 0, 0, "\xC3\x84", &name), 0);
  EXPECT_STREQ(name, "\xC3\x84");
  ASSERT_EQ(fm_sheet_rename_table_column(wb.handle, 0, 0, 1, "\xC3\xA4", &name), 0) << fm_get_last_error_message();
  EXPECT_STREQ(name,
               "\xC3\xA4"
               "2")
      << "a duplicate under the fold takes a number";
}

TEST(FormulonCApiTableStyles, CustomStyleNamesFoldNonAsciiCase) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_table_style_t upper = MakeTableStyle(
      "\xC3\x84"
      "1",
      0, 1, nullptr, 0U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &upper), 0) << fm_get_last_error_message();
  const fm_table_style_t lower = MakeTableStyle(
      "\xC3\xA4"
      "1",
      0, 1, nullptr, 0U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &lower), 0) << fm_get_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_table_style_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U) << "the second set replaces the first";
  EXPECT_EQ(fm_workbook_remove_table_style(wb.handle,
                                           "\xC3\x84"
                                           "1"),
            0);
}

TEST(FormulonCApiTable, UpdateRefusesABadPatchAndChangesNothing) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  fm_table_patch_t clear_range = TablePatch();
  clear_range.range_op = FM_PATCH_OP_CLEAR;
  fm_table_patch_t clear_header = TablePatch();
  clear_header.header_row_op = FM_PATCH_OP_CLEAR;
  fm_table_patch_t bad_op = TablePatch();
  bad_op.totals_row_op = static_cast<fm_patch_op_t>(3);
  bad_op.totals_row = 1;
  fm_table_patch_t bad_flag = TablePatch();
  bad_flag.header_row_op = FM_PATCH_OP_SET;
  bad_flag.header_row = 2;
  // A valid totals-row SET beside a range of another width: nothing is written.
  fm_table_patch_t bad_width = TablePatch();
  bad_width.totals_row_op = FM_PATCH_OP_SET;
  bad_width.totals_row = 1;
  bad_width.range_op = FM_PATCH_OP_SET;
  bad_width.range = fm_cell_rect_t{0, 0, 3, 2};
  for (const fm_table_patch_t* rejected : {&clear_range, &clear_header, &bad_op, &bad_flag, &bad_width}) {
    EXPECT_EQ(fm_sheet_update_table(wb.handle, 0, 0, rejected), kInvalidArgument);
  }
  fm_table_patch_t unsized = TablePatch();
  unsized.struct_size = 0U;
  EXPECT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &unsized), kInvalidArgument);
  EXPECT_EQ(fm_sheet_update_table(wb.handle, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_table(wb.handle, 0, 4, &bad_flag), FM_STATUS_INDEX_OUT_OF_RANGE);
  const fm_table_info_t info = Info(wb.handle);
  EXPECT_EQ(info.range.last_row, 2U);
  EXPECT_EQ(info.range.last_col, 1U);
  EXPECT_EQ(info.header_row_count, 1U);
  EXPECT_EQ(info.totals_row_count, 0U);
}

TEST(FormulonCApiTable, CreateAvoidsTheIdOfAVerbatimXlsbTable) {
  WorkbookGuard wb;
  Load(RetainedSalesPackage(1), wb);
  const char* columns[] = {"A", "B"};
  std::size_t index = 0;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "Fresh", columns, 2, nullptr, 1, 0, &index),
            0);
  // The xlsx part is named after the table id.
  const std::string table_xml = ExtractZipEntry(Save(wb.handle), "xl/tables/table2.xml");
  EXPECT_NE(table_xml.find(" id=\"2\""), std::string::npos) << table_xml;
  EXPECT_NE(table_xml.find("name=\"Fresh\""), std::string::npos) << table_xml;
}

std::uint32_t TableRuntimeId(const fm_workbook_t* wb, std::size_t sheet, std::size_t index) {
  std::uint32_t id = 0;
  EXPECT_EQ(fm_sheet_get_table_runtime_id(wb, sheet, index, &id), 0) << fm_get_last_error_message();
  return id;
}

std::string TableName(const fm_workbook_t* wb, std::size_t sheet, std::size_t index) {
  fm_table_info_t info = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_sheet_get_table_at(wb, sheet, index, &info), 0) << fm_get_last_error_message();
  return info.name != nullptr ? info.name : "";
}

/// The name of the table `id` resolves to, or "" when it resolves to none.
std::string ResolvedTableName(const fm_workbook_t* wb, std::uint32_t id) {
  std::size_t sheet = 0;
  std::size_t index = 0;
  return fm_workbook_resolve_table_runtime_id(wb, id, &sheet, &index) == 0 ? TableName(wb, sheet, index)
                                                                           : std::string();
}

/// Adds a one-column table `name` on `sheet` over rows `row`..`row + 1` of column A.
void AddOneColumnTable(fm_workbook_t* wb, std::size_t sheet, std::uint32_t row, const char* name) {
  const char* columns[] = {"Value"};
  ASSERT_EQ(fm_sheet_set_text(wb, sheet, row, 0, "Value"), 0);
  ASSERT_EQ(AddTable(wb, sheet, record_arg(fm_cell_rect_t{row, 0, row + 1U, 0}), name, columns, 1, nullptr, 1, 0,
                     ignored_index()),
            0)
      << fm_get_last_error_message();
}

TEST(FormulonCApiTable, TableIndexIsLocalToItsSheet) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  std::size_t other = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", append_position(wb.handle), &other), 0)
      << fm_get_last_error_message();
  const char* columns[] = {"Value"};
  ASSERT_EQ(fm_sheet_set_text(wb.handle, other, 0, 0, "Value"), 0);
  std::size_t local = 99;
  ASSERT_EQ(
      AddTable(wb.handle, other, record_arg(fm_cell_rect_t{0, 0, 1, 0}), "OnOther", columns, 1, nullptr, 1, 0, &local),
      0);
  EXPECT_EQ(local, 0U) << "the first table of a sheet has index 0 whatever other sheets hold";

  EXPECT_EQ(count_of(fm_sheet_get_table_count, wb.handle, std::size_t{0}), 1U);
  EXPECT_EQ(count_of(fm_sheet_get_table_count, wb.handle, other), 1U);
  EXPECT_EQ(TableName(wb.handle, 0, 0), "Sales");
  EXPECT_EQ(TableName(wb.handle, other, 0), "OnOther");

  fm_table_info_t info = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_sheet_get_table_at(wb.handle, 0, 1, &info), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_at(wb.handle, 9, 0, &info), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  std::size_t count = 7;
  EXPECT_EQ(fm_sheet_get_table_count(wb.handle, 9, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_count(wb.handle, 0, nullptr), kBindingNullPointer);

  // Removing the table of one sheet leaves the other sheet's index alone.
  ASSERT_EQ(fm_sheet_remove_table(wb.handle, 0, 0), 0);
  EXPECT_EQ(count_of(fm_sheet_get_table_count, wb.handle, std::size_t{0}), 0U);
  EXPECT_EQ(TableName(wb.handle, other, 0), "OnOther");
  EXPECT_EQ(fm_sheet_remove_table(wb.handle, 0, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_remove_table(wb.handle, 9, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiTable, AddAssignsTheNextSheetLocalIndex) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  const char* columns[] = {"Key"};
  std::size_t index = 99;
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{0, 5, 1, 5}), "Second", columns, 1, "", 1, 0, &index), 0);
  EXPECT_EQ(index, 1U);
  ASSERT_EQ(AddTable(wb.handle, 0, record_arg(fm_cell_rect_t{5, 5, 6, 5}), "Third", columns, 1, "", 1, 0, &index), 0);
  EXPECT_EQ(index, 2U);
  ASSERT_EQ(fm_sheet_remove_table(wb.handle, 0, 0), 0);
  EXPECT_EQ(TableName(wb.handle, 0, 0), "Second") << "the tables after a removed one move down";
  EXPECT_EQ(TableName(wb.handle, 0, 1), "Third");
}

TEST(FormulonCApiTable, FindTableLooksUpByNameAcrossTheWorkbook) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  std::size_t other = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", append_position(wb.handle), &other), 0);
  AddOneColumnTable(wb.handle, other, 0U, "Tracked");
  AddOneColumnTable(wb.handle, other, 4U, "Later");

  std::size_t sheet = 99;
  std::size_t index = 99;
  int32_t present = -1;
  ASSERT_EQ(fm_workbook_find_table(wb.handle, "Sales", &present, &sheet, &index), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(sheet, 0U);
  EXPECT_EQ(index, 0U);

  ASSERT_EQ(fm_workbook_find_table(wb.handle, "later", &present, &sheet, &index), 0) << "case-insensitive";
  EXPECT_EQ(present, 1);
  EXPECT_EQ(sheet, other);
  EXPECT_EQ(index, 1U);

  sheet = 99;
  index = 99;
  ASSERT_EQ(fm_workbook_find_table(wb.handle, "Nope", &present, &sheet, &index), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(sheet, 99U) << "the indices are left unchanged when absent";
  EXPECT_EQ(index, 99U);

  EXPECT_EQ(fm_workbook_find_table(wb.handle, nullptr, &present, &sheet, &index), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_find_table(wb.handle, "Sales", &present, nullptr, &index), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_find_table(wb.handle, "Sales", &present, &sheet, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_find_table(wb.handle, "Sales", nullptr, &sheet, &index), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_find_table(nullptr, "Sales", &present, &sheet, &index), kBindingNullPointer);
}

TEST(FormulonCApiTable, AddRefusesATakenNameWithAlreadyExists) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  std::size_t other = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", append_position(wb.handle), &other), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, other, 0, 0, "Value"), 0);
  const char* columns[] = {"Value"};
  std::size_t index = 99;
  // The workbook-wide uniqueness covers a table on another sheet and ignores case.
  EXPECT_EQ(
      AddTable(wb.handle, other, record_arg(fm_cell_rect_t{0, 0, 1, 0}), "SALES", columns, 1, nullptr, 1, 0, &index),
      kAlreadyExists);
  EXPECT_EQ(count_of(fm_sheet_get_table_count, wb.handle, other), 0U);
  EXPECT_EQ(AddTable(wb.handle, 99, record_arg(fm_cell_rect_t{0, 0, 1, 0}), "Fine", columns, 1, nullptr, 1, 0, &index),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiTable, RuntimeIdFollowsTheTableAcrossListRebuilds) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  std::size_t other = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", append_position(wb.handle), &other), 0)
      << fm_get_last_error_message();
  AddOneColumnTable(wb.handle, other, 0U, "Tracked");
  const std::uint32_t tracked = TableRuntimeId(wb.handle, other, 0);
  const std::uint32_t sales = TableRuntimeId(wb.handle, 0, 0);
  EXPECT_NE(tracked, 0U);
  EXPECT_NE(sales, 0U);
  EXPECT_NE(tracked, sales);
  EXPECT_EQ(TableRuntimeId(wb.handle, other, 0), tracked);

  std::size_t sheet = 99;
  std::size_t index = 99;
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, 0U, &sheet, &index), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(sheet, 99U);
  EXPECT_EQ(index, 99U);
  ASSERT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, tracked, &sheet, &index), 0);
  EXPECT_EQ(sheet, other);
  EXPECT_EQ(index, 0U);

  // Enough appends to reallocate the list.
  for (std::uint32_t i = 0; i < 12U; ++i) {
    const std::string name = "Extra" + std::to_string(i);
    AddOneColumnTable(wb.handle, other, 3U * (i + 1U), name.c_str());
  }
  EXPECT_EQ(ResolvedTableName(wb.handle, tracked), "Tracked");

  // A row deletion on another sheet rebuilds the list.
  ASSERT_EQ(fm_sheet_delete_rows(wb.handle, 0, 5U, 6U), 0) << fm_get_last_error_message();
  EXPECT_EQ(ResolvedTableName(wb.handle, tracked), "Tracked");

  // Removing another sheet, with a table of its own, rebuilds it too.
  std::size_t doomed = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Doomed", append_position(wb.handle), &doomed), 0)
      << fm_get_last_error_message();
  AddOneColumnTable(wb.handle, doomed, 0U, "Gone");
  const std::uint32_t gone = TableRuntimeId(wb.handle, doomed, 0);
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, doomed), 0) << fm_get_last_error_message();
  EXPECT_EQ(ResolvedTableName(wb.handle, tracked), "Tracked");
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, gone, &sheet, &index), FM_STATUS_NOT_FOUND);

  // Removing the preceding table on another sheet does not move the tracked one.
  ASSERT_EQ(fm_sheet_remove_table(wb.handle, 0, 0), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, tracked, &sheet, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(sheet, other);
  EXPECT_EQ(index, 0U);
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, sales, &sheet, &index), FM_STATUS_NOT_FOUND);

  // Removing the preceding table on the same sheet shifts the tracked one down by one.
  ASSERT_EQ(fm_sheet_rename_table(wb.handle, other, 0, "Renamed"), 0) << fm_get_last_error_message();
  EXPECT_EQ(ResolvedTableName(wb.handle, tracked), "Renamed") << "rewriting the table in place keeps its identity";
  const std::uint32_t extra0 = TableRuntimeId(wb.handle, other, 1);
  ASSERT_EQ(fm_sheet_remove_table(wb.handle, other, 0), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, extra0, &sheet, &index), 0);
  EXPECT_EQ(sheet, other);
  EXPECT_EQ(index, 0U);
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, tracked, &sheet, &index), FM_STATUS_NOT_FOUND);

  std::uint32_t out = 0;
  EXPECT_EQ(fm_sheet_get_table_runtime_id(wb.handle, other, 99, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_runtime_id(wb.handle, 99, 0, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_runtime_id(wb.handle, other, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, extra0, nullptr, &index), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, extra0, &sheet, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiTable, RuntimeIdsAreUniqueAcrossLoadAddCopyAndRemove) {
  WorkbookGuard source;
  CreateSalesTable(source);
  AddOneColumnTable(source.handle, 0, 10U, "Second");
  WorkbookGuard wb;
  Load(Save(source.handle), wb);

  std::vector<std::uint32_t> seen;
  const auto expect_new = [&seen](std::uint32_t id) {
    EXPECT_NE(id, 0U);
    for (const std::uint32_t earlier : seen) {
      EXPECT_NE(id, earlier);
    }
    seen.push_back(id);
  };
  // Loaded tables.
  expect_new(TableRuntimeId(wb.handle, 0, 0));
  expect_new(TableRuntimeId(wb.handle, 0, 1));

  // Added.
  AddOneColumnTable(wb.handle, 0, 20U, "Third");
  expect_new(TableRuntimeId(wb.handle, 0, 2));

  // Copied: the copies are new tables; the originals keep their ids.
  std::size_t copy = 99;
  ASSERT_EQ(fm_workbook_copy_sheet(wb.handle, 0, 1, "Copy", &copy), 0) << fm_get_last_error_message();
  ASSERT_EQ(count_of(fm_sheet_get_table_count, wb.handle, std::size_t{0}), 3U);
  ASSERT_EQ(count_of(fm_sheet_get_table_count, wb.handle, copy), 3U);
  for (std::size_t i = 0; i < 3U; ++i) {
    EXPECT_EQ(TableRuntimeId(wb.handle, 0, i), seen[i]);
  }
  for (std::size_t i = 0; i < 3U; ++i) {
    expect_new(TableRuntimeId(wb.handle, copy, i));
  }

  // Added after a removal: the removed id is not reused.
  const std::uint32_t removed = TableRuntimeId(wb.handle, 0, 2);
  ASSERT_EQ(fm_sheet_remove_table(wb.handle, 0, 2), 0) << fm_get_last_error_message();
  AddOneColumnTable(wb.handle, 0, 30U, "Fourth");
  const std::size_t count = count_of(fm_sheet_get_table_count, wb.handle, std::size_t{0});
  const std::uint32_t fresh = TableRuntimeId(wb.handle, 0, count - 1U);
  EXPECT_NE(fresh, removed);
  expect_new(fresh);
}

TEST(FormulonCApiTableStyleDefaultsRecord, ReadsRecord) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_table_style_defaults_t defaults = FM_STRUCT_INIT(fm_table_style_defaults_t);
  ASSERT_EQ(UpdateStyleDefaults(wb.handle, "TableStyleLight9", "PivotStyleMedium4"), 0);
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &defaults), 0);
  EXPECT_EQ(defaults.struct_size, sizeof(fm_table_style_defaults_t));
  EXPECT_STREQ(defaults.table_style_name, "TableStyleLight9");
  EXPECT_STREQ(defaults.pivot_style_name, "PivotStyleMedium4");
  ASSERT_EQ(UpdateStyleDefaults(wb.handle, "", nullptr), 0);
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &defaults), 0);
  EXPECT_EQ(defaults.table_style_name, nullptr) << "a cleared default reads NULL";
  EXPECT_STREQ(defaults.pivot_style_name, "PivotStyleMedium4");
  EXPECT_EQ(fm_workbook_get_table_style_defaults(wb.handle, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiTableStyleUsableDefault, ZeroInitStyleUsableByBoth) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_table_style_t style = FM_STRUCT_INIT(fm_table_style_t);
  style.name = "ZeroInit";
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &style), 0) << fm_get_last_error_message();
  fm_table_style_t got = FM_STRUCT_INIT(fm_table_style_t);
  std::int32_t present = 0;
  ASSERT_EQ(fm_workbook_get_table_style(wb.handle, "ZeroInit", &present, &got), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(got.pivot, FM_TRISTATE_TRUE) << "DEFAULT is usable, and a read states it";
  EXPECT_EQ(got.table, FM_TRISTATE_TRUE);

  style.pivot = FM_TRISTATE_FALSE;
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &style), 0);
  ASSERT_EQ(fm_workbook_get_table_style(wb.handle, "ZeroInit", &present, &got), 0);
  EXPECT_EQ(got.pivot, FM_TRISTATE_FALSE);
  EXPECT_EQ(got.table, FM_TRISTATE_TRUE);

  style.table = FM_TRISTATE_TRUE + 1;
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &style), kInvalidArgument);
}

TEST(FormulonCApiEmptyStyleName, TableSpecPivotSlicerDefaultsRefuseEmpty) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "Amount"), 0);
  const char* columns[] = {"Product", "Amount"};
  fm_table_spec_t spec = FM_STRUCT_INIT(fm_table_spec_t);
  spec.range = fm_cell_rect_t{0, 0, 2, 1};
  spec.name = "Sales";
  spec.column_names = columns;
  spec.column_count = 2;
  spec.style_name = "";
  std::size_t index = 99;
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), kInvalidArgument) << "an empty style name is refused";
  spec.styled = FM_TRISTATE_FALSE;
  spec.style_name = "TableStyleLight1";
  EXPECT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), kInvalidArgument)
      << "an unstyled table takes no style name";
  EXPECT_EQ(index, 99U);

  // `styled` FALSE with no name creates the table without <tableStyleInfo>.
  spec.style_name = nullptr;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(Info(wb.handle, index).style_name, nullptr);
  EXPECT_EQ(StyleOptions(wb.handle, index).present, 0);
  ASSERT_EQ(fm_sheet_remove_table(wb.handle, 0, index), 0);

  // A zero-initialised spec is styled with the workbook default.
  spec.styled = FM_TRISTATE_DEFAULT;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, &spec, &index), 0) << fm_get_last_error_message();
  EXPECT_STREQ(Info(wb.handle, index).style_name, "TableStyleMedium2");

  EXPECT_EQ(SetTableStyleName(wb.handle, index, ""), kInvalidArgument);

  fm_table_style_defaults_patch_t defaults = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  defaults.table_style_name_op = FM_PATCH_OP_SET;
  defaults.table_style_name = "";
  EXPECT_EQ(fm_workbook_update_table_style_defaults(wb.handle, &defaults), kInvalidArgument);
  defaults = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  defaults.pivot_style_name_op = FM_PATCH_OP_SET;
  defaults.pivot_style_name = "";
  EXPECT_EQ(fm_workbook_update_table_style_defaults(wb.handle, &defaults), kInvalidArgument);

  EXPECT_EQ(SetTableColumnTotals(wb.handle, 0, index, 0, FM_TABLE_TOTALS_ROW_FUNCTION_NONE, "", nullptr),
            kInvalidArgument);
  EXPECT_EQ(SetTableColumnTotals(wb.handle, 0, index, 1, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr, ""),
            kInvalidArgument);

  WorkbookGuard pivots;
  ASSERT_EQ(fm_workbook_create_default(&pivots.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(pivots.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  fm_pivot_style_options_patch_t pivot_patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  pivot_patch.style_name_op = FM_PATCH_OP_SET;
  pivot_patch.style_name = "";
  EXPECT_EQ(fm_sheet_update_pivot_style_options(pivots.handle, 0, pivot_index, &pivot_patch), kInvalidArgument);
}

// --- Column patches: calculated columns and totals -------------------------------------------------------------

/// Header `a b c` on row 1 (Excel row 2), data rows 2..4 with `a` = 1..3 and `b` = 10..30 and `c` empty; table `Tbl`
/// over A2:C5, or A2:C6 with a totals row. The first data row is Excel row 3, so `=A3*10` is the first-row form.
void CreateCalcTable(WorkbookGuard& wb, int32_t totals = 0) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* columns[] = {"a", "b", "c"};
  for (std::uint32_t col = 0; col < 3; ++col) {
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 1, col, columns[col]), 0);
  }
  for (std::uint32_t row = 2; row <= 4; ++row) {
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, row, 0, static_cast<double>(row - 1)), 0);
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, row, 1, static_cast<double>((row - 1) * 10)), 0);
  }
  const fm_cell_rect_t ref{1, 0, static_cast<std::uint32_t>(totals != 0 ? 5 : 4), 2};
  std::size_t index = 99U;
  ASSERT_EQ(AddTable(wb.handle, 0, &ref, "Tbl", columns, 3, "TableStyleMedium2", 1, totals, &index), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(index, 0U);
}

struct ColumnRead {
  std::int32_t function = 0;
  bool has_label = false;
  std::string label;
  bool has_formula = false;
  std::string formula;
  bool has_calc = false;
  std::string calc;
};

ColumnRead ReadColumn(const fm_workbook_t* wb, std::size_t column) {
  fm_table_column_info_t info = FM_STRUCT_INIT(fm_table_column_info_t);
  EXPECT_EQ(fm_sheet_get_table_column_at(wb, 0, 0, column, &info), 0);
  ColumnRead read;
  read.function = info.totals_row_function;
  read.has_label = info.totals_row_label != nullptr;
  read.label = read.has_label ? info.totals_row_label : "";
  read.has_formula = info.totals_row_formula != nullptr;
  read.formula = read.has_formula ? info.totals_row_formula : "";
  read.has_calc = info.calculated_formula != nullptr;
  read.calc = read.has_calc ? info.calculated_formula : "";
  return read;
}

fm_status_t PatchColumns(fm_workbook_t* wb, const std::vector<TableColumnElement>& columns,
                         fm_table_patch_t patch = TablePatch()) {
  return UpdateTableColumns(wb, 0, 0, columns, patch);
}

std::vector<std::uint8_t> SaveAs(fm_workbook_t* wb, fm_workbook_format_t format) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, format, &saved.data, &saved.len), 0);
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

double Number(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  const fm_value_t v = Value(wb, row, col);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_NUMBER) << row << "," << col;
  return v.u.number;
}

TEST(FormulonCApiTable, ColumnPatchSetsTheCalculatedFormulaAndFillsTheColumn) {
  WorkbookGuard wb;
  CreateCalcTable(wb);
  EXPECT_FALSE(ReadColumn(wb.handle, 2).has_calc);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 2, 99.0), 0);  // a constant in the column is overwritten
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0) << fm_get_last_error_message();
  const ColumnRead read = ReadColumn(wb.handle, 2);
  ASSERT_TRUE(read.has_calc);
  EXPECT_EQ(read.calc, "=A3*10");
  EXPECT_EQ(Formula(wb.handle, 2, 2), "=A3*10");
  EXPECT_EQ(Formula(wb.handle, 3, 2), "=A4*10");
  EXPECT_EQ(Formula(wb.handle, 4, 2), "=A5*10");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 2, 2), 10.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 3, 2), 20.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 2), 30.0);
  EXPECT_FALSE(ReadColumn(wb.handle, 0).has_calc);
  EXPECT_FALSE(ReadColumn(wb.handle, 1).has_calc);
  EXPECT_EQ(Value(wb.handle, 1, 2).kind, FM_VALUE_KIND_TEXT) << "the header cell is not part of the fill";
}

TEST(FormulonCApiTable, ColumnPatchCalculatedFormulaAcceptsEveryInputForm) {
  const struct {
    const char* input;
    const char* read;
  } kCases[] = {
      {"=A3*10", "=A3*10"},                      // with the equals sign
      {"A3*10", "=A3*10"},                       // without it
      {"=[@a]*10", "=[@a]*10"},                  // short this-row form
      {"=Tbl[[#This Row],[a]]*10", "=[@a]*10"},  // long this-row form reads back short
      {"=$A$3*10", "=$A$3*10"},                  // absolute references do not move
      {"=A$3*10+$B3", "=A$3*10+$B3"},            // mixed references move their relative part only
  };
  for (const auto& c : kCases) {
    WorkbookGuard wb;
    CreateCalcTable(wb);
    ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, c.input), 0)
        << c.input << ": " << fm_get_last_error_message();
    EXPECT_EQ(ReadColumn(wb.handle, 2).calc, c.read) << c.input;
  }
  WorkbookGuard wb;
  CreateCalcTable(wb);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=$A$3*10"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 2), 10.0) << "an absolute reference stays on A3 in every row";
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=[@a]*10"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 2, 2), 10.0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 2), 30.0);
}

TEST(FormulonCApiTable, ColumnPatchClearingTheCalculatedFormulaLeavesTheCells) {
  WorkbookGuard wb;
  CreateCalcTable(wb);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, nullptr), 0) << fm_get_last_error_message();
  EXPECT_FALSE(ReadColumn(wb.handle, 2).has_calc);
  EXPECT_EQ(Formula(wb.handle, 2, 2), "=A3*10");
  EXPECT_EQ(Formula(wb.handle, 3, 2), "=A4*10");
  EXPECT_EQ(Formula(wb.handle, 4, 2), "=A5*10");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_DOUBLE_EQ(Number(wb.handle, 4, 2), 30.0);
  EXPECT_EQ(ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml").find("calculatedColumnFormula"),
            std::string::npos);
}

TEST(FormulonCApiTable, ColumnPatchRefusesAnEmptyCalculatedFormulaAndKeepsTheMark) {
  WorkbookGuard wb;
  CreateCalcTable(wb);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0);
  EXPECT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, ""), kInvalidArgument) << "clearing is CLEAR, not \"\"";
  EXPECT_EQ(ReadColumn(wb.handle, 2).calc, "=A3*10");
  EXPECT_EQ(Formula(wb.handle, 3, 2), "=A4*10");
}

TEST(FormulonCApiTable, ColumnPatchRefusesUnusableCalculatedFormulasWithoutChanging) {
  WorkbookGuard wb;
  CreateCalcTable(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 2, 99.0), 0);
  EXPECT_NE(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=SUM(("), 0)
      << "a parse failure carries its own status";
  // A dynamic-array result, and a reference that the shift to the data rows pushes off the grid.
  EXPECT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=SEQUENCE(3)"), kInvalidArgument);
  EXPECT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=FILTER(A3:A5,A3:A5>1)"), kInvalidArgument);
  EXPECT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A1048576+1"), kInvalidArgument);
  EXPECT_FALSE(ReadColumn(wb.handle, 2).has_calc);
  const fm_value_t kept = Value(wb.handle, 3, 2);
  ASSERT_EQ(kept.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(kept.u.number, 99.0);
  EXPECT_EQ(Formula(wb.handle, 2, 2), "");
}

TEST(FormulonCApiTable, ColumnPatchRefusesAnArrayLongerThanTheColumnCountAndAppliesNothing) {
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  const TableColumnElement keep = TableColumnKeep();
  EXPECT_EQ(PatchColumns(wb.handle, {keep, TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_SUM),
                                     TableColumnCalculated("=A3*10"), keep}),
            FM_STATUS_INDEX_OUT_OF_RANGE)
      << "four elements for three columns";
  EXPECT_FALSE(ReadColumn(wb.handle, 2).has_calc);
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  EXPECT_EQ(Formula(wb.handle, 2, 2), "");
  EXPECT_EQ(Formula(wb.handle, 5, 1), "");
  // Exactly the column count is fine.
  EXPECT_EQ(PatchColumns(wb.handle, {keep, keep, keep}), 0);
}

TEST(FormulonCApiTable, ColumnPatchIsCheckedWholeBeforeTheGeometryChanges) {
  WorkbookGuard wb;
  CreateCalcTable(wb);
  fm_table_patch_t grow = TablePatch();
  grow.totals_row_op = FM_PATCH_OP_SET;
  grow.totals_row = 1;
  EXPECT_EQ(PatchColumns(wb.handle, std::vector<TableColumnElement>(7, TableColumnKeep()), grow),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(Info(wb.handle).totals_row_count, 0U) << "a refused column patch leaves the totals row off";
  EXPECT_EQ(Info(wb.handle).range.last_row, 4U);
  EXPECT_NE(PatchColumns(wb.handle, TableColumnsAt(2, TableColumnCalculated("=SUM((")), grow), 0);
  EXPECT_EQ(PatchColumns(wb.handle, TableColumnsAt(2, TableColumnCalculated("=SEQUENCE(3)")), grow), kInvalidArgument);
  EXPECT_EQ(PatchColumns(wb.handle, TableColumnsAt(2, TableColumnCalculated("=A1048576+1")), grow), kInvalidArgument);
  EXPECT_EQ(Info(wb.handle).totals_row_count, 0U);
  EXPECT_EQ(Info(wb.handle).range.last_row, 4U);
}

TEST(FormulonCApiTable, ColumnPatchWithoutColumnsChangesNothing) {
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0);
  ASSERT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr), 0);
  EXPECT_EQ(PatchColumns(wb.handle, {}), 0);
  fm_table_patch_t patch = TablePatch();
  patch.columns_op = FM_PATCH_OP_SET;
  patch.columns = nullptr;
  patch.columns_count = 0;
  EXPECT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &patch), 0);
  EXPECT_EQ(ReadColumn(wb.handle, 2).calc, "=A3*10");
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
  EXPECT_EQ(Formula(wb.handle, 5, 1), "=SUBTOTAL(109,[b])");
  // An all-KEEP element is a no-op, and `columns_op` KEEP ignores the array.
  const TableColumnElement keep = TableColumnKeep();
  EXPECT_EQ(PatchColumns(wb.handle, {keep, keep, keep}), 0);
  const TableColumnElement clear = TableColumnTotalsClear();
  fm_table_patch_t ignored = TablePatch();
  ignored.columns = &clear;
  ignored.columns_count = 1;
  EXPECT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &ignored), 0);
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
  EXPECT_EQ(ReadColumn(wb.handle, 2).calc, "=A3*10");
  // CLEAR is not an operation of the array itself.
  ignored.columns_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &ignored), kInvalidArgument);
}

TEST(FormulonCApiTable, ColumnPatchReapplyingTheSameCalculatedFormulaKeepsConstants) {
  WorkbookGuard wb;
  CreateCalcTable(wb);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 2, 99.0), 0);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0);
  EXPECT_EQ(Value(wb.handle, 3, 2).kind, FM_VALUE_KIND_NUMBER) << "the same formula is no change";
  EXPECT_DOUBLE_EQ(Value(wb.handle, 3, 2).u.number, 99.0);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A$3*10"), 0);
  EXPECT_EQ(Formula(wb.handle, 3, 2), "=A$3*10") << "a different formula fills every data cell";
  EXPECT_EQ(ReadColumn(wb.handle, 2).calc, "=A$3*10");
}

TEST(FormulonCApiTable, ColumnPatchClearsEachFieldAndANonCustomFunctionDropsTheFormula) {
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  TableColumnElement function_only{};
  function_only.totals_row_function_op = FM_PATCH_OP_SET;
  function_only.totals_row_function = FM_TABLE_TOTALS_ROW_FUNCTION_SUM;
  TableColumnElement label_only{};
  label_only.totals_row_label_op = FM_PATCH_OP_SET;
  label_only.totals_row_label = "Sum";
  ASSERT_EQ(PatchColumns(wb.handle, {TableColumnKeep(), function_only}), 0);
  ASSERT_EQ(PatchColumns(wb.handle, {TableColumnKeep(), label_only}), 0) << "the label key leaves the function";
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
  EXPECT_EQ(ReadColumn(wb.handle, 1).label, "Sum");

  TableColumnElement clear_function{};
  clear_function.totals_row_function_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(PatchColumns(wb.handle, {TableColumnKeep(), clear_function}), 0);
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  EXPECT_EQ(ReadColumn(wb.handle, 1).label, "Sum") << "clearing the function leaves the label";
  TableColumnElement clear_label{};
  clear_label.totals_row_label_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(PatchColumns(wb.handle, {TableColumnKeep(), clear_label}), 0);
  EXPECT_FALSE(ReadColumn(wb.handle, 1).has_label);

  ASSERT_EQ(PatchColumns(wb.handle, {TableColumnKeep(), TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr,
                                                                          "=SUM(Tbl[b])*2")}),
            0);
  EXPECT_EQ(ReadColumn(wb.handle, 1).formula, "=SUM([b])*2");
  TableColumnElement to_max{};
  to_max.totals_row_function_op = FM_PATCH_OP_SET;
  to_max.totals_row_function = FM_TABLE_TOTALS_ROW_FUNCTION_MAX;
  ASSERT_EQ(PatchColumns(wb.handle, {TableColumnKeep(), to_max}), 0);
  EXPECT_FALSE(ReadColumn(wb.handle, 1).has_formula) << "a non-custom function reads its formula as null";

  TableColumnElement custom_without_formula{};
  custom_without_formula.totals_row_function_op = FM_PATCH_OP_SET;
  custom_without_formula.totals_row_function = FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM;
  EXPECT_EQ(PatchColumns(wb.handle, {TableColumnKeep(), custom_without_formula}), kInvalidArgument);
  TableColumnElement empty_formula{};
  empty_formula.totals_row_formula_op = FM_PATCH_OP_SET;
  empty_formula.totals_row_formula = "";
  EXPECT_EQ(PatchColumns(wb.handle, {TableColumnKeep(), empty_formula}), kInvalidArgument);
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_MAX);
}

TEST(FormulonCApiTable, ColumnPatchAppliesSeveralColumnsInOneCall) {
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  ASSERT_EQ(
      PatchColumns(wb.handle, {TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_NONE, "Total"),
                               TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_SUM), TableColumnCalculated("=A3*10")}),
      0)
      << fm_get_last_error_message();
  EXPECT_EQ(ReadColumn(wb.handle, 0).label, "Total");
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
  EXPECT_EQ(ReadColumn(wb.handle, 2).calc, "=A3*10");
  EXPECT_EQ(Formula(wb.handle, 5, 1), "=SUBTOTAL(109,[b])");
  EXPECT_EQ(Formula(wb.handle, 4, 2), "=A5*10");
  EXPECT_NE(Formula(wb.handle, 5, 2), "=A6*10") << "the totals row is not part of the fill";
}

TEST(FormulonCApiTable, ColumnPatchTotalsClearBlanksTheShownCellAndHiddenTotalsAreMetadataOnly) {
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  ASSERT_EQ(PatchColumns(wb.handle, TableColumnsAt(1, TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_SUM))), 0);
  EXPECT_EQ(Formula(wb.handle, 5, 1), "=SUBTOTAL(109,[b])");
  ASSERT_EQ(PatchColumns(wb.handle, TableColumnsAt(1, TableColumnTotalsClear())), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  EXPECT_EQ(Value(wb.handle, 5, 1).kind, FM_VALUE_KIND_BLANK);

  WorkbookGuard hidden;
  CreateCalcTable(hidden);
  ASSERT_EQ(PatchColumns(hidden.handle, TableColumnsAt(1, TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_MAX))), 0);
  EXPECT_EQ(ReadColumn(hidden.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_MAX);
  EXPECT_EQ(Value(hidden.handle, 5, 1).kind, FM_VALUE_KIND_BLANK) << "no totals row, so no cell";
  EXPECT_EQ(Info(hidden.handle).range.last_row, 4U);
}

TEST(FormulonCApiTable, ColumnPatchOnATableWithoutDataRowsSetsTheMarkOnly) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* columns[] = {"a", "b", "c"};
  for (std::uint32_t col = 0; col < 3; ++col) {
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 1, col, columns[col]), 0);
  }
  const fm_cell_rect_t ref{1, 0, 2, 2};  // header row and totals row only
  std::size_t index = 99U;
  ASSERT_EQ(AddTable(wb.handle, 0, &ref, "Tbl", columns, 3, "TableStyleMedium2", 1, 1, &index), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0) << fm_get_last_error_message();
  const ColumnRead read = ReadColumn(wb.handle, 2);
  EXPECT_TRUE(read.has_calc);
  EXPECT_NE(read.calc.find("*10"), std::string::npos) << read.calc;
  EXPECT_EQ(Formula(wb.handle, 2, 2).find("*10"), std::string::npos) << "the totals row is not a data row";
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, nullptr), 0);
  EXPECT_FALSE(ReadColumn(wb.handle, 2).has_calc);
}

TEST(FormulonCApiTable, CalculatedColumnSurvivesSingleCellWrites) {
  const struct {
    const char* what;
    std::function<void(fm_workbook_t*)> write;
  } kCases[] = {
      {"a constant", [](fm_workbook_t* wb) { EXPECT_EQ(fm_sheet_set_number(wb, 0, 3, 2, 7.0), 0); }},
      {"a text", [](fm_workbook_t* wb) { EXPECT_EQ(fm_sheet_set_text(wb, 0, 3, 2, "x"), 0); }},
      {"a cleared cell", [](fm_workbook_t* wb) { EXPECT_EQ(fm_sheet_set_blank(wb, 0, 3, 2), 0); }},
      {"another formula", [](fm_workbook_t* wb) { EXPECT_EQ(fm_sheet_set_formula(wb, 0, 3, 2, "=B4+1"), 0); }},
  };
  for (const auto& c : kCases) {
    WorkbookGuard wb;
    CreateCalcTable(wb);
    ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0);
    c.write(wb.handle);
    EXPECT_EQ(ReadColumn(wb.handle, 2).calc, "=A3*10") << c.what;
    EXPECT_EQ(Formula(wb.handle, 2, 2), "=A3*10") << c.what << ": the neighbours are untouched";
    EXPECT_EQ(Formula(wb.handle, 4, 2), "=A5*10") << c.what << ": the neighbours are untouched";
  }
}

TEST(FormulonCApiTable, CalculatedColumnIsDroppedWhenItsLastNonBlankCellIsBlanked) {
  WorkbookGuard wb;
  CreateCalcTable(wb);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0);
  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 2, 2), 0);
  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 4, 2), 0);
  EXPECT_TRUE(ReadColumn(wb.handle, 2).has_calc) << "one non-blank cell is left";
  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 3, 2), 0);
  EXPECT_FALSE(ReadColumn(wb.handle, 2).has_calc);
  EXPECT_EQ(ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml").find("calculatedColumnFormula"),
            std::string::npos);
  // A column that never carried the mark is unaffected by blanking.
  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 2, 0), 0);
  EXPECT_FALSE(ReadColumn(wb.handle, 0).has_calc);
}

TEST(FormulonCApiTable, BlankingATotalsCellDoesNotDropTheCalculatedColumn) {
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  ASSERT_EQ(SetTableColumnCalculatedFormula(wb.handle, 0, 0, 2, "=A3*10"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 5, 2, "x"), 0);
  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 5, 2), 0);
  EXPECT_EQ(ReadColumn(wb.handle, 2).calc, "=A3*10");
}

TEST(FormulonCApiTable, TotalsCellWritesSetTheColumnMetadata) {
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  constexpr std::uint32_t kTotalsRow = 5;

  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, kTotalsRow, 1, 5.0), 0);
  ColumnRead read = ReadColumn(wb.handle, 1);
  EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  EXPECT_EQ(read.label, "5");
  EXPECT_EQ(Value(wb.handle, kTotalsRow, 1).kind, FM_VALUE_KIND_TEXT) << "a numeric label is stored as text";
  EXPECT_STREQ(Value(wb.handle, kTotalsRow, 1).u.text, "5");

  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, kTotalsRow, 1, 5.5), 0);
  EXPECT_EQ(ReadColumn(wb.handle, 1).label, "5.5");

  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, kTotalsRow, 0, "Total"), 0);
  read = ReadColumn(wb.handle, 0);
  EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  EXPECT_TRUE(read.has_label);
  EXPECT_EQ(read.label, "Total");

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kTotalsRow, 1, "=SUBTOTAL(109,Tbl[b])"), 0);
  read = ReadColumn(wb.handle, 1);
  EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
  EXPECT_FALSE(read.has_formula);

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kTotalsRow, 1, "=5"), 0);
  read = ReadColumn(wb.handle, 1);
  EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM) << "a constant formula is custom, not a label";
  EXPECT_EQ(read.formula, "=5");

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kTotalsRow, 1, "=SUBTOTAL(9,Tbl[b])"), 0);
  read = ReadColumn(wb.handle, 1);
  EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM) << "only the 100-series codes are built-in functions";
  EXPECT_EQ(read.formula, "=SUBTOTAL(9,[b])");

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kTotalsRow, 1, "=SUM(Tbl[b])*2"), 0);
  read = ReadColumn(wb.handle, 1);
  EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM);
  EXPECT_EQ(read.formula, "=SUM([b])*2");

  ASSERT_EQ(fm_sheet_set_bool(wb.handle, 0, kTotalsRow, 1, 1), 0);
  read = ReadColumn(wb.handle, 1);
  EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  EXPECT_EQ(read.label, "TRUE");

  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, kTotalsRow, 1), 0);
  read = ReadColumn(wb.handle, 1);
  EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  EXPECT_FALSE(read.has_label);
  EXPECT_FALSE(read.has_formula);

  // Neither a data cell nor a cell under a table without a totals row is a totals cell.
  ASSERT_EQ(PatchColumns(wb.handle, TableColumnsAt(1, TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_SUM))), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 4, 1, "x"), 0);
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
  fm_table_patch_t hide = TablePatch();
  hide.totals_row_op = FM_PATCH_OP_SET;
  hide.totals_row = 0;
  ASSERT_EQ(fm_sheet_update_table(wb.handle, 0, 0, &hide), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 5, 1, "below"), 0);
  EXPECT_EQ(ReadColumn(wb.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM)
      << "the row below is outside the table";
}

enum class TotalsWrite {
  kNumber,
  kText,
  kBool,
  kSubtotalOwn,
  kSubtotalOtherColumn,
  kConstantFormula,
  kCustomFormula,
  kBlank
};
enum class TotalsPrior { kNone, kLabel, kFunction, kCustom };

struct TotalsStep {
  TotalsWrite write;
  std::uint32_t column;
  TotalsPrior prior;
};

// Pairwise rows over write kind x column x prior state.
constexpr TotalsStep kTotalsSteps[] = {
    {TotalsWrite::kBlank, 2, TotalsPrior::kLabel},
    {TotalsWrite::kConstantFormula, 2, TotalsPrior::kFunction},
    {TotalsWrite::kSubtotalOwn, 1, TotalsPrior::kNone},
    {TotalsWrite::kCustomFormula, 0, TotalsPrior::kLabel},
    {TotalsWrite::kNumber, 1, TotalsPrior::kCustom},
    {TotalsWrite::kSubtotalOtherColumn, 2, TotalsPrior::kNone},
    {TotalsWrite::kConstantFormula, 0, TotalsPrior::kCustom},
    {TotalsWrite::kCustomFormula, 1, TotalsPrior::kFunction},
    {TotalsWrite::kSubtotalOwn, 0, TotalsPrior::kFunction},
    {TotalsWrite::kBool, 1, TotalsPrior::kLabel},
    {TotalsWrite::kConstantFormula, 1, TotalsPrior::kNone},
    {TotalsWrite::kNumber, 0, TotalsPrior::kNone},
    {TotalsWrite::kNumber, 2, TotalsPrior::kLabel},
    {TotalsWrite::kBlank, 0, TotalsPrior::kNone},
    {TotalsWrite::kBool, 0, TotalsPrior::kNone},
    {TotalsWrite::kSubtotalOtherColumn, 1, TotalsPrior::kLabel},
    {TotalsWrite::kConstantFormula, 2, TotalsPrior::kLabel},
    {TotalsWrite::kCustomFormula, 2, TotalsPrior::kNone},
    {TotalsWrite::kText, 2, TotalsPrior::kNone},
    {TotalsWrite::kSubtotalOwn, 2, TotalsPrior::kCustom},
    {TotalsWrite::kBool, 2, TotalsPrior::kFunction},
    {TotalsWrite::kBlank, 1, TotalsPrior::kFunction},
    {TotalsWrite::kSubtotalOtherColumn, 0, TotalsPrior::kCustom},
    {TotalsWrite::kText, 0, TotalsPrior::kCustom},
    {TotalsWrite::kCustomFormula, 0, TotalsPrior::kCustom},
    {TotalsWrite::kText, 1, TotalsPrior::kNone},
    {TotalsWrite::kNumber, 2, TotalsPrior::kFunction},
    {TotalsWrite::kText, 1, TotalsPrior::kFunction},
    {TotalsWrite::kBlank, 0, TotalsPrior::kCustom},
    {TotalsWrite::kBool, 2, TotalsPrior::kCustom},
    {TotalsWrite::kSubtotalOtherColumn, 0, TotalsPrior::kFunction},
    {TotalsWrite::kSubtotalOwn, 0, TotalsPrior::kLabel},
    {TotalsWrite::kText, 0, TotalsPrior::kLabel},
};

constexpr const char* kColumnNames[] = {"a", "b", "c"};
constexpr std::uint32_t kCalcTotalsRow = 5;

void WriteTotalsCell(fm_workbook_t* wb, std::uint32_t column, TotalsWrite write) {
  const std::string own = std::string("Tbl[") + kColumnNames[column] + "]";
  const std::string other = std::string("Tbl[") + kColumnNames[(column + 1) % 3] + "]";
  switch (write) {
    case TotalsWrite::kNumber:
      ASSERT_EQ(fm_sheet_set_number(wb, 0, kCalcTotalsRow, column, 5.5), 0);
      break;
    case TotalsWrite::kText:
      ASSERT_EQ(fm_sheet_set_text(wb, 0, kCalcTotalsRow, column, "Total"), 0);
      break;
    case TotalsWrite::kBool:
      ASSERT_EQ(fm_sheet_set_bool(wb, 0, kCalcTotalsRow, column, 1), 0);
      break;
    case TotalsWrite::kSubtotalOwn:
      ASSERT_EQ(fm_sheet_set_formula(wb, 0, kCalcTotalsRow, column, ("=SUBTOTAL(104," + own + ")").c_str()), 0);
      break;
    case TotalsWrite::kSubtotalOtherColumn:
      ASSERT_EQ(fm_sheet_set_formula(wb, 0, kCalcTotalsRow, column, ("=SUBTOTAL(109," + other + ")").c_str()), 0);
      break;
    case TotalsWrite::kConstantFormula:
      ASSERT_EQ(fm_sheet_set_formula(wb, 0, kCalcTotalsRow, column, "=5"), 0);
      break;
    case TotalsWrite::kCustomFormula:
      ASSERT_EQ(fm_sheet_set_formula(wb, 0, kCalcTotalsRow, column, ("=SUM(" + own + ")*2").c_str()), 0);
      break;
    case TotalsWrite::kBlank:
      ASSERT_EQ(fm_sheet_set_blank(wb, 0, kCalcTotalsRow, column), 0);
      break;
  }
}

void ApplyPrior(fm_workbook_t* wb, std::uint32_t column, TotalsPrior prior) {
  switch (prior) {
    case TotalsPrior::kNone:
      WriteTotalsCell(wb, column, TotalsWrite::kBlank);
      break;
    case TotalsPrior::kLabel:
      WriteTotalsCell(wb, column, TotalsWrite::kText);
      break;
    case TotalsPrior::kFunction:
      WriteTotalsCell(wb, column, TotalsWrite::kSubtotalOwn);
      break;
    case TotalsPrior::kCustom:
      WriteTotalsCell(wb, column, TotalsWrite::kConstantFormula);
      break;
  }
}

/// The column's totals metadata, judged against what its totals cell holds now.
void ExpectTotalsAgreeWithTheCell(fm_workbook_t* wb, std::uint32_t column, const std::string& where) {
  static const struct {
    std::int32_t function;
    int code;
  } kBuiltIns[] = {{FM_TABLE_TOTALS_ROW_FUNCTION_SUM, 109},     {FM_TABLE_TOTALS_ROW_FUNCTION_MIN, 105},
                   {FM_TABLE_TOTALS_ROW_FUNCTION_MAX, 104},     {FM_TABLE_TOTALS_ROW_FUNCTION_AVERAGE, 101},
                   {FM_TABLE_TOTALS_ROW_FUNCTION_COUNT, 103},   {FM_TABLE_TOTALS_ROW_FUNCTION_COUNT_NUMBERS, 102},
                   {FM_TABLE_TOTALS_ROW_FUNCTION_STD_DEV, 107}, {FM_TABLE_TOTALS_ROW_FUNCTION_VAR, 110}};
  const ColumnRead read = ReadColumn(wb, column);
  const std::string formula = Formula(wb, kCalcTotalsRow, column);
  const fm_value_t cell = Value(wb, kCalcTotalsRow, column);
  if (!formula.empty()) {
    std::int32_t expected = FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM;
    for (const auto& b : kBuiltIns) {
      if (formula == "=SUBTOTAL(" + std::to_string(b.code) + ",[" + kColumnNames[column] + "])") {
        expected = b.function;
      }
    }
    EXPECT_EQ(read.function, expected) << where << ": " << formula;
    EXPECT_EQ(read.has_formula, expected == FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM) << where;
    if (expected == FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM) {
      EXPECT_EQ(read.formula, formula) << where;
    }
  } else if (cell.kind == FM_VALUE_KIND_BLANK) {
    EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE) << where;
    EXPECT_FALSE(read.has_label) << where;
    EXPECT_FALSE(read.has_formula) << where;
  } else {
    EXPECT_EQ(read.function, FM_TABLE_TOTALS_ROW_FUNCTION_NONE) << where;
    EXPECT_EQ(cell.kind, FM_VALUE_KIND_TEXT) << where << ": a value in a totals cell is stored as its label text";
    EXPECT_TRUE(read.has_label) << where;
    if (cell.kind == FM_VALUE_KIND_TEXT) {
      EXPECT_EQ(read.label, cell.u.text) << where;
    }
  }
}

const char* ExpectedLabel(TotalsWrite write) {
  switch (write) {
    case TotalsWrite::kNumber:
      return "5.5";
    case TotalsWrite::kText:
      return "Total";
    case TotalsWrite::kBool:
      return "TRUE";
    default:
      return nullptr;
  }
}

TEST(FormulonCApiTable, TotalsMetadataAgreesWithTheCellAfterEveryWrite) {
  // Each step on a fresh workbook: bring the column to the prior state, then write.
  int index = 0;
  for (const TotalsStep& step : kTotalsSteps) {
    WorkbookGuard wb;
    CreateCalcTable(wb, 1);
    const std::string where = "step " + std::to_string(index++);
    ApplyPrior(wb.handle, step.column, step.prior);
    ExpectTotalsAgreeWithTheCell(wb.handle, step.column, where + " prior");
    WriteTotalsCell(wb.handle, step.column, step.write);
    ExpectTotalsAgreeWithTheCell(wb.handle, step.column, where + " write");
    if (const char* label = ExpectedLabel(step.write)) {
      EXPECT_EQ(ReadColumn(wb.handle, step.column).label, label) << where;
    }
  }

  // The same writes as one running sequence: every column stays consistent after every step.
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  index = 0;
  for (const TotalsStep& step : kTotalsSteps) {
    WriteTotalsCell(wb.handle, step.column, step.write);
    for (std::uint32_t column = 0; column < 3; ++column) {
      ExpectTotalsAgreeWithTheCell(wb.handle, column,
                                   "sequence " + std::to_string(index) + " column " + std::to_string(column));
    }
    ++index;
  }
}

TEST(FormulonCApiTable, CalculatedColumnAndTotalsRoundTripThroughXlsx) {
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  ASSERT_EQ(
      PatchColumns(wb.handle, {TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_NONE, "Total"),
                               TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_SUM), TableColumnCalculated("=A3*10")}),
      0);
  const std::vector<std::uint8_t> saved = Save(wb.handle);
  const std::string table_xml = ExtractZipEntry(saved, "xl/tables/table1.xml");
  EXPECT_NE(table_xml.find("A3*10</calculatedColumnFormula>"), std::string::npos) << table_xml;
  EXPECT_NE(ExtractZipEntry(saved, "xl/worksheets/sheet1.xml").find("t=\"shared\""), std::string::npos)
      << "an equal-in-R1C1 column is written as a shared formula";

  WorkbookGuard reloaded;
  Load(saved, reloaded);
  EXPECT_EQ(ReadColumn(reloaded.handle, 2).calc, "=A3*10");
  EXPECT_EQ(ReadColumn(reloaded.handle, 0).label, "Total");
  EXPECT_EQ(ReadColumn(reloaded.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
  for (std::uint32_t row = 2; row <= 4; ++row) {
    EXPECT_EQ(Formula(reloaded.handle, row, 2), "=A" + std::to_string(row + 1) + "*10");
  }
  EXPECT_EQ(Formula(reloaded.handle, 5, 1), "=SUBTOTAL(109,[b])");
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  EXPECT_DOUBLE_EQ(Number(reloaded.handle, 4, 2), 30.0);
  EXPECT_DOUBLE_EQ(Number(reloaded.handle, 5, 1), 60.0);
  // The reloaded column still behaves as a calculated column: a one-cell write keeps it.
  ASSERT_EQ(fm_sheet_set_number(reloaded.handle, 0, 3, 2, 7.0), 0);
  EXPECT_EQ(ReadColumn(reloaded.handle, 2).calc, "=A3*10");
}

TEST(FormulonCApiTable, CalculatedColumnAndTotalsRoundTripThroughXlsb) {
  WorkbookGuard wb;
  CreateCalcTable(wb, 1);
  ASSERT_EQ(
      PatchColumns(wb.handle, {TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_NONE, "Total"),
                               TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_SUM), TableColumnCalculated("=A3*10")}),
      0);
  const std::vector<std::uint8_t> saved = SaveAs(wb.handle, FM_WORKBOOK_FORMAT_XLSB);
  WorkbookGuard reloaded;
  Load(saved, reloaded);
  EXPECT_EQ(ReadColumn(reloaded.handle, 2).calc, "=A3*10");
  EXPECT_EQ(ReadColumn(reloaded.handle, 0).label, "Total");
  EXPECT_EQ(ReadColumn(reloaded.handle, 1).function, FM_TABLE_TOTALS_ROW_FUNCTION_SUM);
  EXPECT_EQ(Formula(reloaded.handle, 2, 2), "=A3*10");
  EXPECT_EQ(Formula(reloaded.handle, 4, 2), "=A5*10");
  EXPECT_EQ(Formula(reloaded.handle, 5, 1), "=SUBTOTAL(109,[b])");
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  EXPECT_DOUBLE_EQ(Number(reloaded.handle, 5, 1), 60.0);
}

}  // namespace
