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

std::vector<std::uint8_t> Save(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0);
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
  ASSERT_EQ(fm_workbook_add_table(wb.handle, 0, &ref, "Sales", "Sales", columns, 2, style, 1, totals, &index), 0);
  ASSERT_EQ(index, 0U);
}

fm_table_info_t Info(fm_workbook_t* wb, std::size_t index = 0U) {
  fm_table_info_t info = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_workbook_get_table_info(wb, index, &info), 0);
  return info;
}

std::string Formula(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  const char* text = nullptr;
  EXPECT_EQ(fm_sheet_get_formula(wb, 0, row, col, &text), 0);
  return text != nullptr ? text : "";
}

fm_table_style_t MakeTableStyle(const char* name, int32_t pivot, int32_t table,
                                const fm_table_style_element_t* elements, std::uint32_t element_count) {
  fm_table_style_t style = FM_STRUCT_INIT(fm_table_style_t);
  style.name = name;
  style.pivot = pivot;
  style.table = table;
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
  EXPECT_STREQ(info.style_name, "TableStyleMedium2");
  EXPECT_EQ(info.style_present, 1);
  EXPECT_EQ(info.show_first_column, 0);
  EXPECT_EQ(info.show_last_column, 0);
  EXPECT_EQ(info.show_row_stripes, 1);
  EXPECT_EQ(info.show_column_stripes, 0);
  EXPECT_EQ(info.header_row_count, 1U);
  EXPECT_EQ(info.totals_row_count, 0U);
  EXPECT_EQ(info.column_count, 2U);

  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  ASSERT_EQ(fm_workbook_get_table_column(wb.handle, 0, 1, &column), 0);
  EXPECT_STREQ(column.name, "Amount");
  EXPECT_STREQ(column.totals_row_function, "none");
  EXPECT_STREQ(column.totals_row_label, "");
  EXPECT_STREQ(column.totals_row_formula, "");
  EXPECT_EQ(fm_workbook_get_table_column(wb.handle, 0, 2, &column), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_get_table_info(wb.handle, 1, nullptr), kBindingNullPointer);
  fm_table_info_t ignored = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_workbook_get_table_info(wb.handle, 1, &ignored), FM_STATUS_INDEX_OUT_OF_RANGE);

  WorkbookGuard plain;
  CreateSalesTable(plain, "");
  const fm_table_info_t unstyled = Info(plain.handle);
  EXPECT_EQ(unstyled.style_present, 0);
  EXPECT_EQ(unstyled.show_row_stripes, 0);
  EXPECT_STREQ(unstyled.style_name, "");
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

  ASSERT_EQ(fm_workbook_update_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 2, 1}), "TableStyleMedium4", -1, -1),
            0);
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

TEST(FormulonCApiTable, SetStyleInfoWritesTheElementInSchemaPosition) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "");
  ASSERT_EQ(fm_workbook_set_table_style_info(wb.handle, 0, "TableStyleDark1", 1, 0, 0, 1), 0);
  fm_table_info_t info = Info(wb.handle);
  EXPECT_STREQ(info.style_name, "TableStyleDark1");
  EXPECT_EQ(info.show_first_column, 1);
  EXPECT_EQ(info.show_column_stripes, 1);

  // No style name keeps the element and its toggles, as the "None" style does.
  ASSERT_EQ(fm_workbook_set_table_style_info(wb.handle, 0, nullptr, 1, 0, 0, 1), 0);
  info = Info(wb.handle);
  EXPECT_EQ(info.style_present, 1);
  EXPECT_STREQ(info.style_name, "");
  const std::string xml = ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml");
  const std::size_t columns_end = xml.find("</tableColumns>");
  const std::size_t style = xml.find(
      "<tableStyleInfo showFirstColumn=\"1\" showLastColumn=\"0\" "
      "showRowStripes=\"0\" showColumnStripes=\"1\"/>");
  ASSERT_NE(columns_end, std::string::npos);
  ASSERT_NE(style, std::string::npos);
  EXPECT_LT(columns_end, style);

  // An empty style through update still removes the element.
  ASSERT_EQ(fm_workbook_update_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 2, 1}), "", -1, -1), 0);
  EXPECT_EQ(Info(wb.handle).style_present, 0);
  EXPECT_EQ(fm_workbook_set_table_style_info(wb.handle, 3, "X", 0, 0, 0, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiTable, ColumnTotalsWriteTheTotalsRowCells) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleMedium2", 1, {0, 0, 3, 1});
  ASSERT_EQ(fm_workbook_set_table_column_totals(wb.handle, 0, 0, nullptr, "Total", nullptr), 0);
  ASSERT_EQ(fm_workbook_set_table_column_totals(wb.handle, 0, 1, "sum", nullptr, nullptr), 0);
  EXPECT_EQ(Formula(wb.handle, 3, 1), "=SUBTOTAL(109,[Amount])");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t total = Value(wb.handle, 3, 1);
  ASSERT_EQ(total.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(total.u.number, 30.0);
  fm_value_t label = Value(wb.handle, 3, 0);
  ASSERT_EQ(label.kind, FM_VAL_TEXT);
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
  ASSERT_EQ(fm_workbook_get_table_column(reloaded.handle, 0, 0, &column), 0);
  EXPECT_STREQ(column.totals_row_label, "Total");
  EXPECT_STREQ(column.totals_row_function, "none");
  ASSERT_EQ(fm_workbook_get_table_column(reloaded.handle, 0, 1, &column), 0);
  EXPECT_STREQ(column.totals_row_function, "sum");
  EXPECT_EQ(Info(reloaded.handle).totals_row_count, 1U);
  EXPECT_EQ(Formula(reloaded.handle, 3, 1), "=SUBTOTAL(109,[Amount])");

  // Each built-in token maps to its SUBTOTAL number; `none` clears the cell.
  const struct {
    const char* token;
    const char* formula;
  } kCases[] = {{"average", "=SUBTOTAL(101,[Amount])"}, {"countNums", "=SUBTOTAL(102,[Amount])"},
                {"count", "=SUBTOTAL(103,[Amount])"},   {"max", "=SUBTOTAL(104,[Amount])"},
                {"min", "=SUBTOTAL(105,[Amount])"},     {"stdDev", "=SUBTOTAL(107,[Amount])"},
                {"var", "=SUBTOTAL(110,[Amount])"}};
  for (const auto& c : kCases) {
    ASSERT_EQ(fm_workbook_set_table_column_totals(reloaded.handle, 0, 1, c.token, nullptr, nullptr), 0) << c.token;
    EXPECT_EQ(Formula(reloaded.handle, 3, 1), c.formula);
  }
  ASSERT_EQ(fm_workbook_set_table_column_totals(reloaded.handle, 0, 1, "none", nullptr, nullptr), 0);
  EXPECT_EQ(Formula(reloaded.handle, 3, 1), "");
  EXPECT_EQ(Value(reloaded.handle, 3, 1).kind, FM_VAL_BLANK);
}

TEST(FormulonCApiTable, CustomTotalsFormulaRoundTrips) {
  WorkbookGuard wb;
  CreateSalesTable(wb, "TableStyleMedium2", 1, {0, 0, 3, 1});
  EXPECT_EQ(fm_workbook_set_table_column_totals(wb.handle, 0, 1, "custom", nullptr, nullptr), kInvalidArgument);
  EXPECT_EQ(fm_workbook_set_table_column_totals(wb.handle, 0, 1, "median", nullptr, nullptr), kInvalidArgument);
  EXPECT_NE(fm_workbook_set_table_column_totals(wb.handle, 0, 1, "custom", nullptr, "=SUM(("), 0);
  ASSERT_EQ(fm_workbook_set_table_column_totals(wb.handle, 0, 1, "custom", nullptr, "=SUM(Sales[Amount])*2"), 0);
  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  ASSERT_EQ(fm_workbook_get_table_column(wb.handle, 0, 1, &column), 0);
  EXPECT_STREQ(column.totals_row_function, "custom");
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
  ASSERT_EQ(fm_workbook_get_table_column(reloaded.handle, 0, 1, &column), 0);
  EXPECT_STREQ(column.totals_row_function, "custom");
  EXPECT_STREQ(column.totals_row_formula, "=SUM([Amount])*2");
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  const fm_value_t v = Value(reloaded.handle, 3, 1);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 60.0);
}

TEST(FormulonCApiTable, TurningTheTotalsRowOnRestoresItsCells) {
  // A file whose totals row was shown, then hidden, keeps the column
  // settings; showing it again writes the cells back and drops the
  // "never shown" marker.
  WorkbookGuard source;
  CreateSalesTable(source);
  ASSERT_EQ(fm_workbook_set_table_column_totals(source.handle, 0, 0, nullptr, "Total", nullptr), 0);
  ASSERT_EQ(fm_workbook_set_table_column_totals(source.handle, 0, 1, "max", nullptr, nullptr), 0);
  EXPECT_EQ(Value(source.handle, 3, 1).kind, FM_VAL_BLANK) << "no cell while the totals row is off";
  const std::vector<std::uint8_t> patched =
      RewriteEntry(Save(source.handle), "xl/tables/table1.xml", [](std::string xml) {
        return ReplaceOnce(std::move(xml), " ref=\"A1:B3\">", " ref=\"A1:B3\" totalsRowShown=\"0\">");
      });
  WorkbookGuard wb;
  Load(patched, wb);
  ASSERT_EQ(fm_workbook_update_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 3, 1}), nullptr, -1, 1), 0);
  EXPECT_EQ(Formula(wb.handle, 3, 1), "=SUBTOTAL(104,[Amount])");
  const fm_value_t label = Value(wb.handle, 3, 0);
  ASSERT_EQ(label.kind, FM_VAL_TEXT);
  EXPECT_STREQ(label.u.text, "Total");
  const std::string xml = ExtractZipEntry(Save(wb.handle), "xl/tables/table1.xml");
  EXPECT_EQ(xml.find("totalsRowShown"), std::string::npos);
  EXPECT_NE(xml.find("totalsRowCount=\"1\""), std::string::npos);
}

TEST(FormulonCApiTable, LoadedTotalsShownMarkerRoundTrips) {
  WorkbookGuard source;
  CreateSalesTable(source);
  const std::vector<std::uint8_t> patched =
      RewriteEntry(Save(source.handle), "xl/tables/table1.xml", [](std::string xml) {
        return ReplaceOnce(std::move(xml), " ref=\"A1:B3\">", " ref=\"A1:B3\" totalsRowShown=\"0\">");
      });
  WorkbookGuard wb;
  Load(patched, wb);
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
  ASSERT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 2, 1}), "T", "T", columns, 2, "", 1, 1,
                                  &index),
            0);
  ASSERT_EQ(fm_workbook_set_table_column_totals(wb.handle, 0, 0, "sum", nullptr, nullptr), 0);
  ASSERT_EQ(fm_workbook_set_table_column_totals(wb.handle, 0, 1, "sum", nullptr, nullptr), 0);
  EXPECT_EQ(Formula(wb.handle, 2, 0), "=SUBTOTAL(109,[Q1 '#est])");
  EXPECT_EQ(Formula(wb.handle, 2, 1), "=SUBTOTAL(109,[a,b])");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v = Value(wb.handle, 2, 0);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 4.0);
  v = Value(wb.handle, 2, 1);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 5.0);
}

TEST(FormulonCApiTable, RenameRewritesStructuredReferences) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  const char* columns[] = {"Key"};
  std::size_t other = 0;
  ASSERT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 5, 1, 5}), "Other", "Other", columns, 1,
                                  "", 1, 0, &other),
            0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 3, "=SUM(Sales[Amount])+COUNTA(sales[Product])"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 3, "=Sales"), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Doubled", "=SUM(Sales[Amount])*2"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 3, "=Doubled"), 0);

  EXPECT_EQ(fm_workbook_rename_table(wb.handle, 0, "other", nullptr), kInvalidArgument);
  EXPECT_EQ(fm_workbook_rename_table(wb.handle, 0, "", nullptr), kInvalidArgument);
  EXPECT_EQ(fm_workbook_rename_table(wb.handle, 0, nullptr, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_rename_table(wb.handle, 5, "Revenue", nullptr), FM_STATUS_INDEX_OUT_OF_RANGE);
  // A case-only rename of the table itself is not a collision.
  ASSERT_EQ(fm_workbook_rename_table(wb.handle, 0, "SALES", nullptr), 0);
  ASSERT_EQ(fm_workbook_rename_table(wb.handle, 0, "Revenue", nullptr), 0);

  const char* name = nullptr;
  const char* display = nullptr;
  fm_cell_rect_t rect{};
  std::size_t sheet = 0;
  ASSERT_EQ(fm_workbook_get_table(wb.handle, 0, &name, &display, &rect, &sheet), 0);
  EXPECT_STREQ(name, "Revenue");
  EXPECT_STREQ(display, "Revenue");
  EXPECT_EQ(Formula(wb.handle, 0, 3), "=SUM(Revenue[Amount])+COUNTA(Revenue[Product])");
  EXPECT_EQ(Formula(wb.handle, 1, 3), "=Revenue") << "an entered bare table name binds to the whole table";
  fm_defined_name_t dn = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.formula, "=SUM(Revenue[Amount])*2");

  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v = Value(wb.handle, 0, 3);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 32.0);
  v = Value(wb.handle, 2, 3);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 60.0);

  ASSERT_EQ(fm_workbook_rename_table(wb.handle, 0, "RevenueInternal", "RevenueShown"), 0);
  EXPECT_EQ(Formula(wb.handle, 0, 3), "=SUM(RevenueShown[Amount])+COUNTA(RevenueShown[Product])");

  WorkbookGuard reloaded;
  Load(Save(wb.handle), reloaded);
  ASSERT_EQ(fm_workbook_get_table(reloaded.handle, 0, &name, &display, &rect, &sheet), 0);
  EXPECT_STREQ(name, "RevenueInternal");
  EXPECT_STREQ(display, "RevenueShown");
  ASSERT_EQ(fm_workbook_recalc(reloaded.handle), 0);
  v = Value(reloaded.handle, 0, 3);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 32.0);
}

fm_dxf_record_t FillDxf() {
  fm_dxf_record_t dxf{};
  dxf.fill_engaged = 1;
  dxf.fill.pattern = 1;
  dxf.fill.fg.kind = kFmColorRgb;
  dxf.fill.fg.rgb = 0xFFDDEBF7U;
  return dxf;
}

TEST(FormulonCApiTableStyles, CustomStyleRoundTripsAndATableUsesIt) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  std::size_t dxf = 99;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, record_arg(FillDxf()), &dxf), 0);

  const fm_table_style_element_t elements[] = {
      {"wholeTable", 1U, 1, static_cast<uint32_t>(dxf)},
      {"headerRow", 1U, 1, static_cast<uint32_t>(dxf)},
      {"firstRowStripe", 2U, 1, static_cast<uint32_t>(dxf)},
      {"totalRow", 1U, 0, 0U},
  };
  fm_table_style_t style = MakeTableStyle("Corporate", 0, 1, elements, 4U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &style), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_table_style_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);
  const char* default_table = nullptr;
  const char* default_pivot = nullptr;
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &default_table, &default_pivot), 0);
  EXPECT_STREQ(default_table, "TableStyleMedium2");
  EXPECT_STREQ(default_pivot, "PivotStyleLight16");
  ASSERT_EQ(fm_workbook_set_table_style_info(wb.handle, 0, "Corporate", 0, 0, 1, 0), 0);

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
  EXPECT_EQ(got.pivot, 0);
  EXPECT_EQ(got.table, 1);
  ASSERT_EQ(got.element_count, 4U);
  EXPECT_STREQ(got.elements[2].type, "firstRowStripe");
  EXPECT_EQ(got.elements[2].size, 2U);
  EXPECT_EQ(got.elements[2].has_dxf_id, 1);
  EXPECT_EQ(got.elements[2].dxf_id, dxf);
  EXPECT_EQ(got.elements[3].has_dxf_id, 0);
  fm_dxf_record_t reloaded_dxf{};
  ASSERT_EQ(fm_workbook_get_dxf(reloaded.handle, got.elements[0].dxf_id, &reloaded_dxf), 0);
  EXPECT_EQ(reloaded_dxf.fill_engaged, 1);

  // Setting the record a get returned, views and all, leaves the part as it was.
  const std::string before = ExtractZipEntry(Save(reloaded.handle), "xl/styles.xml");
  ASSERT_EQ(fm_workbook_get_table_style_at(reloaded.handle, 0, &got), 0);
  ASSERT_EQ(fm_workbook_set_table_style(reloaded.handle, &got), 0);
  EXPECT_EQ(ExtractZipEntry(Save(reloaded.handle), "xl/styles.xml"), before);
}

TEST(FormulonCApiTableStyles, SetReplacesByNameAndValidates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t dxf = 0;
  ASSERT_EQ(fm_workbook_add_dxf(wb.handle, record_arg(FillDxf()), &dxf), 0);
  const fm_table_style_element_t one[] = {{"wholeTable", 1U, 1, static_cast<uint32_t>(dxf)}};
  fm_table_style_t style = MakeTableStyle("Mine", 1, 1, one, 1U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &style), 0);
  const fm_table_style_element_t two[] = {{"pageFieldLabels", 1U, 0, 0U},
                                          {"blankRow", 1U, 1, static_cast<uint32_t>(dxf)}};
  fm_table_style_t replaced = MakeTableStyle("MINE", 1, 0, two, 2U);
  ASSERT_EQ(fm_workbook_set_table_style(wb.handle, &replaced), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_table_style_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);
  fm_table_style_t got = FM_STRUCT_INIT(fm_table_style_t);
  ASSERT_EQ(fm_workbook_get_table_style_at(wb.handle, 0, &got), 0);
  EXPECT_STREQ(got.name, "MINE");
  EXPECT_EQ(got.table, 0);
  ASSERT_EQ(got.element_count, 2U);
  EXPECT_STREQ(got.elements[0].type, "pageFieldLabels");

  const fm_table_style_element_t unknown[] = {{"bodyRow", 1U, 0, 0U}};
  fm_table_style_t bad = MakeTableStyle("Other", 1, 1, unknown, 1U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  const fm_table_style_element_t repeated[] = {{"wholeTable", 1U, 0, 0U}, {"wholeTable", 1U, 0, 0U}};
  bad = MakeTableStyle("Other", 1, 1, repeated, 2U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  const fm_table_style_element_t zero_size[] = {{"firstRowStripe", 0U, 0, 0U}};
  bad = MakeTableStyle("Other", 1, 1, zero_size, 1U);
  EXPECT_EQ(fm_workbook_set_table_style(wb.handle, &bad), kInvalidArgument);
  const fm_table_style_element_t dangling[] = {{"wholeTable", 1U, 1, static_cast<uint32_t>(dxf + 1U)}};
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
  ASSERT_EQ(fm_workbook_set_table_style_defaults(wb.handle, "Mine", nullptr), 0);
  const char* table = nullptr;
  const char* pivot = nullptr;
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &table, &pivot), 0);
  EXPECT_STREQ(table, "Mine");
  EXPECT_STREQ(pivot, "PivotStyleLight16");

  EXPECT_EQ(fm_workbook_remove_table_style(wb.handle, "Missing"), FM_STATUS_NOT_FOUND);
  ASSERT_EQ(fm_workbook_remove_table_style(wb.handle, "mine"), 0);
  ASSERT_EQ(fm_workbook_get_table_style_defaults(wb.handle, &table, &pivot), 0);
  EXPECT_STREQ(table, "TableStyleMedium2");

  ASSERT_EQ(fm_workbook_set_table_style_defaults(wb.handle, nullptr, ""), 0);
  const std::string styles_xml = ExtractZipEntry(Save(wb.handle), "xl/styles.xml");
  EXPECT_NE(styles_xml.find("<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\"/>"), std::string::npos);
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
  ASSERT_EQ(count_of(fm_workbook_get_table_count, wb.handle), 0U);
  const char* columns[] = {"A", "B"};
  std::size_t index = 0;
  EXPECT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "sales", "sales", columns, 2,
                                  nullptr, 1, 0, &index),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "Fresh", "SALES", columns, 2,
                                  nullptr, 1, 0, &index),
            kInvalidArgument);
  EXPECT_EQ(count_of(fm_workbook_get_table_count, wb.handle), 0U);
  EXPECT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "Fresh", "Fresh", columns, 2,
                                  nullptr, 1, 0, &index),
            0);
}

TEST(FormulonCApiTable, RenameRejectsTheNameOfAVerbatimXlsbTable) {
  WorkbookGuard wb;
  Load(RetainedSalesPackage(3), wb);
  const char* columns[] = {"A", "B"};
  std::size_t index = 0;
  ASSERT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "Fresh", "Fresh", columns, 2,
                                  nullptr, 1, 0, &index),
            0);
  EXPECT_EQ(fm_workbook_rename_table(wb.handle, index, "Sales", "Sales"), kInvalidArgument);
  EXPECT_EQ(fm_workbook_rename_table(wb.handle, index, "Fresh2", "sALES"), kInvalidArgument);
  EXPECT_EQ(fm_workbook_rename_table(wb.handle, index, "Fresh2", "Fresh2"), 0);
}

TEST(FormulonCApiTable, AddUpdateAndRenameRefuseWhatExcelRefuses) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  const char* columns[] = {"A", "B"};
  std::size_t index = 0;
  EXPECT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "T1", "T1", columns, 2, nullptr,
                                  1, 0, &index),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{2, 1, 4, 2}), "Other", "Other", columns, 2,
                                  nullptr, 1, 0, &index),
            kInvalidArgument);
  ASSERT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "Other", "Other", columns, 2,
                                  nullptr, 1, 0, &index),
            0);
  EXPECT_EQ(fm_workbook_update_table(wb.handle, index, record_arg(fm_cell_rect_t{0, 1, 2, 2}), nullptr, -1, -1),
            kInvalidArgument);
  // A resize keeps the first row and, under a header row, one data row.
  EXPECT_EQ(fm_workbook_update_table(wb.handle, index, record_arg(fm_cell_rect_t{1, 4, 3, 5}), nullptr, -1, -1),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_update_table(wb.handle, index, record_arg(fm_cell_rect_t{0, 4, 0, 5}), nullptr, -1, -1),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_update_table(wb.handle, index, record_arg(fm_cell_rect_t{0, 4, 3, 5}), nullptr, -1, -1), 0);
  EXPECT_EQ(fm_workbook_rename_table(wb.handle, index, "R2D2", nullptr), kInvalidArgument);
  EXPECT_EQ(count_of(fm_workbook_get_table_count, wb.handle), 2U);
}

TEST(FormulonCApiTable, CreateAvoidsTheIdOfAVerbatimXlsbTable) {
  WorkbookGuard wb;
  Load(RetainedSalesPackage(1), wb);
  const char* columns[] = {"A", "B"};
  std::size_t index = 0;
  ASSERT_EQ(fm_workbook_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 4, 2, 5}), "Fresh", "Fresh", columns, 2,
                                  nullptr, 1, 0, &index),
            0);
  // The xlsx part is named after the table id.
  const std::string table_xml = ExtractZipEntry(Save(wb.handle), "xl/tables/table2.xml");
  EXPECT_NE(table_xml.find(" id=\"2\""), std::string::npos) << table_xml;
  EXPECT_NE(table_xml.find("name=\"Fresh\""), std::string::npos) << table_xml;
}

std::uint32_t TableRuntimeId(const fm_workbook_t* wb, std::size_t index) {
  std::uint32_t id = 0;
  EXPECT_EQ(fm_workbook_get_table_runtime_id(wb, index, &id), 0) << fm_get_last_error_message();
  return id;
}

std::string TableName(const fm_workbook_t* wb, std::size_t index) {
  const char* name = nullptr;
  const char* display = nullptr;
  fm_cell_rect_t rect{};
  std::size_t sheet = 0;
  EXPECT_EQ(fm_workbook_get_table(wb, index, &name, &display, &rect, &sheet), 0) << fm_get_last_error_message();
  return name != nullptr ? name : "";
}

/// The name of the table `id` resolves to, or "" when it resolves to none.
std::string ResolvedTableName(const fm_workbook_t* wb, std::uint32_t id) {
  std::size_t index = 0;
  return fm_workbook_resolve_table_runtime_id(wb, id, &index) == 0 ? TableName(wb, index) : std::string();
}

/// Adds a one-column table `name` on `sheet` over rows `row`..`row + 1` of column A.
void AddOneColumnTable(fm_workbook_t* wb, std::size_t sheet, std::uint32_t row, const char* name) {
  const char* columns[] = {"Value"};
  ASSERT_EQ(fm_sheet_set_text(wb, sheet, row, 0, "Value"), 0);
  ASSERT_EQ(fm_workbook_add_table(wb, sheet, record_arg(fm_cell_rect_t{row, 0, row + 1U, 0}), name, name, columns, 1,
                                  nullptr, 1, 0, ignored_index()),
            0)
      << fm_get_last_error_message();
}

TEST(FormulonCApiTable, RuntimeIdFollowsTheTableAcrossListRebuilds) {
  WorkbookGuard wb;
  CreateSalesTable(wb);
  std::size_t other = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", &other), 0) << fm_get_last_error_message();
  AddOneColumnTable(wb.handle, other, 0U, "Tracked");
  const std::uint32_t tracked = TableRuntimeId(wb.handle, 1);
  const std::uint32_t sales = TableRuntimeId(wb.handle, 0);
  EXPECT_NE(tracked, 0U);
  EXPECT_NE(sales, 0U);
  EXPECT_NE(tracked, sales);
  EXPECT_EQ(TableRuntimeId(wb.handle, 1), tracked);

  std::size_t index = 99;
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, 0U, &index), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(index, 99U);

  // Enough appends to reallocate the list.
  for (std::uint32_t i = 0; i < 12U; ++i) {
    const std::string name = "Extra" + std::to_string(i);
    AddOneColumnTable(wb.handle, other, 3U * (i + 1U), name.c_str());
  }
  EXPECT_EQ(ResolvedTableName(wb.handle, tracked), "Tracked");

  // A row deletion on another sheet rebuilds the list.
  ASSERT_EQ(fm_sheet_delete_rows(wb.handle, 0, 5U, 2U), 0) << fm_get_last_error_message();
  EXPECT_EQ(ResolvedTableName(wb.handle, tracked), "Tracked");

  // Removing another sheet, with a table of its own, rebuilds it too.
  std::size_t doomed = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Doomed", &doomed), 0) << fm_get_last_error_message();
  AddOneColumnTable(wb.handle, doomed, 0U, "Gone");
  const std::uint32_t gone = TableRuntimeId(wb.handle, 14);
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, doomed), 0) << fm_get_last_error_message();
  EXPECT_EQ(ResolvedTableName(wb.handle, tracked), "Tracked");
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, gone, &index), FM_STATUS_NOT_FOUND);

  // Removing the preceding table shifts the tracked one down by one.
  ASSERT_EQ(fm_workbook_remove_table(wb.handle, 0), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, tracked, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(index, 0U);
  EXPECT_EQ(TableName(wb.handle, index), "Tracked");
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, sales, &index), FM_STATUS_NOT_FOUND);

  // Rewriting the table in place keeps its identity.
  ASSERT_EQ(fm_workbook_rename_table(wb.handle, 0, "Renamed", "Renamed"), 0) << fm_get_last_error_message();
  EXPECT_EQ(ResolvedTableName(wb.handle, tracked), "Renamed");

  std::uint32_t out = 0;
  EXPECT_EQ(fm_workbook_get_table_runtime_id(wb.handle, 99, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_get_table_runtime_id(wb.handle, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_resolve_table_runtime_id(wb.handle, tracked, nullptr), kBindingNullPointer);
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
  expect_new(TableRuntimeId(wb.handle, 0));
  expect_new(TableRuntimeId(wb.handle, 1));

  // Added.
  AddOneColumnTable(wb.handle, 0, 20U, "Third");
  expect_new(TableRuntimeId(wb.handle, 2));

  // Copied: the copies are new tables; the originals keep their ids.
  std::size_t copy = 99;
  ASSERT_EQ(fm_workbook_copy_sheet(wb.handle, 0, 1, "Copy", &copy), 0) << fm_get_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_table_count(wb.handle, &count), 0);
  ASSERT_EQ(count, 6U);
  for (std::size_t i = 0; i < 3U; ++i) {
    EXPECT_EQ(TableRuntimeId(wb.handle, i), seen[i]);
  }
  for (std::size_t i = 3; i < 6U; ++i) {
    expect_new(TableRuntimeId(wb.handle, i));
  }

  // Added after a removal: the removed id is not reused.
  const std::uint32_t removed = TableRuntimeId(wb.handle, 2);
  ASSERT_EQ(fm_workbook_remove_table(wb.handle, 2), 0) << fm_get_last_error_message();
  AddOneColumnTable(wb.handle, 0, 30U, "Fourth");
  ASSERT_EQ(fm_workbook_get_table_count(wb.handle, &count), 0);
  const std::uint32_t fresh = TableRuntimeId(wb.handle, count - 1U);
  EXPECT_NE(fresh, removed);
  expect_new(fresh);
}

}  // namespace
