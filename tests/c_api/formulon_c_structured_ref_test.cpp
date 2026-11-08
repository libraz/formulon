// C ABI tests for structured references as the formula bar spells them: the
// getters return the display form for the formula's cell, the setters bind
// that form back to the stored one, and a table column renames through
// `fm_sheet_rename_table_column` or by writing its header cell.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cell.h"
#include "formulon_c_evaluate_top_left.h"
#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_setters.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kBindingUtf8EncodingError =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingUtf8EncodingError);
constexpr fm_status_t kParserUnexpectedToken =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kParserUnexpectedToken);

// Cells used throughout: E2 is inside tblSales (its Calc column), H2 is
// outside it, and B2 on Sheet2 is on another sheet.
constexpr std::uint32_t kInsideRow = 1;
constexpr std::uint32_t kInsideCol = 4;
constexpr std::uint32_t kOutsideRow = 1;
constexpr std::uint32_t kOutsideCol = 7;
constexpr std::size_t kOtherSheet = 1;

/// Sheet1 with table tblSales over A1:E5 (columns Qty, Unit Cost, a#b, Price,
/// Calc; three data rows; a totals row), and an empty Sheet2. `profile` NULL
/// creates the default profile.
void BuildSales(WorkbookGuard& wb, const char* profile = nullptr) {
  if (profile == nullptr) {
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  } else {
    ASSERT_EQ(fm_workbook_create_for_profile(profile, &wb.handle), 0) << fm_get_last_error_message();
  }
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", ignored_index()), 0);
  const char* columns[] = {"Qty", "Unit Cost", "a#b", "Price", "Calc"};
  for (std::uint32_t c = 0; c < 5U; ++c) {
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, c, columns[c]), 0);
  }
  for (std::uint32_t r = 1; r <= 3U; ++r) {
    for (std::uint32_t c = 0; c < 4U; ++c) {
      ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, r, c, static_cast<double>(r + 1U + 10U * c)), 0);
    }
  }
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_add_table(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 4, 4}), "tblSales", columns, 5, "", 1, 1,
                               &index),
            0)
      << fm_get_last_error_message();
}

std::string Formula(const WorkbookGuard& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const char* text = nullptr;
  EXPECT_EQ(fm_sheet_get_formula(wb.handle, sheet, row, col, &text), 0) << fm_get_last_error_message();
  return text != nullptr ? text : "";
}

std::string FormulaR1C1(const WorkbookGuard& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const char* text = nullptr;
  EXPECT_EQ(fm_sheet_get_formula_r1c1(wb.handle, sheet, row, col, &text), 0) << fm_get_last_error_message();
  return text != nullptr ? text : "";
}

std::string CellAtFormula(const WorkbookGuard& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_cell_count(wb.handle, sheet, &count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    std::uint32_t r = 0;
    std::uint32_t c = 0;
    const char* formula = nullptr;
    fm_value_t value{};
    EXPECT_EQ(fm_sheet_get_cell_at(wb.handle, sheet, i, &r, &c, &formula, &value), 0);
    if (r == row && c == col) {
      return formula != nullptr ? formula : "";
    }
  }
  ADD_FAILURE() << "no cell at row=" << row << " col=" << col;
  return "";
}

std::string RangeFormula(const WorkbookGuard& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  fm_cell_range_t* range = nullptr;
  EXPECT_EQ(fm_sheet_get_cells_in_range(wb.handle, sheet, record_arg(fm_cell_rect_t{row, col, row, col}), 0, 0, &range),
            0);
  std::string out;
  std::size_t count = 0;
  EXPECT_EQ(fm_cell_range_get_cell_count(range, &count), 0);
  if (count == 1U) {
    std::uint32_t r = 0;
    std::uint32_t c = 0;
    const char* formula = nullptr;
    fm_value_t value{};
    EXPECT_EQ(fm_cell_range_get_cell_at(range, 0, &r, &c, &formula, &value), 0);
    out = formula != nullptr ? formula : "";
  }
  fm_cell_range_destroy(range);
  return out;
}

std::string Stored(const WorkbookGuard& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  const formulon::Cell* cell = wb.handle->workbook().sheet(sheet).cell_at(row, col);
  return cell != nullptr ? cell->formula_text : "";
}

std::string ColumnName(const WorkbookGuard& wb, std::size_t column) {
  fm_table_column_info_t info = FM_STRUCT_INIT(fm_table_column_info_t);
  EXPECT_EQ(fm_sheet_get_table_column(wb.handle, 0, 0, column, &info), 0);
  return info.name != nullptr ? info.name : "";
}

std::string CellText(const WorkbookGuard& wb, std::uint32_t row, std::uint32_t col) {
  fm_value_t value{};
  EXPECT_EQ(fm_sheet_get_value(wb.handle, 0, row, col, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_TEXT);
  return value.kind == FM_VALUE_KIND_TEXT ? value.u.text : "";
}

std::string Rename(const WorkbookGuard& wb, std::size_t column, const char* name) {
  const char* out = nullptr;
  EXPECT_EQ(fm_sheet_rename_table_column(wb.handle, 0, 0, column, name, &out), 0) << fm_get_last_error_message();
  return out != nullptr ? out : "";
}

std::string DefinedName(const WorkbookGuard& wb, const std::string& wanted) {
  const std::size_t count = count_of(fm_workbook_get_defined_name_count, wb.handle);
  for (std::size_t i = 0; i < count; ++i) {
    fm_defined_name_t entry = FM_STRUCT_INIT(fm_defined_name_t);
    EXPECT_EQ(fm_workbook_get_defined_name_at(wb.handle, i, &entry), 0);
    if (wanted == entry.name) {
      return entry.formula;
    }
  }
  ADD_FAILURE() << "no defined name " << wanted;
  return "";
}

struct DisplayCase {
  const char* stored;
  const char* inside;
  const char* outside;
};

// Excel's Range.Formula for each stored formula, in the table's own range
// and outside it (the other sheet reads as outside).
constexpr DisplayCase kDisplayCases[] = {
    {"=tblSales[[#This Row],[Qty]]", "=[@Qty]", "=tblSales[@Qty]"},
    {"=tblSales[[#This Row],[Unit Cost]]", "=[@[Unit Cost]]", "=tblSales[@[Unit Cost]]"},
    {"=SUM(tblSales[Qty])", "=SUM([Qty])", "=SUM(tblSales[Qty])"},
    {"=SUM(tblSales[Unit Cost])", "=SUM([Unit Cost])", "=SUM(tblSales[Unit Cost])"},
    {"=tblSales[[#This Row],[a'#b]]", "=[@[a'#b]]", "=tblSales[@[a'#b]]"},
    {"=SUM(tblSales[[#This Row],[Qty]:[Price]])", "=SUM(tblSales[@[Qty]:[Price]])", "=SUM(tblSales[@[Qty]:[Price]])"},
    {"=SUM(tblSales[#This Row])", "=SUM(tblSales[@])", "=SUM(tblSales[@])"},
    {"=tblSales[[#Headers],[Qty]]", "=tblSales[[#Headers],[Qty]]", "=tblSales[[#Headers],[Qty]]"},
    {"=tblSales[[#Totals],[Qty]]", "=tblSales[[#Totals],[Qty]]", "=tblSales[[#Totals],[Qty]]"},
    {"=ROWS(tblSales[#All])", "=ROWS(tblSales[#All])", "=ROWS(tblSales[#All])"},
    {"=ROWS(tblSales[])", "=ROWS(tblSales)", "=ROWS(tblSales)"},
    {"=ROWS(tblSales[[#Headers],[#Data],[Qty]])", "=ROWS(tblSales[[#Headers],[#Data],[Qty]])",
     "=ROWS(tblSales[[#Headers],[#Data],[Qty]])"},
    {"=SUM(tblSales[[Qty]:[Price]])", "=SUM(tblSales[[Qty]:[Price]])", "=SUM(tblSales[[Qty]:[Price]])"},
};

TEST(FormulonCApiStructuredRef, GettersReturnTheDisplayFormForTheCell) {
  WorkbookGuard wb;
  BuildSales(wb);
  for (const DisplayCase& c : kDisplayCases) {
    SCOPED_TRACE(c.stored);
    ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kInsideRow, kInsideCol, c.stored), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kOutsideRow, kOutsideCol, c.stored), 0);
    ASSERT_EQ(fm_sheet_set_formula(wb.handle, kOtherSheet, 1, 1, c.stored), 0);
    EXPECT_EQ(Stored(wb, 0, kInsideRow, kInsideCol), c.stored);

    EXPECT_EQ(Formula(wb, 0, kInsideRow, kInsideCol), c.inside);
    EXPECT_EQ(Formula(wb, 0, kOutsideRow, kOutsideCol), c.outside);
    EXPECT_EQ(Formula(wb, kOtherSheet, 1, 1), c.outside);

    EXPECT_EQ(FormulaR1C1(wb, 0, kInsideRow, kInsideCol), c.inside);
    EXPECT_EQ(FormulaR1C1(wb, 0, kOutsideRow, kOutsideCol), c.outside);

    EXPECT_EQ(CellAtFormula(wb, 0, kInsideRow, kInsideCol), c.inside);
    EXPECT_EQ(CellAtFormula(wb, kOtherSheet, 1, 1), c.outside);
    EXPECT_EQ(RangeFormula(wb, 0, kInsideRow, kInsideCol), c.inside);
    EXPECT_EQ(RangeFormula(wb, 0, kOutsideRow, kOutsideCol), c.outside);
  }
}

TEST(FormulonCApiStructuredRef, ArrayFormulaReadsAtItsAnchor) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(
      fm_sheet_set_array_formula(
          wb.handle, 0, record_arg(fm_cell_rect_t{kInsideRow, kInsideCol, kInsideRow + 1 - 1U, kInsideCol + 1 - 1U}),
          "=SUM(tblSales[Qty])"),
      0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_array_formula(
                wb.handle, 0,
                record_arg(fm_cell_rect_t{kOutsideRow, kOutsideCol, kOutsideRow + 1 - 1U, kOutsideCol + 1 - 1U}),
                "=SUM(tblSales[Qty])"),
            0);
  fm_array_formula_t info = FM_STRUCT_INIT(fm_array_formula_t);
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, kInsideRow, kInsideCol, &info), 0);
  ASSERT_EQ(info.engaged, 1);
  EXPECT_STREQ(info.formula, "=SUM([Qty])");
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, kOutsideRow, kOutsideCol, &info), 0);
  ASSERT_EQ(info.engaged, 1);
  EXPECT_STREQ(info.formula, "=SUM(tblSales[Qty])");
}

TEST(FormulonCApiStructuredRef, TotalsRowShowsTheInsideForm) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_set_table_column_totals(wb.handle, 0, 0, 0, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr),
            0);
  EXPECT_EQ(Stored(wb, 0, 4, 0), "=SUBTOTAL(109,tblSales[Qty])");
  EXPECT_EQ(Formula(wb, 0, 4, 0), "=SUBTOTAL(109,[Qty])");

  // A custom formula entered in the display form binds at the totals cell.
  ASSERT_EQ(fm_sheet_set_table_column_totals(wb.handle, 0, 0, 3, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr,
                                             "=SUM([Price])*2"),
            0)
      << fm_get_last_error_message();
  fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
  ASSERT_EQ(fm_sheet_get_table_column(wb.handle, 0, 0, 3, &column), 0);
  EXPECT_STREQ(column.totals_row_formula, "=SUM([Price])*2");
  EXPECT_EQ(wb.handle->workbook().tables()[0].columns[3].totals_row_formula, "SUM(tblSales[Price])*2");
  EXPECT_EQ(Stored(wb, 0, 4, 3), "=SUM(tblSales[Price])*2");
}

TEST(FormulonCApiStructuredRef, DefinedNamesShowTheOutsideForm) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(SetDefinedName(wb.handle, "NmR", "=tblSales[[#This Row],[Qty]]"), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetDefinedName(wb.handle, "NmH", "=tblSales[[#Headers],[Qty]]"), 0);
  EXPECT_EQ(DefinedName(wb, "NmR"), "=tblSales[@Qty]");
  EXPECT_EQ(DefinedName(wb, "NmH"), "=tblSales[[#Headers],[Qty]]");
}

TEST(FormulonCApiStructuredRef, DisplayFormSetsBackToTheStoredForm) {
  WorkbookGuard wb;
  BuildSales(wb);
  for (const DisplayCase& c : kDisplayCases) {
    SCOPED_TRACE(c.stored);
    ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kInsideRow, kInsideCol, c.inside), 0) << fm_get_last_error_message();
    EXPECT_EQ(Stored(wb, 0, kInsideRow, kInsideCol), c.stored);
    ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kOutsideRow, kOutsideCol, c.outside), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(Stored(wb, 0, kOutsideRow, kOutsideCol), c.stored);
  }
  // A bare table name entered is a whole-table reference.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, kOtherSheet, 1, 1, "=ROWS(tblSales)"), 0);
  EXPECT_EQ(Stored(wb, kOtherSheet, 1, 1), "=ROWS(tblSales[])");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, kOtherSheet, 1, 1, &v), 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 3.0);
}

TEST(FormulonCApiStructuredRef, UnqualifiedReferenceOutsideEveryTableIsRefused) {
  WorkbookGuard wb;
  BuildSales(wb);
  EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, kOutsideRow, kOutsideCol, "=[Qty]"), kParserUnexpectedToken);
  EXPECT_EQ(Stored(wb, 0, kOutsideRow, kOutsideCol), "");
  EXPECT_EQ(
      fm_sheet_set_array_formula(
          wb.handle, 0,
          record_arg(fm_cell_rect_t{kOutsideRow, kOutsideCol, kOutsideRow + 1 - 1U, kOutsideCol + 1 - 1U}), "=[@Qty]"),
      kParserUnexpectedToken);
}

TEST(FormulonCApiStructuredRef, EvaluateFormulaBindsAtItsCell) {
  WorkbookGuard wb;
  BuildSales(wb);
  fm_value_t v{};
  // Row 3 holds Qty 3.
  ASSERT_EQ(evaluate_top_left(wb.handle, 0, 2, kInsideCol, "=[@Qty]*10", &v), 0) << fm_get_last_error_message();
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 30.0);
  ASSERT_EQ(evaluate_top_left(wb.handle, 0, kOutsideRow, kOutsideCol, "=ROWS(tblSales)", &v), 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 3.0);
}

TEST(FormulonCApiStructuredRef, RenameColumnRefusesBadArguments) {
  WorkbookGuard wb;
  BuildSales(wb);
  const char* out = nullptr;
  EXPECT_EQ(fm_sheet_rename_table_column(nullptr, 0, 0, 0, "X", &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_rename_table_column(wb.handle, 0, 0, 0, nullptr, &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_rename_table_column(wb.handle, 0, 0, 0, "X", nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_rename_table_column(wb.handle, 0, 1, 0, "X", &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_rename_table_column(wb.handle, 0, 0, 5, "X", &out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_rename_table_column(wb.handle, 9, 0, 0, "X", &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_rename_table_column(wb.handle, 0, 0, 0, "\xC3\x28", &out), kBindingUtf8EncodingError);
  EXPECT_EQ(ColumnName(wb, 0), "Qty");
}

TEST(FormulonCApiStructuredRef, RenameColumnRewritesEveryReference) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kInsideRow, kInsideCol, "=tblSales[[#This Row],[Qty]]"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, kInsideCol, "=SUM(tblSales[[#This Row],[Qty]:[Price]])"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, kOtherSheet, 1, 1, "=SUM(tblSales[Qty])"), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "NmQ", "=tblSales[Qty]"), 0);
  ASSERT_EQ(fm_sheet_set_table_column_totals(wb.handle, 0, 0, 0, FM_TABLE_TOTALS_ROW_FUNCTION_SUM, nullptr, nullptr),
            0);

  EXPECT_EQ(Rename(wb, 0, "Amount"), "Amount");
  EXPECT_EQ(ColumnName(wb, 0), "Amount");
  EXPECT_EQ(CellText(wb, 0, 0), "Amount");
  EXPECT_EQ(Formula(wb, 0, kInsideRow, kInsideCol), "=[@Amount]");
  EXPECT_EQ(Stored(wb, 0, kInsideRow, kInsideCol), "=tblSales[[#This Row],[Amount]]");
  EXPECT_EQ(Formula(wb, 0, 2, kInsideCol), "=SUM(tblSales[@[Amount]:[Price]])");
  EXPECT_EQ(Formula(wb, kOtherSheet, 1, 1), "=SUM(tblSales[Amount])");
  EXPECT_EQ(DefinedName(wb, "NmQ"), "=tblSales[Amount]");
  EXPECT_EQ(Formula(wb, 0, 4, 0), "=SUBTOTAL(109,[Amount])");

  EXPECT_EQ(Rename(wb, 1, "Unit-Cost"), "Unit-Cost");
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kOutsideRow, kOutsideCol, "=[@[Unit-Cost]]"), kParserUnexpectedToken);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kInsideRow, kInsideCol, "=[@[Unit-Cost]]"), 0);
  EXPECT_EQ(Stored(wb, 0, kInsideRow, kInsideCol), "=tblSales[[#This Row],[Unit-Cost]]");
}

std::vector<std::uint8_t> SaveAs(const WorkbookGuard& wb, int32_t format) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save_as(wb.handle, format, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

// Excel's save after typing Amount over Qty: the table part names the column
// and keeps the custom totals formula in long form, and cells store long form.
TEST(FormulonCApiStructuredRef, RenamedColumnSavesInExcelsStoredForm) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kInsideRow, kInsideCol, "=[@Qty]"), 0);
  ASSERT_EQ(fm_sheet_set_table_column_totals(wb.handle, 0, 0, 0, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr,
                                             "=SUBTOTAL(109,[Qty])"),
            0)
      << fm_get_last_error_message();
  ASSERT_EQ(Rename(wb, 0, "Amount"), "Amount");

  const std::vector<std::uint8_t> xlsx = SaveAs(wb, FM_WORKBOOK_FORMAT_XLSX);
  const std::string table = ExtractZipEntry(xlsx, "xl/tables/table1.xml");
  EXPECT_NE(table.find("<tableColumn id=\"1\" name=\"Amount\" totalsRowFunction=\"custom\">"), std::string::npos)
      << table;
  EXPECT_NE(table.find("<totalsRowFormula>SUBTOTAL(109,tblSales[Amount])</totalsRowFormula>"), std::string::npos)
      << table;
  const std::string sheet = ExtractZipEntry(xlsx, "xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet.find("<f>tblSales[[#This Row],[Amount]]</f>"), std::string::npos) << sheet;
  EXPECT_NE(sheet.find("<f>SUBTOTAL(109,tblSales[Amount])</f>"), std::string::npos) << sheet;

  for (const int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    const std::vector<std::uint8_t> bytes = SaveAs(wb, format);
    WorkbookGuard reloaded;
    ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &reloaded.handle), 0) << "format=" << format;
    EXPECT_EQ(ColumnName(reloaded, 0), "Amount") << "format=" << format;
    EXPECT_EQ(Stored(reloaded, 0, kInsideRow, kInsideCol), "=tblSales[[#This Row],[Amount]]") << "format=" << format;
    EXPECT_EQ(Formula(reloaded, 0, kInsideRow, kInsideCol), "=[@Amount]") << "format=" << format;
    EXPECT_EQ(Formula(reloaded, 0, 4, 0), "=SUBTOTAL(109,[Amount])") << "format=" << format;
  }
}

TEST(FormulonCApiStructuredRef, RenameColumnNumbersEmptyAndDuplicateNames) {
  WorkbookGuard wb;
  BuildSales(wb);
  // An empty name takes the lowest unused `Column<N>`.
  EXPECT_EQ(Rename(wb, 3, ""), "Column1");
  EXPECT_EQ(CellText(wb, 0, 3), "Column1");
  EXPECT_EQ(Rename(wb, 2, ""), "Column2");
  // A duplicate (ASCII case-insensitive) appends N >= 2 to the whole name.
  EXPECT_EQ(Rename(wb, 2, "qty"), "qty2");
  EXPECT_EQ(Rename(wb, 3, "Qty2"), "Qty22");
  // The same name rewrites nothing and still writes the header cell.
  EXPECT_EQ(Rename(wb, 1, "Unit Cost"), "Unit Cost");
  EXPECT_EQ(CellText(wb, 0, 1), "Unit Cost");
  // A case-only change renames; spaces are kept.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kInsideRow, kInsideCol, "=[@Qty]"), 0);
  EXPECT_EQ(Rename(wb, 0, "qty"), "qty");
  EXPECT_EQ(Formula(wb, 0, kInsideRow, kInsideCol), "=[@qty]");
  EXPECT_EQ(Rename(wb, 1, "  pad  "), "  pad  ");
  EXPECT_EQ(ColumnName(wb, 1), "  pad  ");
}

TEST(FormulonCApiStructuredRef, RenameColumnUsesTheProfileColumnWord) {
  WorkbookGuard wb;
  BuildSales(wb, "mac-365-ja_JP");
  EXPECT_EQ(Rename(wb, 3, ""), "列1");
  EXPECT_EQ(Rename(wb, 2, ""), "列2");
}

TEST(FormulonCApiStructuredRef, HeaderCellValueSettersRenameTheColumn) {
  WorkbookGuard wb;
  BuildSales(wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kInsideRow, kInsideCol, "=[@Qty]+[@[Unit Cost]]"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, kOutsideRow, kOutsideCol, "=SUM(tblSales[Qty])"), 0);

  // A number becomes its General text, and references follow.
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 2024.0), 0) << fm_get_last_error_message();
  EXPECT_EQ(ColumnName(wb, 0), "2024");
  EXPECT_EQ(CellText(wb, 0, 0), "2024");
  EXPECT_EQ(Formula(wb, 0, kOutsideRow, kOutsideCol), "=SUM(tblSales[2024])");

  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Amount"), 0);
  EXPECT_EQ(ColumnName(wb, 0), "Amount");
  EXPECT_EQ(Formula(wb, 0, kInsideRow, kInsideCol), "=[@Amount]+[@[Unit Cost]]");

  // A blank header takes the default name.
  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 0, 1), 0);
  EXPECT_EQ(ColumnName(wb, 1), "Column1");
  EXPECT_EQ(CellText(wb, 0, 1), "Column1");
  EXPECT_EQ(Formula(wb, 0, kInsideRow, kInsideCol), "=[@Amount]+[@Column1]");

  ASSERT_EQ(fm_sheet_set_bool(wb.handle, 0, 0, 2, 1), 0);
  EXPECT_EQ(ColumnName(wb, 2), "TRUE");
  EXPECT_EQ(CellText(wb, 0, 2), "TRUE");
  ASSERT_EQ(fm_sheet_set_error(wb.handle, 0, 0, 3, static_cast<fm_error_code_t>(formulon::ErrorCode::NA)), 0);
  EXPECT_EQ(ColumnName(wb, 3), "#N/A");
  EXPECT_EQ(CellText(wb, 0, 3), "#N/A");

  // A data row and a cell beside the table take the value itself.
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 0, 2024.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, kOutsideCol, 2024.0), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 1, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, kOutsideCol, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(ColumnName(wb, 0), "Amount");
}

}  // namespace
