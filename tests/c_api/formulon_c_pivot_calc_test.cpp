// Calculated fields and calculated items through the C ABI: the grids and the
// saved strings Mac Excel produced for the same operations on the same source
// (Region / Product / Amt / Qty over five rows), the solve order at an
// intersection, the functions that exclude calculated items, and the argument
// boundaries of every calculated-field / -item entry point.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <optional>
#include <string>
#include <vector>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "value.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kBindingUtf8EncodingError =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingUtf8EncodingError);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);

constexpr std::size_t kPivot = 0U;
constexpr std::size_t kRegion = 0U;
constexpr std::size_t kProduct = 1U;
constexpr std::size_t kAmt = 2U;
constexpr std::size_t kQty = 3U;

const char* const kGrandTotal = "\xE7\xB7\x8F\xE8\xA8\x88";  // 総計
const char* const kEmoji = "\xF0\x9F\x98\x80";               // two UTF-16 code units

// "合計 / <name>", the ja-JP name of a Sum data field.
std::string SumOf(const std::string& name) {
  return "\xE5\x90\x88\xE8\xA8\x88 / " + name;
}

struct SourceRow {
  double region;   // North 0, South 1, East 2
  double product;  // A 0, B 1
  double amt;
  double qty;
};

constexpr SourceRow kRows[] = {{0, 0, 120, 4}, {0, 1, 90, 3}, {1, 0, 300, 5}, {1, 1, 50, 2}, {2, 0, 70, 6}};

fm_pivot_data_field_spec_t DataSpec(std::size_t field, fm_pivot_aggregation_t aggregation) {
  fm_pivot_data_field_spec_t spec{};
  spec.name = nullptr;
  spec.field_index = static_cast<std::uint32_t>(field);
  spec.aggregation = aggregation;
  spec.number_format = "";
  spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  spec.show_as_base_field = -1;
  spec.show_as_base_item = -1;
  return spec;
}

void AddField(fm_workbook_t* wb, const char* name, fm_pivot_axis_t axis) {
  fm_pivot_field_spec_t spec{};
  spec.source_name = name;
  spec.custom_name = "";
  spec.axis = axis;
  spec.subtotal_top = 1;
  spec.number_format = "";
  std::size_t index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb, 0U, kPivot, &spec, &index), 0) << fm_get_last_error_message();
}

// The measured source in a ja-JP workbook: Region on rows, Sum of Amt and Sum
// of Qty as data fields; with `product_on_cols`, Product on columns and Amt
// alone. The pivot sits at A3, as in the Excel files.
void BuildSales(WorkbookGuard& wb, std::uint32_t* out_cache, bool product_on_cols = false) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::uint32_t cache = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache, 1, "A1:D6", "Sheet1", nullptr), 0);
  for (const char* name : {"Region", "Product", "Amt", "Qty"}) {
    std::size_t index = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache, name, &index), 0) << fm_get_last_error_message();
  }
  for (const char* item : {"North", "South", "East"}) {
    ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache, kRegion, item), 0);
  }
  for (const char* item : {"A", "B"}) {
    ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache, kProduct, item), 0);
  }
  for (const SourceRow& row : kRows) {
    std::size_t record = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache, &record), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache, record, kRegion, row.region), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache, record, kProduct, row.product), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache, record, kAmt, row.amt), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache, record, kQty, row.qty), 0);
  }
  std::size_t pivot = 99;
  ASSERT_EQ(fm_sheet_add_pivot(wb.handle, 0U, "PivotTable1", cache, 2U, 0U, &pivot), 0) << fm_get_last_error_message();
  AddField(wb.handle, "Region", FM_PIVOT_AXIS_ROW);
  AddField(wb.handle, "Amt", FM_PIVOT_AXIS_VALUE);
  // As Excel creates the field: manual order, its items listed sorted.
  ASSERT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0U, kPivot, kRegion, FM_PIVOT_SORT_MANUAL, nullptr), 0)
      << fm_get_last_error_message();
  std::size_t expected_item = 0;
  for (const char* item : {"East", "North", "South"}) {
    std::size_t item_index = 99;
    ASSERT_EQ(fm_sheet_add_pivot_field_item(wb.handle, 0U, kPivot, kRegion, item, 1, &item_index), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(item_index, expected_item++);
  }
  const std::uint32_t rows[] = {static_cast<std::uint32_t>(kRegion)};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0U, kPivot, rows, 1U), 0) << fm_get_last_error_message();
  if (product_on_cols) {
    AddField(wb.handle, "Product", FM_PIVOT_AXIS_COL);
    ASSERT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0U, kPivot, kProduct, FM_PIVOT_SORT_MANUAL, nullptr), 0);
    for (const char* item : {"A", "B"}) {
      ASSERT_EQ(fm_sheet_add_pivot_field_item(wb.handle, 0U, kPivot, kProduct, item, 1, ignored_index()), 0);
    }
    const std::uint32_t cols[] = {static_cast<std::uint32_t>(kProduct)};
    ASSERT_EQ(fm_sheet_set_pivot_column_field_order(wb.handle, 0U, kPivot, cols, 1U), 0) << fm_get_last_error_message();
  } else {
    AddField(wb.handle, "Qty", FM_PIVOT_AXIS_VALUE);
  }
  std::size_t data = 99;
  fm_pivot_data_field_spec_t amt = DataSpec(kAmt, FM_PIVOT_AGG_SUM);
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, kPivot, &amt, &data), 0) << fm_get_last_error_message();
  if (!product_on_cols) {
    fm_pivot_data_field_spec_t qty = DataSpec(kQty, FM_PIVOT_AGG_SUM);
    ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, kPivot, &qty, &data), 0) << fm_get_last_error_message();
  }
  *out_cache = cache;
}

void LoadFixture(const std::string& name, WorkbookGuard* out) {
  std::string dir = __FILE__;
  dir = dir.substr(0, dir.find_last_of("/\\"));
  std::ifstream in(dir + "/../fixtures/excel/" + name, std::ios::binary);
  ASSERT_TRUE(in.good()) << name;
  const std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out->handle), 0) << fm_get_last_error_message();
  // The files were saved by a ja-JP Excel; the profile is not stored in them.
  ASSERT_EQ(fm_workbook_set_excel_profile_id(out->handle, "win-365-ja_JP"), 0);
}

// One report cell, copied out of the layout handle.
struct Shown {
  int32_t kind = FM_VAL_BLANK;
  double number = 0.0;
  std::int32_t error = 0;
  std::string text;
  bool found = false;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
};

Shown Copy(const fm_value_t& value) {
  Shown out;
  out.kind = value.kind;
  if (value.kind == FM_VAL_NUMBER) {
    out.number = value.u.number;
  } else if (value.kind == FM_VAL_ERROR) {
    out.error = value.u.error_code;
  } else if (value.kind == FM_VAL_TEXT && value.u.text != nullptr) {
    out.text = value.u.text;
  }
  return out;
}

// The report cell in the row labelled `row_label` (first report column) and
// the column headed `column` (any other column), as the layout API draws it.
Shown ReportAt(fm_workbook_t* wb, const std::string& row_label, const std::string& column) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_sheet_get_pivot_layout(wb, 0U, kPivot, &projected.handle), 0) << fm_get_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  EXPECT_EQ(PivotLayoutExtent(projected.handle, &top, &left, &rows, &cols), 0);
  const std::vector<fm_pivot_cell_t> cells = CollectCells(projected.handle);
  const auto is_text = [](const fm_pivot_cell_t& cell, const std::string& text) {
    return cell.value.kind == FM_VAL_TEXT && cell.value.u.text != nullptr && text == cell.value.u.text;
  };
  std::optional<std::uint32_t> row;
  std::optional<std::uint32_t> col;
  for (const fm_pivot_cell_t& cell : cells) {
    if (!row.has_value() && cell.col == left && is_text(cell, row_label)) {
      row = cell.row;
    }
    if (!col.has_value() && cell.col != left && is_text(cell, column)) {
      col = cell.col;
    }
  }
  Shown out;
  if (!row.has_value() || !col.has_value()) {
    ADD_FAILURE() << "no report cell at row '" << row_label << "', column '" << column << "'";
    return out;
  }
  for (const fm_pivot_cell_t& cell : cells) {
    if (cell.row == *row && cell.col == *col) {
      out = Copy(cell.value);
    }
  }
  out.found = true;
  out.row = *row;
  out.col = *col;
  return out;
}

void ExpectNumber(fm_workbook_t* wb, const std::string& row_label, const std::string& column, double expected) {
  const Shown shown = ReportAt(wb, row_label, column);
  ASSERT_EQ(shown.kind, FM_VAL_NUMBER) << row_label << " / " << column;
  EXPECT_DOUBLE_EQ(shown.number, expected) << row_label << " / " << column;
  // The sheet holds the same value at the same cell.
  fm_value_t cell{};
  ASSERT_EQ(fm_sheet_get_value(wb, 0U, shown.row, shown.col, &cell), 0);
  ASSERT_EQ(cell.kind, FM_VAL_NUMBER) << row_label << " / " << column;
  EXPECT_DOUBLE_EQ(cell.u.number, expected) << row_label << " / " << column;
}

void ExpectError(fm_workbook_t* wb, const std::string& row_label, const std::string& column,
                 formulon::ErrorCode expected) {
  const Shown shown = ReportAt(wb, row_label, column);
  ASSERT_EQ(shown.kind, FM_VAL_ERROR) << row_label << " / " << column;
  EXPECT_EQ(shown.error, static_cast<std::int32_t>(expected)) << row_label << " / " << column;
}

std::int32_t Pending(const fm_workbook_t* wb) {
  std::int32_t pending = -1;
  EXPECT_EQ(fm_sheet_is_pivot_projection_pending(wb, 0U, kPivot, &pending), 0);
  return pending;
}

std::string SavedPart(fm_workbook_t* wb, const char* part) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  return ExtractZipEntry(std::vector<std::uint8_t>(saved.data, saved.data + saved.len), part);
}

std::size_t AddPrice(fm_workbook_t* wb, std::uint32_t cache) {
  std::size_t price = 99;
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_field(wb, cache, "Price", "=Amt/Qty", &price), 0)
      << fm_get_last_error_message();
  return price;
}

std::size_t AddDataField(fm_workbook_t* wb, std::size_t field, fm_pivot_aggregation_t aggregation) {
  fm_pivot_data_field_spec_t spec = DataSpec(field, aggregation);
  std::size_t index = 99;
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb, 0U, kPivot, &spec, &index), 0) << fm_get_last_error_message();
  return index;
}

std::string Repeat(const std::string& unit, std::size_t times) {
  std::string out;
  for (std::size_t i = 0; i < times; ++i) {
    out += unit;
  }
  return out;
}

}  // namespace

// ---------------------------------------------------------------------------
// Calculated fields
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotCalc, CalculatedFieldGridMatchesExcel) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  ASSERT_EQ(Pending(wb.handle), 0);
  const std::size_t price = AddPrice(wb.handle, cache);
  EXPECT_EQ(price, 4U);
  // A cache edit leaves the report pending; the table edit projects it.
  EXPECT_EQ(Pending(wb.handle), 1);
  const std::size_t data = AddDataField(wb.handle, price, FM_PIVOT_AGG_SUM);
  EXPECT_EQ(Pending(wb.handle), 0);

  fm_pivot_data_field_spec_t read{};
  ASSERT_EQ(fm_sheet_get_pivot_data_field(wb.handle, 0U, kPivot, data, &read), 0);
  EXPECT_EQ(std::string(read.name), SumOf("Price"));

  // CF1: the formula over the Sums, totals included.
  ExpectNumber(wb.handle, "East", SumOf("Price"), 70.0 / 6.0);
  ExpectNumber(wb.handle, "North", SumOf("Price"), 30.0);
  ExpectNumber(wb.handle, "South", SumOf("Price"), 50.0);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Price"), 31.5);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Amt"), 630.0);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Qty"), 20.0);

  const char* formula = nullptr;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, price, &formula), 0);
  EXPECT_STREQ(formula, "=Amt/Qty");
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, kAmt, &formula), 0);
  EXPECT_STREQ(formula, "");
}

TEST(FormulonCApiPivotCalc, CalculatedFieldSavesExcelStrings) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  AddDataField(wb.handle, AddPrice(wb.handle, cache), FM_PIVOT_AGG_SUM);

  const std::string cache_xml = SavedPart(wb.handle, "xl/pivotCache/pivotCacheDefinition1.xml");
  EXPECT_NE(cache_xml.find("<cacheField name=\"Price\" numFmtId=\"0\" formula=\"Amt/Qty\" databaseField=\"0\"/>"),
            std::string::npos)
      << cache_xml;
  const std::string pivot_xml = SavedPart(wb.handle, "xl/pivotTables/pivotTable1.xml");
  EXPECT_NE(pivot_xml.find("<pivotField dataField=\"1\" dragToRow=\"0\" dragToCol=\"0\" dragToPage=\"0\" "
                           "showAll=\"0\" defaultSubtotal=\"0\"/>"),
            std::string::npos)
      << pivot_xml;
  EXPECT_NE(pivot_xml.find("<dataField name=\"" + SumOf("Price") + "\" fld=\"4\""), std::string::npos) << pivot_xml;

  // The saved file evaluates the same after a reload.
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_excel_profile_id(loaded.handle, "win-365-ja_JP"), 0);
  ExpectNumber(loaded.handle, kGrandTotal, SumOf("Price"), 31.5);
}

TEST(FormulonCApiPivotCalc, ExcelSavedCalculatedFieldLoadsWithTheSameGrid) {
  WorkbookGuard wb;
  LoadFixture("pivot_calc/cf1_price_amt_per_qty.xlsx", &wb);
  ExpectNumber(wb.handle, "East", SumOf("Price"), 70.0 / 6.0);
  ExpectNumber(wb.handle, "North", SumOf("Price"), 30.0);
  ExpectNumber(wb.handle, "South", SumOf("Price"), 50.0);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Price"), 31.5);
}

TEST(FormulonCApiPivotCalc, DivisionByZeroShowsTheErrorCaption) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  std::size_t bad = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_field(wb.handle, cache, "Bad", "Amt/0", &bad), 0)
      << fm_get_last_error_message();
  AddDataField(wb.handle, bad, FM_PIVOT_AGG_SUM);
  // CF2: every cell and total is #DIV/0!.
  for (const char* row : {"East", "North", "South", kGrandTotal}) {
    ExpectError(wb.handle, row, SumOf("Bad"), formulon::ErrorCode::Div0);
  }

  fm_pivot_options_t options{};
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, kPivot, &options), 0);
  options.show_error = 1;
  options.error_caption = "ERR";
  ASSERT_EQ(fm_sheet_set_pivot_options(wb.handle, 0U, kPivot, &options), 0) << fm_get_last_error_message();
  for (const char* row : {"East", "North", "South", kGrandTotal}) {
    const Shown shown = ReportAt(wb.handle, row, SumOf("Bad"));
    ASSERT_EQ(shown.kind, FM_VAL_TEXT) << row;
    EXPECT_EQ(shown.text, "ERR") << row;
  }
  const std::string pivot_xml = SavedPart(wb.handle, "xl/pivotTables/pivotTable1.xml");
  EXPECT_NE(pivot_xml.find("showError=\"1\""), std::string::npos) << pivot_xml;
  EXPECT_NE(pivot_xml.find("errorCaption=\"ERR\""), std::string::npos) << pivot_xml;
}

TEST(FormulonCApiPivotCalc, QuotedNamesAndIfApplyToTheSums) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  // CF3: IF acts on the Sums, and a quoted name is stored as entered.
  std::size_t cond = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_field(wb.handle, cache, "Cond", "=IF(Qty>2,'Amt',0)", &cond), 0)
      << fm_get_last_error_message();
  AddDataField(wb.handle, cond, FM_PIVOT_AGG_SUM);
  ExpectNumber(wb.handle, "East", SumOf("Cond"), 70.0);
  ExpectNumber(wb.handle, "North", SumOf("Cond"), 210.0);
  ExpectNumber(wb.handle, "South", SumOf("Cond"), 350.0);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Cond"), 630.0);
  const std::string cache_xml = SavedPart(wb.handle, "xl/pivotCache/pivotCacheDefinition1.xml");
  EXPECT_NE(cache_xml.find("formula=\"IF(Qty>2,'Amt',0)\""), std::string::npos) << cache_xml;
}

TEST(FormulonCApiPivotCalc, DataFieldFunctionRenamesButKeepsTheSumBasedValues) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  const std::size_t data = AddDataField(wb.handle, AddPrice(wb.handle, cache), FM_PIVOT_AGG_SUM);
  // CF4: Average and Count rename the data field; the values stay.
  const struct {
    fm_pivot_aggregation_t aggregation;
    const char* name;
  } cases[] = {{FM_PIVOT_AGG_AVERAGE, "\xE5\xB9\xB3\xE5\x9D\x87 / Price"},  // 平均 / Price
               {FM_PIVOT_AGG_COUNT, "\xE5\x80\x8B\xE6\x95\xB0 / Price"}};   // 個数 / Price
  for (const auto& c : cases) {
    fm_pivot_data_field_spec_t spec = DataSpec(4U, c.aggregation);
    ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0U, kPivot, data, &spec), 0) << fm_get_last_error_message();
    fm_pivot_data_field_spec_t read{};
    ASSERT_EQ(fm_sheet_get_pivot_data_field(wb.handle, 0U, kPivot, data, &read), 0);
    EXPECT_STREQ(read.name, c.name);
    ExpectNumber(wb.handle, "North", c.name, 30.0);
    ExpectNumber(wb.handle, kGrandTotal, c.name, 31.5);
  }
}

TEST(FormulonCApiPivotCalc, CalculatedFieldOverAnotherCalculatedField) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  const std::size_t price = AddPrice(wb.handle, cache);
  std::size_t price2 = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_field(wb.handle, cache, "Price2", "Price*2", &price2), 0)
      << fm_get_last_error_message();
  AddDataField(wb.handle, price2, FM_PIVOT_AGG_SUM);
  // CF5.
  ExpectNumber(wb.handle, "East", SumOf("Price2"), 140.0 / 6.0);
  ExpectNumber(wb.handle, "North", SumOf("Price2"), 60.0);
  ExpectNumber(wb.handle, "South", SumOf("Price2"), 100.0);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Price2"), 63.0);

  // Price cannot go while Price2 names it, nor be made to name Price2.
  EXPECT_EQ(fm_workbook_remove_pivot_cache_calculated_field(wb.handle, cache, price), kPreconditionFailed);
  EXPECT_EQ(fm_workbook_set_pivot_cache_calculated_field_formula(wb.handle, cache, price, "Price2/2"),
            kInvalidArgument);
  EXPECT_EQ(fm_workbook_set_pivot_cache_calculated_field_formula(wb.handle, cache, price2, "Price2+1"),
            kInvalidArgument);

  // Removing Price2 takes its data field along; Price can go after it.
  ASSERT_EQ(fm_workbook_remove_pivot_cache_calculated_field(wb.handle, cache, price2), 0)
      << fm_get_last_error_message();
  std::size_t data_fields = 0;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0U, kPivot, &data_fields), 0);
  EXPECT_EQ(data_fields, 2U);
  ASSERT_EQ(fm_workbook_remove_pivot_cache_calculated_field(wb.handle, cache, price), 0) << fm_get_last_error_message();
  std::size_t fields = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_count(wb.handle, cache, &fields), 0);
  EXPECT_EQ(fields, 4U);
  ASSERT_EQ(fm_sheet_get_pivot_field_count(wb.handle, 0U, kPivot, &fields), 0);
  EXPECT_EQ(fields, 4U);
}

TEST(FormulonCApiPivotCalc, CalculatedFieldSetFormulaReevaluates) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  const std::size_t price = AddPrice(wb.handle, cache);
  AddDataField(wb.handle, price, FM_PIVOT_AGG_SUM);
  ASSERT_EQ(fm_workbook_set_pivot_cache_calculated_field_formula(wb.handle, cache, price, "=Amt*2"), 0)
      << fm_get_last_error_message();
  const char* formula = nullptr;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, price, &formula), 0);
  EXPECT_STREQ(formula, "=Amt*2");
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Price"), 1260.0);
}

TEST(FormulonCApiPivotCalc, CalculatedFieldArgumentBoundaries) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  std::size_t out = 99;
  const auto add = [&](const char* name, const char* formula) {
    return fm_workbook_add_pivot_cache_calculated_field(wb.handle, cache, name, formula, &out);
  };
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_field(nullptr, cache, "P", "Amt", &out), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_field(wb.handle, cache, "P", "Amt", nullptr), kBindingNullPointer);
  EXPECT_EQ(add(nullptr, "Amt"), kBindingNullPointer);
  EXPECT_EQ(add("P", nullptr), kBindingNullPointer);
  EXPECT_EQ(add("", "Amt"), kInvalidArgument);
  EXPECT_EQ(add("P", ""), kInvalidArgument);
  EXPECT_EQ(add("\xC3", "Amt"), kBindingUtf8EncodingError);
  EXPECT_EQ(add("P", "Amt\xC3"), kBindingUtf8EncodingError);
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_field(wb.handle, cache + 100U, "P", "Amt", &out), kInvalidArgument);
  // Reserved, bracketed and taken names, compared under the collation.
  for (const char* name : {"All", "Blank", "a[b", "a]b", "Amt", "amt"}) {
    EXPECT_EQ(add(name, "Qty"), kInvalidArgument) << name;
  }
  // Formulas that are not pivot formulas.
  for (const char* formula : {"Amt+", "Nope*2", "A1*2", "SUM(A1:A3)", "Table1[Amt]", "[Book2]Sheet1!A1", "{1,2}",
                              "MAX(Amt)", "Region[North]"}) {
    EXPECT_EQ(add("P", formula), kInvalidArgument) << formula;
  }
  // 255 UTF-16 units of name fit, a surrogate pair counting two; 256 do not.
  EXPECT_EQ(add((Repeat(kEmoji, 127) + "a").c_str(), "Amt"), 0) << fm_get_last_error_message();
  EXPECT_EQ(add(Repeat(kEmoji, 128).c_str(), "Amt"), kInvalidArgument);
  EXPECT_EQ(add(Repeat("b", 256).c_str(), "Amt"), kInvalidArgument);
  // 8192 units of formula fit, however many names they hold; 8193 do not.
  EXPECT_EQ(add("Long", ("Amt" + Repeat("+Amt", 2046) + "+1+11").c_str()), 0) << fm_get_last_error_message();
  EXPECT_EQ(add("Longer", ("Amt" + Repeat("+Amt", 2046) + "+1+111").c_str()), kInvalidArgument);
  EXPECT_NE(std::string(fm_get_last_error_message()).find("too long"), std::string::npos);

  // Index edges of the other entry points.
  for (const std::size_t index :
       {std::size_t{2}, std::size_t{99}, std::size_t{0x80000000U}, static_cast<std::size_t>(0x100000000ULL)}) {
    EXPECT_EQ(fm_workbook_set_pivot_cache_calculated_field_formula(wb.handle, cache, index, "Amt"), kInvalidArgument)
        << index;
    EXPECT_EQ(fm_workbook_remove_pivot_cache_calculated_field(wb.handle, cache, index), kInvalidArgument) << index;
  }
  const char* formula = nullptr;
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, 99U, &formula), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_formula(wb.handle, cache, 0U, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_formula(nullptr, cache, 0U, &formula), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_set_pivot_cache_calculated_field_formula(wb.handle, cache, out, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_set_pivot_cache_calculated_field_formula(nullptr, cache, out, "Amt"), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_remove_pivot_cache_calculated_field(nullptr, cache, out), kBindingNullPointer);
}

// ---------------------------------------------------------------------------
// Calculated items
// ---------------------------------------------------------------------------

TEST(FormulonCApiPivotCalc, CalculatedItemGridAndSavedStringsMatchExcel) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  // Typed as `=North+South`; Excel stores the item names qualified.
  std::size_t item_index = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, "X", "=North+South", &item_index), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(item_index, 0U);
  EXPECT_EQ(Pending(wb.handle), 1);
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  EXPECT_EQ(Pending(wb.handle), 0);

  // CI1: the item sums its items, and the grand total includes it.
  ExpectNumber(wb.handle, "X", SumOf("Amt"), 560.0);
  ExpectNumber(wb.handle, "X", SumOf("Qty"), 14.0);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Amt"), 1190.0);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Qty"), 34.0);

  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_calculated_item_count(wb.handle, cache, &count), 0);
  EXPECT_EQ(count, 1U);
  fm_pivot_calculated_item_t item{};
  ASSERT_EQ(fm_workbook_get_pivot_cache_calculated_item(wb.handle, cache, 0U, &item), 0);
  EXPECT_EQ(item.field_index, kRegion);
  EXPECT_STREQ(item.name, "X");
  EXPECT_STREQ(item.formula, "=Region[North]+Region[South]");
  EXPECT_EQ(item.scoped, 0);

  const std::string cache_xml = SavedPart(wb.handle, "xl/pivotCache/pivotCacheDefinition1.xml");
  EXPECT_NE(cache_xml.find("<s v=\"X\" f=\"1\"/>"), std::string::npos) << cache_xml;
  EXPECT_NE(cache_xml.find("<calculatedItems count=\"1\"><calculatedItem formula=\"Region[North]+Region[South]\">"
                           "<pivotArea cacheIndex=\"1\" outline=\"0\" fieldPosition=\"0\"><references count=\"1\">"
                           "<reference field=\"0\" count=\"1\"><x v=\"3\"/></reference></references></pivotArea>"
                           "</calculatedItem></calculatedItems>"),
            std::string::npos)
      << cache_xml;
  const std::string pivot_xml = SavedPart(wb.handle, "xl/pivotTables/pivotTable1.xml");
  EXPECT_NE(pivot_xml.find("<item f=\"1\" x=\"3\"/>"), std::string::npos) << pivot_xml;
}

TEST(FormulonCApiPivotCalc, ExcelSavedCalculatedItemLoadsWithTheSameGrid) {
  WorkbookGuard wb;
  LoadFixture("pivot_calc/ci1_item_sum_two_items.xlsx", &wb);
  ExpectNumber(wb.handle, "X", "Summe von Amt", 560.0);
  ExpectNumber(wb.handle, kGrandTotal, "Summe von Amt", 1190.0);
  ExpectNumber(wb.handle, kGrandTotal, "Summe von Qty", 34.0);
  std::uint32_t cache = 0;
  ASSERT_EQ(fm_sheet_get_pivot_cache_id(wb.handle, 0U, kPivot, &cache), 0);
  fm_pivot_calculated_item_t item{};
  ASSERT_EQ(fm_workbook_get_pivot_cache_calculated_item(wb.handle, cache, 0U, &item), 0);
  EXPECT_STREQ(item.name, "X");
  EXPECT_STREQ(item.formula, "=Region[North]+Region[South]");
}

TEST(FormulonCApiPivotCalc, AveragingFunctionsExcludeCalculatedItems) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, "X", "North+South", ignored_index()),
            0)
      << fm_get_last_error_message();
  // CI2: Average is refused; Max is not.
  for (const fm_pivot_aggregation_t aggregation :
       {FM_PIVOT_AGG_AVERAGE, FM_PIVOT_AGG_STDDEV, FM_PIVOT_AGG_STDDEVP, FM_PIVOT_AGG_VAR, FM_PIVOT_AGG_VARP}) {
    fm_pivot_data_field_spec_t spec = DataSpec(kAmt, aggregation);
    EXPECT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0U, kPivot, 0U, &spec), kPreconditionFailed) << aggregation;
    std::size_t index = 99;
    EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, kPivot, &spec, &index), kPreconditionFailed) << aggregation;
    EXPECT_EQ(index, 99U);
  }
  std::size_t data_fields = 0;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0U, kPivot, &data_fields), 0);
  EXPECT_EQ(data_fields, 2U);

  // CI2: Count counts the records of the items the formula names.
  fm_pivot_data_field_spec_t count = DataSpec(kAmt, FM_PIVOT_AGG_COUNT);
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0U, kPivot, 0U, &count), 0) << fm_get_last_error_message();
  const std::string count_name = "\xE5\x80\x8B\xE6\x95\xB0 / Amt";  // 個数 / Amt
  ExpectNumber(wb.handle, "East", count_name, 1.0);
  ExpectNumber(wb.handle, "North", count_name, 2.0);
  ExpectNumber(wb.handle, "X", count_name, 4.0);
  ExpectNumber(wb.handle, kGrandTotal, count_name, 9.0);
  fm_pivot_data_field_spec_t max = DataSpec(kAmt, FM_PIVOT_AGG_MAX);
  EXPECT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0U, kPivot, 0U, &max), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiPivotCalc, AveragingDataFieldBlocksAddingACalculatedItem) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  fm_pivot_data_field_spec_t average = DataSpec(kAmt, FM_PIVOT_AGG_AVERAGE);
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0U, kPivot, 0U, &average), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, "X", "North+South", ignored_index()),
            kPreconditionFailed);
  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_get_pivot_cache_calculated_item_count(wb.handle, cache, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivotCalc, PositionalReferences) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  // CI3: the previous item's value, then out-of-range positions.
  ASSERT_EQ(
      fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, "Prev", "=Region[-1]", ignored_index()), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  ExpectNumber(wb.handle, "Prev", SumOf("Amt"), 350.0);
  ExpectNumber(wb.handle, "Prev", SumOf("Qty"), 7.0);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Amt"), 980.0);
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Qty"), 27.0);

  ASSERT_EQ(
      fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, "Next", "=Region[+1]", ignored_index()), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(
      fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, "Far", "=Region[-9]", ignored_index()), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  ExpectNumber(wb.handle, "Prev", SumOf("Amt"), 350.0);
  ExpectError(wb.handle, "Next", SumOf("Amt"), formulon::ErrorCode::Ref);
  ExpectError(wb.handle, "Far", SumOf("Qty"), formulon::ErrorCode::Ref);
  ExpectError(wb.handle, kGrandTotal, SumOf("Amt"), formulon::ErrorCode::Ref);
}

TEST(FormulonCApiPivotCalc, LaterItemWinsAtAnIntersectionAndMoveSwapsIt) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache, /*product_on_cols=*/true);
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, "RX", "Region[North]+Region[South]",
                                                        ignored_index()),
            0)
      << fm_get_last_error_message();
  ASSERT_EQ(
      fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kProduct, "PY", "Product[A]+100", ignored_index()),
      0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  // CI4b: RX first, PY last: PY's formula over RX's A.
  ExpectNumber(wb.handle, "RX", "PY", 520.0);
  ExpectNumber(wb.handle, "RX", "A", 420.0);
  ExpectNumber(wb.handle, "North", "PY", 220.0);

  // CI4c: PY first, RX last: RX's formula over the PY cells.
  ASSERT_EQ(fm_workbook_move_pivot_cache_calculated_item(wb.handle, cache, 1U, 0U), 0) << fm_get_last_error_message();
  fm_pivot_calculated_item_t first{};
  ASSERT_EQ(fm_workbook_get_pivot_cache_calculated_item(wb.handle, cache, 0U, &first), 0);
  EXPECT_STREQ(first.name, "PY");
  EXPECT_EQ(first.field_index, kProduct);
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  ExpectNumber(wb.handle, "RX", "PY", 620.0);
  ExpectNumber(wb.handle, "RX", "A", 420.0);
}

TEST(FormulonCApiPivotCalc, CalculatedItemFormulaAndRemoval) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, "X", "North+South", ignored_index()),
            0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, "Y", "X*2", ignored_index()), 0)
      << fm_get_last_error_message();
  // X cannot go while Y names it, nor name itself through Y.
  EXPECT_EQ(fm_workbook_remove_pivot_cache_calculated_item(wb.handle, cache, 0U), kPreconditionFailed);
  EXPECT_EQ(fm_workbook_set_pivot_cache_calculated_item_formula(wb.handle, cache, 0U, "Y/2"), kInvalidArgument);

  ASSERT_EQ(fm_workbook_set_pivot_cache_calculated_item_formula(wb.handle, cache, 1U, "=X*3"), 0)
      << fm_get_last_error_message();
  fm_pivot_calculated_item_t item{};
  ASSERT_EQ(fm_workbook_get_pivot_cache_calculated_item(wb.handle, cache, 1U, &item), 0);
  // Stored as Excel stores it: a bare item name is qualified with its field.
  EXPECT_STREQ(item.formula, "=Region[X]*3");
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  ExpectNumber(wb.handle, "Y", SumOf("Amt"), 1680.0);

  ASSERT_EQ(fm_workbook_remove_pivot_cache_calculated_item(wb.handle, cache, 1U), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_remove_pivot_cache_calculated_item(wb.handle, cache, 0U), 0) << fm_get_last_error_message();
  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_get_pivot_cache_calculated_item_count(wb.handle, cache, &count), 0);
  EXPECT_EQ(count, 0U);
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache, kRegion, &count), 0);
  EXPECT_EQ(count, 3U);
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, kPivot), 0) << fm_get_last_error_message();
  ExpectNumber(wb.handle, kGrandTotal, SumOf("Amt"), 630.0);
}

TEST(FormulonCApiPivotCalc, CalculatedItemPlacementPreconditions) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  const std::size_t price = AddPrice(wb.handle, cache);
  // A calculated field, and a field on no axis.
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, price, "X", "1", ignored_index()),
            kPreconditionFailed);
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kProduct, "X", "Product[A]", ignored_index()),
            kPreconditionFailed);

  // A grouped field.
  WorkbookGuard dated;
  ASSERT_EQ(fm_workbook_create_default(&dated.handle), 0);
  std::uint32_t dated_cache = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(dated.handle, 0, &dated_cache), 0);
  std::size_t index = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(dated.handle, dated_cache, "Date", &index), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(dated.handle, dated_cache, "Amt", &index), 0);
  for (const double serial : {45306.0, 45340.0, 45400.0}) {
    std::size_t record = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(dated.handle, dated_cache, &record), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(dated.handle, dated_cache, record, 0U, serial), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(dated.handle, dated_cache, record, 1U, 1.0), 0);
  }
  std::size_t pivot = 99;
  ASSERT_EQ(fm_sheet_add_pivot(dated.handle, 0U, "PT", dated_cache, 0U, 3U, &pivot), 0);
  AddField(dated.handle, "Date", FM_PIVOT_AXIS_ROW);
  const std::uint32_t rows[] = {0U};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(dated.handle, 0U, kPivot, rows, 1U), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_date_group(dated.handle, 0U, kPivot, 0U, FM_PIVOT_PERIOD_MONTHS,
                                                FM_PIVOT_CALENDAR_GREGORIAN, 1U, std::nan(""), std::nan("")),
            0)
      << fm_get_last_error_message();
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_item(dated.handle, dated_cache, 0U, "X", "1", ignored_index()),
            kPreconditionFailed);
}

TEST(FormulonCApiPivotCalc, CalculatedItemArgumentBoundaries) {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  BuildSales(wb, &cache);
  const auto add = [&](const char* name, const char* formula) {
    return fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, kRegion, name, formula, ignored_index());
  };
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_item(nullptr, cache, kRegion, "X", "North", ignored_index()),
            kBindingNullPointer);
  EXPECT_EQ(add(nullptr, "North"), kBindingNullPointer);
  EXPECT_EQ(add("X", nullptr), kBindingNullPointer);
  EXPECT_EQ(add("", "North"), kInvalidArgument);
  EXPECT_EQ(add("X", ""), kInvalidArgument);
  EXPECT_EQ(add("\xC3", "North"), kBindingUtf8EncodingError);
  EXPECT_EQ(add("X", "\xC3"), kBindingUtf8EncodingError);
  EXPECT_EQ(
      fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache + 100U, kRegion, "X", "North", ignored_index()),
      kInvalidArgument);
  EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_item(wb.handle, cache, 99U, "X", "North", ignored_index()),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  // Reserved, bracketed and existing item names.
  for (const char* name : {"All", "Blank", "X[1]", "north", "South"}) {
    EXPECT_EQ(add(name, "North"), kInvalidArgument) << name;
  }
  // Another field's items, a bare field name, cells and defined names.
  for (const char* formula : {"Product[A]", "Amt", "A1", "North+", "Unknown"}) {
    EXPECT_EQ(add("X", formula), kInvalidArgument) << formula;
  }
  EXPECT_EQ(add(Repeat("x", 256).c_str(), "North"), kInvalidArgument);
  EXPECT_EQ(add("X", ("North" + Repeat(" ", 8188)).c_str()), kInvalidArgument);
  EXPECT_EQ(add((Repeat(kEmoji, 127) + "a").c_str(), "North"), 0) << fm_get_last_error_message();

  // Index edges of the other entry points.
  fm_pivot_calculated_item_t item{};
  for (const std::size_t index : {std::size_t{1}, std::size_t{0x80000000U}, static_cast<std::size_t>(0x100000000ULL)}) {
    EXPECT_EQ(fm_workbook_get_pivot_cache_calculated_item(wb.handle, cache, index, &item), FM_STATUS_INDEX_OUT_OF_RANGE)
        << index;
    EXPECT_EQ(fm_workbook_set_pivot_cache_calculated_item_formula(wb.handle, cache, index, "North"),
              FM_STATUS_INDEX_OUT_OF_RANGE);
    EXPECT_EQ(fm_workbook_move_pivot_cache_calculated_item(wb.handle, cache, index, 0U), FM_STATUS_INDEX_OUT_OF_RANGE);
    EXPECT_EQ(fm_workbook_move_pivot_cache_calculated_item(wb.handle, cache, 0U, index), FM_STATUS_INDEX_OUT_OF_RANGE);
    EXPECT_EQ(fm_workbook_remove_pivot_cache_calculated_item(wb.handle, cache, index), FM_STATUS_INDEX_OUT_OF_RANGE);
  }
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_get_pivot_cache_calculated_item_count(nullptr, cache, &count), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_pivot_cache_calculated_item_count(wb.handle, cache, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_pivot_cache_calculated_item_count(wb.handle, cache + 100U, &count), kInvalidArgument);
  EXPECT_EQ(fm_workbook_get_pivot_cache_calculated_item(wb.handle, cache, 0U, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_set_pivot_cache_calculated_item_formula(wb.handle, cache, 0U, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_set_pivot_cache_calculated_item_formula(wb.handle, cache, 0U, ""), kInvalidArgument);
  EXPECT_EQ(fm_workbook_move_pivot_cache_calculated_item(nullptr, cache, 0U, 0U), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_remove_pivot_cache_calculated_item(nullptr, cache, 0U), kBindingNullPointer);
}
