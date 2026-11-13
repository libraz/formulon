//
// Workbook::copy_sheet against workbooks Excel 365 (Mac, ja-JP) copied with
// Worksheet.Copy. Expected texts are the ones Excel reported for its copies
// of the fixtures under tests/fixtures/copy_sheet/.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <tuple>
#include <utility>
#include <vector>

#include "drawing/drawing_xml.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/auto_filter_xml.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "passthrough_part.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "table.h"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr std::size_t kSrc = 1;    // main.xlsx: A, Src, Other, C, Shx, ...
constexpr std::size_t kOther = 2;  // main.xlsx
constexpr std::size_t kHidden = 12;
constexpr std::size_t kVeryHidden = 13;

Bytes ReadFile(const std::string& path) {
  Bytes out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  out.resize(static_cast<std::size_t>(std::ftell(file)));
  std::fseek(file, 0, SEEK_SET);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(file);
  return out;
}

Workbook ReadBytes(const Bytes& bytes) {
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

Workbook Load(const std::string& dir, const std::string& name) {
  const Bytes bytes = ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/" + dir + "/" + name);
  if (name.size() > 5 && name.substr(name.size() - 5) == ".xlsb") {
    auto read = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
    EXPECT_TRUE(static_cast<bool>(read));
    return read ? std::move(read.value().workbook) : Workbook::create_empty();
  }
  return ReadBytes(bytes);
}

Workbook LoadCopyFixture(const std::string& name) {
  return Load("copy_sheet", name);
}

Workbook Reload(const Workbook& wb) {
  auto saved = wb.save();
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? ReadBytes(saved.value()) : Workbook::create_empty();
}

void Recalc(Workbook& wb) {
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
}

// Zero-based (row, col) of an A1 address with a one-letter column.
std::pair<std::uint32_t, std::uint32_t> At(std::string_view a1) {
  const auto col = static_cast<std::uint32_t>(a1[0] - 'A');
  return {static_cast<std::uint32_t>(std::stoul(std::string(a1.substr(1)))) - 1U, col};
}

std::string FormulaAt(const Workbook& wb, std::size_t sheet, std::string_view a1) {
  const auto [row, col] = At(a1);
  const Cell* cell = wb.sheet(sheet).cell_at(row, col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

Value ValueAt(const Workbook& wb, std::size_t sheet, std::string_view a1) {
  const auto [row, col] = At(a1);
  const Cell* cell = wb.sheet(sheet).cell_at(row, col);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

double NumberAt(const Workbook& wb, std::size_t sheet, std::string_view a1) {
  const Value value = ValueAt(wb, sheet, a1);
  EXPECT_TRUE(value.is_number()) << a1;
  return value.is_number() ? value.as_number() : -1.0;
}

std::vector<std::string> SheetNames(const Workbook& wb) {
  std::vector<std::string> names;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    names.push_back(wb.sheet(i).name());
  }
  return names;
}

std::string Strip(std::string formula) {
  return !formula.empty() && formula.front() == '=' ? formula.substr(1) : formula;
}

// (scope, name, formula) with the scope a sheet name or "" for workbook scope.
using NameRow = std::tuple<std::string, std::string, std::string>;

std::set<NameRow> Names(const Workbook& wb) {
  std::set<NameRow> rows;
  for (const DefinedName& entry : wb.defined_names()) {
    const std::string scope =
        entry.local_sheet_id >= 0 ? wb.sheet(static_cast<std::size_t>(entry.local_sheet_id)).name() : std::string();
    rows.emplace(scope, entry.name, Strip(entry.formula));
  }
  return rows;
}

std::vector<std::string> TableNamesOn(const Workbook& wb, std::size_t sheet) {
  std::vector<std::string> names;
  for (const TableMetadata& table : wb.tables()) {
    if (table.sheet_index == sheet) {
      names.push_back(table.display_name);
    }
  }
  return names;
}

std::string PartText(const Workbook& wb, const std::string& path) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return std::string(part.bytes.begin(), part.bytes.end());
    }
  }
  return {};
}

// Every `<c:f>` / `<f>` body of the charts on `sheet`, in part order.
std::vector<std::string> ChartRefs(const Workbook& wb, std::size_t sheet) {
  std::vector<std::string> refs;
  auto charts = list_chart_parts(wb, sheet);
  EXPECT_TRUE(static_cast<bool>(charts));
  if (!charts) {
    return refs;
  }
  for (const ChartPartRef& chart : charts.value()) {
    const std::string text = PartText(wb, chart.part_path);
    for (const std::string open : {"<f>", "<c:f>"}) {
      for (std::size_t at = text.find(open); at != std::string::npos; at = text.find(open, at)) {
        at += open.size();
        refs.push_back(text.substr(at, text.find("</", at) - at));
      }
    }
  }
  return refs;
}

// Adds `codeName` to the sheet's raw `<sheetPr>`, creating the element when absent.
void SetCodeName(Sheet& sheet, const std::string& code) {
  std::string& pr = sheet.mutable_print_settings().sheet_pr_xml;
  if (pr.find("<sheetPr") == std::string::npos) {
    pr = "<sheetPr codeName=\"" + code + "\"/>";
    return;
  }
  pr.insert(pr.find("<sheetPr") + 8U, " codeName=\"" + code + "\"");
}

bool IsError(const Value& value, ErrorCode code) {
  return value.is_error() && value.as_error() == code;
}

// ---------------------------------------------------------------------------

TEST(WorkbookCopySheet, MeasuredFormulas) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  auto copy = wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(copy)) << copy.error().message;
  ASSERT_EQ(copy.value(), 2U);
  EXPECT_EQ(wb.sheet(2).name(), "Src (2)");
  Recalc(wb);

  // The reader spells a legacy formula's implicit intersection `@`, on the
  // source as on the copy; Excel's `Formula` omits it.
  const std::vector<std::pair<std::string, std::string>> formulas = {
      {"B1", "='Src (2)'!A1"},
      {"B2", "=A1"},
      {"B3", "=Other!A1"},
      {"B4", "=SUM(A:Other!A1)"},
      {"B5", "=LN"},
      {"B6", "=GN"},
      {"B7", "=GO"},
      {"B8", "=@INDIRECT(\"Src!A1\")"},
      {"B9", "=SUM(Tbl_3[V])"},
      {"B10", "=SUM('Src (2)'!A1:A3)"},
      {"B12", "='Src (2)'!$A$1+'Src (2)'!A2"},
      {"B13", "=@CELL(\"address\",'Src (2)'!A1)"},
  };
  for (const auto& [cell, text] : formulas) {
    EXPECT_EQ(FormulaAt(wb, 2, cell), text) << cell;
  }

  const std::vector<std::pair<std::string, double>> values = {
      {"B1", 1.0},  {"B2", 1.0}, {"B3", 10.0}, {"B4", 112.0}, {"B5", 2.0},  {"B6", 1.0},
      {"B7", 10.0}, {"B8", 1.0}, {"B9", 12.0}, {"B10", 6.0},  {"B12", 3.0},
  };
  for (const auto& [cell, number] : values) {
    EXPECT_EQ(NumberAt(wb, 2, cell), number) << cell;
  }
  // A dead 3-D reference keeps its address. Excel shows 1012 until a full
  // recalculation, then #REF!.
  EXPECT_EQ(FormulaAt(wb, 2, "B11"), "=SUM(#REF!A1)");
  EXPECT_TRUE(IsError(ValueAt(wb, 2, "B11"), ErrorCode::Ref));
  // The copy's reference names its own sheet, so Excel answers `$A$1`.
  ASSERT_TRUE(ValueAt(wb, 2, "B13").is_text());
  EXPECT_EQ(ValueAt(wb, 2, "B13").as_text(), "$A$1");

  // Conditional formatting and validation follow the copy (Excel stores
  // these self-sheet references unqualified: `$A$1>0`, `$A$1:$A$3`); both
  // qualifiers are dropped on load.
  const Sheet& sheet = wb.sheet(2);
  ASSERT_EQ(sheet.conditional_formats().size(), 1U);
  ASSERT_EQ(sheet.conditional_formats()[0].rules.size(), 1U);
  ASSERT_TRUE(sheet.conditional_formats()[0].rules[0].formula1.has_value());
  EXPECT_EQ(Strip(*sheet.conditional_formats()[0].rules[0].formula1), "$A$1>0");
  ASSERT_EQ(sheet.validations().size(), 1U);
  EXPECT_EQ(Strip(sheet.validations()[0].formula1), "$A$1:$A$3");
}

TEST(WorkbookCopySheet, ThreeDEndpoints) {
  Workbook wb = LoadCopyFixture("t3.xlsx");
  auto copy = wb.copy_sheet(1, 2, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(copy));
  ASSERT_EQ(wb.sheet(2).name(), "S (2)");
  Recalc(wb);
  // A span with one endpoint on the source loses that sheet and keeps its address.
  EXPECT_EQ(FormulaAt(wb, 2, "B1"), "=SUM(#REF!A1)");
  EXPECT_EQ(FormulaAt(wb, 2, "B2"), "=SUM(#REF!A1)");
  EXPECT_EQ(FormulaAt(wb, 2, "B3"), "=SUM(A:Z!A1)");
  EXPECT_EQ(FormulaAt(wb, 2, "B4"), "=SUM('S (2)'!A1)");
  EXPECT_TRUE(IsError(ValueAt(wb, 2, "B1"), ErrorCode::Ref));
  EXPECT_TRUE(IsError(ValueAt(wb, 2, "B2"), ErrorCode::Ref));
  EXPECT_EQ(NumberAt(wb, 2, "B3"), 109.0);
  EXPECT_EQ(NumberAt(wb, 2, "B4"), 1.0);
}

TEST(WorkbookCopySheet, DefinedNames) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt)));
  const std::set<NameRow> expected = {
      {"Src (2)", "GN", "'Src (2)'!$A$1"},
      {"", "GN", "Src!$A$1"},
      {"", "GO", "Other!$A$1"},
      {"Src (2)", "GUNUSED", "'Src (2)'!$A$3"},
      {"", "GUNUSED", "Src!$A$3"},
      {"Src", "LN", "Src!$A$2"},
      {"Src (2)", "LN", "'Src (2)'!$A$2"},
      {"Src", "_xlnm.Print_Area", "'Src'!$A$1:$B$5"},
      {"Src (2)", "_xlnm.Print_Area", "'Src (2)'!$A$1:$B$5"},
      {"Src", "_xlnm.Print_Titles", "'Src'!$1:$1"},
      {"Src (2)", "_xlnm.Print_Titles", "'Src (2)'!$1:$1"},
  };
  EXPECT_EQ(Names(wb), expected);
}

TEST(WorkbookCopySheet, ConflictLocalWins) {
  Workbook wb = LoadCopyFixture("conflict.xlsx");
  auto copy = wb.copy_sheet(0, 1, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(copy));
  EXPECT_EQ(SheetNames(wb), (std::vector<std::string>{"S", "S (2)", "O"}));
  const std::set<NameRow> expected = {
      {"S", "DUP", "S!$A$3"},
      {"S (2)", "DUP", "'S (2)'!$A$3"},
      {"", "DUP", "O!$A$1"},
  };
  EXPECT_EQ(Names(wb), expected);
  Recalc(wb);
  EXPECT_EQ(FormulaAt(wb, 1, "B1"), "=DUP");
  EXPECT_EQ(NumberAt(wb, 1, "B1"), 3.0);
}

TEST(WorkbookCopySheet, TablesRenamedAndRebound) {
  {
    Workbook wb = LoadCopyFixture("main.xlsx");
    TableMetadata& source = wb.mutable_tables().front();
    source.root_extra_attrs +=
        " xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\""
        " xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\""
        " xr:uid=\"{00000000-000C-0000-FFFF-FFFF00000000}\"";
    source.columns.front().extra_attrs += " xr3:uid=\"{00000000-0010-0000-0000-000004000000}\"";
    const std::string filter_xml = io::auto_filter_xml(wb.tables().front().auto_filter_xml.get());

    ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt)));
    ASSERT_EQ(wb.tables().size(), 2U);
    const TableMetadata& copy = wb.tables().back();
    EXPECT_EQ(copy.name, "Tbl_3");
    EXPECT_EQ(copy.display_name, "Tbl_3");
    EXPECT_EQ(copy.sheet_index, 2U);
    EXPECT_EQ(copy.ref, "D1:D3");
    EXPECT_EQ(copy.id, 2U);
    ASSERT_EQ(copy.columns.size(), 1U);
    EXPECT_EQ(copy.columns.front().name, "V");
    EXPECT_EQ(io::auto_filter_xml(copy.auto_filter_xml.get()), filter_xml);
    EXPECT_EQ(copy.root_extra_attrs.find("{00000000-000C-0000-FFFF-FFFF00000000}"), std::string::npos);
    EXPECT_NE(copy.root_extra_attrs.find(" xr:uid=\"{"), std::string::npos);
    EXPECT_EQ(copy.columns.front().extra_attrs.find("{00000000-0010-0000-0000-000004000000}"), std::string::npos);
    EXPECT_NE(copy.columns.front().extra_attrs.find(" xr3:uid=\"{"), std::string::npos);
    EXPECT_EQ(wb.tables().front().name, "Tbl");
    EXPECT_EQ(wb.tables().front().sheet_index, kSrc);
    EXPECT_EQ(FormulaAt(wb, kSrc, "B9"), "=SUM(Tbl[V])");
    EXPECT_EQ(FormulaAt(wb, 2, "B9"), "=SUM(Tbl_3[V])");
  }
  {
    // Copy to the end, to the front, then a copy of a copy after C.
    Workbook wb = LoadCopyFixture("main.xlsx");
    auto end = wb.copy_sheet(kSrc, wb.sheet_count(), std::nullopt);
    ASSERT_TRUE(static_cast<bool>(end));
    EXPECT_EQ(wb.sheet(end.value()).name(), "Src (2)");
    EXPECT_EQ(TableNamesOn(wb, end.value()), std::vector<std::string>{"Tbl_3"});
    auto front = wb.copy_sheet(kSrc, 0, std::nullopt);
    ASSERT_TRUE(static_cast<bool>(front));
    EXPECT_EQ(front.value(), 0U);
    EXPECT_EQ(wb.sheet(0).name(), "Src (3)");
    EXPECT_EQ(TableNamesOn(wb, 0), std::vector<std::string>{"Tbl_4"});
    const std::size_t second = wb.sheet_index_by_name("Src (2)");
    const std::size_t c_sheet = wb.sheet_index_by_name("C");
    auto nested = wb.copy_sheet(second, c_sheet + 1U, std::nullopt);
    ASSERT_TRUE(static_cast<bool>(nested));
    EXPECT_EQ(nested.value(), 5U);
    EXPECT_EQ(wb.sheet(5).name(), "Src (4)");
    EXPECT_EQ(TableNamesOn(wb, 5), std::vector<std::string>{"Tbl_35"});
    EXPECT_EQ(FormulaAt(wb, 5, "B9"), "=SUM(Tbl_35[V])");
    for (const TableMetadata& table : wb.tables()) {
      EXPECT_EQ(table.ref, "D1:D3");
      EXPECT_EQ(wb.sheet(table.sheet_index).name().rfind("Src", 0), 0U);
    }
  }
  {
    Workbook wb = LoadCopyFixture("t3.xlsx");
    ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(1, 2, std::nullopt)));
    EXPECT_EQ(TableNamesOn(wb, 2), (std::vector<std::string>{"My_Tbl6", "Txn_7", "Q_128", "Plain9"}));
    auto again = wb.copy_sheet(1, wb.sheet_index_by_name("Z") + 1U, std::nullopt);
    ASSERT_TRUE(static_cast<bool>(again));
    EXPECT_EQ(wb.sheet(again.value()).name(), "S (3)");
    EXPECT_EQ(TableNamesOn(wb, again.value()), (std::vector<std::string>{"My_Tbl10", "Txn_11", "Q_1212", "Plain13"}));
    std::set<std::uint32_t> ids;
    for (const TableMetadata& table : wb.tables()) {
      EXPECT_TRUE(ids.insert(table.id).second) << table.name;
    }
  }
}

TEST(WorkbookCopySheet, ChartSeriesRebound) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  const std::vector<std::string> before = ChartRefs(wb, kSrc);
  ASSERT_EQ(before, std::vector<std::string>{"'Src'!$A$1:$A$3"});
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt)));
  EXPECT_EQ(ChartRefs(wb, 2), std::vector<std::string>{"'Src (2)'!$A$1:$A$3"});
  EXPECT_EQ(ChartRefs(wb, kSrc), before);
  EXPECT_NE(wb.sheet(2).drawing_rel_target(), wb.sheet(kSrc).drawing_rel_target());
}

TEST(WorkbookCopySheet, PivotSharesCache) {
  Workbook wb = LoadCopyFixture("pivot_refresh_M2.xlsx");
  ASSERT_EQ(wb.sheet(0).name(), "Sheet2");
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  const auto& source_pivots = wb.sheet(0).pivot_tables();
  ASSERT_FALSE(source_pivots.empty());
  const std::string pivot_name = source_pivots.front()->name();
  const std::uint32_t cache_id = source_pivots.front()->pivot_cache_id();
  const std::string cache_sheet = wb.pivot_caches().front()->worksheet_source().sheet;

  auto copy = wb.copy_sheet(0, 1, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(copy)) << copy.error().message;
  ASSERT_EQ(copy.value(), 1U);
  EXPECT_EQ(wb.sheet(1).name(), "Sheet2 (2)");
  ASSERT_EQ(wb.sheet(1).pivot_tables().size(), wb.sheet(0).pivot_tables().size());
  EXPECT_EQ(wb.sheet(1).pivot_tables().front()->name(), pivot_name);
  EXPECT_EQ(wb.sheet(1).pivot_tables().front()->pivot_cache_id(), cache_id);
  EXPECT_NE(wb.sheet(1).pivot_tables().front().get(), wb.sheet(0).pivot_tables().front().get());
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  EXPECT_EQ(wb.pivot_caches().front()->worksheet_source().sheet, cache_sheet);

  const Workbook reloaded = Reload(wb);
  EXPECT_EQ(reloaded.pivot_caches().size(), 1U);
  ASSERT_EQ(reloaded.sheet_count(), wb.sheet_count());
  ASSERT_EQ(reloaded.sheet(1).pivot_tables().size(), wb.sheet(1).pivot_tables().size());
  EXPECT_EQ(reloaded.sheet(1).pivot_tables().front()->name(), pivot_name);
}

TEST(WorkbookCopySheet, HiddenStaysHidden) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  const std::uint32_t active = wb.active_sheet();
  auto copy = wb.copy_sheet(kHidden, wb.sheet_index_by_name("C") + 1U, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(copy));
  EXPECT_EQ(copy.value(), 4U);
  EXPECT_EQ(wb.sheet(4).name(), "H (2)");
  EXPECT_EQ(wb.sheet(4).view().visibility(), SheetVisibility::kHidden);
  EXPECT_EQ(wb.active_sheet(), active);
}

TEST(WorkbookCopySheet, VeryHiddenRefused) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  ASSERT_EQ(wb.sheet(kVeryHidden).view().visibility(), SheetVisibility::kVeryHidden);
  const std::size_t count = wb.sheet_count();
  auto copy = wb.copy_sheet(kVeryHidden, 4, std::nullopt);
  ASSERT_FALSE(static_cast<bool>(copy));
  EXPECT_EQ(copy.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.sheet_count(), count);
}

// Copies `sheet` and expects a kNotImplemented refusal that leaves the saved
// package byte-identical.
// `what` is the refusal reason the copy reports (`Error::context`).
void ExpectRefused(Workbook& wb, std::size_t sheet, const char* what) {
  auto before = wb.save();
  ASSERT_TRUE(static_cast<bool>(before)) << what;
  const std::size_t count = wb.sheet_count();
  auto copy = wb.copy_sheet(sheet, sheet + 1U, std::nullopt);
  ASSERT_FALSE(static_cast<bool>(copy)) << what;
  EXPECT_EQ(copy.error().code, FormulonErrorCode::kNotImplemented) << what << ": " << copy.error().message;
  EXPECT_EQ(copy.error().context, what);
  EXPECT_EQ(wb.sheet_count(), count) << what;
  auto after = wb.save();
  ASSERT_TRUE(static_cast<bool>(after)) << what;
  EXPECT_TRUE(before.value() == after.value()) << what;
}

std::string DrawingWith(const std::string& anchor) {
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
         "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
         "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
         "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
         anchor + "</xdr:wsDr>";
}

Bytes ToBytes(const std::string& text) {
  return Bytes(text.begin(), text.end());
}

TEST(WorkbookCopySheet, UnsupportedRefusedWithoutChange) {
  {
    Workbook wb = LoadCopyFixture("main.xlsx");
    wb.sheet(kSrc).set_unknown_relationships(
        {UnknownRelationship{"rId9", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp",
                             "../ctrlProps/ctrlProp1.xml", false}});
    ExpectRefused(wb, kSrc, "unrecognised relationship");
  }
  {
    Workbook wb = LoadCopyFixture("main.xlsx");
    ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(PassthroughPart(
        "xl/chartsheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
        ToBytes("<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>")))));
    wb.sheet(kOther).set_opaque_ooxml_sheet(
        "xl/chartsheets/sheet1.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet");
    ExpectRefused(wb, kOther, "not a worksheet");
  }
  {
    Workbook wb = Workbook::create();
    const std::string shape =
        "<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>"
        "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"914400\" cy=\"457200\"/><xdr:sp macro=\"\" "
        "textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"Oval\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/>"
        "</xdr:sp><xdr:clientData/></xdr:oneCellAnchor>";
    ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(
        PassthroughPart("xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml",
                        ToBytes(DrawingWith(shape))))));
    wb.sheet(0).set_drawing_rel_target("xl/drawings/drawing1.xml");
    ExpectRefused(wb, 0, "shape, group, connector, text box or other frame");
  }
  {
    Workbook wb = LoadCopyFixture("main.xlsx");
    const std::string chart_path = ChartRefs(wb, kSrc).empty() ? "" : list_chart_parts(wb, kSrc).value()[0].part_path;
    ASSERT_FALSE(chart_path.empty());
    std::string chart = PartText(wb, chart_path);
    chart.insert(chart.find("<chart>"),
                 "<pivotSource><name>[x]Sheet1!PivotTable1</name><fmtId val=\"0\"/></pivotSource>");
    ASSERT_TRUE(static_cast<bool>(wb.replace_passthrough_part(chart_path, ToBytes(chart))));
    ExpectRefused(wb, kSrc, "pivot chart");
  }
  {
    // The slicer relationship alone already refuses; without it the slicer's
    // graphic frame is an unmodelled frame and refuses as one.
    Workbook wb = Load("excel", "table_slicer.xlsx");
    ExpectRefused(wb, 0, "unrecognised relationship");
    wb.sheet(0).set_unknown_relationships({});
    ExpectRefused(wb, 0, "shape, group, connector, text box or other frame");
  }
  {
    // An .xlsb drawing the reader could not lift stays an unmodelled
    // drawing relationship.
    Workbook wb = Load("excel", "chart_xlsb.xlsb");
    Sheet& sheet = wb.sheet(0);
    ASSERT_FALSE(sheet.drawing_rel_target().empty());
    std::vector<UnknownRelationship> rels = sheet.unknown_relationships();
    rels.push_back(UnknownRelationship{"rId9", std::string(io::kRelDrawing), sheet.drawing_rel_target(), false});
    sheet.set_unknown_relationships(std::move(rels));
    sheet.set_drawing_rel_target(std::string());
    ExpectRefused(wb, 0, "unrecognised relationship");
  }
}

bool HasBinaryIndex(const Sheet& sheet) {
  for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
    if (rel.target.rfind("xl/worksheets/binaryIndex", 0) == 0U) {
      return true;
    }
  }
  return false;
}

TEST(WorkbookCopySheet, XlsbBinaryIndexDropped) {
  Workbook wb = Load("excel", "xlsb_fidelity_base.xlsb");
  ASSERT_TRUE(HasBinaryIndex(wb.sheet(0)));
  auto copy = wb.copy_sheet(0, 1, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(copy)) << copy.error().message << " " << copy.error().context;
  EXPECT_TRUE(HasBinaryIndex(wb.sheet(0)));
  EXPECT_FALSE(HasBinaryIndex(wb.sheet(1)));
  EXPECT_TRUE(wb.sheet(1).unknown_relationships().empty());
  EXPECT_TRUE(static_cast<bool>(wb.save()));
}

TEST(WorkbookCopySheet, OtherSheetsUnchanged) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  std::vector<std::string> src_formulas;
  for (int row = 1; row <= 13; ++row) {
    src_formulas.push_back(FormulaAt(wb, kSrc, "B" + std::to_string(row)));
  }
  const std::string src_cf = *wb.sheet(kSrc).conditional_formats()[0].rules[0].formula1;
  const std::string src_dv = wb.sheet(kSrc).validations()[0].formula1;

  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt)));
  Recalc(wb);
  for (int row = 1; row <= 13; ++row) {
    EXPECT_EQ(FormulaAt(wb, kSrc, "B" + std::to_string(row)), src_formulas[static_cast<std::size_t>(row - 1)]) << row;
  }
  EXPECT_EQ(*wb.sheet(kSrc).conditional_formats()[0].rules[0].formula1, src_cf);
  EXPECT_EQ(wb.sheet(kSrc).validations()[0].formula1, src_dv);

  const std::size_t other = wb.sheet_index_by_name("Other");
  EXPECT_EQ(FormulaAt(wb, other, "B1"), "=SUM(A:C!A1)");
  EXPECT_EQ(FormulaAt(wb, other, "B2"), "=Src!A1");
  EXPECT_EQ(FormulaAt(wb, other, "B3"), "=SUM(Src:C!A1)");
  EXPECT_EQ(NumberAt(wb, other, "B1"), 1112.0);
  EXPECT_EQ(NumberAt(wb, other, "B2"), 1.0);
  EXPECT_EQ(NumberAt(wb, other, "B3"), 1012.0);
  EXPECT_EQ(NumberAt(wb, kSrc, "B4"), 112.0);
}

TEST(WorkbookCopySheet, ThreeDSpanElsewhereIncludesCopy) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt)));
  Recalc(wb);
  EXPECT_EQ(NumberAt(wb, wb.sheet_index_by_name("Other"), "B1"), 1112.0);
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(0, 1, std::nullopt)));
  EXPECT_EQ(wb.sheet(1).name(), "A (2)");
  Recalc(wb);
  const std::size_t other = wb.sheet_index_by_name("Other");
  EXPECT_EQ(NumberAt(wb, other, "B1"), 1212.0);
  EXPECT_EQ(NumberAt(wb, other, "B3"), 1012.0);
}

TEST(WorkbookCopySheet, TabSelectedCleared) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  wb.sheet(kSrc).mutable_view().tab_selected = true;
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt)));
  EXPECT_TRUE(wb.sheet(kSrc).view().tab_selected);
  EXPECT_FALSE(wb.sheet(2).view().tab_selected);
  std::size_t selected = 0;
  const Workbook reloaded = Reload(wb);
  for (std::size_t i = 0; i < reloaded.sheet_count(); ++i) {
    selected += reloaded.sheet(i).view().tab_selected ? 1U : 0U;
  }
  EXPECT_EQ(selected, 1U);
}

TEST(WorkbookCopySheet, HyperlinkLocationKept) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  ASSERT_EQ(wb.sheet(kSrc).hyperlinks().size(), 1U);
  const Hyperlink source = wb.sheet(kSrc).hyperlinks().front();
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt)));
  ASSERT_EQ(wb.sheet(2).hyperlinks().size(), 1U);
  const Hyperlink& copy = wb.sheet(2).hyperlinks().front();
  EXPECT_EQ(copy.location, source.location);
  EXPECT_EQ(copy.target, source.target);
  EXPECT_EQ(copy.row, source.row);
  EXPECT_EQ(copy.col, source.col);

  const Workbook reloaded = Reload(wb);
  ASSERT_EQ(reloaded.sheet(2).hyperlinks().size(), 1U);
  EXPECT_EQ(reloaded.sheet(2).hyperlinks().front().location, reloaded.sheet(kSrc).hyperlinks().front().location);
  EXPECT_EQ(reloaded.sheet(2).hyperlinks().front().target, reloaded.sheet(kSrc).hyperlinks().front().target);
}

TEST(WorkbookCopySheet, AppendAtEndRecalcs) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  Recalc(wb);
  const std::size_t end = wb.sheet_count();
  auto copy = wb.copy_sheet(kSrc, end, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(copy));
  ASSERT_EQ(copy.value(), end);
  Recalc(wb);
  EXPECT_EQ(NumberAt(wb, end, "B1"), 1.0);
  EXPECT_EQ(NumberAt(wb, end, "B10"), 6.0);
  EXPECT_EQ(NumberAt(wb, end, "B9"), 12.0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(end, 0, 0, Value::number(5.0))));
  Recalc(wb);
  EXPECT_EQ(NumberAt(wb, end, "B1"), 5.0);
  EXPECT_EQ(NumberAt(wb, end, "B10"), 10.0);
  EXPECT_EQ(NumberAt(wb, kSrc, "B1"), 1.0);
}

TEST(WorkbookCopySheet, CodeNameStripped) {
  Workbook wb = LoadCopyFixture("main.xlsx");
  SetCodeName(wb.sheet(kSrc), "Sheet2");
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt)));
  EXPECT_NE(wb.sheet(kSrc).print_settings().sheet_pr_xml.find("codeName=\"Sheet2\""), std::string::npos);
  EXPECT_EQ(wb.sheet(2).print_settings().sheet_pr_xml.find("codeName"), std::string::npos);
}

TEST(WorkbookCopySheet, OriginSheetIndexFollowsMove) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("P");
  wb.add_sheet("Q");
  wb.add_sheet("R");
  PassthroughPart part("xl/pivotTables/pivotTable1.xml", "application/xml", ToBytes("<pivotTableDefinition/>"));
  part.retained_origin = PassthroughPart::RetainedOrigin::kPivotTable;
  part.origin_sheet_index = 0;
  ASSERT_TRUE(static_cast<bool>(wb.add_passthrough_part(part)));
  const auto origin = [&wb] { return wb.passthrough_parts().front().origin_sheet_index; };

  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(0, 2)));
  EXPECT_EQ(wb.sheet(2).name(), "P");
  EXPECT_EQ(origin(), 2U);
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(1, 0)));
  EXPECT_EQ(origin(), 2U);
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(1, 0, std::nullopt)));
  EXPECT_EQ(wb.sheet(3).name(), "P");
  EXPECT_EQ(origin(), 3U);
}

TEST(WorkbookCopySheet, ArgumentsValidated) {
  Workbook wb = LoadCopyFixture("conflict.xlsx");
  const auto code = [&wb](std::size_t source, std::size_t before, std::optional<std::string> name) {
    auto result = wb.copy_sheet(source, before, std::move(name));
    return result ? FormulonErrorCode::kOk : result.error().code;
  };
  EXPECT_EQ(code(2, 0, std::nullopt), FormulonErrorCode::kSheetIndexOutOfRange);
  EXPECT_EQ(code(0, 3, std::nullopt), FormulonErrorCode::kSheetIndexOutOfRange);
  EXPECT_EQ(code(0, 0, std::string()), FormulonErrorCode::kInvalidSheetName);
  EXPECT_EQ(code(0, 0, std::string("o")), FormulonErrorCode::kAlreadyExists);
  EXPECT_EQ(code(0, 0, std::string(32, 'x')), FormulonErrorCode::kInvalidSheetName);
  EXPECT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(code(0, 0, std::string(31, 'x')), FormulonErrorCode::kOk);
  EXPECT_EQ(wb.sheet(0).name(), std::string(31, 'x'));
  EXPECT_EQ(FormulaAt(wb, 0, "B1"), "=DUP");
}

// ---------------------------------------------------------------------------
// Round trip: the projection named by the design survives an xlsx save/load.

std::string ValueText(const Value& value) {
  if (value.is_number()) {
    return "n:" + std::to_string(value.as_number());
  }
  if (value.is_text()) {
    return "t:" + std::string(value.as_text());
  }
  if (value.is_error()) {
    return "e:" + std::to_string(static_cast<int>(value.as_error()));
  }
  return value.is_blank() ? "" : "other";
}

std::vector<std::string> Projection(const Workbook& wb) {
  std::vector<std::string> out;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const Sheet& sheet = wb.sheet(i);
    std::string head = "sheet " + sheet.name() + " vis=" + std::to_string(static_cast<int>(sheet.view().visibility())) +
                       " sel=" + std::to_string(sheet.view().tab_selected ? 1 : 0);
    out.push_back(head);
    std::set<std::string> cells;
    for (const auto& [row, cells_in_row] : sheet.rows()) {
      for (const auto& [col, cell] : cells_in_row) {
        cells.insert(std::to_string(row) + "," + std::to_string(col) + " " + cell.formula_text + " " +
                     ValueText(cell.cached_value));
      }
    }
    out.insert(out.end(), cells.begin(), cells.end());
    for (const cf::ConditionalFormat& format : sheet.conditional_formats()) {
      for (const cf::CFRule& rule : format.rules) {
        out.push_back("cf " + Strip(rule.formula1.value_or("")) + " " + Strip(rule.formula2.value_or("")));
      }
    }
    for (const DataValidation& validation : sheet.validations()) {
      out.push_back("dv " + Strip(validation.formula1) + " " + Strip(validation.formula2));
    }
    for (const MergeRange& merge : sheet.merges()) {
      out.push_back("merge " + std::to_string(merge.first_row) + "," + std::to_string(merge.first_col) + ":" +
                    std::to_string(merge.last_row) + "," + std::to_string(merge.last_col));
    }
    for (const std::string& ref : ChartRefs(wb, i)) {
      out.push_back("chart " + ref);
    }
  }
  for (const NameRow& row : Names(wb)) {
    out.push_back("name " + std::get<0>(row) + "!" + std::get<1>(row) + " " + std::get<2>(row));
  }
  std::set<std::string> tables;
  for (const TableMetadata& table : wb.tables()) {
    std::string line = "table " + wb.sheet(table.sheet_index).name() + " " + table.display_name + " " + table.ref;
    for (const TableColumn& column : table.columns) {
      line += " " + column.name;
    }
    tables.insert(line);
  }
  out.insert(out.end(), tables.begin(), tables.end());
  return out;
}

std::string AttrValue(const std::string& attrs, const std::string& key) {
  const std::size_t at = attrs.find(" " + key + "=\"");
  if (at == std::string::npos) {
    return {};
  }
  const std::size_t begin = at + key.size() + 3U;
  return attrs.substr(begin, attrs.find('"', begin) - begin);
}

TEST(WorkbookCopySheet, RoundTripXlsx) {
  // Start from the engine's own save of the fixture, so a 3-D span the writer
  // quotes (`'A:C'!A1`, as Excel stores it) compares equal on both sides.
  Workbook wb = Reload(LoadCopyFixture("main.xlsx"));
  TableMetadata& source = wb.mutable_tables().front();
  source.root_extra_attrs +=
      " xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\""
      " xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\""
      " xr:uid=\"{00000000-000C-0000-FFFF-FFFF00000000}\"";
  source.columns.front().extra_attrs += " xr3:uid=\"{00000000-0010-0000-0000-000004000000}\"";
  SetCodeName(wb.sheet(kSrc), "Sheet2");
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, kSrc + 1U, std::nullopt)));
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(kSrc, wb.sheet_count(), std::nullopt)));
  ASSERT_TRUE(static_cast<bool>(wb.copy_sheet(wb.sheet_index_by_name("H"), 0, std::nullopt)));
  Recalc(wb);

  const Workbook reloaded = Reload(wb);
  const std::vector<std::string> saved = Projection(wb);
  const std::vector<std::string> loaded = Projection(reloaded);
  ASSERT_EQ(loaded.size(), saved.size());
  for (std::size_t i = 0; i < saved.size(); ++i) {
    // A dead sheet's `#REF!A1` is a formula-bar spelling; Excel stores bare `#REF!`.
    std::string expected = saved[i];
    for (std::size_t at = expected.find("#REF!A1"); at != std::string::npos; at = expected.find("#REF!A1", at)) {
      expected.erase(at + 5U, 2U);
    }
    EXPECT_EQ(loaded[i], expected);
  }

  std::set<std::string> paths;
  for (const PassthroughPart& part : reloaded.passthrough_parts()) {
    EXPECT_TRUE(paths.insert(part.path).second) << part.path;
  }
  std::set<std::uint32_t> sheet_ids;
  std::set<std::string> code_names;
  for (std::size_t i = 0; i < reloaded.sheet_count(); ++i) {
    const Sheet& sheet = reloaded.sheet(i);
    EXPECT_NE(sheet.source_sheet_id(), 0U) << sheet.name();
    EXPECT_TRUE(sheet_ids.insert(sheet.source_sheet_id()).second) << sheet.name();
    const std::string code = AttrValue(sheet.print_settings().sheet_pr_xml.substr(0, 200), "codeName");
    if (!code.empty()) {
      EXPECT_TRUE(code_names.insert(code).second) << sheet.name();
    }
    for (std::size_t j = i + 1U; j < reloaded.sheet_count(); ++j) {
      if (!sheet.drawing_rel_target().empty()) {
        EXPECT_NE(sheet.drawing_rel_target(), reloaded.sheet(j).drawing_rel_target());
      }
    }
  }
  EXPECT_EQ(code_names.size(), 1U);
  std::set<std::string> uids;
  for (const TableMetadata& table : reloaded.tables()) {
    const std::string uid = AttrValue(table.root_extra_attrs, "xr:uid");
    EXPECT_FALSE(uid.empty()) << table.name;
    EXPECT_TRUE(uids.insert(uid).second) << table.name;
    for (const TableColumn& column : table.columns) {
      const std::string column_uid = AttrValue(column.extra_attrs, "xr3:uid");
      EXPECT_FALSE(column_uid.empty()) << table.name;
      EXPECT_TRUE(uids.insert(column_uid).second) << table.name;
    }
  }
  EXPECT_EQ(reloaded.tables().size(), 3U);
}

}  // namespace
}  // namespace formulon
