// A defined name stored without the `_xlnm.` prefix but spelled as the load locale names a built-in opens as that
// built-in, in both package formats, and the cells naming it follow.

#include <algorithm>
#include <array>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/read_options.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "support/ooxml_package_fixture.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr ExcelProfile kGerman{ExcelHost::kMac365, ExcelLocale::kDeDE};
constexpr ExcelProfile kFrench{ExcelHost::kMac365, ExcelLocale::kFrFR};
constexpr ExcelProfile kEnglish = mac_365_en_us_profile();

struct FormulaCell {
  int sheet;  // 0-based
  int row;    // 1-based, column E
  const char* text;
};

std::string DefinedNameXml(const char* name, const char* body, int local_sheet_id = -1) {
  std::string out = "<definedName name=\"";
  out += name;
  out += '"';
  if (local_sheet_id >= 0) {
    out += " localSheetId=\"" + std::to_string(local_sheet_id) + "\"";
  }
  out += '>';
  out += body;
  out += "</definedName>";
  return out;
}

std::string InlineText(const char* ref, const char* text) {
  return std::string("<c r=\"") + ref + "\" t=\"inlineStr\"><is><t>" + text + "</t></is></c>";
}

std::string NumberCell(const char* ref, int value) {
  return std::string("<c r=\"") + ref + "\"><v>" + std::to_string(value) + "</v></c>";
}

// Sheet n of three holds k/v, a/(3+n), b/(5+n) in A1:B3 and the given formulas in column E.
std::string SheetXml(int sheet, const std::vector<FormulaCell>& formulas) {
  std::string out = "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>";
  constexpr std::array<const char*, 3> kKeys{"k", "a", "b"};
  for (int row = 1; row <= 3; ++row) {
    const std::string a = "A" + std::to_string(row);
    const std::string b = "B" + std::to_string(row);
    out += "<row r=\"" + std::to_string(row) + "\">";
    out += InlineText(a.c_str(), kKeys.at(static_cast<std::size_t>(row - 1)));
    out += row == 1 ? InlineText(b.c_str(), "v") : NumberCell(b.c_str(), (row == 2 ? 3 : 5) + sheet);
    for (const FormulaCell& f : formulas) {
      if (f.sheet == sheet && f.row == row) {
        out += "<c r=\"E" + std::to_string(row) + "\"><f>" + f.text + "</f></c>";
      }
    }
    out += "</row>";
  }
  out += "</sheetData></worksheet>";
  return out;
}

// A three-sheet xlsx (Sheet1..Sheet3) with `names_xml` as its definedNames and the given formulas.
std::vector<std::uint8_t> AuthorBook(const std::string& names_xml, const std::vector<FormulaCell>& formulas) {
  std::string types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Override PartName=\"/xl/workbook.xml\" ContentType=\"";
  types += test::kXlsxWorkbookContentType;
  types += "\"/>";
  std::string workbook =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>";
  std::string rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
  for (int i = 1; i <= 3; ++i) {
    const std::string n = std::to_string(i);
    types += "<Override PartName=\"/xl/worksheets/sheet" + n +
             ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
    workbook += "<sheet name=\"Sheet" + n + "\" sheetId=\"" + n + "\" r:id=\"rId" + n + "\"/>";
    rels += "<Relationship Id=\"rId" + n +
            "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
            "Target=\"worksheets/sheet" +
            n + ".xml\"/>";
  }
  types += "</Types>";
  workbook += "</sheets><definedNames>" + names_xml + "</definedNames></workbook>";
  rels += "</Relationships>";
  const std::string sheet1 = SheetXml(0, formulas);
  const std::string sheet2 = SheetXml(1, formulas);
  const std::string sheet3 = SheetXml(2, formulas);
  return test::BuildOoxmlPackage(types, sheet1, workbook, rels,
                                 {{"xl/worksheets/sheet2.xml", sheet2}, {"xl/worksheets/sheet3.xml", sheet3}});
}

io::ReadOptions Under(ExcelProfile profile) {
  io::ReadOptions options;
  options.profile = profile;
  return options;
}

Workbook LoadXlsx(const std::vector<std::uint8_t>& bytes, ExcelProfile profile) {
  auto loaded = io::read_ooxml(test::span_of(bytes), Under(profile));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

// The same book as an .xlsb: loaded where no spelling folds, then saved by the engine.
std::vector<std::uint8_t> AsXlsb(const std::vector<std::uint8_t>& xlsx, ExcelProfile unfolding) {
  const Workbook wb = LoadXlsx(xlsx, unfolding);
  auto saved = io::xlsb::write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes, ExcelProfile profile) {
  auto loaded = io::xlsb::read_xlsb(test::span_of(bytes), Under(profile));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

using NameList = std::vector<std::pair<std::string, std::int32_t>>;

NameList Names(const Workbook& wb) {
  NameList out;
  for (const DefinedName& entry : wb.defined_names()) {
    out.emplace_back(entry.name, entry.local_sheet_id);
  }
  std::sort(out.begin(), out.end());
  return out;
}

const DefinedName* Find(const Workbook& wb, std::string_view name, std::int32_t scope) {
  for (const DefinedName& entry : wb.defined_names()) {
    if (entry.name == name && entry.local_sheet_id == scope) {
      return &entry;
    }
  }
  return nullptr;
}

// The column-E formula of `row`, without its leading `=`.
std::string FormulaAt(const Workbook& wb, std::size_t sheet, std::uint32_t row) {
  const Cell* cell = wb.sheet(sheet).cell_at(row - 1U, 4U);
  if (cell == nullptr) {
    return "<no cell>";
  }
  const std::string& text = cell->formula_text;
  return !text.empty() && text.front() == '=' ? text.substr(1) : text;
}

Value ValueAt(Workbook& wb, std::size_t sheet, std::uint32_t row) {
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* cell = wb.sheet(sheet).cell_at(row - 1U, 4U);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

void ExpectNumber(const Value& v, double expected) {
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), expected);
}

void ExpectNameError(const Value& v) {
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

const std::string kControl = DefinedNameXml("Druckzone", "Sheet1!$A$1:$A$4");

// Both package formats of `xlsx`, loaded under `profile`; the .xlsb is written from a load under `unfolding`.
std::vector<std::pair<const char*, Workbook>> BothFormats(const std::vector<std::uint8_t>& xlsx, ExcelProfile profile,
                                                          ExcelProfile unfolding) {
  std::vector<std::pair<const char*, Workbook>> out;
  out.emplace_back("xlsx", LoadXlsx(xlsx, profile));
  out.emplace_back("xlsb", LoadXlsb(AsXlsb(xlsx, unfolding), profile));
  return out;
}

TEST(LoadedBuiltinNameFold, GermanPrintAreaSpellingOpensAsThePrintAreaInItsScope) {
  for (const std::int32_t scope : {-1, 0}) {
    SCOPED_TRACE(scope);
    const auto book = AuthorBook(DefinedNameXml("Druckbereich", "Sheet1!$A$1:$A$3", scope) + kControl,
                                 {{0, 1, "ROWS(Druckbereich)"}, {0, 2, "ROWS(Druckzone)"}});
    for (auto& [format, wb] : BothFormats(book, kGerman, kEnglish)) {
      SCOPED_TRACE(format);
      EXPECT_EQ(Names(wb), (NameList{{"Druckzone", -1}, {"_xlnm.Print_Area", scope}}));
      EXPECT_EQ(Find(wb, "_xlnm.Print_Area", scope)->formula, "Sheet1!$A$1:$A$3");
      EXPECT_EQ(FormulaAt(wb, 0, 1), "ROWS(_xlnm.Print_Area)");
      EXPECT_EQ(FormulaAt(wb, 0, 2), "ROWS(Druckzone)");
      ExpectNumber(ValueAt(wb, 0, 1), 3.0);
      ExpectNumber(ValueAt(wb, 0, 2), 4.0);
    }
  }
}

TEST(LoadedBuiltinNameFold, FoldedNameKeepsItsBody) {
  const auto book = AuthorBook(DefinedNameXml("Druckbereich", "Sheet2!$A$1:$B$2") + kControl,
                               {{0, 1, "ROWS(Druckbereich)"}, {0, 2, "ROWS(Druckzone)"}});
  for (auto& [format, wb] : BothFormats(book, kGerman, kEnglish)) {
    SCOPED_TRACE(format);
    ASSERT_NE(Find(wb, "_xlnm.Print_Area", -1), nullptr);
    EXPECT_EQ(Find(wb, "_xlnm.Print_Area", -1)->formula, "Sheet2!$A$1:$B$2");
    ExpectNumber(ValueAt(wb, 0, 1), 2.0);
  }
}

TEST(LoadedBuiltinNameFold, EveryLocalBuiltinSpellingFolds) {
  struct Case {
    ExcelProfile profile;
    const char* typed;
    const char* stored;
  };
  for (const Case& c :
       {Case{kGerman, "Suchkriterien", "_xlnm.Criteria"}, Case{kGerman, "Drucktitel", "_xlnm.Print_Titles"},
        Case{kGerman, u8"Auto_\u00F6ffnen", "_xlnm.Auto_Open"}, Case{kFrench, "Zone_d_impression", "_xlnm.Print_Area"},
        Case{kEnglish, "Print_Area", "_xlnm.Print_Area"}}) {
    SCOPED_TRACE(c.typed);
    const std::string cell = std::string("ROWS(") + c.typed + ")";
    const std::string folded = std::string("ROWS(") + c.stored + ")";
    const auto book = AuthorBook(DefinedNameXml(c.typed, "Sheet1!$A$1:$A$3") + kControl, {{0, 1, cell.c_str()}});
    const ExcelProfile unfolding = c.profile.locale == ExcelLocale::kEnUS ? kGerman : kEnglish;
    for (auto& [format, wb] : BothFormats(book, c.profile, unfolding)) {
      SCOPED_TRACE(format);
      EXPECT_EQ(Names(wb), (NameList{{"Druckzone", -1}, {c.stored, -1}}));
      EXPECT_EQ(FormulaAt(wb, 0, 1), folded);
      ExpectNumber(ValueAt(wb, 0, 1), 3.0);
    }
  }
}

TEST(LoadedBuiltinNameFold, AnotherLocalesSpellingStaysOrdinary) {
  struct Case {
    ExcelProfile profile;
    const char* typed;
  };
  for (const Case& c : {Case{kEnglish, "Druckbereich"}, Case{kGerman, "Print_Area"}}) {
    SCOPED_TRACE(c.typed);
    const std::string cell = std::string("ROWS(") + c.typed + ")";
    const auto book = AuthorBook(DefinedNameXml(c.typed, "Sheet1!$A$1:$A$3") + kControl, {{0, 1, cell.c_str()}});
    for (auto& [format, wb] : BothFormats(book, c.profile, c.profile)) {
      SCOPED_TRACE(format);
      NameList expected{{c.typed, -1}, {"Druckzone", -1}};
      std::sort(expected.begin(), expected.end());
      EXPECT_EQ(Names(wb), expected);
      EXPECT_EQ(FormulaAt(wb, 0, 1), cell);
      ExpectNumber(ValueAt(wb, 0, 1), 3.0);
    }
  }
}

TEST(LoadedBuiltinNameFold, NameBodiesFollowTheFold) {
  const auto book = AuthorBook(
      DefinedNameXml("Druckbereich", "Sheet1!$A$1:$A$3") + DefinedNameXml("Zeilen", "ROWS(Druckbereich)") + kControl,
      {{0, 1, "Zeilen"}});
  for (auto& [format, wb] : BothFormats(book, kGerman, kEnglish)) {
    SCOPED_TRACE(format);
    ASSERT_NE(Find(wb, "Zeilen", -1), nullptr);
    EXPECT_EQ(Find(wb, "Zeilen", -1)->formula, "ROWS(_xlnm.Print_Area)");
    ExpectNumber(ValueAt(wb, 0, 1), 3.0);
  }
}

// An unbound local spelling in a cell is stored as the English spelling without the prefix, and stays unbound.
TEST(LoadedBuiltinNameFold, UnboundLocalSpellingInACellTakesTheEnglishSpelling) {
  struct Case {
    std::string names;
    FormulaCell cell;
    const char* stored;
  };
  for (const Case& c : {Case{kControl, {0, 1, "ROWS(Druckbereich)"}, "ROWS(Print_Area)"},
                        Case{DefinedNameXml("_xlnm.Print_Area", "Sheet1!$A$1:$A$5", 0) + kControl,
                             {1, 1, "ROWS(Druckbereich)"},
                             "ROWS(Print_Area)"},
                        Case{kControl, {0, 1, "ROWS(Suchkriterien)"}, "ROWS(Criteria)"}}) {
    SCOPED_TRACE(c.stored);
    const auto book = AuthorBook(c.names, {c.cell, {0, 2, "ROWS(Druckzone)"}});
    Workbook wb = LoadXlsx(book, kGerman);
    const auto sheet = static_cast<std::size_t>(c.cell.sheet);
    EXPECT_EQ(FormulaAt(wb, sheet, 1), c.stored);
    ExpectNameError(ValueAt(wb, sheet, 1));
    ExpectNumber(ValueAt(wb, 0, 2), 4.0);
  }
}

// A workbook-scope local spelling beside a Sheet1-local print area folds; the cell on Sheet2 reads the folded entry.
TEST(LoadedBuiltinNameFold, WorkbookSpellingBesideASheetPrintAreaFolds) {
  const auto book = AuthorBook(
      DefinedNameXml("Druckbereich", "Sheet1!$A$1:$A$3") + DefinedNameXml("_xlnm.Print_Area", "Sheet1!$A$1:$A$5", 0),
      {{0, 1, "ROWS(Druckbereich)"}, {1, 1, "ROWS(Druckbereich)"}});
  for (auto& [format, wb] : BothFormats(book, kGerman, kEnglish)) {
    SCOPED_TRACE(format);
    EXPECT_EQ(Names(wb), (NameList{{"_xlnm.Print_Area", -1}, {"_xlnm.Print_Area", 0}}));
    EXPECT_EQ(Find(wb, "_xlnm.Print_Area", -1)->formula, "Sheet1!$A$1:$A$3");
    EXPECT_EQ(FormulaAt(wb, 1, 1), "ROWS(_xlnm.Print_Area)");
    ExpectNumber(ValueAt(wb, 1, 1), 3.0);
  }
}

// Where the built-in of that name is visible from the spelling's scope, Excel demands a new name before it opens the
// book; the load keeps both entries as stored (divergence
// visible_builtin_user_name_equal_to_local_spelling_kept_at_load). A built-in local to another sheet is not visible,
// so that spelling folds.
TEST(LoadedBuiltinNameFold, SpellingWithAVisibleBuiltinIsKeptAsStored) {
  struct Case {
    std::int32_t ordinary_scope;
    std::int32_t builtin_scope;
    bool folds;
  };
  for (const Case& c : {Case{-1, -1, false}, Case{0, 0, false}, Case{0, -1, false}, Case{1, 0, true}}) {
    SCOPED_TRACE(std::to_string(c.ordinary_scope) + "/" + std::to_string(c.builtin_scope));
    const auto book = AuthorBook(DefinedNameXml("Druckbereich", "Sheet1!$A$1:$A$3", c.ordinary_scope) +
                                     DefinedNameXml("_xlnm.Print_Area", "Sheet1!$A$1:$A$5", c.builtin_scope),
                                 {});
    for (auto& [format, wb] : BothFormats(book, kGerman, kEnglish)) {
      SCOPED_TRACE(format);
      const NameList expected =
          c.folds ? NameList{{"_xlnm.Print_Area", c.builtin_scope}, {"_xlnm.Print_Area", c.ordinary_scope}}
                  : NameList{{"Druckbereich", c.ordinary_scope}, {"_xlnm.Print_Area", c.builtin_scope}};
      NameList sorted = expected;
      std::sort(sorted.begin(), sorted.end());
      EXPECT_EQ(Names(wb), sorted);
    }
  }
}

// de-DE, Druckbereich local to one sheet beside a print area stored as _xlnm. on another, and without one: the
// spelling folds, and only the renamed entry is reached by it. On the sheet whose print area was stored as _xlnm., and
// on a sheet with neither, the cell is unbound (Excel's open and .xlsx save of these three books).
TEST(LoadedBuiltinNameFold, OldSpellingReachesOnlyTheRenamedEntry) {
  struct Case {
    std::int32_t ordinary_scope;
    std::int32_t builtin_scope;  // -2: none
  };
  for (const Case& c : {Case{0, 1}, Case{1, 0}, Case{0, -2}}) {
    SCOPED_TRACE(std::to_string(c.ordinary_scope) + "/" + std::to_string(c.builtin_scope));
    const std::string ordinary_sheet = "Sheet" + std::to_string(c.ordinary_scope + 1);
    std::string names = DefinedNameXml("Druckbereich", (ordinary_sheet + "!$A$1:$A$2").c_str(), c.ordinary_scope) +
                        DefinedNameXml("Ctl_8", "Sheet1!$A$1:$A$4");
    if (c.builtin_scope >= 0) {
      const std::string builtin_sheet = "Sheet" + std::to_string(c.builtin_scope + 1);
      names += DefinedNameXml("_xlnm.Print_Area", (builtin_sheet + "!$A$1:$A$3").c_str(), c.builtin_scope);
    }
    const auto book = AuthorBook(names, {{0, 1, "ROWS(Druckbereich)"},
                                         {1, 1, "ROWS(Druckbereich)"},
                                         {2, 1, "ROWS(Druckbereich)"},
                                         {0, 2, "ROWS(Ctl_8)"}});
    Workbook wb = LoadXlsx(book, kGerman);
    NameList expected{{"Ctl_8", -1}, {"_xlnm.Print_Area", c.ordinary_scope}};
    if (c.builtin_scope >= 0) {
      expected.emplace_back("_xlnm.Print_Area", c.builtin_scope);
    }
    std::sort(expected.begin(), expected.end());
    EXPECT_EQ(Names(wb), expected);
    for (std::size_t sheet = 0; sheet < 3U; ++sheet) {
      SCOPED_TRACE(sheet);
      if (static_cast<std::int32_t>(sheet) == c.ordinary_scope) {
        EXPECT_EQ(FormulaAt(wb, sheet, 1), "ROWS(_xlnm.Print_Area)");
        ExpectNumber(ValueAt(wb, sheet, 1), 2.0);
      } else {
        EXPECT_EQ(FormulaAt(wb, sheet, 1), "ROWS(Print_Area)");
        ExpectNameError(ValueAt(wb, sheet, 1));
      }
    }
    ExpectNumber(ValueAt(wb, 0, 2), 4.0);
  }
}

// A folded name is saved as the built-in in both formats.
TEST(LoadedBuiltinNameFold, FoldedNameIsSavedAsTheBuiltin) {
  const auto book =
      AuthorBook(DefinedNameXml("Druckbereich", "Sheet1!$A$1:$A$3", 0) + kControl, {{0, 1, "ROWS(Druckbereich)"}});
  const Workbook wb = LoadXlsx(book, kGerman);
  auto xlsx = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  auto xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  // An English load sees the saved built-in, not a spelling it would leave ordinary.
  std::vector<Workbook> reloaded;
  reloaded.push_back(LoadXlsx(xlsx.value(), kEnglish));
  reloaded.push_back(LoadXlsb(xlsb.value(), kEnglish));
  for (const Workbook& again : reloaded) {
    EXPECT_EQ(Names(again), (NameList{{"Druckzone", -1}, {"_xlnm.Print_Area", 0}}));
    EXPECT_EQ(FormulaAt(again, 0, 1), "ROWS(_xlnm.Print_Area)");
  }
}

}  // namespace
}  // namespace formulon
