//
// Ptg classes and dynamic-array marks against a Mac Excel 365-produced
// `.xlsb` (`tests/fixtures/excel/xlsb_ptg_class.xlsb`).
//
// Sheet1 holds A1:C3 = row numbers, Z1 = SEQUENCE(2) and one formula every
// third row of column E, each typed into Excel with Formula2: IFS / SWITCH /
// XLOOKUP in every kind of parameter slot, calls nested under array-class
// parameters, and the formulas whose dynamic-array mark (`BrtCellMeta`)
// Excel decides from the shapes of their arguments. Sheet2 repeats A1:C3.
// The defined names cover the classes of a name's body.
//
// Every formula is re-encoded here against Excel's own name and sheet-range
// tables, so the token streams compare byte for byte once the unused fields
// Excel leaves uninitialised are cleared. Array constants are left out:
// their `PtgArray` carries 14 such bytes among others.
//
// `xlsb_ptg_class_legacy.xlsx` holds formulas written without the
// dynamic-array mark, as a pre-dynamic-array workbook stores them, and
// `xlsb_ptg_class_legacy_resaved.xlsb` is that workbook opened and saved by
// Excel 365: where it shows an implicit `@`, and the tokens it writes.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "gtest/gtest.h"
#include "io/dynamic_array_formula.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/zip_reader.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

std::vector<std::uint8_t> ReadFixture(const char* file = "xlsb_ptg_class.xlsb") {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + file;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(f);
  return out;
}

std::uint32_t U32(ByteSpan span, std::size_t at) {
  return static_cast<std::uint32_t>(span.data[at]) | (static_cast<std::uint32_t>(span.data[at + 1]) << 8) |
         (static_cast<std::uint32_t>(span.data[at + 2]) << 16) | (static_cast<std::uint32_t>(span.data[at + 3]) << 24);
}

std::vector<XlsbRecord> Records(const std::vector<std::uint8_t>& part) {
  std::vector<XlsbRecord> out;
  ByteSpan cursor{part.data(), part.size()};
  while (cursor.size > 0U) {
    auto record = read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    out.push_back(record.value());
  }
  return out;
}

/// What the fixture package says in its own bytes.
struct ExcelTokens {
  /// `rgce` per (row, col) of Sheet1, a dynamic-array cell's from its `BrtArrFmla`.
  std::map<std::pair<std::uint32_t, std::uint32_t>, std::vector<std::uint8_t>> cells;
  /// `rgce` and `fCalcExp` per user-defined name.
  std::map<std::string, std::vector<std::uint8_t>> names;
  std::map<std::string, bool> calc_exp;
  NameTable name_table;
  SheetRangeTable sheet_ranges;
};

ExcelTokens ReadExcelTokens(const std::vector<std::uint8_t>& package) {
  ExcelTokens out;
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
  auto sheet = zip.read_entry("xl/worksheets/sheet1.bin");
  auto book = zip.read_entry("xl/workbook.bin");
  if (!sheet || !book) {
    ADD_FAILURE() << "fixture parts missing";
    return out;
  }
  std::uint32_t row = 0;
  std::map<std::pair<std::uint32_t, std::uint32_t>, std::vector<std::uint8_t>> arrays;
  for (const XlsbRecord& r : Records(sheet.value())) {
    const ByteSpan p = r.payload;
    if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr)) {
      row = U32(p, 0);
    } else if (r.type >= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaString) &&
               r.type <= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaError)) {
      // Cell header, then the cached value, the flags and the formula.
      std::size_t at = 8;
      if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaString)) {
        at += 4U + 2U * U32(p, at);
      } else if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaNum)) {
        at += 8U;
      } else {
        at += 1U;
      }
      at += 2U;
      const std::uint32_t cce = U32(p, at);
      out.cells[{row, U32(p, 0)}] = std::vector<std::uint8_t>(p.data + at + 4, p.data + at + 4 + cce);
    } else if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla)) {
      const std::uint32_t cce = U32(p, 17);
      arrays[{U32(p, 0), U32(p, 8)}] = std::vector<std::uint8_t>(p.data + 21, p.data + 21 + cce);
    }
  }
  for (auto& [key, rgce] : arrays) {
    out.cells[key] = std::move(rgce);
  }
  std::uint32_t ilbl = 0;
  for (const XlsbRecord& r : Records(book.value())) {
    const ByteSpan p = r.payload;
    if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtName)) {
      ++ilbl;
      const auto itab = static_cast<std::int32_t>(U32(p, 5));
      const std::uint32_t chars = U32(p, 9);
      std::string name;
      for (std::uint32_t i = 0; i < chars; ++i) {
        name.push_back(static_cast<char>(p.data[13 + 2 * i]));  // the fixture's names are ASCII
      }
      const std::size_t at = 13U + 2U * chars;
      const std::uint32_t cce = U32(p, at);
      if (name.rfind("_xl", 0) != 0) {
        out.names[name] = std::vector<std::uint8_t>(p.data + at + 4, p.data + at + 4 + cce);
        out.calc_exp[name] = (p.data[0] & 0x10U) != 0U;
      }
      out.name_table.emplace(itab < 0 ? name : sheet_scoped_name_key(itab, name), ilbl);
    } else if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtExternSheet)) {
      const std::uint32_t count = U32(p, 0);
      for (std::uint32_t i = 0; i < count; ++i) {
        const std::size_t at = 4U + 12U * i;  // iSupBook, itabFirst, itabLast
        out.sheet_ranges.emplace_back(static_cast<std::int32_t>(U32(p, at + 4)),
                                      static_cast<std::int32_t>(U32(p, at + 8)));
      }
    }
  }
  return out;
}

Workbook LoadFixture(const std::vector<std::uint8_t>& package) {
  auto result = read_xlsb(ByteSpan{package.data(), package.size()});
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

/// `rgce` with the unused fields Excel leaves uninitialised cleared: those
/// of a leading `PtgAttrSemi` and of the `PtgMemArea` a root reference
/// operation starts with.
std::vector<std::uint8_t> ClearUnused(std::vector<std::uint8_t> rgce) {
  std::size_t at = 0;
  if (rgce.size() >= 4U && rgce[0] == 0x19 && rgce[1] == 0x01) {
    rgce[2] = rgce[3] = 0;
    at = 4;
  }
  if (rgce.size() >= at + 5U && (rgce[at] & 0x1F) == 0x06) {
    std::fill(rgce.begin() + static_cast<std::ptrdiff_t>(at) + 1, rgce.begin() + static_cast<std::ptrdiff_t>(at) + 5,
              0);
  }
  return rgce;
}

std::string Body(const std::string& formula) {
  return formula.empty() || formula.front() != '=' ? formula : formula.substr(1);
}

TEST(XlsbPtgClassFixture, EntryMarksMatchExcel) {
  const std::vector<std::uint8_t> package = ReadFixture();
  Workbook wb = LoadFixture(package);
  ASSERT_GE(wb.sheet_count(), 1U);
  std::size_t checked = 0;
  for (const CellAddress a : wb.sheet(0).formula_cells_in(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U)) {
    const Cell* cell = wb.sheet(0).cell_at(a.row, a.col);
    const std::string body = Body(cell->formula_text);
    Arena arena;
    parser::Parser p(body, arena);
    const parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << body;
    EXPECT_EQ(entered_as_dynamic_array(wb, 0, *root), cell->dynamic_array) << body;
    ++checked;
  }
  EXPECT_EQ(checked, 147U);
}

// Every formula of Excel-saved probe workbooks, re-entered, takes the mark
// Excel gave it: each built-in called with an area, a cell, an area operand
// and a cell operand (`tools/dev/xlsb_func_id_harvest.py classes-build`), and
// IFS / SWITCH / XLOOKUP and shape-dependent formulas in every slot.
TEST(XlsbPtgClassFixture, EntryMarksMatchExcelAcrossBuiltins) {
  const std::pair<const char*, std::size_t> probes[] = {
      {"xlsb_cm_marks_area.xlsb", 523U},    {"xlsb_cm_marks_cell.xlsb", 523U},  {"xlsb_cm_marks_area_op.xlsb", 491U},
      {"xlsb_cm_marks_cell_op.xlsb", 491U}, {"xlsb_cm_marks_slots.xlsb", 183U}, {"xlsb_cm_marks_shapes.xlsb", 134U}};
  for (const auto& [file, expected] : probes) {
    const std::vector<std::uint8_t> package = ReadFixture(file);
    Workbook wb = LoadFixture(package);
    std::size_t checked = 0;
    for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
      for (const CellAddress a : wb.sheet(s).formula_cells_in(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U)) {
        const Cell* cell = wb.sheet(s).cell_at(a.row, a.col);
        const std::string body = Body(cell->formula_text);
        Arena arena;
        parser::Parser p(body, arena);
        const parser::AstNode* root = p.parse();
        ASSERT_NE(root, nullptr) << file << ": " << body;
        EXPECT_EQ(entered_as_dynamic_array(wb, s, *root), cell->dynamic_array) << file << ": " << body;
        ++checked;
      }
    }
    EXPECT_EQ(checked, expected) << file;
  }
}

TEST(XlsbPtgClassFixture, CellFormulaTokensMatchExcel) {
  const std::vector<std::uint8_t> package = ReadFixture();
  Workbook wb = LoadFixture(package);
  const ExcelTokens excel = ReadExcelTokens(package);
  const std::vector<std::string> sheets = {"Sheet1", "Sheet2"};
  std::size_t checked = 0;
  for (const CellAddress a : wb.sheet(0).formula_cells_in(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U)) {
    const Cell* cell = wb.sheet(0).cell_at(a.row, a.col);
    const std::string body = Body(cell->formula_text);
    if (body.find('{') != std::string::npos) {
      continue;
    }
    Arena arena;
    parser::Parser p(body, arena);
    const parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << body;
    auto encoded =
        encode_ptgs(*root, sheets, excel.sheet_ranges, excel.name_table, PtgRootClass::kValue, std::nullopt,
                    cell->dynamic_array ? PtgEvaluation::kDynamicArray : PtgEvaluation::kLegacy, name_shapes(wb, 0));
    ASSERT_TRUE(static_cast<bool>(encoded)) << body;
    const auto it = excel.cells.find({a.row, a.col});
    ASSERT_NE(it, excel.cells.end()) << body;
    EXPECT_EQ(ClearUnused(encoded.value().rgce), ClearUnused(it->second)) << body;
    ++checked;
  }
  EXPECT_EQ(checked, 132U);
}

TEST(XlsbPtgClassFixture, DefinedNameTokensMatchExcel) {
  const std::vector<std::uint8_t> package = ReadFixture();
  Workbook wb = LoadFixture(package);
  const ExcelTokens excel = ReadExcelTokens(package);
  const std::vector<std::string> sheets = {"Sheet1", "Sheet2"};
  std::size_t checked = 0;
  for (const DefinedName& name : wb.defined_names()) {
    const auto it = excel.names.find(name.name);
    ASSERT_NE(it, excel.names.end()) << name.name;
    const std::string body = Body(name.formula);
    Arena arena;
    parser::Parser p(body, arena);
    const parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << body;
    auto encoded = encode_ptgs(*root, sheets, excel.sheet_ranges, excel.name_table, PtgRootClass::kReference);
    ASSERT_TRUE(static_cast<bool>(encoded)) << name.name;
    EXPECT_EQ(encoded.value().rgce, it->second) << name.name << " = " << body;
    ++checked;
  }
  EXPECT_EQ(checked, 32U);
}

TEST(XlsbPtgClassFixture, LegacyFormulasIntersectAndEncodeAsExcel) {
  const std::vector<std::uint8_t> legacy = ReadFixture("xlsb_ptg_class_legacy.xlsx");
  auto loaded = read_ooxml(ByteSpan{legacy.data(), legacy.size()});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;
  const ExcelTokens excel = ReadExcelTokens(ReadFixture("xlsb_ptg_class_legacy_resaved.xlsb"));
  // Excel's formula2 text of each, one every third row of column E.
  const char* const shown[] = {"=SUMPRODUCT(ABS(A1))",
                               "=ROWS(ABS(A1))",
                               "=ROWS(ABS(A1)*2)",
                               "=SUMPRODUCT(A1*2)",
                               "=IF(ISNUMBER(A1),1,0)",
                               "=SUM(ABS(A1))",
                               "=N(ABS(A1))",
                               "=ROWS(A1*2)",
                               "=INDEX(A1:B2,1,1)+0",
                               "=SUMPRODUCT(INDEX(A1:B2,1,1))",
                               "=LET(x,1,x+1)",
                               "=SUM(@A1:A2*2)",
                               "=ABS(@A1:A2)",
                               "=SUMPRODUCT(ABS(A1:A2))",
                               "=@A1:A2",
                               "=SUM(SEQUENCE(2))",
                               "=@INDEX(A1:B2,1,0)",
                               "=@IF(A1>0,A1:A2,0)"};
  const std::vector<std::string> sheets = {"Sheet1"};
  std::uint32_t row = 0;
  for (const char* text : shown) {
    const Cell* cell = wb.sheet(0).cell_at(row, 4U);
    ASSERT_NE(cell, nullptr) << text;
    EXPECT_FALSE(cell->dynamic_array) << text;
    EXPECT_EQ(cell->formula_text, text);
    const std::string body = Body(cell->formula_text);
    Arena arena;
    parser::Parser p(body, arena);
    const parser::AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << body;
    auto encoded = encode_ptgs(*root, sheets, excel.sheet_ranges, excel.name_table, PtgRootClass::kValue, std::nullopt,
                               PtgEvaluation::kLegacy, name_shapes(wb, 0));
    ASSERT_TRUE(static_cast<bool>(encoded)) << body;
    const auto it = excel.cells.find({row, 4U});
    ASSERT_NE(it, excel.cells.end()) << body;
    EXPECT_EQ(encoded.value().rgce, it->second) << body;
    row += 3U;
  }
}

// `fCalcExp` on every defined name of Excel-saved workbooks: one name per
// built-in called with cells and with areas, those calls nested in others,
// under operators and parentheses, and names referring to flagged names.
TEST(XlsbPtgClassFixture, DefinedNameCalcExpMatchesExcel) {
  const std::pair<const char*, std::size_t> books[] = {{"xlsb_calc_exp_builtins.xlsb", 1042U},
                                                       {"xlsb_calc_exp_nesting.xlsb", 17U},
                                                       {"xlsb_calc_exp_references.xlsb", 17U},
                                                       {"xlsb_calc_exp_localized.xlsb", 3U},
                                                       {"xlsb_ptg_class.xlsb", 32U}};
  for (const auto& [file, expected] : books) {
    const std::vector<std::uint8_t> package = ReadFixture(file);
    Workbook wb = LoadFixture(package);
    const ExcelTokens excel = ReadExcelTokens(package);
    const std::vector<NameShape> shapes = defined_name_shapes(wb);
    ASSERT_EQ(shapes.size(), wb.defined_names().size()) << file;
    std::size_t checked = 0;
    for (std::size_t i = 0; i < shapes.size(); ++i) {
      const DefinedName& name = wb.defined_names()[i];
      // Excel's name API reads the ja-JP function names, where USDOLLAR and
      // DBCS are not function names, so it stored those two harvest calls
      // through undefined names (a leading PtgName), which read back as the
      // built-ins; xlsb_calc_exp_localized.xlsb covers ids 204 and 215.
      const auto rgce = excel.names.find(name.name);
      if (rgce != excel.names.end() && !rgce->second.empty() && rgce->second[0] == 0x23 &&
          (name.formula.rfind("USDOLLAR(", 0) == 0 || name.formula.rfind("DBCS(", 0) == 0)) {
        continue;
      }
      const auto it = excel.calc_exp.find(name.name);
      ASSERT_NE(it, excel.calc_exp.end()) << file << ": " << name.name;
      EXPECT_EQ(shapes[i].calc_exp, it->second) << file << ": " << name.name << " = " << name.formula;
      ++checked;
    }
    EXPECT_EQ(checked, expected) << file;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
