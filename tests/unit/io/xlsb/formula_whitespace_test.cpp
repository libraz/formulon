//
// Formula whitespace through `PtgAttrSpace`, against a Mac Excel 365-produced
// `.xlsb` (`tests/fixtures/excel/xlsb_formula_whitespace.xlsb`).
//
// Sheet S holds A1 = 1, A2 = 2 and, from row 3, one formula per row typed
// with whitespace between its tokens: column B as plain text, column C with
// its leading whitespace kept. Excel drops a space before `,`, around `:`,
// inside an array constant and at the end when it reads the text, and stores
// the rest in the token stream, which reads back to the same text.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/dynamic_array_formula.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "parser/ast_format.h"
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

std::vector<std::uint8_t> ReadFixture() {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_formula_whitespace.xlsb";
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

/// `rgce` per (row, col) of the fixture's sheet, and its sheet-range table.
struct ExcelTokens {
  std::map<std::pair<std::uint32_t, std::uint32_t>, std::vector<std::uint8_t>> cells;
  SheetRangeTable sheet_ranges;
};

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
    }
  }
  for (const XlsbRecord& r : Records(book.value())) {
    const ByteSpan p = r.payload;
    if (r.type == static_cast<std::uint16_t>(XlsbRecordType::BrtExternSheet)) {
      const std::uint32_t count = U32(p, 0);
      for (std::uint32_t i = 0; i < count; ++i) {
        const std::size_t at = 4U + 12U * i;  // iSupBook, itabFirst, itabLast
        out.sheet_ranges.xti.push_back(
            XtiEntry{0U, static_cast<std::int32_t>(U32(p, at + 4)), static_cast<std::int32_t>(U32(p, at + 8))});
      }
    }
  }
  return out;
}

Workbook Load(const std::vector<std::uint8_t>& package) {
  auto result = read_xlsb(ByteSpan{package.data(), package.size()});
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

/// `rgce` with the unused bytes of a leading `PtgMemArea` cleared.
std::vector<std::uint8_t> ClearUnused(std::vector<std::uint8_t> rgce) {
  if (rgce.size() >= 5U && (rgce[0] & 0x1F) == 0x06) {
    std::fill(rgce.begin() + 1, rgce.begin() + 5, 0);
  }
  return rgce;
}

std::string Body(const std::string& formula) {
  return formula.empty() || formula.front() != '=' ? formula : formula.substr(1);
}

struct Row {
  std::uint32_t row;
  const char* plain;      // column B
  const char* preserved;  // column C
};

// What Excel shows for each formula after reopening its own `.xlsb`.
constexpr Row kExcelText[] = {
    {2U, "=- 5", "=- 5"},
    {3U, "=-  5", "=-  5"},
    {4U, "=1 +2", "=1 +2"},
    {5U, "=1+ 2", "=1+ 2"},
    {6U, "=1 + 2", "=1 + 2"},
    {7U, "=1  +   2", "=1  +   2"},
    {8U, "=SUM( 1, 2 )", "=SUM( 1, 2 )"},
    {9U, "=SUM(1, 2)", "=SUM(1, 2)"},
    {10U, "=SUM(1,\n2)", "=SUM(1,\n2)"},
    {11U, "=1+\n2", "=1+\n2"},
    {12U, "=1", "= 1"},
    {13U, "=1+2", "=  1+2"},
    {14U, "=1+2", "=1+2"},
    {15U, "=1+2", "=1+2"},
    {16U, "=1+2", "= 1+2"},
    {17U, "=(1 )+2", "=(1 )+2"},
    {18U, "=( 1+2 )", "=( 1+2 )"},
    {19U, "=SUM(A1:A2)", "=SUM(A1:A2)"},
    {20U, "=SUM(A1:A2)", "=SUM(A1:A2)"},
    {21U, "=SUM(A1:A2)", "=SUM(A1:A2)"},
    {22U, "=1+ 5", "=1+ 5"},
    {23U, "=1 + -5", "=1 + -5"},
    {24U, "=A1 A1", "=A1 A1"},
    {25U, "=A1  A1", "=A1  A1"},
    {26U, "=1 &\"a\"", "=1 &\"a\""},
    {27U, "={1,2}", "={1,2}"},
    {28U, "={1,2}", "={1,2}"},
    {29U, "=-( 5)", "=-( 5)"},
    {30U, "=5 %", "=5 %"},
    {31U, "=SUM(\n1)", "=SUM(\n1)"},
    {32U, "=1+2", "=\n1+2"},
    {33U, "=1+2", "=1+2"},
    {34U, "=IF( A1 >0, 1, 2 )", "=IF( A1 >0, 1, 2 )"},
    {35U, "=A1 =1", "=A1 =1"},
    {36U, "=1 <> 2", "=1 <> 2"},
    {37U, "=1 ^ 2", "=1 ^ 2"},
    {38U, "=\"a\" & \"b\"", "=\"a\" & \"b\""},
    {39U, "=-\n5", "=-\n5"},
};

std::string FormulaAt(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(0).cell_at(row, col);
  return cell == nullptr ? std::string("<no cell>") : cell->formula_text;
}

// An array constant in a formula without the dynamic-array mark reads back
// behind an implicit `@`, which the text column above, Excel's legacy
// spelling, does not show; those rows carry no whitespace and are skipped.
bool IsArrayConstantRow(const Row& r) {
  return std::string(r.plain).find('{') != std::string::npos;
}

TEST(XlsbFormulaWhitespace, ReadsExcelText) {
  const Workbook wb = Load(ReadFixture());
  ASSERT_GE(wb.sheet_count(), 1U);
  for (const Row& r : kExcelText) {
    if (IsArrayConstantRow(r)) {
      continue;
    }
    EXPECT_EQ(FormulaAt(wb, r.row, 1U), r.plain) << "row " << r.row + 1U;
    EXPECT_EQ(FormulaAt(wb, r.row, 2U), r.preserved) << "row " << r.row + 1U;
  }
}

// Re-encoding the text read back writes Excel's own tokens. Array constants
// are left out: their `PtgArray` carries bytes Excel leaves uninitialised.
TEST(XlsbFormulaWhitespace, EncodesExcelTokens) {
  const std::vector<std::uint8_t> package = ReadFixture();
  const Workbook wb = Load(package);
  const ExcelTokens excel = ReadExcelTokens(package);
  const std::vector<std::string> sheets = {wb.sheet(0).name()};
  const NameTable names;
  std::size_t checked = 0;
  for (const Row& r : kExcelText) {
    for (std::uint32_t col = 1U; col <= 2U; ++col) {
      const Cell* cell = wb.sheet(0).cell_at(r.row, col);
      ASSERT_NE(cell, nullptr) << "row " << r.row + 1U;
      const std::string body = Body(cell->formula_text);
      if (body.find('{') != std::string::npos) {
        continue;
      }
      Arena arena;
      parser::Parser p(body, arena);
      const parser::AstNode* root = p.parse();
      ASSERT_NE(root, nullptr) << body;
      auto encoded = encode_ptgs(*root, sheets, excel.sheet_ranges, names, PtgRootClass::kValue, std::nullopt,
                                 cell->dynamic_array ? PtgEvaluation::kDynamicArray : PtgEvaluation::kLegacy,
                                 name_shapes(wb, 0), body);
      ASSERT_TRUE(static_cast<bool>(encoded)) << body;
      const auto it = excel.cells.find({r.row, col});
      ASSERT_NE(it, excel.cells.end()) << body;
      EXPECT_EQ(ClearUnused(encoded.value().rgce), ClearUnused(it->second)) << body;
      ++checked;
    }
  }
  EXPECT_EQ(checked, 72U);
}

TEST(XlsbFormulaWhitespace, WriteReadKeepsText) {
  const Workbook wb = Load(ReadFixture());
  auto bytes = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  const Workbook back = Load(bytes.value());
  for (const Row& r : kExcelText) {
    if (IsArrayConstantRow(r)) {
      continue;
    }
    EXPECT_EQ(FormulaAt(back, r.row, 1U), r.plain) << "row " << r.row + 1U;
    EXPECT_EQ(FormulaAt(back, r.row, 2U), r.preserved) << "row " << r.row + 1U;
  }
}

std::vector<std::uint8_t> Encode(const std::string& body) {
  Arena arena;
  parser::Parser p(body, arena);
  const parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << body;
  if (root == nullptr) {
    return {};
  }
  auto encoded = encode_ptgs(*root, {}, SheetRangeTable{}, NameTable{}, PtgRootClass::kValue, std::nullopt,
                             PtgEvaluation::kDynamicArray, NameShapes(), body);
  EXPECT_TRUE(static_cast<bool>(encoded)) << body;
  return encoded ? encoded.value().rgce : std::vector<std::uint8_t>{};
}

bool HasAttrSpace(const std::vector<std::uint8_t>& rgce) {
  for (std::size_t i = 0; i + 1U < rgce.size(); ++i) {
    if (rgce[i] == 0x19 && rgce[i + 1U] == 0x40) {
      return true;
    }
  }
  return false;
}

// Whitespace Excel drops on reading, or in a position Excel's tokens do not
// carry it, is not stored, and every such formula encodes as its spaceless
// spelling does.
TEST(XlsbFormulaWhitespace, DropsWhitespaceExcelDoesNotStore) {
  const std::pair<const char*, const char*> cases[] = {
      {"SUM(1 ,2)", "SUM(1,2)"},   {"1+2 ", "1+2"},   {"1+2\n", "1+2"},
      {"1+ (2)", "1+(2)"},         {"1+\t2", "1+2"},  {"1+ \n2", "1+2"},
      {"SUM(1,2 ,)", "SUM(1,2,)"}, {"- (5)", "-(5)"}, {"1+ SUM(2)", "1+SUM(2)"},
  };
  for (const auto& [spaced, bare] : cases) {
    const std::vector<std::uint8_t> rgce = Encode(spaced);
    EXPECT_FALSE(HasAttrSpace(rgce)) << spaced;
    EXPECT_EQ(rgce, Encode(bare)) << spaced;
  }
}

// A space between `-` and a number keeps the minus as an operator token, as
// Excel stores it; without one the sign folds into the number.
TEST(XlsbFormulaWhitespace, SpacedMinusStaysAnOperator) {
  EXPECT_EQ(Encode("- 5"), (std::vector<std::uint8_t>{0x19, 0x40, 0x00, 0x01, 0x1E, 0x05, 0x00, 0x13}));
  EXPECT_EQ(Encode(" -5"), (std::vector<std::uint8_t>{0x19, 0x40, 0x00, 0x01, 0x1F, 0, 0, 0, 0, 0, 0, 0x14, 0xC0}));
}

std::string Decode(const std::vector<std::uint8_t>& rgce) {
  Arena arena;
  parser::FormulaWhitespace whitespace;
  auto ast =
      decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, ByteSpan{}, arena, {}, {}, {}, {}, -1, std::nullopt, &whitespace);
  EXPECT_TRUE(static_cast<bool>(ast));
  return ast ? parser::format_formula(*ast.value(), &whitespace) : std::string();
}

// A whitespace type or placement outside the ones Excel writes for formula
// text is skipped rather than guessed at.
TEST(XlsbFormulaWhitespace, SkipsUnknownPlacements) {
  // Spaces before an open parenthesis (type 2), and two attrs before one token.
  EXPECT_EQ(Decode({0x19, 0x40, 0x02, 0x01, 0x1E, 0x01, 0x00, 0x15}), "(1)");
  EXPECT_EQ(Decode({0x19, 0x40, 0x00, 0x01, 0x19, 0x40, 0x01, 0x01, 0x1E, 0x01, 0x00}), "1");
  // Spaces before an operand ahead of a unary minus token.
  EXPECT_EQ(Decode({0x1E, 0x05, 0x00, 0x19, 0x40, 0x00, 0x01, 0x13}), "-5");
  // A control attr between the space and its token is transparent.
  EXPECT_EQ(Decode({0x1E, 0x01, 0x00, 0x19, 0x40, 0x04, 0x02, 0x19, 0x08, 0x03, 0x00, 0x42, 0x01, 0x04, 0x00}),
            "SUM(1  )");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
