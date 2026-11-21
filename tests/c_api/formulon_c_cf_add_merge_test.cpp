// Conditional-format rules added through the C ABI, as a save writes them.
//
// Each case adds the rules of an authored book one by one (each add appends a
// block of its own), saves, and compares the written blocks, sqref text, rule
// types, priorities and dxf fills with Excel's own re-save of that book
// (tests/fixtures/excel/cf_add_merge/<row>_sit.xlsx). The model keeps the rules
// as added; only the written form is merged.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <regex>
#include <string>
#include <vector>

#include "formulon_c_cf_test_helpers.h"
#include "io/ooxml_reader.h"
#include "io/styles_writer.h"
#include "io/xlsb/cf_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "support/roundtrip_symmetry.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace {

using formulon::io::ByteSpan;

constexpr std::uint32_t kRed = 0;
constexpr std::uint32_t kGreen = 1;

struct Add {
  const char* sqref;  // One column, "H1:H3".
  std::int32_t priority;
  std::uint32_t dxf;
};

struct Row {
  const char* name;
  std::vector<Add> adds;
};

const std::vector<Row>& Rows() {
  static const std::vector<Row> rows = {
      {"o00", {{"H1:H3", 1, kRed}, {"I1:I3", 2, kRed}}},
      {"o01", {{"H1:H3", 1, kRed}, {"H1:H3", 2, kRed}}},
      {"o04", {{"H1:H3", 1, kRed}, {"H1:H3", 2, kGreen}, {"H1:H3", 3, kRed}}},
      {"o10", {{"H1:H3", 2, kRed}, {"I1:I3", 1, kRed}}},
      {"o11", {{"H1:H3", 1, kRed}, {"H1:H3", 2, kGreen}, {"I1:I3", 3, kRed}}},
  };
  return rows;
}

formulon::DifferentialFormat SolidDxf(std::uint32_t rgb) {
  formulon::DifferentialFormat dxf;
  dxf.has_fill = true;
  dxf.fill.pattern = 1;
  dxf.fill.fg_argb = 0xFF000000U | rgb;
  dxf.fill.bg_argb = 0xFF000000U | rgb;
  dxf.fill.fg.kind = formulon::ColorSpec::Kind::kRgb;
  dxf.fill.fg.rgb = 0xFF000000U | rgb;
  dxf.fill.bg.kind = formulon::ColorSpec::Kind::kRgb;
  dxf.fill.bg.rgb = 0xFF000000U | rgb;
  return dxf;
}

fm_cell_rect_t RectOf(const std::string& a1) {
  const auto col = static_cast<std::uint32_t>(a1[0] - 'A');
  const std::size_t colon = a1.find(':');
  const auto first = static_cast<std::uint32_t>(std::stoul(a1.substr(1, colon - 1U))) - 1U;
  const auto last = static_cast<std::uint32_t>(std::stoul(a1.substr(colon + 2U))) - 1U;
  return fm_cell_rect_t{first, col, last, col};
}

// A workbook holding the row's rules, added through the C ABI in order.
WorkbookGuard Build(const Row& row) {
  WorkbookGuard wb;
  EXPECT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().dxfs = {SolidDxf(0xCC2222U), SolidDxf(0x22CC22U)};
  for (const Add& add : row.adds) {
    fm_cell_rect_t range = RectOf(add.sqref);
    fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
    rule.type = FM_CF_RULE_TYPE_EXPRESSION;
    rule.formula1 = "TRUE";
    rule.priority = add.priority;
    rule.has_dxf_index = 1;
    rule.dxf_index = add.dxf;
    rule.ranges = &range;
    rule.range_count = 1;
    const char* id = nullptr;
    EXPECT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  }
  return wb;
}

std::vector<std::uint8_t> Save(fm_workbook_t* handle, int32_t format) {
  BufferGuard buf;
  EXPECT_EQ(fm_workbook_save(handle, format, &buf.data, &buf.len), 0) << fm_get_last_error_message();
  return buf.data == nullptr ? std::vector<std::uint8_t>() : std::vector<std::uint8_t>(buf.data, buf.data + buf.len);
}

std::string ColumnText(std::uint32_t col) {
  std::string out;
  for (std::uint32_t n = col + 1U; n != 0U; n = (n - 1U) / 26U) {
    out.insert(out.begin(), static_cast<char>('A' + (n - 1U) % 26U));
  }
  return out;
}

std::string SqrefText(const std::vector<formulon::cf::CFCellRange>& sqref) {
  std::string out;
  for (const formulon::cf::CFCellRange& range : sqref) {
    if (!out.empty()) {
      out.push_back(' ');
    }
    out += ColumnText(range.first.col) + std::to_string(range.first.row + 1U);
    if (range.first != range.last) {
      out += ":" + ColumnText(range.last.col) + std::to_string(range.last.row + 1U);
    }
  }
  return out;
}

// RRGGBB of the fill of dxf `id`; empty when the rule names none.
std::string FillOf(const formulon::Workbook& wb, const std::optional<std::uint32_t>& id) {
  if (!id.has_value()) {
    return std::string();
  }
  if (*id >= wb.styles().dxfs.size()) {
    return "out of range";
  }
  static const std::regex kFg("<fgColor rgb=\"[0-9A-Fa-f]{2}([0-9A-Fa-f]{6})\"");
  static const std::regex kBg("<bgColor rgb=\"[0-9A-Fa-f]{2}([0-9A-Fa-f]{6})\"");
  const std::string body = formulon::io::dxf_body_fragment(wb.styles().dxfs[*id]);
  std::smatch m;
  return std::regex_search(body, m, kFg) || std::regex_search(body, m, kBg) ? m[1].str() : std::string("?");
}

// "sqref{p<priority> <type> <fill>,...}" per block, joined by '|', in written order.
struct Structure {
  std::string text;

  void AddBlock(const std::string& sqref, const std::vector<std::string>& rules) {
    text += (text.empty() ? "" : "|") + sqref + "{";
    for (std::size_t i = 0; i < rules.size(); ++i) {
      text += (i == 0U ? "" : ",") + rules[i];
    }
    text += "}";
  }
};

std::string Attr(const std::string& tag, const char* name) {
  const std::regex pattern(std::string("\\s") + name + "=\"([^\"]*)\"");
  std::smatch m;
  return std::regex_search(tag, m, pattern) ? m[1].str() : std::string();
}

// The blocks of sheet1.xml in written order; fills resolve through the package's own dxf table.
std::string XlsxStructure(const std::vector<std::uint8_t>& bytes) {
  std::string xml;
  EXPECT_TRUE(formulon::test::extract_part(ByteSpan{bytes.data(), bytes.size()}, "xl/worksheets/sheet1.xml", &xml));
  auto read = formulon::io::read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read));
  if (!read) {
    return std::string();
  }
  Structure out;
  const std::regex block_pattern("<conditionalFormatting ([^>]*)>([\\s\\S]*?)</conditionalFormatting>");
  const std::regex rule_pattern("<cfRule ([^>]*)>");
  for (std::sregex_iterator block(xml.begin(), xml.end(), block_pattern), end; block != end; ++block) {
    const std::string body = (*block)[2].str();
    std::vector<std::string> rules;
    for (std::sregex_iterator rule(body.begin(), body.end(), rule_pattern); rule != end; ++rule) {
      const std::string tag = " " + (*rule)[1].str();
      const std::string dxf = Attr(tag, "dxfId");
      rules.push_back(
          "p" + Attr(tag, "priority") + " " + Attr(tag, "type") + " " +
          FillOf(read.value().workbook, dxf.empty() ? std::nullopt : std::optional<std::uint32_t>(std::stoul(dxf))));
    }
    out.AddBlock(Attr(" " + (*block)[1].str(), "sqref"), rules);
  }
  return out.text;
}

const char* TypeName(formulon::cf::RuleType type) {
  return type == formulon::cf::RuleType::Expression ? "expression" : "other";
}

// The blocks of sheet1.bin in written order, decoded from the records themselves.
std::string XlsbStructure(const std::vector<std::uint8_t>& bytes) {
  std::string sheet;
  EXPECT_TRUE(formulon::test::extract_part(ByteSpan{bytes.data(), bytes.size()}, "xl/worksheets/sheet1.bin", &sheet));
  auto read = formulon::io::xlsb::read_xlsb(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read));
  if (!read) {
    return std::string();
  }
  const std::vector<std::string> names = {"Sheet1"};
  const std::vector<formulon::io::xlsb::XlsbName> name_table;
  const std::vector<formulon::io::xlsb::XlsbSheetRange> ranges;
  const formulon::io::xlsb::XlsbExternalBooks books;
  const formulon::io::xlsb::FeatureFormulaReadContext ctx{names, name_table, ranges, books, 0};

  Structure out;
  ByteSpan cursor{reinterpret_cast<const std::uint8_t*>(sheet.data()), sheet.size()};
  const std::uint8_t* block_start = nullptr;
  while (cursor.size != 0U) {
    const std::uint8_t* record_start = cursor.data;
    auto record = formulon::io::xlsb::read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "sheet stream does not frame";
      break;
    }
    if (record.value().type == formulon::io::xlsb::kBrtBeginConditionalFormatting) {
      block_start = record_start;
    } else if (record.value().type == formulon::io::xlsb::kBrtEndConditionalFormatting && block_start != nullptr) {
      const auto decoded = formulon::io::xlsb::decode_cf_block(
          ByteSpan{block_start, static_cast<std::size_t>(cursor.data - block_start)}, ctx);
      EXPECT_TRUE(decoded.has_value());
      if (decoded.has_value()) {
        std::vector<std::string> rules;
        for (const formulon::cf::CFRule& rule : decoded->rules) {
          rules.push_back("p" + std::to_string(rule.priority) + " " + TypeName(rule.type) + " " +
                          FillOf(read.value().workbook, rule.dxf_id));
        }
        out.AddBlock(SqrefText(decoded->sqref), rules);
      }
      block_start = nullptr;
    }
  }
  return out.text;
}

std::vector<std::uint8_t> Fixture(const std::string& name) {
  return formulon::test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_add_merge/" + name + ".xlsx");
}

// The rules in the model as added: one block per add, in add order.
void ExpectModelAsAdded(const WorkbookGuard& wb, const Row& row) {
  const auto& blocks = wb.handle->workbook().sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), row.adds.size()) << row.name;
  for (std::size_t i = 0; i < blocks.size(); ++i) {
    ASSERT_EQ(blocks[i].rules.size(), 1U) << row.name;
    EXPECT_EQ(SqrefText(blocks[i].sqref), row.adds[i].sqref) << row.name;
    EXPECT_EQ(blocks[i].rules[0].priority, row.adds[i].priority) << row.name;
  }
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cf_rule_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, row.adds.size()) << row.name;
}

class CfAddMerge : public ::testing::TestWithParam<Row> {};

TEST_P(CfAddMerge, XlsxSaveHoldsExcelsBlocks) {
  const Row& row = GetParam();
  const WorkbookGuard wb = Build(row);
  const std::string excel = XlsxStructure(Fixture(std::string("cfa_") + row.name + "_sit"));
  ASSERT_FALSE(excel.empty());
  EXPECT_EQ(XlsxStructure(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX)), excel) << row.name;
  ExpectModelAsAdded(wb, row);
}

TEST_P(CfAddMerge, XlsbSaveHoldsExcelsBlocks) {
  const Row& row = GetParam();
  const WorkbookGuard wb = Build(row);
  const std::string excel = XlsxStructure(Fixture(std::string("cfa_") + row.name + "_sit"));
  ASSERT_FALSE(excel.empty());
  EXPECT_EQ(XlsbStructure(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB)), excel) << row.name;
  ExpectModelAsAdded(wb, row);
}

// A loaded workbook saved again writes the same blocks.
TEST_P(CfAddMerge, SavingALoadedSaveChangesNothing) {
  const Row& row = GetParam();
  const WorkbookGuard wb = Build(row);
  for (const int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    const std::vector<std::uint8_t> first = Save(wb.handle, format);
    WorkbookGuard loaded;
    ASSERT_EQ(fm_workbook_load(first.data(), first.size(), nullptr, &loaded.handle), 0) << fm_get_last_error_message();
    const std::vector<std::uint8_t> second = Save(loaded.handle, format);
    if (format == FM_WORKBOOK_FORMAT_XLSX) {
      EXPECT_EQ(XlsxStructure(second), XlsxStructure(first)) << row.name;
    } else {
      EXPECT_EQ(XlsbStructure(second), XlsbStructure(first)) << row.name;
    }
  }
}

// The authored book is the unmerged shape: every added rule is in the package.
TEST_P(CfAddMerge, AuthoredBookHoldsEveryAddedRule) {
  const Row& row = GetParam();
  std::string xml;
  const std::vector<std::uint8_t> bytes = Fixture(std::string("cfa_") + row.name);
  ASSERT_TRUE(formulon::test::extract_part(ByteSpan{bytes.data(), bytes.size()}, "xl/worksheets/sheet1.xml", &xml));
  std::size_t rules = 0;
  for (std::size_t at = xml.find("<cfRule "); at != std::string::npos; at = xml.find("<cfRule ", at + 1U)) {
    ++rules;
  }
  EXPECT_EQ(rules, row.adds.size()) << row.name;
}

std::string RowName(const ::testing::TestParamInfo<Row>& info) {
  return std::string(info.param.name);
}

INSTANTIATE_TEST_SUITE_P(Rows, CfAddMerge, ::testing::ValuesIn(Rows()), RowName);

}  // namespace
