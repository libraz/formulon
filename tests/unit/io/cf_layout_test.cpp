//
// The conditional-format blocks a save writes. Each fixture case loads an
// authored input and saves it; the sequence of blocks (sqref text and rule
// priorities in order) must equal the one Excel 16.115 (Mac) wrote on
// re-saving the same input, transcribed below. The key tests pin the block
// order and grouping on models built in memory.

#include "io/cf_layout.h"

#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "miniz.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io {
namespace {

constexpr std::size_t kCfSheet = 2;

/// One saved block: its sqref text and the priorities of its rules in order.
struct Block {
  std::string sqref;
  std::vector<std::int32_t> priorities;

  friend bool operator==(const Block& a, const Block& b) { return a.sqref == b.sqref && a.priorities == b.priorities; }
};

struct LayoutCase {
  const char* name;
  std::vector<Block> blocks;
};

/// Excel's saved blocks per fixture case, in Excel's order.
const std::vector<LayoutCase>& Cases() {
  static const std::vector<LayoutCase> cases = {
      {"ord_no_merge_doc_vs_prio", {{"H1:H3", {1}}, {"J1:J3", {2}}}},
      {"ord_no_merge_left_low_prio_later", {{"H1:H3", {2}}, {"J1:J3", {1}}}},
      {"ord_pos_row_vs_col_a", {{"H5:H6", {1}}, {"I1:I2", {2}}}},
      {"ord_pos_row_vs_col_b", {{"H5:H6", {2}}, {"I1:I2", {1}}}},
      {"ord_pos_same_col", {{"H1:H2", {2}}, {"H5:H6", {1}}}},
      {"ord_pos_multi_area", {{"I1", {2}}, {"J1:J3 H5", {1}}}},
      {"ord_pos_three", {{"H9", {2}}, {"I1", {3}}, {"K2", {1}}}},
      {"key_h10_vs_h9", {{"H9", {2}}, {"H10", {1}}}},
      {"key_aa1_vs_b1", {{"B1", {2}}, {"AA1", {1}}}},
      {"key_wide_vs_tall", {{"H1:H3", {2}}, {"H1:J1", {1}}}},
      {"key_short_wide_vs_long_tall", {{"H1:H5", {2}}, {"H1:I1", {1}}}},
      {"tie_big_first_low_prio", {{"H1:H3", {2}}, {"H1:H5", {1}}}},
      {"tie_small_first_low_prio", {{"H1:H3", {1}}, {"H1:H5", {2}}}},
      {"tie_same_sqref_two_blocks", {{"H1:H3", {1, 2}}}},
      {"tie_same_sqref_two_blocks_rev", {{"H1:H3", {1, 2}}}},
      {"ord_mixed_pairs", {{"H1:H3", {1}}, {"H1:I3", {2}}, {"I1:I3", {4}}}},
      {"ord_first_block_tail", {{"H1:H3", {2}}, {"H1:I3", {1}}, {"K1:K3", {4}}}},
  };
  return cases;
}

std::string ColumnText(std::uint32_t col) {
  std::string out;
  for (std::uint32_t n = col + 1U; n != 0U; n = (n - 1U) / 26U) {
    out.insert(out.begin(), static_cast<char>('A' + (n - 1U) % 26U));
  }
  return out;
}

std::string CellText(CellAddress cell) {
  return ColumnText(cell.col) + std::to_string(cell.row + 1U);
}

std::string SqrefText(const std::vector<cf::CFCellRange>& sqref) {
  std::string out;
  for (const cf::CFCellRange& range : sqref) {
    if (!out.empty()) {
      out.push_back(' ');
    }
    out += CellText(range.first);
    if (range.first != range.last) {
      out += ":" + CellText(range.last);
    }
  }
  return out;
}

CellAddress ParseCell(std::string_view a1) {
  std::uint32_t col = 0;
  std::size_t i = 0;
  for (; i < a1.size() && a1[i] >= 'A' && a1[i] <= 'Z'; ++i) {
    col = col * 26U + static_cast<std::uint32_t>(a1[i] - 'A' + 1);
  }
  std::uint32_t row = 0;
  for (; i < a1.size(); ++i) {
    row = row * 10U + static_cast<std::uint32_t>(a1[i] - '0');
  }
  return CellAddress{row - 1U, col - 1U};
}

std::vector<cf::CFCellRange> ParseSqref(std::string_view text) {
  std::vector<cf::CFCellRange> out;
  while (!text.empty()) {
    const std::size_t space = text.find(' ');
    const std::string_view area = text.substr(0, space);
    const std::size_t colon = area.find(':');
    const CellAddress first = ParseCell(area.substr(0, colon));
    out.push_back(cf::CFCellRange{first, colon == std::string_view::npos ? first : ParseCell(area.substr(colon + 1U))});
    text = space == std::string_view::npos ? std::string_view() : text.substr(space + 1U);
  }
  return out;
}

std::string Describe(const std::vector<Block>& blocks) {
  std::string out;
  for (const Block& b : blocks) {
    out += "[" + b.sqref + " p";
    for (const std::int32_t p : b.priorities) {
      out += " " + std::to_string(p);
    }
    out += "] ";
  }
  return out;
}

/// The blocks of `sheet`'s model, in model order.
std::vector<Block> ModelBlocks(const Sheet& sheet) {
  std::vector<Block> out;
  for (const cf::ConditionalFormat& format : sheet.conditional_formats()) {
    Block b{SqrefText(format.sqref), {}};
    for (const cf::CFRule& rule : format.rules) {
      b.priorities.push_back(rule.priority);
    }
    out.push_back(std::move(b));
  }
  return out;
}

/// The `<conditionalFormatting>` blocks of an xlsx worksheet part.
std::vector<Block> XmlBlocks(const std::string& xml) {
  std::vector<Block> out;
  const std::string open = "<conditionalFormatting sqref=\"";
  for (std::size_t at = xml.find(open); at != std::string::npos; at = xml.find(open, at + 1U)) {
    const std::size_t sqref_begin = at + open.size();
    const std::size_t sqref_end = xml.find('"', sqref_begin);
    const std::size_t block_end = xml.find("</conditionalFormatting>", sqref_end);
    Block b{xml.substr(sqref_begin, sqref_end - sqref_begin), {}};
    const std::string key = "<cfRule ";
    for (std::size_t r = xml.find(key, sqref_end); r != std::string::npos && r < block_end; r = xml.find(key, r + 1U)) {
      const std::size_t p = xml.find("priority=\"", r) + 10U;
      b.priorities.push_back(std::stoi(xml.substr(p, xml.find('"', p) - p)));
    }
    out.push_back(std::move(b));
  }
  return out;
}

std::string SheetPart(const std::vector<std::uint8_t>& bytes, const char* name) {
  std::string out;
  mz_zip_archive zip{};
  if (mz_zip_reader_init_mem(&zip, bytes.data(), bytes.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return out;
  }
  std::size_t size = 0;
  void* body = mz_zip_reader_extract_file_to_heap(&zip, name, &size, 0);
  if (body != nullptr) {
    out.assign(static_cast<const char*>(body), size);
    mz_free(body);
  } else {
    ADD_FAILURE() << "missing part " << name;
  }
  mz_zip_reader_end(&zip);
  return out;
}

Workbook LoadCase(const char* name) {
  const std::vector<std::uint8_t> bytes =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/cf_layout/" + name + ".xlsx");
  auto read = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << name;
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

TEST(CfSavedLayout, XlsxSaveWritesExcelsBlockSequence) {
  for (const LayoutCase& c : Cases()) {
    const Workbook wb = LoadCase(c.name);
    ASSERT_GT(wb.sheet_count(), kCfSheet) << c.name;
    auto bytes = write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(bytes)) << c.name;
    const std::vector<Block> saved = XmlBlocks(SheetPart(bytes.value(), "xl/worksheets/sheet3.xml"));
    EXPECT_EQ(saved, c.blocks) << c.name << ": got " << Describe(saved) << "expected " << Describe(c.blocks);
  }
}

TEST(CfSavedLayout, XlsbSaveReloadsInExcelsBlockSequence) {
  for (const LayoutCase& c : Cases()) {
    const Workbook wb = LoadCase(c.name);
    auto bytes = xlsb::write_xlsb(wb);
    ASSERT_TRUE(static_cast<bool>(bytes)) << c.name;
    auto reread = xlsb::read_xlsb(ByteSpan{bytes.value().data(), bytes.value().size()});
    ASSERT_TRUE(static_cast<bool>(reread)) << c.name;
    ASSERT_GT(reread.value().workbook.sheet_count(), kCfSheet) << c.name;
    const std::vector<Block> saved = ModelBlocks(reread.value().workbook.sheet(kCfSheet));
    EXPECT_EQ(saved, c.blocks) << c.name << ": got " << Describe(saved) << "expected " << Describe(c.blocks);
  }
}

// ---------------------------------------------------------------------------
// The layout function on models built in memory.

cf::ConditionalFormat MakeBlock(std::string_view sqref, std::vector<std::int32_t> priorities) {
  cf::ConditionalFormat block;
  block.sqref = ParseSqref(sqref);
  for (const std::int32_t priority : priorities) {
    cf::CFRule rule;
    rule.type = cf::RuleType::Expression;
    rule.priority = priority;
    rule.formula1 = "TRUE";
    block.rules.push_back(std::move(rule));
  }
  return block;
}

std::vector<Block> Layout(const std::vector<cf::ConditionalFormat>& formats) {
  std::vector<Block> out;
  for (const CfBlockView& view : cf_saved_layout(formats)) {
    Block b{SqrefText(*view.sqref), {}};
    for (const cf::CFRule* rule : view.rules) {
      b.priorities.push_back(rule->priority);
    }
    out.push_back(std::move(b));
  }
  return out;
}

TEST(CfSavedLayoutKey, ColumnFirstThenRow) {
  EXPECT_EQ(Layout({MakeBlock("H10", {1}), MakeBlock("H9", {2})}), (std::vector<Block>{{"H9", {2}}, {"H10", {1}}}));
}

TEST(CfSavedLayoutKey, ColumnOrderIsNumericNotAlphabetic) {
  EXPECT_EQ(Layout({MakeBlock("AA1", {1}), MakeBlock("B1", {2})}), (std::vector<Block>{{"B1", {2}}, {"AA1", {1}}}));
}

TEST(CfSavedLayoutKey, EqualFirstCellOrdersByLastColumnThenLastRow) {
  EXPECT_EQ(Layout({MakeBlock("H1:J1", {1}), MakeBlock("H1:H3", {2})}),
            (std::vector<Block>{{"H1:H3", {2}}, {"H1:J1", {1}}}));
  EXPECT_EQ(Layout({MakeBlock("H1:I1", {1}), MakeBlock("H1:H5", {2})}),
            (std::vector<Block>{{"H1:H5", {2}}, {"H1:I1", {1}}}));
}

TEST(CfSavedLayoutKey, MultiAreaBlockIsKeyedByItsFirstArea) {
  EXPECT_EQ(Layout({MakeBlock("J1:J3 H5", {1}), MakeBlock("I1", {2})}),
            (std::vector<Block>{{"I1", {2}}, {"J1:J3 H5", {1}}}));
}

TEST(CfSavedLayoutKey, EqualKeysKeepModelOrder) {
  EXPECT_EQ(Layout({MakeBlock("H1:H3 K1", {5}), MakeBlock("H1:H3 A1", {1})}),
            (std::vector<Block>{{"H1:H3 K1", {5}}, {"H1:H3 A1", {1}}}));
}

TEST(CfSavedLayoutKey, BlocksOfOneSqrefCombineWithRulesByPriority) {
  EXPECT_EQ(Layout({MakeBlock("H1:H3", {2, 5}), MakeBlock("J1", {3}), MakeBlock("H1:H3", {1, 5})}),
            (std::vector<Block>{{"H1:H3", {1, 2, 5, 5}}, {"J1", {3}}}));
}

TEST(CfSavedLayoutKey, EqualPrioritiesKeepModelOrderAcrossCombinedBlocks) {
  std::vector<cf::ConditionalFormat> formats = {MakeBlock("H1", {4}), MakeBlock("H1", {4})};
  formats[0].rules[0].formula1 = "first";
  formats[1].rules[0].formula1 = "second";
  const std::vector<CfBlockView> views = cf_saved_layout(formats);
  ASSERT_EQ(views.size(), 1U);
  ASSERT_EQ(views[0].rules.size(), 2U);
  EXPECT_EQ(*views[0].rules[0]->formula1, "first");
  EXPECT_EQ(*views[0].rules[1]->formula1, "second");
}

TEST(CfSavedLayoutKey, SameAreasInADifferentOrderDoNotCombine) {
  EXPECT_EQ(Layout({MakeBlock("H1 J1", {1}), MakeBlock("J1 H1", {2})}),
            (std::vector<Block>{{"H1 J1", {1}}, {"J1 H1", {2}}}));
}

TEST(CfSavedLayoutKey, PivotScopedBlocksCombineOnlyWithPivotScopedBlocks) {
  std::vector<cf::ConditionalFormat> formats = {MakeBlock("H1", {1}), MakeBlock("H1", {2}), MakeBlock("H1", {3})};
  formats[1].pivot_scope = true;
  const std::vector<CfBlockView> views = cf_saved_layout(formats);
  ASSERT_EQ(views.size(), 2U);
  EXPECT_FALSE(views[0].pivot_scope);
  EXPECT_EQ(views[0].rules.size(), 2U);
  EXPECT_TRUE(views[1].pivot_scope);
  EXPECT_EQ(views[1].rules.size(), 1U);
}

TEST(CfSavedLayoutKey, BlocksWithExtLstStayApart) {
  std::vector<cf::ConditionalFormat> formats = {MakeBlock("H1", {1}), MakeBlock("H1", {2}), MakeBlock("H1", {3}),
                                                MakeBlock("H1", {4})};
  formats[1].ext_lst_raw = "<extLst/>";
  formats[3].ext_lst_raw = "<extLst/>";
  const std::vector<CfBlockView> views = cf_saved_layout(formats);
  ASSERT_EQ(views.size(), 3U);
  EXPECT_EQ(Layout(formats), (std::vector<Block>{{"H1", {1, 3}}, {"H1", {2}}, {"H1", {4}}}));
  EXPECT_FALSE(views[0].ext_lst_raw->has_value());
  EXPECT_TRUE(views[1].ext_lst_raw->has_value());
}

}  // namespace
}  // namespace formulon::io
