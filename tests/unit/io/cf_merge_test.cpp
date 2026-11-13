//
// Conditional-format rules merged at load. The fixture cases compare the
// rules Formulon holds after loading each authored input with the rules
// Excel 16.115 (Mac) wrote on re-saving it, transcribed below; every input
// is loaded twice, from `<case>.xlsx` and from `<case>.xlsb`, a pre-merge
// Formulon save of it. The model-level tests pin each merge condition on
// workbooks built in memory.

#include "io/cf_merge.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <regex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/styles_writer.h"
#include "io/xlsb/cf_records.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "sheet_passthrough.h"
#include "styles.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io {
namespace {

using Sqref = std::vector<cf::CFCellRange>;

// ---------------------------------------------------------------------------
// Shared helpers

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

std::string SqrefText(const Sqref& sqref) {
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

/// `text` as space-separated A1 areas, e.g. "H1:H3 J1".
Sqref ParseSqref(std::string_view text) {
  Sqref out;
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

const char* TypeName(cf::RuleType type) {
  static constexpr std::array<const char*, 18> kNames = {
      "expression",        "cellIs",         "colorScale",        "dataBar",    "iconSet",         "top10",
      "aboveAverage",      "containsText",   "notContainsText",   "beginsWith", "endsWith",        "containsBlanks",
      "notContainsBlanks", "containsErrors", "notContainsErrors", "timePeriod", "duplicateValues", "uniqueValues"};
  return kNames[static_cast<std::size_t>(type)];
}

/// RRGGBB of the fill of `rule`'s dxf (its foreground, else its background);
/// empty when the rule names no dxf.
std::string DxfFill(const Workbook& wb, const cf::CFRule& rule) {
  if (!rule.dxf_id.has_value()) {
    return std::string();
  }
  if (*rule.dxf_id >= wb.styles().dxfs.size()) {
    return "out of range";
  }
  static const std::regex kFg("<fgColor rgb=\"[0-9A-Fa-f]{2}([0-9A-Fa-f]{6})\"");
  static const std::regex kBg("<bgColor rgb=\"[0-9A-Fa-f]{2}([0-9A-Fa-f]{6})\"");
  const std::string body = dxf_body_fragment(wb.styles().dxfs[*rule.dxf_id]);
  std::smatch m;
  if (std::regex_search(body, m, kFg) || std::regex_search(body, m, kBg)) {
    return m[1].str();
  }
  return "?";
}

struct FlatRule {
  const Sqref* sqref;
  const cf::CFRule* rule;
};

/// Every rule of `sheet` with its block's sqref, by priority (model order on ties).
std::vector<FlatRule> RulesByPriority(const Sheet& sheet) {
  std::vector<FlatRule> out;
  for (const cf::ConditionalFormat& format : sheet.conditional_formats()) {
    for (const cf::CFRule& rule : format.rules) {
      out.push_back(FlatRule{&format.sqref, &rule});
    }
  }
  std::stable_sort(out.begin(), out.end(),
                   [](const FlatRule& a, const FlatRule& b) { return a.rule->priority < b.rule->priority; });
  return out;
}

/// "p<priority> <sqref>" for every rule of `sheet`, by priority.
std::vector<std::string> Summary(const Sheet& sheet) {
  std::vector<std::string> out;
  for (const FlatRule& flat : RulesByPriority(sheet)) {
    out.push_back("p" + std::to_string(flat.rule->priority) + " " + SqrefText(*flat.sqref));
  }
  return out;
}

std::size_t RuleCount(const Sheet& sheet) {
  std::size_t count = 0;
  for (const cf::ConditionalFormat& format : sheet.conditional_formats()) {
    count += format.rules.size();
  }
  return count;
}

std::string FixturePath(const char* name, const char* extension) {
  return std::string(FORMULON_FIXTURES_DIR) + "/cf_merge/" + name + extension;
}

std::optional<Workbook> LoadXlsx(const char* name) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath(name, ".xlsx"));
  auto read = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  if (!read) {
    return std::nullopt;
  }
  return std::move(read.value().workbook);
}

std::optional<Workbook> LoadXlsb(const char* name) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath(name, ".xlsb"));
  auto read = xlsb::read_xlsb(ByteSpan{bytes.data(), bytes.size()});
  if (!read) {
    return std::nullopt;
  }
  return std::move(read.value().workbook);
}

// ---------------------------------------------------------------------------
// Fixture cases: tests/fixtures/cf_merge/<case>.{xlsx,xlsb}, Sheet3

/// One rule of Excel's re-save. `formula1` / `formula2` are nullptr when
/// absent; `fill` is the dxf fill RRGGBB, empty for a rule with no dxf.
struct ExpectedRule {
  const char* sqref;
  std::int32_t priority;
  const char* type;
  const char* formula1;
  const char* formula2;
  bool stop;
  const char* fill;
};

struct MergeCase {
  // From ranges: GCC 14 -Werror=stringop-overflow misjudges a vector built from a one-element braced list.
  MergeCase(const char* case_name, std::initializer_list<ExpectedRule> case_rules,
            std::initializer_list<ExpectedRule> case_xlsb_rules = {})
      : name(case_name),
        rules(case_rules.begin(), case_rules.end()),
        xlsb_rules(case_xlsb_rules.begin(), case_xlsb_rules.end()) {}

  const char* name;
  std::vector<ExpectedRule> rules;
  /// The rules the xlsb load holds when it differs: the xlsb loader keeps a
  /// pivot-scoped block as raw records, which also keeps the sheet unmerged.
  std::vector<ExpectedRule> xlsb_rules;
};

constexpr std::size_t kCfSheet = 2;

const std::vector<MergeCase>& Cases() {
  static const std::vector<MergeCase> cases = {
      {"fp_chain_out_of_order", {{"H1:J3", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"fp_rect_after_merge", {{"H1:I3", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"fp_rel_formula_chain", {{"H1:J3", 1, "expression", "H1>0", nullptr, false, "CC2222"}}},
      {"fp_between_overlaps_earlier",
       {{"H1:H3", 2, "expression", "TRUE", nullptr, false, "22CC22"},
        {"H1:I3", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"fp_between_overlaps_later",
       {{"H1:H3", 1, "expression", "TRUE", nullptr, false, "CC2222"},
        {"I1:I3", 2, "expression", "TRUE", nullptr, false, "22CC22"},
        {"I1:I3", 3, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"ord_mixed_pairs",
       {{"H1:H3", 1, "expression", "TRUE", nullptr, false, "CC2222"},
        {"H1:I3", 2, "expression", "TRUE", nullptr, false, "22CC22"},
        {"I1:I3", 4, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"ord_earlier_low_prio",
       {{"H1:I3", 1, "expression", "TRUE", nullptr, false, "CC2222"},
        {"I1:I3", 2, "expression", "TRUE", nullptr, false, "22CC22"}}},
      {"ord_first_block_tail",
       {{"H1:H3", 2, "expression", "TRUE", nullptr, false, "22CC22"},
        {"H1:I3", 1, "expression", "TRUE", nullptr, false, "CC2222"},
        {"K1:K3", 4, "expression", "TRUE", nullptr, false, "2222CC"}}},
      {"ord_merged_block_place",
       {{"H1:I3", 1, "expression", "TRUE", nullptr, false, "CC2222"},
        {"K1:K3", 2, "expression", "TRUE", nullptr, false, "22CC22"}}},
      {"nd2_text", {{"H1:I3", 1, "containsText", "NOT(ISERROR(SEARCH(\"a\",H1)))", nullptr, false, "CC2222"}}},
      {"nd3_notcontains", {{"H1:I3", 1, "notContainsText", "ISERROR(SEARCH(\"a\",H1))", nullptr, false, "CC2222"}}},
      {"nd3_begins", {{"H1:I3", 1, "beginsWith", "LEFT(H1,LEN(\"a\"))=\"a\"", nullptr, false, "CC2222"}}},
      {"nd3_ends", {{"H1:I3", 1, "endsWith", "RIGHT(H1,LEN(\"a\"))=\"a\"", nullptr, false, "CC2222"}}},
      {"nd3_notblanks", {{"H1:I3", 1, "notContainsBlanks", "LEN(TRIM(H1))>0", nullptr, false, "CC2222"}}},
      {"nd3_errors", {{"H1:I3", 1, "containsErrors", "ISERROR(H1)", nullptr, false, "CC2222"}}},
      {"nd3_noterrors", {{"H1:I3", 1, "notContainsErrors", "NOT(ISERROR(H1))", nullptr, false, "CC2222"}}},
      {"nd2_top10_same_sqref",
       {{"H1:H3", 1, "top10", nullptr, nullptr, false, "CC2222"},
        {"H1:H3", 2, "top10", nullptr, nullptr, false, "CC2222"}}},
      {"nd_scale_same_sqref",
       {{"H1:H3", 1, "colorScale", nullptr, nullptr, false, ""},
        {"H1:H3", 2, "colorScale", nullptr, nullptr, false, ""}}},
      {"nd3_unique_same_sqref",
       {{"H1:H3", 1, "uniqueValues", nullptr, nullptr, false, "CC2222"},
        {"H1:H3", 2, "uniqueValues", nullptr, nullptr, false, "CC2222"}}},
      {"nd3_above_avg_same_sqref",
       {{"H1:H3", 1, "aboveAverage", nullptr, nullptr, false, "CC2222"},
        {"H1:H3", 2, "aboveAverage", nullptr, nullptr, false, "CC2222"}}},
      {"nd3_dup_same_sqref",
       {{"H1:H3", 1, "duplicateValues", nullptr, nullptr, false, "CC2222"},
        {"H1:H3", 2, "duplicateValues", nullptr, nullptr, false, "CC2222"}}},
      {"nd3_bar_same_sqref",
       {{"H1:H3", 1, "dataBar", nullptr, nullptr, false, ""}, {"H1:H3", 2, "dataBar", nullptr, nullptr, false, ""}}},
      {"nd3_icon_same_sqref",
       {{"H1:H3", 1, "iconSet", nullptr, nullptr, false, ""}, {"H1:H3", 2, "iconSet", nullptr, nullptr, false, ""}}},
      {"x14_expr_adjacent", {{"H1:I3", 1, "expression", "Sheet1!$A$1>0", nullptr, false, "CC2222"}}},
      {"lit_case",
       {{"H1:H3", 1, "expression", "H1=\"a\"", nullptr, false, "CC2222"},
        {"I1:I3", 2, "expression", "I1=\"A\"", nullptr, false, "CC2222"}}},
      {"eq_priority", {{"H1:I3", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"multi_union_rect", {{"H1:J3", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"multi_order_swapped", {{"H1:H3 J1:J3", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"stop_differs",
       {{"H1:H3", 1, "expression", "TRUE", nullptr, false, "CC2222"},
        {"I1:I3", 2, "expression", "TRUE", nullptr, true, "CC2222"}}},
      {"abs_rel_mixed", {{"H1:I3", 1, "expression", "$A1>0", nullptr, false, "CC2222"}}},
      {"no_dxf", {{"H1:I3", 1, "expression", "TRUE", nullptr, false, ""}}},
      {"pivot_block",
       {{"H1:H3", 1, "expression", "TRUE", nullptr, false, "CC2222"},
        {"I1:I3", 2, "expression", "TRUE", nullptr, false, "CC2222"}},
       {{"I1:I3", 2, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"tie_same_sqref_two_blocks",
       {{"H1:H3", 1, "expression", "TRUE", nullptr, false, "CC2222"},
        {"H1:H3", 2, "expression", "TRUE", nullptr, false, "22CC22"}}},
      {"x14_vs_legacy_adjacent",
       {{"H1:H3", 1, "expression", "$A$1>0", nullptr, false, "CC2222"},
        {"I1:I3", 2, "expression", "$A$1>0", nullptr, false, "CC2222"}}},
      {"x14_vs_legacy_adjacent_eq", {{"H1:I3", 1, "expression", "$A$1>0", nullptr, false, "CC2222"}}},
      {"same_block_content_equal", {{"H1:I3", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"two_blocks_overlap", {{"H1:H4", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"same_block_formula_case", {{"H1:I3", 1, "expression", "A1=1", nullptr, false, "CC2222"}}},
      {"containment", {{"H1:H5", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"containment_rev", {{"H1:H5", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"split_diff", {{"H1:H3 J1:J3", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
      {"four_areas", {{"H1:K1", 1, "expression", "TRUE", nullptr, false, "CC2222"}}},
  };
  return cases;
}

void ExpectRules(const Workbook& wb, const std::vector<ExpectedRule>& expected_in_saved_order, const char* format) {
  ASSERT_GT(wb.sheet_count(), kCfSheet) << format;
  const std::vector<FlatRule> rules = RulesByPriority(wb.sheet(kCfSheet));
  ASSERT_EQ(rules.size(), expected_in_saved_order.size()) << format;
  // The table keeps Excel's saved block order; rules compare in priority order.
  std::vector<ExpectedRule> expected = expected_in_saved_order;
  std::stable_sort(expected.begin(), expected.end(),
                   [](const ExpectedRule& lhs, const ExpectedRule& rhs) { return lhs.priority < rhs.priority; });
  const auto text = [](const std::optional<std::string>& formula) {
    return formula.has_value() ? *formula : std::string("<none>");
  };
  for (std::size_t i = 0; i < rules.size(); ++i) {
    const cf::CFRule& rule = *rules[i].rule;
    const ExpectedRule& want = expected[i];
    EXPECT_EQ(SqrefText(*rules[i].sqref), want.sqref) << format << " rule " << i;
    EXPECT_EQ(rule.priority, want.priority) << format << " rule " << i;
    EXPECT_STREQ(TypeName(rule.type), want.type) << format << " rule " << i;
    EXPECT_EQ(text(rule.formula1), want.formula1 != nullptr ? want.formula1 : "<none>") << format << " rule " << i;
    EXPECT_EQ(text(rule.formula2), want.formula2 != nullptr ? want.formula2 : "<none>") << format << " rule " << i;
    EXPECT_EQ(rule.stop_if_true, want.stop) << format << " rule " << i;
    EXPECT_EQ(DxfFill(wb, rule), want.fill) << format << " rule " << i;
  }
}

class CfMergeFixture : public ::testing::TestWithParam<MergeCase> {};

TEST_P(CfMergeFixture, XlsxLoadHoldsExcelsRules) {
  const std::optional<Workbook> wb = LoadXlsx(GetParam().name);
  ASSERT_TRUE(wb.has_value());
  ExpectRules(*wb, GetParam().rules, "xlsx");
}

TEST_P(CfMergeFixture, XlsbLoadHoldsExcelsRules) {
  const std::optional<Workbook> wb = LoadXlsb(GetParam().name);
  ASSERT_TRUE(wb.has_value());
  ExpectRules(*wb, GetParam().xlsb_rules.empty() ? GetParam().rules : GetParam().xlsb_rules, "xlsb");
}

INSTANTIATE_TEST_SUITE_P(Cases, CfMergeFixture, ::testing::ValuesIn(Cases()),
                         [](const ::testing::TestParamInfo<MergeCase>& param_info) {
                           return std::string(param_info.param.name);
                         });

std::size_t CountRecords(const std::vector<std::uint8_t>& slot, std::uint16_t type) {
  std::size_t count = 0;
  ByteSpan cursor{slot.data(), slot.size()};
  while (cursor.size != 0U) {
    auto rec = xlsb::read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << "tail does not frame";
      return count;
    }
    count += rec.value().type == type ? 1U : 0U;
  }
  return count;
}

// The xlsb load keeps both x14 rules' records retained; the save must drop the
// folded rule's and write the merged rule once.
TEST(CfMergeXlsbRoundTrip, MergedX14RuleIsWrittenOnce) {
  const std::optional<Workbook> wb = LoadXlsb("x14_expr_adjacent");
  ASSERT_TRUE(wb.has_value());
  ASSERT_EQ(Summary(wb->sheet(kCfSheet)), (std::vector<std::string>{"p1 H1:I3"}));
  auto bytes = xlsb::write_xlsb(*wb);
  ASSERT_TRUE(static_cast<bool>(bytes));
  auto reread = xlsb::read_xlsb(ByteSpan{bytes.value().data(), bytes.value().size()});
  ASSERT_TRUE(static_cast<bool>(reread));
  const Sheet& sheet = reread.value().workbook.sheet(kCfSheet);
  EXPECT_EQ(Summary(sheet), (std::vector<std::string>{"p1 H1:I3"}));
  std::size_t rule_records = 0;
  for (const std::vector<std::uint8_t>* slot :
       {&sheet.xlsb_tail().before_merges, &sheet.xlsb_tail().after_merges_before_hyperlinks,
        &sheet.xlsb_tail().after_hyperlinks}) {
    rule_records += CountRecords(*slot, xlsb::kBeginCfRule14);
  }
  EXPECT_EQ(rule_records, 1U);
}

// ---------------------------------------------------------------------------
// Model-level pins

constexpr std::uint32_t kRed = 0;
constexpr std::uint32_t kGreen = 1;
constexpr std::uint32_t kRedCopy = 2;

DifferentialFormat SolidDxf(std::uint32_t rgb) {
  DifferentialFormat dxf;
  dxf.has_fill = true;
  dxf.fill.pattern = 1;
  dxf.fill.fg_argb = 0xFF000000U | rgb;
  dxf.fill.bg_argb = 0xFF000000U | rgb;
  dxf.fill.fg.kind = ColorSpec::Kind::kRgb;
  dxf.fill.fg.rgb = 0xFF000000U | rgb;
  dxf.fill.bg.kind = ColorSpec::Kind::kRgb;
  dxf.fill.bg.rgb = 0xFF000000U | rgb;
  return dxf;
}

/// A one-sheet workbook whose dxfs are red, green, and red again.
Workbook Book() {
  Workbook wb = Workbook::create();
  wb.mutable_styles().dxfs = {SolidDxf(0xCC2222U), SolidDxf(0x22CC22U), SolidDxf(0xCC2222U)};
  return wb;
}

cf::CFRule Rule(std::int32_t priority, std::optional<std::uint32_t> dxf = kRed, std::string formula = "TRUE") {
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = priority;
  rule.dxf_id = dxf;
  rule.formula1 = std::move(formula);
  return rule;
}

void AddBlock(Workbook& wb, std::string_view sqref, std::vector<cf::CFRule> rules) {
  cf::ConditionalFormat format;
  format.sqref = ParseSqref(sqref);
  format.rules = std::move(rules);
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(format));
}

/// Summary of sheet 0 after the load-time merge.
std::vector<std::string> Merged(Workbook& wb) {
  merge_cf_rules_on_load(wb);
  return Summary(wb.sheet(0));
}

/// Summary after merging `first` on H1:H3 (p1) and `second` on I1:I3 (p2).
std::vector<std::string> MergedPair(cf::CFRule first, cf::CFRule second) {
  Workbook wb = Book();
  first.priority = 1;
  second.priority = 2;
  AddBlock(wb, "H1:H3", {std::move(first)});
  AddBlock(wb, "I1:I3", {std::move(second)});
  return Merged(wb);
}

using Rules = std::vector<std::string>;

TEST(CfMergeModel, AdjacentRectanglesMerge) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(1)});
  AddBlock(wb, "I1:I3", {Rule(2)});
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:I3"}));
}

TEST(CfMergeModel, OverlapAndContainmentMerge) {
  Workbook overlap = Book();
  AddBlock(overlap, "H1:H3", {Rule(1)});
  AddBlock(overlap, "H2:H4", {Rule(2)});
  EXPECT_EQ(Merged(overlap), (Rules{"p1 H1:H4"}));

  Workbook contained = Book();
  AddBlock(contained, "H2:H3", {Rule(1)});
  AddBlock(contained, "H1:H5", {Rule(2)});
  EXPECT_EQ(Merged(contained), (Rules{"p1 H1:H5"}));
}

TEST(CfMergeModel, MultiAreaUnionThatIsOneRectangleMerges) {
  Workbook wb = Book();
  AddBlock(wb, "H1 J1", {Rule(1)});
  AddBlock(wb, "I1 K1", {Rule(2)});
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:K1"}));
}

TEST(CfMergeModel, EqualCellSetsKeepTheSurvivorsAreaOrder) {
  Workbook swapped = Book();
  AddBlock(swapped, "H1:H3 J1:J3", {Rule(1)});
  AddBlock(swapped, "J1:J3 H1:H3", {Rule(2)});
  EXPECT_EQ(Merged(swapped), (Rules{"p1 H1:H3 J1:J3"}));

  Workbook split = Book();
  AddBlock(split, "H1:H3 J1:J3", {Rule(1)});
  AddBlock(split, "H1 H2:H3 J1:J3", {Rule(2)});
  EXPECT_EQ(Merged(split), (Rules{"p1 H1:H3 J1:J3"}));
}

TEST(CfMergeModel, NonRectangularUnionsStay) {
  Workbook gap = Book();
  AddBlock(gap, "H1:H3", {Rule(1)});
  AddBlock(gap, "J1:J3", {Rule(2)});
  EXPECT_EQ(Merged(gap), (Rules{"p1 H1:H3", "p2 J1:J3"}));

  Workbook staggered = Book();
  AddBlock(staggered, "H1:H3", {Rule(1)});
  AddBlock(staggered, "I2:I4", {Rule(2)});
  EXPECT_EQ(Merged(staggered), (Rules{"p1 H1:H3", "p2 I2:I4"}));
}

// An intermediate rule over the lower-numbered rule's range does not block.
TEST(CfMergeModel, IntermediateOverTheSurvivorDoesNotBlock) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(1, kRed)});
  AddBlock(wb, "H1:H3", {Rule(2, kGreen)});
  AddBlock(wb, "I1:I3", {Rule(3, kRed)});
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:I3", "p2 H1:H3"}));
}

// An intermediate rule over the higher-numbered rule's range blocks.
TEST(CfMergeModel, IntermediateOverTheFoldedRuleBlocks) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(1, kRed)});
  AddBlock(wb, "I1:I3", {Rule(2, kGreen)});
  AddBlock(wb, "I1:I3", {Rule(3, kRed)});
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:H3", "p2 I1:I3", "p3 I1:I3"}));
}

// The same two ranges with the priorities swapped: now the intermediate rule
// covers the folded rule's range.
TEST(CfMergeModel, IntermediateTestIsAsymmetric) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(3, kRed)});
  AddBlock(wb, "H1:H3", {Rule(2, kGreen)});
  AddBlock(wb, "I1:I3", {Rule(1, kRed)});
  EXPECT_EQ(Merged(wb), (Rules{"p1 I1:I3", "p2 H1:H3", "p3 H1:H3"}));
}

TEST(CfMergeModel, DisjointIntermediateDoesNotBlock) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(1, kRed)});
  AddBlock(wb, "K1:K3", {Rule(2, kGreen)});
  AddBlock(wb, "I1:I3", {Rule(3, kRed)});
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:I3", "p2 K1:K3"}));
}

TEST(CfMergeModel, ChainOutOfOrderReachesTheFixedPoint) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(1)});
  AddBlock(wb, "J1:J3", {Rule(2)});
  AddBlock(wb, "I1:I3", {Rule(3)});
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:J3"}));
}

TEST(CfMergeModel, RectangleFormedByAMergeMergesAgain) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(1)});
  AddBlock(wb, "I1:I2", {Rule(2)});
  AddBlock(wb, "I3", {Rule(3)});
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:I3"}));
}

TEST(CfMergeModel, SurvivorLeavesAMultiRuleBlock) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(1, kRed), Rule(2, kGreen)});
  AddBlock(wb, "I1:I3", {Rule(3, kRed)});
  merge_cf_rules_on_load(wb);
  const std::vector<cf::ConditionalFormat>& formats = wb.sheet(0).conditional_formats();
  ASSERT_EQ(formats.size(), 2U);
  EXPECT_EQ(SqrefText(formats[0].sqref), "H1:H3");
  ASSERT_EQ(formats[0].rules.size(), 1U);
  EXPECT_EQ(formats[0].rules[0].priority, 2);
  EXPECT_EQ(SqrefText(formats[1].sqref), "H1:I3");
  ASSERT_EQ(formats[1].rules.size(), 1U);
  EXPECT_EQ(formats[1].rules[0].priority, 1);
}

TEST(CfMergeModel, KeptFormulaIsTheTextAtTheNewAnchor) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(3, kRed, "H1>0")});
  AddBlock(wb, "I1:I3", {Rule(1, kRed, "I1>0")});
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:I3"}));
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].formula1, std::optional<std::string>("H1>0"));
}

TEST(CfMergeModel, EqualPrioritiesKeepTheEarlierRule) {
  Workbook wb = Book();
  cf::CFRule first = Rule(1);
  first.id = "first";
  AddBlock(wb, "I1:I3", {std::move(first)});
  AddBlock(wb, "H1:H3", {Rule(1)});
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:I3"}));
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].id, "first");
}

TEST(CfMergeModel, DefinitionsMustMatch) {
  cf::CFRule stop = Rule(0);
  stop.stop_if_true = true;
  EXPECT_EQ(MergedPair(Rule(0), stop).size(), 2U);
  EXPECT_EQ(MergedPair(Rule(0, kRed, "H1=\"a\""), Rule(0, kRed, "I1=\"A\"")).size(), 2U);
  EXPECT_EQ(MergedPair(Rule(0, kRed, "H1=\"a\""), Rule(0, kRed, "I1=\"a\"")).size(), 1U);
  EXPECT_EQ(MergedPair(Rule(0, kRed, "H1>0"), Rule(0, kRed, "H1>0")).size(), 2U);
  EXPECT_EQ(MergedPair(Rule(0, kRed, "$A$1>0"), Rule(0, kRed, "$A$1>0")).size(), 1U);
  EXPECT_EQ(MergedPair(Rule(0, kRed), Rule(0, kGreen)).size(), 2U);
  EXPECT_EQ(MergedPair(Rule(0, kRed), Rule(0, kRedCopy)).size(), 1U);
  EXPECT_EQ(MergedPair(Rule(0, std::nullopt), Rule(0, std::nullopt)).size(), 1U);
  EXPECT_EQ(MergedPair(Rule(0, kRed), Rule(0, std::nullopt)).size(), 2U);

  cf::CFRule greater = Rule(0, kRed, "1");
  greater.type = cf::RuleType::CellIs;
  greater.op = cf::CellIsOperator::GreaterThan;
  cf::CFRule less = greater;
  less.op = cf::CellIsOperator::LessThan;
  EXPECT_EQ(MergedPair(greater, greater).size(), 1U);
  EXPECT_EQ(MergedPair(greater, less).size(), 2U);
}

TEST(CfMergeModel, NeverMergingTypesStay) {
  for (const cf::RuleType type :
       {cf::RuleType::Top10, cf::RuleType::AboveAverage, cf::RuleType::UniqueValues, cf::RuleType::DuplicateValues,
        cf::RuleType::ColorScale, cf::RuleType::DataBar, cf::RuleType::IconSet}) {
    Workbook wb = Book();
    cf::CFRule first = Rule(1);
    first.type = type;
    first.formula1.reset();
    cf::CFRule second = first;
    second.priority = 2;
    AddBlock(wb, "H1:H3", {std::move(first)});
    AddBlock(wb, "H1:H3", {std::move(second)});
    EXPECT_EQ(Merged(wb).size(), 2U) << TypeName(type);
  }
}

TEST(CfMergeModel, PivotBlocksAndRawExtensionsStay) {
  Workbook pivot = Book();
  AddBlock(pivot, "H1:H3", {Rule(1)});
  AddBlock(pivot, "I1:I3", {Rule(2)});
  pivot.sheet(0).mutable_conditional_formats()[0].pivot_scope = true;
  EXPECT_EQ(Merged(pivot).size(), 2U);

  Workbook block_ext = Book();
  AddBlock(block_ext, "H1:H3", {Rule(1)});
  AddBlock(block_ext, "I1:I3", {Rule(2)});
  block_ext.sheet(0).mutable_conditional_formats()[1].ext_lst_raw = std::string("<extLst/>");
  EXPECT_EQ(Merged(block_ext).size(), 2U);

  cf::CFRule linked = Rule(2);
  linked.ext_lst_raw = std::string("<extLst/>");
  EXPECT_EQ(MergedPair(Rule(1), linked).size(), 2U);
}

TEST(CfMergeModel, NonCandidatesStay) {
  EXPECT_EQ(MergedPair(Rule(0, 99U), Rule(0, 99U)).size(), 2U);
  EXPECT_EQ(MergedPair(Rule(0, kRed, "SUM(1,"), Rule(0, kRed, "SUM(1,")).size(), 2U);

  Workbook empty = Book();
  AddBlock(empty, "", {Rule(1)});
  AddBlock(empty, "", {Rule(2)});
  merge_cf_rules_on_load(empty);
  EXPECT_EQ(RuleCount(empty.sheet(0)), 2U);
}

// ---------------------------------------------------------------------------
// Sheet gate: CF the model does not hold keeps the whole sheet as authored.

constexpr const char* kUnheldId = "{00000000-0000-0000-0000-0000000000AA}";

std::string X14Extension(const std::string& rule_id) {
  return "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
         "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
         "<x14:conditionalFormattings><x14:conditionalFormatting "
         "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
         "<x14:cfRule type=\"iconSet\" priority=\"3\" id=\"" +
         rule_id +
         "\"/><xm:sqref>I1:I3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";
}

/// H1:H3 p1 + I1:I3 p2, mergeable, and a data bar on K1:K3 p3 with id `id`.
Workbook GatedBook(const std::string& id) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(1)});
  AddBlock(wb, "I1:I3", {Rule(2)});
  cf::CFRule bar = Rule(3, std::nullopt);
  bar.type = cf::RuleType::DataBar;
  bar.formula1.reset();
  bar.id = id;
  AddBlock(wb, "K1:K3", {std::move(bar)});
  return wb;
}

TEST(CfMergeGate, XlsxExtensionRuleTheModelLacksKeepsTheSheet) {
  Workbook wb = GatedBook("{00000000-0000-0000-0000-000000000001}");
  wb.sheet(0).set_ext_lst_xml(X14Extension(kUnheldId));
  EXPECT_EQ(Merged(wb).size(), 3U);
}

TEST(CfMergeGate, XlsxExtensionLinkedToAModelRuleDoesNotGate) {
  Workbook wb = GatedBook(kUnheldId);
  wb.sheet(0).set_ext_lst_xml(X14Extension(kUnheldId));
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:I3", "p3 K1:K3"}));
}

TEST(CfMergeGate, XlsbRetainedLegacyBlockKeepsTheSheet) {
  Workbook wb = GatedBook(std::string());
  XlsbSheetTail tail;
  xlsb::emit_record(tail.after_merges_before_hyperlinks, xlsb::kBrtBeginConditionalFormatting,
                    std::vector<std::uint8_t>(28, 0U));
  xlsb::emit_record(tail.after_merges_before_hyperlinks, xlsb::kBrtEndConditionalFormatting,
                    std::vector<std::uint8_t>());
  wb.sheet(0).set_xlsb_tail(std::move(tail));
  EXPECT_EQ(Merged(wb).size(), 3U);
}

/// A tail holding one BrtBeginCFRule14 whose GUID is `id`.
XlsbSheetTail X14RuleTail(const std::string& id) {
  std::array<std::uint8_t, xlsb::kGuidBytes> guid{};
  EXPECT_TRUE(xlsb::ParseGuid(id, guid));
  std::vector<std::uint8_t> rule(70, 0U);
  std::copy(guid.begin(), guid.end(), rule.begin() + 50);
  rule[66] = 1U;
  XlsbSheetTail tail;
  xlsb::emit_record(tail.after_hyperlinks, xlsb::kBeginCfRule14, rule);
  return tail;
}

TEST(CfMergeGate, XlsbRetainedX14RuleTheModelLacksKeepsTheSheet) {
  Workbook wb = GatedBook("{00000000-0000-0000-0000-000000000001}");
  wb.sheet(0).set_xlsb_tail(X14RuleTail(kUnheldId));
  EXPECT_EQ(Merged(wb).size(), 3U);
}

TEST(CfMergeGate, XlsbRetainedX14RuleOfAModelRuleDoesNotGate) {
  Workbook wb = GatedBook(kUnheldId);
  wb.sheet(0).set_xlsb_tail(X14RuleTail(kUnheldId));
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:I3", "p3 K1:K3"}));
}

// ---------------------------------------------------------------------------
// Only the loaders merge.

TEST(CfMergeApi, AddedRulesAreNotMerged) {
  Workbook wb = Book();
  AddBlock(wb, "H1:H3", {Rule(1)});
  AddBlock(wb, "I1:I3", {Rule(2)});
  EXPECT_EQ(RuleCount(wb.sheet(0)), 2U);
  auto bytes = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes));
  EXPECT_EQ(RuleCount(wb.sheet(0)), 2U);
  auto reread = read_ooxml(ByteSpan{bytes.value().data(), bytes.value().size()});
  ASSERT_TRUE(static_cast<bool>(reread));
  EXPECT_EQ(Summary(reread.value().workbook.sheet(0)), (Rules{"p1 H1:I3"}));
}

// ---------------------------------------------------------------------------
// Cost: 2,000 single-cell rules down one column, odd rows taking the low
// priorities, so every merge has the whole interleaved run between its pair.

TEST(CfMergePerf, InterleavedSingleCellsMergeToOneRule) {
  constexpr std::uint32_t kRows = 2000;
  Workbook wb = Book();
  std::vector<cf::ConditionalFormat>& formats = wb.sheet(0).mutable_conditional_formats();
  for (std::uint32_t row = 0; row < kRows; ++row) {
    const std::uint32_t ordinal = row / 2U + 1U;
    cf::ConditionalFormat format;
    format.sqref.push_back(cf::CFCellRange{{row, 7U}, {row, 7U}});
    format.rules.push_back(Rule(static_cast<std::int32_t>(row % 2U == 0U ? ordinal : kRows / 2U + ordinal)));
    formats.push_back(std::move(format));
  }
  EXPECT_EQ(Merged(wb), (Rules{"p1 H1:H2000"}));
}

}  // namespace
}  // namespace formulon::io
