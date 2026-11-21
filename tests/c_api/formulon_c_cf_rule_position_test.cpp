// Stable C ABI conditional-format position-addressed mutators and priority placement.
//
// The fixture is an Excel-saved sheet holding two data-bar rules that share one id, on A1:A10 (priority 1) and
// B1:B10 (priority 2).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "formulon_c_cf_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace {

constexpr const char* kSharedId = "{00000000-000E-0000-0000-000001000000}";
constexpr const char* kStarsId = "{33333333-3333-3333-3333-333333333333}";
constexpr const char* kBoxesId = "{44444444-4444-4444-4444-444444444444}";
constexpr std::int32_t kMaxPriority = std::numeric_limits<std::int32_t>::max();

std::vector<std::uint8_t> ReadFixture() {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_x14_shared_id.xlsx";
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

WorkbookGuard LoadSharedIdFixture() {
  WorkbookGuard guard;
  const std::vector<std::uint8_t> bytes = ReadFixture();
  EXPECT_FALSE(bytes.empty());
  EXPECT_EQ(fm_workbook_load(bytes.data(), bytes.size(), nullptr, &guard.handle), 0) << fm_get_last_error_message();
  return guard;
}

// What a listed rule is, reduced to the fields these tests compare.
struct RuleView {
  std::string id;
  std::int32_t priority = 0;
  std::uint32_t first_col = 0;
  std::int32_t stop_if_true = 0;

  bool operator==(const RuleView& other) const {
    return id == other.id && priority == other.priority && first_col == other.first_col &&
           stop_if_true == other.stop_if_true;
  }
};

std::vector<RuleView> ListRules(const fm_workbook_t* handle) {
  std::vector<RuleView> out;
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_cf_rule_count(handle, 0, &count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
    EXPECT_EQ(fm_sheet_get_cf_rule_at(handle, 0, i, &rule), 0);
    RuleView view;
    view.id = rule.id;
    view.priority = rule.priority;
    view.first_col = rule.range_count > 0U ? rule.ranges[0].first_col : UINT32_MAX;
    view.stop_if_true = rule.stop_if_true;
    out.push_back(std::move(view));
  }
  return out;
}

std::vector<std::int32_t> Priorities(const fm_workbook_t* handle) {
  std::vector<std::int32_t> out;
  for (const RuleView& rule : ListRules(handle)) {
    out.push_back(rule.priority);
  }
  return out;
}

// The rule whose first range starts at `first_col`; the fixture's two rules sit in columns 0 and 1.
RuleView ByColumn(const fm_workbook_t* handle, std::uint32_t first_col) {
  for (const RuleView& rule : ListRules(handle)) {
    if (rule.first_col == first_col) {
      return rule;
    }
  }
  ADD_FAILURE() << "no rule starts at column " << first_col;
  return RuleView{};
}

fm_status_t SetRanges(fm_workbook_t* handle, std::size_t position, const fm_cell_rect_t& rect) {
  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.ranges_op = FM_PATCH_OP_SET;
  patch.ranges = &rect;
  patch.range_count = 1;
  return fm_sheet_update_cf_rule_at(handle, 0, position, &patch);
}

fm_status_t SetPriority(fm_workbook_t* handle, std::size_t position, std::int32_t priority) {
  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.priority_op = FM_PATCH_OP_SET;
  patch.priority = priority;
  return fm_sheet_update_cf_rule_at(handle, 0, position, &patch);
}

// Adds an expression rule on column D with the engine-assigned priority and returns its id.
std::string AddAutoRule(fm_workbook_t* handle, std::int32_t priority = 0) {
  fm_cell_rect_t sqref{0, 3, 9, 3};
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  rule.type = 0;  // Expression
  rule.formula1 = "TRUE";
  rule.ranges = &sqref;
  rule.range_count = 1;
  rule.priority = priority;
  const char* id = nullptr;
  EXPECT_EQ(fm_sheet_add_cf_rule(handle, 0, &rule, &id), 0) << fm_get_last_error_message();
  return id != nullptr ? id : "";
}

RuleView ById(const fm_workbook_t* handle, const std::string& id) {
  for (const RuleView& rule : ListRules(handle)) {
    if (rule.id == id) {
      return rule;
    }
  }
  ADD_FAILURE() << "no rule with id " << id;
  return RuleView{};
}

// Adds retained x14 rules (they never enter the model) to the sheet's extension list.
void RetainX14Rules(fm_workbook_t* handle, const std::string& rules) {
  auto& sheet = handle->workbook().sheet(0);
  std::string xml = sheet.ext_lst_xml();
  const std::size_t at = xml.find("<xm:sqref>");
  if (at != std::string::npos) {
    xml.insert(at, rules);
  } else {
    // The model owns the fixture's rules, so the stored list holds no x14 block to extend; add a whole one.
    const std::string block =
        "<ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
        "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
        "<x14:conditionalFormattings>"
        "<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">" +
        rules + "<xm:sqref>D1:D10</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>";
    const std::size_t close = xml.rfind("</extLst>");
    if (close != std::string::npos) {
      xml.insert(close, block);
    } else {
      xml += "<extLst>" + block + "</extLst>";
    }
  }
  sheet.set_ext_lst_xml(std::move(xml));
}

std::string StarsRule(const char* priority) {
  return std::string("<x14:cfRule type=\"iconSet\" priority=\"") + priority + "\" id=\"" + kStarsId +
         "\"><x14:iconSet iconSet=\"3Stars\"><x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo>"
         "</x14:iconSet></x14:cfRule>";
}

std::string BoxesRule(const char* priority) {
  return std::string("<x14:cfRule type=\"iconSet\" priority=\"") + priority + "\" id=\"" + kBoxesId +
         "\"><x14:iconSet iconSet=\"5Boxes\"><x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo>"
         "</x14:iconSet></x14:cfRule>";
}

// The start tag of the retained x14 rule with `id`; empty when absent.
std::string X14StartTag(const std::string& xml, const char* id) {
  const std::size_t at = xml.find(id);
  if (at == std::string::npos) {
    return std::string();
  }
  const std::size_t open = xml.rfind("<x14:cfRule", at);
  const std::size_t close = xml.find('>', at);
  if (open == std::string::npos || close == std::string::npos) {
    return std::string();
  }
  return xml.substr(open, close - open + 1U);
}

bool HasPriority(const std::string& start_tag, const char* priority) {
  return start_tag.find(std::string("priority=\"") + priority + "\"") != std::string::npos;
}

}  // namespace

TEST(CfRulePosition, FixtureListsTwoRulesSharingOneId) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::vector<RuleView> rules = ListRules(wb.handle);
  ASSERT_EQ(rules.size(), 2U);
  EXPECT_EQ(rules[0].id, kSharedId);
  EXPECT_EQ(rules[1].id, kSharedId);
  EXPECT_EQ(rules[0].priority, 1);
  EXPECT_EQ(rules[1].priority, 2);
  EXPECT_EQ(rules[0].first_col, 0U);
  EXPECT_EQ(rules[1].first_col, 1U);
}

TEST(CfRulePosition, UpdateAtSecondRuleLeavesFirstUntouched) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::vector<RuleView> before = ListRules(wb.handle);

  ASSERT_EQ(SetRanges(wb.handle, 1, fm_cell_rect_t{0, 5, 9, 5}), 0) << fm_get_last_error_message();

  const std::vector<RuleView> after = ListRules(wb.handle);
  ASSERT_EQ(after.size(), 2U);
  EXPECT_EQ(after[0], before[0]);
  EXPECT_EQ(after[1].first_col, 5U);
  EXPECT_EQ(after[1].id, kSharedId);
  EXPECT_EQ(after[1].priority, 2);
}

TEST(CfRulePosition, UpdateAtFirstRuleLeavesSecondUntouched) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::vector<RuleView> before = ListRules(wb.handle);

  ASSERT_EQ(SetRanges(wb.handle, 0, fm_cell_rect_t{0, 6, 9, 6}), 0) << fm_get_last_error_message();

  const std::vector<RuleView> after = ListRules(wb.handle);
  ASSERT_EQ(after.size(), 2U);
  EXPECT_EQ(after[0].first_col, 6U);
  EXPECT_EQ(after[1], before[1]);
}

TEST(CfRulePosition, RemoveAtFirstKeepsSurvivorPriorityAndSharedId) {
  WorkbookGuard wb = LoadSharedIdFixture();

  ASSERT_EQ(fm_sheet_remove_cf_rule_at(wb.handle, 0, 0), 0) << fm_get_last_error_message();

  const std::vector<RuleView> rules = ListRules(wb.handle);
  ASSERT_EQ(rules.size(), 1U);
  EXPECT_EQ(rules[0].id, kSharedId);
  EXPECT_EQ(rules[0].priority, 2);
  EXPECT_EQ(rules[0].first_col, 1U);
}

TEST(CfRulePosition, RemoveAtSecondKeepsFirstRule) {
  WorkbookGuard wb = LoadSharedIdFixture();

  ASSERT_EQ(fm_sheet_remove_cf_rule_at(wb.handle, 0, 1), 0) << fm_get_last_error_message();

  const std::vector<RuleView> rules = ListRules(wb.handle);
  ASSERT_EQ(rules.size(), 1U);
  EXPECT_EQ(rules[0].id, kSharedId);
  EXPECT_EQ(rules[0].priority, 1);
  EXPECT_EQ(rules[0].first_col, 0U);
}

TEST(CfRulePosition, IdAddressedCallsActOnTheFirstOfTheSharedRules) {
  WorkbookGuard wb = LoadSharedIdFixture();

  std::int32_t present = 0;
  fm_cf_rule_t got = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule(wb.handle, 0, kSharedId, &present, &got), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(got.priority, 1);
  EXPECT_EQ(got.ranges[0].first_col, 0U);

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.stop_if_true_op = FM_PATCH_OP_SET;
  patch.stop_if_true = 1;
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, kSharedId, &patch), 0) << fm_get_last_error_message();
  EXPECT_EQ(ByColumn(wb.handle, 0).stop_if_true, 1);
  EXPECT_EQ(ByColumn(wb.handle, 1).stop_if_true, 0);

  ASSERT_EQ(fm_sheet_remove_cf_rule(wb.handle, 0, kSharedId), 0);
  const std::vector<RuleView> rules = ListRules(wb.handle);
  ASSERT_EQ(rules.size(), 1U);
  EXPECT_EQ(rules[0].first_col, 1U);
  EXPECT_EQ(rules[0].priority, 2);
}

TEST(CfRulePosition, IdIsReachableAgainOnceTheFirstSharingRuleIsGone) {
  WorkbookGuard wb = LoadSharedIdFixture();
  ASSERT_EQ(fm_sheet_remove_cf_rule_at(wb.handle, 0, 0), 0);

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.stop_if_true_op = FM_PATCH_OP_SET;
  patch.stop_if_true = 1;
  ASSERT_EQ(fm_sheet_update_cf_rule(wb.handle, 0, kSharedId, &patch), 0);
  EXPECT_EQ(ByColumn(wb.handle, 1).stop_if_true, 1);
  ASSERT_EQ(fm_sheet_remove_cf_rule(wb.handle, 0, kSharedId), 0);
  EXPECT_TRUE(ListRules(wb.handle).empty());
}

// Every position-addressed mutator against every kind of bad argument: the status is the one the kind maps to and
// nothing changes.
TEST(CfRulePosition, MutatorsReportTheStatusOfEachBadArgumentAndChangeNothing) {
  enum Fn { kUpdate, kRemove, kPriority };
  enum Fault { kNullWorkbook, kNullPatch, kSheetPastEnd, kPositionAtCount, kPositionFarPast, kPositionMax };
  struct Case {
    Fn fn;
    Fault fault;
    fm_status_t want;
  };
  const Case cases[] = {
      {kUpdate, kNullWorkbook, FM_STATUS_BINDING_NULL_POINTER},
      {kUpdate, kNullPatch, FM_STATUS_BINDING_NULL_POINTER},
      {kUpdate, kSheetPastEnd, FM_STATUS_SHEET_INDEX_OUT_OF_RANGE},
      {kUpdate, kPositionAtCount, FM_STATUS_INDEX_OUT_OF_RANGE},
      {kUpdate, kPositionFarPast, FM_STATUS_INDEX_OUT_OF_RANGE},
      {kUpdate, kPositionMax, FM_STATUS_INDEX_OUT_OF_RANGE},
      {kRemove, kNullWorkbook, FM_STATUS_BINDING_NULL_POINTER},
      {kRemove, kSheetPastEnd, FM_STATUS_SHEET_INDEX_OUT_OF_RANGE},
      {kRemove, kPositionAtCount, FM_STATUS_INDEX_OUT_OF_RANGE},
      {kRemove, kPositionFarPast, FM_STATUS_INDEX_OUT_OF_RANGE},
      {kRemove, kPositionMax, FM_STATUS_INDEX_OUT_OF_RANGE},
      {kPriority, kNullWorkbook, FM_STATUS_BINDING_NULL_POINTER},
      {kPriority, kSheetPastEnd, FM_STATUS_SHEET_INDEX_OUT_OF_RANGE},
      {kPriority, kPositionAtCount, FM_STATUS_INDEX_OUT_OF_RANGE},
      {kPriority, kPositionFarPast, FM_STATUS_INDEX_OUT_OF_RANGE},
      {kPriority, kPositionMax, FM_STATUS_INDEX_OUT_OF_RANGE},
  };

  WorkbookGuard wb = LoadSharedIdFixture();
  const std::vector<RuleView> before = ListRules(wb.handle);
  const std::string ext_before = wb.handle->workbook().sheet(0).ext_lst_xml();
  ASSERT_EQ(before.size(), 2U);

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.stop_if_true_op = FM_PATCH_OP_SET;
  patch.stop_if_true = 1;

  for (const Case& c : cases) {
    SCOPED_TRACE(::testing::Message() << "fn " << c.fn << " fault " << c.fault);
    fm_workbook_t* handle = c.fault == kNullWorkbook ? nullptr : wb.handle;
    std::size_t sheet = c.fault == kSheetPastEnd ? 99U : 0U;
    std::size_t position = 0;
    switch (c.fault) {
      case kPositionAtCount:
        position = before.size();
        break;
      case kPositionFarPast:
        position = 5;
        break;
      case kPositionMax:
        position = SIZE_MAX;
        break;
      default:
        break;
    }
    const fm_cf_rule_patch_t* patch_arg = c.fault == kNullPatch ? nullptr : &patch;
    fm_status_t got = FM_STATUS_OK;
    switch (c.fn) {
      case kUpdate:
        got = fm_sheet_update_cf_rule_at(handle, sheet, position, patch_arg);
        break;
      case kRemove:
        got = fm_sheet_remove_cf_rule_at(handle, sheet, position);
        break;
      case kPriority:
        got = fm_sheet_set_cf_rule_priority_at(handle, sheet, position, 1);
        break;
    }
    EXPECT_EQ(got, c.want);
    EXPECT_EQ(ListRules(wb.handle), before);
    EXPECT_EQ(wb.handle->workbook().sheet(0).ext_lst_xml(), ext_before);
  }
}

TEST(CfRulePosition, UpdateAtRefusesAnInvalidPatchBeforeChangingAnything) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::vector<RuleView> before = ListRules(wb.handle);

  fm_cf_rule_patch_t patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  patch.stop_if_true_op = FM_PATCH_OP_SET;
  patch.stop_if_true = 1;
  patch.formula1_op = FM_PATCH_OP_SET;  // a data bar does not use formula1
  patch.formula1 = "TRUE";
  EXPECT_EQ(fm_sheet_update_cf_rule_at(wb.handle, 0, 1, &patch), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ListRules(wb.handle), before);
}

TEST(CfRulePosition, ExplicitDuplicatePriorityIsAcceptedOnAddAndUpdate) {
  WorkbookGuard wb = LoadSharedIdFixture();

  const std::string added = AddAutoRule(wb.handle, 2);
  EXPECT_EQ(ById(wb.handle, added).priority, 2);
  EXPECT_EQ(Priorities(wb.handle), (std::vector<std::int32_t>{1, 2, 2}));

  ASSERT_EQ(SetPriority(wb.handle, 1, 1), 0) << fm_get_last_error_message();
  EXPECT_EQ(Priorities(wb.handle), (std::vector<std::int32_t>{1, 1, 2}));
}

TEST(CfRulePosition, ListedRuleReAddedWithIdClearedKeepsItsFields) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::size_t count_before = ListRules(wb.handle).size();

  fm_cf_rule_t copy = FM_STRUCT_INIT(fm_cf_rule_t);
  ASSERT_EQ(fm_sheet_get_cf_rule_at(wb.handle, 0, 0, &copy), 0);
  copy.id = nullptr;
  const char* new_id = nullptr;
  ASSERT_EQ(fm_sheet_add_cf_rule(wb.handle, 0, &copy, &new_id), 0) << fm_get_last_error_message();
  ASSERT_NE(new_id, nullptr);
  const std::string added_id = new_id;

  const std::vector<RuleView> rules = ListRules(wb.handle);
  ASSERT_EQ(rules.size(), count_before + 1U);
  const RuleView added = ById(wb.handle, added_id);
  EXPECT_NE(added.id, kSharedId);
  EXPECT_EQ(added.priority, 1);
  EXPECT_EQ(added.first_col, 0U);
}

TEST(CfRulePosition, SetPriorityOneShiftsEveryLoadedRuleUp) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::string added = AddAutoRule(wb.handle);
  ASSERT_EQ(ById(wb.handle, added).priority, 3);

  const std::size_t position = ListRules(wb.handle).size() - 1U;
  ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, position, 1), 0) << fm_get_last_error_message();

  EXPECT_EQ(ById(wb.handle, added).priority, 1);
  EXPECT_EQ(ByColumn(wb.handle, 0).priority, 2);
  EXPECT_EQ(ByColumn(wb.handle, 1).priority, 3);
}

TEST(CfRulePosition, SetPriorityKeepsTiesTogetherAndInOrder) {
  WorkbookGuard wb = LoadSharedIdFixture();
  ASSERT_EQ(SetPriority(wb.handle, 1, 1), 0);  // both fixture rules at 1
  const std::string added = AddAutoRule(wb.handle);
  ASSERT_EQ(ById(wb.handle, added).priority, 2);

  const std::size_t position = ListRules(wb.handle).size() - 1U;
  ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, position, 1), 0);

  EXPECT_EQ(ById(wb.handle, added).priority, 1);
  EXPECT_EQ(ByColumn(wb.handle, 0).priority, 2);
  EXPECT_EQ(ByColumn(wb.handle, 1).priority, 2);
  const std::vector<RuleView> rules = ListRules(wb.handle);
  EXPECT_EQ(rules[0].first_col, 0U);
  EXPECT_EQ(rules[1].first_col, 1U);
}

TEST(CfRulePosition, SetPriorityLeavesRulesBelowItAndDoesNotRenumberTheOldSlot) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::string added = AddAutoRule(wb.handle);  // priority 3
  const std::size_t position = ListRules(wb.handle).size() - 1U;

  ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, position, 2), 0);

  EXPECT_EQ(ByColumn(wb.handle, 0).priority, 1);
  EXPECT_EQ(ById(wb.handle, added).priority, 2);
  EXPECT_EQ(ByColumn(wb.handle, 1).priority, 3);

  // Moving the rule up again leaves the gap it left: nothing is closed up.
  ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, position, 9), 0);
  EXPECT_EQ(ById(wb.handle, added).priority, 9);
  EXPECT_EQ(ByColumn(wb.handle, 1).priority, 3);
}

TEST(CfRulePosition, SetPriorityToTheCurrentValueChangesNothing) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::vector<RuleView> before = ListRules(wb.handle);
  const std::string ext_before = wb.handle->workbook().sheet(0).ext_lst_xml();

  for (std::size_t i = 0; i < before.size(); ++i) {
    ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, i, before[i].priority), 0);
    EXPECT_EQ(ListRules(wb.handle), before);
  }
  EXPECT_EQ(wb.handle->workbook().sheet(0).ext_lst_xml(), ext_before);
}

TEST(CfRulePosition, SetPriorityAboveTheMaximumShiftsNothing) {
  WorkbookGuard wb = LoadSharedIdFixture();

  ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 0, 10), 0);

  EXPECT_EQ(ByColumn(wb.handle, 0).priority, 10);
  EXPECT_EQ(ByColumn(wb.handle, 1).priority, 2);
}

TEST(CfRulePosition, SetPriorityBelowOneIsRefusedAndChangesNothing) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::vector<RuleView> before = ListRules(wb.handle);

  for (const std::int32_t bad : {0, -1, std::numeric_limits<std::int32_t>::min()}) {
    SCOPED_TRACE(bad);
    EXPECT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 1, bad), FM_STATUS_INVALID_ARGUMENT);
    EXPECT_EQ(ListRules(wb.handle), before);
  }
}

TEST(CfRulePosition, SetPriorityThatWouldOverflowIsRefusedAndChangesNothing) {
  WorkbookGuard wb = LoadSharedIdFixture();
  ASSERT_EQ(SetPriority(wb.handle, 1, kMaxPriority), 0);
  const std::vector<RuleView> before = ListRules(wb.handle);
  ASSERT_EQ(Priorities(wb.handle), (std::vector<std::int32_t>{1, kMaxPriority}));

  EXPECT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 0, kMaxPriority), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 0, 2), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ListRules(wb.handle), before);
}

TEST(CfRulePosition, SetPriorityShiftsRetainedExtensionRulesAtOrAboveIt) {
  WorkbookGuard wb = LoadSharedIdFixture();
  RetainX14Rules(wb.handle, StarsRule("4") + BoxesRule("1"));

  ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 0, 3), 0) << fm_get_last_error_message();

  EXPECT_EQ(ByColumn(wb.handle, 0).priority, 3);
  EXPECT_EQ(ByColumn(wb.handle, 1).priority, 2);
  const std::string& xml = wb.handle->workbook().sheet(0).ext_lst_xml();
  EXPECT_TRUE(HasPriority(X14StartTag(xml, kStarsId), "5")) << xml;
  EXPECT_TRUE(HasPriority(X14StartTag(xml, kBoxesId), "1")) << xml;
}

TEST(CfRulePosition, SetPriorityLeavesRetainedRulesByteIdenticalWhenNothingShifts) {
  WorkbookGuard wb = LoadSharedIdFixture();
  RetainX14Rules(wb.handle, StarsRule("4") + BoxesRule("1"));
  const std::string before = wb.handle->workbook().sheet(0).ext_lst_xml();

  ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 0, 20), 0);

  EXPECT_EQ(ByColumn(wb.handle, 0).priority, 20);
  EXPECT_EQ(wb.handle->workbook().sheet(0).ext_lst_xml(), before);
}

TEST(CfRulePosition, SavedPrioritiesSurviveReload) {
  WorkbookGuard wb = LoadSharedIdFixture();
  RetainX14Rules(wb.handle, StarsRule("4"));
  ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 1, 1), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &reloaded.handle), 0) << fm_get_last_error_message();

  EXPECT_EQ(ByColumn(reloaded.handle, 1).priority, 1);
  EXPECT_EQ(ByColumn(reloaded.handle, 0).priority, 2);
  EXPECT_TRUE(HasPriority(X14StartTag(reloaded.handle->workbook().sheet(0).ext_lst_xml(), kStarsId), "5"));
}

TEST(CfRulePosition, UnparseableExtensionRefusesTheShiftAndChangesNothing) {
  WorkbookGuard wb = LoadSharedIdFixture();
  const std::vector<RuleView> before = ListRules(wb.handle);
  const std::string corrupt = "<extLst><ext><x14:conditionalFormattings>";
  wb.handle->workbook().sheet(0).set_ext_lst_xml(corrupt);

  EXPECT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 1, 1), FM_STATUS_PRECONDITION_FAILED);

  EXPECT_EQ(ListRules(wb.handle), before);
  EXPECT_EQ(wb.handle->workbook().sheet(0).ext_lst_xml(), corrupt);
}

TEST(CfRulePosition, NonIntegerRetainedPriorityRefusesTheShiftAndChangesNothing) {
  WorkbookGuard wb = LoadSharedIdFixture();
  RetainX14Rules(wb.handle, StarsRule("abc"));
  const std::vector<RuleView> before = ListRules(wb.handle);
  const std::string ext_before = wb.handle->workbook().sheet(0).ext_lst_xml();

  EXPECT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 1, 1), FM_STATUS_PRECONDITION_FAILED);

  EXPECT_EQ(ListRules(wb.handle), before);
  EXPECT_EQ(wb.handle->workbook().sheet(0).ext_lst_xml(), ext_before);
}

TEST(CfRulePosition, AutoPriorityCountsRetainedExtensionRules) {
  WorkbookGuard wb = LoadSharedIdFixture();
  RetainX14Rules(wb.handle, StarsRule("9"));

  const std::string added = AddAutoRule(wb.handle);

  EXPECT_EQ(ById(wb.handle, added).priority, 10);
}

TEST(CfRulePosition, AutoPriorityIsOnePastTheModelledMaximumWhenNothingIsRetained) {
  WorkbookGuard wb = LoadSharedIdFixture();

  const std::string added = AddAutoRule(wb.handle);

  EXPECT_EQ(ById(wb.handle, added).priority, 3);
}

// The Excel-saved .xlsb holding one x14 formula rule the reader adopts into the model.
static WorkbookGuard LoadAdoptedX14Xlsb() {
  WorkbookGuard guard;
  const std::string path =
      std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_dv_external_link/expression_cf_unknown_sheet_excel_saved.xlsb";
  FILE* file = std::fopen(path.c_str(), "rb");
  EXPECT_NE(file, nullptr) << "could not open fixture: " << path;
  if (file == nullptr) {
    return guard;
  }
  std::vector<std::uint8_t> bytes;
  std::uint8_t chunk[4096];
  for (std::size_t n = std::fread(chunk, 1, sizeof chunk, file); n != 0; n = std::fread(chunk, 1, sizeof chunk, file)) {
    bytes.insert(bytes.end(), chunk, chunk + n);
  }
  std::fclose(file);
  EXPECT_EQ(fm_workbook_load(bytes.data(), bytes.size(), nullptr, &guard.handle), 0) << fm_get_last_error_message();
  return guard;
}

TEST(CfRulePosition, RemovedAdoptedXlsbRuleNoLongerBlocksSetPriority) {
  WorkbookGuard wb = LoadAdoptedX14Xlsb();
  ASSERT_NE(wb.handle, nullptr);
  const std::vector<RuleView> loaded = ListRules(wb.handle);
  ASSERT_EQ(loaded.size(), 1U);
  const std::string adopted_id = loaded[0].id;
  const std::int32_t adopted_priority = loaded[0].priority;
  ASSERT_GE(adopted_priority, 1);

  const std::string added = AddAutoRule(wb.handle);
  ASSERT_EQ(ById(wb.handle, added).priority, adopted_priority + 1);
  ASSERT_EQ(fm_sheet_remove_cf_rule(wb.handle, 0, adopted_id.c_str()), 0) << fm_get_last_error_message();
  ASSERT_EQ(ListRules(wb.handle).size(), 1U);

  ASSERT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 0, adopted_priority), 0) << fm_get_last_error_message();
  EXPECT_EQ(ById(wb.handle, added).priority, adopted_priority);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &reloaded.handle), 0) << fm_get_last_error_message();
  const std::vector<RuleView> after = ListRules(reloaded.handle);
  ASSERT_EQ(after.size(), 1U);
  EXPECT_EQ(after[0].priority, adopted_priority);
}

TEST(CfRulePosition, UnadoptableXlsbRecordStillBlocksSetPriority) {
  WorkbookGuard wb = LoadAdoptedX14Xlsb();
  ASSERT_NE(wb.handle, nullptr);
  const std::vector<RuleView> loaded = ListRules(wb.handle);
  ASSERT_EQ(loaded.size(), 1U);
  const std::int32_t record_priority = loaded[0].priority;

  const std::string added = AddAutoRule(wb.handle);
  ASSERT_EQ(fm_sheet_remove_cf_rule(wb.handle, 0, loaded[0].id.c_str()), 0) << fm_get_last_error_message();
  // Without its extern-sheet table the record's formula no longer decodes, so a save keeps the record verbatim.
  formulon::XlsbSheetTail tail = wb.handle->workbook().sheet(0).xlsb_tail();
  tail.extern_sheets.clear();
  wb.handle->workbook().sheet(0).set_xlsb_tail(std::move(tail));

  EXPECT_EQ(fm_sheet_set_cf_rule_priority_at(wb.handle, 0, 0, record_priority), FM_STATUS_PRECONDITION_FAILED);
  EXPECT_EQ(ById(wb.handle, added).priority, record_priority + 1);
}
