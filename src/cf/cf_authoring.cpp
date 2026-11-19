#include "cf/cf_authoring.h"

#include <array>
#include <cstdint>
#include <string_view>

#include "utils/a1_ref.h"

namespace formulon::cf {

namespace {

// Templates copied from Excel-saved sheets: '@' is the anchor cell, '#' the
// rule text with its quotes doubled.
constexpr std::array<std::string_view, 8> kTextFamilyTemplates = {
    "NOT(ISERROR(SEARCH(\"#\",@)))",  // ContainsText
    "ISERROR(SEARCH(\"#\",@))",       // NotContainsText
    "LEFT(@,LEN(\"#\"))=\"#\"",       // BeginsWith
    "RIGHT(@,LEN(\"#\"))=\"#\"",      // EndsWith
    "LEN(TRIM(@))=0",                 // ContainsBlanks
    "LEN(TRIM(@))>0",                 // NotContainsBlanks
    "ISERROR(@)",                     // ContainsErrors
    "NOT(ISERROR(@))",                // NotContainsErrors
};

// Indexed by `TimePeriod`.
constexpr std::array<std::string_view, 10> kTimePeriodTemplates = {
    "FLOOR(@,1)=TODAY()",
    "FLOOR(@,1)=TODAY()-1",
    "FLOOR(@,1)=TODAY()+1",
    "AND(TODAY()-FLOOR(@,1)<=6,FLOOR(@,1)<=TODAY())",
    "AND(TODAY()-ROUNDDOWN(@,0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(@,0)-TODAY()<=7-WEEKDAY(TODAY()))",
    "AND(TODAY()-ROUNDDOWN(@,0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(@,0)<(WEEKDAY(TODAY())+7))",
    "AND(ROUNDDOWN(@,0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(@,0)-TODAY()<(15-WEEKDAY(TODAY())))",
    "AND(MONTH(@)=MONTH(TODAY()),YEAR(@)=YEAR(TODAY()))",
    "AND(MONTH(@)=MONTH(EDATE(TODAY(),0-1)),YEAR(@)=YEAR(EDATE(TODAY(),0-1)))",
    "AND(MONTH(@)=MONTH(EDATE(TODAY(),0+1)),YEAR(@)=YEAR(EDATE(TODAY(),0+1)))",
};

}  // namespace

std::optional<std::string> derived_rule_formula(const CFRule& rule, CellAddress anchor) {
  std::string_view tmpl;
  const auto type = static_cast<std::uint8_t>(rule.type);
  if (rule.type == RuleType::TimePeriod) {
    if (!rule.time_period.has_value()) {
      return std::nullopt;
    }
    tmpl = kTimePeriodTemplates[static_cast<std::size_t>(*rule.time_period)];
  } else if (type >= static_cast<std::uint8_t>(RuleType::ContainsText) &&
             type <= static_cast<std::uint8_t>(RuleType::NotContainsErrors)) {
    if (type <= static_cast<std::uint8_t>(RuleType::EndsWith) && !rule.text.has_value()) {
      return std::nullopt;
    }
    tmpl = kTextFamilyTemplates[static_cast<std::size_t>(type) - static_cast<std::size_t>(RuleType::ContainsText)];
  } else {
    return std::nullopt;
  }
  const std::string cell = a1::encode_a1(anchor.row, anchor.col);
  std::string out;
  for (const char c : tmpl) {
    if (c == '@') {
      out += cell;
    } else if (c == '#') {
      for (const char t : *rule.text) {
        out += t;
        if (t == '"') {
          out += '"';
        }
      }
    } else {
      out += c;
    }
  }
  return out;
}

}  // namespace formulon::cf
