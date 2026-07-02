// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the conditional-formatting writer. See cf_writer.h
// for the public contract; see cf_reader.cpp for the symmetric grammar.

#include "io/cf_writer.h"

#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>

#include "cf/cf_types.h"
#include "io/ooxml_writer_cell.h"
#include "io/xml_escape.h"

namespace formulon::io {
namespace {

std::string_view RuleTypeToString(cf::RuleType t) {
  switch (t) {
    case cf::RuleType::Expression:
      return "expression";
    case cf::RuleType::CellIs:
      return "cellIs";
    case cf::RuleType::ColorScale:
      return "colorScale";
    case cf::RuleType::DataBar:
      return "dataBar";
    case cf::RuleType::IconSet:
      return "iconSet";
    case cf::RuleType::Top10:
      return "top10";
    case cf::RuleType::AboveAverage:
      return "aboveAverage";
    case cf::RuleType::ContainsText:
      return "containsText";
    case cf::RuleType::NotContainsText:
      return "notContainsText";
    case cf::RuleType::BeginsWith:
      return "beginsWith";
    case cf::RuleType::EndsWith:
      return "endsWith";
    case cf::RuleType::ContainsBlanks:
      return "containsBlanks";
    case cf::RuleType::NotContainsBlanks:
      return "notContainsBlanks";
    case cf::RuleType::ContainsErrors:
      return "containsErrors";
    case cf::RuleType::NotContainsErrors:
      return "notContainsErrors";
    case cf::RuleType::TimePeriod:
      return "timePeriod";
    case cf::RuleType::DuplicateValues:
      return "duplicateValues";
    case cf::RuleType::UniqueValues:
      return "uniqueValues";
  }
  return "expression";
}

std::string_view CellIsOperatorToString(cf::CellIsOperator op) {
  switch (op) {
    case cf::CellIsOperator::LessThan:
      return "lessThan";
    case cf::CellIsOperator::LessThanOrEqual:
      return "lessThanOrEqual";
    case cf::CellIsOperator::Equal:
      return "equal";
    case cf::CellIsOperator::NotEqual:
      return "notEqual";
    case cf::CellIsOperator::GreaterThanOrEqual:
      return "greaterThanOrEqual";
    case cf::CellIsOperator::GreaterThan:
      return "greaterThan";
    case cf::CellIsOperator::Between:
      return "between";
    case cf::CellIsOperator::NotBetween:
      return "notBetween";
  }
  return "equal";
}

std::string_view CfvoTypeToString(cf::CfvoType t) {
  switch (t) {
    case cf::CfvoType::Number:
      return "num";
    case cf::CfvoType::Percent:
      return "percent";
    case cf::CfvoType::Percentile:
      return "percentile";
    case cf::CfvoType::Min:
      return "min";
    case cf::CfvoType::Max:
      return "max";
    case cf::CfvoType::Formula:
      return "formula";
    case cf::CfvoType::AutoMin:
      return "autoMin";
    case cf::CfvoType::AutoMax:
      return "autoMax";
  }
  return "num";
}

std::string_view IconSetNameToString(cf::IconSetName n) {
  switch (n) {
    case cf::IconSetName::Three_Arrows:
      return "3Arrows";
    case cf::IconSetName::Three_ArrowsGray:
      return "3ArrowsGray";
    case cf::IconSetName::Three_Flags:
      return "3Flags";
    case cf::IconSetName::Three_TrafficLights1:
      return "3TrafficLights1";
    case cf::IconSetName::Three_TrafficLights2:
      return "3TrafficLights2";
    case cf::IconSetName::Three_Signs:
      return "3Signs";
    case cf::IconSetName::Three_Symbols:
      return "3Symbols";
    case cf::IconSetName::Three_Symbols2:
      return "3Symbols2";
    case cf::IconSetName::Four_Arrows:
      return "4Arrows";
    case cf::IconSetName::Four_ArrowsGray:
      return "4ArrowsGray";
    case cf::IconSetName::Four_RedToBlack:
      return "4RedToBlack";
    case cf::IconSetName::Four_Rating:
      return "4Rating";
    case cf::IconSetName::Four_TrafficLights:
      return "4TrafficLights";
    case cf::IconSetName::Five_Arrows:
      return "5Arrows";
    case cf::IconSetName::Five_ArrowsGray:
      return "5ArrowsGray";
    case cf::IconSetName::Five_Rating:
      return "5Rating";
    case cf::IconSetName::Five_Quarters:
      return "5Quarters";
  }
  return "3Arrows";
}

std::string_view TimePeriodToString(cf::TimePeriod p) {
  switch (p) {
    case cf::TimePeriod::Today:
      return "today";
    case cf::TimePeriod::Yesterday:
      return "yesterday";
    case cf::TimePeriod::Tomorrow:
      return "tomorrow";
    case cf::TimePeriod::Last7Days:
      return "last7Days";
    case cf::TimePeriod::ThisWeek:
      return "thisWeek";
    case cf::TimePeriod::LastWeek:
      return "lastWeek";
    case cf::TimePeriod::NextWeek:
      return "nextWeek";
    case cf::TimePeriod::ThisMonth:
      return "thisMonth";
    case cf::TimePeriod::LastMonth:
      return "lastMonth";
    case cf::TimePeriod::NextMonth:
      return "nextMonth";
  }
  return "today";
}

/// Appends the Excel column letters (`A`, `AB`, `XFD`) for a 0-based column
/// index to `out`.
void AppendColumnLetters(std::string& out, std::uint32_t col) {
  char buf[4];
  int len = 0;
  std::uint32_t n = col + 1U;  // 1-based for the base-26 bijection.
  while (n > 0U && len < 4) {
    const std::uint32_t rem = (n - 1U) % 26U;
    buf[len++] = static_cast<char>('A' + static_cast<int>(rem));
    n = (n - 1U) / 26U;
  }
  for (int i = len - 1; i >= 0; --i) {
    out.push_back(buf[i]);
  }
}

/// Formats one `CFCellRange` as A1 (single cell) or A1:B5 (range). The
/// reader accepts both `A1:A1` and `A1` for a single cell; the writer
/// prefers the shorter `A1` form so the round-trip output matches what
/// Excel emits. Whole-column / whole-row ranges re-emit the compact
/// `A:A` / `1:1` form Excel authors.
std::string EncodeA1Range(const cf::CFCellRange& r) {
  if (r.is_full_col()) {
    std::string out;
    AppendColumnLetters(out, r.first.col);
    out.push_back(':');
    AppendColumnLetters(out, r.last.col);
    return out;
  }
  if (r.is_full_row()) {
    std::string out = std::to_string(r.first.row + 1U);
    out.push_back(':');
    out.append(std::to_string(r.last.row + 1U));
    return out;
  }
  std::string a = EncodeA1(r.first.row, r.first.col);
  if (r.first == r.last) {
    return a;
  }
  std::string b = EncodeA1(r.last.row, r.last.col);
  a.push_back(':');
  a.append(b);
  return a;
}

std::string EncodeSqref(const std::vector<cf::CFCellRange>& ranges) {
  std::string out;
  out.reserve(ranges.size() * 12);
  for (std::size_t i = 0; i < ranges.size(); ++i) {
    if (i != 0) {
      out.push_back(' ');
    }
    out.append(EncodeA1Range(ranges[i]));
  }
  return out;
}

void AppendColor(std::string& out, cf::Color c) {
  char buf[16];
  std::snprintf(buf, sizeof(buf), "%02X%02X%02X%02X", c.a, c.r, c.g, c.b);
  out.append("<color rgb=\"");
  out.append(buf);
  out.append("\"/>");
}

void AppendCfvo(std::string& out, const cf::CfValueObject& v) {
  out.append("<cfvo type=\"");
  out.append(CfvoTypeToString(v.type));
  out.push_back('"');
  if (!v.value.empty()) {
    out.append(" val=\"");
    AppendXmlAttrEscaped(out, v.value);
    out.push_back('"');
  }
  if (!v.gte) {
    out.append(" gte=\"0\"");
  }
  out.append("/>");
}

void AppendColorScale(std::string& out, const cf::ColorScaleSpec& s) {
  out.append("<colorScale>");
  for (const auto& th : s.thresholds) {
    AppendCfvo(out, th);
  }
  for (const auto& c : s.colors) {
    AppendColor(out, c);
  }
  out.append("</colorScale>");
}

void AppendDataBar(std::string& out, const cf::DataBarSpec& d) {
  out.append("<dataBar minLength=\"");
  out.append(std::to_string(static_cast<unsigned>(d.min_length_pct)));
  out.append("\" maxLength=\"");
  out.append(std::to_string(static_cast<unsigned>(d.max_length_pct)));
  out.push_back('"');
  if (!d.show_value) {
    out.append(" showValue=\"0\"");
  }
  out.push_back('>');
  AppendCfvo(out, d.min);
  AppendCfvo(out, d.max);
  AppendColor(out, d.fill);
  out.append("</dataBar>");
}

void AppendIconSet(std::string& out, const cf::IconSetSpec& i) {
  out.append("<iconSet iconSet=\"");
  out.append(IconSetNameToString(i.name));
  out.push_back('"');
  if (i.reverse) {
    out.append(" reverse=\"1\"");
  }
  if (!i.show_value) {
    out.append(" showValue=\"0\"");
  }
  if (!i.percent) {
    // Reader defaults `percent` to true; only emit the attribute when
    // the user opted out so the round-trip stays minimal.
    out.append(" percent=\"0\"");
  }
  out.push_back('>');
  // OOXML requires N `<cfvo>` children for an N-icon set, but the model
  // only carries the N-1 real boundary thresholds (see cf_reader.cpp's
  // `ReadIconSet`); re-synthesize the dropped floor cfvo here so the
  // emitted XML stays schema-valid and round-trips through Excel.
  cf::CfValueObject floor;
  floor.type = cf::CfvoType::Percent;
  floor.value = "0";
  AppendCfvo(out, floor);
  for (const auto& th : i.thresholds) {
    AppendCfvo(out, th);
  }
  out.append("</iconSet>");
}

void AppendCfRule(std::string& out, const cf::CFRule& r) {
  out.append("<cfRule type=\"");
  out.append(RuleTypeToString(r.type));
  out.append("\" priority=\"");
  out.append(std::to_string(r.priority));
  out.push_back('"');
  if (r.stop_if_true) {
    out.append(" stopIfTrue=\"1\"");
  }
  if (r.dxf_id.has_value()) {
    out.append(" dxfId=\"");
    out.append(std::to_string(r.dxf_id.value()));
    out.push_back('"');
  }
  if (r.type == cf::RuleType::CellIs && r.op.has_value()) {
    out.append(" operator=\"");
    out.append(CellIsOperatorToString(r.op.value()));
    out.push_back('"');
  }
  if ((r.type == cf::RuleType::ContainsText || r.type == cf::RuleType::NotContainsText ||
       r.type == cf::RuleType::BeginsWith || r.type == cf::RuleType::EndsWith) &&
      r.text.has_value()) {
    // Excel emits operator on these rule types as a hint to the
    // legacy reader (e.g. `operator="containsText"`); the modern
    // reader keys solely off `type`. Emit it for parity with Excel-
    // saved files.
    const char* op_attr = nullptr;
    switch (r.type) {
      case cf::RuleType::ContainsText:
        op_attr = "containsText";
        break;
      case cf::RuleType::NotContainsText:
        op_attr = "notContains";
        break;
      case cf::RuleType::BeginsWith:
        op_attr = "beginsWith";
        break;
      case cf::RuleType::EndsWith:
        op_attr = "endsWith";
        break;
      default:
        break;
    }
    if (op_attr != nullptr) {
      out.append(" operator=\"");
      out.append(op_attr);
      out.push_back('"');
    }
    out.append(" text=\"");
    AppendXmlAttrEscaped(out, r.text.value());
    out.push_back('"');
  }
  if (r.type == cf::RuleType::Top10) {
    if (r.rank.has_value()) {
      out.append(" rank=\"");
      out.append(std::to_string(r.rank.value()));
      out.push_back('"');
    }
    if (r.percent) {
      out.append(" percent=\"1\"");
    }
    if (r.bottom) {
      out.append(" bottom=\"1\"");
    }
  }
  if (r.type == cf::RuleType::AboveAverage) {
    if (!r.above_average) {
      out.append(" aboveAverage=\"0\"");
    }
    if (r.equal_average) {
      out.append(" equalAverage=\"1\"");
    }
    if (r.std_dev.has_value()) {
      char buf[32];
      std::snprintf(buf, sizeof(buf), "%g", r.std_dev.value());
      out.append(" stdDev=\"");
      out.append(buf);
      out.push_back('"');
    }
  }
  if (r.type == cf::RuleType::TimePeriod && r.time_period.has_value()) {
    out.append(" timePeriod=\"");
    out.append(TimePeriodToString(r.time_period.value()));
    out.push_back('"');
  }
  if (!r.id.empty()) {
    out.append(" id=\"");
    AppendXmlAttrEscaped(out, r.id);
    out.push_back('"');
  }
  out.push_back('>');

  if (r.formula1.has_value()) {
    out.append("<formula>");
    AppendXmlEscaped(out, r.formula1.value());
    out.append("</formula>");
  }
  if (r.formula2.has_value()) {
    out.append("<formula>");
    AppendXmlEscaped(out, r.formula2.value());
    out.append("</formula>");
  }
  if (r.color_scale.has_value()) {
    AppendColorScale(out, r.color_scale.value());
  }
  if (r.data_bar.has_value()) {
    AppendDataBar(out, r.data_bar.value());
  }
  if (r.icon_set.has_value()) {
    AppendIconSet(out, r.icon_set.value());
  }
  // `CT_CfRule`'s schema-trailing `extLst?`, round-tripped byte-for-byte
  // (see `CFRule::ext_lst_raw`); already a complete `<extLst>...</extLst>`
  // string captured verbatim by the reader.
  if (r.ext_lst_raw.has_value()) {
    out.append(r.ext_lst_raw.value());
  }

  out.append("</cfRule>");
}

}  // namespace

std::string write_conditional_formattings(const std::vector<cf::ConditionalFormat>& formats) {
  if (formats.empty()) {
    return std::string{};
  }
  std::string out;
  out.reserve(formats.size() * 128);
  for (const auto& cf : formats) {
    out.append("<conditionalFormatting sqref=\"");
    AppendXmlAttrEscaped(out, EncodeSqref(cf.sqref));
    out.push_back('"');
    if (cf.pivot_scope) {
      out.append(" pivot=\"1\"");
    }
    out.push_back('>');
    for (const auto& rule : cf.rules) {
      AppendCfRule(out, rule);
    }
    // `CT_ConditionalFormatting`'s schema-trailing `extLst?`, round-tripped
    // byte-for-byte (see `ConditionalFormat::ext_lst_raw`).
    if (cf.ext_lst_raw.has_value()) {
      out.append(cf.ext_lst_raw.value());
    }
    out.append("</conditionalFormatting>");
  }
  return out;
}

}  // namespace formulon::io
