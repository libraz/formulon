// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the conditional-formatting reader. See cf_reader.h
// for the public contract.

#include "io/cf_reader.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "io/cell_parser.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {
namespace {

bool ParseBoolAttr(const pugi::xml_attribute& attr) {
  if (!attr) {
    return false;
  }
  const std::string_view text = attr.value();
  return text == "1" || text == "true";
}

cf::RuleType ParseRuleType(std::string_view text) {
  if (text == "expression")
    return cf::RuleType::Expression;
  if (text == "cellIs")
    return cf::RuleType::CellIs;
  if (text == "colorScale")
    return cf::RuleType::ColorScale;
  if (text == "dataBar")
    return cf::RuleType::DataBar;
  if (text == "iconSet")
    return cf::RuleType::IconSet;
  if (text == "top10")
    return cf::RuleType::Top10;
  if (text == "aboveAverage")
    return cf::RuleType::AboveAverage;
  if (text == "containsText")
    return cf::RuleType::ContainsText;
  if (text == "notContainsText")
    return cf::RuleType::NotContainsText;
  if (text == "beginsWith")
    return cf::RuleType::BeginsWith;
  if (text == "endsWith")
    return cf::RuleType::EndsWith;
  if (text == "containsBlanks")
    return cf::RuleType::ContainsBlanks;
  if (text == "notContainsBlanks")
    return cf::RuleType::NotContainsBlanks;
  if (text == "containsErrors")
    return cf::RuleType::ContainsErrors;
  if (text == "notContainsErrors")
    return cf::RuleType::NotContainsErrors;
  if (text == "timePeriod")
    return cf::RuleType::TimePeriod;
  if (text == "duplicateValues")
    return cf::RuleType::DuplicateValues;
  if (text == "uniqueValues")
    return cf::RuleType::UniqueValues;
  // Forward-compat: unknown types fold to Expression. The captured
  // formula1 (if any) is still observable to the caller.
  return cf::RuleType::Expression;
}

cf::CellIsOperator ParseCellIsOperator(std::string_view text) {
  if (text == "lessThan")
    return cf::CellIsOperator::LessThan;
  if (text == "lessThanOrEqual")
    return cf::CellIsOperator::LessThanOrEqual;
  if (text == "equal")
    return cf::CellIsOperator::Equal;
  if (text == "notEqual")
    return cf::CellIsOperator::NotEqual;
  if (text == "greaterThanOrEqual")
    return cf::CellIsOperator::GreaterThanOrEqual;
  if (text == "greaterThan")
    return cf::CellIsOperator::GreaterThan;
  if (text == "between")
    return cf::CellIsOperator::Between;
  if (text == "notBetween")
    return cf::CellIsOperator::NotBetween;
  return cf::CellIsOperator::Equal;
}

cf::CfvoType ParseCfvoType(std::string_view text) {
  if (text == "num")
    return cf::CfvoType::Number;
  if (text == "percent")
    return cf::CfvoType::Percent;
  if (text == "percentile")
    return cf::CfvoType::Percentile;
  if (text == "min")
    return cf::CfvoType::Min;
  if (text == "max")
    return cf::CfvoType::Max;
  if (text == "formula")
    return cf::CfvoType::Formula;
  if (text == "autoMin")
    return cf::CfvoType::AutoMin;
  if (text == "autoMax")
    return cf::CfvoType::AutoMax;
  return cf::CfvoType::Number;
}

cf::IconSetName ParseIconSetName(std::string_view text) {
  if (text == "3Arrows")
    return cf::IconSetName::Three_Arrows;
  if (text == "3ArrowsGray")
    return cf::IconSetName::Three_ArrowsGray;
  if (text == "3Flags")
    return cf::IconSetName::Three_Flags;
  if (text == "3TrafficLights1")
    return cf::IconSetName::Three_TrafficLights1;
  if (text == "3TrafficLights2")
    return cf::IconSetName::Three_TrafficLights2;
  if (text == "3Signs")
    return cf::IconSetName::Three_Signs;
  if (text == "3Symbols")
    return cf::IconSetName::Three_Symbols;
  if (text == "3Symbols2")
    return cf::IconSetName::Three_Symbols2;
  if (text == "4Arrows")
    return cf::IconSetName::Four_Arrows;
  if (text == "4ArrowsGray")
    return cf::IconSetName::Four_ArrowsGray;
  if (text == "4RedToBlack")
    return cf::IconSetName::Four_RedToBlack;
  if (text == "4Rating")
    return cf::IconSetName::Four_Rating;
  if (text == "4TrafficLights")
    return cf::IconSetName::Four_TrafficLights;
  if (text == "5Arrows")
    return cf::IconSetName::Five_Arrows;
  if (text == "5ArrowsGray")
    return cf::IconSetName::Five_ArrowsGray;
  if (text == "5Rating")
    return cf::IconSetName::Five_Rating;
  if (text == "5Quarters")
    return cf::IconSetName::Five_Quarters;
  return cf::IconSetName::Three_Arrows;
}

cf::TimePeriod ParseTimePeriod(std::string_view text) {
  if (text == "yesterday")
    return cf::TimePeriod::Yesterday;
  if (text == "tomorrow")
    return cf::TimePeriod::Tomorrow;
  if (text == "last7Days")
    return cf::TimePeriod::Last7Days;
  if (text == "thisWeek")
    return cf::TimePeriod::ThisWeek;
  if (text == "lastWeek")
    return cf::TimePeriod::LastWeek;
  if (text == "nextWeek")
    return cf::TimePeriod::NextWeek;
  if (text == "thisMonth")
    return cf::TimePeriod::ThisMonth;
  if (text == "lastMonth")
    return cf::TimePeriod::LastMonth;
  if (text == "nextMonth")
    return cf::TimePeriod::NextMonth;
  // "today" and any unknown spelling.
  return cf::TimePeriod::Today;
}

/// Decodes one A1 cell-range token (`A1`, `A1:B5`, `$A$1:$B$5`) into a
/// `CFCellRange`. Single-cell tokens land as `first == last`. Returns
/// `nullopt` for unparseable input — the caller folds the surrounding
/// sqref into `kIoSheetCorrupt`.
Expected<cf::CFCellRange, Error> ParseA1Range(std::string_view ref) {
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    auto rc = parse_a1(ref);
    if (!rc) {
      std::string ctx("context=cf_reader ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "conditionalFormatting: sqref token unparseable",
                        std::move(ctx));
    }
    cf::CFCellRange out{};
    out.first = {rc.value().first, rc.value().second};
    out.last = out.first;
    return out;
  }
  const std::string_view a = ref.substr(0, colon);
  const std::string_view b = ref.substr(colon + 1);
  auto a_rc = parse_a1(a);
  auto b_rc = parse_a1(b);
  if (!a_rc || !b_rc) {
    std::string ctx("context=cf_reader ref=");
    ctx.append(ref);
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "conditionalFormatting: sqref token unparseable",
                      std::move(ctx));
  }
  const std::uint32_t r0 = a_rc.value().first;
  const std::uint32_t c0 = a_rc.value().second;
  const std::uint32_t r1 = b_rc.value().first;
  const std::uint32_t c1 = b_rc.value().second;
  cf::CFCellRange out{};
  out.first = {(r0 < r1) ? r0 : r1, (c0 < c1) ? c0 : c1};
  out.last = {(r0 < r1) ? r1 : r0, (c0 < c1) ? c1 : c0};
  return out;
}

/// Splits a whitespace-separated sqref attribute into individual A1
/// range tokens and decodes each. Returns `kIoSheetCorrupt` on the
/// first unparseable token.
Expected<std::vector<cf::CFCellRange>, Error> ParseSqref(std::string_view sqref) {
  std::vector<cf::CFCellRange> out;
  std::size_t i = 0;
  while (i < sqref.size()) {
    while (i < sqref.size() && (sqref[i] == ' ' || sqref[i] == '\t' || sqref[i] == '\n' || sqref[i] == '\r')) {
      ++i;
    }
    const std::size_t start = i;
    while (i < sqref.size() && sqref[i] != ' ' && sqref[i] != '\t' && sqref[i] != '\n' && sqref[i] != '\r') {
      ++i;
    }
    if (start == i) {
      break;
    }
    const std::string_view tok = sqref.substr(start, i - start);
    auto range_or = ParseA1Range(tok);
    if (!range_or) {
      return range_or.error();
    }
    out.push_back(range_or.value());
  }
  if (out.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "conditionalFormatting: sqref attribute is empty",
                      "context=cf_reader");
  }
  return out;
}

/// Decodes one hex digit. Returns -1 on a non-hex character.
int HexDigit(char c) {
  if (c >= '0' && c <= '9')
    return c - '0';
  if (c >= 'a' && c <= 'f')
    return 10 + (c - 'a');
  if (c >= 'A' && c <= 'F')
    return 10 + (c - 'A');
  return -1;
}

/// Parses a `<color rgb="AARRGGBB">` attribute into a `cf::Color`. The
/// alpha channel is optional in the spec (some Excel exports drop it
/// for fully-opaque colours, emitting `RRGGBB` only); both 6- and 8-hex
/// forms are accepted. Returns opaque black on malformed input — the
/// caller is free to drop the rule, but in practice CF colours are
/// either valid or absent.
cf::Color ParseRgbColor(std::string_view rgb) {
  cf::Color out{};
  if (rgb.empty()) {
    return out;
  }
  std::size_t off = 0;
  if (rgb.size() == 8) {
    const int a_hi = HexDigit(rgb[0]);
    const int a_lo = HexDigit(rgb[1]);
    if (a_hi < 0 || a_lo < 0) {
      return cf::Color{};
    }
    out.a = static_cast<std::uint8_t>((a_hi << 4) | a_lo);
    off = 2;
  } else if (rgb.size() == 6) {
    out.a = 255;
  } else {
    return cf::Color{};
  }
  const int r_hi = HexDigit(rgb[off]);
  const int r_lo = HexDigit(rgb[off + 1]);
  const int g_hi = HexDigit(rgb[off + 2]);
  const int g_lo = HexDigit(rgb[off + 3]);
  const int b_hi = HexDigit(rgb[off + 4]);
  const int b_lo = HexDigit(rgb[off + 5]);
  if (r_hi < 0 || r_lo < 0 || g_hi < 0 || g_lo < 0 || b_hi < 0 || b_lo < 0) {
    return cf::Color{};
  }
  out.r = static_cast<std::uint8_t>((r_hi << 4) | r_lo);
  out.g = static_cast<std::uint8_t>((g_hi << 4) | g_lo);
  out.b = static_cast<std::uint8_t>((b_hi << 4) | b_lo);
  return out;
}

cf::CfValueObject ReadCfvo(const pugi::xml_node& cfvo) {
  cf::CfValueObject out{};
  out.type = ParseCfvoType(cfvo.attribute("type").value());
  out.value = cfvo.attribute("val").value();
  // `gte` is "1" by default; only "0"/"false" toggles to false.
  const pugi::xml_attribute gte = cfvo.attribute("gte");
  if (gte) {
    const std::string_view text = gte.value();
    out.gte = !(text == "0" || text == "false");
  } else {
    out.gte = true;
  }
  return out;
}

void ReadColorScale(const pugi::xml_node& scale, cf::ColorScaleSpec* out) {
  for (pugi::xml_node child = scale.first_child(); child; child = child.next_sibling()) {
    const std::string_view name = child.name();
    if (name == "cfvo") {
      out->thresholds.push_back(ReadCfvo(child));
    } else if (name == "color") {
      out->colors.push_back(ParseRgbColor(child.attribute("rgb").value()));
    }
  }
}

void ReadDataBar(const pugi::xml_node& bar, cf::DataBarSpec* out) {
  std::size_t color_idx = 0;
  std::size_t cfvo_idx = 0;
  for (pugi::xml_node child = bar.first_child(); child; child = child.next_sibling()) {
    const std::string_view name = child.name();
    if (name == "cfvo") {
      if (cfvo_idx == 0) {
        out->min = ReadCfvo(child);
      } else if (cfvo_idx == 1) {
        out->max = ReadCfvo(child);
      }
      ++cfvo_idx;
    } else if (name == "color") {
      // Legacy `<dataBar>` (pre-2010) carries one `<color>` element used
      // as the positive fill. The 2010+ extension lifts negative-fill /
      // axis-color into `<extLst>`, which this PR does not yet consume.
      if (color_idx == 0) {
        out->fill = ParseRgbColor(child.attribute("rgb").value());
        out->negative_fill = out->fill;
      }
      ++color_idx;
    }
  }
  const pugi::xml_attribute min_len = bar.attribute("minLength");
  if (min_len) {
    const long parsed = std::strtol(min_len.value(), nullptr, 10);
    if (parsed >= 0 && parsed <= 100) {
      out->min_length_pct = static_cast<std::uint8_t>(parsed);
    }
  }
  const pugi::xml_attribute max_len = bar.attribute("maxLength");
  if (max_len) {
    const long parsed = std::strtol(max_len.value(), nullptr, 10);
    if (parsed >= 0 && parsed <= 100) {
      out->max_length_pct = static_cast<std::uint8_t>(parsed);
    }
  }
  const pugi::xml_attribute show_value = bar.attribute("showValue");
  if (show_value) {
    out->show_value = ParseBoolAttr(show_value);
  }
}

void ReadIconSet(const pugi::xml_node& iset, cf::IconSetSpec* out) {
  out->name = ParseIconSetName(iset.attribute("iconSet").value());
  const pugi::xml_attribute reverse = iset.attribute("reverse");
  if (reverse) {
    out->reverse = ParseBoolAttr(reverse);
  }
  const pugi::xml_attribute show_value = iset.attribute("showValue");
  if (show_value) {
    out->show_value = ParseBoolAttr(show_value);
  }
  const pugi::xml_attribute percent = iset.attribute("percent");
  if (percent) {
    out->percent = ParseBoolAttr(percent);
  }
  for (pugi::xml_node cfvo = iset.child("cfvo"); cfvo; cfvo = cfvo.next_sibling("cfvo")) {
    out->thresholds.push_back(ReadCfvo(cfvo));
  }
}

cf::CFRule ReadCfRule(const pugi::xml_node& rule) {
  cf::CFRule out;
  const pugi::xml_attribute type_attr = rule.attribute("type");
  out.type = ParseRuleType(type_attr.value());
  out.priority = parse_xml_i32_attr(rule.attribute("priority"), 1);
  out.stop_if_true = ParseBoolAttr(rule.attribute("stopIfTrue"));
  const pugi::xml_attribute dxf_attr = rule.attribute("dxfId");
  if (dxf_attr) {
    out.dxf_id = static_cast<std::uint32_t>(parse_xml_i32_attr(dxf_attr, 0));
  }
  const pugi::xml_attribute id_attr = rule.attribute("id");
  if (id_attr) {
    out.id = id_attr.value();
  }

  if (out.type == cf::RuleType::CellIs) {
    out.op = ParseCellIsOperator(rule.attribute("operator").value());
  }
  if (out.type == cf::RuleType::TimePeriod) {
    out.time_period = ParseTimePeriod(rule.attribute("timePeriod").value());
  }
  if (out.type == cf::RuleType::Top10) {
    out.rank = parse_xml_i32_attr(rule.attribute("rank"), 10);
    out.percent = ParseBoolAttr(rule.attribute("percent"));
    out.bottom = ParseBoolAttr(rule.attribute("bottom"));
  }
  if (out.type == cf::RuleType::AboveAverage) {
    const pugi::xml_attribute above_attr = rule.attribute("aboveAverage");
    out.above_average = above_attr ? ParseBoolAttr(above_attr) : true;
    out.equal_average = ParseBoolAttr(rule.attribute("equalAverage"));
    const pugi::xml_attribute std_dev_attr = rule.attribute("stdDev");
    if (std_dev_attr) {
      out.std_dev = std::strtod(std_dev_attr.value(), nullptr);
    }
  }
  if (out.type == cf::RuleType::ContainsText || out.type == cf::RuleType::NotContainsText ||
      out.type == cf::RuleType::BeginsWith || out.type == cf::RuleType::EndsWith) {
    const pugi::xml_attribute text_attr = rule.attribute("text");
    if (text_attr) {
      out.text = text_attr.value();
    }
  }

  // Children: <formula>, <colorScale>, <dataBar>, <iconSet> (mutually
  // exclusive at the visual-payload level).
  std::size_t formula_idx = 0;
  for (pugi::xml_node child = rule.first_child(); child; child = child.next_sibling()) {
    const std::string_view name = child.name();
    if (name == "formula") {
      const std::string_view raw = child.text().get();
      std::string body(raw);
      if (!body.empty() && body[0] == '=') {
        body.erase(body.begin());
      }
      if (formula_idx == 0) {
        out.formula1 = std::move(body);
      } else if (formula_idx == 1) {
        out.formula2 = std::move(body);
      }
      ++formula_idx;
    } else if (name == "colorScale") {
      cf::ColorScaleSpec spec;
      ReadColorScale(child, &spec);
      out.color_scale = std::move(spec);
    } else if (name == "dataBar") {
      cf::DataBarSpec spec;
      ReadDataBar(child, &spec);
      out.data_bar = std::move(spec);
    } else if (name == "iconSet") {
      cf::IconSetSpec spec;
      ReadIconSet(child, &spec);
      out.icon_set = std::move(spec);
    }
  }

  return out;
}

}  // namespace

Expected<std::vector<cf::ConditionalFormat>, Error> read_conditional_formats(const pugi::xml_node& worksheet) {
  std::vector<cf::ConditionalFormat> out;
  for (pugi::xml_node block = worksheet.child("conditionalFormatting"); block;
       block = block.next_sibling("conditionalFormatting")) {
    const pugi::xml_attribute sqref_attr = block.attribute("sqref");
    if (!sqref_attr) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "conditionalFormatting: required sqref attribute missing",
                        "context=cf_reader");
    }
    auto ranges_or = ParseSqref(sqref_attr.value());
    if (!ranges_or) {
      return ranges_or.error();
    }
    cf::ConditionalFormat cfmt;
    cfmt.sqref = std::move(ranges_or.value());
    cfmt.pivot_scope = ParseBoolAttr(block.attribute("pivot"));
    for (pugi::xml_node rule = block.child("cfRule"); rule; rule = rule.next_sibling("cfRule")) {
      cfmt.rules.push_back(ReadCfRule(rule));
    }
    out.push_back(std::move(cfmt));
  }
  return out;
}

}  // namespace formulon::io
