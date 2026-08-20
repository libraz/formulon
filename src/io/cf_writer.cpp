//
// Implementation of the conditional-formatting writer. See cf_writer.h
// for the public contract; see cf_reader.cpp for the symmetric grammar.

#include "io/cf_writer.h"

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>

#include "cf/cf_types.h"
#include "io/ooxml_writer_cell.h"
#include "io/xml_escape.h"
#include "utils/a1_column.h"

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

/// Appends one `CFCellRange` as A1 (single cell) or A1:B5 (range). The
/// reader accepts both `A1:A1` and `A1` for a single cell; the writer
/// prefers the shorter `A1` form so the round-trip output matches what
/// Excel emits. Whole-column / whole-row ranges re-emit the compact
/// `A:A` / `1:1` form Excel authors.
///
/// Returns `false` and leaves `out` untouched for an inverted range or
/// one reaching outside the grid, which has no A1 spelling. The mutation
/// API rejects such a rectangle at set time; this is the defensive half,
/// so a model assembled in-process still cannot make the writer emit a
/// reference Excel reads as a broken block.
bool AppendA1Range(std::string& out, const cf::CFCellRange& r) {
  if (r.first.row > r.last.row || r.first.col > r.last.col || r.last.row >= cf::kCfMaxRows ||
      r.last.col >= cf::kCfMaxCols) {
    return false;
  }
  std::string encoded;
  if (r.is_full_col()) {
    if (!a1::append_column_letters(encoded, r.first.col)) {
      return false;
    }
    encoded.push_back(':');
    if (!a1::append_column_letters(encoded, r.last.col)) {
      return false;
    }
  } else if (r.is_full_row()) {
    encoded = std::to_string(r.first.row + 1U);
    encoded.push_back(':');
    encoded.append(std::to_string(r.last.row + 1U));
  } else {
    encoded = EncodeA1(r.first.row, r.first.col);
    if (encoded.empty()) {
      return false;
    }
    if (r.first != r.last) {
      const std::string last = EncodeA1(r.last.row, r.last.col);
      if (last.empty()) {
        return false;
      }
      encoded.push_back(':');
      encoded.append(last);
    }
  }
  out.append(encoded);
  return true;
}

std::string EncodeSqref(const std::vector<cf::CFCellRange>& ranges) {
  std::string out;
  out.reserve(ranges.size() * 12);
  for (const auto& range : ranges) {
    const std::size_t mark = out.size();
    if (!out.empty()) {
      out.push_back(' ');
    }
    if (!AppendA1Range(out, range)) {
      out.resize(mark);
    }
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

/// URI of the `<ext>` that links a legacy `<cfRule>` to its x14
/// counterpart, and the two namespaces the extension content lives in.
/// All three are fixed constants Excel matches literally.
constexpr std::string_view kX14CfRuleExtUri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";
constexpr std::string_view kX14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
constexpr std::string_view kXmNs = "http://schemas.microsoft.com/office/excel/2006/main";

/// The axis colour a data bar has when the x14 extension says nothing —
/// mirrors `DataBarSpec::axis_color`'s in-class initializer.
constexpr cf::Color kDefaultAxisColor{0, 0, 0, 255};

/// True when `d` carries at least one setting the legacy `<dataBar>`
/// element has no attribute for, so the rule needs an x14 counterpart to
/// survive a save. False for a plain bar, which keeps a legacy-only file
/// legacy-only across a round trip.
bool NeedsX14DataBarPayload(const cf::DataBarSpec& d) {
  return d.border.has_value() || d.negative_border.has_value() || d.negative_fill != d.fill ||
         d.axis_position != cf::DataBarAxisPosition::Automatic || !d.gradient || d.axis_color != kDefaultAxisColor;
}

/// True when the rule both needs an x14 counterpart and has an id to
/// link it by. An id is required: the link and the payload find each
/// other by that GUID and nothing else.
bool RuleNeedsX14Payload(const cf::CFRule& r) {
  return !r.id.empty() && r.data_bar.has_value() && NeedsX14DataBarPayload(r.data_bar.value());
}

/// True when the rule's captured `<extLst>` already carries an x14 link,
/// i.e. the rule was loaded from a file that had one. Re-emitting the
/// capture verbatim is then both sufficient and more faithful than
/// rebuilding the element.
bool HasCapturedX14Link(const cf::CFRule& r) {
  return r.ext_lst_raw.has_value() && r.ext_lst_raw->find("x14:id") != std::string::npos;
}

/// Emits `CT_CfRule`'s schema-trailing `extLst?`, folding in the x14
/// link when the rule needs one and does not already carry it.
///
/// The link is a nested `<ext><x14:id>`, never an `id` attribute on the
/// `<cfRule>` itself: `CT_CfRule` has no such attribute, and Excel
/// discards both the attribute and the worksheet-level payload it was
/// meant to reach when it re-saves a file spelled that way.
void AppendRuleExtLst(std::string& out, const cf::CFRule& r) {
  if (!RuleNeedsX14Payload(r) || HasCapturedX14Link(r)) {
    if (r.ext_lst_raw.has_value()) {
      out.append(r.ext_lst_raw.value());
    }
    return;
  }

  std::string link("<ext uri=\"");
  link.append(kX14CfRuleExtUri);
  link.append("\" xmlns:x14=\"");
  link.append(kX14Ns);
  link.append("\"><x14:id>");
  AppendXmlEscaped(link, r.id);
  link.append("</x14:id></ext>");

  // `CT_CfRule` allows a single `extLst`, so a captured one is extended
  // with another `<ext>` sibling rather than joined by a second element.
  const std::size_t close = r.ext_lst_raw.has_value() ? r.ext_lst_raw->rfind("</extLst>") : std::string::npos;
  if (close == std::string::npos) {
    // No capture, or a self-closing `<extLst/>` capture with no content
    // worth preserving.
    out.append("<extLst>");
    out.append(link);
    out.append("</extLst>");
    return;
  }
  out.append(r.ext_lst_raw.value(), 0, close);
  out.append(link);
  out.append(r.ext_lst_raw.value(), close, std::string::npos);
}

/// Emits one x14 colour element. The x14 schema names each colour slot
/// with its own element rather than reusing `<color>`, so the element
/// name is a parameter here where `AppendColor` can hard-code it.
void AppendX14Color(std::string& out, std::string_view element, cf::Color c) {
  char buf[16];
  std::snprintf(buf, sizeof(buf), "%02X%02X%02X%02X", c.a, c.r, c.g, c.b);
  out.push_back('<');
  out.append(element);
  out.append(" rgb=\"");
  out.append(buf);
  out.append("\"/>");
}

/// Emits one `<x14:cfvo>`, mirroring the legacy `<cfvo>` it accompanies.
/// The x14 shape differs: the value is an `<xm:f>` child, not a `val`
/// attribute.
void AppendX14Cfvo(std::string& out, const cf::CfValueObject& v) {
  out.append("<x14:cfvo type=\"");
  out.append(CfvoTypeToString(v.type));
  out.push_back('"');
  if (!v.gte) {
    out.append(" gte=\"0\"");
  }
  if (v.value.empty()) {
    out.append("/>");
    return;
  }
  out.append("><xm:f>");
  AppendXmlEscaped(out, v.value);
  out.append("</xm:f></x14:cfvo>");
}

/// Emits the `<x14:conditionalFormatting>` entry holding `r`'s data-bar
/// extension payload, scoped to `sqref` (the enclosing block's range
/// union, restated because the entry is a sibling of the legacy block
/// rather than a child of it).
void AppendX14CfRuleEntry(std::string& out, const cf::CFRule& r, const std::vector<cf::CFCellRange>& sqref) {
  const cf::DataBarSpec& d = r.data_bar.value();
  out.append("<x14:conditionalFormatting xmlns:xm=\"");
  out.append(kXmNs);
  out.append("\"><x14:cfRule type=\"dataBar\" id=\"");
  AppendXmlAttrEscaped(out, r.id);
  // The bar-length bounds are restated because the reader lets the
  // extension win over the legacy element, matching Excel: Excel omits
  // them from `<dataBar>` (which means the pre-2010 defaults 10/90) and
  // states the real bounds only here.
  out.append("\"><x14:dataBar minLength=\"");
  out.append(std::to_string(static_cast<unsigned>(d.min_length_pct)));
  out.append("\" maxLength=\"");
  out.append(std::to_string(static_cast<unsigned>(d.max_length_pct)));
  out.push_back('"');
  if (!d.gradient) {
    out.append(" gradient=\"0\"");
  }
  if (d.border.has_value()) {
    out.append(" border=\"1\"");
  }
  if (d.negative_fill != d.fill) {
    out.append(" negativeBarColorSameAsPositive=\"0\"");
  }
  if (d.negative_border.has_value()) {
    out.append(" negativeBarBorderColorSameAsPositive=\"0\"");
  }
  if (d.axis_position == cf::DataBarAxisPosition::Middle) {
    out.append(" axisPosition=\"middle\"");
  } else if (d.axis_position == cf::DataBarAxisPosition::None) {
    out.append(" axisPosition=\"none\"");
  }
  out.push_back('>');
  // Schema order inside `<x14:dataBar>`: two cfvo, then the colour slots
  // border / negativeFill / negativeBorder / axis.
  AppendX14Cfvo(out, d.min);
  AppendX14Cfvo(out, d.max);
  if (d.border.has_value()) {
    AppendX14Color(out, "x14:borderColor", d.border.value());
  }
  if (d.negative_fill != d.fill) {
    AppendX14Color(out, "x14:negativeFillColor", d.negative_fill);
  }
  if (d.negative_border.has_value()) {
    AppendX14Color(out, "x14:negativeBorderColor", d.negative_border.value());
  }
  if (d.axis_color != kDefaultAxisColor) {
    AppendX14Color(out, "x14:axisColor", d.axis_color);
  }
  out.append("</x14:dataBar></x14:cfRule><xm:sqref>");
  AppendXmlEscaped(out, EncodeSqref(sqref));
  out.append("</xm:sqref></x14:conditionalFormatting>");
}

void AppendCfRule(std::string& out, const cf::CFRule& r, std::size_t dxf_count) {
  out.append("<cfRule type=\"");
  out.append(RuleTypeToString(r.type));
  out.append("\" priority=\"");
  out.append(std::to_string(r.priority));
  out.push_back('"');
  if (r.stop_if_true) {
    out.append(" stopIfTrue=\"1\"");
  }
  // A `dxfId` the package's `<dxfs>` table cannot resolve is dropped
  // rather than emitted dangling — see `cf_writer.h` for why Excel's
  // reaction makes that the cheaper loss.
  if (r.dxf_id.has_value() && r.dxf_id.value() < dxf_count) {
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
  // `CFRule::id` is deliberately not emitted as an attribute here — see
  // `AppendRuleExtLst` for where the linkage actually goes.
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
  AppendRuleExtLst(out, r);

  out.append("</cfRule>");
}

}  // namespace

std::string write_conditional_formattings(const std::vector<cf::ConditionalFormat>& formats, std::size_t dxf_count) {
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
      AppendCfRule(out, rule, dxf_count);
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

std::string build_x14_cf_overlay_entries(const std::vector<cf::ConditionalFormat>& formats) {
  std::string out;
  for (const auto& cf : formats) {
    for (const auto& rule : cf.rules) {
      if (RuleNeedsX14Payload(rule)) {
        AppendX14CfRuleEntry(out, rule, cf.sqref);
      }
    }
  }
  return out;
}

}  // namespace formulon::io
