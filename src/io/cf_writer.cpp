//
// Implementation of the conditional-formatting writer. See cf_writer.h
// for the public contract; see cf_reader.cpp for the symmetric grammar.

#include "io/cf_writer.h"

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "cf/cf_types.h"
#include "cf/cf_x14_id.h"
#include "io/cf_attr_names.h"
#include "io/color_spec_xml.h"
#include "io/future_functions.h"
#include "io/ooxml/external_link_writer.h"
#include "io/ooxml_defs.h"
#include "io/styles_writer.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"
#include "utils/arena.h"
#include "utils/number_text.h"

namespace formulon::io {
namespace {

std::string_view IconSetNameToString(cf::IconSetName n) {
  return cf::kIconSetNames[static_cast<std::size_t>(n)];
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
  const std::size_t mark = out.size();
  const auto fail = [&out, mark]() {
    out.resize(mark);
    return false;
  };
  if (r.is_full_col()) {
    if (!a1::append_column_letters(out, r.first.col)) {
      return fail();
    }
    out.push_back(':');
    if (!a1::append_column_letters(out, r.last.col)) {
      return fail();
    }
  } else if (r.is_full_row()) {
    out.append(std::to_string(r.first.row + 1U));
    out.push_back(':');
    out.append(std::to_string(r.last.row + 1U));
  } else {
    if (!a1::append_column_letters(out, r.first.col)) {
      return fail();
    }
    out.append(std::to_string(r.first.row + 1U));
    if (r.first != r.last) {
      out.push_back(':');
      if (!a1::append_column_letters(out, r.last.col)) {
        return fail();
      }
      out.append(std::to_string(r.last.row + 1U));
    }
  }
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

/// Emits the colour attributes: `theme` / `indexed` / `auto` as written, else
/// the literal `rgb`.
void AppendColorAttrs(std::string& out, const cf::Color& c) {
  const std::uint32_t argb = pack_argb(c.a, c.r, c.g, c.b);
  ColorSpec spec = c.spec;
  if (spec.kind == ColorSpec::Kind::kRgb) {
    spec.kind = ColorSpec::Kind::kNone;  // The literal channels are authoritative for rgb.
  }
  io::append_color_spec_attrs(out, spec, argb);
}

void AppendColor(std::string& out, const cf::Color& c) {
  out.append("<color");
  AppendColorAttrs(out, c);
  out.append("/>");
}

void AppendCfvo(std::string& out, const cf::CfValueObject& v, const parser::ExternalBookIndexer* indexer) {
  out.append("<cfvo type=\"");
  out.append(enum_name(kCfvoTypeNames, v.type, "num"));
  out.push_back('"');
  if (!v.value.empty()) {
    append_xml_attr(out, "val", v.type == cf::CfvoType::Formula ? storage_feature_formula(v.value, indexer) : v.value);
  }
  if (!v.gte) {
    out.append(" gte=\"0\"");
  }
  out.append("/>");
}

void AppendColorScale(std::string& out, const cf::ColorScaleSpec& s, const parser::ExternalBookIndexer* indexer) {
  out.append("<colorScale>");
  for (const auto& th : s.thresholds) {
    AppendCfvo(out, th, indexer);
  }
  for (const auto& c : s.colors) {
    AppendColor(out, c);
  }
  out.append("</colorScale>");
}

/// `linked`: the bar has an x14 counterpart, which carries its lengths.
/// Excel then leaves the legacy lengths at the pre-2010 defaults (10/90,
/// omitted), and so does this writer.
void AppendDataBar(std::string& out, const cf::DataBarSpec& d, bool linked,
                   const parser::ExternalBookIndexer* indexer) {
  out.append("<dataBar");
  if (!linked) {
    append_xml_attr_uint(out, "minLength", d.min_length_pct);
    append_xml_attr_uint(out, "maxLength", d.max_length_pct);
  }
  if (!d.show_value) {
    out.append(" showValue=\"0\"");
  }
  out.push_back('>');
  AppendCfvo(out, d.min, indexer);
  AppendCfvo(out, d.max, indexer);
  AppendColor(out, d.fill);
  out.append("</dataBar>");
}

void AppendIconSet(std::string& out, const cf::IconSetSpec& i, const parser::ExternalBookIndexer* indexer) {
  out.append("<iconSet");
  // Excel omits the attribute for the schema default, 3TrafficLights1.
  if (i.name != cf::IconSetName::Three_TrafficLights1) {
    append_xml_attr(out, "iconSet", IconSetNameToString(i.name));
  }
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
  AppendCfvo(out, i.floor, indexer);
  for (const auto& th : i.thresholds) {
    AppendCfvo(out, th, indexer);
  }
  out.append("</iconSet>");
}

bool CfvoReferencesOtherSheet(const cf::CfValueObject& v, std::string_view own_sheet,
                              const std::vector<std::string>* sheet_list) {
  return v.type == cf::CfvoType::Formula && parser::formula_references_other_sheet(v.value, own_sheet, sheet_list);
}

/// True when `r` is written only in the x14 extension: a formula or `cfvo`
/// of it names a sheet other than `own_sheet`, spellings resolved among `sheet_list`.
bool RuleIsX14Only(const cf::CFRule& r, std::string_view own_sheet, const std::vector<std::string>* sheet_list) {
  if (own_sheet.empty()) {
    return false;
  }
  for (const std::optional<std::string>* formula : {&r.formula1, &r.formula2}) {
    if (formula->has_value() && parser::formula_references_other_sheet(**formula, own_sheet, sheet_list)) {
      return true;
    }
  }
  if (r.color_scale.has_value()) {
    for (const cf::CfValueObject& v : r.color_scale->thresholds) {
      if (CfvoReferencesOtherSheet(v, own_sheet, sheet_list)) {
        return true;
      }
    }
  }
  if (r.data_bar.has_value() && (CfvoReferencesOtherSheet(r.data_bar->min, own_sheet, sheet_list) ||
                                 CfvoReferencesOtherSheet(r.data_bar->max, own_sheet, sheet_list))) {
    return true;
  }
  if (r.icon_set.has_value()) {
    if (CfvoReferencesOtherSheet(r.icon_set->floor, own_sheet, sheet_list)) {
      return true;
    }
    for (const cf::CfValueObject& v : r.icon_set->thresholds) {
      if (CfvoReferencesOtherSheet(v, own_sheet, sheet_list)) {
        return true;
      }
    }
  }
  return false;
}

/// The legacy `operator` hint Excel writes on the text rule types, or null
/// for any other type.
const char* TextRuleOperator(cf::RuleType type) {
  switch (type) {
    case cf::RuleType::ContainsText:
      return "containsText";
    case cf::RuleType::NotContainsText:
      return "notContains";
    case cf::RuleType::BeginsWith:
      return "beginsWith";
    case cf::RuleType::EndsWith:
      return "endsWith";
    default:
      return nullptr;
  }
}

/// URI of the `<ext>` that links a legacy `<cfRule>` to its x14
/// counterpart, and the two namespaces the extension content lives in.
/// All three are fixed constants Excel matches literally.
constexpr std::string_view kX14CfRuleExtUri = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}";

/// The axis colour a data bar has when the x14 extension says nothing —
/// mirrors `DataBarSpec::axis_color`'s in-class initializer.
constexpr cf::Color kDefaultAxisColor{0, 0, 0, 255};

/// True when `d` renders differently from a legacy `<dataBar>` (no axis,
/// one fill colour, gradient), so the rule needs an x14 counterpart to
/// survive a save. False for a legacy-shaped bar, which keeps a
/// legacy-only file legacy-only across a round trip.
bool NeedsX14DataBarPayload(const cf::DataBarSpec& d) {
  return d.border.has_value() || d.negative_border.has_value() || d.negative_fill != d.fill ||
         d.axis_position != cf::DataBarAxisPosition::None || !d.gradient || d.axis_color != kDefaultAxisColor ||
         d.direction != cf::DataBarDirection::Context;
}

/// True when the rule both needs an x14 counterpart and has an id to
/// link it by. An id is required: the link and the payload find each
/// other by that GUID and nothing else.
bool RuleNeedsX14Payload(const cf::CFRule& r) {
  if (r.id.empty()) {
    return false;
  }
  return (r.data_bar.has_value() && NeedsX14DataBarPayload(r.data_bar.value())) ||
         (r.icon_set.has_value() && !r.icon_set->custom_icons.empty());
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
  link.append(kNsX14);
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
void AppendX14Color(std::string& out, std::string_view element, const cf::Color& c) {
  out.push_back('<');
  out.append(element);
  AppendColorAttrs(out, c);
  out.append("/>");
}

/// Emits one `<x14:cfvo>`, mirroring the legacy `<cfvo>` it accompanies.
/// The x14 shape differs: the value is an `<xm:f>` child, not a `val`
/// attribute.
void AppendX14Cfvo(std::string& out, const cf::CfValueObject& v, const parser::ExternalBookIndexer* indexer) {
  out.append("<x14:cfvo type=\"");
  out.append(enum_name(kCfvoTypeNames, v.type, "num"));
  out.push_back('"');
  if (!v.gte) {
    out.append(" gte=\"0\"");
  }
  if (v.value.empty()) {
    out.append("/>");
    return;
  }
  out.append("><xm:f>");
  AppendXmlEscaped(out, v.type == cf::CfvoType::Formula ? storage_feature_formula(v.value, indexer) : v.value);
  out.append("</xm:f></x14:cfvo>");
}

/// Emits `d`'s `<x14:dataBar>` element. The bar lengths are always
/// stated here: the legacy element of a linked bar leaves them at the
/// pre-2010 defaults, as Excel does. A `standalone` bar has no legacy
/// element, so its `showValue` and fill colour are stated here too.
void AppendX14DataBar(std::string& out, const cf::DataBarSpec& d, const parser::ExternalBookIndexer* indexer,
                      bool standalone) {
  out.append("<x14:dataBar");
  append_xml_attr_uint(out, "minLength", d.min_length_pct);
  append_xml_attr_uint(out, "maxLength", d.max_length_pct);
  if (standalone && !d.show_value) {
    out.append(" showValue=\"0\"");
  }
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
  if (d.direction == cf::DataBarDirection::LeftToRight) {
    out.append(" direction=\"leftToRight\"");
  } else if (d.direction == cf::DataBarDirection::RightToLeft) {
    out.append(" direction=\"rightToLeft\"");
  }
  out.push_back('>');
  // Schema order inside `<x14:dataBar>`: two cfvo, then the colour slots
  // fill / border / negativeFill / negativeBorder / axis.
  AppendX14Cfvo(out, d.min, indexer);
  AppendX14Cfvo(out, d.max, indexer);
  if (standalone) {
    AppendX14Color(out, "x14:fillColor", d.fill);
  }
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
  out.append("</x14:dataBar>");
}

/// Emits `i`'s `<x14:iconSet>` element: the legacy set's thresholds restated
/// as `<x14:cfvo>` plus one `<x14:cfIcon>` per level.
void AppendX14IconSet(std::string& out, const cf::IconSetSpec& i, const parser::ExternalBookIndexer* indexer) {
  out.append("<x14:iconSet");
  append_xml_attr(out, "iconSet", IconSetNameToString(i.name));
  if (!i.custom_icons.empty()) {
    out.append(" custom=\"1\"");
  }
  if (!i.show_value) {
    out.append(" showValue=\"0\"");
  }
  if (!i.percent) {
    out.append(" percent=\"0\"");
  }
  if (i.reverse) {
    out.append(" reverse=\"1\"");
  }
  out.push_back('>');
  AppendX14Cfvo(out, i.floor, indexer);
  for (const auto& th : i.thresholds) {
    AppendX14Cfvo(out, th, indexer);
  }
  for (const cf::CfIcon& icon : i.custom_icons) {
    out.append("<x14:cfIcon");
    append_xml_attr(out, "iconSet", icon.icon_set);
    append_xml_attr_int(out, "iconId", icon.icon_id);
    out.append("/>");
  }
  out.append("</x14:iconSet>");
}

/// Emits the `<x14:conditionalFormatting>` entry holding `r`'s extension
/// payload (data bar or custom icon set), scoped to `sqref` (the enclosing
/// block's range union, restated because the entry is a sibling of the legacy
/// block rather than a child of it).
void AppendX14CfRuleEntry(std::string& out, const cf::CFRule& r, const std::vector<cf::CFCellRange>& sqref,
                          const parser::ExternalBookIndexer* indexer) {
  out.append("<x14:conditionalFormatting xmlns:xm=\"");
  out.append(kNsExcelMain);
  const bool icon_set = r.icon_set.has_value() && !r.icon_set->custom_icons.empty();
  out.append(icon_set ? "\"><x14:cfRule type=\"iconSet\" priority=\"" : "\"><x14:cfRule type=\"dataBar\" id=\"");
  if (icon_set) {
    out.append(std::to_string(r.priority));
    out.append("\" id=\"");
  }
  AppendXmlAttrEscaped(out, r.id);
  out.append("\">");
  if (icon_set) {
    AppendX14IconSet(out, *r.icon_set, indexer);
  } else {
    AppendX14DataBar(out, r.data_bar.value(), indexer, /*standalone=*/false);
  }
  out.append("</x14:cfRule><xm:sqref>");
  AppendXmlEscaped(out, EncodeSqref(sqref));
  out.append("</xm:sqref></x14:conditionalFormatting>");
}

/// Emits `<x14:dxf>` holding the body of the `<dxf>` `r.dxf_id` names, or an
/// empty `<x14:dxf/>` when the rule sets no formatting.
void AppendX14Dxf(std::string& out, const cf::CFRule& r, const std::vector<DifferentialFormat>* dxfs) {
  if (dxfs != nullptr && r.dxf_id.has_value() && *r.dxf_id < dxfs->size()) {
    const std::string body = dxf_body_fragment((*dxfs)[*r.dxf_id]);
    if (!body.empty()) {
      out.append("<x14:dxf>");
      out.append(body);
      out.append("</x14:dxf>");
      return;
    }
  }
  out.append("<x14:dxf/>");
}

/// Emits `r` whole as an `<x14:conditionalFormatting>` entry under `id`, in
/// the shape Excel stores a rule that names another sheet: the attributes in
/// Excel's order, `<xm:f>` formulas and an inline `<x14:dxf>`, or the x14 form
/// of a colour scale, data bar or icon set.
void AppendX14OnlyRuleEntry(std::string& out, const cf::CFRule& r, std::string_view id,
                            const cf::ConditionalFormat& block, const CfSheetContext& sheet,
                            const parser::ExternalBookIndexer* indexer) {
  out.append("<x14:conditionalFormatting xmlns:xm=\"");
  out.append(kNsExcelMain);
  out.push_back('"');
  if (block.pivot_scope) {
    out.append(" pivot=\"1\"");
  }
  out.append("><x14:cfRule type=\"");
  out.append(enum_name(kCfRuleTypeNames, r.type, "expression"));
  out.append("\" priority=\"");
  out.append(std::to_string(r.priority));
  out.push_back('"');
  if (r.stop_if_true) {
    out.append(" stopIfTrue=\"1\"");
  }
  if (r.type == cf::RuleType::AboveAverage && !r.above_average) {
    out.append(" aboveAverage=\"0\"");
  }
  if (r.type == cf::RuleType::Top10) {
    if (r.percent) {
      out.append(" percent=\"1\"");
    }
    if (r.bottom) {
      out.append(" bottom=\"1\"");
    }
  }
  const char* text_operator = TextRuleOperator(r.type);
  if (r.type == cf::RuleType::CellIs && r.op.has_value()) {
    append_xml_attr(out, "operator", enum_name(kCfCellIsOperatorNames, r.op.value(), "equal"));
  } else if (text_operator != nullptr && r.text.has_value()) {
    append_xml_attr(out, "operator", text_operator);
  }
  if (text_operator != nullptr && r.text.has_value()) {
    append_xml_attr(out, "text", r.text.value());
  }
  if (r.type == cf::RuleType::TimePeriod && r.time_period.has_value()) {
    append_xml_attr(out, "timePeriod", enum_name(kCfTimePeriodNames, r.time_period.value(), "today"));
  }
  if (r.type == cf::RuleType::Top10 && r.rank.has_value()) {
    append_xml_attr_int(out, "rank", r.rank.value());
  }
  if (r.type == cf::RuleType::AboveAverage) {
    if (r.std_dev.has_value()) {
      char buf[32];
      format_general(buf, sizeof(buf), r.std_dev.value(), 6);
      append_xml_attr(out, "stdDev", buf);
    }
    if (r.equal_average) {
      out.append(" equalAverage=\"1\"");
    }
  }
  append_xml_attr(out, "id", id);
  out.push_back('>');
  if (r.color_scale.has_value()) {
    out.append("<x14:colorScale>");
    for (const cf::CfValueObject& th : r.color_scale->thresholds) {
      AppendX14Cfvo(out, th, indexer);
    }
    for (const cf::Color& c : r.color_scale->colors) {
      AppendX14Color(out, "x14:color", c);
    }
    out.append("</x14:colorScale>");
  } else if (r.data_bar.has_value()) {
    AppendX14DataBar(out, *r.data_bar, indexer, /*standalone=*/true);
  } else if (r.icon_set.has_value()) {
    AppendX14IconSet(out, *r.icon_set, indexer);
  } else {
    for (const std::optional<std::string>* formula : {&r.formula1, &r.formula2}) {
      if (formula->has_value()) {
        out.append("<xm:f>");
        AppendXmlEscaped(out, storage_feature_formula(**formula, indexer));
        out.append("</xm:f>");
      }
    }
    AppendX14Dxf(out, r, sheet.dxfs);
  }
  out.append("</x14:cfRule><xm:sqref>");
  AppendXmlEscaped(out, EncodeSqref(block.sqref));
  out.append("</xm:sqref></x14:conditionalFormatting>");
}

void AppendCfRule(std::string& out, const cf::CFRule& r, std::size_t dxf_count,
                  const parser::ExternalBookIndexer* indexer) {
  out.append("<cfRule type=\"");
  out.append(enum_name(kCfRuleTypeNames, r.type, "expression"));
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
    append_xml_attr_uint(out, "dxfId", r.dxf_id.value());
  }
  if (r.type == cf::RuleType::CellIs && r.op.has_value()) {
    append_xml_attr(out, "operator", enum_name(kCfCellIsOperatorNames, r.op.value(), "equal"));
  }
  if ((r.type == cf::RuleType::ContainsText || r.type == cf::RuleType::NotContainsText ||
       r.type == cf::RuleType::BeginsWith || r.type == cf::RuleType::EndsWith) &&
      r.text.has_value()) {
    // Excel emits operator on these rule types as a hint to the
    // legacy reader (e.g. `operator="containsText"`); the modern
    // reader keys solely off `type`. Emit it for parity with Excel-
    // saved files.
    if (const char* op_attr = TextRuleOperator(r.type); op_attr != nullptr) {
      append_xml_attr(out, "operator", op_attr);
    }
    append_xml_attr(out, "text", r.text.value());
  }
  if (r.type == cf::RuleType::Top10) {
    if (r.rank.has_value()) {
      append_xml_attr_int(out, "rank", r.rank.value());
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
      format_general(buf, sizeof(buf), r.std_dev.value(), 6);
      append_xml_attr(out, "stdDev", buf);
    }
  }
  if (r.type == cf::RuleType::TimePeriod && r.time_period.has_value()) {
    append_xml_attr(out, "timePeriod", enum_name(kCfTimePeriodNames, r.time_period.value(), "today"));
  }
  // `CFRule::id` is deliberately not emitted as an attribute here — see
  // `AppendRuleExtLst` for where the linkage actually goes.
  out.push_back('>');

  if (r.formula1.has_value()) {
    out.append("<formula>");
    AppendXmlEscaped(out, storage_feature_formula(r.formula1.value(), indexer));
    out.append("</formula>");
  }
  if (r.formula2.has_value()) {
    out.append("<formula>");
    AppendXmlEscaped(out, storage_feature_formula(r.formula2.value(), indexer));
    out.append("</formula>");
  }
  if (r.color_scale.has_value()) {
    AppendColorScale(out, r.color_scale.value(), indexer);
  }
  if (r.data_bar.has_value()) {
    AppendDataBar(out, r.data_bar.value(), RuleNeedsX14Payload(r) || HasCapturedX14Link(r), indexer);
  }
  if (r.icon_set.has_value()) {
    AppendIconSet(out, r.icon_set.value(), indexer);
  }
  AppendRuleExtLst(out, r);

  out.append("</cfRule>");
}

}  // namespace

std::string write_conditional_formattings(const std::vector<cf::ConditionalFormat>& formats, std::size_t dxf_count,
                                          const parser::ExternalBookIndexer* indexer, const CfSheetContext& sheet) {
  if (formats.empty()) {
    return std::string{};
  }
  std::string out;
  out.reserve(formats.size() * 128);
  std::string rules;
  for (const auto& cf : formats) {
    rules.clear();
    for (const auto& rule : cf.rules) {
      if (!RuleIsX14Only(rule, sheet.sheet_name, sheet.sheet_names)) {
        AppendCfRule(rules, rule, dxf_count, indexer);
      }
    }
    if (rules.empty() && !cf.rules.empty()) {
      continue;
    }
    out.append("<conditionalFormatting sqref=\"");
    AppendXmlAttrEscaped(out, EncodeSqref(cf.sqref));
    out.push_back('"');
    if (cf.pivot_scope) {
      out.append(" pivot=\"1\"");
    }
    out.push_back('>');
    out.append(rules);
    // `CT_ConditionalFormatting`'s schema-trailing `extLst?`, round-tripped
    // byte-for-byte (see `ConditionalFormat::ext_lst_raw`).
    if (cf.ext_lst_raw.has_value()) {
      out.append(cf.ext_lst_raw.value());
    }
    out.append("</conditionalFormatting>");
  }
  return out;
}

std::string storage_feature_formula(std::string_view formula, const parser::ExternalBookIndexer* indexer) {
  Arena arena;
  parser::Parser parser(formula, arena);
  const parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    return std::string(formula);
  }
  const std::vector<const parser::AstNode*> function_values = xlsb::function_value_refs(*root, {});
  std::string storage =
      parser::format_formula_storage(*root, &storage_call_name, /*omitted_at=*/nullptr, indexer, &function_values);
  // A classic formula is written verbatim, keeping its exact spelling.
  return NeedsStorageSpelling(*root, storage) ? storage : std::string(formula);
}

bool data_bar_needs_x14(const cf::DataBarSpec& bar) {
  return NeedsX14DataBarPayload(bar);
}

std::string build_x14_data_bar_element(const cf::DataBarSpec& bar) {
  std::string out;
  AppendX14DataBar(out, bar, /*indexer=*/nullptr, /*standalone=*/false);
  return out;
}

std::string build_x14_icon_set_element(const cf::IconSetSpec& icon_set, const parser::ExternalBookIndexer* indexer) {
  std::string out;
  AppendX14IconSet(out, icon_set, indexer);
  return out;
}

std::string build_x14_cf_overlay_entries(const std::vector<cf::ConditionalFormat>& formats,
                                         const parser::ExternalBookIndexer* indexer, const CfSheetContext& sheet,
                                         const std::unordered_set<std::string>& reserved_ids) {
  std::unordered_set<std::string> used_ids = reserved_ids;
  for (const auto& block : formats) {
    for (const auto& rule : block.rules) {
      if (!rule.id.empty()) {
        used_ids.insert(rule.id);
      }
    }
  }
  std::string out;
  for (const auto& block : formats) {
    for (const auto& rule : block.rules) {
      if (RuleIsX14Only(rule, sheet.sheet_name, sheet.sheet_names)) {
        std::string id = rule.id;
        if (id.empty()) {
          const std::uint32_t ordinal = rule.priority > 0 ? static_cast<std::uint32_t>(rule.priority) : 0U;
          id = cf::x14_rule_id(sheet.sheet_index, ordinal, used_ids);
          used_ids.insert(id);
        }
        AppendX14OnlyRuleEntry(out, rule, id, block, sheet, indexer);
      } else if (RuleNeedsX14Payload(rule)) {
        AppendX14CfRuleEntry(out, rule, block.sqref, indexer);
      }
    }
  }
  return out;
}

}  // namespace formulon::io
