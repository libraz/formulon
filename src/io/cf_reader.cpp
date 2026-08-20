//
// Implementation of the conditional-formatting reader. See cf_reader.h
// for the public contract.

#include "io/cf_reader.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "io/cell_parser.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/structured_log.h"

namespace formulon::io {
namespace {

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

/// Strips `$` absolute markers from an A1 token. CF `sqref` legitimately
/// carries absolute references (`$A$1:$A$10`), which `parse_a1` rejects;
/// the column/row position is identical with or without the markers, so
/// dropping them is a lossless normalisation. The result is copied into
/// `buf` (kept alive by the caller) and returned as a view over it.
std::string_view StripAbsoluteMarkers(std::string_view ref, std::string& buf) {
  if (ref.find('$') == std::string_view::npos) {
    return ref;
  }
  buf.clear();
  buf.reserve(ref.size());
  for (const char ch : ref) {
    if (ch != '$') {
      buf.push_back(ch);
    }
  }
  return buf;
}

/// Decodes one A1 cell-range token (`A1`, `A1:B5`, `$A$1:$B$5`) into a
/// `CFCellRange`. Single-cell tokens land as `first == last`. Absolute
/// markers (`$`) are stripped first since CF sqref legitimately contains
/// them. Returns `kIoSheetCorrupt` for unparseable input — the caller
/// either folds it into the surrounding sqref error or skips the block.
// Decodes a run of column letters (`A`, `AB`, `XFD`) to a 1-based column
// index, or 0 when the run is empty / non-alpha / out of range. Only
// upper-case letters are accepted, matching `parse_a1` (which decodes the
// cell-shaped tokens of the same sqref) and Excel, which never emits a
// lower-case column letter. Both halves of `ParseA1Range` therefore apply
// one case rule.
std::uint32_t DecodeColumnRun(std::string_view s) {
  if (s.empty() || s.size() > 3) {
    return 0;
  }
  std::uint32_t col = 0;
  for (char ch : s) {
    if (ch < 'A' || ch > 'Z') {
      return 0;
    }
    col = col * 26U + static_cast<std::uint32_t>(ch - 'A' + 1);
    if (col > Sheet::kMaxCols) {
      return 0;
    }
  }
  return col;
}

// Decodes a run of digits to a 1-based row index, or 0 when the run is
// empty / non-digit / out of range.
std::uint32_t DecodeRowRun(std::string_view s) {
  if (s.empty() || s.size() > 7) {
    return 0;
  }
  std::uint64_t row = 0;
  for (char ch : s) {
    if (ch < '0' || ch > '9') {
      return 0;
    }
    row = row * 10U + static_cast<std::uint32_t>(ch - '0');
  }
  if (row == 0 || row > Sheet::kMaxRows) {
    return 0;
  }
  return static_cast<std::uint32_t>(row);
}

Expected<cf::CFCellRange, Error> ParseA1Range(std::string_view ref) {
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    std::string norm_buf;
    auto rc = parse_a1(StripAbsoluteMarkers(ref, norm_buf));
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
  std::string a_buf;
  std::string b_buf;
  const std::string_view a = StripAbsoluteMarkers(ref.substr(0, colon), a_buf);
  const std::string_view b = StripAbsoluteMarkers(ref.substr(colon + 1), b_buf);

  // Whole-column (`A:A`, `A:C`) sqref: both endpoints are column-letter runs
  // with no row. Store the full logical row extent and flag `full_col` so
  // membership tests work while range-aware evaluation clamps to the used
  // range and the writer re-emits the compact form.
  const std::uint32_t a_col = DecodeColumnRun(a);
  const std::uint32_t b_col = DecodeColumnRun(b);
  if (a_col != 0 && b_col != 0) {
    cf::CFCellRange out{};
    out.first = {0, (a_col < b_col ? a_col : b_col) - 1U};
    out.last = {Sheet::kMaxRows - 1U, (a_col < b_col ? b_col : a_col) - 1U};
    return out;
  }
  // Whole-row (`1:1`, `1:3`) sqref: both endpoints are digit runs with no
  // column.
  const std::uint32_t a_row = DecodeRowRun(a);
  const std::uint32_t b_row = DecodeRowRun(b);
  if (a_row != 0 && b_row != 0) {
    cf::CFCellRange out{};
    out.first = {(a_row < b_row ? a_row : b_row) - 1U, 0};
    out.last = {(a_row < b_row ? b_row : a_row) - 1U, Sheet::kMaxCols - 1U};
    return out;
  }

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

/// Parses a `<color rgb="AARRGGBB">` attribute into a `cf::Color`. The
/// alpha channel is optional in the spec (some Excel exports drop it
/// for fully-opaque colours, emitting `RRGGBB` only); both 6- and 8-hex
/// forms are accepted. Returns opaque black on malformed input — the
/// caller is free to drop the rule, but in practice CF colours are
/// either valid or absent.
///
/// The hex-decoding loop is shared with `styles_reader.cpp` via
/// `parse_rgb_hex()` in `xml_utils.h`; this wrapper only unpacks the
/// packed `0xAARRGGBB` value into the `cf::Color` channels.
cf::Color ParseRgbColor(std::string_view rgb) {
  // 0xFF000000 (opaque black) preserves the legacy fallback for every
  // malformed case: empty string, wrong length, or non-hex character.
  const std::uint32_t packed = parse_rgb_hex(rgb, 0xFF000000U);
  cf::Color out{};
  out.a = static_cast<std::uint8_t>((packed >> 24U) & 0xFFU);
  out.r = static_cast<std::uint8_t>((packed >> 16U) & 0xFFU);
  out.g = static_cast<std::uint8_t>((packed >> 8U) & 0xFFU);
  out.b = static_cast<std::uint8_t>(packed & 0xFFU);
  return out;
}

cf::CfValueObject ReadCfvo(const pugi::xml_node& cfvo) {
  cf::CfValueObject out{};
  out.type = ParseCfvoType(attr_str(cfvo, "type"));
  out.value = attr_str(cfvo, "val");
  // `gte` is "1" by default; only an explicit "0" / "false" toggles
  // it off. The OOXML schema only emits those two spellings, so an
  // unrecognised token (defensive case) is folded back to the default
  // `true` to preserve the legacy lenient behaviour.
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
      out->colors.push_back(ParseRgbColor(attr_str(child, "rgb")));
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
      // as the positive fill. Negative fill, axis colour/position and
      // gradient live only in the 2010+ `<x14:dataBar>` extension, which
      // `ApplyX14DataBarOverlay` folds on top of this.
      if (color_idx == 0) {
        out->fill = ParseRgbColor(attr_str(child, "rgb"));
        out->negative_fill = out->fill;
      }
      ++color_idx;
    }
  }
  // Note: malformed numeric input falls back to def=0, matching the
  // legacy `std::strtol(_, nullptr, 10)` return-on-error of 0 (which
  // happens to be in-range and would have stored min/max=0).
  if (bar.attribute("minLength")) {
    const std::int32_t parsed = attr_i32(bar, "minLength", 0);
    if (parsed >= 0 && parsed <= 100) {
      out->min_length_pct = static_cast<std::uint8_t>(parsed);
    }
  }
  if (bar.attribute("maxLength")) {
    const std::int32_t parsed = attr_i32(bar, "maxLength", 0);
    if (parsed >= 0 && parsed <= 100) {
      out->max_length_pct = static_cast<std::uint8_t>(parsed);
    }
  }
  if (const pugi::xml_attribute show_value = bar.attribute("showValue"); show_value) {
    out->show_value = parse_xml_bool_attr(show_value);
  }
}

/// Overlays a `<x14:dataBar>` extension element (Excel 2010+) onto a
/// `DataBarSpec` already populated from the legacy `<dataBar>` element.
/// The x14 extension is the only place negative-fill / negative-border /
/// axis colour+position / gradient-vs-solid are expressed; the legacy
/// schema has no attributes for them, so the pre-overlay `DataBarSpec`
/// carries only fallback values (`negative_fill == fill`,
/// `axis_position == Automatic`, `gradient == true`).
void ApplyX14DataBarOverlay(const pugi::xml_node& x14_bar, cf::DataBarSpec* out) {
  if (const pugi::xml_attribute gradient = x14_bar.attribute("gradient"); gradient) {
    out->gradient = parse_xml_bool_attr(gradient);
  }
  const std::string_view axis_position = attr_str(x14_bar, "axisPosition");
  if (axis_position == "middle") {
    out->axis_position = cf::DataBarAxisPosition::Middle;
  } else if (axis_position == "none") {
    out->axis_position = cf::DataBarAxisPosition::None;
  } else {
    // "automatic", absent, or any unrecognised token: fold to the
    // schema default.
    out->axis_position = cf::DataBarAxisPosition::Automatic;
  }
  for (pugi::xml_node child = x14_bar.first_child(); child; child = child.next_sibling()) {
    const std::string_view name = child.name();
    if (name == "x14:negativeFillColor") {
      out->negative_fill = ParseRgbColor(attr_str(child, "rgb"));
    } else if (name == "x14:negativeBorderColor") {
      out->negative_border = ParseRgbColor(attr_str(child, "rgb"));
    } else if (name == "x14:axisColor") {
      out->axis_color = ParseRgbColor(attr_str(child, "rgb"));
    } else if (name == "x14:borderColor") {
      out->border = ParseRgbColor(attr_str(child, "rgb"));
    }
  }
  // `negativeBarColorSameAsPositive="1"` overrides any explicit
  // `<x14:negativeFillColor>` sibling (Excel does not emit both, but a
  // hand-edited file could; the flag wins per the CT_DataBar schema).
  if (attr_bool(x14_bar, "negativeBarColorSameAsPositive", false)) {
    out->negative_fill = out->fill;
  }
  // Bar-length bounds are restated here and win over the legacy element.
  // Excel writes a legacy `<dataBar>` with no minLength/maxLength -- which
  // means the pre-2010 defaults 10/90 -- while stating 0/100 in the
  // extension, and renders 0/100. Taking the legacy defaults would report
  // a bar length Excel does not draw.
  if (const pugi::xml_attribute min_length = x14_bar.attribute("minLength"); min_length) {
    const std::int32_t parsed = attr_i32(x14_bar, "minLength", 0);
    if (parsed >= 0 && parsed <= 100) {
      out->min_length_pct = static_cast<std::uint8_t>(parsed);
    }
  }
  if (const pugi::xml_attribute max_length = x14_bar.attribute("maxLength"); max_length) {
    const std::int32_t parsed = attr_i32(x14_bar, "maxLength", 0);
    if (parsed >= 0 && parsed <= 100) {
      out->max_length_pct = static_cast<std::uint8_t>(parsed);
    }
  }
}

/// Returns the x14 rule id a legacy `<cfRule>` points at, or an empty
/// string when it points at none.
///
/// Excel expresses the link as a nested
/// `<extLst><ext uri="{B025F937-...}"><x14:id>{GUID}</x14:id>`, NOT as an
/// `id` attribute on the `<cfRule>` itself. Both spellings appear in the
/// wild, but only the nested one survives Excel: opening a file that uses
/// the attribute form and saving it back strips the attribute and drops
/// the orphaned worksheet-level x14 block entirely (measured against
/// Excel 365, macOS, 2026-08-15). The attribute is still accepted here as
/// a lenient fallback -- it costs one lookup and keeps hand-written files
/// working -- but the nested form is what real input uses and therefore
/// what decides.
std::string ReadX14RuleId(const pugi::xml_node& rule) {
  for (pugi::xml_node ext_lst = rule.child("extLst"); ext_lst; ext_lst = ext_lst.next_sibling("extLst")) {
    for (pugi::xml_node ext = ext_lst.child("ext"); ext; ext = ext.next_sibling("ext")) {
      if (const pugi::xml_node id = ext.child("x14:id"); id) {
        if (const std::string text(id.child_value()); !text.empty()) {
          return text;
        }
      }
    }
  }
  return std::string(attr_str(rule, "id"));
}

/// Builds an `id -> <x14:dataBar>` lookup from the worksheet-level
/// `<extLst><ext><x14:conditionalFormattings>` overlay (Excel 2010+).
/// Each `<x14:cfRule type="dataBar" id="{GUID}">` cross-references a
/// legacy `<cfRule id="{GUID}">` (see `CFRule::id`); `read_conditional_
/// formats` applies the match via `ApplyX14DataBarOverlay`. Returns an
/// empty map when the sheet carries no such overlay (pre-2010 files, or
/// files with no DataBar rules at all).
std::unordered_map<std::string, pugi::xml_node> CollectX14DataBarOverlay(const pugi::xml_node& worksheet) {
  std::unordered_map<std::string, pugi::xml_node> out;
  const pugi::xml_node ext_lst = worksheet.child("extLst");
  if (!ext_lst) {
    return out;
  }
  for (pugi::xml_node ext = ext_lst.child("ext"); ext; ext = ext.next_sibling("ext")) {
    const pugi::xml_node formattings = ext.child("x14:conditionalFormattings");
    if (!formattings) {
      continue;
    }
    for (pugi::xml_node block = formattings.child("x14:conditionalFormatting"); block;
         block = block.next_sibling("x14:conditionalFormatting")) {
      for (pugi::xml_node rule = block.child("x14:cfRule"); rule; rule = rule.next_sibling("x14:cfRule")) {
        if (attr_str(rule, "type") != "dataBar") {
          continue;
        }
        const pugi::xml_node bar = rule.child("x14:dataBar");
        const std::string id(attr_str(rule, "id"));
        if (bar && !id.empty()) {
          out.emplace(id, bar);
        }
      }
    }
  }
  return out;
}

void ReadIconSet(const pugi::xml_node& iset, cf::IconSetSpec* out) {
  out->name = ParseIconSetName(attr_str(iset, "iconSet"));
  // OOXML defaults: reverse=false, showValue=true, percent=true. The
  // current struct defaults map directly through `attr_bool(_, _, def)`.
  out->reverse = attr_bool(iset, "reverse", out->reverse);
  out->show_value = attr_bool(iset, "showValue", out->show_value);
  out->percent = attr_bool(iset, "percent", out->percent);
  // OOXML emits N `<cfvo>` children for an N-icon set. The first one is
  // the floor of the lowest icon's bucket (conventionally `type="percent"
  // val="0"`) and carries no boundary of its own — `IconSetSpec::thresholds`
  // stores only the N-1 real boundaries, matching `resolve_icon_set()`'s
  // bucket model. The writer re-synthesizes the floor cfvo on output.
  bool skipped_floor = false;
  for (pugi::xml_node cfvo = iset.child("cfvo"); cfvo; cfvo = cfvo.next_sibling("cfvo")) {
    if (!skipped_floor) {
      skipped_floor = true;
      continue;
    }
    out->thresholds.push_back(ReadCfvo(cfvo));
  }
}

cf::CFRule ReadCfRule(const pugi::xml_node& rule) {
  cf::CFRule out;
  out.type = ParseRuleType(attr_str(rule, "type"));
  out.priority = attr_i32(rule, "priority", 1);
  out.stop_if_true = attr_bool(rule, "stopIfTrue");
  if (rule.attribute("dxfId")) {
    out.dxf_id = static_cast<std::uint32_t>(attr_i32(rule, "dxfId", 0));
  }
  out.id = ReadX14RuleId(rule);

  if (out.type == cf::RuleType::CellIs) {
    out.op = ParseCellIsOperator(attr_str(rule, "operator"));
  }
  if (out.type == cf::RuleType::TimePeriod) {
    out.time_period = ParseTimePeriod(attr_str(rule, "timePeriod"));
  }
  if (out.type == cf::RuleType::Top10) {
    out.rank = attr_i32(rule, "rank", 10);
    out.percent = attr_bool(rule, "percent");
    out.bottom = attr_bool(rule, "bottom");
  }
  if (out.type == cf::RuleType::AboveAverage) {
    // `aboveAverage` defaults to true in the OOXML spec.
    out.above_average = attr_bool(rule, "aboveAverage", true);
    out.equal_average = attr_bool(rule, "equalAverage");
    if (rule.attribute("stdDev")) {
      out.std_dev = attr_f64(rule, "stdDev");
    }
  }
  if (out.type == cf::RuleType::ContainsText || out.type == cf::RuleType::NotContainsText ||
      out.type == cf::RuleType::BeginsWith || out.type == cf::RuleType::EndsWith) {
    out.text = attr_str(rule, "text");
  }

  // Children: <formula>, <colorScale>, <dataBar>, <iconSet> (mutually
  // exclusive at the visual-payload level).
  //
  // `cf::CFRule` declares at most one of the three visual payloads
  // engaged, and its consumers rely on that. A document spelling several
  // in one `<cfRule>` therefore keeps the payload the rule's `type`
  // names, or the first one present when `type` names none of them.
  bool visual_engaged = false;
  const auto engage_visual = [&out, &visual_engaged](cf::RuleType kind) {
    if (visual_engaged && out.type != kind) {
      return false;
    }
    out.color_scale.reset();
    out.data_bar.reset();
    out.icon_set.reset();
    visual_engaged = true;
    return true;
  };
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
      if (engage_visual(cf::RuleType::ColorScale)) {
        cf::ColorScaleSpec spec;
        ReadColorScale(child, &spec);
        out.color_scale = std::move(spec);
      }
    } else if (name == "dataBar") {
      if (engage_visual(cf::RuleType::DataBar)) {
        cf::DataBarSpec spec;
        ReadDataBar(child, &spec);
        out.data_bar = std::move(spec);
      }
    } else if (name == "iconSet") {
      if (engage_visual(cf::RuleType::IconSet)) {
        cf::IconSetSpec spec;
        ReadIconSet(child, &spec);
        out.icon_set = std::move(spec);
      }
    }
  }

  // `CT_CfRule`'s schema-trailing `extLst?`: capture verbatim rather than
  // interpreting it (see `CFRule::ext_lst_raw`).
  if (const pugi::xml_node ext_lst = rule.child("extLst"); ext_lst) {
    out.ext_lst_raw = raw_xml(ext_lst);
  }

  return out;
}

}  // namespace

Expected<std::vector<cf::ConditionalFormat>, Error> read_conditional_formats(const pugi::xml_node& worksheet,
                                                                             ReadDiagnostics* diagnostics) {
  std::vector<cf::ConditionalFormat> out;
  const std::unordered_map<std::string, pugi::xml_node> x14_data_bars = CollectX14DataBarOverlay(worksheet);
  for (pugi::xml_node block = worksheet.child("conditionalFormatting"); block;
       block = block.next_sibling("conditionalFormatting")) {
    // A conditional-formatting block is a presentation-only overlay; a
    // single malformed block (missing or unparseable `sqref`) must not
    // reject the whole workbook. Skip the offending block with a WARN
    // diagnostic and continue, mirroring how other optional parts
    // degrade. Genuine sheet data is read by a separate reader and is
    // unaffected by this scope.
    const pugi::xml_attribute sqref_attr = block.attribute("sqref");
    if (!sqref_attr) {
      StructuredLog("io.cf.skip")
          .field("reason", std::string_view("sqref attribute missing"))
          .error_code(FormulonErrorCode::kIoSheetCorrupt)
          .warn();
      if (diagnostics != nullptr) {
        ++diagnostics->skipped_feature_count;
      }
      continue;
    }
    auto ranges_or = ParseSqref(sqref_attr.value());
    if (!ranges_or) {
      StructuredLog("io.cf.skip")
          .field("reason", std::string_view("sqref unparseable"))
          .field("sqref", std::string_view(sqref_attr.value()))
          .error_code(FormulonErrorCode::kIoSheetCorrupt)
          .warn();
      if (diagnostics != nullptr) {
        ++diagnostics->skipped_feature_count;
      }
      continue;
    }
    cf::ConditionalFormat cfmt;
    cfmt.sqref = std::move(ranges_or.value());
    cfmt.pivot_scope = attr_bool(block, "pivot");
    for (pugi::xml_node rule = block.child("cfRule"); rule; rule = rule.next_sibling("cfRule")) {
      cf::CFRule parsed = ReadCfRule(rule);
      if (parsed.type == cf::RuleType::DataBar && parsed.data_bar.has_value() && !parsed.id.empty()) {
        if (const auto it = x14_data_bars.find(parsed.id); it != x14_data_bars.end()) {
          ApplyX14DataBarOverlay(it->second, &*parsed.data_bar);
        }
      }
      cfmt.rules.push_back(std::move(parsed));
    }
    // `CT_ConditionalFormatting`'s schema-trailing `extLst?`: a sibling of
    // the block's `<cfRule>` children, captured verbatim (see
    // `ConditionalFormat::ext_lst_raw`).
    if (const pugi::xml_node ext_lst = block.child("extLst"); ext_lst) {
      cfmt.ext_lst_raw = raw_xml(ext_lst);
    }
    out.push_back(std::move(cfmt));
  }
  return out;
}

void normalize_cf_dxf_ids(std::vector<cf::ConditionalFormat>& formats, std::size_t dxf_count) {
  for (cf::ConditionalFormat& cfmt : formats) {
    for (cf::CFRule& rule : cfmt.rules) {
      if (!rule.dxf_id.has_value() || static_cast<std::size_t>(*rule.dxf_id) < dxf_count) {
        continue;
      }
      StructuredLog("io.cf.dxf_id_unresolved")
          .field("dxf_id", static_cast<std::int64_t>(*rule.dxf_id))
          .field("dxf_count", static_cast<std::int64_t>(dxf_count))
          .warn();
      rule.dxf_id.reset();
    }
  }
}

}  // namespace formulon::io
