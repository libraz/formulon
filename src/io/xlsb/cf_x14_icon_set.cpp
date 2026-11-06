// Custom-icon-set BrtBeginCFRule14 records: reading a rule into the model and writing it back.

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cf/cf_x14_id.h"
#include "io/cf_writer.h"
#include "io/xlsb/cf_records.h"
#include "io/xlsb/cf_x14_link.h"
#include "io/xlsb/cf_x14_records_internal.h"
#include "io/xlsb/dxf_props.h"
#include "io/xlsb/feature_formula.h"
#include "io/xlsb/frt_formula.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/tail_refs.h"
#include "io/xml_utils.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {

using detail::kCfvo14Types;
using detail::kEndCfRule14;
using detail::kRule14GuidOffset;

namespace {

// Custom icon set ([MS-XLSB] 2.4.24 / 2.4.92 / 2.4.332 / 2.4.446): a
// standalone BrtBeginCFRule14 (CF_TYPE_MULTISTATE, CF_TEMPLATE_MULTISTATE,
// iPri > 0), BrtBeginIconSet14 (FRTBlank, KPISets14 iSet, u16 flags: fCustom
// bit 0, fIcon bit 1, fReverse bit 2, then the gte of the first four
// thresholds as in BrtBeginIconSet), one BrtCFVO14 per icon (floor first),
// one BrtCFIcon (FRTBlank, KPISets14, Icon) per icon and BrtEndIconSet14.
constexpr std::uint16_t kBeginIconSet14 = 1052;
constexpr std::uint16_t kCfIcon = 1112;
constexpr std::uint16_t kEndIconSet14 = 1155;
constexpr std::uint32_t kTypeMultistate = 6;
constexpr std::uint32_t kTemplateMultistate = 4;
constexpr std::size_t kRule14Bytes = 70;
constexpr std::size_t kRule14PriorityOffset = 16;
constexpr std::size_t kRule14FGuidOffset = 66;
constexpr std::uint16_t kIconCustom = 0x01;
constexpr std::uint16_t kIconOnly = 0x02;
constexpr std::uint16_t kIconReverse = 0x04;

/// KPINIL_14, spelled `NoIcons`.
constexpr std::int32_t kKpiNil14 = -1;
constexpr std::string_view kNoIconsName = "NoIcons";

/// KPI3STARS_14 / KPI3TRIANGLES_14 / KPI5BOXES_14, which follow the 17 built-in sets.
struct ExtraKpi14 {
  std::int32_t id;
  std::string_view name;
};
constexpr ExtraKpi14 kExtraKpi14[] = {{17, "3Stars"}, {18, "3Triangles"}, {19, "5Boxes"}};

/// KPISets14 value of an `ST_IconSetType` spelling.
std::optional<std::int32_t> Kpi14(std::string_view name) {
  if (name == kNoIconsName) {
    return kKpiNil14;
  }
  for (std::size_t i = 0; i < cf::kIconSetNames.size(); ++i) {
    if (cf::kIconSetNames[i] == name) {
      return static_cast<std::int32_t>(i);
    }
  }
  for (const ExtraKpi14& extra : kExtraKpi14) {
    if (name == extra.name) {
      return extra.id;
    }
  }
  return std::nullopt;
}

std::optional<std::string> Kpi14Name(std::int32_t set) {
  if (set == kKpiNil14) {
    return std::string(kNoIconsName);
  }
  if (set >= 0 && static_cast<std::size_t>(set) < cf::kIconSetNames.size()) {
    return std::string(cf::kIconSetNames[static_cast<std::size_t>(set)]);
  }
  for (const ExtraKpi14& extra : kExtraKpi14) {
    if (set == extra.id) {
      return std::string(extra.name);
    }
  }
  return std::nullopt;
}

/// One icon-set BrtCFVO14: FRTHeader (no formula, or one FRTFormula), iType,
/// numParam, fSaveGTE (1), fGTE, cbFmla. Min / max thresholds are not
/// allowed in an icon set.
std::optional<cf::CfValueObject> DecodeIconCfvo14(ByteSpan p, const MergeRange& base,
                                                  const FeatureFormulaReadContext& ctx) {
  auto flags = read_u32(p);
  if (!flags || (flags.value() != 0U && flags.value() != kFrtHeaderFormulas)) {
    return std::nullopt;
  }
  std::string text;
  if (flags.value() == kFrtHeaderFormulas) {
    auto count = read_u32(p);
    if (!count || count.value() != 1U) {
      return std::nullopt;
    }
    auto f = read_frt_formula(p, base.first_row, base.first_col, ctx);
    if (!f || f.value().empty()) {
      return std::nullopt;
    }
    text = std::move(f).value();
  }
  auto itype = read_u32(p);
  auto num = read_double(p);
  auto save_gte = read_u32(p);
  auto gte = read_u32(p);
  auto cb_fmla = read_u32(p);
  if (!itype || !num || !save_gte || !gte || !cb_fmla || p.size != 0U || save_gte.value() != 1U || gte.value() > 1U) {
    return std::nullopt;
  }
  cf::CfValueObject out;
  out.gte = gte.value() != 0U;
  switch (itype.value()) {
    case 1:
      out.type = cf::CfvoType::Number;
      break;
    case 4:
      out.type = cf::CfvoType::Percent;
      break;
    case 5:
      out.type = cf::CfvoType::Percentile;
      break;
    case 7:
      out.type = cf::CfvoType::Formula;
      break;
    default:
      return std::nullopt;
  }
  const bool formula = out.type == cf::CfvoType::Formula;
  if (formula != (flags.value() == kFrtHeaderFormulas) || formula != (cb_fmla.value() != 0U)) {
    return std::nullopt;
  }
  if (formula) {
    out.value = std::move(text);
  } else {
    if (!std::isfinite(num.value())) {
      return std::nullopt;
    }
    append_xml_number(out.value, num.value());
  }
  return out;
}

}  // namespace

namespace detail {

/// The custom icon-set rule `recs[first..last]`, or nothing when any record
/// is outside the shape above.
std::optional<cf::CFRule> DecodeIconRule14(const std::vector<FramedRecord>& recs, std::size_t first, std::size_t last,
                                           const MergeRange& base, const FeatureFormulaReadContext& ctx) {
  const ByteSpan head = recs[first].payload;
  if (head.size != kRule14Bytes || load_u32(head.data) != 0U || load_u32(head.data + 4) != kTypeMultistate ||
      load_u32(head.data + 8) != kTemplateMultistate || load_u32(head.data + 34) != 0U ||
      load_u32(head.data + 38) != 0U || load_u32(head.data + 42) != 0U ||
      load_u32(head.data + kRule14FGuidOffset) > 1U) {
    return std::nullopt;
  }
  const auto priority = static_cast<std::int32_t>(load_u32(head.data + kRule14PriorityOffset));
  if (priority <= 0 || first + 2U > last || recs[first + 1U].type != kBeginIconSet14) {
    return std::nullopt;
  }
  const ByteSpan set = recs[first + 1U].payload;
  if (set.size != 10U || load_u32(set.data) != 0U) {
    return std::nullopt;
  }
  const std::uint32_t name = load_u32(set.data + 4);
  const auto flags = static_cast<std::uint16_t>(set.data[8] | (set.data[9] << 8U));
  if (name > static_cast<std::uint32_t>(cf::IconSetName::Five_Quarters) || (flags & kIconCustom) == 0U ||
      (flags & ~0x7FU) != 0U) {
    return std::nullopt;
  }
  cf::IconSetSpec spec;
  spec.name = static_cast<cf::IconSetName>(name);
  spec.show_value = (flags & kIconOnly) == 0U;
  spec.reverse = (flags & kIconReverse) != 0U;
  const std::size_t icons = cf::icon_count(spec.name);
  if (last != first + 3U + 2U * icons) {
    return std::nullopt;
  }
  for (std::size_t i = 0; i < icons; ++i) {
    const FramedRecord& rec = recs[first + 2U + i];
    std::optional<cf::CfValueObject> v = rec.type == kCfvo14 ? DecodeIconCfvo14(rec.payload, base, ctx) : std::nullopt;
    if (!v) {
      return std::nullopt;
    }
    (i == 0U ? spec.floor : spec.thresholds.emplace_back()) = std::move(*v);
  }
  for (std::size_t i = 0; i < icons; ++i) {
    const FramedRecord& rec = recs[first + 2U + icons + i];
    if (rec.type != kCfIcon || rec.payload.size != 12U || load_u32(rec.payload.data) != 0U) {
      return std::nullopt;
    }
    const auto kpi = static_cast<std::int32_t>(load_u32(rec.payload.data + 4));
    const auto index = static_cast<std::int32_t>(load_u32(rec.payload.data + 8));
    std::optional<std::string> icon_set = Kpi14Name(kpi);
    if (!icon_set ||
        (kpi == -1 ? index != -1
                   : index < 0 || static_cast<std::size_t>(index) >= cf::cf_icon_set_icon_count(*icon_set))) {
      return std::nullopt;
    }
    spec.custom_icons.push_back(cf::CfIcon{std::move(*icon_set), kpi == -1 ? 0 : index});
  }
  if (recs[last - 1U].type != kEndIconSet14 || recs[last - 1U].payload.size != 0U || recs[last].payload.size != 0U) {
    return std::nullopt;
  }
  cf::CFRule rule;
  rule.type = cf::RuleType::IconSet;
  rule.priority = priority;
  if (load_u32(head.data + kRule14FGuidOffset) == 1U) {
    rule.id = FormatGuid(head.data + kRule14GuidOffset);
  }
  rule.icon_set = std::move(spec);
  return rule;
}

}  // namespace detail

namespace {

Expected<void, Error> EmitIconCfvo14(std::vector<std::uint8_t>& dst, const cf::CfValueObject& v, const MergeRange& base,
                                     const FeatureFormulaWriteContext& ctx) {
  std::vector<std::uint8_t> p;
  double num = 0;
  std::uint32_t cb_fmla = 0;
  if (v.type == cf::CfvoType::Formula) {
    auto f = encode_feature_formula(v.value, base.first_row, base.first_col, PtgRootClass::kValue, ctx,
                                    PtgEvaluation::kConditionalFormat);
    if (!f) {
      return std::move(f.error());
    }
    emit_u32(p, kFrtHeaderFormulas);
    emit_u32(p, 1U);
    emit_frt_formula(p, f.value());
    cb_fmla = f.value().cb_fmla;
  } else if (v.type == cf::CfvoType::Number || v.type == cf::CfvoType::Percent || v.type == cf::CfvoType::Percentile) {
    ASSIGN_OR_RETURN(num, parse_threshold_number(v.value));
    emit_u32(p, 0U);
  } else {
    return refuse_write("xlsb icon set threshold must be a number, percent, percentile or formula");
  }
  emit_u32(p, kCfvo14Types[static_cast<std::size_t>(v.type)]);
  emit_double(p, num);
  emit_u32(p, 1U);
  emit_u32(p, v.gte ? 1U : 0U);
  emit_u32(p, cb_fmla);
  emit_record(dst, kCfvo14, p);
  return Expected<void, Error>::Ok();
}

}  // namespace

namespace detail {

Expected<void, Error> EmitIconRule14(std::vector<std::uint8_t>& dst, const cf::CFRule& rule, const MergeRange& base,
                                     const FeatureFormulaWriteContext& ctx) {
  const cf::IconSetSpec& set = *rule.icon_set;
  const std::size_t icons = cf::icon_count(set.name);
  if (set.thresholds.size() + 1U != icons || set.custom_icons.size() != icons) {
    return refuse_write("xlsb custom icon set needs one threshold and one icon per level");
  }
  if (rule.priority <= 0) {
    return refuse_write("xlsb custom icon set rule needs a positive priority");
  }
  std::vector<std::uint8_t> head(kRule14Bytes, 0U);
  const auto put = [&head](std::size_t at, std::uint32_t v) { store_u32(head.data() + at, v); };
  put(4, kTypeMultistate);
  put(8, kTemplateMultistate);
  put(12, kNoDxf);
  put(kRule14PriorityOffset, static_cast<std::uint32_t>(rule.priority));
  if (!rule.id.empty()) {
    std::array<std::uint8_t, kGuidBytes> guid{};
    if (!ParseGuid(rule.id, guid)) {
      return refuse_write("xlsb conditional format x14 rule id is not a GUID: " + rule.id);
    }
    std::copy(guid.begin(), guid.end(), head.begin() + static_cast<std::ptrdiff_t>(kRule14GuidOffset));
    put(kRule14FGuidOffset, 1U);
  }
  emit_record(dst, kBeginCfRule14, head);
  std::uint16_t flags =
      static_cast<std::uint16_t>(kIconCustom | (set.show_value ? 0U : kIconOnly) | (set.reverse ? kIconReverse : 0U) |
                                 ((set.floor.gte ? 1U : 0U) << 3U));
  for (std::size_t i = 0; i < set.thresholds.size() && i < 3U; ++i) {
    flags = static_cast<std::uint16_t>(flags | ((set.thresholds[i].gte ? 1U : 0U) << (4U + i)));
  }
  std::vector<std::uint8_t> p(4U, 0U);
  emit_u32(p, static_cast<std::uint32_t>(set.name));
  emit_u16(p, flags);
  emit_record(dst, kBeginIconSet14, p);
  if (auto s = EmitIconCfvo14(dst, set.floor, base, ctx); !s) {
    return s;
  }
  for (const cf::CfValueObject& v : set.thresholds) {
    if (auto s = EmitIconCfvo14(dst, v, base, ctx); !s) {
      return s;
    }
  }
  std::vector<std::uint8_t> ip;
  for (const cf::CfIcon& icon : set.custom_icons) {
    const std::optional<std::int32_t> kpi = Kpi14(icon.icon_set);
    if (!kpi || (*kpi >= 0 && (icon.icon_id < 0 ||
                               static_cast<std::size_t>(icon.icon_id) >= cf::cf_icon_set_icon_count(icon.icon_set)))) {
      return refuse_write("xlsb custom icon names no icon of a known set: " + icon.icon_set);
    }
    ip.assign(4U, 0U);
    emit_u32(ip, static_cast<std::uint32_t>(*kpi));
    emit_u32(ip, static_cast<std::uint32_t>(*kpi == -1 ? -1 : icon.icon_id));
    emit_record(dst, kCfIcon, ip);
  }
  emit_record(dst, kEndIconSet14, ByteSpan{});
  emit_record(dst, kEndCfRule14, ByteSpan{});
  return Expected<void, Error>::Ok();
}

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
