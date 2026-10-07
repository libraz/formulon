// @size-budget: 41 KB

#include "io/xlsb/cf_records.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstring>
#include <limits>
#include <optional>
#include <string>
#include <unordered_map>
#include <utility>

#include "io/cf_writer.h"
#include "io/xlsb/brt_color.h"
#include "io/xlsb/cf_x14_link.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "io/xsd_double.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/number_text.h"
#include "utils/structured_log.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// Record ids inside a block.
constexpr std::uint16_t kBeginCfRule = 463;
constexpr std::uint16_t kEndCfRule = 464;
constexpr std::uint16_t kBeginIconSet = 465;
constexpr std::uint16_t kEndIconSet = 466;
constexpr std::uint16_t kBeginDataBar = 467;
constexpr std::uint16_t kEndDataBar = 468;
constexpr std::uint16_t kBeginColorScale = 469;
constexpr std::uint16_t kEndColorScale = 470;
constexpr std::uint16_t kCfvo = 471;
constexpr std::uint16_t kColor = 564;
constexpr std::uint16_t kCfRuleExt = 1146;

constexpr std::uint8_t kLegacyMinLength = 10;
constexpr std::uint8_t kLegacyMaxLength = 90;

// BrtBeginCFRule flag bits; every other bit was zero in every sample.
constexpr std::uint16_t kFlagStopIfTrue = 0x02;
constexpr std::uint16_t kFlagAbove = 0x04;
constexpr std::uint16_t kFlagBottom = 0x08;
constexpr std::uint16_t kFlagPercent = 0x10;
constexpr std::uint16_t kKnownFlags = kFlagStopIfTrue | kFlagAbove | kFlagBottom | kFlagPercent;

// BrtBeginCFRule iType.
constexpr std::uint32_t kTypeCellIs = 1;
constexpr std::uint32_t kTypeExpression = 2;
constexpr std::uint32_t kTypeColorScale = 3;
constexpr std::uint32_t kTypeDataBar = 4;
constexpr std::uint32_t kTypeTop10 = 5;
constexpr std::uint32_t kTypeIconSet = 6;

/// One (rule type, iType, iTemplate, iParam) row. `param` is fixed for
/// every row except cellIs (operator), top10 (rank) and aboveAverage
/// (stdDev), which carry a value there and set `param_is_value`.
struct RuleShape {
  cf::RuleType type;
  std::uint32_t itype;
  std::uint32_t tmpl;
  std::uint32_t param;
  bool param_is_value;
};

constexpr std::array<RuleShape, 17> kShapes = {{
    {cf::RuleType::CellIs, kTypeCellIs, 0, 0, true},
    {cf::RuleType::Expression, kTypeExpression, 1, 0, false},
    {cf::RuleType::ColorScale, kTypeColorScale, 2, 0, false},
    {cf::RuleType::DataBar, kTypeDataBar, 3, 0, false},
    {cf::RuleType::IconSet, kTypeIconSet, 4, 0, false},
    {cf::RuleType::Top10, kTypeTop10, 5, 0, true},
    {cf::RuleType::UniqueValues, kTypeExpression, 7, 0, false},
    {cf::RuleType::ContainsText, kTypeExpression, 8, 0, false},
    {cf::RuleType::NotContainsText, kTypeExpression, 8, 1, false},
    {cf::RuleType::BeginsWith, kTypeExpression, 8, 2, false},
    {cf::RuleType::EndsWith, kTypeExpression, 8, 3, false},
    {cf::RuleType::ContainsBlanks, kTypeExpression, 9, 0, false},
    {cf::RuleType::NotContainsBlanks, kTypeExpression, 10, 0, false},
    {cf::RuleType::ContainsErrors, kTypeExpression, 11, 0, false},
    {cf::RuleType::NotContainsErrors, kTypeExpression, 12, 0, false},
    {cf::RuleType::DuplicateValues, kTypeExpression, 27, 0, false},
    {cf::RuleType::AboveAverage, kTypeExpression, 25, 0, true},
}};

/// aboveAverage templates: above, below, above-or-equal, below-or-equal.
constexpr std::uint32_t kTmplAbove = 25;
constexpr std::uint32_t kTmplBelow = 26;
constexpr std::uint32_t kTmplAboveEqual = 29;
constexpr std::uint32_t kTmplBelowEqual = 30;

/// timePeriod (template, iParam), indexed by `cf::TimePeriod`.
constexpr std::array<std::pair<std::uint32_t, std::uint32_t>, 10> kTimePeriods = {{
    {15, 0},
    {17, 1},
    {16, 6},
    {18, 2},
    {21, 3},
    {23, 4},
    {22, 7},
    {24, 9},
    {19, 5},
    {20, 8},
}};

/// The formula Excel generates for each timePeriod, as it stores it: its own
/// fixed tokens (value classes, a reference-class WEEKDAY, redundant
/// parentheses), the same bytes whatever the anchor. Indexed by
/// `cf::TimePeriod`.
constexpr std::uint8_t kTimePeriodToday[] = {0x19, 0x01, 0xFE, 0xFF, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0,
                                             0x1E, 0x01, 0x00, 0x41, 0x1D, 0x01, 0x41, 0xDD, 0x00, 0x0B};
constexpr std::uint8_t kTimePeriodYesterday[] = {0x19, 0x01, 0xFE, 0xFF, 0x4C, 0x00, 0x00, 0x00, 0x00,
                                                 0x00, 0xC0, 0x1E, 0x01, 0x00, 0x41, 0x1D, 0x01, 0x41,
                                                 0xDD, 0x00, 0x1E, 0x01, 0x00, 0x04, 0x0B};
constexpr std::uint8_t kTimePeriodTomorrow[] = {0x19, 0x01, 0xFE, 0xFF, 0x4C, 0x00, 0x00, 0x00, 0x00,
                                                0x00, 0xC0, 0x1E, 0x01, 0x00, 0x41, 0x1D, 0x01, 0x41,
                                                0xDD, 0x00, 0x1E, 0x01, 0x00, 0x03, 0x0B};
constexpr std::uint8_t kTimePeriodLast7Days[] = {0x19, 0x01, 0x00, 0x00, 0x41, 0xDD, 0x00, 0x4C, 0x00, 0x00, 0x00, 0x00,
                                                 0x00, 0xC0, 0x1E, 0x01, 0x00, 0x41, 0x1D, 0x01, 0x04, 0x1E, 0x06, 0x00,
                                                 0x0A, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x01, 0x00, 0x41,
                                                 0x1D, 0x01, 0x41, 0xDD, 0x00, 0x0A, 0x42, 0x02, 0x24, 0x00};
constexpr std::uint8_t kTimePeriodThisWeek[] = {
    0x19, 0x01, 0x00, 0x00, 0x41, 0xDD, 0x00, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x00, 0x00,
    0x41, 0xD5, 0x00, 0x04, 0x41, 0xDD, 0x00, 0x22, 0x01, 0x46, 0x00, 0x1E, 0x01, 0x00, 0x04, 0x0A, 0x4C,
    0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x00, 0x00, 0x41, 0xD5, 0x00, 0x41, 0xDD, 0x00, 0x04, 0x1E,
    0x07, 0x00, 0x41, 0xDD, 0x00, 0x22, 0x01, 0x46, 0x00, 0x04, 0x0A, 0x42, 0x02, 0x24, 0x00};
constexpr std::uint8_t kTimePeriodLastWeek[] = {
    0x19, 0x01, 0x00, 0x00, 0x41, 0xDD, 0x00, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x00,
    0x00, 0x41, 0xD5, 0x00, 0x04, 0x41, 0xDD, 0x00, 0x22, 0x01, 0x46, 0x00, 0x15, 0x0C, 0x41, 0xDD,
    0x00, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x00, 0x00, 0x41, 0xD5, 0x00, 0x04, 0x41,
    0xDD, 0x00, 0x22, 0x01, 0x46, 0x00, 0x1E, 0x07, 0x00, 0x03, 0x15, 0x09, 0x42, 0x02, 0x24, 0x00};
constexpr std::uint8_t kTimePeriodNextWeek[] = {
    0x19, 0x01, 0xFC, 0xFF, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x00, 0x00, 0x41, 0xD5, 0x00,
    0x41, 0xDD, 0x00, 0x04, 0x1E, 0x07, 0x00, 0x41, 0xDD, 0x00, 0x22, 0x01, 0x46, 0x00, 0x04, 0x15, 0x0D,
    0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x1E, 0x00, 0x00, 0x41, 0xD5, 0x00, 0x41, 0xDD, 0x00, 0x04,
    0x1E, 0x0F, 0x00, 0x41, 0xDD, 0x00, 0x22, 0x01, 0x46, 0x00, 0x04, 0x15, 0x09, 0x42, 0x02, 0x24, 0x00};
constexpr std::uint8_t kTimePeriodThisMonth[] = {0x19, 0x01, 0xFC, 0xFF, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0,
                                                 0x41, 0x44, 0x00, 0x41, 0xDD, 0x00, 0x41, 0x44, 0x00, 0x0B, 0x4C,
                                                 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x41, 0x45, 0x00, 0x41, 0xDD,
                                                 0x00, 0x41, 0x45, 0x00, 0x0B, 0x42, 0x02, 0x24, 0x00};
constexpr std::uint8_t kTimePeriodLastMonth[] = {
    0x19, 0x01, 0xFC, 0xFF, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x41, 0x44, 0x00, 0x41, 0xDD,
    0x00, 0x1E, 0x00, 0x00, 0x1E, 0x01, 0x00, 0x04, 0x42, 0x02, 0xC1, 0x01, 0x41, 0x44, 0x00, 0x0B,
    0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x41, 0x45, 0x00, 0x41, 0xDD, 0x00, 0x1E, 0x00, 0x00,
    0x1E, 0x01, 0x00, 0x04, 0x42, 0x02, 0xC1, 0x01, 0x41, 0x45, 0x00, 0x0B, 0x42, 0x02, 0x24, 0x00};
constexpr std::uint8_t kTimePeriodNextMonth[] = {
    0x19, 0x01, 0xFC, 0xFF, 0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x41, 0x44, 0x00, 0x41, 0xDD,
    0x00, 0x1E, 0x00, 0x00, 0x1E, 0x01, 0x00, 0x03, 0x42, 0x02, 0xC1, 0x01, 0x41, 0x44, 0x00, 0x0B,
    0x4C, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x41, 0x45, 0x00, 0x41, 0xDD, 0x00, 0x1E, 0x00, 0x00,
    0x1E, 0x01, 0x00, 0x03, 0x42, 0x02, 0xC1, 0x01, 0x41, 0x45, 0x00, 0x0B, 0x42, 0x02, 0x24, 0x00};
constexpr std::array<std::pair<const std::uint8_t*, std::size_t>, 10> kTimePeriodFormulas = {{
    {kTimePeriodToday, sizeof(kTimePeriodToday)},
    {kTimePeriodYesterday, sizeof(kTimePeriodYesterday)},
    {kTimePeriodTomorrow, sizeof(kTimePeriodTomorrow)},
    {kTimePeriodLast7Days, sizeof(kTimePeriodLast7Days)},
    {kTimePeriodThisWeek, sizeof(kTimePeriodThisWeek)},
    {kTimePeriodLastWeek, sizeof(kTimePeriodLastWeek)},
    {kTimePeriodNextWeek, sizeof(kTimePeriodNextWeek)},
    {kTimePeriodThisMonth, sizeof(kTimePeriodThisMonth)},
    {kTimePeriodLastMonth, sizeof(kTimePeriodLastMonth)},
    {kTimePeriodNextMonth, sizeof(kTimePeriodNextMonth)},
}};

/// cellIs iParam, indexed by `cf::CellIsOperator`.
constexpr std::array<std::uint32_t, 8> kCellIsParam = {6, 8, 3, 4, 7, 5, 1, 2};

/// BrtCFVO iType, indexed by `cf::CfvoType`; 0 marks the x14-only kinds.
constexpr std::array<std::uint32_t, 8> kCfvoTypes = {1, 4, 5, 2, 3, 7, 0, 0};

/// Whether Excel stores formula slots for the rule type (the rank / average
/// / uniqueness and visual rules carry none).
bool HasFormulas(cf::RuleType t) {
  switch (t) {
    case cf::RuleType::ColorScale:
    case cf::RuleType::DataBar:
    case cf::RuleType::IconSet:
    case cf::RuleType::Top10:
    case cf::RuleType::AboveAverage:
    case cf::RuleType::DuplicateValues:
    case cf::RuleType::UniqueValues:
      return false;
    default:
      return true;
  }
}

std::size_t IconCount(cf::IconSetName name) {
  const auto v = static_cast<std::uint8_t>(name);
  return v <= static_cast<std::uint8_t>(cf::IconSetName::Three_Symbols2)       ? 3U
         : v <= static_cast<std::uint8_t>(cf::IconSetName::Four_TrafficLights) ? 4U
                                                                               : 5U;
}

bool Consumed(const ByteSpan& p) {
  return p.size == 0U;
}

// ---------------------------------------------------------------------------
// Decode
// ---------------------------------------------------------------------------

class BlockCursor {
 public:
  explicit BlockCursor(ByteSpan block) : cursor_(block) {}
  bool next(XlsbRecord& rec) {
    if (cursor_.size == 0U) {
      return false;
    }
    auto r = read_record(cursor_);
    if (!r) {
      return false;
    }
    rec = r.value();
    return true;
  }
  bool at_end() const { return cursor_.size == 0U; }

 private:
  ByteSpan cursor_;
};

std::optional<cf::CfValueObject> DecodeCfvo(ByteSpan p, bool in_icon_set, const MergeRange& base,
                                            const FeatureFormulaReadContext& ctx) {
  auto itype = read_u32(p);
  if (!itype || p.size < 8U) {
    return std::nullopt;
  }
  double num = 0;
  std::memcpy(&num, p.data, sizeof(num));
  p.data += 8;
  p.size -= 8;
  auto saved_gte = read_u32(p);
  auto gte = read_u32(p);
  auto cb_fmla = read_u32(p);
  if (!saved_gte || !gte || !cb_fmla || saved_gte.value() != (in_icon_set ? 1U : 0U) || gte.value() > 1U ||
      (!in_icon_set && gte.value() != 0U)) {
    return std::nullopt;
  }
  cf::CfValueObject out;
  out.gte = in_icon_set ? gte.value() != 0U : true;
  std::size_t kind = 0;
  while (kind < kCfvoTypes.size() && (kCfvoTypes[kind] == 0U || kCfvoTypes[kind] != itype.value())) {
    ++kind;
  }
  if (kind == kCfvoTypes.size()) {
    return std::nullopt;
  }
  out.type = static_cast<cf::CfvoType>(kind);
  if (out.type == cf::CfvoType::Formula) {
    if (cb_fmla.value() == 0U || num != 0.0) {
      return std::nullopt;
    }
    auto text = read_feature_formula(p, base.first_row, base.first_col, ctx);
    if (!text || text.value().empty()) {
      return std::nullopt;
    }
    out.value = std::move(text).value();
  } else {
    if (cb_fmla.value() != 0U || !std::isfinite(num)) {
      return std::nullopt;
    }
    if (out.type == cf::CfvoType::Min || out.type == cf::CfvoType::Max) {
      if (num != 0.0) {
        return std::nullopt;
      }
    } else {
      append_xml_number(out.value, num);
    }
  }
  if (!Consumed(p)) {
    return std::nullopt;
  }
  return out;
}

/// Decodes the visual payload and x14 link that follow a BrtBeginCFRule, up
/// to and including its BrtEndCFRule.
bool DecodeRuleBody(BlockCursor& blocks, cf::CFRule& rule, const MergeRange& base,
                    const FeatureFormulaReadContext& ctx) {
  XlsbRecord rec{};
  bool visual = false;
  while (blocks.next(rec)) {
    switch (rec.type) {
      case kEndCfRule: {
        const bool wants_visual = rule.type == cf::RuleType::ColorScale || rule.type == cf::RuleType::DataBar ||
                                  rule.type == cf::RuleType::IconSet;
        return rec.payload.size == 0U && visual == wants_visual;
      }
      case kBeginColorScale: {
        if (visual || rule.type != cf::RuleType::ColorScale || rec.payload.size != 0U) {
          return false;
        }
        cf::ColorScaleSpec spec;
        while (blocks.next(rec) && rec.type != kEndColorScale) {
          if (rec.type == kCfvo && spec.colors.empty()) {
            auto v = DecodeCfvo(rec.payload, false, base, ctx);
            if (!v) {
              return false;
            }
            spec.thresholds.push_back(std::move(*v));
          } else if (rec.type == kColor) {
            auto c = DecodeColor(rec.payload);
            if (!c) {
              return false;
            }
            spec.colors.push_back(*c);
          } else {
            return false;
          }
        }
        if (rec.type != kEndColorScale || spec.colors.size() != spec.thresholds.size() || spec.colors.size() < 2U ||
            spec.colors.size() > 3U) {
          return false;
        }
        rule.color_scale = std::move(spec);
        visual = true;
        break;
      }
      case kBeginDataBar: {
        if (visual || rule.type != cf::RuleType::DataBar || rec.payload.size != 3U) {
          return false;
        }
        cf::DataBarSpec spec;
        const std::uint8_t* d = rec.payload.data;
        if (d[0] > 100U || d[1] > 100U || (d[2] & ~1U) != 0U) {
          return false;
        }
        spec.min_length_pct = d[0];
        spec.max_length_pct = d[1];
        spec.show_value = (d[2] & 1U) != 0U;
        std::optional<cf::CfValueObject> lo;
        std::optional<cf::CfValueObject> hi;
        std::optional<cf::Color> fill;
        if (blocks.next(rec) && rec.type == kCfvo) {
          lo = DecodeCfvo(rec.payload, false, base, ctx);
        }
        if (lo && blocks.next(rec) && rec.type == kCfvo) {
          hi = DecodeCfvo(rec.payload, false, base, ctx);
        }
        if (hi && blocks.next(rec) && rec.type == kColor) {
          fill = DecodeColor(rec.payload);
        }
        if (!fill || !blocks.next(rec) || rec.type != kEndDataBar) {
          return false;
        }
        spec.min = std::move(*lo);
        spec.max = std::move(*hi);
        spec.fill = *fill;
        spec.negative_fill = *fill;
        rule.data_bar = std::move(spec);
        visual = true;
        break;
      }
      case kBeginIconSet: {
        ByteSpan p = rec.payload;
        auto set = read_u32(p);
        auto flags = read_u16(p);
        if (visual || rule.type != cf::RuleType::IconSet || !set || !flags || !Consumed(p) ||
            set.value() > static_cast<std::uint32_t>(cf::IconSetName::Five_Quarters) ||
            (flags.value() & ~0x7EU) != 0U) {
          return false;
        }
        cf::IconSetSpec spec;
        spec.name = static_cast<cf::IconSetName>(set.value());
        spec.show_value = (flags.value() & 0x02U) == 0U;
        spec.reverse = (flags.value() & 0x04U) != 0U;
        std::size_t index = 0;
        while (blocks.next(rec) && rec.type == kCfvo) {
          auto v = DecodeCfvo(rec.payload, true, base, ctx);
          // Bits 3..6 repeat the gte of the first four cfvos.
          if (!v || (index < 4U && ((flags.value() >> (3U + index)) & 1U) != (v->gte ? 1U : 0U))) {
            return false;
          }
          if (index == 0U) {
            spec.floor = std::move(*v);
          } else {
            spec.thresholds.push_back(std::move(*v));
          }
          ++index;
        }
        if (rec.type != kEndIconSet || index != IconCount(spec.name) ||
            (index < 4U && (flags.value() >> (3U + index)) != 0U)) {
          return false;
        }
        rule.icon_set = std::move(spec);
        visual = true;
        break;
      }
      case kFrtBegin: {
        if (!rule.id.empty() || rec.payload.size != kFrtVersion.size() ||
            !std::equal(kFrtVersion.begin(), kFrtVersion.end(), rec.payload.data) || !blocks.next(rec) ||
            rec.type != kCfRuleExt || rec.payload.size != 4U + kGuidBytes) {
          return false;
        }
        const std::uint8_t* d = rec.payload.data;
        if (d[0] != 0U || d[1] != 0U || d[2] != 0U || d[3] != 0U) {
          return false;
        }
        rule.id = FormatGuid(d + 4);
        if (!blocks.next(rec) || rec.type != kFrtEnd || rec.payload.size != 0U) {
          return false;
        }
        break;
      }
      default:
        return false;
    }
  }
  return false;
}

std::optional<cf::CFRule> DecodeRuleHeader(ByteSpan p, const MergeRange& base, const FeatureFormulaReadContext& ctx) {
  auto itype = read_u32(p);
  auto tmpl = read_u32(p);
  auto dxf = read_u32(p);
  auto pri = read_u32(p);
  auto param = read_u32(p);
  auto reserved1 = read_u32(p);
  auto reserved2 = read_u32(p);
  auto flags = read_u16(p);
  auto cb1 = read_u32(p);
  auto cb2 = read_u32(p);
  auto cb3 = read_u32(p);
  if (!itype || !tmpl || !dxf || !pri || !param || !reserved1 || !reserved2 || !flags || !cb1 || !cb2 || !cb3 ||
      reserved1.value() != 0U || reserved2.value() != 0U || cb3.value() != 0U || (flags.value() & ~kKnownFlags) != 0U ||
      pri.value() > static_cast<std::uint32_t>(std::numeric_limits<std::int32_t>::max())) {
    return std::nullopt;
  }
  auto text = read_xlnullablewidestring(p);
  if (!text) {
    return std::nullopt;
  }
  cf::CFRule rule;
  bool matched = false;
  const std::uint32_t t = tmpl.value();
  const std::uint32_t v = param.value();
  const bool above_template = t == kTmplAbove || t == kTmplBelow || t == kTmplAboveEqual || t == kTmplBelowEqual;
  for (const RuleShape& s : kShapes) {
    const bool tmpl_ok = s.type == cf::RuleType::AboveAverage ? above_template : s.tmpl == t;
    if (s.itype == itype.value() && tmpl_ok && (s.param_is_value || s.param == v)) {
      rule.type = s.type;
      matched = true;
      break;
    }
  }
  if (!matched && itype.value() == kTypeExpression) {
    for (std::size_t i = 0; i < kTimePeriods.size(); ++i) {
      if (kTimePeriods[i].first == t && kTimePeriods[i].second == v) {
        rule.type = cf::RuleType::TimePeriod;
        rule.time_period = static_cast<cf::TimePeriod>(i);
        matched = true;
      }
    }
  }
  if (!matched) {
    return std::nullopt;
  }
  const bool above = (flags.value() & kFlagAbove) != 0U;
  const bool rank_flags = (flags.value() & (kFlagBottom | kFlagPercent)) != 0U;
  if ((above && rule.type != cf::RuleType::AboveAverage) || (rank_flags && rule.type != cf::RuleType::Top10)) {
    return std::nullopt;
  }
  switch (rule.type) {
    case cf::RuleType::CellIs: {
      std::size_t op = 0;
      while (op < kCellIsParam.size() && kCellIsParam[op] != v) {
        ++op;
      }
      if (op == kCellIsParam.size()) {
        return std::nullopt;
      }
      rule.op = static_cast<cf::CellIsOperator>(op);
      break;
    }
    case cf::RuleType::Top10:
      if (v > static_cast<std::uint32_t>(std::numeric_limits<std::int32_t>::max())) {
        return std::nullopt;
      }
      rule.rank = static_cast<std::int32_t>(v);
      rule.percent = (flags.value() & kFlagPercent) != 0U;
      rule.bottom = (flags.value() & kFlagBottom) != 0U;
      break;
    case cf::RuleType::AboveAverage:
      rule.above_average = t == kTmplAbove || t == kTmplAboveEqual;
      rule.equal_average = t == kTmplAboveEqual || t == kTmplBelowEqual;
      if (above != rule.above_average || (rule.equal_average && v != 0U)) {
        return std::nullopt;
      }
      if (v != 0U) {
        rule.std_dev = static_cast<double>(v);
      }
      break;
    case cf::RuleType::ContainsText:
    case cf::RuleType::NotContainsText:
    case cf::RuleType::BeginsWith:
    case cf::RuleType::EndsWith:
      rule.text = std::move(text).value();
      break;
    default:
      if (!text.value().empty()) {
        return std::nullopt;
      }
      break;
  }
  rule.priority = static_cast<std::int32_t>(pri.value());
  rule.stop_if_true = (flags.value() & kFlagStopIfTrue) != 0U;
  if (dxf.value() != 0xFFFFFFFFU) {
    rule.dxf_id = dxf.value();
  }
  if (!HasFormulas(rule.type) && (cb1.value() != 0U || cb2.value() != 0U)) {
    return std::nullopt;
  }
  for (std::uint32_t cb : {cb1.value(), cb2.value()}) {
    if (cb == 0U) {
      continue;
    }
    auto f = read_feature_formula(p, base.first_row, base.first_col, ctx);
    if (!f) {
      StructuredLog("xlsb.cf.formula_not_decoded").field("reason", f.error().message).warn();
      return std::nullopt;
    }
    (rule.formula1 ? rule.formula2 : rule.formula1) = std::move(f).value();
  }
  if (!Consumed(p)) {
    return std::nullopt;
  }
  return rule;
}

// ---------------------------------------------------------------------------
// Encode
// ---------------------------------------------------------------------------

Error Refuse(const std::string& what) {
  return make_error(FormulonErrorCode::kInvalidArgument, what, "context=write_xlsb");
}

void EmitColor(std::vector<std::uint8_t>& dst, cf::Color c) {
  const auto bytes = EncodeColor(c);
  emit_record(dst, kColor, std::vector<std::uint8_t>(bytes.begin(), bytes.end()));
}

Expected<void, Error> EmitCfvo(std::vector<std::uint8_t>& dst, const cf::CfValueObject& v, bool in_icon_set,
                               const MergeRange& base, const FeatureFormulaWriteContext& ctx) {
  const std::uint32_t itype = kCfvoTypes[static_cast<std::size_t>(v.type)];
  if (itype == 0U) {
    return Refuse("xlsb conditional format threshold type has no legacy record form");
  }
  double num = 0;
  EncodedFeatureFormula formula;
  if (v.type == cf::CfvoType::Formula) {
    auto f = encode_feature_formula(v.value, base.first_row, base.first_col, PtgRootClass::kValue, ctx,
                                    PtgEvaluation::kConditionalFormat);
    if (!f) {
      return std::move(f.error());
    }
    formula = std::move(f).value();
  } else if (v.type != cf::CfvoType::Min && v.type != cf::CfvoType::Max &&
             !(v.value.empty() || parse_xsd_double(v.value, &num))) {
    return Refuse("xlsb conditional format threshold is not a number: " + v.value);
  }
  std::vector<std::uint8_t> p;
  emit_u32(p, itype);
  emit_double(p, num);
  emit_u32(p, in_icon_set ? 1U : 0U);
  emit_u32(p, in_icon_set && v.gte ? 1U : 0U);
  emit_u32(p, formula.cb_fmla);
  p.insert(p.end(), formula.bytes.begin(), formula.bytes.end());
  emit_record(dst, kCfvo, p);
  return Expected<void, Error>::Ok();
}

/// iType, template and iParam for `rule`.
Expected<std::array<std::uint32_t, 3>, Error> RuleCodes(const cf::CFRule& rule) {
  if (rule.type == cf::RuleType::TimePeriod) {
    const auto& tp = kTimePeriods[static_cast<std::size_t>(rule.time_period.value_or(cf::TimePeriod::Today))];
    return std::array<std::uint32_t, 3>{kTypeExpression, tp.first, tp.second};
  }
  for (const RuleShape& s : kShapes) {
    if (s.type != rule.type) {
      continue;
    }
    std::array<std::uint32_t, 3> codes{s.itype, s.tmpl, s.param};
    if (rule.type == cf::RuleType::CellIs) {
      codes[2] = kCellIsParam[static_cast<std::size_t>(rule.op.value_or(cf::CellIsOperator::Equal))];
    } else if (rule.type == cf::RuleType::Top10) {
      codes[2] = static_cast<std::uint32_t>(rule.rank.value_or(10));
    } else if (rule.type == cf::RuleType::AboveAverage) {
      codes[1] = rule.above_average ? (rule.equal_average ? kTmplAboveEqual : kTmplAbove)
                                    : (rule.equal_average ? kTmplBelowEqual : kTmplBelow);
      if (rule.std_dev.has_value()) {
        const double sd = *rule.std_dev;
        if (rule.equal_average || !(sd >= 1.0 && sd <= 3.0) || sd != std::floor(sd)) {
          return Refuse("xlsb aboveAverage standard deviation is not representable");
        }
        codes[2] = static_cast<std::uint32_t>(sd);
      }
    }
    return codes;
  }
  return Refuse("xlsb conditional format rule type has no record form");
}

/// Excel's stored form of a timePeriod rule's formula, when the rule carries
/// the one Excel generates for its period at `base`.
std::optional<EncodedFeatureFormula> TimePeriodFormula(const cf::CFRule& rule, const MergeRange& base) {
  if (rule.type != cf::RuleType::TimePeriod || !rule.time_period.has_value() || !rule.formula1.has_value()) {
    return std::nullopt;
  }
  const auto& [bytes, size] = kTimePeriodFormulas[static_cast<std::size_t>(*rule.time_period)];
  Arena arena;
  auto generated = decode_ptgs(ByteSpan{bytes, size}, ByteSpan{}, arena, {}, {}, {}, {}, -1,
                               PtgBaseCell{base.first_row, base.first_col});
  parser::Parser parser(*rule.formula1, arena);
  const parser::AstNode* written = parser.parse();
  if (!generated || written == nullptr || !parser.errors().empty() ||
      parser::format_formula(*generated.value()) != parser::format_formula(*written)) {
    return std::nullopt;
  }
  EncodedFeatureFormula out;
  out.cb_fmla = static_cast<std::uint32_t>(size);
  emit_u32(out.bytes, static_cast<std::uint32_t>(size));
  out.bytes.insert(out.bytes.end(), bytes, bytes + size);
  emit_u32(out.bytes, 0U);
  return out;
}

Expected<void, Error> EmitRule(std::vector<std::uint8_t>& dst, const cf::CFRule& rule, const MergeRange& base,
                               const FeatureFormulaWriteContext& ctx, bool linked) {
  auto codes = RuleCodes(rule);
  if (!codes) {
    return std::move(codes.error());
  }
  std::array<std::uint8_t, kGuidBytes> guid{};
  if (!rule.id.empty() && !ParseGuid(rule.id, guid)) {
    return Refuse("xlsb conditional format x14 link id is not a GUID: " + rule.id);
  }
  std::uint16_t flags = rule.stop_if_true ? kFlagStopIfTrue : 0U;
  if (rule.type == cf::RuleType::AboveAverage && rule.above_average) {
    flags |= kFlagAbove;
  }
  if (rule.type == cf::RuleType::Top10) {
    flags |= (rule.bottom ? kFlagBottom : 0U) | (rule.percent ? kFlagPercent : 0U);
  }
  std::array<EncodedFeatureFormula, 2> formulas;
  if (HasFormulas(rule.type)) {
    const std::optional<std::string>* sources[2] = {&rule.formula1, &rule.formula2};
    for (std::size_t i = 0; i < 2U; ++i) {
      if (!sources[i]->has_value() || (*sources[i])->empty()) {
        continue;
      }
      if (i == 0U) {
        if (std::optional<EncodedFeatureFormula> generated = TimePeriodFormula(rule, base)) {
          formulas[i] = std::move(*generated);
          continue;
        }
      }
      auto f = encode_feature_formula(**sources[i], base.first_row, base.first_col, PtgRootClass::kValue, ctx,
                                      PtgEvaluation::kConditionalFormat);
      if (!f) {
        return std::move(f.error());
      }
      formulas[i] = std::move(f).value();
    }
  }
  std::vector<std::uint8_t> p;
  emit_u32(p, codes.value()[0]);
  emit_u32(p, codes.value()[1]);
  emit_u32(p, rule.dxf_id.value_or(0xFFFFFFFFU));
  emit_u32(p, static_cast<std::uint32_t>(rule.priority));
  emit_u32(p, codes.value()[2]);
  emit_u32(p, 0U);
  emit_u32(p, 0U);
  emit_u16(p, flags);
  emit_u32(p, formulas[0].cb_fmla);
  emit_u32(p, formulas[1].cb_fmla);
  emit_u32(p, 0U);
  const bool has_text =
      rule.text.has_value() && (rule.type == cf::RuleType::ContainsText || rule.type == cf::RuleType::NotContainsText ||
                                rule.type == cf::RuleType::BeginsWith || rule.type == cf::RuleType::EndsWith);
  emit_xlnullablewidestring(p, has_text ? std::optional<std::string_view>(*rule.text) : std::nullopt);
  for (const EncodedFeatureFormula& f : formulas) {
    if (f.cb_fmla != 0U) {
      p.insert(p.end(), f.bytes.begin(), f.bytes.end());
    }
  }
  emit_record(dst, kBeginCfRule, p);

  if (rule.type == cf::RuleType::ColorScale) {
    if (!rule.color_scale || rule.color_scale->colors.size() != rule.color_scale->thresholds.size() ||
        rule.color_scale->colors.size() < 2U || rule.color_scale->colors.size() > 3U) {
      return Refuse("xlsb colour scale needs two or three matched stops");
    }
    emit_record(dst, kBeginColorScale, ByteSpan{});
    for (const cf::CfValueObject& v : rule.color_scale->thresholds) {
      if (auto s = EmitCfvo(dst, v, false, base, ctx); !s) {
        return s;
      }
    }
    for (cf::Color c : rule.color_scale->colors) {
      EmitColor(dst, c);
    }
    emit_record(dst, kEndColorScale, ByteSpan{});
  } else if (rule.type == cf::RuleType::DataBar) {
    if (!rule.data_bar) {
      return Refuse("xlsb data bar rule has no data bar");
    }
    const cf::DataBarSpec& bar = *rule.data_bar;
    // A linked bar's lengths live in its x14 record; Excel leaves the
    // legacy ones at the pre-2010 10/90.
    const std::vector<std::uint8_t> head = {linked ? kLegacyMinLength : bar.min_length_pct,
                                            linked ? kLegacyMaxLength : bar.max_length_pct,
                                            static_cast<std::uint8_t>(bar.show_value ? 1U : 0U)};
    emit_record(dst, kBeginDataBar, head);
    for (const cf::CfValueObject* v : {&bar.min, &bar.max}) {
      if (auto s = EmitCfvo(dst, *v, false, base, ctx); !s) {
        return s;
      }
    }
    EmitColor(dst, bar.fill);
    emit_record(dst, kEndDataBar, ByteSpan{});
  } else if (rule.type == cf::RuleType::IconSet) {
    if (!rule.icon_set || rule.icon_set->thresholds.size() + 1U != IconCount(rule.icon_set->name)) {
      return Refuse("xlsb icon set threshold count does not match its icons");
    }
    const cf::IconSetSpec& set = *rule.icon_set;
    std::uint16_t set_flags = static_cast<std::uint16_t>((set.show_value ? 0U : 0x02U) | (set.reverse ? 0x04U : 0U) |
                                                         ((set.floor.gte ? 1U : 0U) << 3U));
    for (std::size_t i = 0; i < set.thresholds.size() && i < 3U; ++i) {
      set_flags = static_cast<std::uint16_t>(set_flags | ((set.thresholds[i].gte ? 1U : 0U) << (4U + i)));
    }
    std::vector<std::uint8_t> head;
    emit_u32(head, static_cast<std::uint32_t>(set.name));
    emit_u16(head, set_flags);
    emit_record(dst, kBeginIconSet, head);
    if (auto s = EmitCfvo(dst, set.floor, true, base, ctx); !s) {
      return s;
    }
    for (const cf::CfValueObject& v : set.thresholds) {
      if (auto s = EmitCfvo(dst, v, true, base, ctx); !s) {
        return s;
      }
    }
    emit_record(dst, kEndIconSet, ByteSpan{});
  }
  if (!rule.id.empty()) {
    emit_record(dst, kFrtBegin, ByteSpan{kFrtVersion.data(), kFrtVersion.size()});
    std::vector<std::uint8_t> ext(4U, 0U);
    ext.insert(ext.end(), guid.begin(), guid.end());
    emit_record(dst, kCfRuleExt, ext);
    emit_record(dst, kFrtEnd, ByteSpan{});
  }
  emit_record(dst, kEndCfRule, ByteSpan{});
  return Expected<void, Error>::Ok();
}

}  // namespace

std::optional<cf::ConditionalFormat> decode_cf_block(ByteSpan block, const FeatureFormulaReadContext& ctx) {
  BlockCursor blocks(block);
  XlsbRecord rec{};
  if (!blocks.next(rec) || rec.type != kBrtBeginConditionalFormatting) {
    return std::nullopt;
  }
  ByteSpan p = rec.payload;
  auto count = read_u32(p);
  auto pivot = read_u32(p);
  std::vector<MergeRange> ranges;
  if (!count || !pivot || pivot.value() != 0U || !read_sqref(p, ranges) || !Consumed(p)) {
    return std::nullopt;
  }
  const MergeRange base = sqref_base(ranges);
  cf::ConditionalFormat out;
  for (const MergeRange& r : ranges) {
    out.sqref.push_back(cf::CFCellRange{{r.first_row, r.first_col}, {r.last_row, r.last_col}});
  }
  while (blocks.next(rec)) {
    if (rec.type == kBrtEndConditionalFormatting) {
      if (rec.payload.size != 0U || !blocks.at_end() || out.rules.size() != count.value()) {
        return std::nullopt;
      }
      return out;
    }
    if (rec.type != kBeginCfRule) {
      return std::nullopt;
    }
    auto rule = DecodeRuleHeader(rec.payload, base, ctx);
    if (!rule || !DecodeRuleBody(blocks, *rule, base, ctx)) {
      return std::nullopt;
    }
    out.rules.push_back(std::move(*rule));
  }
  return std::nullopt;
}

Expected<void, Error> emit_cf_block(std::vector<std::uint8_t>& dst, const cf::ConditionalFormat& format,
                                    const FeatureFormulaWriteContext& ctx,
                                    const std::unordered_set<std::string>& linked) {
  if (format.sqref.empty() || format.rules.empty()) {
    return Expected<void, Error>::Ok();
  }
  std::vector<MergeRange> ranges;
  for (const cf::CFCellRange& r : format.sqref) {
    ranges.push_back(MergeRange{r.first.row, r.first.col, r.last.row, r.last.col});
  }
  const MergeRange base = sqref_base(ranges);
  std::vector<std::uint8_t> head;
  emit_u32(head, static_cast<std::uint32_t>(format.rules.size()));
  emit_u32(head, format.pivot_scope ? 1U : 0U);
  emit_sqref(head, ranges);
  std::vector<std::uint8_t> body;
  emit_record(body, kBrtBeginConditionalFormatting, head);
  for (const cf::CFRule& rule : format.rules) {
    if (auto s = EmitRule(body, rule, base, ctx, !rule.id.empty() && linked.count(rule.id) != 0U); !s) {
      return s;
    }
  }
  emit_record(body, kBrtEndConditionalFormatting, ByteSpan{});
  dst.insert(dst.end(), body.begin(), body.end());
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
