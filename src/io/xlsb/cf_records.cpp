#include "io/xlsb/cf_records.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdio>
#include <cstring>
#include <limits>
#include <string>
#include <unordered_map>
#include <utility>

#include "io/cf_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "io/xsd_double.h"
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
constexpr std::uint16_t kFrtBegin = 35;
constexpr std::uint16_t kFrtEnd = 36;
constexpr std::uint16_t kCfRuleExt = 1146;

// BrtFRTBegin payload Excel writes around the x14 link, and the
// BrtCFRuleExt layout: a zero u32, then the GUID in Windows byte order.
constexpr std::array<std::uint8_t, 4> kFrtVersion = {0x02, 0x0E, 0x00, 0x00};
constexpr std::size_t kGuidBytes = 16;

// x14 conditional-format records (inside a sheet-level FRT block).
// BrtBeginCFRule14 for a data bar is 70 bytes with the rule GUID at offset
// 50; BrtBeginDataBar14 is a zero u32, minLength, maxLength, a byte always
// 1, direction, axis position (0 automatic, 1 middle, 2 none) and a u16
// flag word; each BrtColor14 is a zero u32 and a BrtColor, positional in
// the order border, negative fill, negative border, axis.
constexpr std::uint16_t kBeginCfRule14 = 1048;
constexpr std::uint16_t kEndCfRule14 = 1049;
constexpr std::uint16_t kBeginDataBar14 = 1051;
constexpr std::uint16_t kColor14 = 1055;
constexpr std::size_t kDataBarRule14Bytes = 70;
constexpr std::size_t kRule14GuidOffset = 50;
constexpr std::size_t kDataBar14Bytes = 11;
constexpr std::uint16_t kBar14Border = 0x01;
constexpr std::uint16_t kBar14Gradient = 0x02;
constexpr std::uint16_t kBar14NegativeFill = 0x04;
constexpr std::uint16_t kBar14NegativeBorder = 0x08;
constexpr std::uint16_t kBeginCfs14 = 1135;
constexpr std::uint16_t kEndCfs14 = 1136;
constexpr std::uint16_t kBeginCondFmt14 = 1046;
constexpr std::uint16_t kEndCondFmt14 = 1047;
constexpr std::uint16_t kCfvo14 = 1050;
constexpr std::uint16_t kEndDataBar14 = 1156;
constexpr std::uint16_t kAcBegin = 37;
constexpr std::uint16_t kPrintOptions = 477;
constexpr std::uint16_t kMargins = 476;
constexpr std::uint16_t kPageSetup = 478;
constexpr std::uint8_t kLegacyMinLength = 10;
constexpr std::uint8_t kLegacyMaxLength = 90;
// BrtCFVO14 type, indexed by `cf::CfvoType`; 7 (formula) is not written.
constexpr std::array<std::uint32_t, 8> kCfvo14Types = {1, 4, 5, 2, 3, 7, 8, 9};

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

std::optional<cf::Color> DecodeColor(ByteSpan p) {
  if (p.size != 8U || p.data[0] != 0x05U || p.data[1] != 0xFFU || p.data[2] != 0U || p.data[3] != 0U) {
    return std::nullopt;
  }
  return cf::Color{p.data[4], p.data[5], p.data[6], p.data[7]};
}

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

std::string FormatGuid(const std::uint8_t* g) {
  char buf[40];
  std::snprintf(buf, sizeof(buf), "{%02X%02X%02X%02X-%02X%02X-%02X%02X-%02X%02X-%02X%02X%02X%02X%02X%02X}", g[3], g[2],
                g[1], g[0], g[5], g[4], g[7], g[6], g[8], g[9], g[10], g[11], g[12], g[13], g[14], g[15]);
  return buf;
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

bool ParseGuid(const std::string& id, std::array<std::uint8_t, kGuidBytes>& out) {
  // {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
  if (id.size() != 38U || id.front() != '{' || id.back() != '}') {
    return false;
  }
  std::array<std::uint8_t, kGuidBytes> be{};
  std::size_t n = 0;
  for (std::size_t i = 1; i + 1U < id.size(); ++i) {
    const char c = id[i];
    if (c == '-') {
      if (i != 9U && i != 14U && i != 19U && i != 24U) {
        return false;
      }
      continue;
    }
    const int d = (c >= '0' && c <= '9')   ? c - '0'
                  : (c >= 'A' && c <= 'F') ? c - 'A' + 10
                  : (c >= 'a' && c <= 'f') ? c - 'a' + 10
                                           : -1;
    if (d < 0 || n >= 2U * kGuidBytes) {
      return false;
    }
    be[n / 2U] = static_cast<std::uint8_t>((be[n / 2U] << 4U) | static_cast<std::uint8_t>(d));
    ++n;
  }
  if (n != 2U * kGuidBytes) {
    return false;
  }
  // Data1..Data3 are little-endian on the wire; Data4 is a byte string.
  constexpr std::array<std::uint8_t, kGuidBytes> kOrder = {3, 2, 1, 0, 5, 4, 7, 6, 8, 9, 10, 11, 12, 13, 14, 15};
  for (std::size_t i = 0; i < kGuidBytes; ++i) {
    out[i] = be[kOrder[i]];
  }
  return true;
}

void EmitColor(std::vector<std::uint8_t>& dst, cf::Color c) {
  const std::vector<std::uint8_t> payload = {0x05, 0xFF, 0x00, 0x00, c.r, c.g, c.b, c.a};
  emit_record(dst, kColor, payload);
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
    auto f = encode_feature_formula(v.value, base.first_row, base.first_col, PtgRootClass::kValue, ctx);
    if (!f) {
      return f.error();
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

Expected<void, Error> EmitRule(std::vector<std::uint8_t>& dst, const cf::CFRule& rule, const MergeRange& base,
                               const FeatureFormulaWriteContext& ctx, bool linked) {
  auto codes = RuleCodes(rule);
  if (!codes) {
    return codes.error();
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
      auto f = encode_feature_formula(**sources[i], base.first_row, base.first_col, PtgRootClass::kValue, ctx);
      if (!f) {
        return f.error();
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

void apply_x14_data_bar_overlays(ByteSpan records, std::vector<cf::ConditionalFormat>& formats) {
  // State of the x14 rule being read; `measured` drops to false at the
  // first record outside the measured data-bar shape.
  struct Rule14 {
    bool measured = false;
    bool has_bar = false;
    std::string id;
    std::uint16_t flags = 0;
    cf::DataBarSpec bar;
    std::vector<cf::Color> colors;
  } rule14;
  ByteSpan cursor = records;
  while (cursor.size != 0U) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return;
    }
    const XlsbRecord& rec = rec_or.value();
    const std::uint8_t* d = rec.payload.data;
    if (rec.type == kBeginCfRule14) {
      rule14 = Rule14{};
      rule14.measured = rec.payload.size == kDataBarRule14Bytes && d[4] == 4U;
      if (rule14.measured) {
        rule14.id = FormatGuid(d + kRule14GuidOffset);
      }
    } else if (rec.type == kBeginDataBar14 && rule14.measured) {
      rule14.measured = rec.payload.size == kDataBar14Bytes && d[0] == 0U && d[1] == 0U && d[2] == 0U && d[3] == 0U &&
                        d[4] <= 100U && d[5] <= 100U && d[6] == 1U && d[8] <= 2U && (d[9] & 0xF0U) == 0U && d[10] == 0U;
      if (rule14.measured) {
        rule14.has_bar = true;
        rule14.bar.min_length_pct = d[4];
        rule14.bar.max_length_pct = d[5];
        rule14.bar.axis_position = static_cast<cf::DataBarAxisPosition>(d[8]);
        rule14.flags = d[9];
        rule14.bar.gradient = (rule14.flags & kBar14Gradient) != 0U;
      }
    } else if (rec.type == kColor14 && rule14.measured) {
      const bool framed = rec.payload.size == 12U && d[0] == 0U && d[1] == 0U && d[2] == 0U && d[3] == 0U;
      const std::optional<cf::Color> color = framed ? DecodeColor(ByteSpan{d + 4, 8U}) : std::nullopt;
      rule14.measured = color.has_value();
      if (color) {
        rule14.colors.push_back(*color);
      }
    } else if (rec.type == kEndCfRule14 && rule14.measured && rule14.has_bar) {
      const bool border = (rule14.flags & kBar14Border) != 0U;
      const bool negative_fill = (rule14.flags & kBar14NegativeFill) != 0U;
      const bool negative_border = border && (rule14.flags & kBar14NegativeBorder) != 0U;
      const bool axis = rule14.bar.axis_position != cf::DataBarAxisPosition::None;
      const std::size_t expected = static_cast<std::size_t>(border) + static_cast<std::size_t>(negative_fill) +
                                   static_cast<std::size_t>(negative_border) + static_cast<std::size_t>(axis);
      for (cf::ConditionalFormat& format : formats) {
        for (cf::CFRule& rule : format.rules) {
          if (rule.id != rule14.id || !rule.data_bar || rule14.colors.size() != expected) {
            continue;
          }
          cf::DataBarSpec& out = *rule.data_bar;
          std::size_t next = 0;
          if (border) {
            out.border = rule14.colors[next++];
          }
          out.negative_fill = negative_fill ? rule14.colors[next++] : out.fill;
          if (negative_border) {
            out.negative_border = rule14.colors[next++];
          }
          if (axis) {
            out.axis_color = rule14.colors[next++];
          }
          out.axis_position = rule14.bar.axis_position;
          out.gradient = rule14.bar.gradient;
          out.min_length_pct = rule14.bar.min_length_pct;
          out.max_length_pct = rule14.bar.max_length_pct;
        }
      }
      rule14.measured = false;
    }
  }
}

namespace {

/// One framed record of a buffer, by offset.
struct FramedRecord {
  std::uint16_t type = 0;
  ByteSpan payload{};
  std::size_t begin = 0;
  std::size_t end = 0;
};

/// Splits `buf` into framed records; empty when it does not parse whole.
std::vector<FramedRecord> SplitRecords(const std::vector<std::uint8_t>& buf) {
  std::vector<FramedRecord> out;
  ByteSpan cursor{buf.data(), buf.size()};
  while (cursor.size != 0U) {
    const std::size_t begin = static_cast<std::size_t>(cursor.data - buf.data());
    auto rec = read_record(cursor);
    if (!rec) {
      return {};
    }
    out.push_back(
        FramedRecord{rec.value().type, rec.value().payload, begin, static_cast<std::size_t>(cursor.data - buf.data())});
  }
  return out;
}

void Append(std::vector<std::uint8_t>& dst, const std::vector<std::uint8_t>& src, const FramedRecord& rec) {
  dst.insert(dst.end(), src.begin() + static_cast<std::ptrdiff_t>(rec.begin),
             src.begin() + static_cast<std::ptrdiff_t>(rec.end));
}

bool IsMeasuredDataBarRule14(const FramedRecord& rec) {
  return rec.type == kBeginCfRule14 && rec.payload.size == kDataBarRule14Bytes && rec.payload.data[4] == 4U;
}

/// BrtBeginDataBar14 and the BrtColor14 run for `bar`. `keep` holds the
/// source record's always-1 byte and direction, which the model does not
/// carry.
void EmitDataBar14(std::vector<std::uint8_t>& head, std::vector<std::uint8_t>& colors, const cf::DataBarSpec& bar,
                   std::uint8_t always_one, std::uint8_t direction) {
  const bool border = bar.border.has_value();
  const bool negative_fill = bar.negative_fill != bar.fill;
  const bool negative_border = border && bar.negative_border.has_value();
  const std::uint16_t flags = static_cast<std::uint16_t>(
      (border ? kBar14Border : 0U) | (bar.gradient ? kBar14Gradient : 0U) | (negative_fill ? kBar14NegativeFill : 0U) |
      (negative_border ? kBar14NegativeBorder : 0U));
  std::vector<std::uint8_t> p(4U, 0U);
  p.push_back(bar.min_length_pct);
  p.push_back(bar.max_length_pct);
  p.push_back(always_one);
  p.push_back(direction);
  p.push_back(static_cast<std::uint8_t>(bar.axis_position));
  emit_u16(p, flags);
  emit_record(head, kBeginDataBar14, p);
  const auto color = [&colors](cf::Color c) {
    const std::vector<std::uint8_t> payload = {0x00, 0x00, 0x00, 0x00, 0x05, 0xFF, 0x00, 0x00, c.r, c.g, c.b, c.a};
    emit_record(colors, kColor14, payload);
  };
  if (border) {
    color(*bar.border);
  }
  if (negative_fill) {
    color(bar.negative_fill);
  }
  if (negative_border) {
    color(*bar.negative_border);
  }
  if (bar.axis_position != cf::DataBarAxisPosition::None) {
    color(bar.axis_color);
  }
}

/// A retained x14 data-bar rule (`recs[first..last]`, BrtBeginCFRule14 to
/// BrtEndCFRule14) rewritten from `bar`, or the source bytes when its
/// BrtBeginDataBar14 is not the measured shape.
std::vector<std::uint8_t> RewriteRule14(const std::vector<std::uint8_t>& buf, const std::vector<FramedRecord>& recs,
                                        std::size_t first, std::size_t last, const cf::DataBarSpec& bar) {
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> colors;
  bool colors_pending = false;
  for (std::size_t i = first; i <= last; ++i) {
    const FramedRecord& rec = recs[i];
    if (rec.type == kBeginDataBar14) {
      if (rec.payload.size != kDataBar14Bytes) {
        std::vector<std::uint8_t> raw;
        for (std::size_t k = first; k <= last; ++k) {
          Append(raw, buf, recs[k]);
        }
        return raw;
      }
      EmitDataBar14(out, colors, bar, rec.payload.data[6], rec.payload.data[7]);
      colors_pending = true;
    } else if (rec.type == kColor14) {
      continue;
    } else {
      if (colors_pending && rec.type != kCfvo14) {
        out.insert(out.end(), colors.begin(), colors.end());
        colors_pending = false;
      }
      Append(out, buf, rec);
    }
  }
  return out;
}

std::size_t FindType(const std::vector<FramedRecord>& recs, std::size_t from, std::size_t until, std::uint16_t type) {
  for (std::size_t i = from; i < until; ++i) {
    if (recs[i].type == type) {
      return i;
    }
  }
  return until;
}

}  // namespace

void reconcile_x14_data_bars(std::vector<std::uint8_t>& records, const std::vector<cf::ConditionalFormat>& formats,
                             std::unordered_set<std::string>& linked) {
  const std::vector<FramedRecord> recs = SplitRecords(records);
  std::unordered_map<std::string, const cf::DataBarSpec*> model;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty() && rule.data_bar) {
        model.emplace(rule.id, &*rule.data_bar);
      }
    }
  }
  std::vector<std::uint8_t> out;
  const std::size_t n = recs.size();
  for (std::size_t i = 0; i < n;) {
    const std::size_t close = FindType(recs, i, n, kEndCfs14);
    const bool container = recs[i].type == kFrtBegin && i + 1U < n && recs[i + 1U].type == kBeginCfs14 &&
                           close + 1U < n && recs[close + 1U].type == kFrtEnd;
    if (!container) {
      Append(out, records, recs[i]);
      ++i;
      continue;
    }
    std::vector<std::uint8_t> inner;
    for (std::size_t k = i + 2U; k < close;) {
      const std::size_t block_end = FindType(recs, k, close, kEndCondFmt14);
      if (recs[k].type != kBeginCondFmt14 || block_end == close) {
        Append(inner, records, recs[k]);
        ++k;
        continue;
      }
      std::vector<std::uint8_t> block;
      std::size_t rules = 0;
      std::size_t kept = 0;
      for (std::size_t l = k + 1U; l < block_end;) {
        const std::size_t rule_end = FindType(recs, l, block_end, kEndCfRule14);
        if (recs[l].type != kBeginCfRule14 || rule_end == block_end) {
          Append(block, records, recs[l]);
          ++l;
          continue;
        }
        ++rules;
        std::vector<std::uint8_t> rule;
        if (IsMeasuredDataBarRule14(recs[l])) {
          const std::string id = FormatGuid(recs[l].payload.data + kRule14GuidOffset);
          const auto it = model.find(id);
          if (it != model.end()) {
            rule = RewriteRule14(records, recs, l, rule_end, *it->second);
            linked.insert(id);
          }
        } else {
          for (std::size_t m = l; m <= rule_end; ++m) {
            Append(rule, records, recs[m]);
          }
        }
        if (!rule.empty()) {
          block.insert(block.end(), rule.begin(), rule.end());
          ++kept;
        }
        l = rule_end + 1U;
      }
      if (kept != 0U || rules == 0U) {
        Append(inner, records, recs[k]);
        inner.insert(inner.end(), block.begin(), block.end());
        Append(inner, records, recs[block_end]);
      }
      k = block_end + 1U;
    }
    if (!inner.empty()) {
      Append(out, records, recs[i]);
      Append(out, records, recs[i + 1U]);
      out.insert(out.end(), inner.begin(), inner.end());
      Append(out, records, recs[close]);
      Append(out, records, recs[close + 1U]);
    }
    i = close + 2U;
  }
  if (n != 0U) {
    records = std::move(out);
  }
}

Expected<void, Error> add_x14_data_bars(std::vector<std::uint8_t>& records,
                                        const std::vector<cf::ConditionalFormat>& formats,
                                        std::unordered_set<std::string>& linked) {
  std::vector<std::uint8_t> groups;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (rule.id.empty() || !rule.data_bar || linked.count(rule.id) != 0U || !data_bar_needs_x14(*rule.data_bar)) {
        continue;
      }
      std::array<std::uint8_t, kGuidBytes> guid{};
      if (!ParseGuid(rule.id, guid)) {
        return Refuse("xlsb conditional format x14 link id is not a GUID: " + rule.id);
      }
      const cf::DataBarSpec& bar = *rule.data_bar;
      std::vector<std::uint8_t> head = {0x02, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00};
      std::vector<MergeRange> ranges;
      for (const cf::CFCellRange& r : format.sqref) {
        ranges.push_back(MergeRange{r.first.row, r.first.col, r.last.row, r.last.col});
      }
      emit_sqref(head, ranges);
      emit_u32(head, 1U);
      emit_u32(head, 0U);
      emit_record(groups, kBeginCondFmt14, head);
      std::vector<std::uint8_t> rule14(kDataBarRule14Bytes, 0U);
      rule14[4] = 4U;
      std::fill(rule14.begin() + 16, rule14.begin() + 20, std::uint8_t{0xFF});
      std::copy(guid.begin(), guid.end(), rule14.begin() + static_cast<std::ptrdiff_t>(kRule14GuidOffset));
      rule14[66] = 1U;
      emit_record(groups, kBeginCfRule14, rule14);
      std::vector<std::uint8_t> colors;
      EmitDataBar14(groups, colors, bar, 1U, 0U);
      for (const cf::CfValueObject* v : {&bar.min, &bar.max}) {
        double num = 0;
        if (v->type == cf::CfvoType::Formula) {
          return Refuse("xlsb x14 data bar with a formula threshold has no measured record form");
        }
        if (v->type == cf::CfvoType::Number || v->type == cf::CfvoType::Percent ||
            v->type == cf::CfvoType::Percentile) {
          if (!v->value.empty() && !parse_xsd_double(v->value, &num)) {
            return Refuse("xlsb conditional format threshold is not a number: " + v->value);
          }
        }
        std::vector<std::uint8_t> p(4U, 0U);
        emit_u32(p, kCfvo14Types[static_cast<std::size_t>(v->type)]);
        emit_double(p, num);
        p.resize(p.size() + 12U, 0U);
        emit_record(groups, kCfvo14, p);
      }
      groups.insert(groups.end(), colors.begin(), colors.end());
      emit_record(groups, kEndDataBar14, ByteSpan{});
      emit_record(groups, kEndCfRule14, ByteSpan{});
      emit_record(groups, kEndCondFmt14, ByteSpan{});
      linked.insert(rule.id);
    }
  }
  if (groups.empty()) {
    return Expected<void, Error>::Ok();
  }
  const std::vector<FramedRecord> recs = SplitRecords(records);
  for (std::size_t i = 0; i + 1U < recs.size(); ++i) {
    if (recs[i].type == kFrtBegin && recs[i + 1U].type == kBeginCfs14) {
      const std::size_t close = FindType(recs, i, recs.size(), kEndCfs14);
      if (close != recs.size()) {
        records.insert(records.begin() + static_cast<std::ptrdiff_t>(recs[close].begin), groups.begin(), groups.end());
        return Expected<void, Error>::Ok();
      }
    }
  }
  std::vector<std::uint8_t> block;
  // Excel refuses an x14 container that directly follows the last legacy
  // block (measured); every sheet it writes has BrtPrintOptions there, so
  // a sheet without print records gets Excel's default one.
  const bool has_print = std::any_of(recs.begin(), recs.end(), [](const FramedRecord& r) {
    return r.type == kPrintOptions || r.type == kMargins || r.type == kPageSetup;
  });
  if (!has_print) {
    const std::array<std::uint8_t, 2> default_print_options = {0x10, 0x00};
    emit_record(block, kPrintOptions, ByteSpan{default_print_options.data(), default_print_options.size()});
  }
  emit_record(block, kFrtBegin, ByteSpan{kFrtVersion.data(), kFrtVersion.size()});
  emit_record(block, kBeginCfs14, ByteSpan{});
  block.insert(block.end(), groups.begin(), groups.end());
  emit_record(block, kEndCfs14, ByteSpan{});
  emit_record(block, kFrtEnd, ByteSpan{});
  // Excel keeps the sheet's revision-uid wrapper last.
  std::size_t at = records.size();
  if (recs.size() >= 3U && recs[recs.size() - 3U].type == kAcBegin) {
    at = recs[recs.size() - 3U].begin;
  }
  records.insert(records.begin() + static_cast<std::ptrdiff_t>(at), block.begin(), block.end());
  return Expected<void, Error>::Ok();
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
