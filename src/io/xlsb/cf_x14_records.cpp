// @size-budget: 15 KB

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include "io/cf_writer.h"
#include "io/xlsb/cf_records.h"
#include "io/xlsb/cf_x14_link.h"
#include "io/xlsb/feature_formula.h"
#include "io/xlsb/frt_formula.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "io/xsd_double.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// x14 conditional-format records (inside a sheet-level FRT block).
// BrtBeginCFRule14 for a data bar is 70 bytes with the rule GUID at offset
// 50; BrtBeginDataBar14 is a zero u32, minLength, maxLength, a byte always
// 1, direction (0 context, 1 left-to-right, 2 right-to-left), axis
// position (0 automatic, 1 middle, 2 none) and a u16
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
// BrtCFVO14 type, indexed by `cf::CfvoType`; 7 (formula) is not written.
constexpr std::array<std::uint32_t, 8> kCfvo14Types = {1, 4, 5, 2, 3, 7, 8, 9};
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

Error Refuse(const std::string& what) {
  return make_error(FormulonErrorCode::kInvalidArgument, what, "context=write_xlsb");
}

}  // namespace

void apply_x14_data_bar_overlays(ByteSpan records, std::vector<cf::ConditionalFormat>& formats) {
  // State of the x14 rule being read; `measured` drops to false at the
  // first record outside the supported data-bar shape.
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
                        d[4] <= 100U && d[5] <= 100U && d[6] == 1U && d[7] <= 2U && d[8] <= 2U &&
                        (d[9] & 0xF0U) == 0U && d[10] == 0U;
      if (rule14.measured) {
        rule14.has_bar = true;
        rule14.bar.min_length_pct = d[4];
        rule14.bar.max_length_pct = d[5];
        rule14.bar.direction = static_cast<cf::DataBarDirection>(d[7]);
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
          out.direction = rule14.bar.direction;
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

bool IsMeasuredDataBarRule14(const FramedRecord& rec) {
  return rec.type == kBeginCfRule14 && rec.payload.size == kDataBarRule14Bytes && rec.payload.data[4] == 4U;
}

/// BrtBeginDataBar14 and the BrtColor14 run for `bar`. `always_one` is the
/// source record's byte that was 1 in every sample and is not modelled.
void EmitDataBar14(std::vector<std::uint8_t>& head, std::vector<std::uint8_t>& colors, const cf::DataBarSpec& bar,
                   std::uint8_t always_one) {
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
  p.push_back(static_cast<std::uint8_t>(bar.direction));
  p.push_back(static_cast<std::uint8_t>(bar.axis_position));
  emit_u16(p, flags);
  emit_record(head, kBeginDataBar14, p);
  const auto color = [&colors](cf::Color c) {
    std::vector<std::uint8_t> payload(4U, 0U);
    const auto bytes = EncodeColor(c);
    payload.insert(payload.end(), bytes.begin(), bytes.end());
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
/// BrtBeginDataBar14 is not the supported shape.
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
          append_record(raw, buf, recs[k]);
        }
        return raw;
      }
      EmitDataBar14(out, colors, bar, rec.payload.data[6]);
      colors_pending = true;
    } else if (rec.type == kColor14) {
      continue;
    } else {
      if (colors_pending && rec.type != kCfvo14) {
        out.insert(out.end(), colors.begin(), colors.end());
        colors_pending = false;
      }
      append_record(out, buf, rec);
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

std::uint32_t LoadU32(const std::uint8_t* p) {
  return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8) |
         (static_cast<std::uint32_t>(p[2]) << 16) | (static_cast<std::uint32_t>(p[3]) << 24);
}

/// The sqref of a BrtBeginConditionalFormatting14: FRTHeader flags 0x02, one
/// FRTSqref (flags, Sqrfx), then ccf and fPivot. False for any other shape.
bool ParseBlockHead(const FramedRecord& rec, std::vector<MergeRange>& ranges) {
  const ByteSpan p = rec.payload;
  if (p.size < 12U || LoadU32(p.data) != 2U || LoadU32(p.data + 4) != 1U) {
    return false;
  }
  ByteSpan cursor{p.data + 12, p.size - 12U};
  return read_sqref(cursor, ranges) && cursor.size == 8U && LoadU32(cursor.data + 4) == 0U;
}

/// What a rule callback decides for one x14 rule.
struct RuleOutcome {
  enum class Kind { kKeep, kDrop, kReplace } kind = Kind::kKeep;
  std::vector<std::uint8_t> bytes;
};

/// Rebuilds the x14 CF containers of `records`, passing every rule
/// (`recs[first..last]`, BrtBeginCFRule14 to BrtEndCFRule14) and its block's
/// ranges (empty when its head is not the supported shape) to `visit`. A
/// block's ccf follows the rules it keeps; a block or
/// container left empty goes. Records outside a container pass through.
template <typename Visit>
void RewriteX14Rules(std::vector<std::uint8_t>& records, Visit&& visit) {
  const std::vector<FramedRecord> recs = split_records(records);
  std::vector<std::uint8_t> out;
  const std::size_t n = recs.size();
  for (std::size_t i = 0; i < n;) {
    const std::size_t close = FindType(recs, i, n, kEndCfs14);
    const bool container = recs[i].type == kFrtBegin && i + 1U < n && recs[i + 1U].type == kBeginCfs14 &&
                           close + 1U < n && recs[close + 1U].type == kFrtEnd;
    if (!container) {
      append_record(out, records, recs[i]);
      ++i;
      continue;
    }
    std::vector<std::uint8_t> inner;
    for (std::size_t k = i + 2U; k < close;) {
      const std::size_t block_end = FindType(recs, k, close, kEndCondFmt14);
      if (recs[k].type != kBeginCondFmt14 || block_end == close) {
        append_record(inner, records, recs[k]);
        ++k;
        continue;
      }
      std::vector<MergeRange> ranges;
      const bool head_ok = ParseBlockHead(recs[k], ranges);
      std::vector<std::uint8_t> block;
      std::size_t rules = 0;
      std::size_t kept = 0;
      for (std::size_t l = k + 1U; l < block_end;) {
        const std::size_t rule_end = FindType(recs, l, block_end, kEndCfRule14);
        if (recs[l].type != kBeginCfRule14 || rule_end == block_end) {
          append_record(block, records, recs[l]);
          ++l;
          continue;
        }
        ++rules;
        RuleOutcome outcome = visit(recs, l, rule_end, head_ok ? ranges : std::vector<MergeRange>{});
        if (outcome.kind == RuleOutcome::Kind::kKeep) {
          for (std::size_t m = l; m <= rule_end; ++m) {
            append_record(block, records, recs[m]);
          }
          ++kept;
        } else if (outcome.kind == RuleOutcome::Kind::kReplace) {
          block.insert(block.end(), outcome.bytes.begin(), outcome.bytes.end());
          ++kept;
        }
        l = rule_end + 1U;
      }
      if (kept != 0U || rules == 0U) {
        if (kept != rules && head_ok) {
          std::vector<std::uint8_t> head(recs[k].payload.data, recs[k].payload.data + recs[k].payload.size);
          const std::uint32_t ccf = static_cast<std::uint32_t>(kept);
          for (std::size_t b = 0; b < 4U; ++b) {
            head[head.size() - 8U + b] = static_cast<std::uint8_t>(ccf >> (8U * b));
          }
          emit_record(inner, kBeginCondFmt14, head);
        } else {
          append_record(inner, records, recs[k]);
        }
        inner.insert(inner.end(), block.begin(), block.end());
        append_record(inner, records, recs[block_end]);
      }
      k = block_end + 1U;
    }
    if (!inner.empty()) {
      append_record(out, records, recs[i]);
      append_record(out, records, recs[i + 1U]);
      out.insert(out.end(), inner.begin(), inner.end());
      append_record(out, records, recs[close]);
      append_record(out, records, recs[close + 1U]);
    }
    i = close + 2U;
  }
  if (n != 0U) {
    records = std::move(out);
  }
}

/// Splices `groups` (whole BrtBeginConditionalFormatting14 blocks) into the
/// retained x14 container of `records`, or appends a new container.
void InsertX14Blocks(std::vector<std::uint8_t>& records, const std::vector<std::uint8_t>& groups) {
  const std::vector<FramedRecord> recs = split_records(records);
  for (std::size_t i = 0; i + 1U < recs.size(); ++i) {
    if (recs[i].type == kFrtBegin && recs[i + 1U].type == kBeginCfs14) {
      const std::size_t close = FindType(recs, i, recs.size(), kEndCfs14);
      if (close != recs.size()) {
        records.insert(records.begin() + static_cast<std::ptrdiff_t>(recs[close].begin), groups.begin(), groups.end());
        return;
      }
    }
  }
  std::vector<std::uint8_t> block;
  // Excel refuses an x14 container that directly follows the last legacy
  // block; every sheet it writes has BrtPrintOptions there, so
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
  // The x14 container leads the worksheet future records; Excel keeps the
  // sheet's revision-uid wrapper last.
  std::size_t at = records.size();
  for (const FramedRecord& r : recs) {
    if (r.type == kFrtBegin) {
      at = r.begin;
      break;
    }
  }
  if (at == records.size() && recs.size() >= 3U && recs[recs.size() - 3U].type == kAcBegin) {
    at = recs[recs.size() - 3U].begin;
  }
  records.insert(records.begin() + static_cast<std::ptrdiff_t>(at), block.begin(), block.end());
}

/// BrtBeginConditionalFormatting14 for `format`'s ranges holding `rules` rules.
void EmitBlockHead(std::vector<std::uint8_t>& dst, const cf::ConditionalFormat& format, std::uint32_t rules) {
  std::vector<std::uint8_t> head = {0x02, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00};
  std::vector<MergeRange> ranges;
  for (const cf::CFCellRange& r : format.sqref) {
    ranges.push_back(MergeRange{r.first.row, r.first.col, r.last.row, r.last.col});
  }
  emit_sqref(head, ranges);
  emit_u32(head, rules);
  emit_u32(head, 0U);
  emit_record(dst, kBeginCondFmt14, head);
}

}  // namespace

void reconcile_x14_data_bars(std::vector<std::uint8_t>& records, const std::vector<cf::ConditionalFormat>& formats,
                             std::unordered_set<std::string>& linked) {
  std::unordered_map<std::string, const cf::DataBarSpec*> model;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty() && rule.data_bar) {
        model.emplace(rule.id, &*rule.data_bar);
      }
    }
  }
  RewriteX14Rules(records, [&](const std::vector<FramedRecord>& recs, std::size_t first, std::size_t last,
                               const std::vector<MergeRange>&) {
    RuleOutcome outcome;
    if (!IsMeasuredDataBarRule14(recs[first])) {
      return outcome;
    }
    const std::string id = FormatGuid(recs[first].payload.data + kRule14GuidOffset);
    const auto it = model.find(id);
    if (it == model.end()) {
      outcome.kind = RuleOutcome::Kind::kDrop;
      return outcome;
    }
    outcome.kind = RuleOutcome::Kind::kReplace;
    outcome.bytes = RewriteRule14(records, recs, first, last, *it->second);
    linked.insert(id);
    return outcome;
  });
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
      EmitBlockHead(groups, format, 1U);
      std::vector<std::uint8_t> rule14(kDataBarRule14Bytes, 0U);
      rule14[4] = 4U;
      std::fill(rule14.begin() + 16, rule14.begin() + 20, std::uint8_t{0xFF});
      std::copy(guid.begin(), guid.end(), rule14.begin() + static_cast<std::ptrdiff_t>(kRule14GuidOffset));
      rule14[66] = 1U;
      emit_record(groups, kBeginCfRule14, rule14);
      std::vector<std::uint8_t> colors;
      EmitDataBar14(groups, colors, bar, 1U);
      for (const cf::CfValueObject* v : {&bar.min, &bar.max}) {
        double num = 0;
        if (v->type == cf::CfvoType::Formula) {
          return Refuse("xlsb x14 data bar with a formula threshold has no known record form");
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
  if (!groups.empty()) {
    InsertX14Blocks(records, groups);
  }
  return Expected<void, Error>::Ok();
}

namespace {

/// KPISets14 value of an `ST_IconSetType` spelling: the 17 built-in sets,
/// KPI3STARS_14 / KPI3TRIANGLES_14 / KPI5BOXES_14, and KPINIL_14 for `NoIcons`.
std::optional<std::int32_t> Kpi14(std::string_view name) {
  if (name == "NoIcons") {
    return -1;
  }
  for (std::size_t i = 0; i < cf::kIconSetNames.size(); ++i) {
    if (cf::kIconSetNames[i] == name) {
      return static_cast<std::int32_t>(i);
    }
  }
  if (name == "3Stars") {
    return 17;
  }
  if (name == "3Triangles") {
    return 18;
  }
  if (name == "5Boxes") {
    return 19;
  }
  return std::nullopt;
}

std::optional<std::string> Kpi14Name(std::int32_t set) {
  if (set == -1) {
    return std::string("NoIcons");
  }
  if (set >= 0 && static_cast<std::size_t>(set) < cf::kIconSetNames.size()) {
    return std::string(cf::kIconSetNames[static_cast<std::size_t>(set)]);
  }
  switch (set) {
    case 17:
      return std::string("3Stars");
    case 18:
      return std::string("3Triangles");
    case 19:
      return std::string("5Boxes");
    default:
      return std::nullopt;
  }
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

/// The custom icon-set rule `recs[first..last]`, or nothing when any record
/// is outside the shape above.
std::optional<cf::CFRule> DecodeIconRule14(const std::vector<FramedRecord>& recs, std::size_t first, std::size_t last,
                                           const MergeRange& base, const FeatureFormulaReadContext& ctx) {
  const ByteSpan head = recs[first].payload;
  if (head.size != kRule14Bytes || LoadU32(head.data) != 0U || LoadU32(head.data + 4) != kTypeMultistate ||
      LoadU32(head.data + 8) != kTemplateMultistate || LoadU32(head.data + 34) != 0U || LoadU32(head.data + 38) != 0U ||
      LoadU32(head.data + 42) != 0U || LoadU32(head.data + kRule14FGuidOffset) > 1U) {
    return std::nullopt;
  }
  const auto priority = static_cast<std::int32_t>(LoadU32(head.data + kRule14PriorityOffset));
  if (priority <= 0 || first + 2U > last || recs[first + 1U].type != kBeginIconSet14) {
    return std::nullopt;
  }
  const ByteSpan set = recs[first + 1U].payload;
  if (set.size != 10U || LoadU32(set.data) != 0U) {
    return std::nullopt;
  }
  const std::uint32_t name = LoadU32(set.data + 4);
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
    if (rec.type != kCfIcon || rec.payload.size != 12U || LoadU32(rec.payload.data) != 0U) {
      return std::nullopt;
    }
    const auto kpi = static_cast<std::int32_t>(LoadU32(rec.payload.data + 4));
    const auto index = static_cast<std::int32_t>(LoadU32(rec.payload.data + 8));
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
  if (LoadU32(head.data + kRule14FGuidOffset) == 1U) {
    rule.id = FormatGuid(head.data + kRule14GuidOffset);
  }
  rule.icon_set = std::move(spec);
  return rule;
}

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
    if (!v.value.empty() && !parse_xsd_double(v.value, &num)) {
      return Refuse("xlsb conditional format threshold is not a number: " + v.value);
    }
    emit_u32(p, 0U);
  } else {
    return Refuse("xlsb icon set threshold must be a number, percent, percentile or formula");
  }
  emit_u32(p, kCfvo14Types[static_cast<std::size_t>(v.type)]);
  emit_double(p, num);
  emit_u32(p, 1U);
  emit_u32(p, v.gte ? 1U : 0U);
  emit_u32(p, cb_fmla);
  emit_record(dst, kCfvo14, p);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> EmitIconRule14(std::vector<std::uint8_t>& dst, const cf::CFRule& rule, const MergeRange& base,
                                     const FeatureFormulaWriteContext& ctx) {
  const cf::IconSetSpec& set = *rule.icon_set;
  const std::size_t icons = cf::icon_count(set.name);
  if (set.thresholds.size() + 1U != icons || set.custom_icons.size() != icons) {
    return Refuse("xlsb custom icon set needs one threshold and one icon per level");
  }
  if (rule.priority <= 0) {
    return Refuse("xlsb custom icon set rule needs a positive priority");
  }
  std::vector<std::uint8_t> head(kRule14Bytes, 0U);
  const auto put = [&head](std::size_t at, std::uint32_t v) {
    for (std::size_t b = 0; b < 4U; ++b) {
      head[at + b] = static_cast<std::uint8_t>(v >> (8U * b));
    }
  };
  put(4, kTypeMultistate);
  put(8, kTemplateMultistate);
  put(12, 0xFFFFFFFFU);
  put(kRule14PriorityOffset, static_cast<std::uint32_t>(rule.priority));
  if (!rule.id.empty()) {
    std::array<std::uint8_t, kGuidBytes> guid{};
    if (!ParseGuid(rule.id, guid)) {
      return Refuse("xlsb conditional format x14 rule id is not a GUID: " + rule.id);
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
  for (const cf::CfIcon& icon : set.custom_icons) {
    const std::optional<std::int32_t> kpi = Kpi14(icon.icon_set);
    if (!kpi || (*kpi >= 0 && (icon.icon_id < 0 ||
                               static_cast<std::size_t>(icon.icon_id) >= cf::cf_icon_set_icon_count(icon.icon_set)))) {
      return Refuse("xlsb custom icon names no icon of a known set: " + icon.icon_set);
    }
    std::vector<std::uint8_t> ip(4U, 0U);
    emit_u32(ip, static_cast<std::uint32_t>(*kpi));
    emit_u32(ip, static_cast<std::uint32_t>(*kpi == -1 ? -1 : icon.icon_id));
    emit_record(dst, kCfIcon, ip);
  }
  emit_record(dst, kEndIconSet14, ByteSpan{});
  emit_record(dst, kEndCfRule14, ByteSpan{});
  return Expected<void, Error>::Ok();
}

bool SameSqref(const std::vector<cf::CFCellRange>& a, const std::vector<MergeRange>& b) {
  if (a.size() != b.size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.size(); ++i) {
    if (a[i].first.row != b[i].first_row || a[i].first.col != b[i].first_col || a[i].last.row != b[i].last_row ||
        a[i].last.col != b[i].last_col) {
      return false;
    }
  }
  return true;
}

}  // namespace

bool is_x14_icon_set_rule(const cf::CFRule& rule) {
  return rule.type == cf::RuleType::IconSet && rule.icon_set.has_value() && !rule.icon_set->custom_icons.empty();
}

void take_x14_icon_sets(std::vector<std::uint8_t>& records, const FeatureFormulaReadContext& ctx,
                        std::vector<cf::ConditionalFormat>& formats) {
  RewriteX14Rules(records, [&](const std::vector<FramedRecord>& recs, std::size_t first, std::size_t last,
                               const std::vector<MergeRange>& ranges) {
    RuleOutcome outcome;
    std::optional<cf::CFRule> rule =
        ranges.empty() ? std::nullopt : DecodeIconRule14(recs, first, last, sqref_base(ranges), ctx);
    if (!rule) {
      return outcome;
    }
    // Excel also writes the rule as a legacy icon set without the custom
    // icons; that copy and this one are a single rule.
    const auto is_shadow = [&rule](const cf::CFRule& other) {
      if (other.type != cf::RuleType::IconSet || !other.icon_set || !other.icon_set->custom_icons.empty()) {
        return false;
      }
      const cf::IconSetSpec& a = *other.icon_set;
      const cf::IconSetSpec& b = *rule->icon_set;
      const auto same_cfvo = [](const cf::CfValueObject& x, const cf::CfValueObject& y) {
        return x.type == y.type && x.value == y.value && x.gte == y.gte;
      };
      return a.name == b.name && a.reverse == b.reverse && a.show_value == b.show_value &&
             a.thresholds.size() == b.thresholds.size() && same_cfvo(a.floor, b.floor) &&
             std::equal(a.thresholds.begin(), a.thresholds.end(), b.thresholds.begin(), same_cfvo);
    };
    bool merged = false;
    for (cf::ConditionalFormat& format : formats) {
      if (format.pivot_scope || !SameSqref(format.sqref, ranges)) {
        continue;
      }
      const auto shadow = std::find_if(format.rules.begin(), format.rules.end(), is_shadow);
      if (shadow != format.rules.end()) {
        shadow->icon_set->custom_icons = std::move(rule->icon_set->custom_icons);
        shadow->id = rule->id;
        merged = true;
        break;
      }
    }
    if (merged) {
      outcome.kind = RuleOutcome::Kind::kDrop;
      return outcome;
    }
    const auto same = std::find_if(formats.begin(), formats.end(), [&ranges](const cf::ConditionalFormat& f) {
      return !f.pivot_scope && SameSqref(f.sqref, ranges);
    });
    if (same != formats.end()) {
      same->rules.push_back(std::move(*rule));
    } else {
      cf::ConditionalFormat format;
      for (const MergeRange& r : ranges) {
        format.sqref.push_back(cf::CFCellRange{{r.first_row, r.first_col}, {r.last_row, r.last_col}});
      }
      format.rules.push_back(std::move(*rule));
      formats.push_back(std::move(format));
    }
    outcome.kind = RuleOutcome::Kind::kDrop;
    return outcome;
  });
}

Expected<void, Error> add_x14_icon_sets(std::vector<std::uint8_t>& records,
                                        const std::vector<cf::ConditionalFormat>& formats,
                                        const FeatureFormulaWriteContext& ctx) {
  std::vector<std::uint8_t> groups;
  for (const cf::ConditionalFormat& format : formats) {
    const auto rules =
        static_cast<std::uint32_t>(std::count_if(format.rules.begin(), format.rules.end(), is_x14_icon_set_rule));
    if (rules == 0U || format.sqref.empty()) {
      continue;
    }
    std::vector<MergeRange> ranges;
    for (const cf::CFCellRange& r : format.sqref) {
      ranges.push_back(MergeRange{r.first.row, r.first.col, r.last.row, r.last.col});
    }
    EmitBlockHead(groups, format, rules);
    for (const cf::CFRule& rule : format.rules) {
      if (!is_x14_icon_set_rule(rule)) {
        continue;
      }
      if (auto s = EmitIconRule14(groups, rule, sqref_base(ranges), ctx); !s) {
        return s;
      }
    }
    emit_record(groups, kEndCondFmt14, ByteSpan{});
  }
  if (!groups.empty()) {
    InsertX14Blocks(records, groups);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
