// @size-budget: 15 KB

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
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/tail_refs.h"
#include "io/xml_utils.h"
#include "parser/parser.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {

using detail::DecodeIconRule14;
using detail::EmitIconRule14;
using detail::kCfvo14Types;
using detail::kEndCfRule14;
using detail::kRule14GuidOffset;

namespace {

// x14 conditional-format records (inside a sheet-level FRT block).
// BrtBeginCFRule14 for a data bar is 70 bytes with the rule GUID at offset
// 50; BrtBeginDataBar14 is a zero u32, minLength, maxLength, a byte always
// 1, direction (0 context, 1 left-to-right, 2 right-to-left), axis
// position (0 automatic, 1 middle, 2 none) and a u16
// flag word; each BrtColor14 is a zero u32 and a BrtColor, positional in
// the order border, negative fill, negative border, axis.
constexpr std::uint16_t kBeginDataBar14 = 1051;
constexpr std::uint16_t kColor14 = 1055;
constexpr std::size_t kDataBarRule14Bytes = 70;
constexpr std::size_t kDataBar14Bytes = 11;
constexpr std::uint16_t kBar14Border = 0x01;
constexpr std::uint16_t kBar14Gradient = 0x02;
constexpr std::uint16_t kBar14NegativeFill = 0x04;
constexpr std::uint16_t kBar14NegativeBorder = 0x08;
constexpr std::uint16_t kEndDataBar14 = 1156;
// A standalone formula rule (one naming another sheet): BrtBeginCFRule14 is
// an FRTHeader (0, or 0x04 with a formula list: u32 count, one FRTFormula
// per non-zero cbFmla), then iType, iTemplate, dxfId, iPri, iParam, two
// zero u32, the BrtBeginCFRule flag word, cbFmla1..3 (cce), strParam (an
// empty string unless a text rule), the rule GUID and fGuid = 1. dxfId
// indexes the styles part's BrtDXF list (-1 for none); BrtEndCFRule14 may
// be preceded by a copy of that dxf: BrtBeginDXFs (FRTBlank, count 1), one
// BrtDXF prefixed by an FRTBlank, BrtEndDXFs.
constexpr std::size_t kRule14TailBytes = kGuidBytes + 4U;

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

/// The sqref of a BrtBeginConditionalFormatting14: FRTHeader flags 0x02, one
/// FRTSqref (flags, Sqrfx), then ccf and fPivot. False for any other shape.
bool ParseBlockHead(const FramedRecord& rec, std::vector<MergeRange>& ranges) {
  const ByteSpan p = rec.payload;
  if (p.size < 12U || load_u32(p.data) != 2U || load_u32(p.data + 4) != 1U) {
    return false;
  }
  ByteSpan cursor{p.data + 12, p.size - 12U};
  return read_sqref(cursor, ranges) && cursor.size == 8U && load_u32(cursor.data + 4) == 0U;
}

/// What a rule callback decides for one x14 rule.
struct RuleOutcome {
  enum class Kind { kKeep, kDrop, kReplace } kind = Kind::kKeep;
  std::vector<std::uint8_t> bytes;
};

/// Rebuilds the x14 CF containers of `records`, passing every rule
/// (`recs[first..last]`, BrtBeginCFRule14 to BrtEndCFRule14), its block's
/// ranges (empty when its head is not the supported shape) and the index of
/// its block's head record to `visit`. A
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
        RuleOutcome outcome = visit(recs, l, rule_end, head_ok ? ranges : std::vector<MergeRange>{}, k);
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
          store_u32(head.data() + head.size() - 8U, static_cast<std::uint32_t>(kept));
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
/// retained x14 container of `records`, at its start when `front`, else at
/// its end, or appends a new container.
void InsertX14Blocks(std::vector<std::uint8_t>& records, const std::vector<std::uint8_t>& groups, bool front = false) {
  const std::vector<FramedRecord> recs = split_records(records);
  for (std::size_t i = 0; i + 1U < recs.size(); ++i) {
    if (recs[i].type == kFrtBegin && recs[i + 1U].type == kBeginCfs14) {
      const std::size_t close = FindType(recs, i, recs.size(), kEndCfs14);
      if (close != recs.size()) {
        const std::size_t at = front ? recs[i + 2U].begin : recs[close].begin;
        records.insert(records.begin() + static_cast<std::ptrdiff_t>(at), groups.begin(), groups.end());
        return;
      }
    }
  }
  const std::vector<std::uint8_t> block = frt_container_block(recs, kBeginCfs14, kEndCfs14, groups);
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

std::vector<std::uint8_t> frt_container_block(const std::vector<FramedRecord>& recs, std::uint16_t open,
                                              std::uint16_t close, const std::vector<std::uint8_t>& body) {
  std::vector<std::uint8_t> block;
  const bool has_print = std::any_of(recs.begin(), recs.end(), [](const FramedRecord& r) {
    return r.type == kBrtPrintOptions || r.type == kBrtMargins || r.type == kBrtPageSetup;
  });
  if (!has_print) {
    const std::array<std::uint8_t, 2> default_print_options = {0x10, 0x00};
    emit_record(block, kBrtPrintOptions, ByteSpan{default_print_options.data(), default_print_options.size()});
  }
  emit_record(block, kFrtBegin, ByteSpan{kFrtVersion.data(), kFrtVersion.size()});
  emit_record(block, open, ByteSpan{});
  block.insert(block.end(), body.begin(), body.end());
  emit_record(block, close, ByteSpan{});
  emit_record(block, kFrtEnd, ByteSpan{});
  return block;
}

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
                               const std::vector<MergeRange>&, std::size_t) {
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
        return refuse_write("xlsb conditional format x14 link id is not a GUID: " + rule.id);
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
          return refuse_write("xlsb x14 data bar with a formula threshold has no known record form");
        }
        if (v->type == cf::CfvoType::Number || v->type == cf::CfvoType::Percent ||
            v->type == cf::CfvoType::Percentile) {
          ASSIGN_OR_RETURN(num, parse_threshold_number(v->value));
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

bool IsVisualRule(cf::RuleType type) {
  return type == cf::RuleType::ColorScale || type == cf::RuleType::DataBar || type == cf::RuleType::IconSet;
}

/// The dxf copy between the rule head and BrtEndCFRule14
/// (`recs[from..until)`): nothing, or BrtBeginDXFs, one BrtDXF, BrtEndDXFs.
/// False for any other run, or for a copy that states formatting the
/// rule's dxfId does not name.
bool IsDxfCopy(const std::vector<FramedRecord>& recs, std::size_t from, std::size_t until, bool has_dxf) {
  if (from == until) {
    return true;
  }
  if (until - from != 3U) {
    return false;
  }
  const ByteSpan begin = recs[from].payload;
  const ByteSpan dxf = recs[from + 1U].payload;
  // BrtDXF after its FRTBlank: u16 flags, u16 reserved, u16 property count.
  return recs[from].type == kBrtBeginDxfs && begin.size == 8U && load_u32(begin.data) == 0U &&
         load_u32(begin.data + 4) == 1U && recs[from + 1U].type == kBrtDxf && dxf.size >= 10U &&
         load_u32(dxf.data) == 0U && (has_dxf || (dxf.data[8] == 0U && dxf.data[9] == 0U)) &&
         recs[from + 2U].type == kBrtEndDxfs && recs[from + 2U].payload.size == 0U;
}

/// The fields of a standalone formula rule's BrtBeginCFRule14, its
/// FRTFormulas left encoded.
struct Rule14Head {
  std::vector<ByteSpan> formulas;
  std::uint32_t dxf = kNoDxf;
  std::int32_t priority = 0;
  bool slot1 = false;
  bool slot2 = false;
  std::string id;
};

/// `p` read as a standalone non-visual rule with iPri > 0 and a GUID, or
/// nothing for any other shape. `rule` receives its kind fields.
std::optional<Rule14Head> ParseRule14Head(ByteSpan p, cf::CFRule& rule) {
  Rule14Head head;
  auto header = read_u32(p);
  if (!header || (header.value() != 0U && header.value() != kFrtHeaderFormulas)) {
    return std::nullopt;
  }
  if (header.value() == kFrtHeaderFormulas) {
    auto count = read_u32(p);
    if (!count || count.value() == 0U || count.value() > 2U) {
      return std::nullopt;
    }
    for (std::uint32_t i = 0; i < count.value(); ++i) {
      if (p.size < 12U) {
        return std::nullopt;
      }
      const std::size_t cce = load_u32(p.data + 4);
      const std::size_t cb = load_u32(p.data + 8);
      if (p.size - 12U < cce || p.size - 12U - cce < cb) {
        return std::nullopt;
      }
      const std::size_t size = 12U + cce + cb;
      head.formulas.push_back(ByteSpan{p.data, size});
      p.data += size;
      p.size -= size;
    }
  }
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
      reserved1.value() != 0U || reserved2.value() != 0U || cb3.value() != 0U ||
      static_cast<std::int32_t>(pri.value()) <= 0) {
    return std::nullopt;
  }
  auto text = read_xlnullablewidestring(p);
  if (!text || p.size != kRule14TailBytes || load_u32(p.data + kGuidBytes) != 1U) {
    return std::nullopt;
  }
  if (!decode_cf_rule_fields(CfRuleFields{itype.value(), tmpl.value(), param.value(), flags.value()},
                             std::move(text).value(), rule) ||
      IsVisualRule(rule.type)) {
    return std::nullopt;
  }
  head.slot1 = cb1.value() != 0U;
  head.slot2 = cb2.value() != 0U;
  const std::size_t slots = static_cast<std::size_t>(head.slot1) + static_cast<std::size_t>(head.slot2);
  if (slots != head.formulas.size() || (slots != 0U && !cf_rule_has_formulas(rule.type))) {
    return std::nullopt;
  }
  head.dxf = dxf.value();
  head.priority = static_cast<std::int32_t>(pri.value());
  head.id = FormatGuid(p.data);
  return head;
}

/// The standalone formula rule `recs[first..last]`, or nothing when any
/// record is outside the shape above or the rule is a visual one.
std::optional<cf::CFRule> DecodeFormulaRule14(const std::vector<FramedRecord>& recs, std::size_t first,
                                              std::size_t last, const MergeRange& base,
                                              const FeatureFormulaReadContext& ctx) {
  cf::CFRule rule;
  std::optional<Rule14Head> head = ParseRule14Head(recs[first].payload, rule);
  if (!head || !IsDxfCopy(recs, first + 1U, last, head->dxf != kNoDxf) || recs[last].payload.size != 0U) {
    return std::nullopt;
  }
  std::vector<std::string> formulas;
  for (ByteSpan span : head->formulas) {
    auto f = read_frt_formula(span, base.first_row, base.first_col, ctx);
    if (!f || f.value().empty()) {
      return std::nullopt;
    }
    formulas.push_back(std::move(f).value());
  }
  std::size_t next = 0;
  if (head->slot1) {
    rule.formula1 = std::move(formulas[next++]);
  }
  if (head->slot2) {
    rule.formula2 = std::move(formulas[next++]);
  }
  rule.priority = head->priority;
  if (head->dxf != kNoDxf) {
    rule.dxf_id = head->dxf;
  }
  rule.id = std::move(head->id);
  return rule;
}

/// `rule` as a standalone BrtBeginCFRule14 .. BrtEndCFRule14 under `id`.
/// `dxf_copy` (framed records, possibly none) goes before BrtEndCFRule14;
/// dxfId alone names the formatting.
Expected<void, Error> EmitFormulaRule14(std::vector<std::uint8_t>& dst, const cf::CFRule& rule, const std::string& id,
                                        const MergeRange& base, const FeatureFormulaWriteContext& ctx,
                                        const std::vector<std::uint8_t>& dxf_copy = {}) {
  ASSIGN_OR_RETURN(CfRuleFields fields, encode_cf_rule_fields(rule));
  std::array<std::uint8_t, kGuidBytes> guid{};
  if (!ParseGuid(id, guid)) {
    return refuse_write("xlsb conditional format x14 rule id is not a GUID: " + id);
  }
  if (rule.priority <= 0) {
    return refuse_write("xlsb conditional format x14 rule needs a positive priority");
  }
  std::array<std::uint32_t, 2> cb{};
  std::vector<std::uint8_t> list;
  std::uint32_t count = 0;
  if (cf_rule_has_formulas(rule.type)) {
    const std::optional<std::string>* sources[2] = {&rule.formula1, &rule.formula2};
    for (std::size_t i = 0; i < 2U; ++i) {
      if (!sources[i]->has_value() || (*sources[i])->empty()) {
        continue;
      }
      auto f = encode_feature_formula(**sources[i], base.first_row, base.first_col, PtgRootClass::kValue, ctx,
                                      PtgEvaluation::kConditionalFormat);
      if (!f) {
        return std::move(f.error());
      }
      emit_frt_formula(list, f.value());
      const std::vector<std::uint8_t>& bytes = f.value().bytes;
      const std::optional<std::uint32_t> size =
          bytes.size() >= 4U ? frt_formula_size_field(bytes.data() + 4, load_u32(bytes.data())) : std::nullopt;
      cb[i] = size.value_or(f.value().cb_fmla);
      ++count;
    }
  }
  std::vector<std::uint8_t> p;
  emit_u32(p, count != 0U ? kFrtHeaderFormulas : 0U);
  if (count != 0U) {
    emit_u32(p, count);
    p.insert(p.end(), list.begin(), list.end());
  }
  emit_u32(p, fields.itype);
  emit_u32(p, fields.tmpl);
  emit_u32(p, rule.dxf_id.value_or(kNoDxf));
  emit_u32(p, static_cast<std::uint32_t>(rule.priority));
  emit_u32(p, fields.param);
  emit_u32(p, 0U);
  emit_u32(p, 0U);
  emit_u16(p, fields.flags);
  emit_u32(p, cb[0]);
  emit_u32(p, cb[1]);
  emit_u32(p, 0U);
  const bool has_text =
      rule.text.has_value() && (rule.type == cf::RuleType::ContainsText || rule.type == cf::RuleType::NotContainsText ||
                                rule.type == cf::RuleType::BeginsWith || rule.type == cf::RuleType::EndsWith);
  emit_xlnullablewidestring(p, has_text ? std::string_view(*rule.text) : std::string_view());
  p.insert(p.end(), guid.begin(), guid.end());
  emit_u32(p, 1U);
  emit_record(dst, kBeginCfRule14, p);
  dst.insert(dst.end(), dxf_copy.begin(), dxf_copy.end());
  emit_record(dst, kEndCfRule14, ByteSpan{});
  return Expected<void, Error>::Ok();
}

}  // namespace

bool is_x14_icon_set_rule(const cf::CFRule& rule) {
  return rule.type == cf::RuleType::IconSet && rule.icon_set.has_value() && !rule.icon_set->custom_icons.empty();
}

bool is_x14_only_rule(const cf::CFRule& rule, std::string_view own_sheet, const std::vector<std::string>* sheet_list) {
  if (is_x14_icon_set_rule(rule)) {
    return true;
  }
  if (IsVisualRule(rule.type)) {
    return false;
  }
  const auto names_other = [&](const std::optional<std::string>& formula) {
    return formula.has_value() && parser::formula_references_other_sheet(*formula, own_sheet, sheet_list);
  };
  return !own_sheet.empty() && (names_other(rule.formula1) || names_other(rule.formula2));
}

void take_x14_rules(std::vector<std::uint8_t>& records, const FeatureFormulaReadContext& ctx,
                    std::vector<cf::ConditionalFormat>& formats) {
  std::unordered_set<std::string> model_ids;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty()) {
        model_ids.insert(rule.id);
      }
    }
  }
  // The model block the current x14 block's formula rules join.
  std::size_t open_block = std::numeric_limits<std::size_t>::max();
  std::size_t open_format = 0;
  RewriteX14Rules(records, [&](const std::vector<FramedRecord>& recs, std::size_t first, std::size_t last,
                               const std::vector<MergeRange>& ranges, std::size_t block) {
    RuleOutcome outcome;
    if (ranges.empty()) {
      return outcome;
    }
    const MergeRange base = sqref_base(ranges);
    std::optional<cf::CFRule> rule = DecodeIconRule14(recs, first, last, base, ctx);
    if (!rule) {
      std::optional<cf::CFRule> formula_rule = DecodeFormulaRule14(recs, first, last, base, ctx);
      if (!formula_rule || model_ids.count(formula_rule->id) != 0U) {
        return outcome;
      }
      model_ids.insert(formula_rule->id);
      if (block != open_block) {
        cf::ConditionalFormat format;
        for (const MergeRange& r : ranges) {
          format.sqref.push_back(cf::CFCellRange{{r.first_row, r.first_col}, {r.last_row, r.last_col}});
        }
        formats.push_back(std::move(format));
        open_block = block;
        open_format = formats.size() - 1U;
      }
      formats[open_format].rules.push_back(std::move(*formula_rule));
      // The records stay as the placeholder `reconcile_x14_rules` rewrites,
      // so the sheet references they index keep their table.
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

Expected<void, Error> reconcile_x14_rules(std::vector<std::uint8_t>& records,
                                          const std::vector<cf::ConditionalFormat>& formats,
                                          const FeatureFormulaWriteContext& ctx, std::string_view own_sheet,
                                          const std::vector<XlsbExternSheetEntry>& extern_sheets,
                                          std::unordered_set<std::string>& linked) {
  std::unordered_map<std::string, std::pair<const cf::ConditionalFormat*, const cf::CFRule*>> model;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty() && !is_x14_icon_set_rule(rule) && is_x14_only_rule(rule, own_sheet, &ctx.sheet_names)) {
        model.emplace(rule.id, std::make_pair(&format, &rule));
      }
    }
  }
  const FeatureFormulaLoadView load_view(ctx, extern_sheets, own_sheet);
  const FeatureFormulaReadContext& read_ctx = load_view.context();
  Expected<void, Error> status = Expected<void, Error>::Ok();
  RewriteX14Rules(records, [&](const std::vector<FramedRecord>& recs, std::size_t first, std::size_t last,
                               const std::vector<MergeRange>& ranges, std::size_t) {
    RuleOutcome outcome;
    cf::CFRule kind;
    const std::optional<Rule14Head> head = ParseRule14Head(recs[first].payload, kind);
    if (!status || ranges.empty() || !head || !IsDxfCopy(recs, first + 1U, last, head->dxf != kNoDxf)) {
      return outcome;
    }
    const auto it = model.find(head->id);
    if (it == model.end()) {
      // Only a rule the reader adopted, since removed from the model or no
      // longer x14-only (the base records carry it), leaves; one it could not
      // adopt stays verbatim.
      if (DecodeFormulaRule14(recs, first, last, sqref_base(ranges), read_ctx).has_value()) {
        outcome.kind = RuleOutcome::Kind::kDrop;
      }
      return outcome;
    }
    // Ranges moved apart from this block's: `add_x14_rules` writes it afresh.
    outcome.kind = RuleOutcome::Kind::kDrop;
    if (!SameSqref(it->second.first->sqref, ranges)) {
      return outcome;
    }
    const cf::CFRule& rule = *it->second.second;
    std::vector<std::uint8_t> dxf_copy;
    if (rule.dxf_id.value_or(kNoDxf) == head->dxf) {
      for (std::size_t i = first + 1U; i < last; ++i) {
        append_record(dxf_copy, records, recs[i]);
      }
    }
    std::vector<std::uint8_t> bytes;
    if (auto s = EmitFormulaRule14(bytes, rule, head->id, sqref_base(ranges), ctx, dxf_copy); !s) {
      status = std::move(s);
      outcome.kind = RuleOutcome::Kind::kKeep;
      return outcome;
    }
    outcome.kind = RuleOutcome::Kind::kReplace;
    outcome.bytes = std::move(bytes);
    linked.insert(head->id);
    return outcome;
  });
  return status;
}

Expected<void, Error> add_x14_rules(std::vector<std::uint8_t>& records,
                                    const std::vector<cf::ConditionalFormat>& formats,
                                    const FeatureFormulaWriteContext& ctx, std::string_view own_sheet,
                                    std::size_t sheet_index, const std::unordered_set<std::string>& linked) {
  std::unordered_set<std::string> used_ids;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty()) {
        used_ids.insert(rule.id);
      }
    }
  }
  std::vector<std::uint8_t> groups;
  for (const cf::ConditionalFormat& format : formats) {
    if (format.sqref.empty()) {
      continue;
    }
    std::vector<MergeRange> ranges;
    for (const cf::CFCellRange& r : format.sqref) {
      ranges.push_back(MergeRange{r.first.row, r.first.col, r.last.row, r.last.col});
    }
    const MergeRange base = sqref_base(ranges);
    const auto icon_rules =
        static_cast<std::uint32_t>(std::count_if(format.rules.begin(), format.rules.end(), is_x14_icon_set_rule));
    if (icon_rules != 0U) {
      EmitBlockHead(groups, format, icon_rules);
      for (const cf::CFRule& rule : format.rules) {
        if (!is_x14_icon_set_rule(rule)) {
          continue;
        }
        if (auto s = EmitIconRule14(groups, rule, base, ctx); !s) {
          return s;
        }
      }
      emit_record(groups, kEndCondFmt14, ByteSpan{});
    }
    for (const cf::CFRule& rule : format.rules) {
      if (is_x14_icon_set_rule(rule) || !is_x14_only_rule(rule, own_sheet, &ctx.sheet_names) ||
          (!rule.id.empty() && linked.count(rule.id) != 0U)) {
        continue;
      }
      std::string id = rule.id;
      if (id.empty()) {
        id = cf::x14_rule_id(sheet_index, rule.priority > 0 ? static_cast<std::uint32_t>(rule.priority) : 0U, used_ids);
        used_ids.insert(id);
      }
      EmitBlockHead(groups, format, 1U);
      if (auto s = EmitFormulaRule14(groups, rule, id, base, ctx); !s) {
        return s;
      }
      emit_record(groups, kEndCondFmt14, ByteSpan{});
    }
  }
  // Excel orders rules stored in the extension alone before data-bar payloads.
  if (!groups.empty()) {
    InsertX14Blocks(records, groups, /*front=*/true);
  }
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
