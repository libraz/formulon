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

#include "cf/cf_rule_traits.h"
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
// 50; BrtBeginDataBar14 is a zero u32, minLength, maxLength, showValue,
// direction (0 context, 1 left-to-right, 2 right-to-left), axis
// position (0 automatic, 1 middle, 2 none) and a u16
// flag word; each BrtColor14 is a zero u32 and a BrtColor, positional in
// the order border, negative fill, negative border, axis. A standalone rule
// (iPri > 0) leads its run with the fill colour.
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

namespace {

/// A linked data-bar payload head: iPri 0xFFFFFFFF and iTemplate 0, as opposed to a standalone rule's iPri > 0.
bool IsLinkedDataBarHead(ByteSpan p) {
  return p.size == kDataBarRule14Bytes && p.data[4] == 4U && static_cast<std::int32_t>(load_u32(p.data + 16)) <= 0 &&
         load_u32(p.data + 8) == 0U;
}

}  // namespace

void apply_x14_data_bar_overlays(ByteSpan records, std::vector<cf::ConditionalFormat>& formats) {
  // State of the x14 rule being read; `measured` drops to false at the
  // first record outside the supported data-bar shape.
  struct Rule14 {
    bool measured = false;
    bool has_bar = false;
    bool show_value = true;
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
      rule14.measured = IsLinkedDataBarHead(rec.payload);
      if (rule14.measured) {
        rule14.id = FormatGuid(d + kRule14GuidOffset);
      }
    } else if (rec.type == kBeginDataBar14 && rule14.measured) {
      rule14.measured = rec.payload.size == kDataBar14Bytes && d[0] == 0U && d[1] == 0U && d[2] == 0U && d[3] == 0U &&
                        d[4] <= 100U && d[5] <= 100U && d[6] <= 1U && d[7] <= 2U && d[8] <= 2U &&
                        (d[9] & 0xF0U) == 0U && d[10] == 0U;
      if (rule14.measured) {
        rule14.has_bar = true;
        rule14.show_value = d[6] == 1U;
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
          // The payload mirrors the legacy rule's showValue.
          if (rule.id != rule14.id || !rule.data_bar || rule14.colors.size() != expected ||
              rule.data_bar->show_value != rule14.show_value) {
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
  return rec.type == kBeginCfRule14 && IsLinkedDataBarHead(rec.payload);
}

/// BrtBeginDataBar14 and the BrtColor14 run for `bar` after the fill. `show_value` is the record's showValue byte.
void EmitDataBar14(std::vector<std::uint8_t>& head, std::vector<std::uint8_t>& colors, const cf::DataBarSpec& bar,
                   std::uint8_t show_value) {
  const bool border = bar.border.has_value();
  const bool negative_fill = bar.negative_fill != bar.fill;
  const bool negative_border = border && bar.negative_border.has_value();
  const std::uint16_t flags = static_cast<std::uint16_t>(
      (border ? kBar14Border : 0U) | (bar.gradient ? kBar14Gradient : 0U) | (negative_fill ? kBar14NegativeFill : 0U) |
      (negative_border ? kBar14NegativeBorder : 0U));
  std::vector<std::uint8_t> p(4U, 0U);
  p.push_back(bar.min_length_pct);
  p.push_back(bar.max_length_pct);
  p.push_back(show_value);
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
  return read_sqref(cursor, ranges) && cursor.size == 8U && load_u32(cursor.data + 4) <= 1U;
}

/// fPivot of a head `ParseBlockHead` accepts: the block holds pivot-scoped rules.
bool BlockHeadPivot(const FramedRecord& rec) {
  return load_u32(rec.payload.data + rec.payload.size - 4U) == 1U;
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
void EmitBlockHead(std::vector<std::uint8_t>& dst, const cf::ConditionalFormat& format, std::uint32_t rules,
                   bool pivot) {
  std::vector<std::uint8_t> head = {0x02, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00};
  std::vector<MergeRange> ranges;
  for (const cf::CFCellRange& r : format.sqref) {
    ranges.push_back(MergeRange{r.first.row, r.first.col, r.last.row, r.last.col});
  }
  emit_sqref(head, ranges);
  emit_u32(head, rules);
  emit_u32(head, pivot ? 1U : 0U);
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
  emit_frt_begin(block, kFrtVersionWord);
  emit_record(block, open, ByteSpan{});
  block.insert(block.end(), body.begin(), body.end());
  emit_record(block, close, ByteSpan{});
  emit_record(block, kFrtEnd, ByteSpan{});
  return block;
}

void reconcile_x14_data_bars(std::vector<std::uint8_t>& records, const std::vector<cf::ConditionalFormat>& formats,
                             std::unordered_set<std::string>& linked, std::string_view own_sheet,
                             const std::vector<std::string>* sheet_list) {
  std::unordered_map<std::string, const cf::DataBarSpec*> model;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty() && rule.data_bar && !is_x14_only_rule(rule, own_sheet, sheet_list)) {
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
                                        std::unordered_set<std::string>& linked, std::size_t sheet_index,
                                        std::string_view own_sheet, const std::vector<std::string>* sheet_list) {
  std::vector<std::uint8_t> groups;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (rule.id.empty() || !rule.data_bar || linked.count(rule.id) != 0U || !data_bar_needs_x14(*rule.data_bar) ||
          is_x14_only_rule(rule, own_sheet, sheet_list)) {
        continue;
      }
      std::array<std::uint8_t, kGuidBytes> guid{};
      if (!ParseGuid(rule.id, guid)) {
        return refuse_write("XLSB conditional format x14 link id is not a GUID: " + rule.id);
      }
      const cf::DataBarSpec& bar = *rule.data_bar;
      EmitBlockHead(groups, format, 1U, /*pivot=*/false);
      std::vector<std::uint8_t> rule14(kDataBarRule14Bytes, 0U);
      rule14[4] = 4U;
      std::fill(rule14.begin() + 16, rule14.begin() + 20, std::uint8_t{0xFF});
      std::copy(guid.begin(), guid.end(), rule14.begin() + static_cast<std::ptrdiff_t>(kRule14GuidOffset));
      rule14[66] = 1U;
      emit_record(groups, kBeginCfRule14, rule14);
      std::vector<std::uint8_t> colors;
      EmitDataBar14(groups, colors, bar, static_cast<std::uint8_t>(bar.show_value ? 1U : 0U));
      for (const cf::CfValueObject* v : {&bar.min, &bar.max}) {
        double num = 0;
        if (v->type == cf::CfvoType::Formula) {
          return not_writable("data_bar", sheet_index);
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
    return refuse_write("XLSB conditional format x14 rule id is not a GUID: " + id);
  }
  if (rule.priority <= 0) {
    return refuse_write("XLSB conditional format x14 rule needs a positive priority");
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
      ASSIGN_OR_RETURN(const EncodedFeatureFormula f, encode_cf_formula(**sources[i], base, ctx));
      emit_frt_formula(list, f);
      const std::vector<std::uint8_t>& bytes = f.bytes;
      const std::optional<std::uint32_t> size =
          bytes.size() >= 4U ? frt_formula_size_field(bytes.data() + 4, load_u32(bytes.data())) : std::nullopt;
      cb[i] = size.value_or(f.cb_fmla);
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

namespace {

void EmitColor14(std::vector<std::uint8_t>& dst, const cf::Color& c) {
  std::vector<std::uint8_t> payload(4U, 0U);
  const auto bytes = EncodeColor(c);
  payload.insert(payload.end(), bytes.begin(), bytes.end());
  emit_record(dst, kColor14, payload);
}

std::optional<cf::Color> DecodeColor14(const FramedRecord& rec) {
  const ByteSpan p = rec.payload;
  if (rec.type != kColor14 || p.size != 12U || load_u32(p.data) != 0U) {
    return std::nullopt;
  }
  return DecodeColor(ByteSpan{p.data + 4, 8U});
}

/// `rule` as a standalone data bar: BrtBeginCFRule14, BrtBeginDataBar14, two BrtCFVO14, the BrtColor14 run (fill
/// first) and BrtEndDataBar14.
Expected<void, Error> EmitDataBarRule14(std::vector<std::uint8_t>& dst, const cf::CFRule& rule, const std::string& id,
                                        const MergeRange& base, const FeatureFormulaWriteContext& ctx) {
  if (!rule.data_bar) {
    return refuse_write("XLSB data bar rule has no data bar");
  }
  const cf::DataBarSpec& bar = *rule.data_bar;
  if (bar.axis_position == cf::DataBarAxisPosition::Middle) {
    return refuse_write("XLSB x14 data bar with a middle axis has no record form");
  }
  if (bar.negative_border.has_value() && !bar.border.has_value()) {
    return refuse_write(
        "XLSB x14 data bar with a negative border and no border has no record form: Excel moves that colour into the "
        "axis slot and loses the axis colour");
  }
  ASSIGN_OR_RETURN(const auto guid, detail::VisualRule14Guid(rule, id));
  detail::EmitRule14Head(dst, detail::kRule14TypeDataBar, detail::kRule14TemplateDataBar, rule.priority, guid);
  std::vector<std::uint8_t> colors;
  EmitDataBar14(dst, colors, bar, static_cast<std::uint8_t>(bar.show_value ? 1U : 0U));
  for (const cf::CfValueObject* v : {&bar.min, &bar.max}) {
    if (auto s = detail::EmitCfvo14(dst, *v, /*icon=*/false, base, ctx); !s) {
      return s;
    }
  }
  EmitColor14(dst, bar.fill);
  dst.insert(dst.end(), colors.begin(), colors.end());
  emit_record(dst, kEndDataBar14, ByteSpan{});
  emit_record(dst, kEndCfRule14, ByteSpan{});
  return Expected<void, Error>::Ok();
}

/// `rule` as a standalone colour scale: BrtBeginCFRule14, BrtBeginColorScale14, one BrtCFVO14 and one BrtColor14
/// per stop, BrtEndColorScale14.
Expected<void, Error> EmitColorScaleRule14(std::vector<std::uint8_t>& dst, const cf::CFRule& rule,
                                           const std::string& id, const MergeRange& base,
                                           const FeatureFormulaWriteContext& ctx) {
  if (!rule.color_scale || rule.color_scale->colors.size() != rule.color_scale->thresholds.size() ||
      rule.color_scale->colors.size() < 2U || rule.color_scale->colors.size() > 3U) {
    return refuse_write("XLSB colour scale needs two or three matched stops");
  }
  ASSIGN_OR_RETURN(const auto guid, detail::VisualRule14Guid(rule, id));
  detail::EmitRule14Head(dst, detail::kRule14TypeColorScale, detail::kRule14TemplateColorScale, rule.priority, guid);
  emit_record(dst, detail::kBeginColorScale14, ByteSpan{});
  for (const cf::CfValueObject& v : rule.color_scale->thresholds) {
    if (auto s = detail::EmitCfvo14(dst, v, /*icon=*/false, base, ctx); !s) {
      return s;
    }
  }
  for (const cf::Color& c : rule.color_scale->colors) {
    EmitColor14(dst, c);
  }
  emit_record(dst, detail::kEndColorScale14, ByteSpan{});
  emit_record(dst, kEndCfRule14, ByteSpan{});
  return Expected<void, Error>::Ok();
}

/// The standalone data bar `recs[first..last]`, or nothing for any record outside the shape
/// `EmitDataBarRule14` writes.
std::optional<cf::CFRule> DecodeDataBarRule14(const std::vector<FramedRecord>& recs, std::size_t first,
                                              std::size_t last, const MergeRange& base,
                                              const FeatureFormulaReadContext& ctx) {
  std::int32_t priority = 0;
  std::string id;
  if (last < first + 5U || recs[first + 1U].type != kBeginDataBar14 ||
      !detail::ParseVisualRule14Head(recs[first].payload, detail::kRule14TypeDataBar, detail::kRule14TemplateDataBar,
                                     priority, id)) {
    return std::nullopt;
  }
  const ByteSpan p = recs[first + 1U].payload;
  const std::uint8_t* d = p.data;
  if (p.size != kDataBar14Bytes || load_u32(d) != 0U || d[4] > 100U || d[5] > 100U || d[6] > 1U || d[7] > 2U ||
      d[8] == 1U || d[8] > 2U || (d[9] & 0xF0U) != 0U || d[10] != 0U) {
    return std::nullopt;
  }
  const bool border = (d[9] & kBar14Border) != 0U;
  const bool negative_fill = (d[9] & kBar14NegativeFill) != 0U;
  const bool negative_border = (d[9] & kBar14NegativeBorder) != 0U;
  const bool axis = d[8] != static_cast<std::uint8_t>(cf::DataBarAxisPosition::None);
  if (negative_border && !border) {
    return std::nullopt;
  }
  const std::size_t colors = 1U + static_cast<std::size_t>(border) + static_cast<std::size_t>(negative_fill) +
                             static_cast<std::size_t>(negative_border) + static_cast<std::size_t>(axis);
  if (last != first + 5U + colors || recs[last - 1U].type != kEndDataBar14 || recs[last - 1U].payload.size != 0U ||
      recs[last].payload.size != 0U) {
    return std::nullopt;
  }
  cf::DataBarSpec bar;
  bar.min_length_pct = d[4];
  bar.max_length_pct = d[5];
  bar.show_value = d[6] == 1U;
  bar.direction = static_cast<cf::DataBarDirection>(d[7]);
  bar.axis_position = static_cast<cf::DataBarAxisPosition>(d[8]);
  bar.gradient = (d[9] & kBar14Gradient) != 0U;
  for (std::size_t i = 0; i < 2U; ++i) {
    const FramedRecord& rec = recs[first + 2U + i];
    std::optional<cf::CfValueObject> v =
        rec.type == kCfvo14 ? detail::DecodeCfvo14(rec.payload, /*icon=*/false, base, ctx) : std::nullopt;
    if (!v) {
      return std::nullopt;
    }
    (i == 0U ? bar.min : bar.max) = std::move(*v);
  }
  std::vector<cf::Color> run;
  for (std::size_t i = 0; i < colors; ++i) {
    std::optional<cf::Color> c = DecodeColor14(recs[first + 4U + i]);
    if (!c) {
      return std::nullopt;
    }
    run.push_back(*c);
  }
  std::size_t next = 0;
  bar.fill = run[next++];
  if (border) {
    bar.border = run[next++];
  }
  bar.negative_fill = negative_fill ? run[next++] : bar.fill;
  if (negative_border) {
    bar.negative_border = run[next++];
  }
  if (axis) {
    bar.axis_color = run[next++];
  }
  cf::CFRule rule;
  rule.type = cf::RuleType::DataBar;
  rule.priority = priority;
  rule.id = std::move(id);
  rule.data_bar = std::move(bar);
  return rule;
}

/// The standalone colour scale `recs[first..last]`, or nothing for any record outside the shape
/// `EmitColorScaleRule14` writes.
std::optional<cf::CFRule> DecodeColorScaleRule14(const std::vector<FramedRecord>& recs, std::size_t first,
                                                 std::size_t last, const MergeRange& base,
                                                 const FeatureFormulaReadContext& ctx) {
  std::int32_t priority = 0;
  std::string id;
  const std::size_t span = last - first;
  if (last < first + 7U || (span - 3U) % 2U != 0U || span > 9U ||
      !detail::ParseVisualRule14Head(recs[first].payload, detail::kRule14TypeColorScale,
                                     detail::kRule14TemplateColorScale, priority, id)) {
    return std::nullopt;
  }
  const std::size_t stops = (span - 3U) / 2U;
  if (recs[first + 1U].type != detail::kBeginColorScale14 || recs[first + 1U].payload.size != 0U ||
      recs[last - 1U].type != detail::kEndColorScale14 || recs[last - 1U].payload.size != 0U ||
      recs[last].payload.size != 0U) {
    return std::nullopt;
  }
  cf::ColorScaleSpec scale;
  for (std::size_t i = 0; i < stops; ++i) {
    const FramedRecord& rec = recs[first + 2U + i];
    std::optional<cf::CfValueObject> v =
        rec.type == kCfvo14 ? detail::DecodeCfvo14(rec.payload, /*icon=*/false, base, ctx) : std::nullopt;
    std::optional<cf::Color> c = DecodeColor14(recs[first + 2U + stops + i]);
    if (!v || !c) {
      return std::nullopt;
    }
    scale.thresholds.push_back(std::move(*v));
    scale.colors.push_back(*c);
  }
  cf::CFRule rule;
  rule.type = cf::RuleType::ColorScale;
  rule.priority = priority;
  rule.id = std::move(id);
  rule.color_scale = std::move(scale);
  return rule;
}

}  // namespace

bool is_x14_icon_set_rule(const cf::CFRule& rule) {
  return rule.type == cf::RuleType::IconSet && rule.icon_set.has_value() && !rule.icon_set->custom_icons.empty();
}

bool is_x14_only_rule(const cf::CFRule& rule, std::string_view own_sheet, const std::vector<std::string>* sheet_list) {
  return is_x14_icon_set_rule(rule) || cf::rule_is_x14_only(rule, own_sheet, sheet_list);
}

bool is_x14_visual_rule(const cf::CFRule& rule, std::string_view own_sheet,
                        const std::vector<std::string>* sheet_list) {
  return IsVisualRule(rule.type) && is_x14_only_rule(rule, own_sheet, sheet_list);
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
    const bool pivot = BlockHeadPivot(recs[block]);
    const MergeRange base = sqref_base(ranges);
    std::optional<cf::CFRule> rule = DecodeIconRule14(recs, first, last, base, ctx);
    if (!rule) {
      rule = DecodeDataBarRule14(recs, first, last, base, ctx);
    }
    if (!rule) {
      rule = DecodeColorScaleRule14(recs, first, last, base, ctx);
    }
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
        format.pivot_scope = pivot;
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
    const bool custom_icons = rule->icon_set.has_value() && !rule->icon_set->custom_icons.empty();
    for (cf::ConditionalFormat& format : formats) {
      if (!custom_icons || format.pivot_scope != pivot || !SameSqref(format.sqref, ranges)) {
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
    const auto same = std::find_if(formats.begin(), formats.end(), [&ranges, pivot](const cf::ConditionalFormat& f) {
      return f.pivot_scope == pivot && SameSqref(f.sqref, ranges);
    });
    if (same != formats.end()) {
      same->rules.push_back(std::move(*rule));
    } else {
      cf::ConditionalFormat format;
      for (const MergeRange& r : ranges) {
        format.sqref.push_back(cf::CFCellRange{{r.first_row, r.first_col}, {r.last_row, r.last_col}});
      }
      format.pivot_scope = pivot;
      format.rules.push_back(std::move(*rule));
      formats.push_back(std::move(format));
    }
    outcome.kind = RuleOutcome::Kind::kDrop;
    return outcome;
  });
}

bool tail_holds_unmodelled_cf(const XlsbSheetTail& tail, const std::vector<cf::ConditionalFormat>& formats) {
  std::unordered_set<std::string> model_ids;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty()) {
        model_ids.insert(rule.id);
      }
    }
  }
  for (const std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    ByteSpan cursor{slot->data(), slot->size()};
    while (cursor.size != 0U) {
      auto rec = read_record(cursor);
      if (!rec) {
        return true;
      }
      const ByteSpan p = rec.value().payload;
      if (rec.value().type == kBrtBeginConditionalFormatting) {
        return true;
      }
      // Every BrtBeginCFRule14 ends with the rule GUID and fGuid.
      if (rec.value().type == kBeginCfRule14 &&
          (p.size < kRule14TailBytes || model_ids.count(FormatGuid(p.data + p.size - kRule14TailBytes)) == 0U)) {
        return true;
      }
    }
  }
  return false;
}

namespace {

/// iPri of a BrtBeginCFRule14 payload `p`, past its optional FRTFormula list; false when the head does not frame.
bool Rule14Priority(ByteSpan p, std::uint32_t& out) {
  auto header = read_u32(p);
  if (!header || (header.value() != 0U && header.value() != kFrtHeaderFormulas)) {
    return false;
  }
  if (header.value() == kFrtHeaderFormulas) {
    auto count = read_u32(p);
    if (!count) {
      return false;
    }
    for (std::uint32_t i = 0; i < count.value(); ++i) {
      if (p.size < 12U) {
        return false;
      }
      const std::size_t cce = load_u32(p.data + 4);
      const std::size_t cb = load_u32(p.data + 8);
      if (p.size - 12U < cce || p.size - 12U - cce < cb) {
        return false;
      }
      const std::size_t size = 12U + cce + cb;
      p.data += size;
      p.size -= size;
    }
  }
  // iType, iTemplate and dxfId precede iPri.
  if (!read_u32(p) || !read_u32(p) || !read_u32(p)) {
    return false;
  }
  auto pri = read_u32(p);
  if (!pri) {
    return false;
  }
  out = pri.value();
  return true;
}

}  // namespace

bool retained_cf_priority_max(const XlsbSheetTail& tail, const std::vector<cf::ConditionalFormat>& formats,
                              std::int64_t* out_max) {
  std::unordered_set<std::string> model_ids;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty()) {
        model_ids.insert(rule.id);
      }
    }
  }
  std::int64_t max = 0;
  for (const std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    ByteSpan cursor{slot->data(), slot->size()};
    while (cursor.size != 0U) {
      auto rec = read_record(cursor);
      if (!rec) {
        return false;
      }
      const ByteSpan p = rec.value().payload;
      if (rec.value().type == kBrtBeginCFRule) {
        // iType, iTemplate, dxfId, then iPri.
        if (p.size < 16U) {
          return false;
        }
        max = std::max<std::int64_t>(max, load_u32(p.data + 12));
      } else if (rec.value().type == kBeginCfRule14) {
        cf::CFRule scratch;
        const std::optional<Rule14Head> head = ParseRule14Head(p, scratch);
        if (head && model_ids.count(head->id) != 0U) {
          // Rewritten from the model on save.
          continue;
        }
        std::uint32_t pri = 0;
        if (!Rule14Priority(p, pri)) {
          return false;
        }
        // A linked data-bar payload carries no priority (iPri 0xFFFFFFFF).
        if (static_cast<std::int32_t>(pri) > 0) {
          max = std::max<std::int64_t>(max, pri);
        }
      }
    }
  }
  *out_max = max;
  return true;
}

void collect_x14_rule_ids(const std::vector<std::uint8_t>& records, std::unordered_set<std::string>& ids) {
  ByteSpan cursor{records.data(), records.size()};
  while (cursor.size != 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      return;
    }
    const ByteSpan p = rec.value().payload;
    // Every BrtBeginCFRule14 ends with the rule GUID and fGuid.
    if (rec.value().type == kBeginCfRule14 && p.size >= kRule14TailBytes && load_u32(p.data + p.size - 4U) == 1U) {
      ids.insert(FormatGuid(p.data + p.size - kRule14TailBytes));
    }
  }
}

Expected<void, Error> reconcile_x14_rules(std::vector<std::uint8_t>& records,
                                          const std::vector<cf::ConditionalFormat>& formats,
                                          const FeatureFormulaWriteContext& ctx, std::string_view own_sheet,
                                          const std::vector<XlsbExternSheetEntry>& extern_sheets,
                                          std::unordered_set<std::string>& linked) {
  std::unordered_map<std::string, std::pair<const cf::ConditionalFormat*, const cf::CFRule*>> model;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty() && !IsVisualRule(rule.type) && is_x14_only_rule(rule, own_sheet, &ctx.sheet_names)) {
        model.emplace(rule.id, std::make_pair(&format, &rule));
      }
    }
  }
  const FeatureFormulaLoadView load_view(ctx, extern_sheets, own_sheet);
  const FeatureFormulaReadContext& read_ctx = load_view.context();
  Expected<void, Error> status = Expected<void, Error>::Ok();
  RewriteX14Rules(records, [&](const std::vector<FramedRecord>& recs, std::size_t first, std::size_t last,
                               const std::vector<MergeRange>& ranges, std::size_t block) {
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
    // Ranges or scope moved apart from this block's: `add_x14_rules` writes it afresh.
    outcome.kind = RuleOutcome::Kind::kDrop;
    if (!SameSqref(it->second.first->sqref, ranges) || it->second.first->pivot_scope != BlockHeadPivot(recs[block])) {
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
                                    std::size_t sheet_index, const std::unordered_set<std::string>& linked,
                                    const std::unordered_set<std::string>& retained_ids) {
  std::unordered_set<std::string> used_ids = retained_ids;
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
    const auto is_visual = [&](const cf::CFRule& rule) {
      return is_x14_visual_rule(rule, own_sheet, &ctx.sheet_names);
    };
    const auto visual_rules =
        static_cast<std::uint32_t>(std::count_if(format.rules.begin(), format.rules.end(), is_visual));
    if (visual_rules != 0U) {
      EmitBlockHead(groups, format, visual_rules, format.pivot_scope);
      for (const cf::CFRule& rule : format.rules) {
        if (!is_visual(rule)) {
          continue;
        }
        std::string id = rule.id;
        if (id.empty()) {
          id = cf::x14_rule_id(sheet_index, rule.priority > 0 ? static_cast<std::uint32_t>(rule.priority) : 0U,
                               used_ids);
          used_ids.insert(id);
        }
        Expected<void, Error> status =
            rule.type == cf::RuleType::DataBar      ? EmitDataBarRule14(groups, rule, id, base, ctx)
            : rule.type == cf::RuleType::ColorScale ? EmitColorScaleRule14(groups, rule, id, base, ctx)
                                                    : EmitIconRule14(groups, rule, id, base, ctx);
        if (!status) {
          return status;
        }
      }
      emit_record(groups, kEndCondFmt14, ByteSpan{});
    }
    for (const cf::CFRule& rule : format.rules) {
      if (is_visual(rule) || !is_x14_only_rule(rule, own_sheet, &ctx.sheet_names) ||
          (!rule.id.empty() && linked.count(rule.id) != 0U)) {
        continue;
      }
      std::string id = rule.id;
      if (id.empty()) {
        id = cf::x14_rule_id(sheet_index, rule.priority > 0 ? static_cast<std::uint32_t>(rule.priority) : 0U, used_ids);
        used_ids.insert(id);
      }
      EmitBlockHead(groups, format, 1U, format.pivot_scope);
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

namespace detail {

Rule14Kind ReadRule14Kind(ByteSpan head) {
  Rule14Kind kind;
  auto header = read_u32(head);
  if (!header || (header.value() != 0U && header.value() != kFrtHeaderFormulas)) {
    return kind;
  }
  if (header.value() == kFrtHeaderFormulas) {
    auto count = read_u32(head);
    if (!count) {
      return kind;
    }
    for (std::uint32_t i = 0; i < count.value(); ++i) {
      if (head.size < 12U) {
        return kind;
      }
      const std::size_t size = 12U + static_cast<std::size_t>(load_u32(head.data + 4)) + load_u32(head.data + 8);
      if (head.size < size) {
        return kind;
      }
      head.data += size;
      head.size -= size;
    }
  }
  auto itype = read_u32(head);
  auto tmpl = read_u32(head);
  auto dxf = read_u32(head);
  auto pri = read_u32(head);
  auto param = read_u32(head);
  auto reserved1 = read_u32(head);
  auto reserved2 = read_u32(head);
  auto flags = read_u16(head);
  auto cb1 = read_u32(head);
  auto cb2 = read_u32(head);
  auto cb3 = read_u32(head);
  if (!itype || !tmpl || !dxf || !pri || !param || !reserved1 || !reserved2 || !flags || !cb1 || !cb2 || !cb3) {
    return kind;
  }
  cf::CFRule rule;
  if (decode_cf_rule_fields(CfRuleFields{itype.value(), tmpl.value(), param.value(), flags.value()}, std::nullopt,
                            rule)) {
    kind.type = rule.type;
  }
  // strParam, then the rule GUID and fGuid close the record.
  if (read_xlnullablewidestring(head) && head.size == kRule14TailBytes && load_u32(head.data + kGuidBytes) == 1U) {
    kind.id = FormatGuid(head.data);
  }
  return kind;
}

bool IsFormulaCfvo14(ByteSpan payload) {
  // iType, numParam, fSaveGTE, fGTE and cbFmla close the record.
  constexpr std::size_t kTailBytes = 24U;
  return payload.size >= kTailBytes + 4U && load_u32(payload.data) == kFrtHeaderFormulas &&
         load_u32(payload.data + payload.size - kTailBytes) ==
             kCfvo14Types[static_cast<std::size_t>(cf::CfvoType::Formula)];
}

}  // namespace detail

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
