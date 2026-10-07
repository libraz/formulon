// @size-budget: 15 KB

#include <algorithm>
#include <array>
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
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
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

Error Refuse(const std::string& what) {
  return make_error(FormulonErrorCode::kInvalidArgument, what, "context=write_xlsb");
}

}  // namespace

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

}  // namespace

void reconcile_x14_data_bars(std::vector<std::uint8_t>& records, const std::vector<cf::ConditionalFormat>& formats,
                             std::unordered_set<std::string>& linked) {
  const std::vector<FramedRecord> recs = split_records(records);
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
            append_record(rule, records, recs[m]);
          }
        }
        if (!rule.empty()) {
          block.insert(block.end(), rule.begin(), rule.end());
          ++kept;
        }
        l = rule_end + 1U;
      }
      if (kept != 0U || rules == 0U) {
        append_record(inner, records, recs[k]);
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
      EmitDataBar14(groups, colors, bar, 1U);
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
  const std::vector<FramedRecord> recs = split_records(records);
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

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
