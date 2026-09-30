#include "io/xlsb/tail_refs.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/feature_formula.h"
#include "io/xlsb/ptg.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "parser/reference.h"

namespace formulon::io::xlsb {
namespace {

constexpr std::uint16_t kFrtBegin = 35;
constexpr std::uint16_t kFrtEnd = 36;
constexpr std::uint16_t kAcBegin = 37;
constexpr std::uint16_t kAcEnd = 38;
constexpr std::uint16_t kBeginSparklineGroup = 1041;
constexpr std::uint16_t kEndSparklineGroup = 1042;
constexpr std::uint16_t kSparkline = 1043;
constexpr std::uint16_t kBeginCondFmt14 = 1046;
constexpr std::uint16_t kEndCondFmt14 = 1047;
constexpr std::uint16_t kBeginCfRule14 = 1048;
constexpr std::uint16_t kCfvo14 = 1050;
constexpr std::uint16_t kBeginSparklines = 1056;
constexpr std::uint16_t kEndSparklines = 1057;
constexpr std::uint16_t kBeginSparklineGroups = 1058;
constexpr std::uint16_t kEndSparklineGroups = 1059;
constexpr std::uint16_t kBeginCfs14 = 1135;
constexpr std::uint16_t kEndCfs14 = 1136;

// Future-record header flags: a range list, plus a formula list.
constexpr std::uint32_t kFrtRanges = 0x02U;
constexpr std::uint32_t kFrtRangesAndFormulas = 0x06U;
constexpr std::size_t kSqrefOffset = 12U;

constexpr std::uint16_t kColMask = 0x3FFF;
constexpr std::uint16_t kColRelBit = 0x4000;
constexpr std::uint16_t kRowRelBit = 0x8000;

std::uint32_t LoadU32(const std::uint8_t* p) {
  return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8) |
         (static_cast<std::uint32_t>(p[2]) << 16) | (static_cast<std::uint32_t>(p[3]) << 24);
}

std::uint16_t LoadU16(const std::uint8_t* p) {
  return static_cast<std::uint16_t>(p[0] | (p[1] << 8));
}

void StoreU32(std::uint8_t* p, std::uint32_t v) {
  for (int i = 0; i < 4; ++i) {
    p[i] = static_cast<std::uint8_t>(v >> (8 * i));
  }
}

void StoreU16(std::uint8_t* p, std::uint16_t v) {
  p[0] = static_cast<std::uint8_t>(v);
  p[1] = static_cast<std::uint8_t>(v >> 8);
}

/// The range list of a 1046 / 1043 record in the measured shape; `rest` is
/// the offset of whatever follows the Sqrfx.
bool ParseRanges(const FramedRecord& rec, std::vector<MergeRange>& ranges, std::size_t& rest) {
  const std::uint32_t want = rec.type == kBeginCondFmt14 ? kFrtRanges : kFrtRangesAndFormulas;
  if ((rec.type != kBeginCondFmt14 && rec.type != kSparkline) || rec.payload.size < kSqrefOffset ||
      LoadU32(rec.payload.data) != want || LoadU32(rec.payload.data + 4) != 1U) {
    return false;
  }
  ByteSpan cursor{rec.payload.data + kSqrefOffset, rec.payload.size - kSqrefOffset};
  if (!read_sqref(cursor, ranges)) {
    return false;
  }
  rest = rec.payload.size - cursor.size;
  return true;
}

bool SameRanges(const std::vector<MergeRange>& a, const std::vector<MergeRange>& b) {
  if (a.size() != b.size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.size(); ++i) {
    if (a[i].first_row != b[i].first_row || a[i].first_col != b[i].first_col || a[i].last_row != b[i].last_row ||
        a[i].last_col != b[i].last_col) {
      return false;
    }
  }
  return true;
}

/// A slot's records with a keep mask and per-record replacement payloads.
struct Edit {
  std::vector<FramedRecord> recs;
  std::vector<bool> keep;
  std::vector<std::optional<std::vector<std::uint8_t>>> payload;

  std::size_t Find(std::size_t from, std::uint16_t type) const {
    for (std::size_t i = from; i < recs.size(); ++i) {
      if (recs[i].type == type) {
        return i;
      }
    }
    return recs.size();
  }

  bool AnyKept(std::size_t from, std::size_t until, std::uint16_t type) const {
    for (std::size_t i = from; i < until; ++i) {
      if (keep[i] && recs[i].type == type) {
        return true;
      }
    }
    return false;
  }

  void Drop(std::size_t first, std::size_t last) {
    for (std::size_t i = first; i <= last && i < recs.size(); ++i) {
      keep[i] = false;
    }
  }

  /// Drops each kept `begin`..`end` span holding no kept `member`, with the
  /// future-record brackets around it.
  void DropEmpty(std::uint16_t begin, std::uint16_t end, std::uint16_t member) {
    for (std::size_t i = 0; i < recs.size(); ++i) {
      if (!keep[i] || recs[i].type != begin) {
        continue;
      }
      const std::size_t close = Find(i, end);
      if (close == recs.size() || AnyKept(i, close, member)) {
        continue;
      }
      const bool framed =
          i > 0U && recs[i - 1U].type == kFrtBegin && close + 1U < recs.size() && recs[close + 1U].type == kFrtEnd;
      Drop(framed ? i - 1U : i, framed ? close + 1U : close);
    }
  }

  void Write(std::vector<std::uint8_t>& buf) const {
    std::vector<std::uint8_t> out;
    for (std::size_t i = 0; i < recs.size(); ++i) {
      if (!keep[i]) {
        continue;
      }
      if (payload[i]) {
        emit_record(out, recs[i].type, *payload[i]);
      } else {
        append_record(out, buf, recs[i]);
      }
    }
    buf = std::move(out);
  }
};

Edit Open(const std::vector<std::uint8_t>& buf) {
  Edit edit;
  edit.recs = split_records(buf);
  edit.keep.assign(edit.recs.size(), true);
  edit.payload.resize(edit.recs.size());
  return edit;
}

void RemapSlotSqrefs(std::vector<std::uint8_t>& buf, const SqrefRemap& remap) {
  Edit edit = Open(buf);
  bool changed = false;
  for (std::size_t i = 0; i < edit.recs.size(); ++i) {
    const FramedRecord& rec = edit.recs[i];
    std::vector<MergeRange> ranges;
    std::size_t rest = 0;
    if (!edit.keep[i] || !ParseRanges(rec, ranges, rest)) {
      continue;
    }
    const std::vector<MergeRange> before = ranges;
    remap(ranges);
    if (ranges.empty()) {
      edit.Drop(i, rec.type == kBeginCondFmt14 ? edit.Find(i, kEndCondFmt14) : i);
      changed = true;
      continue;
    }
    if (SameRanges(before, ranges)) {
      continue;
    }
    std::vector<std::uint8_t> out(rec.payload.data, rec.payload.data + kSqrefOffset);
    emit_sqref(out, ranges);
    out.insert(out.end(), rec.payload.data + rest, rec.payload.data + rec.payload.size);
    edit.payload[i] = std::move(out);
    changed = true;
  }
  if (!changed) {
    return;
  }
  // A group left without sparklines goes whole, with the uid block ahead of it.
  for (std::size_t i = 0; i < edit.recs.size(); ++i) {
    if (!edit.keep[i] || edit.recs[i].type != kBeginSparklines) {
      continue;
    }
    const std::size_t close = edit.Find(i, kEndSparklines);
    if (close == edit.recs.size() || edit.AnyKept(i, close, kSparkline)) {
      continue;
    }
    std::size_t group = i;
    while (group > 0U && edit.recs[group].type != kBeginSparklineGroup) {
      --group;
    }
    const std::size_t group_end = edit.Find(close, kEndSparklineGroup);
    if (edit.recs[group].type != kBeginSparklineGroup || group_end == edit.recs.size()) {
      continue;
    }
    std::size_t first = group;
    if (first > 0U && edit.recs[first - 1U].type == kAcEnd) {
      std::size_t ac = first - 1U;
      while (ac > 0U && edit.recs[ac].type != kAcBegin) {
        --ac;
      }
      first = edit.recs[ac].type == kAcBegin ? ac : first;
    }
    edit.Drop(first, group_end);
  }
  edit.DropEmpty(kBeginSparklineGroups, kEndSparklineGroups, kBeginSparklineGroup);
  edit.DropEmpty(kBeginCfs14, kEndCfs14, kBeginCondFmt14);
  edit.Write(buf);
}

/// Offset and length of each rgce in a future record: the header flags, the
/// range list when flagged, then u32 formula count and per formula u32
/// flags, cce, cb, rgce, rgcb. False for a header carrying anything else
/// before the formulas.
bool FormulaSpans(ByteSpan p, std::vector<std::pair<std::size_t, std::size_t>>& spans) {
  spans.clear();
  if (p.size < 4U) {
    return false;
  }
  const std::uint32_t flags = LoadU32(p.data);
  if ((flags & 0x04U) == 0U) {
    return true;
  }
  if ((flags & 0x01U) != 0U) {
    return false;
  }
  ByteSpan cursor{p.data + 4, p.size - 4U};
  if ((flags & 0x02U) != 0U) {
    auto count = read_u32(cursor);
    if (!count) {
      return false;
    }
    for (std::uint32_t i = 0; i < count.value(); ++i) {
      std::vector<MergeRange> ranges;
      if (!read_u32(cursor) || !read_sqref(cursor, ranges)) {
        return false;
      }
    }
  }
  auto count = read_u32(cursor);
  if (!count) {
    return false;
  }
  for (std::uint32_t i = 0; i < count.value(); ++i) {
    if (cursor.size < 12U) {
      return false;
    }
    const std::size_t cce = LoadU32(cursor.data + 4);
    const std::size_t cb = LoadU32(cursor.data + 8);
    if (cursor.size - 12U < cce || cursor.size - 12U - cce < cb) {
      return false;
    }
    spans.emplace_back(static_cast<std::size_t>(cursor.data - p.data) + 12U, cce);
    cursor = ByteSpan{cursor.data + 12U + cce + cb, cursor.size - 12U - cce - cb};
  }
  return true;
}

/// What a formula's references resolve through: the sheet names behind
/// `ixti`, and for a formula anchored to a range, its base cell before and
/// after the edit.
struct RefContext {
  const std::vector<XlsbExternSheetEntry>& sheets;
  const parser::RefTransform& transform;
  std::optional<PtgBaseCell> old_base;
  std::optional<PtgBaseCell> new_base;
};

/// Payload bytes of the token at `p`, or nothing for one this walk does
/// not step over (array constants, mem areas, reserved Ptgs).
std::optional<std::size_t> TokenSize(const std::uint8_t* p, std::size_t left) {
  const std::uint8_t tag = p[0];
  if (tag >= 0x80U) {
    return std::nullopt;
  }
  if (tag < 0x20U) {
    if (tag >= 0x03U && tag <= 0x16U) {
      return 0U;
    }
    switch (tag) {
      case 0x17U:
        return left >= 3U ? std::optional<std::size_t>(2U + 2U * LoadU16(p + 1)) : std::nullopt;
      case 0x19U:
        if (left < 4U) {
          return std::nullopt;
        }
        return (p[1] & 0x04U) != 0U ? 3U + 2U * (LoadU16(p + 2) + 1U) : 3U;
      case 0x1CU:
      case 0x1DU:
        return 1U;
      case 0x1EU:
        return 2U;
      case 0x1FU:
        return 8U;
      default:
        return std::nullopt;
    }
  }
  switch (tag & 0x1FU) {
    case 0x01U:
    case 0x09U:
      return 2U;
    case 0x02U:
      return 3U;
    case 0x03U:
      return 4U;
    case 0x04U:
    case 0x0AU:
    case 0x0CU:
    case 0x19U:
      return 6U;
    case 0x05U:
    case 0x0BU:
    case 0x0DU:
      return 12U;
    case 0x1AU:
    case 0x1CU:
      return 8U;
    case 0x1BU:
    case 0x1DU:
      return 14U;
    default:
      return std::nullopt;
  }
}

/// Rewrites one PtgRef / PtgArea (plain, N or 3d) whose payload starts at
/// `at` through `ctx`. A reference the transform removes becomes the error
/// form of the same Ptg. False when a relative token has no base to resolve
/// against.
bool RemapRefToken(std::uint8_t* tag, const RefContext& ctx) {
  const std::uint8_t kind = *tag & 0x1FU;
  const bool three_d = kind == 0x1AU || kind == 0x1BU;
  const bool relative = kind == 0x0CU || kind == 0x0DU;
  const bool area = kind == 0x05U || kind == 0x0DU || kind == 0x1BU;
  std::uint8_t* at = tag + 1 + (three_d ? 2 : 0);
  std::string_view sheet;
  if (three_d) {
    const std::uint16_t ixti = LoadU16(tag + 1);
    // A multi-sheet span keeps its coordinates under a row/column edit.
    if (ixti >= ctx.sheets.size() || ctx.sheets[ixti].unresolved || ctx.sheets[ixti].first.empty() ||
        ctx.sheets[ixti].first != ctx.sheets[ixti].last) {
      return true;
    }
    sheet = ctx.sheets[ixti].first;
  }
  if (relative && (!ctx.old_base || !ctx.new_base)) {
    return false;
  }
  constexpr std::uint32_t kRowMask = (1U << 20) - 1U;
  const auto load = [&](std::uint32_t row, std::uint16_t col) {
    parser::Reference r;
    r.sheet = sheet;
    r.row = row;
    r.col = col & kColMask;
    r.col_abs = (col & kColRelBit) == 0U;
    r.row_abs = (col & kRowRelBit) == 0U;
    if (relative) {
      r.row = r.row_abs ? r.row : (ctx.old_base->row + r.row) & kRowMask;
      r.col = r.col_abs ? r.col : (ctx.old_base->col + r.col) & kColMask;
    }
    return r;
  };
  const auto store = [&](std::uint8_t* row_at, std::uint8_t* col_at, const parser::Reference& r) {
    std::uint32_t row = r.row;
    std::uint32_t col = r.col;
    if (relative) {
      row = r.row_abs ? row : (row - ctx.new_base->row) & kRowMask;
      col = r.col_abs ? col : (col - ctx.new_base->col) & kColMask;
    }
    StoreU32(row_at, row);
    StoreU16(col_at, static_cast<std::uint16_t>(col | (r.col_abs ? 0U : kColRelBit) | (r.row_abs ? 0U : kRowRelBit)));
  };
  const auto to_error = [&]() {
    const std::uint8_t error_kind = three_d ? kind + 2U : (area ? 0x0BU : 0x0AU);
    *tag = static_cast<std::uint8_t>((*tag & 0x60U) | error_kind);
    std::fill(at, at + (area ? 12 : 6), std::uint8_t{0});
  };
  if (!area) {
    const std::optional<parser::Reference> out = ctx.transform.apply(load(LoadU32(at), LoadU16(at + 4)));
    if (!out) {
      to_error();
    } else {
      store(at, at + 4, *out);
    }
    return true;
  }
  parser::Reference first = load(LoadU32(at), LoadU16(at + 8));
  parser::Reference last = load(LoadU32(at + 4), LoadU16(at + 10));
  const bool cols = first.row == 0U && last.row == Sheet::kMaxRows - 1U && first.row_abs && last.row_abs;
  const bool rows = !cols && first.col == 0U && last.col == Sheet::kMaxCols - 1U && first.col_abs && last.col_abs;
  first.is_full_col = last.is_full_col = cols;
  first.is_full_row = last.is_full_row = rows;
  const auto out = ctx.transform.apply_range(first, last);
  if (!out) {
    to_error();
    return true;
  }
  store(at, at + 8, out->first);
  store(at + 4, at + 10, out->second);
  return true;
}

/// Rewrites every reference in `rgce[0, cce)` in place. False when a token
/// cannot be stepped over or resolved; the caller then keeps the record.
bool RemapRgce(std::uint8_t* rgce, std::size_t cce, const RefContext& ctx) {
  for (std::size_t i = 0; i < cce;) {
    const std::optional<std::size_t> size = TokenSize(rgce + i, cce - i);
    if (!size || *size > cce - i - 1U) {
      return false;
    }
    const std::uint8_t kind = rgce[i] & 0x1FU;
    const bool ref = rgce[i] >= 0x20U && (kind == 0x04U || kind == 0x05U || kind == 0x0CU || kind == 0x0DU ||
                                          kind == 0x1AU || kind == 0x1BU);
    if (ref && !RemapRefToken(rgce + i, ctx)) {
      return false;
    }
    i += 1U + *size;
  }
  return true;
}

/// The base cell of `ranges` once `transform` has moved them, or nothing
/// when it removes them all.
std::optional<PtgBaseCell> MovedBase(const std::vector<MergeRange>& ranges, const parser::RefTransform& transform) {
  std::optional<PtgBaseCell> base;
  for (const MergeRange& r : ranges) {
    parser::Reference first;
    first.row = r.first_row;
    first.col = r.first_col;
    parser::Reference last;
    last.row = r.last_row;
    last.col = r.last_col;
    const auto out = transform.apply_range(first, last);
    if (!out) {
      continue;
    }
    base = PtgBaseCell{base ? std::min(base->row, out->first.row) : out->first.row,
                       base ? std::min(base->col, out->first.col) : out->first.col};
  }
  return base;
}

void RemapSlotFormulas(std::vector<std::uint8_t>& buf, const std::vector<XlsbExternSheetEntry>& sheets,
                       const parser::RefTransform& transform) {
  Edit edit = Open(buf);
  bool changed = false;
  RefContext ctx{sheets, transform, std::nullopt, std::nullopt};
  for (std::size_t i = 0; i < edit.recs.size(); ++i) {
    const FramedRecord& rec = edit.recs[i];
    if (rec.type == kBeginCondFmt14 || rec.type == kEndCondFmt14) {
      // Formulas inside a block are anchored to its range's top-left cell.
      std::vector<MergeRange> ranges;
      std::size_t rest = 0;
      const bool anchored = rec.type == kBeginCondFmt14 && ParseRanges(rec, ranges, rest);
      const MergeRange base = sqref_base(ranges);
      ctx.old_base = anchored ? std::optional<PtgBaseCell>(PtgBaseCell{base.first_row, base.first_col}) : std::nullopt;
      ctx.new_base = anchored ? MovedBase(ranges, transform) : std::nullopt;
      continue;
    }
    if (rec.type != kSparkline && rec.type != kBeginCfRule14 && rec.type != kCfvo14) {
      continue;
    }
    std::vector<std::pair<std::size_t, std::size_t>> spans;
    if (!FormulaSpans(rec.payload, spans) || spans.empty()) {
      continue;
    }
    RefContext record_ctx = ctx;
    if (rec.type == kSparkline) {
      record_ctx.old_base = record_ctx.new_base = std::nullopt;
    }
    std::vector<std::uint8_t> out(rec.payload.data, rec.payload.data + rec.payload.size);
    bool ok = true;
    for (const auto& [offset, cce] : spans) {
      ok = ok && RemapRgce(out.data() + offset, cce, record_ctx);
    }
    if (!ok || std::equal(out.begin(), out.end(), rec.payload.data)) {
      continue;
    }
    edit.payload[i] = std::move(out);
    changed = true;
  }
  if (changed) {
    edit.Write(buf);
  }
}

}  // namespace

void remap_tail_sqrefs(XlsbSheetTail& tail, const SqrefRemap& remap) {
  for (std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    RemapSlotSqrefs(*slot, remap);
  }
}

void remap_tail_formulas(XlsbSheetTail& tail, const parser::RefTransform& transform) {
  for (std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    RemapSlotFormulas(*slot, tail.extern_sheets, transform);
  }
}

}  // namespace formulon::io::xlsb
