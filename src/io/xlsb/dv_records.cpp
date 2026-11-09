// @size-budget: 56 KB

#include "io/xlsb/dv_records.h"

#include <algorithm>
#include <array>
#include <string>
#include <utility>

#include "io/dv_overlay.h"
#include "io/xlsb/cf_x14_link.h"
#include "io/xlsb/frt_formula.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "utils/status_macros.h"
#include "utils/structured_log.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::array<std::uint8_t, 6> kAcBeginPayload = {0x01, 0x00, 0x00, 0x10, 0x00, 0x80};
constexpr std::size_t kUidBytes = 16;
constexpr std::size_t kHeaderZeroBytes = 14;
constexpr std::size_t kFrtHeaderBytes = 4;

constexpr std::uint8_t kTypeList = 3;
constexpr std::uint8_t kMaxType = 7;
constexpr std::uint8_t kMaxErrorStyle = 2;
constexpr std::uint8_t kMaxOperator = 7;
constexpr std::uint32_t kStrLookup = 1U << 7U;
constexpr std::uint32_t kAllowBlank = 1U << 8U;
constexpr std::uint32_t kSuppressCombo = 1U << 9U;
constexpr std::uint32_t kImeMode = 0xFFU << 10U;
constexpr std::uint32_t kShowInput = 1U << 18U;
constexpr std::uint32_t kShowError = 1U << 19U;
constexpr std::uint32_t kHighBits = 0xFFU << 24U;
// BrtDVal14 flag bits naming the FRTFormulas present.
constexpr std::uint32_t kFormula1Present = 1U << 24U;
constexpr std::uint32_t kFormula2Present = 1U << 25U;

bool IsInlineList(const DataValidation& dv) {
  return dv.type == kTypeList && !dv.formula1.empty() && dv.formula1[0] == '"';
}

PtgRootClass RootClass(std::uint8_t type) {
  return type == kTypeList ? PtgRootClass::kReference : PtgRootClass::kValue;
}

/// Reads BrtDVal's flag word into `dv`; false for a value outside the model.
bool DecodeFlags(std::uint32_t flags, DataValidation& dv) {
  dv.type = static_cast<std::uint8_t>(flags & 0x0FU);
  dv.error_style = static_cast<std::uint8_t>((flags >> 4U) & 0x07U);
  dv.op = static_cast<std::uint8_t>((flags >> 20U) & 0x0FU);
  if (dv.type > kMaxType || dv.error_style > kMaxErrorStyle || dv.op > kMaxOperator || (flags & kImeMode) != 0U ||
      (flags & kHighBits) != 0U) {
    return false;
  }
  dv.allow_blank = (flags & kAllowBlank) != 0U;
  dv.show_dropdown = (flags & kSuppressCombo) == 0U;
  dv.show_input_message = (flags & kShowInput) != 0U;
  dv.show_error_message = (flags & kShowError) != 0U;
  return true;
}

std::uint32_t EncodeFlags(const DataValidation& dv) {
  std::uint32_t flags =
      dv.type | (static_cast<std::uint32_t>(dv.error_style) << 4U) | (static_cast<std::uint32_t>(dv.op) << 20U);
  flags |= (IsInlineList(dv) ? kStrLookup : 0U) | (dv.allow_blank ? kAllowBlank : 0U) |
           (dv.show_dropdown ? 0U : kSuppressCombo) | (dv.show_input_message ? kShowInput : 0U) |
           (dv.show_error_message ? kShowError : 0U);
  return flags;
}

bool ReadStrings(ByteSpan& p, DataValidation& dv) {
  for (std::string* s : {&dv.error_title, &dv.error_message, &dv.prompt_title, &dv.prompt_message}) {
    auto text = read_xlnullablewidestring(p);
    if (!text) {
      return false;
    }
    *s = std::move(text).value();
  }
  return true;
}

void EmitStrings(std::vector<std::uint8_t>& p, const DataValidation& dv) {
  for (const std::string* s : {&dv.error_title, &dv.error_message, &dv.prompt_title, &dv.prompt_message}) {
    emit_xlnullablewidestring(p, s->empty() ? std::nullopt : std::optional<std::string_view>(*s));
  }
}

Expected<void, Error> CheckWritable(const DataValidation& dv) {
  if (dv.type > kMaxType || dv.error_style > kMaxErrorStyle || dv.op > kMaxOperator) {
    return make_error(FormulonErrorCode::kInvalidArgument, "xlsb data validation has no record form",
                      kWriteXlsbContext);
  }
  return Expected<void, Error>::Ok();
}

std::optional<DataValidation> DecodeDVal(ByteSpan p, const FeatureFormulaReadContext& ctx) {
  auto flags = read_u32(p);
  DataValidation dv;
  if (!flags || !DecodeFlags(flags.value(), dv) || !read_sqref(p, dv.ranges) || !ReadStrings(p, dv)) {
    return std::nullopt;
  }
  const MergeRange base = sqref_base(dv.ranges);
  for (std::string* f : {&dv.formula1, &dv.formula2}) {
    auto text = read_feature_formula(p, base.first_row, base.first_col, ctx);
    if (!text) {
      StructuredLog("xlsb.dv.formula_not_decoded").field("reason", text.error().message).warn();
      return std::nullopt;
    }
    *f = std::move(text).value();
  }
  if (p.size != 0U || ((flags.value() & kStrLookup) != 0U) != IsInlineList(dv)) {
    return std::nullopt;
  }
  return dv;
}

/// BrtDVal14 as laid out in `dv_records.h`, or nothing for any other shape.
std::optional<DataValidation> DecodeDVal14(ByteSpan p, const FeatureFormulaReadContext& ctx) {
  auto header = read_u32(p);
  auto sqrefs = read_u32(p);
  auto sqref_flags = read_u32(p);
  if (!header || !sqrefs || !sqref_flags ||
      (header.value() != kFrtHeaderSqrefs && header.value() != (kFrtHeaderSqrefs | kFrtHeaderFormulas)) ||
      sqrefs.value() != 1U || sqref_flags.value() != kFrtItemFlags) {
    return std::nullopt;
  }
  DataValidation dv;
  if (!read_sqref(p, dv.ranges) || dv.ranges.empty()) {
    return std::nullopt;
  }
  const MergeRange base = sqref_base(dv.ranges);
  std::vector<std::string> formulas;
  if ((header.value() & kFrtHeaderFormulas) != 0U) {
    auto count = read_u32(p);
    if (!count || count.value() == 0U || count.value() > 2U) {
      return std::nullopt;
    }
    for (std::uint32_t i = 0; i < count.value(); ++i) {
      auto text = read_frt_formula(p, base.first_row, base.first_col, ctx);
      if (!text) {
        StructuredLog("xlsb.dv14.formula_not_decoded").field("reason", text.error().message).warn();
        return std::nullopt;
      }
      if (text.value().empty()) {
        return std::nullopt;
      }
      formulas.push_back(std::move(text).value());
    }
  }
  auto flags = read_u32(p);
  if (!flags) {
    return std::nullopt;
  }
  const bool has1 = (flags.value() & kFormula1Present) != 0U;
  const bool has2 = (flags.value() & kFormula2Present) != 0U;
  if (formulas.size() != static_cast<std::size_t>(has1) + static_cast<std::size_t>(has2) ||
      !DecodeFlags(flags.value() & ~(kFormula1Present | kFormula2Present), dv) || !ReadStrings(p, dv) || p.size != 0U) {
    return std::nullopt;
  }
  std::size_t next = 0;
  if (has1) {
    dv.formula1 = std::move(formulas[next++]);
  }
  if (has2) {
    dv.formula2 = std::move(formulas[next++]);
  }
  if (((flags.value() & kStrLookup) != 0U) != IsInlineList(dv)) {
    return std::nullopt;
  }
  return dv;
}

Expected<void, Error> EmitDVal14(std::vector<std::uint8_t>& dst, const DataValidation& dv,
                                 const FeatureFormulaWriteContext& ctx) {
  RETURN_IF_ERROR(CheckWritable(dv));
  const MergeRange base = sqref_base(dv.ranges);
  std::vector<std::uint8_t> list;
  std::uint32_t count = 0;
  std::uint32_t present = 0;
  const std::string* sources[2] = {&dv.formula1, &dv.formula2};
  const std::uint32_t bits[2] = {kFormula1Present, kFormula2Present};
  for (std::size_t i = 0; i < 2U; ++i) {
    if (sources[i]->empty()) {
      continue;
    }
    ASSIGN_OR_RETURN(EncodedFeatureFormula f,
                     encode_feature_formula(*sources[i], base.first_row, base.first_col, RootClass(dv.type), ctx));
    emit_frt_formula(list, f);
    present |= bits[i];
    ++count;
  }
  std::vector<std::uint8_t> p;
  emit_u32(p, kFrtHeaderSqrefs | (count != 0U ? kFrtHeaderFormulas : 0U));
  emit_u32(p, 1U);
  emit_u32(p, kFrtItemFlags);
  emit_sqref(p, dv.ranges);
  if (count != 0U) {
    emit_u32(p, count);
    p.insert(p.end(), list.begin(), list.end());
  }
  emit_u32(p, EncodeFlags(dv) | present);
  EmitStrings(p, dv);
  emit_record(dst, kDVal14, p);
  return Expected<void, Error>::Ok();
}

/// True when `recs[i..i+2]` is a BrtDVal's revision-uid wrapper.
bool IsUidWrapper(const std::vector<FramedRecord>& recs, std::size_t i) {
  if (i + 2U >= recs.size()) {
    return false;
  }
  const ByteSpan begin = recs[i].payload;
  return recs[i].type == kAcBegin && begin.size == kAcBeginPayload.size() &&
         std::equal(kAcBeginPayload.begin(), kAcBeginPayload.end(), begin.data) && recs[i + 1U].type == kUid &&
         recs[i + 1U].payload.size == kUidBytes && recs[i + 2U].type == kAcEnd && recs[i + 2U].payload.size == 0U;
}

/// One BrtDVal14 of a container: `first` is its uid wrapper, or `record`
/// itself when it has none.
struct Dv14Item {
  std::size_t first = 0;
  std::size_t record = 0;
};

/// The DVals14 container opening at `recs[open]` (BrtFRTBegin) and closing at
/// `recs[close]` (BrtFRTEnd).
struct Dv14Container {
  std::size_t open = 0;
  std::size_t close = 0;
  std::vector<Dv14Item> items;
};

/// The container at `recs[i]`, or nothing when it is not one of the shape
/// in `dv_records.h` with a matching record count.
std::optional<Dv14Container> ParseContainer(const std::vector<FramedRecord>& recs, std::size_t i) {
  const std::size_t n = recs.size();
  if (recs[i].type != kFrtBegin || i + 1U >= n || recs[i + 1U].type != kBeginDVals14) {
    return std::nullopt;
  }
  Dv14Container c;
  c.open = i;
  std::size_t k = i + 2U;
  while (k < n && recs[k].type != kEndDVals14) {
    Dv14Item item{k, IsUidWrapper(recs, k) ? k + 3U : k};
    if (item.record >= n || recs[item.record].type != kDVal14) {
      return std::nullopt;
    }
    c.items.push_back(item);
    k = item.record + 1U;
  }
  if (k + 1U >= n || recs[k].payload.size != 0U || recs[k + 1U].type != kFrtEnd) {
    return std::nullopt;
  }
  const ByteSpan head = recs[i + 1U].payload;
  constexpr std::size_t kZeroBytes = kFrtHeaderBytes + kHeaderZeroBytes;
  if (head.size != kZeroBytes + 4U ||
      std::any_of(head.data, head.data + kZeroBytes, [](std::uint8_t b) { return b != 0U; }) ||
      load_u32(head.data + kZeroBytes) != c.items.size()) {
    return std::nullopt;
  }
  c.close = k + 1U;
  return c;
}

void EmitDVals14Head(std::vector<std::uint8_t>& dst, std::uint32_t count) {
  std::vector<std::uint8_t> head(kFrtHeaderBytes + kHeaderZeroBytes, 0U);
  emit_u32(head, count);
  emit_record(dst, kBeginDVals14, head);
}

/// Where a new DVals14 container goes: after the x14 CF container, ahead of
/// any other worksheet future record and the sheet's trailing revision-uid
/// block.
std::size_t NewContainerAt(const std::vector<FramedRecord>& recs, std::size_t size) {
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (recs[i].type != kFrtBegin) {
      continue;
    }
    if (i + 1U < recs.size() && recs[i + 1U].type == kBeginCfs14) {
      std::size_t k = i + 2U;
      while (k < recs.size() && recs[k].type != kEndCfs14) {
        ++k;
      }
      if (k + 1U < recs.size() && recs[k + 1U].type == kFrtEnd) {
        i = k + 1U;
        continue;
      }
    }
    return recs[i].begin;
  }
  if (recs.size() >= 3U && recs[recs.size() - 3U].type == kAcBegin && recs.back().type == kAcEnd) {
    return recs[recs.size() - 3U].begin;
  }
  return size;
}

}  // namespace

std::optional<std::vector<DataValidation>> decode_dv_block(ByteSpan block, const FeatureFormulaReadContext& ctx) {
  ByteSpan cursor = block;
  auto head = read_record(cursor);
  if (!head || head.value().type != kBrtBeginDVals || head.value().payload.size != kHeaderZeroBytes + 4U) {
    return std::nullopt;
  }
  ByteSpan hp = head.value().payload;
  for (std::size_t i = 0; i < kHeaderZeroBytes; ++i) {
    if (hp.data[i] != 0U) {
      return std::nullopt;
    }
  }
  hp.data += kHeaderZeroBytes;
  hp.size -= kHeaderZeroBytes;
  const std::uint32_t count = read_u32(hp).value();
  std::vector<DataValidation> out;
  while (cursor.size != 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      return std::nullopt;
    }
    const XlsbRecord& r = rec.value();
    if (r.type == kBrtEndDVals) {
      if (r.payload.size != 0U || cursor.size != 0U || out.size() != count) {
        return std::nullopt;
      }
      return out;
    }
    if (r.type == kAcBegin) {
      auto uid = read_record(cursor);
      if (!uid) {
        return std::nullopt;
      }
      auto end = read_record(cursor);
      if (r.payload.size != kAcBeginPayload.size() ||
          !std::equal(kAcBeginPayload.begin(), kAcBeginPayload.end(), r.payload.data) || uid.value().type != kUid ||
          uid.value().payload.size != kUidBytes || !end || end.value().type != kAcEnd ||
          end.value().payload.size != 0U) {
        return std::nullopt;
      }
      continue;
    }
    if (r.type != kBrtDVal) {
      return std::nullopt;
    }
    auto dv = DecodeDVal(r.payload, ctx);
    if (!dv) {
      return std::nullopt;
    }
    out.push_back(std::move(*dv));
  }
  return std::nullopt;
}

Expected<void, Error> emit_dv_block(std::vector<std::uint8_t>& dst, const std::vector<DataValidation>& validations,
                                    const FeatureFormulaWriteContext& ctx, std::string_view own_sheet) {
  // A validation covering no cells has no Sqrfx to carry; like an .xlsx
  // entry without `sqref`, it is left out.
  std::vector<const DataValidation*> listed;
  for (const DataValidation& dv : validations) {
    if (!dv.ranges.empty() && !data_validation_in_x14(dv, own_sheet, &ctx.sheet_names)) {
      listed.push_back(&dv);
    }
  }
  if (listed.empty()) {
    return Expected<void, Error>::Ok();
  }
  std::vector<std::uint8_t> body;
  std::vector<std::uint8_t> head(kHeaderZeroBytes, 0U);
  emit_u32(head, static_cast<std::uint32_t>(listed.size()));
  emit_record(body, kBrtBeginDVals, head);
  for (const DataValidation* dv : listed) {
    RETURN_IF_ERROR(CheckWritable(*dv));
    std::vector<std::uint8_t> p;
    emit_u32(p, EncodeFlags(*dv));
    emit_sqref(p, dv->ranges);
    EmitStrings(p, *dv);
    const MergeRange base = sqref_base(dv->ranges);
    for (const std::string* f : {&dv->formula1, &dv->formula2}) {
      auto encoded = encode_feature_formula(*f, base.first_row, base.first_col, RootClass(dv->type), ctx);
      if (!encoded) {
        return std::move(encoded.error());
      }
      p.insert(p.end(), encoded.value().bytes.begin(), encoded.value().bytes.end());
    }
    emit_record(body, kBrtDVal, p);
  }
  emit_record(body, kBrtEndDVals, ByteSpan{});
  dst.insert(dst.end(), body.begin(), body.end());
  return Expected<void, Error>::Ok();
}

void take_x14_validations(const std::vector<std::uint8_t>& records, const FeatureFormulaReadContext& ctx,
                          std::vector<DataValidation>& validations) {
  const std::vector<FramedRecord> recs = split_records(records);
  for (std::size_t i = 0; i < recs.size();) {
    const std::optional<Dv14Container> c = ParseContainer(recs, i);
    if (!c) {
      ++i;
      continue;
    }
    for (const Dv14Item& item : c->items) {
      if (std::optional<DataValidation> dv = DecodeDVal14(recs[item.record].payload, ctx)) {
        validations.push_back(std::move(*dv));
      } else {
        StructuredLog("xlsb.dv14.retained_raw").warn();
      }
    }
    i = c->close + 1U;
  }
}

Expected<void, Error> place_x14_validations(XlsbSheetTail& tail, const std::vector<DataValidation>& validations,
                                            const FeatureFormulaWriteContext& ctx, std::string_view own_sheet) {
  std::vector<const DataValidation*> pending;
  for (const DataValidation& dv : validations) {
    if (!dv.ranges.empty() && data_validation_in_x14(dv, own_sheet, &ctx.sheet_names)) {
      pending.push_back(&dv);
    }
  }
  const FeatureFormulaLoadView load_view(ctx, tail.extern_sheets, own_sheet);
  const FeatureFormulaReadContext& read_ctx = load_view.context();
  std::size_t next = 0;
  bool filled = false;
  for (std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    const std::vector<FramedRecord> recs = split_records(*slot);
    std::vector<std::uint8_t> out;
    bool changed = false;
    for (std::size_t i = 0; i < recs.size();) {
      const std::optional<Dv14Container> c = ParseContainer(recs, i);
      if (!c) {
        append_record(out, *slot, recs[i]);
        ++i;
        continue;
      }
      changed = true;
      std::vector<std::uint8_t> items;
      std::uint32_t count = 0;
      for (const Dv14Item& item : c->items) {
        // An adopted record is a slot for the model's next validation.
        if (DecodeDVal14(recs[item.record].payload, read_ctx).has_value()) {
          if (!filled && next < pending.size()) {
            RETURN_IF_ERROR(EmitDVal14(items, *pending[next++], ctx));
            ++count;
          }
          continue;
        }
        for (std::size_t m = item.first; m <= item.record; ++m) {
          append_record(items, *slot, recs[m]);
        }
        ++count;
      }
      if (!filled) {
        for (; next < pending.size(); ++next) {
          RETURN_IF_ERROR(EmitDVal14(items, *pending[next], ctx));
          ++count;
        }
        filled = true;
      }
      if (count != 0U) {
        append_record(out, *slot, recs[c->open]);
        EmitDVals14Head(out, count);
        out.insert(out.end(), items.begin(), items.end());
        append_record(out, *slot, recs[c->close - 1U]);
        append_record(out, *slot, recs[c->close]);
      }
      i = c->close + 1U;
    }
    if (changed) {
      *slot = std::move(out);
    }
  }
  if (next == pending.size()) {
    return Expected<void, Error>::Ok();
  }
  std::vector<std::uint8_t> items;
  for (std::size_t i = next; i < pending.size(); ++i) {
    RETURN_IF_ERROR(EmitDVal14(items, *pending[i], ctx));
  }
  std::vector<std::uint8_t>& records = tail.after_hyperlinks;
  const std::vector<FramedRecord> recs = split_records(records);
  std::vector<std::uint8_t> block = frt_container_block(recs, kBeginDVals14, kEndDVals14, items);
  // The container helper opens with an empty record; BrtBeginDVals14 carries
  // the header and count.
  for (const FramedRecord& r : split_records(block)) {
    if (r.type == kBeginDVals14) {
      std::vector<std::uint8_t> head;
      EmitDVals14Head(head, static_cast<std::uint32_t>(pending.size() - next));
      block.erase(block.begin() + static_cast<std::ptrdiff_t>(r.begin),
                  block.begin() + static_cast<std::ptrdiff_t>(r.end));
      block.insert(block.begin() + static_cast<std::ptrdiff_t>(r.begin), head.begin(), head.end());
      break;
    }
  }
  const std::size_t at = NewContainerAt(recs, records.size());
  records.insert(records.begin() + static_cast<std::ptrdiff_t>(at), block.begin(), block.end());
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
