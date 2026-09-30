// @size-budget: 56 KB

#include "io/xlsb/dv_records.h"

#include <algorithm>
#include <array>
#include <string>
#include <utility>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/structured_log.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kDVal = 64;
constexpr std::uint16_t kAcBegin = 37;
constexpr std::uint16_t kAcEnd = 38;
constexpr std::uint16_t kUid = 3072;
constexpr std::array<std::uint8_t, 6> kAcBeginPayload = {0x01, 0x00, 0x00, 0x10, 0x00, 0x80};
constexpr std::size_t kUidBytes = 16;
constexpr std::size_t kHeaderZeroBytes = 14;

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

bool IsInlineList(const DataValidation& dv) {
  return dv.type == kTypeList && !dv.formula1.empty() && dv.formula1[0] == '"';
}

PtgRootClass RootClass(std::uint8_t type) {
  return type == kTypeList ? PtgRootClass::kReference : PtgRootClass::kValue;
}

std::optional<DataValidation> DecodeDVal(ByteSpan p, const FeatureFormulaReadContext& ctx) {
  auto flags_or = read_u32(p);
  if (!flags_or) {
    return std::nullopt;
  }
  const std::uint32_t flags = flags_or.value();
  DataValidation dv;
  dv.type = static_cast<std::uint8_t>(flags & 0x0FU);
  dv.error_style = static_cast<std::uint8_t>((flags >> 4U) & 0x07U);
  dv.op = static_cast<std::uint8_t>((flags >> 20U) & 0x0FU);
  if (dv.type > kMaxType || dv.error_style > kMaxErrorStyle || dv.op > kMaxOperator || (flags & kImeMode) != 0U ||
      (flags & kHighBits) != 0U) {
    return std::nullopt;
  }
  dv.allow_blank = (flags & kAllowBlank) != 0U;
  dv.show_dropdown = (flags & kSuppressCombo) == 0U;
  dv.show_input_message = (flags & kShowInput) != 0U;
  dv.show_error_message = (flags & kShowError) != 0U;
  if (!read_sqref(p, dv.ranges)) {
    return std::nullopt;
  }
  std::string* strings[4] = {&dv.error_title, &dv.error_message, &dv.prompt_title, &dv.prompt_message};
  for (std::string* s : strings) {
    auto text = read_xlnullablewidestring(p);
    if (!text) {
      return std::nullopt;
    }
    *s = std::move(text).value();
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
  if (p.size != 0U || ((flags & kStrLookup) != 0U) != IsInlineList(dv)) {
    return std::nullopt;
  }
  return dv;
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
    if (r.type != kDVal) {
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
                                    const FeatureFormulaWriteContext& ctx) {
  // A validation covering no cells has no Sqrfx to carry; like an .xlsx
  // entry without `sqref`, it is left out.
  std::uint32_t count = 0;
  for (const DataValidation& dv : validations) {
    count += dv.ranges.empty() ? 0U : 1U;
  }
  if (count == 0U) {
    return Expected<void, Error>::Ok();
  }
  std::vector<std::uint8_t> body;
  std::vector<std::uint8_t> head(kHeaderZeroBytes, 0U);
  emit_u32(head, count);
  emit_record(body, kBrtBeginDVals, head);
  for (const DataValidation& dv : validations) {
    if (dv.ranges.empty()) {
      continue;
    }
    if (dv.type > kMaxType || dv.error_style > kMaxErrorStyle || dv.op > kMaxOperator) {
      return make_error(FormulonErrorCode::kInvalidArgument, "xlsb data validation has no record form",
                        "context=write_xlsb");
    }
    std::uint32_t flags =
        dv.type | (static_cast<std::uint32_t>(dv.error_style) << 4U) | (static_cast<std::uint32_t>(dv.op) << 20U);
    flags |= (IsInlineList(dv) ? kStrLookup : 0U) | (dv.allow_blank ? kAllowBlank : 0U) |
             (dv.show_dropdown ? 0U : kSuppressCombo) | (dv.show_input_message ? kShowInput : 0U) |
             (dv.show_error_message ? kShowError : 0U);
    std::vector<std::uint8_t> p;
    emit_u32(p, flags);
    emit_sqref(p, dv.ranges);
    for (const std::string* s : {&dv.error_title, &dv.error_message, &dv.prompt_title, &dv.prompt_message}) {
      emit_xlnullablewidestring(p, s->empty() ? std::nullopt : std::optional<std::string_view>(*s));
    }
    const MergeRange base = sqref_base(dv.ranges);
    for (const std::string* f : {&dv.formula1, &dv.formula2}) {
      auto encoded = encode_feature_formula(*f, base.first_row, base.first_col, RootClass(dv.type), ctx);
      if (!encoded) {
        return encoded.error();
      }
      p.insert(p.end(), encoded.value().bytes.begin(), encoded.value().bytes.end());
    }
    emit_record(body, kDVal, p);
  }
  emit_record(body, kBrtEndDVals, ByteSpan{});
  dst.insert(dst.end(), body.begin(), body.end());
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
