//
// Implementation of the XLSB shared-string-table builder + emitter.

#include "io/xlsb/sst_writer.h"

#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/utf8_length.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "phonetic.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// `RichStr` flag bit announcing the phonetic tail ([MS-XLSB] §2.5.87).
constexpr std::uint8_t kRichStrPhonetic = 0x02U;

/// The constant high bits of the phonetic tail's closing flags word.
///
/// The word packs the annotation type in bits 0-1 and its alignment in
/// bits 2-3; bits 4-5 are set on every guide Excel writes. An all-default
/// `PhoneticProperties` therefore encodes as `0x0030`, which is
/// `halfwidthKatakana` / `noControl` — the pair Excel infers for a guide
/// that arrived without a `<phoneticPr>`.
constexpr std::uint16_t kPhoneticFlagsBase = 0x0030U;

std::uint16_t pack_phonetic_flags(PhoneticProperties props) {
  return static_cast<std::uint16_t>(kPhoneticFlagsBase | (props.type & 0x03U) |
                                    static_cast<std::uint16_t>((props.alignment & 0x03U) << 2U));
}

/// Builds the interner key for one payload.
///
/// Length-prefixes every field for the same reason the OOXML shared
/// strings writer does: without it, adjacent fields could be re-cut into
/// the same byte sequence and collide two distinct annotations onto one
/// entry.
std::string key_for(std::string_view text, const std::vector<PhoneticRun>& phonetic,
                    PhoneticProperties phonetic_props) {
  std::string key;
  key.reserve(text.size() + 32U + phonetic.size() * 16U);
  key.append(std::to_string(text.size()));
  key.push_back(':');
  key.append(text);
  for (const PhoneticRun& run : phonetic) {
    key.push_back(';');
    key.append(std::to_string(run.sb));
    key.push_back(',');
    key.append(std::to_string(run.eb));
    key.push_back(',');
    key.append(std::to_string(run.text.size()));
    key.push_back(':');
    key.append(run.text);
  }
  if (!phonetic.empty()) {
    key.push_back('|');
    key.append(std::to_string(phonetic_props.font_id));
    key.push_back(',');
    key.append(std::to_string(phonetic_props.type));
    key.push_back(',');
    key.append(std::to_string(phonetic_props.alignment));
  }
  return key;
}

/// Narrows a UTF-16 offset to the `u16` the record format stores.
///
/// Excel's own string ceiling is 32767 characters, so a well-formed
/// annotation always fits; clamping rather than truncating keeps a
/// hostile in-memory workbook from wrapping an offset around into a
/// different, valid-looking span.
std::uint16_t narrow_offset(std::uint32_t units) {
  constexpr std::uint32_t kMax = std::numeric_limits<std::uint16_t>::max();
  return static_cast<std::uint16_t>(units < kMax ? units : kMax);
}

/// Appends the phonetic tail for `runs` to `payload`.
void emit_phonetic_tail(std::vector<std::uint8_t>& payload, const std::vector<PhoneticRun>& runs,
                        PhoneticProperties props) {
  std::string kana;
  for (const PhoneticRun& run : runs) {
    kana.append(run.text);
  }
  emit_xlwidestring(payload, kana);
  emit_u32(payload, static_cast<std::uint32_t>(runs.size()));
  // `ichFirst` is a running offset into the concatenation above, so the
  // slice boundaries are recovered on read from consecutive runs.
  std::uint32_t kana_offset = 0;
  for (const PhoneticRun& run : runs) {
    emit_u16(payload, narrow_offset(kana_offset));
    emit_u16(payload, narrow_offset(run.sb));
    emit_u16(payload, narrow_offset(run.eb > run.sb ? run.eb - run.sb : 0U));
    kana_offset += eval::utf16_units_in(run.text);
  }
  emit_u16(payload, props.font_id);
  emit_u16(payload, pack_phonetic_flags(props));
}

}  // namespace

std::uint32_t SstBuilder::intern(std::string_view text, const std::vector<PhoneticRun>& phonetic,
                                 PhoneticProperties phonetic_props) {
  std::string key = key_for(text, phonetic, phonetic_props);
  if (auto it = index_.find(key); it != index_.end()) {
    return it->second;
  }
  const std::uint32_t idx = static_cast<std::uint32_t>(entries_.size());
  entries_.push_back(SstEntry{std::string(text), phonetic, phonetic_props});
  index_.emplace(std::move(key), idx);
  return idx;
}

Expected<std::vector<std::uint8_t>, Error> emit_sst(const SstBuilder& sst) {
  std::vector<std::uint8_t> body;
  // Reserve a generous lower-bound: 16 bytes for the framing
  // (begin + end records) plus 8 for the begin payload, plus a
  // per-item average of ~8 bytes of framing on top of the string
  // body. Resizing is amortised; this just prevents pathological
  // small reallocations on workbooks with many strings.
  body.reserve(32U + sst.size() * 16U);

  // BrtBeginSst payload: (u32 cstTotal, u32 cstUnique). We don't
  // separately track multiplicity (the interner has already deduped
  // every text cell to a single index), so cstTotal == cstUnique.
  {
    std::vector<std::uint8_t> p;
    emit_u32(p, sst.size());
    emit_u32(p, sst.size());
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSst), p);
  }

  // One BrtSSTItem per entry: (u8 flags, XLWideString[, phonetic tail]).
  for (const SstEntry& entry : sst.entries()) {
    std::vector<std::uint8_t> p;
    const bool has_phonetic = !entry.phonetic.empty();
    emit_u8(p, has_phonetic ? kRichStrPhonetic : 0U);
    emit_xlwidestring(p, entry.text);
    if (has_phonetic) {
      emit_phonetic_tail(p, entry.phonetic, entry.phonetic_props);
    }
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem), p);
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSst), ByteSpan{});
  return body;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
