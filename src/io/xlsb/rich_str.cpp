// @size-budget: 1 KB
//
// Implementation of the `RichStr` codec. See `io/xlsb/rich_str.h`.

#include "io/xlsb/rich_str.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/record_writer.h"
#include "utils/status_macros.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// One `StrRun` ([MS-XLSB] §2.5.94) is `(u16 ich, u16 ifnt)`.
constexpr std::size_t kStrRunSize = 4U;

/// UTF-16 offsets into one phonetic string, indexed once before its runs are
/// sliced. At a supplementary-plane midpoint, the floor and ceil entries
/// intentionally differ just as they do in `utf16_substring`.
class Utf16SliceIndex {
 public:
  explicit Utf16SliceIndex(std::string_view text) : text_(text) {
    floor_bytes_.push_back(0U);
    ceil_bytes_.push_back(0U);
    std::size_t byte = 0U;
    while (byte < text_.size()) {
      std::size_t byte_length = 0U;
      const std::uint32_t codepoint = decode_utf8_step(text_, byte, &byte_length);
      if (codepoint > 0xFFFFU) {
        floor_bytes_.push_back(byte);
        ceil_bytes_.push_back(byte + byte_length);
      }
      byte += byte_length;
      floor_bytes_.push_back(byte);
      ceil_bytes_.push_back(byte);
    }
  }

  std::uint32_t utf16_units() const noexcept { return static_cast<std::uint32_t>(floor_bytes_.size() - 1U); }

  std::string substring(std::uint32_t start_units, std::uint32_t length_units) const {
    if (length_units == 0U) {
      return {};
    }
    std::uint32_t end_units = start_units + length_units;
    if (end_units < start_units) {
      end_units = UINT32_MAX;
    }
    const std::size_t last_unit = floor_bytes_.size() - 1U;
    const std::size_t start_byte = floor_bytes_[std::min<std::size_t>(start_units, last_unit)];
    const std::size_t end_byte = ceil_bytes_[std::min<std::size_t>(end_units, last_unit)];
    if (end_byte <= start_byte) {
      return {};
    }
    return std::string(text_.substr(start_byte, end_byte - start_byte));
  }

 private:
  std::string_view text_;
  std::vector<std::size_t> floor_bytes_;
  std::vector<std::size_t> ceil_bytes_;
};

/// Decodes the phonetic tail a `RichStr` carries when `fExtStr` is set,
/// appending one `PhoneticRun` per `PhRun` to `out`.
///
/// The binary shape stores the kana once, concatenated across every run,
/// and each `PhRun` names where its slice starts rather than carrying the
/// slice: `(u16 ichFirst, u16 ichMom, u16 cchMom)` is the start of this
/// run's kana inside the concatenation, the surface-text offset it reads,
/// and how many surface characters it covers. A run's kana therefore ends
/// where the next run's begins, and the last run's runs to the end. All
/// offsets are UTF-16 code units, which is what `PhoneticRun` uses too.
///
/// Excel elides the runs entirely for a whole-string reading (the
/// concatenation is then simply the whole annotation), so an empty run
/// array with a non-empty phonetic string is the single-run case rather
/// than an absent one -- the same shape `<rPh sb="0" eb="len">` takes in
/// OOXML.
///
/// The trailing `(u16 ifnt, u16 flags)` -- phonetic font, plus the
/// annotation type and alignment `<phoneticPr>` carries in OOXML -- is
/// decoded into `out_props` via `DecodePhoneticProperties` below. The
/// OOXML reader models the same element too, parsing `<phoneticPr>` in
/// its DOM, SAX and SST paths (`cell_parser.cpp`, `sax_xml_reader.cpp`,
/// `sst_reader.cpp`).

/// Reads the phonetic tail's closing `(u16 ifnt, u16 flags)` pair.
///
/// Leniently: a record that stops short of the pair keeps the defaults
/// rather than failing the load. The runs have already been decoded by
/// then, and a guide missing only its rendering hints is still the
/// reading the user typed.
void DecodePhoneticProperties(ByteSpan& cursor, PhoneticProperties& out) {
  auto font_or = read_u16(cursor);
  if (!font_or) {
    return;
  }
  auto flags_or = read_u16(cursor);
  if (!flags_or) {
    return;
  }
  out.font_id = font_or.value();
  out.type = static_cast<std::uint8_t>(flags_or.value() & 0x03U);
  out.alignment = static_cast<std::uint8_t>((flags_or.value() >> 2U) & 0x03U);
}

Expected<void, Error> DecodePhoneticTail(ByteSpan& cursor, std::string_view surface, std::vector<PhoneticRun>& out,
                                         PhoneticProperties& out_props) {
  auto phonetic_or = read_xlwidestring(cursor);
  if (!phonetic_or) {
    return std::move(phonetic_or.error());
  }
  const std::string phonetic = std::move(phonetic_or.value());
  auto count_or = read_u32(cursor);
  if (!count_or) {
    return std::move(count_or.error());
  }
  const std::uint32_t run_count = count_or.value();
  if (run_count == 0U) {
    if (!phonetic.empty()) {
      out.push_back(PhoneticRun{0U, utf16_units_in(surface), phonetic});
    }
    DecodePhoneticProperties(cursor, out_props);
    return {};
  }

  // Each PhRun is three `u16`s; bound the count against the remaining
  // payload before reserving so a corrupt count cannot drive a large
  // allocation.
  constexpr std::size_t kPhRunSize = 6U;
  if (static_cast<std::size_t>(run_count) > cursor.size / kPhRunSize) {
    std::string ctx("context=xlsb.rich_str run_count=");
    ctx.append(std::to_string(run_count));
    ctx.append(" cursor_size=").append(std::to_string(cursor.size));
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb phonetic run array truncated", std::move(ctx));
  }
  std::vector<std::array<std::uint16_t, 3>> runs;
  runs.reserve(run_count);
  for (std::uint32_t i = 0; i < run_count; ++i) {
    std::array<std::uint16_t, 3> fields{};
    for (std::uint16_t& field : fields) {
      auto field_or = read_u16(cursor);
      if (!field_or) {
        return std::move(field_or.error());
      }
      field = field_or.value();
    }
    runs.push_back(fields);
  }

  const Utf16SliceIndex phonetic_index(phonetic);
  const std::uint32_t phonetic_units = phonetic_index.utf16_units();
  for (std::size_t i = 0; i < runs.size(); ++i) {
    const std::uint32_t kana_start = runs[i][0];
    const std::uint32_t surface_start = runs[i][1];
    const std::uint32_t surface_length = runs[i][2];
    const std::uint32_t kana_end = (i + 1U < runs.size()) ? runs[i + 1U][0] : phonetic_units;
    // A backwards or out-of-range slice yields an empty reading rather
    // than an error: the surrounding cell is still usable, and the run
    // boundaries are Excel's own bookkeeping rather than user data.
    const std::uint32_t kana_length = kana_end > kana_start ? kana_end - kana_start : 0U;
    out.push_back(
        PhoneticRun{surface_start, surface_start + surface_length, phonetic_index.substring(kana_start, kana_length)});
  }
  DecodePhoneticProperties(cursor, out_props);
  return {};
}
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
    kana_offset += utf16_units_in(run.text);
  }
  emit_u16(payload, props.font_id);
  emit_u16(payload, pack_phonetic_flags(props));
}

}  // namespace

Expected<RichStr, Error> read_rich_str(ByteSpan& cursor) {
  RichStr out;
  ASSIGN_OR_RETURN(const std::uint8_t flags, read_u8(cursor));
  ASSIGN_OR_RETURN(out.text, read_xlwidestring(cursor));
  if ((flags & kRichStrRichRuns) != 0U) {
    ASSIGN_OR_RETURN(const std::uint32_t run_count, read_u32(cursor));
    if (static_cast<std::size_t>(run_count) > cursor.size / kStrRunSize) {
      std::string ctx("context=xlsb.rich_str rich_runs=");
      ctx.append(std::to_string(run_count));
      ctx.append(" cursor_size=").append(std::to_string(cursor.size));
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb rich-text run array truncated",
                        std::move(ctx));
    }
    out.runs.resize(run_count);
    for (StrRun& run : out.runs) {
      ASSIGN_OR_RETURN(run.ich, read_u16(cursor));
      ASSIGN_OR_RETURN(run.ifnt, read_u16(cursor));
    }
  }
  if ((flags & kRichStrPhonetic) != 0U) {
    RETURN_IF_ERROR(DecodePhoneticTail(cursor, out.text, out.phonetic, out.phonetic_props));
  }
  return out;
}

void emit_rich_str(std::vector<std::uint8_t>& out, std::string_view text, const std::vector<StrRun>& runs,
                   const std::vector<PhoneticRun>& phonetic, PhoneticProperties phonetic_props) {
  const std::uint8_t flags =
      static_cast<std::uint8_t>((runs.empty() ? 0U : kRichStrRichRuns) | (phonetic.empty() ? 0U : kRichStrPhonetic));
  emit_u8(out, flags);
  emit_xlwidestring(out, text);
  if (!runs.empty()) {
    emit_u32(out, static_cast<std::uint32_t>(runs.size()));
    for (const StrRun& run : runs) {
      emit_u16(out, run.ich);
      emit_u16(out, run.ifnt);
    }
  }
  if (!phonetic.empty()) {
    emit_phonetic_tail(out, phonetic, phonetic_props);
  }
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
