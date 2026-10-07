#include "io/xlsb/sst_reader.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/record.h"
#include "phonetic.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// `RichStr` flag bits ([MS-XLSB] §2.5.87): rich-text runs follow the
/// string when the first is set, a phonetic guide when the second is.
constexpr std::uint8_t kRichStrRichRuns = 0x01U;
constexpr std::uint8_t kRichStrPhonetic = 0x02U;

/// One `StrRun` ([MS-XLSB] §2.5.94) is `(u16 ich, u16 ifnt)`.
constexpr std::size_t kStrRunSize = 4U;

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
    std::string ctx("context=xlsb.sst run_count=");
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

  const std::uint32_t phonetic_units = utf16_units_in(phonetic);
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
        PhoneticRun{surface_start, surface_start + surface_length, utf16_substring(phonetic, kana_start, kana_length)});
  }
  DecodePhoneticProperties(cursor, out_props);
  return {};
}
}  // namespace

Expected<std::vector<std::string_view>, Error> DecodeSharedStringsBin(
    const std::vector<std::uint8_t>& body, std::deque<std::string>& text_storage,
    std::vector<std::vector<PhoneticRun>>& out_phonetic, std::vector<PhoneticProperties>& out_phonetic_props) {
  std::vector<std::string_view> entries;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return std::move(rec_or.error());
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem)) {
      continue;
    }
    // BrtSSTItem ([MS-XLSB] §2.4.293) is a `RichStr`: a flags byte, the
    // string, then the optional rich-format runs and phonetic guide the
    // flags announce.
    ByteSpan p = rec.payload;
    ASSIGN_OR_RETURN(const std::uint8_t flags, read_u8(p));
    ASSIGN_OR_RETURN(auto str, read_xlwidestring(p));
    text_storage.push_back(std::move(str));
    entries.push_back(text_storage.back());
    out_phonetic.emplace_back();
    out_phonetic_props.emplace_back();

    if ((flags & kRichStrPhonetic) == 0U) {
      continue;
    }
    // The rich-format runs sit between the string and the phonetic tail,
    // so they have to be stepped over even though the reader models
    // plain text only.
    if ((flags & kRichStrRichRuns) != 0U) {
      ASSIGN_OR_RETURN(auto run_count, read_u32(p));
      const std::uint32_t rich_runs = run_count;
      if (static_cast<std::size_t>(rich_runs) > p.size / kStrRunSize) {
        std::string ctx("context=xlsb.sst rich_runs=");
        ctx.append(std::to_string(rich_runs));
        ctx.append(" cursor_size=").append(std::to_string(p.size));
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb rich-text run array truncated",
                          std::move(ctx));
      }
      p.data += static_cast<std::size_t>(rich_runs) * kStrRunSize;
      p.size -= static_cast<std::size_t>(rich_runs) * kStrRunSize;
    }
    if (auto decoded = DecodePhoneticTail(p, entries.back(), out_phonetic.back(), out_phonetic_props.back());
        !decoded) {
      return std::move(decoded.error());
    }
  }
  return entries;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
