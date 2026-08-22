//
// Phonetic (furigana) annotations attached to a Text cell.
//
// OOXML records a cell's ruby text as one or more `<rPh sb="S" eb="E">`
// blocks hanging off the string item, each carrying the kana for the
// half-open span `[sb, eb)` of the surface text. The offsets are UTF-16
// code units, matching how Excel indexes string positions everywhere
// else.
//
// The spans are load-bearing rather than presentation detail: Excel's
// `PHONETIC` substitutes each annotated span with its kana and passes
// the text outside every span through unchanged, so a partially
// annotated string surfaces a mix of kana and original characters. That
// makes the run boundaries observable through the engine, and they are
// preserved end to end -- reader, cell storage and writer -- instead of
// being flattened into a single kana string.

#ifndef FORMULON_PHONETIC_H_
#define FORMULON_PHONETIC_H_

#include <cstdint>
#include <string>
#include <vector>

namespace formulon {

/// One `<rPh>` block: the kana covering the half-open surface-text span
/// `[sb, eb)`, measured in UTF-16 code units.
///
/// A whole-string annotation is the single run `{0, utf16_units_in(text),
/// kana}`, which is what Excel emits for a cell whose reading was typed
/// in one go and what the authoring API (`Sheet::set_cell_phonetic`)
/// produces.
struct PhoneticRun {
  std::uint32_t sb = 0;
  std::uint32_t eb = 0;
  std::string text;
};

/// The `<phoneticPr>` block that sits beside a string item's `<rPh>` runs:
/// which font renders the ruby, which kana form Excel generates for it, and
/// how it is distributed over the surface text.
///
/// The ordinals are the ones XLSB packs into `BrtSSTItem`'s phonetic
/// trailer (`flags = 0x30 | type | (alignment << 2)`), so the XML and
/// binary paths share this struct without a mapping table.
///
/// Every field's `0` is the value Excel normalises an absent `<phoneticPr>`
/// to, which is why a default-constructed instance needs no presence flag:
/// writing the element out with all-zero fields reproduces what Excel would
/// have inferred anyway.
struct PhoneticProperties {
  /// Index into the workbook's font table for the ruby text.
  std::uint16_t font_id = 0;
  /// 0=halfwidthKatakana, 1=fullwidthKatakana, 2=Hiragana, 3=noConversion.
  std::uint8_t type = 0;
  /// 0=noControl, 1=left, 2=center, 3=distributed.
  std::uint8_t alignment = 0;
};

/// Returns every run's kana concatenated in run order.
///
/// This is the cell's reading with no regard for which characters each
/// run covers -- the shape a furigana field wants, and what the C ABI's
/// `fm_workbook_get_cell_phonetic` returns. `PHONETIC`'s result is a
/// different projection and is built by `eval::compose_phonetic`.
inline std::string flatten_phonetic(const std::vector<PhoneticRun>& runs) {
  std::string out;
  std::size_t total = 0;
  for (const PhoneticRun& run : runs) {
    total += run.text.size();
  }
  out.reserve(total);
  for (const PhoneticRun& run : runs) {
    out.append(run.text);
  }
  return out;
}

}  // namespace formulon

#endif  // FORMULON_PHONETIC_H_
