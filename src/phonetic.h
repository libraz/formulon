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
