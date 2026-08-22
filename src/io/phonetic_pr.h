//
// `<phoneticPr>` attribute vocabulary, shared by every path that reads or
// writes a furigana guide.
//
// The element hangs off a string item (`<si>` in the shared-string table,
// `<is>` for an inline string) beside that item's `<rPh>` runs and states
// how Excel renders and generates the ruby: which font, which kana form,
// and how the kana is distributed over the characters it covers.
//
// It is kept here rather than in `phonetic.h` because the model header is
// deliberately free of XML: `PhoneticProperties` holds ordinals, and this
// is the one place that knows their OOXML spelling. The binary container
// needs no counterpart -- XLSB packs the same ordinals into `BrtSSTItem`'s
// trailer numerically.
//
// An absent element and a present-but-bare one resolve differently, which
// is the trap in this part of the format. Both were measured by handing
// Excel a package that varies in exactly one attribute and reading the
// ordinals back out of the `.xlsb` it converts to:
//
//   * No `<phoneticPr>` at all -> `fontId="0"`, `halfwidthKatakana`,
//     `noControl`: the all-zero triple, so a default-constructed
//     `PhoneticProperties` is right and a reader can leave it alone.
//   * `<phoneticPr fontId="0"/>` -> `fullwidthKatakana`, `left`: each
//     attribute falls back to its own schema default, which is a
//     different state from the element being missing.
//
// The writer sidesteps the distinction by spelling all three attributes
// out beside any non-empty run list, which is what Excel does too.

#ifndef FORMULON_IO_PHONETIC_PR_H_
#define FORMULON_IO_PHONETIC_PR_H_

#include <cstdint>
#include <string>
#include <string_view>

#include "phonetic.h"

namespace formulon {
namespace io {

/// Maps a `type` attribute to `PhoneticProperties::type`.
///
/// A missing or unrecognised attribute is `fullwidthKatakana`, which is
/// *not* the same as the all-zero state a missing element resolves to.
/// Call this only for an element that is actually present.
inline std::uint8_t parse_phonetic_type(std::string_view value) {
  if (value == "halfwidthKatakana") {
    return 0U;
  }
  if (value == "Hiragana") {
    return 2U;
  }
  if (value == "noConversion") {
    return 3U;
  }
  return 1U;  // fullwidthKatakana, the attribute's schema default
}

/// Maps an `alignment` attribute to `PhoneticProperties::alignment`.
/// A missing or unrecognised attribute is `left`, per the same rule.
inline std::uint8_t parse_phonetic_alignment(std::string_view value) {
  if (value == "noControl") {
    return 0U;
  }
  if (value == "center") {
    return 2U;
  }
  if (value == "distributed") {
    return 3U;
  }
  return 1U;  // left, the attribute's schema default
}

inline const char* phonetic_type_name(std::uint8_t type) {
  switch (type) {
    case 1U:
      return "fullwidthKatakana";
    case 2U:
      return "Hiragana";
    case 3U:
      return "noConversion";
    default:
      return "halfwidthKatakana";
  }
}

inline const char* phonetic_alignment_name(std::uint8_t alignment) {
  switch (alignment) {
    case 1U:
      return "left";
    case 2U:
      return "center";
    case 3U:
      return "distributed";
    default:
      return "noControl";
  }
}

/// Appends `<phoneticPr fontId="F" type="T" alignment="A"/>`.
///
/// All three attributes are always written, matching Excel: it spells the
/// defaults out rather than relying on them being inferred, so a saved
/// package compares byte for byte against the one it was read from.
inline void append_phonetic_pr(std::string& out, const PhoneticProperties& props) {
  out.append("<phoneticPr fontId=\"");
  out.append(std::to_string(props.font_id));
  out.append("\" type=\"");
  out.append(phonetic_type_name(props.type));
  out.append("\" alignment=\"");
  out.append(phonetic_alignment_name(props.alignment));
  out.append("\"/>");
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_PHONETIC_PR_H_
