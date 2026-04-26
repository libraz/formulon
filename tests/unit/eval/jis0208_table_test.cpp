// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the JIS X 0208 reverse table accessors used by Mac Excel
// ja-JP CODE / CHAR. The corresponding oracle suite is
// tests/oracle/cases/code_char_jp_probes.yaml.

#include "eval/jis0208_table.h"

#include <cstdint>

#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace {

TEST(Jis0208Table, UnicodeToJisHiragana) {
  // U+3042 (あ) maps to row 4, cell 2 -> ((4+0x20)<<8) | (2+0x20) = 0x2422.
  EXPECT_EQ(lookup_unicode_to_jis0208(0x3042u), 0x2422u);
}

TEST(Jis0208Table, UnicodeToJisCjkLevel1) {
  // U+4E00 (一) is at row 16, cell 76 -> ((16+0x20)<<8) | (76+0x20) = 0x306C.
  EXPECT_EQ(lookup_unicode_to_jis0208(0x4E00u), 0x306Cu);
}

TEST(Jis0208Table, UnicodeToJisNecExtensionUnmapped) {
  // U+9AD9 (髙) is an NEC extension, not in JIS X 0208 -> sentinel 0.
  EXPECT_EQ(lookup_unicode_to_jis0208(0x9AD9u), 0u);
}

TEST(Jis0208Table, UnicodeToJisEmojiUnmapped) {
  // Supplementary plane (U+1F600) cannot live in JIS X 0208 -> sentinel 0.
  EXPECT_EQ(lookup_unicode_to_jis0208(0x1F600u), 0u);
}

TEST(Jis0208Table, JisToUnicodeHiragana) {
  // Inverse of UnicodeToJisHiragana.
  EXPECT_EQ(lookup_jis0208_to_unicode(0x24, 0x22), 0x3042u);
}

TEST(Jis0208Table, JisToUnicodeIdeographicSpace) {
  // (0x21, 0x21) -> row 1, cell 1 -> U+3000 (ideographic space).
  EXPECT_EQ(lookup_jis0208_to_unicode(0x21, 0x21), 0x3000u);
}

TEST(Jis0208Table, JisToUnicodeOutOfRange) {
  // Either byte outside [0x21, 0x7E] yields the sentinel.
  EXPECT_EQ(lookup_jis0208_to_unicode(0x20, 0x22), 0u);
  EXPECT_EQ(lookup_jis0208_to_unicode(0x7F, 0x22), 0u);
  EXPECT_EQ(lookup_jis0208_to_unicode(0x24, 0x20), 0u);
  EXPECT_EQ(lookup_jis0208_to_unicode(0x24, 0x7F), 0u);
}

TEST(Jis0208Table, RoundTripSampleCodepoints) {
  // For a variety of mapped slots, JIS -> Unicode -> JIS returns the
  // original encoding. Slots picked from the Mac probe corpus.
  struct Sample {
    std::uint8_t hi;
    std::uint8_t lo;
  };
  constexpr Sample kSamples[] = {
      {0x21, 0x21},  // U+3000
      {0x24, 0x22},  // U+3042 (あ)
      {0x30, 0x6C},  // U+4E00 (一)
      {0x23, 0x30},  // U+FF10 (full-width 0)
      {0x23, 0x41},  // U+FF21 (full-width A)
  };
  for (const auto& s : kSamples) {
    const std::uint16_t cp = lookup_jis0208_to_unicode(s.hi, s.lo);
    ASSERT_NE(cp, 0u) << "slot " << static_cast<int>(s.hi) << ',' << static_cast<int>(s.lo)
                      << " is unexpectedly unmapped";
    const std::uint16_t encoded = lookup_unicode_to_jis0208(cp);
    EXPECT_EQ(encoded, static_cast<std::uint16_t>((s.hi << 8) | s.lo))
        << "round-trip mismatch for U+" << std::hex << cp;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
