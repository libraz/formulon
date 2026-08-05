//
// Unit tests for the shared half-to-full katakana base lookup
// (`eval/jp_kana_table.h`). Boundary checks ensure the accessors return
// zero outside their respective Unicode ranges and a representative
// sample of mappings to confirm table integrity.

#include "eval/jp_kana_table.h"

#include <cstdint>

#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace {

TEST(JpKanaTable, HalfToFullBaseRangeBoundary) {
  // One below the supported range.
  EXPECT_EQ(half_to_full_katakana_base(0xFF65u), 0u);
  // One above the supported range.
  EXPECT_EQ(half_to_full_katakana_base(0xFF9Eu), 0u);
  // Far outside.
  EXPECT_EQ(half_to_full_katakana_base(0u), 0u);
  EXPECT_EQ(half_to_full_katakana_base(0x30A2u), 0u);  // already full-width.
}

TEST(JpKanaTable, HalfToFullBaseSpotChecks) {
  EXPECT_EQ(half_to_full_katakana_base(0xFF66u), 0x30F2u);  // ｦ -> ヲ
  EXPECT_EQ(half_to_full_katakana_base(0xFF67u), 0x30A1u);  // ｧ -> ァ
  EXPECT_EQ(half_to_full_katakana_base(0xFF70u), 0x30FCu);  // ｰ -> ー
  EXPECT_EQ(half_to_full_katakana_base(0xFF76u), 0x30ABu);  // ｶ -> カ
  EXPECT_EQ(half_to_full_katakana_base(0xFF8Au), 0x30CFu);  // ﾊ -> ハ
  EXPECT_EQ(half_to_full_katakana_base(0xFF9Du), 0x30F3u);  // ﾝ -> ン
}

TEST(JpKanaTable, HalfToFullPunctuationRangeBoundary) {
  // One below.
  EXPECT_EQ(half_to_full_punctuation(0xFF60u), 0u);
  // One above.
  EXPECT_EQ(half_to_full_punctuation(0xFF66u), 0u);
  // Spot check inside.
  EXPECT_EQ(half_to_full_punctuation(0xFF61u), 0x3002u);  // ｡ -> 。
  EXPECT_EQ(half_to_full_punctuation(0xFF62u), 0x300Cu);  // ｢ -> 「
  EXPECT_EQ(half_to_full_punctuation(0xFF63u), 0x300Du);  // ｣ -> 」
  EXPECT_EQ(half_to_full_punctuation(0xFF64u), 0x3001u);  // ､ -> 、
  EXPECT_EQ(half_to_full_punctuation(0xFF65u), 0x30FBu);  // ･ -> ・
}

}  // namespace
}  // namespace eval
}  // namespace formulon
