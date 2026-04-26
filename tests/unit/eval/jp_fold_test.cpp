// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `fold_jp_text`. Each case fixes a single fold rule against
// its exact UTF-8 byte sequence so a future regression in either the
// hiragana / half-width / full-width tables or the voicing-composition
// branch fails loudly.

#include "eval/jp_fold.h"

#include <string>

#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace {

TEST(JpFold, AsciiPassthrough) {
  EXPECT_EQ(fold_jp_text(""), "");
  EXPECT_EQ(fold_jp_text("hello"), "hello");
  EXPECT_EQ(fold_jp_text("Hello, World!"), "Hello, World!");
}

TEST(JpFold, HiraganaToKatakana) {
  // U+3042 あ -> U+30A2 ア.
  EXPECT_EQ(fold_jp_text("\xE3\x81\x82"), "\xE3\x82\xA2");
  // Sokuon: U+3063 っ -> U+30C3 ッ.
  EXPECT_EQ(fold_jp_text("\xE3\x81\xA3"), "\xE3\x83\x83");
  // Word: あっぷる -> アップル.
  EXPECT_EQ(fold_jp_text("\xE3\x81\x82\xE3\x81\xA3\xE3\x81\xB7\xE3\x82\x8B"),
            "\xE3\x82\xA2\xE3\x83\x83\xE3\x83\x97\xE3\x83\xAB");
}

TEST(JpFold, HalfWidthToFullWidthKatakana) {
  // U+FF71 ｱ -> U+30A2 ア.
  EXPECT_EQ(fold_jp_text("\xEF\xBD\xB1"), "\xE3\x82\xA2");
  // U+FF9D ﾝ -> U+30F3 ン.
  EXPECT_EQ(fold_jp_text("\xEF\xBE\x9D"), "\xE3\x83\xB3");
  // U+FF70 ｰ -> U+30FC ー (long mark).
  EXPECT_EQ(fold_jp_text("\xEF\xBD\xB0"), "\xE3\x83\xBC");
}

TEST(JpFold, HalfWidthVoicingCompose) {
  // U+FF76 ｶ + U+FF9E ﾞ -> U+30AC ガ.
  EXPECT_EQ(fold_jp_text("\xEF\xBD\xB6\xEF\xBE\x9E"), "\xE3\x82\xAC");
  // U+FF8A ﾊ + U+FF9E ﾞ -> U+30D0 バ.
  EXPECT_EQ(fold_jp_text("\xEF\xBE\x8A\xEF\xBE\x9E"), "\xE3\x83\x90");
  // U+FF73 ｳ + U+FF9E ﾞ -> U+30F4 ヴ (special-case).
  EXPECT_EQ(fold_jp_text("\xEF\xBD\xB3\xEF\xBE\x9E"), "\xE3\x83\xB4");
}

TEST(JpFold, HalfWidthSemiVoicingCompose) {
  // U+FF8A ﾊ + U+FF9F ﾟ -> U+30D1 パ.
  EXPECT_EQ(fold_jp_text("\xEF\xBE\x8A\xEF\xBE\x9F"), "\xE3\x83\x91");
}

TEST(JpFold, StandaloneVoicingMarks) {
  // 'a' + U+FF9E ﾞ (no composable base) -> 'a' + U+309B ゛.
  EXPECT_EQ(fold_jp_text("a\xEF\xBE\x9E"), "a\xE3\x82\x9B");
  // U+FF9F ﾟ alone -> U+309C ゜.
  EXPECT_EQ(fold_jp_text("\xEF\xBE\x9F"), "\xE3\x82\x9C");
  // ｱ (cannot voice) + ﾞ -> ア (full-width) + U+309B ゛.
  EXPECT_EQ(fold_jp_text("\xEF\xBD\xB1\xEF\xBE\x9E"), "\xE3\x82\xA2\xE3\x82\x9B");
}

TEST(JpFold, FullWidthAsciiToHalfWidth) {
  // U+FF21..U+FF23 ＡＢＣ -> ABC.
  EXPECT_EQ(fold_jp_text("\xEF\xBC\xA1\xEF\xBC\xA2\xEF\xBC\xA3"), "ABC");
}

TEST(JpFold, FullWidthDigits) {
  // U+FF11..U+FF13 １２３ -> 123.
  EXPECT_EQ(fold_jp_text("\xEF\xBC\x91\xEF\xBC\x92\xEF\xBC\x93"), "123");
}

TEST(JpFold, FullWidthPunctuation) {
  // U+FF0A ＊ -> '*'.
  EXPECT_EQ(fold_jp_text("\xEF\xBC\x8A"), "*");
  // U+FF1F ？ -> '?'.
  EXPECT_EQ(fold_jp_text("\xEF\xBC\x9F"), "?");
  // U+FF20 ＠ -> '@'.
  EXPECT_EQ(fold_jp_text("\xEF\xBC\xA0"), "@");
}

TEST(JpFold, IdeographicSpace) {
  // 'a' + U+3000 + 'b' -> "a b".
  EXPECT_EQ(fold_jp_text("a\xE3\x80\x80"
                         "b"),
            "a b");
}

TEST(JpFold, FullWidthAsciiBoundary) {
  // U+FF01 ！ (first) -> '!'.
  EXPECT_EQ(fold_jp_text("\xEF\xBC\x81"), "!");
  // U+FF5E ～ (last) -> '~'.
  EXPECT_EQ(fold_jp_text("\xEF\xBD\x9E"), "~");
}

TEST(JpFold, MixedFolding) {
  // あ ＊ ｶﾞ -> ア * ガ.
  EXPECT_EQ(fold_jp_text("\xE3\x81\x82\xEF\xBC\x8A\xEF\xBD\xB6\xEF\xBE\x9E"), "\xE3\x82\xA2*\xE3\x82\xAC");
}

TEST(JpFold, NonJpCodepointsPassThrough) {
  // U+4E2D 中 should not change.
  EXPECT_EQ(fold_jp_text("\xE4\xB8\xAD"), "\xE4\xB8\xAD");
  // Emoji U+1F600 😀 should not change.
  EXPECT_EQ(fold_jp_text("\xF0\x9F\x98\x80"), "\xF0\x9F\x98\x80");
}

// `fold_halfwidth_kana = false` covers the D-function header path
// (DSUM / DCOUNT / DGET / etc.). Mac Excel deliberately leaves
// half-width katakana — including the FF9E / FF9F voicing-mark
// composition — unfolded when resolving the `field` argument or a
// criteria-block header against the database header row, even though it
// folds them in COUNTIF cell-vs-criterion comparisons.

TEST(JpFold, HalfWidthKanaNoFoldPassThrough) {
  // ﾌﾙｰﾂ (FF8C FF99 FF70 FF82) must pass through byte-for-byte under the
  // D-function header rule; the corresponding default-mode result would
  // fold to フルーツ (U+30D5 U+30EB U+30FC U+30C4).
  EXPECT_EQ(fold_jp_text("\xEF\xBE\x8C\xEF\xBE\x99\xEF\xBD\xB0\xEF\xBE\x82",
                         /*fold_fullwidth_digits=*/true, /*fold_halfwidth_kana=*/false),
            "\xEF\xBE\x8C\xEF\xBE\x99\xEF\xBD\xB0\xEF\xBE\x82");
}

TEST(JpFold, HalfWidthVoicingNoCompose) {
  // ｶﾞ (FF76 FF9E) must NOT compose to ガ (U+30AC); both codepoints pass
  // through unchanged so the D-function header comparison sees the raw
  // two-codepoint sequence.
  EXPECT_EQ(fold_jp_text("\xEF\xBD\xB6\xEF\xBE\x9E",
                         /*fold_fullwidth_digits=*/true, /*fold_halfwidth_kana=*/false),
            "\xEF\xBD\xB6\xEF\xBE\x9E");
  // Standalone FF9F also passes through unchanged (default mode would
  // map to U+309C ゜).
  EXPECT_EQ(fold_jp_text("\xEF\xBE\x9F",
                         /*fold_fullwidth_digits=*/true, /*fold_halfwidth_kana=*/false),
            "\xEF\xBE\x9F");
}

TEST(JpFold, HiraAndFullWidthLatinStillFoldUnderHalfKanaDisabled) {
  // ふるーつ (U+3075 U+308B U+30FC U+3064) -> フルーツ (U+30D5 U+30EB
  // U+30FC U+30C4); ＦＲＵＩＴ (U+FF26 U+FF32 U+FF35 U+FF29 U+FF34) ->
  // FRUIT. Neither rule depends on `fold_halfwidth_kana`, so they must
  // still apply under the D-function header asymmetric mode.
  EXPECT_EQ(fold_jp_text("\xE3\x81\xB5\xE3\x82\x8B\xE3\x83\xBC\xE3\x81\xA4\xEF\xBC\xA6\xEF\xBC\xB2\xEF\xBC\xB5\xEF\xBC"
                         "\xA9\xEF\xBC\xB4",
                         /*fold_fullwidth_digits=*/true, /*fold_halfwidth_kana=*/false),
            "\xE3\x83\x95\xE3\x83\xAB\xE3\x83\xBC\xE3\x83\x84"
            "FRUIT");
  // Full-width digit text ('１', U+FF11) still folds to ASCII '1' under
  // the asymmetric mode; this is the rule the D-function field-arg path
  // relies on.
  EXPECT_EQ(fold_jp_text("\xEF\xBC\x91", /*fold_fullwidth_digits=*/true, /*fold_halfwidth_kana=*/false), "1");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
