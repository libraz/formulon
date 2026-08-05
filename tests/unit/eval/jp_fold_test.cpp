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

TEST(JpFold, EveryVoicedKatakanaComposes) {
  const std::pair<const char*, const char*> cases[] = {
      {"\xEF\xBD\xB3\xEF\xBE\x9E", "\xE3\x83\xB4"},  // ｳﾞ -> ヴ
      {"\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC"}, {"\xEF\xBD\xB7\xEF\xBE\x9E", "\xE3\x82\xAE"},
      {"\xEF\xBD\xB8\xEF\xBE\x9E", "\xE3\x82\xB0"}, {"\xEF\xBD\xB9\xEF\xBE\x9E", "\xE3\x82\xB2"},
      {"\xEF\xBD\xBA\xEF\xBE\x9E", "\xE3\x82\xB4"}, {"\xEF\xBD\xBB\xEF\xBE\x9E", "\xE3\x82\xB6"},
      {"\xEF\xBD\xBC\xEF\xBE\x9E", "\xE3\x82\xB8"}, {"\xEF\xBD\xBD\xEF\xBE\x9E", "\xE3\x82\xBA"},
      {"\xEF\xBD\xBE\xEF\xBE\x9E", "\xE3\x82\xBC"}, {"\xEF\xBD\xBF\xEF\xBE\x9E", "\xE3\x82\xBE"},
      {"\xEF\xBE\x80\xEF\xBE\x9E", "\xE3\x83\x80"}, {"\xEF\xBE\x81\xEF\xBE\x9E", "\xE3\x83\x82"},
      {"\xEF\xBE\x82\xEF\xBE\x9E", "\xE3\x83\x85"}, {"\xEF\xBE\x83\xEF\xBE\x9E", "\xE3\x83\x87"},
      {"\xEF\xBE\x84\xEF\xBE\x9E", "\xE3\x83\x89"}, {"\xEF\xBE\x8A\xEF\xBE\x9E", "\xE3\x83\x90"},
      {"\xEF\xBE\x8B\xEF\xBE\x9E", "\xE3\x83\x93"}, {"\xEF\xBE\x8C\xEF\xBE\x9E", "\xE3\x83\x96"},
      {"\xEF\xBE\x8D\xEF\xBE\x9E", "\xE3\x83\x99"}, {"\xEF\xBE\x8E\xEF\xBE\x9E", "\xE3\x83\x9C"},
  };
  for (const auto& [halfwidth, fullwidth] : cases) {
    EXPECT_EQ(fold_jp_text(halfwidth), fullwidth) << halfwidth;
  }
}

TEST(JpFold, HalfWidthSemiVoicingCompose) {
  // U+FF8A ﾊ + U+FF9F ﾟ -> U+30D1 パ.
  EXPECT_EQ(fold_jp_text("\xEF\xBE\x8A\xEF\xBE\x9F"), "\xE3\x83\x91");
}

TEST(JpFold, EverySemiVoicedKatakanaComposes) {
  const std::pair<const char*, const char*> cases[] = {
      {"\xEF\xBE\x8A\xEF\xBE\x9F", "\xE3\x83\x91"}, {"\xEF\xBE\x8B\xEF\xBE\x9F", "\xE3\x83\x94"},
      {"\xEF\xBE\x8C\xEF\xBE\x9F", "\xE3\x83\x97"}, {"\xEF\xBE\x8D\xEF\xBE\x9F", "\xE3\x83\x9A"},
      {"\xEF\xBE\x8E\xEF\xBE\x9F", "\xE3\x83\x9D"},
  };
  for (const auto& [halfwidth, fullwidth] : cases) {
    EXPECT_EQ(fold_jp_text(halfwidth), fullwidth) << halfwidth;
  }
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

// `fold_and_lower` composes the kana fold with an ASCII lowercase pass.
// These tests pin the equality verdict (cell vs. lookup) the lookup
// family relies on so xlookup and classic stay byte-identical.
TEST(FoldAndLower, AsciiLowercase) {
  EXPECT_EQ(fold_and_lower("HELLO"), "hello");
  EXPECT_EQ(fold_and_lower("Hello"), "hello");
  EXPECT_EQ(fold_and_lower("hello"), "hello");
}

TEST(FoldAndLower, FullWidthAsciiFoldsThenLowercases) {
  // Ａ (U+FF21) -> 'A' -> 'a'; Ｂ -> 'b'.
  EXPECT_EQ(fold_and_lower("\xEF\xBC\xA1\xEF\xBC\xA2"), "ab");
}

TEST(FoldAndLower, HalfWidthKatakanaWithVoicingComposes) {
  // ｶﾞ (FF76 FF9E) -> ガ (U+30AC). Lower-pass leaves CJK bytes alone.
  EXPECT_EQ(fold_and_lower("\xEF\xBD\xB6\xEF\xBE\x9E"), "\xE3\x82\xAC");
}

TEST(FoldAndLower, MixedAsciiAndKana) {
  // Ｓｍｉｔｈガ -> "smith" + "ガ".
  EXPECT_EQ(fold_and_lower("\xEF\xBC\xB3\xEF\xBD\x8D\xEF\xBD\x89\xEF\xBD\x94\xEF\xBD\x88\xE3\x82\xAC"),
            "smith\xE3\x82\xAC");
}

TEST(FoldAndLower, FullWidthDigitsNotFoldedByDefault) {
  // Default mode (lookup parity): full-width digit '１' (U+FF11) stays
  // unfolded; ASCII pass leaves it alone.
  EXPECT_EQ(fold_and_lower("\xEF\xBC\x91"), "\xEF\xBC\x91");
  // Opt in to digit folding (criteria parity): '１' -> '1'.
  EXPECT_EQ(fold_and_lower("\xEF\xBC\x91", /*fold_fullwidth_digits=*/true), "1");
}

// ---------------------------------------------------------------------------
// equal_folded / compare_folded
// ---------------------------------------------------------------------------
//
// These pin the high-level helpers introduced for the lookup / groupby /
// database refactor. Each test fixes a single flag combination so a
// future regression in the option plumbing fails loudly.

TEST(EqualFolded, AsciiNoFoldingDefaultCaseInsensitive) {
  // ASCII pass-through under default options (case_insensitive=true).
  EXPECT_TRUE(equal_folded("abc", "abc"));
  EXPECT_TRUE(equal_folded("abc", "ABC"));
  EXPECT_TRUE(equal_folded("Hello", "HELLO"));
  EXPECT_FALSE(equal_folded("abc", "abd"));
  EXPECT_FALSE(equal_folded("abc", "abcd"));
  EXPECT_TRUE(equal_folded("", ""));
}

TEST(EqualFolded, CaseSensitiveOption) {
  FoldCompareOptions cs;
  cs.case_insensitive = false;
  EXPECT_TRUE(equal_folded("abc", "abc", cs));
  EXPECT_FALSE(equal_folded("abc", "ABC", cs));
  // Folded equality still works under case-sensitive mode: Ａ -> 'A'.
  EXPECT_TRUE(equal_folded("\xEF\xBC\xA1", "A", cs));
  EXPECT_FALSE(equal_folded("\xEF\xBC\xA1", "a", cs));
}

TEST(EqualFolded, HalfwidthFullwidthAscii) {
  // ＡＢＣ (U+FF21..U+FF23) folds to "ABC"; case-insensitive default
  // makes both sides equal to "abc".
  EXPECT_TRUE(equal_folded("\xEF\xBC\xA1\xEF\xBC\xA2\xEF\xBC\xA3", "abc"));
  EXPECT_TRUE(equal_folded("\xEF\xBC\xA1\xEF\xBC\xA2\xEF\xBC\xA3", "ABC"));
  // Mixed: ＡBＣ should still fold and compare equal to "abc".
  EXPECT_TRUE(
      equal_folded("\xEF\xBC\xA1"
                   "B\xEF\xBC\xA3",
                   "abc"));
}

TEST(EqualFolded, HiraganaToKatakana) {
  // あ (U+3042) folds to ア (U+30A2); compare against the literal
  // full-width katakana ア.
  EXPECT_TRUE(equal_folded("\xE3\x81\x82", "\xE3\x82\xA2"));
  // Word: あっぷる == アップル.
  EXPECT_TRUE(equal_folded("\xE3\x81\x82\xE3\x81\xA3\xE3\x81\xB7\xE3\x82\x8B",
                           "\xE3\x82\xA2\xE3\x83\x83\xE3\x83\x97\xE3\x83\xAB"));
}

TEST(EqualFolded, LengthDifferenceAfterFolding) {
  // ｶﾞ (FF76 FF9E, 6 bytes) folds to ガ (U+30AC, 3 bytes). Pre-fold
  // byte lengths differ; post-fold the strings are byte-equal.
  EXPECT_TRUE(equal_folded("\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC"));
  // Disable half-width katakana folding (D-function path): the same
  // pair no longer compares equal because the LHS stays as the raw
  // two-codepoint sequence.
  FoldCompareOptions dfn;
  dfn.fold_halfwidth_kana = false;
  EXPECT_FALSE(equal_folded("\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC", dfn));
}

TEST(EqualFolded, FullwidthDigitOptionGate) {
  // '１' (U+FF11) folds to '1' when fold_fullwidth_digits=true (criteria
  // / database parity); stays as the original codepoint and does NOT
  // compare equal to '1' when fold_fullwidth_digits=false (lookup
  // parity).
  FoldCompareOptions criteria;  // defaults
  EXPECT_TRUE(equal_folded("\xEF\xBC\x91", "1", criteria));
  FoldCompareOptions lookup;
  lookup.fold_fullwidth_digits = false;
  EXPECT_FALSE(equal_folded("\xEF\xBC\x91", "1", lookup));
}

TEST(CompareFolded, AsciiOrdering) {
  EXPECT_EQ(compare_folded("abc", "abc"), 0);
  EXPECT_LT(compare_folded("abc", "abd"), 0);
  EXPECT_GT(compare_folded("abd", "abc"), 0);
  // Case-insensitive: "ABC" == "abc".
  EXPECT_EQ(compare_folded("ABC", "abc"), 0);
  // Shorter prefix orders before its longer continuation.
  EXPECT_LT(compare_folded("abc", "abcd"), 0);
}

TEST(CompareFolded, CaseSensitiveOption) {
  FoldCompareOptions cs;
  cs.case_insensitive = false;
  // 'A' (0x41) < 'a' (0x61) under exact-byte compare.
  EXPECT_LT(compare_folded("ABC", "abc", cs), 0);
  // Same byte sequence after the fold compares equal even under
  // case-sensitive mode.
  EXPECT_EQ(compare_folded("\xEF\xBC\xA1", "A", cs), 0);
}

TEST(CompareFolded, FoldedThenCompare) {
  // ｶﾞ (FF76 FF9E) folds to ガ (U+30AC); equal under default options.
  EXPECT_EQ(compare_folded("\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC"), 0);
  // Disable half-width kana folding: now they differ; LHS starts with
  // 0xEF which sorts after 0xE3.
  FoldCompareOptions dfn;
  dfn.fold_halfwidth_kana = false;
  EXPECT_GT(compare_folded("\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC", dfn), 0);
}

// ---------------------------------------------------------------------------
// value_equal_folded_text / value_compare_folded_text
// ---------------------------------------------------------------------------

TEST(ValueEqualFoldedText, TextFoldsHiraganaKatakana) {
  const std::string hira = "\xE3\x81\x82\xE3\x81\xA3";  // あっ
  const std::string kata = "\xE3\x82\xA2\xE3\x83\x83";  // アッ
  EXPECT_TRUE(value_equal_folded_text(Value::text(hira), Value::text(kata)));
}

TEST(ValueEqualFoldedText, NumberCompareDoesNotFold) {
  // Numbers compare by IEEE-754 `==`; no fold pass is invoked.
  EXPECT_TRUE(value_equal_folded_text(Value::number(1.5), Value::number(1.5)));
  EXPECT_FALSE(value_equal_folded_text(Value::number(1.5), Value::number(1.6)));
  // +0.0 == -0.0 under `==` (mirroring group_cell_equal).
  EXPECT_TRUE(value_equal_folded_text(Value::number(0.0), Value::number(-0.0)));
}

TEST(ValueEqualFoldedText, BoolErrorBlankFallbacks) {
  EXPECT_TRUE(value_equal_folded_text(Value::boolean(true), Value::boolean(true)));
  EXPECT_FALSE(value_equal_folded_text(Value::boolean(true), Value::boolean(false)));
  EXPECT_TRUE(value_equal_folded_text(Value::error(ErrorCode::NA), Value::error(ErrorCode::NA)));
  EXPECT_FALSE(value_equal_folded_text(Value::error(ErrorCode::NA), Value::error(ErrorCode::Value)));
  EXPECT_TRUE(value_equal_folded_text(Value::blank(), Value::blank()));
}

TEST(ValueEqualFoldedText, CrossKindNeverEqual) {
  // Text "1" vs Number 1 -> not equal (Excel-canonical: no coercion).
  EXPECT_FALSE(value_equal_folded_text(Value::text("1"), Value::number(1.0)));
  // Number 0 vs Bool FALSE -> distinct kinds.
  EXPECT_FALSE(value_equal_folded_text(Value::number(0.0), Value::boolean(false)));
  // Blank vs Text "" -> distinct kinds (Blank != empty Text in
  // group_cell_equal contract).
  EXPECT_FALSE(value_equal_folded_text(Value::blank(), Value::text("")));
}

TEST(ValueCompareFoldedText, OrdersWithinSameKind) {
  EXPECT_LT(value_compare_folded_text(Value::number(1.0), Value::number(2.0)), 0);
  EXPECT_GT(value_compare_folded_text(Value::number(2.0), Value::number(1.0)), 0);
  EXPECT_EQ(value_compare_folded_text(Value::number(1.0), Value::number(1.0)), 0);

  // Text compare uses fold + case-insensitive default.
  const std::string hira_a = "\xE3\x81\x82";  // あ
  const std::string kata_a = "\xE3\x82\xA2";  // ア
  EXPECT_EQ(value_compare_folded_text(Value::text(hira_a), Value::text(kata_a)), 0);

  // Bool: FALSE < TRUE.
  EXPECT_LT(value_compare_folded_text(Value::boolean(false), Value::boolean(true)), 0);
}

TEST(ValueCompareFoldedText, CrossKindReturnsZero) {
  // The contract: cross-kind compare returns 0 so the caller can
  // impose a bucket-rank tiebreaker. This mirrors the way
  // `groupby_pivotby_lazy.cpp` separates bucket vs intra-kind
  // ordering.
  EXPECT_EQ(value_compare_folded_text(Value::text("a"), Value::number(1.0)), 0);
  EXPECT_EQ(value_compare_folded_text(Value::number(0.0), Value::boolean(false)), 0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
