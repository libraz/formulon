//
// End-to-end tests for the width-conversion text built-ins: ASC (full-width
// to half-width), JIS (half-width to full-width), and DBCS (alias of JIS).
// Each test parses a formula source, evaluates the AST through the default
// registry, and asserts the resulting Value. Voiced and semi-voiced
// katakana are exercised on both sides to pin the decompose / recompose
// paths (ガ <-> ｶﾞ, パ <-> ﾊﾟ, ヴ <-> ｳﾞ). Hiragana / kanji passthrough
// behaviour is pinned so any future regression surfaces here rather than
// in the oracle diff.

#include <string>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSourceIn;

Value EvalWithProfile(std::string_view src, ExcelProfile profile) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(profile);
  return EvalSourceIn(src, wb, wb.sheet(0));
}

Value EvalJapanese(std::string_view src) {
  return EvalWithProfile(src, mac_365_ja_jp_profile());
}

// ---------------------------------------------------------------------------
// Registry pin: all three names are registered.
// ---------------------------------------------------------------------------

TEST(BuiltinsWidthRegistry, AllNamesRegistered) {
  const FunctionRegistry& r = default_registry();
  for (const char* name : {"ASC", "JIS", "DBCS"}) {
    EXPECT_NE(r.lookup(name), nullptr) << "missing registration: " << name;
  }
}

// ---------------------------------------------------------------------------
// ASC
// ---------------------------------------------------------------------------

TEST(BuiltinsWidthAsc, FullwidthAscii) {
  const Value v = EvalJapanese(u8"=ASC(\"Ａ\")");  // "Ａ"
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "A");
}

TEST(BuiltinsWidthAsc, FullwidthAsciiAndSpace) {
  const Value v = EvalJapanese(u8"=ASC(\"ＡＢＣ　１２３\")");  // "ＡＢＣ　１２３"
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ABC 123");
}

TEST(BuiltinsWidthAsc, VoicedKatakanaGa) {
  const Value v = EvalJapanese(u8"=ASC(\"ガ\")");  // "ガ"
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ｶﾞ");  // "ｶﾞ"
}

TEST(BuiltinsWidthAsc, SemiVoicedKatakanaRow) {
  // ASC("パピプペポ") -> "ﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ"
  const Value v = EvalJapanese(u8"=ASC(\"パピプペポ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ");
}

TEST(BuiltinsWidthAsc, VoicedVu) {
  // U+30F4 ヴ has no half-width counterpart in JIS X 0201, so Mac Excel
  // 365 (ja-JP) leaves it unchanged rather than decomposing to ｳ + ﾞ.
  const Value v = EvalJapanese(u8"=ASC(\"ヴ\")");  // "ヴ"
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ヴ");  // "ヴ" (passthrough)
}

TEST(BuiltinsWidthAsc, HiraganaPassthrough) {
  const Value v = EvalJapanese(u8"=ASC(\"ひらがな\")");  // "ひらがな"
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ひらがな");
}

TEST(BuiltinsWidthAsc, MixedKanjiAndFullwidthAscii) {
  // "混在ＡＢ" -> "混在AB"
  const Value v = EvalJapanese(u8"=ASC(\"混在ＡＢ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"混在AB");
}

TEST(BuiltinsWidthAsc, EmptyString) {
  const Value v = EvalJapanese("=ASC(\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsWidthAsc, AlreadyHalfwidth) {
  // Plain ASCII passes through unchanged.
  const Value v = EvalJapanese("=ASC(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello");
}

TEST(BuiltinsWidthAsc, NumberCoerces) {
  // Number coerces to text first, then ASC is identity on ASCII.
  const Value v = EvalJapanese("=ASC(123)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "123");
}

TEST(BuiltinsWidthAsc, ArchaicKatakanaPassthrough) {
  // U+30F7 ヷ has no half-width equivalent in the standard table - pass
  // through verbatim.
  const Value v = EvalJapanese(u8"=ASC(\"ヷ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ヷ");
}

// ---------------------------------------------------------------------------
// JIS
// ---------------------------------------------------------------------------

TEST(BuiltinsWidthJis, AsciiAndSpace) {
  const Value v = EvalJapanese("=JIS(\"ABC 123\")");
  ASSERT_TRUE(v.is_text());
  // "ＡＢＣ　１２３"
  EXPECT_EQ(v.as_text(), u8"ＡＢＣ　１２３");
}

TEST(BuiltinsWidthJis, VoicedRecompose) {
  // JIS("ｶﾞ") -> "ガ"
  const Value v = EvalJapanese(u8"=JIS(\"ｶﾞ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ガ");
}

TEST(BuiltinsWidthJis, SemiVoicedRow) {
  // JIS("ﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ") -> "パピプペポ"
  const Value v = EvalJapanese(u8"=JIS(\"ﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"パピプペポ");
}

TEST(BuiltinsWidthJis, UnvoicedKatakanaRow) {
  // JIS("ｱｲｳｴｵ") -> "アイウエオ"
  const Value v = EvalJapanese(u8"=JIS(\"ｱｲｳｴｵ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"アイウエオ");
}

TEST(BuiltinsWidthJis, BlankStringPassesThrough) {
  const Value v = EvalJapanese("=JIS(\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsWidthJis, HalfwidthPunctuation) {
  // JIS("｡｢｣､･") -> "。「」、・"
  const Value v = EvalJapanese(u8"=JIS(\"｡｢｣､･\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"。「」、・");
}

TEST(BuiltinsWidthJis, LoneDakutenMapsToSpacingMark) {
  // A U+FF9E not preceded by a voice-accepting katakana maps to U+309B.
  const Value v = EvalJapanese(u8"=JIS(\"ﾞ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"゛");
}

TEST(BuiltinsWidthJis, DakutenAfterNonVoicedBaseSplits) {
  // ｧ (small a, no voiced form) + U+FF9E -> small a + spacing dakuten.
  const Value v = EvalJapanese(u8"=JIS(\"ｧﾞ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ァ゛");
}

TEST(BuiltinsWidthJis, VuRecompose) {
  // JIS("ｳﾞ") -> "ヴ"
  const Value v = EvalJapanese(u8"=JIS(\"ｳﾞ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ヴ");
}

// ---------------------------------------------------------------------------
// DBCS is an alias of JIS
// ---------------------------------------------------------------------------

TEST(BuiltinsWidthDbcs, AliasOfJisAscii) {
  const Value v = EvalJapanese("=DBCS(\"ABC 123\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ＡＢＣ　１２３");
}

TEST(BuiltinsWidthDbcs, AliasOfJisVoiced) {
  // DBCS("ｶﾞ") -> "ガ"
  const Value v = EvalJapanese(u8"=DBCS(\"ｶﾞ\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ガ");
}

TEST(BuiltinsWidthDbcs, BothSpellingsAgreeAcrossTheConversionCorpus) {
  // `JIS` is the ja-JP formula-bar spelling of `DBCS`; Excel rewrites it
  // to `DBCS` when the formula is committed, so no oracle case can check
  // that the two agree — by the time Excel evaluates, only one name is
  // left. The equivalence is an internal invariant of the shared
  // registration in `register_text_width_builtins`, and this is where it
  // is pinned. The argument list mirrors the width oracle corpus so the
  // alias is exercised over exactly the inputs Excel answers for.
  const char* const kArguments[] = {
      "\"ABC 123\"",
      u8"\"ｶﾞ\"",
      u8"\"ﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ\"",
      u8"\"ｱｲｳｴｵ\"",
      u8"\"ＡＢＣ\"",
      u8"\"漢字\"",
      u8"\"漢ｱ字\"",
      u8"\"ﾞ\"",
      "\"!@#$%\"",
      "TRUE",
      "#DIV/0!",
      u8"\"ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ\"",
      "ASC(\"ＡＢＣ１２３\")",
  };
  for (const char* argument : kArguments) {
    SCOPED_TRACE(argument);
    const Value from_jis = EvalJapanese(std::string("=JIS(") + argument + ")");
    const Value from_dbcs = EvalJapanese(std::string("=DBCS(") + argument + ")");
    ASSERT_EQ(from_jis.kind(), from_dbcs.kind());
    if (from_jis.is_text()) {
      EXPECT_EQ(from_jis.as_text(), from_dbcs.as_text());
    } else if (from_jis.is_error()) {
      EXPECT_EQ(from_jis.as_error(), from_dbcs.as_error());
    }
  }

  // The nesting order must not matter either, which is what the corpus's
  // two round-trip cases assert against Excel under the stored spelling.
  EXPECT_EQ(EvalJapanese("=ASC(JIS(\"ABC123\"))").as_text(), EvalJapanese("=ASC(DBCS(\"ABC123\"))").as_text());
}

// ---------------------------------------------------------------------------
// Round-trips
// ---------------------------------------------------------------------------

TEST(BuiltinsWidthRoundTrip, AscJisAsciiIdentity) {
  const Value v = EvalJapanese("=ASC(JIS(\"ABC\"))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ABC");
}

TEST(BuiltinsWidthRoundTrip, JisAscVoicedIdentity) {
  // JIS(ASC("ガ")) -> "ガ"
  const Value v = EvalJapanese(u8"=JIS(ASC(\"ガ\"))");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), u8"ガ");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
