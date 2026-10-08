//
// End-to-end tests for the byte-oriented text built-ins: LENB, LEFTB,
// RIGHTB, MIDB, CHAR, CODE. Pins the Mac Excel 365 (ja-JP) DBCS byte-count
// rule across ASCII, half-width katakana, BMP CJK, and supplementary-plane
// codepoints, plus the "pad with ASCII space on a 1-byte overflow of a
// 2-byte character" rule shared by LEFTB / RIGHTB / MIDB.

#include <array>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "test_eval_helpers.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry. Mirrors
// the helper used by `builtins_text_test.cpp`.
Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena, default_registry(), test::mac_context());
}

Value EvalSourceWithHost(std::string_view src, ExcelHost host) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  const EvalContext ctx = test::host_context(host);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

struct TextProfileCase {
  ExcelProfile profile;
  const char* id;
};

constexpr std::array<TextProfileCase, 4> kTextProfiles = {{
    {mac_365_ja_jp_profile(), "mac-365-ja_JP"},
    {win_365_ja_jp_profile(), "win-365-ja_JP"},
    {mac_365_en_us_profile(), "mac-365-en_US"},
    {win_365_en_us_profile(), "win-365-en_US"},
}};

Value EvalWithProfile(std::string_view src, ExcelProfile profile) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(profile);
  return formulon::test::EvalSourceIn(src, wb, wb.sheet(0));
}

// ---------------------------------------------------------------------------
// LENB
// ---------------------------------------------------------------------------

TEST(TextLenb, AsciiEachByteIsOne) {
  const Value v = EvalSource("=LENB(\"hello\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(TextLenb, EmptyStringIsZero) {
  const Value v = EvalSource("=LENB(\"\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(TextLenb, HiraganaEachIsTwoBytes) {
  // "あいう" = 3 hiragana chars * 2 bytes.
  const Value v = EvalSource("=LENB(\"あいう\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(TextLenb, KanjiEachIsTwoBytes) {
  // "日本語" = 3 kanji chars * 2 bytes.
  const Value v = EvalSource("=LENB(\"日本語\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(TextLenb, MixedAsciiAndHiragana) {
  // "abcあ" = 3 ASCII (1 byte each) + 1 hiragana (2 bytes) = 5.
  const Value v = EvalSource("=LENB(\"abcあ\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(TextLenb, HalfWidthKatakanaEachIsOneByte) {
  // "ｱｲｳ" (half-width katakana) = 3 chars * 1 byte each.
  const Value v = EvalSource("=LENB(\"ｱｲｳ\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(TextLenb, FullWidthDigitsAreTwoBytes) {
  // Full-width digits U+FF10..U+FF19 are outside the half-width katakana
  // block so they should count as 2 bytes each.
  const Value v = EvalSource("=LENB(\"１２３\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(TextLenb, SupplementaryPlaneIsTwoBytes) {
  // "😀" U+1F600: Mac Excel ja-JP counts supplementary-plane codepoints as
  // 2 bytes (oracle-verified 2026-04-23), not 4 as a naive
  // surrogate-pair-times-two would suggest.
  const Value v = EvalSource("=LENB(\"😀\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(TextLenb, CoercesNumber) {
  // LENB(123) -> LENB("123") -> 3.
  const Value v = EvalSource("=LENB(123)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(TextLenb, ErrorPropagates) {
  const Value v = EvalSource("=LENB(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// LEFTB
// ---------------------------------------------------------------------------

TEST(TextLeftb, AsciiFirstThree) {
  const Value v = EvalSource("=LEFTB(\"hello\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hel");
}

TEST(TextLeftb, DefaultCountIsOne) {
  const Value v = EvalSource("=LEFTB(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "h");
}

TEST(TextLeftb, HiraganaExactBoundary) {
  // "あいう" 4 bytes -> 2 full hiragana chars.
  const Value v = EvalSource("=LEFTB(\"あいう\", 4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "あい");
}

TEST(TextLeftb, HiraganaSplitPadsSpace) {
  // Budget = 3: one full char (2 bytes) + 1 byte remaining; the next char
  // costs 2 bytes, so we pad with a single ASCII space.
  const Value v = EvalSource("=LEFTB(\"あいう\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "あ ");
}

TEST(TextLeftb, ZeroCount) {
  const Value v = EvalSource("=LEFTB(\"hello\", 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextLeftb, NegativeCountIsValueError) {
  const Value v = EvalSource("=LEFTB(\"hello\", -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextLeftb, CountExceedsLength) {
  const Value v = EvalSource("=LEFTB(\"ab\", 100)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

TEST(TextLeftb, HugeCountBeyondIntRangeReturnsWholeText) {
  // 1E+15 overflows int32; the clamp-before-cast must be host-independent.
  const Value v = EvalSource("=LEFTB(\"ab\", 1E+15)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

TEST(TextLeftb, MixedAsciiHiragana) {
  // "aあb" = 'a' (1) + 'あ' (2) + 'b' (1) = 4 bytes.
  // LEFTB(...,2) consumes 'a' (1) then tries 'あ' (2); 1 byte remaining,
  // so pad with a space. Expected: "a ".
  const Value v = EvalSource("=LEFTB(\"aあb\", 2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a ");
}

// ---------------------------------------------------------------------------
// RIGHTB
// ---------------------------------------------------------------------------

TEST(TextRightb, AsciiLastThree) {
  const Value v = EvalSource("=RIGHTB(\"hello\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "llo");
}

TEST(TextRightb, DefaultCountIsOne) {
  const Value v = EvalSource("=RIGHTB(\"hello\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "o");
}

TEST(TextRightb, HiraganaExactBoundary) {
  // "あいう" -> rightmost 4 bytes -> "いう".
  const Value v = EvalSource("=RIGHTB(\"あいう\", 4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "いう");
}

TEST(TextRightb, HiraganaSplitPadsSpace) {
  // "あいう" with budget 3: rightmost full "う" (2 bytes) + 1 byte; next
  // char costs 2 -> pad with space on the left.
  const Value v = EvalSource("=RIGHTB(\"あいう\", 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), " う");
}

TEST(TextRightb, ZeroCount) {
  const Value v = EvalSource("=RIGHTB(\"hello\", 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextRightb, NegativeCountIsValueError) {
  const Value v = EvalSource("=RIGHTB(\"hello\", -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextRightb, CountExceedsLength) {
  const Value v = EvalSource("=RIGHTB(\"ab\", 100)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ab");
}

// ---------------------------------------------------------------------------
// MIDB
// ---------------------------------------------------------------------------

TEST(TextMidb, AsciiMiddleWindow) {
  const Value v = EvalSource("=MIDB(\"hello\", 2, 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ell");
}

TEST(TextMidb, HiraganaCleanBoundary) {
  // "あいうえお" bytes: [あ:1-2][い:3-4][う:5-6][え:7-8][お:9-10].
  // MIDB(...,3,4) = bytes 3..6 = "いう".
  const Value v = EvalSource("=MIDB(\"あいうえお\", 3, 4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "いう");
}

TEST(TextMidb, HiraganaStartSplitsCharPadsSpace) {
  // MIDB(...,2,4): start_byte=2 falls inside "あ" (bytes 1..2). Excel pads
  // the first byte with a space and continues from byte 3. Remaining budget
  // = 3 bytes: include "い" (2) then 1 byte overflow on "う" -> trailing
  // space. Expected: " い ".
  const Value v = EvalSource("=MIDB(\"あいうえお\", 2, 4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), " い ");
}

TEST(TextMidb, StartBeyondLengthYieldsEmpty) {
  const Value v = EvalSource("=MIDB(\"abc\", 10, 3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextMidb, StartZeroIsValueError) {
  const Value v = EvalSource("=MIDB(\"abc\", 0, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextMidb, NegativeNumBytesIsValueError) {
  const Value v = EvalSource("=MIDB(\"abc\", 1, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextMidb, ZeroLengthYieldsEmpty) {
  const Value v = EvalSource("=MIDB(\"abc\", 2, 0)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

// ---------------------------------------------------------------------------
// CHAR
// ---------------------------------------------------------------------------

TEST(TextChar, AsciiLetterA) {
  const Value v = EvalSource("=CHAR(65)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "A");
}

TEST(TextChar, Space) {
  const Value v = EvalSource("=CHAR(32)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), " ");
}

TEST(TextChar, ZeroIsValueError) {
  const Value v = EvalSource("=CHAR(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextChar, TooLargeIsValueError) {
  const Value v = EvalSource("=CHAR(256)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextChar, NegativeIsValueError) {
  const Value v = EvalSource("=CHAR(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextChar, TruncatesFractional) {
  // 65.9 truncates to 65 -> "A".
  const Value v = EvalSource("=CHAR(65.9)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "A");
}

TEST(TextChar, HalfWidthKatakanaViaCp932) {
  // CP932 maps 0xA9 to half-width katakana small-u U+FF69.
  // 169 - 0xA1 = 8, so U+FF61 + 8 = U+FF69 = "ｩ".
  // UTF-8 encoding of U+FF69 = 0xEF 0xBD 0xA9.
  const Value v = EvalSource("=CHAR(169)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xEF\xBD\xA9");
}

TEST(TextChar, HalfWidthKatakanaFirstSlot) {
  // 0xA1 -> U+FF61 = "｡" (half-width ideographic full stop).
  const Value v = EvalSource("=CHAR(161)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "\xEF\xBD\xA1");
}

// ---------------------------------------------------------------------------
// CODE
// ---------------------------------------------------------------------------

TEST(TextCode, AsciiLetterA) {
  const Value v = EvalSource("=CODE(\"A\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 65.0);
}

TEST(TextCode, EmptyStringIsValueError) {
  const Value v = EvalSource("=CODE(\"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextCode, UsesFirstCharOnly) {
  const Value v = EvalSource("=CODE(\"Apple\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 65.0);
}

TEST(TextCode, MacHostHiraganaReturnsJisRowCell) {
  // "あ" = U+3042. Mac Excel ja-JP returns the JIS X 0208 row-cell
  // encoding (row 4, cell 2) -> ((4+0x20)<<8) | (2+0x20) = 0x2422 = 9250.
  const Value v = EvalSource("=CODE(\"あ\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 9250.0);
}

TEST(TextCode, MacHostEmojiFallsBackToUnderscore) {
  // Codepoints outside JIS X 0208 (here U+1F600) fall back to 95 (the
  // underscore byte) per Mac Excel's empirically confirmed behaviour.
  const Value v = EvalSource("=CODE(\"\xF0\x9F\x98\x80\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 95.0);
}

TEST(TextCode, WinHostUsesCp932Fallbacks) {
  Value v = EvalSourceWithHost("=CODE(\"髙\")", ExcelHost::kWin365);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 38526.0);

  v = EvalSourceWithHost("=CODE(\"\xF0\x9F\x98\x80\")", ExcelHost::kWin365);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 63.0);
}

TEST(TextCode, ErrorPropagates) {
  const Value v = EvalSource("=CODE(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(TextProfileCodepage, CharAndCodeUseTheSelectedProfile) {
  for (const TextProfileCase& test_case : kTextProfiles) {
    SCOPED_TRACE(test_case.id);
    const bool japanese = test_case.profile.locale == ExcelLocale::kJaJP;
    const bool mac = test_case.profile.host == ExcelHost::kMac365;

    if (!japanese) {
      const Value char_high = EvalWithProfile("=CHAR(130)", test_case.profile);
      ASSERT_TRUE(char_high.is_text());
      if (mac) {
        EXPECT_EQ(char_high.as_text(), "\xC3\x87");
      } else {
        EXPECT_EQ(char_high.as_text(), "\xE2\x80\x9A");
      }
    }

    const Value char_katakana = EvalWithProfile("=CHAR(169)", test_case.profile);
    ASSERT_TRUE(char_katakana.is_text());
    EXPECT_EQ(char_katakana.as_text(), japanese ? "\xEF\xBD\xA9" : "\xC2\xA9");

    const Value code_accent = EvalWithProfile("=CODE(\"é\")", test_case.profile);
    ASSERT_TRUE(code_accent.is_number());
    const double expected_accent = japanese ? 95.0 : (mac ? 142.0 : 233.0);
    EXPECT_DOUBLE_EQ(code_accent.as_number(), expected_accent);

    const Value code_hiragana = EvalWithProfile("=CODE(\"あ\")", test_case.profile);
    ASSERT_TRUE(code_hiragana.is_number());
    EXPECT_DOUBLE_EQ(code_hiragana.as_number(), japanese ? 9250.0 : 95.0);
  }
}

TEST(TextProfileCodepage, CharArgumentsKeepRangeAndFallbackRules) {
  for (const TextProfileCase& test_case : kTextProfiles) {
    SCOPED_TRACE(test_case.id);
    const Value near_integer = EvalWithProfile("=CODE(CHAR(64.9999999))", test_case.profile);
    ASSERT_TRUE(near_integer.is_number());
    const double expected_near_integer = test_case.profile.locale == ExcelLocale::kEnUS ? 65.0 : 64.0;
    EXPECT_DOUBLE_EQ(near_integer.as_number(), expected_near_integer);

    const Value out_of_range = EvalWithProfile("=CHAR(256)", test_case.profile);
    ASSERT_TRUE(out_of_range.is_error());
    EXPECT_EQ(out_of_range.as_error(), ErrorCode::Value);

    const Value unmapped = EvalWithProfile("=CODE(\"☃\")", test_case.profile);
    ASSERT_TRUE(unmapped.is_number());
    EXPECT_DOUBLE_EQ(unmapped.as_number(), 95.0);
  }
}

TEST(TextProfileByteFunctions, NonDbcsUsesUtf16UnitsForSupplementaryText) {
  const std::string text =
      "あ\xF0\x9F\x98\x80"
      "B";
  const std::string lenb_formula = "=LENB(\"" + text + "\")";
  const std::string leftb_formula = "=LEFTB(\"" + text + "\",3)";
  const std::string rightb_formula = "=RIGHTB(\"" + text + "\",2)";
  const std::string midb_formula = "=MIDB(\"" + text + "\",2,1)";
  const std::string findb_formula = "=FINDB(\"B\",\"" + text + "\")";
  const std::string searchb_formula = "=SEARCHB(\"B\",\"" + text + "\")";
  const std::string replaceb_formula = "=REPLACEB(\"abcdef\",2,3,\"X\")";
  const std::string replaceb_non_ascii_formula = "=REPLACEB(\"あいう\",2,1,\"X\")";

  for (const TextProfileCase& test_case : kTextProfiles) {
    SCOPED_TRACE(test_case.id);
    const bool japanese = test_case.profile.locale == ExcelLocale::kJaJP;

    const Value lenb = EvalWithProfile(lenb_formula, test_case.profile);
    ASSERT_TRUE(lenb.is_number());
    EXPECT_DOUBLE_EQ(lenb.as_number(), japanese ? 5.0 : 4.0);

    const Value leftb = EvalWithProfile(leftb_formula, test_case.profile);
    ASSERT_TRUE(leftb.is_text());
    EXPECT_EQ(leftb.as_text(), japanese ? "あ " : "あ\xF0\x9F\x98\x80");

    const Value rightb = EvalWithProfile(rightb_formula, test_case.profile);
    ASSERT_TRUE(rightb.is_text());
    const std::string expected_right = japanese ? std::string(" ") + "B" : std::string("\xF0\x9F\x98\x80") + "B";
    EXPECT_EQ(rightb.as_text(), expected_right);

    const Value midb = EvalWithProfile(midb_formula, test_case.profile);
    ASSERT_TRUE(midb.is_text());
    EXPECT_EQ(midb.as_text(), japanese ? " " : "\xF0\x9F\x98\x80");

    const Value findb = EvalWithProfile(findb_formula, test_case.profile);
    ASSERT_TRUE(findb.is_number());
    EXPECT_DOUBLE_EQ(findb.as_number(), japanese ? 5.0 : 4.0);

    const Value searchb = EvalWithProfile(searchb_formula, test_case.profile);
    ASSERT_TRUE(searchb.is_number());
    EXPECT_DOUBLE_EQ(searchb.as_number(), japanese ? 5.0 : 4.0);

    const Value replaceb = EvalWithProfile(replaceb_formula, test_case.profile);
    ASSERT_TRUE(replaceb.is_text());
    EXPECT_EQ(replaceb.as_text(), "aXef");

    const Value replaceb_non_ascii = EvalWithProfile(replaceb_non_ascii_formula, test_case.profile);
    ASSERT_TRUE(replaceb_non_ascii.is_text());
    EXPECT_EQ(replaceb_non_ascii.as_text(), japanese ? " Xいう" : "あXう");
  }
}

TEST(TextProfileLazyDispatch, CodeAndLenbUseLocaleFactsOnWindows) {
  for (const TextProfileCase& test_case : kTextProfiles) {
    SCOPED_TRACE(test_case.id);
    const bool japanese = test_case.profile.locale == ExcelLocale::kJaJP;
    const bool win = test_case.profile.host == ExcelHost::kWin365;

    const Value code = EvalWithProfile("=CODE(\"\xF0\x9F\x98\x80\")", test_case.profile);
    ASSERT_TRUE(code.is_number());
    EXPECT_DOUBLE_EQ(code.as_number(), japanese && win ? 63.0 : 95.0);

    const Value lenb = EvalWithProfile("=LENB(\"あ\")", test_case.profile);
    ASSERT_TRUE(lenb.is_number());
    EXPECT_DOUBLE_EQ(lenb.as_number(), japanese ? 2.0 : 1.0);
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
