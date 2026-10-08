//
// End-to-end tests for TEXT, VALUE, and NUMBERVALUE. Each test parses a
// formula source and evaluates it through the default registry.

#include "value.h"

#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/text_format/number_format.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;

// ---------------------------------------------------------------------------
// TEXT
// ---------------------------------------------------------------------------

TEST(TextFunctionText, IntegerFormat) {
  const Value v = EvalSource("=TEXT(1234, \"#,##0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1,234");
}

TEST(TextFunctionText, BoundedRangeSpillsElementwise) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0U, 0U, Value::number(12.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1U, 0U, Value::number(34.0))));
  const Value result = EvalSourceIn("=TEXT(A1:A2,\"00\")", wb, wb.sheet(0));
  ASSERT_TRUE(result.is_array());
  ASSERT_EQ(result.as_array_rows(), 2U);
  ASSERT_EQ(result.as_array_cols(), 1U);
  EXPECT_EQ(result.as_array()->cells[0].as_text(), "12");
  EXPECT_EQ(result.as_array()->cells[1].as_text(), "34");
}

TEST(TextFunctionText, TwoDecimals) {
  const Value v = EvalSource("=TEXT(3.14159, \"0.00\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3.14");
}

TEST(TextFunctionText, RoundsTiesAwayFromZero) {
  EXPECT_EQ(EvalSource("=TEXT(2.5, \"0\")").as_text(), "3");
  EXPECT_EQ(EvalSource("=TEXT(1234.5, \"0\")").as_text(), "1235");
  EXPECT_EQ(EvalSource("=TEXT(-2.5, \"0\")").as_text(), "-3");
}

TEST(TextFunctionText, LimitsDisplayToFifteenSignificantDigits) {
  const Value v = EvalSource("=TEXT(0.1+0.2, \"0.00000000000000000\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "0.30000000000000000");
}

TEST(TextFunctionText, Percent) {
  const Value v = EvalSource("=TEXT(0.1234, \"0.00%\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "12.34%");
}

TEST(TextFunctionText, Negative) {
  const Value v = EvalSource("=TEXT(-5, \"0.00;(0.00)\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "(5.00)");
}

TEST(TextFunctionText, EmptyFormat) {
  const Value v = EvalSource("=TEXT(42, \"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextFunctionText, IsoDate) {
  const Value v = EvalSource("=TEXT(45366, \"yyyy-mm-dd\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "2024-03-15");
}

TEST(TextFunctionText, KanjiDate) {
  const Value v = EvalSource("=TEXT(45366, \"yyyy年m月d日\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "2024年3月15日");
}

TEST(TextFunctionText, MalformedFormatWithNonNumericTextIsValueError) {
  // This malformed format remains a value error even when the first
  // argument is non-numeric text.
  const Value v = EvalSource("=TEXT(\"abc\", \"text is @\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextFunctionText, BoolTrueReturnsUppercaseText) {
  // A numeric format has no text placeholder, so the original bool spelling
  // passes through the text fallback.
  const Value v = EvalSource("=TEXT(TRUE, \"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "TRUE");
}

TEST(TextFunctionText, BoolFalseReturnsUppercaseText) {
  const Value v = EvalSource("=TEXT(FALSE, \"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "FALSE");
}

TEST(TextFunctionText, NumericTextStillCoerces) {
  // Numeric strings still parse and render through the format.
  const Value v = EvalSource("=TEXT(\"42\", \"0.00\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "42.00");
}

TEST(TextFunctionText, DateTextCoercesThroughTheSharedLadder) {
  // The first argument goes through `coerce_to_number`, so every text shape
  // arithmetic accepts is formattable here too. A date string reaches the
  // date fallback and renders through the format codes.
  const Value iso = EvalSource("=TEXT(\"2024-03-15\", \"yyyy\")");
  ASSERT_TRUE(iso.is_text()) << "date text must not be rejected as non-coercible";
  EXPECT_EQ(iso.as_text(), "2024");

  const Value slash = EvalSource("=TEXT(\"2024/3/15\", \"yyyy/m/d\")");
  ASSERT_TRUE(slash.is_text());
  EXPECT_EQ(slash.as_text(), "2024/3/15");
}

TEST(TextFunctionText, TimeTextCoercesThroughTheSharedLadder) {
  const Value v = EvalSource("=TEXT(\"13:30\", \"h:mm\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "13:30");
}

TEST(TextFunctionText, EmptyStringRetainsTextForNumericFormat) {
  // An empty text value falls back to text rendering when numeric coercion
  // fails, even with a numeric-looking format.
  const Value v = EvalSource("=TEXT(\"\", \"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextFunctionText, NonNumericTextFallsBackToTextFormats) {
  const Value plain = EvalSource("=TEXT(\"abc\", \"@\")");
  ASSERT_TRUE(plain.is_text());
  EXPECT_EQ(plain.as_text(), "abc");

  const Value suffix = EvalSource("=TEXT(\"abc\", \"@\"\"post\"\"\")");
  ASSERT_TRUE(suffix.is_text());
  EXPECT_EQ(suffix.as_text(), "abcpost");

  const Value fourth = EvalSource("=TEXT(\"abc\", \"0;0;0;\"\"pre\"\"@\")");
  ASSERT_TRUE(fourth.is_text());
  EXPECT_EQ(fourth.as_text(), "preabc");

  const Value numeric_format = EvalSource("=TEXT(\"abc\", \"0.00\")");
  ASSERT_TRUE(numeric_format.is_text());
  EXPECT_EQ(numeric_format.as_text(), "abc");
}

TEST(TextFunctionText, NumericTextOnlyFormatRetainsNumericRendering) {
  const Value at = EvalSource("=TEXT(12, \"@\")");
  ASSERT_TRUE(at.is_text());
  EXPECT_EQ(at.as_text(), "12");

  const Value literal_at = EvalSource("=TEXT(12, \"\"\"pre\"\"@\")");
  ASSERT_TRUE(literal_at.is_text());
  EXPECT_EQ(literal_at.as_text(), "12");
}

TEST(TextFunctionText, BoolUsesTextFormatsAndValidatesFormatErrors) {
  const Value suffix = EvalSource("=TEXT(TRUE, \"@\"\"post\"\"\")");
  ASSERT_TRUE(suffix.is_text());
  EXPECT_EQ(suffix.as_text(), "TRUEpost");

  const Value fourth = EvalSource("=TEXT(TRUE, \"0;0;0;\"\"pre\"\"@\")");
  ASSERT_TRUE(fourth.is_text());
  EXPECT_EQ(fourth.as_text(), "preTRUE");

  const Value invalid = EvalSource("=TEXT(TRUE, \"[invalid]0\")");
  ASSERT_TRUE(invalid.is_error());
  EXPECT_EQ(invalid.as_error(), ErrorCode::Value);
}

TEST(TextFunctionText, NonNumericTextValidatesEveryFormatSection) {
  const Value invalid = EvalSource("=TEXT(\"abc\", \"@;[invalid]0\")");
  ASSERT_TRUE(invalid.is_error());
  EXPECT_EQ(invalid.as_error(), ErrorCode::Value);
}

TEST(TextFunctionText, NumericValuesValidateEveryFormatSection) {
  const Value positive_unused = EvalSource("=TEXT(1, \"0;[invalid]0\")");
  ASSERT_TRUE(positive_unused.is_error());
  EXPECT_EQ(positive_unused.as_error(), ErrorCode::Value);

  const Value negative_unused = EvalSource("=TEXT(-1, \"[invalid]0;0\")");
  ASSERT_TRUE(negative_unused.is_error());
  EXPECT_EQ(negative_unused.as_error(), ErrorCode::Value);

  const Value text_placeholder_unused = EvalSource("=TEXT(1, \"0;0 @\")");
  ASSERT_TRUE(text_placeholder_unused.is_error());
  EXPECT_EQ(text_placeholder_unused.as_error(), ErrorCode::Value);
}

TEST(TextFunctionText, TooManySectionsAreValueError) {
  const Value numeric = EvalSource("=TEXT(1, \"0;0;0;@;0\")");
  ASSERT_TRUE(numeric.is_error());
  EXPECT_EQ(numeric.as_error(), ErrorCode::Value);

  const Value text = EvalSource("=TEXT(\"abc\", \"0;0;0;@;0\")");
  ASSERT_TRUE(text.is_error());
  EXPECT_EQ(text.as_error(), ErrorCode::Value);

  const Value boolean = EvalSource("=TEXT(TRUE, \"0;0;0;@;0\")");
  ASSERT_TRUE(boolean.is_error());
  EXPECT_EQ(boolean.as_error(), ErrorCode::Value);
}

TEST(TextFunctionText, NumericTextSectionFallsBackToGeneralForNegativeValues) {
  const Value v = EvalSource("=TEXT(-12, \"0;@\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "-12");
}

TEST(TextFunctionText, ConditionalFormatsValidateAndKeepOutputAtomic) {
  std::string out = "prefix";
  EXPECT_EQ(::formulon::text_format::apply_format(1.0, "[>10]0;[<0]0", out),
            ::formulon::text_format::FormatStatus::kOverflow);
  EXPECT_EQ(out, "prefix");

  out = "prefix";
  EXPECT_EQ(::formulon::text_format::apply_format(1.0, "[>10]0;[<0]0;[=0]0", out),
            ::formulon::text_format::FormatStatus::kValueError);
  EXPECT_EQ(out, "prefix");
}

TEST(TextFunctionText, ConditionalFormatsIntegrateThroughText) {
  const Value invalid = EvalSource("=TEXT(1, \"[>10]0;[<0]0\")");
  ASSERT_TRUE(invalid.is_error());
  EXPECT_EQ(invalid.as_error(), ErrorCode::Value);

  const Value literal_fallback = EvalSource("=TEXT(1, \"[>10]0;[<0]0;\"\"else\"\"\")");
  ASSERT_TRUE(literal_fallback.is_text());
  EXPECT_EQ(literal_fallback.as_text(), "else");

  const Value third_predicate = EvalSource("=TEXT(1, \"[>10]0;[<0]0;[=0]0\")");
  ASSERT_TRUE(third_predicate.is_error());
  EXPECT_EQ(third_predicate.as_error(), ErrorCode::Value);
}

TEST(TextFunctionText, ConditionalNegativeSelectionUsesPredicateMagnitude) {
  struct FormatCase {
    std::string_view format;
    std::string_view expected;
  };
  constexpr FormatCase cases[] = {
      {"[<0]0;0", "5"},  {"[<-1]0;0", "5"},  {"[<=-1]0;0", "5"}, {"[=-5]0;0", "5"},   {"[<=0]0;0", "-5"},
      {"[<1]0;0", "-5"}, {"[<=1]0;0", "-5"}, {"[<>0]0;0", "-5"}, {"[>-10]0;0", "-5"}, {"[>0]0;0", "5"},
      {"[>=0]0;0", "5"}, {"[>10]0;0", "-5"}, {"[>=1]0;0", "-5"},
  };
  for (const FormatCase& test_case : cases) {
    const std::string formula = "=TEXT(-5, \"" + std::string(test_case.format) + "\")";
    const Value value = EvalSource(formula);
    ASSERT_TRUE(value.is_text()) << test_case.format;
    EXPECT_EQ(value.as_text(), test_case.expected) << test_case.format;
  }
}

TEST(TextFunctionText, ConditionalNegativeExplicitMinusIsNotDuplicated) {
  const Value value = EvalSource("=TEXT(-1, \"[>10]-0;[<0]-0\")");
  ASSERT_TRUE(value.is_text());
  EXPECT_EQ(value.as_text(), "-1");
}

TEST(TextFunctionText, ConditionalLiteralSignsFollowSelectedPredicate) {
  const Value leading = EvalSource("=TEXT(-1.5, \"[<=0]-0.00;0.00\")");
  ASSERT_TRUE(leading.is_text());
  EXPECT_EQ(leading.as_text(), "--1.50");

  const Value prefix = EvalSource("=TEXT(-1.5, \"[<=0]\"\"pre-\"\"0.00;0.00\")");
  ASSERT_TRUE(prefix.is_text());
  EXPECT_EQ(prefix.as_text(), "-pre-1.50");

  const Value suffix = EvalSource("=TEXT(-1.5, \"[<=0]0.00-;0.00\")");
  ASSERT_TRUE(suffix.is_text());
  EXPECT_EQ(suffix.as_text(), "-1.50-");
}

TEST(TextFunctionText, ConditionalSecondPredicateAndThreeSectionFallbackUseMagnitude) {
  const Value selected = EvalSource("=TEXT(-1, \"[>10]0;[<0]0\")");
  ASSERT_TRUE(selected.is_text());
  EXPECT_EQ(selected.as_text(), "1");

  const Value fallback = EvalSource("=TEXT(-1, \"[>10]0;0;0\")");
  ASSERT_TRUE(fallback.is_text());
  EXPECT_EQ(fallback.as_text(), "1");
}

TEST(TextFunctionText, ConditionalSingleSectionNoMatchFallsThrough) {
  const Value value = EvalSource("=TEXT(1, \"[>10]0\")");
  ASSERT_TRUE(value.is_text());
  EXPECT_EQ(value.as_text(), "1");
}

TEST(TextFunctionText, ConditionalDispatchUsesImplicitSignArms) {
  const auto eval_text = [](std::string_view number, std::string_view format) {
    std::string escaped_format;
    escaped_format.reserve(format.size());
    for (const char ch : format) {
      if (ch == '"') {
        escaped_format.append("\"\"");
      } else {
        escaped_format.push_back(ch);
      }
    }
    const std::string formula = "=TEXT(" + std::string(number) + ",\"" + escaped_format + "\")";
    return EvalSource(formula);
  };
  const auto expect_text = [&](std::string_view number, std::string_view format, std::string_view expected) {
    const Value value = eval_text(number, format);
    ASSERT_TRUE(value.is_text()) << format;
    EXPECT_EQ(value.as_text(), expected) << format;
  };
  const auto expect_value_error = [&](std::string_view number, std::string_view format) {
    const Value value = eval_text(number, format);
    ASSERT_TRUE(value.is_error()) << format;
    EXPECT_EQ(value.as_error(), ErrorCode::Value) << format;
  };

  constexpr std::string_view first_predicate = "[>10]\"first\"0;\"second\"0;\"third\"0";
  expect_text("-5", first_predicate, "second5");
  expect_text("5", first_predicate, "third5");
  expect_text("0", first_predicate, "third0");
  expect_text("15", first_predicate, "first15");
  expect_text("-5", "[<0]\"first\"0;\"second\"0;\"third\"0", "first5");

  constexpr std::string_view second_predicate = "\"first\"0;[>10]\"second\"0;\"third\"0";
  expect_text("-5", second_predicate, "-third5");
  expect_text("5", second_predicate, "first5");
  expect_text("0", second_predicate, "third0");
  expect_text("15", second_predicate, "first15");
  expect_text("-5", "\"first\"0;[<0]\"second\"0;\"third\"0", "second5");

  expect_text("5", "0;[>10]0", "5");
  expect_text("15", "0;[>10]0", "15");
  expect_value_error("-5", "0;[>10]0");
  expect_value_error("0", "0;[>10]0");
  expect_text("-5", "0;[<0]0", "5");
  expect_value_error("0", "0;[<0]0");

  expect_text("-5", "[>=-10]0", "-5");
  expect_text("-15", "[>=-10]0", "15");
}

TEST(TextFunctionText, ConditionalSingleSectionSignBoundaries) {
  const Value equal_zero = EvalSource("=TEXT(-5, \"[=0]0\")");
  ASSERT_TRUE(equal_zero.is_text());
  EXPECT_EQ(equal_zero.as_text(), "-5");

  const Value equal_negative = EvalSource("=TEXT(-5, \"[=-10]0\")");
  ASSERT_TRUE(equal_negative.is_text());
  EXPECT_EQ(equal_negative.as_text(), "-5");

  const Value not_equal_at_point = EvalSource("=TEXT(-5, \"[<>-5]0\")");
  ASSERT_TRUE(not_equal_at_point.is_text());
  EXPECT_EQ(not_equal_at_point.as_text(), "5");

  const Value not_equal_fallback = EvalSource("=TEXT(-5, \"[<>-5]0;0\")");
  ASSERT_TRUE(not_equal_fallback.is_text());
  EXPECT_EQ(not_equal_fallback.as_text(), "5");

  const Value not_equal_match = EvalSource("=TEXT(-5, \"[<>-10]0\")");
  ASSERT_TRUE(not_equal_match.is_text());
  EXPECT_EQ(not_equal_match.as_text(), "-5");

  const Value less_than_false = EvalSource("=TEXT(-5, \"[<-10]0\")");
  ASSERT_TRUE(less_than_false.is_text());
  EXPECT_EQ(less_than_false.as_text(), "5");
}

TEST(TextFunctionText, EmptyFormatKeepsKindSpecificFallbacks) {
  const Value numeric = EvalSource("=TEXT(12, \"\")");
  ASSERT_TRUE(numeric.is_text());
  EXPECT_EQ(numeric.as_text(), "");

  const Value text = EvalSource("=TEXT(\"abc\", \"\")");
  ASSERT_TRUE(text.is_text());
  EXPECT_EQ(text.as_text(), "abc");

  const Value boolean = EvalSource("=TEXT(TRUE, \"\")");
  ASSERT_TRUE(boolean.is_text());
  EXPECT_EQ(boolean.as_text(), "TRUE");
}

TEST(TextFunctionText, NumErrorPropagatesBeforeFormatRouting) {
  const Value v = EvalSource("=TEXT(#NUM!, \"0\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(TextFunctionText, FormatErrorPropagatesBeforeKindRouting) {
  const Value v = EvalSource("=TEXT(TRUE, #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(TextFunctionText, DateTextUsesWorkbookEpochAfterNumericCoercion) {
  Workbook wb = Workbook::create();
  wb.set_date1904(true);
  const Value v = EvalSourceIn("=TEXT(\"2024-03-15\", \"yyyy-mm-dd\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "2024-03-15");
}

TEST(TextFunctionText, ReferencesPreserveCellKindsAndFormatCells) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::text("123"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::text("0.00"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::text("abc"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 1U, Value::text("@\"post\""))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 2U, 0U, Value::boolean(true))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 2U, 1U, Value::text("0;0;0;\"pre\"@"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 3U, 0U, Value::text(""))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 3U, 1U, Value::text("0"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=TEXT(A1,B1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 2U, "=TEXT(A2,B2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 2U, "=TEXT(A3,B3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 3U, 2U, "=TEXT(A4,B4)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));

  const auto expect_text = [&](std::uint32_t row, std::string_view expected) {
    const Value value = wb.sheet(0).resolve_cell_value(row, 2U);
    ASSERT_TRUE(value.is_text()) << "C" << (row + 1U);
    EXPECT_EQ(value.as_text(), expected) << "C" << (row + 1U);
  };
  expect_text(0U, "123.00");
  expect_text(1U, "abcpost");
  expect_text(2U, "preTRUE");
  expect_text(3U, "");
}

TEST(TextFunctionText, ReferenceFormatAndValueMutationsRecalculateDependents) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::text("abc"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::text("\"pre\"@"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=TEXT(A1,B1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  ASSERT_TRUE(wb.sheet(0).resolve_cell_value(0U, 2U).is_text());
  EXPECT_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_text(), "preabc");

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::text("@\"post\""))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  ASSERT_TRUE(wb.sheet(0).resolve_cell_value(0U, 2U).is_text());
  EXPECT_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_text(), "abcpost");

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::boolean(true))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  ASSERT_TRUE(wb.sheet(0).resolve_cell_value(0U, 2U).is_text());
  EXPECT_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_text(), "TRUEpost");

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=1/0")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  const Value result = wb.sheet(0).resolve_cell_value(0U, 2U);
  ASSERT_TRUE(result.is_error());
  EXPECT_EQ(result.as_error(), ErrorCode::Div0);
}

TEST(TextFunctionText, ErrorPropagates) {
  const Value v = EvalSource("=TEXT(#REF!, \"0\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(TextFunctionText, LocalizedFormatTokensFollowWorkbookLocale) {
  const ExcelProfile profiles[] = {mac_365_ja_jp_profile(), win_365_ja_jp_profile(), mac_365_en_us_profile(),
                                   win_365_en_us_profile()};
  for (const ExcelProfile profile : profiles) {
    SCOPED_TRACE(excel_profile_id(profile));
    Workbook workbook = Workbook::create();
    workbook.set_excel_profile(profile);

    const Value localized = EvalSourceIn("=TEXT(5, \"[赤]0.00\")", workbook, workbook.sheet(0));
    if (profile.locale == ExcelLocale::kJaJP) {
      ASSERT_TRUE(localized.is_text());
      const std::string localized_text(localized.as_text());
      EXPECT_EQ(localized_text, "5.00");
    } else {
      ASSERT_TRUE(localized.is_error());
      EXPECT_EQ(localized.as_error(), ErrorCode::Value);
    }

    const Value stored = EvalSourceIn("=TEXT(5, \"[Red]0.00\")", workbook, workbook.sheet(0));
    if (profile.locale == ExcelLocale::kJaJP) {
      ASSERT_TRUE(stored.is_error());
      EXPECT_EQ(stored.as_error(), ErrorCode::Value);
    } else {
      ASSERT_TRUE(stored.is_text());
      EXPECT_EQ(stored.as_text(), "5.00");
    }
  }
}

TEST(TextFunctionText, LocaleSpecificEscapesAndDbNumFollowWorkbookLocale) {
  const ExcelProfile profiles[] = {mac_365_ja_jp_profile(), win_365_ja_jp_profile(), mac_365_en_us_profile(),
                                   win_365_en_us_profile()};
  for (const ExcelProfile profile : profiles) {
    SCOPED_TRACE(excel_profile_id(profile));
    Workbook workbook = Workbook::create();
    workbook.set_excel_profile(profile);

    const Value dbnum = EvalSourceIn("=TEXT(1234, \"[DBNum1]0\")", workbook, workbook.sheet(0));
    ASSERT_TRUE(dbnum.is_text());
    EXPECT_EQ(dbnum.as_text(), profile.locale == ExcelLocale::kJaJP ? "一二三四" : "1234");

    const Value bang = EvalSourceIn("=TEXT(\"hello\", \"0;0;0;@!\")", workbook, workbook.sheet(0));
    if (profile.locale == ExcelLocale::kJaJP) {
      ASSERT_TRUE(bang.is_error());
      EXPECT_EQ(bang.as_error(), ErrorCode::Value);
    } else {
      ASSERT_TRUE(bang.is_text());
      EXPECT_EQ(bang.as_text(), "hello!");
    }
  }
}

TEST(TextFunctionText, DollarDefaultsAndNegativeZeroFollowWorkbookLocale) {
  struct Expected {
    const char* positive;
    const char* negative;
    const char* negative_zero;
  };
  const ExcelProfile profiles[] = {mac_365_ja_jp_profile(), win_365_ja_jp_profile(), mac_365_en_us_profile(),
                                   win_365_en_us_profile()};
  for (const ExcelProfile profile : profiles) {
    SCOPED_TRACE(excel_profile_id(profile));
    Workbook workbook = Workbook::create();
    workbook.set_excel_profile(profile);
    const Expected expected = profile.locale == ExcelLocale::kJaJP ? Expected{"¥1,235", "¥-1,235", "¥-0.00"}
                                                                   : Expected{"$1,234.50", "($1,234.50)", "($0.00)"};

    const Value positive = EvalSourceIn("=DOLLAR(1234.5)", workbook, workbook.sheet(0));
    ASSERT_TRUE(positive.is_text());
    EXPECT_EQ(positive.as_text(), expected.positive);
    const Value negative = EvalSourceIn("=DOLLAR(-1234.5)", workbook, workbook.sheet(0));
    ASSERT_TRUE(negative.is_text());
    EXPECT_EQ(negative.as_text(), expected.negative);
    const Value negative_zero = EvalSourceIn("=DOLLAR(-0.001,2)", workbook, workbook.sheet(0));
    ASSERT_TRUE(negative_zero.is_text());
    EXPECT_EQ(negative_zero.as_text(), expected.negative_zero);
  }
}

// ---------------------------------------------------------------------------
// VALUE
// ---------------------------------------------------------------------------

TEST(ValueFunction, IntegerString) {
  const Value v = EvalSource("=VALUE(\"123\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 123.0);
}

TEST(ValueFunction, DecimalString) {
  const Value v = EvalSource("=VALUE(\"1.5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.5);
}

TEST(ValueFunction, ScientificString) {
  const Value v = EvalSource("=VALUE(\"1.5e2\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 150.0);
}

TEST(ValueFunction, WithCommaThousands) {
  const Value v = EvalSource("=VALUE(\"1,234.5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234.5);
}

TEST(ValueFunction, WithPercent) {
  const Value v = EvalSource("=VALUE(\"50%\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(ValueFunction, WithDollarPrefix) {
  const Value v = EvalSource("=VALUE(\"$100\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 100.0);
}

TEST(ValueFunction, WithLeadingWhitespace) {
  const Value v = EvalSource("=VALUE(\"   42   \")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(ValueFunction, NegativeSign) {
  const Value v = EvalSource("=VALUE(\"-3.14\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -3.14);
}

TEST(ValueFunction, NumberPassthrough) {
  const Value v = EvalSource("=VALUE(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(ValueFunction, BoolRejected) {
  const Value v = EvalSource("=VALUE(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, NonNumericString) {
  const Value v = EvalSource("=VALUE(\"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, ErrorPropagates) {
  const Value v = EvalSource("=VALUE(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(ValueFunction, EuroSuffixInteger) {
  // Mac Excel 365 ja-JP accepts Euro as a trailing currency suffix.
  const Value v = EvalSource("=VALUE(\"23\xE2\x82\xAC\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 23.0);
}

TEST(ValueFunction, EuroSuffixSmall) {
  const Value v = EvalSource("=VALUE(\"12\xE2\x82\xAC\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

TEST(ValueFunction, EuroPrefix) {
  const Value v = EvalSource(
      "=VALUE(\"\xE2\x82\xAC"
      "23\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 23.0);
}

TEST(ValueFunction, DollarSuffixRejected) {
  // Mac Excel ja-JP rejects `$` as a trailing suffix (only Euro is
  // bidirectional).
  const Value v = EvalSource("=VALUE(\"23$\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, YenKanjiSuffixRejected) {
  // The yen kanji (U+5186) is not accepted by Mac Excel.
  const Value v = EvalSource("=VALUE(\"23\xE5\x86\x86\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, InvalidThousandsGrouping) {
  // `"12,34"` is rejected by Mac: first group is 2 digits (OK), but the
  // final group must be exactly 3 digits.
  const Value v = EvalSource("=VALUE(\"12,34\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, ValidThousandsGrouping) {
  const Value v = EvalSource("=VALUE(\"1,234\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234.0);
}

TEST(ValueFunction, ValidThousandsGroupingMultiple) {
  const Value v = EvalSource("=VALUE(\"1,234,567\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234567.0);
}

TEST(ValueFunction, FourDigitFinalGroupRejected) {
  const Value v = EvalSource("=VALUE(\"1,2345\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, LeadingGroupSepRejected) {
  const Value v = EvalSource("=VALUE(\",234\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, TrailingGroupSepRejected) {
  const Value v = EvalSource("=VALUE(\"1,\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, ThousandsWithDecimalAccepted) {
  const Value v = EvalSource("=VALUE(\"1,234.56\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1234.56);
}

TEST(ValueFunction, IsoDate) {
  const Value v = EvalSource("=VALUE(\"2024-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(ValueFunction, TimeOnly) {
  const Value v = EvalSource("=VALUE(\"13:30\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5625);
}

TEST(ValueFunction, DateAndTime) {
  const Value v = EvalSource("=VALUE(\"2024-03-15 12:00\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 45366.5);
}

// Year-less date_text ("3/15") falls back to the wall clock's year, sharing
// DATEVALUE's rule (see DatevalueYearLess in
// builtins_datevalue_timevalue_test.cpp for the pinned-clock coverage of
// that rule itself). This confirms VALUE reaches the same fallback.
TEST(ValueFunction, YearLessDateUsesCurrentYear) {
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=VALUE(\"3/15\")", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  constexpr date_time::CivilTime kPinned{{2026, 4U, 23U}, {15U, 30U, 45U}};
  const Value v = evaluate(*root, eval_arena, default_registry(), EvalContext().with_pinned_now(kPinned));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 46096.0);  // 2026-03-15
}

// ---------------------------------------------------------------------------
// NUMBERVALUE
// ---------------------------------------------------------------------------

TEST(NumberValueFunction, DefaultSeparators) {
  const Value v = EvalSource("=NUMBERVALUE(\"1,234.5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234.5);
}

TEST(NumberValueFunction, CustomDecimalSep) {
  // European-style: decimal is "," and group is "."
  const Value v = EvalSource("=NUMBERVALUE(\"1.234,5\", \",\", \".\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234.5);
}

TEST(NumberValueFunction, TrailingPercent) {
  const Value v = EvalSource("=NUMBERVALUE(\"50%\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(NumberValueFunction, MultipleTrailingPercents) {
  // Each `%` multiplies by 0.01: 50 * 0.01 * 0.01 = 0.005.
  const Value v = EvalSource("=NUMBERVALUE(\"50%%\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.005);
}

TEST(NumberValueFunction, CommaDecimalOnlyNoGroupCollision) {
  // With only `decimal_sep = ","` supplied, grouping is disabled so the
  // default group sep of `,` cannot collide.
  const Value v = EvalSource("=NUMBERVALUE(\"3,14\", \",\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.14);
}

TEST(NumberValueFunction, WhitespaceTrimmed) {
  const Value v = EvalSource("=NUMBERVALUE(\"  3.14  \")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.14);
}

TEST(NumberValueFunction, Scientific) {
  const Value v = EvalSource("=NUMBERVALUE(\"1.5E3\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1500.0);
}

TEST(NumberValueFunction, RejectNonNumeric) {
  const Value v = EvalSource("=NUMBERVALUE(\"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(NumberValueFunction, AcceptsDateString) {
  // Mac Excel ja-JP NUMBERVALUE accepts date strings after the numeric
  // path fails (unlike Microsoft's docs). 2024-03-15 -> serial 45366.
  const Value v = EvalSource("=NUMBERVALUE(\"2024-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 45366.0);
}

TEST(NumberValueFunction, SameSeparatorRejected) {
  const Value v = EvalSource("=NUMBERVALUE(\"1.5\", \".\", \".\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(NumberValueFunction, EmptyDecimalSepRejected) {
  const Value v = EvalSource("=NUMBERVALUE(\"1.5\", \"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(NumberValueFunction, NumericPassthrough) {
  // First arg coerces to text via Formulon's shortest-double form, so this
  // succeeds.
  const Value v = EvalSource("=NUMBERVALUE(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
