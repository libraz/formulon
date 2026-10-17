//
// Unit tests for `formulon::parse_double_prefix` / `parse_double_exact`.
//
// On the decimal grammar the parser must agree with the C library's
// `strtod` bit for bit, in the value and in the consumed length. The tests
// compare against the host `strtod` over boundary values, every combination
// of the grammar's parts, and generated inputs. The deliberate differences
// (hexadecimal floats, inf / nan spellings) are pinned separately.

#include "utils/double_parse.h"

#include <algorithm>
#include <cfloat>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <random>
#include <string>
#include <vector>

#include "gtest/gtest.h"

namespace formulon {
namespace {

std::uint64_t Bits(double v) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof(bits));
  return bits;
}

// Asserts that the parser and the host strtod read `text` identically.
void ExpectMatchesStrtod(const std::string& text) {
  char* end = nullptr;
  const double expected = std::strtod(text.c_str(), &end);
  const auto expected_consumed = static_cast<std::size_t>(end - text.c_str());
  const ParsedDouble parsed = parse_double_prefix(text);
  ASSERT_EQ(parsed.consumed, expected_consumed) << "text: " << text;
  ASSERT_EQ(Bits(parsed.value), Bits(expected)) << "text: " << text;
  const std::size_t mantissa_end = std::min(expected_consumed, text.find_first_of("eE"));
  const bool nonzero_digits = text.find_first_of("123456789") < mantissa_end;
  const bool expected_out_of_range = nonzero_digits && (std::isinf(expected) || std::fabs(expected) < DBL_MIN);
  ASSERT_EQ(parsed.out_of_range, expected_out_of_range) << "text: " << text;
}

TEST(DoubleParse, BoundaryValuesMatchStrtod) {
  for (const char* text : {
           "0",
           "-0",
           "+0",
           "0.0",
           "00000",
           "0e0",
           "0e999",
           "-0e-999",
           "1",
           "-1",
           "+1",
           ".5",
           "5.",
           "-.5",
           "+5.",
           "0.1",
           "0.2",
           "0.3",
           "1.005",
           "123456789012345678",
           "9007199254740993",
           "9007199254740992.5",
           "1.7976931348623157e308",
           "1.7976931348623158e308",
           "1.797693134862315807e308",
           "1.7976931348623159e308",
           "1e308",
           "1e309",
           "-1e309",
           "2.2250738585072014e-308",
           "2.2250738585072011e-308",
           "2.2250738585072009e-308",
           "4.9406564584124654e-324",
           "2.4703282292062327e-324",
           "2.4703282292062328e-324",
           "1e-320",
           "1e-400",
           "-1e-400",
           "1e2147483647",
           "1e-2147483648",
           "1e99999999999",
           "0.000000000000000000000000000001e30",
           "100000000000000000000000000000000000000e-38",
           "1E5",
           "1e+5",
           "1e-5",
           "1e",
           "1e+",
           "1e-",
           "1ex",
           "1.2.3",
           "1..2",
           "..5",
           ".",
           "-",
           "+",
           "",
           "e5",
           "-e5",
           " 1",
           "\t\n\v\f\r 1",
           " -1.5e3xyz",
           "12abc",
           "1,5",
           "1_000",
           "0.0000000000000000000000000000000000000000000000000000001",
       }) {
    ExpectMatchesStrtod(text);
  }
}

TEST(DoubleParse, HalfwayCasesRoundToEven) {
  // 2^53 + 1 and 2^53 + 3 sit exactly between two doubles.
  ExpectMatchesStrtod("9007199254740993");
  ExpectMatchesStrtod("9007199254740995");
  // A nonzero digit past the 780th significant digit lifts an exact halfway
  // value, so it rounds up; dropping it would round to even instead.
  const std::string zeros(800, '0');
  ExpectMatchesStrtod("9007199254740993." + zeros + "1");
  ExpectMatchesStrtod("9007199254740993" + zeros + "1e-801");
  ExpectMatchesStrtod("-9007199254740993." + zeros + "1");
  // The halfway point between DBL_MAX and the next power of two overflows.
  ExpectMatchesStrtod("179769313486231580793728971405301e276");
  // A 767-digit halfway value one digit away from rounding the other way.
  std::string halfway =
      "2.4703282292062327208828439643411068618252990130716238221279284125033775363510437593264991818081799618";
  halfway += std::string(660, '0');
  halfway += "e-324";
  ExpectMatchesStrtod(halfway);
}

TEST(DoubleParse, LongMantissasMatchStrtod) {
  // Past 780 significant digits the tail is folded into a sticky digit.
  for (const std::size_t length : {779U, 780U, 781U, 800U, 2000U}) {
    for (const char tail : {'0', '1', '5', '9'}) {
      std::string text = "1";
      text += std::string(length - 2, '0');
      text.push_back(tail);
      ExpectMatchesStrtod(text + "e-" + std::to_string(length + 300));
      ExpectMatchesStrtod("0." + text);
      ExpectMatchesStrtod(text + "." + text);
    }
  }
}

TEST(DoubleParse, EveryGrammarShapeMatchesStrtod) {
  // Full product of the grammar's parts; small enough to enumerate outright.
  const std::vector<std::string> spaces = {"", " ", "\t"};
  const std::vector<std::string> signs = {"", "+", "-"};
  const std::vector<std::string> integers = {"", "0", "7", "00012", "123456789012345678901234567890"};
  const std::vector<std::string> fractions = {"", ".", ".0", ".25", ".000001", ".12345678901234567890123"};
  const std::vector<std::string> exponents = {"", "e", "E+", "e7", "E-7", "e+308", "e-330", "e0400", "e-"};
  const std::vector<std::string> tails = {"", "x", " ", "e5"};
  for (const auto& space : spaces) {
    for (const auto& sign : signs) {
      for (const auto& integer : integers) {
        for (const auto& fraction : fractions) {
          for (const auto& exponent : exponents) {
            for (const auto& tail : tails) {
              ExpectMatchesStrtod(space + sign + integer + fraction + exponent + tail);
            }
          }
        }
      }
    }
  }
}

TEST(DoubleParse, GeneratedDecimalsMatchStrtod) {
  std::mt19937_64 rng(0x5eed5eedULL);
  std::uniform_int_distribution<int> digit(0, 9);
  std::uniform_int_distribution<int> length(1, 40);
  std::uniform_int_distribution<int> exponent(-360, 330);
  std::uniform_int_distribution<int> coin(0, 3);
  for (int n = 0; n < 50000; ++n) {
    std::string text;
    if (coin(rng) == 0) {
      text.push_back('-');
    }
    const int int_len = coin(rng) == 0 ? 0 : length(rng);
    for (int i = 0; i < int_len; ++i) {
      text.push_back(static_cast<char>('0' + digit(rng)));
    }
    if (int_len == 0 || coin(rng) != 0) {
      text.push_back('.');
      const int frac_len = length(rng);
      for (int i = 0; i < frac_len; ++i) {
        text.push_back(static_cast<char>('0' + digit(rng)));
      }
    }
    if (coin(rng) != 0) {
      text += "e" + std::to_string(exponent(rng));
    }
    ExpectMatchesStrtod(text);
  }
}

TEST(DoubleParse, GeneratedRoundTripsOfEveryBitPattern) {
  // Shortest round-trip spellings of random doubles, including subnormals.
  std::mt19937_64 rng(0xd0b1eULL);
  char buf[64];
  for (int n = 0; n < 20000; ++n) {
    std::uint64_t bits = rng();
    double v = 0.0;
    std::memcpy(&v, &bits, sizeof(v));
    if (!std::isfinite(v)) {
      continue;
    }
    std::snprintf(buf, sizeof(buf), "%.17g", v);
    ExpectMatchesStrtod(buf);
    std::snprintf(buf, sizeof(buf), "%.16e", v);
    ExpectMatchesStrtod(buf);
  }
}

TEST(DoubleParse, HexAndInfNanSpellingsAreNotNumbers) {
  // strtod reads these; Excel treats all of them as text (#VALUE! when
  // coerced), so the parser stops where the decimal grammar does.
  EXPECT_EQ(parse_double_prefix("0x10").consumed, 1U);
  EXPECT_EQ(parse_double_prefix("-0X1A").consumed, 2U);
  EXPECT_EQ(parse_double_prefix("0x1p3").consumed, 1U);
  for (const char* text : {"inf", "INF", "-inf", "infinity", "nan", "NaN", "nan(0x1)", " inf", "x1"}) {
    const ParsedDouble parsed = parse_double_prefix(text);
    EXPECT_EQ(parsed.consumed, 0U) << text;
    EXPECT_EQ(parsed.value, 0.0) << text;
    EXPECT_FALSE(parsed.out_of_range) << text;
  }
}

TEST(DoubleParse, ExactRequiresTheWholeText) {
  double out = -1.0;
  EXPECT_TRUE(parse_double_exact("1.5", &out));
  EXPECT_EQ(out, 1.5);
  EXPECT_TRUE(parse_double_exact(" 2", &out));
  EXPECT_EQ(out, 2.0);
  out = -1.0;
  EXPECT_FALSE(parse_double_exact("", &out));
  EXPECT_FALSE(parse_double_exact("1,5", &out));
  EXPECT_FALSE(parse_double_exact("1 ", &out));
  EXPECT_FALSE(parse_double_exact("1e", &out));
  EXPECT_FALSE(parse_double_exact("0x10", &out));
  EXPECT_EQ(out, -1.0);
  EXPECT_TRUE(parse_double_exact("1e999", &out));
  EXPECT_TRUE(std::isinf(out));
}

TEST(DoubleParse, ReadsOnlyTheViewItIsGiven) {
  // No NUL terminator is needed: the digits past the view are not read.
  const std::string backing = "12345";
  const ParsedDouble parsed = parse_double_prefix(std::string_view(backing).substr(0, 3));
  EXPECT_EQ(parsed.consumed, 3U);
  EXPECT_EQ(parsed.value, 123.0);
}

TEST(ExcelDecimal, MatchesExcelValues) {
  struct Case {
    const char* text;
    double expected;
  };
  const Case cases[] = {
      {"3.35404999999997", 3.3540499999999698},
      {"37131855.1", 37131855.1},
      {"697959.037122062", 697959.0371220619},
      {"1.0307514780773E-16", 1.0307514780773e-16},
      {"9.74985033627417E-25", 9.74985033627417e-25},
      {"69628.8437749113", 69628.8437749113},
      {"17177295334.7191", 17177295334.7191},
      {"746090443.926817", 746090443.926817},
      {"6.759420838998E+48", 6.759420838998e+48},
      {"0.01489433427931", 0.01489433427931},
      {"0.0107954085429028", 0.0107954085429028},
      {"984999686.6074", 984999686.6074},
      {"11261435778.9136", 11261435778.9136},
      {"1.65212142008856", 1.65212142008856},
      {"10000000000", 10000000000.0},
      {"2.24525923513E-15", 2.24525923513e-15},
      {"78.662361911501", 78.662361911501},
      {"6313011.2652404", 6313011.2652404},
      {"9.89354736655", 9.89354736655},
      {"7.63641851158249E-16", 7.63641851158249e-16},
      {"1.30937622098065E-12", 1.30937622098065e-12},
      {"0.0003806309377209", 0.0003806309377209},
      {"0.00000869396673210", 8.6939667321e-06},
      {"3.1381947382267E+54", 3.1381947382267003e+54},
      {"1.59236848910830E+3", 1592.3684891083},
      {"0.0003018059348", 0.0003018059348},
      {"8.01519121728257E+53", 8.01519121728257e+53},
      {"5.77103972365227E-21", 5.77103972365227e-21},
      {"9.0432937918753E-24", 9.0432937918753e-24},
      {"9.88156248108836", 9.88156248108836},
      {"0.0815376397103211", 0.0815376397103211},
      {"5.55167543372546E-308", 5.55167543372546e-308},
      {"0.00000600156152566619", 6.00156152566619e-06},
      {"9.50281034636159E-13", 9.50281034636159e-13},
      {"961787.972791237", 961787.972791237},
      {"999.618842087830", 999.61884208783},
      {"0.00000008717204495052266", 8.71720449505226e-08},
      {"0.00000140829246479952", 1.40829246479952e-06},
      {"60463834.73965609", 60463834.739656},
      {"69001.1774372", 69001.1774372},
      {"15298.13051894", 15298.13051894},
      {"0.333333333333333333", 0.333333333333333},
      {"1.797693134862316E308", 1.79769313486231e+308},
  };
  for (const Case& c : cases) {
    double value = 0.0;
    ASSERT_TRUE(parse_excel_decimal(c.text, &value)) << c.text;
    EXPECT_EQ(value, c.expected) << c.text;
  }
}

TEST(ExcelDecimal, ScaleRangeLimits) {
  EXPECT_EQ(excel_scale_decimal(0U, 400), 0.0);
  EXPECT_EQ(excel_scale_decimal(1U, 400), HUGE_VAL);
  EXPECT_EQ(excel_scale_decimal(1U, 100000), HUGE_VAL);
  EXPECT_EQ(excel_scale_decimal(1U, -400), 0.0);
  EXPECT_EQ(excel_scale_decimal(1U, -100000), 0.0);
  EXPECT_EQ(excel_scale_decimal(1U, -310), 0.0);
  EXPECT_EQ(excel_scale_decimal(123456789012345U, 0), 123456789012345.0);
}

TEST(ExcelDecimal, RejectsMalformedText) {
  double value = 0.0;
  EXPECT_FALSE(parse_excel_decimal("", &value));
  EXPECT_FALSE(parse_excel_decimal("1.5x", &value));
  EXPECT_FALSE(parse_excel_decimal("e5", &value));
}

TEST(ExcelDecimal, FileParserStaysCorrectlyRounded) {
  double value = 0.0;
  ASSERT_TRUE(parse_double_exact("3.35404999999997", &value));
  double excel = 0.0;
  ASSERT_TRUE(parse_excel_decimal("3.35404999999997", &excel));
  EXPECT_LT(excel, value);
}

}  // namespace
}  // namespace formulon
