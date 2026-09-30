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

}  // namespace
}  // namespace formulon
