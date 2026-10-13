// Unit tests for the x87 extended-precision emulation in `eval/x87_arith.h`
// and its use by the Windows-profile arithmetic operators. Expectations are
// bit-exact.

#include "eval/x87_arith.h"

#include "eval/coerce.h"
#include "eval/eval_profile_scope.h"
#include "eval/scalar_ops.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

struct TieCase {
  double a;
  double b;
  double ieee;
  double x87;
};

TEST(X87ArithTest, AddAndSubTies) {
  const TieCase cases[] = {
      {-0.9129955133469774, -5.879469383574229e-08, -0.9129955721416713, -0.9129955721416712},
      {1.2208398285616757, 0.00010733837723753141, 1.2209471669389134, 1.2209471669389131},
      {-0.4581922305687231, 1.2420342457268487e-12, -0.4581922305674811, -0.45819223056748104},
      {7.703547992360871, 0.00011844907489910741, 7.703666441435771, 7.70366644143577},
  };
  for (const TieCase& c : cases) {
    volatile double a = c.a;
    volatile double b = c.b;
    EXPECT_EQ(a + b, c.ieee);
    EXPECT_EQ(x87_add(c.a, c.b), c.x87);
    EXPECT_EQ(x87_sub(c.a, -c.b), c.x87);
  }
}

TEST(X87ArithTest, MulTies) {
  const TieCase cases[] = {
      {-0.3656043332756239, -1.0011108631918133, 0.3660104696722272, 0.3660104696722273},
      {-0.2747879133066195, -0.24452578313715034, 0.06719272969792452, 0.0671927296979245},
      {7.293599685543779, 3.120008013493573, 22.75608946611079, 22.756089466110794},
      {-1.0080076485166398, 0.9979718422059631, -1.005963249947852, -1.0059632499478521},
  };
  for (const TieCase& c : cases) {
    volatile double a = c.a;
    volatile double b = c.b;
    EXPECT_EQ(a * b, c.ieee);
    EXPECT_EQ(x87_mul(c.a, c.b), c.x87);
  }
}

TEST(X87ArithTest, DivTies) {
  const TieCase cases[] = {
      {-1.024374217800573, -0.8410527840589299, 1.2179666213777118, 1.2179666213777116},
      {5.321356791975804, -10.731808001387783, -0.49584904904072763, -0.4958490490407276},
      {-2.832347425338331, 0.1818382513899044, -15.57619149815242, -15.576191498152419},
      {-0.14368235354888226, -0.2265327376244943, 0.6342675017111802, 0.6342675017111801},
      {1.0, 1.9999999999999998, 0.5000000000000001, 0.5},
  };
  for (const TieCase& c : cases) {
    volatile double a = c.a;
    volatile double b = c.b;
    EXPECT_EQ(a / b, c.ieee);
    EXPECT_EQ(x87_div(c.a, c.b), c.x87);
  }
}

TEST(X87ArithTest, ExactResultsAreUnchanged) {
  EXPECT_EQ(x87_add(0.5, 0.25), 0.75);
  EXPECT_EQ(x87_mul(3.0, 7.0), 21.0);
  EXPECT_EQ(x87_div(1.0, 4.0), 0.25);
  EXPECT_EQ(x87_sub(1.0, 1.0), 0.0);
}

TEST(X87ArithTest, Exp) {
  EXPECT_EQ(x87_exp(230.25850929940455), 9.999999999999827e+99);
  EXPECT_EQ(x87_exp(230.25850929940458), 1.000000000000011e+100);
  EXPECT_EQ(x87_exp(230.25850929940452), 9.999999999999542e+99);
  EXPECT_EQ(x87_exp(-0.693147180559945), 0.5000000000000002);
  EXPECT_EQ(x87_exp(1.0), 2.718281828459045);
  EXPECT_EQ(x87_exp(-230.25850929940455), 1.0000000000000174e-100);
}

TEST(X87ArithTest, Ln) {
  struct Case {
    double x;
    double expected;
  };
  const Case cases[] = {
      {8, 2.0794415416798357}, {125, 4.8283137373023015}, {2, 0.6931471805599453},     {1E300, 690.7755278982137},
      {7, 1.9459101490553132}, {1000, 6.907755278982137}, {0.001, -6.907755278982137}, {10, 2.302585092994046},
      {3, 1.0986122886681098}, {32, 3.4657359027997265},  {27, 3.295836866004329},     {1E-300, -690.7755278982137},
  };
  for (const Case& c : cases) {
    EXPECT_EQ(x87_ln(c.x), c.expected) << c.x;
  }
}

TEST(X87ArithTest, WinHostPower) {
  struct Case {
    double x;
    double y;
    double expected;
  };
  const Case cases[] = {
      {8, 1.0 / 3, 1.9999999999999998},
      {8, -1.0 / 3, 0.5},
      {125, -1.0 / 3, 0.19999999999999996},
      {1e300, 1.0 / 3, 9.999999999999827e+99},
      {2, 0.5, 1.4142135623730951},
      {2, 1.0 / 3, 1.2599210498948732},
      {7, 1.0 / 3, 1.912931182772389},
      {1000, 1.0 / 3, 9.999999999999998},
      {0.001, 1.0 / 3, 0.10000000000000002},
      {10, 0.1, 1.2589254117941673},
      {3, 1.0 / 7, 1.169930812758687},
      {32, 0.2, 2.0},
      {2, 0.2, 1.148698354997035},
      {2, 1.5, 2.82842712474619},
      {1e-300, 1.0 / 3, 1.0000000000000174e-100},
      {27, 1.0 / 3, 2.9999999999999996},
  };
  ScopedEvalProfile scope(win_365_ja_jp_profile());
  for (const Case& c : cases) {
    auto r = apply_pow(c.x, c.y);
    ASSERT_TRUE(r.has_value()) << c.x << "^" << c.y;
    EXPECT_EQ(r.value(), c.expected) << c.x << "^" << c.y;
  }
}

TEST(X87ArithTest, ArithmeticFollowsHost) {
  {
    ScopedEvalProfile scope(win_365_ja_jp_profile());
    EXPECT_EQ(apply_arithmetic(parser::BinOp::Div, 1.0, 1.9999999999999998).as_number(), 0.5);
  }
  {
    ScopedEvalProfile scope(mac_365_ja_jp_profile());
    EXPECT_EQ(apply_arithmetic(parser::BinOp::Div, 1.0, 1.9999999999999998).as_number(), 0.5000000000000001);
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
