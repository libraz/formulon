//
// utils/elementary_math pinned bit for bit. sin and cos: each reduction path and quadrant. atan: each range and range
// edge. tan: each of the three argument reductions (none below pi/4, the three-part pi/2 below 2^19, the bits of 2/pi
// above) and both quadrant parities; TAN and COT evaluate through it.

#include "utils/elementary_math.h"

#include <cmath>
#include <limits>
#include <string_view>

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "value.h"

namespace formulon {
namespace {

struct Pin {
  double x;
  double tan;
};

constexpr Pin kPins[] = {
    {0.5, 0.5463024898437905},                        // no reduction
    {-0.785, -0.9992039901050427},                    // no reduction
    {1e-09, 1e-09},                                   // no reduction
    {1.0536712127723509e-08, 1.053671212772351e-08},  // no reduction
    {1.0, 1.557407724654902},                         // three-part pi/2, odd quadrant
    {-1.0, -1.557407724654902},                       // three-part pi/2, odd quadrant
    {2.5, -0.7470222972386602},                       // three-part pi/2, even quadrant
    {3.0, -0.1425465430742778},                       // three-part pi/2, even quadrant
    {4.0, 1.1578212823495775},                        // three-part pi/2, odd quadrant
    {-5.5, 0.995584052213885},                        // three-part pi/2, even quadrant
    {100.0, -0.5872139151569291},                     // three-part pi/2, even quadrant
    {-12345.678, 0.9914971407432148},                 // three-part pi/2, even quadrant
    {524287.5, -0.34429901008431474},                 // three-part pi/2, even quadrant
    {524289.0, 2.349594260882002},                    // bits of 2/pi, odd quadrant
    {-1000000.0, 0.373624453987599},                  // bits of 2/pi, even quadrant
    {134217727.0, 3.2584564256367834},                // bits of 2/pi, odd quadrant
    {123456789.0, 7.059154426802704},                 // bits of 2/pi, odd quadrant
    {1e15, -1.672414782127583},                       // bits of 2/pi, odd quadrant
    {-7e22, -0.5324221632283528},                     // bits of 2/pi, even quadrant
    {1e300, 1.4214488238747243},                      // bits of 2/pi, even quadrant
};

TEST(ElementaryTan, MatchesPinnedValuesOnEveryReductionPath) {
  for (const Pin& p : kPins) {
    EXPECT_EQ(elementary::tan(p.x), p.tan) << p.x;
  }
}

TEST(ElementaryTan, KeepsZeroAndOddSymmetry) {
  EXPECT_TRUE(std::signbit(elementary::tan(-0.0)));
  EXPECT_EQ(elementary::tan(0.0), 0.0);
  for (const Pin& p : kPins) {
    EXPECT_EQ(elementary::tan(-p.x), -p.tan) << p.x;
  }
}

double Number(std::string_view src) {
  const Value v = formulon::test::EvalSource(src);
  EXPECT_TRUE(v.is_number()) << src;
  return v.is_number() ? v.as_number() : 0.0;
}

TEST(ElementaryTan, TanAndCotEvaluateThroughTheKernel) {
  EXPECT_EQ(Number("=TAN(TRUE)"), 1.557407724654902);
  EXPECT_EQ(Number("=COT(TRUE)"), 0.6420926159343308);
  EXPECT_EQ(Number("=TAN(134217727)"), 3.2584564256367834);
  EXPECT_EQ(Number("=COT(134217727)"), 0.30689377710631044);
}

struct SinCosPin {
  double x;
  double sin;
  double cos;
};

// sin and cos on each reduction path and in all four quadrants.
constexpr SinCosPin kSinCosPins[] = {
    {0.5, 0.479425538604203, 0.8775825618903728},                   // none, quadrant 0
    {-0.785, -0.706825181105366, 0.7073882691671998},               // none, quadrant 0
    {1e-09, 1e-09, 1.0},                                            // none, quadrant 0
    {0.7853981633974483, 0.7071067811865475, 0.7071067811865476},   // none, quadrant 0
    {1.0, 0.8414709848078965, 0.5403023058681398},                  // three-part pi/2, quadrant 1
    {-1.0, -0.8414709848078965, 0.5403023058681398},                // three-part pi/2, quadrant 3
    {2.5, 0.5984721441039564, -0.8011436155469337},                 // three-part pi/2, quadrant 2
    {3.0, 0.1411200080598672, -0.9899924966004454},                 // three-part pi/2, quadrant 2
    {4.0, -0.7568024953079283, -0.6536436208636119},                // three-part pi/2, quadrant 3
    {-5.5, 0.7055403255703919, 0.70866977429126},                   // three-part pi/2, quadrant 0
    {100.0, -0.5063656411097588, 0.8623188722876839},               // three-part pi/2, quadrant 0
    {-12345.678, 0.7040813137533816, 0.7101193587160628},           // three-part pi/2, quadrant 0
    {524287.5, -0.32554397588580297, 0.9455269006032899},           // three-part pi/2, quadrant 0
    {524289.0, 0.9201302731771119, 0.3916124109154527},             // bits of 2/pi, quadrant 1
    {-1000000.0, 0.34999350217129294, 0.9367521275331447},          // bits of 2/pi, quadrant 0
    {134217727.5, -0.9796209680098831, 0.2008550697278502},         // bits of 2/pi, quadrant 3
    {123456789.0, 0.9901147518020355, 0.14025968153390964},         // bits of 2/pi, quadrant 1
    {1000000000000000.0, 0.8582727931702359, -0.5131937377869702},  // bits of 2/pi, quadrant 1
    {-7e+22, -0.4699619490293273, 0.8826866751370817},              // bits of 2/pi, quadrant 0
    {1e+300, -0.8178819121159085, -0.5753861119575491},             // bits of 2/pi, quadrant 0
};

TEST(ElementarySinCos, MatchesPinnedValuesOnEveryReductionPath) {
  for (const SinCosPin& p : kSinCosPins) {
    EXPECT_EQ(elementary::sin(p.x), p.sin) << p.x;
    EXPECT_EQ(elementary::cos(p.x), p.cos) << p.x;
  }
}

TEST(ElementarySinCos, SineIsOddAndCosineEven) {
  for (const SinCosPin& p : kSinCosPins) {
    EXPECT_EQ(elementary::sin(-p.x), -p.sin) << p.x;
    EXPECT_EQ(elementary::cos(-p.x), p.cos) << p.x;
  }
  EXPECT_EQ(elementary::cos(0.0), 1.0);
}

struct AtanPin {
  double x;
  double atan;
};

// atan in each of its ranges and on the range edges.
constexpr AtanPin kAtanPins[] = {
    {1e-09, 1e-09},
    {0.3, 0.2914567944778671},
    {0.5, 0.46364760900080615},
    {0.6, 0.5404195002705842},
    {0.75, 0.6435011087932844},
    {0.8, 0.6747409422235527},
    {1.0, 0.7853981633974483},
    {1.2, 0.8760580505981934},
    {1.3333333333333333, 0.9272952180016122},
    {1.5, 0.982793723247329},
    {1.6666666666666667, 1.0303768265243125},
    {1.9, 1.0863183977578734},
    {2.0, 1.1071487177940906},
    {2.5, 1.1902899496825317},
    {100.0, 1.5607966601082315},
    {10000000000.0, 1.5707963266948965},
    {1e+300, 1.5707963267948966},
};

TEST(ElementaryAtan, MatchesPinnedValuesInEveryRange) {
  for (const AtanPin& p : kAtanPins) {
    EXPECT_EQ(elementary::atan(p.x), p.atan) << p.x;
    EXPECT_EQ(elementary::atan(-p.x), -p.atan) << p.x;
  }
  EXPECT_EQ(elementary::atan(std::numeric_limits<double>::infinity()), 1.5707963267948966);
}

struct ExpPin {
  double x;
  double exp;
};

// exp across the normal range, near overflow and into subnormal results.
constexpr ExpPin kExpPins[] = {
    {-745.2, 0.0},
    {-720.0, 2.0322308024e-313},
    {-1.0, 0.36787944117144233},
    {1e-10, 1.0000000001},
    {0.5, 1.6487212707001282},
    {1.0, 2.718281828459045},
    {100.0, 2.6881171418161356e+43},
    {709.78, 1.7928227943945155e+308},
};

struct LogPin {
  double x;
  double log;
  double log10;
};

// log and log10 on subnormal and normal arguments, including the exponent-0 path near 1.
constexpr LogPin kLogPins[] = {
    {5e-324, -744.4400719213812, -323.3062153431158},
    {1e-310, -713.8013788281542, -310.0},
    {0.5, -0.6931471805599453, -0.3010299956639812},
    {0.75, -0.2876820724517809, -0.12493873660829995},
    {1.0, 0.0, 0.0},
    {1.3, 0.26236426446749106, 0.11394335230683679},
    {2.0, 0.6931471805599453, 0.3010299956639812},
    {1e+300, 690.7755278982137, 300.0},
};

struct PowPin {
  double x;
  double y;
  double pow;
};

// pow on exact powers of two and ten, a near-1 base with a large exponent, a negative base and tiny exponents.
constexpr PowPin kPowPins[] = {
    {2.0, 0.5, 1.4142135623730951},
    {10.0, -5.0, 1e-05},
    {10.0, 22.0, 1e+22},
    {10.0, -308.0, 1e-308},
    {1.0001, 100000.0, 22015.456048527954},
    {-2.0, 3.0, -8.0},
    {0.5, -1074.0, std::numeric_limits<double>::infinity()},
    {2.5, 1e-70, 1.0},
    {7.0, 0.3333333333333333, 1.912931182772389},
    {1e-300, -1.5, std::numeric_limits<double>::infinity()},
};

TEST(ElementaryExpLog, MatchesPinnedValues) {
  for (const ExpPin& p : kExpPins) {
    EXPECT_EQ(elementary::exp(p.x), p.exp) << p.x;
  }
  for (const LogPin& p : kLogPins) {
    EXPECT_EQ(elementary::log(p.x), p.log) << p.x;
    EXPECT_EQ(elementary::log10(p.x), p.log10) << p.x;
  }
  for (const PowPin& p : kPowPins) {
    EXPECT_EQ(elementary::pow(p.x, p.y), p.pow) << p.x << "^" << p.y;
  }
}

// log10 arguments whose result moves by one ulp when the kernel's multiply-adds are contracted.
struct Log10Pin {
  double x;
  double log10;
};

constexpr Log10Pin kLog10ContractionPins[] = {
    {689.875, 2.83877040707332},   {518.203, 2.7144999228907833}, {34.6829, 1.5401154037217513},
    {402.459, 2.6047216437580167}, {8.7077, 0.9399034581976156},  {276.38, 2.4415066124941145},
    {17.2043, 1.2356370069585836},
};

TEST(ElementaryExpLog, Log10IsNotContracted) {
  for (const Log10Pin& p : kLog10ContractionPins) {
    EXPECT_EQ(elementary::log10(p.x), p.log10) << p.x;
  }
}

TEST(ElementaryExpLog, SpecialArguments) {
  EXPECT_EQ(elementary::log(0.0), -std::numeric_limits<double>::infinity());
  EXPECT_TRUE(std::isnan(elementary::log(-1.0)));
  EXPECT_EQ(elementary::exp(-1000.0), 0.0);
  EXPECT_EQ(elementary::exp(1000.0), std::numeric_limits<double>::infinity());
  EXPECT_TRUE(std::isnan(elementary::pow(-8.0, 1.0 / 3.0)));
  EXPECT_EQ(elementary::pow(0.0, -1.0), std::numeric_limits<double>::infinity());
  EXPECT_EQ(elementary::pow(-0.0, -3.0), -std::numeric_limits<double>::infinity());
  EXPECT_EQ(elementary::pow(1.0, std::numeric_limits<double>::quiet_NaN()), 1.0);
}

}  // namespace
}  // namespace formulon
