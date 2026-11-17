//
// utils/elementary_math pinned bit for bit. tan: each of the three argument reductions (none below pi/4, the
// three-part pi/2 below 2^19, the bits of 2/pi above) and both quadrant parities; TAN and COT evaluate through it.

#include "utils/elementary_math.h"

#include <cmath>
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

}  // namespace
}  // namespace formulon
