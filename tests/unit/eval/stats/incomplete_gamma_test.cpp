//
// Reference values for the incomplete gamma ratio and the Poisson density
// prefactor, across the shape/argument branches (small and large shape,
// integer and half-integer shape, series, continued fraction, asymptotic and
// uniform expansions).

#include "eval/stats/incomplete_gamma.h"

#include <cmath>

#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace stats {
namespace {

struct RatioPoint {
  double a;
  double x;
  bool upper;
  double value;
};

const RatioPoint kRatioPoints[] = {
    {0.3, 15.25, false, 0.9999999886654591},      {0.3, 23.6875, false, 0.9999999999981701},
    {0.5, 4.21875, false, 0.9963243878822587},    {0.75, 30.03125, false, 0.9999999999999687},
    {0.75, 15.0, false, 0.9999998751140187},      {1.0, 30.125, false, 0.9999999999999174},
    {1.5, 25.4375, false, 0.9999999999479855},    {1.5, 14.84375, false, 0.9999983944617273},
    {2.5, 24.96875, false, 0.9999999985727879},   {3.7, 4.125, false, 0.6501515940662252},
    {3.7, 2.75, false, 0.3572832271497379},       {7.2, 2.0, false, 0.003445808302479514},
    {12.5, 30.28125, false, 0.9999125745245394},  {12.5, 27.34375, false, 0.9994592616579983},
    {25.5, 28.46875, false, 0.7362982937597935},  {45.2, 23.40625, false, 4.1420558189526576e-05},
    {45.2, 2.3125, false, 1.161068715979077e-41}, {150.3, 35.0, false, 3.7743237691254834e-47},
    {20.5, 10.03125, true, 0.9975606635882226},   {20.5, 5.46875, true, 0.9999993202096574},
    {22.5, 37.75, true, 0.002953524806645362},    {30.5, 5.5625, true, 0.9999999999998297},
    {30.5, 31.25, true, 0.422701196981119},       {50.0, 52.0, true, 0.37214970157386756},
    {100.5, 40.0, true, 0.9999999999999992},      {100.5, 66.0, true, 0.9999529561780236},
    {250.0, 175.0, true, 0.9999999426125555},     {0.5, 1.4453125, true, 0.08909635592130864},
    {1.0, 10.59375, true, 2.507222244856129e-05}, {1.5, 8.9375, true, 0.0004667454848969492},
    {2.0, 7.5546875, true, 0.004479660029670486}, {2.5, 9.9140625, true, 0.0013460140864643606},
    {3.0, 2.46875, true, 0.5518539095634257},     {3.5, 34.25, true, 2.9675396444978065e-12},
    {5.0, 6.3203125, true, 0.24446453643415025},  {7.5, 9.359375, true, 0.22676282986357815},
    {15.0, 5.5625, true, 0.9993301551058015},     {0.5, 10.875, false, 0.999996894101332},
    {1.0, 11.53125, false, 0.9999901815766159},   {1.5, 12.09375, false, 0.9999771733878133},
    {2.0, 14.734375, false, 0.9999937224274613},  {2.5, 12.359375, false, 0.9998421204274461},
    {3.0, 3.2890625, false, 0.6383722372608964},  {3.5, 8.6875, false, 0.984868263184437},
    {5.0, 7.609375, false, 0.8757121788606055},   {7.5, 14.0, false, 0.9784309660267487},
    {15.0, 26.875, false, 0.9950986610799109},
};

struct DensityPoint {
  double a;
  double x;
  double value;
};

const DensityPoint kDensityPoints[] = {
    {0.3, 16.53125, 3.1035327883892773e-09}, {0.5, 19.28125, 5.433966936703543e-10},
    {0.75, 39.5, 2.280000254261986e-18},     {1.0, 108.0, 1.2479464629129515e-47},
    {1.5, 67.0, 7.375002521750191e-29},      {2.5, 26.09375, 4.664412052266333e-10},
    {3.7, 14.875, 0.0001217024948963807},    {7.2, 1.84375, 0.006688079441663087},
    {12.5, 140.5, 3.4968682901193297e-45},   {25.5, 20.90625, 0.05995039617111003},
    {45.2, 106.0, 5.3559448615515165e-12},   {150.3, 7.1875, 3.406098293950535e-137},
};

struct BinomialPoint {
  double n;
  double p;
  double k;
  double value;
};

const BinomialPoint kBinomialPoints[] = {
    {10.0, 0.3, 9.0, 0.0001377809999999999},       {50.0, 0.5, 47.0, 1.7408297026122522e-11},
    {200.0, 0.02, 176.0, 3.7882218378969775e-269}, {400.0, 0.9, 373.0, 0.005771563995077984},
    {1000.0, 0.37, 14.0, 1.3319799761792194e-173}, {25.0, 0.001, 24.0, 2.4975000000000673e-71},
};

TEST(GammaRatioExcel, MatchesExcelValues) {
  for (const RatioPoint& p : kRatioPoints) {
    EXPECT_EQ(gamma_ratio_excel(p.a, p.x, p.upper), p.value) << p.a << " " << p.x;
  }
}

TEST(GammaRatioExcel, DensityMatchesExcelValues) {
  for (const DensityPoint& p : kDensityPoints) {
    EXPECT_EQ(dpois_raw(p.x, p.a) * (p.a / p.x), p.value) << p.a << " " << p.x;
  }
}

TEST(GammaRatioExcel, BinomialDensityMatchesExcelValues) {
  for (const BinomialPoint& b : kBinomialPoints) {
    EXPECT_EQ(dbinom_raw(b.n, b.p, b.k), b.value) << b.n << " " << b.p << " " << b.k;
  }
}

TEST(GammaRatioExcel, Boundaries) {
  EXPECT_EQ(gamma_ratio_excel(2.0, 0.0, false), 0.0);
  EXPECT_EQ(gamma_ratio_excel(2.0, 0.0, true), 1.0);
  EXPECT_EQ(gamma_ratio_excel(0.0, 3.0, false), 1.0);
  EXPECT_EQ(gamma_ratio_excel(0.0, 3.0, true), 0.0);
  EXPECT_NEAR(gamma_ratio_excel(3.5, 2.0, false) + gamma_ratio_excel(3.5, 2.0, true), 1.0, 1e-15);
}

TEST(GammaRatioExcel, LoaderHelpers) {
  EXPECT_EQ(log1pmx(0.0), 0.0);
  EXPECT_NEAR(log1pmx(0.5), std::log(1.5) - 0.5, 1e-15);
  EXPECT_NEAR(lgamma_excel(10.5), std::lgamma(10.5), 1e-13);
  EXPECT_NEAR(lgamma_excel(0.3), std::lgamma(0.3), 1e-14);
  EXPECT_NEAR(lgamma_excel(2.7), std::lgamma(2.7), 1e-14);
  EXPECT_NEAR(stirlerr(20.0), std::lgamma(21.0) - (20.5 * std::log(20.0) - 20.0 + 0.9189385332046728), 1e-14);
  EXPECT_NEAR(bd0(3.0, 2.0), 3.0 * std::log(1.5) + 2.0 - 3.0, 1e-14);
}

}  // namespace
}  // namespace stats
}  // namespace eval
}  // namespace formulon
