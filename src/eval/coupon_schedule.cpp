//
// Implementation of the shared coupon-schedule engine declared in
// `coupon_schedule.h`. See that header for the semantic contract and
// the list of callers that will reuse this module.

#include "eval/coupon_schedule.h"

#include <cmath>
#include <cstdint>

#include "eval/date_time.h"

namespace formulon {
namespace eval {
namespace {

// Calendar day-count helpers shared with the COUP* engine and the date
// builtins; see `eval/date_time.h`.
using date_time::basis_days_between;
using date_time::days_in_month;

// Constructs the serial for (y, m, anchor_day), clamping the day to
// the target month's last day when shorter. This implements Excel's
// coupon-date day-of-month preservation rule: if the maturity is
// Aug-31 and we step back three months, we land on May-31; if we step
// back six, we land on Feb-28 (or Feb-29 in a leap year).
double coupon_serial(int y, unsigned m, unsigned anchor_day) noexcept {
  const unsigned last = days_in_month(y, m);
  const unsigned d = anchor_day > last ? last : anchor_day;
  return date_time::serial_from_ymd(y, m, d);
}

// Shifts (y, m) backward by `months` (always positive). Uses
// floor-division so negative month remainders wrap correctly.
void shift_months_back(int& y, unsigned& m, int months) noexcept {
  long long mm0 = static_cast<long long>(m) - 1 - static_cast<long long>(months);
  long long year_shift = mm0 / 12;
  long long rem = mm0 % 12;
  if (rem < 0) {
    rem += 12;
    year_shift -= 1;
  }
  y += static_cast<int>(year_shift);
  m = static_cast<unsigned>(rem + 1);
}

}  // namespace

bool compute_coupon_dates(double settlement, double maturity, int frequency, int basis, CouponDates* out) noexcept {
  if (out == nullptr) {
    return false;
  }
  const double s = std::trunc(settlement);
  const double m = std::trunc(maturity);

  // Decompose maturity into (y, m, d). The anchor day-of-month is
  // preserved across the backward walk and clamped to each target
  // month's last day as needed.
  const date_time::YMD mat = date_time::ymd_from_serial(m);
  const unsigned anchor_day = mat.d;
  const int step_months = 12 / frequency;

  // Walk backwards from maturity in `step_months` increments until
  // the candidate is `<= settlement`. The "one step forward" from
  // that candidate is the NCD; the candidate itself is the PCD.
  //
  // We also count how many coupon dates lie strictly after settlement
  // and up to maturity — this is the loop-count of full step iterations
  // we performed before reaching PCD.
  int y_walk = mat.y;
  unsigned m_walk = mat.m;
  double current = date_time::serial_from_ymd(y_walk, m_walk, anchor_day);
  std::int32_t coupons = 0;

  // If maturity's own serial is already <= settlement (shouldn't
  // happen given the caller's pre-validation of settlement < maturity,
  // but guard defensively), there is one coupon remaining.
  while (current > s) {
    // Record that `current` is a coupon strictly after settlement
    // and <= maturity.
    ++coupons;
    // Step one period back.
    shift_months_back(y_walk, m_walk, step_months);
    current = coupon_serial(y_walk, m_walk, anchor_day);
  }

  const double pcd = current;
  // NCD is one step forward from PCD.
  int y_ncd = y_walk;
  unsigned m_ncd = m_walk;
  // Forward step by step_months using the same floor-mod logic.
  {
    long long mm0 = static_cast<long long>(m_ncd) - 1 + static_cast<long long>(step_months);
    long long year_shift = mm0 / 12;
    long long rem = mm0 % 12;
    if (rem < 0) {
      rem += 12;
      year_shift -= 1;
    }
    y_ncd += static_cast<int>(year_shift);
    m_ncd = static_cast<unsigned>(rem + 1);
  }
  const double ncd = coupon_serial(y_ncd, m_ncd, anchor_day);

  // Basis-adjusted day counts. `days_bs` = settlement - PCD;
  // `days_nc` = NCD - settlement. Rounded to match Excel's
  // integer day-count output.
  const double days_bs_raw = basis_days_between(pcd, s, basis);
  const double days_nc_raw = basis_days_between(s, ncd, basis);

  // Period length. Bases 0/4 always use 360/freq; basis 2 uses
  // 360/freq; basis 3 uses 365/freq; basis 1 uses the actual NCD-PCD
  // gap (integer, since both are date serials).
  double period_days = 0.0;
  switch (basis) {
    case 0:
    case 2:
    case 4:
      period_days = 360.0 / static_cast<double>(frequency);
      break;
    case 3:
      period_days = 365.0 / static_cast<double>(frequency);
      break;
    case 1:
      period_days = ncd - pcd;
      break;
    default:
      return false;
  }

  out->pcd = pcd;
  out->ncd = ncd;
  out->days_bs = std::round(days_bs_raw);
  out->days_nc = std::round(days_nc_raw);
  out->period_days = period_days;
  out->coupons_remaining = coupons;
  // `bond_dsc` is the effective DSC for bond pricing. For bases 0/1/4
  // the raw `days_nc` already satisfies `days_bs + days_nc ==
  // period_days`, so `bond_dsc == days_nc`. For bases 2 / 3 the raw
  // counts use actual days but `period_days` uses the nominal
  // `360/freq` or `365/freq`, so the identity breaks; we re-derive
  // `bond_dsc = period_days - days_bs` to restore `A/E + DSC/E == 1`,
  // matching Mac Excel 365's PRICE / YIELD / DURATION / MDURATION
  // output for those bases.
  out->bond_dsc = (basis == 2 || basis == 3) ? (period_days - out->days_bs) : out->days_nc;
  return true;
}

}  // namespace eval
}  // namespace formulon
