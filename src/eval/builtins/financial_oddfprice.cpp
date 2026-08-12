//
// Implementation of the irregular-first-period bond-pricing built-in:
//
//   * ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld,
//               redemption, frequency, [basis=0])
//
// Returns the clean price per 100 face. The "odd first" family
// handles a bond whose initial coupon period does not align with the
// regular `12/freq`-month grid: the bond is issued at `issue`, pays
// its first (potentially irregular) coupon at `first_coupon`, and
// then pays subsequent coupons on the regular schedule up to
// `maturity`.
//
// Microsoft's documentation publishes two distinct formulas keyed on
// whether the irregular first span fits within a single normal period
// (NC == 1, "short first") or spans multiple normal periods
// (NC > 1, "long first"). This implementation uses a single general
// formulation that walks the quasi-coupon schedule backward from
// `first_coupon` and discovers NC dynamically; when NC == 1 the
// formula collapses to the short-first-period case automatically.
//
// The formulation:
//
//   nc, qp[0..nc-1] = backward-walk quasi-coupon schedule from
//                     first_coupon (12/freq months per step) until
//                     qp[0].start <= issue
//   For each i:
//     nl[i] = basis-adjusted length of qp[i]
//     dc[i] = (i == 0) ? days(max(qp[0].start, issue), qp[0].end)
//                      : nl[i]
//     a[i]  = days within qp[i] that are <= settlement (with `issue`
//             substituted for qp[0].start when i == 0); zero for
//             quasi-periods strictly after the settlement-bearing one
//   dsc     = days from settlement to first_coupon (sum across the
//             settlement-bearing quasi-period plus subsequent ones)
//   E       = normal coupon-period length (basis-adjusted)
//   N       = compute_coupon_dates(first_coupon, maturity, ...).coupons_remaining
//
//   v   = 1 / (1 + yld/freq)
//   cf  = 100 * rate / freq
//
//   first_period_pv = cf * sum_{i=0..nc-1} (dc[i]/nl[i]) * v^((nc-1-i) + dsc/E)
//   reg_coupons_pv  = cf * sum_{j=1..N} v^(dsc/E + j)
//   redemption_pv   = redemption * v^(dsc/E + N)
//   ai              = cf * sum_i (a[i]/nl[i])
//
//   ODDFPRICE = first_period_pv + reg_coupons_pv + redemption_pv - ai
//
// Coupon-schedule mechanics for the trailing regular periods are
// delegated to the shared engine in `eval/coupon_schedule.h` (as in
// PRICE / YIELD / DURATION / ACCRINT).

#include "eval/builtins/financial_oddfprice.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/financial_oddf_helpers.h"
#include "eval/coupon_schedule.h"
#include "eval/date_time.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Calendar day-count helpers shared with the COUP* engine and the date
// builtins; see `eval/date_time.h`.
using date_time::basis_days_between;
using date_time::days_in_month;

// Constructs the serial for (y, m, anchor_day), clamping the day to the
// target month's last day when shorter. Matches the COUP* engine's
// month-end-preservation rule: if `first_coupon` is Aug-31 and we step
// three months backward, we land on May-31; six months -> Feb-28 (or
// Feb-29 in a leap year).
double quasi_serial(int y, unsigned m, unsigned anchor_day) noexcept {
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

// Period length E for the basis. Bases 0/2/4 use 360/freq; basis 3
// uses 365/freq; basis 1 uses the actual length of the most-recent
// quasi-period (the one anchored on `first_coupon`, walking 12/freq
// months back and forward).
//
// Bases 2 / 3 deliberately reuse the nominal `360/freq` / `365/freq`
// here (and feed the same value into `nl[i]` for full quasi-periods
// in the schedule below) so that `dc[i] / nl[i]` and `a[i] / nl[i]`
// stay coherent with Mac Excel's published ODDFPRICE / ODDFYIELD
// output. Without this — i.e. if `nl[i]` came from the *actual*
// `qend - qstart` while `e` came from the nominal `360/freq`, as the
// legacy `basis_days_between` path produced — the per-quasi-period
// PV contributions would mix two different period lengths and drift
// from Excel by ~0.012 per 100 face on the docs canonical case.
double normal_period_days(int basis, int frequency, double first_coupon) noexcept {
  switch (basis) {
    case 0:
    case 2:
    case 4:
      return 360.0 / static_cast<double>(frequency);
    case 3:
      return 365.0 / static_cast<double>(frequency);
    case 1: {
      // Actual length of the regular quasi-period ending on
      // `first_coupon`: step back 12/freq months and measure the gap.
      const date_time::YMD fc = date_time::ymd_from_serial(first_coupon);
      int y = fc.y;
      unsigned m = fc.m;
      shift_months_back(y, m, 12 / frequency);
      const double q_prev = quasi_serial(y, m, fc.d);
      return first_coupon - q_prev;
    }
    default:
      return 0.0;
  }
}

// Full quasi-period length used by `nl[i]` and by the `a[i]` /
// `dc[i]` slots that represent a *full* quasi-period (rather than a
// partial span carved out by `issue` or `settlement`). This must
// match `normal_period_days(...)` so that for full periods
// `dc[i] / nl[i] == 1` (full coupon paid) — see the explanation
// above for why bases 2 / 3 use the nominal length.
double full_quasi_period_length(int basis, int frequency, double first_coupon) noexcept {
  return normal_period_days(basis, frequency, first_coupon);
}

}  // namespace

Expected<OddFirstSchedule, ErrorCode> compute_odd_first_schedule(double settlement, double maturity, double issue,
                                                                 double first_coupon, int frequency,
                                                                 int basis) noexcept {
  const double s = std::trunc(settlement);
  const double mat = std::trunc(maturity);
  const double iss = std::trunc(issue);
  const double fc = std::trunc(first_coupon);

  // --- Walk quasi-coupon schedule backward from first_coupon by
  // 12/freq months until we cross or touch `issue`. We collect the
  // (start, end) serials, oldest first.
  const date_time::YMD fc_ymd = date_time::ymd_from_serial(fc);
  const unsigned anchor_day = fc_ymd.d;
  const int step_months = 12 / frequency;

  // Temporary storage: walk backward, stamp the (end, start) of each
  // quasi-period, then reverse so the oldest is at index 0. We cap at
  // `kMaxOddFirstQuasiPeriods` to defend against pathological inputs
  // (a 200-period span at semi-annual frequency is 100 years).
  double end_serials[kMaxOddFirstQuasiPeriods + 1];
  end_serials[0] = fc;
  int y_walk = fc_ymd.y;
  unsigned m_walk = fc_ymd.m;
  int count = 0;  // number of quasi-periods discovered so far
  while (true) {
    if (count >= kMaxOddFirstQuasiPeriods) {
      return ErrorCode::Num;
    }
    shift_months_back(y_walk, m_walk, step_months);
    const double prev = quasi_serial(y_walk, m_walk, anchor_day);
    end_serials[count + 1] = prev;
    ++count;
    if (prev <= iss) {
      break;
    }
  }
  // count = NC, end_serials[0] = first_coupon, end_serials[NC] = qp[0].start.

  OddFirstSchedule out{};
  out.nc = count;

  // Populate qp[0..nc-1] sorted oldest -> newest. qp[i].start is the
  // earlier serial, qp[i].end is the later one.
  for (int i = 0; i < count; ++i) {
    const int rev = count - 1 - i;  // walk i=0 = oldest
    out.qp[i].start = end_serials[rev + 1];
    out.qp[i].end = end_serials[rev];
  }

  // --- Per-quasi-period day counts: nl, dc, a.
  //
  //   nl[i] = full-period length for the basis (must match `e` so the
  //           `dc[i] / nl[i]` ratio is unitless and full periods give
  //           ratio == 1 — see `full_quasi_period_length` above)
  //   dc[i] = (i == 0) ? basis_days_between(max(qp[0].start, issue), qp[0].end)
  //                    : nl[i]
  //   a[i]  = days within qp[i] that fall in [start_or_issue, settlement]:
  //             - i strictly before settlement-bearing: full nl[i]
  //               (with the i==0 substitution start := issue applying
  //               only when issue is *inside* qp[0])
  //             - i equal to settlement-bearing: days from start_or_issue
  //               to settlement
  //             - i strictly after: 0
  //
  // For bases 0 / 4 the basis-adjusted day-count call rounds via
  // std::round to match Excel's integer day-count grid. For basis 1
  // (actual day counts coherent with `e = ncd - pcd`) the same path
  // produces integer-valued doubles. For bases 2 / 3 the *full*
  // period quantities (`nl[i]`, plus `dc[i]` / `a[i]` when they
  // represent a full quasi-period) come from the nominal length so
  // they stay coherent with `e`; partial spans (issue->qend,
  // start->settle) keep their actual day counts because that is what
  // the published `dc / nl` and `a / nl` ratios assume.
  const double full_nl = full_quasi_period_length(basis, frequency, fc);
  if (full_nl <= 0.0) {
    return ErrorCode::Num;
  }
  for (int i = 0; i < count; ++i) {
    const double qstart = out.qp[i].start;
    const double qend = out.qp[i].end;
    const double effective_start = (i == 0 && qstart < iss) ? iss : qstart;

    out.qp[i].nl = full_nl;
    if (i == 0) {
      out.qp[i].dc = std::round(basis_days_between(effective_start, qend, basis));
    } else {
      out.qp[i].dc = out.qp[i].nl;
    }

    if (s >= qend) {
      if (i == 0 && qstart < iss) {
        // Issue is inside qp[0] *and* settlement is past qend ->
        // accrual is the partial issue-to-qend span (actual days).
        out.qp[i].a = std::round(basis_days_between(effective_start, qend, basis));
      } else {
        // Full quasi-period accrued -> use the basis full-period
        // length so the ratio `a[i] / nl[i] == 1`.
        out.qp[i].a = full_nl;
      }
    } else if (s > effective_start) {
      // Settlement falls inside qp[i] -> partial accrual (actual days).
      out.qp[i].a = std::round(basis_days_between(effective_start, s, basis));
    } else {
      // Settlement is at or before this quasi-period's effective start
      // -> no accrual on this or subsequent quasi-periods. (Shouldn't
      // happen given `issue < settlement`, but guard defensively.)
      out.qp[i].a = 0.0;
    }
  }

  // --- DSC: days from settlement to first_coupon, summed across the
  // settlement-bearing quasi-period plus all subsequent ones in the
  // irregular span. By design, `settlement < first_coupon`, so DSC > 0.
  double dsc = 0.0;
  for (int i = 0; i < count; ++i) {
    const double qstart = out.qp[i].start;
    const double qend = out.qp[i].end;
    if (s >= qend) {
      // No DSC contribution from this quasi-period.
      continue;
    }
    if (s > qstart) {
      // Settlement falls inside this quasi-period.
      dsc += basis_days_between(s, qend, basis);
    } else {
      // Settlement is before this quasi-period entirely.
      dsc += basis_days_between(qstart, qend, basis);
    }
  }
  out.dsc = std::round(dsc);
  if (out.dsc <= 0.0) {
    return ErrorCode::Num;
  }

  // --- Normal period length E.
  out.e = normal_period_days(basis, frequency, fc);
  if (out.e <= 0.0 || std::isnan(out.e) || std::isinf(out.e)) {
    return ErrorCode::Num;
  }

  // --- Regular coupons N from first_coupon to maturity. We borrow the
  // shared coupon-schedule engine, treating `first_coupon` as a
  // synthetic "settlement" date strictly before any actual coupon
  // (since first_coupon IS itself a coupon, the engine reports the
  // count of coupons strictly after settlement and <= maturity).
  CouponDates cd{};
  if (!compute_coupon_dates(fc, mat, frequency, basis, &cd)) {
    return ErrorCode::Num;
  }
  if (cd.coupons_remaining < 0) {
    return ErrorCode::Num;
  }
  // `compute_coupon_dates` reports coupons STRICTLY AFTER settlement.
  // With `settlement = first_coupon` it returns the count of regular
  // coupons in (first_coupon, maturity]; we add 1 for the
  // first_coupon itself only when it's not already counted -- but
  // first_coupon is the irregular coupon (handled by first_period_pv),
  // not part of the regular series. So `n_regular = coupons_remaining`
  // counts the regular coupons after first_coupon up to and including
  // maturity. When first_coupon == maturity, coupons_remaining == 0
  // and there are no regular coupons (a degenerate case caller already
  // rejected via first_coupon < maturity).
  out.n_regular = cd.coupons_remaining;
  if (out.n_regular <= 0) {
    return ErrorCode::Num;
  }

  return out;
}

Expected<double, ErrorCode> compute_oddf_clean_price(const Value* args, std::uint32_t arity) {
  auto settlement = read_financial_date(args, 0);
  if (!settlement) {
    return settlement.error();
  }
  auto maturity = read_financial_date(args, 1);
  if (!maturity) {
    return maturity.error();
  }
  auto issue = read_financial_date(args, 2);
  if (!issue) {
    return issue.error();
  }
  auto first_coupon = read_financial_date(args, 3);
  if (!first_coupon) {
    return first_coupon.error();
  }
  auto rate = read_required_number(args, 4);
  if (!rate) {
    return rate.error();
  }
  auto yld = read_required_number(args, 5);
  if (!yld) {
    return yld.error();
  }
  auto redemption = read_required_number(args, 6);
  if (!redemption) {
    return redemption.error();
  }
  auto frequency_e = read_coupon_frequency(args, 7);
  if (!frequency_e) {
    return frequency_e.error();
  }
  auto basis_e = read_day_count_basis(args, arity, 8);
  if (!basis_e) {
    return basis_e.error();
  }

  // Validation order mirrors PRICE / ODDLPRICE: date ordering ->
  // frequency -> basis -> rate / yld signs -> redemption sign.
  // ODDFPRICE has the strict ordering constraint
  // `issue < settlement < first_coupon < maturity` -- the issue date
  // must precede settlement (the holder is past the issue at trade
  // time), settlement must precede the first coupon (hence "first"
  // means the first coupon yet to come), and first_coupon must
  // precede maturity (else there are no regular coupons after the
  // irregular first one).
  if (issue.value() >= settlement.value()) {
    return ErrorCode::Num;
  }
  if (settlement.value() >= first_coupon.value()) {
    return ErrorCode::Num;
  }
  if (first_coupon.value() >= maturity.value()) {
    return ErrorCode::Num;
  }
  const int frequency = frequency_e.value();
  const int basis = basis_e.value();
  if (rate.value() < 0.0) {
    return ErrorCode::Num;
  }
  if (yld.value() < 0.0) {
    return ErrorCode::Num;
  }
  if (redemption.value() <= 0.0) {
    return ErrorCode::Num;
  }

  auto sched_e = compute_odd_first_schedule(settlement.value(), maturity.value(), issue.value(), first_coupon.value(),
                                            frequency, basis);
  if (!sched_e) {
    return sched_e.error();
  }
  const OddFirstSchedule& sched = sched_e.value();

  const double freq_d = static_cast<double>(frequency);
  const double cf = 100.0 * rate.value() / freq_d;
  const double v_denom = 1.0 + yld.value() / freq_d;
  if (v_denom == 0.0) {
    return ErrorCode::Num;
  }
  const double v = 1.0 / v_denom;
  const double dsc_over_e = sched.dsc / sched.e;

  // First-period present value: each quasi-period i contributes
  //   (dc[i] / nl[i]) * cf * v^((nc-1-i) + dsc/E)
  // i.e. quasi-period i is discounted by (nc-1-i) full periods after
  // first_coupon's discount factor v^(dsc/E). For NC == 1 the sole
  // term is (dc[0] / nl[0]) * cf * v^(dsc/E), matching Microsoft's
  // documented short-first-period formula.
  double first_period_pv = 0.0;
  for (int i = 0; i < sched.nc; ++i) {
    const double exp_periods = static_cast<double>(sched.nc - 1 - i) + dsc_over_e;
    const double disc = std::pow(v, exp_periods);
    if (std::isnan(disc) || std::isinf(disc)) {
      return ErrorCode::Num;
    }
    if (sched.qp[i].nl <= 0.0) {
      return ErrorCode::Num;
    }
    first_period_pv += (sched.qp[i].dc / sched.qp[i].nl) * cf * disc;
  }

  // Regular coupons: cf * sum_{j=1..N} v^(dsc/E + j).
  double reg_coupons_pv = 0.0;
  for (std::int32_t j = 1; j <= sched.n_regular; ++j) {
    const double disc = std::pow(v, dsc_over_e + static_cast<double>(j));
    if (std::isnan(disc) || std::isinf(disc)) {
      return ErrorCode::Num;
    }
    reg_coupons_pv += cf * disc;
  }

  // Redemption at maturity: redemption * v^(dsc/E + N).
  const double red_disc = std::pow(v, dsc_over_e + static_cast<double>(sched.n_regular));
  if (std::isnan(red_disc) || std::isinf(red_disc)) {
    return ErrorCode::Num;
  }
  const double redemption_pv = redemption.value() * red_disc;

  // Accrued interest: cf * sum_i (a[i] / nl[i]).
  double ai_units = 0.0;
  for (int i = 0; i < sched.nc; ++i) {
    if (sched.qp[i].nl <= 0.0) {
      return ErrorCode::Num;
    }
    ai_units += sched.qp[i].a / sched.qp[i].nl;
  }
  const double ai = cf * ai_units;

  const double price = first_period_pv + reg_coupons_pv + redemption_pv - ai;
  if (std::isnan(price) || std::isinf(price)) {
    return ErrorCode::Num;
  }
  return price;
}

// --- ODDFPRICE(settlement, maturity, issue, first_coupon, rate, yld,
//              redemption, frequency, [basis=0]) ----------------------------
//
// Clean price per 100 face for a security whose first coupon period
// is irregular (the bond is issued at `issue`, pays its first coupon
// at `first_coupon`, and pays subsequent coupons on the regular grid
// up to `maturity`).
Value OddfPrice(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto p = compute_oddf_clean_price(args, arity);
  if (!p) {
    return Value::error(p.error());
  }
  return finalize(p.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
