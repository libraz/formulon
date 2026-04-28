// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the irregular-last-period bond-pricing built-in:
//
//   * ODDLPRICE(settlement, maturity, last_interest, rate, yld,
//               redemption, frequency, [basis=0])
//
// Returns the clean price per 100 face. The "odd last" family handles
// bonds whose final coupon period does not align with the regular
// `12/freq`-month grid: the bond pays its periodic coupons up to
// `last_interest` and then a single irregular coupon + redemption at
// `maturity`. Microsoft's documented closed form (semi-annual / annual
// / quarterly):
//
//   cf   = 100 * rate / freq * (DC_total / E)
//   ai   = 100 * rate / freq * (A_total  / E)
//   disc = 1 + DSC * yld / freq / E
//   ODDLPRICE = (redemption + cf) / disc - ai
//
// where (DC_total, A_total, DSC, E) come from
// `compute_odd_last_schedule` -- see `financial_oddl_helpers.h` for the
// derivation. The companion ODDLYIELD inverts this in closed form.
//
// Implementation choice: rather than refactor the existing backward
// walker in `coupon_schedule.cpp`, the schedule walker is replicated
// locally in this TU (anchored on `last_interest` rather than
// `maturity`, walking forward instead of backward). The two walkers
// share month-end-clamping semantics by construction; folding them into
// one would either inflate `coupon_schedule.h`'s API surface or force
// an awkward "direction" parameter. Replication keeps the helpers
// small and the existing COUP* / PRICE / YIELD code paths untouched.

#include "eval/builtins/financial_oddlprice.h"

#include <cmath>
#include <cstdint>

#include "eval/builtins/financial_helpers.h"
#include "eval/builtins/financial_oddl_helpers.h"
#include "eval/coerce.h"
#include "eval/date_time.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Reads a required date argument, truncating toward zero. Negative
// serials are rejected as `#NUM!`. Mirrors the helper in
// `financial_price.cpp` / `financial_yield.cpp`; replicated per-TU to
// keep the helper local to its callers (see those files for the
// established pattern).
Expected<double, ErrorCode> read_date(const Value* args, std::uint32_t index) {
  auto raw = read_required_number(args, index);
  if (!raw) {
    return raw.error();
  }
  const double t = std::trunc(raw.value());
  if (t < 0.0) {
    return ErrorCode::Num;
  }
  return t;
}

// Last day-of-month for a given Gregorian (year, month) pair. Identical
// to the helper in `coupon_schedule.cpp`; replicated here to keep this
// TU's schedule walker self-contained without exporting symbols out of
// the existing COUP* engine.
unsigned last_day_of_month(int y, unsigned m) noexcept {
  static constexpr unsigned kTable[] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
  if (m < 1u || m > 12u) {
    return 31u;
  }
  if (m == 2u) {
    const bool leap = (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0);
    return leap ? 29u : 28u;
  }
  return kTable[m - 1u];
}

// Constructs the serial for (y, m, anchor_day), clamping the day to the
// target month's last day when shorter. Matches the COUP* engine's
// month-end-preservation rule: if last_interest is Aug-31 and we step
// three months forward, we land on Nov-30 (Nov has only 30 days).
double quasi_serial(int y, unsigned m, unsigned anchor_day) noexcept {
  const unsigned last = last_day_of_month(y, m);
  const unsigned d = anchor_day > last ? last : anchor_day;
  return date_time::serial_from_ymd(y, m, d);
}

// Shifts (y, m) forward by `months` (always positive). Uses
// floor-division so positive month overflow wraps correctly into the
// next year.
void shift_months_forward(int& y, unsigned& m, int months) noexcept {
  long long mm0 = static_cast<long long>(m) - 1 + static_cast<long long>(months);
  long long year_shift = mm0 / 12;
  long long rem = mm0 % 12;
  if (rem < 0) {
    rem += 12;
    year_shift -= 1;
  }
  y += static_cast<int>(year_shift);
  m = static_cast<unsigned>(rem + 1);
}

// Basis-adjusted days-between-two-serials. `a <= b`. The 30/360 bases
// decompose each serial via `ymd_from_serial` and apply the NASD / EU
// day-count rules (multiplied by 360 to match the integer day-count
// output the COUP* engine produces); bases 1, 2, 3 use raw serial
// difference. Identical to the helper in `coupon_schedule.cpp`.
double basis_days_between(double a, double b, int basis) noexcept {
  if (basis == 0 || basis == 4) {
    const date_time::YMD ya = date_time::ymd_from_serial(a);
    const date_time::YMD yb = date_time::ymd_from_serial(b);
    const double yf = basis == 0 ? date_time::yearfrac_us30_360(ya.y, ya.m, ya.d, yb.y, yb.m, yb.d)
                                 : date_time::yearfrac_eu30_360(ya.y, ya.m, ya.d, yb.y, yb.m, yb.d);
    return yf * 360.0;
  }
  return b - a;
}

// Period length E for the basis.
//   * Bases 0 / 4 (30/360 family): 360/freq — the formula's DC / A /
//     DSC ratios already use 30/360 day counts that sum to E, so the
//     nominal coupon length is consistent.
//   * Bases 1 / 2 / 3 (actual day-count family): the actual length of
//     the first regular quasi-period from `last_interest` forward by
//     `12/freq` months. Mac Excel 365 uses the same actual span on
//     all three of bases 1, 2, 3 inside ODDLPRICE / ODDLYIELD even
//     though their COUPDAYS values diverge — the function family
//     ignores the basis-specific year length here so that the DC / A
//     / DSC ratios stay coherent with the actual spans the schedule
//     walker produced. (Without this fall-through, basis 2 reuses
//     360/freq and basis 3 reuses 365/freq, and the `DC/E` ratio
//     becomes inconsistent with the actual day counts, producing
//     observable price drift versus Mac Excel for any coupon period
//     whose actual length is not exactly 360/freq or 365/freq.)
double normal_period_days(int basis, int frequency, double last_interest) noexcept {
  switch (basis) {
    case 0:
    case 4:
      return 360.0 / static_cast<double>(frequency);
    case 1:
    case 2:
    case 3: {
      // Actual length of the first quasi-period from `last_interest`
      // forward by `12/freq` months.
      const date_time::YMD li = date_time::ymd_from_serial(last_interest);
      int y = li.y;
      unsigned m = li.m;
      shift_months_forward(y, m, 12 / frequency);
      const double q2 = quasi_serial(y, m, li.d);
      return q2 - last_interest;
    }
    default:
      return 0.0;
  }
}

}  // namespace

Expected<OddLastSchedule, ErrorCode> compute_odd_last_schedule(double settlement, double maturity, double last_interest,
                                                               int frequency, int basis) noexcept {
  const double s = std::trunc(settlement);
  const double m = std::trunc(maturity);
  const double li = std::trunc(last_interest);

  // Clean-case day counts straight off the basis day-count engine.
  // Microsoft's published ODDLPRICE formula evaluates DC_total /
  // A_total / DSC as the basis-adjusted day spans across the irregular
  // period, with E being the *normal* coupon-period length. For all
  // five bases this collapses to a basis-driven yearfrac call,
  // multiplied by 360 to match the COUP* engine's integer-day output
  // for bases 0 / 4. For bases 1 / 2 / 3 the call returns the raw
  // serial difference (actual days), which agrees with Excel's
  // documented A / DC / DSC for those bases.
  const double dc_total_raw = basis_days_between(li, m, basis);
  const double a_total_raw = basis_days_between(li, s, basis);
  const double dsc_raw = basis_days_between(s, m, basis);

  // Round to integer days to match the COUP* engine's output style
  // (Excel reports COUPDAYBS / COUPDAYSNC / COUPDAYS as integers; the
  // odd-period formula expects the same "integer-grid" convention).
  // For bases 1 / 2 / 3 the inputs are already integer differences of
  // truncated serials, so std::round is a no-op there.
  OddLastSchedule out{};
  out.dc_total = std::round(dc_total_raw);
  out.a_total = std::round(a_total_raw);
  out.dsc = std::round(dsc_raw);
  out.e = normal_period_days(basis, frequency, li);

  if (std::isnan(out.dc_total) || std::isnan(out.a_total) || std::isnan(out.dsc) || std::isnan(out.e)) {
    return ErrorCode::Num;
  }
  if (out.e <= 0.0 || out.dc_total <= 0.0 || out.dsc <= 0.0) {
    return ErrorCode::Num;
  }
  return out;
}

Expected<double, ErrorCode> compute_oddl_clean_price(const Value* args, std::uint32_t arity) {
  auto settlement = read_date(args, 0);
  if (!settlement) {
    return settlement.error();
  }
  auto maturity = read_date(args, 1);
  if (!maturity) {
    return maturity.error();
  }
  auto last_interest = read_date(args, 2);
  if (!last_interest) {
    return last_interest.error();
  }
  auto rate = read_required_number(args, 3);
  if (!rate) {
    return rate.error();
  }
  auto yld = read_required_number(args, 4);
  if (!yld) {
    return yld.error();
  }
  auto redemption = read_required_number(args, 5);
  if (!redemption) {
    return redemption.error();
  }
  auto frequency_e = read_required_number(args, 6);
  if (!frequency_e) {
    return frequency_e.error();
  }
  auto basis_e = read_optional_number(args, arity, 7, 0.0);
  if (!basis_e) {
    return basis_e.error();
  }

  // Validation order mirrors PRICE: date ordering -> frequency -> basis
  // -> rate / yld signs -> redemption sign. ODDLPRICE has the extra
  // ordering constraint `last_interest < settlement` (the bond's last
  // regular coupon must precede settlement).
  if (last_interest.value() >= settlement.value()) {
    return ErrorCode::Num;
  }
  if (settlement.value() >= maturity.value()) {
    return ErrorCode::Num;
  }
  const int frequency = static_cast<int>(std::trunc(frequency_e.value()));
  if (frequency != 1 && frequency != 2 && frequency != 4) {
    return ErrorCode::Num;
  }
  const int basis = static_cast<int>(std::trunc(basis_e.value()));
  if (basis < 0 || basis > 4) {
    return ErrorCode::Num;
  }
  if (rate.value() < 0.0) {
    return ErrorCode::Num;
  }
  if (yld.value() < 0.0) {
    return ErrorCode::Num;
  }
  if (redemption.value() <= 0.0) {
    return ErrorCode::Num;
  }

  auto sched = compute_odd_last_schedule(settlement.value(), maturity.value(), last_interest.value(), frequency, basis);
  if (!sched) {
    return sched.error();
  }
  const double dc = sched.value().dc_total;
  const double a = sched.value().a_total;
  const double dsc = sched.value().dsc;
  const double e = sched.value().e;

  const double freq_d = static_cast<double>(frequency);
  const double cf = 100.0 * rate.value() / freq_d * (dc / e);
  const double ai = 100.0 * rate.value() / freq_d * (a / e);
  const double disc = 1.0 + dsc * yld.value() / freq_d / e;
  if (disc == 0.0) {
    return ErrorCode::Num;
  }
  const double price = (redemption.value() + cf) / disc - ai;
  if (std::isnan(price) || std::isinf(price)) {
    return ErrorCode::Num;
  }
  return price;
}

// --- ODDLPRICE(settlement, maturity, last_interest, rate, yld,
//              redemption, frequency, [basis=0]) ----------------------------
//
// Clean price per 100 face for a security whose final coupon period is
// irregular (the bond pays periodic coupons up to `last_interest` and a
// single irregular coupon + redemption at `maturity`).
Value OddlPrice(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto p = compute_oddl_clean_price(args, arity);
  if (!p) {
    return Value::error(p.error());
  }
  return finalize(p.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
