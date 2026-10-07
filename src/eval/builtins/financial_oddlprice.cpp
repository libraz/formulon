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
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {
namespace {

// Calendar day-count helpers shared with the COUP* engine and the date
// builtins; see `eval/date_time.h`.
using date_time::basis_days_between;

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
double normal_period_days(int basis, int frequency, double last_interest, bool date1904) noexcept {
  switch (basis) {
    case 0:
    case 4:
      return 360.0 / static_cast<double>(frequency);
    case 1:
    case 2:
    case 3: {
      // Actual length of the first quasi-period from `last_interest`
      // forward by `12/freq` months.
      const date_time::YMD li = date_time::ymd_from_serial(last_interest, date1904);
      int y = li.y;
      unsigned m = li.m;
      shift_months(y, m, 12 / frequency);
      const double q2 = quasi_serial(y, m, li.d, date1904);
      return q2 - last_interest;
    }
    default:
      return 0.0;
  }
}

}  // namespace

Expected<OddLastSchedule, ErrorCode> compute_odd_last_schedule(double settlement, double maturity, double last_interest,
                                                               int frequency, int basis, bool date1904) noexcept {
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
  const double dc_total_raw = basis_days_between(li, m, basis, date1904);
  const double a_total_raw = basis_days_between(li, s, basis, date1904);
  const double dsc_raw = basis_days_between(s, m, basis, date1904);

  // Round to integer days to match the COUP* engine's output style
  // (Excel reports COUPDAYBS / COUPDAYSNC / COUPDAYS as integers; the
  // odd-period formula expects the same "integer-grid" convention).
  // For bases 1 / 2 / 3 the inputs are already integer differences of
  // truncated serials, so std::round is a no-op there.
  OddLastSchedule out{};
  out.dc_total = std::round(dc_total_raw);
  out.a_total = std::round(a_total_raw);
  out.dsc = std::round(dsc_raw);
  out.e = normal_period_days(basis, frequency, li, date1904);

  if (std::isnan(out.dc_total) || std::isnan(out.a_total) || std::isnan(out.dsc) || std::isnan(out.e)) {
    return ErrorCode::Num;
  }
  if (out.e <= 0.0 || out.dc_total <= 0.0 || out.dsc <= 0.0) {
    return ErrorCode::Num;
  }
  return out;
}

Expected<OddLastInputs, ErrorCode> read_odd_last_inputs(const Value* args, std::uint32_t arity,
                                                        bool slot4_must_be_positive, bool date1904) {
  double dates[3];
  if (auto read = read_required_numbers(args, "ddd", dates); !read) {
    return std::move(read.error());
  }
  const double settlement = dates[0];
  const double maturity = dates[1];
  const double last_interest = dates[2];
  auto tail = read_coupon_bond_tail(args, arity, 3);
  if (!tail) {
    return std::move(tail.error());
  }
  const auto [rate, slot4, redemption, frequency, basis] = tail.value();

  // Validation mirrors PRICE. The odd-last family adds the
  // ordering constraint `last_interest < settlement` (the bond's last
  // regular coupon must precede settlement).
  if (last_interest >= settlement) {
    return ErrorCode::Num;
  }
  if (settlement >= maturity) {
    return ErrorCode::Num;
  }
  if (slot4_must_be_positive ? (slot4 <= 0.0) : (slot4 < 0.0)) {
    return ErrorCode::Num;
  }

  auto sched = compute_odd_last_schedule(settlement, maturity, last_interest, frequency, basis, date1904);
  if (!sched) {
    return std::move(sched.error());
  }

  const double freq_d = static_cast<double>(frequency);
  OddLastInputs out{};
  out.slot4 = slot4;
  out.redemption = redemption;
  out.dsc = sched.value().dsc;
  out.e = sched.value().e;
  out.freq_d = freq_d;
  out.cf = 100.0 * rate / freq_d * (sched.value().dc_total / out.e);
  out.ai = 100.0 * rate / freq_d * (sched.value().a_total / out.e);
  return out;
}

Expected<double, ErrorCode> compute_oddl_clean_price(const Value* args, std::uint32_t arity, bool date1904) {
  auto in = read_odd_last_inputs(args, arity, /*slot4_must_be_positive=*/false, date1904);
  if (!in) {
    return std::move(in.error());
  }
  const double disc = 1.0 + in.value().dsc * in.value().slot4 / in.value().freq_d / in.value().e;
  if (disc == 0.0) {
    return ErrorCode::Num;
  }
  const double price = (in.value().redemption + in.value().cf) / disc - in.value().ai;
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
Value OddlPrice(const Value* args, std::uint32_t arity, Arena& /*arena*/, bool date1904) {
  auto p = compute_oddl_clean_price(args, arity, date1904);
  if (!p) {
    return Value::error(p.error());
  }
  return finalize(p.value());
}

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon
