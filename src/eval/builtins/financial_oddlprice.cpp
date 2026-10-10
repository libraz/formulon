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
//   cf   = 100 * rate / freq * sum_i DC_i / NL_i
//   ai   = 100 * rate / freq * sum_i A_i  / NL_i
//   disc = 1 + (sum_i DSC_i / NL_i) * yld / freq
//   ODDLPRICE = (redemption + cf) / disc - ai
//
// summed over the quasi-coupon periods from `last_interest` to `maturity`
// as built by `compute_odd_last_schedule` -- see `financial_oddl_helpers.h` for the
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

// Quasi-coupon serial `k` periods after `last_interest`, keeping its day
// of month (clamped to shorter months).
double quasi_after(const date_time::YMD& anchor, int k, int frequency, bool date1904) noexcept {
  int y = anchor.y;
  unsigned m = anchor.m;
  date_time::shift_year_month(y, m, static_cast<std::int64_t>(k) * (12 / frequency));
  return date_time::serial_from_ymd_clamped(y, m, anchor.d, date1904);
}

}  // namespace

Expected<OddLastSchedule, ErrorCode> compute_odd_last_schedule(double settlement, double maturity, double last_interest,
                                                               int frequency, int basis, bool date1904) noexcept {
  const double s = std::trunc(settlement);
  const double m = std::trunc(maturity);
  const double li = std::trunc(last_interest);
  const date_time::YMD anchor = date_time::ymd_from_serial(li, date1904);
  const auto days = [&](double a, double b) { return std::round(basis_days_between(a, b, basis, date1904)); };

  // Each quasi-period (q(k-1), q(k)] contributes its DC / A / DSC days over
  // its own normal length NL: 360/freq for the 30/360 bases, the period's
  // actual length for bases 1 / 2 / 3.
  OddLastSchedule out{};
  for (int k = 1;; ++k) {
    const double qstart = quasi_after(anchor, k - 1, frequency, date1904);
    const double qend = quasi_after(anchor, k, frequency, date1904);
    const double nl = (basis == 0 || basis == 4) ? 360.0 / static_cast<double>(frequency) : qend - qstart;
    if (nl <= 0.0) {
      return ErrorCode::Num;
    }
    const double end = qend < m ? qend : m;
    out.dc_units += days(qstart, end) / nl;
    if (s > qstart) {
      out.a_units += days(qstart, s < end ? s : end) / nl;
    }
    if (end > s) {
      out.dsc_units += days(s > qstart ? s : qstart, end) / nl;
    }
    if (qend >= m) {
      break;
    }
  }
  if (std::isnan(out.dc_units) || std::isnan(out.a_units) || std::isnan(out.dsc_units)) {
    return ErrorCode::Num;
  }
  if (out.dc_units <= 0.0 || out.dsc_units <= 0.0) {
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
  out.dsc_units = sched.value().dsc_units;
  out.freq_d = freq_d;
  out.cf = 100.0 * rate / freq_d * sched.value().dc_units;
  out.ai = 100.0 * rate / freq_d * sched.value().a_units;
  return out;
}

Expected<double, ErrorCode> compute_oddl_clean_price(const Value* args, std::uint32_t arity, bool date1904) {
  auto in = read_odd_last_inputs(args, arity, /*slot4_must_be_positive=*/false, date1904);
  if (!in) {
    return std::move(in.error());
  }
  const double disc = 1.0 + in.value().dsc_units * in.value().slot4 / in.value().freq_d;
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
