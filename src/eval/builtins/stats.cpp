// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Statistical built-ins: dispersion family (VAR.S / VAR.P / STDEV.S /
// STDEV.P) plus the "A" variants (AVERAGEA / MAXA / MINA / VARA / VARPA /
// STDEVA / STDEVPA), the shared collector / moment helpers consumed by
// the sibling stats TUs, and the central `register_stats_builtins`
// registrar that wires every entry into the registry.
//
// The order-statistic family (MEDIAN / MODE / LARGE / SMALL / PERCENTILE
// / QUARTILE / TRIMMEAN) lives in `stats/stats_order.cpp`. The higher-
// moment descriptive family (GEOMEAN / HARMEAN / DEVSQ / AVEDEV / SKEW /
// SKEW.P / KURT / STANDARDIZE) lives in `stats/stats_moments.cpp`. The
// probability-distribution catalog (NORM.*, BINOM.DIST, POISSON.DIST,
// EXPON.DIST, CHISQ.*, T.*, F.*, plus legacy NORMSDIST / TDIST) lives
// in `stats/stats_distributions.cpp` and `stats/stats_distributions_misc.cpp`.
// All four TUs share `stats/stats_helpers.h` and register here.
//
// Argument-type rule (DIFFERENT from SUM / AVERAGE / MIN / MAX / PRODUCT):
// these functions silently SKIP text, boolean, and blank inputs instead of
// coercing them. Only values whose kind is `Number` participate. The
// dispatcher runs with `propagate_errors = true` so error-typed arguments
// still short-circuit before the impl executes; once inside the body we
// only need to filter the non-numeric non-error kinds. Contrast this with
// `Sum` in `builtins/aggregate.cpp`, which coerces every argument through
// `coerce_to_number` and surfaces `#VALUE!` on text like `"abc"`.

#include "eval/builtins/stats.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/aggregate_kernels.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/builtins/stats/stats_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace stats_detail {

// --- Shared helpers (declared in `stats/stats_helpers.h`). ---------------
//
// Each of these is a thin wrapper around `eval::collect_numerics(Value*,
// count, NumericCollectPolicy)` -- the policy flags pick the per-family
// type-coercion rule (numeric-only AVERAGE / SUM family, "A"-family with
// Bool+Text+Blank coercion, SMALL / LARGE direct-scalar with Bool+Text
// but no Blank). The shared helper consolidates the per-kind switch so
// the SMALL / LARGE / A-family rule cannot drift from its companions in
// `eval/coerce.cpp`.

std::vector<double> collect_numerics(const Value* args, std::uint32_t count) {
  // Errors never reach this helper because every caller is registered
  // with `propagate_errors = true`, so the dispatcher short-circuits
  // before the impl runs. The default policy still has
  // `error_on_error_cell = true`, but an Error cell at this stage is
  // unreachable; calling `.value()` directly is safe.
  NumericCollectPolicy policy;  // Defaults: numbers only.
  auto collected = eval::collect_numerics(args, count, policy);
  if (!collected) {
    // Unreachable in normal operation; return an empty slice so the
    // caller's subsequent emptiness guard fires gracefully.
    return {};
  }
  return std::move(collected.value());
}

// Direct-scalar-aware collector for SMALL / LARGE. Mirrors the "A"-family
// rule for direct scalar arguments but differs for the "anything else"
// category: range-sourced Text / Bool / Blank cells have already been
// dropped by the dispatcher's `range_filter_numeric_only` filter, so by
// the time they reach this helper a Text or Bool kind can only come from
// a direct scalar literal (e.g. `SMALL("3.4", 1)` or `SMALL(TRUE, 1)`)
// or from an array-literal element that survived the filter (it cannot --
// the filter drops those too, so those never appear here either).
// Direct Number -> kept as-is; direct Bool -> 1.0 / 0.0;
// direct Text -> strict `coerce_to_number` (propagates #VALUE! on
// unparseable text like `"Hello"`). Anything else (Blank left over from
// a dropped optional argument, unresolved Ref, Array, Lambda) is skipped
// silently, matching AVERAGE-family leniency.
Expected<std::vector<double>, ErrorCode> collect_small_large(const Value* args, std::uint32_t count) {
  NumericCollectPolicy policy;
  policy.include_bool = true;                  // Direct Bool -> 1 / 0.
  policy.include_text_numeric_literal = true;  // Direct Text -> strict coerce.
  policy.error_on_text = true;                 // "Hello" -> #VALUE!.
  // Blank stays dropped (default `blank_as_zero = false`) -- the SMALL /
  // LARGE direct-scalar rule diverges from the "A"-family here.
  return eval::collect_numerics(args, count, policy);
}

Expected<std::vector<double>, ErrorCode> collect_a(const Value* args, std::uint32_t count) {
  NumericCollectPolicy policy;
  policy.include_bool = true;                  // Direct Bool -> 1 / 0.
  policy.include_text_numeric_literal = true;  // Direct Text -> strict coerce.
  policy.error_on_text = true;                 // "Hola" -> #VALUE!.
  policy.blank_as_zero = true;                 // Direct Blank -> 0 ("A"-family rule).
  return eval::collect_numerics(args, count, policy);
}

MeanSS compute_mean_ss(const std::vector<double>& xs) {
  if (xs.empty()) {
    return {0.0, 0.0};
  }
  double sum = 0.0;
  for (double x : xs) {
    sum += x;
  }
  const double mean = sum / static_cast<double>(xs.size());
  double ss = 0.0;
  for (double x : xs) {
    const double d = x - mean;
    ss += d * d;
  }
  return {mean, ss};
}

Expected<CenteredScaled, ErrorCode> centered_and_scaled(const Value* args, std::uint32_t arity, std::size_t min_n,
                                                        double sample_offset) {
  std::vector<double> xs = collect_numerics(args, arity);
  if (xs.size() < min_n) {
    return ErrorCode::Div0;
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double n = static_cast<double>(xs.size());
  const double variance = ms.ss / (n - sample_offset);
  if (variance == 0.0) {
    return ErrorCode::Div0;
  }
  return CenteredScaled{std::move(xs), n, ms.mean, std::sqrt(variance)};
}

double mean_of(const std::vector<double>& xs) noexcept {
  double s = 0.0;
  for (double x : xs) {
    s += x;
  }
  return s / static_cast<double>(xs.size());
}

// Core dispersion kernel shared by VAR.S / VAR.P / STDEV.S / STDEV.P (and
// indirectly by VARA / VARPA / STDEVA / STDEVPA through
// `variance_or_stdev_a`). `sample = true` selects the (n - 1) divisor;
// `square_root = true` selects the stdev branch. Empty / single-element
// inputs collapse to `#DIV/0!` per the relevant Excel rule.
static Value variance_or_stdev(const std::vector<double>& xs, bool sample, bool square_root) {
  if (sample ? xs.size() < 2u : xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  const MeanSS ms = compute_mean_ss(xs);
  const double divisor = sample ? static_cast<double>(xs.size() - 1u) : static_cast<double>(xs.size());
  const double variance = ms.ss / divisor;
  const double r = square_root ? std::sqrt(variance) : variance;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// "A"-family dispatch onto `variance_or_stdev`: collects via `collect_a`
// (Bool / Text / Blank coerced or dropped per the AVERAGEA rule) then
// delegates to the shared kernel.
static Value variance_or_stdev_a(const Value* args, std::uint32_t arity, bool sample, bool square_root) {
  auto collected = collect_a(args, arity);
  if (!collected) {
    return Value::error(collected.error());
  }
  return variance_or_stdev(collected.value(), sample, square_root);
}

Expected<double, ErrorCode> read_kth_arg(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return d;
}

Value percentile_inc_sorted(const std::vector<double>& xs, double k) {
  // Delegates to the shared kernel; callers (`PercentileInc`, `QuartileInc`)
  // have already guaranteed `xs` is non-empty and `k in [0, 1]`, so the
  // kernel's domain guards are no-ops on this path. The kernel returns
  // `Expected<double, ErrorCode>` so we lift through `finite_number_result`
  // for the success path and propagate `#NUM!` verbatim on failure (which
  // can still arise on a non-finite interpolated blend).
  auto r = aggregate_kernels::percentile_sorted_inc(xs, k);
  if (!r) {
    return Value::error(r.error());
  }
  return finite_number_result(r.value());
}

Value percentile_exc_sorted(const std::vector<double>& xs, double k) {
  // Same delegation pattern as `percentile_inc_sorted`. The shared kernel
  // enforces the exclusive-method boundary check `idx < 1 || idx >= n`
  // (matching Mac Excel 365's `#NUM!` at the open-interval boundaries),
  // so callers do not have to re-validate `k` against `1/(n+1)` or
  // `n/(n+1)` themselves.
  auto r = aggregate_kernels::percentile_sorted_exc(xs, k);
  if (!r) {
    return Value::error(r.error());
  }
  return finite_number_result(r.value());
}

ModeFrequencies build_mode_frequencies(const std::vector<double>& xs) {
  ModeFrequencies freq;
  freq.values.reserve(xs.size());
  freq.counts.reserve(xs.size());
  for (double v : xs) {
    bool found = false;
    for (std::size_t i = 0; i < freq.values.size(); ++i) {
      if (freq.values[i] == v) {
        ++freq.counts[i];
        if (freq.counts[i] > freq.best_count) {
          freq.best_count = freq.counts[i];
        }
        found = true;
        break;
      }
    }
    if (!found) {
      freq.values.push_back(v);
      freq.counts.push_back(1u);
      if (freq.best_count == 0u) {
        freq.best_count = 1u;
      }
    }
  }
  return freq;
}

double InverseStandardNormal(double p) {
  // Acklam's coefficient tables. The 'a' / 'b' coefficients cover the
  // central region; 'c' / 'd' cover both tails.
  static constexpr double a1 = -3.969683028665376e+01;
  static constexpr double a2 = 2.209460984245205e+02;
  static constexpr double a3 = -2.759285104469687e+02;
  static constexpr double a4 = 1.383577518672690e+02;
  static constexpr double a5 = -3.066479806614716e+01;
  static constexpr double a6 = 2.506628277459239e+00;

  static constexpr double b1 = -5.447609879822406e+01;
  static constexpr double b2 = 1.615858368580409e+02;
  static constexpr double b3 = -1.556989798598866e+02;
  static constexpr double b4 = 6.680131188771972e+01;
  static constexpr double b5 = -1.328068155288572e+01;

  static constexpr double c1 = -7.784894002430293e-03;
  static constexpr double c2 = -3.223964580411365e-01;
  static constexpr double c3 = -2.400758277161838e+00;
  static constexpr double c4 = -2.549732539343734e+00;
  static constexpr double c5 = 4.374664141464968e+00;
  static constexpr double c6 = 2.938163982698783e+00;

  static constexpr double d1 = 7.784695709041462e-03;
  static constexpr double d2 = 3.224671290700398e-01;
  static constexpr double d3 = 2.445134137142996e+00;
  static constexpr double d4 = 3.754408661907416e+00;

  static constexpr double p_low = 0.02425;
  static constexpr double p_high = 1.0 - p_low;

  // Constant pi used by the Halley refinement below. Shared with the
  // distribution TUs via `stats_helpers.h::kStatsPi`.

  double z;
  if (p < p_low) {
    const double q = std::sqrt(-2.0 * std::log(p));
    z = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / ((((d1 * q + d2) * q + d3) * q + d4) * q + 1.0);
  } else if (p <= p_high) {
    const double q = p - 0.5;
    const double r = q * q;
    z = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q /
        (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1.0);
  } else {
    const double q = std::sqrt(-2.0 * std::log(1.0 - p));
    z = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / ((((d1 * q + d2) * q + d3) * q + d4) * q + 1.0);
  }
  // Halley polish: one iteration is enough to converge from Acklam's
  // ~1e-6 residual down to ~1e-14. The update formula uses:
  //   residual e  = Phi(z) - p   (standard-normal CDF minus target)
  //   derivative  f'(z) = phi(z) (standard-normal PDF)
  //   f''(z) / f'(z) = -z  (ratio of PDF derivative to PDF itself)
  // giving z_new = z - (e / f'(z)) * (1 + (e / f'(z)) * z / 2). Two
  // iterations are harmless but unnecessary; one pass brings the
  // absolute error well below `1e-12` for p in (1e-300, 1 - 1e-16).
  for (int i = 0; i < 2; ++i) {
    const double phi_cdf = 0.5 * std::erfc(-z / std::sqrt(2.0));
    const double phi_pdf = std::exp(-0.5 * z * z) / std::sqrt(2.0 * kStatsPi);
    if (phi_pdf == 0.0) {
      break;
    }
    const double e = phi_cdf - p;
    const double u = e / phi_pdf;
    z -= u * (1.0 + 0.5 * z * u);
  }
  return z;
}

// ---------------------------------------------------------------------------
// Dispersion family: VAR.S / VAR.P / STDEV.S / STDEV.P. Each is a thin
// wrapper that collects the numeric slice and delegates to
// `variance_or_stdev`.
// ---------------------------------------------------------------------------

// VAR.S(value, ...) / VAR(value, ...) - sample variance with divisor n - 1.
// Fewer than 2 numeric inputs yields `#DIV/0!`.
static Value VarS(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  return variance_or_stdev(xs, true, false);
}

// VAR.P(value, ...) - population variance with divisor n. A single numeric
// input yields 0; no numeric inputs yields `#DIV/0!`.
static Value VarP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  return variance_or_stdev(xs, false, false);
}

// STDEV.S(value, ...) / STDEV(value, ...) - sample standard deviation,
// `sqrt(VAR.S)`. Fewer than 2 numeric inputs yields `#DIV/0!`.
static Value StdevS(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  return variance_or_stdev(xs, true, true);
}

// STDEV.P(value, ...) - population standard deviation, `sqrt(VAR.P)`.
static Value StdevP(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::vector<double> xs = collect_numerics(args, arity);
  return variance_or_stdev(xs, false, true);
}

// ---------------------------------------------------------------------------
// "A" family: AVERAGEA / MAXA / MINA / VARA / VARPA / STDEVA / STDEVPA.
// Differ from the non-A counterparts by evaluating text as 0 and Bool as
// 0 / 1 instead of skipping them. Range-sourced Blanks are still dropped;
// direct Blank arguments are counted as 0 (see `collect_a`). Registered
// with `range_filter_a_coerce = true` so the dispatcher performs the
// provenance-aware transform on range cells before the impl runs.
// ---------------------------------------------------------------------------

static Value AverageA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto collected = collect_a(args, arity);
  if (!collected) {
    return Value::error(collected.error());
  }
  const std::vector<double>& xs = collected.value();
  if (xs.empty()) {
    return Value::error(ErrorCode::Div0);
  }
  double total = 0.0;
  for (double x : xs) {
    total += x;
  }
  const double r = total / static_cast<double>(xs.size());
  return finite_number_result(r);
}

static Value extreme_a(const Value* args, std::uint32_t arity, bool want_max) {
  auto collected = collect_a(args, arity);
  if (!collected) {
    return Value::error(collected.error());
  }
  const std::vector<double>& xs = collected.value();
  if (xs.empty()) {
    return Value::number(0.0);
  }
  double best = xs[0];
  for (std::size_t i = 1; i < xs.size(); ++i) {
    if (want_max ? xs[i] > best : xs[i] < best) {
      best = xs[i];
    }
  }
  return finite_number_result(best);
}

static Value MaxA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return extreme_a(args, arity, true);
}

static Value MinA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return extreme_a(args, arity, false);
}

static Value VarA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return variance_or_stdev_a(args, arity, true, false);
}

static Value VarPA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return variance_or_stdev_a(args, arity, false, false);
}

static Value StdevA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return variance_or_stdev_a(args, arity, true, true);
}

static Value StdevPA(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return variance_or_stdev_a(args, arity, false, true);
}

}  // namespace stats_detail

void register_stats_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration range_stats[] = {
      {"MEDIAN", 1u, kVariadic, &stats_detail::Median, true, true},
      {"MODE", 1u, kVariadic, &stats_detail::Mode, true, true},
      {"MODE.SNGL", 1u, kVariadic, &stats_detail::Mode, true, true},
      {"MODE.MULT", 1u, kVariadic, &stats_detail::ModeMult, true, true},
      {"LARGE", 2u, kVariadic, &stats_detail::Large, true, true, true},
      {"SMALL", 2u, kVariadic, &stats_detail::Small, true, true, true},
      {"PERCENTILE.INC", 2u, kVariadic, &stats_detail::PercentileInc, true, true},
      {"PERCENTILE", 2u, kVariadic, &stats_detail::PercentileInc, true, true},
      {"PERCENTILE.EXC", 2u, kVariadic, &stats_detail::PercentileExc, true, true, true},
      {"QUARTILE.INC", 2u, kVariadic, &stats_detail::QuartileInc, true, true},
      {"QUARTILE", 2u, kVariadic, &stats_detail::QuartileInc, true, true},
      {"QUARTILE.EXC", 2u, kVariadic, &stats_detail::QuartileExc, true, true, true},
      {"STDEV.S", 1u, kVariadic, &stats_detail::StdevS, true, true},
      {"STDEV", 1u, kVariadic, &stats_detail::StdevS, true, true},
      {"STDEV.P", 1u, kVariadic, &stats_detail::StdevP, true, true},
      {"VAR.S", 1u, kVariadic, &stats_detail::VarS, true, true},
      {"VAR", 1u, kVariadic, &stats_detail::VarS, true, true},
      {"VAR.P", 1u, kVariadic, &stats_detail::VarP, true, true},
      {"VARP", 1u, kVariadic, &stats_detail::VarP, true, true},
      {"STDEVP", 1u, kVariadic, &stats_detail::StdevP, true, true},
      {"AVERAGEA", 1u, kVariadic, &stats_detail::AverageA, true, true, false, false, true},
      {"MAXA", 1u, kVariadic, &stats_detail::MaxA, true, true, false, false, true},
      {"MINA", 1u, kVariadic, &stats_detail::MinA, true, true, false, false, true},
      {"VARA", 1u, kVariadic, &stats_detail::VarA, true, true, false, false, true},
      {"VARPA", 1u, kVariadic, &stats_detail::VarPA, true, true, false, false, true},
      {"STDEVA", 1u, kVariadic, &stats_detail::StdevA, true, true, false, false, true},
      {"STDEVPA", 1u, kVariadic, &stats_detail::StdevPA, true, true, false, false, true},
      {"GEOMEAN", 1u, kVariadic, &stats_detail::GeoMean, true, true, true},
      {"HARMEAN", 1u, kVariadic, &stats_detail::HarMean, true, true, true},
      {"DEVSQ", 1u, kVariadic, &stats_detail::DevSq, true, true, true},
      {"AVEDEV", 1u, kVariadic, &stats_detail::AveDev, true, true, true},
      {"TRIMMEAN", 2u, kVariadic, &stats_detail::TrimMean, true, true, true},
      {"SKEW", 1u, kVariadic, &stats_detail::Skew, true, true, true},
      {"SKEW.P", 1u, kVariadic, &stats_detail::SkewP, true, true, true},
      {"KURT", 1u, kVariadic, &stats_detail::Kurt, true, true, true},
  };
  builtins_detail::register_builtin_functions(registry, range_stats, sizeof(range_stats) / sizeof(range_stats[0]));

  // Probability-distribution family. Scalar-only: `accepts_ranges` is left
  // at its default `false`, and `propagate_errors` stays `true` so the
  // dispatcher short-circuits error arguments before the impl runs. Impls
  // live in `stats/stats_distributions.cpp`.
  static constexpr builtins_detail::BuiltinRegistration scalar_stats[] = {
      {"STANDARDIZE", 3u, 3u, &stats_detail::Standardize},
      {"NORM.DIST", 4u, 4u, &stats_detail::NormDist},
      {"NORM.S.DIST", 2u, 2u, &stats_detail::NormSDist},
      {"NORM.INV", 3u, 3u, &stats_detail::NormInv},
      {"NORM.S.INV", 1u, 1u, &stats_detail::NormSInv},
      {"BINOM.DIST", 4u, 4u, &stats_detail::BinomDist},
      {"POISSON.DIST", 3u, 3u, &stats_detail::PoissonDist},
      {"EXPON.DIST", 3u, 3u, &stats_detail::ExponDist},
      {"CHISQ.DIST", 3u, 3u, &stats_detail::ChisqDist},
      {"CHISQ.DIST.RT", 2u, 2u, &stats_detail::ChisqDistRt},
      {"CHISQ.INV", 2u, 2u, &stats_detail::ChisqInv},
      {"CHISQ.INV.RT", 2u, 2u, &stats_detail::ChisqInvRt},
      {"T.DIST", 3u, 3u, &stats_detail::TDist},
      {"T.DIST.2T", 2u, 2u, &stats_detail::TDist2T},
      {"T.DIST.RT", 2u, 2u, &stats_detail::TDistRt},
      {"T.INV", 2u, 2u, &stats_detail::TInv},
      {"T.INV.2T", 2u, 2u, &stats_detail::TInv2T},
      {"F.DIST", 4u, 4u, &stats_detail::FDist},
      {"F.DIST.RT", 3u, 3u, &stats_detail::FDistRt},
      {"F.INV", 3u, 3u, &stats_detail::FInv},
      {"F.INV.RT", 3u, 3u, &stats_detail::FInvRt},
      {"NORMDIST", 4u, 4u, &stats_detail::NormDist},
      {"NORMINV", 3u, 3u, &stats_detail::NormInv},
      {"NORMSDIST", 1u, 1u, &stats_detail::NormSDistLegacy},
      {"NORMSINV", 1u, 1u, &stats_detail::NormSInv},
      {"BINOMDIST", 4u, 4u, &stats_detail::BinomDist},
      {"POISSON", 3u, 3u, &stats_detail::PoissonDist},
      {"EXPONDIST", 3u, 3u, &stats_detail::ExponDist},
      {"CHIDIST", 2u, 2u, &stats_detail::ChisqDistRt},
      {"CHIINV", 2u, 2u, &stats_detail::ChisqInvRt},
      {"FDIST", 3u, 3u, &stats_detail::FDistRt},
      {"FINV", 3u, 3u, &stats_detail::FInvRt},
      {"TDIST", 3u, 3u, &stats_detail::TDistLegacy},
      {"TINV", 2u, 2u, &stats_detail::TInv2T},
      {"CONFIDENCE", 3u, 3u, &stats_detail::ConfidenceNorm},
      {"CONFIDENCE.NORM", 3u, 3u, &stats_detail::ConfidenceNorm},
      {"CONFIDENCE.T", 3u, 3u, &stats_detail::ConfidenceT},
      {"BINOM.INV", 3u, 3u, &stats_detail::BinomInv},
      {"CRITBINOM", 3u, 3u, &stats_detail::BinomInv},
      {"BINOM.DIST.RANGE", 3u, 4u, &stats_detail::BinomDistRange},
      {"FISHER", 1u, 1u, &stats_detail::Fisher},
      {"FISHERINV", 1u, 1u, &stats_detail::FisherInv},
      {"GAUSS", 1u, 1u, &stats_detail::Gauss},
      {"PHI", 1u, 1u, &stats_detail::Phi},
      {"NEGBINOM.DIST", 4u, 4u, &stats_detail::NegBinomDist},
      {"NEGBINOMDIST", 3u, 3u, &stats_detail::NegBinomDistLegacy},
  };
  builtins_detail::register_builtin_functions(registry, scalar_stats, sizeof(scalar_stats) / sizeof(scalar_stats[0]));

  // The pairwise linear-regression family (CORREL, COVARIANCE.P,
  // COVARIANCE.S, SLOPE, INTERCEPT, RSQ, FORECAST / FORECAST.LINEAR)
  // is routed through the lazy dispatch table (see
  // `eval_*_lazy` in `regression_lazy.cpp`) because each array
  // argument must preserve its (rows, cols) shape so the two inputs
  // can be shape-matched cell-by-cell. Pre-evaluating every arg via
  // the eager path would erase that shape.
}

}  // namespace eval
}  // namespace formulon
