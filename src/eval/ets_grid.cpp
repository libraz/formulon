//
// Timeline step detection and regular-grid rebuild for the FORECAST.ETS
// family, behind `prepare_series`.

#include <algorithm>
#include <cfloat>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/ets_series.h"
#include "eval/ets_series_internal.h"
#include "eval/excel_math.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace ets {
namespace series_detail {

namespace {

constexpr double kStepTolerance = 1e-5;
// Percentage of the steps that must fall in the chosen histogram window.
constexpr std::int64_t kMinWindowPercent = 50;
constexpr std::int64_t kMaxHistogramBins = 3000;
constexpr std::int64_t kMaxRefinePasses = 11;
constexpr double kSameStepTolerance = 3e-7;
// Longest run of regular steps tracked when anchoring the grid.
constexpr std::size_t kMaxAnchorRun = 20U;
constexpr std::int64_t kMaxGridNodes = 1000000;

// -- step detection -----------------------------------------------------

// Working state of the step histogram.
struct Histogram {
  std::vector<std::int64_t> counts;
  std::vector<double> reps;
  std::vector<std::int64_t> same;
  std::int64_t nbins = 0;
  std::int64_t zeros = 0;
};

// Do [a, b] and [lo, hi] intersect (inclusive)?
bool ranges_overlap(double a, double b, double lo, double hi) {
  const auto pos = [&](double v) {
    if (bit_compare(v, a) == -1) {
      return -1;
    }
    return bit_compare(v, b) == 1 ? 1 : 0;
  };
  const int pa = pos(lo);
  const int pb = pos(hi);
  if (pa != pb) {
    return true;
  }
  return pa == 0;
}

// Number of steps whose bins intersect [lo, hi].
std::int64_t count_in(double low, double bw, const std::vector<std::int64_t>& counts, std::int64_t nbins, double lo,
                      double hi) {
  const auto index = [&](double v) -> std::int64_t {
    const double q = div_zero(v - low, bw);
    if (!top_bits_zero(q) && std::isfinite(q) && sign_bit_set(q)) {
      return -1;
    }
    return static_cast<std::int64_t>(to_uint32(q));
  };
  const std::int64_t i0 = index(lo);
  const std::int64_t i1 = index(hi);
  const std::int64_t start = i1 < 0 ? i0 : std::max<std::int64_t>(i0, 0);
  const std::int64_t last = (start < nbins && i1 >= nbins) ? nbins - 1 : i1;
  const std::int64_t first = std::max<std::int64_t>(start, 0);
  if (first > last || first >= nbins) {
    return 0;
  }
  std::int64_t sum = 0;
  for (std::int64_t j = first; j < nbins; ++j) {
    sum += elem(counts, j);
    if (j >= last) {
      break;
    }
  }
  return sum;
}

// Steps of about k years.
void classify_multiyear(Timeline* tl, double lo, double hi) {
  const std::int64_t k = static_cast<std::int64_t>(to_uint32(round_away(div_zero(lo, 365.0))));
  if (k < 2 || k > 100) {
    return;
  }
  const std::int64_t nominal = k * 365 + (k >> 2);
  // floor(log2(k)) + 2
  std::int64_t tol = 2;
  for (std::int64_t v = k; v > 1; v >>= 1) {
    ++tol;
  }
  if (ranges_overlap(lo, hi, static_cast<double>(nominal - tol), static_cast<double>(nominal + tol))) {
    tl->months = 12 * k;
  }
}

// Calendar class of a constant step; `m` is the number of steps.
void classify_constant(Timeline* tl, double step, std::int64_t m) {
  tl->step = step;
  tl->months = 0;
  tl->count = m;
  if (!tl->is_date) {
    return;
  }
  if (ranges_overlap(step, step, 30.0, 31.0)) {
    const std::int64_t lim = step == 29.5 ? 2 : 2 + (step > 29.5 ? 1 : 0);
    if (m > lim) {
      return;
    }
    tl->months = 1;
  } else if (ranges_overlap(step, step, 363.0, 367.0)) {
    tl->months = 12;
  } else if (ranges_overlap(step, step, 90.0, 92.0)) {
    const std::int64_t lim = step == 90.5 ? 2 : 2 + (step > 90.5 ? 1 : 0);
    if (m > lim) {
      return;
    }
    tl->months = 3;
  } else if (ranges_overlap(step, step, 181.5, 184.5)) {
    if (m > 2) {
      return;
    }
    tl->months = 6;
  } else if (bit_compare(step, 365.0) == 1) {
    classify_multiyear(tl, step, step);
  }
}

// Calendar class of the chosen histogram bin [lo, hi]; `rep` is the common
// step of a pure bin.
void classify_bin(Timeline* tl, double lo, double hi, bool has_rep, double rep, double low, double bw,
                  const std::vector<std::int64_t>& counts, std::int64_t nbins) {
  tl->step = has_rep ? rep : lo;
  tl->months = 0;
  if (!tl->is_date) {
    return;
  }
  const auto cnt = [&](double a, double b) { return count_in(low, bw, counts, nbins, a, b); };
  const auto set = [&](std::int64_t months) { tl->months = months; };
  if (ranges_overlap(lo, hi, 30.0, 31.0)) {
    const std::int64_t c1 = cnt(27.5, 31.5);
    if (c1 < 2) {
      return set(1);
    }
    const std::int64_t c2 = cnt(27.5, 29.5);
    const std::int64_t c3 = cnt(29.5, 30.5);
    const std::int64_t c4 = cnt(30.5, 31.5);
    const std::int64_t c5 = cnt(29.5, 31.5);
    if (c4 == 0 && (c3 > 2 || c2 > 2)) {
      return;
    }
    if (c3 == 0) {
      if (c4 > 4 || c2 > 2) {
        return;
      }
      return set(1);
    }
    if (c2 < 3) {
      return set(1);
    }
    if (c5 > 9) {
      return set(1);
    }
    if (c5 + c2 != c1) {
      return set(1);
    }
    return;
  }
  if (ranges_overlap(lo, hi, 363.0, 367.0)) {
    return set(12);
  }
  if (ranges_overlap(lo, hi, 90.0, 92.0)) {
    const std::int64_t c1 = cnt(89.5, 92.5);
    if (c1 < 2) {
      return set(3);
    }
    const std::int64_t c2 = cnt(89.5, 90.5);
    const std::int64_t c3 = cnt(90.5, 91.5);
    const std::int64_t c4 = cnt(91.5, 92.5);
    const std::int64_t c5 = cnt(90.5, 92.5);
    if (c3 == 0 && (c4 > 3 || c2 > 2)) {
      return;
    }
    if (c4 == 0) {
      if (c3 > 3 || c2 > 2) {
        return;
      }
      return set(3);
    }
    if (c2 < 3) {
      return set(3);
    }
    if (c5 > 1) {
      return set(3);
    }
    if (c5 + c2 == c1) {
      return;
    }
    return set(3);
  }
  if (ranges_overlap(lo, hi, 181.5, 184.5)) {
    const std::int64_t c1 = cnt(89.5, 92.5);
    if (c1 < 2) {
      return set(6);
    }
    const std::int64_t c2 = cnt(180.5, 181.5);
    const std::int64_t c3 = cnt(181.5, 182.5);
    const std::int64_t c4 = cnt(183.5, 184.5);
    if (c2 == 0 && (c4 > 2 || c3 > 2)) {
      return;
    }
    if (c4 == 0) {
      if (c2 > 2 || c3 > 2) {
        return;
      }
      return set(6);
    }
    if (c3 < 3) {
      return set(6);
    }
    const std::int64_t s = c4 + c2;
    if (s + c3 != c1) {
      return set(6);
    }
    if (s < 6) {
      return;
    }
    return set(6);
  }
  if (bit_compare(lo, 365.0) == 1) {
    classify_multiyear(tl, lo, hi);
  }
}

struct Binning {
  double bw = 0.0;
  std::int64_t nbins = 0;
  double bw0 = 0.0;
  bool capped = false;
};

// Histogram geometry for the full range of steps.
Binning make_binning(const StepStats& st) {
  double base = 0.0;
  if (st.min == 0.0) {
    base = kStepTolerance;
  } else {
    base = (bit_compare(kStepTolerance, st.min) == 1 && bit_compare(kStepTolerance, st.max) == -1) ? kStepTolerance
                                                                                                   : st.min;
  }
  base = div_zero(base, 5.0);
  const double bw0 = div_zero(base, 5.0);
  const double range = st.max - st.min;
  double q = div_zero(range, bw0);
  q = q + 1.0;
  Binning b;
  if (bit_compare(q, static_cast<double>(kMaxHistogramBins)) == 1) {
    b.bw = div_zero(range, static_cast<double>(kMaxHistogramBins - 1));
    b.nbins = kMaxHistogramBins;
    b.bw0 = bw0;
    b.capped = true;
    return b;
  }
  const std::uint32_t nb = to_uint32(q);
  b.bw = div_zero(range, static_cast<double>(nb));
  const double q2 = div_zero(range, b.bw) + 1.5;
  b.nbins = static_cast<std::int64_t>(to_uint32(q2));
  b.bw0 = b.bw;
  return b;
}

void fill_histogram(Histogram* h, const std::vector<double>& t, std::int64_t nbins, double bw, double low, double high,
                    bool bounded) {
  h->counts.assign(static_cast<std::size_t>(nbins), 0);
  h->reps.assign(static_cast<std::size_t>(nbins), 0.0);
  h->same.assign(static_cast<std::size_t>(nbins), 0);
  h->nbins = nbins;
  h->zeros = 0;
  for (std::size_t i = 1; i < t.size(); ++i) {
    const double d = t[i] - t[i - 1];
    if (top_bits_zero(d) || !std::isfinite(d)) {
      ++h->zeros;
      continue;
    }
    if (bounded && (bit_compare(d, low) == -1 || bit_compare(d, high) == 1)) {
      continue;
    }
    double q = d - low;
    q = div_zero(q, bw);
    std::int64_t b = static_cast<std::int64_t>(to_uint32(q));
    if (b >= nbins) {
      b = nbins - 1;
    }
    const std::int64_t c = elem(h->counts, b);
    elem(h->counts, b) = c + 1;
    if (c == 0) {
      elem(h->reps, b) = d;
      elem(h->same, b) = 1;
    } else if (std::fabs(elem(h->reps, b) - d) <= kSameStepTolerance) {
      elem(h->same, b) += 1;
    }
  }
}

// Picks the step bin: a sliding window over the histogram for wide
// histograms, the largest bin otherwise. Returns true when the bin is pure.
bool choose_bin(const Histogram& h, double low, double bw, Timeline* tl) {
  const std::vector<std::int64_t>& counts = h.counts;
  const std::int64_t nbins = h.nbins;
  std::int64_t b = 0;
  std::int64_t wtotal = 0;
  if (nbins <= 4) {
    std::int64_t total = 0;
    for (std::int64_t i = 0; i < nbins; ++i) {
      total += elem(counts, i);
    }
    std::int64_t best_i = 0;
    if (nbins >= 1 && total > 0) {
      for (std::int64_t i = 1; i < nbins; ++i) {
        if (elem(counts, i) > elem(counts, best_i)) {
          best_i = i;
        }
      }
    }
    b = best_i;
    wtotal = nbins >= 1 ? total : 0;
  } else {
    const double scaled_low = div_zero(low, bw);
    std::int64_t w8 = 0;
    std::int64_t w11 = 0;
    std::int64_t w10 = 0;
    std::int64_t w12 = 0;
    std::int64_t w13 = 1;
    std::int64_t w14 = 10;
    std::int64_t acc = elem(counts, 0);
    std::int64_t best = elem(counts, 0);
    while (true) {
      if (w13 < nbins) {
        acc += elem(counts, w13);
      }
      const std::int64_t mid = (w12 + w13) / 2;
      const double thr = static_cast<double>(w14 - (5 * w12 + mid));
      if (bit_compare(scaled_low, thr) == -1) {
        if (w12 >= 0) {
          acc -= elem(counts, w12);
        }
        ++w12;
      } else {
        --w12;
        if (w12 >= 0) {
          acc += elem(counts, w12);
        }
      }
      const std::int64_t mid2 = (w12 + w13) / 2;
      if (mid2 >= nbins) {
        break;
      }
      if (best < acc || (best == acc && elem(counts, w8) < elem(counts, mid2))) {
        best = acc;
        w8 = mid2;
        w11 = w12;
        w10 = w13;
      }
      ++w13;
      w14 += 5;
    }
    std::int64_t w28 = std::max<std::int64_t>(w11, 0);
    std::int64_t w27 = std::min<std::int64_t>(w10, nbins - 1);
    std::int64_t w25 = 0;
    std::int64_t w9 = 0;
    if (w28 <= w27) {
      w25 = w28;
      w9 = elem(counts, w28);
      for (std::int64_t i = w28 + 1; i <= w27; ++i) {
        if (w9 < elem(counts, i)) {
          w25 = i;
          w9 = elem(counts, i);
        }
      }
    }
    if (w25 != w8 && !(w9 < best - 1 && 15 * best <= 100 * (best - w9))) {
      double v = mul_flush(static_cast<double>(w25), bw) + low;
      v = div_zero(v, 10.0);
      v = div_zero(v, bw);
      const std::int64_t r = static_cast<std::int64_t>(to_uint32(v));
      w28 = std::max<std::int64_t>(w25 - r, 0);
      w27 = std::min<std::int64_t>(w25 + r, nbins - 1);
      wtotal = 0;
      for (std::int64_t i = w28; i <= w27; ++i) {
        wtotal += elem(counts, i);
      }
      b = w25;
    } else {
      b = w25 != w8 ? w8 : w25;
      wtotal = best;
    }
  }
  bool found = false;
  double rep = 0.0;
  if (elem(counts, b) >= 1 && elem(h.same, b) == elem(counts, b)) {
    rep = elem(h.reps, b);
    found = true;
  }
  const double lo = mul_flush(static_cast<double>(b), bw) + low;
  const double hi = lo + bw;
  classify_bin(tl, lo, hi, found, rep, low, bw, counts, nbins);
  tl->count = wtotal;
  return found;
}

// 4 * max(0.4 a, b).
double window_width(double a, double b) {
  double v = div_zero(a, 5.0);
  v = mul_flush(v, 2.0);
  if (bit_compare(v, b) == -1) {
    v = b;
  }
  return mul_flush(v, 4.0);
}

struct Refined {
  double low = 0.0;
  double high = 0.0;
  double bw = 0.0;
  std::int64_t nbins = 0;
  bool capped = false;
};

// Narrower histogram around the current step estimate.
Refined refine_histogram(const StepStats& st, double bw, double bw0, double step) {
  const double lowbase = st.min;
  const double width = window_width(step, bw);
  bool first = true;
  while (true) {
    const double v = div_zero(width, bw0);
    if (bit_compare(v, 2900.0) == -1) {
      break;
    }
    bw0 = mul_flush(bw0, bit_compare(v, 48000.0) == 1 ? 16.0 : 2.0);
    first = false;
  }
  double newlow = step;
  double v = div_zero(step - lowbase, bw0);
  if (!(bit_compare(v, -1.0) == -1)) {
    const std::uint32_t n0 = to_uint32(v);
    v = mul_flush(static_cast<double>(n0), bw0);
    newlow = v + lowbase;
  }
  const double w2 = window_width(newlow, bw);
  const std::uint32_t k = to_uint32(div_zero(w2, bw0));
  const double half = div_zero(mul_flush(static_cast<double>(k), bw0), 2.0);
  Refined r;
  r.low = newlow - half;
  r.high = newlow + half;
  if (bit_compare(r.low, lowbase) == -1) {
    r.low = lowbase;
  }
  if (bit_compare(r.high, st.max) == 1) {
    r.high = st.max;
  }
  r.nbins = static_cast<std::int64_t>(to_uint32(div_zero(r.high - r.low, bw0) + 1.0));
  r.bw = bw0;
  r.capped = first && bit_compare(step, bw0) == 1;
  return r;
}

}  // namespace

// Detects the timeline step. `*irregular` is set when the series has gaps or
// duplicate timestamps and must be rebuilt.
bool detect_step(const std::vector<double>& t, const StepStats& st, Timeline* tl, bool* irregular, ErrorCode* err) {
  const std::int64_t n = static_cast<std::int64_t>(t.size());
  tl->is_date = is_date_serial(t.front()) && is_date_serial(t.back());
  tl->months = tl->is_date ? -1 : 0;
  if (st.min == st.max || (st.nonpositive != 0 && st.max == 0.0)) {
    if (st.max == 0.0) {
      *err = ErrorCode::Num;
      return false;
    }
    classify_constant(tl, st.max, n - st.nonpositive - 1);
    *irregular = st.nonpositive != 0;
    return true;
  }
  const Binning bin = make_binning(st);
  double bw = bin.bw;
  std::int64_t nbins = bin.nbins;
  double bw0 = bin.bw0;
  bool capped = bin.capped;
  double low = st.min;
  Histogram h;
  fill_histogram(&h, t, nbins, bw, low, st.max, false);
  bool found = choose_bin(h, low, bw, tl);
  if (!found && capped) {
    std::int64_t pass = 2;
    while (true) {
      const Refined r = refine_histogram(st, bw, bw0, tl->step);
      low = r.low;
      bw = r.bw;
      nbins = r.nbins;
      capped = r.capped;
      fill_histogram(&h, t, nbins, bw, low, r.high, true);
      found = choose_bin(h, low, bw, tl);
      if (found || !capped) {
        break;
      }
      if (pass >= kMaxRefinePasses) {
        *err = ErrorCode::Num;
        return false;
      }
      bw0 = bw;
      ++pass;
    }
  }
  *irregular = tl->count != n - 1;
  const std::int64_t m = n - h.zeros;
  if (!(tl->count * 100 > (m - 1) * kMinWindowPercent)) {
    *err = ErrorCode::Num;
    return false;
  }
  return true;
}

namespace {

// -- grid ---------------------------------------------------------------

struct GridContext {
  const Timeline* tl;
  double step;
  AggregationMode agg;
  bool count_mode;
  int completion;
  double tol;
  double half;
};

// Maximum number of grid nodes.
std::int64_t grid_capacity(std::int64_t n, double first, double last, double step) {
  const double span = div_zero(last - first, step);
  const std::int64_t by_count = (50 * n) / 35;
  const double c = std::ceil(span);
  std::int64_t w8 = 0;
  if (bit_compare(c, static_cast<double>(kMaxGridNodes)) == 1) {
    w8 = std::max<std::int64_t>(by_count, kMaxGridNodes);
  } else {
    w8 = static_cast<std::int64_t>(to_uint32(c)) + 1;
  }
  w8 += w8 > 54 ? w8 / 5 : 10;
  std::int64_t w9 = std::min(by_count, w8);
  if (!(w8 < kMaxGridNodes)) {
    w9 = by_count;
  }
  return w9;
}

// Aggregates the points from `end` backwards while t >= `lower`; returns the
// value and the index of the group's first point.
bool aggregate_group(const GridContext& g, double lower, std::int64_t end, const std::vector<double>& t,
                     std::vector<double>* yv, double* value, std::int64_t* first, ErrorCode* err) {
  std::vector<double>& y = *yv;
  if (end <= 0 || bit_compare(lower, elem(t, end - 1)) == 1) {
    *value = g.count_mode ? 1.0 : elem(y, end);
    *first = end;
    return true;
  }
  std::int64_t j = end;
  switch (g.agg) {
    case AggregationMode::kAverage:
    case AggregationMode::kSum: {
      double acc = elem(y, end);
      std::int64_t cnt = 1;
      while (j > 0) {
        if (bit_compare(lower, elem(t, j - 1)) == 1) {
          break;
        }
        --j;
        acc = acc + elem(y, j);
        ++cnt;
      }
      *value = g.agg == AggregationMode::kAverage ? div_zero(acc, static_cast<double>(cnt)) : acc;
      break;
    }
    case AggregationMode::kCount:
    case AggregationMode::kCountA: {
      std::int64_t cnt = 1;
      while (j > 0) {
        if (bit_compare(lower, elem(t, j - 1)) == 1) {
          break;
        }
        --j;
        ++cnt;
      }
      *value = static_cast<double>(cnt);
      break;
    }
    case AggregationMode::kMax:
    case AggregationMode::kMin: {
      const int want = g.agg == AggregationMode::kMax ? -1 : 1;
      double val = elem(y, end);
      while (j > 0) {
        if (bit_compare(lower, elem(t, j - 1)) == 1) {
          break;
        }
        --j;
        if (bit_compare(val, elem(y, j)) == want) {
          val = elem(y, j);
        }
      }
      *value = val;
      break;
    }
    case AggregationMode::kMedian: {
      while (j > 0 && bit_compare(lower, elem(t, j - 1)) != 1) {
        --j;
      }
      const std::int64_t cnt = end - j + 1;
      if (cnt > 1) {
        sort_pairs(yv, nullptr, j, end);
      }
      if ((cnt & 1) != 0) {
        *value = elem(y, j + (cnt >> 1));
      } else {
        const std::int64_t m = j + (cnt >> 1);
        const double s = elem(y, m) + elem(y, m - 1);
        *value = div_zero(s, 2.0);
      }
      break;
    }
    default:
      *err = ErrorCode::Num;
      return false;
  }
  *first = j;
  return true;
}

// Index of the first of a run of equal timestamps ending at `j`.
std::int64_t group_start(const std::vector<double>& t, std::int64_t j) {
  while (j > 0 && elem(t, j - 1) == elem(t, j)) {
    --j;
  }
  return j;
}

// Index of the point nearest `node` among the groups from `j` backwards.
std::int64_t nearest_group(const std::vector<double>& t, std::int64_t j, std::int64_t floor_index, double node,
                           double* best_distance) {
  double bestd = DBL_MAX;
  std::int64_t at = j;
  while (j >= floor_index) {
    const double d = std::fabs(node - elem(t, j));
    if (bit_compare(d, bestd) >= 0) {
      break;
    }
    bestd = d;
    at = j;
    if (j == 0) {
      break;
    }
    j = group_start(t, j - 1);
  }
  *best_distance = bestd;
  return at;
}

struct GridStart {
  double node = 0.0;
  double value = 0.0;
  std::int64_t first = 0;
  double stored = 0.0;
};

// Last grid node: anchored on the run of regular steps that covers most of
// the series.
bool first_node(const GridContext& g, const std::vector<double>& t, std::vector<double>* y, GridStart* out,
                ErrorCode* err) {
  const std::int64_t last = static_cast<std::int64_t>(t.size()) - 1;
  const std::int64_t x23 = group_start(t, last);
  std::size_t maxk = 0;
  std::int64_t table[kMaxAnchorRun + 1];
  for (std::int64_t& e : table) {
    e = -1;
  }
  const auto record = [&](std::size_t run, std::int64_t runstart) {
    const std::size_t k = std::min(run, kMaxAnchorRun);
    if (table[k] < 0) {
      table[k] = runstart;
    }
    if (k > maxk) {
      maxk = k;
    }
  };
  if (x23 > 0) {
    std::size_t run = 0;
    std::int64_t runstart = -1;
    std::int64_t cur = x23;
    while (true) {
      const std::int64_t prev = group_start(t, cur - 1);
      double d = (elem(t, cur) - elem(t, prev)) - g.step;
      d = std::fabs(d);
      if (bit_compare(d, g.tol) == 1) {
        if (run >= 2 && run > maxk) {
          record(run, runstart);
        }
        runstart = -1;
        run = 0;
      } else if (run == 0) {
        run = 1;
        runstart = cur;
      } else {
        ++run;
        if (run >= kMaxAnchorRun) {
          if (table[kMaxAnchorRun] < 0) {
            table[kMaxAnchorRun] = runstart;
          }
          if (maxk <= kMaxAnchorRun - 1U) {
            runstart = -1;
            maxk = kMaxAnchorRun;
          } else {
            runstart = -1;
          }
        }
      }
      cur = prev;
      if (prev <= 0) {
        break;
      }
    }
    if (run >= 2) {
      if (run > maxk) {
        record(run, runstart);
      }
    } else if (run == 1) {
      if (maxk == 0) {
        if (table[2] < 0) {
          table[2] = runstart;
        }
        maxk = 2;
      }
    } else if (maxk == 0) {
      *err = ErrorCode::Num;
      return false;
    }
  } else if (maxk == 0) {
    *err = ErrorCode::Num;
    return false;
  }
  const double span = elem(t, last) - t.front();
  double best = 0.0;
  std::int64_t anchor = -1;
  for (std::size_t k = maxk; k >= 2U; --k) {
    const std::int64_t p = table[k];
    if (p < 0) {
      continue;
    }
    double v = div_zero(elem(t, last) - elem(t, p), span);
    v = 1.0 - v;
    v = mul_flush(v, static_cast<double>(k));
    if (bit_compare(v, best) == 1) {
      best = v;
      anchor = p;
    }
  }
  if (anchor < 0) {
    *err = ErrorCode::Value;
    return false;
  }
  const std::int64_t lastg = group_start(t, last);
  bool has_stored = false;
  double stored = 0.0;
  double cur_t = 0.0;
  if (lastg == anchor) {
    cur_t = elem(t, last);
  } else {
    if (g.tl->months >= 1) {
      std::int64_t unused = 0;
      if (!month_walk(g.tl->months, elem(t, anchor), elem(t, lastg), &cur_t, &unused, err)) {
        return false;
      }
    } else {
      double v = div_zero(elem(t, last) - elem(t, anchor), g.step);
      v = round_away(v);
      v = mul_flush(v, g.step);
      cur_t = v + elem(t, anchor);
    }
    double bestd = 0.0;
    const std::int64_t at = nearest_group(t, lastg, anchor, cur_t, &bestd);
    if (bit_compare(bestd, g.tol) == 1) {
      has_stored = true;
      stored = elem(t, at);
    } else {
      cur_t = elem(t, at);
    }
  }
  const double lower = cur_t - g.half;
  double value = 0.0;
  std::int64_t first = 0;
  if (!aggregate_group(g, lower, last, t, y, &value, &first, err)) {
    return false;
  }
  out->node = cur_t;
  out->value = value;
  out->first = first;
  out->stored = has_stored ? stored : cur_t;
  return true;
}

struct PrevNode {
  double node = 0.0;
  std::int64_t steps = 0;
  double value = 0.0;
  std::int64_t first = 0;
  double stored = 0.0;
};

// Node preceding `cur_t`, with the number of steps between them.
bool previous_node(const GridContext& g, const std::vector<double>& t, std::vector<double>* y, std::int64_t ptr,
                   double cur_t, PrevNode* out, ErrorCode* err) {
  const std::int64_t x22 = group_start(t, ptr - 1);
  const double t_prev = elem(t, x22);
  double node = 0.0;
  std::int64_t steps = 0;
  if (g.tl->months >= 1) {
    std::int64_t cnt = 0;
    if (!month_walk(g.tl->months, cur_t, t_prev, &node, &cnt, err)) {
      return false;
    }
    steps = -cnt;
  } else {
    const double tmp = cur_t - g.step;
    const double lower = tmp - g.half;
    if (bit_compare(lower, t_prev) != 1) {
      steps = 1;
      node = tmp;
    } else {
      double v = div_zero(cur_t - t_prev, g.step);
      v = round_away(v);
      steps = static_cast<std::int64_t>(to_uint32(v));
      v = mul_flush(v, g.step);
      node = cur_t - v;
    }
  }
  double bestd = 0.0;
  const std::int64_t at = nearest_group(t, x22, 0, node, &bestd);
  bool has_stored = false;
  double stored = 0.0;
  if (bit_compare(bestd, g.tol) == 1) {
    has_stored = true;
    stored = elem(t, at);
  } else {
    node = elem(t, at);
  }
  const double lower2 = node - g.half;
  double value = 0.0;
  std::int64_t first = 0;
  if (!aggregate_group(g, lower2, ptr - 1, t, y, &value, &first, err)) {
    return false;
  }
  out->node = node;
  out->steps = steps;
  out->value = value;
  out->first = first;
  out->stored = has_stored ? stored : node;
  return true;
}

}  // namespace

// Rebuilds the series on a regular grid, filling skipped nodes per the data
// completion option. `*missing` counts the filled nodes.
bool build_grid(std::vector<double>* tv, std::vector<double>* yv, const Timeline& tl, int completion,
                AggregationMode agg, std::vector<double>* out_t, std::vector<double>* out_y, std::int64_t* missing,
                ErrorCode* err) {
  const std::vector<double>& t = *tv;
  const std::int64_t n = static_cast<std::int64_t>(t.size());
  GridContext g{&tl,
                tl.step,
                agg,
                agg == AggregationMode::kCount || agg == AggregationMode::kCountA,
                completion,
                div_zero(tl.step, 10.0),
                div_zero(tl.step, 2.0)};
  const std::int64_t cap = grid_capacity(n, t.front(), t.back(), g.step);
  std::int64_t room = cap - 1;
  *missing = 0;
  const auto append = [&](double time, double value, bool real) {
    if (room < 0 || room > cap) {
      *err = ErrorCode::Num;
      return false;
    }
    out_t->push_back(time);
    out_y->push_back(value);
    if (!real) {
      ++*missing;
    }
    --room;
    return true;
  };
  GridStart start;
  if (!first_node(g, t, yv, &start, err)) {
    return false;
  }
  double cur_t = start.node;
  double cur_y = start.value;
  std::int64_t ptr = start.first;
  if (!append(start.stored, cur_y, true)) {
    return false;
  }
  while (ptr > 0) {
    PrevNode p;
    if (!previous_node(g, t, yv, ptr, cur_t, &p, err)) {
      return false;
    }
    if (p.steps >= 2) {
      if (tl.months > 0) {
        const double kd = static_cast<double>(p.steps);
        Ymd ymd = ymd_of(cur_t);
        std::int64_t gcount = p.steps;
        while (gcount >= 2) {
          --gcount;
          const double node_t = add_months(&ymd, -tl.months);
          double value = 0.0;
          if (completion == 1) {
            const double diff = cur_y - p.value;
            double v = static_cast<double>(gcount);
            v = mul_flush(v, diff);
            v = div_zero(v, kd);
            value = v + p.value;
          }
          if (!append(node_t, value, false)) {
            return false;
          }
        }
      } else {
        const double span = cur_t - p.node;
        std::int64_t gcount = p.steps;
        double node_t = cur_t;
        double value = 0.0;
        while (gcount >= 2) {
          node_t = node_t - g.step;
          if (completion == 1) {
            const double diff = cur_y - p.value;
            double v = node_t - p.node;
            v = mul_flush(v, diff);
            v = div_zero(v, span);
            value = v + p.value;
          }
          if (!append(node_t, value, false)) {
            return false;
          }
          --gcount;
        }
      }
    }
    if (!append(p.stored, p.value, true)) {
      return false;
    }
    cur_t = p.node;
    cur_y = p.value;
    ptr = p.first;
  }
  std::reverse(out_t->begin(), out_t->end());
  std::reverse(out_y->begin(), out_y->end());
  return true;
}

}  // namespace series_detail
}  // namespace ets
}  // namespace eval
}  // namespace formulon
