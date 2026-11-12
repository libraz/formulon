//
// Series preparation for the FORECAST.ETS family; the contract is declared in
// `eval/ets_series.h`.

#include "eval/ets_series.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/ets_series_internal.h"
#include "eval/excel_math.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace ets {

namespace series_detail {

// Sorts `keys` ascending over [lo, hi], permuting `vals` alongside. Runs of
// at most eight elements use a selection pass; larger ones are partitioned
// around a median-of-three pivot with an explicit stack. The order of equal
// keys is part of the contract: it fixes the summation order of grouped
// values.
void sort_pairs(std::vector<double>* keys, std::vector<double>* vals, std::int64_t lo, std::int64_t hi) {
  std::vector<double>& k = *keys;
  const auto cmp = [&](std::int64_t i, std::int64_t j) { return bit_compare(elem(k, i), elem(k, j)); };
  const auto swap_at = [&](std::int64_t i, std::int64_t j) {
    std::swap(elem(k, i), elem(k, j));
    if (vals != nullptr) {
      std::swap(elem(*vals, i), elem(*vals, j));
    }
  };
  std::vector<std::pair<std::int64_t, std::int64_t>> stack;
  while (true) {
    const std::int64_t size = hi - lo + 1;
    if (size <= 8) {
      while (hi > lo) {
        std::int64_t mx = lo;
        for (std::int64_t p = lo + 1; p <= hi; ++p) {
          if (cmp(p, mx) > 0) {
            mx = p;
          }
        }
        if (mx != hi) {
          swap_at(mx, hi);
        }
        --hi;
      }
    } else {
      std::int64_t mid = lo + size / 2;
      if (cmp(lo, mid) > 0) {
        swap_at(lo, mid);
      }
      if (cmp(lo, hi) > 0) {
        swap_at(lo, hi);
      }
      if (cmp(mid, hi) > 0) {
        swap_at(mid, hi);
      }
      std::int64_t loguy = lo;
      std::int64_t higuy = hi;
      while (true) {
        if (mid > loguy) {
          ++loguy;
          while (loguy < mid && cmp(loguy, mid) <= 0) {
            ++loguy;
          }
        }
        if (mid <= loguy) {
          ++loguy;
          while (loguy <= hi && cmp(loguy, mid) <= 0) {
            ++loguy;
          }
        }
        --higuy;
        while (higuy > mid && cmp(higuy, mid) > 0) {
          --higuy;
        }
        if (higuy < loguy) {
          break;
        }
        swap_at(loguy, higuy);
        if (higuy == mid) {
          mid = loguy;
        }
      }
      ++higuy;
      if (mid < higuy) {
        --higuy;
        while (higuy > mid && cmp(higuy, mid) == 0) {
          --higuy;
        }
      }
      if (mid >= higuy) {
        --higuy;
        while (higuy > lo && cmp(higuy, mid) == 0) {
          --higuy;
        }
      }
      if (higuy - lo >= hi - loguy) {
        if (lo < higuy) {
          stack.emplace_back(lo, higuy);
        }
        if (loguy < hi) {
          lo = loguy;
          continue;
        }
      } else {
        if (loguy < hi) {
          stack.emplace_back(loguy, hi);
        }
        if (lo < higuy) {
          hi = higuy;
          continue;
        }
      }
    }
    if (!stack.empty()) {
      lo = stack.back().first;
      hi = stack.back().second;
      stack.pop_back();
      continue;
    }
    break;
  }
}

}  // namespace series_detail

using namespace series_detail;

namespace {

// Largest share (percent) of grid nodes that may be missing.
// @limit: internal
constexpr std::int64_t kMaxMissingPercent = 30;

// Pairs the cells in order. A pair with a blank, text or boolean timeline or
// value cell is dropped, except that an error cell propagates, text values
// count as 1 under COUNTA and blank values read as 0 without interpolation.
// Under COUNT/COUNTA every kept value is 1.
bool read_pairs(const std::vector<SeriesCell>& values, const std::vector<SeriesCell>& times, int completion,
                AggregationMode agg, std::vector<double>* t, std::vector<double>* y, StepStats* stats, ErrorCode* err) {
  if (values.size() != times.size()) {
    *err = ErrorCode::NA;
    return false;
  }
  const bool count_mode = agg == AggregationMode::kCount || agg == AggregationMode::kCountA;
  for (std::size_t i = 0; i < values.size(); ++i) {
    const SeriesCell& tc = times[i];
    SeriesCell yc = values[i];
    if (tc.kind == SeriesCell::Kind::kError) {
      *err = tc.error;
      return false;
    }
    if (yc.kind == SeriesCell::Kind::kError) {
      *err = yc.error;
      return false;
    }
    if (tc.kind != SeriesCell::Kind::kNumber) {
      continue;
    }
    if (yc.kind == SeriesCell::Kind::kBlank && completion == 0) {
      yc.kind = SeriesCell::Kind::kNumber;
      yc.number = 0.0;
    }
    double yv = 0.0;
    if (yc.kind == SeriesCell::Kind::kText) {
      if (agg != AggregationMode::kCountA) {
        continue;
      }
      yv = 1.0;
    } else if (yc.kind == SeriesCell::Kind::kNumber) {
      yv = count_mode ? 1.0 : yc.number;
    } else {
      continue;
    }
    if (!t->empty()) {
      stats->add(tc.number - t->back());
    }
    t->push_back(tc.number);
    y->push_back(yv);
  }
  return true;
}

}  // namespace

bool prepare_series(const std::vector<SeriesCell>& values, const std::vector<SeriesCell>& timeline, int data_completion,
                    AggregationMode aggregation, PreparedSeries* out, ErrorCode* err) {
  std::vector<double> t;
  std::vector<double> y;
  StepStats stats;
  if (!read_pairs(values, timeline, data_completion, aggregation, &t, &y, &stats, err)) {
    return false;
  }
  if (t.size() <= 1U) {
    *err = ErrorCode::Div0;
    return false;
  }
  if (stats.needs_sort) {
    sort_pairs(&t, &y, 0, static_cast<std::int64_t>(t.size()) - 1);
    stats = StepStats();
    for (std::size_t i = 1; i < t.size(); ++i) {
      stats.add(t[i] - t[i - 1]);
    }
    if (stats.needs_sort) {
      *err = ErrorCode::Value;
      return false;
    }
  }
  Timeline tl;
  bool irregular = false;
  if (!detect_step(t, stats, &tl, &irregular, err)) {
    return false;
  }
  std::vector<double> t2;
  std::vector<double> y2;
  if (irregular) {
    std::int64_t missing = 0;
    if (!build_grid(&t, &y, tl, data_completion, aggregation, &t2, &y2, &missing, err)) {
      return false;
    }
    if (kMaxMissingPercent * static_cast<std::int64_t>(t2.size()) < 100 * missing) {
      *err = ErrorCode::Num;
      return false;
    }
  } else {
    t2 = std::move(t);
    y2 = std::move(y);
  }
  if (t2.size() <= 1U) {
    *err = ErrorCode::Div0;
    return false;
  }
  tl.last = t2.back();
  out->y = std::move(y2);
  out->timeline = tl;
  return true;
}

bool target_position(const Timeline& tl, double target, double* position, bool* interpolated, ErrorCode* err) {
  *interpolated = false;
  if (tl.step == 0.0) {
    *err = ErrorCode::Value;
    return false;
  }
  if (tl.months >= 1) {
    if (!whole_count(std::floor(target))) {
      *err = ErrorCode::Num;
      return false;
    }
    Ymd ymd = ymd_of(tl.last);
    const Ymd at = ymd_of(target);
    // The calendar month difference puts the first node within a step or two
    // of the target, so the walks below stay short.
    std::int64_t k = floor_div((at.year - ymd.year) * 12 + (at.month - ymd.month), tl.months);
    double node = add_months(&ymd, tl.months * k);
    while (bit_compare(node, target) == 1) {
      --k;
      node = add_months(&ymd, -tl.months);
    }
    if (node == target) {
      *position = static_cast<double>(k);
      return true;
    }
    std::int64_t extra = 0;
    double hi = node;
    while (bit_compare(hi, target) == -1) {
      hi = add_months(&ymd, tl.months);
      --extra;
    }
    if (hi == target) {
      *position = static_cast<double>(k - extra);
      return true;
    }
    if (extra != -1) {
      k += (-extra - 1);
      node = add_months(&ymd, -tl.months);
    }
    const double span = hi - node;
    double v = target - node;
    v = div_zero(v, span);
    *position = v + static_cast<double>(k);
    *interpolated = true;
    return true;
  }
  const double pos = div_zero(target - tl.last, tl.step);
  const double r = round_away(pos);
  if (r == pos) {
    *position = pos;
    return true;
  }
  double d = std::fabs(r - pos);
  d = mul_flush(d, 1000.0);
  if (bit_compare(d, 1.0) == 1) {
    *position = pos;
    *interpolated = true;
    return true;
  }
  *position = r;
  return true;
}

}  // namespace ets
}  // namespace eval
}  // namespace formulon
