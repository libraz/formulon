//
// Implementation of the lazy-dispatch table seam. See
// `tree_walker_lazy_table.h` for rationale and contract.
//
// This TU exists to absorb the per-family `#include` fan-out so adding a
// new lazy function only rebuilds this file plus the family TU. The
// main evaluator (`tree_walker.cpp`) does NOT include any per-family
// lazy header; it reaches the table through `find_lazy_impl`.

#include "eval/tree_walker_lazy_table.h"

#include <cstddef>

#include "eval/aggregate_lazy.h"
#include "eval/areas_lazy.h"
#include "eval/builtins/aggregate.h"
#include "eval/builtins/text_format.h"
#include "eval/cell_lazy.h"
#include "eval/conditional_aggregates.h"
#include "eval/database_lazy.h"
#include "eval/datetime_lazy.h"
#include "eval/dynamic_array/anchor.h"
#include "eval/dynamic_array/filtering.h"
#include "eval/dynamic_array/indexing.h"
#include "eval/dynamic_array/layout.h"
#include "eval/dynamic_array/reshape.h"
#include "eval/financial_lazy.h"
#include "eval/forecast_ets_lazy.h"
#include "eval/getpivotdata_lazy.h"
#include "eval/groupby_pivotby/groupby.h"
#include "eval/groupby_pivotby/pivotby.h"
#include "eval/hypothesis_lazy.h"
#include "eval/info_lazy.h"
#include "eval/lambda_helpers_lazy.h"
#include "eval/linest_lazy.h"
#include "eval/lookups/classic.h"
#include "eval/lookups/xlookup.h"
#include "eval/matrix_ops_lazy.h"
#include "eval/phonetic_lazy.h"
#include "eval/rank_lazy.h"
#include "eval/reference/indirect.h"  // INDIRECT lazy impl
#include "eval/reference/offset.h"    // OFFSET lazy impl
#include "eval/regex_lazy.h"
#include "eval/regression_lazy.h"
#include "eval/series_sum_lazy.h"
#include "eval/shape_ops_lazy.h"
#include "eval/special_forms_lazy.h"
#include "eval/textsplit_lazy.h"
#include "eval/trimrange_lazy.h"
#include "eval/workdays_lazy.h"
#include "utils/strings.h"

namespace formulon {
namespace eval {
namespace {

struct LazyEntry {
  const char* name;  // canonical UPPERCASE
  LazyImpl impl;
  // Unannotated lazy forms are conservatively array-capable. A new lazy
  // implementation must opt into kScalar/kReduce/kBroadcast only after its
  // return paths have been checked; this prevents a catalog addition from
  // creating a silent partial-recalc false negative.
  LazyResultShape shape = LazyResultShape::kArray;
};

// The single source of truth for lazy-form routing. Sorted alphabetically
// by canonical UPPERCASE name so a quick visual diff catches accidental
// duplicates. Comments preserved verbatim from the prior in-place table.
constexpr LazyEntry kLazyDispatch[] = {
    {"AGGREGATE", &eval_aggregate_lazy, LazyResultShape::kReduce},
    // ANCHORARRAY is the OOXML internal encoding of the postfix `#`
    // spill operator. The xlsx-only `_xlfn.` prefix is stripped by
    // `strip_future_prefix`, so callers register the canonical bare name.
    // See `eval_anchorarray_lazy`.
    {"ANCHORARRAY", &eval_anchorarray_lazy, LazyResultShape::kArray},
    {"AND", &eval_and_lazy, LazyResultShape::kReduce},
    {"AREAS", &eval_areas_lazy, LazyResultShape::kReduce},
    {"ARRAYTOTEXT", &eval_arraytotext_lazy, LazyResultShape::kScalar},
    {"AVERAGEIF", &eval_averageif_lazy, LazyResultShape::kReduce},
    {"AVERAGEIFS", &eval_averageifs_lazy, LazyResultShape::kReduce},
    {"BYCOL", &eval_bycol_lazy, LazyResultShape::kArray},
    {"BYROW", &eval_byrow_lazy, LazyResultShape::kArray},
    {"CELL", &eval_cell_lazy, LazyResultShape::kArray},
    // CHITEST is the pre-2010 legacy spelling of CHISQ.TEST; same impl.
    {"CHISQ.TEST", &eval_chisq_test_lazy, LazyResultShape::kReduce},
    {"CHITEST", &eval_chisq_test_lazy, LazyResultShape::kReduce},
    {"CHOOSE", &eval_choose_lazy, LazyResultShape::kBroadcast},
    {"CHOOSECOLS", &eval_choosecols_lazy, LazyResultShape::kArray},
    {"CHOOSEROWS", &eval_chooserows_lazy, LazyResultShape::kArray},
    {"CODE", &eval_code_lazy, LazyResultShape::kScalar},
    {"COLUMN", &eval_column_lazy, LazyResultShape::kBroadcast},
    {"COLUMNS", &eval_columns_lazy, LazyResultShape::kReduce},
    {"CORREL", &eval_correl_lazy, LazyResultShape::kReduce},
    {"COUNT", &eval_count_lazy, LazyResultShape::kReduce},
    {"COUNTIF", &eval_countif_lazy, LazyResultShape::kReduce},
    {"COUNTIFS", &eval_countifs_lazy, LazyResultShape::kReduce},
    // COVAR is the pre-2010 legacy spelling of COVARIANCE.P; both compute
    // the population covariance with identical semantics.
    {"COVAR", &eval_covariance_p_lazy, LazyResultShape::kReduce},
    {"COVARIANCE.P", &eval_covariance_p_lazy, LazyResultShape::kReduce},
    {"COVARIANCE.S", &eval_covariance_s_lazy, LazyResultShape::kReduce},
    // Calendar family: date1904-sensitive functions share one lazy impl
    // (`eval_datetime_lazy`) so the workbook epoch reaches the calendar math.
    // WEEKNUM is served by `eval_weeknum_lazy` (it layers a Win365 quirk) and
    // is registered separately below.
    {"DATE", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"DATEDIF", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"DATEVALUE", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"DAVERAGE", &eval_daverage_lazy, LazyResultShape::kReduce},
    {"DAY", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"DAYS", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"DAYS360", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"DCOUNT", &eval_dcount_lazy, LazyResultShape::kReduce},
    {"DCOUNTA", &eval_dcounta_lazy, LazyResultShape::kReduce},
    {"DGET", &eval_dget_lazy, LazyResultShape::kReduce},
    {"DMAX", &eval_dmax_lazy, LazyResultShape::kReduce},
    {"DMIN", &eval_dmin_lazy, LazyResultShape::kReduce},
    {"DPRODUCT", &eval_dproduct_lazy, LazyResultShape::kReduce},
    {"DROP", &eval_drop_lazy, LazyResultShape::kArray},
    {"DSTDEV", &eval_dstdev_lazy, LazyResultShape::kReduce},
    {"DSTDEVP", &eval_dstdevp_lazy, LazyResultShape::kReduce},
    {"DSUM", &eval_dsum_lazy, LazyResultShape::kReduce},
    {"DVAR", &eval_dvar_lazy, LazyResultShape::kReduce},
    {"DVARP", &eval_dvarp_lazy, LazyResultShape::kReduce},
    {"EDATE", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"EOMONTH", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"EXPAND", &eval_expand_lazy, LazyResultShape::kArray},
    {"F.TEST", &eval_f_test_lazy, LazyResultShape::kReduce},
    {"FILTER", &eval_filter_lazy, LazyResultShape::kArray},
    // FORECAST is the legacy spelling kept by Excel for back-compat;
    // its impl and arity are identical to FORECAST.LINEAR.
    {"FORECAST", &eval_forecast_linear_lazy, LazyResultShape::kScalar},
    // The FORECAST.ETS family rides the lazy seam because both the
    // values and timeline arguments may be Range refs that must reach
    // the impl with their (rows, cols) shape preserved for pairing.
    {"FORECAST.ETS", &eval_forecast_ets_lazy, LazyResultShape::kScalar},
    {"FORECAST.ETS.CONFINT", &eval_forecast_ets_confint_lazy, LazyResultShape::kScalar},
    {"FORECAST.ETS.SEASONALITY", &eval_forecast_ets_seasonality_lazy, LazyResultShape::kScalar},
    {"FORECAST.ETS.STAT", &eval_forecast_ets_stat_lazy, LazyResultShape::kScalar},
    {"FORECAST.LINEAR", &eval_forecast_linear_lazy, LazyResultShape::kScalar},
    // FORMULATEXT returns the source text of the referenced cell's formula,
    // so it must inspect the un-evaluated Ref AST and the bound Sheet's
    // `formula_text` directly — the eager path would flatten the argument
    // to a Value before we could see the reference.
    {"FORMULATEXT", &eval_formulatext_lazy, LazyResultShape::kScalar},
    {"FREQUENCY", &eval_frequency_lazy, LazyResultShape::kArray},
    // FTEST is the pre-2010 legacy spelling of F.TEST; same impl.
    {"FTEST", &eval_f_test_lazy, LazyResultShape::kReduce},
    // GETPIVOTDATA recovers the (sheet, row, col) anchor of its second
    // argument from the un-evaluated Reference AST and reads the
    // freshest `PivotResult` off the bound Workbook on EvalContext;
    // the eager dispatcher would flatten the anchor to a Value before
    // the impl could see the reference.
    {"GETPIVOTDATA", &eval_getpivotdata_lazy, LazyResultShape::kScalar},
    {"GROUPBY", &eval_groupby_lazy, LazyResultShape::kArray},
    {"GROWTH", &eval_growth_lazy, LazyResultShape::kArray},
    {"HLOOKUP", &eval_hlookup_lazy, LazyResultShape::kArray},
    {"HSTACK", &eval_hstack_lazy, LazyResultShape::kArray},
    {"IF", &eval_if_lazy, LazyResultShape::kBroadcast},
    {"IFERROR", &eval_iferror_lazy, LazyResultShape::kBroadcast},
    {"IFNA", &eval_ifna_lazy, LazyResultShape::kBroadcast},
    {"IFS", &eval_ifs_lazy, LazyResultShape::kBroadcast},
    {"INDEX", &eval_index_lazy, LazyResultShape::kArray},
    {"INDIRECT", &eval_indirect_lazy, LazyResultShape::kArray},
    {"INTERCEPT", &eval_intercept_lazy, LazyResultShape::kScalar},
    // ISFORMULA / ISREF inspect the un-evaluated AST of their argument;
    // they cannot ride the eager path because it flattens references to
    // `Value` before the impl runs.
    {"IRR", &eval_irr_lazy, LazyResultShape::kScalar},
    {"ISFORMULA", &eval_isformula_lazy, LazyResultShape::kScalar},
    {"ISOMITTED", &eval_isomitted_lazy, LazyResultShape::kScalar},
    {"ISOWEEKNUM", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"ISREF", &eval_isref_lazy, LazyResultShape::kScalar},
    {"LENB", &eval_lenb_lazy, LazyResultShape::kScalar},
    {"LINEST", &eval_linest_lazy, LazyResultShape::kArray},
    {"LOGEST", &eval_logest_lazy, LazyResultShape::kArray},
    {"LOOKUP", &eval_lookup_lazy, LazyResultShape::kScalar},
    {"MAKEARRAY", &eval_makearray_lazy, LazyResultShape::kArray},
    {"MAP", &eval_map_lazy, LazyResultShape::kArray},
    {"MATCH", &eval_match_lazy, LazyResultShape::kArray},
    {"MAXIFS", &eval_maxifs_lazy, LazyResultShape::kReduce},
    {"MDETERM", &eval_mdeterm_lazy, LazyResultShape::kScalar},
    {"MINIFS", &eval_minifs_lazy, LazyResultShape::kReduce},
    {"MINVERSE", &eval_minverse_lazy, LazyResultShape::kArray},
    {"MIRR", &eval_mirr_lazy, LazyResultShape::kScalar},
    {"MMULT", &eval_mmult_lazy, LazyResultShape::kArray},
    {"MONTH", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"NETWORKDAYS", &eval_networkdays_lazy, LazyResultShape::kScalar},
    {"NETWORKDAYS.INTL", &eval_networkdays_intl_lazy, LazyResultShape::kScalar},
    {"NOW", &eval_datetime_lazy, LazyResultShape::kScalar},
    {"OFFSET", &eval_offset_lazy, LazyResultShape::kArray},
    {"OR", &eval_or_lazy, LazyResultShape::kReduce},
    // PEARSON is mathematically identical to CORREL (Pearson product-moment
    // correlation coefficient); Excel keeps both names for back-compat.
    {"PEARSON", &eval_correl_lazy, LazyResultShape::kReduce},
    {"PERCENTOF", &eval_percentof_lazy, LazyResultShape::kScalar},
    {"PERCENTRANK", &eval_percentrank_inc_lazy, LazyResultShape::kScalar},
    {"PERCENTRANK.EXC", &eval_percentrank_exc_lazy, LazyResultShape::kScalar},
    {"PERCENTRANK.INC", &eval_percentrank_inc_lazy, LazyResultShape::kScalar},
    // PHONETIC reads the cell's <rPh> annotation off the un-evaluated
    // Ref AST; the eager dispatcher would flatten the argument to a
    // Value before the impl could consult `Cell::phonetic_text`.
    {"PHONETIC", &eval_phonetic_lazy, LazyResultShape::kScalar},
    {"PIVOTBY", &eval_pivotby_lazy, LazyResultShape::kArray},
    {"PROB", &eval_prob_lazy, LazyResultShape::kScalar},
    {"RANK", &eval_rank_eq_lazy, LazyResultShape::kScalar},
    {"RANK.AVG", &eval_rank_avg_lazy, LazyResultShape::kScalar},
    {"RANK.EQ", &eval_rank_eq_lazy, LazyResultShape::kScalar},
    {"REDUCE", &eval_reduce_lazy, LazyResultShape::kArray},
    {"REGEXEXTRACT", &eval_regexextract_lazy, LazyResultShape::kArray},
    {"REGEXREPLACE", &eval_regexreplace_lazy, LazyResultShape::kBroadcast},
    {"REGEXTEST", &eval_regextest_lazy, LazyResultShape::kBroadcast},
    {"ROW", &eval_row_lazy, LazyResultShape::kBroadcast},
    {"ROWS", &eval_rows_lazy, LazyResultShape::kReduce},
    {"RSQ", &eval_rsq_lazy, LazyResultShape::kScalar},
    {"SCAN", &eval_scan_lazy, LazyResultShape::kArray},
    {"SERIESSUM", &eval_series_sum_lazy, LazyResultShape::kScalar},
    // SHEET / SHEETS consult the bound Workbook + current Sheet on the
    // EvalContext; AST introspection of an optional reference argument
    // tells them which sheet to answer for.
    {"SHEET", &eval_sheet_lazy, LazyResultShape::kScalar},
    {"SHEETS", &eval_sheets_lazy, LazyResultShape::kReduce},
    // SINGLE is the explicit-name form of the `@` implicit-intersection
    // operator; xlsx serialises `@range` as `_xlfn.SINGLE(range)` (the
    // `_xlfn.` prefix is stripped by `strip_future_prefix`). Routes to a
    // lazy impl so the un-evaluated RangeOp AST can be projected onto the
    // formula cell's row / column via `EvalContext::formula_row` /
    // `formula_col` — the same logic the `@` operator uses.
    {"SINGLE", &eval_single_lazy, LazyResultShape::kScalar},
    {"SLOPE", &eval_slope_lazy, LazyResultShape::kScalar},
    {"SORT", &eval_sort_lazy, LazyResultShape::kArray},
    {"SORTBY", &eval_sortby_lazy, LazyResultShape::kArray},
    {"STEYX", &eval_steyx_lazy, LazyResultShape::kScalar},
    // SUBTOTAL is lazy so codes 101..111 can read row visibility off the
    // sheet; the eager registration stays for the bytecode VM. See
    // `eval_subtotal_lazy`.
    {"SUBTOTAL", &eval_subtotal_lazy, LazyResultShape::kReduce},
    {"SUMIF", &eval_sumif_lazy, LazyResultShape::kReduce},
    {"SUMIFS", &eval_sumifs_lazy, LazyResultShape::kReduce},
    {"SUMPRODUCT", &eval_sumproduct_lazy, LazyResultShape::kReduce},
    {"SUMX2MY2", &eval_sumx2my2_lazy, LazyResultShape::kReduce},
    {"SUMX2PY2", &eval_sumx2py2_lazy, LazyResultShape::kReduce},
    {"SUMXMY2", &eval_sumxmy2_lazy, LazyResultShape::kReduce},
    {"SWITCH", &eval_switch_lazy, LazyResultShape::kBroadcast},
    {"T.TEST", &eval_t_test_lazy, LazyResultShape::kReduce},
    {"TAKE", &eval_take_lazy, LazyResultShape::kArray},
    {"TEXT", &eval_text_lazy, LazyResultShape::kBroadcast},
    {"TEXTSPLIT", &eval_textsplit_lazy, LazyResultShape::kArray},
    {"TOCOL", &eval_tocol_lazy, LazyResultShape::kArray},
    {"TODAY", &eval_datetime_lazy, LazyResultShape::kScalar},
    {"TOROW", &eval_torow_lazy, LazyResultShape::kArray},
    {"TRANSPOSE", &eval_transpose_lazy, LazyResultShape::kArray},
    {"TREND", &eval_trend_lazy, LazyResultShape::kArray},
    {"TRIMRANGE", &eval_trimrange_lazy, LazyResultShape::kArray},
    // TTEST is the pre-2010 legacy spelling of T.TEST; same impl.
    {"TTEST", &eval_t_test_lazy, LazyResultShape::kReduce},
    {"UNIQUE", &eval_unique_lazy, LazyResultShape::kArray},
    {"VLOOKUP", &eval_vlookup_lazy, LazyResultShape::kArray},
    {"VSTACK", &eval_vstack_lazy, LazyResultShape::kArray},
    {"WEEKDAY", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"WEEKNUM", &eval_weeknum_lazy, LazyResultShape::kScalar},
    {"WORKDAY", &eval_workday_lazy, LazyResultShape::kScalar},
    {"WORKDAY.INTL", &eval_workday_intl_lazy, LazyResultShape::kScalar},
    {"WRAPCOLS", &eval_wrapcols_lazy, LazyResultShape::kArray},
    {"WRAPROWS", &eval_wraprows_lazy, LazyResultShape::kArray},
    {"XIRR", &eval_xirr_lazy, LazyResultShape::kScalar},
    {"XLOOKUP", &eval_xlookup_lazy, LazyResultShape::kArray},
    {"XMATCH", &eval_xmatch_lazy, LazyResultShape::kArray},
    {"XNPV", &eval_xnpv_lazy, LazyResultShape::kScalar},
    {"YEAR", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"YEARFRAC", &eval_datetime_lazy, LazyResultShape::kBroadcast},
    {"Z.TEST", &eval_z_test_lazy, LazyResultShape::kReduce},
    // ZTEST is the pre-2010 legacy spelling of Z.TEST; same impl.
    {"ZTEST", &eval_z_test_lazy, LazyResultShape::kReduce},
};

constexpr std::size_t kLazyDispatchCount = sizeof(kLazyDispatch) / sizeof(kLazyDispatch[0]);

constexpr int compare_canonical_names(const char* lhs, const char* rhs) {
  for (std::size_t i = 0;; ++i) {
    if (lhs[i] < rhs[i]) {
      return -1;
    }
    if (lhs[i] > rhs[i]) {
      return 1;
    }
    if (lhs[i] == '\0') {
      return 0;
    }
  }
}

constexpr bool lazy_dispatch_is_strictly_sorted() {
  for (std::size_t i = 1; i < kLazyDispatchCount; ++i) {
    if (compare_canonical_names(kLazyDispatch[i - 1].name, kLazyDispatch[i].name) >= 0) {
      return false;
    }
  }
  return true;
}

static_assert(lazy_dispatch_is_strictly_sorted(), "kLazyDispatch must stay in canonical-name order");

}  // namespace

LazyImpl find_lazy_impl(std::string_view name) noexcept {
  std::size_t first = 0;
  std::size_t last = kLazyDispatchCount;
  while (first < last) {
    const std::size_t middle = first + (last - first) / 2;
    const int cmp = strings::case_insensitive_compare(name, kLazyDispatch[middle].name);
    if (cmp == 0) {
      return kLazyDispatch[middle].impl;
    }
    if (cmp < 0) {
      last = middle;
    } else {
      first = middle + 1;
    }
  }
  return nullptr;
}

LazyResultShape find_lazy_result_shape(std::string_view name) noexcept {
  std::size_t first = 0;
  std::size_t last = kLazyDispatchCount;
  while (first < last) {
    const std::size_t middle = first + (last - first) / 2;
    const int cmp = strings::case_insensitive_compare(name, kLazyDispatch[middle].name);
    if (cmp == 0) {
      return kLazyDispatch[middle].shape;
    }
    if (cmp < 0) {
      last = middle;
    } else {
      first = middle + 1;
    }
  }
  return LazyResultShape::kNotLazy;
}

const char* const* lazy_table_names() {
  // Storage has static duration and is initialized exactly once
  // (C++11 magic statics); the lambda builds the nullptr-terminated view
  // from the adjacent dispatch table. No allocation, no ordering concerns.
  static const char* const* kTable = [] {
    static const char* storage[kLazyDispatchCount + 1] = {};
    for (std::size_t i = 0; i < kLazyDispatchCount; ++i) {
      storage[i] = kLazyDispatch[i].name;
    }
    storage[kLazyDispatchCount] = nullptr;
    return static_cast<const char* const*>(storage);
  }();
  return kTable;
}

}  // namespace eval
}  // namespace formulon
