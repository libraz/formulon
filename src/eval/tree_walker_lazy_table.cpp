// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
};

// The single source of truth for lazy-form routing. Sorted alphabetically
// by canonical UPPERCASE name so a quick visual diff catches accidental
// duplicates. Comments preserved verbatim from the prior in-place table.
constexpr LazyEntry kLazyDispatch[] = {
    {"AGGREGATE", &eval_aggregate_lazy},
    {"AND", &eval_and_lazy},
    // ANCHORARRAY is the OOXML internal encoding of the postfix `#`
    // spill operator. The xlsx-only `_xlfn.` prefix is stripped by
    // `strip_future_prefix`, so callers register the canonical bare name.
    // See `eval_anchorarray_lazy`.
    {"ANCHORARRAY", &eval_anchorarray_lazy},
    {"ARRAYTOTEXT", &eval_arraytotext_lazy},
    {"AREAS", &eval_areas_lazy},
    {"AVERAGEIF", &eval_averageif_lazy},
    {"AVERAGEIFS", &eval_averageifs_lazy},
    {"BYCOL", &eval_bycol_lazy},
    {"BYROW", &eval_byrow_lazy},
    {"CELL", &eval_cell_lazy},
    // CHITEST is the pre-2010 legacy spelling of CHISQ.TEST; same impl.
    {"CHISQ.TEST", &eval_chisq_test_lazy},
    {"CHITEST", &eval_chisq_test_lazy},
    {"CHOOSE", &eval_choose_lazy},
    {"CHOOSECOLS", &eval_choosecols_lazy},
    {"CHOOSEROWS", &eval_chooserows_lazy},
    {"CODE", &eval_code_lazy},
    {"COLUMN", &eval_column_lazy},
    {"COLUMNS", &eval_columns_lazy},
    {"CORREL", &eval_correl_lazy},
    {"COUNT", &eval_count_lazy},
    {"COUNTIF", &eval_countif_lazy},
    {"COUNTIFS", &eval_countifs_lazy},
    // COVAR is the pre-2010 legacy spelling of COVARIANCE.P; both compute
    // the population covariance with identical semantics.
    {"COVAR", &eval_covariance_p_lazy},
    {"COVARIANCE.P", &eval_covariance_p_lazy},
    {"COVARIANCE.S", &eval_covariance_s_lazy},
    // Calendar family: date1904-sensitive functions share one lazy impl
    // (`eval_datetime_lazy`) so the workbook epoch reaches the calendar math.
    // WEEKNUM is served by `eval_weeknum_lazy` (it layers a Win365 quirk) and
    // is registered separately below.
    {"DATE", &eval_datetime_lazy},
    {"DATEDIF", &eval_datetime_lazy},
    {"DATEVALUE", &eval_datetime_lazy},
    {"DAY", &eval_datetime_lazy},
    {"DAYS360", &eval_datetime_lazy},
    {"DAVERAGE", &eval_daverage_lazy},
    {"DCOUNT", &eval_dcount_lazy},
    {"DCOUNTA", &eval_dcounta_lazy},
    {"DGET", &eval_dget_lazy},
    {"DMAX", &eval_dmax_lazy},
    {"DMIN", &eval_dmin_lazy},
    {"DPRODUCT", &eval_dproduct_lazy},
    {"DROP", &eval_drop_lazy},
    {"DSTDEV", &eval_dstdev_lazy},
    {"DSTDEVP", &eval_dstdevp_lazy},
    {"DSUM", &eval_dsum_lazy},
    {"DVAR", &eval_dvar_lazy},
    {"DVARP", &eval_dvarp_lazy},
    {"EDATE", &eval_datetime_lazy},
    {"EOMONTH", &eval_datetime_lazy},
    {"EXPAND", &eval_expand_lazy},
    {"F.TEST", &eval_f_test_lazy},
    {"FILTER", &eval_filter_lazy},
    // FORECAST is the legacy spelling kept by Excel for back-compat;
    // its impl and arity are identical to FORECAST.LINEAR.
    {"FORECAST", &eval_forecast_linear_lazy},
    // The FORECAST.ETS family rides the lazy seam because both the
    // values and timeline arguments may be Range refs that must reach
    // the impl with their (rows, cols) shape preserved for pairing.
    {"FORECAST.ETS", &eval_forecast_ets_lazy},
    {"FORECAST.ETS.CONFINT", &eval_forecast_ets_confint_lazy},
    {"FORECAST.ETS.SEASONALITY", &eval_forecast_ets_seasonality_lazy},
    {"FORECAST.ETS.STAT", &eval_forecast_ets_stat_lazy},
    {"FORECAST.LINEAR", &eval_forecast_linear_lazy},
    // FORMULATEXT returns the source text of the referenced cell's formula,
    // so it must inspect the un-evaluated Ref AST and the bound Sheet's
    // `formula_text` directly — the eager path would flatten the argument
    // to a Value before we could see the reference.
    {"FORMULATEXT", &eval_formulatext_lazy},
    {"FREQUENCY", &eval_frequency_lazy},
    // FTEST is the pre-2010 legacy spelling of F.TEST; same impl.
    {"FTEST", &eval_f_test_lazy},
    // GETPIVOTDATA recovers the (sheet, row, col) anchor of its second
    // argument from the un-evaluated Reference AST and reads the
    // freshest `PivotResult` off the bound Workbook on EvalContext;
    // the eager dispatcher would flatten the anchor to a Value before
    // the impl could see the reference.
    {"GETPIVOTDATA", &eval_getpivotdata_lazy},
    {"GROUPBY", &eval_groupby_lazy},
    {"GROWTH", &eval_growth_lazy},
    {"HLOOKUP", &eval_hlookup_lazy},
    {"HSTACK", &eval_hstack_lazy},
    {"IF", &eval_if_lazy},
    {"IFERROR", &eval_iferror_lazy},
    {"IFNA", &eval_ifna_lazy},
    {"IFS", &eval_ifs_lazy},
    {"INDEX", &eval_index_lazy},
    {"INDIRECT", &eval_indirect_lazy},
    {"INTERCEPT", &eval_intercept_lazy},
    // ISFORMULA / ISREF inspect the un-evaluated AST of their argument;
    // they cannot ride the eager path because it flattens references to
    // `Value` before the impl runs.
    {"ISFORMULA", &eval_isformula_lazy},
    {"ISOMITTED", &eval_isomitted_lazy},
    {"ISOWEEKNUM", &eval_datetime_lazy},
    {"ISREF", &eval_isref_lazy},
    {"IRR", &eval_irr_lazy},
    {"LENB", &eval_lenb_lazy},
    {"LINEST", &eval_linest_lazy},
    {"LOGEST", &eval_logest_lazy},
    {"LOOKUP", &eval_lookup_lazy},
    {"MAKEARRAY", &eval_makearray_lazy},
    {"MAP", &eval_map_lazy},
    {"MATCH", &eval_match_lazy},
    {"MAXIFS", &eval_maxifs_lazy},
    {"MDETERM", &eval_mdeterm_lazy},
    {"MINIFS", &eval_minifs_lazy},
    {"MINVERSE", &eval_minverse_lazy},
    {"MIRR", &eval_mirr_lazy},
    {"MMULT", &eval_mmult_lazy},
    {"MONTH", &eval_datetime_lazy},
    {"NETWORKDAYS", &eval_networkdays_lazy},
    {"NETWORKDAYS.INTL", &eval_networkdays_intl_lazy},
    {"NOW", &eval_datetime_lazy},
    {"OFFSET", &eval_offset_lazy},
    {"OR", &eval_or_lazy},
    // PEARSON is mathematically identical to CORREL (Pearson product-moment
    // correlation coefficient); Excel keeps both names for back-compat.
    {"PEARSON", &eval_correl_lazy},
    {"PERCENTOF", &eval_percentof_lazy},
    {"PERCENTRANK", &eval_percentrank_inc_lazy},
    {"PERCENTRANK.EXC", &eval_percentrank_exc_lazy},
    {"PERCENTRANK.INC", &eval_percentrank_inc_lazy},
    // PHONETIC reads the cell's <rPh> annotation off the un-evaluated
    // Ref AST; the eager dispatcher would flatten the argument to a
    // Value before the impl could consult `Cell::phonetic_text`.
    {"PHONETIC", &eval_phonetic_lazy},
    {"PIVOTBY", &eval_pivotby_lazy},
    {"PROB", &eval_prob_lazy},
    {"RANK", &eval_rank_eq_lazy},
    {"RANK.AVG", &eval_rank_avg_lazy},
    {"RANK.EQ", &eval_rank_eq_lazy},
    {"REDUCE", &eval_reduce_lazy},
    {"REGEXEXTRACT", &eval_regexextract_lazy},
    {"REGEXREPLACE", &eval_regexreplace_lazy},
    {"REGEXTEST", &eval_regextest_lazy},
    {"ROW", &eval_row_lazy},
    {"ROWS", &eval_rows_lazy},
    {"RSQ", &eval_rsq_lazy},
    {"SCAN", &eval_scan_lazy},
    {"SERIESSUM", &eval_series_sum_lazy},
    // SHEET / SHEETS consult the bound Workbook + current Sheet on the
    // EvalContext; AST introspection of an optional reference argument
    // tells them which sheet to answer for.
    {"SHEET", &eval_sheet_lazy},
    {"SHEETS", &eval_sheets_lazy},
    // SINGLE is the explicit-name form of the `@` implicit-intersection
    // operator; xlsx serialises `@range` as `_xlfn.SINGLE(range)` (the
    // `_xlfn.` prefix is stripped by `strip_future_prefix`). Routes to a
    // lazy impl so the un-evaluated RangeOp AST can be projected onto the
    // formula cell's row / column via `EvalContext::formula_row` /
    // `formula_col` — the same logic the `@` operator uses.
    {"SINGLE", &eval_single_lazy},
    {"SLOPE", &eval_slope_lazy},
    {"SORT", &eval_sort_lazy},
    {"SORTBY", &eval_sortby_lazy},
    {"STEYX", &eval_steyx_lazy},
    {"SUMIF", &eval_sumif_lazy},
    {"SUMIFS", &eval_sumifs_lazy},
    {"SUMPRODUCT", &eval_sumproduct_lazy},
    {"SUMX2MY2", &eval_sumx2my2_lazy},
    {"SUMX2PY2", &eval_sumx2py2_lazy},
    {"SUMXMY2", &eval_sumxmy2_lazy},
    {"SWITCH", &eval_switch_lazy},
    {"T.TEST", &eval_t_test_lazy},
    {"TAKE", &eval_take_lazy},
    {"TEXT", &eval_text_lazy},
    {"TEXTSPLIT", &eval_textsplit_lazy},
    {"TOCOL", &eval_tocol_lazy},
    {"TODAY", &eval_datetime_lazy},
    {"TOROW", &eval_torow_lazy},
    {"TRANSPOSE", &eval_transpose_lazy},
    {"TREND", &eval_trend_lazy},
    {"TRIMRANGE", &eval_trimrange_lazy},
    // TTEST is the pre-2010 legacy spelling of T.TEST; same impl.
    {"TTEST", &eval_t_test_lazy},
    {"UNIQUE", &eval_unique_lazy},
    {"VLOOKUP", &eval_vlookup_lazy},
    {"VSTACK", &eval_vstack_lazy},
    {"WEEKDAY", &eval_datetime_lazy},
    {"WEEKNUM", &eval_weeknum_lazy},
    {"WORKDAY", &eval_workday_lazy},
    {"WORKDAY.INTL", &eval_workday_intl_lazy},
    {"WRAPCOLS", &eval_wrapcols_lazy},
    {"WRAPROWS", &eval_wraprows_lazy},
    {"XIRR", &eval_xirr_lazy},
    {"XLOOKUP", &eval_xlookup_lazy},
    {"XMATCH", &eval_xmatch_lazy},
    {"XNPV", &eval_xnpv_lazy},
    {"YEAR", &eval_datetime_lazy},
    {"YEARFRAC", &eval_datetime_lazy},
    {"Z.TEST", &eval_z_test_lazy},
    // ZTEST is the pre-2010 legacy spelling of Z.TEST; same impl.
    {"ZTEST", &eval_z_test_lazy},
};

constexpr std::size_t kLazyDispatchCount = sizeof(kLazyDispatch) / sizeof(kLazyDispatch[0]);

}  // namespace

LazyImpl find_lazy_impl(std::string_view name) noexcept {
  for (const auto& e : kLazyDispatch) {
    if (strings::case_insensitive_eq(name, std::string_view(e.name))) {
      return e.impl;
    }
  }
  return nullptr;
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
