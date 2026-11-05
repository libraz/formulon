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
#include "eval/function_registry.h"
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
#include "eval/textjoin_lazy.h"
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
  // Written argument counts Excel accepts at entry, as `FunctionDef`'s
  // min_arity / max_arity. The impl still checks its own arguments.
  std::uint32_t min_arity;
  std::uint32_t max_arity;
};

// The single source of truth for lazy-form routing. Sorted alphabetically
// by canonical UPPERCASE name so a quick visual diff catches accidental
// duplicates. Comments preserved verbatim from the prior in-place table.
constexpr LazyEntry kLazyDispatch[] = {
    // The trim-reference operators `.:`, `:.`, `.:.` (see
    // `parser::trim_ref_function_name`); `_` sorts first under the lookup's
    // lower-case folding.
    {"_TRO_ALL", &eval_trim_ref_lazy, LazyResultShape::kArray, 1U, 1U},
    {"_TRO_LEADING", &eval_trim_ref_lazy, LazyResultShape::kArray, 1U, 1U},
    {"_TRO_TRAILING", &eval_trim_ref_lazy, LazyResultShape::kArray, 1U, 1U},
    // Financial date1904-sensitive family: shares one lazy impl
    // (`eval_financial_date_lazy`) the same way the calendar family shares
    // `eval_datetime_lazy` -- see `eval/financial_lazy.h`.
    {"ACCRINT", &eval_financial_date_lazy, LazyResultShape::kScalar, 6U, 8U},
    {"ACCRINTM", &eval_financial_date_lazy, LazyResultShape::kScalar, 4U, 5U},
    {"AGGREGATE", &eval_aggregate_lazy, LazyResultShape::kReduce, 3U, kVariadic},
    {"AMORDEGRC", &eval_financial_date_lazy, LazyResultShape::kScalar, 6U, 7U},
    {"AMORLINC", &eval_financial_date_lazy, LazyResultShape::kScalar, 6U, 7U},
    // ANCHORARRAY is the OOXML internal encoding of the postfix `#`
    // spill operator. The xlsx-only `_xlfn.` prefix is stripped by
    // `strip_future_prefix`, so callers register the canonical bare name.
    // See `eval_anchorarray_lazy`.
    {"ANCHORARRAY", &eval_anchorarray_lazy, LazyResultShape::kArray, 1U, 1U},
    {"AND", &eval_and_lazy, LazyResultShape::kReduce, 1U, kVariadic},
    {"AREAS", &eval_areas_lazy, LazyResultShape::kReduce, 1U, 1U},
    {"ARRAYTOTEXT", &eval_arraytotext_lazy, LazyResultShape::kScalar, 1U, 2U},
    {"AVERAGEIF", &eval_averageif_lazy, LazyResultShape::kReduce, 2U, 3U},
    {"AVERAGEIFS", &eval_averageifs_lazy, LazyResultShape::kReduce, 3U, kVariadic},
    {"BYCOL", &eval_bycol_lazy, LazyResultShape::kArray, 1U, 2U},
    {"BYROW", &eval_byrow_lazy, LazyResultShape::kArray, 1U, 2U},
    {"CELL", &eval_cell_lazy, LazyResultShape::kArray, 1U, 2U},
    // CHITEST is the pre-2010 legacy spelling of CHISQ.TEST; same impl.
    {"CHISQ.TEST", &eval_chisq_test_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"CHITEST", &eval_chisq_test_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"CHOOSE", &eval_choose_lazy, LazyResultShape::kBroadcast, 2U, kVariadic},
    {"CHOOSECOLS", &eval_choosecols_lazy, LazyResultShape::kArray, 2U, kVariadic},
    {"CHOOSEROWS", &eval_chooserows_lazy, LazyResultShape::kArray, 2U, kVariadic},
    {"CODE", &eval_code_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"COLUMN", &eval_column_lazy, LazyResultShape::kBroadcast, 0U, 1U},
    {"COLUMNS", &eval_columns_lazy, LazyResultShape::kReduce, 1U, 1U},
    {"CORREL", &eval_correl_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"COUNT", &eval_count_lazy, LazyResultShape::kReduce, 1U, kVariadic},
    {"COUNTBLANK", &eval_countblank_lazy, LazyResultShape::kReduce, 1U, 1U},
    {"COUNTIF", &eval_countif_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"COUNTIFS", &eval_countifs_lazy, LazyResultShape::kReduce, 2U, kVariadic},
    {"COUPDAYBS", &eval_financial_date_lazy, LazyResultShape::kScalar, 3U, 4U},
    {"COUPDAYS", &eval_financial_date_lazy, LazyResultShape::kScalar, 3U, 4U},
    {"COUPDAYSNC", &eval_financial_date_lazy, LazyResultShape::kScalar, 3U, 4U},
    {"COUPNCD", &eval_financial_date_lazy, LazyResultShape::kScalar, 3U, 4U},
    {"COUPNUM", &eval_financial_date_lazy, LazyResultShape::kScalar, 3U, 4U},
    {"COUPPCD", &eval_financial_date_lazy, LazyResultShape::kScalar, 3U, 4U},
    // COVAR is the pre-2010 legacy spelling of COVARIANCE.P; both compute
    // the population covariance with identical semantics.
    {"COVAR", &eval_covariance_p_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"COVARIANCE.P", &eval_covariance_p_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"COVARIANCE.S", &eval_covariance_s_lazy, LazyResultShape::kReduce, 2U, 2U},
    // Calendar family: date1904-sensitive functions share one lazy impl
    // (`eval_datetime_lazy`) so the workbook epoch reaches the calendar math.
    // WEEKNUM is served by `eval_weeknum_lazy` and is registered separately
    // below.
    {"DATE", &eval_datetime_lazy, LazyResultShape::kBroadcast, 3U, 3U},
    {"DATEDIF", &eval_datetime_lazy, LazyResultShape::kBroadcast, 3U, 3U},
    {"DATEVALUE", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"DAVERAGE", &eval_daverage_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DAY", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"DAYS", &eval_datetime_lazy, LazyResultShape::kBroadcast, 2U, 2U},
    {"DAYS360", &eval_datetime_lazy, LazyResultShape::kBroadcast, 2U, 3U},
    {"DCOUNT", &eval_dcount_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DCOUNTA", &eval_dcounta_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DGET", &eval_dget_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DISC", &eval_financial_date_lazy, LazyResultShape::kScalar, 4U, 5U},
    {"DMAX", &eval_dmax_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DMIN", &eval_dmin_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DPRODUCT", &eval_dproduct_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DROP", &eval_drop_lazy, LazyResultShape::kArray, 2U, 3U},
    {"DSTDEV", &eval_dstdev_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DSTDEVP", &eval_dstdevp_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DSUM", &eval_dsum_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DURATION", &eval_financial_date_lazy, LazyResultShape::kScalar, 5U, 6U},
    {"DVAR", &eval_dvar_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"DVARP", &eval_dvarp_lazy, LazyResultShape::kReduce, 3U, 3U},
    {"EDATE", &eval_datetime_lazy, LazyResultShape::kBroadcast, 2U, 2U},
    {"EOMONTH", &eval_datetime_lazy, LazyResultShape::kBroadcast, 2U, 2U},
    {"EXPAND", &eval_expand_lazy, LazyResultShape::kArray, 2U, 4U},
    {"F.TEST", &eval_f_test_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"FILTER", &eval_filter_lazy, LazyResultShape::kArray, 2U, 3U},
    // FORECAST is the legacy spelling kept by Excel for back-compat;
    // its impl and arity are identical to FORECAST.LINEAR.
    {"FORECAST", &eval_forecast_linear_lazy, LazyResultShape::kBroadcast, 3U, 3U},
    // The FORECAST.ETS family rides the lazy seam because both the
    // values and timeline arguments may be Range refs that must reach
    // the impl with their (rows, cols) shape preserved for pairing.
    {"FORECAST.ETS", &eval_forecast_ets_lazy, LazyResultShape::kBroadcast, 3U, 6U},
    {"FORECAST.ETS.CONFINT", &eval_forecast_ets_confint_lazy, LazyResultShape::kBroadcast, 3U, 7U},
    {"FORECAST.ETS.SEASONALITY", &eval_forecast_ets_seasonality_lazy, LazyResultShape::kScalar, 2U, 4U},
    {"FORECAST.ETS.STAT", &eval_forecast_ets_stat_lazy, LazyResultShape::kScalar, 3U, 6U},
    {"FORECAST.LINEAR", &eval_forecast_linear_lazy, LazyResultShape::kBroadcast, 3U, 3U},
    // FORMULATEXT returns the source text of the referenced cell's formula,
    // so it must inspect the un-evaluated Ref AST and the bound Sheet's
    // `formula_text` directly — the eager path would flatten the argument
    // to a Value before we could see the reference.
    {"FORMULATEXT", &eval_formulatext_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"FREQUENCY", &eval_frequency_lazy, LazyResultShape::kArray, 2U, 2U},
    // FTEST is the pre-2010 legacy spelling of F.TEST; same impl.
    {"FTEST", &eval_f_test_lazy, LazyResultShape::kReduce, 2U, 2U},
    // GETPIVOTDATA recovers the (sheet, row, col) anchor of its second
    // argument from the un-evaluated Reference AST and reads the
    // freshest `PivotResult` off the bound Workbook on EvalContext;
    // the eager dispatcher would flatten the anchor to a Value before
    // the impl could see the reference.
    {"GETPIVOTDATA", &eval_getpivotdata_lazy, LazyResultShape::kScalar, 2U, kVariadic},
    {"GROUPBY", &eval_groupby_lazy, LazyResultShape::kArray, 3U, 8U},
    {"GROWTH", &eval_growth_lazy, LazyResultShape::kArray, 1U, 4U},
    {"HLOOKUP", &eval_hlookup_lazy, LazyResultShape::kArray, 3U, 4U},
    {"HOUR", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"HSTACK", &eval_hstack_lazy, LazyResultShape::kArray, 1U, kVariadic},
    {"IF", &eval_if_lazy, LazyResultShape::kBroadcast, 2U, 3U},
    {"IFERROR", &eval_iferror_lazy, LazyResultShape::kBroadcast, 2U, 2U},
    {"IFNA", &eval_ifna_lazy, LazyResultShape::kBroadcast, 2U, 2U},
    {"IFS", &eval_ifs_lazy, LazyResultShape::kBroadcast, 2U, kVariadic},
    {"INDEX", &eval_index_lazy, LazyResultShape::kArray, 2U, 4U},
    {"INDIRECT", &eval_indirect_lazy, LazyResultShape::kArray, 1U, 2U},
    {"INTERCEPT", &eval_intercept_lazy, LazyResultShape::kScalar, 2U, 2U},
    {"INTRATE", &eval_financial_date_lazy, LazyResultShape::kScalar, 4U, 5U},
    // ISFORMULA / ISREF inspect the un-evaluated AST of their argument;
    // they cannot ride the eager path because it flattens references to
    // `Value` before the impl runs.
    {"IRR", &eval_irr_lazy, LazyResultShape::kScalar, 1U, 2U},
    {"ISFORMULA", &eval_isformula_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"ISOMITTED", &eval_isomitted_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"ISOWEEKNUM", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"ISREF", &eval_isref_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"LENB", &eval_lenb_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"LINEST", &eval_linest_lazy, LazyResultShape::kArray, 1U, 4U},
    {"LOGEST", &eval_logest_lazy, LazyResultShape::kArray, 1U, 4U},
    {"LOOKUP", &eval_lookup_lazy, LazyResultShape::kScalar, 2U, 3U},
    {"MAKEARRAY", &eval_makearray_lazy, LazyResultShape::kArray, 3U, 3U},
    {"MAP", &eval_map_lazy, LazyResultShape::kArray, 2U, kVariadic},
    {"MATCH", &eval_match_lazy, LazyResultShape::kArray, 2U, 3U},
    {"MAXIFS", &eval_maxifs_lazy, LazyResultShape::kReduce, 3U, kVariadic},
    {"MDETERM", &eval_mdeterm_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"MDURATION", &eval_financial_date_lazy, LazyResultShape::kScalar, 5U, 6U},
    {"MINIFS", &eval_minifs_lazy, LazyResultShape::kReduce, 3U, kVariadic},
    {"MINUTE", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"MINVERSE", &eval_minverse_lazy, LazyResultShape::kArray, 1U, 1U},
    {"MIRR", &eval_mirr_lazy, LazyResultShape::kScalar, 3U, 3U},
    {"MMULT", &eval_mmult_lazy, LazyResultShape::kArray, 2U, 2U},
    {"MONTH", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"N", &eval_n_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"NETWORKDAYS", &eval_networkdays_lazy, LazyResultShape::kScalar, 2U, 3U},
    {"NETWORKDAYS.INTL", &eval_networkdays_intl_lazy, LazyResultShape::kScalar, 2U, 4U},
    {"NOW", &eval_datetime_lazy, LazyResultShape::kScalar, 0U, 0U},
    {"NUMBERVALUE", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 3U},
    {"ODDFPRICE", &eval_financial_date_lazy, LazyResultShape::kScalar, 8U, 9U},
    {"ODDFYIELD", &eval_financial_date_lazy, LazyResultShape::kScalar, 8U, 9U},
    {"ODDLPRICE", &eval_financial_date_lazy, LazyResultShape::kScalar, 7U, 8U},
    {"ODDLYIELD", &eval_financial_date_lazy, LazyResultShape::kScalar, 7U, 8U},
    {"OFFSET", &eval_offset_lazy, LazyResultShape::kArray, 3U, 5U},
    {"OR", &eval_or_lazy, LazyResultShape::kReduce, 1U, kVariadic},
    // PEARSON is mathematically identical to CORREL (Pearson product-moment
    // correlation coefficient); Excel keeps both names for back-compat.
    {"PEARSON", &eval_correl_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"PERCENTOF", &eval_percentof_lazy, LazyResultShape::kScalar, 2U, 2U},
    {"PERCENTRANK", &eval_percentrank_inc_lazy, LazyResultShape::kScalar, 2U, 3U},
    {"PERCENTRANK.EXC", &eval_percentrank_exc_lazy, LazyResultShape::kScalar, 2U, 3U},
    {"PERCENTRANK.INC", &eval_percentrank_inc_lazy, LazyResultShape::kScalar, 2U, 3U},
    // PHONETIC reads the cell's <rPh> annotation off the un-evaluated
    // Ref AST; the eager dispatcher would flatten the argument to a
    // Value before the impl could consult `Cell::phonetic_text`.
    {"PHONETIC", &eval_phonetic_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"PIVOTBY", &eval_pivotby_lazy, LazyResultShape::kArray, 4U, 11U},
    {"PRICE", &eval_financial_date_lazy, LazyResultShape::kScalar, 6U, 7U},
    {"PRICEDISC", &eval_financial_date_lazy, LazyResultShape::kScalar, 4U, 5U},
    {"PRICEMAT", &eval_financial_date_lazy, LazyResultShape::kScalar, 5U, 6U},
    {"PROB", &eval_prob_lazy, LazyResultShape::kScalar, 3U, 4U},
    {"RANK", &eval_rank_eq_lazy, LazyResultShape::kScalar, 2U, 3U},
    {"RANK.AVG", &eval_rank_avg_lazy, LazyResultShape::kScalar, 2U, 3U},
    {"RANK.EQ", &eval_rank_eq_lazy, LazyResultShape::kScalar, 2U, 3U},
    {"RECEIVED", &eval_financial_date_lazy, LazyResultShape::kScalar, 4U, 5U},
    {"REDUCE", &eval_reduce_lazy, LazyResultShape::kArray, 3U, 3U},
    {"REGEXEXTRACT", &eval_regexextract_lazy, LazyResultShape::kArray, 2U, 4U},
    {"REGEXREPLACE", &eval_regexreplace_lazy, LazyResultShape::kBroadcast, 3U, 5U},
    {"REGEXTEST", &eval_regextest_lazy, LazyResultShape::kBroadcast, 2U, 3U},
    {"ROW", &eval_row_lazy, LazyResultShape::kBroadcast, 0U, 1U},
    {"ROWS", &eval_rows_lazy, LazyResultShape::kReduce, 1U, 1U},
    {"RSQ", &eval_rsq_lazy, LazyResultShape::kScalar, 2U, 2U},
    {"SCAN", &eval_scan_lazy, LazyResultShape::kArray, 3U, 3U},
    {"SECOND", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"SERIESSUM", &eval_series_sum_lazy, LazyResultShape::kScalar, 4U, 4U},
    // SHEET / SHEETS consult the bound Workbook + current Sheet on the
    // EvalContext; AST introspection of an optional reference argument
    // tells them which sheet to answer for.
    {"SHEET", &eval_sheet_lazy, LazyResultShape::kScalar, 0U, 1U},
    {"SHEETS", &eval_sheets_lazy, LazyResultShape::kReduce, 0U, 1U},
    // SINGLE is the explicit-name form of the `@` implicit-intersection
    // operator; xlsx serialises `@range` as `_xlfn.SINGLE(range)` (the
    // `_xlfn.` prefix is stripped by `strip_future_prefix`). Routes to a
    // lazy impl so the un-evaluated RangeOp AST can be projected onto the
    // formula cell's row / column via `EvalContext::formula_row` /
    // `formula_col` — the same logic the `@` operator uses.
    {"SINGLE", &eval_single_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"SLOPE", &eval_slope_lazy, LazyResultShape::kScalar, 2U, 2U},
    {"SORT", &eval_sort_lazy, LazyResultShape::kArray, 1U, 4U},
    {"SORTBY", &eval_sortby_lazy, LazyResultShape::kArray, 2U, kVariadic},
    {"STEYX", &eval_steyx_lazy, LazyResultShape::kScalar, 2U, 2U},
    // SUBTOTAL is lazy so codes 101..111 can read row visibility off the
    // sheet. See `eval_subtotal_lazy`.
    {"SUBTOTAL", &eval_subtotal_lazy, LazyResultShape::kReduce, 2U, kVariadic},
    {"SUMIF", &eval_sumif_lazy, LazyResultShape::kReduce, 2U, 3U},
    {"SUMIFS", &eval_sumifs_lazy, LazyResultShape::kReduce, 3U, kVariadic},
    {"SUMPRODUCT", &eval_sumproduct_lazy, LazyResultShape::kReduce, 1U, kVariadic},
    {"SUMX2MY2", &eval_sumx2my2_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"SUMX2PY2", &eval_sumx2py2_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"SUMXMY2", &eval_sumxmy2_lazy, LazyResultShape::kReduce, 2U, 2U},
    {"SWITCH", &eval_switch_lazy, LazyResultShape::kBroadcast, 3U, kVariadic},
    {"T", &eval_t_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"T.TEST", &eval_t_test_lazy, LazyResultShape::kReduce, 4U, 4U},
    {"TAKE", &eval_take_lazy, LazyResultShape::kArray, 2U, 3U},
    {"TBILLEQ", &eval_financial_date_lazy, LazyResultShape::kScalar, 3U, 3U},
    {"TBILLPRICE", &eval_financial_date_lazy, LazyResultShape::kScalar, 3U, 3U},
    {"TBILLYIELD", &eval_financial_date_lazy, LazyResultShape::kScalar, 3U, 3U},
    {"TEXT", &eval_text_lazy, LazyResultShape::kBroadcast, 2U, 2U},
    {"TEXTJOIN", &eval_textjoin_lazy, LazyResultShape::kScalar, 3U, kVariadic},
    {"TEXTSPLIT", &eval_textsplit_lazy, LazyResultShape::kArray, 2U, 6U},
    {"TOCOL", &eval_tocol_lazy, LazyResultShape::kArray, 1U, 3U},
    {"TODAY", &eval_datetime_lazy, LazyResultShape::kScalar, 0U, 0U},
    {"TOROW", &eval_torow_lazy, LazyResultShape::kArray, 1U, 3U},
    {"TRANSPOSE", &eval_transpose_lazy, LazyResultShape::kArray, 1U, 1U},
    {"TREND", &eval_trend_lazy, LazyResultShape::kArray, 1U, 4U},
    {"TRIMRANGE", &eval_trimrange_lazy, LazyResultShape::kArray, 1U, 3U},
    // TTEST is the pre-2010 legacy spelling of T.TEST; same impl.
    {"TTEST", &eval_t_test_lazy, LazyResultShape::kReduce, 4U, 4U},
    {"TYPE", &eval_type_lazy, LazyResultShape::kScalar, 1U, 1U},
    {"UNIQUE", &eval_unique_lazy, LazyResultShape::kArray, 1U, 3U},
    {"VALUE", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"VLOOKUP", &eval_vlookup_lazy, LazyResultShape::kArray, 3U, 4U},
    {"VSTACK", &eval_vstack_lazy, LazyResultShape::kArray, 1U, kVariadic},
    {"WEEKDAY", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 2U},
    {"WEEKNUM", &eval_weeknum_lazy, LazyResultShape::kScalar, 1U, 2U},
    {"WORKDAY", &eval_workday_lazy, LazyResultShape::kScalar, 2U, 3U},
    {"WORKDAY.INTL", &eval_workday_intl_lazy, LazyResultShape::kScalar, 2U, 4U},
    {"WRAPCOLS", &eval_wrapcols_lazy, LazyResultShape::kArray, 2U, 3U},
    {"WRAPROWS", &eval_wraprows_lazy, LazyResultShape::kArray, 2U, 3U},
    {"XIRR", &eval_xirr_lazy, LazyResultShape::kScalar, 2U, 3U},
    {"XLOOKUP", &eval_xlookup_lazy, LazyResultShape::kArray, 3U, 6U},
    {"XMATCH", &eval_xmatch_lazy, LazyResultShape::kArray, 2U, 4U},
    {"XNPV", &eval_xnpv_lazy, LazyResultShape::kScalar, 3U, 3U},
    {"YEAR", &eval_datetime_lazy, LazyResultShape::kBroadcast, 1U, 1U},
    {"YEARFRAC", &eval_datetime_lazy, LazyResultShape::kBroadcast, 2U, 3U},
    {"YIELD", &eval_financial_date_lazy, LazyResultShape::kScalar, 6U, 7U},
    {"YIELDDISC", &eval_financial_date_lazy, LazyResultShape::kScalar, 4U, 5U},
    {"YIELDMAT", &eval_financial_date_lazy, LazyResultShape::kScalar, 5U, 6U},
    {"Z.TEST", &eval_z_test_lazy, LazyResultShape::kReduce, 2U, 3U},
    // ZTEST is the pre-2010 legacy spelling of Z.TEST; same impl.
    {"ZTEST", &eval_z_test_lazy, LazyResultShape::kReduce, 2U, 3U},
};

constexpr std::size_t kLazyDispatchCount = sizeof(kLazyDispatch) / sizeof(kLazyDispatch[0]);

// The order `find_lazy_entry` searches in: bytes compared ASCII-lower-cased.
constexpr int compare_canonical_names(const char* lhs, const char* rhs) {
  for (std::size_t i = 0;; ++i) {
    const char l = (lhs[i] >= 'A' && lhs[i] <= 'Z') ? static_cast<char>(lhs[i] - 'A' + 'a') : lhs[i];
    const char r = (rhs[i] >= 'A' && rhs[i] <= 'Z') ? static_cast<char>(rhs[i] - 'A' + 'a') : rhs[i];
    if (l < r) {
      return -1;
    }
    if (l > r) {
      return 1;
    }
    if (l == '\0') {
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

// Binary search over the canonical-name-ordered table; nullptr when absent.
const LazyEntry* find_lazy_entry(std::string_view name) noexcept {
  std::size_t first = 0;
  std::size_t last = kLazyDispatchCount;
  while (first < last) {
    const std::size_t middle = first + (last - first) / 2;
    const int cmp = strings::case_insensitive_compare(name, kLazyDispatch[middle].name);
    if (cmp == 0) {
      return &kLazyDispatch[middle];
    }
    if (cmp < 0) {
      last = middle;
    } else {
      first = middle + 1;
    }
  }
  return nullptr;
}

}  // namespace

LazyImpl find_lazy_impl(std::string_view name) noexcept {
  const LazyEntry* entry = find_lazy_entry(name);
  return entry != nullptr ? entry->impl : nullptr;
}

LazyResultShape find_lazy_result_shape(std::string_view name) noexcept {
  const LazyEntry* entry = find_lazy_entry(name);
  return entry != nullptr ? entry->shape : LazyResultShape::kNotLazy;
}

bool find_lazy_arity(std::string_view name, std::uint32_t* min_arity, std::uint32_t* max_arity) noexcept {
  const LazyEntry* entry = find_lazy_entry(name);
  if (entry == nullptr) {
    return false;
  }
  *min_arity = entry->min_arity;
  *max_arity = entry->max_arity;
  return true;
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
