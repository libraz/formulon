//
// Implementation of the CF evaluator's shared helpers. See cf_helpers.h
// for the contract.

#include "cf/cf_helpers.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_evaluator.h"
#include "cf/cf_types.h"
#include "eval/builtin_names.h"
#include "eval/coerce.h"
#include "eval/collation.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/eval_state.h"
#include "eval/scalar_ops.h"
#include "eval/tree_walker.h"
#include "eval/value_identity.h"
#include "parser/ast.h"
#include "parser/ast_shift.h"
#include "sheet.h"
#include "utils/date_time.h"
#include "utils/double_parse.h"
#include "utils/rect_iterator.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon::cf {

void PreparedCfFormulaCache::initialize(const Arena& caller_arena) {
  if (persistent_arena.has_value()) {
    return;
  }
  initial_chunk_bytes = caller_arena.initial_chunk_bytes();
  const std::size_t allocated = caller_arena.bytes_allocated();
  available_bytes = caller_arena.max_bytes() > allocated ? caller_arena.max_bytes() - allocated : 0U;
  const std::size_t initial = std::min(initial_chunk_bytes, available_bytes);
  persistent_arena.emplace(initial == 0U ? initial_chunk_bytes : initial, available_bytes);
}

void PreparedCfFormulaCache::create_scratch() {
  scratch_arena.reset();
  const std::size_t persistent_bytes = persistent_arena.has_value() ? persistent_arena->bytes_allocated() : 0U;
  const std::size_t remaining = available_bytes > persistent_bytes ? available_bytes - persistent_bytes : 0U;
  const std::size_t initial = std::min(initial_chunk_bytes, remaining);
  scratch_arena.emplace(initial == 0U ? initial_chunk_bytes : initial, remaining);
}

void PreparedCfFormulaCache::capture_resource_failure(bool state_out_of_memory) noexcept {
  resource_failure = resource_failure || state_out_of_memory ||
                     (persistent_arena.has_value() && persistent_arena->exhausted()) ||
                     (scratch_arena.has_value() && scratch_arena->exhausted());
}

}  // namespace formulon::cf

namespace formulon::cf::helpers {

// ---------------------------------------------------------------------------
// Literal-operand helpers.
// ---------------------------------------------------------------------------

std::optional<LiteralOperand> parse_literal(const std::string& source) {
  if (source.empty()) {
    return std::nullopt;
  }

  // Quoted text literal with `""` -> `"` unescape.
  if (source.size() >= 2 && source.front() == '"' && source.back() == '"') {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Text;
    std::string& out = operand.text_value;
    out.reserve(source.size() - 2);
    for (std::size_t i = 1; i + 1 < source.size(); ++i) {
      if (source[i] == '"' && i + 2 < source.size() && source[i + 1] == '"') {
        out.push_back('"');
        ++i;
      } else {
        out.push_back(source[i]);
      }
    }
    return operand;
  }

  // Boolean keywords (case-insensitive, matching Excel's tolerance).
  if (strings::case_insensitive_eq(source, "TRUE")) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Bool;
    operand.bool_value = true;
    return operand;
  }
  if (strings::case_insensitive_eq(source, "FALSE")) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Bool;
    operand.bool_value = false;
    return operand;
  }

  // Numeric literal — consumed in full or rejected. Route through the
  // shared decimal parser (`utils/double_parse.h`) so the cellIs literal parser
  // and the engine's text-to-number coercion accept exactly the same syntax.
  double parsed = 0.0;
  if (parse_double_exact(source, &parsed)) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Number;
    operand.number_value = parsed;
    return operand;
  }
  return std::nullopt;
}

std::optional<int> compare_cell_to_literal(const Value& cell, const LiteralOperand& operand) {
  if (cell.is_error()) {
    return std::nullopt;
  }

  // The cellIs rule is the worksheet comparison `cell <op> operand`, so it
  // routes through the evaluator's `compare_values`: blank is 0 against a
  // number and "" against text, and numbers < text < booleans. `LiteralOperand`
  // owns its text payload, so the non-owning `Value::text` view is valid for
  // the duration of this call.
  Value rhs = Value::blank();
  switch (operand.kind) {
    case LiteralOperand::Kind::Text:
      rhs = Value::text(operand.text_value);
      break;
    case LiteralOperand::Kind::Number:
      rhs = Value::number(operand.number_value);
      break;
    case LiteralOperand::Kind::Bool:
      rhs = Value::boolean(operand.bool_value);
      break;
  }
  bool unordered = false;
  return eval::compare_values(cell, rhs, &unordered);
}

std::optional<LiteralOperand> value_to_operand(const Value& evaluated) {
  if (evaluated.is_number()) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Number;
    operand.number_value = evaluated.as_number();
    return operand;
  }
  if (evaluated.is_boolean()) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Bool;
    operand.bool_value = evaluated.as_boolean();
    return operand;
  }
  if (evaluated.is_text()) {
    LiteralOperand operand;
    operand.kind = LiteralOperand::Kind::Text;
    operand.text_value = std::string(evaluated.as_text());
    return operand;
  }
  return std::nullopt;
}

Value parse_shift_evaluate(const std::string& source, const CFEvalContext& ctx) {
  if (ctx.arena == nullptr || ctx.registry == nullptr || ctx.eval_ctx == nullptr) {
    return Value::error(ErrorCode::Name);
  }

  // The model holds CF formulas in canonical formula-bar spelling (the
  // readers strip Excel's storage prefixes on load), without a leading `=`.
  // Prepared walks keep roots in a private persistent arena and build each
  // shifted tree / evaluation result in a reusable scratch arena. The
  // caller-owned arena remains untouched by this path.
  const parser::AstNode* root = nullptr;
  Arena* formula_arena = ctx.arena;
  if (ctx.formula_cache != nullptr) {
    PreparedCfFormulaCache& cache = *ctx.formula_cache;
    cache.initialize(*ctx.arena);
    if (cache.resource_failure) {
      return Value::error(ErrorCode::Name);
    }

    auto it = cache.indices.find(source);
    if (it == cache.indices.end()) {
      // Give persistent roots the full private budget while a new root is
      // parsed, then hand only the remaining capacity to scratch.
      cache.scratch_arena.reset();
      const std::size_t root_index = cache.roots.size();
      cache.indices.emplace(source, root_index);
      cache.roots.push_back(nullptr);
      root = eval::parse_formula_entry(source, *cache.persistent_arena);
      cache.roots[root_index] = root;
      cache.capture_resource_failure();
      cache.create_scratch();
    } else {
      if (!cache.scratch_arena.has_value()) {
        cache.create_scratch();
      } else {
        cache.scratch_arena->reset();
      }
      root = cache.roots[it->second];
    }
    formula_arena = &*cache.scratch_arena;
  } else {
    root = eval::parse_formula_entry(source, *ctx.arena);
  }
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }

  const std::int32_t row_delta = static_cast<std::int32_t>(ctx.target.row) - static_cast<std::int32_t>(ctx.anchor.row);
  const std::int32_t col_delta = static_cast<std::int32_t>(ctx.target.col) - static_cast<std::int32_t>(ctx.anchor.col);

  const parser::AstNode* shifted = parser::shift_relative_refs(*root, *formula_arena, row_delta, col_delta);
  if (shifted == nullptr) {
    if (ctx.formula_cache != nullptr) {
      ctx.formula_cache->capture_resource_failure();
    }
    return Value::error(ErrorCode::Name);
  }

  const eval::EvalContext target_ctx = ctx.eval_ctx->with_formula_cell(ctx.target.row, ctx.target.col);
  eval::EvalState* const caller_state = ctx.eval_ctx->state();
  if (ctx.formula_cache != nullptr && caller_state != nullptr && !caller_state->observing_reads()) {
    // Scratch-backed results must not enter the caller's memo; read observers keep the caller state.
    eval::EvalState local_state;
    const eval::EvalContext local_ctx = target_ctx.with_state(local_state);
    Value result = eval::evaluate_first_element(*shifted, *formula_arena, *ctx.registry, local_ctx);
    ctx.formula_cache->capture_resource_failure(local_state.out_of_memory());
    return result;
  }
  Value result = eval::evaluate_first_element(*shifted, *formula_arena, *ctx.registry, target_ctx);
  if (ctx.formula_cache != nullptr) {
    ctx.formula_cache->capture_resource_failure();
  }
  return result;
}

std::optional<LiteralOperand> cell_is_operand(const std::string& source, const CFEvalContext& ctx) {
  auto literal = parse_literal(source);
  if (literal.has_value()) {
    return literal;
  }
  const Value evaluated = parse_shift_evaluate(source, ctx);
  return value_to_operand(evaluated);
}

// ---------------------------------------------------------------------------
// Population gathering / caching.
// ---------------------------------------------------------------------------

bool cf_values_equal(const Value& lhs, const Value& rhs) {
  if (lhs.is_number() && rhs.is_number()) {
    return lhs.as_number() == rhs.as_number();
  }
  if (lhs.is_boolean() && rhs.is_boolean()) {
    return lhs.as_boolean() == rhs.as_boolean();
  }
  if (lhs.is_text() && rhs.is_text()) {
    return eval::collate(lhs.as_text(), rhs.as_text(), eval::current_eval_profile(),
                         eval::CollationStrength::kCompare) == 0;
  }
  return false;
}

namespace {

// Small rectangles avoid the callback machinery and retain the direct path
// that dominates ordinary CF sqrefs. Larger rectangles walk populated cells
// through `Sheet::cells_in_range`, so two distant populated corners do not
// turn their entire bounding rectangle into work.
constexpr std::uint64_t kSqrefDenseThresholdCells = 65536U;

// Keep the lightweight grid constants in `cf_types.h` in sync with the
// canonical `Sheet` extents used everywhere else.
static_assert(kCfMaxRows == Sheet::kMaxRows, "kCfMaxRows must match Sheet::kMaxRows");
static_assert(kCfMaxCols == Sheet::kMaxCols, "kCfMaxCols must match Sheet::kMaxCols");

// Calls `visit` with the effective values relevant to population rules in
// `sqref`. Blank cells never contribute to numeric populations or duplicate
// counts, so oversized ranges can visit populated cells only. Each range is
// traversed independently to preserve multiplicity for overlapping sqrefs.
template <typename Visit>
void for_each_sqref_value(const std::vector<CFCellRange>& sqref, const Sheet& sheet, Visit visit) {
  for (const CFCellRange& range : sqref) {
    const std::uint32_t r0 = range.first.row;
    const std::uint32_t c0 = range.first.col;
    const std::uint32_t r1 = std::min(range.last.row, kCfMaxRows - 1U);
    const std::uint32_t c1 = std::min(range.last.col, kCfMaxCols - 1U);
    if (r0 > r1 || c0 > c1 || r0 >= kCfMaxRows || c0 >= kCfMaxCols) {
      continue;
    }

    if (utils::RectRange(r0, c0, r1, c1).size() <= kSqrefDenseThresholdCells) {
      for (auto [row, col] : utils::RectRange(r0, c0, r1, c1)) {
        visit(sheet.resolve_cell_value(row, col));
      }
      continue;
    }

    struct VisitContext {
      Visit* visit;
    } visit_context{&visit};
    const auto visit_populated_cell = [](const Sheet::RangeCell& cell, void* opaque) {
      auto* context = static_cast<VisitContext*>(opaque);
      (*context->visit)(cell.value);
    };

    // A spill in the rectangle makes `cells_in_range` probe every row, so bound it first.
    const auto extent = sheet.populated_extent(r0, c0, r1, c1);
    if (!extent.has_value()) {
      continue;
    }
    const std::uint32_t walk_r1 = std::min(r1, extent->last_row);
    const std::uint32_t walk_c1 = std::min(c1, extent->last_col);
    std::uint64_t cursor = 0;
    while (cursor != Sheet::kCellCursorEnd) {
      cursor = sheet.cells_in_range(r0, c0, walk_r1, walk_c1, cursor, std::numeric_limits<std::uint32_t>::max(),
                                    visit_populated_cell, &visit_context);
    }
  }
}

}  // namespace

std::size_t count_matches_in_sqref(const Value& target, const std::vector<CFCellRange>& sqref, const Sheet& sheet) {
  std::size_t count = 0;
  for_each_sqref_value(sqref, sheet, [&](const Value& cell) {
    if (cf_values_equal(target, cell)) {
      ++count;
    }
  });
  return count;
}

std::vector<double> collect_numeric_values(const std::vector<CFCellRange>& sqref, const Sheet& sheet) {
  std::vector<double> values;
  for_each_sqref_value(sqref, sheet, [&](const Value& cell) {
    if (cell.is_number()) {
      values.push_back(cell.as_number());
    }
  });
  return values;
}

double mean_of(const std::vector<double>& values) {
  double sum = 0.0;
  for (double sample : values) {
    sum += sample;
  }
  return sum / static_cast<double>(values.size());
}

double sample_stddev(const std::vector<double>& values, double mean) {
  if (values.size() < 2) {
    return 0.0;
  }
  double sumsq = 0.0;
  for (double sample : values) {
    const double delta = sample - mean;
    sumsq += delta * delta;
  }
  return std::sqrt(sumsq / static_cast<double>(values.size() - 1));
}

ColorScalePopulation gather_population(const std::vector<CFCellRange>& sqref, const Sheet& sheet) {
  std::vector<double> values = collect_numeric_values(sqref, sheet);
  std::sort(values.begin(), values.end());
  ColorScalePopulation pop;
  if (!values.empty()) {
    pop.min = values.front();
    pop.max = values.back();
    // Reduced once in ascending order; AboveAverage reads these per cell.
    pop.mean = mean_of(values);
    pop.sample_stddev = sample_stddev(values, pop.mean);
  }
  pop.sorted = std::move(values);
  return pop;
}

DuplicateValuePopulation gather_duplicate_population(const std::vector<CFCellRange>& sqref, const Sheet& sheet,
                                                     ExcelProfile profile) {
  DuplicateValuePopulation population;
  population.profile = profile;
  for_each_sqref_value(sqref, sheet, [&](const Value& cell) {
    if ((!cell.is_number() && !cell.is_boolean() && !cell.is_text()) ||
        (cell.is_number() && std::isnan(cell.as_number()))) {
      return;
    }
    std::string key;
    if (eval::append_value_identity_key(cell, profile, eval::CollationStrength::kCompare, &key)) {
      ++population.keys[key];
    }
  });
  return population;
}

std::size_t count_matches_in_population(const Value& target, const DuplicateValuePopulation& population) {
  if ((!target.is_number() && !target.is_boolean() && !target.is_text()) ||
      (target.is_number() && std::isnan(target.as_number()))) {
    return 0;
  }
  std::string key;
  if (!eval::append_value_identity_key(target, population.profile, eval::CollationStrength::kCompare, &key)) {
    return 0;
  }
  const auto it = population.keys.find(key);
  return it == population.keys.end() ? 0U : it->second;
}

const ColorScalePopulation& ensure_population(const CFEvalContext& ctx, ColorScalePopulation& fallback) {
  if (ctx.cached_population != nullptr) {
    return *ctx.cached_population;
  }
  fallback = gather_population(*ctx.sqref, *ctx.eval_ctx->current_sheet());
  return fallback;
}

const ColorScalePopulation* numeric_population_for_cell(const Value& cell_value, const CFEvalContext& ctx,
                                                        ColorScalePopulation& fallback) {
  if (ctx.sqref == nullptr || ctx.eval_ctx == nullptr || ctx.eval_ctx->current_sheet() == nullptr) {
    return nullptr;
  }
  if (!cell_value.is_number()) {
    return nullptr;
  }
  const ColorScalePopulation& pop = ensure_population(ctx, fallback);
  if (pop.sorted.empty()) {
    return nullptr;
  }
  return &pop;
}

// ---------------------------------------------------------------------------
// Numeric helpers.
// ---------------------------------------------------------------------------

double interval_fraction(double value, double low, double high) {
  const double span = high - low;
  if (std::isfinite(span) || !std::isfinite(value) || !std::isfinite(low) || !std::isfinite(high)) {
    return (value - low) / span;
  }

  // Opposite-sign finite endpoints can overflow the span; rescale by the larger magnitude.
  const double scale = std::max(std::abs(low), std::abs(high));
  return (value / scale - low / scale) / (high / scale - low / scale);
}

double interpolate_interval(double low, double high, double fraction) {
  const double span = high - low;
  if (std::isfinite(span) || !std::isfinite(low) || !std::isfinite(high) || fraction < 0.0 || fraction > 1.0) {
    return low + fraction * span;
  }
  if (fraction == 0.0) {
    return low;
  }
  if (fraction == 1.0) {
    return high;
  }
  return (1.0 - fraction) * low + fraction * high;
}

double percentile_inc(const std::vector<double>& sorted, double percentile) {
  const std::size_t count = sorted.size();
  if (count == 1 || !(percentile > 0.0)) {
    return sorted.front();
  }
  if (percentile >= 1.0) {
    return sorted.back();
  }
  const double position = percentile * static_cast<double>(count - 1);
  const auto lower_index = static_cast<std::size_t>(std::floor(position));
  if (lower_index + 1 >= count) {
    return sorted.back();
  }
  const double fraction = position - static_cast<double>(lower_index);
  return interpolate_interval(sorted[lower_index], sorted[lower_index + 1], fraction);
}

std::optional<double> parse_double(std::string_view source) {
  if (source.empty()) {
    return std::nullopt;
  }
  // Shared parser with the engine's text-to-number coercion so CFVO
  // thresholds parse identically to `=`/`<`/`>` numeric operands.
  double parsed = 0.0;
  if (parse_double_exact(source, &parsed)) {
    return parsed;
  }
  return std::nullopt;
}

// ---------------------------------------------------------------------------
// Date helpers.
// ---------------------------------------------------------------------------

namespace {

constexpr int kMonthsPerYear = 12;

}  // namespace

int weekday_sunday_one(double serial_floor, bool date1904) {
  return date_time::weekday_sun0(serial_floor, date1904) + 1;
}

double sunday_of_week(double serial_floor, bool date1904) {
  const int weekday = weekday_sunday_one(serial_floor, date1904);
  return serial_floor - (weekday - 1);
}

YearMonth year_month_from_serial(double serial_floor, bool date1904) {
  const date_time::YMD ymd = date_time::ymd_from_serial(serial_floor, date1904);
  return {ymd.y, ymd.m};
}

YearMonth shift_year_month(YearMonth anchor, int delta_months) {
  const int month_index = static_cast<int>(anchor.month) - 1 + delta_months;
  // Floor-divide month_index by 12 so negative offsets borrow a year
  // correctly (e.g. month_index = -1 → year - 1, month = 12).
  int year_delta = month_index / kMonthsPerYear;
  int normalised_index = month_index % kMonthsPerYear;
  if (normalised_index < 0) {
    normalised_index += kMonthsPerYear;
    --year_delta;
  }
  return {anchor.year + year_delta, static_cast<unsigned>(normalised_index + 1)};
}

}  // namespace formulon::cf::helpers
