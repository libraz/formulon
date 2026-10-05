#include "auto_filter_eval.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstdlib>
#include <functional>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_evaluator.h"
#include "cf/cf_match.h"
#include "color_resolve.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/text_format/display_text.h"
#include "eval/wildcard.h"
#include "sheet.h"
#include "style_resolve.h"
#include "styles.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

/// Shared state of one evaluation: the sheet, the evaluator collaborators
/// and the per-column cell values read once.
struct Run {
  const Workbook& wb;
  const Sheet& sheet;
  Arena* arena;
  const eval::FunctionRegistry* registry;
  const eval::EvalContext* eval_ctx;
  double today;
  std::uint32_t first_body_row;
  std::uint32_t body_rows;
};

// --- Text helpers -----------------------------------------------------------

bool is_blank_value(const Value& v) {
  return v.is_blank() || (v.is_text() && v.as_text().empty());
}

/// Sort class of one byte in Excel's text order: symbols and spaces, then
/// digits, then letters, then everything beyond ASCII.
int collation_class(unsigned char c) {
  if (c >= 0x80U) {
    return 3;
  }
  if (c >= '0' && c <= '9') {
    return 1;
  }
  const char lower = strings::ascii_to_lower(static_cast<char>(c));
  if (lower >= 'a' && lower <= 'z') {
    return 2;
  }
  return 0;
}

/// Case-insensitive text comparison in Excel's sort order (-1 / 0 / +1).
int collate(std::string_view a, std::string_view b) {
  const std::size_t n = std::min(a.size(), b.size());
  for (std::size_t i = 0; i < n; ++i) {
    const auto ca = static_cast<unsigned char>(strings::ascii_to_lower(a[i]));
    const auto cb = static_cast<unsigned char>(strings::ascii_to_lower(b[i]));
    if (ca == cb) {
      continue;
    }
    const int ka = collation_class(ca);
    const int kb = collation_class(cb);
    if (ka != kb) {
      return ka < kb ? -1 : 1;
    }
    return ca < cb ? -1 : 1;
  }
  if (a.size() != b.size()) {
    return a.size() < b.size() ? -1 : 1;
  }
  return 0;
}

/// Parses a criterion value as a number; the whole string must be numeric.
std::optional<double> parse_number(std::string_view text) {
  if (text.empty()) {
    return std::nullopt;
  }
  const std::string owned(text);
  char* end = nullptr;
  const double v = std::strtod(owned.c_str(), &end);
  if (end != owned.c_str() + owned.size() || !std::isfinite(v)) {
    return std::nullopt;
  }
  return v;
}

bool ordered(FilterOperator op, int cmp) {
  switch (op) {
    case FilterOperator::kLessThan:
      return cmp < 0;
    case FilterOperator::kLessThanOrEqual:
      return cmp <= 0;
    case FilterOperator::kGreaterThanOrEqual:
      return cmp >= 0;
    case FilterOperator::kGreaterThan:
      return cmp > 0;
    case FilterOperator::kEqual:
      return cmp == 0;
    case FilterOperator::kNotEqual:
      return cmp != 0;
  }
  return false;
}

// --- Per-kind matchers ------------------------------------------------------

bool match_date_group(const DateGroupItem& g, double serial, bool date1904) {
  if (serial < 0.0) {
    return false;
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(std::floor(serial), date1904);
  const date_time::HMS hms = date_time::hms_from_fraction(serial);
  if (ymd.y != static_cast<int>(g.year)) {
    return false;
  }
  if (g.grouping >= DateTimeGrouping::kMonth && ymd.m != g.month) {
    return false;
  }
  if (g.grouping >= DateTimeGrouping::kDay && ymd.d != g.day) {
    return false;
  }
  if (g.grouping >= DateTimeGrouping::kHour && hms.h != g.hour) {
    return false;
  }
  if (g.grouping >= DateTimeGrouping::kMinute && hms.m != g.minute) {
    return false;
  }
  return g.grouping < DateTimeGrouping::kSecond || hms.s == g.second;
}

bool match_values(const ValueFilters& f, const Run& run, std::uint32_t row, std::uint32_t col, const Value& v) {
  const text_format::DisplayText shown = text_format::format_cell_for_display(run.wb, run.sheet, row, col);
  if (shown.status == text_format::DisplayStatus::kOverflow) {
    return false;
  }
  if (shown.text.empty()) {
    return f.blank;
  }
  for (const std::string& want : f.values) {
    if (strings::case_insensitive_eq(want, shown.text)) {
      return true;
    }
  }
  if (v.is_number()) {
    for (const DateGroupItem& g : f.date_groups) {
      if (match_date_group(g, v.as_number(), run.wb.date1904())) {
        return true;
      }
    }
  }
  return false;
}

/// Equality of one custom condition: a single space (or nothing) names the
/// blank cell, wildcards apply to text only, and a numeric literal compares
/// numerically against numbers.
bool custom_equals(const CustomFilter& c, const Value& v, const Run& run, std::uint32_t row, std::uint32_t col) {
  if (c.val.empty() || c.val == " ") {
    return is_blank_value(v);
  }
  if (is_blank_value(v)) {
    return false;
  }
  if (v.is_number()) {
    const std::optional<double> n = parse_number(c.val);
    return n.has_value() && !eval::scan_has_wildcard(c.val) && *n == v.as_number();
  }
  if (v.is_text()) {
    return eval::wildcard_match(c.val, v.as_text());
  }
  const text_format::DisplayText shown = text_format::format_cell_for_display(run.wb, run.sheet, row, col);
  return eval::wildcard_match(c.val, shown.text);
}

bool match_custom_one(const CustomFilter& c, const Value& v, const Run& run, std::uint32_t row, std::uint32_t col) {
  if (c.op == FilterOperator::kEqual) {
    return custom_equals(c, v, run, row, col);
  }
  if (c.op == FilterOperator::kNotEqual) {
    return !custom_equals(c, v, run, row, col);
  }
  if (const std::optional<double> n = parse_number(c.val)) {
    if (!v.is_number()) {
      return false;
    }
    const double x = v.as_number();
    return ordered(c.op, x < *n ? -1 : (x > *n ? 1 : 0));
  }
  if (!v.is_text() || v.as_text().empty()) {
    return false;
  }
  return ordered(c.op, collate(v.as_text(), c.val));
}

bool match_custom(const CustomFilters& f, const Value& v, const Run& run, std::uint32_t row, std::uint32_t col) {
  if (f.filters.empty()) {
    return true;
  }
  bool result = match_custom_one(f.filters[0], v, run, row, col);
  for (std::size_t i = 1; i < f.filters.size(); ++i) {
    const bool next = match_custom_one(f.filters[i], v, run, row, col);
    result = f.and_join ? (result && next) : (result || next);
  }
  return result;
}

/// The value a cell must reach to be in the top / bottom set: the k-th
/// number in rank order, so every tie with it is kept. Percent counts round
/// down and keep at least one item.
std::optional<double> top10_threshold(const Top10Filter& f, std::vector<double> numbers) {
  if (numbers.empty() || !(f.val > 0.0)) {
    return std::nullopt;
  }
  const double n = static_cast<double>(numbers.size());
  double count = f.percent ? std::floor(n * f.val / 100.0) : std::floor(f.val);
  count = std::max(1.0, std::min(count, n));
  const auto k = static_cast<std::size_t>(count);
  if (f.top) {
    std::sort(numbers.begin(), numbers.end(), std::greater<double>());
  } else {
    std::sort(numbers.begin(), numbers.end());
  }
  return numbers[k - 1U];
}

/// Half-open serial interval `[lo, hi)` of a calendar-relative filter.
struct Period {
  double lo;
  double hi;
};

double month_start(int y, int m, bool date1904) {
  // Normalises month offsets outside 1..12 into the neighbouring years.
  const int zero_based = (y * 12) + (m - 1);
  const int ny = zero_based >= 0 ? zero_based / 12 : -((-zero_based + 11) / 12);
  const int nm = zero_based - (ny * 12) + 1;
  return date_time::serial_from_ymd(ny, static_cast<unsigned>(nm), 1U, date1904);
}

std::optional<Period> relative_period(DynamicFilterType type, double today, bool date1904) {
  const date_time::YMD t = date_time::ymd_from_serial(today, date1904);
  const int y = t.y;
  const int m = static_cast<int>(t.m);
  const double week = today - static_cast<double>(date_time::weekday_sun0(today, date1904));
  const int quarter_month = (((m - 1) / 3) * 3) + 1;
  switch (type) {
    case DynamicFilterType::kYesterday:
      return Period{today - 1.0, today};
    case DynamicFilterType::kToday:
      return Period{today, today + 1.0};
    case DynamicFilterType::kTomorrow:
      return Period{today + 1.0, today + 2.0};
    case DynamicFilterType::kLastWeek:
      return Period{week - 7.0, week};
    case DynamicFilterType::kThisWeek:
      return Period{week, week + 7.0};
    case DynamicFilterType::kNextWeek:
      return Period{week + 7.0, week + 14.0};
    case DynamicFilterType::kLastMonth:
      return Period{month_start(y, m - 1, date1904), month_start(y, m, date1904)};
    case DynamicFilterType::kThisMonth:
      return Period{month_start(y, m, date1904), month_start(y, m + 1, date1904)};
    case DynamicFilterType::kNextMonth:
      return Period{month_start(y, m + 1, date1904), month_start(y, m + 2, date1904)};
    case DynamicFilterType::kLastQuarter:
      return Period{month_start(y, quarter_month - 3, date1904), month_start(y, quarter_month, date1904)};
    case DynamicFilterType::kThisQuarter:
      return Period{month_start(y, quarter_month, date1904), month_start(y, quarter_month + 3, date1904)};
    case DynamicFilterType::kNextQuarter:
      return Period{month_start(y, quarter_month + 3, date1904), month_start(y, quarter_month + 6, date1904)};
    case DynamicFilterType::kLastYear:
      return Period{month_start(y - 1, 1, date1904), month_start(y, 1, date1904)};
    case DynamicFilterType::kThisYear:
      return Period{month_start(y, 1, date1904), month_start(y + 1, 1, date1904)};
    case DynamicFilterType::kNextYear:
      return Period{month_start(y + 1, 1, date1904), month_start(y + 2, 1, date1904)};
    case DynamicFilterType::kYearToDate:
      return Period{month_start(y, 1, date1904), today + 1.0};
    default:
      return std::nullopt;
  }
}

bool match_dynamic(const DynamicFilter& f, const Value& v, const Run& run, std::optional<double> average) {
  if (f.type == DynamicFilterType::kNull) {
    return true;
  }
  if (!v.is_number()) {
    return false;
  }
  const double x = v.as_number();
  const bool date1904 = run.wb.date1904();
  switch (f.type) {
    case DynamicFilterType::kAboveAverage:
      return average.has_value() && x > *average;
    case DynamicFilterType::kBelowAverage:
      return average.has_value() && x < *average;
    default:
      break;
  }
  if (f.type >= DynamicFilterType::kQ1 && f.type <= DynamicFilterType::kQ4) {
    if (x < 0.0) {
      return false;
    }
    const unsigned q = static_cast<unsigned>(f.type) - static_cast<unsigned>(DynamicFilterType::kQ1);
    const unsigned month = date_time::ymd_from_serial(std::floor(x), date1904).m;
    return (month - 1U) / 3U == q;
  }
  if (f.type >= DynamicFilterType::kM1 && f.type <= DynamicFilterType::kM12) {
    if (x < 0.0) {
      return false;
    }
    const unsigned want = static_cast<unsigned>(f.type) - static_cast<unsigned>(DynamicFilterType::kM1) + 1U;
    return date_time::ymd_from_serial(std::floor(x), date1904).m == want;
  }
  const std::optional<Period> p = relative_period(f.type, run.today, date1904);
  return p.has_value() && x >= p->lo && x < p->hi;
}

// --- Colour and icon filters ------------------------------------------------

ColorSpec literal_or_spec(const ColorSpec& spec, std::uint32_t argb) {
  if (spec.kind != ColorSpec::Kind::kNone || argb == 0U) {
    return spec;
  }
  ColorSpec literal;
  literal.kind = ColorSpec::Kind::kRgb;
  literal.rgb = argb;
  return literal;
}

bool has_color(const ColorSpec& spec, std::uint32_t argb) {
  return spec.kind != ColorSpec::Kind::kNone || argb != 0U;
}

std::uint32_t rgb_of(const Workbook& wb, const ColorSpec& spec, std::uint32_t argb, ColorContext context) {
  return resolve_color(wb, literal_or_spec(spec, argb), context).argb & 0x00FFFFFFU;
}

/// Solid colour of a dxf fill: its foreground, else its background.
std::uint32_t dxf_fill_rgb(const Workbook& wb, const FillRecord& fill) {
  if (has_color(fill.fg, fill.fg_argb)) {
    return rgb_of(wb, fill.fg, fill.fg_argb, ColorContext::kFillForeground);
  }
  return rgb_of(wb, fill.bg, fill.bg_argb, ColorContext::kFillBackground);
}

/// The colours a cell shows: its effective style with the first matching
/// conditional format's fill and font colour laid over it.
struct CellColors {
  bool has_fill = false;
  std::uint32_t fill_rgb = 0;
  std::uint32_t font_rgb = 0;
  std::optional<cf::IconRender> icon;
};

CellColors cell_colors(const Run& run, std::uint32_t row, std::uint32_t col) {
  CellColors out;
  const EffectiveStyle style = effective_style(run.wb, run.sheet, row, col);
  out.has_fill = style.fill_pattern != 0U;
  out.fill_rgb = style.fill_foreground.argb & 0x00FFFFFFU;
  out.font_rgb = style.font_color.argb & 0x00FFFFFFU;
  if (run.sheet.conditional_formats().empty() || run.arena == nullptr || run.eval_ctx == nullptr) {
    return out;
  }
  cf::CFHost host;
  host.arena = run.arena;
  host.registry = run.registry;
  host.eval_ctx = run.eval_ctx;
  host.today_serial = run.today;
  bool fill_set = false;
  bool font_set = false;
  const auto& dxfs = run.wb.styles().dxfs;
  for (const cf::CFMatch& m : cf::evaluate_cf_at(run.sheet, CellAddress{row, col}, host)) {
    if (m.icon_render.has_value() && !out.icon.has_value()) {
      out.icon = m.icon_render;
    }
    if (m.resolved_fill_color.has_value() && !fill_set) {
      const cf::Color& c = *m.resolved_fill_color;
      out.fill_rgb = c.is_symbolic()
                         ? rgb_of(run.wb, c.spec, 0U, ColorContext::kFillForeground)
                         : (static_cast<std::uint32_t>(c.r) << 16U) | (static_cast<std::uint32_t>(c.g) << 8U) | c.b;
      out.has_fill = true;
      fill_set = true;
    }
    if (!m.dxf_id.has_value() || *m.dxf_id >= dxfs.size()) {
      continue;
    }
    const DifferentialFormat& dxf = dxfs[*m.dxf_id];
    if (!fill_set && dxf.has_fill &&
        (has_color(dxf.fill.fg, dxf.fill.fg_argb) || has_color(dxf.fill.bg, dxf.fill.bg_argb))) {
      out.fill_rgb = dxf_fill_rgb(run.wb, dxf.fill);
      out.has_fill = true;
      fill_set = true;
    }
    if (!font_set && dxf.has_font && has_color(dxf.font.color, dxf.font.color_argb)) {
      out.font_rgb = rgb_of(run.wb, dxf.font.color, dxf.font.color_argb, ColorContext::kFont);
      font_set = true;
    }
  }
  return out;
}

/// Colour filters name a dxf. A fill filter whose dxf has no solid fill
/// selects unfilled cells. Excel stores a font-colour filter's colour as the
/// dxf's fill foreground, so a dxf font colour is used only when present.
bool match_color(const ColorFilter& f, const Run& run, std::uint32_t row, std::uint32_t col) {
  const CellColors colors = cell_colors(run, row, col);
  const auto& dxfs = run.wb.styles().dxfs;
  const DifferentialFormat* dxf = (f.dxf_id.has_value() && *f.dxf_id < dxfs.size()) ? &dxfs[*f.dxf_id] : nullptr;
  if (f.cell_color) {
    if (dxf == nullptr || !dxf->has_fill || dxf->fill.pattern == 0U) {
      return !colors.has_fill;
    }
    return colors.has_fill && colors.fill_rgb == dxf_fill_rgb(run.wb, dxf->fill);
  }
  std::uint32_t want = rgb_of(run.wb, ColorSpec{}, 0U, ColorContext::kFont);
  if (dxf != nullptr && dxf->has_font && has_color(dxf->font.color, dxf->font.color_argb)) {
    want = rgb_of(run.wb, dxf->font.color, dxf->font.color_argb, ColorContext::kFont);
  } else if (dxf != nullptr && dxf->has_fill) {
    want = dxf_fill_rgb(run.wb, dxf->fill);
  }
  return colors.font_rgb == want;
}

bool match_icon(const IconFilter& f, const Run& run, std::uint32_t row, std::uint32_t col) {
  const CellColors colors = cell_colors(run, row, col);
  if (!f.icon_id.has_value()) {
    return !colors.icon.has_value();
  }
  return colors.icon.has_value() && colors.icon->set_name == f.icon_set && colors.icon->icon_index == *f.icon_id;
}

// --- Column driver ----------------------------------------------------------

/// ANDs one filter column's matches into `visible`.
void apply_column(const FilterColumn& column, const Run& run, std::uint32_t col, std::vector<bool>& visible) {
  if (column.kind == FilterKind::kNone) {
    return;
  }
  std::vector<Value> values;
  values.reserve(run.body_rows);
  std::vector<double> numbers;
  for (std::uint32_t i = 0; i < run.body_rows; ++i) {
    values.push_back(run.sheet.resolve_cell_value(run.first_body_row + i, col));
    if (values.back().is_number()) {
      numbers.push_back(values.back().as_number());
    }
  }
  std::optional<double> threshold;
  if (column.kind == FilterKind::kTop10) {
    threshold = top10_threshold(column.top10, numbers);
  }
  std::optional<double> average;
  if (column.kind == FilterKind::kDynamic && !numbers.empty()) {
    double sum = 0.0;
    for (double x : numbers) {
      sum += x;
    }
    average = sum / static_cast<double>(numbers.size());
  }
  for (std::uint32_t i = 0; i < run.body_rows; ++i) {
    if (!visible[i]) {
      continue;
    }
    const std::uint32_t row = run.first_body_row + i;
    const Value& v = values[i];
    bool match = true;
    switch (column.kind) {
      case FilterKind::kValues:
        match = match_values(column.values, run, row, col, v);
        break;
      case FilterKind::kCustom:
        match = match_custom(column.custom, v, run, row, col);
        break;
      case FilterKind::kTop10:
        match = threshold.has_value() && v.is_number() &&
                (column.top10.top ? v.as_number() >= *threshold : v.as_number() <= *threshold);
        break;
      case FilterKind::kDynamic:
        match = match_dynamic(column.dynamic, v, run, average);
        break;
      case FilterKind::kColor:
        match = match_color(column.color, run, row, col);
        break;
      case FilterKind::kIcon:
        match = match_icon(column.icon, run, row, col);
        break;
      case FilterKind::kNone:
        break;
    }
    visible[i] = match;
  }
}

double today_serial(const eval::EvalContext& ctx, bool date1904) {
  const date_time::CivilTime now = ctx.wall_clock();
  return date_time::serial_from_ymd(now.date.y, now.date.m, now.date.d, date1904);
}

}  // namespace

Expected<std::vector<bool>, Error> evaluate_auto_filter(const Workbook& wb, const Sheet& sheet,
                                                        const AutoFilter& filter, const AutoFilterEvalDeps& deps) {
  if (filter.is_opaque()) {
    return make_error(FormulonErrorCode::kAutoFilterInvalid, "evaluate_auto_filter: AutoFilter is opaque");
  }
  RETURN_IF_ERROR(validate_auto_filter(filter));

  // Collaborators the caller did not supply are built here and live for the call.
  std::optional<Arena> own_arena;
  eval::EvalState own_state;
  std::optional<eval::EvalContext> own_ctx;
  Arena* arena = deps.arena;
  if (arena == nullptr) {
    own_arena.emplace(/*initial_chunk_bytes=*/4096, kMaxEvalArenaBytes);
    arena = &*own_arena;
  }
  const eval::EvalContext* ctx = deps.eval_ctx;
  if (ctx == nullptr) {
    own_ctx.emplace(eval::EvalContext(wb, sheet, own_state).with_pinned_now(wb.pinned_now()));
    ctx = &*own_ctx;
  }

  const MergeRange& range = filter.range;
  const std::uint32_t body_rows = range.last_row - range.first_row;
  Run run{wb,
          sheet,
          arena,
          deps.registry != nullptr ? deps.registry : &eval::default_registry(),
          ctx,
          today_serial(*ctx, wb.date1904()),
          range.first_row + 1U,
          body_rows};
  std::vector<bool> visible(body_rows, true);
  for (const FilterColumn& column : filter.columns) {
    apply_column(column, run, range.first_col + column.col_id, visible);
  }
  return visible;
}

Expected<void, Error> apply_auto_filter(Workbook& wb, std::size_t sheet_index, const AutoFilter& filter,
                                        const AutoFilterEvalDeps& deps) {
  if (sheet_index >= wb.sheet_count()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "apply_auto_filter: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  Sheet& sheet = wb.sheet(sheet_index);
  auto visible = evaluate_auto_filter(wb, sheet, filter, deps);
  if (!visible) {
    return visible.error();
  }
  const std::uint32_t first = filter.range.first_row + 1U;
  std::vector<RowLayout>& overrides = sheet.mutable_layout().row_overrides;
  bool changed = false;
  for (std::uint32_t i = 0; i < visible.value().size(); ++i) {
    const std::uint32_t row = first + i;
    const bool hide = !visible.value()[i];
    auto it = std::find_if(overrides.begin(), overrides.end(), [row](const RowLayout& r) { return r.row == row; });
    if (it == overrides.end()) {
      if (!hide) {
        continue;
      }
      RowLayout fresh;
      fresh.row = row;
      overrides.push_back(fresh);
      it = overrides.end() - 1;
    }
    if (it->hidden != hide) {
      it->hidden = hide;
      changed = true;
    }
  }
  if (changed) {
    wb.mark_row_visibility_dependents_dirty();
  }
  return Expected<void, Error>::Ok();
}

bool sheet_has_filter_criteria(const Workbook* wb, const Sheet& sheet) noexcept {
  if (const AutoFilter* own = sheet.auto_filter(); own != nullptr && own->has_criteria()) {
    return true;
  }
  if (wb == nullptr) {
    return false;
  }
  for (const TableMetadata& table : wb->tables()) {
    const AutoFilter* f = table.auto_filter_xml.get();
    if (f == nullptr || !f->has_criteria() || table.sheet_index >= wb->sheet_count()) {
      continue;
    }
    if (&wb->sheet(table.sheet_index) == &sheet) {
      return true;
    }
  }
  return false;
}

}  // namespace formulon
