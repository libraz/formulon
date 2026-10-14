#include "chart/chart_data.h"

#include <algorithm>
#include <cmath>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_source.h"
#include "eval/eval_profile_scope.h"
#include "eval/function_registry.h"
#include "eval/text_format/display_text.h"
#include "num_fmt_locale.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "style_resolve.h"
#include "styles.h"
#include "utils/number_text.h"
#include "utils/resource_budget.h"
#include "workbook.h"

namespace formulon {
namespace {

/// One point of a source before Excel's plotting rules are applied.
struct Pt {
  enum class Kind : std::uint8_t { kBlank, kNumber, kText, kBool, kError };
  Kind kind = Kind::kBlank;
  double number = 0.0;
  bool is_na = false;
  bool empty_text = false;  ///< `=""`: plots as text (0) but labels as a blank.
  std::string display;      ///< Displayed text; filled only for labels and names.
};

/// A source flattened into points, plus the upper levels of a multi-level category.
struct Gathered {
  bool resolved = false;
  std::vector<SheetRect> ranges;  ///< As referenced (whole columns untrimmed).
  std::vector<Pt> cells;
  std::vector<std::vector<std::string>> levels;  ///< Outermost first.
  std::string format_code;
  bool had_cells = false;  ///< The reference covered cells before hidden ones were dropped.
};

/// How blanks and `#N/A` plot for one plot group.
struct PlotRules {
  ChartBlanksAs blanks = ChartBlanksAs::kZero;
  bool na_as_blank = false;
  bool connects = false;  ///< Line / scatter / radar draw across a connect point.
};

Error chart_cell_budget_error(std::uint64_t cells) {
  return make_error(FormulonErrorCode::kSecResourceLimit, "chart: source exceeds materialization limit",
                    "cells=" + std::to_string(cells) + " ceiling=" + std::to_string(kMaxRangeExpansionCells));
}

Error chart_point_limit_error(std::uint64_t points) {
  return make_error(FormulonErrorCode::kSecResourceLimit, "chart: point count exceeds Excel row limit",
                    "points=" + std::to_string(points) + " ceiling=" + std::to_string(Sheet::kMaxRows));
}

bool connects(ChartPlotType type) {
  return type == ChartPlotType::kLine || type == ChartPlotType::kScatter || type == ChartPlotType::kRadar;
}

bool is_xy(ChartPlotType type) {
  return type == ChartPlotType::kScatter || type == ChartPlotType::kBubble;
}

std::string count_label(std::size_t n) {
  char buf[24];
  const int len = format_unsigned(buf, sizeof(buf), n);
  return std::string(buf, len > 0 ? static_cast<std::size_t>(len) : 0U);
}

class Reader {
 public:
  /// `pivot` takes points from the cache: a pivot chart's references span the
  /// pivot's grand total, which Excel leaves out, and its series are named by
  /// the pivot rather than by the cell.
  Reader(const Workbook& wb, std::size_t host_sheet, bool plot_vis_only, bool pivot)
      : wb_(wb), host_sheet_(host_sheet), plot_vis_only_(plot_vis_only), pivot_(pivot) {}

  /// Resolves `src` from its cells, or from its literal / cache when it has no usable reference.
  Expected<Gathered, Error> gather(const ChartDataSource& src, bool want_display, bool multi_level) const {
    Gathered out;
    if (src.kind == ChartDataSource::Kind::kLiteral) {
      out.resolved = true;
      auto materialized = from_points(src.literal, {}, src.point_count, src.cache_format_code, want_display, out);
      if (!materialized) {
        return materialized.error();
      }
      return out;
    }
    if (src.kind != ChartDataSource::Kind::kRef) {
      out.resolved = true;
      return out;
    }
    // A category filter narrows c:f to the shown cells; the data keeps every category so
    // that the model's hidden category indices address it.
    const std::string& formula = src.full_formula.empty() ? src.formula : src.full_formula;
    auto resolved = resolve_chart_source(wb_, host_sheet_, formula, eval::default_registry());
    if (!resolved) {
      auto materialized =
          from_points(src.cache, src.cache_levels, src.point_count, src.cache_format_code, want_display, out);
      if (!materialized) {
        return materialized.error();
      }
      return out;
    }
    out.resolved = true;
    out.ranges = resolved.value().areas;
    if (pivot_) {
      auto materialized =
          from_points(src.cache, src.cache_levels, src.point_count, src.cache_format_code, want_display, out);
      if (!materialized) {
        return materialized.error();
      }
      return out;
    }
    const std::vector<SheetRect>& areas = out.ranges;
    if (multi_level && areas.size() == 1 && areas[0].top != areas[0].bottom && areas[0].left != areas[0].right) {
      const SheetRect r = trim(areas[0]);
      if (r.top <= r.bottom && r.left <= r.right) {
        const std::uint64_t cells =
            (static_cast<std::uint64_t>(r.bottom) - r.top + 1U) * (static_cast<std::uint64_t>(r.right) - r.left + 1U);
        if (cells > kMaxRangeExpansionCells) {
          return chart_cell_budget_error(cells);
        }
      }
      gather_levels(r, want_display, out);
      return out;
    }
    std::vector<Pos> positions;
    std::vector<SheetRect> trimmed;
    trimmed.reserve(areas.size());
    std::uint64_t cells = 0;
    for (const SheetRect& area : areas) {
      const SheetRect r = trim(area);
      if (r.top > r.bottom || r.left > r.right) {
        continue;
      }
      const std::uint64_t area_cells =
          (static_cast<std::uint64_t>(r.bottom) - r.top + 1U) * (static_cast<std::uint64_t>(r.right) - r.left + 1U);
      if (area_cells > kMaxRangeExpansionCells - std::min(cells, kMaxRangeExpansionCells)) {
        return chart_cell_budget_error(cells + area_cells);
      }
      cells += area_cells;
      trimmed.push_back(r);
    }
    positions.reserve(static_cast<std::size_t>(cells));
    for (const SheetRect& r : trimmed) {
      for (std::uint32_t col = r.left; col <= r.right && r.top <= r.bottom; ++col) {
        for (std::uint32_t row = r.top; row <= r.bottom; ++row) {
          positions.push_back(Pos{r.sheet, row, col});
        }
      }
    }
    out.had_cells = !positions.empty();
    out.format_code = format_code_of(positions);
    for (const Pos& p : positions) {
      if (!hidden(p.sheet, p.row, p.col)) {
        out.cells.push_back(read(p, want_display));
      }
    }
    return out;
  }

 private:
  struct Pos {
    std::size_t sheet;
    std::uint32_t row;
    std::uint32_t col;
  };

  /// Whole columns / rows stop at the sheet's populated extent.
  SheetRect trim(SheetRect r) const {
    const bool whole_col = r.top == 0 && r.bottom == Sheet::kMaxRows - 1;
    const bool whole_row = r.left == 0 && r.right == Sheet::kMaxCols - 1;
    if (!whole_col && !whole_row) {
      return r;
    }
    const auto extent = wb_.sheet(r.sheet).populated_extent(0, 0, Sheet::kMaxRows - 1, Sheet::kMaxCols - 1);
    if (!extent.has_value()) {
      r.bottom = 0;
      r.top = 1;  // empty
      return r;
    }
    if (whole_col) {
      r.bottom = std::min(r.bottom, extent->last_row);
    }
    if (whole_row) {
      r.right = std::min(r.right, extent->last_col);
    }
    return r;
  }

  bool row_hidden(const Sheet& sheet, std::uint32_t row) const {
    if (const RowLayout* entry = find_row_override(sheet.layout(), row); entry != nullptr) {
      return entry->hidden;
    }
    return sheet.format_defaults().zero_height;
  }

  bool col_hidden(const Sheet& sheet, std::uint32_t col) const {
    for (const ColumnLayout& span : sheet.layout().columns) {
      if (col >= span.first && col <= span.last) {
        return span.hidden;
      }
    }
    return false;
  }

  /// Hidden rows and columns drop out when only visible cells are plotted;
  /// a hidden sheet's cells do not.
  bool hidden(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) const {
    if (!plot_vis_only_) {
      return false;
    }
    const Sheet& sheet = wb_.sheet(sheet_index);
    return row_hidden(sheet, row) || col_hidden(sheet, col);
  }

  std::string cell_format_code(const Pos& p) const {
    const std::vector<CellXf>& xfs = wb_.styles().cell_xfs;
    if (xfs.empty()) {
      return {};
    }
    const EffectiveXf choice = select_effective_xf(wb_.sheet(p.sheet), p.row, p.col);
    const CellXf* xf = &xfs[choice.xf_index < xfs.size() ? choice.xf_index : 0U];
    return std::string(text_format::number_format_code_for_xf(wb_.styles(), xf, wb_.excel_profile()));
  }

  /// The second point's format, or the only point's.
  std::string format_code_of(const std::vector<Pos>& positions) const {
    if (positions.empty()) {
      return {};
    }
    return cell_format_code(positions[positions.size() >= 2 ? 1 : 0]);
  }

  Pt read(const Pos& p, bool want_display) const {
    const Sheet& sheet = wb_.sheet(p.sheet);
    const Value v = sheet.resolve_cell_value(p.row, p.col);
    Pt cell;
    if (v.is_number()) {
      cell.kind = Pt::Kind::kNumber;
      cell.number = v.as_number();
    } else if (v.is_text()) {
      cell.kind = Pt::Kind::kText;
      cell.empty_text = v.as_text().empty();
    } else if (v.is_boolean()) {
      cell.kind = Pt::Kind::kBool;
      cell.number = v.as_boolean() ? 1.0 : 0.0;
    } else if (v.is_error()) {
      cell.kind = Pt::Kind::kError;
      cell.is_na = v.as_error() == ErrorCode::NA;
    }
    if (want_display && cell.kind != Pt::Kind::kBlank) {
      cell.display = text_format::format_cell_for_display(wb_, sheet, p.row, p.col).text;
    }
    return cell;
  }

  /// A 2-D category: the last column (or row) is the lowest level, the others
  /// carry their group labels at the points where they are not blank.
  void gather_levels(const SheetRect& r, bool want_display, Gathered& out) const {
    if (r.top > r.bottom) {
      return;
    }
    const bool down = (r.bottom - r.top) >= (r.right - r.left);
    const std::uint32_t first = down ? r.top : r.left;
    const std::uint32_t last = down ? r.bottom : r.right;
    const std::uint32_t lvl_first = down ? r.left : r.top;
    const std::uint32_t lvl_last = down ? r.right : r.bottom;
    const auto at = [&](std::uint32_t point, std::uint32_t level) {
      return down ? Pos{r.sheet, point, level} : Pos{r.sheet, level, point};
    };
    out.levels.assign(lvl_last - lvl_first, {});
    std::vector<Pos> lowest;
    for (std::uint32_t i = first; i <= last; ++i) {
      lowest.push_back(at(i, lvl_last));
    }
    out.had_cells = true;
    out.format_code = format_code_of(lowest);
    for (std::uint32_t i = first; i <= last; ++i) {
      const Sheet& sheet = wb_.sheet(r.sheet);
      if (plot_vis_only_ && (down ? row_hidden(sheet, i) : col_hidden(sheet, i))) {
        continue;
      }
      out.cells.push_back(read(at(i, lvl_last), want_display));
      for (std::uint32_t lv = lvl_first; lv < lvl_last; ++lv) {
        out.levels[lv - lvl_first].push_back(read(at(i, lv), true).display);
      }
    }
  }

  /// Sparse literal / cache points to a dense list; missing points are blank.
  Expected<void, Error> from_points(const std::vector<ChartValue>& points,
                                    const std::vector<std::vector<ChartValue>>& levels, std::uint32_t count,
                                    const std::string& format_code, bool want_display, Gathered& out) const {
    out.format_code = format_code;
    std::uint64_t used = 0;
    for (const ChartValue& point : points) {
      used = std::max(used, static_cast<std::uint64_t>(point.idx) + 1U);
    }
    // A whole column / row cache is trimmed like its cells would be.
    const bool whole_axis = count == Sheet::kMaxRows || count == Sheet::kMaxCols;
    const std::uint64_t requested = whole_axis ? used : std::max<std::uint64_t>(count, used);
    if (requested > Sheet::kMaxRows) {
      return chart_point_limit_error(requested);
    }
    const std::uint64_t cells = requested * (static_cast<std::uint64_t>(levels.size()) + 1U);
    if (cells > kMaxRangeExpansionCells) {
      return chart_cell_budget_error(cells);
    }
    const std::size_t n = static_cast<std::size_t>(requested);
    out.cells.assign(n, Pt{});
    for (const ChartValue& v : points) {
      Pt& cell = out.cells[static_cast<std::size_t>(v.idx)];
      if (v.kind == ChartValue::Kind::kNumber) {
        cell.kind = Pt::Kind::kNumber;
        cell.number = v.number;
        if (want_display) {
          const std::string& stored = v.format_code.empty() ? format_code : v.format_code;
          // A cache spells built-in formats in en-US; render by the workbook's own form of that id.
          const std::optional<std::uint16_t> id = builtin_num_fmt_id_for_en_us_code(stored);
          const std::string code = id ? std::string(localized_builtin_num_fmt(*id, wb_.excel_profile())) : stored;
          cell.display =
              text_format::format_value_for_display(Value::number(v.number), code, wb_.date1904(), wb_.excel_profile())
                  .text;
        }
      } else if (v.kind == ChartValue::Kind::kText) {
        cell.is_na = v.text == "#N/A";
        cell.kind = cell.is_na ? Pt::Kind::kError : Pt::Kind::kText;
        cell.display = v.text;
      }
    }
    for (const std::vector<ChartValue>& level : levels) {
      std::vector<std::string> texts(n);
      for (const ChartValue& v : level) {
        if (v.idx < n) {
          texts[v.idx] = v.text;
        }
      }
      out.levels.push_back(std::move(texts));
    }
    return Expected<void, Error>::Ok();
  }

  const Workbook& wb_;
  std::size_t host_sheet_;
  bool plot_vis_only_;
  bool pivot_;
};

/// Rules 3-5: blanks, `#N/A`, then any other error / text / boolean.
ChartDataPoint plot_point(const Pt& cell, const PlotRules& rules) {
  ChartDataPoint point;
  bool blank = cell.kind == Pt::Kind::kBlank;
  if (cell.kind == Pt::Kind::kError && cell.is_na) {
    if (!rules.na_as_blank) {
      point.state = rules.connects ? ChartPointState::kConnect : ChartPointState::kGap;
      return point;
    }
    blank = true;
  }
  if (blank) {
    switch (rules.blanks) {
      case ChartBlanksAs::kGap:
        point.state = ChartPointState::kGap;
        break;
      case ChartBlanksAs::kSpan:
        point.state = rules.connects ? ChartPointState::kConnect : ChartPointState::kGap;
        break;
      default:
        break;
    }
    return point;
  }
  if (cell.kind == Pt::Kind::kNumber || cell.kind == Pt::Kind::kBool) {
    point.value = cell.number;
  }
  return point;
}

/// A multi-cell name joins its non-blank displayed texts with one space.
std::string joined_text(const Gathered& g) {
  std::string out;
  for (const Pt& cell : g.cells) {
    if (cell.display.empty()) {
      continue;
    }
    if (!out.empty()) {
      out += ' ';
    }
    out += cell.display;
  }
  return out;
}

void note_unresolved(ChartData& data, std::uint32_t idx, const ChartDataSource& src, const Gathered& g,
                     const char* role) {
  if (src.kind == ChartDataSource::Kind::kRef && !g.resolved) {
    data.unresolved.push_back(ChartUnresolved{idx, role});
  }
}

Expected<std::optional<std::string>, Error> title_text(const Workbook& wb, std::size_t host_sheet,
                                                       const ChartTitle& title, bool* unresolved) {
  if (title.kind == ChartTitleKind::kText) {
    return std::optional<std::string>(title.text);
  }
  if (title.kind != ChartTitleKind::kRef) {
    return std::optional<std::string>{};
  }
  ChartDataSource src;
  src.kind = ChartDataSource::Kind::kRef;
  src.formula = title.ref_formula;
  auto gathered = Reader(wb, host_sheet, false, false).gather(src, true, false);
  if (!gathered) {
    return gathered.error();
  }
  const Gathered& g = gathered.value();
  if (!g.resolved) {
    *unresolved = true;
    return std::optional<std::string>{};
  }
  return std::optional<std::string>(joined_text(g));
}

}  // namespace

Expected<ChartData, Error> resolve_chart_data(const Workbook& wb, std::size_t host_sheet, const ChartModel& model) {
  eval::ScopedEvalProfile profile_scope(wb.excel_profile());
  const bool plot_vis_only = model.plot_vis_only != Tri::kFalse;
  const Reader reader(wb, host_sheet, plot_vis_only, model.is_pivot_chart);
  // A series name keeps its text in a hidden row or column.
  const Reader name_reader(wb, host_sheet, false, model.is_pivot_chart);
  ChartData data;

  bool all_xy = !model.plots.empty();
  for (const ChartPlot& plot : model.plots) {
    all_xy = all_xy && is_xy(plot.type);
  }
  data.domain = all_xy ? ChartData::Domain::kXY : ChartData::Domain::kCategory;

  // Categories come from the first series with a category source, preferring shown series.
  const ChartSeries* cat_owner = nullptr;
  for (int pass = 0; pass < 2 && cat_owner == nullptr && !all_xy; ++pass) {
    for (const ChartPlot& plot : model.plots) {
      for (const ChartSeries& s : plot.series) {
        if (cat_owner == nullptr && s.cat.kind != ChartDataSource::Kind::kNone && (pass == 1 || !s.filtered)) {
          cat_owner = &s;
        }
      }
    }
  }

  std::size_t longest = 0;
  for (std::size_t pi = 0; pi < model.plots.size(); ++pi) {
    const ChartPlot& plot = model.plots[pi];
    PlotRules rules;
    rules.blanks = model.disp_blanks_as == ChartBlanksAs::kAbsent ? ChartBlanksAs::kZero : model.disp_blanks_as;
    rules.na_as_blank = model.disp_na_as_blank == Tri::kTrue;
    rules.connects = connects(plot.type);
    for (const ChartSeries& s : plot.series) {
      ChartDataSeries out;
      out.plot_index = pi;
      out.idx = s.idx;
      out.filtered = s.filtered;

      auto name_result = name_reader.gather(s.name, true, false);
      if (!name_result) {
        return name_result.error();
      }
      const Gathered& name = name_result.value();
      note_unresolved(data, s.idx, s.name, name, "name");
      out.name_ranges = name.ranges;
      if (s.name.kind != ChartDataSource::Kind::kNone) {
        out.name = joined_text(name);
      }

      if (s.cat.kind != ChartDataSource::Kind::kNone) {
        if (&s == cat_owner) {
          auto cat_result = reader.gather(s.cat, true, /*multi_level=*/true);
          if (!cat_result) {
            return cat_result.error();
          }
          const Gathered& cat = cat_result.value();
          note_unresolved(data, s.idx, s.cat, cat, "cat");
          out.cat_ranges = cat.ranges;
          ChartDataCategories& categories = data.categories;
          categories.format_code = cat.format_code;
          categories.levels = cat.levels;
          bool numeric = false;
          bool text = false;
          for (const Pt& c : cat.cells) {
            categories.labels.push_back(c.display);
            numeric = numeric || c.kind == Pt::Kind::kNumber;
            text = text || (c.kind != Pt::Kind::kNumber && c.kind != Pt::Kind::kBlank && !c.empty_text);
          }
          if (numeric && !text && cat.levels.empty()) {
            for (const Pt& c : cat.cells) {
              categories.values.push_back(c.kind == Pt::Kind::kNumber ? c.number
                                                                      : std::numeric_limits<double>::quiet_NaN());
            }
          }
        } else {
          auto cat_result = reader.gather(s.cat, false, false);
          if (!cat_result) {
            return cat_result.error();
          }
          const Gathered& cat = cat_result.value();
          note_unresolved(data, s.idx, s.cat, cat, "cat");
          out.cat_ranges = cat.ranges;
        }
      }
      if (is_xy(plot.type)) {
        const ChartDataSource& y_src = s.y.kind != ChartDataSource::Kind::kNone ? s.y : s.val;
        auto y_result = reader.gather(y_src, false, false);
        if (!y_result) {
          return y_result.error();
        }
        auto x_result = reader.gather(s.x, false, false);
        if (!x_result) {
          return x_result.error();
        }
        auto size_result = reader.gather(s.size, false, false);
        if (!size_result) {
          return size_result.error();
        }
        const Gathered& y = y_result.value();
        const Gathered& x = x_result.value();
        const Gathered& size = size_result.value();
        note_unresolved(data, s.idx, s.x, x, "x");
        note_unresolved(data, s.idx, y_src, y, s.y.kind != ChartDataSource::Kind::kNone ? "y" : "val");
        note_unresolved(data, s.idx, s.size, size, "size");
        if (y.had_cells && y.cells.empty()) {
          continue;  // Every value is hidden: Excel drops the series.
        }
        out.x_ranges = x.ranges;
        (s.y.kind != ChartDataSource::Kind::kNone ? out.y_ranges : out.val_ranges) = y.ranges;
        out.size_ranges = size.ranges;
        out.format_code = y.format_code;
        out.x_format_code = x.format_code;
        const bool text_x = std::any_of(x.cells.begin(), x.cells.end(),
                                        [](const Pt& c) { return c.kind == Pt::Kind::kText && !c.empty_text; });
        for (std::size_t i = 0; i < y.cells.size(); ++i) {
          const ChartDataPoint py = plot_point(y.cells[i], rules);
          ChartXYPoint p;
          p.y = py.value;
          p.state = py.state;
          if (s.x.kind == ChartDataSource::Kind::kNone || text_x || i >= x.cells.size()) {
            p.x = static_cast<double>(i + 1);
          } else {
            const Pt& c = x.cells[i];
            p.x = (c.kind == Pt::Kind::kNumber || c.kind == Pt::Kind::kBool) ? c.number : 0.0;
            // A blank X has no value, so the point is not drawn.
            if (c.kind == Pt::Kind::kBlank || c.empty_text) {
              p.state = ChartPointState::kGap;
            }
          }
          out.xy.push_back(p);
          if (plot.type == ChartPlotType::kBubble) {
            const bool has = i < size.cells.size() &&
                             (size.cells[i].kind == Pt::Kind::kNumber || size.cells[i].kind == Pt::Kind::kBool);
            out.sizes.push_back(has ? size.cells[i].number : 0.0);
          }
        }
        longest = std::max(longest, out.xy.size());
      } else {
        auto val_result = reader.gather(s.val, false, false);
        if (!val_result) {
          return val_result.error();
        }
        const Gathered& val = val_result.value();
        note_unresolved(data, s.idx, s.val, val, "val");
        if (val.had_cells && val.cells.empty()) {
          continue;  // Every value is hidden: Excel drops the series.
        }
        out.val_ranges = val.ranges;
        out.format_code = val.format_code;
        for (const Pt& c : val.cells) {
          out.points.push_back(plot_point(c, rules));
        }
        longest = std::max(longest, out.points.size());
      }

      data.series.push_back(std::move(out));
    }
  }

  if (data.domain == ChartData::Domain::kCategory) {
    longest = std::max(longest, data.categories.labels.size());
    if (cat_owner == nullptr) {
      data.categories.generated = true;
      for (std::size_t i = 1; i <= longest; ++i) {
        data.categories.labels.push_back(count_label(i));
      }
    }
  }
  data.point_count = longest;

  bool unresolved = false;
  auto title_result = title_text(wb, host_sheet, model.title, &unresolved);
  if (!title_result) {
    return title_result.error();
  }
  data.title_text = std::move(title_result.value());
  if (unresolved) {
    data.unresolved.push_back(ChartUnresolved{0, "title"});
  }
  for (const ChartAxis& axis : model.axes) {
    unresolved = false;
    auto axis_title_result = title_text(wb, host_sheet, axis.title, &unresolved);
    if (!axis_title_result) {
      return axis_title_result.error();
    }
    if (axis_title_result.value().has_value()) {
      data.axis_title_texts.push_back(ChartAxisTitleText{axis.id, std::move(*axis_title_result.value())});
    }
    if (unresolved) {
      data.unresolved.push_back(ChartUnresolved{0, "axis_title", axis.id});
    }
  }
  return data;
}

}  // namespace formulon
