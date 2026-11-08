#include "chart/chart_template.h"

#include <algorithm>
#include <array>
#include <cstdint>
#include <iterator>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "chart/chart_cache.h"
#include "chart/chart_ns.h"
#include "eval/a1_parse.h"
#include "excel_profile.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "parser/ast_format.h"
#include "parser/reference.h"
#include "sheet.h"
#include "style_resolve.h"
#include "utils/a1_column.h"
#include "utils/index_sort.h"
#include "utils/number_text.h"
#include "workbook.h"

namespace formulon {
namespace {

namespace a1p = eval::refs_internal;

Error invalid_ref(const char* what) {
  return make_error(FormulonErrorCode::kIoChartInvalidReference, what);
}

// ---------------------------------------------------------------------------
// Reference text
// ---------------------------------------------------------------------------

struct RectRef {
  std::string sheet;  ///< Empty when unqualified.
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  bool full_col = false;
  bool full_row = false;
};

bool parse_rect(std::string_view text, RectRef* out) {
  const a1p::A1Parse p = a1p::parse_a1_ref(text);
  if (!p.valid) {
    return false;
  }
  out->sheet.assign(p.sheet);
  const std::uint32_t row2 = p.is_range ? p.row2 : p.row;
  const std::uint32_t col2 = p.is_range ? p.col2 : p.col;
  out->top = std::min(p.row, row2);
  out->bottom = std::max(p.row, row2);
  out->left = std::min(p.col, col2);
  out->right = std::max(p.col, col2);
  out->full_col = p.is_full_col;
  out->full_row = p.is_full_row;
  return true;
}

std::string sheet_prefix(const std::string& name) {
  std::string out;
  parser::append_sheet_name(name, /*force_quote=*/false, out);
  out += '!';
  return out;
}

void append_abs_cell(std::string& out, std::uint32_t row, std::uint32_t col) {
  out += '$';
  a1::append_column_letters(out, col);
  out += '$';
  out += std::to_string(row + 1U);
}

std::string rect_ref(const std::string& prefix, std::uint32_t r0, std::uint32_t c0, std::uint32_t r1,
                     std::uint32_t c1) {
  std::string out = prefix;
  append_abs_cell(out, r0, c0);
  if (r0 != r1 || c0 != c1) {
    out += ':';
    append_abs_cell(out, r1, c1);
  }
  return out;
}

/// Resolves the sheet of `rect` (`host_sheet` when unqualified); size_t(-1) if none.
std::size_t rect_sheet(const Workbook& wb, std::size_t host_sheet, const RectRef& rect) {
  const std::size_t index = rect.sheet.empty() ? host_sheet : wb.sheet_index_by_name(rect.sheet);
  return index < wb.sheet_count() ? index : static_cast<std::size_t>(-1);
}

// ---------------------------------------------------------------------------
// Cell classification
// ---------------------------------------------------------------------------

enum class Cls : std::uint8_t { kBlank, kText, kDate, kNumber, kOther };

bool is_label(Cls c) {
  return c == Cls::kText || c == Cls::kDate;
}

bool ieq_prefix(std::string_view s, std::size_t pos, std::string_view word) {
  if (s.size() - pos < word.size()) {
    return false;
  }
  for (std::size_t i = 0; i < word.size(); ++i) {
    if ((s[pos + i] | 0x20) != word[i]) {
      return false;
    }
  }
  return true;
}

/// Source rectangle on one sheet, read through a class lookup.
struct Grid {
  const Workbook* wb = nullptr;
  const Sheet* sheet = nullptr;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t nr = 0;
  std::uint32_t nc = 0;
  mutable std::unordered_map<std::uint32_t, bool> date_xf;

  Cls cls(std::uint32_t row, std::uint32_t col) const {
    const Value v = sheet->resolve_cell_value(row, col);
    if (v.is_blank()) {
      return Cls::kBlank;
    }
    if (v.is_text()) {
      return Cls::kText;
    }
    if (!v.is_number()) {
      return Cls::kOther;
    }
    const std::uint32_t xf = select_effective_xf(*sheet, row, col).xf_index;
    auto it = date_xf.find(xf);
    if (it == date_xf.end()) {
      it = date_xf.emplace(xf, chart_is_date_format(effective_style(*wb, *sheet, row, col).num_fmt_code)).first;
    }
    return it->second ? Cls::kDate : Cls::kNumber;
  }

  Cls at(std::uint32_t i, std::uint32_t j) const { return cls(top + i, left + j); }
};

/// Header rows / columns Excel takes from the rectangle's contents.
void detect_headers(const Grid& g, std::uint32_t* hr, std::uint32_t* hc) {
  *hr = 0;
  *hc = 0;
  if (g.at(0, 0) == Cls::kBlank) {
    while (*hr < g.nr && g.at(*hr, 0) == Cls::kBlank) {
      ++*hr;
    }
    while (*hc < g.nc && g.at(0, *hc) == Cls::kBlank) {
      ++*hc;
    }
    return;
  }
  for (std::uint32_t i = g.nr - 1U; i-- > 0U;) {
    if (is_label(g.at(i, g.nc - 1U))) {
      *hr = i + 1U;
      break;
    }
  }
  for (std::uint32_t j = g.nc - 1U; j-- > 0U;) {
    if (is_label(g.at(g.nr - 1U, j))) {
      *hc = j + 1U;
      break;
    }
  }
}

bool header_cells_all_dates(const Grid& g, std::uint32_t hr, std::uint32_t hc) {
  for (std::uint32_t i = 0; i < hr; ++i) {
    for (std::uint32_t j = hc; j < g.nc; ++j) {
      if (g.at(i, j) != Cls::kDate) {
        return false;
      }
    }
  }
  return true;
}

// ---------------------------------------------------------------------------
// Canonical reference keys (derivation)
// ---------------------------------------------------------------------------

/// Sheet-independent identity of a reference formula; empty text keys to
/// empty, anything unparseable to a key no split produces.
std::string canon_ref(const Workbook& wb, std::size_t host_sheet, std::string_view formula) {
  if (formula.empty()) {
    return {};
  }
  RectRef r;
  if (!parse_rect(formula, &r) || r.full_col || r.full_row) {
    return "?";
  }
  const std::size_t si = rect_sheet(wb, host_sheet, r);
  if (si == static_cast<std::size_t>(-1)) {
    return "?";
  }
  return std::to_string(si) + ':' + std::to_string(r.top) + ',' + std::to_string(r.left) + ',' +
         std::to_string(r.bottom) + ',' + std::to_string(r.right);
}

// ---------------------------------------------------------------------------
// XML fragments
// ---------------------------------------------------------------------------

constexpr const char* kRoot =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
    "xmlns:c16r2=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">";

constexpr const char* kStyleAlt =
    "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
    "<mc:Choice Requires=\"c14\" xmlns:c14=\"http://schemas.microsoft.com/office/drawing/2007/8/2/chart\">"
    "<c14:style val=\"102\"/></mc:Choice><mc:Fallback><c:style val=\"2\"/></mc:Fallback></mc:AlternateContent>";

constexpr const char* kNoFillSpPr = "<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>";

constexpr std::uint32_t kAxisId1 = 507100001U;
constexpr std::uint32_t kAxisId2 = 507100002U;

void append_def_rpr(std::string& x, const char* sz, bool title) {
  x += "<a:defRPr sz=\"";
  x += sz;
  x += "\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\"";
  x += title ? " spc=\"0\"" : "";
  x += " baseline=\"0\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/>"
       "</a:schemeClr></a:solidFill><a:latin typeface=\"+mn-lt\"/><a:ea typeface=\"+mn-ea\"/>"
       "<a:cs typeface=\"+mn-cs\"/></a:defRPr>";
}

void append_body_pr(std::string& x, const char* rot) {
  x += "<a:bodyPr rot=\"";
  x += rot;
  x += "\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\" wrap=\"square\" anchor=\"ctr\" "
       "anchorCtr=\"1\"/><a:lstStyle/>";
}

void append_tx_pr(std::string& x, const char* rot, const char* sz, bool title, const std::string& lang) {
  x += "<c:txPr>";
  append_body_pr(x, rot);
  x += "<a:p><a:pPr>";
  append_def_rpr(x, sz, title);
  x += "</a:pPr><a:endParaRPr lang=\"" + lang + "\"/></a:p></c:txPr>";
}

// A chart title is 14 pt and drops `c:txPr` once it has text; an axis title is 10 pt and keeps it.
void append_title(std::string& x, const std::optional<std::string>& text, const std::string& lang, bool axis,
                  const char* rot) {
  const char* sz = axis ? "1000" : "1400";
  x += "<c:title>";
  if (text.has_value()) {
    x += "<c:tx><c:rich>";
    append_body_pr(x, rot);
    x += "<a:p><a:pPr>";
    append_def_rpr(x, sz, !axis);
    x += "</a:pPr><a:r><a:rPr lang=\"" + lang + "\"/><a:t>";
    io::AppendXmlRawTextEscaped(x, *text);
    x += "</a:t></a:r></a:p></c:rich></c:tx>";
  }
  x += "<c:overlay val=\"0\"/>";
  x += kNoFillSpPr;
  if (axis || !text.has_value()) {
    append_tx_pr(x, rot, sz, !axis, lang);
  }
  x += "</c:title>";
}

std::string chart_lang(const Workbook& wb) {
  std::string lang = excel_profile_id(wb.excel_profile());
  lang.erase(0, lang.rfind('-') + 1U);
  std::replace(lang.begin(), lang.end(), '_', '-');
  return lang;
}

void append_source(std::string& x, const char* tag, const std::string& f, ChartRefKind kind) {
  const char* ref = chart_ref_tag(kind);
  const char* cache = chart_cache_tag(kind);
  x += std::string("<c:") + tag + "><c:" + ref + "><c:f>";
  io::AppendXmlRawTextEscaped(x, f);
  x += std::string("</c:f><c:") + cache + ">";
  x += kind == ChartRefKind::kNum ? "<c:formatCode>General</c:formatCode>" : "";
  x += std::string("<c:ptCount val=\"0\"/></c:") + cache + "></c:" + ref + "></c:" + tag + ">";
}

void append_accent(std::string& x, const pugi::xml_node& colors, std::size_t idx) {
  x += chart_cycle_color(colors, idx);
}

// A category axis with a `time_unit` is a date axis.
void append_axis(std::string& x, bool category, std::uint32_t id, std::uint32_t cross, char pos, const char* tick,
                 int line_lum, std::string_view num_fmt, const char* cross_between, bool grid, const std::string& lang,
                 const char* time_unit) {
  const char* tag = !category ? "valAx" : time_unit != nullptr ? "dateAx" : "catAx";
  x += std::string("<c:") + tag + "><c:axId val=\"" + std::to_string(id) +
       "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"" + pos + "\"/>";
  if (grid) {
    x += gridlines_xml();
  }
  x += "<c:numFmt formatCode=\"";
  io::AppendXmlAttrEscaped(x, num_fmt);
  x += std::string("\" sourceLinked=\"1\"/><c:majorTickMark val=\"") + tick +
       "\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:spPr><a:noFill/>";
  if (line_lum != 0) {
    append_grey_ln(x, line_lum);
  } else {
    x += "<a:ln><a:noFill/></a:ln>";
  }
  x += "<a:effectLst/></c:spPr>";
  append_tx_pr(x, "-60000000", "900", false, lang);
  x += "<c:crossAx val=\"" + std::to_string(cross) + "\"/><c:crosses val=\"autoZero\"/>";
  if (time_unit != nullptr) {
    x += std::string("<c:auto val=\"1\"/><c:lblOffset val=\"100\"/><c:baseTimeUnit val=\"") + time_unit + "\"/>";
  } else if (category) {
    x += "<c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/><c:noMultiLvlLbl val=\"0\"/>";
  } else {
    x += std::string("<c:crossBetween val=\"") + cross_between + "\"/>";
  }
  x += std::string("</c:") + tag + ">";
}

const char* legend_pos(ChartLegendPosition p) {
  const int i = static_cast<int>(p);
  return kChartLegendTokens[i > 0 && i < static_cast<int>(std::size(kChartLegendTokens)) ? i : 1];
}

/// Cell class of the first cell of `formula` (a qualified rectangle).
Cls first_cell_class(const Workbook& wb, const std::string& formula) {
  RectRef r;
  if (!parse_rect(formula, &r)) {
    return Cls::kBlank;
  }
  const std::size_t si = rect_sheet(wb, 0, r);
  if (si == static_cast<std::size_t>(-1)) {
    return Cls::kBlank;
  }
  Grid g;
  g.wb = &wb;
  g.sheet = &wb.sheet(si);
  return g.cls(r.top, r.left);
}

/// Whether category range `cat` holds more than one level for series laid out in `direction`.
bool multi_level(const std::string& cat, ChartSeriesDirection direction) {
  RectRef r;
  return parse_rect(cat, &r) && (direction == ChartSeriesDirection::kRows ? r.bottom > r.top : r.right > r.left);
}

/// What an axis linked to the cells of `formula` takes from them.
struct LinkedCells {
  std::string format = "General";  ///< The first cell's number format; General when any cell holds text.
  std::vector<double> dates;       ///< The serials when every cell that is not blank holds a date; else empty.
};

LinkedCells linked_cells(const Workbook& wb, const std::string& formula) {
  LinkedCells out;
  RectRef r;
  if (!parse_rect(formula, &r) || r.full_col || r.full_row) {
    return out;
  }
  const std::size_t si = rect_sheet(wb, 0, r);
  if (si == static_cast<std::size_t>(-1)) {
    return out;
  }
  Grid g;
  g.wb = &wb;
  g.sheet = &wb.sheet(si);
  bool text = false;
  bool dated = true;
  for (std::uint32_t row = r.top; row <= r.bottom; ++row) {
    for (std::uint32_t col = r.left; col <= r.right; ++col) {
      const Cls c = g.cls(row, col);
      text = text || c == Cls::kText;
      dated = dated && (c == Cls::kDate || c == Cls::kBlank);
      if (dated && c == Cls::kDate) {
        out.dates.push_back(g.sheet->resolve_cell_value(row, col).as_number());
      }
    }
  }
  if (!dated) {
    out.dates.clear();
  }
  const std::string code = text ? std::string() : chart_cell_format_code(wb, si, r.top, r.left);
  if (!code.empty()) {
    out.format = code;
  }
  return out;
}

std::uint64_t point_count(const std::string& formula) {
  RectRef r;
  if (!parse_rect(formula, &r) || r.full_col || r.full_row) {
    return 0;
  }
  return static_cast<std::uint64_t>(r.bottom - r.top + 1U) * (r.right - r.left + 1U);
}

}  // namespace

bool chart_is_date_format(std::string_view code) {
  for (std::size_t i = 0; i < code.size(); ++i) {
    const char c = code[i];
    if (c == '"') {
      i = code.find('"', i + 1);
      if (i == std::string_view::npos) {
        return false;
      }
    } else if (c == '\\' || c == '_' || c == '*') {
      ++i;
    } else if (c == '[') {
      const std::size_t end = code.find(']', i);
      if (end == std::string_view::npos) {
        return false;
      }
      const char k = static_cast<char>(code[i + 1] | 0x20);
      bool elapsed = (k == 'h' || k == 'm' || k == 's');
      for (std::size_t j = i + 1; elapsed && j < end; ++j) {
        elapsed = (code[j] | 0x20) == k;
      }
      if (elapsed) {
        return true;
      }
      i = end;
    } else if (ieq_prefix(code, i, "general")) {
      i += 6;
    } else {
      const char k = static_cast<char>(c | 0x20);
      if (k == 'y' || k == 'm' || k == 'd' || k == 'h' || k == 's') {
        return true;
      }
    }
  }
  return false;
}

const char* chart_base_time_unit(const std::vector<double>& dates) {
  std::vector<std::uint32_t> order;
  sorted_index_order_ascending(order, dates);
  double gap = 365.0;
  for (std::size_t i = 1; i < order.size(); ++i) {
    gap = std::min(gap, dates[order[i]] - dates[order[i - 1U]]);
  }
  return order.size() < 2U || gap < 28.0 ? "days" : gap < 365.0 ? "months" : "years";
}

void append_grey_ln(std::string& x, int lum) {
  x += "<a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"";
  x += std::to_string(lum * 1000);
  x += "\"/><a:lumOff val=\"";
  x += std::to_string(100000 - lum * 1000);
  x += "\"/></a:schemeClr></a:solidFill><a:round/></a:ln>";
}

std::string gridlines_xml(bool minor) {
  const std::string tag = minor ? "c:minorGridlines" : "c:majorGridlines";
  std::string x = "<" + tag + "><c:spPr>";
  append_grey_ln(x, minor ? 5 : 15);
  return x + "<a:effectLst/></c:spPr></" + tag + ">";
}

Expected<ChartSplit, Error> split_chart_source(const Workbook& wb, std::size_t host_sheet, std::string_view source,
                                               ChartInsertKind kind, ChartSeriesDirection direction,
                                               ChartHeaderMode header_row, ChartHeaderMode header_column) {
  RectRef rr;
  if (!parse_rect(source, &rr)) {
    return invalid_ref("chart source is not a rectangle reference");
  }
  const std::size_t si = rect_sheet(wb, host_sheet, rr);
  if (si == static_cast<std::size_t>(-1)) {
    return invalid_ref("chart source names no sheet");
  }
  const Sheet& sheet = wb.sheet(si);
  if (rr.full_col || rr.full_row) {
    const auto extent = sheet.populated_extent(rr.top, rr.left, rr.bottom, rr.right);
    if (!extent) {
      return invalid_ref("chart source is empty");
    }
    rr.bottom = rr.full_col ? extent->last_row : rr.bottom;
    rr.right = rr.full_row ? extent->last_col : rr.right;
  }

  Grid g;
  g.wb = &wb;
  g.sheet = &sheet;
  g.top = rr.top;
  g.left = rr.left;
  g.nr = rr.bottom - rr.top + 1U;
  g.nc = rr.right - rr.left + 1U;

  std::uint32_t hr = 0;
  std::uint32_t hc = 0;
  detect_headers(g, &hr, &hc);
  if (header_row != ChartHeaderMode::kAuto) {
    hr = header_row == ChartHeaderMode::kYes ? 1U : 0U;
  }
  if (header_column != ChartHeaderMode::kAuto) {
    hc = header_column == ChartHeaderMode::kYes ? 1U : 0U;
  }
  if (hr >= g.nr || hc >= g.nc) {
    return invalid_ref("chart source has no data cells");
  }

  ChartSplit out;
  out.header_row = hr > 0U;
  out.header_column = hc > 0U;
  out.direction = direction;
  if (direction == ChartSeriesDirection::kAuto) {
    const bool tall = (g.nr - hr) > (g.nc - hc);
    const bool dates = hr > 0U && hc > 0U && header_cells_all_dates(g, hr, hc);
    out.direction = tall && !dates ? ChartSeriesDirection::kColumns : ChartSeriesDirection::kRows;
  }

  // Oriented coordinates: p runs along a series, q across series.
  const bool rows = out.direction == ChartSeriesDirection::kRows;
  const std::uint32_t hp = rows ? hc : hr;
  const std::uint32_t hq = rows ? hr : hc;
  const std::uint32_t pn = rows ? g.nc : g.nr;
  const std::uint32_t qn = rows ? g.nr : g.nc;
  const std::string prefix = sheet_prefix(sheet.name());
  const auto ref = [&](std::uint32_t p0, std::uint32_t q0, std::uint32_t p1, std::uint32_t q1) {
    return rows ? rect_ref(prefix, g.top + q0, g.left + p0, g.top + q1, g.left + p1)
                : rect_ref(prefix, g.top + p0, g.left + q0, g.top + p1, g.left + q1);
  };
  const auto at = [&](std::uint32_t p, std::uint32_t q) { return rows ? g.at(q, p) : g.at(p, q); };

  std::uint32_t q_first = hq;
  std::string x_ref;
  const bool scatter = kind == ChartInsertKind::kScatter;
  if (scatter) {
    if (hq > 0U) {
      x_ref = ref(hp, 0, pn - 1U, hq - 1U);
      if (at(0, 0) != Cls::kBlank && !is_label(at(hp, 0)) && qn - hq >= 2U) {
        x_ref = ref(hp, 0, pn - 1U, hq);
        q_first = hq + 1U;
      }
    } else if (qn >= 2U && at(hp, 0) != Cls::kText) {
      x_ref = ref(hp, 0, pn - 1U, 0);
      q_first = 1U;
    }
  }
  const std::string cat_ref = hq > 0U ? ref(hp, 0, pn - 1U, hq - 1U) : std::string();

  for (std::uint32_t q = q_first; q < qn; ++q) {
    ChartSeriesRefs s;
    if (hp > 0U) {
      s.name = ref(0, q, hp - 1U, q);
    }
    const std::string vals = ref(hp, q, pn - 1U, q);
    if (scatter) {
      s.x = x_ref;
      s.y = vals;
    } else {
      s.cat = cat_ref;
      s.val = vals;
    }
    out.series.push_back(std::move(s));
  }
  return out;
}

Expected<std::string, Error> build_chart_xml(const Workbook& wb, const ChartInsertOptions& options,
                                             const ChartSplit& split) {
  pugi::xml_document color_style;
  color_style.load_string(kDefaultChartColorStyle);
  return build_chart_xml(wb, options, split, color_style.document_element());
}

Expected<std::string, Error> build_chart_xml(const Workbook& wb, const ChartInsertOptions& options,
                                             const ChartSplit& split, const pugi::xml_node& colors) {
  if (split.series.empty()) {
    return invalid_ref("chart has no series");
  }
  const ChartInsertKind kind = options.kind;
  const bool is_bar = kind == ChartInsertKind::kColumn || kind == ChartInsertKind::kBar;
  const bool is_line = kind == ChartInsertKind::kLine;
  const bool is_area = kind == ChartInsertKind::kArea;
  // A doughnut shares the pie form: per-point colours, leader lines, no axes.
  const bool is_ring = kind == ChartInsertKind::kDoughnut;
  const bool is_pie = kind == ChartInsertKind::kPie || is_ring;
  const bool is_scatter = kind == ChartInsertKind::kScatter;
  const bool grouped = is_bar || is_line || is_area;
  const ChartGrouping grouping = grouped ? options.grouping : ChartGrouping::kClustered;
  const bool stacked = grouping == ChartGrouping::kStacked || grouping == ChartGrouping::kPercentStacked;
  const bool percent = grouping == ChartGrouping::kPercentStacked;

  const std::string lang = chart_lang(wb);

  std::string x = kRoot;
  x += std::string("<c:date1904 val=\"") + (wb.date1904() ? "1" : "0") + "\"/><c:lang val=\"" + lang +
       "\"/><c:roundedCorners val=\"0\"/>" + kStyleAlt + "<c:chart>";
  const bool no_title = options.title.has_value() && options.title->empty();
  if (!no_title) {
    append_title(x, options.title, lang, false, "0");
  }
  x += no_title ? "<c:autoTitleDeleted val=\"1\"/>" : "<c:autoTitleDeleted val=\"0\"/>";
  x += "<c:plotArea><c:layout/>";

  const char* plot = is_bar    ? "barChart"
                     : is_line ? "lineChart"
                     : is_area ? "areaChart"
                     : is_ring ? "doughnutChart"
                     : is_pie  ? "pieChart"
                               : "scatterChart";
  x += std::string("<c:") + plot + ">";
  if (is_bar) {
    x += kind == ChartInsertKind::kBar ? "<c:barDir val=\"bar\"/>" : "<c:barDir val=\"col\"/>";
  }
  if (grouped) {
    x += std::string("<c:grouping val=\"") +
         (percent   ? "percentStacked"
          : stacked ? "stacked"
          : is_bar  ? "clustered"
                    : "standard") +
         "\"/>";
  }
  if (is_scatter) {
    x += "<c:scatterStyle val=\"lineMarker\"/>";
  }
  x += is_pie ? "<c:varyColors val=\"1\"/>" : "<c:varyColors val=\"0\"/>";

  for (std::size_t i = 0; i < split.series.size(); ++i) {
    const ChartSeriesRefs& s = split.series[i];
    x += "<c:ser><c:idx val=\"" + std::to_string(i) + "\"/><c:order val=\"" + std::to_string(i) + "\"/>";
    if (!s.name.empty()) {
      append_source(x, "tx", s.name, ChartRefKind::kStr);
    }
    if (is_bar || is_area) {
      x += "<c:spPr><a:solidFill>";
      append_accent(x, colors, i);
      x += std::string("</a:solidFill><a:ln") + (is_area && i > 0U ? " w=\"25400\"" : "") +
           "><a:noFill/></a:ln><a:effectLst/></c:spPr>";
      x += is_bar ? "<c:invertIfNegative val=\"0\"/>" : "";
    } else if (is_line || is_scatter) {
      x += std::string("<c:spPr><a:ln w=\"") + (is_line ? "28575" : "25400") + "\" cap=\"rnd\">";
      if (is_line) {
        x += "<a:solidFill>";
        append_accent(x, colors, i);
        x += "</a:solidFill>";
      } else {
        x += "<a:noFill/>";
      }
      x += "<a:round/></a:ln><a:effectLst/></c:spPr>";
      if (is_line) {
        x += "<c:marker><c:symbol val=\"none\"/></c:marker>";
      } else {
        x += "<c:marker><c:symbol val=\"circle\"/><c:size val=\"5\"/><c:spPr><a:solidFill>";
        append_accent(x, colors, i);
        x += "</a:solidFill><a:ln w=\"9525\"><a:solidFill>";
        append_accent(x, colors, i);
        x += "</a:solidFill></a:ln><a:effectLst/></c:spPr></c:marker>";
      }
    }
    if (is_pie) {
      const std::uint64_t points = point_count(s.val);
      if (points > Sheet::kMaxRows) {
        return make_error(FormulonErrorCode::kSecResourceLimit, "chart: point count exceeds Excel row limit",
                          "points=" + std::to_string(points) + " ceiling=" + std::to_string(Sheet::kMaxRows));
      }
      for (std::uint32_t p = 0; p < points; ++p) {
        x += "<c:dPt><c:idx val=\"" + std::to_string(p) + "\"/><c:bubble3D val=\"0\"/><c:spPr><a:solidFill>";
        append_accent(x, colors, p);
        x += "</a:solidFill><a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:ln>"
             "<a:effectLst/></c:spPr></c:dPt>";
      }
    }
    if (is_scatter) {
      append_source(x, "xVal", s.x, first_cell_class(wb, s.x) == Cls::kText ? ChartRefKind::kStr : ChartRefKind::kNum);
      append_source(x, "yVal", s.y, ChartRefKind::kNum);
    } else {
      if (!s.cat.empty()) {
        const Cls first = first_cell_class(wb, s.cat);
        append_source(x, "cat", s.cat,
                      multi_level(s.cat, split.direction)
                          ? ChartRefKind::kMulti
                          : (first == Cls::kNumber || first == Cls::kDate ? ChartRefKind::kNum : ChartRefKind::kStr));
      }
      append_source(x, "val", s.val, ChartRefKind::kNum);
    }
    if (is_line || is_scatter) {
      x += "<c:smooth val=\"0\"/>";
    }
    char uid_digits[8];
    format_hex(uid_digits, sizeof(uid_digits), i & 0xFFFFU, 4, true);
    const std::string uid = std::string("{00000000-") + uid_digits + "-0000-0000-000000000000}";
    x += "<c:extLst><c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\" "
         "xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\"><c16:uniqueId val=\"" +
         uid + "\"/></c:ext></c:extLst></c:ser>";
  }

  x += std::string("<c:dLbls>") + kChartShowFlags + (is_pie ? "<c:showLeaderLines val=\"1\"/>" : "") + "</c:dLbls>";
  if (is_bar) {
    x += kind == ChartInsertKind::kBar ? "<c:gapWidth val=\"182\"/>" : "<c:gapWidth val=\"219\"/>";
    x += stacked ? "<c:overlap val=\"100\"/>" : "";
  }
  x += is_line ? "<c:smooth val=\"0\"/>" : "";
  if (is_pie) {
    x += is_ring ? "<c:firstSliceAng val=\"0\"/><c:holeSize val=\"50\"/>" : "<c:firstSliceAng val=\"0\"/>";
  } else {
    x += "<c:axId val=\"" + std::to_string(kAxisId1) + "\"/><c:axId val=\"" + std::to_string(kAxisId2) + "\"/>";
  }
  x += std::string("</c:") + plot + ">";

  // Each axis links to the first cells of the first series; a single-level category range of dates makes a date axis.
  const ChartSeriesRefs& first = split.series[0];
  const LinkedCells values = linked_cells(wb, is_scatter ? first.y : first.val);
  if (is_scatter) {
    append_axis(x, false, kAxisId1, kAxisId2, 'b', "none", 25, linked_cells(wb, first.x).format, "midCat", true, lang,
                nullptr);
    append_axis(x, false, kAxisId2, kAxisId1, 'l', "none", 25, values.format, "midCat", true, lang, nullptr);
  } else if (!is_pie) {
    const bool horizontal = kind == ChartInsertKind::kBar;
    LinkedCells cats;
    if (!first.cat.empty() && !multi_level(first.cat, split.direction)) {
      cats = linked_cells(wb, first.cat);
    }
    append_axis(x, true, kAxisId1, kAxisId2, horizontal ? 'l' : 'b', is_area ? "out" : "none", 15, cats.format, "",
                false, lang, cats.dates.empty() ? nullptr : chart_base_time_unit(cats.dates));
    append_axis(x, false, kAxisId2, kAxisId1, horizontal ? 'b' : 'l', "none", 0,
                percent ? std::string_view("0%") : std::string_view(values.format), is_area ? "midCat" : "between",
                true, lang, nullptr);
  }
  x += kNoFillSpPr;
  x += "</c:plotArea>";

  if (options.legend != ChartLegendPosition::kNone) {
    x += std::string("<c:legend><c:legendPos val=\"") + legend_pos(options.legend) + "\"/><c:overlay val=\"0\"/>" +
         kNoFillSpPr;
    append_tx_pr(x, "0", "900", false, lang);
    x += "</c:legend>";
  }
  x += "<c:plotVisOnly val=\"1\"/><c:dispBlanksAs val=\"gap\"/><c:showDLblsOverMax val=\"0\"/><c:extLst>";
  x += chart_ns::data_display_ext_xml();
  x += "</c:extLst></c:chart><c:spPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill>";
  append_grey_ln(x, 15);
  x += "<a:effectLst/></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang=\"" +
       lang +
       "\"/></a:p></c:txPr><c:printSettings><c:headerFooter/><c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" "
       "t=\"0.75\" header=\"0.3\" footer=\"0.3\"/><c:pageSetup/></c:printSettings></c:chartSpace>";
  return x;
}

std::optional<ChartSource> derive_chart_source(const Workbook& wb, std::size_t host_sheet, const ChartModel& model) {
  std::vector<const ChartSeries*> series;
  bool scatter = false;
  for (const ChartPlot& plot : model.plots) {
    const bool is_scatter = plot.type == ChartPlotType::kScatter;
    const bool is_category = plot.type <= ChartPlotType::kOfPie;
    if (!is_scatter && !is_category) {
      return std::nullopt;
    }
    if (!series.empty() && is_scatter != scatter) {
      return std::nullopt;
    }
    scatter = is_scatter;
    for (const ChartSeries& s : plot.series) {
      series.push_back(&s);
    }
  }
  if (series.empty()) {
    return std::nullopt;
  }
  // Filtered series sit after the visible ones in the file; the split follows c:order.
  sort_by_index(series, [](const ChartSeries* a, const ChartSeries* b) { return a->order < b->order; });

  // A filtered ref keeps its full range in c15:fullRef; the c:f union lists only visible cells.
  const auto ref_of = [](const ChartDataSource& d) -> const std::string& {
    return d.full_formula.empty() ? d.formula : d.full_formula;
  };
  // Each series reduces to the keys of its name / category / value references.
  const auto key = [&](const ChartDataSource& d) {
    return d.kind == ChartDataSource::Kind::kLiteral ? std::string("?") : canon_ref(wb, host_sheet, ref_of(d));
  };
  std::vector<std::array<std::string, 3>> want;
  bool have = false;
  std::size_t sheet = 0;
  RectRef box;
  for (const ChartSeries* s : series) {
    const ChartDataSource* parts[3] = {&s->name, scatter ? &s->x : &s->cat, scatter ? &s->y : &s->val};
    std::array<std::string, 3> keys;
    for (std::size_t i = 0; i < 3; ++i) {
      keys[i] = key(*parts[i]);
      if (keys[i] == "?") {
        return std::nullopt;
      }
      if (ref_of(*parts[i]).empty() || parts[i]->kind != ChartDataSource::Kind::kRef) {
        continue;
      }
      RectRef r;
      parse_rect(ref_of(*parts[i]), &r);
      const std::size_t si = rect_sheet(wb, host_sheet, r);
      if (have && si != sheet) {
        return std::nullopt;
      }
      if (!have) {
        box = r;
        sheet = si;
        have = true;
      }
      box.top = std::min(box.top, r.top);
      box.left = std::min(box.left, r.left);
      box.bottom = std::max(box.bottom, r.bottom);
      box.right = std::max(box.right, r.right);
    }
    want.push_back(keys);
  }
  if (!have) {
    return std::nullopt;
  }

  const std::string prefix = sheet_prefix(wb.sheet(sheet).name());
  const std::string range = rect_ref(prefix, box.top, box.left, box.bottom, box.right);
  static constexpr ChartHeaderMode kModes[][2] = {
      {ChartHeaderMode::kAuto, ChartHeaderMode::kAuto}, {ChartHeaderMode::kAuto, ChartHeaderMode::kNo},
      {ChartHeaderMode::kNo, ChartHeaderMode::kAuto},   {ChartHeaderMode::kNo, ChartHeaderMode::kNo},
      {ChartHeaderMode::kAuto, ChartHeaderMode::kYes},  {ChartHeaderMode::kYes, ChartHeaderMode::kAuto},
      {ChartHeaderMode::kYes, ChartHeaderMode::kYes},   {ChartHeaderMode::kYes, ChartHeaderMode::kNo},
      {ChartHeaderMode::kNo, ChartHeaderMode::kYes},
  };
  const ChartInsertKind kind = scatter ? ChartInsertKind::kScatter : ChartInsertKind::kColumn;
  for (const ChartSeriesDirection dir : {ChartSeriesDirection::kColumns, ChartSeriesDirection::kRows}) {
    for (const auto& mode : kModes) {
      const auto split = split_chart_source(wb, host_sheet, range, kind, dir, mode[0], mode[1]);
      if (!split || split.value().series.size() != want.size()) {
        continue;
      }
      bool same = true;
      for (std::size_t i = 0; same && i < want.size(); ++i) {
        const ChartSeriesRefs& s = split.value().series[i];
        same = canon_ref(wb, host_sheet, s.name) == want[i][0] &&
               canon_ref(wb, host_sheet, scatter ? s.x : s.cat) == want[i][1] &&
               canon_ref(wb, host_sheet, scatter ? s.y : s.val) == want[i][2];
      }
      if (same) {
        return ChartSource{range, split.value().direction, split.value().header_row, split.value().header_column};
      }
    }
  }
  return std::nullopt;
}

const char* const kDefaultChartColorStyle =
    "<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" "
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" meth=\"cycle\" id=\"10\">"
    "<a:schemeClr val=\"accent1\"/><a:schemeClr val=\"accent2\"/><a:schemeClr val=\"accent3\"/>"
    "<a:schemeClr val=\"accent4\"/><a:schemeClr val=\"accent5\"/><a:schemeClr val=\"accent6\"/><cs:variation/>"
    "<cs:variation><a:lumMod val=\"60000\"/></cs:variation><cs:variation><a:lumMod val=\"80000\"/>"
    "<a:lumOff val=\"20000\"/></cs:variation><cs:variation><a:lumMod val=\"80000\"/></cs:variation><cs:variation>"
    "<a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/></cs:variation><cs:variation><a:lumMod val=\"50000\"/>"
    "</cs:variation><cs:variation><a:lumMod val=\"70000\"/><a:lumOff val=\"30000\"/></cs:variation><cs:variation>"
    "<a:lumMod val=\"70000\"/></cs:variation><cs:variation><a:lumMod val=\"50000\"/><a:lumOff val=\"50000\"/>"
    "</cs:variation></cs:colorStyle>";

std::string build_axis_title_xml(const Workbook& wb, const std::optional<std::string>& text, bool vertical) {
  std::string x;
  append_title(x, text, chart_lang(wb), true, vertical ? "-5400000" : "0");
  return x;
}

std::string chart_cycle_color(const pugi::xml_node& style, std::size_t idx) {
  std::vector<pugi::xml_node> bases;
  std::vector<pugi::xml_node> variations;
  // DrawingML colour choices; anything else in the part (an extension list) is neither base nor variation.
  static constexpr std::string_view kColours[] = {"scrgbClr", "srgbClr", "hslClr", "sysClr", "schemeClr", "prstClr"};
  for (pugi::xml_node n = style.first_child(); n; n = n.next_sibling()) {
    if (n.type() != pugi::node_element) {
      continue;
    }
    const std::string_view local = io::local_name(n);
    if (local == "variation") {
      variations.push_back(n);
    } else if (std::find(std::begin(kColours), std::end(kColours), local) != std::end(kColours)) {
      bases.push_back(n);
    }
  }
  if (bases.empty()) {
    return {};
  }
  const pugi::xml_node base = bases[idx % bases.size()];
  std::string inner;
  for (pugi::xml_node n = base.first_child(); n; n = n.next_sibling()) {
    io::append_raw_xml(inner, n);
  }
  if (!variations.empty()) {
    for (pugi::xml_node n = variations[(idx / bases.size()) % variations.size()].first_child(); n;
         n = n.next_sibling()) {
      io::append_raw_xml(inner, n);
    }
  }
  std::string x = std::string("<") + base.name();
  for (const pugi::xml_attribute& a : base.attributes()) {
    x += std::string(" ") + a.name() + "=\"";
    io::AppendXmlAttrEscaped(x, a.value());
    x += "\"";
  }
  return inner.empty() ? x + "/>" : x + ">" + inner + "</" + base.name() + ">";
}

std::string chart_rect_ref(const std::string& sheet_name, std::uint32_t top, std::uint32_t left, std::uint32_t bottom,
                           std::uint32_t right) {
  return rect_ref(sheet_prefix(sheet_name), top, left, bottom, right);
}

}  // namespace formulon
