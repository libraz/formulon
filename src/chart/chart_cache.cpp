#include "chart/chart_cache.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <set>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_model.h"
#include "chart/chart_ns.h"
#include "chart/chart_source.h"
#include "chart/chart_xml.h"
#include "drawing/drawing_xml.h"
#include "eval/function_registry.h"
#include "eval/text_format/display_text.h"
#include "io/passthrough_part_index.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_double.h"
#include "num_fmt_locale.h"
#include "sheet.h"
#include "style_resolve.h"
#include "styles.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace {

/// One `c:pt`. A numeric point compares bit-exact; any other value as text.
struct CachePt {
  std::uint64_t idx = 0;
  bool numeric = false;
  double number = 0.0;
  std::string text;
  std::string format_code;  ///< `pt@formatCode`; empty when absent.
};

struct Cache {
  bool present = true;
  std::uint64_t point_count = 0;
  std::string format_code;                 ///< Cache level; numCache only.
  std::vector<std::vector<CachePt>> lvls;  ///< Lowest level first; one entry unless multi-level.
};

enum class RefKind : std::uint8_t { kNum, kStr, kMultiLvl };

bool is_role(std::string_view n) {
  return n == "tx" || n == "cat" || n == "val" || n == "xVal" || n == "yVal" || n == "bubbleSize";
}

bool ref_kind_of(std::string_view n, RefKind& kind) {
  if (n == "numRef") {
    kind = RefKind::kNum;
  } else if (n == "strRef") {
    kind = RefKind::kStr;
  } else if (n == "multiLvlStrRef") {
    kind = RefKind::kMultiLvl;
  } else {
    return false;
  }
  return true;
}

/// Series references: `ser/<role>/<ref>`, including `c15:ser` of filtered series.
void collect_refs(const pugi::xml_node& node, std::vector<pugi::xml_node>& out) {
  for (pugi::xml_node c = node.first_child(); c && c != node;) {
    if (c.type() == pugi::node_element) {
      RefKind kind;
      const pugi::xml_node role = c.parent();
      if (ref_kind_of(local_name(c.name()), kind) && is_role(local_name(role.name())) &&
          local_name(role.parent().name()) == "ser") {
        out.push_back(c);
      } else if (pugi::xml_node child = c.first_child()) {
        c = child;
        continue;
      }
    }

    while (c && c != node && !c.next_sibling()) {
      c = c.parent();
    }
    if (!c || c == node) {
      break;
    }
    c = c.next_sibling();
  }
}

/// `c:f`, or the `c15:formulaRef/c15:sqref` a filtered series keeps instead.
std::string ref_formula(const pugi::xml_node& ref) {
  std::string text = child_local(ref, "f").text().get();
  if (text.empty()) {
    pugi::xml_node ext = child_local(ref, "extLst");
    for (pugi::xml_node e = ext.first_child(); e && text.empty(); e = e.next_sibling()) {
      text = child_local(child_local(e, "formulaRef"), "sqref").text().get();
    }
  }
  return text;
}

struct Pos {
  std::size_t sheet;
  std::uint32_t row;
  std::uint32_t col;
};

class Gatherer {
 public:
  Gatherer(const Workbook& wb, bool vis_only) : wb_(wb), vis_only_(vis_only) {}

  /// Builds the cache `ref` would carry on save. `kind` may change from
  /// numeric to string when a category holds text.
  Cache build(const std::vector<SheetRect>& areas, RefKind& kind, std::string_view role, bool apply_hidden) const {
    const bool hide = vis_only_ && apply_hidden;
    if (kind == RefKind::kMultiLvl && areas.size() == 1 && areas[0].top != areas[0].bottom &&
        areas[0].left != areas[0].right) {
      return build_levels(areas[0], hide);
    }
    std::vector<Item> items;
    Cache cache;
    bool had_cells = false;
    cache.point_count = flatten(areas, hide, items, had_cells);
    if (kind == RefKind::kNum && (role == "cat" || role == "xVal") && has_text(items)) {
      kind = RefKind::kStr;
    }
    cache.lvls.emplace_back();
    if (kind == RefKind::kNum) {
      cache.format_code = format_code_at(second_position(areas));
      if (items.empty() && had_cells && hide) {
        cache.present = false;  // every cell hidden: the ref keeps `c:f` alone
        return cache;
      }
      for (const Item& it : items) {
        numeric_point(it, cache);
      }
    } else if (role == "tx") {
      std::string name;  // a multi-cell name is one point, cells joined by a space
      for (const Item& it : items) {
        const std::string text = display(it.pos, value_at(it.pos));
        name += (name.empty() || text.empty() ? "" : " ") + text;
      }
      cache.point_count = 1;
      if (!name.empty()) {
        cache.lvls[0].push_back(CachePt{0, false, 0.0, name, {}});
      }
    } else if (items.empty() && had_cells && hide) {
      cache.present = false;  // every cell hidden
    } else {
      for (const Item& it : items) {
        const std::string text = display(it.pos, value_at(it.pos));
        if (!text.empty()) {
          cache.lvls[0].push_back(CachePt{it.idx, false, 0.0, text, {}});
        }
      }
    }
    return cache;
  }

 private:
  struct Item {
    std::uint64_t idx;
    Pos pos;
  };

  Value value_at(const Pos& p) const { return wb_.sheet(p.sheet).resolve_cell_value(p.row, p.col); }

  std::string display(const Pos& p, const Value& v) const {
    if (v.is_blank()) {
      return {};
    }
    if (v.is_text()) {
      return std::string(v.as_text());
    }
    return text_format::format_cell_for_display(wb_, wb_.sheet(p.sheet), p.row, p.col).text;
  }

  bool has_text(const std::vector<Item>& items) const {
    for (const Item& it : items) {
      const Value v = value_at(it.pos);
      if (v.is_text() && !v.as_text().empty()) {
        return true;
      }
    }
    return false;
  }

  void numeric_point(const Item& it, Cache& cache) const {
    const Value v = value_at(it.pos);
    CachePt pt;
    pt.idx = it.idx;
    pt.numeric = true;
    if (v.is_blank()) {
      return;  // a formula's "" is text, plotted as 0
    }
    if (v.is_number()) {
      const double number = v.as_number();
      // Non-finite, signed-zero and subnormal values all cache as +0.
      pt.number = std::isnormal(number) ? number : 0.0;
    } else if (v.is_boolean()) {
      pt.number = v.as_boolean() ? 1.0 : 0.0;
    } else if (v.is_error() && v.as_error() == ErrorCode::NA) {
      pt.numeric = false;
      pt.text = "#N/A";
    }
    const std::string code = format_code_at(it.pos);
    if (code != cache.format_code) {
      pt.format_code = code;
    }
    cache.lvls[0].push_back(std::move(pt));
  }

  /// A 2-D category: the last column (or row) is the lowest level.
  Cache build_levels(const SheetRect& r, bool hide) const {
    Cache cache;
    const bool down = (r.bottom - r.top) >= (r.right - r.left);
    const std::uint32_t first = down ? r.top : r.left;
    const std::uint32_t last = down ? r.bottom : r.right;
    const std::uint32_t lvl_first = down ? r.left : r.top;
    const std::uint32_t lvl_last = down ? r.right : r.bottom;
    const Sheet& sheet = wb_.sheet(r.sheet);
    cache.lvls.assign(lvl_last - lvl_first + 1U, {});
    std::uint64_t next = 0;
    for (std::uint32_t i = first; i <= last; ++i) {
      if (hide && (down ? row_hidden(sheet, i) : col_hidden(sheet, i))) {
        continue;
      }
      for (std::uint32_t lv = lvl_last + 1U; lv-- > lvl_first;) {
        const Pos p = down ? Pos{r.sheet, i, lv} : Pos{r.sheet, lv, i};
        const std::string text = display(p, value_at(p));
        if (!text.empty()) {
          cache.lvls[lvl_last - lv].push_back(CachePt{next, false, 0.0, text, {}});
        }
      }
      ++next;
    }
    cache.point_count = next;
    cache.present = next > 0 || !hide;  // every point hidden
    return cache;
  }

  /// Cells in point order (column by column) with hidden ones dropped
  /// and the rest renumbered; returns the point count.
  std::uint64_t flatten(const std::vector<SheetRect>& areas, bool hide, std::vector<Item>& items,
                        bool& had_cells) const {
    std::uint64_t next = 0;
    for (const SheetRect& area : areas) {
      const Sheet& sheet = wb_.sheet(area.sheet);
      const SheetRect r = trim_to_populated(wb_, area);
      std::uint64_t kept = 0;
      for (std::uint32_t col = r.left; col <= r.right && r.top <= r.bottom; ++col) {
        if (hide && col_hidden(sheet, col)) {
          continue;
        }
        for (std::uint32_t row = r.top; row <= r.bottom; ++row) {
          if (!hide || !row_hidden(sheet, row)) {
            items.push_back(Item{next + kept++, Pos{area.sheet, row, col}});
          }
        }
      }
      const std::uint64_t full =
          static_cast<std::uint64_t>(area.bottom - area.top + 1U) * (area.right - area.left + 1U);
      const std::uint64_t seen =
          r.top > r.bottom ? 0U : static_cast<std::uint64_t>(r.bottom - r.top + 1U) * (r.right - r.left + 1U);
      had_cells = had_cells || seen > 0;
      next += kept + (full - seen);
    }
    return next;
  }

  /// The cache's format is that of the range's second cell, or the only one.
  Pos second_position(const std::vector<SheetRect>& areas) const {
    const SheetRect& a = areas[0];
    if (a.bottom > a.top) {
      return Pos{a.sheet, a.top + 1U, a.left};
    }
    if (a.right > a.left) {
      return Pos{a.sheet, a.top, a.left + 1U};
    }
    if (areas.size() > 1) {
      return Pos{areas[1].sheet, areas[1].top, areas[1].left};
    }
    return Pos{a.sheet, a.top, a.left};
  }

  /// Built-in ids are cached with their en-US codes whatever the UI locale, and
  /// a custom code spelling the locale's own short date (id 14) is taken for
  /// that built-in.
  std::string format_code_at(const Pos& p) const {
    const std::vector<CellXf>& xfs = wb_.styles().cell_xfs;
    if (xfs.empty()) {
      return {};
    }
    const EffectiveXf choice = select_effective_xf(wb_.sheet(p.sheet), p.row, p.col);
    const CellXf& xf = xfs[choice.xf_index < xfs.size() ? choice.xf_index : 0U];
    const std::string_view code = text_format::number_format_code_for_xf(wb_.styles(), &xf, mac_365_en_us_profile());
    if (code == localized_builtin_num_fmt(14, wb_.excel_profile())) {
      return std::string(localized_builtin_num_fmt(14, mac_365_en_us_profile()));
    }
    return std::string(code);
  }

  const Workbook& wb_;
  bool vis_only_;
};

std::vector<CachePt> read_points(const pugi::xml_node& parent) {
  std::vector<CachePt> pts;
  for (pugi::xml_node c = parent.first_child(); c; c = c.next_sibling()) {
    if (c.type() == pugi::node_element && local_name(c.name()) == "pt") {
      CachePt pt;
      pt.idx = io::parse_xml_u32_attr(c.attribute("idx"), 0);
      pt.format_code = c.attribute("formatCode").value();
      io::AppendOoxmlTextUnescaped(pt.text, child_local(c, "v").text().get());
      pts.push_back(std::move(pt));
    }
  }
  return pts;
}

/// The cache element under `ref`, or an empty node.
pugi::xml_node find_cache(const pugi::xml_node& ref) {
  for (pugi::xml_node c = ref.first_child(); c; c = c.next_sibling()) {
    if (c.type() != pugi::node_element) {
      continue;
    }
    const std::string_view n = local_name(c.name());
    if (n == "numCache" || n == "strCache" || n == "multiLvlStrCache") {
      return c;
    }
  }
  return {};
}

Cache read_cache(const pugi::xml_node& node) {
  Cache cache;
  cache.present = static_cast<bool>(node);
  if (!node) {
    return cache;
  }
  cache.point_count = io::parse_xml_u32_attr(child_local(node, "ptCount").attribute("val"), 0);
  cache.format_code = child_local(node, "formatCode").text().get();
  if (local_name(node.name()) == "multiLvlStrCache") {
    for (pugi::xml_node c = node.first_child(); c; c = c.next_sibling()) {
      if (c.type() == pugi::node_element && local_name(c.name()) == "lvl") {
        cache.lvls.push_back(read_points(c));
      }
    }
  } else {
    cache.lvls.push_back(read_points(node));
  }
  return cache;
}

bool same_point(const CachePt& fresh, const CachePt& old) {
  if (fresh.idx != old.idx || fresh.format_code != old.format_code) {
    return false;
  }
  if (!fresh.numeric) {
    return fresh.text == old.text;
  }
  double parsed = 0.0;
  if (!io::parse_xsd_double(old.text, &parsed)) {
    return false;
  }
  return std::memcmp(&parsed, &fresh.number, sizeof(double)) == 0;
}

bool same_cache(const Cache& fresh, const Cache& old, RefKind kind) {
  if (fresh.present != old.present) {
    return false;
  }
  if (!fresh.present) {
    return true;
  }
  if (fresh.point_count != old.point_count || (kind == RefKind::kNum && fresh.format_code != old.format_code) ||
      fresh.lvls.size() != old.lvls.size()) {
    return false;
  }
  for (std::size_t l = 0; l < fresh.lvls.size(); ++l) {
    if (fresh.lvls[l].size() != old.lvls[l].size()) {
      return false;
    }
    for (std::size_t i = 0; i < fresh.lvls[l].size(); ++i) {
      if (!same_point(fresh.lvls[l][i], old.lvls[l][i])) {
        return false;
      }
    }
  }
  return true;
}

std::string number_or_text(const CachePt& pt) {
  if (!pt.numeric) {
    return pt.text;
  }
  std::string out;
  io::append_xml_number(out, pt.number);
  return out;
}

void write_points(pugi::xml_node parent, const std::string& prefix, const std::vector<CachePt>& pts) {
  for (const CachePt& pt : pts) {
    pugi::xml_node node = parent.append_child((prefix + "pt").c_str());
    node.append_attribute("idx") = count_text(pt.idx).c_str();
    if (!pt.format_code.empty()) {
      node.append_attribute("formatCode") = pt.format_code.c_str();
    }
    std::string value = number_or_text(pt);
    if (!pt.numeric) {
      std::string encoded;
      io::AppendOoxmlTextEscaped(encoded, value);
      value = std::move(encoded);
    }
    node.append_child((prefix + "v").c_str()).append_child(pugi::node_pcdata).set_value(value.c_str());
  }
}

/// Replaces the cache under `ref` (or inserts one after `c:f`) with `cache`.
void write_cache(pugi::xml_node ref, const std::string& prefix, RefKind kind, const Cache& cache) {
  const pugi::xml_node old = find_cache(ref);
  if (!cache.present) {
    if (old) {
      ref.remove_child(old);
    }
    return;
  }
  const char* name = kind == RefKind::kNum ? "numCache" : kind == RefKind::kStr ? "strCache" : "multiLvlStrCache";
  const std::string qualified = prefix + name;
  pugi::xml_node node;
  if (old) {
    node = ref.insert_child_after(qualified.c_str(), old);
    ref.remove_child(old);
  } else if (const pugi::xml_node f = child_local(ref, "f"); f) {
    node = ref.insert_child_after(qualified.c_str(), f);
  } else {
    node = ref.append_child(qualified.c_str());
  }
  if (kind == RefKind::kNum) {
    node.append_child((prefix + "formatCode").c_str())
        .append_child(pugi::node_pcdata)
        .set_value(cache.format_code.c_str());
  }
  node.append_child((prefix + "ptCount").c_str()).append_attribute("val") = count_text(cache.point_count).c_str();
  if (kind == RefKind::kMultiLvl) {
    for (const std::vector<CachePt>& lvl : cache.lvls) {
      write_points(node.append_child((prefix + "lvl").c_str()), prefix, lvl);
    }
  } else {
    write_points(node, prefix, cache.lvls[0]);
  }
}

}  // namespace

bool refresh_chart_caches(const Workbook& wb, std::size_t host_sheet, pugi::xml_document& doc) {
  const pugi::xml_node root = doc.document_element();
  const std::string_view root_name = root.name();
  const std::size_t root_colon = root_name.find(':');
  const std::string_view root_prefix =
      root_colon == std::string_view::npos ? std::string_view() : root_name.substr(0, root_colon);
  if (local_name(root.name()) != "chartSpace" || io::find_namespace_binding(root, root_prefix) == chart_ns::kNsCx ||
      child_local(root, "pivotSource")) {
    return false;
  }
  const pugi::xml_node vis = child_local(child_local(root, "chart"), "plotVisOnly");
  const char* vis_val = vis.attribute("val").value();
  const Gatherer gatherer(wb, !vis || (std::strcmp(vis_val, "0") != 0 && std::strcmp(vis_val, "false") != 0));
  std::vector<pugi::xml_node> refs;
  collect_refs(root, refs);
  bool changed = false;
  for (pugi::xml_node ref : refs) {
    RefKind kind = RefKind::kNum;
    ref_kind_of(local_name(ref.name()), kind);
    const std::string_view role = local_name(ref.parent().name());
    const std::string formula = ref_formula(ref);
    if (formula.empty()) {
      continue;
    }
    auto source = resolve_chart_source(wb, host_sheet, formula, eval::default_registry());
    if (!source || source.value().areas.empty()) {
      continue;
    }
    const RefKind original = kind;
    // A series name is not plotted, so hidden cells do not drop out of it.
    const Cache fresh = gatherer.build(source.value().areas, kind, role, role != "tx");
    const pugi::xml_node cache_node = find_cache(ref);
    const Cache old = read_cache(cache_node);
    if (kind == original && same_cache(fresh, old, kind)) {
      continue;
    }
    const std::string name = ref.name();
    const std::size_t colon = name.rfind(':');
    const std::string prefix = colon == std::string::npos ? std::string() : name.substr(0, colon + 1);
    if (kind != original) {
      ref.set_name((prefix + "strRef").c_str());
    }
    write_cache(ref, prefix, kind, fresh);
    changed = true;
  }
  return changed;
}

Expected<std::vector<PassthroughPart>, Error> refreshed_chart_parts(const Workbook& wb) {
  ASSIGN_OR_RETURN(std::vector<SheetChartPartRef> charts, all_chart_parts(wb));
  std::vector<PassthroughPart> out;
  const io::PassthroughPartIndex source_index(wb.passthrough_parts());
  std::set<std::string> seen;
  for (const SheetChartPartRef& entry : charts) {
    const std::string& path = entry.ref.part_path;
    if (path.empty() || entry.ref.is_chartex || !seen.insert(path).second) {
      continue;
    }
    const PassthroughPart* part = source_index.find(path);
    pugi::xml_document doc;
    if (part == nullptr || !parse_chart_part(part->bytes, doc)) {
      continue;
    }
    const std::size_t host = wb.sheet(entry.sheet).is_opaque_ooxml_sheet() ? kChartNoHostSheet : entry.sheet;
    if (!refresh_chart_caches(wb, host, doc)) {
      continue;
    }
    std::string xml(io::kXmlDecl);
    io::append_raw_xml(xml, doc.document_element());
    PassthroughPart updated = *part;
    updated.bytes.assign(xml.begin(), xml.end());
    out.push_back(std::move(updated));
  }
  return out;
}

}  // namespace formulon
