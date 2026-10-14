#include "chart/chart_ref.h"

#include <cstdint>
#include <iterator>
#include <optional>
#include <set>
#include <string>
#include <string_view>

#include "chart/chart_ns.h"
#include "drawing/drawing_xml.h"
#include "io/ooxml/part_dom.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace chart_ref {
namespace {

constexpr const char* kRefError = "#REF!";
using chart_ns::kNsC15;
using chart_ns::kUriC15;

bool is_site(const pugi::xml_node& node) {
  const std::string_view name = node.name();
  if (name == "c:f" || name == "c15:f" || name == "cx:f") {
    return true;
  }
  if (name == "c15:sqref") {
    const std::string_view parent = node.parent().name();
    return parent == "c15:fullRef" || parent == "c15:formulaRef";
  }
  return false;
}

struct SiteCollector : pugi::xml_tree_walker {
  std::vector<pugi::xml_node> sites;
  bool for_each(pugi::xml_node& node) override {
    if (node.type() == pugi::node_element && is_site(node)) {
      sites.push_back(node);
    }
    return true;
  }
};

void drop_caches(pugi::xml_node ref) {
  for (const char* name : {"c:strCache", "c:numCache", "c:multiLvlStrCache"}) {
    while (ref.child(name)) {
      ref.remove_child(ref.child(name));
    }
  }
}

/// Sheet qualifier of a reference area; empty when it carries none.
std::string_view area_sheet(const parser::AstNode& area) {
  switch (area.kind()) {
    case parser::NodeKind::Ref:
      return area.as_ref().sheet;
    case parser::NodeKind::RangeOp:
      return area_sheet(area.as_range_lhs());
    case parser::NodeKind::NameRef:
      return area.as_name_sheet();
    default:
      return {};
  }
}

/// Columns and rows of a cell or rectangle area; false for any other form.
bool area_span(const parser::AstNode& area, std::uint32_t& cols, std::uint32_t& rows) {
  if (area.kind() == parser::NodeKind::Ref) {
    const parser::Reference& ref = area.as_ref();
    cols = rows = 1U;
    return !ref.is_full_col && !ref.is_full_row;
  }
  if (area.kind() != parser::NodeKind::RangeOp || area.as_range_lhs().kind() != parser::NodeKind::Ref ||
      area.as_range_rhs().kind() != parser::NodeKind::Ref) {
    return false;
  }
  const parser::Reference& lhs = area.as_range_lhs().as_ref();
  const parser::Reference& rhs = area.as_range_rhs().as_ref();
  if (lhs.is_full_col || lhs.is_full_row || rhs.is_full_col || rhs.is_full_row) {
    return false;
  }
  cols = (lhs.col > rhs.col ? lhs.col - rhs.col : rhs.col - lhs.col) + 1U;
  rows = (lhs.row > rhs.row ? lhs.row - rhs.row : rhs.row - lhs.row) + 1U;
  return true;
}

/// A rewritten site: its text, whether any of its areas became `#REF!`, and
/// the spans of a single rectangle before and after (zero when not one).
struct SiteRewrite {
  std::string text;
  bool deleted = false;
  std::uint32_t old_cols = 0;
  std::uint32_t old_rows = 0;
  std::uint32_t new_cols = 0;
  std::uint32_t new_rows = 0;
};

/// Rewrites a chart reference area by area. An area whose cells are deleted
/// keeps its sheet qualifier (`Sheet1!#REF!`) unless the sheet itself is gone.
/// Empty when the text does not parse or nothing changed.
std::optional<SiteRewrite> rewrite_site(std::string_view text, const parser::RefTransform& transform) {
  Arena arena;
  parser::Parser parser(text, arena);
  const parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    return std::nullopt;
  }
  const bool is_union = root->kind() == parser::NodeKind::UnionOp;
  const std::uint32_t areas = is_union ? root->as_union_arity() : 1U;
  SiteRewrite out;
  bool changed = false;
  for (std::uint32_t i = 0; i < areas; ++i) {
    const parser::AstNode& area = is_union ? root->as_union_child(i) : *root;
    const parser::AstNode* shifted = parser::shift_refs(area, arena, transform);
    if (shifted == nullptr) {
      return std::nullopt;
    }
    changed = changed || shifted != &area;
    if (!is_union && area_span(area, out.old_cols, out.old_rows) && !area_span(*shifted, out.new_cols, out.new_rows)) {
      out.old_cols = out.old_rows = 0U;
    }
    if (i > 0U) {
      out.text.push_back(',');
    }
    std::string spelled = parser::format_formula(*shifted);
    if (shifted != &area && spelled == kRefError) {
      const std::string_view sheet = area_sheet(area);
      const std::optional<std::string_view> kept = sheet.empty() ? std::nullopt : transform.apply_name_sheet(sheet);
      if (kept.has_value()) {
        spelled.clear();
        parser::append_sheet_name(*kept, false, spelled);
        spelled.push_back('!');
        spelled.append(kRefError);
      }
      out.deleted = true;
    }
    out.text.append(spelled);
  }
  if (!changed) {
    return std::nullopt;
  }
  if (is_union) {
    out.text = "(" + out.text + ")";
  }
  return out;
}

/// Replaces the cache beside a site with a newly deleted area: a name gets
/// one `#REF!` string point, a category cache is dropped, a value cache holds
/// the single point 1.
void reset_cache(const pugi::xml_node& site) {
  pugi::xml_node ref = site.parent();
  const std::string_view kind = ref.name();
  if (std::string_view(site.name()) != "c:f" ||
      (kind != "c:strRef" && kind != "c:numRef" && kind != "c:multiLvlStrRef")) {
    return;
  }
  const std::string_view role = ref.parent().name();
  drop_caches(ref);
  if (role == "c:cat" || kind == "c:multiLvlStrRef") {
    return;
  }
  const bool is_name = role == "c:tx";
  if (!is_name && role != "c:val" && role != "c:xVal" && role != "c:yVal" && role != "c:bubbleSize") {
    return;
  }
  pugi::xml_node cache = ref.insert_child_after(is_name ? "c:strCache" : "c:numCache", site);
  if (!is_name) {
    cache.append_child("c:formatCode").text().set("General");
  }
  cache.append_child("c:ptCount").append_attribute("val").set_value("1");
  pugi::xml_node pt = cache.append_child("c:pt");
  pt.append_attribute("idx").set_value("0");
  pt.append_child("c:v").text().set(is_name ? "#REF!" : "1");
}

/// Moves a series name or category whose reference was deleted out of `c:tx`
/// / `c:cat` into `c15:filteredSeriesTitle` / `c15:filteredCategoryTitle`
/// under the series `c:extLst`, as Excel does: the cache travels with it and
/// the reference text becomes a `c15:formulaRef`.
void file_as_filtered_title(const pugi::xml_node& site) {
  pugi::xml_node ref = site.parent();
  const std::string_view kind = ref.name();
  pugi::xml_node holder = ref.parent();
  const std::string_view role = holder.name();
  pugi::xml_node ser = holder.parent();
  if ((kind != "c:strRef" && kind != "c:numRef" && kind != "c:multiLvlStrRef") || (role != "c:tx" && role != "c:cat") ||
      std::string_view(ser.name()) != "c:ser") {
    return;
  }
  const bool is_name = role == "c:tx";
  pugi::xml_node ext_list = ser.child("c:extLst");
  if (!ext_list) {
    ext_list = ser.append_child("c:extLst");
  }
  // Filtered titles lead the list, the name before the category.
  pugi::xml_node after;
  for (pugi::xml_node ext : ext_list.children("c:ext")) {
    if (std::string_view(ext.first_child().name()).substr(0, 12) == "c15:filtered") {
      after = ext;
    }
  }
  pugi::xml_node ext = after ? ext_list.insert_child_after("c:ext", after) : ext_list.prepend_child("c:ext");
  ext.append_attribute("uri").set_value(kUriC15);
  ext.append_attribute("xmlns:c15").set_value(kNsC15);
  pugi::xml_node title = ext.append_child(is_name ? "c15:filteredSeriesTitle" : "c15:filteredCategoryTitle");
  pugi::xml_node moved = title.append_child(is_name ? "c15:tx" : "c15:cat").append_move(ref);
  const std::string text = site.child_value();
  moved.remove_child(site);
  pugi::xml_node ref_ext_list = moved.child("c:extLst");
  if (!ref_ext_list) {
    ref_ext_list = moved.prepend_child("c:extLst");
  }
  pugi::xml_node formula_ext = ref_ext_list.append_child("c:ext");
  formula_ext.append_attribute("uri").set_value(kUriC15);
  formula_ext.append_child("c15:formulaRef").append_child("c15:sqref").text().set(text.c_str());
  ser.remove_child(holder);
}

/// Whether `ser` has lost its values: its `c:val` / `c:yVal` is one `#REF!` area.
bool lost_values(const pugi::xml_node& ser) {
  pugi::xml_node holder = ser.child("c:val");
  if (!holder) {
    holder = ser.child("c:yVal");
  }
  const std::string_view text = holder.first_child().child("c:f").child_value();
  const std::string_view tail = "!#REF!";
  return text == kRefError || (text.size() > tail.size() && text.substr(text.size() - tail.size()) == tail);
}

/// Whether a deleted site stays where it is. Excel keeps the name of a series
/// that lost its values in `c:tx` while another series of the chart keeps
/// its own; every other deleted name or category is filed as filtered.
bool keeps_place(const pugi::xml_node& site, bool chart_has_values) {
  const pugi::xml_node holder = site.parent().parent();
  return chart_has_values && std::string_view(holder.name()) == "c:tx" &&
         std::string_view(holder.parent().name()) == "c:ser" && lost_values(holder.parent());
}

/// Turns a multi-level category whose levels shrank to one into the plain
/// `c:strRef` Excel writes; the save-time cache refresh rebuilds its cache.
void flatten_category(const pugi::xml_node& site, const SiteRewrite& rewrite) {
  pugi::xml_node ref = site.parent();
  if (std::string_view(ref.name()) != "c:multiLvlStrRef") {
    return;
  }
  const auto lvls = ref.child("c:multiLvlStrCache").children("c:lvl");
  const auto levels = static_cast<std::uint32_t>(std::distance(lvls.begin(), lvls.end()));
  const bool by_cols = rewrite.old_cols > 1U && rewrite.old_cols == levels && rewrite.new_cols == 1U;
  // A square area reads its levels across columns, the by-column default.
  const bool by_rows =
      rewrite.old_rows > 1U && rewrite.old_rows == levels && rewrite.old_cols != levels && rewrite.new_rows == 1U;
  if (!by_cols && !by_rows) {
    return;
  }
  drop_caches(ref);
  ref.set_name("c:strRef");
}

}  // namespace

std::vector<pugi::xml_node> enumerate_sites(const pugi::xml_document& doc) {
  SiteCollector collector;
  pugi::xml_node root = doc.document_element();
  root.traverse(collector);
  return std::move(collector.sites);
}

std::size_t rewrite_sites(pugi::xml_document& doc, const parser::RefTransform& transform) {
  std::size_t changed = 0;
  std::vector<pugi::xml_node> deleted;
  for (pugi::xml_node site : enumerate_sites(doc)) {
    const std::string_view text = site.child_value();
    if (text.empty()) {
      continue;
    }
    const std::optional<SiteRewrite> result = rewrite_site(text, transform);
    if (!result.has_value()) {
      continue;
    }
    site.text().set(result->text.c_str());
    ++changed;
    if (result->deleted) {
      reset_cache(site);
      deleted.push_back(site);
    } else {
      flatten_category(site, *result);
    }
  }
  bool chart_has_values = false;
  for (const pugi::xpath_node& ser : doc.select_nodes("//c:ser")) {
    chart_has_values = chart_has_values || !lost_values(ser.node());
  }
  for (const pugi::xml_node& site : deleted) {
    if (!keeps_place(site, chart_has_values)) {
      file_as_filtered_title(site);
    }
  }
  return changed;
}

Expected<void, Error> rewrite_all(Workbook& wb, const parser::RefTransform& transform) {
  auto parts = all_chart_parts(wb);
  if (!parts) {
    return parts.error();
  }
  std::set<std::string> seen;
  for (const SheetChartPartRef& entry : parts.value()) {
    const std::string& path = entry.ref.part_path;
    if (path.empty() || !seen.insert(path).second) {
      continue;
    }
    pugi::xml_document doc;
    if (!io::ooxml::load_part_dom(wb, path, "chart_ref", doc)) {
      continue;
    }
    if (rewrite_sites(doc, transform) == 0) {
      continue;
    }
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace chart_ref
}  // namespace formulon
