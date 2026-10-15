#include "chart/chart_ref.h"

#include <cstdint>
#include <initializer_list>
#include <optional>
#include <set>
#include <string>
#include <string_view>

#include "chart/chart_ns.h"
#include "drawing/drawing_xml.h"
#include "io/ooxml/part_dom.h"
#include "io/xml_utils.h"
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
using chart_ns::kNsC;
using chart_ns::kNsC15;
using chart_ns::kNsCx;
using chart_ns::kUriC15;

std::string_view prefix_of_qname(std::string_view name) {
  const std::size_t colon = name.find(':');
  return colon == std::string_view::npos ? std::string_view() : name.substr(0, colon);
}

std::string_view namespace_uri(const pugi::xml_node& node) {
  return io::find_namespace_binding(node, prefix_of_qname(node.name()));
}

bool is_element(const pugi::xml_node& node, std::string_view uri, std::string_view local) {
  return node.type() == pugi::node_element && namespace_uri(node) == uri && local_name(node) == local;
}

bool is_chart_element(const pugi::xml_node& node, std::string_view local) {
  return is_element(node, kNsC, local);
}

std::string qualified_name_like(const pugi::xml_node& node, std::string_view local) {
  const std::string_view name = node.name();
  const std::size_t colon = name.find(':');
  if (colon == std::string_view::npos) {
    return std::string(local);
  }
  std::string out(name.substr(0, colon));
  out.push_back(':');
  out.append(local);
  return out;
}

std::string qualified_name_with_prefix(std::string_view prefix, std::string_view local) {
  std::string out(prefix);
  out.push_back(':');
  out.append(local);
  return out;
}

/// Finds a c15 namespace prefix that will not change the expanded name of the
/// chart element or any content that is moved under it. A candidate already
/// used for another namespace by either side of the move is left untouched.
std::string extension_prefix(const pugi::xml_node& source_scope, const pugi::xml_node& destination_scope,
                             std::string_view chart_qname) {
  for (std::uint32_t suffix = 0;; ++suffix) {
    std::string prefix = suffix == 0 ? "c15" : "c15ext";
    if (suffix > 1) {
      prefix += std::to_string(suffix - 1);
    }
    if (std::string_view(prefix) == prefix_of_qname(chart_qname)) {
      continue;
    }
    const std::string_view source_binding = io::find_namespace_binding(source_scope, prefix);
    const std::string_view destination_binding = io::find_namespace_binding(destination_scope, prefix);
    if ((!source_binding.empty() && source_binding != kNsC15) ||
        (!destination_binding.empty() && destination_binding != kNsC15)) {
      continue;
    }
    return prefix;
  }
}

void declare_namespace(pugi::xml_node node, std::string_view prefix, std::string_view uri) {
  const std::string name = "xmlns:" + std::string(prefix);
  const std::string value(uri);
  node.append_attribute(name.c_str()).set_value(value.c_str());
}

pugi::xml_node child_element(const pugi::xml_node& parent, std::string_view uri, std::string_view local) {
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    if (is_element(child, uri, local)) {
      return child;
    }
  }
  return {};
}

void preserve_local_namespace_declarations(const pugi::xml_node& source, pugi::xml_node destination) {
  for (pugi::xml_attribute attr = source.first_attribute(); attr; attr = attr.next_attribute()) {
    const std::string_view name = attr.name();
    if (name != "xmlns" && name.rfind("xmlns:", 0) != 0) {
      continue;
    }
    if (destination.attribute(attr.name())) {
      continue;
    }
    destination.append_attribute(attr.name()).set_value(attr.value());
  }
}

void preserve_bindings_shadowed_by(const pugi::xml_node& source_scope, const pugi::xml_node& new_ancestor,
                                   pugi::xml_node moved_root) {
  for (pugi::xml_attribute attr = new_ancestor.first_attribute(); attr; attr = attr.next_attribute()) {
    const std::string_view name = attr.name();
    if (name != "xmlns" && name.rfind("xmlns:", 0) != 0) {
      continue;
    }
    const std::string_view prefix = name == "xmlns" ? std::string_view() : name.substr(6);
    const std::string declaration_name(name);
    if (moved_root.attribute(declaration_name.c_str())) {
      continue;
    }
    const std::string_view old_binding = io::find_namespace_binding(source_scope, prefix);
    if (old_binding == attr.value() || (old_binding.empty() && !prefix.empty())) {
      continue;
    }
    const std::string old_value(old_binding);
    moved_root.append_attribute(declaration_name.c_str()).set_value(old_value.c_str());
  }
}

bool is_site(const pugi::xml_node& node) {
  if (is_chart_element(node, "f") || is_element(node, kNsC15, "f") || is_element(node, kNsCx, "f")) {
    return true;
  }
  if (!is_element(node, kNsC15, "sqref")) {
    return false;
  }
  const pugi::xml_node parent = node.parent();
  return is_element(parent, kNsC15, "fullRef") || is_element(parent, kNsC15, "formulaRef");
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
  for (const std::string_view local : {"strCache", "numCache", "multiLvlStrCache"}) {
    for (pugi::xml_node cache = child_element(ref, kNsC, local); cache; cache = child_element(ref, kNsC, local)) {
      ref.remove_child(cache);
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
  const std::string_view kind = local_name(ref);
  if (!is_chart_element(site, "f") || namespace_uri(ref) != kNsC ||
      (kind != "strRef" && kind != "numRef" && kind != "multiLvlStrRef")) {
    return;
  }
  const pugi::xml_node holder = ref.parent();
  if (namespace_uri(holder) != kNsC) {
    return;
  }
  const std::string_view role = local_name(holder);
  drop_caches(ref);
  if (role == "cat" || kind == "multiLvlStrRef") {
    return;
  }
  const bool is_name = role == "tx";
  if (!is_name && role != "val" && role != "xVal" && role != "yVal" && role != "bubbleSize") {
    return;
  }
  const std::string cache_name = qualified_name_like(ref, is_name ? "strCache" : "numCache");
  pugi::xml_node cache = ref.insert_child_after(cache_name.c_str(), site);
  if (!is_name) {
    const std::string format_code_name = qualified_name_like(ref, "formatCode");
    cache.append_child(format_code_name.c_str()).text().set("General");
  }
  const std::string pt_count_name = qualified_name_like(ref, "ptCount");
  cache.append_child(pt_count_name.c_str()).append_attribute("val").set_value("1");
  const std::string pt_name = qualified_name_like(ref, "pt");
  pugi::xml_node pt = cache.append_child(pt_name.c_str());
  pt.append_attribute("idx").set_value("0");
  const std::string value_name = qualified_name_like(ref, "v");
  pt.append_child(value_name.c_str()).text().set(is_name ? "#REF!" : "1");
}

/// Moves a series name or category whose reference was deleted out of `c:tx`
/// / `c:cat` into `c15:filteredSeriesTitle` / `c15:filteredCategoryTitle`
/// under the series `c:extLst`, as Excel does: the cache travels with it and
/// the reference text becomes a `c15:formulaRef`.
void file_as_filtered_title(const pugi::xml_node& site) {
  pugi::xml_node ref = site.parent();
  const std::string_view kind = local_name(ref);
  pugi::xml_node holder = ref.parent();
  const std::string_view role = local_name(holder);
  pugi::xml_node ser = holder.parent();
  if (namespace_uri(ref) != kNsC || namespace_uri(holder) != kNsC || !is_chart_element(ser, "ser") ||
      (kind != "strRef" && kind != "numRef" && kind != "multiLvlStrRef") || (role != "tx" && role != "cat")) {
    return;
  }
  const bool is_name = role == "tx";
  // The holder may be removed; keep its local xmlns on the moved reference.
  preserve_local_namespace_declarations(holder, ref);
  pugi::xml_node ext_list = child_element(ser, kNsC, "extLst");
  if (!ext_list) {
    const std::string ext_list_name = qualified_name_like(ser, "extLst");
    ext_list = ser.append_child(ext_list_name.c_str());
  }
  // Keep the source binding wherever `ext_list` shadows it.
  preserve_bindings_shadowed_by(ref, ext_list, ref);
  // Filtered titles lead the list, the name before the category.
  pugi::xml_node after;
  for (pugi::xml_node ext = child_element(ext_list, kNsC, "ext"); ext; ext = ext.next_sibling()) {
    if (!is_chart_element(ext, "ext")) {
      continue;
    }
    const pugi::xml_node title = io::first_element_child(ext);
    const std::string_view title_local = local_name(title);
    if (namespace_uri(title) == kNsC15 && title_local.substr(0, 8) == "filtered") {
      after = ext;
    }
  }
  const std::string ext_name = qualified_name_like(ext_list, "ext");
  const std::string wrapper_prefix = extension_prefix(ref, ext_list, ext_name);
  pugi::xml_node ext =
      after ? ext_list.insert_child_after(ext_name.c_str(), after) : ext_list.prepend_child(ext_name.c_str());
  ext.append_attribute("uri").set_value(kUriC15);
  declare_namespace(ext, wrapper_prefix, kNsC15);
  const std::string title_name =
      qualified_name_with_prefix(wrapper_prefix, is_name ? "filteredSeriesTitle" : "filteredCategoryTitle");
  const std::string holder_name = qualified_name_with_prefix(wrapper_prefix, is_name ? "tx" : "cat");
  pugi::xml_node title = ext.append_child(title_name.c_str());
  pugi::xml_node moved = title.append_child(holder_name.c_str()).append_move(ref);
  const std::string text = site.child_value();
  moved.remove_child(site);
  pugi::xml_node ref_ext_list = child_element(moved, kNsC, "extLst");
  if (!ref_ext_list) {
    const std::string ref_ext_list_name = qualified_name_like(moved, "extLst");
    ref_ext_list = moved.prepend_child(ref_ext_list_name.c_str());
  }
  const std::string formula_ext_name = qualified_name_like(ref_ext_list, "ext");
  const std::string formula_prefix = extension_prefix(moved, ref_ext_list, formula_ext_name);
  pugi::xml_node formula_ext = ref_ext_list.append_child(formula_ext_name.c_str());
  formula_ext.append_attribute("uri").set_value(kUriC15);
  if (io::find_namespace_binding(ref_ext_list, formula_prefix) != kNsC15) {
    declare_namespace(formula_ext, formula_prefix, kNsC15);
  }
  const std::string formula_ref_name = qualified_name_with_prefix(formula_prefix, "formulaRef");
  const std::string sqref_name = qualified_name_with_prefix(formula_prefix, "sqref");
  formula_ext.append_child(formula_ref_name.c_str()).append_child(sqref_name.c_str()).text().set(text.c_str());
  ser.remove_child(holder);
}

/// Whether `ser` has lost its values: its `c:val` / `c:yVal` is one `#REF!` area.
bool lost_values(const pugi::xml_node& ser) {
  pugi::xml_node holder = child_element(ser, kNsC, "val");
  if (!holder) {
    holder = child_element(ser, kNsC, "yVal");
  }
  const pugi::xml_node ref = io::first_element_child(holder);
  const pugi::xml_node formula = child_element(ref, kNsC, "f");
  const std::string_view text = formula.child_value();
  const std::string_view tail = "!#REF!";
  return text == kRefError || (text.size() > tail.size() && text.substr(text.size() - tail.size()) == tail);
}

/// Whether a deleted site stays where it is. Excel keeps the name of a series
/// that lost its values in `c:tx` while another series of the chart keeps
/// its own; every other deleted name or category is filed as filtered.
bool keeps_place(const pugi::xml_node& site, bool chart_has_values) {
  const pugi::xml_node holder = site.parent().parent();
  return chart_has_values && is_chart_element(holder, "tx") && is_chart_element(holder.parent(), "ser") &&
         lost_values(holder.parent());
}

/// Turns a multi-level category whose levels shrank to one into the plain
/// `c:strRef` Excel writes; the save-time cache refresh rebuilds its cache.
void flatten_category(const pugi::xml_node& site, const SiteRewrite& rewrite) {
  pugi::xml_node ref = site.parent();
  if (!is_chart_element(ref, "multiLvlStrRef")) {
    return;
  }
  std::uint32_t levels = 0;
  const pugi::xml_node cache = child_element(ref, kNsC, "multiLvlStrCache");
  for (pugi::xml_node level = child_element(cache, kNsC, "lvl"); level; level = level.next_sibling()) {
    if (is_chart_element(level, "lvl")) {
      ++levels;
    }
  }
  const bool by_cols = rewrite.old_cols > 1U && rewrite.old_cols == levels && rewrite.new_cols == 1U;
  // A square area reads its levels across columns, the by-column default.
  const bool by_rows =
      rewrite.old_rows > 1U && rewrite.old_rows == levels && rewrite.old_cols != levels && rewrite.new_rows == 1U;
  if (!by_cols && !by_rows) {
    return;
  }
  drop_caches(ref);
  const std::string str_ref_name = qualified_name_like(ref, "strRef");
  ref.set_name(str_ref_name.c_str());
}

struct SeriesValueCollector : pugi::xml_tree_walker {
  bool chart_has_values = false;
  bool for_each(pugi::xml_node& node) override {
    if (is_chart_element(node, "ser") && !lost_values(node)) {
      chart_has_values = true;
    }
    return true;
  }
};

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
  SeriesValueCollector series_collector;
  pugi::xml_node root = doc.document_element();
  root.traverse(series_collector);
  const bool chart_has_values = series_collector.chart_has_values;
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
