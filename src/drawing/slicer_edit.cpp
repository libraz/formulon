#include "drawing/slicer_edit.h"

#include <algorithm>
#include <cctype>
#include <cstdlib>
#include <limits>
#include <string>
#include <string_view>
#include <utility>

#include "drawing/drawing_edit.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::string_view kRelSlicer = "http://schemas.microsoft.com/office/2007/relationships/slicer";
constexpr std::string_view kRelCache = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
constexpr std::int64_t kEmuLimit = std::int64_t{1} << 31;
constexpr std::int64_t kDefaultRowHeight = 241300;
constexpr std::uint32_t kMaxColumnCount = 20000;

// Skeletons of the new parts and of the extension-list entries that point at them.
constexpr const char* kSlicersRoot =
    "<slicers xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" "
    "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>";
constexpr const char* kCacheRoot =
    "<slicerCacheDefinition xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" "
    "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><extLst><x:ext "
    "uri=\"{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/"
    "2010/11/main\"><x15:tableSlicerCache/></x:ext></extLst></slicerCacheDefinition>";
constexpr const char* kDrawingRoot =
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>";
constexpr const char* kSheetExt =
    "<ext uri=\"{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}\" xmlns:x15=\"http://schemas.microsoft.com/office/"
    "spreadsheetml/2010/11/main\"><x14:slicerList xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/"
    "9/main\"/></ext>";
constexpr const char* kBookExt =
    "<ext uri=\"{46BE6895-7355-4a93-B00E-2C351335B9C9}\" xmlns:x15=\"http://schemas.microsoft.com/office/"
    "spreadsheetml/2010/11/main\"><x15:slicerCaches xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/"
    "2009/9/main\"/></ext>";

Error invalid(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
}

Error unparseable(const char* message) {
  return make_error(FormulonErrorCode::kIoDrawingUnparseable, message);
}

bool is_element(const pugi::xml_node& node, std::string_view name) {
  return node.type() == pugi::node_element && local_name(node) == name;
}

pugi::xml_node find_element(const pugi::xml_node& root, std::string_view name) {
  return root.find_node([name](const pugi::xml_node& node) { return is_element(node, name); });
}

// The `tableSlicerCache` of a loaded cache part; null for a pivot cache.
pugi::xml_node table_cache_of(const pugi::xml_document& doc) {
  return find_element(doc, "tableSlicerCache");
}

// Loads the table-slicer cache named `name` into `doc` and returns its part path, or empty.
std::string load_cache(const Workbook& wb, std::string_view name, pugi::xml_document& doc) {
  for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
    if (rel.type == kRelCache && !rel.target_external && io::ooxml::load_part_dom(wb, rel.target, "slicer", doc) &&
        table_cache_of(doc) && name == doc.document_element().attribute("name").value()) {
      return rel.target;
    }
  }
  return {};
}

// The drawing anchor holding the frame of the slicer called `name`.
pugi::xml_node find_frame(const pugi::xml_document& drawing, std::string_view name) {
  for (const pugi::xml_node& anchor : drawing_anchors(drawing.document_element(), false)) {
    if (anchor.find_node([name](const pugi::xml_node& node) {
          return is_element(node, "slicer") && name == node.attribute("name").value();
        })) {
      return anchor;
    }
  }
  return {};
}

// One table slicer and the parts that hold it.
struct Slot {
  SlicerInfo info;
  std::uint32_t table_id = 0;
  std::string slicers_part;
  std::string cache_part;
};

Expected<void, Error> load_drawing(const Workbook& wb, std::size_t sheet, pugi::xml_document& drawing) {
  const std::string& path = wb.sheet(sheet).drawing_rel_target();
  if (path.empty()) {
    return Expected<void, Error>::Ok();
  }
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path);
  if (part == nullptr) {
    return unparseable("slicer: drawing part is missing");
  }
  return parse_drawing_part(part->bytes, drawing);
}

// Walks the table slicers of `sheet` in list order. In list mode (`all`) every slicer's info is appended; in lookup
// mode (`out`) the walk stops at the slicer with ordinal `want`, leaves it in `*out` and returns true.
Expected<bool, Error> scan(const Workbook& wb, std::size_t sheet, std::size_t want, Slot* out,
                           std::vector<SlicerInfo>* all) {
  if (sheet >= wb.sheet_count()) {
    return invalid("slicer: sheet index out of range");
  }
  pugi::xml_document drawing;
  bool drawing_loaded = false;
  std::size_t ordinal = 0;
  for (const UnknownRelationship& rel : wb.sheet(sheet).unknown_relationships()) {
    pugi::xml_document doc;
    if (rel.type != kRelSlicer || rel.target_external || !io::ooxml::load_part_dom(wb, rel.target, "slicer", doc)) {
      continue;
    }
    if (!drawing_loaded) {
      RETURN_IF_ERROR(load_drawing(wb, sheet, drawing));
      drawing_loaded = true;
    }
    for (const pugi::xml_node& node : children_local(doc.document_element(), "slicer")) {
      Slot slot;
      SlicerInfo& info = slot.info;
      info.cache_name = node.attribute("cache").value();
      pugi::xml_document cache;
      slot.cache_part = load_cache(wb, info.cache_name, cache);
      if (slot.cache_part.empty()) {
        continue;
      }
      slot.slicers_part = rel.target;
      info.name = node.attribute("name").value();
      info.caption = node.attribute("caption").value();
      info.style = node.attribute("style").value();
      info.column_count = node.attribute("columnCount").as_uint(1);
      info.row_height = node.attribute("rowHeight").as_llong(0);
      const pugi::xml_node table_cache = table_cache_of(cache);
      slot.table_id = table_cache.attribute("tableId").as_uint();
      info.column_name = cache.document_element().attribute("sourceName").value();
      for (const TableMetadata& table : wb.tables()) {
        if (table.id != slot.table_id) {
          continue;
        }
        info.table_name = table.name;
        for (const TableColumn& column : table.columns) {
          if (column.id == table_cache.attribute("column").as_uint()) {
            info.column_name = column.name;
          }
        }
      }
      const std::string_view cross = table_cache.attribute("crossFilter").value();
      info.sort_order = std::string_view(table_cache.attribute("sortOrder").value()) == "descending"
                            ? SlicerSortOrder::kDescending
                            : SlicerSortOrder::kAscending;
      info.cross_filter = cross == "none"                  ? SlicerCrossFilter::kNone
                          : cross == "showItemsWithNoData" ? SlicerCrossFilter::kShowItemsWithNoData
                                                           : SlicerCrossFilter::kShowItemsWithDataAtTop;
      if (const pugi::xml_node anchor = find_frame(drawing, info.name)) {
        info.object = read_drawing_object(anchor, {});
      }
      if (all != nullptr) {
        all->push_back(std::move(info));
      } else if (ordinal == want) {
        *out = std::move(slot);
        return true;
      }
      ++ordinal;
    }
  }
  return false;
}

// Loads slicer `index` of `sheet` into `slot`.
Expected<void, Error> slot_at(const Workbook& wb, std::size_t sheet, std::size_t index, Slot& slot) {
  ASSIGN_OR_RETURN(const bool found, scan(wb, sheet, index, &slot, nullptr));
  if (!found) {
    return invalid("slicer: index out of range");
  }
  return Expected<void, Error>::Ok();
}

// Loads the slicers part of `slot` into `doc` and returns the element of its slicer.
Expected<pugi::xml_node, Error> load_slicer_node(const Workbook& wb, const Slot& slot, pugi::xml_document& doc) {
  RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, slot.slicers_part, "slicer", doc));
  for (const pugi::xml_node& node : children_local(doc.document_element(), "slicer")) {
    if (slot.info.name == node.attribute("name").value()) {
      return node;
    }
  }
  return invalid("slicer: slicer element is missing");
}

void put_attr(pugi::xml_node node, const char* name, const std::string& value) {
  pugi::xml_attribute attr = node.attribute(name);
  (attr ? attr : node.append_attribute(name)).set_value(value.c_str());
}

// Sets `name` to `value`, or removes it when `value` is the schema default.
void put_attr_unless(pugi::xml_node node, const char* name, const std::string& value, std::string_view dflt) {
  if (value == dflt) {
    node.remove_attribute(name);
  } else {
    put_attr(node, name, value);
  }
}

void apply_slicer(pugi::xml_node node, const SlicerSpec& s) {
  using namespace slicer_field;
  if ((s.fields & kCaption) != 0U) {
    put_attr(node, "caption", s.caption);
  }
  if ((s.fields & kStyle) != 0U) {
    put_attr_unless(node, "style", s.style, "");
  }
  if ((s.fields & kColumnCount) != 0U) {
    put_attr_unless(node, "columnCount", std::to_string(s.column_count), "1");
  }
  if ((s.fields & kRowHeight) != 0U) {
    put_attr(node, "rowHeight", std::to_string(s.row_height));
  }
}

void apply_cache(pugi::xml_node table_cache, const SlicerSpec& s) {
  constexpr const char* kCrossNames[] = {"showItemsWithDataAtTop", "showItemsWithNoData", "none"};
  if ((s.fields & slicer_field::kSortOrder) != 0U) {
    put_attr_unless(table_cache, "sortOrder", s.sort_order == SlicerSortOrder::kDescending ? "descending" : "ascending",
                    "ascending");
  }
  if ((s.fields & slicer_field::kCrossFilter) != 0U) {
    put_attr_unless(table_cache, "crossFilter", kCrossNames[static_cast<std::size_t>(s.cross_filter)], kCrossNames[0]);
  }
}

// Moves a two-cell frame to the spec's markers and its fallback shape to the matching box.
void apply_anchor(const Workbook& wb, std::size_t sheet, pugi::xml_node anchor, const SlicerSpec& s) {
  write_anchor_point(child_local(anchor, "from"), s.from);
  write_anchor_point(child_local(anchor, "to"), s.to);
  const pugi::xml_node fallback = find_element(anchor, "Fallback");
  const pugi::xml_node off = find_element(fallback, "off");
  const pugi::xml_node ext = find_element(fallback, "ext");
  if (!off || !ext) {
    return;
  }
  const auto pos = [&](bool rows, const AnchorPoint& p) {
    return sheet_offset_emu(wb, sheet, rows, rows ? p.row : p.col) + (rows ? p.row_off : p.col_off);
  };
  const std::int64_t x = pos(false, s.from);
  const std::int64_t y = pos(true, s.from);
  put_attr(off, "x", std::to_string(x));
  put_attr(off, "y", std::to_string(y));
  put_attr(ext, "cx", std::to_string(pos(false, s.to) - x));
  put_attr(ext, "cy", std::to_string(pos(true, s.to) - y));
}

bool has_control(const std::string& text) {
  return std::any_of(text.begin(), text.end(), [](char c) { return static_cast<unsigned char>(c) < 0x20U; });
}

bool bad_marker(const AnchorPoint& p) {
  return p.row >= Sheet::kMaxRows || p.col >= Sheet::kMaxCols || p.row_off < 0 || p.row_off >= kEmuLimit ||
         p.col_off < 0 || p.col_off >= kEmuLimit;
}

Expected<void, Error> check_spec(const SlicerSpec& s) {
  using namespace slicer_field;
  if ((s.fields & ~kAll) != 0U || has_control(s.name) || has_control(s.caption) || has_control(s.style) ||
      ((s.fields & kName) != 0U && s.name.empty())) {
    return invalid("slicer: unknown field bit, empty name or control character in text");
  }
  if (((s.fields & kColumnCount) != 0U && (s.column_count == 0U || s.column_count > kMaxColumnCount)) ||
      ((s.fields & kRowHeight) != 0U && (s.row_height <= 0 || s.row_height >= kEmuLimit))) {
    return invalid("slicer: column count or row height out of range");
  }
  if ((s.fields & kAnchor) != 0U &&
      (bad_marker(s.from) || bad_marker(s.to) || s.to.row < s.from.row || s.to.col < s.from.col)) {
    return invalid("slicer: anchor lies outside the sheet or ends before it starts");
  }
  return Expected<void, Error>::Ok();
}

// First `<stem><N>.xml` with no part.
std::string unused_path(const Workbook& wb, const std::string& stem) {
  for (std::uint32_t n = 1;; ++n) {
    std::string path = stem + std::to_string(n) + ".xml";
    if (io::ooxml::find_passthrough_part(wb, path) == nullptr) {
      return path;
    }
  }
}

// `rId<N>` past every id in `rels`.
std::string fresh_rid(const std::vector<UnknownRelationship>& rels) {
  std::uint32_t top = 0;
  for (const UnknownRelationship& rel : rels) {
    top = std::max(top, static_cast<std::uint32_t>(
                            std::strtoul(rel.id.c_str() + std::min<std::size_t>(3, rel.id.size()), nullptr, 10)));
  }
  return "rId" + std::to_string(top + 1U);
}

// Appends `entry` to the list element of the extension `skeleton` describes, in the `<extLst>` text `xml`,
// copying the extension from `skeleton` when absent.
void add_ext_entry(std::string& xml, const char* skeleton, const std::string& entry) {
  pugi::xml_document doc;
  pugi::xml_document tmpl;
  doc.load_string(xml.empty() ? "<extLst/>" : xml.c_str());
  tmpl.load_string(skeleton);
  pugi::xml_node root = doc.document_element();
  const pugi::xml_node want = tmpl.document_element();
  pugi::xml_node ext;
  for (const pugi::xml_node& candidate : children_local(root, "ext")) {
    if (std::string_view(candidate.attribute("uri").value()) == want.attribute("uri").value()) {
      ext = candidate;
    }
  }
  if (!ext) {
    ext = root.append_copy(want);
  }
  pugi::xml_node list = child_local(ext, local_name(want.first_child()));
  if (!list) {
    list = ext.append_copy(want.first_child());
  }
  list.append_buffer(entry.data(), entry.size());
  xml = io::raw_xml(root);
}

// Drops the element naming relationship `rid` from the `<extLst>` text `xml`, with the parents it empties.
void drop_ext_entry(std::string& xml, std::string_view rid) {
  pugi::xml_document doc;
  if (xml.empty() || !doc.load_string(xml.c_str())) {
    return;
  }
  pugi::xml_node root = doc.document_element();
  pugi::xml_node doomed =
      root.find_node([rid](const pugi::xml_node& node) { return rid == node.attribute("r:id").value(); });
  while (doomed && doomed.parent() != root) {
    pugi::xml_node parent = doomed.parent();
    parent.remove_child(doomed);
    if (io::has_element_child(parent)) {
      doomed = pugi::xml_node();
    } else {
      doomed = parent;
    }
    if (!doomed) {
      xml = io::raw_xml(root);
      return;
    }
  }
  root.remove_child(doomed);
  xml = io::has_element_child(root) ? io::raw_xml(root) : std::string();
}

// Removes the relationship to `path` from `rels` and returns its id.
std::string detach(std::vector<UnknownRelationship>& rels, std::string_view path) {
  std::string id;
  rels.erase(std::remove_if(rels.begin(), rels.end(),
                            [&](const UnknownRelationship& rel) {
                              const bool hit = rel.target == path;
                              if (hit) {
                                id = rel.id;
                              }
                              return hit;
                            }),
             rels.end());
  return id;
}

// Records a relationship of `type` to `path` in `rels` and its `<tag r:id>` entry in the `<extLst>` text `ext`.
void link_part(std::vector<UnknownRelationship>& rels, std::string& ext, std::string_view type, const std::string& path,
               const char* skeleton, const char* tag) {
  const std::string rid = fresh_rid(rels);
  rels.push_back({rid, std::string(type), path, false});
  add_ext_entry(ext, skeleton, std::string("<") + tag + " r:id=\"" + rid + "\"/>");
}

// True when a slicer part of any sheet still names cache `name`.
bool cache_in_use(const Workbook& wb, std::string_view name) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    pugi::xml_document doc;
    if (strings::starts_with(part.path, "xl/slicers/") && io::ooxml::load_part_dom(wb, part.path, "slicer", doc) &&
        doc.find_node([name](const pugi::xml_node& node) {
          return is_element(node, "slicer") && name == node.attribute("cache").value();
        })) {
      return true;
    }
  }
  return false;
}

std::string lower(std::string_view text) {
  return strings::to_ascii_lower(text);
}

// `base`, then `base<sep>1`, `base<sep>2`... until it is not in `taken` (lower-cased names).
std::string unique_name(const std::string& base, const char* sep, const std::vector<std::string>& taken) {
  const auto used = [&taken](const std::string& name) {
    return std::find(taken.begin(), taken.end(), lower(name)) != taken.end();
  };
  std::string name = base;
  for (std::uint32_t n = 1; used(name); ++n) {
    name = base + sep + std::to_string(n);
  }
  return name;
}

// Lower-cased `name` attributes of every element in the parts under `dir`.
void collect_names(const Workbook& wb, std::string_view dir, std::vector<std::string>& names) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    pugi::xml_document doc;
    if (strings::starts_with(part.path, dir) && io::ooxml::load_part_dom(wb, part.path, "slicer", doc)) {
      doc.find_node([&names](const pugi::xml_node& node) {
        if (node.attribute("name")) {
          names.push_back(lower(node.attribute("name").value()));
        }
        return false;
      });
    }
  }
}

// Creates the part at `path` when absent and stores `doc` in it.
Expected<void, Error> write_part(Workbook& wb, const std::string& path, const char* content_type,
                                 const pugi::xml_document& doc) {
  if (io::ooxml::find_passthrough_part(wb, path) == nullptr) {
    RETURN_IF_ERROR(wb.add_passthrough_part(PassthroughPart(path, content_type, {})));
  }
  return io::ooxml::store_part_dom(wb, path, doc);
}

void ensure_attr(pugi::xml_node node, const char* name, const char* value) {
  if (!node.attribute(name)) {
    node.append_attribute(name).set_value(value);
  }
}

// Anchor of a new slicer's frame with its fallback shape; the id, names and markers are set after parsing.
constexpr const char* kFrameXml =
    "<xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row>"
    "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row>"
    "<xdr:rowOff>0</xdr:rowOff></xdr:to><mc:AlternateContent "
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice "
    "xmlns:sle15=\"http://schemas.microsoft.com/office/drawing/2012/slicer\" Requires=\"sle15\"><xdr:graphicFrame "
    "macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"0\" name=\"\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>"
    "<xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData "
    "uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"><sle:slicer "
    "xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" "
    "name=\"\"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro=\"\" "
    "textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"0\" name=\"\"/><xdr:cNvSpPr><a:spLocks "
    "noTextEdit=\"1\"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" "
    "cy=\"0\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr "
    "val=\"white\"/></a:solidFill><a:ln w=\"1\"><a:solidFill><a:prstClr val=\"green\"/></a:solidFill></a:ln>"
    "</xdr:spPr><xdr:txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"/><a:lstStyle/><a:p><a:r><a:rPr "
    "lang=\"en-US\" sz=\"1100\"/><a:t>This shape represents a slicer. Slicers need an application that supports "
    "them.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent><xdr:clientData/>"
    "</xdr:twoCellAnchor>";

}  // namespace

Expected<std::vector<SlicerInfo>, Error> list_slicers(const Workbook& wb, std::size_t sheet) {
  std::vector<SlicerInfo> infos;
  auto walked = scan(wb, sheet, 0, nullptr, &infos);
  if (!walked) {
    return walked.error();
  }
  return infos;
}

Expected<std::uint32_t, Error> add_slicer(Workbook& wb, std::size_t sheet, std::size_t table_index, std::size_t column,
                                          const SlicerSpec& request) {
  using namespace slicer_field;
  SlicerSpec spec = request;
  RETURN_IF_ERROR(check_spec(spec));
  if (sheet >= wb.sheet_count() || wb.sheet(sheet).is_opaque_ooxml_sheet() || table_index >= wb.tables().size() ||
      column >= wb.tables()[table_index].columns.size() || (spec.fields & kAnchor) == 0U) {
    return invalid("slicer: bad sheet, table or column, or no anchor");
  }
  const std::vector<UnknownRelationship>& sheet_rels = wb.sheet(sheet).unknown_relationships();
  if (std::any_of(sheet_rels.begin(), sheet_rels.end(),
                  [](const UnknownRelationship& rel) { return rel.type == io::kRelDrawing; })) {
    return unparseable("slicer: drawing retained from an xlsb package is not editable");
  }
  const TableMetadata& table = wb.tables()[table_index];
  const TableColumn& col = table.columns[column];

  std::vector<std::string> slicer_names;
  std::vector<std::string> cache_names;
  collect_names(wb, "xl/slicers/", slicer_names);
  collect_names(wb, "xl/slicerCaches/", cache_names);
  for (const DefinedName& defined : wb.defined_names()) {
    cache_names.push_back(lower(defined.name));
  }
  const std::string name = unique_name((spec.fields & kName) != 0U ? spec.name : col.name, " ", slicer_names);
  std::string cache_base = "Slicer_";
  for (const char c : col.name) {
    const auto byte = static_cast<unsigned char>(c);
    cache_base.push_back(std::isalnum(byte) != 0 || c == '_' || c == '.' || byte >= 0x80U ? c : '_');
  }
  const std::string cache_name = unique_name(cache_base, "", cache_names);
  if ((spec.fields & kCaption) == 0U) {
    spec.caption = col.name;
  }
  if ((spec.fields & kRowHeight) == 0U) {
    spec.row_height = kDefaultRowHeight;
  }
  spec.fields |= kCaption | kRowHeight;

  pugi::xml_document drawing;
  std::string drawing_path = wb.sheet(sheet).drawing_rel_target();
  const bool fresh_drawing = drawing_path.empty();
  if (fresh_drawing) {
    drawing_path = unused_path(wb, "xl/drawings/drawing");
    drawing.load_string(kDrawingRoot);
  } else {
    RETURN_IF_ERROR(load_drawing(wb, sheet, drawing));
  }
  std::uint32_t frame_id = 1;
  drawing.find_node([&frame_id](const pugi::xml_node& node) {
    frame_id = is_element(node, "cNvPr") ? std::max(frame_id, node.attribute("id").as_uint()) : frame_id;
    return false;
  });
  if (frame_id == std::numeric_limits<std::uint32_t>::max()) {
    return invalid("slicer: drawing object id space exhausted");
  }
  ++frame_id;
  pugi::xml_node drawing_root = drawing.document_element();
  ensure_attr(drawing_root, "xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
  ensure_attr(drawing_root, "xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
  if (!drawing_root.append_buffer(kFrameXml, std::char_traits<char>::length(kFrameXml))) {
    return invalid("slicer: frame could not be built");
  }
  const pugi::xml_node frame = drawing_root.last_child();
  const pugi::xml_node cnvpr = find_element(frame, "cNvPr");
  put_attr(cnvpr, "id", std::to_string(frame_id));
  put_attr(cnvpr, "name", name);
  put_attr(find_element(frame, "slicer"), "name", name);
  apply_anchor(wb, sheet, frame, spec);

  std::string slicers_path;
  for (const UnknownRelationship& rel : sheet_rels) {
    if (rel.type == kRelSlicer && !rel.target_external && io::ooxml::find_passthrough_part(wb, rel.target) != nullptr) {
      slicers_path = rel.target;
      break;
    }
  }
  const bool fresh_slicers = slicers_path.empty();
  pugi::xml_document slicers;
  if (fresh_slicers) {
    slicers_path = unused_path(wb, "xl/slicers/slicer");
    slicers.load_string(kSlicersRoot);
  } else {
    RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, slicers_path, "slicer", slicers));
  }
  const std::string_view root_name = slicers.document_element().name();
  const std::string prefix(
      root_name.substr(0, root_name.find(':') == std::string_view::npos ? 0 : root_name.find(':') + 1));
  pugi::xml_node node = slicers.document_element().append_child((prefix + "slicer").c_str());
  put_attr(node, "name", name);
  put_attr(node, "cache", cache_name);
  apply_slicer(node, spec);

  pugi::xml_document cache;
  cache.load_string(kCacheRoot);
  put_attr(cache.document_element(), "name", cache_name);
  put_attr(cache.document_element(), "sourceName", col.name);
  const pugi::xml_node table_cache = table_cache_of(cache);
  put_attr(table_cache, "tableId", std::to_string(table.id));
  put_attr(table_cache, "column", std::to_string(col.id));
  apply_cache(table_cache, spec);

  const std::string cache_path = unused_path(wb, "xl/slicerCaches/slicerCache");
  RETURN_IF_ERROR(write_part(wb, cache_path, "application/vnd.ms-excel.slicerCache+xml", cache));
  RETURN_IF_ERROR(write_part(wb, slicers_path, "application/vnd.ms-excel.slicer+xml", slicers));
  RETURN_IF_ERROR(write_part(wb, drawing_path, "application/vnd.openxmlformats-officedocument.drawing+xml", drawing));
  if (fresh_drawing) {
    wb.sheet(sheet).set_drawing_rel_target(drawing_path);
  }
  if (fresh_slicers) {
    std::vector<UnknownRelationship> rels = wb.sheet(sheet).unknown_relationships();
    std::string ext = wb.sheet(sheet).ext_lst_xml();
    link_part(rels, ext, kRelSlicer, slicers_path, kSheetExt, "x14:slicer");
    wb.sheet(sheet).set_unknown_relationships(std::move(rels));
    wb.sheet(sheet).set_ext_lst_xml(std::move(ext));
  }
  std::vector<UnknownRelationship> rels = wb.unknown_workbook_rels();
  std::string ext = wb.workbook_ext_lst_xml();
  link_part(rels, ext, kRelCache, cache_path, kBookExt, "x14:slicerCache");
  wb.set_unknown_workbook_rels(std::move(rels));
  wb.set_workbook_ext_lst_xml(std::move(ext));
  RETURN_IF_ERROR(wb.set_defined_name(cache_name, "#N/A"));
  return frame_id;
}

Expected<void, Error> update_slicer(Workbook& wb, std::size_t sheet, std::size_t index, const SlicerSpec& spec) {
  using namespace slicer_field;
  RETURN_IF_ERROR(check_spec(spec));
  Slot slot;
  RETURN_IF_ERROR(slot_at(wb, sheet, index, slot));
  if ((spec.fields & kName) != 0U) {
    return invalid("slicer: a slicer cannot be renamed");
  }
  pugi::xml_document drawing;
  RETURN_IF_ERROR(load_drawing(wb, sheet, drawing));
  const pugi::xml_node frame = find_frame(drawing, slot.info.name);
  if ((spec.fields & kAnchor) != 0U && (!frame || !is_element(frame, "twoCellAnchor"))) {
    return invalid("slicer: the slicer has no two-cell frame to move");
  }
  if ((spec.fields & (kCaption | kStyle | kColumnCount | kRowHeight)) != 0U) {
    pugi::xml_document doc;
    ASSIGN_OR_RETURN(pugi::xml_node node, load_slicer_node(wb, slot, doc));
    apply_slicer(node, spec);
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, slot.slicers_part, doc));
  }
  if ((spec.fields & (kSortOrder | kCrossFilter)) != 0U) {
    pugi::xml_document doc;
    RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, slot.cache_part, "slicer", doc));
    apply_cache(table_cache_of(doc), spec);
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, slot.cache_part, doc));
  }
  if ((spec.fields & kAnchor) != 0U) {
    apply_anchor(wb, sheet, frame, spec);
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, wb.sheet(sheet).drawing_rel_target(), drawing));
  }
  return Expected<void, Error>::Ok();
}

namespace {

// Removes the part at `path` together with the relationship and extension entry that reach it, from the sheet
// (`sheet` given) or the workbook, and then the part itself.
void drop_linked_part(Workbook& wb, const std::size_t* sheet, const std::string& path) {
  std::vector<UnknownRelationship> rels =
      sheet != nullptr ? wb.sheet(*sheet).unknown_relationships() : wb.unknown_workbook_rels();
  std::string ext = sheet != nullptr ? wb.sheet(*sheet).ext_lst_xml() : wb.workbook_ext_lst_xml();
  drop_ext_entry(ext, detach(rels, path));
  if (sheet != nullptr) {
    wb.sheet(*sheet).set_unknown_relationships(std::move(rels));
    wb.sheet(*sheet).set_ext_lst_xml(std::move(ext));
  } else {
    wb.set_unknown_workbook_rels(std::move(rels));
    wb.set_workbook_ext_lst_xml(std::move(ext));
  }
  io::ooxml::remove_orphans(wb, {path});
}

// Removes slicer `index`; the `Slicer_*` names of caches that went with it are appended to `dropped`.
Expected<void, Error> remove_slicer_at(Workbook& wb, std::size_t sheet, std::size_t index,
                                       std::vector<std::string>& dropped) {
  Slot slot;
  RETURN_IF_ERROR(slot_at(wb, sheet, index, slot));
  pugi::xml_document drawing;
  RETURN_IF_ERROR(load_drawing(wb, sheet, drawing));
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(pugi::xml_node node, load_slicer_node(wb, slot, doc));
  if (pugi::xml_node top = find_frame(drawing, slot.info.name)) {
    pugi::xml_node root = drawing.document_element();
    while (top.parent() != root) {
      top = top.parent();
    }
    root.remove_child(top);
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, wb.sheet(sheet).drawing_rel_target(), drawing));
  }
  const bool drawing_empty = drawing.document_element() && !io::has_element_child(drawing.document_element());
  doc.document_element().remove_child(node);
  if (children_local(doc.document_element(), "slicer").empty()) {
    drop_linked_part(wb, &sheet, slot.slicers_part);
  } else {
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, slot.slicers_part, doc));
  }
  if (drawing_empty) {
    const std::string drawing_path = wb.sheet(sheet).drawing_rel_target();
    wb.sheet(sheet).set_drawing_rel_target({});
    io::ooxml::remove_orphans(wb, {drawing_path});
  }
  if (!cache_in_use(wb, slot.info.cache_name)) {
    dropped.push_back(slot.info.cache_name);
    drop_linked_part(wb, nullptr, slot.cache_part);
  }
  return Expected<void, Error>::Ok();
}

// Which slicers a cascade removes: those of a table, those that lost the frame they had (`framed` lists the
// slicers that had one), or every slicer of the sheet.
struct Match {
  bool by_table = false;
  std::uint32_t table_id = 0;
  const std::vector<std::string>* framed = nullptr;
};

bool matches(const Match& match, const Slot& slot) {
  if (match.by_table) {
    return slot.table_id == match.table_id;
  }
  return match.framed == nullptr ||
         (slot.info.object.object_id == 0 &&
          std::find(match.framed->begin(), match.framed->end(), slot.info.name) != match.framed->end());
}

void remove_matching(Workbook& wb, std::size_t sheet, const Match& match, std::vector<std::string>& dropped) {
  Slot slot;
  for (std::size_t i = 0;;) {
    auto found = scan(wb, sheet, i, &slot, nullptr);
    if (!found || !found.value()) {
      return;
    }
    if (!matches(match, slot) || !remove_slicer_at(wb, sheet, i, dropped)) {
      ++i;
    }
  }
}

void remove_table_slicers_at(Workbook& wb, std::uint32_t table_id, std::vector<std::string>& dropped) {
  for (std::size_t sheet = 0; sheet < wb.sheet_count(); ++sheet) {
    remove_matching(wb, sheet, Match{true, table_id, nullptr}, dropped);
  }
  for (const UnknownRelationship& rel : std::vector<UnknownRelationship>(wb.unknown_workbook_rels())) {
    pugi::xml_document doc;
    if (rel.type == kRelCache && io::ooxml::load_part_dom(wb, rel.target, "slicer", doc) &&
        table_cache_of(doc).attribute("tableId").as_uint() == table_id) {
      const std::string name = doc.document_element().attribute("name").value();
      if (!cache_in_use(wb, name)) {
        dropped.push_back(name);
        drop_linked_part(wb, nullptr, rel.target);
      }
    }
  }
}

}  // namespace

Expected<void, Error> remove_slicer(Workbook& wb, std::size_t sheet, std::size_t index) {
  std::vector<std::string> dropped;
  RETURN_IF_ERROR(remove_slicer_at(wb, sheet, index, dropped));
  for (const std::string& name : dropped) {
    static_cast<void>(wb.set_defined_name(name, ""));
  }
  return Expected<void, Error>::Ok();
}

void remove_table_slicers(Workbook& wb, std::uint32_t table_id) {
  std::vector<std::string> dropped;
  remove_table_slicers_at(wb, table_id, dropped);
  for (const std::string& name : dropped) {
    static_cast<void>(wb.set_defined_name(name, ""));
  }
}

std::vector<std::string> remove_sheet_slicers(Workbook& wb, std::size_t sheet) {
  std::vector<std::string> dropped;
  for (const TableMetadata& table : wb.tables()) {
    if (table.sheet_index == sheet) {
      remove_table_slicers_at(wb, table.id, dropped);
    }
  }
  remove_matching(wb, sheet, Match{}, dropped);
  return dropped;
}

std::vector<std::string> framed_slicers(const Workbook& wb, std::size_t sheet) {
  std::vector<std::string> names;
  auto infos = list_slicers(wb, sheet);
  for (const SlicerInfo& info : infos ? infos.value() : std::vector<SlicerInfo>()) {
    if (info.object.object_id != 0) {
      names.push_back(info.name);
    }
  }
  return names;
}

std::vector<std::string> remove_unframed_slicers(Workbook& wb, std::size_t sheet,
                                                 const std::vector<std::string>& framed) {
  std::vector<std::string> dropped;
  if (!framed.empty()) {
    remove_matching(wb, sheet, Match{false, 0, &framed}, dropped);
  }
  return dropped;
}

}  // namespace formulon
