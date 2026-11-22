#include "drawing/drawing_clone.h"

#include <algorithm>
#include <set>
#include <string>

#include "chart/chart_ref.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "drawing/slicer_edit.h"
#include "io/ooxml/part_dom.h"
#include "pugixml.hpp"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace {

Error unsupported(const char* what) {
  return make_error(FormulonErrorCode::kNotImplemented, "Drawing content cannot be copied", what);
}

bool is_pivot_chart(const Workbook& wb, const std::string& path) {
  pugi::xml_document doc;
  if (!io::ooxml::load_part_dom(wb, path, "drawing_clone", doc)) {
    return false;
  }
  return static_cast<bool>(doc.find_node([](const pugi::xml_node& n) { return local_name(n) == "pivotSource"; }));
}

}  // namespace

Expected<DrawingCopyPlan, Error> plan_drawing_copy(const Workbook& wb, std::size_t source_sheet) {
  DrawingCopyPlan plan;
  auto path = editable_drawing_path(wb, source_sheet);
  if (!path) {
    if (path.error().code == FormulonErrorCode::kIoDrawingUnparseable) {
      return unsupported("drawing retained from an xlsb package");
    }
    return path.error();
  }
  if (path.value().empty()) {
    return plan;
  }
  if (!framed_slicers(wb, source_sheet).empty()) {
    return unsupported("slicer");
  }

  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing_part(wb, path.value(), doc));
  ASSIGN_OR_RETURN(auto rels, load_part_rels(wb, path.value()));
  for (const DrawingObject& obj : read_drawing_objects(doc, rels)) {
    if (obj.kind != DrawingObjectKind::kPicture && obj.kind != DrawingObjectKind::kChart) {
      return unsupported("shape, group, connector, text box or other frame");
    }
  }

  ASSIGN_OR_RETURN(auto charts, list_chart_parts(wb, source_sheet));
  for (const ChartPartRef& chart : charts) {
    if (!chart.part_path.empty() && is_pivot_chart(wb, chart.part_path)) {
      return unsupported("pivot chart");
    }
  }

  ASSIGN_OR_RETURN(auto objects, list_drawing_objects(wb, source_sheet));
  for (const DrawingObject& obj : objects) {
    ASSIGN_OR_RETURN(auto bytes, snapshot_drawing(wb, source_sheet, obj.object_id));
    plan.snapshots.push_back(std::move(bytes));
  }
  return plan;
}

Expected<void, Error> apply_drawing_copy(Workbook& wb, std::size_t target_sheet, const DrawingCopyPlan& plan,
                                         const parser::RefTransform& transform) {
  if (plan.snapshots.empty()) {
    return Expected<void, Error>::Ok();
  }
  std::set<std::string> before;
  {
    ASSIGN_OR_RETURN(auto existing, list_chart_parts(wb, target_sheet));
    for (const ChartPartRef& ref : existing) {
      before.insert(ref.part_path);
    }
  }
  for (const std::vector<std::uint8_t>& bytes : plan.snapshots) {
    auto id = restore_drawing(wb, target_sheet, bytes.data(), bytes.size(), kDrawingRestoreNewId);
    if (!id) {
      return id.error();
    }
  }
  ASSIGN_OR_RETURN(auto added, list_chart_parts(wb, target_sheet));
  for (const ChartPartRef& ref : added) {
    if (ref.part_path.empty() || before.count(ref.part_path) != 0U) {
      continue;
    }
    pugi::xml_document doc;
    if (!io::ooxml::load_part_dom(wb, ref.part_path, "drawing_clone", doc)) {
      continue;
    }
    if (chart_ref::rewrite_sites(doc, transform) != 0U) {
      RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, ref.part_path, doc));
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
