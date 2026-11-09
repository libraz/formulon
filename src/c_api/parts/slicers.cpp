//
// C ABI - the table slicers of a worksheet: reading, adding and updating them.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/drawing_anchor.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_geometry.h"
#include "drawing/slicer_edit.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::AnchorPlan;
using formulon::c_api::parts::apply_placement_extras;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::drawing_patch_from_c;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::plan_add_anchor;
using formulon::c_api::parts::resolve_sheet_table;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::slicer_sort_order_from_c;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::to_c;

namespace {

bool is_set(fm_patch_op_t op) {
  return op == FM_PATCH_OP_SET;
}

// Converts `p` to the model spec. On add every field may be CLEAR (the default); on update only `style_name`.
fm_status_t to_spec(const fm_slicer_patch_t& p, bool add, formulon::SlicerSpec& spec) {
  namespace field = formulon::slicer_field;
  const std::pair<fm_patch_op_t, const char*> ops[] = {
      {p.column_count_op, "column_count"},   {p.sort_order_op, "sort_order"}, {p.cross_filter_op, "cross_filter"},
      {p.row_height_pt_op, "row_height_pt"}, {p.caption_op, "caption"},       {p.style_name_op, "style_name"}};
  for (const auto& [op, name] : ops) {
    const bool clearable = add || std::string_view(name) == "style_name";
    if (!(clearable ? patch_op_ok(op) : patch_op_ok_without_clear(op))) {
      return invalid_argument(patch_op_ok(op) ? "field cannot be cleared" : "op must be KEEP, CLEAR or SET",
                              std::string("field=") + name);
    }
  }
  const std::pair<fm_patch_op_t, const char*> texts[] = {{p.caption_op, p.caption}, {p.style_name_op, p.style_name}};
  const char* const text_names[] = {"caption", "style_name"};
  for (std::size_t i = 0; i < 2; ++i) {
    if (!is_set(texts[i].first)) {
      continue;
    }
    if (texts[i].second == nullptr) {
      return invalid_argument("a SET string cannot be NULL", std::string("field=") + text_names[i]);
    }
    if (auto rc = check_utf8(texts[i].second, "field", text_names[i]); rc != 0) {
      return rc;
    }
  }
  // No style is spelled CLEAR (or NULL on add), never an empty name.
  if (is_set(p.style_name_op) && p.style_name[0] == '\0') {
    return invalid_argument("a style name is not empty", "field=style_name");
  }

  spec.fields = 0;
  if (is_set(p.caption_op)) {
    spec.fields |= field::kCaption;
    spec.caption = p.caption;
  }
  if (is_set(p.style_name_op) || (!add && p.style_name_op == FM_PATCH_OP_CLEAR)) {
    spec.fields |= field::kStyle;
    spec.style = is_set(p.style_name_op) ? std::string(p.style_name) : std::string();
  }
  if (is_set(p.column_count_op)) {
    spec.fields |= field::kColumnCount;
    spec.column_count = p.column_count;
  }
  if (is_set(p.row_height_pt_op)) {
    if (!std::isfinite(p.row_height_pt) || p.row_height_pt < 0.0) {
      return invalid_argument("row height is negative or not finite", "field=row_height_pt");
    }
    spec.fields |= field::kRowHeight;
    spec.row_height = std::llround(p.row_height_pt * static_cast<double>(formulon::kEmuPerPoint));
  }
  if (is_set(p.sort_order_op)) {
    if (auto rc = check_enum_domain(p.sort_order, FM_SLICER_SORT_ORDER_ASCENDING, FM_SLICER_SORT_ORDER_DESCENDING,
                                    "sort_order");
        rc != 0) {
      return rc;
    }
    spec.fields |= field::kSortOrder;
    spec.sort_order = slicer_sort_order_from_c(p.sort_order);
  }
  if (is_set(p.cross_filter_op)) {
    if (auto rc = check_enum_domain(p.cross_filter, 0, FM_SLICER_CROSS_FILTER_NONE, "cross_filter"); rc != 0) {
      return rc;
    }
    spec.fields |= field::kCrossFilter;
    spec.cross_filter = static_cast<formulon::SlicerCrossFilter>(p.cross_filter);
  }
  return 0;
}

// Finds the position of the slicer whose frame id is `object_id` in `list_slicers`. A non-zero
// return is the status to report; `*found` is false when no slicer carries the id.
fm_status_t find_slicer(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t* index, bool* found) {
  auto slicers = formulon::list_slicers(wb->workbook(), sheet_index);
  if (!slicers) {
    return set_last_error(slicers.error());
  }
  const std::vector<formulon::SlicerInfo>& list = slicers.value();
  *found = false;
  for (size_t i = 0; object_id != 0U && i < list.size(); ++i) {
    if (list[i].object.object_id == object_id) {
      *index = i;
      *found = true;
      break;
    }
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_slicer(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                           int32_t* out_present, fm_slicer_info_t* out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_slicer_info_t{};
  out->struct_size = struct_size;
  size_t index = 0;
  bool found = false;
  if (auto rc = find_slicer(wb, sheet_index, object_id, &index, &found); rc != 0 || !found) {
    return rc;
  }
  auto slicers = formulon::list_slicers(wb->workbook(), sheet_index);
  if (!slicers) {
    return set_last_error(slicers.error());
  }
  const formulon::SlicerInfo& info = slicers.value()[index];
  TextStore& store = fresh_read_scratch(wb);
  out->object_id = info.object.object_id;
  out->column_count = info.column_count;
  out->sort_order = to_c(info.sort_order);
  out->cross_filter = static_cast<int32_t>(info.cross_filter);
  out->row_height_pt = static_cast<double>(info.row_height) / static_cast<double>(formulon::kEmuPerPoint);
  out->caption = store_cstr(store, info.caption);
  out->cache_name = store_cstr(store, info.cache_name);
  out->table_name = store_cstr(store, info.table_name);
  out->column_name = store_cstr(store, info.column_name);
  out->style_name = info.style.empty() ? nullptr : store_cstr(store, info.style);
  *out_present = 1;
  return 0;
}

extern "C" fm_status_t fm_sheet_add_slicer(fm_workbook_t* wb, size_t sheet_index, size_t table_sheet_index,
                                           size_t table_index, size_t column_index, const fm_slicer_spec_t* spec,
                                           uint32_t* out_object_id) {
  clear_last_error();
  if (spec == nullptr || out_object_id == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*spec); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(spec->placement); rc != 0) {
    return rc;
  }
  const fm_drawing_placement_spec_t& placement = spec->placement;
  if (auto rc = check_utf8({placement.name, placement.description, spec->caption, spec->style_name}); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(placement.hidden, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "placement.hidden");
      rc != 0) {
    return rc;
  }
  // The engine addresses a table by its workbook-wide position.
  size_t global = 0;
  if (auto rc = resolve_sheet_table(wb, table_sheet_index, table_index, &global); rc != 0) {
    return rc;
  }
  AnchorPlan plan;
  if (auto rc =
          plan_add_anchor(wb->workbook(), sheet_index, placement.anchor, formulon::AnchorKind::kTwoCell, true, &plan);
      rc != 0) {
    return rc;
  }
  if (plan.kind != formulon::AnchorKind::kTwoCell) {
    return invalid_argument("a slicer frame is two-cell", "field=placement.anchor.kind");
  }
  // A zero field takes the default; the sort order and cross filter are plain values.
  fm_slicer_patch_t fields{};
  fields.struct_size = sizeof(fields);
  if (spec->column_count != 0U) {
    fields.column_count_op = FM_PATCH_OP_SET;
    fields.column_count = spec->column_count;
  }
  fields.sort_order_op = fields.cross_filter_op = FM_PATCH_OP_SET;
  fields.sort_order = spec->sort_order;
  fields.cross_filter = spec->cross_filter;
  // A zero height is Excel's default; anything else, a negative or NaN included, goes to the range check.
  if (!(spec->row_height_pt == 0.0)) {
    fields.row_height_pt_op = FM_PATCH_OP_SET;
    fields.row_height_pt = spec->row_height_pt;
  }
  if (spec->caption != nullptr && spec->caption[0] != '\0') {
    fields.caption_op = FM_PATCH_OP_SET;
    fields.caption = spec->caption;
  }
  if (spec->style_name != nullptr) {
    fields.style_name_op = FM_PATCH_OP_SET;
    fields.style_name = spec->style_name;
  }
  formulon::SlicerSpec model;
  if (auto rc = to_spec(fields, true, model); rc != 0) {
    return rc;
  }
  namespace field = formulon::slicer_field;
  model.fields |= field::kAnchor;
  model.from = plan.from;
  model.to = plan.to;
  if (placement.name != nullptr && placement.name[0] != '\0') {
    model.fields |= field::kName;
    model.name = placement.name;
  }
  auto id = formulon::add_slicer(wb->workbook(), sheet_index, global, column_index, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return apply_placement_extras(wb->workbook(), sheet_index, id.value(), placement, true);
}

extern "C" fm_status_t fm_sheet_update_slicer(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                              const fm_drawing_patch_t* common, const fm_slicer_patch_t* slicer) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::DrawingCommonPatch common_model;
  if (common != nullptr) {
    if (auto rc = drawing_patch_from_c(*common, &common_model); rc != 0) {
      return rc;
    }
  }
  formulon::SlicerSpec slicer_model;
  if (slicer != nullptr) {
    if (auto rc = check_record_size(*slicer); rc != 0) {
      return rc;
    }
    if (auto rc = to_spec(*slicer, false, slicer_model); rc != 0) {
      return rc;
    }
  }
  return finish_status(formulon::update_slicer(wb->workbook(), sheet_index, object_id,
                                               common != nullptr ? &common_model : nullptr,
                                               slicer != nullptr ? &slicer_model : nullptr));
}
