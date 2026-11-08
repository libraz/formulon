#include "pivot/field_patch.h"

#include <cstddef>
#include <iterator>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "eval/collation.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "utils/status_macros.h"
#include "utils/utf8_length.h"
#include "workbook.h"

namespace formulon::pivot {
namespace {

Error invalid(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
}

Error invalid(const char* message, std::string context) {
  return make_error(FormulonErrorCode::kInvalidArgument, message, std::move(context));
}

Expected<void, Error> ok() {
  return Expected<void, Error>::Ok();
}

// True when `a` and `b` are one item under the profile's pivot collation.
bool same_label(const Workbook& wb, std::string_view a, std::string_view b) {
  return eval::collate(a, b, wb.excel_profile(), eval::CollationStrength::kDynamic) == 0;
}

// The label `item` takes from its value alone: its own name, else the bound cache value's text.
std::string value_label(const Workbook& wb, const PivotTable& table, std::size_t field_index, const PivotItem& item) {
  if (const PivotCache* cache = wb.find_pivot_cache(table.pivot_cache_id()); cache != nullptr) {
    return pivot_item_label(*cache, field_index, item);
  }
  return item.name;
}

// True when a caption equal to the item's own label removes the caption instead.
bool caption_clears(const Workbook& wb, const PivotTable& table, std::size_t field_index, const PivotItem& item,
                    std::string_view caption) {
  return caption.empty() || same_label(wb, caption, value_label(wb, table, field_index, item));
}

}  // namespace

bool field_show_detail(const PivotField& field) noexcept {
  for (const PivotItem& item : field.items) {
    if (!item.show_detail) {
      return false;
    }
  }
  for (const PivotMarkerItem& marker : field.marker_items) {
    if (marker.type == "default" && !marker.show_detail) {
      return false;
    }
  }
  return true;
}

Expected<void, Error> check_field_patch(const Workbook& wb, const PivotTable& table, std::size_t field_index,
                                        const PivotFieldPatch& patch) {
  if (field_index >= table.fields().size()) {
    return invalid("pivot: field index out of range");
  }
  const PivotField& field = table.fields()[field_index];
  if (patch.custom_name.has_value() && !patch.custom_name->empty()) {
    const std::string& name = *patch.custom_name;
    if (utf16_units_in(name) > kMaxPivotNameUnits) {
      return invalid("pivot: name longer than 255 UTF-16 code units");
    }
    const auto taken = [&](const std::string& existing) { return !existing.empty() && same_label(wb, name, existing); };
    for (std::size_t i = 0; i < table.fields().size(); ++i) {
      const PivotField& other = table.fields()[i];
      if (i != field_index && (taken(other.source_name) || taken(other.custom_name))) {
        return invalid("pivot: name matches another field", "field_index=" + std::to_string(i));
      }
    }
    for (std::size_t i = 0; i < table.data_fields().size(); ++i) {
      if (taken(table.data_fields()[i].name)) {
        return invalid("pivot: name matches a data field", "data_field_index=" + std::to_string(i));
      }
    }
  }
  const bool manual = patch.sort_order.has_value() ? *patch.sort_order == FieldSortOrder::kManual : field.sort.manual;
  const std::string& by_field = patch.sort_by_field.has_value() ? *patch.sort_by_field : field.sort.by_field;
  if (manual && !by_field.empty()) {
    return invalid("pivot: manual order takes no by-field");
  }
  return ok();
}

Expected<void, Error> apply_field_patch(const Workbook& wb, PivotTable& table, std::size_t field_index,
                                        const PivotFieldPatch& patch) {
  RETURN_IF_ERROR(check_field_patch(wb, table, field_index, patch));
  PivotField& field = table.mutable_fields()[field_index];
  if (patch.custom_name.has_value()) {
    field.custom_name = *patch.custom_name;
  }
  if (patch.number_format.has_value()) {
    field.number_format = *patch.number_format;
  }
  const std::optional<bool> PivotFieldPatch::*const flag_keys[] = {
      &PivotFieldPatch::default_subtotal, &PivotFieldPatch::insert_blank_row, &PivotFieldPatch::fill_down_labels,
      &PivotFieldPatch::show_all, &PivotFieldPatch::subtotal_top};
  bool PivotField::*const flag_fields[] = {&PivotField::default_subtotal, &PivotField::insert_blank_row,
                                           &PivotField::fill_down_labels, &PivotField::show_all,
                                           &PivotField::subtotal_top};
  for (std::size_t i = 0; i < std::size(flag_keys); ++i) {
    if (const std::optional<bool>& v = patch.*flag_keys[i]; v.has_value()) {
      field.*flag_fields[i] = *v;
    }
  }
  if (patch.sort_order.has_value()) {
    field.sort.manual = *patch.sort_order == FieldSortOrder::kManual;
    // Manual keeps the ascending form the OOXML reader yields for it.
    field.sort.ascending = *patch.sort_order != FieldSortOrder::kDescending;
  }
  if (patch.sort_by_field.has_value()) {
    field.sort.by_field = *patch.sort_by_field;
  }
  if (patch.sort_order.has_value() || patch.sort_by_field.has_value()) {
    field.sort.stated = true;
  }
  if (patch.show_detail.has_value()) {
    for (PivotItem& item : field.items) {
      item.show_detail = *patch.show_detail;
    }
    for (PivotMarkerItem& marker : field.marker_items) {
      if (marker.type == "default") {
        marker.show_detail = *patch.show_detail;
      }
    }
  }
  return ok();
}

Expected<void, Error> check_item_patch(const Workbook& wb, const PivotTable& table, std::size_t field_index,
                                       std::size_t item_index, const PivotItemPatch& patch) {
  if (field_index >= table.fields().size()) {
    return invalid("pivot: field index out of range");
  }
  const PivotField& field = table.fields()[field_index];
  if (item_index >= field.items.size()) {
    return invalid("pivot: item index out of range");
  }
  if (!patch.caption.has_value()) {
    return ok();
  }
  const std::string& caption = *patch.caption;
  if (utf16_units_in(caption) > kMaxPivotNameUnits) {
    return invalid("pivot: caption longer than 255 UTF-16 code units");
  }
  if (caption_clears(wb, table, field_index, field.items[item_index], caption)) {
    return ok();
  }
  for (std::size_t i = 0; i < field.items.size(); ++i) {
    const PivotItem& other = field.items[i];
    const std::string shown = !other.caption.empty() ? other.caption : value_label(wb, table, field_index, other);
    if (i != item_index && same_label(wb, caption, shown)) {
      return invalid("pivot: caption matches another item of the field", "item_index=" + std::to_string(i));
    }
  }
  return ok();
}

Expected<void, Error> apply_item_patch(const Workbook& wb, PivotTable& table, std::size_t field_index,
                                       std::size_t item_index, const PivotItemPatch& patch) {
  RETURN_IF_ERROR(check_item_patch(wb, table, field_index, item_index, patch));
  PivotItem& item = table.mutable_fields()[field_index].items[item_index];
  if (patch.visible.has_value()) {
    item.visible = *patch.visible;
  }
  if (patch.show_detail.has_value()) {
    item.show_detail = *patch.show_detail;
  }
  if (patch.caption.has_value()) {
    item.caption = caption_clears(wb, table, field_index, item, *patch.caption) ? std::string() : *patch.caption;
  }
  return ok();
}

}  // namespace formulon::pivot
