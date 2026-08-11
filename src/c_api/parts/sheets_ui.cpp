//
// C ABI - sheet UI surface: hyperlinks, merges, comments, data
// validations, and sheet protection.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_range_count;
using formulon::c_api::parts::check_sheet_u32;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;

// ---------------------------------------------------------------------------
// Hyperlinks
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_add_hyperlink(fm_workbook_t* wb, std::uint32_t sheet, fm_hyperlink hl) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_add_hyperlink"); rc != 0) {
    return rc;
  }
  formulon::Hyperlink out;
  out.row = hl.row;
  out.col = hl.col;
  out.target = (hl.target != nullptr) ? std::string(hl.target) : std::string();
  out.location = (hl.location != nullptr) ? std::string(hl.location) : std::string();
  out.display = (hl.display != nullptr) ? std::string(hl.display) : std::string();
  out.tooltip = (hl.tooltip != nullptr) ? std::string(hl.tooltip) : std::string();
  // rid stays empty; the writer mints a fresh rIdN on save.
  wb->workbook().sheet(sheet).mutable_hyperlinks().push_back(std::move(out));
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_hyperlink(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                                 std::uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_hyperlink"); rc != 0) {
    return rc;
  }
  auto& hls = wb->workbook().sheet(sheet).mutable_hyperlinks();
  hls.erase(std::remove_if(hls.begin(), hls.end(),
                           [&](const formulon::Hyperlink& h) { return h.row == row && h.col == col; }),
            hls.end());
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_hyperlink_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_hyperlink_at"); rc != 0) {
    return rc;
  }
  auto& hls = wb->workbook().sheet(sheet).mutable_hyperlinks();
  if (static_cast<std::size_t>(index) >= hls.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_remove_hyperlink_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(hls.size()));
  }
  hls.erase(hls.begin() + static_cast<std::ptrdiff_t>(index));
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_hyperlinks(fm_workbook_t* wb, std::uint32_t sheet) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_clear_hyperlinks"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet).mutable_hyperlinks().clear();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_hyperlink_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index,
                                                 fm_hyperlink* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_hyperlink_at: out is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_hyperlink_at"); rc != 0) {
    return rc;
  }
  const auto& hls = wb->workbook().sheet(sheet).hyperlinks();
  if (static_cast<std::size_t>(index) >= hls.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_get_hyperlink_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(hls.size()));
  }
  const formulon::Hyperlink& h = hls[index];
  out->row = h.row;
  out->col = h.col;
  out->target = h.target.c_str();
  out->location = h.location.c_str();
  out->display = h.display.c_str();
  out->tooltip = h.tooltip.c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_hyperlink_count(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_hyperlink_count: out_count is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_hyperlink_count"); rc != 0) {
    return rc;
  }
  *out_count = static_cast<std::uint32_t>(wb->workbook().sheet(sheet).hyperlinks().size());
  return 0;
}

// ---------------------------------------------------------------------------
// Merges
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_add_merge(fm_workbook_t* wb, std::uint32_t sheet, fm_merge_range merge) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_add_merge"); rc != 0) {
    return rc;
  }
  // Normalise corners so first <= last componentwise; mirrors the
  // reader's behaviour and keeps downstream consumers simple.
  formulon::MergeRange m;
  m.first_row = (merge.first_row < merge.last_row) ? merge.first_row : merge.last_row;
  m.first_col = (merge.first_col < merge.last_col) ? merge.first_col : merge.last_col;
  m.last_row = (merge.first_row < merge.last_row) ? merge.last_row : merge.first_row;
  m.last_col = (merge.first_col < merge.last_col) ? merge.last_col : merge.first_col;
  const auto result = wb->workbook().add_merge(sheet, m);
  if (!result) {
    return formulon::c_api::parts::set_last_error(result.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_merge(fm_workbook_t* wb, std::uint32_t sheet, fm_merge_range range) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_merge"); rc != 0) {
    return rc;
  }
  // Normalise corners so first <= last componentwise; mirrors fm_sheet_add_merge.
  formulon::MergeRange q;
  q.first_row = (range.first_row < range.last_row) ? range.first_row : range.last_row;
  q.first_col = (range.first_col < range.last_col) ? range.first_col : range.last_col;
  q.last_row = (range.first_row < range.last_row) ? range.last_row : range.first_row;
  q.last_col = (range.first_col < range.last_col) ? range.last_col : range.first_col;
  const auto result = wb->workbook().remove_merges_intersecting(sheet, q);
  if (!result) {
    return formulon::c_api::parts::set_last_error(result.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_merge_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_merge_at"); rc != 0) {
    return rc;
  }
  const auto result = wb->workbook().remove_merge_at(sheet, index);
  if (!result) {
    return formulon::c_api::parts::set_last_error(result.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_merges(fm_workbook_t* wb, std::uint32_t sheet) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_clear_merges"); rc != 0) {
    return rc;
  }
  const auto result = wb->workbook().clear_merges(sheet);
  if (!result) {
    return formulon::c_api::parts::set_last_error(result.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_merge_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index,
                                             fm_merge_range* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_merge_at: out is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_merge_at"); rc != 0) {
    return rc;
  }
  const auto& merges = wb->workbook().sheet(sheet).merges();
  if (static_cast<std::size_t>(index) >= merges.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_get_merge_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(merges.size()));
  }
  const formulon::MergeRange& m = merges[index];
  out->first_row = m.first_row;
  out->first_col = m.first_col;
  out->last_row = m.last_row;
  out->last_col = m.last_col;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_merge_count(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_merge_count: out_count is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_merge_count"); rc != 0) {
    return rc;
  }
  *out_count = static_cast<std::uint32_t>(wb->workbook().sheet(sheet).merges().size());
  return 0;
}

// ---------------------------------------------------------------------------
// Comments
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_get_comment_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                               std::uint32_t col, fm_comment* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_comment_at: out is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_comment_at"); rc != 0) {
    return rc;
  }
  for (const formulon::CellComment& c : wb->workbook().sheet(sheet).comments()) {
    if (c.row == row && c.col == col) {
      out->row = c.row;
      out->col = c.col;
      out->author = c.author.c_str();
      out->text = c.text.c_str();
      return 0;
    }
  }
  return set_binding_error(formulon::FormulonErrorCode::kNotFound, "fm_sheet_get_comment_at: no comment at cell",
                           "row=" + std::to_string(row) + " col=" + std::to_string(col));
}

extern "C" fm_status_t fm_sheet_get_comment_count(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_comment_count: out_count is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_comment_count"); rc != 0) {
    return rc;
  }
  *out_count = static_cast<std::uint32_t>(wb->workbook().sheet(sheet).comments().size());
  return 0;
}

extern "C" fm_status_t fm_sheet_get_comment_at_index(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index,
                                                     fm_comment* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_comment_at_index: out is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_comment_at_index"); rc != 0) {
    return rc;
  }
  const auto& comments = wb->workbook().sheet(sheet).comments();
  if (static_cast<std::size_t>(index) >= comments.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_get_comment_at_index: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(comments.size()));
  }
  const formulon::CellComment& c = comments[index];
  out->row = c.row;
  out->col = c.col;
  out->author = c.author.c_str();
  out->text = c.text.c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_set_comment(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                            std::uint32_t col, const char* author, const char* text) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_set_comment"); rc != 0) {
    return rc;
  }
  auto& list = wb->workbook().sheet(sheet).mutable_comments();
  // Locate any existing entry first; mutating the list invalidates
  // iterators so we capture the index instead of an iterator.
  std::size_t found = list.size();
  for (std::size_t i = 0; i < list.size(); ++i) {
    if (list[i].row == row && list[i].col == col) {
      found = i;
      break;
    }
  }
  // Empty/NULL text => removal.
  if (text == nullptr || text[0] == '\0') {
    if (found < list.size()) {
      list.erase(list.begin() + static_cast<std::ptrdiff_t>(found));
    }
    return 0;
  }
  formulon::CellComment c;
  c.row = row;
  c.col = col;
  c.author = (author != nullptr) ? std::string(author) : std::string();
  c.text = std::string(text);
  if (found < list.size()) {
    list[found] = std::move(c);
  } else {
    list.push_back(std::move(c));
  }
  return 0;
}

// ---------------------------------------------------------------------------
// Validations
// ---------------------------------------------------------------------------

// `fm_merge_range` and `formulon::MergeRange` must remain layout-compatible
// so `fm_sheet_get_validation_at` can hand out the engine-side
// `std::vector<MergeRange>::data()` directly via `reinterpret_cast`. Both
// types are POD with four `uint32_t` fields in identical order; the static
// asserts below pin that contract so any future change to either struct
// trips a compile-time error.
static_assert(sizeof(fm_merge_range) == sizeof(formulon::MergeRange),
              "fm_merge_range / MergeRange size mismatch breaks validation getter");
static_assert(alignof(fm_merge_range) == alignof(formulon::MergeRange),
              "fm_merge_range / MergeRange alignment mismatch breaks validation getter");
static_assert(offsetof(fm_merge_range, first_row) == offsetof(formulon::MergeRange, first_row),
              "fm_merge_range::first_row layout mismatch");
static_assert(offsetof(fm_merge_range, first_col) == offsetof(formulon::MergeRange, first_col),
              "fm_merge_range::first_col layout mismatch");
static_assert(offsetof(fm_merge_range, last_row) == offsetof(formulon::MergeRange, last_row),
              "fm_merge_range::last_row layout mismatch");
static_assert(offsetof(fm_merge_range, last_col) == offsetof(formulon::MergeRange, last_col),
              "fm_merge_range::last_col layout mismatch");

extern "C" fm_status_t fm_sheet_get_validation_count(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_validation_count: out_count is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_validation_count"); rc != 0) {
    return rc;
  }
  *out_count = static_cast<std::uint32_t>(wb->workbook().sheet(sheet).validations().size());
  return 0;
}

extern "C" fm_status_t fm_sheet_get_validation_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index,
                                                  fm_data_validation* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_validation_at: out is NULL");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_get_validation_at"); rc != 0) {
    return rc;
  }
  const auto& list = wb->workbook().sheet(sheet).validations();
  if (static_cast<std::size_t>(index) >= list.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_get_validation_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(list.size()));
  }
  const formulon::DataValidation& v = list[index];
  // Ranges are layout-compatible (see static_asserts above), so we hand
  // out the engine-side vector buffer directly. The lifetime contract
  // documented in the header (valid until the next mutation that
  // touches the validation list) follows naturally from
  // `std::vector::data()` invalidation rules.
  out->ranges = v.ranges.empty() ? nullptr : reinterpret_cast<const fm_merge_range*>(v.ranges.data());
  out->range_count = static_cast<std::uint32_t>(v.ranges.size());
  out->type = v.type;
  out->op = v.op;
  out->error_style = v.error_style;
  out->allow_blank = v.allow_blank ? 1 : 0;
  out->show_input_message = v.show_input_message ? 1 : 0;
  out->show_error_message = v.show_error_message ? 1 : 0;
  out->show_dropdown = v.show_dropdown ? 1 : 0;
  out->formula1 = v.formula1.c_str();
  out->formula2 = v.formula2.c_str();
  out->error_title = v.error_title.c_str();
  out->error_message = v.error_message.c_str();
  out->prompt_title = v.prompt_title.c_str();
  out->prompt_message = v.prompt_message.c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_add_validation(fm_workbook_t* wb, std::uint32_t sheet, fm_data_validation v) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_add_validation"); rc != 0) {
    return rc;
  }
  if (v.range_count > 0 && v.ranges == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_add_validation: ranges is NULL but range_count > 0",
                             "range_count=" + std::to_string(v.range_count));
  }
  if (!check_range_count(v.range_count, "fm_sheet_add_validation")) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  formulon::DataValidation out;
  out.ranges.reserve(v.range_count);
  for (std::uint32_t i = 0; i < v.range_count; ++i) {
    formulon::MergeRange r;
    r.first_row = v.ranges[i].first_row;
    r.first_col = v.ranges[i].first_col;
    r.last_row = v.ranges[i].last_row;
    r.last_col = v.ranges[i].last_col;
    out.ranges.push_back(r);
  }
  out.type = v.type;
  out.op = v.op;
  out.error_style = v.error_style;
  out.allow_blank = v.allow_blank != 0;
  out.show_input_message = v.show_input_message != 0;
  out.show_error_message = v.show_error_message != 0;
  out.show_dropdown = v.show_dropdown != 0;
  out.formula1 = (v.formula1 != nullptr) ? std::string(v.formula1) : std::string();
  out.formula2 = (v.formula2 != nullptr) ? std::string(v.formula2) : std::string();
  out.error_title = (v.error_title != nullptr) ? std::string(v.error_title) : std::string();
  out.error_message = (v.error_message != nullptr) ? std::string(v.error_message) : std::string();
  out.prompt_title = (v.prompt_title != nullptr) ? std::string(v.prompt_title) : std::string();
  out.prompt_message = (v.prompt_message != nullptr) ? std::string(v.prompt_message) : std::string();
  wb->workbook().sheet(sheet).mutable_validations().push_back(std::move(out));
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_validation_at(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t index) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_remove_validation_at"); rc != 0) {
    return rc;
  }
  auto& list = wb->workbook().sheet(sheet).mutable_validations();
  if (static_cast<std::size_t>(index) >= list.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_remove_validation_at: index out of range",
                             "index=" + std::to_string(index) + " count=" + std::to_string(list.size()));
  }
  list.erase(list.begin() + static_cast<std::ptrdiff_t>(index));
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_validations(fm_workbook_t* wb, std::uint32_t sheet) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_sheet_clear_validations"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet).mutable_validations().clear();
  return 0;
}

// ---------------------------------------------------------------------------
// Sheet protection
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_get_protection(const fm_workbook_t* wb, uint32_t sheet_index,
                                               fm_sheet_protection_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_protection: NULL argument");
  }
  if (auto rc = check_sheet_u32(wb, sheet_index, "fm_sheet_get_protection"); rc != 0) {
    return rc;
  }
  const formulon::SheetProtection& p = wb->workbook().sheet(sheet_index).protection();
  out->enabled = p.enabled ? 1 : 0;
  out->algorithm_name = p.algorithm_name.c_str();
  out->hash_value = p.hash_value.c_str();
  out->salt_value = p.salt_value.c_str();
  out->spin_count = p.spin_count;
  out->legacy_password = p.legacy_password.c_str();
  out->sheet = p.sheet ? 1 : 0;
  out->objects = p.objects ? 1 : 0;
  out->scenarios = p.scenarios ? 1 : 0;
  out->format_cells = p.format_cells ? 1 : 0;
  out->format_columns = p.format_columns ? 1 : 0;
  out->format_rows = p.format_rows ? 1 : 0;
  out->insert_columns = p.insert_columns ? 1 : 0;
  out->insert_rows = p.insert_rows ? 1 : 0;
  out->insert_hyperlinks = p.insert_hyperlinks ? 1 : 0;
  out->delete_columns = p.delete_columns ? 1 : 0;
  out->delete_rows = p.delete_rows ? 1 : 0;
  out->select_locked_cells = p.select_locked_cells ? 1 : 0;
  out->select_unlocked_cells = p.select_unlocked_cells ? 1 : 0;
  out->sort = p.sort ? 1 : 0;
  out->auto_filter = p.auto_filter ? 1 : 0;
  out->pivot_tables = p.pivot_tables ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_protection(fm_workbook_t* wb, uint32_t sheet_index,
                                               const fm_sheet_protection_t* in) {
  clear_last_error();
  if (in == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_set_protection: NULL argument");
  }
  if (auto rc = check_sheet_u32(wb, sheet_index, "fm_sheet_set_protection"); rc != 0) {
    return rc;
  }
  formulon::SheetProtection& p = wb->workbook().sheet(sheet_index).mutable_protection();
  // Helper for pointer->string deep copy with NULL-as-empty semantics.
  const auto copy_or_empty = [](const char* s) -> std::string { return s == nullptr ? std::string() : std::string(s); };
  p.enabled = in->enabled != 0;
  p.algorithm_name = copy_or_empty(in->algorithm_name);
  p.hash_value = copy_or_empty(in->hash_value);
  p.salt_value = copy_or_empty(in->salt_value);
  p.spin_count = in->spin_count;
  p.legacy_password = copy_or_empty(in->legacy_password);
  p.sheet = in->sheet != 0;
  p.objects = in->objects != 0;
  p.scenarios = in->scenarios != 0;
  p.format_cells = in->format_cells != 0;
  p.format_columns = in->format_columns != 0;
  p.format_rows = in->format_rows != 0;
  p.insert_columns = in->insert_columns != 0;
  p.insert_rows = in->insert_rows != 0;
  p.insert_hyperlinks = in->insert_hyperlinks != 0;
  p.delete_columns = in->delete_columns != 0;
  p.delete_rows = in->delete_rows != 0;
  p.select_locked_cells = in->select_locked_cells != 0;
  p.select_unlocked_cells = in->select_unlocked_cells != 0;
  p.sort = in->sort != 0;
  p.auto_filter = in->auto_filter != 0;
  p.pivot_tables = in->pivot_tables != 0;
  return 0;
}
