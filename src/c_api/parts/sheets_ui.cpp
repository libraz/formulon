//
// C ABI - sheet UI surface: hyperlinks, merges, comments, data
// validations, and sheet protection.

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "sheet.h"
#include "threaded_comment.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::c_str_or_null;
using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_stored_formula;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::display_formula_text;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::ingest_entered_formula;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::to_cell_rect;
using formulon::c_api::parts::to_merge_range;
using formulon::c_api::parts::validate;
using formulon::c_api::parts::validation_operator_from_c;
using formulon::c_api::parts::validation_operator_to_c;
using formulon::c_api::parts::with_formula_prefix;

namespace {

using SheetItemSize = std::size_t (*)(const formulon::Sheet& sheet);

// Shared body of the `fm_sheet_get_*_count` getters.
fm_status_t sheet_item_count(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t* out_count,
                             SheetItemSize size_of) {
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_count = size_of(wb->workbook().sheet(sheet_index));
  return 0;
}

template <typename GetItems>
fm_status_t remove_sheet_item_at(fm_workbook_t* wb, std::size_t sheet_index, std::size_t index, GetItems get_items) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto& list = get_items(wb->workbook().sheet(sheet_index));
  if (auto rc = check_index(index, list.size(), "index"); rc != 0) {
    return rc;
  }
  list.erase(list.begin() + static_cast<std::ptrdiff_t>(index));
  return 0;
}

/// Validates a hyperlink record (rect inside the grid, utf8 strings) and converts it to the model type.
fm_status_t convert_hyperlink(const fm_hyperlink_t& hl, formulon::Hyperlink* out) {
  if (auto rc = check_sheet_rect(hl.rect); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({hl.target, hl.location, hl.display, hl.tooltip}); rc != 0) {
    return rc;
  }
  out->row = hl.rect.first_row;
  out->col = hl.rect.first_col;
  out->last_row = hl.rect.last_row;
  out->last_col = hl.rect.last_col;
  out->target = text_or_empty(hl.target);
  out->location = text_or_empty(hl.location);
  out->display = text_or_empty(hl.display);
  out->tooltip = text_or_empty(hl.tooltip);
  // rid stays empty; the writer mints a fresh rIdN on save.
  return 0;
}

void fill_comment(const formulon::CellComment& c, fm_comment_t* out) {
  out->row = c.row;
  out->col = c.col;
  out->author = c_str_or_null(c.author);
  out->text = c.text.c_str();
}

/// Reads of a tristate field never report DEFAULT.
std::int32_t tristate_of(bool value) {
  return value ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
}

/// DEFAULT resolves to true, Excel's default for every tristate in this file's records.
bool tristate_to_bool(std::int32_t value) {
  return value != FM_TRISTATE_FALSE;
}

}  // namespace

// ---------------------------------------------------------------------------
// Hyperlinks
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_set_hyperlink(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                              std::uint32_t col, const fm_hyperlink_t* hl) {
  clear_last_error();
  if (hl != nullptr) {
    if (auto rc = check_record_size(*hl); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  auto& hls = wb->workbook().sheet(sheet_index).mutable_hyperlinks();
  const auto anchored = [&](const formulon::Hyperlink& h) { return h.row == row && h.col == col; };
  if (hl == nullptr) {
    hls.erase(std::remove_if(hls.begin(), hls.end(), anchored), hls.end());
    return 0;
  }
  formulon::Hyperlink out;
  if (auto rc = convert_hyperlink(*hl, &out); rc != 0) {
    return rc;
  }
  if (out.row != row || out.col != col) {
    return formulon::c_api::parts::invalid_argument("hyperlink.rect must start at row and col");
  }
  // As Excel adds a link: one covering exactly the same range is replaced, any other is kept.
  const auto same_range = [&out](const formulon::Hyperlink& h) {
    return h.row == out.row && h.col == out.col && h.last_row == out.last_row && h.last_col == out.last_col;
  };
  const auto existing = std::find_if(hls.begin(), hls.end(), same_range);
  if (existing == hls.end()) {
    hls.push_back(std::move(out));
  } else {
    *existing = std::move(out);
    // A repeated range from a loaded file would otherwise survive the replacement.
    hls.erase(std::remove_if(existing + 1, hls.end(), same_range), hls.end());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_insert_hyperlink_at(fm_workbook_t* wb, std::size_t sheet_index, std::size_t index,
                                                    const fm_hyperlink_t* hl) {
  clear_last_error();
  if (hl == nullptr) {
    return null_argument("hl");
  }
  if (auto rc = check_record_size(*hl); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto& hls = wb->workbook().sheet(sheet_index).mutable_hyperlinks();
  if (auto rc = check_index(index, hls.size() + 1, "index"); rc != 0) {
    return rc;
  }
  formulon::Hyperlink out;
  if (auto rc = convert_hyperlink(*hl, &out); rc != 0) {
    return rc;
  }
  hls.insert(hls.begin() + static_cast<std::ptrdiff_t>(index), std::move(out));
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_hyperlink_at(fm_workbook_t* wb, std::size_t sheet_index, std::size_t index) {
  clear_last_error();
  return remove_sheet_item_at(wb, sheet_index, index,
                              [](formulon::Sheet& s) -> decltype(auto) { return s.mutable_hyperlinks(); });
}

extern "C" fm_status_t fm_sheet_clear_hyperlinks(fm_workbook_t* wb, std::size_t sheet_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_hyperlinks().clear();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_hyperlink_at(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t index,
                                                 fm_hyperlink_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& hls = wb->workbook().sheet(sheet_index).hyperlinks();
  if (auto rc = check_index(index, hls.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::Hyperlink& h = hls[index];
  out->rect = fm_cell_rect_t{h.row, h.col, h.last_row, h.last_col};
  out->target = c_str_or_null(h.target);
  out->location = c_str_or_null(h.location);
  out->display = c_str_or_null(h.display);
  out->tooltip = c_str_or_null(h.tooltip);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_hyperlink(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                              std::uint32_t col, std::int32_t* out_present, fm_hyperlink_t* out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  *out_present = 0;
  {
    const std::uint32_t struct_size = out->struct_size;
    *out = fm_hyperlink_t{};
    out->struct_size = struct_size;
  }
  const auto& hls = wb->workbook().sheet(sheet_index).hyperlinks();
  for (std::size_t i = 0; i < hls.size(); ++i) {
    if (hls[i].row == row && hls[i].col == col) {
      *out_present = 1;
      return fm_sheet_get_hyperlink_at(wb, sheet_index, i, out);
    }
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_hyperlink_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                    std::size_t* out_count) {
  clear_last_error();
  return sheet_item_count(wb, sheet_index, out_count, [](const formulon::Sheet& s) { return s.hyperlinks().size(); });
}

// ---------------------------------------------------------------------------
// Merges
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_add_merge(fm_workbook_t* wb, std::size_t sheet_index, const fm_cell_rect_t* merge,
                                          fm_cell_rect_t* out_merge) {
  clear_last_error();
  if (out_merge != nullptr) {
    *out_merge = fm_cell_rect_t{};
  }
  if (merge == nullptr) {
    return null_argument("merge");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*merge); rc != 0) {
    return rc;
  }
  const formulon::MergeRange range = to_merge_range(*merge);
  if (auto rc = finish_status(wb->workbook().add_merge(sheet_index, range)); rc != 0) {
    return rc;
  }
  if (out_merge != nullptr) {
    *out_merge = to_cell_rect(range);
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_merge(fm_workbook_t* wb, std::size_t sheet_index, const fm_cell_rect_t* range) {
  clear_last_error();
  if (range == nullptr) {
    return null_argument("range");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*range); rc != 0) {
    return rc;
  }
  const formulon::MergeRange target = to_merge_range(*range);
  const auto& merges = wb->workbook().sheet(sheet_index).merges();
  const bool any = std::any_of(merges.begin(), merges.end(), [&](const formulon::MergeRange& m) {
    return m.first_row <= target.last_row && target.first_row <= m.last_row && m.first_col <= target.last_col &&
           target.first_col <= m.last_col;
  });
  if (!any) {
    return not_found("no merged range intersects that range");
  }
  return finish_status(wb->workbook().remove_merges_intersecting(sheet_index, target));
}

extern "C" fm_status_t fm_sheet_remove_merge_at(fm_workbook_t* wb, std::size_t sheet_index, std::size_t index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(index, wb->workbook().sheet(sheet_index).merges().size(), "index"); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().remove_merge_at(sheet_index, index));
}

extern "C" fm_status_t fm_sheet_clear_merges(fm_workbook_t* wb, std::size_t sheet_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().clear_merges(sheet_index));
}

extern "C" fm_status_t fm_sheet_get_merge_at(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t index,
                                             fm_cell_rect_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& merges = wb->workbook().sheet(sheet_index).merges();
  if (auto rc = check_index(index, merges.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::MergeRange& m = merges[index];
  *out = to_cell_rect(m);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_merge_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                std::size_t* out_count) {
  clear_last_error();
  return sheet_item_count(wb, sheet_index, out_count, [](const formulon::Sheet& s) { return s.merges().size(); });
}

extern "C" fm_status_t fm_sheet_get_merges_in_range(const fm_workbook_t* wb, std::size_t sheet_index,
                                                    const fm_cell_rect_t* range, fm_cell_rect_t* out,
                                                    std::size_t capacity, std::size_t* out_count) {
  clear_last_error();
  if (range == nullptr || out_count == nullptr || (out == nullptr && capacity > 0U)) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*range); rc != 0) {
    return rc;
  }
  const formulon::MergeRange query = to_merge_range(*range);
  std::size_t count = 0;
  for (const formulon::MergeRange& m : wb->workbook().sheet(sheet_index).merges()) {
    if (m.last_row < query.first_row || m.first_row > query.last_row || m.last_col < query.first_col ||
        m.first_col > query.last_col) {
      continue;
    }
    if (count < capacity) {
      out[count] = to_cell_rect(m);
    }
    ++count;
  }
  *out_count = count;
  return 0;
}

// ---------------------------------------------------------------------------
// Comments
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_get_comment(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                            std::uint32_t col, int32_t* out_present, fm_comment_t* out) {
  clear_last_error();
  if (out_present == nullptr) {
    return null_argument("out_present");
  }
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_comment_t{};
  out->struct_size = struct_size;
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  for (const formulon::CellComment& c : wb->workbook().sheet(sheet_index).comments()) {
    if (c.row == row && c.col == col) {
      fill_comment(c, out);
      *out_present = 1;
      return 0;
    }
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_comment_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                  std::size_t* out_count) {
  clear_last_error();
  return sheet_item_count(wb, sheet_index, out_count, [](const formulon::Sheet& s) { return s.comments().size(); });
}

extern "C" fm_status_t fm_sheet_get_comment_at(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t index,
                                               fm_comment_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& comments = wb->workbook().sheet(sheet_index).comments();
  if (auto rc = check_index(index, comments.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::CellComment& c = comments[index];
  fill_comment(c, out);
  return 0;
}

namespace {

const formulon::CellComment* FindNote(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                      std::uint32_t col) {
  for (const formulon::CellComment& c : wb->workbook().sheet(sheet_index).comments()) {
    if (c.row == row && c.col == col) {
      return &c;
    }
  }
  return nullptr;
}

formulon::CellComment* FindNote(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  return const_cast<formulon::CellComment*>(FindNote(static_cast<const fm_workbook_t*>(wb), sheet_index, row, col));
}

}  // namespace

extern "C" fm_status_t fm_sheet_set_comment(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                            std::uint32_t col, const fm_comment_t* comment) {
  clear_last_error();
  if (comment != nullptr) {
    if (auto rc = check_record_size(*comment); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  auto& list = wb->workbook().sheet(sheet_index).mutable_comments();
  formulon::CellComment* existing = FindNote(wb, sheet_index, row, col);
  if (comment == nullptr) {
    if (existing != nullptr) {
      list.erase(list.begin() + (existing - list.data()));
    }
    return 0;
  }
  if (auto rc = check_utf8({comment->author, comment->text}); rc != 0) {
    return rc;
  }
  // A cell holds a note or a threaded comment, never both.
  for (const formulon::ThreadedComment& thread : wb->workbook().sheet(sheet_index).threaded_comments()) {
    if (thread.parent_id.empty() && thread.row == row && thread.col == col) {
      return formulon::c_api::parts::set_last_error(formulon::make_error(
          formulon::FormulonErrorCode::kAlreadyExists, "cell already has a threaded comment", thread.id));
    }
  }
  formulon::CellComment c;
  c.row = row;
  c.col = col;
  c.author = text_or_empty(comment->author);
  c.text = text_or_empty(comment->text);
  if (existing != nullptr) {
    c.uid = std::move(existing->uid);
    c.visible = existing->visible;
    c.has_box = existing->has_box;
    c.box = existing->box;
    c.vml_shape = std::move(existing->vml_shape);
    *existing = std::move(c);
  } else {
    list.push_back(std::move(c));
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_comment_visible(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                                    std::uint32_t col, std::int32_t* out_present,
                                                    std::int32_t* out_visible) {
  clear_last_error();
  if (out_present == nullptr || out_visible == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  *out_present = 0;
  *out_visible = 0;
  const formulon::CellComment* note = FindNote(wb, sheet_index, row, col);
  if (note == nullptr) {
    return 0;
  }
  *out_present = 1;
  *out_visible = note->visible ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_comment_visible(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                                    std::uint32_t col, std::int32_t visible) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  formulon::CellComment* note = FindNote(wb, sheet_index, row, col);
  if (note == nullptr) {
    return not_found("no comment at cell", "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  note->visible = visible != 0;
  return 0;
}

// ---------------------------------------------------------------------------
// Validations
// ---------------------------------------------------------------------------

// `fm_cell_rect_t` and `formulon::MergeRange` must remain layout-compatible
// so `fm_sheet_get_validation_at` can hand out the engine-side
// `std::vector<MergeRange>::data()` directly via `reinterpret_cast`. Both
// types are POD with four `uint32_t` fields in identical order; the static
// asserts below pin that contract so any future change to either struct
// trips a compile-time error.
static_assert(sizeof(fm_cell_rect_t) == sizeof(formulon::MergeRange),
              "fm_cell_rect_t / MergeRange size mismatch breaks validation getter");
static_assert(alignof(fm_cell_rect_t) == alignof(formulon::MergeRange),
              "fm_cell_rect_t / MergeRange alignment mismatch breaks validation getter");
static_assert(offsetof(fm_cell_rect_t, first_row) == offsetof(formulon::MergeRange, first_row),
              "cell rect field first_row layout mismatch");
static_assert(offsetof(fm_cell_rect_t, first_col) == offsetof(formulon::MergeRange, first_col),
              "cell rect field first_col layout mismatch");
static_assert(offsetof(fm_cell_rect_t, last_row) == offsetof(formulon::MergeRange, last_row),
              "cell rect field last_row layout mismatch");
static_assert(offsetof(fm_cell_rect_t, last_col) == offsetof(formulon::MergeRange, last_col),
              "cell rect field last_col layout mismatch");

extern "C" fm_status_t fm_sheet_get_validation_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t* out_count) {
  clear_last_error();
  return sheet_item_count(wb, sheet_index, out_count, [](const formulon::Sheet& s) { return s.validations().size(); });
}

namespace {

/// Writes validation `index` of `sheet_index` into `out`; its strings live in the read scratch.
void fill_validation(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t index, fm_data_validation_t* out) {
  const formulon::Workbook& book = wb->workbook();
  const formulon::DataValidation& v = book.sheet(sheet_index).validations()[index];
  // Ranges are layout-compatible (see static_asserts above), so we hand
  // out the engine-side vector buffer directly. The lifetime contract
  // documented in the header (valid until the next mutation that
  // touches the validation list) follows naturally from
  // `std::vector::data()` invalidation rules.
  out->ranges = v.ranges.empty() ? nullptr : reinterpret_cast<const fm_cell_rect_t*>(v.ranges.data());
  out->range_count = static_cast<std::uint32_t>(v.ranges.size());
  out->type = v.type;
  out->op = validation_operator_to_c(v.op);
  out->error_style = v.error_style;
  out->allow_blank = tristate_of(v.allow_blank);
  out->show_input_message = tristate_of(v.show_input_message);
  out->show_error_message = tristate_of(v.show_error_message);
  out->show_in_cell_dropdown = tristate_of(v.show_dropdown);
  formulon::c_api::parts::TextStore& store = fresh_read_scratch(wb);
  out->id =
      store_cstr(store, formulon::c_api::parts::format_runtime_id(book.validation_runtime_id(sheet_index, index)));
  out->formula1 = nullptr;
  out->formula2 = nullptr;
  if (!v.formula1.empty()) {
    out->formula1 = store_cstr(store, with_formula_prefix(display_formula_text(book, v.formula1, {})));
  }
  if (!v.formula2.empty()) {
    out->formula2 = store_cstr(store, with_formula_prefix(display_formula_text(book, v.formula2, {})));
  }
  out->error_title = c_str_or_null(v.error_title);
  out->error_message = c_str_or_null(v.error_message);
  out->prompt_title = c_str_or_null(v.prompt_title);
  out->prompt_message = c_str_or_null(v.prompt_message);
}

/// Position of the validation with id text `id` on `sheet_index`; nullopt when the text names none.
std::optional<std::size_t> locate_validation(const fm_workbook_t* wb, std::size_t sheet_index, const char* id) {
  std::uint32_t runtime_id = 0;
  if (!formulon::c_api::parts::parse_runtime_id(id, &runtime_id)) {
    return std::nullopt;
  }
  return wb->workbook().find_validation(sheet_index, runtime_id);
}

fm_status_t validation_not_found(const char* id) {
  return not_found("no data validation has that id", "id=" + std::string(id));
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_validation_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                  std::size_t validation_index, fm_data_validation_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& list = wb->workbook().sheet(sheet_index).validations();
  if (auto rc = check_index(validation_index, list.size(), "validation_index"); rc != 0) {
    return rc;
  }
  fill_validation(wb, sheet_index, validation_index, out);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_validation(const fm_workbook_t* wb, std::size_t sheet_index, const char* id,
                                               std::int32_t* out_present, fm_data_validation_t* out) {
  clear_last_error();
  if (id == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument(id == nullptr ? "id" : (out_present == nullptr ? "out_present" : "out"));
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({id}); rc != 0) {
    return rc;
  }
  const std::optional<std::size_t> index = locate_validation(wb, sheet_index, id);
  if (!index.has_value()) {
    const std::uint32_t struct_size = out->struct_size;
    *out = FM_STRUCT_INIT(fm_data_validation_t);
    out->struct_size = struct_size;
    *out_present = 0;
    return 0;
  }
  fill_validation(wb, sheet_index, *index, out);
  *out_present = 1;
  return 0;
}

extern "C" fm_status_t fm_sheet_add_validation(fm_workbook_t* wb, std::size_t sheet_index,
                                               const fm_data_validation_t* v, const char** out_validation_id) {
  clear_last_error();
  if (v == nullptr || out_validation_id == nullptr) {
    return null_argument(v == nullptr ? "v" : "out_validation_id");
  }
  if (auto rc = check_record_size(*v); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = validate(*v); rc != 0) {
    return rc;
  }
  const char* const formula_fields[] = {v->formula1, v->formula2};
  formulon::DataValidation out;
  std::string* const formula_slots[] = {&out.formula1, &out.formula2};
  for (std::size_t i = 0; i < 2; ++i) {
    if (formula_fields[i] != nullptr) {
      if (auto rc = check_stored_formula(wb, formula_fields[i], i == 0 ? "field=formula1" : "field=formula2",
                                         formula_slots[i]);
          rc != 0) {
        return rc;
      }
    }
  }
  out.ranges.reserve(v->range_count);
  for (std::size_t i = 0; i < v->range_count; ++i) {
    out.ranges.push_back(to_merge_range(v->ranges[i]));
  }
  // `validate` bounded both enumerations to their uint8_t model domains.
  out.type = static_cast<std::uint8_t>(v->type);
  out.op = validation_operator_from_c(v->op);
  out.error_style = static_cast<std::uint8_t>(v->error_style);
  out.allow_blank = tristate_to_bool(v->allow_blank);
  out.show_input_message = tristate_to_bool(v->show_input_message);
  out.show_error_message = tristate_to_bool(v->show_error_message);
  out.show_dropdown = tristate_to_bool(v->show_in_cell_dropdown);
  out.error_title = text_or_empty(v->error_title);
  out.error_message = text_or_empty(v->error_message);
  out.prompt_title = text_or_empty(v->prompt_title);
  out.prompt_message = text_or_empty(v->prompt_message);
  for (std::size_t i = 0; i < 2; ++i) {
    if (formula_fields[i] != nullptr) {
      *formula_slots[i] = ingest_entered_formula(wb, formula_fields[i]);
    }
  }
  auto added = wb->workbook().add_validation(sheet_index, std::move(out));
  if (!added) {
    return set_last_error(added.error());
  }
  *out_validation_id = publish_scratch(wb, formulon::c_api::parts::format_runtime_id(added.value()));
  return 0;
}

namespace {

/// A patch string field: checked text and, for a formula, its stored spelling.
struct PatchedText {
  fm_patch_op_t op;
  const char* text;
  const char* field;
  bool is_formula;
  std::string stored;
};

fm_status_t check_validation_patch(const fm_workbook_t* wb, const fm_data_validation_patch_t& p,
                                   std::array<PatchedText, 6>* texts) {
  if (!formulon::c_api::parts::patch_op_ok_without_clear(p.ranges_op) ||
      !formulon::c_api::parts::patch_op_ok_without_clear(p.type_op) ||
      !formulon::c_api::parts::patch_op_ok_without_clear(p.op_op) ||
      !formulon::c_api::parts::patch_op_ok_without_clear(p.error_style_op) ||
      !formulon::c_api::parts::patch_op_ok_without_clear(p.allow_blank_op) ||
      !formulon::c_api::parts::patch_op_ok_without_clear(p.show_input_message_op) ||
      !formulon::c_api::parts::patch_op_ok_without_clear(p.show_error_message_op) ||
      !formulon::c_api::parts::patch_op_ok_without_clear(p.show_in_cell_dropdown_op)) {
    return formulon::c_api::parts::invalid_argument("a patch op is not allowed for a field that cannot be cleared");
  }
  if (p.ranges_op == FM_PATCH_OP_SET) {
    if (auto rc = formulon::c_api::parts::check_rect_list(p.ranges, p.range_count); rc != 0) {
      return rc;
    }
  }
  struct Domain {
    fm_patch_op_t op;
    std::int32_t value;
    std::int64_t max;
    const char* field;
    std::int64_t min;
  };
  const Domain domains[] = {
      {p.type_op, p.type, 7, "type", 0},
      {p.op_op, p.op, FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL, "op", 0},
      {p.error_style_op, p.error_style, 2, "error_style", 0},
      {p.allow_blank_op, p.allow_blank, FM_TRISTATE_TRUE, "allow_blank", FM_TRISTATE_DEFAULT},
      {p.show_input_message_op, p.show_input_message, FM_TRISTATE_TRUE, "show_input_message", FM_TRISTATE_DEFAULT},
      {p.show_error_message_op, p.show_error_message, FM_TRISTATE_TRUE, "show_error_message", FM_TRISTATE_DEFAULT},
      {p.show_in_cell_dropdown_op, p.show_in_cell_dropdown, FM_TRISTATE_TRUE, "show_in_cell_dropdown",
       FM_TRISTATE_DEFAULT},
  };
  for (const Domain& d : domains) {
    if (d.op == FM_PATCH_OP_SET) {
      if (auto rc = formulon::c_api::parts::check_enum_domain(d.value, d.min, d.max, d.field); rc != 0) {
        return rc;
      }
    }
  }
  *texts = {{{p.formula1_op, p.formula1, "field=formula1", true, {}},
             {p.formula2_op, p.formula2, "field=formula2", true, {}},
             {p.error_title_op, p.error_title, "field=error_title", false, {}},
             {p.error_message_op, p.error_message, "field=error_message", false, {}},
             {p.prompt_title_op, p.prompt_title, "field=prompt_title", false, {}},
             {p.prompt_message_op, p.prompt_message, "field=prompt_message", false, {}}}};
  for (PatchedText& t : *texts) {
    if (!formulon::c_api::parts::patch_op_ok(t.op)) {
      return formulon::c_api::parts::invalid_argument((std::string(t.field + 6) + "_op is not a patch op").c_str(),
                                                      t.field);
    }
    if (t.op != FM_PATCH_OP_SET) {
      continue;
    }
    if (t.text == nullptr) {
      return formulon::c_api::parts::null_field(t.field + 6);
    }
    if (auto rc = check_utf8({t.text}); rc != 0) {
      return rc;
    }
    if (t.is_formula) {
      if (auto rc = check_stored_formula(wb, t.text, t.field, &t.stored); rc != 0) {
        return rc;
      }
    }
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_update_validation(fm_workbook_t* wb, std::size_t sheet_index, const char* id,
                                                  const fm_data_validation_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || id == nullptr || patch == nullptr) {
    return null_argument(wb == nullptr ? "wb" : (id == nullptr ? "id" : "patch"));
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({id}); rc != 0) {
    return rc;
  }
  const std::optional<std::size_t> index = locate_validation(wb, sheet_index, id);
  if (!index.has_value()) {
    return validation_not_found(id);
  }
  std::array<PatchedText, 6> texts{};
  if (auto rc = check_validation_patch(wb, *patch, &texts); rc != 0) {
    return rc;
  }
  formulon::DataValidation& v = wb->workbook().sheet(sheet_index).mutable_validations()[*index];
  if (patch->ranges_op == FM_PATCH_OP_SET) {
    v.ranges.clear();
    v.ranges.reserve(patch->range_count);
    for (std::size_t i = 0; i < patch->range_count; ++i) {
      v.ranges.push_back(to_merge_range(patch->ranges[i]));
    }
  }
  if (patch->type_op == FM_PATCH_OP_SET) {
    v.type = static_cast<std::uint8_t>(patch->type);
  }
  if (patch->op_op == FM_PATCH_OP_SET) {
    v.op = validation_operator_from_c(patch->op);
  }
  if (patch->error_style_op == FM_PATCH_OP_SET) {
    v.error_style = static_cast<std::uint8_t>(patch->error_style);
  }
  if (patch->allow_blank_op == FM_PATCH_OP_SET) {
    v.allow_blank = tristate_to_bool(patch->allow_blank);
  }
  if (patch->show_input_message_op == FM_PATCH_OP_SET) {
    v.show_input_message = tristate_to_bool(patch->show_input_message);
  }
  if (patch->show_error_message_op == FM_PATCH_OP_SET) {
    v.show_error_message = tristate_to_bool(patch->show_error_message);
  }
  if (patch->show_in_cell_dropdown_op == FM_PATCH_OP_SET) {
    v.show_dropdown = tristate_to_bool(patch->show_in_cell_dropdown);
  }
  std::string* const slots[] = {&v.formula1,      &v.formula2,     &v.error_title,
                                &v.error_message, &v.prompt_title, &v.prompt_message};
  for (std::size_t i = 0; i < texts.size(); ++i) {
    const PatchedText& t = texts[i];
    if (t.op == FM_PATCH_OP_CLEAR) {
      slots[i]->clear();
    } else if (t.op == FM_PATCH_OP_SET) {
      *slots[i] = t.is_formula ? ingest_entered_formula(wb, t.text) : std::string(t.text);
    }
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_validation(fm_workbook_t* wb, std::size_t sheet_index, const char* id) {
  clear_last_error();
  if (wb == nullptr || id == nullptr) {
    return null_argument(wb == nullptr ? "wb" : "id");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({id}); rc != 0) {
    return rc;
  }
  const std::optional<std::size_t> index = locate_validation(wb, sheet_index, id);
  if (!index.has_value()) {
    return validation_not_found(id);
  }
  std::uint32_t runtime_id = 0;
  formulon::c_api::parts::parse_runtime_id(id, &runtime_id);
  return finish_status(wb->workbook().remove_validation(sheet_index, runtime_id));
}

extern "C" fm_status_t fm_sheet_clear_validations(fm_workbook_t* wb, std::size_t sheet_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_validations().clear();
  return 0;
}

// ---------------------------------------------------------------------------
// Sheet protection
// ---------------------------------------------------------------------------

namespace {

using formulon::SheetProtection;

/// One boolean `allow_*` field: the engine member holds the inverse (the OOXML "locked" meaning).
struct AllowFlag {
  bool SheetProtection::*locked;
  std::int32_t fm_sheet_protection_t::*value;
  fm_patch_op_t fm_sheet_protection_patch_t::*patch_op;
  std::int32_t fm_sheet_protection_patch_t::*patch_value;
  bool is_select;  // fm_tristate_t instead of 0 / 1
  const char* name;
};

#define FM_ALLOW_FLAG(field, member, select) \
  {&SheetProtection::member,                 \
   &fm_sheet_protection_t::field,            \
   &fm_sheet_protection_patch_t::field##_op, \
   &fm_sheet_protection_patch_t::field,      \
   select,                                   \
   #field}

const AllowFlag kAllowFlags[] = {
    FM_ALLOW_FLAG(allow_edit_objects, objects, false),
    FM_ALLOW_FLAG(allow_edit_scenarios, scenarios, false),
    FM_ALLOW_FLAG(allow_format_cells, format_cells, false),
    FM_ALLOW_FLAG(allow_format_columns, format_columns, false),
    FM_ALLOW_FLAG(allow_format_rows, format_rows, false),
    FM_ALLOW_FLAG(allow_insert_columns, insert_columns, false),
    FM_ALLOW_FLAG(allow_insert_rows, insert_rows, false),
    FM_ALLOW_FLAG(allow_insert_hyperlinks, insert_hyperlinks, false),
    FM_ALLOW_FLAG(allow_delete_columns, delete_columns, false),
    FM_ALLOW_FLAG(allow_delete_rows, delete_rows, false),
    FM_ALLOW_FLAG(allow_select_locked_cells, select_locked_cells, true),
    FM_ALLOW_FLAG(allow_select_unlocked_cells, select_unlocked_cells, true),
    FM_ALLOW_FLAG(allow_sort, sort, false),
    FM_ALLOW_FLAG(allow_auto_filter, auto_filter, false),
    FM_ALLOW_FLAG(allow_pivot_tables, pivot_tables, false),
};

#undef FM_ALLOW_FLAG

bool flag_in_domain(const AllowFlag& flag, std::int32_t value) {
  return value >= 0 && value <= (flag.is_select ? static_cast<std::int32_t>(FM_TRISTATE_TRUE) : 1);
}

/// Whether the flag value permits the action (DEFAULT permits, as Excel's select flags do).
bool permits(const AllowFlag& flag, std::int32_t value) {
  return flag.is_select ? value != FM_TRISTATE_FALSE : value != 0;
}

void apply_flag(SheetProtection& p, const AllowFlag& flag, std::int32_t value) {
  p.*flag.locked = !permits(flag, value);
}

std::int32_t read_flag(const SheetProtection& p, const AllowFlag& flag) {
  const bool allowed = !(p.*flag.locked);
  return flag.is_select ? (allowed ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE) : (allowed ? 1 : 0);
}

SheetProtection protection_from_record(const fm_sheet_protection_t& in) {
  SheetProtection p;
  p.enabled = true;
  p.sheet = true;
  p.algorithm_name = text_or_empty(in.algorithm_name);
  p.hash_value = text_or_empty(in.hash_value);
  p.salt_value = text_or_empty(in.salt_value);
  p.spin_count = in.spin_count;
  p.legacy_password = text_or_empty(in.legacy_password_hash);
  for (const AllowFlag& flag : kAllowFlags) {
    apply_flag(p, flag, in.*flag.value);
  }
  return p;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_sheet_protection(const fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::int32_t* out_present, fm_sheet_protection_t* out) {
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
  const SheetProtection& p = wb->workbook().sheet(sheet_index).protection();
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_sheet_protection_t{};
  out->struct_size = struct_size;
  if (!p.enabled) {
    *out_present = 0;
    return 0;
  }
  *out_present = 1;
  out->algorithm_name = c_str_or_null(p.algorithm_name);
  out->hash_value = c_str_or_null(p.hash_value);
  out->salt_value = c_str_or_null(p.salt_value);
  out->spin_count = p.spin_count;
  out->legacy_password_hash = c_str_or_null(p.legacy_password);
  for (const AllowFlag& flag : kAllowFlags) {
    out->*flag.value = read_flag(p, flag);
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_set_sheet_protection(fm_workbook_t* wb, std::size_t sheet_index,
                                                     const fm_sheet_protection_t* in) {
  clear_last_error();
  if (in != nullptr) {
    if (auto rc = check_record_size(*in); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (in == nullptr) {
    if (wb->workbook().sheet(sheet_index).protection().enabled) {
      wb->workbook().sheet(sheet_index).mutable_protection() = SheetProtection{};
    }
    return 0;
  }
  if (auto rc = check_utf8({in->algorithm_name, in->hash_value, in->salt_value, in->legacy_password_hash}); rc != 0) {
    return rc;
  }
  for (const AllowFlag& flag : kAllowFlags) {
    if (!flag_in_domain(flag, in->*flag.value)) {
      return formulon::c_api::parts::invalid_argument((std::string(flag.name) + " is out of range").c_str());
    }
  }
  wb->workbook().sheet(sheet_index).mutable_protection() = protection_from_record(*in);
  return 0;
}

extern "C" fm_status_t fm_sheet_update_sheet_protection(fm_workbook_t* wb, std::size_t sheet_index,
                                                        const fm_sheet_protection_patch_t* patch) {
  clear_last_error();
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  struct StringField {
    fm_patch_op_t op;
    const char* value;
    std::string SheetProtection::*member;
    const char* name;
  };
  const StringField strings[] = {
      {patch->algorithm_name_op, patch->algorithm_name, &SheetProtection::algorithm_name, "algorithm_name"},
      {patch->hash_value_op, patch->hash_value, &SheetProtection::hash_value, "hash_value"},
      {patch->salt_value_op, patch->salt_value, &SheetProtection::salt_value, "salt_value"},
      {patch->legacy_password_hash_op, patch->legacy_password_hash, &SheetProtection::legacy_password,
       "legacy_password_hash"},
  };
  bool any_change = false;
  for (const StringField& f : strings) {
    if (!patch_op_ok(f.op)) {
      return formulon::c_api::parts::invalid_argument((std::string(f.name) + "_op is not a patch op").c_str());
    }
    if (f.op == FM_PATCH_OP_SET && f.value == nullptr) {
      return formulon::c_api::parts::null_field(f.name);
    }
    if (f.op == FM_PATCH_OP_SET) {
      if (auto rc = check_utf8(f.value, "patch string"); rc != 0) {
        return rc;
      }
    }
    any_change = any_change || f.op != FM_PATCH_OP_KEEP;
  }
  if (!patch_op_ok(patch->spin_count_op)) {
    return formulon::c_api::parts::invalid_argument("spin_count_op is not a patch op");
  }
  any_change = any_change || patch->spin_count_op != FM_PATCH_OP_KEEP;
  for (const AllowFlag& flag : kAllowFlags) {
    const fm_patch_op_t op = patch->*flag.patch_op;
    if (!patch_op_ok_without_clear(op)) {
      return formulon::c_api::parts::invalid_argument((std::string(flag.name) + "_op is not a patch op").c_str());
    }
    // A patch value is a plain 0 / 1; only the record's select flags carry a tristate.
    if (op == FM_PATCH_OP_SET && patch->*flag.patch_value != 0 && patch->*flag.patch_value != 1) {
      return formulon::c_api::parts::invalid_argument((std::string(flag.name) + " must be in 0..1").c_str());
    }
    any_change = any_change || op != FM_PATCH_OP_KEEP;
  }
  if (!any_change) {
    return 0;
  }
  SheetProtection p = wb->workbook().sheet(sheet_index).protection();
  if (!p.enabled) {
    p = protection_from_record(FM_STRUCT_INIT(fm_sheet_protection_t));
  }
  for (const StringField& f : strings) {
    if (f.op == FM_PATCH_OP_SET) {
      p.*f.member = f.value;
    } else if (f.op == FM_PATCH_OP_CLEAR) {
      (p.*f.member).clear();
    }
  }
  if (patch->spin_count_op == FM_PATCH_OP_SET) {
    p.spin_count = patch->spin_count;
  } else if (patch->spin_count_op == FM_PATCH_OP_CLEAR) {
    p.spin_count = 0;
  }
  for (const AllowFlag& flag : kAllowFlags) {
    if (patch->*flag.patch_op == FM_PATCH_OP_SET) {
      const std::int32_t value = patch->*flag.patch_value;
      apply_flag(p, flag, flag.is_select ? (value != 0 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE) : value);
    }
  }
  wb->workbook().sheet(sheet_index).mutable_protection() = std::move(p);
  return 0;
}
