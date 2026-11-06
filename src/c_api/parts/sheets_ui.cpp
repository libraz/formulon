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
#include "threaded_comment.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_stored_formula;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::display_formula;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::ingest_entered_formula;
using formulon::c_api::parts::may_respell_formula;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::to_cell_rect;
using formulon::c_api::parts::to_merge_range;
using formulon::c_api::parts::validate;

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

void fill_comment(const formulon::CellComment& c, fm_comment_t* out) {
  out->row = c.row;
  out->col = c.col;
  out->author = c.author.c_str();
  out->text = c.text.c_str();
}

}  // namespace

// ---------------------------------------------------------------------------
// Hyperlinks
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_add_hyperlink(fm_workbook_t* wb, std::size_t sheet_index, const fm_hyperlink_t* hl) {
  clear_last_error();
  if (hl == nullptr) {
    return null_argument("hl");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(hl->rect); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({hl->target, hl->location, hl->display, hl->tooltip}); rc != 0) {
    return rc;
  }
  formulon::Hyperlink out;
  out.row = hl->rect.first_row;
  out.col = hl->rect.first_col;
  out.last_row = hl->rect.last_row;
  out.last_col = hl->rect.last_col;
  out.target = text_or_empty(hl->target);
  out.location = text_or_empty(hl->location);
  out.display = text_or_empty(hl->display);
  out.tooltip = text_or_empty(hl->tooltip);
  // rid stays empty; the writer mints a fresh rIdN on save.
  wb->workbook().sheet(sheet_index).mutable_hyperlinks().push_back(std::move(out));
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_hyperlink(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                                 std::uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto& hls = wb->workbook().sheet(sheet_index).mutable_hyperlinks();
  const auto removed = std::remove_if(hls.begin(), hls.end(),
                                      [&](const formulon::Hyperlink& h) { return h.row == row && h.col == col; });
  if (removed == hls.end()) {
    return not_found("no hyperlink at that cell", "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  hls.erase(removed, hls.end());
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
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& hls = wb->workbook().sheet(sheet_index).hyperlinks();
  if (auto rc = check_index(index, hls.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::Hyperlink& h = hls[index];
  out->rect = fm_cell_rect_t{h.row, h.col, h.last_row, h.last_col};
  out->target = h.target.c_str();
  out->location = h.location.c_str();
  out->display = h.display.c_str();
  out->tooltip = h.tooltip.c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_hyperlink(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                              std::uint32_t col, std::int32_t* out_present, fm_hyperlink_t* out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_present = 0;
  *out = fm_hyperlink_t{};
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

extern "C" fm_status_t fm_sheet_add_merge(fm_workbook_t* wb, std::size_t sheet_index, const fm_cell_rect_t* merge) {
  clear_last_error();
  if (merge == nullptr) {
    return null_argument("merge");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*merge); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().add_merge(sheet_index, to_merge_range(*merge)));
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
                                                    std::uint32_t capacity, std::uint32_t* out_count) {
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
  std::uint32_t count = 0;
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
  *out_present = 0;
  *out = fm_comment_t{};
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
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
                                            std::uint32_t col, const char* author, const char* text) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({author, text}); rc != 0) {
    return rc;
  }
  auto& list = wb->workbook().sheet(sheet_index).mutable_comments();
  formulon::CellComment* existing = FindNote(wb, sheet_index, row, col);
  // Empty/NULL text => removal.
  if (text == nullptr || text[0] == '\0') {
    if (existing != nullptr) {
      list.erase(list.begin() + (existing - list.data()));
    }
    return 0;
  }
  // A cell holds a note or a threaded comment, never both.
  for (const formulon::ThreadedComment& thread : wb->workbook().sheet(sheet_index).threaded_comments()) {
    if (thread.parent_id.empty() && thread.row == row && thread.col == col) {
      return formulon::c_api::parts::set_last_error(formulon::make_error(
          formulon::FormulonErrorCode::kThreadedCommentInvalid, "cell already has a threaded comment", thread.id));
    }
  }
  formulon::CellComment c;
  c.row = row;
  c.col = col;
  c.author = text_or_empty(author);
  c.text = std::string(text);
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
                                                    std::uint32_t col, std::int32_t* out_visible) {
  clear_last_error();
  if (out_visible == nullptr) {
    return null_argument("out_visible");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const formulon::CellComment* note = FindNote(wb, sheet_index, row, col);
  if (note == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kNotFound, "no comment at cell",
                             "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  *out_visible = note->visible ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_comment_visible(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                                    std::uint32_t col, std::int32_t visible) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::CellComment* note = FindNote(wb, sheet_index, row, col);
  if (note == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kNotFound, "no comment at cell",
                             "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  note->visible = visible != 0;
  return 0;
}

// ---------------------------------------------------------------------------
// Validations
// ---------------------------------------------------------------------------

// `fm_cell_rect_t` and `formulon::MergeRange` must remain layout-compatible
// so `fm_sheet_get_validation` can hand out the engine-side
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

extern "C" fm_status_t fm_sheet_get_validation(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t index,
                                               fm_data_validation_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& list = wb->workbook().sheet(sheet_index).validations();
  if (auto rc = check_index(index, list.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::DataValidation& v = list[index];
  // Ranges are layout-compatible (see static_asserts above), so we hand
  // out the engine-side vector buffer directly. The lifetime contract
  // documented in the header (valid until the next mutation that
  // touches the validation list) follows naturally from
  // `std::vector::data()` invalidation rules.
  out->ranges = v.ranges.empty() ? nullptr : reinterpret_cast<const fm_cell_rect_t*>(v.ranges.data());
  out->range_count = v.ranges.size();
  out->type = v.type;
  out->op = v.op;
  out->error_style = v.error_style;
  out->allow_blank = v.allow_blank ? 1 : 0;
  out->show_input_message = v.show_input_message ? 1 : 0;
  out->show_error_message = v.show_error_message ? 1 : 0;
  out->show_dropdown = v.show_dropdown ? 1 : 0;
  const formulon::Workbook& book = wb->workbook();
  if (may_respell_formula(v.formula1) || may_respell_formula(v.formula2)) {
    formulon::c_api::parts::TextStore& store = fresh_read_scratch(wb);
    out->formula1 = display_formula(book, v.formula1, store, {});
    out->formula2 = display_formula(book, v.formula2, store, {});
  } else {
    out->formula1 = v.formula1.c_str();
    out->formula2 = v.formula2.c_str();
  }
  out->error_title = v.error_title.c_str();
  out->error_message = v.error_message.c_str();
  out->prompt_title = v.prompt_title.c_str();
  out->prompt_message = v.prompt_message.c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_add_validation(fm_workbook_t* wb, std::size_t sheet_index,
                                               const fm_data_validation_t* v, std::size_t* out_index) {
  clear_last_error();
  if (v == nullptr || out_index == nullptr) {
    return null_argument(v == nullptr ? "v" : "out_index");
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
      if (auto rc =
              check_stored_formula(formula_fields[i], i == 0 ? "field=formula1" : "field=formula2", formula_slots[i]);
          rc != 0) {
        return rc;
      }
    }
  }
  out.ranges.reserve(v->range_count);
  for (std::size_t i = 0; i < v->range_count; ++i) {
    out.ranges.push_back(to_merge_range(v->ranges[i]));
  }
  out.type = v->type;
  out.op = v->op;
  out.error_style = v->error_style;
  out.allow_blank = v->allow_blank != 0;
  out.show_input_message = v->show_input_message != 0;
  out.show_error_message = v->show_error_message != 0;
  out.show_dropdown = v->show_dropdown != 0;
  out.error_title = text_or_empty(v->error_title);
  out.error_message = text_or_empty(v->error_message);
  out.prompt_title = text_or_empty(v->prompt_title);
  out.prompt_message = text_or_empty(v->prompt_message);
  for (std::size_t i = 0; i < 2; ++i) {
    if (formula_fields[i] != nullptr) {
      *formula_slots[i] = ingest_entered_formula(wb, formula_fields[i]);
    }
  }
  auto& validations = wb->workbook().sheet(sheet_index).mutable_validations();
  validations.push_back(std::move(out));
  *out_index = validations.size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_validation(fm_workbook_t* wb, std::size_t sheet_index, std::size_t index) {
  clear_last_error();
  return remove_sheet_item_at(wb, sheet_index, index,
                              [](formulon::Sheet& s) -> decltype(auto) { return s.mutable_validations(); });
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

extern "C" fm_status_t fm_sheet_get_sheet_protection(const fm_workbook_t* wb, std::size_t sheet_index,
                                                     fm_sheet_protection_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
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

extern "C" fm_status_t fm_sheet_set_sheet_protection(fm_workbook_t* wb, std::size_t sheet_index,
                                                     const fm_sheet_protection_t* in) {
  clear_last_error();
  if (in == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({in->algorithm_name, in->hash_value, in->salt_value, in->legacy_password}); rc != 0) {
    return rc;
  }
  formulon::SheetProtection& p = wb->workbook().sheet(sheet_index).mutable_protection();
  // Helper for pointer->string deep copy with NULL-as-empty semantics.
  p.enabled = in->enabled != 0;
  p.algorithm_name = text_or_empty(in->algorithm_name);
  p.hash_value = text_or_empty(in->hash_value);
  p.salt_value = text_or_empty(in->salt_value);
  p.spin_count = in->spin_count;
  p.legacy_password = text_or_empty(in->legacy_password);
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
