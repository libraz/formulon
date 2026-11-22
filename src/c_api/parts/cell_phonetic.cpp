//
// C ABI - cell phonetic guides: furigana text, runs and properties, and the
// per-cell show-phonetic flag.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "cell.h"
#include "phonetic.h"
#include "sheet.h"
#include "workbook.h"

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_cell_text_length;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_limit;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::code_points_to_utf16_units;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::reset_record;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::SheetUse;
using formulon::c_api::parts::utf16_units_to_code_points;

extern "C" fm_status_t fm_sheet_set_cell_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                  const char* utf8) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  // NULL and "" are one state (no reading) and both clear it.
  if (utf8 == nullptr) {
    utf8 = "";
  }
  if (auto rc = check_utf8({utf8}); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_text_length(utf8, "text"); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).set_cell_phonetic(row, col, utf8);
  // `PHONETIC()` reads this cell's phonetic reading through the same
  // reference edge any other function argument creates, so its dependents
  // need the same dirty mark a value/text write gives them -- the reading
  // itself changed without the cell's own Value changing, so the dep-graph
  // registration is untouched.
  wb->workbook().mark_cell_dependents_dirty(sheet_index, row, col);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_cell_phonetic_runs(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                       uint32_t col, int32_t offset_unit, const fm_phonetic_run_t* runs,
                                                       size_t count) {
  clear_last_error();
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (runs == nullptr && count != 0U) {
    return null_argument("runs");
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  if (auto rc =
          check_enum_domain(offset_unit, FM_TEXT_OFFSET_UNIT_UTF16, FM_TEXT_OFFSET_UNIT_CODE_POINT, "offset_unit");
      rc != 0) {
    return rc;
  }
  // The surface text the spans index; a cell without text yet has none.
  const formulon::Cell* existing = wb->workbook().sheet(sheet_index).cell_at(row, col);
  const std::string_view surface =
      existing != nullptr && existing->cached_value.is_text() ? existing->cached_value.as_text() : std::string_view();
  const bool by_code_point = offset_unit == FM_TEXT_OFFSET_UNIT_CODE_POINT;

  std::vector<formulon::PhoneticRun> parsed;
  parsed.reserve(count);
  // `PHONETIC` walks the runs once in order, emitting each run's kana when
  // the walk reaches its start. Runs that overlap or run backwards would
  // still all be emitted, silently attaching kana to characters they do not
  // read, so the ordered-partition rule is enforced here rather than
  // repaired at composition time.
  std::uint32_t previous_end = 0;
  for (size_t i = 0; i < count; ++i) {
    const fm_phonetic_run_t& run = runs[i];
    if (run.text == nullptr) {
      return null_field("text");
    }
    if (auto rc = check_utf8(run.text, "run_index", i); rc != 0) {
      return rc;
    }
    if (auto rc = check_cell_text_length(run.text, ("runs[" + std::to_string(i) + "].text").c_str()); rc != 0) {
      return rc;
    }
    const std::uint32_t start = by_code_point ? code_points_to_utf16_units(surface, run.start_cp) : run.start;
    const std::uint32_t end = by_code_point ? code_points_to_utf16_units(surface, run.end_cp) : run.end;
    if (end < start) {
      return invalid_argument(
          ("runs[" + std::to_string(i) + "].end must not be before start").c_str(),
          "run_index=" + std::to_string(i) + " start=" + std::to_string(start) + " end=" + std::to_string(end));
    }
    if (i != 0U && start < previous_end) {
      return invalid_argument(("runs[" + std::to_string(i) + "].start must not be before the previous end").c_str(),
                              "run_index=" + std::to_string(i) + " start=" + std::to_string(start) +
                                  " previous_end=" + std::to_string(previous_end));
    }
    previous_end = end;
    parsed.push_back(formulon::PhoneticRun{start, end, std::string(run.text)});
  }

  wb->workbook().sheet(sheet_index).set_cell_phonetic_runs(row, col, std::move(parsed));
  wb->workbook().mark_cell_dependents_dirty(sheet_index, row, col);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_cell_phonetic_properties(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                             uint32_t col, const fm_phonetic_properties_t* properties) {
  clear_last_error();
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  if (properties == nullptr) {
    // A cell that does not exist already reads as the unstated triple.
    if (sheet.cell_at(row, col) == nullptr) {
      return 0;
    }
    sheet.set_cell_phonetic_props(row, col, formulon::PhoneticProperties{}, /*explicit_pr=*/false);
    wb->workbook().mark_cell_dependents_dirty(sheet_index, row, col);
    return 0;
  }
  if (auto rc = check_record_size(*properties); rc != 0) {
    return rc;
  }
  const std::uint32_t font_index = properties->font_index;
  const std::int32_t type = properties->type;
  const std::int32_t alignment = properties->alignment;
  // Both enumerations occupy two bits of the XLSB trailer's flags word, so
  // an out-of-range value would not survive an `.xlsb` save intact. Reject
  // it here rather than truncate it into a different, valid-looking form.
  if (type < FM_PHONETIC_TYPE_HALFWIDTH_KATAKANA || type > FM_PHONETIC_TYPE_NO_CONVERSION) {
    return invalid_argument(("type " + std::to_string(type) + " is not a PhoneticType").c_str(),
                            "type=" + std::to_string(type));
  }
  if (alignment < FM_PHONETIC_ALIGNMENT_NO_CONTROL || alignment > FM_PHONETIC_ALIGNMENT_DISTRIBUTED) {
    return invalid_argument(("alignment " + std::to_string(alignment) + " is not a PhoneticAlignment").c_str(),
                            "alignment=" + std::to_string(alignment));
  }
  constexpr std::uint32_t kMaxFontIndex = 0xFFFFU;
  if (auto rc = check_limit(font_index, kMaxFontIndex + 1U, "font_index"); rc != 0) {
    return rc;
  }

  sheet.set_cell_phonetic_props(
      row, col,
      formulon::PhoneticProperties{static_cast<std::uint16_t>(font_index), static_cast<std::uint8_t>(type),
                                   static_cast<std::uint8_t>(alignment)});
  wb->workbook().mark_cell_dependents_dirty(sheet_index, row, col);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_cell_show_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                       uint32_t col, int32_t show) {
  clear_last_error();
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).set_cell_show_phonetic(row, col, show != 0);
  return 0;
}

namespace {

// Shared prologue of the phonetic getters: a non-NULL `out` (named `out_name`
// in the diagnostic; skipped when `out_name` is NULL), a valid sheet, then the
// cell at (row, col), left NULL when the cell does not exist.
fm_status_t lookup_phonetic_cell(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                 const void* out, const char* out_name, const formulon::Cell** cell) {
  if (out_name != nullptr && out == nullptr) {
    return null_argument(out_name);
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  *cell = wb->workbook().sheet(sheet_index).cell_at(row, col);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_cell_show_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                       uint32_t col, int32_t* out_show) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (out_show == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, out_show, "out_show", &cell); rc != 0) {
    return rc;
  }
  *out_show = cell != nullptr && cell->show_phonetic ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_show_phonetic_cells(const fm_workbook_t* wb, size_t sheet_index,
                                                        const fm_cell_rect_t* range, fm_cell_address_t* out,
                                                        size_t capacity, size_t* out_count) {
  clear_last_error();
  if (out_count != nullptr) {
    *out_count = 0U;
  }
  if (range == nullptr || out_count == nullptr || (out == nullptr && capacity > 0U)) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*range); rc != 0) {
    return rc;
  }
  const formulon::StyleRect rect{range->first_row, range->first_col, range->last_row, range->last_col};
  auto cells = wb->workbook().show_phonetic_cells(sheet_index, rect);
  if (!cells) {
    return set_last_error(cells.error());
  }
  const size_t total = cells.value().size();
  *out_count = total;
  for (size_t i = 0; i < std::min(capacity, total); ++i) {
    out[i] = {cells.value()[i].row, cells.value()[i].col};
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_cell_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                  uint32_t col, int32_t* out_present, const char** out_text) {
  clear_last_error();
  if (out_present == nullptr) {
    return null_argument("out_present");
  }
  if (out_text == nullptr) {
    return null_argument("out_text");
  }
  const formulon::Cell* cell = nullptr;
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, out_text, "out_text", &cell); rc != 0) {
    return rc;
  }
  *out_present = 0;
  *out_text = nullptr;
  if (cell == nullptr || cell->phonetic_runs.empty()) {
    return 0;
  }
  // The ABI surfaces a bare reading, so the runs' kana concatenate: the
  // span each covers is `PHONETIC`'s business, not a furigana field's.
  *out_present = 1;
  *out_text = publish_scratch(wb, formulon::flatten_phonetic(cell->phonetic_runs));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_cell_phonetic_run_count(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                            uint32_t col, size_t* out_count) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (out_count == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, out_count, "out_count", &cell); rc != 0) {
    return rc;
  }
  *out_count = cell == nullptr ? 0U : cell->phonetic_runs.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_cell_phonetic_run_at(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                         uint32_t col, size_t run_index, fm_phonetic_run_t* out) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, out, "out", &cell); rc != 0) {
    return rc;
  }
  const size_t run_count = cell == nullptr ? 0U : cell->phonetic_runs.size();
  if (auto rc = check_index(run_index, run_count, "run_index"); rc != 0) {
    return rc;
  }
  const formulon::PhoneticRun& run = cell->phonetic_runs[run_index];
  const std::string_view surface = cell->cached_value.is_text() ? cell->cached_value.as_text() : std::string_view();
  out->start = run.sb;
  out->end = run.eb;
  out->start_cp = utf16_units_to_code_points(surface, run.sb);
  out->end_cp = utf16_units_to_code_points(surface, run.eb);
  out->text = publish_scratch(wb, run.text);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_cell_phonetic_properties(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                             uint32_t col, fm_phonetic_properties_t* out) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, out, "out", &cell); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  // An absent cell reports the defaults rather than failing, matching the
  // run-count reader: "no annotation" and "no cell" are the same answer to
  // a host walking a range.
  const formulon::PhoneticProperties props = cell == nullptr ? formulon::PhoneticProperties{} : cell->phonetic_props;
  reset_record(out);
  out->font_index = props.font_id;
  out->type = props.type;
  out->alignment = props.alignment;
  out->stated = cell != nullptr && cell->phonetic_present ? 1 : 0;
  return 0;
}
