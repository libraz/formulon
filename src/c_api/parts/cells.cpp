//
// C ABI - cell mutation (set_*), cell read (get_value, lambda_text_at),
// flat-enumeration (cell_count, cell_at), and dynamic-array spill info.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "c_api/formulon_c.h"
#include "c_api/parts/cell_range.h"
#include "c_api/parts/common.h"
#include "cell.h"
#include "eval/lambda_format.h"
#include "eval/lambda_value.h"
#include "parser/ast_format_r1c1.h"
#include "parser/parser.h"
#include "phonetic.h"
#include "sheet.h"
#include "style_resolve.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/double_format.h"
#include "utils/error.h"
#include "utils/formula_text.h"
#include "utils/index_sort.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::cell_range_page_size;
using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_cell_cursor;
using formulon::c_api::parts::check_entered_cell_formula;
using formulon::c_api::parts::check_finite;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_sheet_u32;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::display_formula;
using formulon::c_api::parts::display_formula_text;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::table_display_name_at;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::value_to_fm;

// ---------------------------------------------------------------------------
// Cell mutation
// ---------------------------------------------------------------------------

namespace {

// Writes an entry into (`row`, `col`) the way Excel stores one: a cell that
// did not own a style takes its row's style, else its column's.
template <typename Write>
fm_status_t enter_cell(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col, Write&& write) {
  formulon::Workbook& book = wb->workbook();
  const formulon::EffectiveXf inherited = formulon::select_effective_xf(book.sheet(sheet_index), row, col);
  auto r = write(book);
  if (!r) {
    return set_last_error(r.error());
  }
  if (inherited.source == formulon::StyleSource::kRow || inherited.source == formulon::StyleSource::kColumn) {
    auto styled = book.set_cell_xf_index(sheet_index, row, col, inherited.xf_index);
    if (!styled) {
      return set_last_error(styled.error());
    }
  }
  return 0;
}

// Table and column whose header cell is (`row`, `col`).
struct HeaderCell {
  std::size_t table;
  std::size_t column;
};

std::optional<HeaderCell> table_header_cell(const formulon::Workbook& book, size_t sheet_index, uint32_t row,
                                            uint32_t col) {
  const std::optional<std::size_t> table = book.table_containing(sheet_index, row, col);
  if (!table) {
    return std::nullopt;
  }
  const formulon::TableMetadata& meta = book.tables()[*table];
  const std::optional<formulon::MergeRange> rect = formulon::parse_a1_rectangle(meta.ref);
  if (!meta.header_row || !rect || row != rect->first_row) {
    return std::nullopt;
  }
  return HeaderCell{*table, static_cast<std::size_t>(col - rect->first_col)};
}

// A value entered into a table's header cell renames the column to its text,
// as Excel does; the header cell then holds the final name.
fm_status_t rename_from_header(fm_workbook_t* wb, const HeaderCell& header, std::string_view text) {
  auto renamed = wb->workbook().rename_table_column(header.table, header.column, text);
  return renamed ? 0 : set_last_error(renamed.error());
}

}  // namespace

extern "C" fm_status_t fm_workbook_set_number(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              double value) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_number"); rc != 0) {
    return rc;
  }
  if (auto rc = check_finite(value, "fm_workbook_set_number", "value"); rc != 0) {
    return rc;
  }
  if (const auto header = table_header_cell(wb->workbook(), sheet_index, row, col)) {
    std::string text;
    formulon::format_double(text, value);
    return rename_from_header(wb, *header, text);
  }
  return enter_cell(wb, sheet_index, row, col, [&](formulon::Workbook& book) {
    return book.set_cell_value(sheet_index, row, col, formulon::Value::number(value));
  });
}

extern "C" fm_status_t fm_workbook_set_bool(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                            int32_t value) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_bool"); rc != 0) {
    return rc;
  }
  if (const auto header = table_header_cell(wb->workbook(), sheet_index, row, col)) {
    return rename_from_header(wb, *header, value != 0 ? "TRUE" : "FALSE");
  }
  return enter_cell(wb, sheet_index, row, col, [&](formulon::Workbook& book) {
    return book.set_cell_value(sheet_index, row, col, formulon::Value::boolean(value != 0));
  });
}

extern "C" fm_status_t fm_workbook_set_error(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             fm_error_code_t error) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_error"); rc != 0) {
    return rc;
  }
  if (error < 0 || error > static_cast<fm_error_code_t>(formulon::ErrorCode::Unknown)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_error: error code out of range", "error=" + std::to_string(error));
  }
  if (const auto header = table_header_cell(wb->workbook(), sheet_index, row, col)) {
    return rename_from_header(wb, *header, formulon::display_name(static_cast<formulon::ErrorCode>(error)));
  }
  return enter_cell(wb, sheet_index, row, col, [&](formulon::Workbook& book) {
    return book.set_cell_value(sheet_index, row, col, formulon::Value::error(static_cast<formulon::ErrorCode>(error)));
  });
}

extern "C" fm_status_t fm_workbook_set_text(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                            const char* utf8) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_text"); rc != 0) {
    return rc;
  }
  if (utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_set_text: utf8 is NULL");
  }
  if (auto rc = check_utf8("fm_workbook_set_text", {utf8}); rc != 0) {
    return rc;
  }
  if (const auto header = table_header_cell(wb->workbook(), sheet_index, row, col)) {
    return rename_from_header(wb, *header, utf8);
  }
  // Copy directly into the destination cell. Keeping every overwritten
  // input in a handle-global deque leaked one string per C-ABI write.
  return enter_cell(wb, sheet_index, row, col,
                    [&](formulon::Workbook& book) { return book.set_cell_text(sheet_index, row, col, utf8); });
}

extern "C" fm_status_t fm_workbook_set_cell_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                     const char* utf8) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_cell_phonetic"); rc != 0) {
    return rc;
  }
  if (utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_cell_phonetic: utf8 is NULL");
  }
  if (auto rc = check_utf8("fm_workbook_set_cell_phonetic", {utf8}); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col, "fm_workbook_set_cell_phonetic"); rc != 0) {
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

extern "C" fm_status_t fm_workbook_set_cell_phonetic_runs(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                          uint32_t col, const fm_phonetic_run_t* runs, size_t count) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_cell_phonetic_runs"); rc != 0) {
    return rc;
  }
  if (runs == nullptr && count != 0U) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_cell_phonetic_runs: runs is NULL");
  }
  if (auto rc = check_cell_coord(row, col, "fm_workbook_set_cell_phonetic_runs"); rc != 0) {
    return rc;
  }

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
      return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                               "fm_workbook_set_cell_phonetic_runs: run text is NULL",
                               "run_index=" + std::to_string(i));
    }
    if (auto rc = check_utf8(run.text, "fm_workbook_set_cell_phonetic_runs", "run_index", i); rc != 0) {
      return rc;
    }
    if (run.eb < run.sb) {
      return set_binding_error(
          formulon::FormulonErrorCode::kInvalidArgument,
          "fm_workbook_set_cell_phonetic_runs: run ends before it starts",
          "run_index=" + std::to_string(i) + " sb=" + std::to_string(run.sb) + " eb=" + std::to_string(run.eb));
    }
    if (i != 0U && run.sb < previous_end) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_set_cell_phonetic_runs: runs overlap or are out of order",
                               "run_index=" + std::to_string(i) + " sb=" + std::to_string(run.sb) +
                                   " previous_eb=" + std::to_string(previous_end));
    }
    previous_end = run.eb;
    parsed.push_back(formulon::PhoneticRun{run.sb, run.eb, std::string(run.text)});
  }

  wb->workbook().sheet(sheet_index).set_cell_phonetic_runs(row, col, std::move(parsed));
  wb->workbook().mark_cell_dependents_dirty(sheet_index, row, col);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_cell_phonetic_properties(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                                uint32_t col, uint32_t font_id, uint32_t type,
                                                                uint32_t alignment) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_cell_phonetic_properties"); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col, "fm_workbook_set_cell_phonetic_properties"); rc != 0) {
    return rc;
  }
  // Both enumerations occupy two bits of the XLSB trailer's flags word, so
  // an out-of-range value would not survive an `.xlsb` save intact. Reject
  // it here rather than truncate it into a different, valid-looking form.
  constexpr uint32_t kMaxOrdinal = 3U;
  if (type > kMaxOrdinal) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_cell_phonetic_properties: type out of range",
                             "type=" + std::to_string(type));
  }
  if (alignment > kMaxOrdinal) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_cell_phonetic_properties: alignment out of range",
                             "alignment=" + std::to_string(alignment));
  }
  constexpr uint32_t kMaxFontId = 0xFFFFU;
  if (font_id > kMaxFontId) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_cell_phonetic_properties: font_id out of range",
                             "font_id=" + std::to_string(font_id));
  }

  wb->workbook()
      .sheet(sheet_index)
      .set_cell_phonetic_props(
          row, col,
          formulon::PhoneticProperties{static_cast<std::uint16_t>(font_id), static_cast<std::uint8_t>(type),
                                       static_cast<std::uint8_t>(alignment)});
  wb->workbook().mark_cell_dependents_dirty(sheet_index, row, col);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_cell_show_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                          uint32_t col, int32_t show) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_cell_show_phonetic"); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col, "fm_workbook_set_cell_show_phonetic"); rc != 0) {
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
                                 const char* fn, const void* out, const char* out_name, const formulon::Cell** cell) {
  if (out_name != nullptr && out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(fn) + ": " + out_name + " is NULL").c_str());
  }
  if (auto rc = check_sheet_index(wb, sheet_index, fn); rc != 0) {
    return rc;
  }
  *cell = wb->workbook().sheet(sheet_index).cell_at(row, col);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_get_cell_show_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                          uint32_t col, int32_t* out_show) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, "fm_workbook_get_cell_show_phonetic", out_show,
                                     "out_show", &cell);
      rc != 0) {
    return rc;
  }
  *out_show = cell != nullptr && cell->show_phonetic ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_workbook_has_cell_phonetic_properties(const fm_workbook_t* wb, size_t sheet_index,
                                                                uint32_t row, uint32_t col, int32_t* out_present) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, "fm_workbook_has_cell_phonetic_properties", out_present,
                                     "out_present", &cell);
      rc != 0) {
    return rc;
  }
  *out_present = cell != nullptr && cell->phonetic_present ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_workbook_set_blank(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_blank"); rc != 0) {
    return rc;
  }
  if (const auto header = table_header_cell(wb->workbook(), sheet_index, row, col)) {
    return rename_from_header(wb, *header, std::string_view());
  }
  return finish_status(wb->workbook().set_cell_value(sheet_index, row, col, formulon::Value::blank()));
}

extern "C" fm_status_t fm_workbook_set_formula(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                               const char* formula) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_set_formula"); rc != 0) {
    return rc;
  }
  if (formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_formula: formula is NULL");
  }
  std::string bound;
  if (auto rc = check_entered_cell_formula("fm_workbook_set_formula", wb, sheet_index, row, col, formula,
                                           "row=" + std::to_string(row) + " col=" + std::to_string(col), &bound);
      rc != 0) {
    return rc;
  }
  return enter_cell(wb, sheet_index, row, col, [&](formulon::Workbook& book) {
    return book.set_cell_formula(sheet_index, row, col, std::move(bound));
  });
}

// ---------------------------------------------------------------------------
// Cell read
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_get_value(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             fm_value_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_get_value: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_get_value"); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col, "fm_workbook_get_value"); rc != 0) {
    return rc;
  }
  // `resolve_cell_value` is the spill-aware accessor: phantoms of a
  // dynamic-array spill surface their array cell rather than the raw
  // (blank) cached value.
  const formulon::Value v = wb->workbook().sheet(sheet_index).resolve_cell_value(row, col);
  // The returned text pointer is valid until the next read on this handle
  // (see `formulon_c.h`).
  TextStore& store = fresh_read_scratch(wb);
  value_to_fm(v, store, out);
  return 0;
}

extern "C" fm_status_t fm_workbook_get_cell_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                     uint32_t col, const char** out_text) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (auto rc =
          lookup_phonetic_cell(wb, sheet_index, row, col, "fm_workbook_get_cell_phonetic", out_text, "out_text", &cell);
      rc != 0) {
    return rc;
  }
  TextStore& store = fresh_read_scratch(wb);
  // The ABI surfaces a bare reading, so the runs' kana concatenate: the
  // span each covers is `PHONETIC`'s business, not a furigana field's.
  store.emplace_back(cell == nullptr ? std::string() : formulon::flatten_phonetic(cell->phonetic_runs));
  *out_text = store.back().c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_cell_phonetic_run_count(const fm_workbook_t* wb, size_t sheet_index,
                                                               uint32_t row, uint32_t col, uint32_t* out_count) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, "fm_workbook_get_cell_phonetic_run_count", out_count,
                                     "out_count", &cell);
      rc != 0) {
    return rc;
  }
  *out_count = cell == nullptr ? 0U : static_cast<uint32_t>(cell->phonetic_runs.size());
  return 0;
}

extern "C" fm_status_t fm_workbook_get_cell_phonetic_run(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                         uint32_t col, uint32_t run_index, fm_phonetic_run_t* out) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, "fm_workbook_get_cell_phonetic_run", out, "out", &cell);
      rc != 0) {
    return rc;
  }
  const size_t run_count = cell == nullptr ? 0U : cell->phonetic_runs.size();
  if (auto rc = check_index(run_index, run_count, "fm_workbook_get_cell_phonetic_run", "run_index"); rc != 0) {
    return rc;
  }
  const formulon::PhoneticRun& run = cell->phonetic_runs[run_index];
  TextStore& store = fresh_read_scratch(wb);
  store.emplace_back(run.text);
  out->sb = run.sb;
  out->eb = run.eb;
  out->text = store.back().c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_cell_phonetic_properties(const fm_workbook_t* wb, size_t sheet_index,
                                                                uint32_t row, uint32_t col, uint32_t* out_font_id,
                                                                uint32_t* out_type, uint32_t* out_alignment) {
  clear_last_error();
  const formulon::Cell* cell = nullptr;
  if (auto rc = lookup_phonetic_cell(wb, sheet_index, row, col, "fm_workbook_get_cell_phonetic_properties", nullptr,
                                     nullptr, &cell);
      rc != 0) {
    return rc;
  }
  // An absent cell reports the defaults rather than failing, matching the
  // run-count reader: "no annotation" and "no cell" are the same answer to
  // a host walking a range.
  const formulon::PhoneticProperties props = cell == nullptr ? formulon::PhoneticProperties{} : cell->phonetic_props;
  if (out_font_id != nullptr) {
    *out_font_id = props.font_id;
  }
  if (out_type != nullptr) {
    *out_type = props.type;
  }
  if (out_alignment != nullptr) {
    *out_alignment = props.alignment;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_lambda_text_at(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                  const char** out_text) {
  clear_last_error();
  if (out_text == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_lambda_text_at: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_lambda_text_at"); rc != 0) {
    return rc;
  }
  const formulon::Cell* cell = wb->workbook().sheet(sheet_index).cell_at(row, col);
  if (cell == nullptr || !cell->cached_value.is_lambda()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_lambda_text_at: cell does not hold a lambda value",
        "sheet_index=" + std::to_string(sheet_index) + " row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  const formulon::eval::LambdaValue* lv = cell->cached_value.as_lambda();
  if (lv == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_lambda_text_at: lambda payload is NULL");
  }
  std::string formatted = wb->workbook().spell_self_book_names(formulon::eval::format_lambda_value(*lv));
  // Read-path scratch: cleared per call so the lambda text does not
  // accumulate across repeated reads. The returned pointer is valid until
  // the next read on this handle (see `formulon_c.h`).
  TextStore& store = fresh_read_scratch(wb);
  store.emplace_back(std::move(formatted));
  *out_text = store.back().c_str();
  return 0;
}

// ---------------------------------------------------------------------------
// Iteration / dump
// ---------------------------------------------------------------------------
//
// Sheets are stored row-sparse (`unordered_map<row, RowCells>`); we
// surface a flat enumeration to bindings by materialising a sorted
// `(row, col)` index on demand. The cache is purely an optimisation;
// correctness does not depend on it.

namespace {

// Returns the `(row, col)` indices of every written coordinate on `sheet`,
// sorted by `(row, col)` ascending. Cleared cells keep their slot and stay
// listed (the CLI filters blanks at render time); unwritten columns are not
// enumerated. This matches `Sheet::cell_count`.
std::vector<std::pair<std::uint32_t, std::uint32_t>> collect_cell_addresses(const formulon::Sheet& sheet) {
  std::vector<std::pair<std::uint32_t, std::uint32_t>> out;
  for (const auto& [row, cells] : sheet.rows()) {
    for (const auto& [col, cell] : cells) {
      (void)cell;
      out.emplace_back(row, col);
    }
  }
  // Dynamic-array spill phantoms carry an effective value through
  // `resolve_cell_value` but live only in the spill table, absent from
  // `rows()`. Merge them so the flat enumeration surfaces spilled cells, then
  // sort + unique: a phantom that coincides with a stored slot must appear
  // only once.
  //
  // Streamed rather than collected first: a phantom set is a spill
  // rectangle's whole area, so `=SEQUENCE(1048576)` would otherwise build an
  // eight-megabyte vector only to append it to this one.
  sheet.for_each_spill_phantom(
      [](formulon::CellAddress address, void* ctx) {
        static_cast<std::vector<std::pair<std::uint32_t, std::uint32_t>>*>(ctx)->emplace_back(address.row, address.col);
      },
      &out);
  formulon::sort_by_index(out, [](const std::pair<std::uint32_t, std::uint32_t>& lhs,
                                  const std::pair<std::uint32_t, std::uint32_t>& rhs) { return lhs < rhs; });
  out.erase(std::unique(out.begin(), out.end()), out.end());
  return out;
}

const std::vector<std::pair<std::uint32_t, std::uint32_t>>& cached_cell_addresses(const fm_workbook_t* wb,
                                                                                  std::size_t sheet_index) {
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  const std::uint64_t revision = sheet.cell_enumeration_revision();
  auto& caches = wb->cell_enumeration_cache;

  // Reuse an entry for this sheet even when its revision changed: rebuilding
  // in place preserves the other sheet's warm entry. Promote a slot-1 hit so
  // the next lookup follows the hot path.
  for (std::size_t i = 0; i < caches.size(); ++i) {
    auto& cache = caches[i];
    if (cache.sheet_index != sheet_index) {
      continue;
    }
    if (cache.revision != revision) {
      cache.addresses = collect_cell_addresses(sheet);
      cache.revision = revision;
    }
    if (i != 0U) {
      std::swap(caches[0], cache);
    }
    return caches[0].addresses;
  }

  // Collect before evicting the least-recently-used entry.
  std::vector<std::pair<std::uint32_t, std::uint32_t>> addresses = collect_cell_addresses(sheet);
  caches[1] = std::move(caches[0]);
  caches[0].addresses = std::move(addresses);
  caches[0].sheet_index = sheet_index;
  caches[0].revision = revision;
  return caches[0].addresses;
}

}  // namespace

// `fm_workbook_cell_count` is now emitted by the binding codegen (see
// `src/c_api/generated/sheet_counts.cpp`).

extern "C" fm_status_t fm_workbook_cell_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, uint32_t* out_row,
                                           uint32_t* out_col, const char** out_formula, fm_value_t* out_value) {
  clear_last_error();
  if (out_row == nullptr || out_col == nullptr || out_value == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_cell_at: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_cell_at"); rc != 0) {
    return rc;
  }
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  // The sorted address vector is cached on the handle until the Sheet's
  // enumeration revision changes. A `cell_count` / `cell_at(0..N)` pass
  // therefore pays the O(N log N) collection once rather than per cell.
  const auto& addrs = cached_cell_addresses(wb, sheet_index);
  if (auto rc = check_index(idx, addrs.size(), "fm_workbook_cell_at", "idx"); rc != 0) {
    return rc;
  }
  const auto [row, col] = addrs[idx];
  *out_row = row;
  *out_col = col;
  // `(row, col)` may be a spill phantom, which has no stored `Cell`: that is
  // a normal case here, not an error. Phantoms carry no formula text and
  // their value is resolved below via `resolve_cell_value`. Only a stored
  // cell can contribute a formula pointer.
  const formulon::Cell* cell = sheet.cell_at(row, col);
  const formulon::DataTable* table =
      cell != nullptr && !cell->formula_text.empty() ? nullptr : sheet.data_table_covering(row, col);
  // `out_value`'s text payload goes to `read_scratch` (cleared per call,
  // valid until the next read), as does a formula the workbook respells;
  // any other formula points into the cell's own `formula_text`.
  TextStore& store = fresh_read_scratch(wb);
  if (out_formula != nullptr) {
    *out_formula = (cell != nullptr && !cell->formula_text.empty())
                       ? display_formula(wb->workbook(), cell->formula_text, store,
                                         table_display_name_at(wb->workbook(), sheet_index, row, col))
                       : nullptr;
  }
  // Use the spill-aware accessor so phantoms surface their owning anchor's
  // value. The phantoms themselves are still indexed via `cell_at`'s
  // implicit default cells; users that want only stored formulae filter
  // by `out_formula != NULL`.
  const formulon::Value v = sheet.resolve_cell_value(row, col);
  if (out_formula != nullptr && table != nullptr) {
    *out_formula = store.emplace_back(formulon::data_table_formula_text(*table)).c_str();
  }
  value_to_fm(v, store, out_value);
  return 0;
}

// ---------------------------------------------------------------------------
// Dynamic-array spill payload
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_spill_info(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                              std::uint32_t col, fm_spill_info_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_spill_info: NULL argument");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_workbook_spill_info"); rc != 0) {
    return rc;
  }
  *out = fm_spill_info_t{};
  const auto& s = wb->workbook().sheet(sheet);
  // Try anchor lookup first; fall back to phantom-coverage map.
  const formulon::SpillRegion* region = s.spill_region_at_anchor(row, col);
  if (region == nullptr) {
    region = s.spill_region_covering(row, col);
  }
  if (region == nullptr) {
    // Only a footprint that fits the grid is reported: an overflowing one is
    // recorded on some spill paths and not others.
    const auto blocked = s.blocked_spill_footprint_at(row, col);
    if (blocked && static_cast<std::uint64_t>(blocked->anchor_row) + blocked->rows <= formulon::Sheet::kMaxRows &&
        static_cast<std::uint64_t>(blocked->anchor_col) + blocked->cols <= formulon::Sheet::kMaxCols) {
      out->anchor_row = blocked->anchor_row;
      out->anchor_col = blocked->anchor_col;
      out->rows = blocked->rows;
      out->cols = blocked->cols;
      out->blocked = 1;
    }
    return 0;
  }
  out->anchor_row = region->anchor_row;
  out->anchor_col = region->anchor_col;
  out->rows = region->rows;
  out->cols = region->cols;
  out->engaged = 1;
  return 0;
}

// ---------------------------------------------------------------------------
// Legacy array (CSE) formulas
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_get_array_formula(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                                     std::uint32_t col, fm_array_formula_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_get_array_formula: NULL argument");
  }
  if (auto rc = check_sheet_u32(wb, sheet, "fm_workbook_get_array_formula"); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col, "fm_workbook_get_array_formula"); rc != 0) {
    return rc;
  }
  *out = fm_array_formula_t{};
  const std::optional<formulon::Workbook::ArrayFormulaInfo> info = wb->workbook().array_formula_at(sheet, row, col);
  if (!info.has_value()) {
    return 0;
  }
  out->anchor_row = info->anchor_row;
  out->anchor_col = info->anchor_col;
  out->rows = info->rows;
  out->cols = info->cols;
  out->formula = publish_scratch(
      wb, display_formula_text(wb->workbook(), info->formula,
                               table_display_name_at(wb->workbook(), sheet, info->anchor_row, info->anchor_col)));
  out->engaged = 1;
  return 0;
}

extern "C" fm_status_t fm_workbook_set_array_formula(fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                                     std::uint32_t col, std::uint32_t rows, std::uint32_t cols,
                                                     const char* formula) {
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, "fm_workbook_set_array_formula"); rc != 0) {
    return rc;
  }
  if (formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_array_formula: formula is NULL");
  }
  if (auto rc = check_cell_coord(row, col, "fm_workbook_set_array_formula"); rc != 0) {
    return rc;
  }
  std::string bound;
  if (auto rc = check_entered_cell_formula("fm_workbook_set_array_formula", wb, sheet, row, col, formula,
                                           "row=" + std::to_string(row) + " col=" + std::to_string(col), &bound);
      rc != 0) {
    return rc;
  }
  return enter_cell(wb, sheet, row, col, [&](formulon::Workbook& book) {
    return book.set_array_formula(sheet, row, col, rows, cols, std::move(bound));
  });
}

// ---------------------------------------------------------------------------
// Formula text and range enumeration
// ---------------------------------------------------------------------------

namespace {

// Shared front half of the two formula getters: validates, then copies the
// stored formula text (empty for a non-formula) out under the sheet lock.
fm_status_t read_stored_formula(const fm_workbook_t* wb, size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                const char** out_formula, const char* api, std::string* out_text) {
  if (out_formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(api) + ": NULL argument").c_str());
  }
  if (auto rc = check_sheet_index(wb, sheet_index, api); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col, api); rc != 0) {
    return rc;
  }
  formulon::Sheet::CellRead read;
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  sheet.read_formula_cell(row, col, read);
  out_text->assign(read.formula_text());
  if (const formulon::DataTable* table = sheet.data_table_covering(row, col); out_text->empty() && table != nullptr) {
    *out_text = formulon::data_table_formula_text(*table);
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_get_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                               const char** out_formula) {
  clear_last_error();
  std::string text;
  if (auto rc = read_stored_formula(wb, sheet_index, row, col, out_formula, "fm_workbook_get_formula", &text);
      rc != 0) {
    return rc;
  }
  *out_formula = publish_scratch(
      wb, display_formula_text(wb->workbook(), text, table_display_name_at(wb->workbook(), sheet_index, row, col)));
  return 0;
}

extern "C" fm_status_t fm_workbook_get_formula_r1c1(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                    uint32_t col, const char** out_formula) {
  clear_last_error();
  std::string text;
  if (auto rc = read_stored_formula(wb, sheet_index, row, col, out_formula, "fm_workbook_get_formula_r1c1", &text);
      rc != 0) {
    return rc;
  }
  if (text.empty()) {
    *out_formula = publish_scratch(wb, std::string());
    return 0;
  }
  formulon::Arena arena;
  const formulon::parser::AstNode* root = formulon::parser::parse_strict(formulon::strip_formula_prefix(text), arena);
  if (root == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kParserUnexpectedToken,
                             "fm_workbook_get_formula_r1c1: stored formula does not parse",
                             "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  const formulon::parser::StructuredRefDisplay display{table_display_name_at(wb->workbook(), sheet_index, row, col)};
  *out_formula = publish_scratch(
      wb, wb->workbook().spell_self_book_names(formulon::parser::format_formula_r1c1(*root, row, col, &display)));
  return 0;
}

extern "C" fm_status_t fm_sheet_cells_in_range(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                               uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                               uint64_t cursor, uint32_t limit, fm_cell_range_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cells_in_range: NULL out");
  }
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_cells_in_range"); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(first_row, first_col, last_row, last_col, "fm_sheet_cells_in_range"); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_cursor(cursor, "fm_sheet_cells_in_range"); rc != 0) {
    return rc;
  }
  auto handle = std::unique_ptr<fm_cell_range_t>(new fm_cell_range_t{});
  handle->next_cursor = wb->workbook()
                            .sheet(sheet_index)
                            .cells_in_range(
                                first_row, first_col, last_row, last_col, cursor, cell_range_page_size(limit),
                                [](const formulon::Sheet::RangeCell& cell, void* ctx) {
                                  static_cast<fm_cell_range_t*>(ctx)->append(cell);
                                },
                                handle.get());
  // A data table's result cells show its `TABLE()` formula, as `fm_workbook_get_formula` does.
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  for (fm_cell_range_t::Entry& entry : handle->entries) {
    if (const formulon::DataTable* table = sheet.data_table_covering(entry.row, entry.col);
        entry.formula == nullptr && table != nullptr) {
      entry.formula = handle->text.emplace_back(formulon::data_table_formula_text(*table)).c_str();
    }
  }
  handle->spell_formulas(wb->workbook(), sheet_index);
  *out = handle.release();
  return 0;
}

extern "C" fm_status_t fm_cell_range_count(const fm_cell_range_t* range, size_t* out_count) {
  clear_last_error();
  if (range == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cell_range_count: NULL argument");
  }
  *out_count = range->entries.size();
  return 0;
}

extern "C" fm_status_t fm_cell_range_at(const fm_cell_range_t* range, size_t idx, uint32_t* out_row, uint32_t* out_col,
                                        const char** out_formula, fm_value_t* out_value) {
  clear_last_error();
  if (range == nullptr || out_row == nullptr || out_col == nullptr || out_value == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cell_range_at: NULL argument");
  }
  if (auto rc = check_index(idx, range->entries.size(), "fm_cell_range_at", "idx"); rc != 0) {
    return rc;
  }
  const fm_cell_range_t::Entry& entry = range->entries[idx];
  *out_row = entry.row;
  *out_col = entry.col;
  if (out_formula != nullptr) {
    *out_formula = entry.formula;
  }
  *out_value = entry.value;
  return 0;
}

extern "C" fm_status_t fm_cell_range_next_cursor(const fm_cell_range_t* range, uint64_t* out_cursor) {
  clear_last_error();
  if (range == nullptr || out_cursor == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_cell_range_next_cursor: NULL argument");
  }
  *out_cursor = range->next_cursor;
  return 0;
}

extern "C" void fm_cell_range_destroy(fm_cell_range_t* range) {
  delete range;
}
