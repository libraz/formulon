//
// Style-record helpers shared by the C-ABI parts that write the font, fill and border tables
// (`styles_mutate.cpp`, `cell_format_patch.cpp`).

#ifndef FORMULON_C_API_PARTS_STYLES_RECORDS_H_
#define FORMULON_C_API_PARTS_STYLES_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "styles.h"

namespace formulon {
namespace c_api {
namespace parts {

// Dedup identity of a font / fill / border record: the XML fragment the styles writer emits for it.
std::string font_key(const formulon::FontRecord& value);
std::string fill_key(const formulon::FillRecord& value);
std::string border_key(const formulon::BorderRecord& value);

// Index of the record in `table` whose key equals `candidate`'s, appending `candidate` when none does. The
// candidate's key is built once, so the scan stays linear in the table size.
template <typename Record>
uint32_t find_or_append_record(std::vector<Record>& table, Record candidate, std::string (*key)(const Record&)) {
  const std::string candidate_key = key(candidate);
  for (std::size_t i = 0; i < table.size(); ++i) {
    if (key(table[i]) == candidate_key) {
      return static_cast<uint32_t>(i);
    }
  }
  table.push_back(std::move(candidate));
  return static_cast<uint32_t>(table.size() - 1U);
}

// A colour specification from the caller. An out-of-range `kind` degrades to `kNone`; `out_argb` receives the
// record's legacy ARGB field.
formulon::ColorSpec color_from_c(fm_color_t src, std::uint32_t* out_argb);

// The engine fill of a validated `fm_fill_t`; a gradient fill carries no pattern fields.
formulon::FillRecord fill_from_c(const fm_fill_t& record);

// Seeds the reserved index-0 font, fill and border (and cell xf) of a table a package without `xl/styles.xml` left
// empty.
void ensure_default_style_roots(formulon::StylesTable& styles);
void ensure_default_cell_xf(formulon::StylesTable& styles);

// A stated font size is `1..409` points: 0 when `size_pt` is in range, else `FM_STATUS_INVALID_ARGUMENT`.
fm_status_t check_font_size_pt(double size_pt);

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_STYLES_RECORDS_H_
