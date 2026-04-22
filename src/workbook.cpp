// Copyright 2026 libraz. Licensed under the MIT License.
//
// Workbook implementation. The current surface is intentionally thin: a
// factory that builds a one-sheet workbook and a `save()` method that
// delegates to the OOXML writer slice.

#include "workbook.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml_writer.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"

namespace formulon {

Workbook Workbook::create() {
  Workbook wb;
  wb.sheets_.emplace_back(Sheet{std::string("Sheet1")});
  return wb;
}

Sheet& Workbook::add_sheet(std::string name) {
  sheets_.emplace_back(Sheet{std::move(name)});
  return sheets_.back();
}

const Sheet* Workbook::sheet_by_name(std::string_view name) const noexcept {
  for (const Sheet& s : sheets_) {
    if (strings::case_insensitive_eq(s.name(), name)) {
      return &s;
    }
  }
  return nullptr;
}

Expected<std::vector<std::uint8_t>, Error> Workbook::save() const {
  return io::write_ooxml(*this);
}

}  // namespace formulon
