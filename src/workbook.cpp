// Copyright 2026 libraz. Licensed under the MIT License.
//
// Workbook implementation. The M1 surface is intentionally thin: a factory
// that builds a one-sheet workbook and a `save()` method that delegates to
// the OOXML writer slice.

#include "workbook.h"

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "io/ooxml_writer.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

Workbook Workbook::create() {
  Workbook wb;
  wb.sheets_.emplace_back(Sheet{std::string("Sheet1")});
  return wb;
}

Expected<std::vector<std::uint8_t>, Error> Workbook::save() const {
  return io::write_ooxml(*this);
}

}  // namespace formulon
