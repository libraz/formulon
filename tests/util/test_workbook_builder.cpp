//
// Implementation of the fluent WorkbookBuilder declared in
// `test_workbook_builder.h`.

#include "util/test_workbook_builder.h"

#include <gtest/gtest.h>

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>

#include "io/a1_ref.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace test {

namespace {

constexpr std::string_view kDefaultSheetName = "Sheet1";

bool StartsWithEquals(std::string_view s) noexcept {
  return !s.empty() && s.front() == '=';
}

}  // namespace

WorkbookBuilder::WorkbookBuilder() : wb_(Workbook::create_empty()), active_sheet_index_(0), has_active_sheet_(false) {}

WorkbookBuilder& WorkbookBuilder::sheet(std::string_view name) {
  const std::size_t existing = wb_.sheet_index_by_name(name);
  if (existing != static_cast<std::size_t>(-1)) {
    active_sheet_index_ = existing;
  } else {
    wb_.add_sheet(std::string(name));
    active_sheet_index_ = wb_.sheet_count() - 1U;
  }
  has_active_sheet_ = true;
  return *this;
}

WorkbookBuilder& WorkbookBuilder::cell(std::string_view a1, double value) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  if (!ParseA1(a1, &row, &col)) {
    return *this;
  }
  active_sheet().set_cell_value(row, col, Value::number(value));
  return *this;
}

WorkbookBuilder& WorkbookBuilder::cell(std::string_view a1, bool value) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  if (!ParseA1(a1, &row, &col)) {
    return *this;
  }
  active_sheet().set_cell_value(row, col, Value::boolean(value));
  return *this;
}

WorkbookBuilder& WorkbookBuilder::cell(std::string_view a1, std::string_view text_or_formula) {
  if (StartsWithEquals(text_or_formula)) {
    return formula_cell(a1, text_or_formula);
  }
  return text_cell(a1, text_or_formula);
}

WorkbookBuilder& WorkbookBuilder::text_cell(std::string_view a1, std::string_view text) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  if (!ParseA1(a1, &row, &col)) {
    return *this;
  }
  // Sheet::set_cell_value deep-copies text payloads into the cell's
  // internal storage (see sheet.h `update_cached_value` semantics),
  // so the source string_view does NOT need to outlive this call.
  active_sheet().set_cell_value(row, col, Value::text(text));
  return *this;
}

WorkbookBuilder& WorkbookBuilder::formula_cell(std::string_view a1, std::string_view formula) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  if (!ParseA1(a1, &row, &col)) {
    return *this;
  }
  active_sheet().set_cell_formula(row, col, std::string(formula));
  return *this;
}

WorkbookBuilder& WorkbookBuilder::cell(std::string_view a1, Value value) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  if (!ParseA1(a1, &row, &col)) {
    return *this;
  }
  active_sheet().set_cell_value(row, col, value);
  return *this;
}

Workbook WorkbookBuilder::build() {
  return std::move(wb_);
}

bool WorkbookBuilder::ParseA1(std::string_view a1, std::uint32_t* out_row, std::uint32_t* out_col) {
  if (!io::parse_a1_ref(a1, out_row, out_col)) {
    ADD_FAILURE() << "WorkbookBuilder: malformed A1 reference \"" << a1 << "\"";
    return false;
  }
  return true;
}

Sheet& WorkbookBuilder::active_sheet() {
  if (!has_active_sheet_) {
    wb_.add_sheet(std::string(kDefaultSheetName));
    active_sheet_index_ = wb_.sheet_count() - 1U;
    has_active_sheet_ = true;
  }
  return wb_.sheet(active_sheet_index_);
}

}  // namespace test
}  // namespace formulon
