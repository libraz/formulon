// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Fluent builder for populating small `Workbook` fixtures inside unit
// tests.
//
// Typical usage:
//
//   Workbook wb = WorkbookBuilder()
//                     .sheet("Sheet1")
//                     .cell("A1", 42.0)
//                     .cell("B1", "=A1+1")
//                     .cell("C1", "hello")
//                     .cell("D1", true)
//                     .build();
//
// Semantics:
//   * `sheet(name)` selects (or creates) the sheet that subsequent
//     `cell(...)` calls write to. The most recent `sheet()` call is
//     the active sheet.
//   * `cell(a1, double)` stores a numeric literal.
//   * `cell(a1, bool)` stores a boolean literal.
//   * `cell(a1, string_view)` is overloaded by content: a string that
//     starts with `=` is stored as a formula (via
//     `Sheet::set_cell_formula`), otherwise as a text literal.
//   * Calls return `*this` so they chain.
//   * `build()` returns the populated workbook by value (Workbook is
//     move-only). The builder is single-use after `build()` runs.
//
// The builder intentionally returns a `Workbook` without driving
// `recalc()`: tests that need cached formula values should call
// `wb.recalc(eval::default_registry())` themselves. This keeps the
// builder usable in tests that exercise the un-evaluated formula
// surface (parser, dep-graph, save/load).

#ifndef FORMULON_TESTS_UTIL_TEST_WORKBOOK_BUILDER_H_
#define FORMULON_TESTS_UTIL_TEST_WORKBOOK_BUILDER_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "value.h"
#include "workbook.h"

namespace formulon {
namespace test {

class WorkbookBuilder {
 public:
  // Constructs an empty builder. Internally allocates an empty
  // workbook; the first `sheet()` call appends its first sheet. If a
  // caller writes a `cell()` without an explicit `sheet()` first, the
  // builder creates a default `"Sheet1"` on demand.
  WorkbookBuilder();

  WorkbookBuilder(const WorkbookBuilder&) = delete;
  WorkbookBuilder& operator=(const WorkbookBuilder&) = delete;
  WorkbookBuilder(WorkbookBuilder&&) noexcept = default;
  WorkbookBuilder& operator=(WorkbookBuilder&&) noexcept = default;

  // Selects (or appends) the sheet with display name `name` and marks
  // it as the active target for subsequent `cell(...)` calls.
  // Case-insensitive match, matching Excel and `Workbook::sheet_by_name`.
  WorkbookBuilder& sheet(std::string_view name);

  // Stores a numeric literal at A1-notation address `a1` on the active
  // sheet. Triggers `EXPECT_NONFATAL_FAILURE` (via a gtest ADD_FAILURE)
  // when `a1` cannot be parsed; the cell is then silently skipped so
  // chaining can continue.
  WorkbookBuilder& cell(std::string_view a1, double value);

  // Stores a boolean literal at A1-notation address `a1`. Overloaded
  // explicitly so `cell("A1", true)` does NOT silently convert to the
  // `string_view` overload via the `const char*` -> `bool` promotion.
  WorkbookBuilder& cell(std::string_view a1, bool value);

  // Stores either a formula or a text literal at A1-notation address
  // `a1`, depending on the leading character:
  //   * `text_or_formula.front() == '='` -> stored as formula
  //     (`Sheet::set_cell_formula`).
  //   * otherwise stored as a Text value (`Sheet::set_cell_value`).
  // Empty `text_or_formula` stores an empty Text payload, not Blank.
  WorkbookBuilder& cell(std::string_view a1, std::string_view text_or_formula);

  // Explicit text-only overload. Identical to passing a `string_view`
  // that does not start with `=`, but unambiguous at call sites where
  // the literal might begin with `=`.
  WorkbookBuilder& text_cell(std::string_view a1, std::string_view text);

  // Explicit formula overload. Stores `formula` (with or without the
  // leading `=`) as a formula via `Sheet::set_cell_formula`.
  WorkbookBuilder& formula_cell(std::string_view a1, std::string_view formula);

  // Stores an arbitrary `Value` (Blank / Number / Bool / Text / Error
  // / Array / Ref / Lambda) at A1-notation address `a1` via
  // `Sheet::set_cell_value`. Text payload lifetime is the caller's
  // responsibility (this overload does NOT intern); prefer the
  // `string_view` overload for plain text literals.
  WorkbookBuilder& cell(std::string_view a1, Value value);

  // Returns the populated workbook and leaves the builder in a moved-from
  // state. Subsequent calls on the same builder are undefined.
  Workbook build();

 private:
  // Resolves `a1` to (row, col) and asserts via gtest ADD_FAILURE on
  // parse failure. Returns false on failure; callers should skip the
  // cell write.
  static bool ParseA1(std::string_view a1, std::uint32_t* out_row, std::uint32_t* out_col);

  // Returns the active sheet, creating "Sheet1" on demand if no
  // sheet() call has happened yet.
  Sheet& active_sheet();

  Workbook wb_;
  std::size_t active_sheet_index_;
  bool has_active_sheet_;
};

}  // namespace test
}  // namespace formulon

#endif  // FORMULON_TESTS_UTIL_TEST_WORKBOOK_BUILDER_H_
