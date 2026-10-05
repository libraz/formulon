#pragma once

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/borrowed_arena.h"
#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace {

// RAII guard so the workbook handle is released even on test failure.
// Move-only so factory functions can return one by value.
struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
  WorkbookGuard(WorkbookGuard&& other) noexcept : handle(other.handle) { other.handle = nullptr; }
  WorkbookGuard& operator=(WorkbookGuard&& other) noexcept {
    if (this != &other) {
      fm_workbook_destroy(handle);
      handle = other.handle;
      other.handle = nullptr;
    }
    return *this;
  }
};

// RAII guard for `fm_cf_results_t` handles.
struct CfResultsGuard {
  fm_cf_results_t* handle = nullptr;
  ~CfResultsGuard() { fm_cf_results_destroy(handle); }
  CfResultsGuard() = default;
  CfResultsGuard(const CfResultsGuard&) = delete;
  CfResultsGuard& operator=(const CfResultsGuard&) = delete;
};

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

// Builds a `Workbook`, applies `mutate` to seed cells / CF blocks, then
// serialises it through OOXML and loads the resulting bytes through the
// C ABI. The returned guard owns a populated `fm_workbook_t*` ready for
// CF evaluation.
template <typename MutateFn>
[[maybe_unused]] inline WorkbookGuard WorkbookFromMutator(MutateFn&& mutate) {
  formulon::Workbook wb = formulon::Workbook::create();
  // Registers the `<dxf>` records the CF rules below reference. The
  // writer omits any `dxfId` its styles part cannot resolve, so a rule
  // whose differential format was never registered would come back from
  // the load with no `dxf_id` at all.
  wb.mutable_styles().dxfs.resize(8);
  mutate(wb);
  auto bytes = wb.save();
  EXPECT_TRUE(static_cast<bool>(bytes)) << "Workbook::save: " << (bytes ? "" : bytes.error().message);
  WorkbookGuard guard;
  if (!bytes) {
    return guard;
  }
  const auto& src = bytes.value();
  EXPECT_EQ(fm_workbook_load(src.data(), src.size(), &guard.handle), 0)
      << "fm_workbook_load: " << fm_last_error_message();
  return guard;
}

// Registers `count` `<dxf>` slots on an existing handle. `fm_sheet_cf_add_rule`
// refuses a `dxf_id` past the end of the table, so a mutation test that
// engages a differential format has to seed the table first.
[[maybe_unused]] inline void SeedDxfs(fm_workbook_t* handle, std::size_t count) {
  handle->workbook().mutable_styles().dxfs.resize(count);
}

[[maybe_unused]] inline formulon::cf::CFCellRange MakeRange(std::uint32_t r1, std::uint32_t c1, std::uint32_t r2,
                                                            std::uint32_t c2) {
  return {{r1, c1}, {r2, c2}};
}

// Minimal valid DataBar rule. Extension fields stay zero-initialized so
// each test below engages only what it is about to assert.
[[maybe_unused]] inline fm_cf_rule_t MakeDataBarRule(const fm_cf_cell_range_t* sqref) {
  fm_cf_rule_t rule{};
  rule.type = 3;  // DataBar
  rule.sqref = sqref;
  rule.sqref_count = 1;
  rule.data_bar_engaged = 1;
  rule.data_bar_min.type = 3;  // Min
  rule.data_bar_min.gte = 1;
  rule.data_bar_max.type = 4;  // Max
  rule.data_bar_max.gte = 1;
  rule.data_bar_fill = fm_cf_color_t{99, 142, 198, 255};
  rule.data_bar_show_value = 1;
  rule.data_bar_min_length_pct = 10;
  rule.data_bar_max_length_pct = 90;
  return rule;
}

}  // namespace
