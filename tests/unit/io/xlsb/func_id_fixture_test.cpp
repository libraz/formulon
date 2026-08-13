//
// Function-id evidence test against a real Mac Excel 365-produced
// `.xlsb` (`tests/fixtures/excel/xlsb_func_ids.xlsb`).
//
// The workbook is a probe grid: column A holds a function name as a
// label and column B a call to that function, one row per function per
// probed arity, starting at row 5 (rows 1-3 hold the helper block the
// range-taking probes read). Excel encoded each call as a `PtgFunc` /
// `PtgFuncVar` token carrying the function's 16-bit id, so reading the
// fixture back exercises `lookup_func_by_id` on ids Excel itself wrote.
//
// A wrong id in `func_id_table` is a silent substitution of a different
// function, and a wrong `arg_min` on a fixed-arity row makes the Reader
// pop the wrong number of operands; both surface here as a formula
// whose callee (or argument count) no longer matches the label Excel
// stored next to it.
//
// The workbook is regenerated with
// `tools/dev/xlsb_func_id_harvest.py build --out <path>`.

#include <cstdint>
#include <cstdio>
#include <set>
#include <string>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/xlsb/func_id_table.h"
#include "io/xlsb/reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Zero-based sheet coordinates of the probe grid.
constexpr std::uint32_t kFirstProbeRow = 4U;
constexpr std::uint32_t kProbeRowBound = 512U;  ///< Scanned upper bound.
constexpr std::uint32_t kLabelCol = 0U;
constexpr std::uint32_t kFormulaCol = 1U;

/// Number of distinct functions the fixture resolves through the table.
/// Every probe row must match its label, so this is also the count of
/// harvested ids the table carries.
constexpr std::size_t kCoveredFunctionCount = 90U;

std::string FixturePath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_func_ids.xlsb";
}

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    const std::size_t read = std::fread(out.data(), 1, out.size(), f);
    if (read != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(f);
  return out;
}

Workbook LoadFixture() {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath());
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  auto result_or = read_xlsb(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << "read_xlsb failed: " << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

/// Returns the callee of a single-call formula (`"=MROUND(10,3)"` ->
/// `"MROUND"`), or an empty string when `text` is not of that shape.
std::string CalleeOf(const std::string& text) {
  if (text.size() < 3U || text[0] != '=') {
    return std::string();
  }
  const std::string::size_type open = text.find('(');
  if (open == std::string::npos) {
    return std::string();
  }
  return text.substr(1U, open - 1U);
}

TEST(XlsbFuncIdFixture, EveryProbeFormulaResolvesBackToItsLabelledFunction) {
  Workbook wb = LoadFixture();
  ASSERT_GE(wb.sheet_count(), 1U);
  const Sheet& sheet = wb.sheet(0);

  std::set<std::string> covered;
  std::size_t rows_seen = 0U;
  for (std::uint32_t row = kFirstProbeRow; row < kProbeRowBound; ++row) {
    const Cell* label = sheet.cell_at(row, kLabelCol);
    if (label == nullptr || !label->cached_value.is_text()) {
      continue;
    }
    const std::string name(label->cached_value.as_text());
    const Cell* formula = sheet.cell_at(row, kFormulaCol);
    ASSERT_NE(formula, nullptr) << "row " << (row + 1U) << " labelled " << name << " has no formula cell";
    ++rows_seen;
    EXPECT_EQ(CalleeOf(formula->formula_text), name) << "row " << (row + 1U) << " formula=" << formula->formula_text;
    if (CalleeOf(formula->formula_text) == name) {
      covered.insert(name);
    }
  }

  EXPECT_GT(rows_seen, 100U) << "probe grid looks truncated";
  // `ISO.CEILING` is stored as a hidden `_xlfn.` name rather than a
  // function id, so it resolves through the future-function path and is
  // not one of the ids this table carries.
  covered.erase("ISO.CEILING");
  EXPECT_EQ(covered.size(), kCoveredFunctionCount);
}

TEST(XlsbFuncIdFixture, EveryResolvedNameLooksUpBackToTheSameRow) {
  // The name the Reader produced came from `lookup_func_by_id(id)` with
  // the id Excel wrote; looking that name back up must land on the same
  // row, which is what the Writer will emit for a call to it.
  Workbook wb = LoadFixture();
  ASSERT_GE(wb.sheet_count(), 1U);
  const Sheet& sheet = wb.sheet(0);

  for (std::uint32_t row = kFirstProbeRow; row < kProbeRowBound; ++row) {
    const Cell* formula = sheet.cell_at(row, kFormulaCol);
    if (formula == nullptr || formula->formula_text.empty()) {
      continue;
    }
    const std::string callee = CalleeOf(formula->formula_text);
    if (callee.empty() || callee == "ISO.CEILING") {
      continue;
    }
    const XlsbFuncEntry* by_name = lookup_func_by_name(callee);
    ASSERT_NE(by_name, nullptr) << "row " << (row + 1U) << " callee " << callee;
    const XlsbFuncEntry* by_id = lookup_func_by_id(by_name->id);
    ASSERT_NE(by_id, nullptr) << "id " << by_name->id;
    EXPECT_EQ(std::string(by_id->name), callee);
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
