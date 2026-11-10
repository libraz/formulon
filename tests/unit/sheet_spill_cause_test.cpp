//
// Why a dynamic-array spill is refused: the first obstructed cell in
// row-major order decides between a merge and any other occupant, and a
// rectangle leaving the grid is past the edge whatever lies before it. The
// refused footprint carries the cause, and only a refusal made at the anchor
// keeps its footprint across the commit.

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "merge_range.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::uint32_t kC = 2U;
constexpr std::uint32_t kD = 3U;
constexpr std::uint32_t kE = 4U;
constexpr std::uint32_t kX = 23U;

void Set(Workbook& wb, std::uint32_t row, std::uint32_t col, const std::string& formula) {
  const Expected<void, Error> set = wb.set_cell_formula(0U, row, col, formula);
  ASSERT_TRUE(static_cast<bool>(set)) << formula;
}

void Recalc(Workbook& wb) {
  const Expected<void, Error> done = wb.ensure_calculated(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(done)) << done.error().message;
}

// The anchor holds `#SPILL!` and a footprint of `rows` x `cols` refused for `cause`.
void ExpectRefused(const Workbook& wb, std::uint32_t row, std::uint32_t col, std::uint32_t rows, std::uint32_t cols,
                   SpillCause cause) {
  SCOPED_TRACE("anchor row " + std::to_string(row) + " col " + std::to_string(col));
  const Value v = wb.sheet(0U).resolve_cell_value(row, col);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
  const std::optional<BlockedSpillFootprint> footprint = wb.sheet(0U).blocked_spill_footprint_at(row, col);
  ASSERT_TRUE(footprint.has_value());
  EXPECT_EQ(footprint->rows, rows);
  EXPECT_EQ(footprint->cols, cols);
  EXPECT_EQ(footprint->cause, cause);
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(row, col).has_value());
}

TEST(SpillCause, FirstObstructionInRowMajorOrderDecides) {
  struct Case {
    const char* name;
    std::uint32_t spill_rows;
    void (*place)(Workbook& wb);
    SpillCause cause;
  };
  const Case cases[] = {
      // A merge is a merge whether or not it holds a value.
      {"merge with value C3:D3", 3U,
       [](Workbook& wb) {
         ASSERT_TRUE(static_cast<bool>(wb.add_merge(0U, MergeRange{2U, kC, 2U, kD})));
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 2U, kC, "m")));
       },
       SpillCause::kMerged},
      {"value D2 before empty merge C4:D4", 4U,
       [](Workbook& wb) {
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 1U, kD, "v")));
         ASSERT_TRUE(static_cast<bool>(wb.add_merge(0U, MergeRange{3U, kC, 3U, kD})));
       },
       SpillCause::kOccupied},
      {"empty merge C2:D2 before value D4", 4U,
       [](Workbook& wb) {
         ASSERT_TRUE(static_cast<bool>(wb.add_merge(0U, MergeRange{1U, kC, 1U, kD})));
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 3U, kD, "v")));
       },
       SpillCause::kMerged},
      // Row-major, not column-major: D2 comes before C3.
      {"empty merge D2:D3 before value C3", 4U,
       [](Workbook& wb) {
         ASSERT_TRUE(static_cast<bool>(wb.add_merge(0U, MergeRange{1U, kD, 2U, kD})));
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 2U, kC, "v")));
       },
       SpillCause::kMerged},
  };
  for (const Case& entry : cases) {
    SCOPED_TRACE(entry.name);
    Workbook wb = Workbook::create();
    entry.place(wb);
    Set(wb, 0U, kC, "=SEQUENCE(" + std::to_string(entry.spill_rows) + ",2)");
    Recalc(wb);
    ExpectRefused(wb, 0U, kC, entry.spill_rows, 2U, entry.cause);
  }
}

TEST(SpillCause, EveryNonMergeOccupantIsOccupied) {
  struct Case {
    const char* name;
    void (*place)(Workbook& wb);
  };
  const Case cases[] = {
      {"value", [](Workbook& wb) { ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 2U, kC, "x"))); }},
      {"formula", [](Workbook& wb) { Set(wb, 2U, kC, "=\"\""); }},
      {"legacy array",
       [](Workbook& wb) { ASSERT_TRUE(static_cast<bool>(wb.set_array_formula(0U, 2U, kC, 1U, 2U, "=1"))); }},
      {"table",
       [](Workbook& wb) {
         TableMetadata table;
         table.name = "Blocker";
         table.display_name = "Blocker";
         table.ref = "C3:D4";
         table.sheet_index = 0U;
         for (const char* name : {"A", "B"}) {
           TableColumn column;
           column.name = name;
           table.columns.push_back(std::move(column));
         }
         ASSERT_TRUE(static_cast<bool>(wb.add_table(std::move(table))));
       }},
      // Another spill's phantom cell, entered first.
      {"phantom", [](Workbook& wb) { Set(wb, 2U, kC - 1U, "=SEQUENCE(1,3)"); }},
  };
  for (const Case& entry : cases) {
    SCOPED_TRACE(entry.name);
    Workbook wb = Workbook::create();
    entry.place(wb);
    Recalc(wb);
    Set(wb, 0U, kC, "=SEQUENCE(3)");
    Recalc(wb);
    ExpectRefused(wb, 0U, kC, 3U, 1U, SpillCause::kOccupied);
  }
}

TEST(SpillCause, PastTheEdgeRecordsAFootprint) {
  Workbook wb = Workbook::create();
  // Bottom edge, right edge, and an in-grid value before the bottom edge.
  Set(wb, Sheet::kMaxRows - 2U, kC, "=SEQUENCE(5)");
  Set(wb, 0U, Sheet::kMaxCols - 3U, "=SEQUENCE(1,5)");
  Set(wb, Sheet::kMaxRows - 3U, kE, "=SEQUENCE(5)");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, Sheet::kMaxRows - 2U, kE, "v")));
  Recalc(wb);
  ExpectRefused(wb, Sheet::kMaxRows - 2U, kC, 5U, 1U, SpillCause::kBeyondEdge);
  ExpectRefused(wb, 0U, Sheet::kMaxCols - 3U, 1U, 5U, SpillCause::kBeyondEdge);
  ExpectRefused(wb, Sheet::kMaxRows - 3U, kE, 5U, 1U, SpillCause::kBeyondEdge);
}

TEST(SpillCause, PropagatedSpillDropsAStaleFootprint) {
  Workbook wb = Workbook::create();
  // E10 is blocked by E11; C1 spills when X1 = 1 (blocked by C2) and reads E10 otherwise.
  Set(wb, 9U, kE, "=SEQUENCE(3)");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 10U, kE, "x")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, kX, Value::number(1.0))));
  Set(wb, 0U, kC, "=IF(X1=1,SEQUENCE(3),E10)");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 1U, kC, "y")));
  Recalc(wb);
  ExpectRefused(wb, 0U, kC, 3U, 1U, SpillCause::kOccupied);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, kX, Value::number(0.0))));
  Recalc(wb);
  const Value v = wb.sheet(0U).resolve_cell_value(0U, kC);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
  EXPECT_FALSE(wb.sheet(0U).blocked_spill_footprint_at(0U, kC).has_value());
  const std::optional<CellErrorDetail> detail = wb.sheet(0U).error_detail_at(0U, kC);
  ASSERT_TRUE(detail.has_value());
  EXPECT_EQ(detail->code, ErrorCode::Spill);
  EXPECT_EQ(detail->shape, CellErrorDetail::Shape::kPropagated);
}

TEST(SpillCause, RefusalAtTheAnchorKeepsItsFootprint) {
  Workbook wb = Workbook::create();
  // A whole column refuses its own footprint before it is materialised.
  for (std::uint32_t r = 0; r < 3U; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, r, kX, Value::number(r + 1.0))));
  }
  Set(wb, 0U, kC, "=X:X");
  ASSERT_TRUE(static_cast<bool>(wb.add_merge(0U, MergeRange{1U, kC, 1U, kD})));
  Recalc(wb);
  ExpectRefused(wb, 0U, kC, Sheet::kMaxRows, 1U, SpillCause::kMerged);
}

TEST(SpillCause, ProbeReportsTheCauseOnlyWhenRefused) {
  Sheet sheet("Sheet1");
  EXPECT_EQ(sheet.probe_spill_footprint(0U, 0U, 3U, 1U).admission, Sheet::SpillAdmission::kAdmissible);
  sheet.set_cell_value(2U, 0U, Value::number(1.0));
  Sheet::SpillProbe probe = sheet.probe_spill_footprint(0U, 0U, 3U, 1U);
  EXPECT_EQ(probe.admission, Sheet::SpillAdmission::kBlocked);
  EXPECT_EQ(probe.cause, SpillCause::kOccupied);
  sheet.add_merge(MergeRange{1U, 0U, 1U, 1U});
  probe = sheet.probe_spill_footprint(0U, 0U, 3U, 1U);
  EXPECT_EQ(probe.admission, Sheet::SpillAdmission::kBlocked);
  EXPECT_EQ(probe.cause, SpillCause::kMerged);
  probe = sheet.probe_spill_footprint(Sheet::kMaxRows - 1U, 0U, 3U, 1U);
  EXPECT_EQ(probe.admission, Sheet::SpillAdmission::kOutsideGrid);
  EXPECT_EQ(probe.cause, SpillCause::kBeyondEdge);
}

TEST(SpillCause, RestoreKeepsAPastEdgeFootprint) {
  Sheet sheet("Sheet1");
  sheet.restore_blocked_spill_footprints(
      {BlockedSpillFootprint{Sheet::kMaxRows - 1U, 0U, 5U, 1U, SpillCause::kBeyondEdge},
       BlockedSpillFootprint{0U, 0U, 0U, 1U, SpillCause::kOccupied}});
  const std::vector<BlockedSpillFootprint> kept = sheet.blocked_spill_footprints();
  ASSERT_EQ(kept.size(), 1U);
  EXPECT_EQ(kept[0].anchor_row, Sheet::kMaxRows - 1U);
  EXPECT_EQ(kept[0].rows, 5U);
  EXPECT_EQ(kept[0].cause, SpillCause::kBeyondEdge);
}

TEST(SpillCause, RestoredBlockedSpillReadsAsTheFileRecordedIt) {
  Workbook wb = Workbook::create();
  Set(wb, 0U, kC, "=SEQUENCE(3)");
  ASSERT_TRUE(static_cast<bool>(wb.restore_blocked_spill(0U, 0U, kC, 1U, 1U, SpillCause::kBeyondEdge)));
  ExpectRefused(wb, 0U, kC, 1U, 1U, SpillCause::kBeyondEdge);
  EXPECT_FALSE(static_cast<bool>(wb.restore_blocked_spill(0U, 0U, kC, 0U, 1U, SpillCause::kOccupied)));
  // The next recalc recomputes the footprint: nothing blocks it any more.
  Recalc(wb);
  EXPECT_FALSE(wb.sheet(0U).blocked_spill_footprint_at(0U, kC).has_value());
  EXPECT_EQ(wb.sheet(0U).resolve_cell_value(2U, kC).as_number(), 3.0);
}

}  // namespace
}  // namespace formulon
