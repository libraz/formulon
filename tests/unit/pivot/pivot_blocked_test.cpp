//
// A PivotTable report whose range is occupied or overflows the sheet: the
// edit applies, the report is withheld behind `#SPILL!` in its anchor, and the
// next recalc renders it once the range is free. Creation refuses instead,
// and a move clears what a paste would overwrite.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <utility>

#include "calc_settings.h"
#include "cell.h"
#include "data_table.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "merge_range.h"
#include "pivot/field_remap.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "styles.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using pivot::PivotAxis;
using pivot::PivotCache;
using pivot::PivotCacheField;
using pivot::PivotCacheRecord;
using pivot::PivotDataField;
using pivot::PivotField;
using pivot::PivotItem;
using pivot::PivotRect;
using pivot::PivotTable;

constexpr std::size_t kDataSheet = 0U;
constexpr std::size_t kPivotSheet = 1U;
constexpr std::uint32_t kCacheId = 1U;
constexpr std::uint32_t kRegion = 0U;
constexpr std::uint32_t kAmt = 2U;
// The placeholder frame's outer edge colour and its inner edge palette index.
constexpr std::uint32_t kGrey = 0xFF999999U;
constexpr std::uint32_t kWindow = 65U;

PivotField Field(std::string name, PivotAxis axis = PivotAxis::None) {
  PivotField field;
  field.source_name = std::move(name);
  field.axis = axis;
  return field;
}

PivotCache BaseCache(const std::string& sheet_name) {
  PivotCache cache;
  cache.set_cache_id(kCacheId);
  cache.mutable_fields().push_back(
      PivotCacheField{"Region", {Value::text("North"), Value::text("South"), Value::text("East")}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {Value::text("A"), Value::text("B")}});
  cache.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  const double rows[][3] = {{0, 0, 100}, {1, 1, 200}, {2, 0, 300}, {0, 1, 400}, {1, 0, 500}};
  for (const auto& row : rows) {
    PivotCacheRecord record;
    record.cells = {Value::number(row[0]), Value::number(row[1]), Value::number(row[2])};
    record.cell_is_index = {true, true, false};
    cache.mutable_records().push_back(std::move(record));
  }
  pivot::WorksheetSource& source = cache.mutable_worksheet_source();
  source.present = true;
  source.ref = "A1:C6";
  source.sheet = sheet_name;
  return cache;
}

std::unique_ptr<PivotTable> RegionPivot(const char* name, std::uint32_t row, std::uint32_t col) {
  auto table = std::make_unique<PivotTable>();
  table->set_name(name);
  table->set_pivot_cache_id(kCacheId);
  table->mutable_fields().push_back(Field("Region", PivotAxis::Row));
  table->mutable_fields().push_back(Field("Product"));
  table->mutable_fields().push_back(Field("Amt", PivotAxis::Value));
  for (std::uint32_t x = 0; x < 3U; ++x) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    table->mutable_fields()[kRegion].items.push_back(std::move(item));
  }
  table->mutable_row_field_order() = {kRegion};
  PivotDataField sum;
  sum.name = "Sum of Amt";
  sum.field_index = kAmt;
  table->mutable_data_fields().push_back(std::move(sum));
  table->set_anchor(row, col, 1U, 1U);
  return table;
}

// Region / Product / Amt on the first sheet and one Region pivot anchored at Pivot!A3.
Workbook Book() {
  Workbook wb = Workbook::create_for_profile(mac_365_en_us_profile());
  const char* regions[] = {"North", "South", "East", "North", "South"};
  const char* products[] = {"A", "B", "A", "B", "A"};
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, 0U, 0U, "Region")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, 0U, 1U, "Product")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, 0U, 2U, "Amt")));
  for (std::uint32_t r = 0; r < 5U; ++r) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, r + 1U, 0U, regions[r])));
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, r + 1U, 1U, products[r])));
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(kDataSheet, r + 1U, 2U, Value::number(100.0 * (r + 1U)))));
  }
  wb.add_pivot_cache(std::make_unique<PivotCache>(BaseCache(wb.sheet(kDataSheet).name())));
  EXPECT_TRUE(static_cast<bool>(wb.add_sheet_entered("Pivot")));
  EXPECT_TRUE(static_cast<bool>(wb.add_pivot(kPivotSheet, RegionPivot("PT1", 2U, 0U))));
  EXPECT_TRUE(static_cast<bool>(wb.project_pivot(kPivotSheet, 0U, ReportPlacement::kRefuse)));
  return wb;
}

PivotTable& PivotAt(Workbook& wb, std::size_t index = 0U) {
  return *wb.sheet(kPivotSheet).mutable_pivot_tables()[index];
}

PivotRect Extent(Workbook& wb, std::size_t index = 0U) {
  EXPECT_TRUE(PivotAt(wb, index).projected_extent().has_value());
  return PivotAt(wb, index).projected_extent().value_or(PivotRect{});
}

Value At(const Workbook& wb, std::uint32_t row, std::uint32_t col, std::size_t sheet = kPivotSheet) {
  return wb.sheet(sheet).resolve_cell_value(row, col);
}

bool HoldsSpill(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Value v = At(wb, row, col);
  return v.is_error() && v.as_error() == ErrorCode::Spill;
}

void Recalc(Workbook& wb) {
  const Expected<void, Error> done = wb.ensure_calculated(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(done)) << done.error().message;
}

// The edit every case grows the report with: a second data field, one more column to the right.
Expected<void, Error> AddDataField(Workbook& wb, std::size_t index = 0U) {
  PivotTable& table = PivotAt(wb, index);
  PivotDataField second;
  second.name = "Sum of Amt 2";
  second.field_index = kAmt;
  table.mutable_data_fields().push_back(std::move(second));
  pivot::clear_derived_state(table);
  return wb.project_pivot(kPivotSheet, index, ReportPlacement::kBlock);
}

// The range the edit gives the report with nothing in the way.
PivotRect ControlExtent() {
  Workbook wb = Book();
  EXPECT_TRUE(static_cast<bool>(AddDataField(wb)));
  return Extent(wb);
}

std::string A1(std::uint32_t row, std::uint32_t col) {
  std::string letters;
  for (std::uint32_t c = col + 1U; c > 0U; c = (c - 1U) / 26U) {
    letters.insert(letters.begin(), static_cast<char>('A' + (c - 1U) % 26U));
  }
  return letters + std::to_string(row + 1U);
}

TableMetadata OneColumnTable(const char* name, std::uint32_t row, std::uint32_t col) {
  TableMetadata table;
  table.name = name;
  table.display_name = name;
  table.ref = A1(row, col) + ":" + A1(row + 1U, col);
  table.sheet_index = kPivotSheet;
  TableColumn column;
  column.name = "Col";
  table.columns.push_back(std::move(column));
  return table;
}

DataTableInput InputAt(std::uint32_t row, std::uint32_t col) {
  DataTableInput input;
  input.present = true;
  input.row = row;
  input.col = col;
  return input;
}

// The report is withheld at its anchor: the error, a one-cell span, offsets 0 and no projected merges.
void ExpectBlockedAt(Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const PivotTable& table = PivotAt(wb);
  ASSERT_TRUE(table.blocked().has_value());
  EXPECT_TRUE(HoldsSpill(wb, row, col));
  EXPECT_EQ(table.anchor_row(), row);
  EXPECT_EQ(table.anchor_col(), col);
  EXPECT_EQ(table.span_rows(), 1U);
  EXPECT_EQ(table.span_cols(), 1U);
  EXPECT_EQ(table.location_first_header_row(), std::optional<std::uint32_t>(0U));
  EXPECT_EQ(table.location_first_data_row(), std::optional<std::uint32_t>(0U));
  EXPECT_EQ(table.location_first_data_col(), std::optional<std::uint32_t>(0U));
  EXPECT_FALSE(table.projection_pending());
  EXPECT_TRUE(table.projected_merges().empty());
}

void ExpectRendered(Workbook& wb, const PivotRect& control) {
  const PivotTable& table = PivotAt(wb);
  EXPECT_FALSE(table.blocked().has_value());
  ASSERT_TRUE(table.projected_extent().has_value());
  EXPECT_EQ(*table.projected_extent(), control);
  EXPECT_FALSE(HoldsSpill(wb, control.row0, control.col0));
  EXPECT_EQ(table.span_rows(), control.row1 - control.row0 + 1U);
  EXPECT_EQ(table.span_cols(), control.col1 - control.col0 + 1U);
}

// ---------------------------------------------------------------------------
// Blocking and recovery
// ---------------------------------------------------------------------------

TEST(PivotBlocked, EditOverAValueBlocksAndKeepsTheDefinition) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  const PivotRect before = Extent(wb);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row1, control.col1, "blocker")));
  const Expected<void, Error> edited = AddDataField(wb);
  ASSERT_TRUE(static_cast<bool>(edited)) << edited.error().message;
  ExpectBlockedAt(wb, before.row0, before.col0);
  const PivotTable& table = PivotAt(wb);
  EXPECT_EQ(table.data_fields().size(), 2U);
  EXPECT_EQ(table.blocked()->rows, control.row1 - control.row0 + 1U);
  EXPECT_EQ(table.blocked()->cols, control.col1 - control.col0 + 1U);
  EXPECT_EQ(table.blocked()->rows_above, 0U);
  EXPECT_FALSE(table.blocked()->past_edge);
  EXPECT_FALSE(table.auto_formatting());
  EXPECT_EQ(*table.projected_extent(), (PivotRect{before.row0, before.col0, before.row0, before.col0}));
  EXPECT_EQ(At(wb, control.row1, control.col1).as_text(), "blocker");
  EXPECT_TRUE(At(wb, before.row0 + 1U, before.col0).is_blank());
}

Workbook ReloadXlsx(const Workbook& wb) {
  auto bytes = io::write_ooxml(wb);
  if (!bytes) {
    ADD_FAILURE() << bytes.error().message;
    return Workbook::create();
  }
  auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  if (!read) {
    ADD_FAILURE() << read.error().message;
    return Workbook::create();
  }
  return std::move(read.value().workbook);
}

// Excel keeps a saved report blocked on reopen and draws it once the range comes free.
void ExpectReloadedBlockedReportRedrawsWhenFreed(Workbook (*reload)(const Workbook&)) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row1, control.col1, "blocker")));
  ASSERT_TRUE(static_cast<bool>(AddDataField(wb)));
  Recalc(wb);
  ASSERT_TRUE(PivotAt(wb).blocked().has_value());

  Workbook loaded = reload(wb);
  ASSERT_GT(loaded.sheet_count(), kPivotSheet);
  ASSERT_EQ(loaded.sheet(kPivotSheet).pivot_tables().size(), 1U);
  Recalc(loaded);
  ASSERT_TRUE(PivotAt(loaded).blocked().has_value());
  EXPECT_TRUE(At(loaded, control.row0, control.col0).is_error());
  EXPECT_EQ(At(loaded, control.row1, control.col1).as_text(), "blocker");
  ASSERT_TRUE(static_cast<bool>(loaded.set_cell_value(kPivotSheet, control.row1, control.col1, Value::blank())));
  Recalc(loaded);
  ExpectRendered(loaded, control);
  EXPECT_EQ(PivotAt(loaded).data_fields().size(), 2U);
}

TEST(PivotBlocked, ReloadedXlsxBlockedReportRedrawsWhenItsRangeIsFreed) {
  ExpectReloadedBlockedReportRedrawsWhenFreed(&ReloadXlsx);
}

TEST(PivotBlocked, ReloadedDrawnReportIsNotBlocked) {
  Workbook wb = Book();
  Recalc(wb);
  const PivotRect drawn = Extent(wb);
  Workbook loaded = ReloadXlsx(wb);
  Recalc(loaded);
  EXPECT_FALSE(PivotAt(loaded).blocked().has_value());
  EXPECT_FALSE(HoldsSpill(loaded, drawn.row0, drawn.col0));
}

TEST(PivotBlocked, ClearingTheBlockerRendersOnRecalc) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row1, control.col1, "blocker")));
  ASSERT_TRUE(static_cast<bool>(AddDataField(wb)));
  Recalc(wb);
  ASSERT_TRUE(PivotAt(wb).blocked().has_value());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, control.row1, control.col1, Value::blank())));
  Recalc(wb);
  ExpectRendered(wb, control);
  EXPECT_EQ(PivotAt(wb).data_fields().size(), 2U);
  // Excel leaves autofit off after the report recovers.
  EXPECT_FALSE(PivotAt(wb).auto_formatting());
}

TEST(PivotBlocked, FieldlessReportIsTheBorderedPlaceholder) {
  Workbook wb = Book();
  auto table = std::make_unique<PivotTable>();
  table->set_name("Empty");
  table->set_pivot_cache_id(kCacheId);
  table->set_anchor(2U, 6U, 1U, 1U);
  const Expected<std::size_t, Error> added = wb.add_pivot(kPivotSheet, std::move(table));
  ASSERT_TRUE(static_cast<bool>(added)) << added.error().message;
  const Expected<void, Error> projected = wb.project_pivot(kPivotSheet, added.value(), ReportPlacement::kRefuse);
  ASSERT_TRUE(static_cast<bool>(projected)) << projected.error().message;
  // G3:I20: three columns by eighteen rows from the anchor.
  EXPECT_EQ(Extent(wb, added.value()), (PivotRect{2U, 6U, 19U, 8U}));
  const StylesTable& styles = wb.styles();
  const auto expect_side = [](const BorderSide& side, std::optional<bool> outer, const std::string& where) {
    if (!outer.has_value()) {
      EXPECT_EQ(side.style, 0U) << where;
      return;
    }
    EXPECT_EQ(side.style, 1U) << where;  // thin
    if (*outer) {
      EXPECT_EQ(side.color.kind, ColorSpec::Kind::kRgb) << where;
      EXPECT_EQ(side.color.rgb, kGrey) << where;
    } else {
      EXPECT_EQ(side.color.kind, ColorSpec::Kind::kIndexed) << where;
      EXPECT_EQ(side.color.indexed, kWindow) << where;
    }
  };
  for (std::uint32_t r = 0; r < 18U; ++r) {
    for (std::uint32_t c = 0; c < 3U; ++c) {
      const std::string where = A1(2U + r, 6U + c);
      const Cell* cell = wb.sheet(kPivotSheet).cell_at(2U + r, 6U + c);
      ASSERT_NE(cell, nullptr) << where;
      EXPECT_TRUE(cell->formula_text.empty() && cell->cached_value.is_blank()) << where;
      ASSERT_LT(cell->xf_index, styles.cell_xfs.size()) << where;
      const CellXf& xf = styles.cell_xfs[cell->xf_index];
      EXPECT_TRUE(xf.apply_border) << where;
      ASSERT_LT(xf.border_index, styles.borders.size()) << where;
      const BorderRecord& border = styles.borders[xf.border_index];
      // Outer edges grey; inner left and top edges in palette colour 65; no inner right or bottom edge.
      expect_side(border.left, c == 0U, where + " left");
      expect_side(border.top, r == 0U, where + " top");
      expect_side(border.right, c == 2U ? std::optional<bool>(true) : std::nullopt, where + " right");
      expect_side(border.bottom, r == 17U ? std::optional<bool>(true) : std::nullopt, where + " bottom");
    }
  }
  // Nine positions, nine borders.
  std::size_t distinct = 0;
  for (std::size_t i = 0; i < styles.borders.size(); ++i) {
    distinct += styles.borders[i].left.style == 1U ? 1U : 0U;
  }
  EXPECT_EQ(distinct, 9U);
}

TEST(PivotBlocked, CreationOverAnOccupantIsRefused) {
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, 19U, 8U, "taken")));  // I20, the frame's last cell
  auto table = std::make_unique<PivotTable>();
  table->set_name("Empty");
  table->set_pivot_cache_id(kCacheId);
  table->set_anchor(2U, 6U, 1U, 1U);
  const Expected<std::size_t, Error> added = wb.add_pivot(kPivotSheet, std::move(table));
  ASSERT_TRUE(static_cast<bool>(added));
  const Expected<void, Error> projected = wb.project_pivot(kPivotSheet, added.value(), ReportPlacement::kRefuse);
  ASSERT_FALSE(static_cast<bool>(projected));
  EXPECT_EQ(projected.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_NE(projected.error().message.find("non-empty cell Pivot!I20"), std::string::npos) << projected.error().message;
  EXPECT_FALSE(PivotAt(wb, added.value()).blocked().has_value());
  EXPECT_FALSE(PivotAt(wb, added.value()).projected_extent().has_value());
  EXPECT_EQ(wb.sheet(kPivotSheet).cell_at(2U, 6U), nullptr);
  EXPECT_EQ(At(wb, 19U, 8U).as_text(), "taken");
}

TEST(PivotBlocked, EveryOccupantBlocks) {
  const PivotRect control = ControlExtent();
  const std::uint32_t row = control.row1;
  const std::uint32_t col = control.col1;
  struct Case {
    const char* name;
    void (*place)(Workbook& wb, std::uint32_t row, std::uint32_t col, std::uint32_t top);
  };
  const Case cases[] = {
      {"value",
       [](Workbook& wb, std::uint32_t r, std::uint32_t c, std::uint32_t) {
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, r, c, Value::number(1.0))));
       }},
      {"empty-text formula",
       [](Workbook& wb, std::uint32_t r, std::uint32_t c, std::uint32_t) {
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(kPivotSheet, r, c, "=\"\"")));
       }},
      {"empty merge",
       [](Workbook& wb, std::uint32_t r, std::uint32_t c, std::uint32_t) {
         ASSERT_TRUE(static_cast<bool>(wb.add_merge(kPivotSheet, MergeRange{r - 1U, c, r, c})));
       }},
      {"straddling merge",
       [](Workbook& wb, std::uint32_t r, std::uint32_t c, std::uint32_t) {
         ASSERT_TRUE(static_cast<bool>(wb.add_merge(kPivotSheet, MergeRange{r, c, r + 1U, c + 1U})));
       }},
      {"legacy array",
       [](Workbook& wb, std::uint32_t r, std::uint32_t c, std::uint32_t) {
         ASSERT_TRUE(static_cast<bool>(wb.set_array_formula(kPivotSheet, r, c, 2U, 1U, "={1;2}")));
       }},
      {"data-table result",
       [](Workbook& wb, std::uint32_t r, std::uint32_t c, std::uint32_t) {
         ASSERT_TRUE(static_cast<bool>(
             wb.add_data_table(kPivotSheet, MergeRange{r, c, r + 1U, c}, DataTableInput{}, InputAt(0U, 20U))));
       }},
      {"table", [](Workbook& wb, std::uint32_t r, std::uint32_t c,
                   std::uint32_t) { ASSERT_TRUE(static_cast<bool>(wb.add_table(OneColumnTable("Blocker", r, c)))); }},
      {"other pivot",
       [](Workbook& wb, std::uint32_t, std::uint32_t c, std::uint32_t top) {
         ASSERT_TRUE(static_cast<bool>(wb.add_pivot(kPivotSheet, RegionPivot("PT2", top, c))));
         ASSERT_TRUE(static_cast<bool>(wb.project_pivot(kPivotSheet, 1U, ReportPlacement::kRefuse)));
       }},
      {"spill range",
       [](Workbook& wb, std::uint32_t, std::uint32_t c, std::uint32_t top) {
         // Anchored above the report, so only spilled cells reach into its range.
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(kPivotSheet, top - 1U, c, "=SEQUENCE(2)")));
         Recalc(wb);
         ASSERT_EQ(At(wb, top, c).as_number(), 2.0);
       }},
  };
  for (const Case& entry : cases) {
    SCOPED_TRACE(entry.name);
    Workbook wb = Book();
    entry.place(wb, row, col, control.row0);
    const Expected<void, Error> edited = AddDataField(wb);
    ASSERT_TRUE(static_cast<bool>(edited)) << edited.error().message;
    ExpectBlockedAt(wb, control.row0, control.col0);
    Recalc(wb);
    EXPECT_TRUE(PivotAt(wb).blocked().has_value());
  }
}

TEST(PivotBlocked, GrowthPastTheEdgeBlocksWithTheClippedSize) {
  Workbook wb = Book();
  const PivotRect before = Extent(wb);
  const std::uint32_t width = before.col1 - before.col0 + 1U;
  const std::uint32_t height = before.row1 - before.row0 + 1U;
  // Move the report flush against the last column, then widen it by one.
  ASSERT_TRUE(static_cast<bool>(wb.move_pivot(kPivotSheet, 0U, before.row0, Sheet::kMaxCols - width)));
  ASSERT_EQ(Extent(wb).col1, Sheet::kMaxCols - 1U);
  const Expected<void, Error> edited = AddDataField(wb);
  ASSERT_TRUE(static_cast<bool>(edited)) << edited.error().message;
  ExpectBlockedAt(wb, before.row0, Sheet::kMaxCols - width);
  EXPECT_TRUE(PivotAt(wb).blocked()->past_edge);
  // The overflowing dimension records the room left, the other the needed size.
  EXPECT_EQ(PivotAt(wb).blocked()->cols, width);
  EXPECT_EQ(PivotAt(wb).blocked()->rows, height);
  Recalc(wb);
  EXPECT_TRUE(PivotAt(wb).blocked().has_value());
}

TEST(PivotBlocked, AnchorKeepsItsFormatAndTheBodyLosesValuesAndFormats) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  const PivotRect before = Extent(wb);
  const Cell* anchor = wb.sheet(kPivotSheet).cell_at(before.row0, before.col0);
  ASSERT_NE(anchor, nullptr);
  const std::uint32_t anchor_xf = anchor->xf_index;
  ASSERT_NE(anchor_xf, 0U);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row1, control.col1, "blocker")));
  ASSERT_TRUE(static_cast<bool>(AddDataField(wb)));
  anchor = wb.sheet(kPivotSheet).cell_at(before.row0, before.col0);
  ASSERT_NE(anchor, nullptr);
  EXPECT_EQ(anchor->xf_index, anchor_xf);
  for (std::uint32_t r = before.row0; r <= before.row1; ++r) {
    for (std::uint32_t c = before.col0; c <= before.col1; ++c) {
      if (r != before.row0 || c != before.col0) {
        EXPECT_EQ(wb.sheet(kPivotSheet).cell_at(r, c), nullptr) << A1(r, c);
      }
    }
  }
}

TEST(PivotBlocked, TwoBlockersNeedBothCleared) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row0, control.col1, "one")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row1, control.col1, "two")));
  ASSERT_TRUE(static_cast<bool>(AddDataField(wb)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, control.row0, control.col1, Value::blank())));
  Recalc(wb);
  EXPECT_TRUE(PivotAt(wb).blocked().has_value());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, control.row1, control.col1, Value::blank())));
  Recalc(wb);
  ExpectRendered(wb, control);
}

TEST(PivotBlocked, ManualCalculationStaysBlockedUntilAnExplicitRecalc) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  wb.set_calc_mode(CalcMode::kManual);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row1, control.col1, "blocker")));
  ASSERT_TRUE(static_cast<bool>(AddDataField(wb)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, control.row1, control.col1, Value::blank())));
  Recalc(wb);
  EXPECT_TRUE(PivotAt(wb).blocked().has_value());
  EXPECT_TRUE(HoldsSpill(wb, control.row0, control.col0));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ExpectRendered(wb, control);
}

TEST(PivotBlocked, UnmergingAloneDoesNotRecoverButALaterValueWriteDoes) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  const MergeRange merge{control.row1 - 1U, control.col1, control.row1, control.col1};
  ASSERT_TRUE(static_cast<bool>(wb.add_merge(kPivotSheet, merge)));
  ASSERT_TRUE(static_cast<bool>(AddDataField(wb)));
  Recalc(wb);
  ASSERT_TRUE(PivotAt(wb).blocked().has_value());
  ASSERT_TRUE(static_cast<bool>(wb.remove_merges_intersecting(kPivotSheet, merge)));
  Recalc(wb);
  EXPECT_TRUE(PivotAt(wb).blocked().has_value());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, 49U, 25U, Value::number(1.0))));  // Z50
  Recalc(wb);
  ExpectRendered(wb, control);
}

TEST(PivotBlocked, GetPivotDataIsRefWhileBlockedAndLiveAfterRecovery) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  ASSERT_EQ(Extent(wb).row0, 2U);
  ASSERT_EQ(Extent(wb).col0, 0U);
  ASSERT_TRUE(static_cast<bool>(
      wb.set_cell_formula(kDataSheet, 0U, 4U, "=GETPIVOTDATA(\"Sum of Amt\",Pivot!$A$3,\"Region\",\"North\")")));
  ASSERT_TRUE(static_cast<bool>(
      wb.set_cell_formula(kDataSheet, 1U, 4U, "=GETPIVOTDATA(\"Sum of Amt\",Pivot!$B$4,\"Region\",\"North\")")));
  Recalc(wb);
  EXPECT_EQ(At(wb, 0U, 4U, kDataSheet).as_number(), 500.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row1, control.col1, "blocker")));
  ASSERT_TRUE(static_cast<bool>(AddDataField(wb)));
  Recalc(wb);
  for (const std::uint32_t r : {0U, 1U}) {
    const Value v = At(wb, r, 4U, kDataSheet);
    ASSERT_TRUE(v.is_error()) << r;
    EXPECT_EQ(v.as_error(), ErrorCode::Ref) << r;
  }

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, control.row1, control.col1, Value::blank())));
  Recalc(wb);
  ASSERT_FALSE(PivotAt(wb).blocked().has_value());
  EXPECT_EQ(At(wb, 0U, 4U, kDataSheet).as_number(), 500.0);
  EXPECT_EQ(At(wb, 1U, 4U, kDataSheet).as_number(), 500.0);
}

TEST(PivotBlocked, RowInsertAboveMovesTheBlockedAnchor) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row1, control.col1, "blocker")));
  ASSERT_TRUE(static_cast<bool>(AddDataField(wb)));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(kPivotSheet, 0U, 2U)));
  ExpectBlockedAt(wb, control.row0 + 2U, control.col0);
  EXPECT_EQ(PivotAt(wb).blocked()->rows, control.row1 - control.row0 + 1U);
  EXPECT_EQ(At(wb, control.row1 + 2U, control.col1).as_text(), "blocker");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, control.row1 + 2U, control.col1, Value::blank())));
  Recalc(wb);
  ExpectRendered(wb, PivotRect{control.row0 + 2U, control.col0, control.row1 + 2U, control.col1});
}

// ---------------------------------------------------------------------------
// Moving a report
// ---------------------------------------------------------------------------

TEST(PivotBlocked, MoveClearsValuesAndMergesAndDeletesATableAtTheDestination) {
  Workbook wb = Book();
  const PivotRect before = Extent(wb);
  const std::uint32_t to = 20U;
  const std::uint32_t height = before.row1 - before.row0 + 1U;
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, to + 1U, before.col0 + 1U, "overwritten")));
  ASSERT_TRUE(
      static_cast<bool>(wb.add_merge(kPivotSheet, MergeRange{to + 2U, before.col0, to + 2U, before.col0 + 1U})));
  // A table only partly inside the destination goes as a whole.
  ASSERT_TRUE(static_cast<bool>(wb.add_table(OneColumnTable("Pasted", to + height - 1U, before.col0 + 1U))));
  const Expected<void, Error> moved = wb.move_pivot(kPivotSheet, 0U, to, before.col0);
  ASSERT_TRUE(static_cast<bool>(moved)) << moved.error().message;
  EXPECT_FALSE(PivotAt(wb).blocked().has_value());
  EXPECT_EQ(Extent(wb), (PivotRect{to, before.col0, to + height - 1U, before.col1}));
  EXPECT_TRUE(wb.tables().empty());
  EXPECT_TRUE(wb.sheet(kPivotSheet).merges().empty());
  EXPECT_FALSE(At(wb, to + 1U, before.col0 + 1U).is_text());
  EXPECT_EQ(wb.sheet(kPivotSheet).cell_at(before.row0, before.col0), nullptr);
}

TEST(PivotBlocked, MovePastesOverAnArrayOrADataTableInsideTheDestination) {
  // Every case targets A21, whose destination is A21:B25.
  struct Case {
    const char* name;
    void (*place)(Workbook& wb);
  };
  const Case cases[] = {
      {"legacy array",
       [](Workbook& wb) {
         ASSERT_TRUE(static_cast<bool>(wb.set_array_formula(kPivotSheet, 21U, 0U, 2U, 2U, "={1,2;3,4}")));
       }},
      {"full data table",
       [](Workbook& wb) {
         // Formula B22, inputs A23:A24, results B23:B24 reading column input U1.
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(kPivotSheet, 21U, 1U, "=U1*2")));
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, 22U, 0U, Value::number(1.0))));
         ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kPivotSheet, 23U, 0U, Value::number(2.0))));
         ASSERT_TRUE(static_cast<bool>(
             wb.add_data_table(kPivotSheet, MergeRange{22U, 1U, 23U, 1U}, DataTableInput{}, InputAt(0U, 20U))));
       }},
      {"data-table results only",
       [](Workbook& wb) {
         ASSERT_TRUE(static_cast<bool>(
             wb.add_data_table(kPivotSheet, MergeRange{22U, 1U, 23U, 1U}, DataTableInput{}, InputAt(0U, 20U))));
       }},
  };
  for (const Case& entry : cases) {
    SCOPED_TRACE(entry.name);
    Workbook wb = Book();
    const PivotRect before = Extent(wb);
    const std::uint32_t height = before.row1 - before.row0 + 1U;
    entry.place(wb);
    const Expected<void, Error> moved = wb.move_pivot(kPivotSheet, 0U, 20U, before.col0);
    ASSERT_TRUE(static_cast<bool>(moved)) << moved.error().message;
    // A move never blocks: what lies wholly inside the destination is overwritten.
    EXPECT_FALSE(PivotAt(wb).blocked().has_value());
    EXPECT_EQ(Extent(wb), (PivotRect{20U, before.col0, 20U + height - 1U, before.col1}));
    EXPECT_FALSE(wb.array_formula_at(kPivotSheet, 21U, 0U).has_value());
    EXPECT_EQ(wb.data_table_at(kPivotSheet, 22U, 1U), nullptr);
    EXPECT_TRUE(wb.sheet(kPivotSheet).data_tables().empty());
  }
}

TEST(PivotBlocked, MoveIsRefusedOntoAPivotAStraddlingArrayOrDataTableOrPastTheEdge) {
  struct Case {
    const char* name;
    void (*place)(Workbook& wb);
    std::uint32_t row;
    std::uint32_t col;
  };
  // Every case but the edges targets A21, whose destination is A21:B25.
  const Case cases[] = {
      {"other pivot",
       [](Workbook& wb) {
         ASSERT_TRUE(static_cast<bool>(wb.add_pivot(kPivotSheet, RegionPivot("PT2", 21U, 0U))));
         ASSERT_TRUE(static_cast<bool>(wb.project_pivot(kPivotSheet, 1U, ReportPlacement::kRefuse)));
       },
       20U, 0U},
      {"array crossing the right edge",
       [](Workbook& wb) {
         ASSERT_TRUE(static_cast<bool>(wb.set_array_formula(kPivotSheet, 22U, 1U, 1U, 2U, "={1,2}")));
       },
       20U, 0U},
      {"data table crossing the right edge",
       [](Workbook& wb) {
         ASSERT_TRUE(static_cast<bool>(
             wb.add_data_table(kPivotSheet, MergeRange{23U, 1U, 23U, 2U}, DataTableInput{}, InputAt(0U, 20U))));
       },
       20U, 0U},
      {"past the bottom edge", [](Workbook&) {}, Sheet::kMaxRows - 2U, 0U},
      {"past the right edge", [](Workbook&) {}, 2U, Sheet::kMaxCols - 1U},
  };
  for (const Case& entry : cases) {
    SCOPED_TRACE(entry.name);
    Workbook wb = Book();
    const PivotRect before = Extent(wb);
    const Value label = At(wb, before.row0 + 1U, before.col0);
    entry.place(wb);
    const Expected<void, Error> moved = wb.move_pivot(kPivotSheet, 0U, entry.row, entry.col);
    ASSERT_FALSE(static_cast<bool>(moved));
    EXPECT_EQ(moved.error().code, FormulonErrorCode::kPreconditionFailed);
    EXPECT_EQ(PivotAt(wb).anchor_row(), before.row0);
    EXPECT_EQ(PivotAt(wb).anchor_col(), before.col0);
    EXPECT_EQ(Extent(wb), before);
    EXPECT_EQ(At(wb, before.row0 + 1U, before.col0).as_text(), label.as_text());
  }
  // The straddling array is left intact.
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(wb.set_array_formula(kPivotSheet, 22U, 1U, 1U, 2U, "={1,2}")));
  ASSERT_FALSE(static_cast<bool>(wb.move_pivot(kPivotSheet, 0U, 20U, 0U)));
  EXPECT_TRUE(wb.array_formula_at(kPivotSheet, 22U, 2U).has_value());
}

TEST(PivotBlocked, MovingABlockedReportCarriesOnlyItsAnchor) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, control.row1, control.col1, "blocker")));
  ASSERT_TRUE(static_cast<bool>(AddDataField(wb)));
  ASSERT_TRUE(PivotAt(wb).blocked().has_value());
  const Expected<void, Error> moved = wb.move_pivot(kPivotSheet, 0U, 30U, 0U);
  ASSERT_TRUE(static_cast<bool>(moved)) << moved.error().message;
  // Free at the destination, it renders there; the old anchor is gone.
  EXPECT_FALSE(PivotAt(wb).blocked().has_value());
  EXPECT_EQ(Extent(wb), (PivotRect{30U, control.col0, 30U + control.row1 - control.row0, control.col1}));
  EXPECT_EQ(wb.sheet(kPivotSheet).cell_at(control.row0, control.col0), nullptr);
}

// ---------------------------------------------------------------------------
// What a render writes, and the empty report
// ---------------------------------------------------------------------------

TEST(PivotBlocked, AFormatOnlyCellDoesNotBlockAndLosesItsFormatWhenTheReportRenders) {
  const PivotRect control = ControlExtent();
  Workbook wb = Book();
  CellXf formatted;
  formatted.num_fmt_id = 2;  // 0.00
  formatted.apply_number_format = true;
  const std::uint32_t custom = intern_cell_xf(wb.mutable_styles(), formatted);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_xf_index(kPivotSheet, control.row1, control.col1, custom)));
  const Expected<void, Error> edited = AddDataField(wb);
  ASSERT_TRUE(static_cast<bool>(edited)) << edited.error().message;
  EXPECT_FALSE(PivotAt(wb).blocked().has_value());
  EXPECT_EQ(PivotAt(wb).projected_extent(), std::optional<PivotRect>(control));
  const Cell* cell = wb.sheet(kPivotSheet).cell_at(control.row1, control.col1);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(At(wb, control.row1, control.col1).is_number());
  EXPECT_NE(cell->xf_index, custom);
  const CellXf& xf = wb.styles().cell_xfs[cell->xf_index];
  EXPECT_EQ(xf.num_fmt_id, 0U);
  EXPECT_EQ(xf.fill_index, 0U);
  EXPECT_EQ(xf.border_index, 0U);
}

TEST(PivotBlocked, FieldlessReportWritesNoCellValues) {
  // Region / Amt over A1:B2 and a pivot with no fields at A5, on one sheet.
  Workbook wb = Workbook::create_for_profile(mac_365_en_us_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, 0U, 0U, "Region")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, 0U, 1U, "Amt")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, 1U, 0U, "East")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kDataSheet, 1U, 1U, Value::number(5.0))));
  auto cache = std::make_unique<PivotCache>();
  cache->set_cache_id(kCacheId);
  cache->mutable_fields().push_back(PivotCacheField{"Region", {Value::text("East")}});
  cache->mutable_fields().push_back(PivotCacheField{"Amt", {}});
  PivotCacheRecord record;
  record.cells = {Value::number(0.0), Value::number(5.0)};
  record.cell_is_index = {true, false};
  cache->mutable_records().push_back(std::move(record));
  cache->mutable_worksheet_source().present = true;
  cache->mutable_worksheet_source().ref = "A1:B2";
  cache->mutable_worksheet_source().sheet = wb.sheet(kDataSheet).name();
  wb.add_pivot_cache(std::move(cache));
  auto table = std::make_unique<PivotTable>();
  table->set_name("PivotTable1");
  table->set_pivot_cache_id(kCacheId);
  table->set_anchor(4U, 0U, 1U, 1U);
  const Expected<std::size_t, Error> added = wb.add_pivot(kDataSheet, std::move(table));
  ASSERT_TRUE(static_cast<bool>(added)) << added.error().message;
  const Expected<void, Error> projected = wb.project_pivot(kDataSheet, added.value(), ReportPlacement::kRefuse);
  ASSERT_TRUE(static_cast<bool>(projected)) << projected.error().message;
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(kDataSheet, 0U, 4U, "=ISBLANK(A5)")));
  Recalc(wb);
  // A5:C22 holds styled empty cells only; the anchor carries no "Row Labels" caption.
  for (std::uint32_t r = 4U; r <= 21U; ++r) {
    for (std::uint32_t c = 0U; c <= 2U; ++c) {
      const Cell* cell = wb.sheet(kDataSheet).cell_at(r, c);
      ASSERT_NE(cell, nullptr) << A1(r, c);
      EXPECT_TRUE(cell->formula_text.empty()) << A1(r, c);
      EXPECT_TRUE(cell->cached_value.is_blank()) << A1(r, c);
    }
  }
  EXPECT_TRUE(At(wb, 4U, 0U, kDataSheet).is_blank());
  const Value isblank = At(wb, 0U, 4U, kDataSheet);
  ASSERT_TRUE(isblank.is_boolean());
  EXPECT_TRUE(isblank.as_boolean());
}

// ---------------------------------------------------------------------------
// Two reports competing for the same new cells
// ---------------------------------------------------------------------------

std::unique_ptr<PivotTable> RegionColumnsPivot(const char* name, std::uint32_t row, std::uint32_t col) {
  std::unique_ptr<PivotTable> table = RegionPivot(name, row, col);
  table->mutable_fields()[kRegion].axis = PivotAxis::Col;
  table->mutable_row_field_order().clear();
  table->mutable_col_field_order() = {kRegion};
  return table;
}

std::size_t IndexOf(Workbook& wb, const char* name) {
  for (std::size_t i = 0; i < wb.sheet(kPivotSheet).pivot_tables().size(); ++i) {
    if (wb.sheet(kPivotSheet).pivot_tables()[i]->name() == name) {
      return i;
    }
  }
  ADD_FAILURE() << "no pivot " << name;
  return 0U;
}

void AddAndProject(Workbook& wb, std::unique_ptr<PivotTable> table) {
  const Expected<std::size_t, Error> added = wb.add_pivot(kPivotSheet, std::move(table));
  ASSERT_TRUE(static_cast<bool>(added)) << added.error().message;
  ASSERT_TRUE(static_cast<bool>(wb.project_pivot(kPivotSheet, added.value(), ReportPlacement::kRefuse)));
}

TEST(PivotBlocked, NewRangesMeetingOnlyEachOtherLeaveTheEarlierCreatedReportRendered) {
  for (const bool columns_first : {false, true}) {
    for (const bool swapped : {false, true}) {
      SCOPED_TRACE(std::string(columns_first ? "columns pivot created first" : "rows pivot created first") +
                   (swapped ? ", sheet list swapped" : ""));
      Workbook wb = Book();
      // PT1 stays at A3, out of the way; "Rows" grows down and "Cols" grows right on a fourth region.
      if (columns_first) {
        AddAndProject(wb, RegionColumnsPivot("Cols", 40U, 0U));
        AddAndProject(wb, RegionPivot("Rows", 2U, 12U));
      } else {
        AddAndProject(wb, RegionPivot("Rows", 2U, 12U));
        AddAndProject(wb, RegionColumnsPivot("Cols", 40U, 0U));
      }
      // "Cols" goes just below "Rows" and ends one column left of it: neither new range reaches the other's
      // current cells, but both claim the cell diagonally between them.
      const PivotRect rows = Extent(wb, IndexOf(wb, "Rows"));
      const PivotRect cols = Extent(wb, IndexOf(wb, "Cols"));
      const std::uint32_t width = cols.col1 - cols.col0 + 1U;
      ASSERT_TRUE(
          static_cast<bool>(wb.move_pivot(kPivotSheet, IndexOf(wb, "Cols"), rows.row1 + 1U, rows.col0 - width)));
      if (swapped) {
        std::swap(wb.sheet(kPivotSheet).mutable_pivot_tables()[1], wb.sheet(kPivotSheet).mutable_pivot_tables()[2]);
      }
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, 3U, 0U, "West")));
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, 1U, 0U, "Central")));
      const Expected<void, Error> refreshed = wb.refresh_pivot_cache(kCacheId, eval::default_registry());
      ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
      Recalc(wb);
      const PivotTable& earlier = PivotAt(wb, IndexOf(wb, columns_first ? "Cols" : "Rows"));
      const PivotTable& later = PivotAt(wb, IndexOf(wb, columns_first ? "Rows" : "Cols"));
      EXPECT_FALSE(earlier.blocked().has_value());
      ASSERT_TRUE(later.blocked().has_value());
      EXPECT_TRUE(HoldsSpill(wb, later.anchor_row(), later.anchor_col()));
    }
  }
}

}  // namespace
}  // namespace formulon
