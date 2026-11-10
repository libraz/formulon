// Where a PivotTable report may be drawn, through the C ABI. An edit, a
// refresh or a grouping whose new range is occupied or runs past the sheet
// edge is applied and leaves the report blocked behind `#SPILL!`; creating a
// report there is refused; moving one is a cut and paste that overwrites the
// destination or is refused. `fm_sheet_get_spill` and
// `fm_sheet_get_pivot_spill` report the range a blocked report needs.
//
// Each case first runs its trigger on a book without the occupant (the
// control) to learn the range the report grows to, then places the occupant at
// that range's bottom-right cell, outside the report's current range.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "sheet.h"
#include "value.h"

namespace {

constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);
constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

constexpr std::size_t kSheet = 0U;
constexpr std::size_t kSourceSheet = 1U;
constexpr std::uint32_t kMaxRows = formulon::Sheet::kMaxRows;
constexpr std::uint32_t kMaxCols = formulon::Sheet::kMaxCols;

// The report's anchor; a page field sits two rows above it.
constexpr std::uint32_t kAnchorRow = 4U;
constexpr std::uint32_t kAnchorCol = 3U;
// Where a move takes the anchor, and where a created report starts.
constexpr std::uint32_t kMoveRow = 40U;
constexpr std::uint32_t kCreateRow = 40U;
// A report with no field is a 3 x 18 placeholder.
constexpr std::uint32_t kPlaceholderRows = 18U;
constexpr std::uint32_t kPlaceholderCols = 3U;

// Cache fields, in source column order.
constexpr std::size_t kDateField = 0U;
constexpr std::size_t kRegionField = 1U;
constexpr std::size_t kAmountField = 2U;
constexpr std::size_t kItemField = 3U;
constexpr std::size_t kKindField = 4U;
constexpr std::uint32_t kRecords = 20U;

enum class Trigger { kRefresh, kEditField, kEditOption, kGrouping, kMove, kCreate };
enum class Occupant {
  kNone,
  kValue,
  kEmptyStringFormula,
  kTable,
  kOtherPivot,
  kSpillRange,
  kEmptyMerge,
  kLegacyArray,
  kDataTableResult,
  kPastEdge,
};
enum class Kind { kNormal, kWithPageFields, kFieldless };
enum class Outcome { kRendered, kBlocked, kBlockedPastEdge, kRefused, kMoved };

const char* Name(Trigger trigger) {
  switch (trigger) {
    case Trigger::kRefresh:
      return "refresh";
    case Trigger::kEditField:
      return "edit_field";
    case Trigger::kEditOption:
      return "edit_option";
    case Trigger::kGrouping:
      return "grouping";
    case Trigger::kMove:
      return "move";
    case Trigger::kCreate:
      return "create";
  }
  return "?";
}

const char* Name(Occupant occupant) {
  switch (occupant) {
    case Occupant::kNone:
      return "none";
    case Occupant::kValue:
      return "value";
    case Occupant::kEmptyStringFormula:
      return "empty_string_formula";
    case Occupant::kTable:
      return "table";
    case Occupant::kOtherPivot:
      return "other_pivot";
    case Occupant::kSpillRange:
      return "spill_range";
    case Occupant::kEmptyMerge:
      return "empty_merge";
    case Occupant::kLegacyArray:
      return "legacy_array";
    case Occupant::kDataTableResult:
      return "data_table_result";
    case Occupant::kPastEdge:
      return "past_edge";
  }
  return "?";
}

const char* Name(Kind kind) {
  switch (kind) {
    case Kind::kNormal:
      return "normal";
    case Kind::kWithPageFields:
      return "with_page_fields";
    case Kind::kFieldless:
      return "fieldless";
  }
  return "?";
}

struct Case {
  Trigger trigger;
  Occupant occupant;
  Kind kind;
  Outcome expected;
};

// The pairwise placement model: every pair of trigger x occupant x pivot kind
// appears in some row. A fieldless report keeps its placeholder under refresh,
// option edits and grouping, so it pairs only with create, move and field
// edits. The C ABI creates every pivot without fields, so a created report is
// the placeholder whatever the row's kind.
constexpr Case kCases[] = {
    {Trigger::kMove, Occupant::kEmptyMerge, Kind::kNormal, Outcome::kMoved},
    {Trigger::kMove, Occupant::kSpillRange, Kind::kFieldless, Outcome::kMoved},
    {Trigger::kEditField, Occupant::kEmptyMerge, Kind::kFieldless, Outcome::kBlocked},
    {Trigger::kRefresh, Occupant::kTable, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kEditField, Occupant::kEmptyStringFormula, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kGrouping, Occupant::kDataTableResult, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kGrouping, Occupant::kNone, Kind::kNormal, Outcome::kRendered},
    {Trigger::kRefresh, Occupant::kOtherPivot, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kGrouping, Occupant::kEmptyMerge, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kMove, Occupant::kTable, Kind::kWithPageFields, Outcome::kMoved},
    {Trigger::kEditField, Occupant::kSpillRange, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kCreate, Occupant::kSpillRange, Kind::kWithPageFields, Outcome::kRefused},
    {Trigger::kEditField, Occupant::kNone, Kind::kFieldless, Outcome::kRendered},
    {Trigger::kCreate, Occupant::kOtherPivot, Kind::kFieldless, Outcome::kRefused},
    {Trigger::kEditOption, Occupant::kNone, Kind::kWithPageFields, Outcome::kRendered},
    {Trigger::kMove, Occupant::kPastEdge, Kind::kWithPageFields, Outcome::kRefused},
    {Trigger::kRefresh, Occupant::kValue, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kEditField, Occupant::kTable, Kind::kFieldless, Outcome::kBlocked},
    {Trigger::kGrouping, Occupant::kSpillRange, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kGrouping, Occupant::kLegacyArray, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kMove, Occupant::kDataTableResult, Kind::kFieldless, Outcome::kMoved},
    {Trigger::kMove, Occupant::kEmptyStringFormula, Kind::kFieldless, Outcome::kMoved},
    {Trigger::kEditOption, Occupant::kOtherPivot, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kCreate, Occupant::kLegacyArray, Kind::kFieldless, Outcome::kRefused},
    {Trigger::kMove, Occupant::kNone, Kind::kWithPageFields, Outcome::kMoved},
    {Trigger::kCreate, Occupant::kEmptyStringFormula, Kind::kNormal, Outcome::kRefused},
    {Trigger::kCreate, Occupant::kTable, Kind::kFieldless, Outcome::kRefused},
    {Trigger::kEditOption, Occupant::kLegacyArray, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kCreate, Occupant::kNone, Kind::kWithPageFields, Outcome::kRendered},
    {Trigger::kEditOption, Occupant::kDataTableResult, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kEditField, Occupant::kLegacyArray, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kCreate, Occupant::kPastEdge, Kind::kFieldless, Outcome::kRefused},
    {Trigger::kMove, Occupant::kLegacyArray, Kind::kFieldless, Outcome::kMoved},
    {Trigger::kEditOption, Occupant::kEmptyMerge, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kGrouping, Occupant::kPastEdge, Kind::kNormal, Outcome::kBlockedPastEdge},
    {Trigger::kGrouping, Occupant::kValue, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kRefresh, Occupant::kDataTableResult, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kRefresh, Occupant::kLegacyArray, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kEditField, Occupant::kDataTableResult, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kRefresh, Occupant::kPastEdge, Kind::kWithPageFields, Outcome::kBlockedPastEdge},
    {Trigger::kCreate, Occupant::kValue, Kind::kNormal, Outcome::kRefused},
    {Trigger::kCreate, Occupant::kEmptyMerge, Kind::kNormal, Outcome::kRefused},
    {Trigger::kEditField, Occupant::kValue, Kind::kFieldless, Outcome::kBlocked},
    {Trigger::kEditField, Occupant::kOtherPivot, Kind::kFieldless, Outcome::kBlocked},
    {Trigger::kEditOption, Occupant::kPastEdge, Kind::kWithPageFields, Outcome::kBlockedPastEdge},
    {Trigger::kMove, Occupant::kOtherPivot, Kind::kFieldless, Outcome::kRefused},
    {Trigger::kEditOption, Occupant::kTable, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kCreate, Occupant::kDataTableResult, Kind::kNormal, Outcome::kRefused},
    {Trigger::kGrouping, Occupant::kEmptyStringFormula, Kind::kWithPageFields, Outcome::kBlocked},
    {Trigger::kRefresh, Occupant::kEmptyStringFormula, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kEditOption, Occupant::kValue, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kEditOption, Occupant::kEmptyStringFormula, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kRefresh, Occupant::kEmptyMerge, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kEditOption, Occupant::kSpillRange, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kRefresh, Occupant::kSpillRange, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kEditField, Occupant::kPastEdge, Kind::kWithPageFields, Outcome::kBlockedPastEdge},
    {Trigger::kGrouping, Occupant::kOtherPivot, Kind::kNormal, Outcome::kBlocked},
    {Trigger::kMove, Occupant::kValue, Kind::kFieldless, Outcome::kMoved},
    {Trigger::kRefresh, Occupant::kNone, Kind::kNormal, Outcome::kRendered},
    {Trigger::kGrouping, Occupant::kTable, Kind::kNormal, Outcome::kBlocked},
};
static_assert(sizeof(kCases) / sizeof(kCases[0]) == 60U, "the pairwise model has 60 rows");

bool Inside(const fm_cell_rect_t& rect, std::uint32_t row, std::uint32_t col) {
  return row >= rect.first_row && row <= rect.last_row && col >= rect.first_col && col <= rect.last_col;
}

void ExpectRect(const fm_cell_rect_t& actual, const fm_cell_rect_t& expected) {
  EXPECT_EQ(actual.first_row, expected.first_row);
  EXPECT_EQ(actual.first_col, expected.first_col);
  EXPECT_EQ(actual.last_row, expected.last_row);
  EXPECT_EQ(actual.last_col, expected.last_col);
}

fm_value_t ValueAt(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  fm_value_t value{};
  EXPECT_EQ(fm_sheet_get_value(wb, kSheet, row, col, &value), 0) << fm_get_last_error_message();
  return value;
}

bool IsSpillError(const fm_value_t& value) {
  return value.kind == FM_VALUE_KIND_ERROR &&
         value.u.error_code == static_cast<std::int32_t>(formulon::ErrorCode::Spill);
}

std::size_t PivotCount(const fm_workbook_t* wb) {
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_pivot_count(wb, kSheet, &count), 0);
  return count;
}

// The report's range as its layout states it, page fields included.
fm_cell_rect_t Bounds(fm_workbook_t* wb, std::size_t pivot) {
  PivotCellsGuard layout;
  EXPECT_EQ(fm_sheet_get_pivot_layout(wb, kSheet, pivot, &layout.handle), 0) << fm_get_last_error_message();
  std::int32_t present = 0;
  fm_cell_rect_t rect{};
  EXPECT_EQ(fm_pivot_report_get_bounds(layout.handle, &present, &rect), 0);
  EXPECT_EQ(present, 1);
  return rect;
}

// `fm_sheet_get_pivot_spill`, as nullopt when nothing is reported.
std::optional<fm_spill_info_t> PivotSpill(const fm_workbook_t* wb, std::size_t pivot) {
  std::int32_t present = -1;
  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
  EXPECT_EQ(fm_sheet_get_pivot_spill(wb, kSheet, pivot, &present, &info), 0) << fm_get_last_error_message();
  if (present == 0) {
    return std::nullopt;
  }
  EXPECT_EQ(present, 1);
  return info;
}

// `fm_sheet_get_spill` at a cell, as nullopt when nothing is reported.
std::optional<fm_spill_info_t> CellSpill(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  std::int32_t present = -1;
  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
  EXPECT_EQ(fm_sheet_get_spill(wb, kSheet, row, col, &present, &info), 0) << fm_get_last_error_message();
  if (present == 0) {
    return std::nullopt;
  }
  return info;
}

bool HasArrayAt(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  std::int32_t present = 0;
  fm_array_formula_info_t info = FM_STRUCT_INIT(fm_array_formula_info_t);
  EXPECT_EQ(fm_sheet_get_array_formula(wb, kSheet, row, col, &present, &info), 0);
  return present != 0;
}

bool HasDataTableAt(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  std::int32_t present = 0;
  fm_data_table_t info = FM_STRUCT_INIT(fm_data_table_t);
  EXPECT_EQ(fm_sheet_get_data_table(wb, kSheet, row, col, &present, &info), 0);
  return present != 0;
}

// What an occupant leaves at the corner cell and on the sheet.
struct Snapshot {
  std::int32_t corner_kind = 0;
  std::string corner_text;
  std::size_t merges = 0;
  std::size_t tables = 0;
  std::size_t pivots = 0;
  bool array = false;
  bool data_table = false;
  bool spill = false;

  bool operator==(const Snapshot& o) const {
    return corner_kind == o.corner_kind && corner_text == o.corner_text && merges == o.merges && tables == o.tables &&
           pivots == o.pivots && array == o.array && data_table == o.data_table && spill == o.spill;
  }
};

Snapshot Take(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  Snapshot snap;
  const fm_value_t corner = ValueAt(wb, row, col);
  snap.corner_kind = corner.kind;
  if (corner.kind == FM_VALUE_KIND_TEXT && corner.u.text != nullptr) {
    snap.corner_text = corner.u.text;
  }
  EXPECT_EQ(fm_sheet_get_merge_count(wb, kSheet, &snap.merges), 0);
  EXPECT_EQ(fm_sheet_get_table_count(wb, kSheet, &snap.tables), 0);
  snap.pivots = PivotCount(wb);
  snap.array = HasArrayAt(wb, row, col);
  snap.data_table = HasDataTableAt(wb, row, col);
  snap.spill = CellSpill(wb, row, col).has_value();
  return snap;
}

// Sheet1 holds the reports; Src!A1:E21 is the source of the main cache, with
// three dates (one in each of 1901, 1902 and 1903), two regions, twenty items
// and two kinds. A second cache feeds the pivots that serve as occupants.
struct Book {
  WorkbookGuard wb;
  std::uint32_t cache = 0;
  std::uint32_t other_cache = 0;

  void Build() {
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    std::size_t source = 99;
    ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Src", append_position(wb.handle), &source), 0)
        << fm_get_last_error_message();
    ASSERT_EQ(source, kSourceSheet);
    const char* const headers[] = {"Date", "Region", "Amount", "Item", "Kind"};
    for (std::uint32_t c = 0; c < 5U; ++c) {
      ASSERT_EQ(fm_sheet_set_text(wb.handle, kSourceSheet, 0U, c, headers[c]), 0);
    }
    ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0U, &cache), 0) << fm_get_last_error_message();
    ASSERT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache, /*present=*/1, "A1:E21", "Src", nullptr), 0);
    for (const char* header : headers) {
      ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache, header, ignored_index()), 0);
    }
    const double dates[] = {400.0, 800.0, 1200.0};
    for (std::uint32_t i = 0; i < kRecords; ++i) {
      const double date = dates[i % 3U];
      const char* region = i % 2U == 0U ? "East" : "West";
      const std::string item = std::string(i < 9U ? "I0" : "I") + std::to_string(i + 1U);
      const char* kind = i % 4U < 2U ? "A" : "B";
      const double amount = static_cast<double>(i + 1U);
      const std::uint32_t row = i + 1U;
      ASSERT_EQ(fm_sheet_set_number(wb.handle, kSourceSheet, row, 0U, date), 0);
      ASSERT_EQ(fm_sheet_set_text(wb.handle, kSourceSheet, row, 1U, region), 0);
      ASSERT_EQ(fm_sheet_set_number(wb.handle, kSourceSheet, row, 2U, amount), 0);
      ASSERT_EQ(fm_sheet_set_text(wb.handle, kSourceSheet, row, 3U, item.c_str()), 0);
      ASSERT_EQ(fm_sheet_set_text(wb.handle, kSourceSheet, row, 4U, kind), 0);
      std::size_t record = 0;
      ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache, &record), 0);
      ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache, record, kDateField, date), 0);
      ASSERT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, cache, record, kRegionField, region), 0);
      ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache, record, kAmountField, amount), 0);
      ASSERT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, cache, record, kItemField, item.c_str()), 0);
      ASSERT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, cache, record, kKindField, kind), 0);
    }
    ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0U, &other_cache), 0) << fm_get_last_error_message();
    ASSERT_EQ(SetPivotCacheWorksheetSource(wb.handle, other_cache, /*present=*/1, "G1:G2", "Src", nullptr), 0);
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, other_cache, "X", ignored_index()), 0);
    std::size_t record = 0;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, other_cache, &record), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, other_cache, record, 0U, "x"), 0);
  }

  // The report under test at (`row`, `col`): rows by Date with Sum of Amount, plus Region as a page field, or no
  // field at all.
  void AddMainPivot(Kind kind, std::size_t* out_pivot) {
    ASSERT_EQ(AddPivot(wb.handle, kSheet, "Main", cache, kAnchorRow, kAnchorCol, out_pivot), 0)
        << fm_get_last_error_message();
    if (kind == Kind::kFieldless) {
      return;
    }
    ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, kSheet, *out_pivot, kDateField, FM_PIVOT_AXIS_ROW), 0)
        << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, kSheet, *out_pivot, kAmountField, FM_PIVOT_AXIS_VALUE), 0)
        << fm_get_last_error_message();
    fm_pivot_data_field_t sum = FM_STRUCT_INIT(fm_pivot_data_field_t);
    sum.name = "Sum of Amount";
    sum.field_index = static_cast<std::uint32_t>(kAmountField);
    sum.aggregation = FM_PIVOT_AGGREGATION_SUM;
    sum.number_format = "";
    sum.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
    ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, kSheet, *out_pivot, &sum, ignored_index()), 0)
        << fm_get_last_error_message();
    if (kind == Kind::kWithPageFields) {
      ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, kSheet, *out_pivot, kRegionField, FM_PIVOT_AXIS_PAGE), 0)
          << fm_get_last_error_message();
    }
  }

  // Runs an edit trigger on `pivot`.
  fm_status_t Edit(Trigger trigger, Kind kind, std::size_t pivot) {
    switch (trigger) {
      case Trigger::kRefresh: {
        // A fourth date adds a row.
        const fm_status_t rc = fm_sheet_set_number(wb.handle, kSourceSheet, 1U, 0U, 1600.0);
        return rc != 0 ? rc : fm_workbook_refresh_pivot_cache(wb.handle, cache);
      }
      case Trigger::kEditField:
        // Twenty items outgrow the placeholder; two kinds on columns widen and lengthen the report.
        return kind == Kind::kFieldless
                   ? fm_sheet_set_pivot_field_axis(wb.handle, kSheet, pivot, kItemField, FM_PIVOT_AXIS_ROW)
                   : fm_sheet_set_pivot_field_axis(wb.handle, kSheet, pivot, kKindField, FM_PIVOT_AXIS_COLUMN);
      case Trigger::kEditOption: {
        fm_pivot_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_options_patch_t);
        patch.grid_drop_zones_op = FM_PATCH_OP_SET;
        patch.grid_drop_zones = 1;
        return fm_sheet_update_pivot_options(wb.handle, kSheet, pivot, &patch);
      }
      case Trigger::kGrouping:
        // A year row above each month row.
        return SetPivotFieldDateGroup(wb.handle, kSheet, pivot, kDateField,
                                      FM_PIVOT_PERIOD_MONTHS | FM_PIVOT_PERIOD_YEARS, FM_PIVOT_CALENDAR_GREGORIAN, 1U,
                                      std::numeric_limits<double>::quiet_NaN(),
                                      std::numeric_limits<double>::quiet_NaN());
      case Trigger::kMove:
      case Trigger::kCreate:
        break;
    }
    return kInvalidArgument;
  }

  // Puts `occupant` at the bottom-right cell of `target`. `current` is the report's range now, which a spill keeps
  // clear of: the spill runs along the target's last row from the column left of it, or, when the target grows only
  // to the right, down its last column from the row above it.
  void Place(Occupant occupant, const fm_cell_rect_t& target, const std::optional<fm_cell_rect_t>& current) {
    const std::uint32_t r = target.last_row;
    const std::uint32_t c = target.last_col;
    switch (occupant) {
      case Occupant::kValue:
        ASSERT_EQ(fm_sheet_set_text(wb.handle, kSheet, r, c, "x"), 0) << fm_get_last_error_message();
        break;
      case Occupant::kEmptyStringFormula:
        ASSERT_EQ(fm_sheet_set_formula(wb.handle, kSheet, r, c, "=\"\""), 0) << fm_get_last_error_message();
        break;
      case Occupant::kTable: {
        fm_table_spec_t spec = FM_STRUCT_INIT(fm_table_spec_t);
        spec.range = fm_cell_rect_t{r, c, r + 1U, c};
        spec.name = "Occupant";
        ASSERT_EQ(fm_sheet_add_table(wb.handle, kSheet, &spec, ignored_index()), 0) << fm_get_last_error_message();
        break;
      }
      case Occupant::kOtherPivot:
        ASSERT_EQ(AddPivot(wb.handle, kSheet, "Occupant", other_cache, r, c, ignored_index()), 0)
            << fm_get_last_error_message();
        break;
      case Occupant::kSpillRange: {
        const bool along_row = !current.has_value() || r > current->last_row;
        const std::uint32_t row = along_row ? r : target.first_row - 1U;
        const std::uint32_t col = along_row ? target.first_col - 1U : c;
        const std::string formula = along_row
                                        ? "=SEQUENCE(1," + std::to_string(target.last_col - target.first_col + 2U) + ")"
                                        : "=SEQUENCE(" + std::to_string(target.last_row - target.first_row + 2U) + ")";
        ASSERT_EQ(fm_sheet_set_formula(wb.handle, kSheet, row, col, formula.c_str()), 0) << fm_get_last_error_message();
        ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
        const std::optional<fm_spill_info_t> spill = CellSpill(wb.handle, r, c);
        ASSERT_TRUE(spill.has_value());
        ASSERT_EQ(spill->blocked, 0);
        break;
      }
      case Occupant::kEmptyMerge: {
        const fm_cell_rect_t merge{r, c, r, c + 1U};
        ASSERT_EQ(fm_sheet_add_merge(wb.handle, kSheet, &merge, nullptr), 0) << fm_get_last_error_message();
        break;
      }
      case Occupant::kLegacyArray: {
        const fm_cell_rect_t block{r, c, r, c};
        ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, kSheet, &block, "=1"), 0) << fm_get_last_error_message();
        break;
      }
      case Occupant::kDataTableResult: {
        fm_data_table_t spec = FM_STRUCT_INIT(fm_data_table_t);
        spec.rect = fm_cell_rect_t{r, c, r, c};
        spec.has_column_input = 1;
        spec.column_input_row = 0U;
        spec.column_input_col = 0U;
        ASSERT_EQ(fm_sheet_add_data_table(wb.handle, kSheet, &spec, nullptr), 0) << fm_get_last_error_message();
        break;
      }
      case Occupant::kNone:
      case Occupant::kPastEdge:
        break;
    }
  }
};

// An edit of an existing report: refresh, a field edit, an option edit or a grouping.
void RunEdit(const Case& c) {
  // The control gives the range the report grows to.
  fm_cell_rect_t current{};
  fm_cell_rect_t grown{};
  {
    Book control;
    control.Build();
    std::size_t pivot = 0;
    control.AddMainPivot(c.kind, &pivot);
    current = Bounds(control.wb.handle, pivot);
    ASSERT_EQ(control.Edit(c.trigger, c.kind, pivot), 0) << fm_get_last_error_message();
    ASSERT_FALSE(PivotSpill(control.wb.handle, pivot).has_value());
    grown = Bounds(control.wb.handle, pivot);
    ASSERT_FALSE(Inside(current, grown.last_row, grown.last_col)) << "the trigger must grow the report";
  }

  Book book;
  book.Build();
  std::size_t pivot = 0;
  book.AddMainPivot(c.kind, &pivot);
  std::uint32_t anchor_row = kAnchorRow;
  std::uint32_t anchor_col = kAnchorCol;
  if (c.occupant == Occupant::kPastEdge) {
    // Move the report flush with the edge it grows past.
    if (grown.last_row > current.last_row) {
      anchor_row += (kMaxRows - 1U) - current.last_row;
    } else {
      anchor_col += (kMaxCols - 1U) - current.last_col;
    }
    ASSERT_EQ(fm_sheet_set_pivot_anchor(book.wb.handle, kSheet, pivot, anchor_row, anchor_col), 0)
        << fm_get_last_error_message();
    current = Bounds(book.wb.handle, pivot);
  } else {
    book.Place(c.occupant, grown, current);
  }
  const Snapshot before = Take(book.wb.handle, grown.last_row, grown.last_col);

  ASSERT_EQ(book.Edit(c.trigger, c.kind, pivot), 0) << fm_get_last_error_message();
  std::int32_t pending = -1;
  ASSERT_EQ(fm_sheet_is_pivot_projection_pending(book.wb.handle, kSheet, pivot, &pending), 0);
  EXPECT_EQ(pending, 0);
  const fm_value_t anchor = ValueAt(book.wb.handle, anchor_row, anchor_col);
  const std::optional<fm_spill_info_t> spill = PivotSpill(book.wb.handle, pivot);
  switch (c.expected) {
    case Outcome::kRendered:
      EXPECT_FALSE(spill.has_value());
      ExpectRect(Bounds(book.wb.handle, pivot), grown);
      EXPECT_FALSE(IsSpillError(anchor));
      break;
    case Outcome::kBlocked: {
      ASSERT_TRUE(spill.has_value());
      EXPECT_EQ(spill->blocked, 1);
      ExpectRect(spill->rect, grown);
      EXPECT_TRUE(IsSpillError(anchor));
      // The occupant is left as it was.
      EXPECT_TRUE(Take(book.wb.handle, grown.last_row, grown.last_col) == before);
      const std::optional<fm_spill_info_t> at_anchor = CellSpill(book.wb.handle, anchor_row, anchor_col);
      ASSERT_TRUE(at_anchor.has_value());
      EXPECT_EQ(at_anchor->blocked, 1);
      ExpectRect(at_anchor->rect, grown);
      break;
    }
    case Outcome::kBlockedPastEdge:
      // A range past the edge is not reported, as for a formula's spill.
      EXPECT_FALSE(spill.has_value());
      EXPECT_TRUE(IsSpillError(anchor));
      EXPECT_FALSE(CellSpill(book.wb.handle, anchor_row, anchor_col).has_value());
      break;
    case Outcome::kRefused:
    case Outcome::kMoved:
      ADD_FAILURE() << "an edit is never refused or moved";
      break;
  }
}

void RunMove(const Case& c) {
  Book book;
  book.Build();
  std::size_t pivot = 0;
  book.AddMainPivot(c.kind, &pivot);
  const fm_cell_rect_t current = Bounds(book.wb.handle, pivot);
  // Past the edge, the destination runs off the last row.
  const std::uint32_t row = c.occupant == Occupant::kPastEdge ? kMaxRows - 2U : kMoveRow;
  const std::uint32_t shift = row - kAnchorRow;
  const fm_cell_rect_t dest{current.first_row + shift, current.first_col, current.last_row + shift, current.last_col};
  if (c.occupant != Occupant::kPastEdge) {
    book.Place(c.occupant, dest, current);
  }
  const std::uint32_t corner_row = c.occupant == Occupant::kPastEdge ? current.last_row : dest.last_row;
  const Snapshot before = Take(book.wb.handle, corner_row, dest.last_col);

  const fm_status_t rc = fm_sheet_set_pivot_anchor(book.wb.handle, kSheet, pivot, row, kAnchorCol);
  switch (c.expected) {
    case Outcome::kMoved: {
      ASSERT_EQ(rc, 0) << fm_get_last_error_message();
      ExpectRect(Bounds(book.wb.handle, pivot), dest);
      EXPECT_FALSE(PivotSpill(book.wb.handle, pivot).has_value());
      EXPECT_FALSE(IsSpillError(ValueAt(book.wb.handle, row, kAnchorCol)));
      // The old range is empty.
      EXPECT_EQ(ValueAt(book.wb.handle, current.last_row, current.first_col).kind, FM_VALUE_KIND_BLANK);
      // What the paste overwrote is gone.
      const Snapshot after = Take(book.wb.handle, dest.last_row, dest.last_col);
      switch (c.occupant) {
        case Occupant::kValue:
        case Occupant::kEmptyStringFormula:
          EXPECT_NE(after.corner_kind, FM_VALUE_KIND_TEXT);
          break;
        case Occupant::kTable:
          EXPECT_EQ(after.tables, 0U);
          break;
        case Occupant::kEmptyMerge:
          EXPECT_EQ(after.merges, 0U);
          break;
        case Occupant::kLegacyArray:
          EXPECT_FALSE(after.array);
          break;
        case Occupant::kDataTableResult:
          EXPECT_FALSE(after.data_table);
          break;
        default:
          break;
      }
      break;
    }
    case Outcome::kRefused:
      EXPECT_EQ(rc, kPreconditionFailed);
      ExpectRect(Bounds(book.wb.handle, pivot), current);
      EXPECT_TRUE(Take(book.wb.handle, corner_row, dest.last_col) == before);
      break;
    default:
      ADD_FAILURE() << "a move renders or is refused";
      break;
  }
}

void RunCreate(const Case& c) {
  Book book;
  book.Build();
  const std::uint32_t row = c.occupant == Occupant::kPastEdge ? kMaxRows - 5U : kCreateRow;
  const fm_cell_rect_t needed{row, kAnchorCol, row + kPlaceholderRows - 1U, kAnchorCol + kPlaceholderCols - 1U};
  if (c.occupant != Occupant::kPastEdge) {
    book.Place(c.occupant, needed, std::nullopt);
  }
  const std::uint32_t corner_row = c.occupant == Occupant::kPastEdge ? kMaxRows - 1U : needed.last_row;
  const Snapshot before = Take(book.wb.handle, corner_row, needed.last_col);
  const std::size_t pivots = PivotCount(book.wb.handle);

  std::size_t created = 99;
  const fm_status_t rc = AddPivot(book.wb.handle, kSheet, "Created", book.cache, row, kAnchorCol, &created);
  switch (c.expected) {
    case Outcome::kRendered:
      ASSERT_EQ(rc, 0) << fm_get_last_error_message();
      EXPECT_EQ(PivotCount(book.wb.handle), pivots + 1U);
      ExpectRect(Bounds(book.wb.handle, created), needed);
      EXPECT_FALSE(PivotSpill(book.wb.handle, created).has_value());
      break;
    case Outcome::kRefused:
      EXPECT_EQ(rc, kPreconditionFailed);
      EXPECT_EQ(created, 99U);
      EXPECT_EQ(PivotCount(book.wb.handle), pivots);
      EXPECT_TRUE(Take(book.wb.handle, corner_row, needed.last_col) == before);
      break;
    default:
      ADD_FAILURE() << "a creation renders or is refused";
      break;
  }
}

}  // namespace

TEST(FormulonCApiPivotBlocked, PlacementMatrix) {
  std::size_t index = 0;
  for (const Case& c : kCases) {
    SCOPED_TRACE(std::to_string(index++) + ": " + Name(c.trigger) + "," + Name(c.occupant) + "," + Name(c.kind));
    switch (c.trigger) {
      case Trigger::kMove:
        RunMove(c);
        break;
      case Trigger::kCreate:
        RunCreate(c);
        break;
      default:
        RunEdit(c);
        break;
    }
  }
}

TEST(FormulonCApiPivotBlocked, GetSpillReportsTheBlockedAnchor) {
  // A refresh growing the report onto a value blocks it.
  fm_cell_rect_t grown{};
  {
    Book control;
    control.Build();
    std::size_t pivot = 0;
    control.AddMainPivot(Kind::kNormal, &pivot);
    ASSERT_EQ(control.Edit(Trigger::kRefresh, Kind::kNormal, pivot), 0) << fm_get_last_error_message();
    grown = Bounds(control.wb.handle, pivot);
  }
  Book book;
  book.Build();
  std::size_t pivot = 0;
  book.AddMainPivot(Kind::kNormal, &pivot);
  fm_workbook_t* wb = book.wb.handle;
  // A drawn report reports nothing.
  EXPECT_FALSE(PivotSpill(wb, pivot).has_value());
  EXPECT_FALSE(CellSpill(wb, kAnchorRow, kAnchorCol).has_value());
  book.Place(Occupant::kValue, grown, std::nullopt);
  ASSERT_EQ(book.Edit(Trigger::kRefresh, Kind::kNormal, pivot), 0) << fm_get_last_error_message();

  // The anchor reports the needed range; a cleared body cell and the occupant do not.
  const std::optional<fm_spill_info_t> at_anchor = CellSpill(wb, kAnchorRow, kAnchorCol);
  ASSERT_TRUE(at_anchor.has_value());
  EXPECT_EQ(at_anchor->blocked, 1);
  ExpectRect(at_anchor->rect, grown);
  EXPECT_FALSE(CellSpill(wb, kAnchorRow + 1U, kAnchorCol).has_value());
  EXPECT_FALSE(CellSpill(wb, grown.last_row, grown.last_col).has_value());
  const std::optional<fm_spill_info_t> own = PivotSpill(wb, pivot);
  ASSERT_TRUE(own.has_value());
  EXPECT_EQ(own->blocked, 1);
  ExpectRect(own->rect, grown);

  // Argument checks.
  std::int32_t present = -1;
  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
  EXPECT_EQ(fm_sheet_get_pivot_spill(nullptr, kSheet, pivot, &present, &info), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_pivot_spill(wb, kSheet, pivot, nullptr, &info), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_pivot_spill(wb, kSheet, pivot, &present, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  for (const std::size_t bad : {std::size_t{2}, static_cast<std::size_t>(0x100000000ULL)}) {
    EXPECT_EQ(fm_sheet_get_pivot_spill(wb, bad, pivot, &present, &info), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE) << bad;
  }
  for (const std::size_t bad : {std::size_t{1}, static_cast<std::size_t>(0x100000000ULL)}) {
    EXPECT_EQ(fm_sheet_get_pivot_spill(wb, kSheet, bad, &present, &info), FM_STATUS_INDEX_OUT_OF_RANGE) << bad;
  }
  info.struct_size = 4U;
  EXPECT_EQ(fm_sheet_get_pivot_spill(wb, kSheet, pivot, &present, &info), kInvalidArgument);

  // Clearing the occupant renders the report on the next read; nothing is reported then.
  ASSERT_EQ(fm_sheet_set_blank(wb, kSheet, grown.last_row, grown.last_col), 0);
  EXPECT_FALSE(CellSpill(wb, kAnchorRow, kAnchorCol).has_value());
  EXPECT_FALSE(PivotSpill(wb, pivot).has_value());
  EXPECT_FALSE(IsSpillError(ValueAt(wb, kAnchorRow, kAnchorCol)));
  ExpectRect(Bounds(wb, pivot), grown);
}

TEST(FormulonCApiPivotBlocked, GetSpillDoesNotReportARangePastTheEdge) {
  Book book;
  book.Build();
  std::size_t pivot = 0;
  book.AddMainPivot(Kind::kNormal, &pivot);
  const fm_cell_rect_t current = Bounds(book.wb.handle, pivot);
  const std::uint32_t anchor_row = kAnchorRow + (kMaxRows - 1U) - current.last_row;
  ASSERT_EQ(fm_sheet_set_pivot_anchor(book.wb.handle, kSheet, pivot, anchor_row, kAnchorCol), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(book.Edit(Trigger::kGrouping, Kind::kNormal, pivot), 0) << fm_get_last_error_message();
  EXPECT_TRUE(IsSpillError(ValueAt(book.wb.handle, anchor_row, kAnchorCol)));
  EXPECT_FALSE(CellSpill(book.wb.handle, anchor_row, kAnchorCol).has_value());
  EXPECT_FALSE(PivotSpill(book.wb.handle, pivot).has_value());
}
