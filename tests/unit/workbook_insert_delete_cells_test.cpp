// Insert Cells / Delete Cells (`Workbook::insert_cells` / `delete_cells`) and
// the table totals-row toggle built on them. Expected values are the Excel
// measurements of the band edits on B3:C4, the totals-row band shift and the
// PivotTable cases, or follow from the band rules where Excel was not measured
// (the `Rule*` cases, one row each of a pairwise direction x position x object
// model).

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_types.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "merge_range.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "sheet_layout.h"
#include "sheet_layout_edit.h"
#include "sheet_metadata.h"
#include "styles.h"
#include "table.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// ---------------------------------------------------------------------------
// Addressing and editing helpers
// ---------------------------------------------------------------------------

CellAddress Addr(std::string_view a1) {
  CellAddress out{0U, 0U};
  EXPECT_TRUE(a1::parse_a1_ref(a1, &out.row, &out.col)) << a1;
  return out;
}

// "B3" or "B3:C4" as a rectangle.
MergeRange Rect(std::string_view a1) {
  const std::size_t colon = a1.find(':');
  const CellAddress first = Addr(a1.substr(0, colon));
  const CellAddress last = colon == std::string_view::npos ? first : Addr(a1.substr(colon + 1U));
  return MergeRange{first.row, first.col, last.row, last.col};
}

// A rectangle as "B3" (one cell) or "B3:C4".
std::string A1(const MergeRange& r) {
  std::string out = a1::encode_a1(r.first_row, r.first_col);
  if (r.first_row != r.last_row || r.first_col != r.last_col) {
    out += ":" + a1::encode_a1(r.last_row, r.last_col);
  }
  return out;
}

// Sorted, space-separated rectangles, so a split compares without depending
// on the order the parts are stored in.
std::string Joined(std::vector<std::string> parts) {
  std::sort(parts.begin(), parts.end());
  std::string out;
  for (const std::string& part : parts) {
    out += out.empty() ? part : " " + part;
  }
  return out;
}

bool IsInsert(CellShift shift) {
  return shift == CellShift::kDown || shift == CellShift::kRight;
}

Expected<void, Error> Edit(Workbook& wb, const MergeRange& rect, CellShift shift, std::size_t sheet = 0U) {
  return IsInsert(shift) ? wb.insert_cells(sheet, rect, shift) : wb.delete_cells(sheet, rect, shift);
}

Expected<void, Error> Edit(Workbook& wb, std::string_view range, CellShift shift, std::size_t sheet = 0U) {
  return Edit(wb, Rect(range), shift, sheet);
}

void ExpectOk(const Expected<void, Error>& result) {
  EXPECT_TRUE(result.has_value()) << (result.has_value() ? "" : result.error().message);
}

void ExpectCode(const Expected<void, Error>& result, FormulonErrorCode code) {
  ASSERT_FALSE(result.has_value());
  EXPECT_EQ(result.error().code, code);
}

void ExpectRefused(const Expected<void, Error>& result) {
  ExpectCode(result, FormulonErrorCode::kPreconditionFailed);
}

const char* ShiftName(CellShift shift) {
  switch (shift) {
    case CellShift::kDown:
      return "down";
    case CellShift::kRight:
      return "right";
    case CellShift::kUp:
      return "up";
    case CellShift::kLeft:
      return "left";
  }
  return "?";
}

constexpr CellShift kAllShifts[] = {CellShift::kDown, CellShift::kRight, CellShift::kUp, CellShift::kLeft};

void Recalc(Workbook& wb) {
  ASSERT_TRUE(wb.recalc(eval::default_registry()).has_value());
}

void SetNumber(Workbook& wb, std::string_view a1, double value, std::size_t sheet = 0U) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_value(sheet, at.row, at.col, Value::number(value)).has_value()) << a1;
}

void SetText(Workbook& wb, std::string_view a1, std::string_view text, std::size_t sheet = 0U) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_text(sheet, at.row, at.col, std::string(text)).has_value()) << a1;
}

void SetFormula(Workbook& wb, std::string_view a1, std::string_view formula, std::size_t sheet = 0U) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_formula(sheet, at.row, at.col, std::string(formula)).has_value()) << a1;
}

const Cell* CellOf(const Workbook& wb, std::string_view a1, std::size_t sheet = 0U) {
  const CellAddress at = Addr(a1);
  return wb.sheet(sheet).cell_at(at.row, at.col);
}

// The stored formula text of a cell, empty when it holds none.
std::string FormulaAt(const Workbook& wb, std::string_view a1, std::size_t sheet = 0U) {
  const Cell* cell = CellOf(wb, a1, sheet);
  return cell == nullptr ? std::string() : cell->formula_text;
}

Value At(const Workbook& wb, std::string_view a1, std::size_t sheet = 0U) {
  const CellAddress at = Addr(a1);
  return wb.sheet(sheet).resolve_cell_value(at.row, at.col);
}

std::string TextAt(const Workbook& wb, std::string_view a1, std::size_t sheet = 0U) {
  const Value value = At(wb, a1, sheet);
  return value.is_text() ? std::string(value.as_text()) : std::string();
}

std::string NameFormula(const Workbook& wb, std::string_view name) {
  for (const DefinedName& dn : wb.defined_names()) {
    if (dn.name == name) {
      return dn.formula;
    }
  }
  ADD_FAILURE() << "no defined name " << name;
  return {};
}

void SetName(Workbook& wb, std::string_view name, std::string_view formula) {
  ASSERT_TRUE(wb.set_defined_name(std::string(name), std::string(formula)).has_value()) << name;
}

// ---------------------------------------------------------------------------
// Sheet object helpers
// ---------------------------------------------------------------------------

void AddCf(Workbook& wb, const std::vector<std::string_view>& ranges, std::size_t sheet = 0U) {
  cf::ConditionalFormat format;
  for (const std::string_view range : ranges) {
    const MergeRange r = Rect(range);
    cf::CFCellRange cell_range;
    cell_range.first = CellAddress{r.first_row, r.first_col};
    cell_range.last = CellAddress{r.last_row, r.last_col};
    format.sqref.push_back(cell_range);
  }
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = static_cast<std::int32_t>(wb.sheet(sheet).conditional_formats().size() + 1U);
  rule.formula1 = "TRUE";
  format.rules.push_back(rule);
  wb.sheet(sheet).mutable_conditional_formats().push_back(std::move(format));
}

std::string CfSqref(const Workbook& wb, std::size_t index = 0U, std::size_t sheet = 0U) {
  const auto& formats = wb.sheet(sheet).conditional_formats();
  if (index >= formats.size()) {
    return "<none>";
  }
  std::vector<std::string> parts;
  for (const cf::CFCellRange& r : formats[index].sqref) {
    parts.push_back(A1(MergeRange{r.first.row, r.first.col, r.last.row, r.last.col}));
  }
  return Joined(std::move(parts));
}

void AddDv(Workbook& wb, const std::vector<std::string_view>& ranges, std::size_t sheet = 0U) {
  DataValidation dv;
  dv.type = 3;  // list
  dv.formula1 = "\"a,b\"";
  for (const std::string_view range : ranges) {
    dv.ranges.push_back(Rect(range));
  }
  ASSERT_TRUE(wb.add_validation(sheet, std::move(dv)).has_value());
}

std::string DvRanges(const Workbook& wb, std::size_t index = 0U, std::size_t sheet = 0U) {
  const auto& validations = wb.sheet(sheet).validations();
  if (index >= validations.size()) {
    return "<none>";
  }
  std::vector<std::string> parts;
  for (const MergeRange& r : validations[index].ranges) {
    parts.push_back(A1(r));
  }
  return Joined(std::move(parts));
}

void AddMerge(Workbook& wb, std::string_view range) {
  ASSERT_TRUE(wb.add_merge(0U, Rect(range)).has_value()) << range;
}

std::string Merges(const Workbook& wb) {
  std::vector<std::string> parts;
  for (const MergeRange& r : wb.sheet(0).merges()) {
    parts.push_back(A1(r));
  }
  return Joined(std::move(parts));
}

void AddLink(Workbook& wb, std::string_view range, std::string target) {
  const MergeRange r = Rect(range);
  Hyperlink link;
  link.row = r.first_row;
  link.col = r.first_col;
  link.last_row = r.last_row;
  link.last_col = r.last_col;
  link.target = std::move(target);
  wb.sheet(0).mutable_hyperlinks().push_back(std::move(link));
}

// The range of the hyperlink to `target`, empty when it is gone.
std::string LinkAt(const Workbook& wb, std::string_view target) {
  for (const Hyperlink& link : wb.sheet(0).hyperlinks()) {
    if (link.target == target) {
      return A1(MergeRange{link.row, link.col, link.last_row, link.last_col});
    }
  }
  return {};
}

void AddNote(Workbook& wb, std::string_view cell, std::string text) {
  const CellAddress at = Addr(cell);
  wb.sheet(0).mutable_comments().emplace_back(at.row, at.col, "Author", std::move(text));
}

// The cell of the note reading `text`, empty when it is gone.
std::string NoteAt(const Workbook& wb, std::string_view text) {
  for (const CellComment& note : wb.sheet(0).comments()) {
    if (note.text == text) {
      return a1::encode_a1(note.row, note.col);
    }
  }
  return {};
}

TableMetadata MakeTable(std::string name, std::string_view ref, std::size_t sheet = 0U) {
  const MergeRange r = Rect(ref);
  TableMetadata table;
  table.sheet_index = sheet;
  table.name = name;
  table.display_name = std::move(name);
  table.ref = std::string(ref);
  for (std::uint32_t c = r.first_col; c <= r.last_col; ++c) {
    TableColumn column;
    column.name = "C" + std::to_string(c - r.first_col + 1U);
    table.columns.push_back(std::move(column));
  }
  return table;
}

void AddTable(Workbook& wb, std::string name, std::string_view ref) {
  ASSERT_TRUE(wb.add_table(MakeTable(std::move(name), ref)).has_value()) << ref;
}

std::string TableRef(const Workbook& wb, std::string_view name) {
  for (const TableMetadata& table : wb.tables()) {
    if (table.name == name) {
      return table.ref;
    }
  }
  return {};
}

void AddArray(Workbook& wb, std::string_view range) {
  const MergeRange r = Rect(range);
  ASSERT_TRUE(wb.set_array_formula(0U, r.first_row, r.first_col, r.last_row - r.first_row + 1U,
                                   r.last_col - r.first_col + 1U, "=1")
                  .has_value())
      << range;
}

// The extent of the legacy array block whose anchor is `anchor`, empty when
// no block is anchored there.
std::string ArrayAt(const Workbook& wb, std::string_view anchor) {
  const CellAddress at = Addr(anchor);
  const std::optional<Workbook::ArrayFormulaInfo> info = wb.array_formula_at(0U, at.row, at.col);
  if (!info.has_value() || info->anchor_row != at.row || info->anchor_col != at.col) {
    return {};
  }
  return A1(MergeRange{info->anchor_row, info->anchor_col, info->anchor_row + info->rows - 1U,
                       info->anchor_col + info->cols - 1U});
}

// A PivotTable occupying `range`, with no cache: enough for the placement rules.
void AddBarePivot(Workbook& wb, std::string_view range) {
  const MergeRange r = Rect(range);
  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("PivotTable1");
  table->set_anchor(r.first_row, r.first_col, r.last_row - r.first_row + 1U, r.last_col - r.first_col + 1U);
  wb.sheet(0).mutable_pivot_tables().push_back(std::move(table));
}

std::string PivotAt(const Workbook& wb, std::size_t index = 0U) {
  const auto& pivots = wb.sheet(0).pivot_tables();
  if (index >= pivots.size()) {
    return "<none>";
  }
  const pivot::PivotTable& p = *pivots[index];
  return A1(MergeRange{p.anchor_row(), p.anchor_col(), p.anchor_row() + p.span_rows() - 1U,
                       p.anchor_col() + p.span_cols() - 1U});
}

// ---------------------------------------------------------------------------
// Argument and boundary checks
// ---------------------------------------------------------------------------

TEST(InsertDeleteCells, InsertRefusesAnUpOrLeftShift) {
  Workbook wb = Workbook::create();
  SetNumber(wb, "B3", 1.0);
  ExpectCode(wb.insert_cells(0U, Rect("B3:C4"), CellShift::kUp), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.insert_cells(0U, Rect("B3:C4"), CellShift::kLeft), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(At(wb, "B3"), Value::number(1.0));
}

TEST(InsertDeleteCells, DeleteRefusesADownOrRightShift) {
  Workbook wb = Workbook::create();
  SetNumber(wb, "B3", 1.0);
  ExpectCode(wb.delete_cells(0U, Rect("B3:C4"), CellShift::kDown), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.delete_cells(0U, Rect("B3:C4"), CellShift::kRight), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(At(wb, "B3"), Value::number(1.0));
}

TEST(InsertDeleteCells, AnOutOfRangeSheetIsAnInvalidArgument) {
  Workbook wb = Workbook::create();
  ExpectCode(wb.insert_cells(1U, Rect("B3:C4"), CellShift::kDown), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.delete_cells(1U, Rect("B3:C4"), CellShift::kUp), FormulonErrorCode::kInvalidArgument);
}

TEST(InsertDeleteCells, ARangeOutsideTheGridIsAnInvalidArgument) {
  Workbook wb = Workbook::create();
  const MergeRange below{Sheet::kMaxRows, 1U, Sheet::kMaxRows, 2U};
  const MergeRange right{2U, Sheet::kMaxCols, 3U, Sheet::kMaxCols};
  for (const CellShift shift : kAllShifts) {
    SCOPED_TRACE(ShiftName(shift));
    ExpectCode(Edit(wb, below, shift), FormulonErrorCode::kInvalidArgument);
    ExpectCode(Edit(wb, right, shift), FormulonErrorCode::kInvalidArgument);
  }
}

TEST(InsertDeleteCells, AWholeColumnRangeCannotShiftVertically) {
  Workbook wb = Workbook::create();
  const MergeRange columns{0U, 1U, Sheet::kMaxRows - 1U, 2U};  // B:C
  ExpectCode(Edit(wb, columns, CellShift::kDown), FormulonErrorCode::kInvalidArgument);
  ExpectCode(Edit(wb, columns, CellShift::kUp), FormulonErrorCode::kInvalidArgument);
}

TEST(InsertDeleteCells, AWholeRowRangeCannotShiftHorizontally) {
  Workbook wb = Workbook::create();
  const MergeRange rows{2U, 0U, 3U, Sheet::kMaxCols - 1U};  // 3:4
  ExpectCode(Edit(wb, rows, CellShift::kRight), FormulonErrorCode::kInvalidArgument);
  ExpectCode(Edit(wb, rows, CellShift::kLeft), FormulonErrorCode::kInvalidArgument);
}

// Builds the same small sheet twice, for comparing a band edit covering the
// whole perpendicular extent with the matching whole row/column edit.
void FillForWholeLineComparison(Workbook& wb) {
  SetNumber(wb, "B1", 1.0);
  SetNumber(wb, "D3", 2.0);
  SetNumber(wb, "A6", 3.0);
  SetFormula(wb, "J20", "=SUM(A6:C6)");
  SetFormula(wb, "J21", "=SUM(D3:E5)");
  SetFormula(wb, "J22", "=SUM(B:B)");
  SetFormula(wb, "J23", "=SUM(3:3)");
}

void ExpectSameSheet(const Workbook& lhs, const Workbook& rhs) {
  for (const std::string_view a1 : {"B1", "D1", "D3", "F3", "A6", "A8", "A4", "B3"}) {
    EXPECT_EQ(lhs.sheet(0).resolve_cell_value(Addr(a1).row, Addr(a1).col),
              rhs.sheet(0).resolve_cell_value(Addr(a1).row, Addr(a1).col))
        << a1;
  }
  for (std::uint32_t row = 0U; row < 40U; ++row) {
    for (std::uint32_t col = 0U; col < 12U; ++col) {
      const Cell* l = lhs.sheet(0).cell_at(row, col);
      const Cell* r = rhs.sheet(0).cell_at(row, col);
      EXPECT_EQ(l == nullptr ? std::string() : l->formula_text, r == nullptr ? std::string() : r->formula_text)
          << a1::encode_a1(row, col);
    }
  }
}

TEST(InsertDeleteCells, AWholeColumnRangeShiftedRightIsAColumnInsert) {
  Workbook band = Workbook::create();
  Workbook whole = Workbook::create();
  FillForWholeLineComparison(band);
  FillForWholeLineComparison(whole);
  ExpectOk(Edit(band, MergeRange{0U, 1U, Sheet::kMaxRows - 1U, 2U}, CellShift::kRight));
  ExpectOk(whole.insert_cols(0U, 1U, 2U));
  ExpectSameSheet(band, whole);
  EXPECT_EQ(FormulaAt(band, "L22"), "=SUM(D:D)");
}

TEST(InsertDeleteCells, AWholeRowRangeShiftedDownIsARowInsert) {
  Workbook band = Workbook::create();
  Workbook whole = Workbook::create();
  FillForWholeLineComparison(band);
  FillForWholeLineComparison(whole);
  ExpectOk(Edit(band, MergeRange{2U, 0U, 3U, Sheet::kMaxCols - 1U}, CellShift::kDown));
  ExpectOk(whole.insert_rows(0U, 2U, 2U));
  ExpectSameSheet(band, whole);
  EXPECT_EQ(FormulaAt(band, "J25"), "=SUM(5:5)");
}

TEST(InsertDeleteCells, ARectangleCoveringEveryColumnIsARowEdit) {
  for (const CellShift shift : {CellShift::kDown, CellShift::kUp}) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook band = Workbook::create();
    Workbook whole = Workbook::create();
    FillForWholeLineComparison(band);
    FillForWholeLineComparison(whole);
    ExpectOk(Edit(band, "A3:XFD4", shift));
    ExpectOk(shift == CellShift::kDown ? whole.insert_rows(0U, 2U, 2U) : whole.delete_rows(0U, 2U, 2U));
    ExpectSameSheet(band, whole);
  }
}

TEST(InsertDeleteCells, ARectangleCoveringEveryRowIsAColumnEdit) {
  for (const CellShift shift : {CellShift::kRight, CellShift::kLeft}) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook band = Workbook::create();
    Workbook whole = Workbook::create();
    FillForWholeLineComparison(band);
    FillForWholeLineComparison(whole);
    ExpectOk(Edit(band, "B1:C1048576", shift));
    ExpectOk(shift == CellShift::kRight ? whole.insert_cols(0U, 1U, 2U) : whole.delete_cols(0U, 1U, 2U));
    ExpectSameSheet(band, whole);
  }
}

TEST(InsertDeleteCells, ASingleCellInsertShiftsOnlyItsColumn) {
  Workbook wb = Workbook::create();
  SetNumber(wb, "B3", 1.0);
  SetNumber(wb, "C3", 2.0);
  ExpectOk(Edit(wb, "B3", CellShift::kDown));
  EXPECT_TRUE(At(wb, "B3").is_blank());
  EXPECT_EQ(At(wb, "B4"), Value::number(1.0));
  EXPECT_EQ(At(wb, "C3"), Value::number(2.0));
}

TEST(InsertDeleteCells, InsertRefusesPushingAValueOffTheGrid) {
  {
    Workbook wb = Workbook::create();
    SetNumber(wb, "B1048576", 1.0);
    SetNumber(wb, "B3", 2.0);
    ExpectRefused(Edit(wb, "B3:C4", CellShift::kDown));
    EXPECT_EQ(At(wb, "B1048576"), Value::number(1.0));
    EXPECT_EQ(At(wb, "B3"), Value::number(2.0));
  }
  {
    Workbook wb = Workbook::create();
    SetFormula(wb, "C1048575", "=1");
    ExpectRefused(Edit(wb, "B3:C4", CellShift::kDown));
    EXPECT_EQ(FormulaAt(wb, "C1048575"), "=1");
  }
  {
    Workbook wb = Workbook::create();
    SetNumber(wb, "XFD3", 1.0);
    ExpectRefused(Edit(wb, "B3:C4", CellShift::kRight));
    EXPECT_EQ(At(wb, "XFD3"), Value::number(1.0));
  }
}

TEST(InsertDeleteCells, DeleteNearTheGridEndPullsTheLastCellsBack) {
  {
    Workbook wb = Workbook::create();
    SetNumber(wb, "B1048576", 1.0);
    ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
    EXPECT_EQ(At(wb, "B1048574"), Value::number(1.0));
    EXPECT_TRUE(At(wb, "B1048576").is_blank());
  }
  {
    Workbook wb = Workbook::create();
    SetNumber(wb, "XFD3", 1.0);
    ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
    EXPECT_EQ(At(wb, "XFB3"), Value::number(1.0));
    EXPECT_TRUE(At(wb, "XFD3").is_blank());
  }
}

// Gives the strip cell (C1048576 on the last row, XFD3 on the last column) a
// cell format built by `edit`.
template <typename Edit>
void StyleCell(Workbook& wb, bool last_row, Edit&& edit) {
  CellXf xf;
  edit(xf);
  const std::uint32_t index = intern_cell_xf(wb.mutable_styles(), xf);
  wb.sheet(0).set_cell_xf_index(last_row ? Sheet::kMaxRows - 1U : 2U, last_row ? 2U : Sheet::kMaxCols - 1U, index);
}

// A note, a fill or border, or a merge pushed off the grid blocks an insert
// as a value does, on the last row and on the last column.
TEST(InsertDeleteCells, InsertRefusesPushingANoteFormatOrMergeOffTheGrid) {
  struct Blocker {
    const char* name;
    void (*place)(Workbook& wb, bool last_row);
  };
  const Blocker blockers[] = {
      {"note", [](Workbook& wb, bool last_row) { AddNote(wb, last_row ? "C1048576" : "XFD3", "pushed"); }},
      {"fill",
       [](Workbook& wb, bool last_row) {
         FillRecord fill;
         fill.pattern = 1U;
         fill.fg_argb = 0xFFFFFF00U;
         StyleCell(wb, last_row, [&wb, &fill](CellXf& xf) {
           wb.mutable_styles().fills.push_back(fill);
           xf.fill_index = static_cast<std::uint32_t>(wb.styles().fills.size() - 1U);
         });
       }},
      {"border",
       [](Workbook& wb, bool last_row) {
         BorderRecord border;
         border.bottom.style = 1U;
         StyleCell(wb, last_row, [&wb, &border](CellXf& xf) {
           wb.mutable_styles().borders.push_back(border);
           xf.border_index = static_cast<std::uint32_t>(wb.styles().borders.size() - 1U);
         });
       }},
      {"merge",
       [](Workbook& wb, bool last_row) {
         ExpectOk(wb.add_merge(0, last_row ? Rect("B1048575:C1048576") : Rect("XFC3:XFD4")));
       }},
  };
  for (const Blocker& blocker : blockers) {
    for (const bool last_row : {true, false}) {
      SCOPED_TRACE(std::string(blocker.name) + (last_row ? " last row" : " last column"));
      Workbook wb = Workbook::create();
      blocker.place(wb, last_row);
      SetNumber(wb, "B3", 2.0);
      ExpectRefused(Edit(wb, "B3:C4", last_row ? CellShift::kDown : CellShift::kRight));
      EXPECT_EQ(At(wb, "B3"), Value::number(2.0));
      ExpectRefused(last_row ? wb.insert_rows(0U, 2U, 1U) : wb.insert_cols(0U, 1U, 1U));
      EXPECT_EQ(At(wb, "B3"), Value::number(2.0));
    }
  }
  Workbook wb = Workbook::create();
  AddNote(wb, "C1048576", "kept");
  ExpectRefused(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(NoteAt(wb, "kept"), "C1048576");
}

// A cell format that draws nothing, a bold font or a number format, is
// dropped instead of blocking.
TEST(InsertDeleteCells, InsertDropsAFontOrNumberFormatPushedOffTheGrid) {
  for (const bool last_row : {true, false}) {
    SCOPED_TRACE(last_row ? "last row" : "last column");
    Workbook wb = Workbook::create();
    StyleCell(wb, last_row, [&wb](CellXf& xf) {
      FontRecord bold = wb.styles().fonts.front();
      bold.bold = true;
      wb.mutable_styles().fonts.push_back(bold);
      xf.font_index = static_cast<std::uint32_t>(wb.styles().fonts.size() - 1U);
      xf.num_fmt_id = 2U;
    });
    ExpectOk(Edit(wb, "B3:C4", last_row ? CellShift::kDown : CellShift::kRight));
  }
}

TEST(InsertDeleteCells, InsertDropsAHyperlinkPushedOffTheGrid) {
  Workbook wb = Workbook::create();
  AddLink(wb, "B1048576", "https://pushed.example");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(LinkAt(wb, "https://pushed.example"), "");
}

// A cell format with a fill, for whole-row and whole-column styles.
std::uint32_t FillXf(Workbook& wb) {
  FillRecord fill;
  fill.pattern = 1U;
  fill.fg_argb = 0xFFFFFF00U;
  wb.mutable_styles().fills.push_back(fill);
  CellXf xf;
  xf.fill_index = static_cast<std::uint32_t>(wb.styles().fills.size() - 1U);
  return intern_cell_xf(wb.mutable_styles(), xf);
}

void StyleColumn(Workbook& wb, std::uint32_t col, std::uint32_t xf_index) {
  ColumnLayout column;
  column.first = col;
  column.last = col;
  column.has_style = true;
  column.style_xf = xf_index;
  wb.sheet(0).mutable_layout().columns.push_back(column);
}

void StyleRow(Workbook& wb, std::uint32_t row, std::uint32_t xf_index) {
  RowLayout& entry = upsert_row_override(wb.sheet(0).mutable_layout(), row, /*initial_hidden=*/false);
  entry.has_style = true;
  entry.style_xf = xf_index;
}

// The cell record at `a1`, which must be an explicit default-style one.
void ExpectDefaultStyleRecord(Workbook& wb, std::string_view a1) {
  SCOPED_TRACE(std::string(a1));
  const Cell* cell = wb.sheet(0).cell_at(Rect(a1).first_row, Rect(a1).first_col);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->has_explicit_xf);
  EXPECT_EQ(cell->xf_index, 0U);
  EXPECT_TRUE(cell->cached_value.is_blank());
}

// Shifting a band's cells toward a styled last column or row stays inside the
// grid: the style stays, and the blank shifted into the edge cell becomes an
// explicit default-style record that overrides it there.
TEST(InsertDeleteCells, ShiftTowardAStyledEdgeWritesADefaultCellThere) {
  {
    Workbook wb = Workbook::create();
    StyleColumn(wb, Sheet::kMaxCols - 1U, FillXf(wb));
    ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
    ExpectDefaultStyleRecord(wb, "XFD3");
    ExpectDefaultStyleRecord(wb, "XFD4");
    EXPECT_EQ(wb.sheet(0).cell_at(1U, Sheet::kMaxCols - 1U), nullptr) << "XFD2 lies outside the band";
    ASSERT_EQ(wb.sheet(0).layout().columns.size(), 1U);
    EXPECT_EQ(wb.sheet(0).layout().columns.front().first, Sheet::kMaxCols - 1U);
  }
  {
    Workbook wb = Workbook::create();
    StyleRow(wb, Sheet::kMaxRows - 1U, FillXf(wb));
    ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
    ExpectDefaultStyleRecord(wb, "B1048576");
    ExpectDefaultStyleRecord(wb, "C1048576");
    EXPECT_EQ(wb.sheet(0).cell_at(Sheet::kMaxRows - 1U, 0U), nullptr) << "A1048576 lies outside the band";
    ASSERT_NE(find_row_override(wb.sheet(0).layout(), Sheet::kMaxRows - 1U), nullptr);
  }
  {
    // A style on the other axis covers the source and the edge alike, so nothing is written.
    Workbook wb = Workbook::create();
    StyleRow(wb, 2U, FillXf(wb));
    StyleColumn(wb, 0U, FillXf(wb));
    ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
    ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
    EXPECT_EQ(wb.sheet(0).cell_at(2U, Sheet::kMaxCols - 1U), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(Sheet::kMaxRows - 1U, 0U), nullptr);
    ASSERT_NE(find_row_override(wb.sheet(0).layout(), 2U), nullptr);
  }
}

// A whole-column or whole-row style pushed off the grid is dropped and does
// not block; a style on the other axis stays where it is.
TEST(InsertDeleteCells, InsertDropsAnAxisStylePushedOffTheGrid) {
  {
    Workbook wb = Workbook::create();
    StyleColumn(wb, Sheet::kMaxCols - 1U, FillXf(wb));
    StyleRow(wb, 2U, FillXf(wb));
    ExpectOk(wb.insert_cols(0U, 0U, 1U));
    EXPECT_TRUE(wb.sheet(0).layout().columns.empty());
    ASSERT_NE(find_row_override(wb.sheet(0).layout(), 2U), nullptr);
    EXPECT_TRUE(find_row_override(wb.sheet(0).layout(), 2U)->has_style);
  }
  {
    Workbook wb = Workbook::create();
    StyleRow(wb, Sheet::kMaxRows - 1U, FillXf(wb));
    StyleColumn(wb, 0U, FillXf(wb));
    ExpectOk(wb.insert_rows(0U, 0U, 1U));
    EXPECT_EQ(find_row_override(wb.sheet(0).layout(), Sheet::kMaxRows - 1U), nullptr);
    ASSERT_EQ(wb.sheet(0).layout().columns.size(), 1U);
    EXPECT_EQ(wb.sheet(0).layout().columns.front().first, 0U);
  }
}

TEST(InsertDeleteCells, InsertRowsRefusesPushingAValueOffTheGrid) {
  Workbook wb = Workbook::create();
  SetNumber(wb, "A1048576", 1.0);
  ExpectRefused(wb.insert_rows(0U, 2U, 2U));
  EXPECT_EQ(At(wb, "A1048576"), Value::number(1.0));
}

TEST(InsertDeleteCells, InsertColsRefusesPushingAValueOffTheGrid) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "XFD1", "=1");
  ExpectRefused(wb.insert_cols(0U, 1U, 1U));
  EXPECT_EQ(FormulaAt(wb, "XFD1"), "=1");
}

// ---------------------------------------------------------------------------
// Measured: B3:C4 in four directions, reference rewrites
// ---------------------------------------------------------------------------

// One formula and its text after each edit; nullptr means unchanged.
struct ReferenceRow {
  const char* formula;
  const char* down;
  const char* right;
  const char* up;
  const char* left;
};

constexpr ReferenceRow kReferenceRows[] = {
    {"=B6", "=B8", nullptr, "=B4", nullptr},
    {"=D3", nullptr, "=F3", nullptr, "=B3"},
    {"=SUM(A6:C6)", nullptr, nullptr, nullptr, nullptr},
    {"=SUM(B6:C6)", "=SUM(B8:C8)", nullptr, "=SUM(B4:C4)", nullptr},
    {"=SUM(B2:C5)", "=SUM(B2:C7)", nullptr, "=SUM(B2:C3)", nullptr},
    {"=SUM(D3:E4)", nullptr, "=SUM(F3:G4)", nullptr, "=SUM(B3:C4)"},
    {"=SUM(D3:E5)", nullptr, nullptr, nullptr, nullptr},
    {"=SUM(B:B)", nullptr, nullptr, nullptr, nullptr},
    {"=SUM(3:3)", nullptr, nullptr, nullptr, nullptr},
    {"=B3", "=B5", "=D3", "=#REF!", "=#REF!"},
    {"=C4", "=C6", "=E4", "=#REF!", "=#REF!"},
    {"=SUM(B3:C4)", "=SUM(B5:C6)", "=SUM(D3:E4)", "=SUM(#REF!)", "=SUM(#REF!)"},
    {"=SUM(B3:B10)", "=SUM(B5:B12)", nullptr, "=SUM(B3:B8)", "=SUM(B5:B10)"},
    {"=SUM(A3:D3)", nullptr, "=SUM(A3:F3)", nullptr, "=SUM(A3:B3)"},
    {"=SUM(B1:C10)", "=SUM(B1:C12)", nullptr, "=SUM(B1:C8)", nullptr},
    {"=SUM(A3:C4)", nullptr, "=SUM(A3:E4)", "=SUM(A3:A4)", "=SUM(A3:A4)"},
    {"=SUM(B3:D4)", nullptr, "=SUM(D3:F4)", "=SUM(D3:D4)", "=SUM(B3:B4)"},
};

const char* ExpectedText(const ReferenceRow& row, CellShift shift) {
  const char* out = shift == CellShift::kDown    ? row.down
                    : shift == CellShift::kRight ? row.right
                    : shift == CellShift::kUp    ? row.up
                                                 : row.left;
  return out == nullptr ? row.formula : out;
}

// Every cell of A1:F12 holds 100 * row + column (1-based); the formulas sit at
// J20 onwards, outside every band.
double Marker(std::uint32_t row1, std::uint32_t col1) {
  return 100.0 * row1 + col1;
}

void FillMeasuredGrid(Workbook& wb) {
  for (std::uint32_t r = 0U; r < 12U; ++r) {
    for (std::uint32_t c = 0U; c < 6U; ++c) {
      ASSERT_TRUE(wb.set_cell_value(0U, r, c, Value::number(Marker(r + 1U, c + 1U))).has_value());
    }
  }
}

void CheckReferenceTable(CellShift shift) {
  Workbook wb = Workbook::create();
  FillMeasuredGrid(wb);
  std::uint32_t host = 19U;
  for (const ReferenceRow& row : kReferenceRows) {
    ASSERT_TRUE(wb.set_cell_formula(0U, host++, 9U, row.formula).has_value()) << row.formula;
  }
  ExpectOk(Edit(wb, "B3:C4", shift));
  host = 19U;
  for (const ReferenceRow& row : kReferenceRows) {
    const Cell* cell = wb.sheet(0).cell_at(host++, 9U);
    ASSERT_NE(cell, nullptr) << row.formula;
    EXPECT_EQ(cell->formula_text, ExpectedText(row, shift)) << row.formula;
  }
}

TEST(InsertDeleteCells, ReferencesFollowAnInsertShiftingDown) {
  CheckReferenceTable(CellShift::kDown);
}

TEST(InsertDeleteCells, ReferencesFollowAnInsertShiftingRight) {
  CheckReferenceTable(CellShift::kRight);
}

TEST(InsertDeleteCells, ReferencesFollowADeleteShiftingUp) {
  CheckReferenceTable(CellShift::kUp);
}

TEST(InsertDeleteCells, ReferencesFollowADeleteShiftingLeft) {
  CheckReferenceTable(CellShift::kLeft);
}

TEST(InsertDeleteCells, CellsMoveOnlyInsideTheBand) {
  for (const CellShift shift : kAllShifts) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook wb = Workbook::create();
    FillMeasuredGrid(wb);
    ExpectOk(Edit(wb, "B3:C4", shift));
    switch (shift) {
      case CellShift::kDown:
        EXPECT_TRUE(At(wb, "B3").is_blank());
        EXPECT_TRUE(At(wb, "C4").is_blank());
        EXPECT_EQ(At(wb, "B5"), Value::number(Marker(3, 2)));
        EXPECT_EQ(At(wb, "C14"), Value::number(Marker(12, 3)));
        EXPECT_EQ(At(wb, "A3"), Value::number(Marker(3, 1)));
        EXPECT_EQ(At(wb, "D3"), Value::number(Marker(3, 4)));
        break;
      case CellShift::kRight:
        EXPECT_TRUE(At(wb, "B3").is_blank());
        EXPECT_TRUE(At(wb, "C4").is_blank());
        EXPECT_EQ(At(wb, "D3"), Value::number(Marker(3, 2)));
        EXPECT_EQ(At(wb, "H4"), Value::number(Marker(4, 6)));
        EXPECT_EQ(At(wb, "B5"), Value::number(Marker(5, 2)));
        EXPECT_EQ(At(wb, "B2"), Value::number(Marker(2, 2)));
        break;
      case CellShift::kUp:
        EXPECT_EQ(At(wb, "B3"), Value::number(Marker(5, 2)));
        EXPECT_EQ(At(wb, "C10"), Value::number(Marker(12, 3)));
        EXPECT_TRUE(At(wb, "C11").is_blank());
        EXPECT_EQ(At(wb, "D3"), Value::number(Marker(3, 4)));
        EXPECT_EQ(At(wb, "A12"), Value::number(Marker(12, 1)));
        break;
      case CellShift::kLeft:
        EXPECT_EQ(At(wb, "B3"), Value::number(Marker(3, 4)));
        EXPECT_EQ(At(wb, "D4"), Value::number(Marker(4, 6)));
        EXPECT_TRUE(At(wb, "E4").is_blank());
        EXPECT_EQ(At(wb, "B5"), Value::number(Marker(5, 2)));
        EXPECT_EQ(At(wb, "F2"), Value::number(Marker(2, 6)));
        break;
    }
  }
}

TEST(InsertDeleteCells, DefinedNamesFollowTheBandRule) {
  struct Case {
    CellShift shift;
    const char* b6;
    const char* d3;
    const char* block;
  };
  const Case cases[] = {
      {CellShift::kDown, "=Sheet1!$B$8", "=Sheet1!$D$3", "=Sheet1!$B$5:$C$6"},
      {CellShift::kRight, "=Sheet1!$B$6", "=Sheet1!$F$3", "=Sheet1!$D$3:$E$4"},
      {CellShift::kUp, "=Sheet1!$B$4", "=Sheet1!$D$3", "=Sheet1!#REF!"},
      {CellShift::kLeft, "=Sheet1!$B$6", "=Sheet1!$B$3", "=Sheet1!#REF!"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(ShiftName(c.shift));
    Workbook wb = Workbook::create();
    SetName(wb, "PointB6", "=Sheet1!$B$6");
    SetName(wb, "PointD3", "=Sheet1!$D$3");
    SetName(wb, "NBlock", "=Sheet1!$B$3:$C$4");
    SetName(wb, "NWide", "=Sheet1!$A$1:$C$10");
    ExpectOk(Edit(wb, "B3:C4", c.shift));
    EXPECT_EQ(NameFormula(wb, "PointB6"), c.b6);
    EXPECT_EQ(NameFormula(wb, "PointD3"), c.d3);
    EXPECT_EQ(NameFormula(wb, "NBlock"), c.block);
    EXPECT_EQ(NameFormula(wb, "NWide"), "=Sheet1!$A$1:$C$10");
  }
}

TEST(InsertDeleteCells, ReferencesFromAnotherSheetFollowTheBandRule) {
  struct Case {
    CellShift shift;
    const char* b6;
    const char* block;
    const char* b3;
    const char* d3;
  };
  const Case cases[] = {
      {CellShift::kDown, "=Sheet1!B8", "=SUM(Sheet1!B5:C6)", "=Sheet1!B5", "=Sheet1!D3"},
      {CellShift::kRight, "=Sheet1!B6", "=SUM(Sheet1!D3:E4)", "=Sheet1!D3", "=Sheet1!F3"},
      {CellShift::kUp, "=Sheet1!B4", "=SUM(Sheet1!#REF!)", "=Sheet1!#REF!", "=Sheet1!D3"},
      {CellShift::kLeft, "=Sheet1!B6", "=SUM(Sheet1!#REF!)", "=Sheet1!#REF!", "=Sheet1!B3"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(ShiftName(c.shift));
    Workbook wb = Workbook::create();
    const std::size_t other = wb.add_sheet("Other");
    SetFormula(wb, "B6", "=Sheet1!B6", other);
    SetFormula(wb, "B7", "=SUM(Sheet1!B3:C4)", other);
    SetFormula(wb, "B3", "=Sheet1!B3", other);
    SetFormula(wb, "D3", "=Sheet1!D3", other);
    ExpectOk(Edit(wb, "B3:C4", c.shift));
    EXPECT_EQ(FormulaAt(wb, "B6", other), c.b6);
    EXPECT_EQ(FormulaAt(wb, "B7", other), c.block);
    EXPECT_EQ(FormulaAt(wb, "B3", other), c.b3);
    EXPECT_EQ(FormulaAt(wb, "D3", other), c.d3);
  }
}

TEST(InsertDeleteCells, AnEditOnAnotherSheetLeavesThisSheetsCellsAlone) {
  Workbook wb = Workbook::create();
  const std::size_t other = wb.add_sheet("Other");
  SetNumber(wb, "B3", 1.0);
  SetFormula(wb, "A1", "=B3");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown, other));
  EXPECT_EQ(At(wb, "B3"), Value::number(1.0));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=B3");
}

// ---------------------------------------------------------------------------
// Measured: B3:C4 in four directions, objects
// ---------------------------------------------------------------------------

TEST(InsertDeleteCells, ConditionalFormatsMoveOrSplitAtTheBand) {
  struct Case {
    CellShift shift;
    const char* in_band_down;
    const char* straddling;
    const char* in_band_right;
  };
  const Case cases[] = {
      {CellShift::kDown, "B8:C10", "A6:A8 B8:C10", "D3:E4"},
      {CellShift::kRight, "B6:C8", "A6:C8", "F3:G4"},
      {CellShift::kUp, "B4:C6", "A6:A8 B4:C6", "D3:E4"},
      {CellShift::kLeft, "B6:C8", "A6:C8", "B3:C4"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(ShiftName(c.shift));
    Workbook wb = Workbook::create();
    AddCf(wb, {"B6:C8"});
    AddCf(wb, {"A6:C8"});
    AddCf(wb, {"D3:E4"});
    ExpectOk(Edit(wb, "B3:C4", c.shift));
    ASSERT_EQ(wb.sheet(0).conditional_formats().size(), 3U);
    EXPECT_EQ(CfSqref(wb, 0U), c.in_band_down);
    EXPECT_EQ(CfSqref(wb, 1U), c.straddling);
    EXPECT_EQ(CfSqref(wb, 2U), c.in_band_right);
  }
}

// The ranges of the first `<xm:sqref>` of the sheet's retained extLst, sorted.
std::string X14Sqref(const Workbook& wb) {
  const std::string& xml = wb.sheet(0).ext_lst_xml();
  const std::size_t open = xml.find("<xm:sqref>");
  if (open == std::string::npos) {
    return "<none>";
  }
  const std::size_t begin = open + std::string_view("<xm:sqref>").size();
  const std::string text = xml.substr(begin, xml.find("</xm:sqref>", begin) - begin);
  std::vector<std::string> parts;
  for (std::size_t at = 0; at <= text.size();) {
    const std::size_t space = std::min(text.find(' ', at), text.size());
    parts.push_back(text.substr(at, space - at));
    at = space + 1U;
  }
  return Joined(std::move(parts));
}

TEST(InsertDeleteCells, AnX14ConditionalFormatSqrefSplitsAndRejoinsWithTheLegacyOne) {
  Workbook wb = Workbook::create();
  AddCf(wb, {"A6:C8"});
  wb.sheet(0).set_ext_lst_xml(
      "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
      "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
      "<x14:conditionalFormattings><x14:conditionalFormatting>"
      "<x14:cfRule type=\"expression\" id=\"{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}\"><xm:f>TRUE</xm:f></x14:cfRule>"
      "<xm:sqref>A6:C8</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>");
  ExpectOk(Edit(wb, "A5:B5", CellShift::kDown));
  EXPECT_EQ(CfSqref(wb), "A7:B9 C6:C8");
  EXPECT_EQ(X14Sqref(wb), CfSqref(wb));
  ExpectOk(Edit(wb, "A5:B5", CellShift::kUp));
  EXPECT_EQ(CfSqref(wb), "A6:C8");
  EXPECT_EQ(X14Sqref(wb), CfSqref(wb));
}

TEST(InsertDeleteCells, ValidationsMoveOrSplitAtTheBand) {
  struct Case {
    CellShift shift;
    const char* in_band_down;
    const char* straddling;
    const char* in_band_right;
  };
  const Case cases[] = {
      {CellShift::kDown, "B8:C10", "A6:A8 B8:C10", "D3:E4"},
      {CellShift::kRight, "B6:C8", "A6:C8", "F3:G4"},
      {CellShift::kUp, "B4:C6", "A6:A8 B4:C6", "D3:E4"},
      {CellShift::kLeft, "B6:C8", "A6:C8", "B3:C4"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(ShiftName(c.shift));
    Workbook wb = Workbook::create();
    AddDv(wb, {"B6:C8"});
    AddDv(wb, {"A6:C8"});
    AddDv(wb, {"D3:E4"});
    ExpectOk(Edit(wb, "B3:C4", c.shift));
    ASSERT_EQ(wb.sheet(0).validations().size(), 3U);
    EXPECT_EQ(DvRanges(wb, 0U), c.in_band_down);
    EXPECT_EQ(DvRanges(wb, 1U), c.straddling);
    EXPECT_EQ(DvRanges(wb, 2U), c.in_band_right);
  }
}

TEST(InsertDeleteCells, HyperlinksAndNotesGoWithTheirCells) {
  struct Case {
    CellShift shift;
    const char* b3;  // empty: deleted with the cell
    const char* b6;
    const char* d3;
  };
  const Case cases[] = {
      {CellShift::kDown, "B5", "B8", "D3"},
      {CellShift::kRight, "D3", "B6", "F3"},
      {CellShift::kUp, "", "B4", "D3"},
      {CellShift::kLeft, "", "B6", "B3"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(ShiftName(c.shift));
    Workbook wb = Workbook::create();
    for (const char* cell : {"B3", "B6", "D3"}) {
      AddLink(wb, cell, std::string("https://") + cell + ".example");
      AddNote(wb, cell, cell);
    }
    ExpectOk(Edit(wb, "B3:C4", c.shift));
    EXPECT_EQ(LinkAt(wb, "https://B3.example"), c.b3);
    EXPECT_EQ(LinkAt(wb, "https://B6.example"), c.b6);
    EXPECT_EQ(LinkAt(wb, "https://D3.example"), c.d3);
    EXPECT_EQ(NoteAt(wb, "B3"), c.b3);
    EXPECT_EQ(NoteAt(wb, "B6"), c.b6);
    EXPECT_EQ(NoteAt(wb, "D3"), c.d3);
  }
}

TEST(InsertDeleteCells, HyperlinkRangesInsideTheBandMoveAndStraddlingOnesStay) {
  Workbook wb = Workbook::create();
  AddLink(wb, "B6:C7", "https://inside.example");
  AddLink(wb, "A6:B7", "https://straddling.example");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(LinkAt(wb, "https://inside.example"), "B8:C9");
  EXPECT_EQ(LinkAt(wb, "https://straddling.example"), "A6:B7");
}

TEST(InsertDeleteCells, MergesInsideTheBandMove) {
  struct Case {
    CellShift shift;
    const char* merges;
  };
  const Case cases[] = {
      {CellShift::kDown, "B10:C10 E3:F3"},
      {CellShift::kRight, "B8:C8 G3:H3"},
      {CellShift::kUp, "B6:C6 E3:F3"},
      {CellShift::kLeft, "B8:C8 C3:D3"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(ShiftName(c.shift));
    Workbook wb = Workbook::create();
    AddMerge(wb, "B8:C8");
    AddMerge(wb, "E3:F3");
    ExpectOk(Edit(wb, "B3:C4", c.shift));
    EXPECT_EQ(Merges(wb), c.merges);
  }
}

TEST(InsertDeleteCells, AMergeStraddlingTheBandIsUnmerged) {
  struct Case {
    CellShift shift;
    const char* merges;
  };
  // A6:B6 straddles the B:C band of a vertical edit and lies outside the 3:4
  // band of a horizontal one.
  const Case cases[] = {
      {CellShift::kDown, ""},
      {CellShift::kRight, "A6:B6"},
      {CellShift::kUp, ""},
      {CellShift::kLeft, "A6:B6"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(ShiftName(c.shift));
    Workbook wb = Workbook::create();
    AddMerge(wb, "A6:B6");
    ExpectOk(Edit(wb, "B3:C4", c.shift));
    EXPECT_EQ(Merges(wb), c.merges);
  }
}

TEST(InsertDeleteCells, AMergeAcrossTheBandEdgeIsUnmergedOrMoved) {
  for (const CellShift shift : kAllShifts) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook wb = Workbook::create();
    AddMerge(wb, "C3:D3");
    SetText(wb, "C3", "m");
    ExpectOk(Edit(wb, "B3:C4", shift));
    switch (shift) {
      case CellShift::kDown:  // C moves, D stays: unmerged.
        EXPECT_EQ(Merges(wb), "");
        EXPECT_EQ(TextAt(wb, "C5"), "m");
        break;
      case CellShift::kRight:  // Wholly in rows 3:4 past the edge: moves.
        EXPECT_EQ(Merges(wb), "E3:F3");
        EXPECT_EQ(TextAt(wb, "E3"), "m");
        break;
      case CellShift::kUp:  // Its anchor is deleted.
        EXPECT_EQ(Merges(wb), "");
        EXPECT_NE(TextAt(wb, "C3"), "m");
        break;
      case CellShift::kLeft:  // Partly deleted.
        EXPECT_EQ(Merges(wb), "");
        EXPECT_NE(TextAt(wb, "C3"), "m");
        EXPECT_NE(TextAt(wb, "B3"), "m");
        break;
    }
  }
}

TEST(InsertDeleteCells, ATableStraddlingAVerticalBandRefusesTheEdit) {
  for (const CellShift shift : kAllShifts) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook wb = Workbook::create();
    AddTable(wb, "T", "A6:B7");
    SetNumber(wb, "B3", 1.0);
    const auto result = Edit(wb, "B3:C4", shift);
    if (shift == CellShift::kDown || shift == CellShift::kUp) {
      ExpectRefused(result);
      EXPECT_EQ(At(wb, "B3"), Value::number(1.0));
    } else {
      ExpectOk(result);
    }
    EXPECT_EQ(TableRef(wb, "T"), "A6:B7");
  }
}

TEST(InsertDeleteCells, ATableStraddlingAHorizontalBandRefusesTheEdit) {
  for (const CellShift shift : kAllShifts) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook wb = Workbook::create();
    AddTable(wb, "T", "D2:E3");
    SetNumber(wb, "B3", 1.0);
    const auto result = Edit(wb, "B3:C4", shift);
    if (shift == CellShift::kRight || shift == CellShift::kLeft) {
      ExpectRefused(result);
      EXPECT_EQ(At(wb, "B3"), Value::number(1.0));
    } else {
      ExpectOk(result);
    }
    EXPECT_EQ(TableRef(wb, "T"), "D2:E3");
  }
}

TEST(InsertDeleteCells, ATableInsideTheBandMovesWithIt) {
  const std::pair<CellShift, const char*> cases[] = {
      {CellShift::kDown, "D3:E4"},
      {CellShift::kRight, "F3:G4"},
      {CellShift::kUp, "D3:E4"},
      {CellShift::kLeft, "B3:C4"},
  };
  for (const auto& [shift, ref] : cases) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook wb = Workbook::create();
    AddTable(wb, "T", "D3:E4");
    ExpectOk(Edit(wb, "B3:C4", shift));
    EXPECT_EQ(TableRef(wb, "T"), ref);
  }
}

TEST(InsertDeleteCells, AnArrayFormulaStraddlingTheBandRefusesTheEdit) {
  for (const CellShift shift : kAllShifts) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook wb = Workbook::create();
    AddArray(wb, "A6:B7");
    SetNumber(wb, "B3", 1.0);
    const auto result = Edit(wb, "B3:C4", shift);
    if (shift == CellShift::kDown || shift == CellShift::kUp) {
      ExpectRefused(result);
      EXPECT_EQ(At(wb, "B3"), Value::number(1.0));
    } else {
      ExpectOk(result);
    }
    EXPECT_EQ(ArrayAt(wb, "A6"), "A6:B7");
  }
}

TEST(InsertDeleteCells, AnArrayFormulaInsideTheBandMoves) {
  const std::pair<CellShift, const char*> cases[] = {
      {CellShift::kDown, "B8"},
      {CellShift::kRight, "B6"},
      {CellShift::kUp, "B4"},
      {CellShift::kLeft, "B6"},
  };
  for (const auto& [shift, anchor] : cases) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook wb = Workbook::create();
    AddArray(wb, "B6:C7");
    ExpectOk(Edit(wb, "B3:C4", shift));
    const MergeRange at = Rect(anchor);
    EXPECT_EQ(ArrayAt(wb, anchor), A1(MergeRange{at.first_row, at.first_col, at.first_row + 1U, at.first_col + 1U}));
  }
}

TEST(InsertDeleteCells, ASpillAnchorMovesOnlyInsideTheBand) {
  for (const CellShift shift : kAllShifts) {
    SCOPED_TRACE(ShiftName(shift));
    Workbook wb = Workbook::create();
    SetFormula(wb, "A6", "=SEQUENCE(3,2)");  // A6:B8 straddles B:C
    SetFormula(wb, "B12", "=SEQUENCE(2)");   // inside B:C
    SetFormula(wb, "A3", "=SEQUENCE(1,3)");  // A3:C3 straddles B:C, inside rows 3:4
    Recalc(wb);
    ExpectOk(Edit(wb, "B3:C4", shift));
    Recalc(wb);
    EXPECT_EQ(FormulaAt(wb, "A6"), "=SEQUENCE(3,2)");
    EXPECT_EQ(FormulaAt(wb, "A3"), "=SEQUENCE(1,3)");
    const char* moved = shift == CellShift::kDown ? "B14" : shift == CellShift::kUp ? "B10" : "B12";
    EXPECT_EQ(FormulaAt(wb, moved), "=SEQUENCE(2)");
    EXPECT_EQ(At(wb, moved), Value::number(1.0));
  }
}

TEST(InsertDeleteCells, ASpillStraddlingTheBandReSpillsAfterTheEdit) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A6", "=SEQUENCE(3,2)");
  Recalc(wb);
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  Recalc(wb);
  EXPECT_EQ(At(wb, "A6"), Value::number(1.0));
  EXPECT_EQ(At(wb, "B6"), Value::number(2.0));
  EXPECT_EQ(At(wb, "B8"), Value::number(6.0));
}

// ---------------------------------------------------------------------------
// Inserted cells take the format of the cell above or to the left
// ---------------------------------------------------------------------------

std::uint32_t XfAt(const Workbook& wb, std::string_view a1) {
  const Cell* cell = CellOf(wb, a1);
  return cell == nullptr ? 0U : cell->xf_index;
}

TEST(InsertDeleteCells, InsertedCellsTakeTheFormatAboveWhenShiftingDown) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_xf_index(1U, 1U, 7U);  // B2
  wb.sheet(0).set_cell_xf_index(1U, 2U, 8U);  // C2
  wb.sheet(0).set_cell_xf_index(2U, 1U, 9U);  // B3, moves to B5
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(XfAt(wb, "B3"), 7U);
  EXPECT_EQ(XfAt(wb, "B4"), 7U);
  EXPECT_EQ(XfAt(wb, "C3"), 8U);
  EXPECT_EQ(XfAt(wb, "C4"), 8U);
  EXPECT_EQ(XfAt(wb, "B5"), 9U);
}

TEST(InsertDeleteCells, InsertedCellsTakeTheFormatLeftWhenShiftingRight) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_xf_index(2U, 0U, 5U);  // A3
  wb.sheet(0).set_cell_xf_index(3U, 0U, 6U);  // A4
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(XfAt(wb, "B3"), 5U);
  EXPECT_EQ(XfAt(wb, "C3"), 5U);
  EXPECT_EQ(XfAt(wb, "B4"), 6U);
  EXPECT_EQ(XfAt(wb, "C4"), 6U);
}

TEST(InsertDeleteCells, InsertedCellsOnTheFirstRowTakeNoFormat) {
  Workbook wb = Workbook::create();
  RowLayout row;
  row.row = 0U;
  row.style_xf = 4U;
  row.has_style = true;
  wb.sheet(0).mutable_layout().row_overrides.push_back(row);
  ExpectOk(Edit(wb, "B1:C1", CellShift::kDown));
  EXPECT_EQ(XfAt(wb, "B1"), 0U);
}

TEST(InsertDeleteCells, RowAndColumnLayoutDoNotMove) {
  Workbook wb = Workbook::create();
  RowLayout row;
  row.row = 5U;
  row.height = 42.0;
  row.has_height = true;
  wb.sheet(0).mutable_layout().row_overrides.push_back(row);
  ColumnLayout col;
  col.first = 3U;
  col.last = 3U;
  col.width = 33.5;
  wb.sheet(0).mutable_layout().columns.push_back(col);
  ManualBreak brk;
  brk.id = 5U;
  brk.manual = true;
  wb.sheet(0).mutable_print_settings().manual_row_breaks.push_back(brk);
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  ASSERT_EQ(wb.sheet(0).layout().row_overrides.size(), 1U);
  EXPECT_EQ(wb.sheet(0).layout().row_overrides[0].row, 5U);
  ASSERT_EQ(wb.sheet(0).layout().columns.size(), 1U);
  EXPECT_EQ(wb.sheet(0).layout().columns[0].first, 3U);
  ASSERT_EQ(wb.sheet(0).print_settings().manual_row_breaks.size(), 1U);
  EXPECT_EQ(wb.sheet(0).print_settings().manual_row_breaks[0].id, 5U);
}

// ---------------------------------------------------------------------------
// Measured: PivotTable at A3:B6 (one row field, grand total row)
// ---------------------------------------------------------------------------

// Region (South, West) by Sum of Amount at A3, laid out by the engine:
// header A3:B3, South 150 / West 200, grand total 350 on row 6.
void AddProjectedPivot(Workbook& wb) {
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(1U);
  cache->mutable_fields().push_back(pivot::PivotCacheField{"Region", {Value::text("South"), Value::text("West")}});
  cache->mutable_fields().push_back(pivot::PivotCacheField{"Amount", {}});
  const std::pair<double, double> rows[] = {{0.0, 100.0}, {1.0, 200.0}, {0.0, 50.0}};
  for (const auto& [region, amount] : rows) {
    pivot::PivotCacheRecord record;
    record.cells = {Value::number(region), Value::number(amount)};
    record.cell_is_index = {true, false};
    cache->mutable_records().push_back(std::move(record));
  }
  wb.add_pivot_cache(std::move(cache));

  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("PivotTable1");
  table->set_pivot_cache_id(1U);
  pivot::PivotField region;
  region.axis = pivot::PivotAxis::Row;
  region.subtotal_top = false;
  for (const std::uint32_t index : {0U, 1U}) {
    pivot::PivotItem item;
    item.has_cache_index = true;
    item.cache_index = index;
    region.items.push_back(item);
  }
  table->mutable_fields().push_back(std::move(region));
  pivot::PivotField amount;
  amount.axis = pivot::PivotAxis::Value;
  amount.subtotal_top = false;
  table->mutable_fields().push_back(std::move(amount));
  table->mutable_row_field_order() = {0U};
  pivot::PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1U;
  table->mutable_data_fields().push_back(std::move(sum));
  table->set_grand_totals(true, true);
  table->set_anchor(2U, 0U, 4U, 2U);
  wb.sheet(0).add_pivot_table(std::move(table));
  ASSERT_TRUE(wb.project_pivot(0U, 0U, ReportPlacement::kBlock).has_value());
  ASSERT_EQ(PivotAt(wb), "A3:B6");
  ASSERT_EQ(At(wb, "B6"), Value::number(350.0));
}

TEST(InsertDeleteCells, APivotOutsideTheEditStays) {
  const std::pair<const char*, CellShift> cases[] = {
      {"A8:B8", CellShift::kDown},  {"D3:D6", CellShift::kDown}, {"C3:C6", CellShift::kRight},
      {"A8:A9", CellShift::kRight}, {"A8:B8", CellShift::kUp},   {"C3:C6", CellShift::kLeft},
      {"C4", CellShift::kLeft},
  };
  for (const auto& [range, shift] : cases) {
    SCOPED_TRACE(std::string(range) + " " + ShiftName(shift));
    Workbook wb = Workbook::create();
    AddProjectedPivot(wb);
    ExpectOk(Edit(wb, range, shift));
    EXPECT_EQ(PivotAt(wb), "A3:B6");
    EXPECT_EQ(At(wb, "B6"), Value::number(350.0));
  }
}

TEST(InsertDeleteCells, APivotTheBandCoversMovesAsABlock) {
  struct Case {
    const char* range;
    CellShift shift;
    const char* pivot;
    const char* total;
  };
  const Case cases[] = {
      {"A1:B1", CellShift::kUp, "A2:B5", "B5"},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(std::string(c.range) + " " + ShiftName(c.shift));
    Workbook wb = Workbook::create();
    AddProjectedPivot(wb);
    ExpectOk(Edit(wb, c.range, c.shift));
    EXPECT_EQ(PivotAt(wb), c.pivot);
    EXPECT_EQ(At(wb, c.total), Value::number(350.0));
  }
}

void ExpectPivotEditRefused(const char* range, CellShift shift) {
  SCOPED_TRACE(std::string(range) + " " + ShiftName(shift));
  Workbook wb = Workbook::create();
  AddProjectedPivot(wb);
  ExpectRefused(Edit(wb, range, shift));
  EXPECT_EQ(PivotAt(wb), "A3:B6");
  EXPECT_EQ(At(wb, "B6"), Value::number(350.0));
}

TEST(InsertDeleteCells, ABandStraddlingAPivotIsRefused) {
  ExpectPivotEditRefused("B1:C1", CellShift::kDown);
  ExpectPivotEditRefused("B5:C6", CellShift::kDown);
  ExpectPivotEditRefused("A2:A4", CellShift::kRight);
  // Excel returns success and changes nothing here.
  ExpectPivotEditRefused("B1:C1", CellShift::kUp);
}

TEST(InsertDeleteCells, AnInsertInsideAPivotIsRefused) {
  ExpectPivotEditRefused("A4:B5", CellShift::kDown);
  ExpectPivotEditRefused("B4", CellShift::kDown);
  ExpectPivotEditRefused("A3:A6", CellShift::kRight);
  ExpectPivotEditRefused("B4", CellShift::kRight);
}

TEST(InsertDeleteCells, ADeleteInsideAPivotIsRefused) {
  ExpectPivotEditRefused("A4:B4", CellShift::kUp);
  ExpectPivotEditRefused("A4", CellShift::kLeft);
}

TEST(InsertDeleteCells, AnInsertWhosePathHoldsAPivotIsRefused) {
  ExpectPivotEditRefused("A1:B1", CellShift::kDown);
  ExpectPivotEditRefused("A1:A9", CellShift::kRight);
  ExpectPivotEditRefused("A1:C2", CellShift::kDown);
}

TEST(InsertDeleteCells, AnInsertAboveAPivotOutsideItsColumnsSucceeds) {
  Workbook wb = Workbook::create();
  AddBarePivot(wb, "A6:B7");
  ExpectOk(Edit(wb, "C2:D2", CellShift::kDown));
  EXPECT_EQ(PivotAt(wb), "A6:B7");
  ExpectRefused(Edit(wb, "B2", CellShift::kDown));
  EXPECT_EQ(PivotAt(wb), "A6:B7");
}

// ---------------------------------------------------------------------------
// Insert Cells across a table
// ---------------------------------------------------------------------------

// Table T on A1:E4: Item / Qty / Dbl (B*2) / Arr (B+1, array) / Bad (does
// not parse), data rows 2-4 with column C's formulas written.
void AddCalcTable(Workbook& wb, bool totals_row = false) {
  TableMetadata table = MakeTable("T", totals_row ? "A1:E5" : "A1:E4");
  table.totals_row = totals_row;
  const char* names[] = {"Item", "Qty", "Dbl", "Arr", "Bad"};
  for (std::size_t c = 0; c < 5U; ++c) {
    table.columns[c].name = names[c];
    SetText(wb, std::string(1, static_cast<char>('A' + c)) + "1", names[c]);
  }
  table.columns[2].calculated_column_formula = "B2*2";
  table.columns[3].calculated_column_formula = "B2+1";
  table.columns[3].calculated_column_formula_array = true;
  table.columns[4].calculated_column_formula = "SUM(";
  for (int r = 2; r <= 4; ++r) {
    const std::string row = std::to_string(r);
    SetText(wb, "A" + row, "i" + row);
    SetNumber(wb, "B" + row, static_cast<double>(r));
    SetFormula(wb, "C" + row, "=B" + row + "*2");
  }
  if (totals_row) {
    SetText(wb, "A5", "Total");
  }
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
}

TEST(InsertDeleteCells, AnInsertDownInsideATableFillsItsCalculatedColumns) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb);
  ExpectOk(Edit(wb, "A3:E3", CellShift::kDown));
  EXPECT_EQ(TableRef(wb, "T"), "A1:E5");
  EXPECT_EQ(FormulaAt(wb, "C3"), "=B3*2");
  EXPECT_EQ(FormulaAt(wb, "C4"), "=B4*2");
  EXPECT_EQ(FormulaAt(wb, "D3"), "=B3+1");
  EXPECT_EQ(ArrayAt(wb, "D3"), "D3");
  EXPECT_TRUE(At(wb, "A3").is_blank());
  EXPECT_TRUE(At(wb, "B3").is_blank());
  EXPECT_TRUE(FormulaAt(wb, "E3").empty());
  EXPECT_TRUE(FormulaAt(wb, "D4").empty());
  SetNumber(wb, "B3", 5.0);
  Recalc(wb);
  EXPECT_EQ(At(wb, "C3"), Value::number(10.0));
  EXPECT_EQ(At(wb, "D3"), Value::number(6.0));
}

TEST(InsertDeleteCells, AnInsertDownBesideATableLeavesItsCalculatedColumnsAlone) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb);
  SetNumber(wb, "C3", 99.0);
  ExpectOk(Edit(wb, "G3:H3", CellShift::kDown));
  EXPECT_EQ(TableRef(wb, "T"), "A1:E4");
  EXPECT_TRUE(FormulaAt(wb, "C3").empty());
  EXPECT_EQ(At(wb, "C3"), Value::number(99.0));
}

TEST(InsertDeleteCells, AnInsertAtTheTotalsRowAddsDataRowsAboveIt) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb, /*totals_row=*/true);
  ExpectOk(Edit(wb, "A5:E6", CellShift::kDown));
  EXPECT_EQ(TableRef(wb, "T"), "A1:E7");
  EXPECT_TRUE(wb.tables()[0].totals_row);
  EXPECT_EQ(TextAt(wb, "A7"), "Total");
  EXPECT_EQ(FormulaAt(wb, "C5"), "=B5*2");
  EXPECT_EQ(FormulaAt(wb, "C6"), "=B6*2");
  EXPECT_TRUE(FormulaAt(wb, "C7").empty());
}

TEST(InsertDeleteCells, AnInsertDirectlyBelowATableLeavesItUnchanged) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb);
  SetText(wb, "A5", "below");
  ExpectOk(Edit(wb, "A5:E5", CellShift::kDown));
  EXPECT_EQ(TableRef(wb, "T"), "A1:E4");
  EXPECT_EQ(TextAt(wb, "A6"), "below");
  EXPECT_TRUE(FormulaAt(wb, "C5").empty());
}

TEST(InsertDeleteCells, AWholeRowInsertInsideATableFillsItsCalculatedColumns) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb);
  ExpectOk(wb.insert_rows(0U, 2U, 2U));
  EXPECT_EQ(TableRef(wb, "T"), "A1:E6");
  EXPECT_EQ(FormulaAt(wb, "C3"), "=B3*2");
  EXPECT_EQ(FormulaAt(wb, "C4"), "=B4*2");
  EXPECT_EQ(FormulaAt(wb, "C5"), "=B5*2");
}

TEST(InsertDeleteCells, AWholeRowInsertAtTheFirstDataRowKeepsTheFormulasAnchored) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb);
  ExpectOk(wb.insert_rows(0U, 1U, 1U));
  EXPECT_EQ(TableRef(wb, "T"), "A1:E5");
  EXPECT_EQ(wb.tables()[0].columns[2].calculated_column_formula, "B2*2");
  EXPECT_EQ(FormulaAt(wb, "C2"), "=B2*2");
  EXPECT_EQ(FormulaAt(wb, "C3"), "=B3*2");
  EXPECT_EQ(FormulaAt(wb, "D2"), "=B2+1");
}

TEST(InsertDeleteCells, AnInsertOverhangingATableIsRefused) {
  struct Case {
    const char* range;
    CellShift shift;
  };
  const Case cases[] = {
      {"A3:F3", CellShift::kDown},
      {"A3:E3", CellShift::kRight},
      {"B1:B5", CellShift::kRight},
      {"B2:B4", CellShift::kRight},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(std::string(c.range) + " " + ShiftName(c.shift));
    Workbook wb = Workbook::create();
    AddCalcTable(wb);
    ExpectRefused(Edit(wb, c.range, c.shift));
    EXPECT_EQ(TableRef(wb, "T"), "A1:E4");
    EXPECT_EQ(wb.tables()[0].columns.size(), 5U);
  }
}

TEST(InsertDeleteCells, AnInsertRightMustCoverTheTotalsRow) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb, /*totals_row=*/true);
  ExpectRefused(Edit(wb, "B1:B4", CellShift::kRight));
  EXPECT_EQ(TableRef(wb, "T"), "A1:E5");
  ExpectOk(Edit(wb, "B1:B5", CellShift::kRight));
  EXPECT_EQ(TableRef(wb, "T"), "A1:F5");
  EXPECT_EQ(wb.tables()[0].columns.size(), 6U);
}

TEST(InsertDeleteCells, AnInsertRightAtATablesFirstColumnMovesItWhole) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb);
  ExpectOk(Edit(wb, "A1:A4", CellShift::kRight));
  EXPECT_EQ(TableRef(wb, "T"), "B1:F4");
  EXPECT_EQ(wb.tables()[0].columns.size(), 5U);
  EXPECT_EQ(TextAt(wb, "B1"), "Item");
}

TEST(InsertDeleteCells, AnInsertRightInsideATableAddsANamedColumn) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb);
  ExpectOk(Edit(wb, "C1:C4", CellShift::kRight));
  EXPECT_EQ(TableRef(wb, "T"), "A1:F4");
  ASSERT_EQ(wb.tables()[0].columns.size(), 6U);
  EXPECT_EQ(wb.tables()[0].columns[2].name, "Column1");
  EXPECT_EQ(TextAt(wb, "C1"), "Column1");
  EXPECT_EQ(TextAt(wb, "D1"), "Dbl");
}

TEST(InsertDeleteCells, AnotherTableInsideTheShiftedRegionMovesAlong) {
  Workbook wb = Workbook::create();
  AddCalcTable(wb);
  AddTable(wb, "Below", "A8:B9");
  ExpectOk(Edit(wb, "A3:E3", CellShift::kDown));
  EXPECT_EQ(TableRef(wb, "T"), "A1:E5");
  EXPECT_EQ(TableRef(wb, "Below"), "A9:B10");
}

// ---------------------------------------------------------------------------
// Table totals row: ON inserts cells below the table, OFF deletes them
// ---------------------------------------------------------------------------

// Table T on A1:B4: header K / V, three data rows.
void AddKv(Workbook& wb) {
  TableMetadata table = MakeTable("T", "A1:B4");
  table.columns[0].name = "K";
  table.columns[1].name = "V";
  SetText(wb, "A1", "K");
  SetText(wb, "B1", "V");
  for (std::uint32_t r = 1U; r <= 3U; ++r) {
    ASSERT_TRUE(wb.set_cell_text(0U, r, 0U, "k" + std::to_string(r)).has_value());
    ASSERT_TRUE(wb.set_cell_value(0U, r, 1U, Value::number(static_cast<double>(r))).has_value());
  }
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
}

void TotalsOn(Workbook& wb) {
  ExpectOk(wb.set_table_totals_row(0U, true));
}

void TotalsOff(Workbook& wb) {
  ExpectOk(wb.set_table_totals_row(0U, false));
}

void ExpectTotalsRefused(Workbook& wb, bool shown, const char* ref) {
  ExpectRefused(wb.set_table_totals_row(0U, shown));
  EXPECT_EQ(wb.tables()[0].ref, ref);
  EXPECT_EQ(wb.tables()[0].totals_row, !shown);
}

TEST(InsertDeleteCells, TotalsOnWithAFreeRowAndNothingBelowGrowsTheTable) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_TRUE(wb.tables()[0].totals_row);
  // A table that never showed a totals row gets the default label and, over numbers, sum.
  EXPECT_EQ(TextAt(wb, "A5"), "Total");
  EXPECT_EQ(FormulaAt(wb, "B5"), "=SUBTOTAL(109,T[V])");
}

TEST(InsertDeleteCells, TotalsOnShiftsTheTableColumnsDownEvenWhenTheRowBelowIsFree) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  SetText(wb, "A7", "g");
  SetText(wb, "B9", "h");
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_EQ(TextAt(wb, "A8"), "g");
  EXPECT_EQ(TextAt(wb, "B10"), "h");
  EXPECT_TRUE(At(wb, "A7").is_blank());
}

TEST(InsertDeleteCells, TotalsOnShiftsAnOccupiedRowAndLeavesOtherColumns) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  SetText(wb, "A5", "a5");
  SetText(wb, "B5", "b5");
  SetText(wb, "C5", "c5");
  SetText(wb, "A6", "row6");
  SetText(wb, "C4", "c4");
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_EQ(TextAt(wb, "A6"), "a5");
  EXPECT_EQ(TextAt(wb, "B6"), "b5");
  EXPECT_EQ(TextAt(wb, "A7"), "row6");
  EXPECT_EQ(TextAt(wb, "C5"), "c5");
  EXPECT_EQ(TextAt(wb, "C4"), "c4");
}

TEST(InsertDeleteCells, TotalsOnShiftsASingleOccupiedCell) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  SetText(wb, "B5", "onlyB");
  SetText(wb, "C5", "c5");
  TotalsOn(wb);
  EXPECT_EQ(TextAt(wb, "B6"), "onlyB");
  EXPECT_EQ(TextAt(wb, "C5"), "c5");
}

TEST(InsertDeleteCells, TotalsOnAdjustsReferencesToTheShiftedCells) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  SetText(wb, "A5", "x");
  SetFormula(wb, "D1", "=SUM(B2:B5)");
  SetFormula(wb, "D2", "=A5");
  TotalsOn(wb);
  EXPECT_EQ(TextAt(wb, "A6"), "x");
  EXPECT_EQ(FormulaAt(wb, "D1"), "=SUM(B2:B6)");
  EXPECT_EQ(FormulaAt(wb, "D2"), "=A6");
}

TEST(InsertDeleteCells, TotalsOnIsRefusedForATableOnTheLastGridRow) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(MakeTable("T", "A1048573:B1048576")).has_value());
  ExpectTotalsRefused(wb, true, "A1048573:B1048576");
}

TEST(InsertDeleteCells, TotalsOnIsRefusedWhenItWouldPushAValueOffTheGrid) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(MakeTable("T", "A1048572:B1048575")).has_value());
  SetText(wb, "A1048576", "z");
  ExpectTotalsRefused(wb, true, "A1048572:B1048575");
  EXPECT_EQ(TextAt(wb, "A1048576"), "z");
}

// One formula at E1 with A5:C10 filled: its text after ON, and after OFF
// (nullptr: OFF is refused because the formula references the totals row).
struct TotalsFormulaCase {
  const char* formula;
  const char* after_on;
  const char* after_off;
};

constexpr TotalsFormulaCase kTotalsFormulaCases[] = {
    {"=SUM(A1:C10)", "=SUM(A1:C10)", nullptr},     {"=SUM(B5:C5)", "=SUM(B5:C5)", nullptr},
    {"=SUM(A6:C6)", "=SUM(A6:C6)", "=SUM(A6:C6)"}, {"=COUNTA(A:A)", "=COUNTA(A:A)", nullptr},
    {"=SUM(5:5)", "=SUM(5:5)", nullptr},           {"=SUM(A6:B9)", "=SUM(A7:B10)", "=SUM(A6:B9)"},
    {"=SUM(A3:B8)", "=SUM(A3:B9)", nullptr},
};

void FillA5C10(Workbook& wb) {
  for (std::uint32_t r = 4U; r < 10U; ++r) {
    for (std::uint32_t c = 0U; c < 3U; ++c) {
      ASSERT_TRUE(wb.set_cell_value(0U, r, c, Value::number(Marker(r + 1U, c + 1U))).has_value());
    }
  }
}

TEST(InsertDeleteCells, TotalsOnAndOffAdjustOnlyReferencesInsideTheBand) {
  for (const TotalsFormulaCase& c : kTotalsFormulaCases) {
    SCOPED_TRACE(c.formula);
    Workbook wb = Workbook::create();
    AddKv(wb);
    FillA5C10(wb);
    SetFormula(wb, "E1", c.formula);
    TotalsOn(wb);
    EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
    EXPECT_EQ(At(wb, "A6"), Value::number(Marker(5, 1)));
    EXPECT_EQ(At(wb, "C5"), Value::number(Marker(5, 3)));
    EXPECT_EQ(FormulaAt(wb, "E1"), c.after_on);
    if (c.after_off == nullptr) {
      ExpectTotalsRefused(wb, false, "A1:B5");
      EXPECT_EQ(FormulaAt(wb, "E1"), c.after_on);
      EXPECT_EQ(At(wb, "A6"), Value::number(Marker(5, 1)));
    } else {
      TotalsOff(wb);
      EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
      EXPECT_EQ(FormulaAt(wb, "E1"), c.after_off);
      EXPECT_EQ(At(wb, "A5"), Value::number(Marker(5, 1)));
    }
  }
}

TEST(InsertDeleteCells, TotalsOnAndOffAdjustNamesInsideTheBand) {
  {
    Workbook wb = Workbook::create();
    AddKv(wb);
    SetName(wb, "markOne", "=Sheet1!$A$6");
    TotalsOn(wb);
    EXPECT_EQ(NameFormula(wb, "markOne"), "=Sheet1!$A$7");
    TotalsOff(wb);
    EXPECT_EQ(NameFormula(wb, "markOne"), "=Sheet1!$A$6");
  }
  {
    Workbook wb = Workbook::create();
    AddKv(wb);
    SetName(wb, "markTwo", "=Sheet1!$A$1:$C$6");
    TotalsOn(wb);
    EXPECT_EQ(NameFormula(wb, "markTwo"), "=Sheet1!$A$1:$C$6");
    ExpectTotalsRefused(wb, false, "A1:B5");
  }
}

TEST(InsertDeleteCells, TotalsOnAndOffMoveAMergeInsideTheBand) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddMerge(wb, "A6:B6");
  TotalsOn(wb);
  EXPECT_EQ(Merges(wb), "A7:B7");
  TotalsOff(wb);
  EXPECT_EQ(Merges(wb), "A6:B6");
}

TEST(InsertDeleteCells, TotalsOnTakesTheFreeRowUnderAStraddlingMerge) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddMerge(wb, "A6:C6");
  SetText(wb, "A8", "m");
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_EQ(Merges(wb), "A6:C6");
  EXPECT_EQ(TextAt(wb, "A8"), "m");
  TotalsOff(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_EQ(Merges(wb), "A6:C6");
  EXPECT_EQ(TextAt(wb, "A8"), "m");
}

TEST(InsertDeleteCells, TotalsOnAndOffMoveOrSplitConditionalFormats) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddCf(wb, {"A6:B8"});
  AddCf(wb, {"A6:C8"});
  TotalsOn(wb);
  EXPECT_EQ(CfSqref(wb, 0U), "A7:B9");
  EXPECT_EQ(CfSqref(wb, 1U), "A7:B9 C6:C8");
  TotalsOff(wb);
  EXPECT_EQ(CfSqref(wb, 0U), "A6:B8");
  EXPECT_EQ(CfSqref(wb, 1U), "A6:C8");
}

TEST(InsertDeleteCells, TotalsOffLeavesASplitValidationSplit) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddDv(wb, {"A6:B8"});
  AddDv(wb, {"A6:C8"});
  TotalsOn(wb);
  EXPECT_EQ(DvRanges(wb, 0U), "A7:B9");
  EXPECT_EQ(DvRanges(wb, 1U), "A7:B9 C6:C8");
  TotalsOff(wb);
  EXPECT_EQ(DvRanges(wb, 0U), "A6:B8");
  EXPECT_EQ(DvRanges(wb, 1U), "A6:B8 C6:C8");
}

TEST(InsertDeleteCells, TotalsOnAndOffMoveASecondTableInsideTheBand) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddTable(wb, "Below", "A7:B8");
  TotalsOn(wb);
  EXPECT_EQ(TableRef(wb, "Below"), "A8:B9");
  TotalsOff(wb);
  EXPECT_EQ(TableRef(wb, "Below"), "A7:B8");
}

TEST(InsertDeleteCells, TotalsOnIsRefusedWithAPivotBelowInTheTableColumns) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddBarePivot(wb, "A7:B8");
  ExpectTotalsRefused(wb, true, "A1:B4");
  EXPECT_EQ(PivotAt(wb), "A7:B8");
}

TEST(InsertDeleteCells, TotalsOnTakesTheFreeRowAboveAStraddlingTable) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddTable(wb, "Below", "A7:C8");
  SetText(wb, "A10", "m");
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_EQ(TableRef(wb, "Below"), "A7:C8");
  TotalsOff(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_EQ(TableRef(wb, "Below"), "A7:C8");
  EXPECT_EQ(TextAt(wb, "A10"), "m");
}

TEST(InsertDeleteCells, TotalsOnAndOffMoveSpillAnchorsInsideTheBand) {
  for (const char* formula : {"=SEQUENCE(3)", "=SEQUENCE(2,3)"}) {
    SCOPED_TRACE(formula);
    Workbook wb = Workbook::create();
    AddKv(wb);
    SetFormula(wb, "A7", formula);
    Recalc(wb);
    TotalsOn(wb);
    Recalc(wb);
    EXPECT_EQ(FormulaAt(wb, "A8"), formula);
    EXPECT_EQ(At(wb, "A8"), Value::number(1.0));
    TotalsOff(wb);
    Recalc(wb);
    EXPECT_EQ(FormulaAt(wb, "A7"), formula);
    EXPECT_EQ(At(wb, "A7"), Value::number(1.0));
  }
}

TEST(InsertDeleteCells, TotalsOnAndOffAdjustReferencesFromAnotherSheet) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  const std::size_t xref = wb.add_sheet("xref");
  SetFormula(wb, "A1", "=Sheet1!A6", xref);
  SetFormula(wb, "A2", "=SUM(Sheet1!A6:B6)", xref);
  SetFormula(wb, "A3", "=Sheet1!A5", xref);
  TotalsOn(wb);
  EXPECT_EQ(FormulaAt(wb, "A1", xref), "=Sheet1!A7");
  EXPECT_EQ(FormulaAt(wb, "A2", xref), "=SUM(Sheet1!A7:B7)");
  EXPECT_EQ(FormulaAt(wb, "A3", xref), "=Sheet1!A6");
  TotalsOff(wb);
  EXPECT_EQ(FormulaAt(wb, "A1", xref), "=Sheet1!A6");
  EXPECT_EQ(FormulaAt(wb, "A2", xref), "=SUM(Sheet1!A6:B6)");
  EXPECT_EQ(FormulaAt(wb, "A3", xref), "=Sheet1!A5");
}

TEST(InsertDeleteCells, TotalsOnIgnoresAFarStraddlingMergeAndTakesTheFreeRow) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddMerge(wb, "A100:C100");
  SetText(wb, "B200", "marker");
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_EQ(Merges(wb), "A100:C100");
  EXPECT_EQ(TextAt(wb, "B200"), "marker");
}

TEST(InsertDeleteCells, TotalsOnShiftsAFarMergeInsideTheBand) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddMerge(wb, "A100:B100");
  SetText(wb, "B200", "marker");
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B5");
  EXPECT_EQ(Merges(wb), "A101:B101");
  EXPECT_EQ(TextAt(wb, "B201"), "marker");
}

TEST(InsertDeleteCells, TotalsOnIsRefusedWhenTheRowBelowIsBusyAndAMergeStraddles) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  SetText(wb, "A5", "x");
  AddMerge(wb, "A6:C6");
  ExpectTotalsRefused(wb, true, "A1:B4");
  EXPECT_EQ(TextAt(wb, "A5"), "x");
  EXPECT_EQ(Merges(wb), "A6:C6");
}

TEST(InsertDeleteCells, TotalsOffRemovesTheRowAndKeepsTheLabelsForTheNextOn) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  // As loaded from a file whose table has shown its totals row, so ON restores the columns' settings.
  wb.mutable_tables()[0].totals_row_shown.reset();
  wb.mutable_tables()[0].columns[0].totals_label = "mylabel";
  wb.mutable_tables()[0].columns[1].totals_function = "sum";
  TotalsOn(wb);
  EXPECT_EQ(TextAt(wb, "A5"), "mylabel");
  EXPECT_NE(FormulaAt(wb, "B5").find("SUBTOTAL(109,"), std::string::npos);
  TotalsOff(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_TRUE(At(wb, "A5").is_blank());
  EXPECT_EQ(FormulaAt(wb, "B5"), "");
  EXPECT_EQ(wb.tables()[0].columns[0].totals_label, "mylabel");
  EXPECT_EQ(wb.tables()[0].columns[1].totals_function, "sum");
  TotalsOn(wb);
  EXPECT_EQ(TextAt(wb, "A5"), "mylabel");
  EXPECT_NE(FormulaAt(wb, "B5").find("SUBTOTAL(109,"), std::string::npos);
}

TEST(InsertDeleteCells, TotalsOffPullsTheCellsBelowBackUp) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  SetText(wb, "A5", "belowA");
  SetText(wb, "B5", "belowB");
  SetText(wb, "A6", "row6");
  SetText(wb, "C5", "c5");
  TotalsOn(wb);
  EXPECT_EQ(TextAt(wb, "A6"), "belowA");
  EXPECT_EQ(TextAt(wb, "B6"), "belowB");
  EXPECT_EQ(TextAt(wb, "A7"), "row6");
  TotalsOff(wb);
  EXPECT_EQ(TextAt(wb, "A5"), "belowA");
  EXPECT_EQ(TextAt(wb, "B5"), "belowB");
  EXPECT_EQ(TextAt(wb, "A6"), "row6");
  EXPECT_EQ(TextAt(wb, "C5"), "c5");
}

TEST(InsertDeleteCells, TotalsOffAdjustsAReferenceBelowTheTotalsRow) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  TotalsOn(wb);
  SetFormula(wb, "D2", "=A6");
  SetFormula(wb, "D3", "=SUM(B6:B7)");
  TotalsOff(wb);
  EXPECT_EQ(FormulaAt(wb, "D2"), "=A5");
  EXPECT_EQ(FormulaAt(wb, "D3"), "=SUM(B5:B6)");
}

TEST(InsertDeleteCells, TotalsOffIsRefusedWhileAFormulaReferencesTheTotalsRow) {
  for (const char* formula : {"=B5", "=SUM(B2:B5)", "=SUM(B5:B6)", "=SUM(B2:B6)", "=SUM(A5:B5)"}) {
    SCOPED_TRACE(formula);
    Workbook wb = Workbook::create();
    AddKv(wb);
    TotalsOn(wb);
    SetFormula(wb, "D1", formula);
    ExpectTotalsRefused(wb, false, "A1:B5");
    EXPECT_EQ(FormulaAt(wb, "D1"), formula);
  }
}

TEST(InsertDeleteCells, TotalsOffIsRefusedWhileANameReferencesTheTotalsRow) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  TotalsOn(wb);
  SetName(wb, "TotalCell", "=Sheet1!$B$5");
  ExpectTotalsRefused(wb, false, "A1:B5");
}

TEST(InsertDeleteCells, TotalsOffIsRefusedByAStructuredTotalsReference) {
  for (const char* formula : {"=SUM(T[#Totals])", "=ROWS(T[#All])"}) {
    SCOPED_TRACE(formula);
    Workbook wb = Workbook::create();
    AddKv(wb);
    TotalsOn(wb);
    SetFormula(wb, "D1", formula);
    ExpectTotalsRefused(wb, false, "A1:B5");
  }
}

TEST(InsertDeleteCells, TotalsOffOnlyClearsTheRowAboveAStraddlingMerge) {
  Workbook wb = Workbook::create();
  AddKv(wb);
  AddMerge(wb, "A6:C6");
  SetText(wb, "A7", "m");
  wb.mutable_tables()[0].columns[0].totals_label = "Total";
  TotalsOn(wb);
  ASSERT_EQ(TextAt(wb, "A5"), "Total");
  TotalsOff(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_TRUE(At(wb, "A5").is_blank());
  EXPECT_EQ(Merges(wb), "A6:C6");
  EXPECT_EQ(TextAt(wb, "A7"), "m");
}

// The totals-row expectations on table A1:B3 (band A:B), with Region / Amount
// totals settings.
void AddRegions(Workbook& wb) {
  TableMetadata table = MakeTable("Regions", "A1:B3");
  table.columns[0].name = "Region";
  table.columns[0].totals_label = "Total";
  table.columns[1].name = "Amount";
  table.columns[1].totals_function = "sum";
  SetText(wb, "A2", "East");
  SetNumber(wb, "B2", 10.0);
  SetText(wb, "A3", "West");
  SetNumber(wb, "B3", 30.0);
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
}

TEST(InsertDeleteCells, TotalsOnShiftsAValueInTheRowBelow) {
  Workbook wb = Workbook::create();
  AddRegions(wb);
  SetText(wb, "B4", "x");
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_EQ(TextAt(wb, "B5"), "x");
  EXPECT_EQ(TextAt(wb, "A4"), "Total");
}

TEST(InsertDeleteCells, TotalsOnShiftsAFormulaInTheRowBelow) {
  Workbook wb = Workbook::create();
  AddRegions(wb);
  SetFormula(wb, "A4", "=1+1");
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "A1:B4");
  EXPECT_EQ(FormulaAt(wb, "A5"), "=1+1");
}

TEST(InsertDeleteCells, TotalsOnIsRefusedUnderAStraddlingMergeOnTheRowBelow) {
  Workbook wb = Workbook::create();
  AddRegions(wb);
  AddMerge(wb, "B4:C5");
  ExpectTotalsRefused(wb, true, "A1:B3");
  EXPECT_EQ(Merges(wb), "B4:C5");
}

TEST(InsertDeleteCells, TotalsOnIsRefusedUnderAStraddlingTableOnTheRowBelow) {
  Workbook wb = Workbook::create();
  AddRegions(wb);
  AddTable(wb, "Below", "B4:C6");
  ExpectTotalsRefused(wb, true, "A1:B3");
  EXPECT_EQ(TableRef(wb, "Below"), "B4:C6");
}

TEST(InsertDeleteCells, TotalsOnBlocksASpillCrossingTheNewTotalsRow) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(MakeTable("Right", "C1:D3")).has_value());
  SetFormula(wb, "A4", "=SEQUENCE(1,4)");
  Recalc(wb);
  TotalsOn(wb);
  EXPECT_EQ(wb.tables()[0].ref, "C1:D4");
  Recalc(wb);
  EXPECT_EQ(FormulaAt(wb, "A4"), "=SEQUENCE(1,4)");
  EXPECT_EQ(At(wb, "A4"), Value::error(ErrorCode::Spill));
}

TEST(InsertDeleteCells, TotalsOnIsRefusedForAThreeRowTableOnTheLastGridRow) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.add_table(MakeTable("Bottom", "A1048574:B1048576")).has_value());
  ExpectTotalsRefused(wb, true, "A1048574:B1048576");
}

// ---------------------------------------------------------------------------
// Rule-derived cases: one row each of the pairwise model (direction x
// position against the moving region x object). Vertical edits use band B:C,
// horizontal ones band 3:4; the whole-extent edits are A3:XFD4 / B1:C1048576.
// ---------------------------------------------------------------------------

TEST(InsertDeleteCells, RuleMergeInsideMovesDown) {
  Workbook wb = Workbook::create();
  AddMerge(wb, "B6:C7");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(Merges(wb), "B8:C9");
}

TEST(InsertDeleteCells, RuleSpillAnchorFollowsAWholeWidthDeleteUp) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A6", "=SEQUENCE(2,2)");
  Recalc(wb);
  ExpectOk(Edit(wb, "A3:XFD4", CellShift::kUp));
  Recalc(wb);
  EXPECT_EQ(FormulaAt(wb, "A4"), "=SEQUENCE(2,2)");
  EXPECT_EQ(At(wb, "B5"), Value::number(4.0));
}

TEST(InsertDeleteCells, RuleNoteAtTheEdgeMovesRight) {
  Workbook wb = Workbook::create();
  AddNote(wb, "B3", "n");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(NoteAt(wb, "n"), "D3");
}

TEST(InsertDeleteCells, RuleValidationStraddlingSplitsOnDeleteUp) {
  Workbook wb = Workbook::create();
  AddDv(wb, {"A6:B7"});
  ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(DvRanges(wb), "A6:A7 B4:B5");
}

TEST(InsertDeleteCells, RulePivotOutsideStaysOnInsertDown) {
  Workbook wb = Workbook::create();
  AddBarePivot(wb, "D6:E7");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(PivotAt(wb), "D6:E7");
}

TEST(InsertDeleteCells, RuleConditionalFormatInsideMovesUp) {
  Workbook wb = Workbook::create();
  AddCf(wb, {"B6:C7"});
  ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(CfSqref(wb), "B4:C5");
}

TEST(InsertDeleteCells, RuleDefinedNameOutsideStaysOnInsertRight) {
  Workbook wb = Workbook::create();
  SetName(wb, "N", "=Sheet1!$D$6:$E$7");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(NameFormula(wb, "N"), "=Sheet1!$D$6:$E$7");
}

TEST(InsertDeleteCells, RuleMergeOutsideStaysOnDeleteUp) {
  Workbook wb = Workbook::create();
  AddMerge(wb, "D6:E7");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(Merges(wb), "D6:E7");
}

TEST(InsertDeleteCells, RuleMergeFollowsAWholeHeightDeleteLeft) {
  Workbook wb = Workbook::create();
  AddMerge(wb, "D2:E3");
  ExpectOk(Edit(wb, "B1:C1048576", CellShift::kLeft));
  EXPECT_EQ(Merges(wb), "B2:C3");
}

TEST(InsertDeleteCells, RuleValidationFollowsAWholeWidthInsertDown) {
  Workbook wb = Workbook::create();
  AddDv(wb, {"A6:B7"});
  ExpectOk(Edit(wb, "A3:XFD4", CellShift::kDown));
  EXPECT_EQ(DvRanges(wb), "A8:B9");
}

TEST(InsertDeleteCells, RuleValidationAcrossTheEdgeShrinksOnDeleteLeft) {
  Workbook wb = Workbook::create();
  AddDv(wb, {"A3:D4"});
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  EXPECT_EQ(DvRanges(wb), "A3:B4");
}

TEST(InsertDeleteCells, RuleSpillStraddlingKeepsItsAnchorOnDeleteLeft) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "D2", "=SEQUENCE(2,2)");  // D2:E3 straddles rows 3:4
  Recalc(wb);
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  Recalc(wb);
  EXPECT_EQ(FormulaAt(wb, "D2"), "=SEQUENCE(2,2)");
  EXPECT_EQ(At(wb, "D3"), Value::number(3.0));
}

TEST(InsertDeleteCells, RuleDefinedNameFollowsAWholeWidthInsertDown) {
  Workbook wb = Workbook::create();
  SetName(wb, "N", "=Sheet1!$A$6:$B$7");
  ExpectOk(Edit(wb, "A3:XFD4", CellShift::kDown));
  EXPECT_EQ(NameFormula(wb, "N"), "=Sheet1!$A$8:$B$9");
}

TEST(InsertDeleteCells, RuleMergeStraddlingIsUnmergedOnInsertRight) {
  Workbook wb = Workbook::create();
  AddMerge(wb, "D2:E3");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(Merges(wb), "");
}

TEST(InsertDeleteCells, RuleArrayFormulaOutsideStaysOnDeleteLeft) {
  Workbook wb = Workbook::create();
  AddArray(wb, "D6:E7");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  EXPECT_EQ(ArrayAt(wb, "D6"), "D6:E7");
}

TEST(InsertDeleteCells, RuleSpillInsideMovesRight) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "D3", "=SEQUENCE(2,2)");
  Recalc(wb);
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  Recalc(wb);
  EXPECT_EQ(FormulaAt(wb, "F3"), "=SEQUENCE(2,2)");
  EXPECT_EQ(At(wb, "G4"), Value::number(4.0));
  EXPECT_EQ(FormulaAt(wb, "D3"), "");
}

TEST(InsertDeleteCells, RuleTableStraddlingRefusesInsertDown) {
  Workbook wb = Workbook::create();
  AddTable(wb, "T", "A6:B7");
  ExpectRefused(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(TableRef(wb, "T"), "A6:B7");
}

TEST(InsertDeleteCells, RuleMergeAcrossTheEdgeGrowsOnInsertRight) {
  Workbook wb = Workbook::create();
  AddMerge(wb, "A3:D4");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(Merges(wb), "A3:F4");
}

TEST(InsertDeleteCells, RulePivotStraddlingRefusesInsertRight) {
  Workbook wb = Workbook::create();
  AddBarePivot(wb, "D2:E3");
  ExpectRefused(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(PivotAt(wb), "D2:E3");
}

TEST(InsertDeleteCells, RuleRangeReferenceOutsideStaysOnDeleteUp) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=SUM(D6:E7)");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=SUM(D6:E7)");
}

TEST(InsertDeleteCells, RuleRangeReferenceFollowsAWholeHeightInsertRight) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=SUM(D2:E3)");
  ExpectOk(Edit(wb, "B1:C1048576", CellShift::kRight));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=SUM(F2:G3)");
}

TEST(InsertDeleteCells, RuleDefinedNameAcrossTheEdgeShrinksOnDeleteUp) {
  Workbook wb = Workbook::create();
  SetName(wb, "N", "=Sheet1!$B$2:$C$5");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(NameFormula(wb, "N"), "=Sheet1!$B$2:$C$3");
}

TEST(InsertDeleteCells, RuleArrayFormulaStraddlingRefusesInsertRight) {
  Workbook wb = Workbook::create();
  AddArray(wb, "D2:E3");
  ExpectRefused(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(ArrayAt(wb, "D2"), "D2:E3");
}

TEST(InsertDeleteCells, RuleRangeReferenceInsideMovesLeft) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=SUM(D3:E4)");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=SUM(B3:C4)");
}

TEST(InsertDeleteCells, RuleSpillAcrossTheEdgeKeepsItsAnchorOnInsertDown) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "B2", "=SEQUENCE(4,2)");  // B2:C5
  Recalc(wb);
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  Recalc(wb);
  EXPECT_EQ(FormulaAt(wb, "B2"), "=SEQUENCE(4,2)");
  EXPECT_EQ(At(wb, "B3"), Value::number(3.0));
  EXPECT_EQ(At(wb, "C5"), Value::number(8.0));
}

TEST(InsertDeleteCells, RulePivotFollowsAWholeWidthDeleteUp) {
  Workbook wb = Workbook::create();
  AddBarePivot(wb, "A6:B7");
  ExpectOk(Edit(wb, "A3:XFD4", CellShift::kUp));
  EXPECT_EQ(PivotAt(wb), "A4:B5");
}

TEST(InsertDeleteCells, RuleArrayFormulaInsideMovesDown) {
  Workbook wb = Workbook::create();
  AddArray(wb, "B6:C7");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(ArrayAt(wb, "B8"), "B8:C9");
  EXPECT_EQ(ArrayAt(wb, "B6"), "");
}

TEST(InsertDeleteCells, RuleHyperlinkFollowsAWholeHeightDeleteLeft) {
  Workbook wb = Workbook::create();
  AddLink(wb, "D2:E3", "https://x.example");
  ExpectOk(Edit(wb, "B1:C1048576", CellShift::kLeft));
  EXPECT_EQ(LinkAt(wb, "https://x.example"), "B2:C3");
}

TEST(InsertDeleteCells, RulePivotTheBandCoversMovesLeftAsABlock) {
  Workbook wb = Workbook::create();
  AddBarePivot(wb, "D3:E4");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  EXPECT_EQ(PivotAt(wb), "B3:C4");
}

TEST(InsertDeleteCells, RuleHyperlinkStraddlingStaysWithItsKeyCellOnInsertDown) {
  Workbook wb = Workbook::create();
  AddLink(wb, "A6:B7", "https://x.example");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(LinkAt(wb, "https://x.example"), "A6:B7");
}

TEST(InsertDeleteCells, RuleNoteFollowsAWholeHeightDeleteLeft) {
  Workbook wb = Workbook::create();
  AddNote(wb, "D2", "n");
  ExpectOk(Edit(wb, "B1:C1048576", CellShift::kLeft));
  EXPECT_EQ(NoteAt(wb, "n"), "B2");
}

TEST(InsertDeleteCells, RulePivotAcrossTheEdgeRefusesDeleteUp) {
  Workbook wb = Workbook::create();
  AddBarePivot(wb, "B2:C5");
  ExpectRefused(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(PivotAt(wb), "B2:C5");
}

TEST(InsertDeleteCells, RuleCellReferenceFollowsAWholeWidthDeleteUp) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=A6");
  ExpectOk(Edit(wb, "A3:XFD4", CellShift::kUp));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=A4");
}

TEST(InsertDeleteCells, RuleRangeReferenceAcrossTheEdgeGrowsOnInsertDown) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=SUM(B2:C5)");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=SUM(B2:C7)");
}

TEST(InsertDeleteCells, RuleConditionalFormatFollowsAWholeHeightDeleteLeft) {
  Workbook wb = Workbook::create();
  AddCf(wb, {"D2:E3"});
  ExpectOk(Edit(wb, "B1:C1048576", CellShift::kLeft));
  EXPECT_EQ(CfSqref(wb), "B2:C3");
}

TEST(InsertDeleteCells, RuleConditionalFormatAcrossTheEdgeGrowsOnInsertDown) {
  Workbook wb = Workbook::create();
  AddCf(wb, {"B2:C5"});
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(CfSqref(wb), "B2:C7");
}

TEST(InsertDeleteCells, RuleValidationOutsideStaysOnInsertRight) {
  Workbook wb = Workbook::create();
  AddDv(wb, {"D6:E7"});
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(DvRanges(wb), "D6:E7");
}

TEST(InsertDeleteCells, RuleNoteInsideMovesDown) {
  Workbook wb = Workbook::create();
  AddNote(wb, "B6", "n");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(NoteAt(wb, "n"), "B8");
}

TEST(InsertDeleteCells, RuleTableTheBandCoversShrinksOnDeleteUp) {
  Workbook wb = Workbook::create();
  AddTable(wb, "T", "B2:C5");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(TableRef(wb, "T"), "B2:C3");
}

TEST(InsertDeleteCells, RuleRangeReferenceStraddlingStaysOnInsertRight) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=SUM(D2:E3)");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=SUM(D2:E3)");
}

TEST(InsertDeleteCells, RuleTableInsideMovesLeft) {
  Workbook wb = Workbook::create();
  AddTable(wb, "T", "D3:E4");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  EXPECT_EQ(TableRef(wb, "T"), "B3:C4");
}

TEST(InsertDeleteCells, RuleNoteOutsideStaysOnDeleteUp) {
  Workbook wb = Workbook::create();
  AddNote(wb, "D6", "n");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(NoteAt(wb, "n"), "D6");
}

TEST(InsertDeleteCells, RuleArrayFormulaFollowsAWholeWidthDeleteUp) {
  Workbook wb = Workbook::create();
  AddArray(wb, "A6:B7");
  ExpectOk(Edit(wb, "A3:XFD4", CellShift::kUp));
  EXPECT_EQ(ArrayAt(wb, "A4"), "A4:B5");
}

TEST(InsertDeleteCells, RuleTableFollowsAWholeHeightInsertRight) {
  Workbook wb = Workbook::create();
  AddTable(wb, "T", "D2:E3");
  ExpectOk(Edit(wb, "B1:C1048576", CellShift::kRight));
  EXPECT_EQ(TableRef(wb, "T"), "F2:G3");
}

TEST(InsertDeleteCells, RuleConditionalFormatStraddlingSplitsOnDeleteUp) {
  Workbook wb = Workbook::create();
  AddCf(wb, {"A6:B7"});
  ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(CfSqref(wb), "A6:A7 B4:B5");
}

TEST(InsertDeleteCells, RuleCellReferenceInsideMovesDown) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=B6");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=B8");
}

TEST(InsertDeleteCells, RuleDefinedNameInsideMovesLeft) {
  Workbook wb = Workbook::create();
  SetName(wb, "N", "=Sheet1!$D$3:$E$4");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  EXPECT_EQ(NameFormula(wb, "N"), "=Sheet1!$B$3:$C$4");
}

TEST(InsertDeleteCells, RuleConditionalFormatOutsideStaysOnInsertRight) {
  Workbook wb = Workbook::create();
  AddCf(wb, {"D6:E7"});
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(CfSqref(wb), "D6:E7");
}

TEST(InsertDeleteCells, RuleCellReferenceAtTheEdgeMovesRight) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=B3");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=D3");
}

TEST(InsertDeleteCells, RuleHyperlinkInsideMovesRight) {
  Workbook wb = Workbook::create();
  AddLink(wb, "D3:E4", "https://x.example");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(LinkAt(wb, "https://x.example"), "F3:G4");
}

TEST(InsertDeleteCells, RuleCellReferenceOutsideStaysOnInsertDown) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=D6");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kDown));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=D6");
}

TEST(InsertDeleteCells, RuleCellReferenceFollowsAWholeHeightDeleteLeft) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A1", "=D2");
  ExpectOk(Edit(wb, "B1:C1048576", CellShift::kLeft));
  EXPECT_EQ(FormulaAt(wb, "A1"), "=B2");
}

TEST(InsertDeleteCells, RuleArrayFormulaAcrossTheEdgeRefusesDeleteLeft) {
  Workbook wb = Workbook::create();
  AddArray(wb, "A3:D4");
  ExpectRefused(Edit(wb, "B3:C4", CellShift::kLeft));
  EXPECT_EQ(ArrayAt(wb, "A3"), "A3:D4");
}

TEST(InsertDeleteCells, RuleValidationInsideMovesLeft) {
  Workbook wb = Workbook::create();
  AddDv(wb, {"D3:E4"});
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  EXPECT_EQ(DvRanges(wb), "B3:C4");
}

TEST(InsertDeleteCells, RuleSpillOutsideStaysOnDeleteLeft) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "D6", "=SEQUENCE(2,2)");
  Recalc(wb);
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  Recalc(wb);
  EXPECT_EQ(FormulaAt(wb, "D6"), "=SEQUENCE(2,2)");
  EXPECT_EQ(At(wb, "E7"), Value::number(4.0));
}

TEST(InsertDeleteCells, RuleDefinedNameStraddlingStaysOnInsertRight) {
  Workbook wb = Workbook::create();
  SetName(wb, "N", "=Sheet1!$D$2:$E$3");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(NameFormula(wb, "N"), "=Sheet1!$D$2:$E$3");
}

TEST(InsertDeleteCells, RuleHyperlinkFollowsAWholeWidthDeleteUp) {
  Workbook wb = Workbook::create();
  AddLink(wb, "A6:B7", "https://x.example");
  ExpectOk(Edit(wb, "A3:XFD4", CellShift::kUp));
  EXPECT_EQ(LinkAt(wb, "https://x.example"), "A4:B5");
}

TEST(InsertDeleteCells, RuleHyperlinkOutsideStaysOnInsertRight) {
  Workbook wb = Workbook::create();
  AddLink(wb, "D6:E7", "https://x.example");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kRight));
  EXPECT_EQ(LinkAt(wb, "https://x.example"), "D6:E7");
}

TEST(InsertDeleteCells, RuleTableOutsideStaysOnDeleteUp) {
  Workbook wb = Workbook::create();
  AddTable(wb, "T", "D6:E7");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kUp));
  EXPECT_EQ(TableRef(wb, "T"), "D6:E7");
}

TEST(InsertDeleteCells, RuleHyperlinkAcrossTheEdgeShrinksOnDeleteLeft) {
  Workbook wb = Workbook::create();
  AddLink(wb, "A3:D4", "https://x.example");
  ExpectOk(Edit(wb, "B3:C4", CellShift::kLeft));
  EXPECT_EQ(LinkAt(wb, "https://x.example"), "A3:B4");
}

}  // namespace
}  // namespace formulon
