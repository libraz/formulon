// Same-sheet cut / paste and copy / paste (`Workbook::move_cells` /
// `copy_cells`): the supported set, and one case per input the preflight
// refuses. A refused call leaves the workbook unchanged.

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
#include "io/ooxml_writer.h"
#include "io/zip_reader.h"
#include "merge_range.h"
#include "picture_table.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "table.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

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

void Recalc(Workbook& wb) {
  ASSERT_TRUE(wb.recalc(eval::default_registry()).has_value());
}

void SetNumber(Workbook& wb, std::string_view a1, double value, std::size_t sheet = 0U) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_value(sheet, at.row, at.col, Value::number(value)).has_value()) << a1;
}

void SetText(Workbook& wb, std::string_view a1, std::string_view text) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_text(0U, at.row, at.col, std::string(text)).has_value()) << a1;
}

void SetFormula(Workbook& wb, std::string_view a1, std::string_view formula, std::size_t sheet = 0U) {
  const CellAddress at = Addr(a1);
  ASSERT_TRUE(wb.set_cell_formula(sheet, at.row, at.col, std::string(formula)).has_value()) << a1;
}

std::string FormulaAt(const Workbook& wb, std::string_view a1) {
  const CellAddress at = Addr(a1);
  const Cell* cell = wb.sheet(0).cell_at(at.row, at.col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

Value At(const Workbook& wb, std::string_view a1) {
  const CellAddress at = Addr(a1);
  return wb.sheet(0).resolve_cell_value(at.row, at.col);
}

// Every cell of A1:L12 as formula text or value, so two states compare equal
// exactly when no cell changed.
std::string Fingerprint(const Workbook& wb) {
  std::string out;
  for (std::uint32_t row = 0; row < 12U; ++row) {
    for (std::uint32_t col = 0; col < 12U; ++col) {
      const Cell* cell = wb.sheet(0).cell_at(row, col);
      if (cell == nullptr) {
        continue;
      }
      out += a1::encode_a1(row, col) + "=";
      if (!cell->formula_text.empty()) {
        out += cell->formula_text;
      } else if (cell->cached_value.is_number()) {
        out += "n" + std::to_string(cell->cached_value.as_number());
      } else if (cell->cached_value.is_text()) {
        out += "t" + std::string(cell->cached_value.as_text());
      } else if (!cell->cached_value.is_blank()) {
        out += "?";
      }
      out += ";";
    }
  }
  return out;
}

void ExpectOk(const Expected<void, Error>& result) {
  ASSERT_TRUE(result.has_value()) << result.error().message;
}

void ExpectCode(const Expected<void, Error>& result, FormulonErrorCode code) {
  ASSERT_FALSE(result.has_value());
  EXPECT_EQ(result.error().code, code) << result.error().message;
}

// A move or copy refused as unsupported, with every cell as it was.
void ExpectMoveRefused(Workbook& wb, std::string_view src, std::string_view dst) {
  const std::string before = Fingerprint(wb);
  const CellAddress to = Addr(dst);
  ExpectCode(wb.move_cells(0U, Rect(src), to.row, to.col), FormulonErrorCode::kNotImplemented);
  EXPECT_EQ(Fingerprint(wb), before);
}

void ExpectCopyRefused(Workbook& wb, std::string_view src, std::string_view dst) {
  const std::string before = Fingerprint(wb);
  const CellAddress to = Addr(dst);
  ExpectCode(wb.copy_cells(0U, Rect(src), to.row, to.col), FormulonErrorCode::kNotImplemented);
  EXPECT_EQ(Fingerprint(wb), before);
}

// A1 = 5 and A2 `=A1*2`: a source A1:B2 whose formula stays inside it.
Workbook SourceBook() {
  Workbook wb = Workbook::create();
  SetNumber(wb, "A1", 5.0);
  SetFormula(wb, "A2", "=A1*2");
  return wb;
}

std::string SavedSheetXml(const Workbook& wb) {
  const Expected<std::vector<std::uint8_t>, Error> saved = io::write_ooxml(wb);
  if (!saved.has_value()) {
    ADD_FAILURE() << "save failed";
    return {};
  }
  io::ZipReader zip;
  if (!zip.open(io::ByteSpan{saved.value().data(), saved.value().size()}).has_value()) {
    ADD_FAILURE() << "saved package does not open";
    return {};
  }
  const Expected<std::vector<std::uint8_t>, Error> part = zip.read_entry("xl/worksheets/sheet1.xml");
  if (!part.has_value()) {
    ADD_FAILURE() << "no sheet1.xml";
    return {};
  }
  return std::string(part.value().begin(), part.value().end());
}

// ---------------------------------------------------------------------------
// Supported set
// ---------------------------------------------------------------------------

TEST(RangeMove, MoveRewritesInsideReferencesAndEmptiesTheSource) {
  Workbook wb = Workbook::create();
  SetNumber(wb, "A1", 5.0);
  SetFormula(wb, "A2", "=A1*2");
  SetFormula(wb, "B2", "=$A$1+A$1+$A1");
  SetFormula(wb, "G1", "=Z9");  // untouched by the move
  ASSERT_TRUE(wb.move_cells(0U, Rect("A1:B2"), 4U, 3U).has_value());

  EXPECT_EQ(FormulaAt(wb, "D6"), "=D5*2");
  EXPECT_EQ(FormulaAt(wb, "E6"), "=$D$5+D$5+$D5");
  EXPECT_EQ(FormulaAt(wb, "G1"), "=Z9");
  EXPECT_EQ(At(wb, "D5"), Value::number(5.0));
  for (const char* gone : {"A1", "A2", "B1", "B2"}) {
    EXPECT_TRUE(At(wb, gone).is_blank()) << gone;
    EXPECT_EQ(FormulaAt(wb, gone), "") << gone;
  }
  EXPECT_FALSE(wb.sheet(0).populated_extent(0U, 0U, 1U, 1U).has_value());

  Recalc(wb);
  EXPECT_EQ(At(wb, "D6"), Value::number(10.0));
  EXPECT_EQ(At(wb, "E6"), Value::number(15.0));

  const std::string xml = SavedSheetXml(wb);
  EXPECT_NE(xml.find("D5*2</f>"), std::string::npos);
  EXPECT_NE(xml.find("$D$5+D$5+$D5</f>"), std::string::npos);
}

TEST(RangeMove, MoveCarriesTextAndBooleanValues) {
  Workbook wb = Workbook::create();
  SetText(wb, "A1", "hello");
  ASSERT_TRUE(wb.set_cell_value(0U, 0U, 1U, Value::boolean(true)).has_value());
  ASSERT_TRUE(wb.move_cells(0U, Rect("A1:B1"), 5U, 5U).has_value());
  EXPECT_EQ(At(wb, "F6").as_text(), "hello");
  EXPECT_EQ(At(wb, "G6"), Value::boolean(true));
  EXPECT_FALSE(wb.sheet(0).populated_extent(0U, 0U, 0U, 1U).has_value());
}

TEST(RangeMove, MoveKeepsTheTextOfFormulasElsewhere) {
  Workbook wb = SourceBook();
  SetFormula(wb, "J1", "=SUM(K1:K3)");
  ASSERT_TRUE(wb.move_cells(0U, Rect("A1:A2"), 4U, 3U).has_value());
  EXPECT_EQ(FormulaAt(wb, "J1"), "=SUM(K1:K3)");
  EXPECT_EQ(FormulaAt(wb, "D6"), "=D5*2");
}

TEST(RangeMove, CopyShiftsRelativeReferencesByTheColumnOffset) {
  Workbook wb = Workbook::create();
  SetNumber(wb, "D5", 3.0);
  SetFormula(wb, "D6", "=D5*2");
  SetFormula(wb, "E6", "=$D$5+D$5+$D5");
  ASSERT_TRUE(wb.copy_cells(0U, Rect("D5:E6"), 4U, 7U).has_value());

  EXPECT_EQ(At(wb, "H5"), Value::number(3.0));
  EXPECT_EQ(FormulaAt(wb, "H6"), "=H5*2");
  EXPECT_EQ(FormulaAt(wb, "I6"), "=$D$5+H$5+$D5");
  EXPECT_EQ(At(wb, "D5"), Value::number(3.0));
  EXPECT_EQ(FormulaAt(wb, "D6"), "=D5*2");
  EXPECT_EQ(FormulaAt(wb, "E6"), "=$D$5+D$5+$D5");

  Recalc(wb);
  EXPECT_EQ(At(wb, "H6"), Value::number(6.0));
  EXPECT_EQ(At(wb, "I6"), Value::number(9.0));
}

TEST(RangeMove, CopyShiftsRelativeReferencesByTheRowOffset) {
  Workbook wb = Workbook::create();
  SetNumber(wb, "B2", 1.0);
  SetNumber(wb, "C2", 2.0);
  SetFormula(wb, "D3", "=B2+C2");
  SetFormula(wb, "E3", "=$B$2+B$2+$B2");
  SetFormula(wb, "D10", "=B9");
  SetFormula(wb, "D12", "=B2");
  ASSERT_NO_FATAL_FAILURE(ExpectOk(wb.copy_cells(0U, Rect("D3:D3"), 5U, 3U)));
  ASSERT_NO_FATAL_FAILURE(ExpectOk(wb.copy_cells(0U, Rect("E3:E3"), 6U, 4U)));
  ASSERT_NO_FATAL_FAILURE(ExpectOk(wb.copy_cells(0U, Rect("D10:D10"), 3U, 3U)));
  ASSERT_NO_FATAL_FAILURE(ExpectOk(wb.copy_cells(0U, Rect("D12:D12"), 12U, 3U)));
  EXPECT_EQ(FormulaAt(wb, "D6"), "=B5+C5");
  EXPECT_EQ(FormulaAt(wb, "E7"), "=$B$2+B$2+$B6");
  EXPECT_EQ(FormulaAt(wb, "D4"), "=B3");
  EXPECT_EQ(FormulaAt(wb, "D13"), "=B3");
  EXPECT_EQ(FormulaAt(wb, "D3"), "=B2+C2");
}

TEST(RangeMove, CopyIgnoresReferencesElsewhereToTheDestination) {
  Workbook wb = Workbook::create();
  SetNumber(wb, "D5", 3.0);
  SetFormula(wb, "A1", "=H5");
  ASSERT_TRUE(wb.copy_cells(0U, Rect("D5:D5"), 4U, 7U).has_value());
  EXPECT_EQ(FormulaAt(wb, "A1"), "=H5");
  Recalc(wb);
  EXPECT_EQ(At(wb, "A1"), Value::number(3.0));
}

// ---------------------------------------------------------------------------
// Arguments
// ---------------------------------------------------------------------------

TEST(RangeMove, ADestinationOffTheGridIsAnInvalidArgument) {
  Workbook wb = SourceBook();
  const std::string before = Fingerprint(wb);
  ExpectCode(wb.move_cells(0U, Rect("A1:B2"), 1048575U, 3U), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.move_cells(0U, Rect("A1:B2"), 4U, 16383U), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.copy_cells(0U, Rect("A1:B2"), 0U, 16383U), FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(Fingerprint(wb), before);
}

TEST(RangeMove, ASourceOffTheGridOrInvertedIsAnInvalidArgument) {
  Workbook wb = SourceBook();
  ExpectCode(wb.move_cells(0U, MergeRange{2U, 0U, 1U, 1U}, 4U, 3U), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.move_cells(0U, MergeRange{0U, 0U, 1048576U, 1U}, 4U, 3U), FormulonErrorCode::kInvalidArgument);
}

TEST(RangeMove, ABadSheetIndexIsAnInvalidArgument) {
  Workbook wb = SourceBook();
  ExpectCode(wb.move_cells(5U, Rect("A1:B2"), 4U, 3U), FormulonErrorCode::kInvalidArgument);
  ExpectCode(wb.copy_cells(5U, Rect("A1:B2"), 0U, 3U), FormulonErrorCode::kInvalidArgument);
}

// ---------------------------------------------------------------------------
// Refused formulas and references
// ---------------------------------------------------------------------------

TEST(RangeMove, ARangeReferenceInTheSourceRefuses) {
  Workbook wb = SourceBook();
  SetFormula(wb, "B2", "=SUM(A1:B1)");
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ASourceFormulaPointingOutsideTheSourceRefuses) {
  Workbook wb = SourceBook();
  SetFormula(wb, "B2", "=Z1+1");
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ASourceFormulaWithAQualifiedReferenceRefuses) {
  Workbook wb = SourceBook();
  SetFormula(wb, "B2", "=" + wb.sheet(0).name() + "!A1");
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ASourceFormulaThatBuildsAReferenceRefuses) {
  Workbook wb = SourceBook();
  SetFormula(wb, "B2", "=OFFSET(A1,1,0)");
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ACellOutsideReferringIntoTheSourceRefuses) {
  Workbook wb = SourceBook();
  SetFormula(wb, "G1", "=A1");
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ARangeOutsideCoveringTheSourceRefuses) {
  Workbook wb = SourceBook();
  SetFormula(wb, "G1", "=SUM(A1:B2)");
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ADefinedNamePointingIntoTheSourceRefuses) {
  Workbook wb = SourceBook();
  ASSERT_TRUE(wb.set_defined_name("Rate", wb.sheet(0).name() + "!$A$1").has_value());
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, AnotherSheetPointingIntoTheSourceRefuses) {
  Workbook wb = SourceBook();
  const std::size_t other = wb.add_sheet("Other");
  SetFormula(wb, "A1", "=" + wb.sheet(0).name() + "!A1", other);
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, AThreeDimensionalReferenceSpanningTheSheetRefuses) {
  Workbook wb = SourceBook();
  const std::size_t other = wb.add_sheet("Other");
  SetFormula(wb, "A1", "=SUM(" + wb.sheet(0).name() + ":Other!Z9)", other);
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, AReferenceIntoTheDestinationRefuses) {
  Workbook wb = SourceBook();
  SetFormula(wb, "G1", "=D5");
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, AFormulaThatDoesNotParseRefuses) {
  Workbook wb = SourceBook();
  SetFormula(wb, "G1", "=1+(");
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

// ---------------------------------------------------------------------------
// Refused destinations and sources
// ---------------------------------------------------------------------------

TEST(RangeMove, ANonEmptyDestinationRefuses) {
  Workbook wb = SourceBook();
  SetNumber(wb, "E6", 1.0);
  ExpectMoveRefused(wb, "A1:B2", "D5");
  SetNumber(wb, "E2", 1.0);
  ExpectCopyRefused(wb, "A1:B2", "E1");
}

TEST(RangeMove, AFormattedDestinationRefuses) {
  Workbook wb = SourceBook();
  ASSERT_TRUE(wb.set_cell_xf_index(0U, 5U, 3U, 1U).has_value());
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, AFormattedSourceRefuses) {
  Workbook wb = SourceBook();
  ASSERT_TRUE(wb.set_cell_xf_index(0U, 0U, 1U, 1U).has_value());
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ARowWithAFormatOverTheDestinationRefuses) {
  Workbook wb = SourceBook();
  RowLayout row;
  row.row = 5U;
  row.has_style = true;
  row.style_xf = 1U;
  wb.sheet(0).mutable_layout().row_overrides.push_back(row);
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ACustomRowHeightOverTheSourceRefuses) {
  Workbook wb = SourceBook();
  RowLayout row;
  row.row = 1U;
  row.has_height = true;
  row.custom_height = true;
  row.height = 30.0;
  wb.sheet(0).mutable_layout().row_overrides.push_back(row);
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, OverlappingRectanglesRefuse) {
  Workbook wb = SourceBook();
  ExpectMoveRefused(wb, "A1:B2", "B2");
  ExpectCopyRefused(wb, "A1:B2", "B1");
}

TEST(RangeMove, AWholeRowSourceRefuses) {
  Workbook wb = SourceBook();
  ExpectMoveRefused(wb, "A1:XFD1", "A5");
}

TEST(RangeMove, AProtectedSheetRefuses) {
  Workbook wb = SourceBook();
  wb.sheet(0).mutable_protection().enabled = true;
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

// ---------------------------------------------------------------------------
// Refused sheet objects. A destination holding any object refuses too, so no
// object is ever deleted by a move.
// ---------------------------------------------------------------------------

TEST(RangeMove, AConditionalFormatOverTheSourceRefuses) {
  Workbook wb = SourceBook();
  cf::ConditionalFormat format;
  const MergeRange r = Rect("A1:B2");
  cf::CFCellRange range;
  range.first = CellAddress{r.first_row, r.first_col};
  range.last = CellAddress{r.last_row, r.last_col};
  format.sqref.push_back(range);
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = 1;
  rule.formula1 = "TRUE";
  format.rules.push_back(rule);
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(format));
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ADataValidationOverTheDestinationRefuses) {
  Workbook wb = SourceBook();
  DataValidation dv;
  dv.type = 3;
  dv.formula1 = "\"a,b\"";
  dv.ranges.push_back(Rect("D5:E6"));
  dv.ranges.push_back(Rect("D1:E2"));
  ASSERT_TRUE(wb.add_validation(0U, std::move(dv)).has_value());
  ExpectMoveRefused(wb, "A1:B2", "D5");
  ExpectCopyRefused(wb, "A1:B2", "D1");
}

TEST(RangeMove, AMergedRangeTouchingTheSourceRefuses) {
  Workbook wb = SourceBook();
  ASSERT_TRUE(wb.add_merge(0U, Rect("B1:C1")).has_value());
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ANoteOnTheSourceRefuses) {
  Workbook wb = SourceBook();
  wb.sheet(0).mutable_comments().emplace_back(0U, 0U, "Author", "note");
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, AHyperlinkOverTheSourceRefuses) {
  Workbook wb = SourceBook();
  Hyperlink link;
  link.row = 1U;
  link.col = 0U;
  link.last_row = 1U;
  link.last_col = 0U;
  link.target = "https://example.com/";
  wb.sheet(0).mutable_hyperlinks().push_back(std::move(link));
  ExpectMoveRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ATableOverTheSourceRefuses) {
  Workbook wb = Workbook::create();
  TableMetadata table;
  table.sheet_index = 0U;
  table.name = "T";
  table.display_name = "T";
  table.ref = "B1:C3";
  for (const char* name : {"C1", "C2"}) {
    TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
  ExpectMoveRefused(wb, "A1:B2", "E5");
}

TEST(RangeMove, ASpillOverTheSourceRefuses) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "H1", "=SEQUENCE(3)");
  Recalc(wb);
  ExpectMoveRefused(wb, "G1:H2", "J5");
}

TEST(RangeMove, AnArrayFormulaBlockOverTheSourceRefuses) {
  Workbook wb = SourceBook();
  ASSERT_TRUE(wb.set_array_formula(0U, 1U, 1U, 2U, 2U, "=1").has_value());
  ExpectMoveRefused(wb, "A1:B2", "E5");
}

// ---------------------------------------------------------------------------
// Refused copies
// ---------------------------------------------------------------------------

TEST(RangeMove, ACopyWithBothARowAndAColumnOffsetRefuses) {
  Workbook wb = SourceBook();
  ExpectCopyRefused(wb, "A1:B2", "D5");
}

TEST(RangeMove, ACopyShiftingAReferenceAboveTheFirstRowRefuses) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "D3", "=B1");
  ExpectCopyRefused(wb, "D3:D3", "D1");
}

TEST(RangeMove, ACopyShiftingAReferenceOffTheGridRefuses) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A2", "=XFD1");
  ExpectCopyRefused(wb, "A2:A2", "C2");
}

TEST(RangeMove, ACopyOfAFormulaWithARangeRefuses) {
  Workbook wb = Workbook::create();
  SetFormula(wb, "A2", "=SUM(B1:B3)");
  ExpectCopyRefused(wb, "A2:A2", "C2");
}

// A picture of one fresh media, in the table the move and copy cells read.
PicturePayload AddPicture(Workbook& wb) {
  PictureTable& table = wb.picture_table();
  const std::uint32_t media = table.add_media_at(table.media_count(), MediaKey{8U, 1U + table.media_count()});
  return PicturePayload{table.add_entry(media, std::nullopt), media};
}

void ExpectPlacedPicture(const Value& v, PicturePayload expected) {
  ASSERT_TRUE(v.is_picture());
  EXPECT_EQ(v.as_picture().picture_id, expected.picture_id);
  EXPECT_EQ(v.as_picture().media_id, expected.media_id);
  EXPECT_EQ(v.picture_origin(), 5U);
}

TEST(RangeMove, MovingAConstantPictureCarriesItWithPlacedOrigin) {
  Workbook wb = Workbook::create();
  const PicturePayload picture = AddPicture(wb);
  ASSERT_TRUE(wb.set_cell_value(0U, 0U, 0U, Value::picture(picture, 4U)).has_value());
  ExpectOk(wb.move_cells(0U, Rect("A1"), 0U, 2U));
  ExpectPlacedPicture(At(wb, "C1"), picture);
  EXPECT_FALSE(At(wb, "A1").is_picture());
}

TEST(RangeMove, CopyingAConstantPictureCarriesItWithPlacedOrigin) {
  Workbook wb = Workbook::create();
  const PicturePayload picture = AddPicture(wb);
  ASSERT_TRUE(wb.set_cell_value(0U, 0U, 0U, Value::picture(picture, 4U)).has_value());
  ExpectOk(wb.copy_cells(0U, Rect("A1"), 0U, 2U));
  ExpectPlacedPicture(At(wb, "C1"), picture);
  ASSERT_TRUE(At(wb, "A1").is_picture());
  EXPECT_EQ(At(wb, "A1").picture_origin(), 4U);
}

TEST(RangeMove, CopyingAFormulaOverAPictureRecomputesItsOriginAtTheNewPlace) {
  Workbook wb = Workbook::create();
  const PicturePayload picture = AddPicture(wb);
  ASSERT_TRUE(wb.set_cell_value(0U, 1U, 4U, Value::picture(picture, 4U)).has_value());
  SetFormula(wb, "A1", "=$E$2");
  ExpectOk(wb.copy_cells(0U, Rect("A1"), 3U, 0U));
  Recalc(wb);
  const Value copied = At(wb, "A4");
  ASSERT_TRUE(copied.is_picture());
  EXPECT_EQ(copied.as_picture().picture_id, picture.picture_id);
  EXPECT_EQ(copied.picture_origin(), 4U);
}

// A loaded rich value (verbatim when `vm` is set, else an opaque structure) on a #VALUE! cell.
OpaqueRichValue PutRichValue(Workbook& wb, std::string_view a1, std::uint32_t vm) {
  const CellAddress at = Addr(a1);
  EXPECT_TRUE(wb.set_cell_value(0U, at.row, at.col, Value::error(ErrorCode::Value)).has_value());
  OpaqueRichValue value;
  if (vm != 0U) {
    value.verbatim_vm = vm;
  } else {
    value.structure_type = "_linkedEntity";
    value.keys.push_back(OpaqueRichValueKey{"_DisplayString", "s", "", "Tokyo", "", ""});
  }
  wb.sheet(0).set_rich_value(at.row, at.col, value);
  return value;
}

std::optional<OpaqueRichValue> RichAt(const Workbook& wb, std::string_view a1) {
  const CellAddress at = Addr(a1);
  return wb.sheet(0).rich_value_at(at.row, at.col);
}

TEST(RangeMove, MovingRichValuesCarriesThemAndEmptiesTheSource) {
  Workbook wb = Workbook::create();
  const OpaqueRichValue verbatim = PutRichValue(wb, "A1", 2U);
  const OpaqueRichValue opaque = PutRichValue(wb, "A2", 0U);
  ExpectOk(wb.move_cells(0U, Rect("A1:A2"), 0U, 2U));
  EXPECT_EQ(RichAt(wb, "C1"), verbatim);
  EXPECT_EQ(RichAt(wb, "C2"), opaque);
  EXPECT_FALSE(RichAt(wb, "A1").has_value());
  EXPECT_FALSE(RichAt(wb, "A2").has_value());
  EXPECT_EQ(wb.sheet(0).rich_values().size(), 2U);
}

TEST(RangeMove, CopyingARichValueSharesItWithTheDestination) {
  Workbook wb = Workbook::create();
  const OpaqueRichValue verbatim = PutRichValue(wb, "A1", 3U);
  ExpectOk(wb.copy_cells(0U, Rect("A1"), 0U, 2U));
  EXPECT_EQ(RichAt(wb, "C1"), verbatim);
  EXPECT_EQ(RichAt(wb, "A1"), verbatim);
}

}  // namespace
}  // namespace formulon
