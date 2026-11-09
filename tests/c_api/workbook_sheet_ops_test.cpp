//
// C ABI smoke tests for the workbook structural mutation surface added
// in the sheet-rename / move / remove + defined-name editing bundle.
// Each test drives the public C entry points (`fm_workbook_*`) only;
// no C++ engine type leaks across the boundary.

#include <cstdint>
#include <cstring>
#include <limits>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_count_of.h"
#include "formulon_c_out_index.h"
#include "formulon_c_setters.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "value.h"

namespace {

// Sentinel index past `sheet_count()` for any test workbook in this
// file, used to drive the out-of-range rejection paths.
constexpr std::uint32_t kOutOfRangeIndex = 99U;

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

// Builds a workbook with three sheets — `Alpha`, `Beta`, `Gamma`.
void MakeThreeSheets(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Alpha", append_position(wb.handle), ignored_index()), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Beta", append_position(wb.handle), ignored_index()), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Gamma", append_position(wb.handle), ignored_index()), 0);
}

const char* SheetName(fm_workbook_t* wb, std::size_t idx) {
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_get_sheet_name(wb, idx, &name), 0);
  return name != nullptr ? name : "";
}

}  // namespace

TEST(WorkbookSheetOpsCApi, RenameUpdatesName) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 1, "Charlie"), 0);
  EXPECT_STREQ(SheetName(wb.handle, 1), "Charlie");
}

TEST(WorkbookSheetOpsCApi, RenameRejectsBadName) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  // Empty name.
  EXPECT_NE(fm_workbook_rename_sheet(wb.handle, 0, ""), 0);
  // Forbidden character.
  EXPECT_NE(fm_workbook_rename_sheet(wb.handle, 0, "Bad/Name"), 0);
  // Collision (case-insensitive).
  EXPECT_NE(fm_workbook_rename_sheet(wb.handle, 0, "BETA"), 0);
  // Original name must still be intact.
  EXPECT_STREQ(SheetName(wb.handle, 0), "Alpha");
}

TEST(WorkbookSheetOpsCApi, UnicodeSimpleFoldAddRejectsDuplicateAndRenameWorks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "\xC3\x84", append_position(wb.handle), ignored_index()), 0);  // Ä
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "\xC3\x96", append_position(wb.handle), ignored_index()), 0);  // Ö
  EXPECT_NE(fm_workbook_add_sheet(wb.handle, "\xC3\xA4", append_position(wb.handle), ignored_index()), 0);  // ä
  const fm_status_t collision = fm_workbook_rename_sheet(wb.handle, 1U, "\xC3\xA4");                        // ä
  EXPECT_NE(collision, 0);
  EXPECT_EQ(collision, FM_STATUS_ALREADY_EXISTS);
  EXPECT_STREQ(SheetName(wb.handle, 0U), "\xC3\x84");  // Ä
  EXPECT_STREQ(SheetName(wb.handle, 1U), "\xC3\x96");  // Ö
  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 0U, "\xC3\xA4"), 0);
  EXPECT_STREQ(SheetName(wb.handle, 0U), "\xC3\xA4");
  EXPECT_STREQ(SheetName(wb.handle, 1U), "\xC3\x96");  // Ö
}

TEST(WorkbookSheetOpsCApi, UnicodeMalformedUtf8Rejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Valid", append_position(wb.handle), ignored_index()), 0);
  const char malformed[] = {'\xC3', '\0'};
  EXPECT_NE(fm_workbook_add_sheet(wb.handle, malformed, append_position(wb.handle), ignored_index()), 0);
  EXPECT_NE(fm_workbook_rename_sheet(wb.handle, 0U, malformed), 0);
}

TEST(WorkbookSheetOpsCApi, AddRejectsNamesExcelsUiRefusesAsRenameDoes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Valid", append_position(wb.handle), ignored_index()), 0);
  for (const char* name : {"History", "HISTORY", "'a", "a'"}) {
    const fm_status_t added = fm_workbook_add_sheet(wb.handle, name, append_position(wb.handle), ignored_index());
    EXPECT_STREQ(fm_get_status_name(added), "INVALID_SHEET_NAME") << name;
    EXPECT_EQ(added, fm_workbook_rename_sheet(wb.handle, 0U, name)) << name;
  }
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 1U);
  EXPECT_STREQ(SheetName(wb.handle, 0U), "Valid");
  // An inner apostrophe is an ordinary character.
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "a'b", append_position(wb.handle), ignored_index()), 0);
}

TEST(WorkbookSheetOpsCApi, RenameRejectsOutOfRange) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  fm_status_t rc = fm_workbook_rename_sheet(wb.handle, kOutOfRangeIndex, "Whatever");
  EXPECT_NE(rc, 0);
}

TEST(WorkbookSheetOpsCApi, RenameNullArgumentsRejected) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  EXPECT_NE(fm_workbook_rename_sheet(nullptr, 0, "X"), 0);
  EXPECT_NE(fm_workbook_rename_sheet(wb.handle, 0, nullptr), 0);
}

TEST(WorkbookSheetOpsCApi, RemoveDropsSheet) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, 1), 0);
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 2U);
  EXPECT_STREQ(SheetName(wb.handle, 0), "Alpha");
  EXPECT_STREQ(SheetName(wb.handle, 1), "Gamma");
}

TEST(WorkbookSheetOpsCApi, RemoveRejectsLastSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);  // single Sheet1
  fm_status_t rc = fm_workbook_remove_sheet(wb.handle, 0);
  EXPECT_NE(rc, 0);
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 1U);
}

TEST(WorkbookSheetOpsCApi, RemoveRejectsOutOfRange) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  fm_status_t rc = fm_workbook_remove_sheet(wb.handle, kOutOfRangeIndex);
  EXPECT_NE(rc, 0);
}

TEST(WorkbookSheetOpsCApi, MoveForward) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  // Alpha -> end: before the position one past the last sheet.
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 0, 3), 0);
  EXPECT_STREQ(SheetName(wb.handle, 0), "Beta");
  EXPECT_STREQ(SheetName(wb.handle, 1), "Gamma");
  EXPECT_STREQ(SheetName(wb.handle, 2), "Alpha");
}

TEST(WorkbookSheetOpsCApi, MoveBackward) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 2, 0), 0);
  EXPECT_STREQ(SheetName(wb.handle, 0), "Gamma");
  EXPECT_STREQ(SheetName(wb.handle, 1), "Alpha");
  EXPECT_STREQ(SheetName(wb.handle, 2), "Beta");
}

TEST(WorkbookSheetOpsCApi, MoveNoOp) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 1, 1), 0);
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 1, 2), 0);
  EXPECT_STREQ(SheetName(wb.handle, 1), "Beta");
}

TEST(WorkbookSheetOpsCApi, MoveRejectsOutOfRange) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  EXPECT_NE(fm_workbook_move_sheet(wb.handle, kOutOfRangeIndex, 0), 0);
  EXPECT_NE(fm_workbook_move_sheet(wb.handle, 0, kOutOfRangeIndex), 0);
}

TEST(WorkbookSheetOpsCApi, SetDefinedNameAddsAndUpdates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 0U);

  ASSERT_EQ(SetDefinedName(wb.handle, "MyName", "=42"), 0);
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 1U);

  fm_defined_name_t dn = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.name, "MyName");
  EXPECT_STREQ(dn.formula, "=42");

  // Update the formula text via case-insensitive match.
  ASSERT_EQ(SetDefinedName(wb.handle, "MYNAME", "=99"), 0);
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 1U);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.name, "MyName");  // authored case preserved
  EXPECT_STREQ(dn.formula, "=99");
}

TEST(WorkbookSheetOpsCApi, SetDefinedNameScopedAllowsWorkbookAndSheetNames) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);

  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "=1"), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "=2", 1), 0);
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 2U);

  fm_defined_name_t dn = FM_STRUCT_INIT(fm_defined_name_t);
  dn.local_sheet_id = -99;
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.name, "Rate");
  EXPECT_STREQ(dn.formula, "=1");
  EXPECT_EQ(dn.local_sheet_id, -1);

  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 1, &dn), 0);
  EXPECT_STREQ(dn.name, "Rate");
  EXPECT_STREQ(dn.formula, "=2");
  EXPECT_EQ(dn.local_sheet_id, 1);

  ASSERT_EQ(SetDefinedName(wb.handle, "RATE", "=3", 1), 0);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 1, &dn), 0);
  EXPECT_STREQ(dn.name, "Rate");
  EXPECT_STREQ(dn.formula, "=3");
  EXPECT_EQ(dn.local_sheet_id, 1);

  ASSERT_EQ(fm_workbook_remove_defined_name(wb.handle, "Rate", 1), 0);
  EXPECT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 1U);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_EQ(dn.local_sheet_id, -1);
}

// The name stores its formula as files carry it, so a save / load cycle
// returns the same text the setter stored.
TEST(WorkbookSheetOpsCApi, SetDefinedNameDropsTheLeadingEquals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Foo", "=Sheet1!$A$1"), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Bar", "=Sheet1!$B$2", 0), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Baz", "Sheet1!$C$3"), 0);
  EXPECT_EQ(SetDefinedName(wb.handle, "Bad", "==1"),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kParserUnexpectedToken));

  const char* expected[] = {"=Sheet1!$A$1", "=Sheet1!$B$2", "=Sheet1!$C$3"};
  fm_defined_name_t dn = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 3U);
  for (std::size_t i = 0; i < 3U; ++i) {
    ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, i, &dn), 0);
    EXPECT_STREQ(dn.formula, expected[i]) << dn.name;
  }

  uint8_t* bytes = nullptr;
  size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, &bytes, &len), 0);
  WorkbookGuard reloaded;
  const fm_status_t loaded = fm_workbook_load(bytes, len, &reloaded.handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(loaded, 0);
  ASSERT_EQ(count_of(fm_workbook_get_defined_name_count, reloaded.handle), 3U);
  // The writer may reorder names, so match them by name.
  const char* names[] = {"Foo", "Bar", "Baz"};
  for (std::size_t i = 0; i < 3U; ++i) {
    ASSERT_EQ(fm_workbook_get_defined_name_at(reloaded.handle, i, &dn), 0);
    std::size_t k = 0;
    while (k < 3U && std::strcmp(dn.name, names[k]) != 0) {
      ++k;
    }
    ASSERT_LT(k, 3U) << dn.name;
    EXPECT_STREQ(dn.formula, expected[k]) << dn.name;
  }
}

TEST(WorkbookSheetOpsCApi, SetDefinedNameScopedRejectsOutOfRangeScope) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  EXPECT_NE(SetDefinedName(wb.handle, "Bad", "=1", -2), 0);
  EXPECT_NE(SetDefinedName(wb.handle, "Bad", "=1", 3), 0);
}

// Excel has no defined name without a formula, so an empty one is refused
// rather than read as a removal; `fm_workbook_remove_defined_name` removes.
TEST(WorkbookSheetOpsCApi, SetDefinedNameRejectsEmptyFormulaAndKeepsTheName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetDefinedName(wb.handle, "Tmp", "=1"), 0);
  EXPECT_EQ(SetDefinedName(wb.handle, "Tmp", ""),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(SetDefinedName(wb.handle, "Fresh", ""),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  ASSERT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 1U);
  fm_defined_name_t dn = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.formula, "=1");
}

// The record replaces the whole entry: attributes it leaves at their
// defaults reset on a name that had them.
TEST(WorkbookSheetOpsCApi, SetDefinedNameReplacesCommentAndHidden) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_defined_name_t full{sizeof(fm_defined_name_t), "Rate", "=0.1", -1, "Tax rate", 1};
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &full), 0);
  fm_defined_name_t dn = FM_STRUCT_INIT(fm_defined_name_t);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.comment, "Tax rate");
  EXPECT_EQ(dn.hidden, 1);

  ASSERT_EQ(SetDefinedName(wb.handle, "Rate", "=0.2"), 0);
  ASSERT_EQ(count_of(fm_workbook_get_defined_name_count, wb.handle), 1U);
  ASSERT_EQ(fm_workbook_get_defined_name_at(wb.handle, 0, &dn), 0);
  EXPECT_STREQ(dn.formula, "=0.2");
  EXPECT_EQ(dn.local_sheet_id, -1);
  EXPECT_EQ(dn.comment, nullptr);
  EXPECT_EQ(dn.hidden, 0);
}

TEST(WorkbookSheetOpsCApi, SetDefinedNameRejectsEmptyName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_status_t rc = SetDefinedName(wb.handle, "", "=1");
  EXPECT_NE(rc, 0);
}

TEST(WorkbookSheetOpsCApi, SetDefinedNameNullArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_NE(SetDefinedName(nullptr, "n", "=1"), 0);
  EXPECT_NE(SetDefinedName(wb.handle, nullptr, "=1"), 0);
  EXPECT_NE(SetDefinedName(wb.handle, "n", nullptr), 0);
  EXPECT_NE(fm_workbook_set_defined_name(wb.handle, nullptr), 0);
  EXPECT_NE(fm_workbook_get_defined_name_at(wb.handle, 0, nullptr), 0);
}

TEST(WorkbookSheetOpsCApi, DeleteRejectsCountPastSheetBoundWithoutMutatingCells) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 42.0), 0);

  // Negative values from JavaScript become large uint32_t counts at the C
  // ABI. They must be rejected before `origin + count` can wrap in the
  // structural-edit implementation.
  constexpr std::uint32_t kWrappedNegative = std::numeric_limits<std::uint32_t>::max();
  EXPECT_NE(fm_sheet_delete_rows(wb.handle, 0, 0, kWrappedNegative), 0);
  EXPECT_NE(fm_sheet_delete_columns(wb.handle, 0, 0, kWrappedNegative), 0);

  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 1, 1, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 42.0);
}

TEST(WorkbookSheetOpsCApi, InsertRowsAndColumnsShiftCells) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2, 1, 42.0), 0);  // B3

  ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0, 1, 1), 0);
  ASSERT_EQ(fm_sheet_insert_columns(wb.handle, 0, 1, 1), 0);

  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 3, 2, &value), 0);  // C4
  EXPECT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 42.0);
}

TEST(WorkbookSheetOpsCApi, EmptyMetadataListsExposeTheirCAbiContracts) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(count_of(fm_sheet_get_table_count, wb.handle, std::size_t{0}), 0U);
  EXPECT_EQ(count_of(fm_workbook_get_passthrough_count, wb.handle), 0U);

  fm_table_info_t table = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_NE(fm_sheet_get_table_at(wb.handle, 0, 0, &table), 0);

  const char* path = nullptr;
  EXPECT_NE(fm_workbook_get_passthrough_at(wb.handle, 0, &path), 0);
}

TEST(WorkbookSheetOpsCApi, RenameAndRemoveRewriteFormulaThroughPublicPath) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 2, 0, 0, 42.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=Gamma!A1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 2, "Delta"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  size_t cell_count = 0;
  ASSERT_EQ(fm_sheet_get_cell_count(wb.handle, 0, &cell_count), 0);
  ASSERT_EQ(cell_count, 1U);
  uint32_t row = 99;
  uint32_t col = 99;
  const char* formula = nullptr;
  fm_value_t cached{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, &formula, &cached), 0);
  EXPECT_EQ(row, 0U);
  EXPECT_EQ(col, 0U);
  ASSERT_NE(formula, nullptr);
  EXPECT_STREQ(formula, "=Delta!A1");
  EXPECT_EQ(cached.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(cached.u.number, 42.0);

  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, 2), 0);
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, &formula, &cached), 0);
  ASSERT_NE(formula, nullptr);
  EXPECT_STREQ(formula, "=#REF!A1");
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Delta", append_position(wb.handle), ignored_index()), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 2, 0, 0, 99.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &cached), 0);
  EXPECT_EQ(cached.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(cached.u.error_code, static_cast<fm_error_code_t>(formulon::ErrorCode::Ref));
}

TEST(WorkbookSheetOpsCApi, AddWithoutNameTakesTheNextDefaultName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::size_t added = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, nullptr, append_position(wb.handle), &added), 0);
  EXPECT_EQ(added, 1U);
  EXPECT_STREQ(SheetName(wb.handle, 1), "Sheet2");
  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 0, "Sheet3"), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, nullptr, append_position(wb.handle), &added), 0);
  EXPECT_STREQ(SheetName(wb.handle, 2), "Sheet1");
}

TEST(WorkbookSheetOpsCApi, AddPlacesTheSheetBeforeTheGivenPosition) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  std::size_t added = 99;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Front", 0, &added), 0);
  EXPECT_EQ(added, 0U);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Middle", 2, &added), 0);
  EXPECT_EQ(added, 2U);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Back", 5, &added), 0);
  EXPECT_EQ(added, 5U);
  const char* const expected[] = {"Front", "Alpha", "Middle", "Beta", "Gamma", "Back"};
  for (std::size_t i = 0; i < 6; ++i) {
    EXPECT_STREQ(SheetName(wb.handle, i), expected[i]) << i;
  }
}

TEST(WorkbookSheetOpsCApi, AddRejectsAPositionPastTheSheetCount) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  std::size_t added = 0;
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "Late", 4, &added), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, nullptr, 4, &added), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 3U);
}

TEST(WorkbookSheetOpsCApi, MoveTakesAnInsertBeforePosition) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 0, 2), 0);  // Alpha goes before Gamma.
  EXPECT_STREQ(SheetName(wb.handle, 0), "Beta");
  EXPECT_STREQ(SheetName(wb.handle, 1), "Alpha");
  EXPECT_STREQ(SheetName(wb.handle, 2), "Gamma");
  EXPECT_EQ(fm_workbook_move_sheet(wb.handle, 0, 4), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_move_sheet(wb.handle, 3, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(WorkbookSheetOpsCApi, StructuralEditsTakeAnInclusiveSpan) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 5, 5, 7.0), 0);  // F6
  ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0, 1, 3), 0);      // three rows open at row 1
  ASSERT_EQ(fm_sheet_insert_columns(wb.handle, 0, 0, 1), 0);   // two columns open at column 0
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 8, 7, &value), 0);
  EXPECT_DOUBLE_EQ(value.u.number, 7.0);
  ASSERT_EQ(fm_sheet_delete_rows(wb.handle, 0, 1, 3), 0);
  ASSERT_EQ(fm_sheet_delete_columns(wb.handle, 0, 0, 1), 0);
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 5, 5, &value), 0);
  EXPECT_DOUBLE_EQ(value.u.number, 7.0);
}

TEST(WorkbookSheetOpsCApi, StructuralEditsRejectAnEmptySpanOrOneBeyondTheSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_insert_rows(wb.handle, 0, 4, 3), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_delete_rows(wb.handle, 0, 4, 3), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_insert_rows(wb.handle, 0, 0, 1048576U), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_insert_columns(wb.handle, 0, 4, 3), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_delete_columns(wb.handle, 0, 0, 16384U), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_insert_rows(wb.handle, 9, 0, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST(WorkbookSheetOpsCApi, ProfileIdsAreListedByPosition) {
  const std::size_t count = count_of(fm_get_profile_id_count);
  ASSERT_GT(count, 0U);
  for (std::size_t i = 0; i < count; ++i) {
    const char* id = nullptr;
    ASSERT_EQ(fm_get_profile_id_at(i, &id), 0);
    fm_workbook_t* wb = nullptr;
    ASSERT_EQ(fm_workbook_create_for_profile(id, &wb), 0) << id;
    fm_workbook_destroy(wb);
  }
  const char* id = nullptr;
  EXPECT_EQ(fm_get_profile_id_at(count, &id), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_get_profile_id_at(0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
}
