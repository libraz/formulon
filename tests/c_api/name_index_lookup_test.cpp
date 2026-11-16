// `fm_sheet_get_pivot_index`, `fm_workbook_get_pivot_cache_field_index`,
// `fm_sheet_get_table_column_index` and `fm_workbook_get_table_index`: keyed
// lookups through the shared name resolver (exact, then case, then a unique
// width fold). A miss, an empty name and a width-fold tie are successful calls
// with `*out_present == 0`.

#include <cstddef>
#include <cstdint>
#include <functional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_pivot_test_helpers.h"
#include "gtest/gtest.h"

namespace {

// Fullwidth forms of "sales", split so a hex escape never swallows the next letter.
constexpr const char* kWideSales =
    "\xEF\xBD\x93"
    "\xEF\xBD\x81"
    "\xEF\xBD\x8C"
    "\xEF\xBD\x85"
    "\xEF\xBD\x93";
// "sales" with only its first letter fullwidth, and with only its second.
constexpr const char* kWideFirst =
    "\xEF\xBD\x93"
    "ales";
constexpr const char* kWideSecond =
    "s"
    "\xEF\xBD\x81"
    "les";

constexpr size_t kSentinelIndex = 77U;
constexpr int32_t kSentinelPresent = -1;

struct Found {
  fm_status_t rc;
  size_t index;
  int32_t present;
};

Found PivotLookup(const fm_workbook_t* wb, size_t sheet, const char* name) {
  Found found{0, kSentinelIndex, kSentinelPresent};
  found.rc = fm_sheet_get_pivot_index(wb, sheet, name, &found.present, &found.index);
  return found;
}

Found FieldLookup(const fm_workbook_t* wb, uint32_t cache_id, const char* name) {
  Found found{0, kSentinelIndex, kSentinelPresent};
  found.rc = fm_workbook_get_pivot_cache_field_index(wb, cache_id, name, &found.present, &found.index);
  return found;
}

Found ColumnLookup(const fm_workbook_t* wb, size_t sheet, size_t table, const char* name) {
  Found found{0, kSentinelIndex, kSentinelPresent};
  found.rc = fm_sheet_get_table_column_index(wb, sheet, table, name, &found.present, &found.index);
  return found;
}

void ExpectHit(const Found& found, size_t index, const char* what) {
  EXPECT_EQ(found.rc, 0) << what;
  EXPECT_EQ(found.present, 1) << what;
  EXPECT_EQ(found.index, index) << what;
}

void ExpectMiss(const Found& found, const char* what) {
  EXPECT_EQ(found.rc, 0) << what;
  EXPECT_EQ(found.present, 0) << what;
  EXPECT_EQ(found.index, 0U) << what;
  EXPECT_STREQ(fm_get_last_error_message(), "") << what;
}

/// A cache with one text field per entry of `fields`.
uint32_t AddCache(fm_workbook_t* wb, const std::vector<const char*>& fields) {
  uint32_t cache_id = 0;
  EXPECT_EQ(fm_workbook_add_pivot_cache(wb, 0, &cache_id), 0);
  for (const char* field : fields) {
    EXPECT_EQ(fm_workbook_add_pivot_cache_field(wb, cache_id, field, ignored_index()), 0) << field;
  }
  return cache_id;
}

/// Pivots named `names` on sheet 0, side by side so their report ranges never overlap.
void AddPivots(WorkbookGuard& wb, uint32_t cache_id, const std::vector<const char*>& names) {
  uint32_t column = 0;
  for (const char* name : names) {
    ASSERT_EQ(AddPivot(wb.handle, 0, name, cache_id, 0U, column, ignored_index()), 0)
        << name << ": " << fm_get_last_error_message();
    column += 4U;
  }
}

/// Sheet1 and Sheet2 on a default workbook.
void CreateTwoSheets(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  size_t added = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", append_position(wb.handle), &added), 0);
}

/// A table `name` over `rect` of sheet `sheet`, its header cells holding `columns`.
void AddTableWithHeaders(WorkbookGuard& wb, size_t sheet, const char* name, fm_cell_rect_t rect,
                         const std::vector<const char*>& columns) {
  for (size_t i = 0; i < columns.size(); ++i) {
    ASSERT_EQ(
        fm_sheet_set_text(wb.handle, sheet, rect.first_row, rect.first_col + static_cast<uint32_t>(i), columns[i]), 0);
  }
  ASSERT_EQ(AddTable(wb.handle, sheet, &rect, name, columns.data(), columns.size(), nullptr, 1, 0, ignored_index()), 0)
      << name << ": " << fm_get_last_error_message();
}

std::vector<std::uint8_t> Save(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

using EntryEdit = std::pair<std::string, std::function<std::string(std::string)>>;

/// `archive` with the body of each edited part replaced by its function of the old body.
std::vector<std::uint8_t> RewriteEntry(const std::vector<std::uint8_t>& archive, const std::vector<EntryEdit>& edits) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, archive.data(), archive.size(), 0), MZ_FALSE);
  std::vector<std::string> names;
  std::vector<std::string> bodies;
  for (mz_uint i = 0; i < mz_zip_reader_get_num_files(&reader); ++i) {
    char name[512];
    mz_zip_reader_get_filename(&reader, i, name, sizeof(name));
    names.emplace_back(name);
    bodies.push_back(ExtractZipEntry(archive, name));
    for (const auto& edit : edits) {
      if (names.back() == edit.first) {
        bodies.back() = edit.second(bodies.back());
      }
    }
  }
  mz_zip_reader_end(&reader);
  std::vector<PartFile> parts;
  for (size_t i = 0; i < names.size(); ++i) {
    parts.push_back(PartFile{names[i].c_str(), bodies[i]});
  }
  return BuildZip(parts);
}

std::string ReplaceAll(std::string text, std::string_view from, std::string_view to) {
  EXPECT_NE(text.find(from), std::string::npos) << from;
  for (size_t at = text.find(from); at != std::string::npos; at = text.find(from, at + to.size())) {
    text.replace(at, from.size(), to);
  }
  return text;
}

}  // namespace

// --- fm_sheet_get_pivot_index ----------------------------------------------

TEST(NameIndexPivot, FindsAPivotByExactCaseAndUniqueWidthFold) {
  WorkbookGuard wb;
  CreateTwoSheets(wb);
  const uint32_t cache = AddCache(wb.handle, {"Region"});
  AddPivots(wb, cache, {"Totals", "Sales"});

  ExpectHit(PivotLookup(wb.handle, 0, "Totals"), 0U, "exact first");
  ExpectHit(PivotLookup(wb.handle, 0, "Sales"), 1U, "exact second");
  ExpectHit(PivotLookup(wb.handle, 0, "sALES"), 1U, "case");
  ExpectHit(PivotLookup(wb.handle, 0, kWideSales), 1U, "unique width fold");
}

TEST(NameIndexPivot, AMissIsASuccessfulCallWithoutAPresentFlag) {
  WorkbookGuard wb;
  CreateTwoSheets(wb);
  const uint32_t cache = AddCache(wb.handle, {"Region"});
  AddPivots(wb, cache, {"Sales"});

  for (const char* name : {"", "Missing", "Sales ", "Sale"}) {
    ExpectMiss(PivotLookup(wb.handle, 0, name), name);
  }
  ExpectMiss(PivotLookup(wb.handle, 1, "Sales"), "a pivot of another sheet");
}

TEST(NameIndexPivot, ReportsBoundaryStatuses) {
  WorkbookGuard wb;
  CreateTwoSheets(wb);
  size_t index = 0;
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_pivot_index(nullptr, 0, "Sales", &present, &index), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_pivot_index(wb.handle, 0, nullptr, &present, &index), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_pivot_index(wb.handle, 0, "Sales", nullptr, &index), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_pivot_index(wb.handle, 0, "Sales", &present, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_pivot_index(wb.handle, 9, "Sales", &present, &index), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_pivot_index(wb.handle, 0, "\xFF\xFE", &present, &index), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  // A NULL argument is reported before the sheet index, the sheet index before the text.
  EXPECT_EQ(fm_sheet_get_pivot_index(wb.handle, 9, nullptr, &present, &index), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_pivot_index(wb.handle, 9, "\xFF\xFE", &present, &index), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

// --- fm_workbook_get_pivot_cache_field_index --------------------------------

TEST(NameIndexCacheField, FindsAFieldByExactCaseAndUniqueWidthFold) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const uint32_t cache = AddCache(wb.handle, {"Region", "Sales", "Amount"});

  ExpectHit(FieldLookup(wb.handle, cache, "Region"), 0U, "exact first");
  ExpectHit(FieldLookup(wb.handle, cache, "Amount"), 2U, "exact last");
  ExpectHit(FieldLookup(wb.handle, cache, "SALES"), 1U, "case");
  ExpectHit(FieldLookup(wb.handle, cache, kWideSales), 1U, "unique width fold");
}

TEST(NameIndexCacheField, ATieUnderTheWidthFoldResolvesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const uint32_t cache = AddCache(wb.handle, {kWideFirst, kWideSecond, "Amount"});

  ExpectMiss(FieldLookup(wb.handle, cache, "sales"), "two wide-only matches");
  ExpectHit(FieldLookup(wb.handle, cache, kWideSecond), 1U, "exact still resolves");
}

TEST(NameIndexCacheField, AMissIsASuccessfulCallWithoutAPresentFlag) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const uint32_t cache = AddCache(wb.handle, {"Region"});

  for (const char* name : {"", "Missing", "Region ", "Reg"}) {
    ExpectMiss(FieldLookup(wb.handle, cache, name), name);
  }
}

TEST(NameIndexCacheField, ReportsBoundaryStatuses) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const uint32_t cache = AddCache(wb.handle, {"Region"});
  size_t index = 0;
  int32_t present = 0;
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_index(nullptr, cache, "Region", &present, &index),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_index(wb.handle, cache, nullptr, &present, &index),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_index(wb.handle, cache, "Region", nullptr, &index),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_index(wb.handle, cache, "Region", &present, nullptr),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_index(wb.handle, cache + 100U, "Region", &present, &index),
            FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_index(wb.handle, cache, "\xFF\xFE", &present, &index),
            FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  // The cache is looked up before the text is checked.
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_index(wb.handle, cache + 100U, "\xFF\xFE", &present, &index),
            FM_STATUS_NOT_FOUND);
}

// --- fm_sheet_get_table_column_index ----------------------------------------

TEST(NameIndexTableColumn, FindsAColumnByExactCaseAndUniqueWidthFold) {
  WorkbookGuard wb;
  CreateTwoSheets(wb);
  AddTableWithHeaders(wb, 0, "First", {0, 0, 2, 2}, {"Product", "Sales", "Amount"});
  AddTableWithHeaders(wb, 0, "Second", {0, 4, 2, 5}, {"Plan", "Actual"});

  ExpectHit(ColumnLookup(wb.handle, 0, 0, "Product"), 0U, "exact first");
  ExpectHit(ColumnLookup(wb.handle, 0, 0, "Amount"), 2U, "exact last");
  ExpectHit(ColumnLookup(wb.handle, 0, 0, "sALES"), 1U, "case");
  ExpectHit(ColumnLookup(wb.handle, 0, 0, kWideSales), 1U, "unique width fold");
  ExpectHit(ColumnLookup(wb.handle, 0, 1, "Actual"), 1U, "second table");
  ExpectMiss(ColumnLookup(wb.handle, 0, 1, "Product"), "a column of another table");
}

TEST(NameIndexTableColumn, AMissIsASuccessfulCallWithoutAPresentFlag) {
  WorkbookGuard wb;
  CreateTwoSheets(wb);
  AddTableWithHeaders(wb, 0, "First", {0, 0, 2, 1}, {"Product", "Amount"});

  for (const char* name : {"", "Missing", "Product ", "Prod"}) {
    ExpectMiss(ColumnLookup(wb.handle, 0, 0, name), name);
  }
}

TEST(NameIndexTableColumn, ReportsBoundaryStatuses) {
  WorkbookGuard wb;
  CreateTwoSheets(wb);
  AddTableWithHeaders(wb, 0, "First", {0, 0, 2, 1}, {"Product", "Amount"});
  size_t index = 0;
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_table_column_index(nullptr, 0, 0, "Product", &present, &index),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 0, 0, nullptr, &present, &index),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 0, 0, "Product", nullptr, &index),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 0, 0, "Product", &present, nullptr),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 9, 0, "Product", &present, &index),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 0, 5, "Product", &present, &index),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 1, 0, "Product", &present, &index), FM_STATUS_INDEX_OUT_OF_RANGE)
      << "a sheet with no table has no table 0";
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 0, 0, "\xFF\xFE", &present, &index),
            FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  // NULL, then the sheet index, then the table index, then the text.
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 9, 5, nullptr, &present, &index),
            FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 9, 5, "\xFF\xFE", &present, &index),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_table_column_index(wb.handle, 0, 5, "\xFF\xFE", &present, &index),
            FM_STATUS_INDEX_OUT_OF_RANGE);
}

// --- fm_workbook_get_table_index --------------------------------------------

namespace {

struct TableFound {
  fm_status_t rc;
  size_t sheet;
  size_t index;
  int32_t present;
};

TableFound TableLookup(const fm_workbook_t* wb, const char* name) {
  TableFound found{0, kSentinelIndex, kSentinelIndex, kSentinelPresent};
  found.rc = fm_workbook_get_table_index(wb, name, &found.present, &found.sheet, &found.index);
  return found;
}

void ExpectTableHit(const TableFound& found, size_t sheet, size_t index, const char* what) {
  EXPECT_EQ(found.rc, 0) << what;
  EXPECT_EQ(found.present, 1) << what;
  EXPECT_EQ(found.sheet, sheet) << what;
  EXPECT_EQ(found.index, index) << what;
}

void ExpectTableMiss(const TableFound& found, const char* what) {
  EXPECT_EQ(found.rc, 0) << what;
  EXPECT_EQ(found.present, 0) << what;
  EXPECT_EQ(found.sheet, 0U) << what << ": the sheet index is zeroed";
  EXPECT_EQ(found.index, 0U) << what << ": the table index is zeroed";
  EXPECT_STREQ(fm_get_last_error_message(), "") << what;
}

/// Sheet1 holds the tables `Alpha` and `Sales`; Other holds `Budget`.
void CreateTables(WorkbookGuard& wb) {
  CreateTwoSheets(wb);
  AddTableWithHeaders(wb, 0, "Alpha", {0, 0, 2, 1}, {"A", "B"});
  AddTableWithHeaders(wb, 0, "Sales", {0, 4, 2, 5}, {"Product", "Amount"});
  AddTableWithHeaders(wb, 1, "Budget", {0, 0, 2, 1}, {"Plan", "Actual"});
}

}  // namespace

TEST(NameIndexTable, FindsATableByExactCaseAndUniqueWidthFold) {
  WorkbookGuard wb;
  CreateTables(wb);

  ExpectTableHit(TableLookup(wb.handle, "Alpha"), 0U, 0U, "exact first");
  ExpectTableHit(TableLookup(wb.handle, "Sales"), 0U, 1U, "exact on the same sheet");
  ExpectTableHit(TableLookup(wb.handle, "Budget"), 1U, 0U, "exact on another sheet");
  ExpectTableHit(TableLookup(wb.handle, "sALES"), 0U, 1U, "case");
  ExpectTableHit(TableLookup(wb.handle, kWideSales), 0U, 1U, "unique width fold");
}

TEST(NameIndexTable, AMissIsASuccessfulCallWithZeroedOutputs) {
  WorkbookGuard wb;
  CreateTables(wb);

  for (const char* name : {"", "Missing", "Sales ", "Sale"}) {
    ExpectTableMiss(TableLookup(wb.handle, name), name);
  }
}

TEST(NameIndexTable, ResolvesANameAndADisplayName) {
  WorkbookGuard source;
  ASSERT_EQ(fm_workbook_create_default(&source.handle), 0);
  AddTableWithHeaders(source, 0, "Sales", {0, 0, 2, 1}, {"Product", "Amount"});
  // The file's `name` attribute and its `displayName` differ.
  const std::vector<std::uint8_t> edited =
      RewriteEntry(Save(source.handle), {{"xl/tables/table1.xml", [](std::string xml) {
                                            return ReplaceAll(std::move(xml), " name=\"Sales\"", " name=\"Internal\"");
                                          }}});
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(edited.data(), edited.size(), nullptr, &wb.handle), 0) << fm_get_last_error_message();

  ExpectTableHit(TableLookup(wb.handle, "Sales"), 0U, 0U, "display name");
  ExpectTableHit(TableLookup(wb.handle, "Internal"), 0U, 0U, "name");
  ExpectTableHit(TableLookup(wb.handle, "internal"), 0U, 0U, "name, case");
  ExpectTableMiss(TableLookup(wb.handle, "Other"), "neither");
}

TEST(NameIndexTable, ATieUnderTheWidthFoldResolvesNothing) {
  WorkbookGuard source;
  ASSERT_EQ(fm_workbook_create_default(&source.handle), 0);
  AddTableWithHeaders(source, 0, "TwinA", {0, 0, 2, 1}, {"A", "B"});
  AddTableWithHeaders(source, 0, "TwinB", {0, 4, 2, 5}, {"C", "D"});
  // The engine refuses to write two names that collide under the width fold, so a file carries them.
  const std::vector<std::uint8_t> edited = RewriteEntry(
      Save(source.handle),
      {{"xl/tables/table1.xml", [](std::string xml) { return ReplaceAll(std::move(xml), "TwinA", kWideFirst); }},
       {"xl/tables/table2.xml", [](std::string xml) { return ReplaceAll(std::move(xml), "TwinB", kWideSecond); }}});
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(edited.data(), edited.size(), nullptr, &wb.handle), 0) << fm_get_last_error_message();

  ExpectTableMiss(TableLookup(wb.handle, "sales"), "two wide-only matches");
  ExpectTableHit(TableLookup(wb.handle, kWideFirst), 0U, 0U, "exact still resolves");
  ExpectTableHit(TableLookup(wb.handle, kWideSecond), 0U, 1U, "exact still resolves");
}

TEST(NameIndexTable, ReportsBoundaryStatuses) {
  WorkbookGuard wb;
  CreateTables(wb);
  size_t sheet = 0;
  size_t index = 0;
  int32_t present = 0;
  EXPECT_EQ(fm_workbook_get_table_index(nullptr, "Sales", &present, &sheet, &index), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_table_index(wb.handle, nullptr, &present, &sheet, &index), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_table_index(wb.handle, "Sales", nullptr, &sheet, &index), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_table_index(wb.handle, "Sales", &present, nullptr, &index), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_table_index(wb.handle, "Sales", &present, &sheet, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_table_index(wb.handle, "\xFF\xFE", &present, &sheet, &index),
            FM_STATUS_BINDING_UTF8ENCODING_ERROR);
}
