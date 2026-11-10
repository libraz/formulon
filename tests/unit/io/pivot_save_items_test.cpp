//
// What an .xlsx save writes for a pivot's `<rowItems>` / `<colItems>` and its
// cache's `refreshOnLoad`: lines generated from the projected report for an
// edited table, the bytes read for an untouched one, and a rebuild on open
// whenever an edited table gets no lines.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/retained_part_fingerprint.h"
#include "io/xlsb/writer.h"
#include "miniz.h"
#include "passthrough_part.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

std::string ExcelFixture(std::string_view relative) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + std::string(relative);
}

/// An Excel-saved .xlsx under the ja-JP profile it was laid out in.
Workbook LoadXlsx(std::string_view relative) {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(ExcelFixture(relative));
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << relative << ": " << (read_or ? "" : read_or.error().message);
  if (!read_or) {
    return Workbook::create_empty();
  }
  Workbook wb = std::move(read_or.value().workbook);
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id("mac-365-ja_JP", &profile));
  wb.set_excel_profile(profile);
  return wb;
}

Workbook LoadXlsb(std::string_view relative) {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(ExcelFixture(relative));
  auto read_or = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << relative << ": " << (read_or ? "" : read_or.error().message);
  if (!read_or) {
    return Workbook::create_empty();
  }
  return std::move(read_or.value().workbook);
}

pivot::PivotTable* FirstPivot(Workbook& wb) {
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    if (!wb.sheet(s).pivot_tables().empty()) {
      return wb.sheet(s).mutable_pivot_tables().front().get();
    }
  }
  return nullptr;
}

/// The saved pivot table part and cache definition part of a one-pivot book.
struct SavedPivot {
  std::string table;
  std::string cache;
};

SavedPivot SaveXlsx(const Workbook& wb) {
  SavedPivot saved;
  auto bytes_or = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << (bytes_or ? "" : bytes_or.error().message);
  if (!bytes_or) {
    return saved;
  }
  const std::vector<std::uint8_t>& archive = bytes_or.value();
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, archive.data(), archive.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return saved;
  }
  for (mz_uint i = 0; i < mz_zip_reader_get_num_files(&reader); ++i) {
    char name[512] = {};
    if (mz_zip_reader_get_filename(&reader, i, name, sizeof(name)) == 0) {
      continue;
    }
    const std::string_view entry(name);
    std::string* target = nullptr;
    if (entry.rfind("xl/pivotTables/pivotTable", 0) == 0) {
      target = &saved.table;
    } else if (entry.rfind("xl/pivotCache/pivotCacheDefinition", 0) == 0) {
      target = &saved.cache;
    }
    if (target == nullptr) {
      continue;
    }
    std::size_t size = 0;
    void* body = mz_zip_reader_extract_to_heap(&reader, i, &size, 0);
    if (body != nullptr) {
      target->assign(static_cast<const char*>(body), size);
      mz_free(body);
    }
  }
  mz_zip_reader_end(&reader);
  return saved;
}

bool Contains(const std::string& haystack, std::string_view needle) {
  return haystack.find(needle) != std::string::npos;
}

// Two row fields with Values below them on the rows, compact form.
constexpr std::string_view kValuesOnRows = "pivot_layout/ja_d8_two_rows_values_rows.xlsx";

TEST(PivotSaveItems, EditedTableSavesItsProjectedItemsWithoutRefreshOnLoad) {
  Workbook wb = LoadXlsx(kValuesOnRows);
  pivot::PivotTable* table = FirstPivot(wb);
  ASSERT_NE(table, nullptr);
  ASSERT_EQ(table->definition_source(), pivot::PivotDefinitionSource::kXlsx);
  const std::string row_items = table->raw_passthrough_after_row_fields();
  const std::string col_items = table->raw_col_items();
  ASSERT_TRUE(Contains(row_items, "<rowItems count=\"26\">")) << row_items;
  // An edit that leaves the report as Excel drew it: the lines must come back as read.
  table->mark_layout_dirty();
  const SavedPivot saved = SaveXlsx(wb);
  EXPECT_TRUE(Contains(saved.table, "</rowFields>" + row_items)) << saved.table;
  EXPECT_TRUE(Contains(saved.table, col_items)) << saved.table;
  EXPECT_FALSE(Contains(saved.cache, "refreshOnLoad=\"1\"")) << saved.cache;
}

TEST(PivotSaveItems, ProjectionPendingTableSavesNoItemsAndRefreshesOnLoad) {
  Workbook wb = LoadXlsx(kValuesOnRows);
  pivot::PivotTable* table = FirstPivot(wb);
  ASSERT_NE(table, nullptr);
  table->mark_layout_dirty();
  table->set_projection_pending(true);
  const SavedPivot saved = SaveXlsx(wb);
  EXPECT_FALSE(Contains(saved.table, "<rowItems")) << saved.table;
  EXPECT_FALSE(Contains(saved.table, "<colItems")) << saved.table;
  EXPECT_TRUE(Contains(saved.cache, " refreshOnLoad=\"1\"")) << saved.cache;
}

TEST(PivotSaveItems, RefusedShapeSavesNoItemsAndRefreshesOnLoad) {
  Workbook wb = LoadXlsx(kValuesOnRows);
  pivot::PivotTable* table = FirstPivot(wb);
  ASSERT_NE(table, nullptr);
  ASSERT_FALSE(table->fields().empty());
  // A field listing its items without data has no measured line encoding.
  table->mutable_fields()[0].show_all = true;
  table->mark_layout_dirty();
  const SavedPivot saved = SaveXlsx(wb);
  EXPECT_FALSE(Contains(saved.table, "<rowItems")) << saved.table;
  EXPECT_FALSE(Contains(saved.table, "<colItems")) << saved.table;
  EXPECT_TRUE(Contains(saved.cache, " refreshOnLoad=\"1\"")) << saved.cache;
}

TEST(PivotSaveItems, XlsbReadTableSavedAsXlsxGetsNoItemsAndRefreshesOnLoad) {
  Workbook wb = LoadXlsb("xlsb_pivot_base.xlsb");
  pivot::PivotTable* table = FirstPivot(wb);
  ASSERT_NE(table, nullptr);
  EXPECT_EQ(table->definition_source(), pivot::PivotDefinitionSource::kXlsb);
  // The .xlsb reader decodes no cache source, which an .xlsx save requires.
  ASSERT_EQ(wb.mutable_pivot_caches().size(), 1U);
  pivot::WorksheetSource& source = wb.mutable_pivot_caches().front()->mutable_worksheet_source();
  if (!source.present) {
    source.present = true;
    source.sheet = wb.sheet(0).name();
    source.ref = "A1:C4";
  }
  table->mark_layout_dirty();
  const SavedPivot saved = SaveXlsx(wb);
  EXPECT_FALSE(Contains(saved.table, "<rowItems")) << saved.table;
  EXPECT_FALSE(Contains(saved.table, "<colItems")) << saved.table;
  EXPECT_TRUE(Contains(saved.cache, " refreshOnLoad=\"1\"")) << saved.cache;
}

TEST(PivotSaveItems, UneditedXlsxTableSavesTheItemsItWasReadWith) {
  Workbook wb = LoadXlsx(kValuesOnRows);
  pivot::PivotTable* table = FirstPivot(wb);
  ASSERT_NE(table, nullptr);
  const std::string row_items = table->raw_passthrough_after_row_fields();
  const std::string col_items = table->raw_col_items();
  ASSERT_FALSE(row_items.empty());
  ASSERT_FALSE(col_items.empty());
  const SavedPivot saved = SaveXlsx(wb);
  EXPECT_TRUE(Contains(saved.table, "</rowFields>" + row_items)) << saved.table;
  EXPECT_TRUE(Contains(saved.table, col_items)) << saved.table;
  EXPECT_FALSE(Contains(saved.cache, "refreshOnLoad=\"1\"")) << saved.cache;
}

TEST(PivotSaveItems, UneditedXlsbPivotKeepsItsRetainedPartFingerprint) {
  Workbook wb = LoadXlsb("xlsb_pivot_base.xlsb");
  std::size_t checked = 0;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.retained_origin != PassthroughPart::RetainedOrigin::kPivotTable) {
      continue;
    }
    ++checked;
    ASSERT_TRUE(part.model_fingerprint.has_value()) << part.path;
    EXPECT_EQ(io::xlsb::current_retained_part_fingerprint(wb, part), part.model_fingerprint) << part.path;
  }
  EXPECT_EQ(checked, 1U);
  auto saved = io::xlsb::write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
}

}  // namespace
}  // namespace formulon
