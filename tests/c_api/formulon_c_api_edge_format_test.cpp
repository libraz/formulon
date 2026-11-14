//
// Edge values of the sheet format defaults and the outline level: what the
// typed setters accept and refuse, and what an xlsx Excel refuses to open does
// on load and on an xlsb save.

#include <cmath>
#include <cstdint>
#include <cstdio>
#include <limits>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNotWritable = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoFeatureNotWritable);

enum class Field { kWidth, kHeight, kBase };

fm_sheet_format_defaults_patch_t SetPatch(Field field, double value) {
  fm_sheet_format_defaults_patch_t patch = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  switch (field) {
    case Field::kWidth:
      patch.default_column_width_chars_op = FM_PATCH_OP_SET;
      patch.default_column_width_chars = value;
      break;
    case Field::kHeight:
      patch.default_row_height_pt_op = FM_PATCH_OP_SET;
      patch.default_row_height_pt = value;
      break;
    case Field::kBase:
      patch.base_column_width_chars_op = FM_PATCH_OP_SET;
      patch.base_column_width_chars = value;
      break;
  }
  return patch;
}

fm_status_t Set(const WorkbookGuard& wb, Field field, double value) {
  const fm_sheet_format_defaults_patch_t patch = SetPatch(field, value);
  return fm_sheet_update_format_defaults(wb.handle, 0, &patch);
}

fm_sheet_format_defaults_t Read(const WorkbookGuard& wb) {
  fm_sheet_format_defaults_t out = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  EXPECT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &out), 0) << fm_get_last_error_message();
  return out;
}

double Stored(const fm_sheet_format_defaults_t& defaults, Field field) {
  switch (field) {
    case Field::kWidth:
      return defaults.default_column_width_chars;
    case Field::kHeight:
      return defaults.default_row_height_pt;
    case Field::kBase:
      return defaults.base_column_width_chars;
  }
  return 0.0;
}

void SaveAndReload(const WorkbookGuard& wb, std::int32_t format, WorkbookGuard* reloaded) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, format, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded->handle), 0) << fm_get_last_error_message();
}

std::vector<std::uint8_t> Fixture(const std::string& name) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/format/" + name;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(f);
  return out;
}

void Load(const std::string& fixture, WorkbookGuard* wb) {
  const std::vector<std::uint8_t> bytes = Fixture(fixture);
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb->handle), 0) << fm_get_last_error_message();
}

fm_status_t SaveStatus(const WorkbookGuard& wb, std::int32_t format) {
  BufferGuard saved;
  return fm_workbook_save(wb.handle, format, &saved.data, &saved.len);
}

struct Bounds {
  Field field;
  double max;
  double past_max;
};

constexpr Bounds kBounds[] = {
    {Field::kWidth, 255.0, 255.5},
    {Field::kHeight, 409.6, 409.65},
    {Field::kBase, 255.0, 256.0},
};

TEST(ApiEdge, FormatDefaultsZeroAndMaximumAreAccepted) {
  for (const Bounds& bounds : kBounds) {
    for (const double value : {0.0, bounds.max}) {
      WorkbookGuard wb;
      ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
      ASSERT_EQ(Set(wb, bounds.field, value), 0) << value << fm_get_last_error_message();
      EXPECT_EQ(Stored(Read(wb), bounds.field), value);
    }
  }
}

TEST(ApiEdge, FormatDefaultsOutOfRangeIsRefusedAndNothingChanges) {
  for (const Bounds& bounds : kBounds) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const double before = Stored(Read(wb), bounds.field);
    for (const double value : {bounds.past_max, 1000.0, -1.0, std::numeric_limits<double>::quiet_NaN(),
                               std::numeric_limits<double>::infinity(), -std::numeric_limits<double>::infinity()}) {
      EXPECT_EQ(Set(wb, bounds.field, value), kInvalidArgument) << value;
    }
    const double after = Stored(Read(wb), bounds.field);
    EXPECT_TRUE(after == before || (std::isnan(after) && std::isnan(before)));
  }
}

TEST(ApiEdge, FormatDefaultsFractionalBaseIsRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const double before = Read(wb).base_column_width_chars;
  EXPECT_EQ(Set(wb, Field::kBase, 8.5), kInvalidArgument);
  EXPECT_EQ(Read(wb).base_column_width_chars, before);
  EXPECT_EQ(Set(wb, Field::kWidth, 8.5), 0) << fm_get_last_error_message();
}

TEST(ApiEdge, FormatDefaultsPatchWithOneBadFieldWritesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sheet_format_defaults_t before = Read(wb);
  fm_sheet_format_defaults_patch_t patch = SetPatch(Field::kWidth, 12.0);
  patch.default_row_height_pt_op = FM_PATCH_OP_SET;
  patch.default_row_height_pt = 500.0;
  EXPECT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &patch), kInvalidArgument);
  const fm_sheet_format_defaults_t after = Read(wb);
  EXPECT_EQ(after.has_default_column_width_chars, before.has_default_column_width_chars);
  EXPECT_EQ(after.default_column_width_chars, before.default_column_width_chars);
  EXPECT_EQ(after.has_default_row_height_pt, before.has_default_row_height_pt);
}

TEST(ApiEdge, FormatDefaultsMaximaSurviveBothFormats) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(Set(wb, Field::kWidth, 255.0), 0);
  ASSERT_EQ(Set(wb, Field::kBase, 255.0), 0);
  ASSERT_EQ(Set(wb, Field::kHeight, 409.6), 0);
  for (const std::int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    WorkbookGuard reloaded;
    SaveAndReload(wb, format, &reloaded);
    const fm_sheet_format_defaults_t back = Read(reloaded);
    EXPECT_EQ(back.default_column_width_chars, 255.0) << format;
    EXPECT_EQ(back.base_column_width_chars, 255.0) << format;
    // Excel clamps an xlsb default row height to 409.5 on open.
    EXPECT_EQ(back.default_row_height_pt, format == FM_WORKBOOK_FORMAT_XLSB ? 409.5 : 409.6) << format;
  }
}

TEST(ApiEdge, FormatDefaultsExcelRefusesLoadLenientlyAndXlsbSaveRefuses) {
  for (const char* name : {"base_256_src.xlsx", "base_large_src.xlsx", "width_256_src.xlsx", "width_large_src.xlsx",
                           "width_record_max_src.xlsx"}) {
    SCOPED_TRACE(name);
    WorkbookGuard wb;
    Load(name, &wb);
    EXPECT_EQ(SaveStatus(wb, FM_WORKBOOK_FORMAT_XLSX), 0) << fm_get_last_error_message();
    EXPECT_EQ(SaveStatus(wb, FM_WORKBOOK_FORMAT_XLSB), kNotWritable);
    EXPECT_NE(std::string(fm_get_last_error_context()).find("kind=format_defaults sheet_index=0"), std::string::npos)
        << fm_get_last_error_context();
  }
}

TEST(ApiEdge, OutlineLevel) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const std::uint32_t level : {0U, 7U}) {
    EXPECT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 3U, 3U, level), 0) << level;
    EXPECT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2U, 2U, level), 0) << level;
  }
  EXPECT_EQ(fm_sheet_set_row_outline(wb.handle, 0, 3U, 3U, 8U), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_column_outline(wb.handle, 0, 2U, 2U, 8U), kInvalidArgument);

  // Level 8 from an xlsx Excel refuses loads as written and is never wrapped into the xlsb's three bits.
  for (const char* name : {"outline_row_8_src.xlsx", "outline_col_8_src.xlsx"}) {
    SCOPED_TRACE(name);
    WorkbookGuard loaded;
    Load(name, &loaded);
    std::size_t rows = 0U;
    ASSERT_EQ(fm_sheet_get_row_layout_count(loaded.handle, 0, &rows), 0);
    std::size_t columns = 0U;
    ASSERT_EQ(fm_sheet_get_column_layout_count(loaded.handle, 0, &columns), 0);
    std::uint32_t deepest = 0U;
    for (std::size_t i = 0; i < rows; ++i) {
      fm_row_layout_info_t row = FM_STRUCT_INIT(fm_row_layout_info_t);
      ASSERT_EQ(fm_sheet_get_row_layout_at(loaded.handle, 0, i, &row), 0);
      deepest = row.outline_level > deepest ? row.outline_level : deepest;
    }
    for (std::size_t i = 0; i < columns; ++i) {
      fm_column_layout_info_t column = FM_STRUCT_INIT(fm_column_layout_info_t);
      ASSERT_EQ(fm_sheet_get_column_layout_at(loaded.handle, 0, i, &column), 0);
      deepest = column.outline_level > deepest ? column.outline_level : deepest;
    }
    EXPECT_EQ(deepest, 8U);
    EXPECT_EQ(SaveStatus(loaded, FM_WORKBOOK_FORMAT_XLSB), kNotWritable);
    EXPECT_NE(std::string(fm_get_last_error_context()).find("kind=outline_level sheet_index=0"), std::string::npos)
        << fm_get_last_error_context();
  }
}

}  // namespace
