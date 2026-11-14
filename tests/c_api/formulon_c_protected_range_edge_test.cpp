//
// C ABI edge values of protected ranges: the name holds 1..255 UTF-16 units and
// the range list 1..8191 rectangles, the limits Excel keeps in a
// `<protectedRange>`. A stored state past them loads and saves as .xlsx, and
// an .xlsb save refuses it.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_destroy(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNotWritable = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoFeatureNotWritable);

// "\xF0\x9F\x98\x80" (U+1F600) is two UTF-16 units.
constexpr const char* kAstral = "\xF0\x9F\x98\x80";

void Create(WorkbookGuard* wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb->handle), 0) << fm_get_last_error_message();
}

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/protection/" + name;
  std::vector<std::uint8_t> out;
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

std::string Repeat(const char* unit, std::size_t times) {
  std::string out;
  for (std::size_t i = 0; i < times; ++i) {
    out += unit;
  }
  return out;
}

// `count` single cells down column A, every other row.
std::vector<fm_cell_rect_t> Cells(std::size_t count) {
  std::vector<fm_cell_rect_t> out;
  for (std::size_t i = 0; i < count; ++i) {
    const auto row = static_cast<uint32_t>(i * 2U);
    out.push_back({row, 0, row, 0});
  }
  return out;
}

fm_status_t Add(const WorkbookGuard& wb, const std::string& name, const std::vector<fm_cell_rect_t>& ranges) {
  fm_protected_range_t r = FM_STRUCT_INIT(fm_protected_range_t);
  r.name = name.c_str();
  r.ranges = ranges.data();
  r.range_count = static_cast<uint32_t>(ranges.size());
  return fm_sheet_add_protected_range(wb.handle, 0, &r, nullptr);
}

std::string NameAt(const WorkbookGuard& wb, std::size_t index) {
  fm_protected_range_t out = FM_STRUCT_INIT(fm_protected_range_t);
  EXPECT_EQ(fm_sheet_get_protected_range_at(wb.handle, 0, index, &out), 0) << fm_get_last_error_message();
  return out.name != nullptr ? out.name : "";
}

TEST(ApiEdge, ProtectedRangeTitleEmptyRefused) {
  WorkbookGuard wb;
  Create(&wb);
  EXPECT_EQ(Add(wb, "", Cells(1)), kInvalidArgument);
}

TEST(ApiEdge, ProtectedRangeTitle255Ok) {
  WorkbookGuard wb;
  Create(&wb);
  EXPECT_EQ(Add(wb, Repeat("a", 255), Cells(1)), 0) << fm_get_last_error_message();
  EXPECT_EQ(Add(wb, Repeat(kAstral, 127) + "a", Cells(1)), 0) << fm_get_last_error_message();
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
}

TEST(ApiEdge, ProtectedRangeTitle256Refused) {
  WorkbookGuard wb;
  Create(&wb);
  EXPECT_EQ(Add(wb, Repeat("a", 256), Cells(1)), kInvalidArgument);
  // UTF-16 units, not code points: 128 astral characters are 256 units.
  EXPECT_EQ(Add(wb, Repeat(kAstral, 128), Cells(1)), kInvalidArgument);
  size_t count = 1;
  ASSERT_EQ(fm_sheet_get_protected_range_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(ApiEdge, ProtectedRangeRectsCapOk) {
  WorkbookGuard wb;
  Create(&wb);
  EXPECT_EQ(Add(wb, "R", Cells(8191)), 0) << fm_get_last_error_message();
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
}

TEST(ApiEdge, ProtectedRangeRectsPastCapRefused) {
  WorkbookGuard wb;
  Create(&wb);
  EXPECT_EQ(Add(wb, "R", Cells(8192)), kInvalidArgument);
  EXPECT_NE(std::string(fm_get_last_error_context()).find("part=ranges requested=8192 limit=8191"), std::string::npos);
}

TEST(ApiEdge, ProtectedRangeUpdateNameTooLongRefused) {
  WorkbookGuard wb;
  Create(&wb);
  ASSERT_EQ(Add(wb, "R", Cells(1)), 0) << fm_get_last_error_message();
  const std::string long_name = Repeat("a", 256);
  fm_protected_range_patch_t patch = FM_STRUCT_INIT(fm_protected_range_patch_t);
  patch.name_op = FM_PATCH_OP_SET;
  patch.name = long_name.c_str();
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "R", &patch), kInvalidArgument);
  EXPECT_EQ(NameAt(wb, 0), "R");
}

TEST(ApiEdge, ProtectedRangeUpdateRectsPastCapRefused) {
  WorkbookGuard wb;
  Create(&wb);
  ASSERT_EQ(Add(wb, "R", Cells(1)), 0) << fm_get_last_error_message();
  const std::vector<fm_cell_rect_t> cells = Cells(8192);
  fm_protected_range_patch_t patch = FM_STRUCT_INIT(fm_protected_range_patch_t);
  patch.ranges_op = FM_PATCH_OP_SET;
  patch.ranges = cells.data();
  patch.range_count = static_cast<uint32_t>(cells.size());
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, "R", &patch), kInvalidArgument);
}

// A name past the limit read from .xlsx stays editable through the fields the
// patch leaves alone, and the .xlsb save refuses it.
TEST(ApiEdge, ProtectedRangeUpdateUnrelatedFieldOnLongNameStillOk) {
  const std::vector<std::uint8_t> source = ReadFixture("title_256_src.xlsx");
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(source.data(), source.size(), &wb.handle), 0) << fm_get_last_error_message();
  const std::string name = NameAt(wb, 0);
  ASSERT_EQ(name.size(), 256U);
  fm_protected_range_patch_t patch = FM_STRUCT_INIT(fm_protected_range_patch_t);
  patch.spin_count_op = FM_PATCH_OP_SET;
  patch.spin_count = 10;
  EXPECT_EQ(fm_sheet_update_protected_range(wb.handle, 0, name.c_str(), &patch), 0) << fm_get_last_error_message();
  BufferGuard xlsx;
  EXPECT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &xlsx.data, &xlsx.len), 0)
      << fm_get_last_error_message();
  BufferGuard xlsb;
  EXPECT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSB, &xlsb.data, &xlsb.len), kNotWritable);
  const std::string context = fm_get_last_error_context();
  EXPECT_NE(context.find("kind=protected_range_title"), std::string::npos) << context;
  EXPECT_NE(context.find("sheet_index=0"), std::string::npos) << context;
}

}  // namespace
