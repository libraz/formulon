// The workbook location through the C ABI: every status row of
// `fm_workbook_set_location`, the canonical read back, CELL("filename") and
// CELL("address") from it, and the relative local hyperlink targets a save
// writes.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

struct Location {
  bool present = false;
  bool has_directory = false;
  std::string directory;
  std::string file_name;
};

Location Read(fm_workbook_t* wb) {
  Location out;
  std::int32_t present = -1;
  fm_workbook_location_t record = FM_STRUCT_INIT(fm_workbook_location_t);
  EXPECT_EQ(fm_workbook_get_location(wb, &present, &record), 0) << fm_get_last_error_message();
  out.present = present != 0;
  out.has_directory = record.directory != nullptr;
  out.directory = record.directory != nullptr ? record.directory : "";
  out.file_name = record.file_name != nullptr ? record.file_name : "";
  return out;
}

fm_status_t Set(fm_workbook_t* wb, const char* directory, const char* file_name) {
  fm_workbook_location_t record = FM_STRUCT_INIT(fm_workbook_location_t);
  record.directory = directory;
  record.file_name = file_name;
  return fm_workbook_set_location(wb, &record);
}

std::string TextAt(fm_workbook_t* wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, sheet, row, col, &v), 0);
  return v.kind == FM_VALUE_KIND_TEXT ? std::string(v.u.text) : std::string("<not text>");
}

class FormulonCApiLocation : public ::testing::Test {
 protected:
  void SetUp() override { ASSERT_EQ(fm_workbook_create_default(&wb_.handle), 0); }
  fm_workbook_t* wb() { return wb_.handle; }

 private:
  WorkbookGuard wb_;
};

TEST_F(FormulonCApiLocation, UnsetByDefaultAndByNull) {
  EXPECT_FALSE(Read(wb()).present);
  ASSERT_EQ(Set(wb(), "/d", "b.xlsx"), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_location(wb(), nullptr), 0);
  EXPECT_FALSE(Read(wb()).present);
}

TEST_F(FormulonCApiLocation, NullPointersAndShortRecordsAreRefused) {
  std::int32_t present = 0;
  fm_workbook_location_t out = FM_STRUCT_INIT(fm_workbook_location_t);
  EXPECT_EQ(fm_workbook_get_location(nullptr, &present, &out), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_location(wb(), nullptr, &out), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_location(wb(), &present, nullptr), kBindingNullPointer);
  out.struct_size = 4U;
  EXPECT_EQ(fm_workbook_get_location(wb(), &present, &out), kInvalidArgument);

  fm_workbook_location_t in = FM_STRUCT_INIT(fm_workbook_location_t);
  in.file_name = "b.xlsx";
  EXPECT_EQ(fm_workbook_set_location(nullptr, &in), kBindingNullPointer);
  in.struct_size = 4U;
  EXPECT_EQ(fm_workbook_set_location(wb(), &in), kInvalidArgument);
  in = FM_STRUCT_INIT(fm_workbook_location_t);
  in.directory = "/d";
  EXPECT_EQ(fm_workbook_set_location(wb(), &in), kBindingNullPointer);
  EXPECT_FALSE(Read(wb()).present);
}

TEST_F(FormulonCApiLocation, ReadsBackTheCanonicalFormAsAFixedPoint) {
  struct Row {
    const char* input;
    const char* canonical;
  };
  const Row rows[] = {
      {"/a/b/", "/a/b"}, {"/", "/"}, {"c:\\a\\", "C:/a"}, {"C:\\", "C:/"}, {"\\\\server\\share\\a", "//server/share/a"},
  };
  for (const Row& row : rows) {
    ASSERT_EQ(Set(wb(), row.input, "b.xlsx"), 0) << row.input << ": " << fm_get_last_error_message();
    Location got = Read(wb());
    ASSERT_TRUE(got.has_directory) << row.input;
    EXPECT_EQ(got.directory, row.canonical) << row.input;
    ASSERT_EQ(Set(wb(), got.directory.c_str(), got.file_name.c_str()), 0) << row.input;
    EXPECT_EQ(Read(wb()).directory, row.canonical) << row.input;
  }
  ASSERT_EQ(Set(wb(), nullptr, "b.xlsx"), 0);
  const Location name_only = Read(wb());
  EXPECT_TRUE(name_only.present);
  EXPECT_FALSE(name_only.has_directory);
  EXPECT_EQ(name_only.file_name, "b.xlsx");
}

TEST_F(FormulonCApiLocation, EveryRefusalIsInvalidArgumentAndChangesNothing) {
  ASSERT_EQ(Set(wb(), "/keep", "Keep.xlsx"), 0);
  const std::string long_name = std::string(214, 'n') + ".xlsx";
  struct Row {
    const char* directory;
    const char* file_name;
  };
  const Row rows[] = {
      {"/d", ""},
      {"/d", "Book"},
      {"/d", "a[1].xlsx"},
      {"/d", "a:b.xlsx"},
      {"/d", "\xFF.xlsx"},
      {"/d", long_name.c_str()},
      {"", "b.xlsx"},
      {"relative", "b.xlsx"},
      {"C:", "b.xlsx"},
      {"C:rel", "b.xlsx"},
      {"\\\\server", "b.xlsx"},
      {"//server", "b.xlsx"},
      {"//a", "b.xlsx"},
      {"/a[b", "b.xlsx"},
      {"/a\xFF", "b.xlsx"},
      {"/a/../b", "b.xlsx"},
      {"/a//b", "b.xlsx"},
  };
  for (const Row& row : rows) {
    EXPECT_EQ(Set(wb(), row.directory, row.file_name), kInvalidArgument) << row.directory << " | " << row.file_name;
    const Location kept = Read(wb());
    EXPECT_EQ(kept.directory, "/keep");
    EXPECT_EQ(kept.file_name, "Keep.xlsx");
  }
}

TEST_F(FormulonCApiLocation, APathPast1024BytesIsAResourceLimitAndChangesNothing) {
  ASSERT_EQ(Set(wb(), "/keep", "Keep.xlsx"), 0);
  const std::string long_directory = "/" + std::string(1017, 'd');
  EXPECT_EQ(Set(wb(), long_directory.c_str(), "b.xlsx"), FM_STATUS_SECURITY_RESOURCE_LIMIT);
  EXPECT_STREQ(fm_get_last_error_context(), "part=location_path requested=1025 limit=1024");
  const Location kept = Read(wb());
  EXPECT_EQ(kept.directory, "/keep");
  EXPECT_EQ(kept.file_name, "Keep.xlsx");
}

TEST_F(FormulonCApiLocation, CellFilenameAndAddressFollowTheLocation) {
  ASSERT_EQ(fm_workbook_set_profile_id(wb(), "mac-365-ja_JP"), 0);
  std::size_t index = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb(), "Sheet2", 1U, &index), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb(), "S 2", 2U, &index), 0);
  const char* formulas[] = {
      "=CELL(\"filename\",A1)",      "=CELL(\"filename\",'S 2'!A1)", "=CELL(\"address\",Sheet2!A1)",
      "=CELL(\"address\",'S 2'!A1)", "=CELL(\"address\",Sheet1!A1)",
  };
  for (std::uint32_t row = 0; row < 5U; ++row) {
    ASSERT_EQ(fm_sheet_set_formula(wb(), 0, row, 0, formulas[row]), 0) << fm_get_last_error_message();
  }

  ASSERT_EQ(fm_workbook_recalc(wb()), 0);
  EXPECT_EQ(TextAt(wb(), 0, 0, 0), "");
  EXPECT_EQ(TextAt(wb(), 0, 2, 0), "Sheet2!$A$1");
  EXPECT_EQ(TextAt(wb(), 0, 4, 0), "$A$1");

  ASSERT_EQ(Set(wb(), "/U/x/a b", "Book 1.xlsx"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb()), 0);
  EXPECT_EQ(TextAt(wb(), 0, 0, 0), "/U/x/a b/[Book 1.xlsx]Sheet1");
  EXPECT_EQ(TextAt(wb(), 0, 1, 0), "/U/x/a b/[Book 1.xlsx]S 2");
  EXPECT_EQ(TextAt(wb(), 0, 2, 0), "'[Book 1.xlsx]Sheet2'!$A$1");
  EXPECT_EQ(TextAt(wb(), 0, 3, 0), "'[Book 1.xlsx]S 2'!$A$1");
  EXPECT_EQ(TextAt(wb(), 0, 4, 0), "$A$1");

  ASSERT_EQ(Set(wb(), nullptr, "Book.xlsx"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb()), 0);
  EXPECT_EQ(TextAt(wb(), 0, 0, 0), "");
  EXPECT_EQ(TextAt(wb(), 0, 2, 0), "[Book.xlsx]Sheet2!$A$1");
}

struct SavedLink {
  std::string target;
  std::string display;
};

std::vector<SavedLink> SaveAndReadLinks(fm_workbook_t* wb, std::int32_t format) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, format, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  EXPECT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();
  std::vector<SavedLink> out;
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_hyperlink_count(reloaded.handle, 0, &count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
    EXPECT_EQ(fm_sheet_get_hyperlink_at(reloaded.handle, 0, i, &hl), 0);
    out.push_back({hl.target != nullptr ? hl.target : "", hl.display != nullptr ? hl.display : ""});
  }
  return out;
}

TEST_F(FormulonCApiLocation, SaveWritesSharedPrefixTargetsRelative) {
  const std::string data = "/Users/libraz/Library/Containers/com.microsoft.Excel/Data";
  const std::string targets[] = {
      "/Users/libraz/Documents/a b/f.xlsx",
      "/private/tmp/x.xlsx",
      data + "/loc_probe3/p q/sib.xlsx",
  };
  for (std::uint32_t row = 0; row < 3U; ++row) {
    fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
    hl.rect = fm_cell_rect_t{row, 0, row, 0};
    hl.target = targets[row].c_str();
    ASSERT_EQ(fm_sheet_set_hyperlink(wb(), 0, &hl), 0) << fm_get_last_error_message();
  }
  for (const std::int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    const std::vector<SavedLink> unset = SaveAndReadLinks(wb(), format);
    ASSERT_EQ(unset.size(), 3U);
    EXPECT_EQ(unset[0].target, "file:///Users/libraz/Documents/a%20b/f.xlsx");
    EXPECT_EQ(unset[0].display, "");
  }

  ASSERT_EQ(Set(wb(), (data + "/loc_probe3/p q").c_str(), "Book.xlsx"), 0);
  for (const std::int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    const std::vector<SavedLink> saved = SaveAndReadLinks(wb(), format);
    ASSERT_EQ(saved.size(), 3U);
    EXPECT_EQ(saved[0].target, "../../../../../../Documents/a%20b/f.xlsx");
    EXPECT_EQ(saved[0].display, "../../../../../../Documents/a b/f.xlsx");
    EXPECT_EQ(saved[1].target, "file:///private/tmp/x.xlsx");
    EXPECT_EQ(saved[1].display, "");
    EXPECT_EQ(saved[2].target, "sib.xlsx");
    EXPECT_EQ(saved[2].display, "sib.xlsx");
  }

  // Save As: a new folder changes the saved relative form.
  ASSERT_EQ(Set(wb(), (data + "/loc_probe4/r").c_str(), "Book.xlsx"), 0);
  const std::vector<SavedLink> moved = SaveAndReadLinks(wb(), FM_WORKBOOK_FORMAT_XLSX);
  ASSERT_EQ(moved.size(), 3U);
  EXPECT_EQ(moved[2].target, "../../loc_probe3/p%20q/sib.xlsx");
  EXPECT_EQ(moved[2].display, "../../loc_probe3/p q/sib.xlsx");
}

}  // namespace
