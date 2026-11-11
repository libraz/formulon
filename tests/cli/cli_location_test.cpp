// The CLI sets the workbook location from the paths it reads and writes: an
// input path made absolute against the current directory, the `-o` path before
// a save (Excel's Save As), and none for stdin.
#include <sys/stat.h>
#include <unistd.h>

#include <climits>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <string>
#include <vector>

#include "cli_test_support.h"
#include "support/roundtrip_symmetry.h"

using namespace formulon::cli_test;

namespace {

struct Location {
  bool present = false;
  std::string directory;
  std::string file_name;
};

Location ReadLocation(const fm_workbook_t* wb) {
  Location out;
  std::int32_t present = 0;
  fm_workbook_location_t record = FM_STRUCT_INIT(fm_workbook_location_t);
  EXPECT_EQ(fm_workbook_get_location(wb, &present, &record), 0);
  out.present = present != 0;
  out.directory = record.directory != nullptr ? record.directory : "";
  out.file_name = record.file_name != nullptr ? record.file_name : "";
  return out;
}

std::string Directory(const std::string& path) {
  return path.substr(0, path.find_last_of('/'));
}

std::string Base(const std::string& path) {
  return path.substr(path.find_last_of('/') + 1);
}

std::vector<std::uint8_t> ReadAll(const std::string& path) {
  std::ifstream in(path, std::ios::binary);
  return std::vector<std::uint8_t>((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
}

// A1 = CELL("filename",A1); B1 links to `link_target`.
bool WriteLocationFixture(const std::string& path, const std::string& link_target) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create_default(&wb) != 0) {
    return false;
  }
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  hl.rect = fm_cell_rect_t{0, 1, 0, 1};
  hl.target = link_target.c_str();
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  const bool ok = fm_sheet_set_formula(wb, 0, 0, 0, "=CELL(\"filename\",A1)") == 0 &&
                  fm_sheet_set_hyperlink(wb, 0, &hl) == 0 &&
                  fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len) == 0 &&
                  write_bytes(path, std::string_view(reinterpret_cast<const char*>(bytes), len));
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(wb);
  return ok;
}

struct Output {
  std::string a1;
  std::string link_target;
  std::string workbook_xml;
};

Output ReadOutput(const std::string& path) {
  Output out;
  const std::vector<std::uint8_t> bytes = ReadAll(path);
  EXPECT_FALSE(bytes.empty());
  fm_workbook_t* wb = nullptr;
  EXPECT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  if (wb == nullptr) {
    return out;
  }
  // A loaded formula is pending, and a read in automatic mode would recalculate
  // it here, where the book has no location; manual mode reads the saved value.
  EXPECT_EQ(fm_workbook_set_calc_mode(wb, FM_CALC_MODE_MANUAL), 0);
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, 0, 0, &v), 0);
  out.a1 = v.kind == FM_VALUE_KIND_TEXT ? v.u.text : "<not text>";
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  EXPECT_EQ(fm_sheet_get_hyperlink_at(wb, 0, 0, &hl), 0);
  out.link_target = hl.target != nullptr ? hl.target : "";
  fm_workbook_destroy(wb);
  EXPECT_TRUE(formulon::test::extract_part(formulon::test::span_of(bytes), "xl/workbook.xml", &out.workbook_xml));
  return out;
}

}  // namespace

TEST(FormulonCliLocation, RelativeInputPathIsMadeAbsolute) {
  char cwd[PATH_MAX + 1] = {};
  ASSERT_NE(::getcwd(cwd, sizeof(cwd)), nullptr);
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create_default(&wb), 0);
  formulon::cli::set_location_from_path(wb, "sub/./x/../Book.xlsx");
  const Location got = ReadLocation(wb);
  EXPECT_TRUE(got.present);
  EXPECT_EQ(got.directory, std::string(cwd) + "/sub");
  EXPECT_EQ(got.file_name, "Book.xlsx");

  // A name the setter refuses leaves the workbook without a location.
  formulon::cli::set_location_from_path(wb, "/tmp/notes.txt");
  EXPECT_FALSE(ReadLocation(wb).present);
  fm_workbook_destroy(wb);
}

TEST(FormulonCliLocation, AbsoluteInputPathDrivesCellFilename) {
  const std::string in = temp_path("loc_abs_in.xlsx");
  const std::string out = temp_path("loc_abs_out.xlsx");
  PathGuard in_guard(in);
  PathGuard out_guard(out);
  ASSERT_TRUE(WriteLocationFixture(in, "https://example.com/"));

  const CliRun r = run_cli({"recalc", "--profile", "mac-365-ja_JP", in, "-o", out, "--quiet"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  // The recalculation ran with the input's location; Save As does not recalculate.
  EXPECT_EQ(ReadOutput(out).a1, Directory(in) + "/[" + Base(in) + "]Sheet1");
}

TEST(FormulonCliLocation, OutputPathIsTheLocationTheSaveWritesAgainst) {
  const std::string in = temp_path("loc_out_in.xlsx");
  const std::string out_dir = temp_path("loc_out_dir");
  const std::string out = out_dir + "/out.xlsx";
  const std::string sibling = Directory(in) + "/sib.xlsx";
  PathGuard in_guard(in);
  ASSERT_EQ(::mkdir(out_dir.c_str(), 0700), 0);
  DirGuard dir_guard(out_dir);
  ASSERT_TRUE(WriteLocationFixture(in, sibling));

  const CliRun r = run_cli({"recalc", "--profile", "mac-365-ja_JP", in, "-o", out, "--quiet"});
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  const Output saved = ReadOutput(out);
  EXPECT_EQ(saved.link_target, "../sib.xlsx");
  EXPECT_NE(saved.workbook_xml.find("<x15ac:absPath url=\"" + out_dir + "/\""), std::string::npos)
      << saved.workbook_xml;
}

TEST(FormulonCliLocation, StdinInputHasNoLocation) {
  const std::string in = temp_path("loc_stdin_in.xlsx");
  const std::string out = temp_path("loc_stdin_out.xlsx");
  PathGuard in_guard(in);
  PathGuard out_guard(out);
  ASSERT_TRUE(WriteLocationFixture(in, "https://example.com/"));

  const CliRun r = run_cli({"recalc", "--profile", "mac-365-ja_JP", "-", "-o", out, "--quiet"}, false, false, in);
  ASSERT_EQ(r.exit_code, 0) << r.stderr_text;
  EXPECT_EQ(ReadOutput(out).a1, "");
}
