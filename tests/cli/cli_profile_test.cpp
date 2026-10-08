#include <array>
#include <fstream>
#include <string>
#include <string_view>

#include "cli_test_support.h"

namespace {

struct ProfileCase {
  const char* id;
  const char* expected;
};

constexpr std::array<ProfileCase, 4> kProfiles = {{
    {"win-365-ja_JP", "¥1,235"},
    {"mac-365-ja_JP", "¥1,235"},
    {"win-365-en_US", "$1,234.50"},
    {"mac-365-en_US", "$1,234.50"},
}};

bool write_profile_fixture(const std::string& path) {
  fm_workbook_t* workbook = nullptr;
  if (fm_workbook_create(&workbook) != 0) {
    return false;
  }
  const bool formula_set = fm_workbook_set_formula(workbook, 0, 0, 0, "=DOLLAR(1234.5)") == 0;
  std::uint8_t* bytes = nullptr;
  std::size_t length = 0;
  const bool saved = formula_set && fm_workbook_save(workbook, &bytes, &length) == 0;
  bool wrote = false;
  if (saved) {
    std::ofstream output(path, std::ios::binary | std::ios::trunc);
    if (output) {
      output.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(length));
      wrote = output.good();
    }
  }
  fm_buffer_free(bytes);
  fm_workbook_destroy(workbook);
  return saved && wrote;
}

std::string first_cell_text(const std::string& path) {
  std::ifstream input(path, std::ios::binary);
  if (!input) {
    return {};
  }
  const std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
  fm_workbook_t* workbook = nullptr;
  if (fm_workbook_load(bytes.data(), bytes.size(), &workbook) != 0) {
    return {};
  }
  std::size_t cell_count = 0;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  fm_value_t value{};
  const bool read = fm_workbook_cell_count(workbook, 0, &cell_count) == 0 && cell_count != 0 &&
                    fm_workbook_cell_at(workbook, 0, 0, &row, &col, nullptr, &value) == 0;
  const std::string text = read && value.kind == FM_VAL_TEXT && value.u.text != nullptr ? value.u.text : std::string{};
  fm_workbook_destroy(workbook);
  return text;
}

}  // namespace

using namespace formulon::cli_test;

TEST(FormulonCli, ProfileFlagSelectsCurrencyForEval) {
  for (const ProfileCase& profile : kProfiles) {
    const CliRun result = run_cli({"eval", "--profile", profile.id, "=DOLLAR(1234.5)"});
    EXPECT_EQ(result.exit_code, 0) << profile.id << " stderr=" << result.stderr_text;
    EXPECT_EQ(result.stdout_text, std::string(profile.expected) + '\n') << profile.id;
  }
}

TEST(FormulonCli, ProfileFlagOmittedUsesEnglishDefault) {
  const CliRun result = run_cli({"eval", "=DOLLAR(1234.5)"});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;
  EXPECT_EQ(result.stdout_text, "$1,234.50\n");
}

TEST(FormulonCli, ProfileFlagAppliesAfterRecalcLoad) {
  const std::string input = temp_path("profile_recalc_input.xlsx");
  PathGuard input_guard(input);
  ASSERT_TRUE(write_profile_fixture(input));

  for (const ProfileCase& profile : kProfiles) {
    const std::string output = temp_path(std::string("profile_recalc_output_") + profile.id + ".xlsx");
    PathGuard output_guard(output);
    const CliRun result = run_cli({"recalc", "--profile", profile.id, input, "-o", output, "--quiet"});
    ASSERT_EQ(result.exit_code, 0) << profile.id << " stderr=" << result.stderr_text;
    EXPECT_EQ(first_cell_text(output), profile.expected) << profile.id;
  }
}

TEST(FormulonCli, ProfileFlagAppliesAfterDumpLoad) {
  const std::string input = temp_path("profile_dump_input.xlsx");
  PathGuard input_guard(input);
  ASSERT_TRUE(write_profile_fixture(input));

  for (const ProfileCase& profile : kProfiles) {
    const CliRun result = run_cli({"dump", "--profile", profile.id, "--values", input});
    ASSERT_EQ(result.exit_code, 0) << profile.id << " stderr=" << result.stderr_text;
    EXPECT_NE(result.stdout_text.find(profile.expected), std::string::npos) << profile.id;
  }
}

TEST(FormulonCli, ProfileFlagIsAcceptedByPaginate) {
  const std::string input = temp_path("profile_paginate_input.xlsx");
  PathGuard input_guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  for (const ProfileCase& profile : kProfiles) {
    const CliRun result = run_cli({"paginate", "--profile", profile.id, input});
    EXPECT_EQ(result.exit_code, 0) << profile.id << " stderr=" << result.stderr_text;
  }
}

TEST(FormulonCli, ProfileFlagRejectsInvalidAndMissingValue) {
  for (const char* command : {"eval", "recalc", "dump", "paginate"}) {
    const CliRun missing = run_cli({command, "--profile"});
    EXPECT_EQ(missing.exit_code, 64) << command << " stderr=" << missing.stderr_text;
    EXPECT_NE(missing.stderr_text.find("--profile requires a value"), std::string::npos) << command;

    const CliRun invalid = run_cli({command, "--profile", "unknown-profile"});
    EXPECT_EQ(invalid.exit_code, 64) << command << " stderr=" << invalid.stderr_text;
    EXPECT_NE(invalid.stderr_text.find("invalid Excel profile id"), std::string::npos) << command;
  }
}

TEST(FormulonCli, ProfileFlagStopsAtTerminator) {
  const CliRun result = run_cli({"eval", "--", "--profile"});
  EXPECT_EQ(result.exit_code, 0) << result.stderr_text;
  EXPECT_FALSE(result.stdout_text.empty());
}

TEST(FormulonCli, LastProfileFlagWins) {
  const CliRun result =
      run_cli({"eval", "--profile", "win-365-en_US", "--profile", "win-365-ja_JP", "=DOLLAR(1234.5)"});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;
  EXPECT_EQ(result.stdout_text, "¥1,235\n");
}

TEST(FormulonCli, ProfileFlagAppearsInTopAndSubcommandHelp) {
  const std::array<const char*, 4> commands = {"eval", "recalc", "dump", "paginate"};
  const std::array<std::string_view, 4> ids = {"win-365-ja_JP", "mac-365-ja_JP", "win-365-en_US", "mac-365-en_US"};

  const CliRun top = run_cli({"--help"});
  ASSERT_EQ(top.exit_code, 0) << top.stderr_text;
  for (const std::string_view id : ids) {
    EXPECT_NE(top.stdout_text.find(id), std::string::npos) << id;
  }
  EXPECT_NE(top.stdout_text.find("--profile <id>"), std::string::npos);
  EXPECT_NE(top.stdout_text.find("win-365-en_US"), std::string::npos);

  for (const char* command : commands) {
    const CliRun help = run_cli({command, "--help"});
    ASSERT_EQ(help.exit_code, 0) << command << " stderr=" << help.stderr_text;
    EXPECT_NE(help.stdout_text.find("--profile <id>"), std::string::npos) << command;
    for (const std::string_view id : ids) {
      EXPECT_NE(help.stdout_text.find(id), std::string::npos) << command << " " << id;
    }
    EXPECT_NE(help.stdout_text.find("win-365-en_US"), std::string::npos) << command;
  }
}
