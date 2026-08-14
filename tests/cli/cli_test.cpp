//
// `formulon_cli` end-to-end tests. Each test spawns the binary at the
// CMake-injected `FORMULON_CLI_PATH` and asserts on (exit_code, stdout,
// stderr). The harness uses `popen` because it is the simplest
// portable way to run a subprocess and capture output without pulling
// in a third-party library; `popen` does not give us stderr capture
// natively, so we redirect stderr to stdout via the shell. Tests that
// need to distinguish the streams use the explicit `--quiet` flag or
// inspect output ordering.

#include "cli/cli.h"

#include <fcntl.h>
#include <gtest/gtest.h>
#include <sys/stat.h>
#include <unistd.h>

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <fstream>
#include <ios>
#include <iterator>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "io/defined_names.h"
#include "io/format_detect.h"
#include "support/ooxml_package_fixture.h"
#include "workbook.h"

#ifndef FORMULON_CLI_PATH
#error "FORMULON_CLI_PATH must be defined by the build system"
#endif

namespace {

// Output of one CLI subprocess invocation. `stdout_text` is captured
// verbatim; `stderr_text` is only populated when `redirect_stderr` was
// requested. `exit_code` is the wait-status's low byte.
struct CliRun {
  int exit_code = -1;
  std::string stdout_text;
  std::string stderr_text;
};

// Quotes `s` in single quotes for `/bin/sh`, doubling any embedded
// single quotes via `'\''`. Sufficient for the file paths and short
// flag values the tests pass; not a full shell quoter.
std::string sh_quote(const std::string& s) {
  std::string out;
  out.reserve(s.size() + 2);
  out.push_back('\'');
  for (char c : s) {
    if (c == '\'') {
      out.append("'\\''");
    } else {
      out.push_back(c);
    }
  }
  out.push_back('\'');
  return out;
}

// Spawns the CLI with `args` (each argument is sh-quoted before being
// joined into a single command line). When `merge_streams` is true,
// stderr is folded into stdout via `2>&1`; otherwise stderr is
// captured to a temp file via `2>` and read back. The temp-file path
// avoids races with shell escaping.
CliRun run_cli(const std::vector<std::string>& args, bool merge_streams = false, bool close_stdout = false) {
  CliRun out;
  std::string cmd = sh_quote(FORMULON_CLI_PATH);
  for (const auto& a : args) {
    cmd.push_back(' ');
    cmd.append(sh_quote(a));
  }
  // Capture stderr separately to a temp file so we can assert on it
  // without ambiguity. The shell's `>` redirect is portable on macOS
  // and Linux. We append rather than truncate so the temp file is
  // self-contained per call.
  char tmpl[] = "/tmp/fm_cli_stderr_XXXXXX";
  const int fd = mkstemp(tmpl);
  if (fd < 0) {
    out.exit_code = -2;
    return out;
  }
  ::close(fd);
  if (close_stdout) {
    cmd.append(" 1>&- 2>");
    cmd.append(sh_quote(tmpl));
  } else if (merge_streams) {
    cmd.append(" 2>&1");
  } else {
    cmd.append(" 2>");
    cmd.append(sh_quote(tmpl));
  }

  FILE* pipe = popen(cmd.c_str(), "r");
  if (pipe == nullptr) {
    out.exit_code = -3;
    std::remove(tmpl);
    return out;
  }
  char buf[4096];
  while (std::fgets(buf, sizeof(buf), pipe) != nullptr) {
    out.stdout_text.append(buf);
  }
  const int wait_rc = pclose(pipe);
  // `pclose` returns the wait status; on POSIX, `WEXITSTATUS` is in
  // the upper byte. We approximate it with `(wait_rc >> 8) & 0xff`,
  // which matches `/bin/sh` semantics.
  if (wait_rc < 0) {
    out.exit_code = -4;
  } else {
    out.exit_code = (wait_rc >> 8) & 0xff;
  }

  if (!merge_streams) {
    std::ifstream in(tmpl);
    if (in) {
      std::string line;
      while (std::getline(in, line)) {
        out.stderr_text.append(line);
        out.stderr_text.push_back('\n');
      }
    }
  }
  std::remove(tmpl);
  return out;
}

// Small RAII guard: removes the file at `path` on destruction. Used to
// keep temp xlsx files from leaking when a test asserts mid-flow.
struct PathGuard {
  std::string path;
  explicit PathGuard(std::string p) : path(std::move(p)) {}
  PathGuard(const PathGuard&) = delete;
  PathGuard& operator=(const PathGuard&) = delete;
  ~PathGuard() {
    if (!path.empty()) {
      std::remove(path.c_str());
    }
  }
};

// Builds a tiny `.xlsx` workbook in memory via the C ABI and writes
// it to `path`. Used by `recalc` and `dump` tests that need a fixture
// on disk. Returns `true` on success.
bool write_fixture_workbook(const std::string& path, std::string_view extra_text = {}, bool iterative = false,
                            int32_t max_iterations = 100, double max_change = 0.001) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create(&wb) != 0) {
    return false;
  }
  // A1 = 7 (literal), B1 = =A1+1 (formula).
  if (fm_workbook_set_number(wb, 0, 0, 0, 7.0) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  if (fm_workbook_set_formula(wb, 0, 0, 1, "=A1+1") != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  if (iterative && fm_workbook_set_iterative(wb, 1, max_iterations, max_change) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  if (fm_workbook_recalc(wb) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  if (!extra_text.empty() && fm_workbook_set_text(wb, 0, 1, 0, std::string(extra_text).c_str()) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  if (fm_workbook_save(wb, &bytes, &len) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  if (!out) {
    fm_buffer_free(bytes);
    fm_workbook_destroy(wb);
    return false;
  }
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  out.close();
  fm_buffer_free(bytes);
  fm_workbook_destroy(wb);
  return out.good();
}

// A wide, pure formula layer makes the CLI's parallel route observable
// without relying on volatile functions or timing. The caller receives an
// ordinary .xlsx that the command must load, recalculate, save, and whose
// last lane we inspect after reopening.
bool write_wide_recalc_fixture(const std::string& path) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create(&wb) != 0) {
    return false;
  }
  constexpr std::uint32_t kRows = 48U;
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const std::string formula = "=A" + std::to_string(row + 1U) + "*2+1";
    if (fm_workbook_set_number(wb, 0U, row, 0U, static_cast<double>(row + 1U)) != 0 ||
        fm_workbook_set_formula(wb, 0U, row, 1U, formula.c_str()) != 0) {
      fm_workbook_destroy(wb);
      return false;
    }
  }
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0U;
  if (fm_workbook_save(wb, &bytes, &len) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  if (!out) {
    fm_buffer_free(bytes);
    fm_workbook_destroy(wb);
    return false;
  }
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  out.close();
  fm_buffer_free(bytes);
  fm_workbook_destroy(wb);
  return out.good();
}

bool write_out_of_range_defined_names_fixture(const std::string& path) {
  formulon::Workbook wb = formulon::Workbook::create();
  for (std::size_t i = 1; i < 7; ++i) {
    wb.add_sheet("Sheet" + std::to_string(i + 1));
  }

  std::vector<formulon::io::DefinedName> names;
  names.push_back(formulon::io::DefinedName{"Before", "=1", -1, false, ""});
  names.push_back(formulon::io::DefinedName{"Bad", "=Sheet1!$A$1", 99, false, ""});
  names.push_back(formulon::io::DefinedName{"After", "=2", -1, false, ""});
  wb.set_defined_names(std::move(names));

  auto saved = wb.save();
  if (!saved) {
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  if (!out) {
    return false;
  }
  const std::vector<std::uint8_t>& bytes = saved.value();
  out.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
  out.close();
  return out.good();
}

bool write_lossy_xlsx_fixture(const std::string& path) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create(&wb) != 0) {
    return false;
  }
  if (fm_workbook_set_formula(wb, 0, 0, 0, "=@A1:A10") != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  fm_hyperlink hyperlink{};
  hyperlink.target = "https://example.com";
  if (fm_sheet_add_hyperlink(wb, 0, hyperlink) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  fm_merge_range range{0, 0, 0, 0};
  fm_data_validation validation{};
  validation.ranges = &range;
  validation.range_count = 1;
  validation.type = 3;
  validation.formula1 = "$A$1:$A$2";
  if (fm_sheet_add_validation(wb, 0, validation) != 0 ||
      fm_sheet_set_auto_filter_xml(wb, 0, "<autoFilter ref=\"A1:B2\"/>") != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  if (fm_workbook_save(wb, &bytes, &len) != 0) {
    fm_workbook_destroy(wb);
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  if (!out) {
    fm_buffer_free(bytes);
    fm_workbook_destroy(wb);
    return false;
  }
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  out.close();
  fm_buffer_free(bytes);
  fm_workbook_destroy(wb);
  return out.good();
}

std::vector<std::uint8_t> append_empty_zip_entry(const std::vector<std::uint8_t>& bytes, std::string_view name) {
  const std::uint8_t sig[] = {0x50, 0x4b, 0x05, 0x06};
  const auto it = std::find_end(bytes.begin(), bytes.end(), std::begin(sig), std::end(sig));
  if (it == bytes.end())
    return {};
  const std::size_t eocd = static_cast<std::size_t>(it - bytes.begin());
  const auto read16 = [&](std::size_t p) { return static_cast<std::uint16_t>(bytes[p] | (bytes[p + 1] << 8U)); };
  const auto read32 = [&](std::size_t p) {
    return static_cast<std::uint32_t>(bytes[p] | (bytes[p + 1] << 8U) | (bytes[p + 2] << 16U) | (bytes[p + 3] << 24U));
  };
  const std::uint16_t count = read16(eocd + 10U);
  const std::uint32_t central_size = read32(eocd + 12U);
  const std::uint32_t central_offset = read32(eocd + 16U);
  const auto put16 = [](std::vector<std::uint8_t>& out, std::uint16_t x) {
    out.push_back(static_cast<std::uint8_t>(x));
    out.push_back(static_cast<std::uint8_t>(x >> 8U));
  };
  const auto put32 = [](std::vector<std::uint8_t>& out, std::uint32_t x) {
    out.push_back(static_cast<std::uint8_t>(x));
    out.push_back(static_cast<std::uint8_t>(x >> 8U));
    out.push_back(static_cast<std::uint8_t>(x >> 16U));
    out.push_back(static_cast<std::uint8_t>(x >> 24U));
  };
  const std::string n(name);
  std::vector<std::uint8_t> local;
  put32(local, 0x04034b50U);
  put16(local, 20);
  put16(local, 0);
  put16(local, 0);
  put16(local, 0);
  put16(local, 0);
  put32(local, 0);
  put32(local, 0);
  put32(local, 0);
  put16(local, static_cast<std::uint16_t>(n.size()));
  put16(local, 0);
  local.insert(local.end(), n.begin(), n.end());
  std::vector<std::uint8_t> central;
  put32(central, 0x02014b50U);
  put16(central, 20);
  put16(central, 20);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put32(central, 0);
  put32(central, 0);
  put32(central, 0);
  put16(central, static_cast<std::uint16_t>(n.size()));
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put16(central, 0);
  put32(central, 0);
  put32(central, central_offset);
  central.insert(central.end(), n.begin(), n.end());
  std::vector<std::uint8_t> out(bytes.begin(), bytes.begin() + central_offset);
  out.insert(out.end(), local.begin(), local.end());
  out.insert(out.end(), bytes.begin() + central_offset, bytes.begin() + central_offset + central_size);
  out.insert(out.end(), central.begin(), central.end());
  put32(out, 0x06054b50U);
  put16(out, 0);
  put16(out, 0);
  put16(out, count + 1U);
  put16(out, count + 1U);
  put32(out, central_size + static_cast<std::uint32_t>(central.size()));
  put32(out, central_offset + static_cast<std::uint32_t>(local.size()));
  put16(out, 0);
  return out;
}

// An `.xlsx` whose sheet carries one unusable reference per overlay kind and
// whose workbook part declares a content type the reader does not recognise.
// The engine's writer never produces either, so the package is assembled by
// hand.
bool write_lossy_ooxml_fixture(const std::string& path) {
  const std::string content_types = formulon::test::OoxmlContentTypes("application/vnd.formulon-test.unknown+xml");
  const std::vector<std::uint8_t> bytes = formulon::test::BuildOoxmlPackage(
      content_types,
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/>"
      "<mergeCells><mergeCell ref=\"nope\"/></mergeCells>"
      "<conditionalFormatting><cfRule type=\"expression\" priority=\"1\"/></conditionalFormatting>"
      "</worksheet>");
  if (bytes.empty()) {
    return false;
  }
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  out.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
  return out.good();
}

bool write_dropped_xlsb_fixture(const std::string& path) {
  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create(&wb) != 0)
    return false;
  std::uint8_t* raw = nullptr;
  std::size_t len = 0;
  const bool saved = fm_workbook_save_as(wb, FM_WORKBOOK_FORMAT_XLSB, &raw, &len) == 0;
  fm_workbook_destroy(wb);
  if (!saved)
    return false;
  std::vector<std::uint8_t> fixture =
      append_empty_zip_entry(std::vector<std::uint8_t>(raw, raw + len), "xl/dropped.unknown");
  fm_buffer_free(raw);
  if (fixture.empty())
    return false;
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  out.write(reinterpret_cast<const char*>(fixture.data()), static_cast<std::streamsize>(fixture.size()));
  return out.good();
}

}  // namespace

// ---------------------------------------------------------------------------
// Top-level usage / version
// ---------------------------------------------------------------------------

TEST(FormulonCli, VersionPrintsNonEmpty) {
  CliRun r = run_cli({"--version"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_GT(r.stdout_text.size(), 0U);
}

TEST(FormulonCli, HelpExitsZero) {
  CliRun r = run_cli({"--help"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_NE(r.stdout_text.find("Usage"), std::string::npos);
}

TEST(FormulonCli, VersionIsAcceptedBySubcommandsLikeHelp) {
  // The usage banner lists `--version` under the same "Common options"
  // heading as `-h`, so every subcommand must accept it and print the
  // same line the top-level flag prints.
  CliRun top = run_cli({"--version"});
  ASSERT_EQ(top.exit_code, 0);
  ASSERT_FALSE(top.stdout_text.empty());

  for (const char* subcommand : {"eval", "recalc", "dump", "paginate"}) {
    CliRun r = run_cli({subcommand, "--version"});
    EXPECT_EQ(r.exit_code, 0) << subcommand << " stderr=" << r.stderr_text;
    EXPECT_EQ(r.stdout_text, top.stdout_text) << subcommand;
  }
}

TEST(FormulonCli, RecalcHelpDocumentsActualSuccessStatus) {
  CliRun r = run_cli({"recalc", "--help"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_NE(r.stdout_text.find("formulon: recalc: ok, wrote M bytes to 'OUT'"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("--threads N"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("recalc remains serial"), std::string::npos);
}

TEST(FormulonCli, PaginatePrintsResolvedGeometry) {
  const std::string path = "/tmp/fm_cli_paginate.xlsx";
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));

  CliRun r = run_cli({"paginate", path});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "sheet=0\npages=1\nprint_area=\nhorizontal_breaks=\nvertical_breaks=\n");
}

TEST(FormulonCli, PaginateWriteFailureReturnsOutputError) {
  const std::string path = "/tmp/fm_cli_paginate_write_failure.xlsx";
  PathGuard guard(path);
  ASSERT_TRUE(write_fixture_workbook(path));

  CliRun r = run_cli({"paginate", path}, /*merge_streams=*/false, /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
}

TEST(FormulonCli, NoArgsExits64) {
  CliRun r = run_cli({});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, UnknownCommandExits64) {
  CliRun r = run_cli({"frobnicate"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, ExitCodesNeverEncodeLargeInternalStatusValues) {
  EXPECT_EQ(formulon::cli::exit_code_for_status(0), 0);
  EXPECT_EQ(formulon::cli::exit_code_for_status(formulon::cli::kExitUsage), 64);
  EXPECT_EQ(formulon::cli::exit_code_for_status(128), 1);
  EXPECT_EQ(formulon::cli::exit_code_for_status(7000), 1);
  EXPECT_EQ(formulon::cli::exit_code_for_status(8000), 1);
}

// ---------------------------------------------------------------------------
// eval
// ---------------------------------------------------------------------------

TEST(FormulonCli, EvalSumLiterals) {
  CliRun r = run_cli({"eval", "=SUM(1,2,3)"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "6\n");
}

TEST(FormulonCli, EvalIfTrue) {
  CliRun r = run_cli({"eval", "=IF(TRUE,1,2)"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "1\n");
}

TEST(FormulonCli, EvalConcatTextNoQuoting) {
  CliRun r = run_cli({"eval", "=CONCAT(\"hello\",\" \",\"world\")"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "hello world\n");
}

TEST(FormulonCli, EvalCellLevelErrorIsStillExitZero) {
  // The ad-hoc evaluator reads A1 without first writing the expression to
  // A1, so it sees the empty cell rather than making a self-reference cycle.
  CliRun r = run_cli({"eval", "=A1+1"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "1\n");
}

TEST(FormulonCli, EvalSyntaxErrorFailsInsteadOfReturningNameError) {
  CliRun r = run_cli({"eval", "=1+"});
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_TRUE(r.stdout_text.empty());
  EXPECT_NE(r.stderr_text.find("invalid formula syntax"), std::string::npos);
}

TEST(FormulonCli, EvalUnknownFunctionRemainsCellLevelNameError) {
  CliRun r = run_cli({"eval", "=NOTAREALFUNC(1,2)"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "#NAME?\n");
}

TEST(FormulonCli, EvalDynamicArrayPrintsWholeGrid) {
  CliRun r = run_cli({"eval", "=SEQUENCE(2,3)"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "1\t2\t3\n4\t5\t6\n");
}

TEST(FormulonCli, EvalDynamicArrayJsonPrintsNestedArrays) {
  CliRun r = run_cli({"eval", "--json", "=SEQUENCE(2,2)"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text,
            "[[{\"kind\":\"number\",\"value\":1},{\"kind\":\"number\",\"value\":2}],"
            "[{\"kind\":\"number\",\"value\":3},{\"kind\":\"number\",\"value\":4}]]\n");
}

TEST(FormulonCli, EvalDivByZeroSurfacesAsExcelErrorString) {
  CliRun r = run_cli({"eval", "=1/0"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "#DIV/0!\n");
}

TEST(FormulonCli, EvalJsonOutputShape) {
  CliRun r = run_cli({"eval", "--json", "=SUM(1,2,3)"});
  EXPECT_EQ(r.exit_code, 0);
  // Compact form: `{"kind":"number","value":6}`.
  EXPECT_NE(r.stdout_text.find("\"kind\":\"number\""), std::string::npos) << "stdout=" << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("\"value\":6"), std::string::npos) << "stdout=" << r.stdout_text;
}

TEST(FormulonCli, EvalMissingFormulaExits64) {
  CliRun r = run_cli({"eval"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, EvalUnknownFlagExits64) {
  CliRun r = run_cli({"eval", "--bogus"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, EvalNegativeFormulaWithoutTerminatorExits64) {
  CliRun r = run_cli({"eval", "-1+2"});
  EXPECT_EQ(r.exit_code, 64);
  EXPECT_TRUE(r.stdout_text.empty());
}

TEST(FormulonCli, EvalOptionTerminatorAcceptsNegativeFormula) {
  CliRun r = run_cli({"eval", "--", "-1+2"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "1\n");
}

TEST(FormulonCli, EvalRepeatRunsMultipleTimes) {
  CliRun r = run_cli({"eval", "--repeat", "3", "=SUM(1,2,3)"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "6\n");
  // The timing line goes to stderr.
  EXPECT_NE(r.stderr_text.find("3 iterations"), std::string::npos);
}

TEST(FormulonCli, EvalWriteFailureReturnsOutputError) {
  CliRun r = run_cli({"eval", "=SUM(1,2,3)"}, /*merge_streams=*/false, /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
}

TEST(FormulonCli, EvalRepeatWriteFailureSkipsTimingOutput) {
  CliRun r = run_cli({"eval", "--repeat", "3", "=SUM(1,2,3)"}, /*merge_streams=*/false,
                     /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
  EXPECT_EQ(r.stderr_text.find("iterations"), std::string::npos);
}

TEST(FormulonCli, EvalRepeatOverflowExits64) {
  CliRun r = run_cli({"eval", "--repeat", "999999999999999999999999999999", "=1"});
  EXPECT_EQ(r.exit_code, 64);
  EXPECT_NE(r.stderr_text.find("positive integer"), std::string::npos);
}

// Each --repeat pass executes the ad-hoc evaluator, so a high repeat count
// still yields the result without mutating the temporary workbook.
TEST(FormulonCli, EvalRepeatHighCountStillCorrect) {
  CliRun r = run_cli({"eval", "--repeat", "100", "=SUM(1,2,3)"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "6\n");
  EXPECT_NE(r.stderr_text.find("100 iterations"), std::string::npos);
}

// ---------------------------------------------------------------------------
// recalc
// ---------------------------------------------------------------------------

TEST(FormulonCli, RecalcRoundTripsFormulae) {
  // Build a fixture in /tmp.
  std::string in = "/tmp/fm_cli_in.xlsx";
  std::string out_path = "/tmp/fm_cli_out.xlsx";
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  ASSERT_TRUE(write_fixture_workbook(in));

  CliRun r = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  // Load the output back through the C ABI and assert B1 still
  // recalculates correctly.
  std::ifstream f(out_path, std::ios::binary);
  ASSERT_TRUE(f);
  f.seekg(0, std::ios::end);
  const std::streamsize size = f.tellg();
  ASSERT_GT(size, 0);
  f.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(static_cast<std::size_t>(size));
  f.read(reinterpret_cast<char*>(bytes.data()), size);
  ASSERT_TRUE(f);

  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0, 0, 1, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 8.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcWithoutThreadsKeepsSerialStatusShape) {
  const std::string input = "/tmp/fm_cli_default_recalc_input.xlsx";
  const std::string output = "/tmp/fm_cli_default_recalc_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", input, "-o", output});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;
  EXPECT_EQ(result.stderr_text.find("threads="), std::string::npos);
  EXPECT_NE(result.stderr_text.find("formulon: recalc: ok, wrote "), std::string::npos);
  EXPECT_NE(result.stderr_text.find(" bytes to '" + output + "'"), std::string::npos);
}

TEST(FormulonCli, RecalcThreadsUsesParallelSchedulerAndSavesWideDag) {
  const std::string input = "/tmp/fm_cli_parallel_recalc_input.xlsx";
  const std::string output = "/tmp/fm_cli_parallel_recalc_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_wide_recalc_fixture(input));

  const CliRun result = run_cli({"recalc", "--threads", "4", input, "-o", output});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;
  EXPECT_TRUE(result.stdout_text.empty());
  EXPECT_NE(result.stderr_text.find("threads=4"), std::string::npos);
  EXPECT_NE(result.stderr_text.find("cells_evaluated=48"), std::string::npos);
  EXPECT_NE(result.stderr_text.find("sccs_processed="), std::string::npos);
  EXPECT_NE(result.stderr_text.find("parallel_steps="), std::string::npos);
  EXPECT_NE(result.stderr_text.find("worker_threads_started="), std::string::npos);
  EXPECT_NE(result.stderr_text.find("worker_threads_used="), std::string::npos);

  std::ifstream in(output, std::ios::binary);
  ASSERT_TRUE(in);
  const std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
  ASSERT_FALSE(bytes.empty());
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0U, 47U, 1U, &value), 0);
  EXPECT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 97.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcThreadsOneStartsNoWorkers) {
  const std::string input = "/tmp/fm_cli_serial_threads_input.xlsx";
  const std::string output = "/tmp/fm_cli_serial_threads_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", "--threads", "1", input, "-o", output});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;
  EXPECT_NE(result.stderr_text.find("threads=1"), std::string::npos);
  EXPECT_NE(result.stderr_text.find("worker_threads_started=0"), std::string::npos);
  EXPECT_NE(result.stderr_text.find("worker_threads_used=0"), std::string::npos);
}

TEST(FormulonCli, RecalcThreadsRejectsInvalidValuesBeforeWriting) {
  const std::string input = "/tmp/fm_cli_invalid_threads_input.xlsx";
  const std::string output = "/tmp/fm_cli_invalid_threads_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_fixture_workbook(input));
  std::remove(output.c_str());

  for (const char* value : {"9", "-1", "abc", "4294967296", "--"}) {
    const CliRun result = run_cli({"recalc", "--threads", value, input, "-o", output});
    EXPECT_EQ(result.exit_code, 64) << value << " stderr=" << result.stderr_text;
    EXPECT_NE(result.stderr_text.find("--threads"), std::string::npos);
    std::ifstream absent(output, std::ios::binary);
    EXPECT_FALSE(absent.good()) << "invalid value wrote output: " << value;
  }
}

TEST(FormulonCli, RecalcLossWarningsAreNonfatalAndNotSuppressedByQuiet) {
  const std::string input = "/tmp/fm_cli_lossy_input.xlsx";
  const std::string output = "/tmp/fm_cli_lossy_output.xlsb";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_lossy_xlsx_fixture(input));

  CliRun r = run_cli({"recalc", "--quiet", input, "-o", output});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_TRUE(r.stdout_text.empty());
  EXPECT_NE(r.stderr_text.find("warning: XLSB write diagnostics"), std::string::npos);
  EXPECT_NE(r.stderr_text.find("downgraded_formula_count=1"), std::string::npos);
  EXPECT_NE(r.stderr_text.find("deferred_feature_count=2"), std::string::npos);
}

TEST(FormulonCli, RecalcReportsOoxmlReadDiagnosticsSeparatelyFromXlsbOnes) {
  const std::string input = "/tmp/fm_cli_lossy_ooxml_input.xlsx";
  const std::string output = "/tmp/fm_cli_lossy_ooxml_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_lossy_ooxml_fixture(input));

  CliRun r = run_cli({"recalc", "--quiet", input, "-o", output});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stderr_text.find("warning: OOXML read diagnostics"), std::string::npos) << r.stderr_text;
  // One unparseable merge ref plus one conditional-formatting block with no
  // `sqref`; the unrecognised workbook content type is reported separately.
  EXPECT_NE(r.stderr_text.find("skipped_feature_count=2"), std::string::npos) << r.stderr_text;
  EXPECT_NE(r.stderr_text.find("unknown_content_type_count=1"), std::string::npos) << r.stderr_text;
  // The XLSB line must not appear: an `.xlsx` load produces none of its
  // counters, and a zero counter is never printed.
  EXPECT_EQ(r.stderr_text.find("XLSB read diagnostics"), std::string::npos) << r.stderr_text;
}

TEST(FormulonCli, RecalcDroppedPartWarningIsNonfatalAndQuietStillReportsIt) {
  const std::string input = "/tmp/fm_cli_dropped_input.xlsb";
  const std::string output = "/tmp/fm_cli_dropped_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  ASSERT_TRUE(write_dropped_xlsb_fixture(input));

  CliRun r = run_cli({"recalc", "--quiet", input, "-o", output});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stderr_text.find("undecoded_part_count=1"), std::string::npos);
  EXPECT_NE(r.stderr_text.find("warning: XLSB read diagnostics"), std::string::npos);
}

TEST(FormulonCli, RecalcInPlaceOverwritesSamePathAndLeavesNoTemp) {
  // Input and output are the same path: the atomic write must serialize the
  // recalculated workbook fully before replacing the original, and must not
  // leave its sibling temp file behind on success.
  std::string path = "/tmp/fm_cli_inplace.xlsx";
  std::string tmp_sidecar = path + ".formulon-tmp";
  PathGuard g_path(path);
  PathGuard g_tmp(tmp_sidecar);
  ASSERT_TRUE(write_fixture_workbook(path));

  CliRun r = run_cli({"recalc", path, "-o", path, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  // The temp sidecar must be gone (renamed into place).
  {
    std::ifstream leftover(tmp_sidecar, std::ios::binary);
    EXPECT_FALSE(leftover.good()) << "temp sidecar left behind after in-place recalc";
  }

  // The overwritten file still loads and recalculates correctly.
  std::ifstream f(path, std::ios::binary);
  ASSERT_TRUE(f);
  f.seekg(0, std::ios::end);
  const std::streamsize size = f.tellg();
  ASSERT_GT(size, 0);
  f.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(static_cast<std::size_t>(size));
  f.read(reinterpret_cast<char*>(bytes.data()), size);
  ASSERT_TRUE(f);

  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0, 0, 1, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 8.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcDoesNotReusePredictableLegacyTempPath) {
  const std::string input = "/tmp/fm_cli_safe_temp_in.xlsx";
  const std::string output = "/tmp/fm_cli_safe_temp_out.xlsx";
  const std::string legacy_temp = output + ".formulon-tmp";
  PathGuard g_input(input);
  PathGuard g_output(output);
  PathGuard g_legacy_temp(legacy_temp);
  ASSERT_TRUE(write_fixture_workbook(input));
  {
    std::ofstream sentinel(legacy_temp, std::ios::binary);
    ASSERT_TRUE(sentinel);
    sentinel << "unrelated file";
  }

  CliRun r = run_cli({"recalc", input, "-o", output, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  std::ifstream sentinel(legacy_temp, std::ios::binary);
  ASSERT_TRUE(sentinel);
  std::string contents;
  std::getline(sentinel, contents);
  EXPECT_EQ(contents, "unrelated file");
}

TEST(FormulonCli, RecalcThroughSymlinkUpdatesTargetAndKeepsLink) {
  // Saving through a symlink must update the file the link names. A
  // rename onto the link path would replace the link with a regular
  // file and leave the real workbook holding stale values -- the same
  // silent-loss shape the atomic write exists to prevent.
  const std::string input = "/tmp/fm_cli_symlink_in.xlsx";
  const std::string target = "/tmp/fm_cli_symlink_target.xlsx";
  const std::string link = "/tmp/fm_cli_symlink_link.xlsx";
  PathGuard g_input(input);
  PathGuard g_target(target);
  PathGuard g_link(link);
  ASSERT_TRUE(write_fixture_workbook(input));
  ASSERT_TRUE(write_fixture_workbook(target));
  std::remove(link.c_str());
  ASSERT_EQ(::symlink(target.c_str(), link.c_str()), 0);

  CliRun r = run_cli({"recalc", input, "-o", link, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  // The link survives as a link...
  struct stat link_stat {};
  ASSERT_EQ(::lstat(link.c_str(), &link_stat), 0);
  EXPECT_TRUE(S_ISLNK(link_stat.st_mode)) << "symlink was replaced by a regular file";

  // ...and the file it names is the one that got rewritten.
  std::ifstream written(target, std::ios::binary);
  ASSERT_TRUE(written);
  written.seekg(0, std::ios::end);
  EXPECT_GT(written.tellg(), 0);

  // No temp sidecar is left next to either path.
  for (const std::string& sidecar : {target + ".formulon-tmp", link + ".formulon-tmp"}) {
    std::ifstream leftover(sidecar, std::ios::binary);
    EXPECT_FALSE(leftover.good()) << "temp sidecar left behind: " << sidecar;
  }
}

TEST(FormulonCli, RecalcDanglingSymlinkIsReplacedInPlace) {
  // A link with no target has nothing to preserve, so the write lands
  // on the link path itself rather than failing.
  const std::string input = "/tmp/fm_cli_dangling_in.xlsx";
  const std::string link = "/tmp/fm_cli_dangling_link.xlsx";
  PathGuard g_input(input);
  PathGuard g_link(link);
  ASSERT_TRUE(write_fixture_workbook(input));
  std::remove(link.c_str());
  ASSERT_EQ(::symlink("/tmp/fm_cli_dangling_absent.xlsx", link.c_str()), 0);

  CliRun r = run_cli({"recalc", input, "-o", link, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  std::ifstream written(link, std::ios::binary);
  ASSERT_TRUE(written);
  written.seekg(0, std::ios::end);
  EXPECT_GT(written.tellg(), 0);
}

TEST(FormulonCli, RecalcLeavesOutputIntactWhenTheWriteCannotStart) {
  // Stand-in for a disk-full failure: an unwritable directory makes the
  // temp file impossible to create. The pre-existing output must be left
  // exactly as it was, with no partial content and no sidecar.
  const std::string dir = "/tmp/fm_cli_readonly_dir";
  const std::string input = "/tmp/fm_cli_readonly_in.xlsx";
  const std::string output = dir + "/out.xlsx";
  PathGuard g_input(input);
  ASSERT_TRUE(write_fixture_workbook(input));
  ::rmdir(dir.c_str());
  ASSERT_EQ(::mkdir(dir.c_str(), 0700), 0);
  ASSERT_TRUE(write_fixture_workbook(output));

  std::vector<std::uint8_t> before;
  {
    std::ifstream original(output, std::ios::binary);
    ASSERT_TRUE(original);
    before.assign(std::istreambuf_iterator<char>(original), std::istreambuf_iterator<char>());
  }
  ASSERT_FALSE(before.empty());
  ASSERT_EQ(::chmod(dir.c_str(), 0500), 0);

  CliRun r = run_cli({"recalc", input, "-o", output, "--quiet"});
  EXPECT_NE(r.exit_code, 0) << "write into an unwritable directory should fail";

  // Restore write permission so the fixture can be inspected and removed.
  ASSERT_EQ(::chmod(dir.c_str(), 0700), 0);
  std::vector<std::uint8_t> after;
  {
    std::ifstream survivor(output, std::ios::binary);
    ASSERT_TRUE(survivor) << "existing output was destroyed by a failed write";
    after.assign(std::istreambuf_iterator<char>(survivor), std::istreambuf_iterator<char>());
  }
  EXPECT_EQ(before, after) << "existing output was modified by a failed write";

  std::remove(output.c_str());
  ::rmdir(dir.c_str());
}

TEST(FormulonCli, RecalcPreservesExistingOutputPermissions) {
  std::string in = "/tmp/fm_cli_mode_in.xlsx";
  std::string out_path = "/tmp/fm_cli_mode_out.xlsx";
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  ASSERT_TRUE(write_fixture_workbook(in));
  ASSERT_TRUE(write_fixture_workbook(out_path));
  ASSERT_EQ(::chmod(out_path.c_str(), 0600), 0);

  CliRun r = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  struct stat saved_stat {};
  ASSERT_EQ(::stat(out_path.c_str(), &saved_stat), 0);
  EXPECT_EQ(saved_stat.st_mode & 0777U, 0600U);
}

TEST(FormulonCli, RecalcCreatesNewOutputWithUmaskDefaultPermissions) {
  // A fresh output path must come out with the mode an ordinary file
  // creation would produce, not the private mode the atomic-write
  // temporary is created with. The reference file below is created the
  // ordinary way, so the expectation follows whatever umask is in effect.
  std::string in = "/tmp/fm_cli_newmode_in.xlsx";
  std::string out_path = "/tmp/fm_cli_newmode_out.xlsx";
  std::string reference = "/tmp/fm_cli_newmode_reference";
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  PathGuard g_reference(reference);
  ASSERT_TRUE(write_fixture_workbook(in));
  std::remove(out_path.c_str());
  std::remove(reference.c_str());

  const int reference_fd = ::open(reference.c_str(), O_CREAT | O_WRONLY | O_EXCL, 0666);
  ASSERT_GE(reference_fd, 0);
  ASSERT_EQ(::close(reference_fd), 0);
  struct stat reference_stat {};
  ASSERT_EQ(::stat(reference.c_str(), &reference_stat), 0);

  CliRun r = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  ASSERT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  struct stat created_stat {};
  ASSERT_EQ(::stat(out_path.c_str(), &created_stat), 0);
  EXPECT_EQ(created_stat.st_mode & 0777U, reference_stat.st_mode & 0777U);

  // Re-running against the now-existing path keeps that same mode, so the
  // permissions do not depend on whether the output already existed.
  CliRun again = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  ASSERT_EQ(again.exit_code, 0) << "stderr=" << again.stderr_text;
  struct stat rerun_stat {};
  ASSERT_EQ(::stat(out_path.c_str(), &rerun_stat), 0);
  EXPECT_EQ(rerun_stat.st_mode & 0777U, created_stat.st_mode & 0777U);
}

TEST(FormulonCli, RecalcIterativePreservesExistingIterationSettings) {
  std::string in = "/tmp/fm_cli_iterative_in.xlsx";
  std::string out_path = "/tmp/fm_cli_iterative_out.xlsx";
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  ASSERT_TRUE(write_fixture_workbook(in, {}, /*iterative=*/true, /*max_iterations=*/500, /*max_change=*/0.01));

  CliRun r = run_cli({"recalc", "--iterative", in, "-o", out_path, "--quiet"});
  ASSERT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  std::ifstream f(out_path, std::ios::binary);
  ASSERT_TRUE(f);
  f.seekg(0, std::ios::end);
  const std::streamsize size = f.tellg();
  ASSERT_GT(size, 0);
  f.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(static_cast<std::size_t>(size));
  ASSERT_TRUE(f.read(reinterpret_cast<char*>(bytes.data()), size));

  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  int32_t enabled = 0;
  uint32_t max_iterations = 0;
  double max_change = 0.0;
  ASSERT_EQ(fm_workbook_get_iterative(wb, &enabled, &max_iterations, &max_change), 0);
  EXPECT_EQ(enabled, 1);
  EXPECT_EQ(max_iterations, 500U);
  EXPECT_DOUBLE_EQ(max_change, 0.01);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcXlsbOutputExtensionWritesXlsbContainer) {
  // `-o out.xlsb` must select the MS-XLSB writer, not silently emit an
  // OOXML package under an `.xlsb` name.
  std::string in = "/tmp/fm_cli_in_for_xlsb.xlsx";
  std::string out_path = "/tmp/fm_cli_out.xlsb";
  PathGuard g_in(in);
  PathGuard g_out(out_path);
  ASSERT_TRUE(write_fixture_workbook(in));

  CliRun r = run_cli({"recalc", in, "-o", out_path, "--quiet"});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;

  std::ifstream f(out_path, std::ios::binary);
  ASSERT_TRUE(f);
  f.seekg(0, std::ios::end);
  const std::streamsize size = f.tellg();
  ASSERT_GT(size, 0);
  f.seekg(0, std::ios::beg);
  std::vector<std::uint8_t> bytes(static_cast<std::size_t>(size));
  f.read(reinterpret_cast<char*>(bytes.data()), size);
  ASSERT_TRUE(f);

  // The written package must declare `xl/workbook.bin` (xlsb), not
  // `xl/workbook.xml` (ooxml).
  formulon::io::ByteSpan span{bytes.data(), bytes.size()};
  EXPECT_EQ(formulon::io::detect_workbook_format(span), formulon::io::WorkbookFormat::Xlsb);

  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0, 0, 1, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 8.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcMissingOutputExits64) {
  std::string in = "/tmp/fm_cli_in_missing_o.xlsx";
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"recalc", in});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, RecalcMissingInputExits64) {
  CliRun r = run_cli({"recalc", "-o", "/tmp/fm_cli_unused.xlsx"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, RecalcNonexistentFileFailsCleanly) {
  CliRun r = run_cli({"recalc", "/tmp/this_path_does_not_exist_abc123.xlsx", "-o", "/tmp/fm_cli_unused.xlsx"});
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_EQ(r.stderr_text.find("warning:"), std::string::npos);
}

TEST(FormulonCli, RecalcOptionTerminatorAcceptsDashLeadingRelativePaths) {
  const std::string input = "-fm_cli_recalc_dash_input.xlsx";
  const std::string output = "-fm_cli_recalc_dash_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  std::remove(output.c_str());
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", "--threads", "1", "-o", output, "--", input});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;

  std::ifstream saved(output, std::ios::binary);
  ASSERT_TRUE(saved);
  const std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(saved)), std::istreambuf_iterator<char>());
  ASSERT_FALSE(bytes.empty());
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0U, 0U, 1U, &value), 0);
  EXPECT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 8.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, RecalcDashLeadingInputWithoutTerminatorIsUsageError) {
  const std::string input = "-fm_cli_recalc_dash_without_terminator_input.xlsx";
  const std::string output = "-fm_cli_recalc_dash_without_terminator_output.xlsx";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  std::remove(output.c_str());
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", input, "-o", output});
  EXPECT_EQ(result.exit_code, 64);
  EXPECT_TRUE(result.stdout_text.empty());
  EXPECT_FALSE(std::ifstream(output, std::ios::binary).good());
}

TEST(FormulonCli, RecalcAcceptsDoubleDashAsOutputPathValue) {
  const std::string input = "fm_cli_recalc_double_dash_input.xlsx";
  const std::string output = "--";
  PathGuard input_guard(input);
  PathGuard output_guard(output);
  std::remove(output.c_str());
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"recalc", input, "-o", output});
  ASSERT_EQ(result.exit_code, 0) << result.stderr_text;

  std::ifstream saved(output, std::ios::binary);
  ASSERT_TRUE(saved);
  const std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(saved)), std::istreambuf_iterator<char>());
  ASSERT_FALSE(bytes.empty());
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb, 0U, 0U, 1U, &value), 0);
  EXPECT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 8.0);
  fm_workbook_destroy(wb);
}

TEST(FormulonCli, OptionTerminatorRequiresExactlyOnePositional) {
  const std::string output = "/tmp/fm_cli_terminator_recalc_output.xlsx";
  PathGuard output_guard(output);

  for (const std::vector<std::string>& args : {
           std::vector<std::string>{"eval", "--"},
           std::vector<std::string>{"dump", "--"},
           std::vector<std::string>{"paginate", "--"},
       }) {
    const CliRun result = run_cli(args);
    EXPECT_EQ(result.exit_code, 64) << "bare terminator command=" << args.front();
  }

  std::remove(output.c_str());
  CliRun recalc_bare = run_cli({"recalc", "-o", output, "--"});
  EXPECT_EQ(recalc_bare.exit_code, 64);
  EXPECT_FALSE(std::ifstream(output, std::ios::binary).good());

  for (const std::vector<std::string>& args : {
           std::vector<std::string>{"eval", "--", "-1+2", "--json"},
           std::vector<std::string>{"dump", "--", "-fm_cli_missing.xlsx", "--sheets"},
           std::vector<std::string>{"paginate", "--", "-fm_cli_missing.xlsx", "--sheet"},
       }) {
    const CliRun result = run_cli(args);
    EXPECT_EQ(result.exit_code, 64) << "post-terminator extra command=" << args.front();
  }

  std::remove(output.c_str());
  CliRun recalc_extra = run_cli({"recalc", "-o", output, "--", "-fm_cli_missing.xlsx", "--quiet"});
  EXPECT_EQ(recalc_extra.exit_code, 64);
  EXPECT_FALSE(std::ifstream(output, std::ios::binary).good());
}

// ---------------------------------------------------------------------------
// dump
// ---------------------------------------------------------------------------

TEST(FormulonCli, DumpSheetsListsInOrder) {
  std::string in = "/tmp/fm_cli_dump_sheets.xlsx";
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--sheets", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_EQ(r.stdout_text, "Sheet1\n");
}

TEST(FormulonCli, DumpFormulasListsFormulaCells) {
  std::string in = "/tmp/fm_cli_dump_formulas.xlsx";
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--formulas", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  // B1 = `=A1+1` is the only formula; A1 is a literal.
  EXPECT_NE(r.stdout_text.find("Sheet1!B1"), std::string::npos) << "stdout=" << r.stdout_text;
  EXPECT_EQ(r.stdout_text.find("Sheet1!A1"), std::string::npos) << "literal A1 must not appear in --formulas dump";
}

TEST(FormulonCli, DumpMetadataPrintsSectionHeaders) {
  std::string in = "/tmp/fm_cli_dump_metadata.xlsx";
  PathGuard g_in(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--metadata", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stdout_text.find("[defined_names]"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("[tables]"), std::string::npos);
  EXPECT_NE(r.stdout_text.find("[passthrough_parts]"), std::string::npos);
}

TEST(FormulonCli, DumpWriteFailureReturnsOutputError) {
  std::string in = "/tmp/fm_cli_dump_write_failure.xlsx";
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in));
  CliRun r = run_cli({"dump", "--sheets", in}, /*merge_streams=*/false, /*close_stdout=*/true);
  EXPECT_EQ(r.exit_code, 1);
  EXPECT_NE(r.stderr_text.find("failed to write output"), std::string::npos);
}

TEST(FormulonCli, DumpValuesEscapesEmbeddedNewlines) {
  std::string in = "/tmp/fm_cli_dump_escaped.xlsx";
  PathGuard guard(in);
  ASSERT_TRUE(write_fixture_workbook(in, "first\nsecond"));
  CliRun r = run_cli({"dump", "--values", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_NE(r.stdout_text.find("Sheet1!A2 first\\nsecond\n"), std::string::npos) << r.stdout_text;
  EXPECT_EQ(r.stdout_text.find("first\nsecond"), std::string::npos) << r.stdout_text;
}

TEST(FormulonCli, DumpMissingInputExits64) {
  CliRun r = run_cli({"dump", "--sheets"});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, DumpOptionTerminatorAcceptsDashLeadingRelativePath) {
  const std::string input = "-fm_cli_dump_dash_input.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun without_terminator = run_cli({"dump", "--sheets", input});
  EXPECT_EQ(without_terminator.exit_code, 64);

  const CliRun with_terminator = run_cli({"dump", "--sheets", "--", input});
  EXPECT_EQ(with_terminator.exit_code, 0) << with_terminator.stderr_text;
  EXPECT_EQ(with_terminator.stdout_text, "Sheet1\n");
}

TEST(FormulonCli, DumpPostTerminatorHelpIsAnInputPath) {
  const std::string input = "-h";
  PathGuard guard(input);
  std::remove(input.c_str());

  const CliRun result = run_cli({"dump", "--", input});
  EXPECT_EQ(result.exit_code, 1);
  EXPECT_TRUE(result.stdout_text.empty());
  EXPECT_NE(result.stderr_text.find("cannot read"), std::string::npos);
}

TEST(FormulonCli, SecondLiteralTerminatorIsAnExtraPositional) {
  const std::string input = "-fm_cli_second_literal_terminator.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun dump = run_cli({"dump", "--", input, "--"});
  EXPECT_EQ(dump.exit_code, 64);
  const CliRun paginate = run_cli({"paginate", "--", input, "--"});
  EXPECT_EQ(paginate.exit_code, 64);
}

TEST(FormulonCli, DumpMetadataDistinguishesDefinedNameScope) {
  std::string in = "/tmp/fm_cli_dump_metadata_scoped.xlsx";
  PathGuard g_in(in);

  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create(&wb), 0);
  ASSERT_EQ(fm_workbook_set_number(wb, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_defined_name(wb, "WorkbookConst", "=1"), 0);
  ASSERT_EQ(fm_workbook_set_defined_name_scoped(wb, "SheetLocal", "=Sheet1!$A$1", 0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, &bytes, &len), 0);
  std::ofstream out(in, std::ios::binary | std::ios::trunc);
  ASSERT_TRUE(out);
  out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
  out.close();
  fm_buffer_free(bytes);
  fm_workbook_destroy(wb);

  CliRun r = run_cli({"dump", "--metadata", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  // Workbook-scoped names print bare; sheet-scoped names are prefixed
  // with `SheetName!` so the two scopes don't collide in the dump.
  EXPECT_NE(r.stdout_text.find("WorkbookConst =1"), std::string::npos) << "stdout=" << r.stdout_text;
  EXPECT_NE(r.stdout_text.find("Sheet1!SheetLocal ="), std::string::npos) << "stdout=" << r.stdout_text;
}

TEST(FormulonCli, DumpMetadataRecoversFromOutOfRangeDefinedNameScope) {
  std::string in = "/tmp/fm_cli_dump_metadata_out_of_range_scope.xlsx";
  PathGuard guard(in);
  ASSERT_TRUE(write_out_of_range_defined_names_fixture(in));

  CliRun r = run_cli({"dump", "--metadata", in});
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_TRUE(r.stderr_text.empty());
  EXPECT_EQ(r.stdout_text,
            "[defined_names]\n"
            "Before =1\n"
            "#99!Bad =Sheet1!$A$1\n"
            "After =2\n"
            "[tables]\n"
            "[passthrough_parts]\n");
}

TEST(FormulonCli, PaginateOptionTerminatorAcceptsDashLeadingRelativePath) {
  const std::string input = "-fm_cli_paginate_dash_input.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun without_terminator = run_cli({"paginate", "--sheet", "0", input});
  EXPECT_EQ(without_terminator.exit_code, 64);

  const CliRun with_terminator = run_cli({"paginate", "--sheet", "0", "--", input});
  EXPECT_EQ(with_terminator.exit_code, 0) << with_terminator.stderr_text;
  EXPECT_EQ(with_terminator.stdout_text, "sheet=0\npages=1\nprint_area=\nhorizontal_breaks=\nvertical_breaks=\n");
}

TEST(FormulonCli, PaginatePostTerminatorHelpIsAnInputPath) {
  const std::string input = "-h";
  PathGuard guard(input);
  std::remove(input.c_str());

  const CliRun result = run_cli({"paginate", "--", input});
  EXPECT_EQ(result.exit_code, 1);
  EXPECT_TRUE(result.stdout_text.empty());
  EXPECT_NE(result.stderr_text.find("cannot read"), std::string::npos);
}

TEST(FormulonCli, PaginateConsumesTerminatorAsSheetValue) {
  const std::string input = "-fm_cli_paginate_sheet_value_terminator.xlsx";
  PathGuard guard(input);
  ASSERT_TRUE(write_fixture_workbook(input));

  const CliRun result = run_cli({"paginate", "--sheet", "--", input});
  EXPECT_EQ(result.exit_code, 64);
  EXPECT_TRUE(result.stdout_text.empty());
}
