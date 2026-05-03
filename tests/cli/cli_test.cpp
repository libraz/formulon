// Copyright 2026 libraz. Licensed under the MIT License.
//
// `formulon_cli` end-to-end tests. Each test spawns the binary at the
// CMake-injected `FORMULON_CLI_PATH` and asserts on (exit_code, stdout,
// stderr). The harness uses `popen` because it is the simplest
// portable way to run a subprocess and capture output without pulling
// in a third-party library; `popen` does not give us stderr capture
// natively, so we redirect stderr to stdout via the shell. Tests that
// need to distinguish the streams use the explicit `--quiet` flag or
// inspect output ordering.

#include <gtest/gtest.h>
#include <unistd.h>

#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <fstream>
#include <ios>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"

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
CliRun run_cli(const std::vector<std::string>& args, bool merge_streams = false) {
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
  if (merge_streams) {
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
bool write_fixture_workbook(const std::string& path) {
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
  if (fm_workbook_recalc(wb) != 0) {
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

TEST(FormulonCli, NoArgsExits64) {
  CliRun r = run_cli({});
  EXPECT_EQ(r.exit_code, 64);
}

TEST(FormulonCli, UnknownCommandExits64) {
  CliRun r = run_cli({"frobnicate"});
  EXPECT_EQ(r.exit_code, 64);
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
  // `=A1+1` references an empty cell on the empty workbook. The cell
  // value is the engine's behavior for empty refs (typically 1, since
  // blank coerces to 0 and `0+1=1`), but the call must not fail.
  CliRun r = run_cli({"eval", "=A1+1"});
  // Excel cell-level errors (#REF!, etc.) print to stdout and the
  // process exits 0. We do not pin the exact value here — the engine
  // may evolve — but the call must succeed.
  EXPECT_EQ(r.exit_code, 0) << "stderr=" << r.stderr_text;
  EXPECT_GT(r.stdout_text.size(), 0U);
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

TEST(FormulonCli, EvalRepeatRunsMultipleTimes) {
  CliRun r = run_cli({"eval", "--repeat", "3", "=SUM(1,2,3)"});
  EXPECT_EQ(r.exit_code, 0);
  EXPECT_EQ(r.stdout_text, "6\n");
  // The timing line goes to stderr.
  EXPECT_NE(r.stderr_text.find("3 iterations"), std::string::npos);
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
  // Non-zero exit; specific code is the low byte of `kCliFileNotFound`.
  EXPECT_NE(r.exit_code, 0);
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

TEST(FormulonCli, DumpMissingInputExits64) {
  CliRun r = run_cli({"dump", "--sheets"});
  EXPECT_EQ(r.exit_code, 64);
}
