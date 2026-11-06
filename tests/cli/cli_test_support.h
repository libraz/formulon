#pragma once

#include <dirent.h>
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

#ifndef FORMULON_CLI_PATH
#error "FORMULON_CLI_PATH must be defined by the build system"
#endif

#include "c_api/formulon_c.h"
#include "cli/file_io.h"
#include "defined_name.h"
#include "io/format_detect.h"
#include "support/ooxml_package_fixture.h"
#include "workbook.h"

namespace formulon::cli_test {

struct CliRun {
  int exit_code = -1;
  std::string stdout_text;
  std::string stderr_text;
};

// All subprocess and fixture helpers have external linkage in this one
// implementation unit. In particular, temp_prefix() owns one process-wide
// prefix and static initialization, so every split test file shares it.
std::string sh_quote(const std::string& s);
const std::string& temp_prefix();
std::string temp_path(std::string_view name);
std::vector<std::string> list_directory(const std::string& dir);
CliRun run_cli(const std::vector<std::string>& args, bool merge_streams = false, bool close_stdout = false,
               const std::string& stdin_path = {});

/// How `run_cli_spawn` wires the child's standard streams. Streams not made a
/// terminal or a broken pipe read from `/dev/null` (stdin) or are captured
/// (stdout); stderr is always captured.
struct SpawnOptions {
  bool stdin_tty = false;           ///< stdin is the slave side of a pseudo-terminal.
  bool stdout_tty = false;          ///< stdout is the slave side of a pseudo-terminal.
  bool stdout_broken_pipe = false;  ///< stdout is a pipe whose read end is already closed.
};

/// Runs the CLI without a shell so the child's descriptors are exactly the
/// ones `options` names. A child that does not exit within a few seconds is
/// killed and reported with exit code -5.
CliRun run_cli_spawn(const std::vector<std::string>& args, const SpawnOptions& options);

/// Writes `bytes` to `path`. Returns true on success.
bool write_bytes(const std::string& path, std::string_view bytes);

struct PathGuard {
  std::string path;
  explicit PathGuard(std::string p);
  PathGuard(const PathGuard&) = delete;
  PathGuard& operator=(const PathGuard&) = delete;
  ~PathGuard();
};

struct DirGuard {
  std::string path;
  explicit DirGuard(std::string p);
  DirGuard(const DirGuard&) = delete;
  DirGuard& operator=(const DirGuard&) = delete;
  ~DirGuard();
};

std::vector<std::string> read_diagnostic_lines(const std::string& text);
int open_descriptor_count();
bool write_fixture_workbook(const std::string& path, std::string_view extra_text = {}, bool iterative = false,
                            int32_t max_iterations = 100, double max_change = 0.001);
bool write_wide_recalc_fixture(const std::string& path);
bool write_out_of_range_defined_names_fixture(const std::string& path);
bool write_lossy_xlsx_fixture(const std::string& path);
bool write_lossy_ooxml_fixture(const std::string& path);
bool write_dropped_xlsb_fixture(const std::string& path);

}  // namespace formulon::cli_test
