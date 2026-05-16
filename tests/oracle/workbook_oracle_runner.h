// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Fixture support for the workbook oracle gtest target.
//
// The workbook oracle is a parameterized Google Test covering workbook-
// level features that are NOT formula results -- pivot tables and print
// areas. It pulls cases from the committed golden JSON files under
// `tests/oracle/golden_wb/`. Each JSON file is a `suite` whose `kind` is
// `"workbook"` and contains N `cases`; every (suite, case) pair is
// expanded into one TEST_P parameter so failures point at a specific
// case id, not a whole suite.
//
// The runner is deliberately pure-C++ + stdlib: no YAML, no third-party
// JSON. The `json_reader.{h,cpp}` in this directory is the tiny in-test
// JSON parser; golden files are produced by the Python workbook oracle
// generator (`tools/oracle/workbook_oracle_gen.py`).
//
// The golden directory is passed by the build system as a compile-time
// define (`FORMULON_WORKBOOK_ORACLE_GOLDEN_DIR`); when that path is empty
// or the directory does not exist, the parameter list is empty and the
// workbook oracle test suite silently registers zero cases -- the build
// stays green before any goldens have been generated.

#ifndef FORMULON_TESTS_ORACLE_WORKBOOK_ORACLE_RUNNER_H_
#define FORMULON_TESTS_ORACLE_WORKBOOK_ORACLE_RUNNER_H_

#include <ostream>
#include <string>
#include <utility>
#include <vector>

#include "tests/oracle/json_reader.h"

namespace formulon {
namespace tests {
namespace oracle {

/// One (suite, case) pair as loaded from a workbook golden JSON file.
///
/// `suite` is the golden file's basename (without the `.golden.json`
/// suffix) and doubles as the gtest parameter name prefix. `case_id` is
/// the `id` field of the case inside the file. `spec` is the declarative
/// mini-workbook spec (sheets / column widths / row heights / pivot /
/// print blocks); `expect` is the observed pivot / print result the
/// verifier compares against. `environment` is the suite's environment
/// record (Excel version / locale). `variant` carries a non-empty tag for
/// non-primary golden trees and is empty for the primary track.
struct WorkbookOracleCase {
  std::string suite;
  std::string case_id;
  std::string source_file;  // absolute path to the golden JSON, for diagnostics
  JsonValue spec;           // the declarative workbook spec
  JsonValue expect;         // the observed pivot / print result
  JsonValue environment;    // the suite's environment record
  std::string variant;      // "" for primary; non-empty for a variant tree
};

/// Loads every `*.golden.json` file under `golden_dir` and flattens them
/// into a list of `WorkbookOracleCase` parameters. Files that fail to
/// parse are surfaced as a single `WorkbookOracleCase` with
/// `case_id = "<load-error>"` so the failure is visible in the test
/// output rather than silently dropped.
///
/// Returns an empty vector when `golden_dir` is empty or does not exist --
/// the expected state while the workbook oracle generator has not yet
/// been run on a Windows + Excel host.
///
/// `variant_tag` is stamped onto every loaded `WorkbookOracleCase::variant`.
/// The default empty string keeps the primary-only behaviour; callers
/// loading a non-primary golden tree should pass the matching tag.
std::vector<WorkbookOracleCase> load_workbook_oracle_cases(const std::string& golden_dir,
                                                           const std::string& variant_tag = "");

/// Returns variant workbook-golden roots configured for this build, as
/// (tag, golden_wb_dir) pairs. Mirrors `configured_variant_dirs()` in
/// `oracle_runner.h` but resolves the workbook track's `golden_wb/` tree.
/// Source priority:
///   1. Env var FORMULON_WORKBOOK_ORACLE_VARIANT_DIRS (':' separated
///      entries of the form "tag=/abs/path"), if set.
///   2. Compile-time FORMULON_WORKBOOK_ORACLE_VARIANT_DIRS_DEFAULT, if set.
///   3. Empty vector.
std::vector<std::pair<std::string, std::string>> configured_workbook_variant_dirs();

/// Keeps gtest from printing the raw `WorkbookOracleCase` byte dump when
/// it renders test names. Without this overload every discovered test
/// gets a long `# GetParam() = <bytes>` trailer.
inline std::ostream& operator<<(std::ostream& os, const WorkbookOracleCase& c) {
  os << c.suite << '.' << c.case_id;
  return os;
}

}  // namespace oracle
}  // namespace tests
}  // namespace formulon

#endif  // FORMULON_TESTS_ORACLE_WORKBOOK_ORACLE_RUNNER_H_
