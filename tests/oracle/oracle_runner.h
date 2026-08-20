//
// Fixture support for the oracle gtest target.
//
// The oracle verifier is a parameterized Google Test that pulls cases from
// the committed golden JSON files under `tests/oracle/golden/`. Each JSON
// file is a `suite` and contains N `cases`; every (suite, case) pair is
// expanded into one TEST_P parameter so failures point at a specific case
// id, not a whole suite.
//
// The runner is deliberately pure-C++ + stdlib: no YAML, no third-party
// JSON. The `json_reader.{h,cpp}` in this directory is our tiny in-test JSON
// parser; golden files are produced by the Python oracle-gen pipeline.
//
// Case discovery happens at static-init time via a helper function the test
// main harness invokes. The golden directory is passed by the build system
// as a compile-time define (`FORMULON_ORACLE_GOLDEN_DIR`); when that path is
// empty or the directory does not exist, the parameter list is empty and
// the oracle test suite silently registers zero cases — the build stays
// green before any goldens have been generated.

#ifndef FORMULON_TESTS_ORACLE_ORACLE_RUNNER_H_
#define FORMULON_TESTS_ORACLE_ORACLE_RUNNER_H_

#include <cstdint>
#include <map>
#include <ostream>
#include <string>
#include <utility>
#include <vector>

#include "tests/oracle/json_reader.h"

namespace formulon {
namespace tests {
namespace oracle {

/// One (suite, case) pair as loaded from a golden JSON file.
///
/// `suite` is the golden file's basename (without `.golden.json` suffix) and
/// doubles as the gtest parameter name prefix. `case_id` is the `id` field
/// of the case inside the file. `raw_case` is the original JSON object for
/// the case, kept so the fixture can read `formula`, `setup`, `expect`, and
/// any suite-level defaults the Python generator inlined per case.
struct OracleCase {
  std::string suite;
  std::string case_id;
  std::string source_file;  // absolute path to the golden JSON, for diagnostics
  JsonValue raw_case;       // the case object (has formula/setup/expect/...)
  JsonValue environment;    // the suite's environment record (excel_version etc.)
  double tolerance_abs = 0.0;
  double tolerance_rel = 0.0;
  // "" or "exact" -> byte-equality; other values (e.g. "complex_text") select
  // a structured comparator in the verifier.
  std::string compare_mode;
  // Variant tag (e.g. "win_excel"). Empty string means primary (Mac/ja-JP)
  // and is used as the historical default; non-empty values are appended to
  // the gtest parameter name as `__<tag>` so primary and variant cases
  // never collide.
  std::string variant;
  // Non-empty when the case must not be verified: either the golden was
  // captured with the `skipped` field already baked in, or the divergence
  // registry names the case id. Empty means "verify this case".
  std::string skipped_reason;
};

/// Case ids the divergence registry excludes from verification, mapped to
/// the registry's reason.
///
/// `mode: skip-oracle` in tests/divergence.yaml is a verification-time
/// policy, so it has to apply to goldens captured before the entry existed.
/// Baking the skip into the golden -- all the generator can do -- would make
/// a newly adjudicated divergence wait on a fresh Excel capture to take
/// effect on every machine, which is the deadlock this path exists to
/// avoid. CMake projects the registry to JSON at configure time
/// (tools/oracle/emit_skip_list.py) and hands the path in as
/// `FORMULON_ORACLE_SKIP_FILE`; an unset or unreadable file yields no
/// skips.
const std::map<std::string, std::string>& oracle_skip_registry();

/// Loads every `*.golden.json` file under `golden_dir` and flattens them
/// into a list of `OracleCase` parameters. Files that fail to parse are
/// surfaced as a single `OracleCase` with `case_id = "<load-error>"` so the
/// failure is visible in the test output rather than silently dropped.
///
/// Returns an empty vector when `golden_dir` is empty or does not exist —
/// the expected state while the oracle-gen pipeline has not yet been run
/// on a Mac.
///
/// `variant_tag` is stamped onto every loaded `OracleCase::variant`. The
/// default empty string preserves the historical primary-only behaviour;
/// callers loading a non-primary golden tree (e.g. a Win-Excel variant)
/// should pass the matching tag so the verifier can disambiguate parameter
/// names and report results per variant.
std::vector<OracleCase> load_oracle_cases(const std::string& golden_dir, const std::string& variant_tag = "");

/// Returns variant golden roots configured for this build, as
/// (tag, golden_dir) pairs. Source priority:
///   1. Env var FORMULON_ORACLE_VARIANT_DIRS (':' separated entries
///      of the form "tag=/abs/path"), if set.
///   2. Compile-time FORMULON_ORACLE_VARIANT_DIRS_DEFAULT, if set.
///   3. Empty vector.
std::vector<std::pair<std::string, std::string>> configured_variant_dirs();

/// Translates an A1 address ("A1", "BC42") into 0-based `(row, col)`.
/// Returns `false` on malformed input. Sheet qualifiers are rejected;
/// callers wanting to parse `"Sheet2!A1"` should first feed the key
/// through `split_sheet_qualified_addr` (below) and pass only the
/// returned `addr` part to this function.
bool a1_to_row_col(const std::string& a1, std::uint32_t* out_row, std::uint32_t* out_col);

/// Splits a setup-key string of the form ``"Sheet2!A1"`` or
/// ``"'My Sheet'!B5"`` into a sheet name and a bare A1 address. Returns
/// `(empty_string, key)` for bare keys (no `!` present). Single-quoted
/// sheet names are unquoted, and the Excel-standard `''` escape for an
/// embedded apostrophe is collapsed to a single quote. The split is on
/// the last `!` so a future stray separator in a quoted name does not
/// confuse the parse.
std::pair<std::string, std::string> split_sheet_qualified_addr(const std::string& key);

/// Returns the compile-time-configured golden directory, or an empty
/// string when the build didn't set one. Declared here so tests can call
/// the same lookup the default INSTANTIATE_TEST_SUITE_P uses.
std::string configured_golden_dir();

/// Keeps gtest from printing the raw 168-byte `OracleCase` dump when it
/// renders test names via `--gtest_list_tests`. Without this overload every
/// discovered test has a long `# GetParam() = <bytes>` trailer that
/// `gtest_discover_tests` mistakenly treats as part of the test name.
inline std::ostream& operator<<(std::ostream& os, const OracleCase& c) {
  os << c.suite << '.' << c.case_id;
  return os;
}

}  // namespace oracle
}  // namespace tests
}  // namespace formulon

#endif  // FORMULON_TESTS_ORACLE_ORACLE_RUNNER_H_
