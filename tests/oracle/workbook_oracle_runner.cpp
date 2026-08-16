
#include "tests/oracle/workbook_oracle_runner.h"

#include <algorithm>
#include <cstdlib>
#include <filesystem>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "tests/oracle/json_reader.h"

// Path to the JSON projection of tests/divergence.yaml that CMake emits at
// configure time (tools/oracle/emit_skip_list.py). Empty disables the
// registry lookup, which is the state in a build configured without the
// oracle Python venv.
#ifndef FORMULON_WORKBOOK_ORACLE_SKIP_FILE
#define FORMULON_WORKBOOK_ORACLE_SKIP_FILE ""
#endif

#ifndef FORMULON_WORKBOOK_ORACLE_VARIANT_DIRS_DEFAULT
#define FORMULON_WORKBOOK_ORACLE_VARIANT_DIRS_DEFAULT ""
#endif

namespace formulon {
namespace tests {
namespace oracle {

namespace {

// Length of the ".golden" infix left on the stem after stripping the
// ".json" extension from a "<suite>.golden.json" filename.
constexpr std::size_t kGoldenInfixLen = 7;

// Emits one synthetic case whose id is `<load-error>` so the failure shows
// up as a failing TEST_P parameter rather than silently dropping the file.
WorkbookOracleCase make_load_error(const std::string& path, const std::string& detail) {
  WorkbookOracleCase c;
  c.suite = std::filesystem::path(path).stem().string();
  // Strip the ".golden" left over after `.golden.json` -> `.golden` stem.
  if (c.suite.size() >= kGoldenInfixLen &&
      c.suite.compare(c.suite.size() - kGoldenInfixLen, kGoldenInfixLen, ".golden") == 0) {
    c.suite.resize(c.suite.size() - kGoldenInfixLen);
  }
  c.case_id = "<load-error>";
  c.source_file = path;

  std::map<std::string, JsonValue> err_case;
  err_case.emplace("id", JsonValue::make_string("<load-error>"));
  err_case.emplace("error", JsonValue::make_string(detail));
  c.spec = JsonValue::make_object(std::move(err_case));
  c.expect = JsonValue::make_object({});
  c.environment = JsonValue::make_object({});
  return c;
}

}  // namespace

std::vector<std::pair<std::string, std::string>> configured_workbook_variant_dirs() {
  // Env var wins; the compile-time scan keeps `ctest` runs in a vanilla
  // build tree functional without per-shell setup. Both sources share the
  // same `tag=path[:tag=path]*` grammar as the formula track.
  std::string raw;
  if (const char* env = std::getenv("FORMULON_WORKBOOK_ORACLE_VARIANT_DIRS"); env != nullptr) {
    raw = env;
  } else {
    raw = FORMULON_WORKBOOK_ORACLE_VARIANT_DIRS_DEFAULT;
  }

  std::vector<std::pair<std::string, std::string>> out;
  if (raw.empty()) {
    return out;
  }

  std::size_t start = 0;
  while (start <= raw.size()) {
    std::size_t end = raw.find(':', start);
    if (end == std::string::npos) {
      end = raw.size();
    }
    std::string entry = raw.substr(start, end - start);
    start = end + 1;
    if (entry.empty()) {
      continue;
    }
    std::size_t eq = entry.find('=');
    if (eq == std::string::npos || eq == 0 || eq + 1 >= entry.size()) {
      // Silently skip malformed entries: the env var is operator-supplied
      // and a typo should not crash the test binary at static-init time.
      continue;
    }
    std::string tag = entry.substr(0, eq);
    std::string path = entry.substr(eq + 1);
    if (tag.empty() || path.empty()) {
      continue;
    }
    out.emplace_back(std::move(tag), std::move(path));
  }
  return out;
}

namespace {

/// Loads `{case_id: reason}` from the divergence projection CMake emits.
///
/// The registry is a verification-time policy, so it must apply to
/// goldens that were captured before the entry existed. Baking the skip
/// into the golden -- which is all the generator can do -- would make a
/// newly adjudicated divergence wait on a fresh Excel capture to take
/// effect, which is the deadlock this whole path exists to avoid.
/// Read once; a missing or malformed file simply yields no skips.
const std::map<std::string, std::string>& registry_skips() {
  static const std::map<std::string, std::string> table = [] {
    std::map<std::string, std::string> out;
    const std::string path = FORMULON_WORKBOOK_ORACLE_SKIP_FILE;
    if (path.empty()) {
      return out;
    }
    auto parsed = parse_json_file(path);
    if (!parsed.has_value() || !parsed.value().is_object()) {
      return out;
    }
    const JsonValue* skips = parsed.value().find("skips");
    if (skips == nullptr || !skips->is_object()) {
      return out;
    }
    for (const auto& [case_id, reason] : skips->as_object()) {
      if (reason.is_string()) {
        out.emplace(case_id, reason.as_string());
      }
    }
    return out;
  }();
  return table;
}

}  // namespace

std::vector<WorkbookOracleCase> load_workbook_oracle_cases(const std::string& golden_dir,
                                                           const std::string& variant_tag) {
  std::vector<WorkbookOracleCase> out;
  if (golden_dir.empty()) {
    return out;
  }
  std::error_code ec;
  if (!std::filesystem::exists(golden_dir, ec) || !std::filesystem::is_directory(golden_dir, ec)) {
    return out;
  }

  std::vector<std::string> files;
  for (const auto& entry : std::filesystem::directory_iterator(golden_dir, ec)) {
    if (!entry.is_regular_file(ec)) {
      continue;
    }
    const auto& path = entry.path();
    if (path.extension() != ".json") {
      continue;
    }
    // Accept only `*.golden.json`; other files in the dir are ignored so
    // a README or .gitkeep does not trip the loader.
    const std::string fname = path.filename().string();
    const std::string marker = ".golden.json";
    if (fname.size() < marker.size() || fname.compare(fname.size() - marker.size(), marker.size(), marker) != 0) {
      continue;
    }
    files.push_back(path.string());
  }
  std::sort(files.begin(), files.end());

  for (const std::string& file_path : files) {
    auto parsed = parse_json_file(file_path);
    if (!parsed.has_value()) {
      out.push_back(make_load_error(file_path, "parse: " + parsed.error().message));
      continue;
    }
    const JsonValue& doc = parsed.value();
    if (!doc.is_object()) {
      out.push_back(make_load_error(file_path, "root is not an object"));
      continue;
    }
    const JsonValue* suite_v = doc.find("suite");
    const JsonValue* kind_v = doc.find("kind");
    const JsonValue* env_v = doc.find("environment");
    const JsonValue* cases_v = doc.find("cases");
    if (suite_v == nullptr || !suite_v->is_string() || cases_v == nullptr || !cases_v->is_array()) {
      out.push_back(make_load_error(file_path, "missing 'suite' or 'cases'"));
      continue;
    }
    if (kind_v == nullptr || !kind_v->is_string() || kind_v->as_string() != "workbook") {
      out.push_back(make_load_error(file_path, "missing or wrong 'kind' (expected \"workbook\")"));
      continue;
    }
    const std::string suite = suite_v->as_string();
    const JsonValue env = env_v != nullptr ? *env_v : JsonValue::make_object({});

    std::size_t fallback_idx = 0;
    for (const JsonValue& c : cases_v->as_array()) {
      if (!c.is_object())
        continue;
      const JsonValue* id_v = c.find("id");
      std::string case_id;
      if (id_v && id_v->is_string()) {
        case_id = id_v->as_string();
      } else {
        case_id = "case_" + std::to_string(fallback_idx);
      }
      ++fallback_idx;

      WorkbookOracleCase wc;
      wc.suite = suite;
      wc.case_id = case_id;
      wc.source_file = file_path;
      if (const JsonValue* spec_v = c.find("spec"); spec_v != nullptr) {
        wc.spec = *spec_v;
      } else {
        wc.spec = JsonValue::make_object({});
      }
      if (const JsonValue* expect_v = c.find("expect"); expect_v != nullptr) {
        wc.expect = *expect_v;
      } else {
        wc.expect = JsonValue::make_object({});
      }
      if (const JsonValue* skipped_v = c.find("skipped"); skipped_v != nullptr && skipped_v->is_string()) {
        wc.skipped_reason = skipped_v->as_string();
      } else if (const auto it = registry_skips().find(case_id); it != registry_skips().end()) {
        wc.skipped_reason = it->second;
      }
      wc.environment = env;
      wc.variant = variant_tag;
      out.push_back(std::move(wc));
    }
  }
  // Stamp variant on load-error sentinels too so a malformed variant
  // golden surfaces under its own parameter name.
  for (WorkbookOracleCase& wc : out) {
    if (wc.variant.empty())
      wc.variant = variant_tag;
  }
  return out;
}

}  // namespace oracle
}  // namespace tests
}  // namespace formulon
