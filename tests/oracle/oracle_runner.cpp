// Copyright 2026 libraz. Licensed under the MIT License.

#include "tests/oracle/oracle_runner.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <filesystem>
#include <string>
#include <string_view>
#include <vector>

#include "tests/oracle/json_reader.h"

#ifndef FORMULON_ORACLE_GOLDEN_DIR
#define FORMULON_ORACLE_GOLDEN_DIR ""
#endif

namespace formulon {
namespace tests {
namespace oracle {

namespace {

// Emits one synthetic case whose id is `<load-error>` so the failure shows
// up as a failing TEST_P parameter rather than silently dropping the file.
OracleCase make_load_error(const std::string& path, const std::string& detail) {
  OracleCase c;
  c.suite = std::filesystem::path(path).stem().string();
  // Strip the ".golden" left over after `.golden.json` → `.golden` stem.
  if (c.suite.size() >= 7 &&
      c.suite.compare(c.suite.size() - 7, 7, ".golden") == 0) {
    c.suite.resize(c.suite.size() - 7);
  }
  c.case_id = "<load-error>";
  c.source_file = path;

  std::map<std::string, JsonValue> err_case;
  err_case.emplace("id", JsonValue::make_string("<load-error>"));
  err_case.emplace("error", JsonValue::make_string(detail));
  c.raw_case = JsonValue::make_object(std::move(err_case));
  c.environment = JsonValue::make_object({});
  return c;
}

}  // namespace

bool a1_to_row_col(const std::string& a1, std::uint32_t* out_row,
                   std::uint32_t* out_col) {
  if (a1.empty() || out_row == nullptr || out_col == nullptr) return false;
  std::size_t i = 0;
  std::uint32_t col = 0;
  // Column letters: A..Z, AA..ZZ, ... (base-26 with digits 1..26, not 0..25).
  while (i < a1.size()) {
    char c = a1[i];
    if (c >= 'A' && c <= 'Z') {
      col = col * 26 + static_cast<std::uint32_t>(c - 'A' + 1);
      ++i;
    } else if (c >= 'a' && c <= 'z') {
      col = col * 26 + static_cast<std::uint32_t>(c - 'a' + 1);
      ++i;
    } else {
      break;
    }
  }
  if (i == 0 || i == a1.size()) return false;  // missing letters or missing row
  std::uint64_t row = 0;
  for (; i < a1.size(); ++i) {
    char c = a1[i];
    if (c < '0' || c > '9') return false;
    row = row * 10 + static_cast<std::uint64_t>(c - '0');
    if (row > 1048576ULL) return false;  // Excel max row
  }
  if (row == 0 || col == 0) return false;
  *out_row = static_cast<std::uint32_t>(row - 1);
  *out_col = col - 1;
  return true;
}

std::string configured_golden_dir() {
  return std::string{FORMULON_ORACLE_GOLDEN_DIR};
}

std::vector<OracleCase> load_oracle_cases(const std::string& golden_dir) {
  std::vector<OracleCase> out;
  if (golden_dir.empty()) return out;
  std::error_code ec;
  if (!std::filesystem::exists(golden_dir, ec) ||
      !std::filesystem::is_directory(golden_dir, ec)) {
    return out;
  }

  std::vector<std::string> files;
  for (const auto& entry :
       std::filesystem::directory_iterator(golden_dir, ec)) {
    if (!entry.is_regular_file(ec)) continue;
    const auto& path = entry.path();
    if (path.extension() != ".json") continue;
    // Accept only `*.golden.json`; other files in the dir are ignored so
    // things like a README or .gitkeep don't trip the loader.
    const std::string fname = path.filename().string();
    const std::string marker = ".golden.json";
    if (fname.size() < marker.size() ||
        fname.compare(fname.size() - marker.size(), marker.size(), marker) !=
            0) {
      continue;
    }
    files.push_back(path.string());
  }
  std::sort(files.begin(), files.end());

  for (const std::string& file_path : files) {
    auto parsed = parse_json_file(file_path);
    if (!parsed.has_value()) {
      out.push_back(
          make_load_error(file_path, "parse: " + parsed.error().message));
      continue;
    }
    const JsonValue& doc = parsed.value();
    if (!doc.is_object()) {
      out.push_back(make_load_error(file_path, "root is not an object"));
      continue;
    }
    const JsonValue* suite_v = doc.find("suite");
    const JsonValue* env_v = doc.find("environment");
    const JsonValue* cases_v = doc.find("cases");
    if (suite_v == nullptr || !suite_v->is_string() || cases_v == nullptr ||
        !cases_v->is_array()) {
      out.push_back(make_load_error(file_path, "missing 'suite' or 'cases'"));
      continue;
    }
    const std::string suite = suite_v->as_string();

    // Pull suite-level tolerance defaults so cases can omit them when they
    // agree with the suite default.
    double tol_abs = 0.0;
    double tol_rel = 0.0;
    if (const JsonValue* tol = doc.find("tolerance"); tol && tol->is_object()) {
      if (const JsonValue* a = tol->find("abs"); a && a->is_number()) {
        tol_abs = a->as_number();
      }
      if (const JsonValue* r = tol->find("rel"); r && r->is_number()) {
        tol_rel = r->as_number();
      }
    }

    const JsonValue env = env_v ? *env_v : JsonValue::make_object({});
    std::size_t fallback_idx = 0;
    for (const JsonValue& c : cases_v->as_array()) {
      if (!c.is_object()) continue;
      const JsonValue* id_v = c.find("id");
      std::string case_id;
      if (id_v && id_v->is_string()) {
        case_id = id_v->as_string();
      } else {
        case_id = "case_" + std::to_string(fallback_idx);
      }
      ++fallback_idx;

      OracleCase oc;
      oc.suite = suite;
      oc.case_id = case_id;
      oc.source_file = file_path;
      oc.raw_case = c;
      oc.environment = env;
      oc.tolerance_abs = tol_abs;
      oc.tolerance_rel = tol_rel;
      if (const JsonValue* tol = c.find("tolerance");
          tol && tol->is_object()) {
        if (const JsonValue* a = tol->find("abs"); a && a->is_number()) {
          oc.tolerance_abs = a->as_number();
        }
        if (const JsonValue* r = tol->find("rel"); r && r->is_number()) {
          oc.tolerance_rel = r->as_number();
        }
      }
      if (const JsonValue* cm = c.find("compare_mode");
          cm && cm->is_string()) {
        oc.compare_mode = cm->as_string();
      }
      out.push_back(std::move(oc));
    }
  }
  return out;
}

}  // namespace oracle
}  // namespace tests
}  // namespace formulon
