//
// Chart oracle verifier: replays every case of the committed chart goldens
// (`golden_chart/<suite>.golden.json`) through Formulon and compares the
// projections the golden's suite names (`views`) with what Excel saved.
// The case spec travels inside the golden, so this binary reads no YAML.
//
// Formulon's side is built, saved and read back before projecting, so the
// comparison covers the cache refresh the writer performs on save. Numbers
// compare exactly: both sides carry the double Excel wrote. Setting
// FORMULON_CHART_GOLDEN_DIR in the environment replays a staged capture.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <filesystem>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "tests/oracle/json_reader.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace tests {
namespace oracle {
namespace chart_oracle {

// Defined in chart_model_dump.cpp.
Expected<Workbook, Error> read_workbook_bytes(const std::vector<std::uint8_t>& bytes, bool xlsb);
Expected<Workbook, Error> build_chart_case(const JsonValue& spec, const std::string& fixtures_dir);
Expected<std::string, Error> dump_charts_json(const Workbook& wb, const std::vector<std::string>& views);

namespace {

struct ChartCase {
  std::string suite;
  std::string id;
  std::vector<std::string> views;
  JsonValue spec;
  JsonValue expect;
  std::string skipped;  ///< Reason the golden recorded in place of `expect`.
};

/// `{case_id: reason}` from the divergence projection CMake emits, so an entry
/// registered after a capture applies without a new Excel session.
const std::map<std::string, std::string>& RegistrySkips() {
  static const std::map<std::string, std::string> table = [] {
    std::map<std::string, std::string> out;
    const std::string path = FORMULON_CHART_ORACLE_SKIP_FILE;
    if (path.empty()) {
      return out;
    }
    auto parsed = parse_json_file(path);
    const JsonValue* skips = parsed && parsed.value().is_object() ? parsed.value().find("skips") : nullptr;
    if (skips == nullptr || !skips->is_object()) {
      return out;
    }
    for (const auto& [id, reason] : skips->as_object()) {
      if (reason.is_string()) {
        out.emplace(id, reason.as_string());
      }
    }
    return out;
  }();
  return table;
}

std::vector<ChartCase> LoadCases() {
  std::vector<ChartCase> out;
  const char* override_dir = std::getenv("FORMULON_CHART_GOLDEN_DIR");
  const std::filesystem::path dir(override_dir != nullptr && *override_dir != '\0' ? override_dir
                                                                                   : FORMULON_CHART_GOLDEN_DIR);
  std::error_code ec;
  if (!std::filesystem::is_directory(dir, ec)) {
    return out;
  }
  std::vector<std::filesystem::path> files;
  for (const auto& entry : std::filesystem::directory_iterator(dir, ec)) {
    const std::string name = entry.path().filename().string();
    if (name.size() > 12 && name.compare(name.size() - 12, 12, ".golden.json") == 0) {
      files.push_back(entry.path());
    }
  }
  std::sort(files.begin(), files.end());
  for (const auto& path : files) {
    auto doc = parse_json_file(path.string());
    if (!doc) {
      ADD_FAILURE() << path << ": " << doc.error().message;
      continue;
    }
    const JsonValue& root = doc.value();
    const JsonValue* name = root.find("name");
    const JsonValue* views = root.find("views");
    const JsonValue* cases = root.find("cases");
    if (name == nullptr || views == nullptr || cases == nullptr) {
      ADD_FAILURE() << path << ": golden needs name, views and cases";
      continue;
    }
    std::vector<std::string> view_names;
    for (const JsonValue& v : views->as_array()) {
      view_names.push_back(v.as_string());
    }
    for (const JsonValue& c : cases->as_array()) {
      const JsonValue* id = c.find("id");
      const JsonValue* spec = c.find("spec");
      const JsonValue* expect = c.find("expect");
      const JsonValue* skipped = c.find("skipped");
      if (id == nullptr || spec == nullptr || (expect == nullptr && skipped == nullptr)) {
        ADD_FAILURE() << path << ": case without id / spec / expect";
        continue;
      }
      out.push_back(ChartCase{name->as_string(), id->as_string(), view_names, *spec,
                              expect != nullptr ? *expect : JsonValue(),
                              skipped != nullptr && skipped->is_string() ? skipped->as_string() : std::string()});
    }
  }
  return out;
}

const std::vector<ChartCase>& Cases() {
  static const std::vector<ChartCase> cases = LoadCases();
  return cases;
}

std::string Num(double v) {
  char buf[32];
  std::snprintf(buf, sizeof buf, "%.17g", v);
  return buf;
}

/// Appends a description of the first difference under `path`, if any.
void Diff(const JsonValue& want, const JsonValue& got, const std::string& path, std::vector<std::string>* out) {
  if (want.kind() != got.kind()) {
    out->push_back(path + ": kind differs");
    return;
  }
  switch (want.kind()) {
    case JsonKind::Null:
      return;
    case JsonKind::Bool:
      if (want.as_bool() != got.as_bool()) {
        out->push_back(path + ": " + (want.as_bool() ? "true" : "false") + " != " + (got.as_bool() ? "true" : "false"));
      }
      return;
    case JsonKind::Number:
      if (want.as_number() != got.as_number()) {
        out->push_back(path + ": " + Num(want.as_number()) + " != " + Num(got.as_number()));
      }
      return;
    case JsonKind::String:
      if (want.as_string() != got.as_string()) {
        out->push_back(path + ": \"" + want.as_string() + "\" != \"" + got.as_string() + "\"");
      }
      return;
    case JsonKind::Array: {
      const auto& a = want.as_array();
      const auto& b = got.as_array();
      if (a.size() != b.size()) {
        out->push_back(path + ": length " + std::to_string(a.size()) + " != " + std::to_string(b.size()));
        return;
      }
      for (std::size_t i = 0; i < a.size(); ++i) {
        Diff(a[i], b[i], path + "[" + std::to_string(i) + "]", out);
      }
      return;
    }
    case JsonKind::Object: {
      for (const auto& [key, value] : want.as_object()) {
        const JsonValue* other = got.find(key);
        if (other == nullptr) {
          out->push_back(path + "." + key + ": missing in Formulon");
          continue;
        }
        Diff(value, *other, path + "." + key, out);
      }
      for (const auto& [key, value] : got.as_object()) {
        if (want.find(key) == nullptr) {
          out->push_back(path + "." + key + ": not in golden");
        }
      }
      return;
    }
  }
}

class ChartOracleTest : public ::testing::TestWithParam<std::size_t> {};

TEST_P(ChartOracleTest, MatchesExcel) {
  const ChartCase& c = Cases()[GetParam()];
  const auto registry = RegistrySkips().find(c.id);
  if (registry != RegistrySkips().end()) {
    GTEST_SKIP() << "divergence.yaml skip-oracle: " << registry->second;
  }
  if (!c.skipped.empty()) {
    GTEST_SKIP() << "divergence.yaml skip-oracle: " << c.skipped;
  }
  auto built = build_chart_case(c.spec, FORMULON_CHART_FIXTURES_DIR);
  ASSERT_TRUE(static_cast<bool>(built)) << built.error().message;
  auto bytes = built.value().save();
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  auto reloaded = read_workbook_bytes(bytes.value(), false);
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  auto json = dump_charts_json(reloaded.value(), c.views);
  ASSERT_TRUE(static_cast<bool>(json)) << json.error().message;
  auto got = parse_json(json.value());
  ASSERT_TRUE(static_cast<bool>(got)) << got.error().message;

  std::vector<std::string> diffs;
  Diff(c.expect, got.value(), "expect", &diffs);
  std::string report;
  for (const std::string& d : diffs) {
    report += "\n  " + d;
  }
  EXPECT_TRUE(diffs.empty()) << c.suite << "/" << c.id << ":" << report;
}

std::string CaseName(const ::testing::TestParamInfo<std::size_t>& info) {
  const ChartCase& c = Cases()[info.param];
  std::string name = c.suite + "_" + c.id;
  for (char& ch : name) {
    if (!((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9'))) {
      ch = '_';
    }
  }
  return name;
}

std::vector<std::size_t> CaseIndices() {
  std::vector<std::size_t> out(Cases().size());
  for (std::size_t i = 0; i < out.size(); ++i) {
    out[i] = i;
  }
  return out;
}

INSTANTIATE_TEST_SUITE_P(Chart, ChartOracleTest, ::testing::ValuesIn(CaseIndices()), CaseName);
GTEST_ALLOW_UNINSTANTIATED_PARAMETERIZED_TEST(ChartOracleTest);

}  // namespace
}  // namespace chart_oracle
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
