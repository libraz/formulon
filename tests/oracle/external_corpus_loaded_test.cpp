//
// Loaded-file mode over the external corpus: every vendored workbook under
// `external/*/fixtures/` is read through the package reader -- once on the
// DOM sheet path and once on the streaming one -- recalculated in full, and
// each formula cell is compared with the value Excel cached for it. This is
// the track that reads the files, and it covers what the formula-entry import
// leaves out: the per-sheet case cap, defined names, structured references
// and array formulas.
//
// Left out, each counted in the test output:
//   * a fixture not saved by Excel (`kNotSavedByExcel`);
//   * a cell with no cached value;
//   * a cell `tests/ironcalc_divergence.yaml` lists, by the importer's
//     `<suite>.<addr>` id whether or not the importer took the cell in;
//   * a cell depending at any depth on RAND / RANDBETWEEN / RANDARRAY;
//   * a cell depending on TODAY or NOW when the saving instant cannot be
//     recovered from a bare `=TODAY()` / `=NOW()` cache: the day from either,
//     the time of day only from a NOW cached on a whole second.
// Profile and tolerance come from the sheet's external golden; a sheet the
// importer produced no golden for takes the same rules from the workbook.

#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iterator>
#include <map>
#include <set>
#include <sstream>
#include <string>
#include <string_view>
#include <tuple>
#include <utility>
#include <vector>

#include "eval/dep_graph.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "tests/oracle/json_reader.h"
#include "tests/oracle/oracle_runner.h"
#include "tests/oracle/oracle_test_support.h"
#include "utils/a1_ref.h"
#include "utils/date_time.h"
#include "value.h"
#include "workbook.h"

namespace formulon::tests::oracle {
namespace {

using support::compare_value;
using support::fold_param_name;

constexpr double kDefaultAbsTolerance = 5e-8;
constexpr double kDefaultRelTolerance = 1e-12;

// What the external golden of one sheet says about it.
struct SheetRules {
  std::string profile_id;
  double tol_abs = kDefaultAbsTolerance;
  double tol_rel = kDefaultRelTolerance;
};

// (fixture path relative to its `fixtures/` directory, sheet name) -> rules.
using RulesIndex = std::map<std::pair<std::string, std::string>, SheetRules>;

struct CorpusFile {
  std::string corpus;    // e.g. "ironcalc"
  std::string relative;  // path under the corpus's `fixtures/`
  std::string path;
  bool streaming = false;
};

void PrintTo(const CorpusFile& file, std::ostream* os) {
  *os << file.relative << (file.streaming ? " (streaming)" : " (dom)");
}

// `Derived from <corpus> fixture <relative>, sheet '<sheet>'.`
bool parse_golden_origin(const std::string& description, std::string* relative, std::string* sheet) {
  const std::string_view marker = " fixture ";
  const std::size_t at = description.find(marker);
  const std::size_t sep = description.rfind(", sheet '");
  if (at == std::string::npos || sep == std::string::npos || sep < at || description.size() < 2U ||
      description.compare(description.size() - 2U, 2U, "'.") != 0) {
    return false;
  }
  *relative = description.substr(at + marker.size(), sep - at - marker.size());
  const std::size_t sheet_begin = sep + 9U;
  *sheet = description.substr(sheet_begin, description.size() - 2U - sheet_begin);
  return true;
}

// The goldens of one fixture: the importer names each
// `<relative, '/' as "__", no extension>__<sheet>.golden.json`, and the
// description is checked as well because one fixture's stem can prefix
// another's.
RulesIndex rules_for(const CorpusFile& file) {
  std::string stem = file.relative.substr(0, file.relative.size() - 5U);
  for (std::size_t at = stem.find('/'); at != std::string::npos; at = stem.find('/', at + 2U)) {
    stem.replace(at, 1U, "__");
  }
  stem += "__";
  RulesIndex index;
  const std::filesystem::path golden_dir = std::filesystem::path(FORMULON_EXTERNAL_CORPUS_DIR) / file.corpus / "golden";
  std::error_code ec;
  for (const auto& entry : std::filesystem::directory_iterator(golden_dir, ec)) {
    const std::string name = entry.path().filename().string();
    if (name.compare(0, stem.size(), stem) != 0) {
      continue;
    }
    const auto doc = parse_json_file(entry.path().string());
    if (!doc) {
      continue;
    }
    const JsonValue* description = doc.value().find("description");
    std::string relative;
    std::string sheet;
    if (description == nullptr || !description->is_string() ||
        !parse_golden_origin(description->as_string(), &relative, &sheet) || relative != file.relative) {
      continue;
    }
    SheetRules& rules = index[{relative, sheet}];
    if (const JsonValue* tol = doc.value().find("tolerance"); tol != nullptr && tol->is_object()) {
      if (const JsonValue* v = tol->find("abs"); v != nullptr && v->is_number()) {
        rules.tol_abs = v->as_number();
      }
      if (const JsonValue* v = tol->find("rel"); v != nullptr && v->is_number()) {
        rules.tol_rel = v->as_number();
      }
    }
    if (const JsonValue* env = doc.value().find("environment"); env != nullptr && env->is_object()) {
      if (const JsonValue* profile = env->find("profile_id"); profile != nullptr && profile->is_string()) {
        rules.profile_id = profile->as_string();
      }
    }
  }
  return index;
}

// `ironcalc_import._scrub_filename`: every run of characters outside
// [A-Za-z0-9._-] becomes one '_', then '.' and '_' are trimmed from both ends.
std::string scrub(std::string_view part) {
  std::string out;
  bool in_run = false;
  for (const char ch : part) {
    const auto u = static_cast<unsigned char>(ch);
    if (std::isalnum(u) != 0 && u < 0x80U) {
      out += ch;
      in_run = false;
    } else if (ch == '.' || ch == '_' || ch == '-') {
      out += ch;
      in_run = false;
    } else if (!in_run) {
      out += '_';
      in_run = true;
    }
  }
  const std::size_t begin = out.find_first_not_of("._");
  if (begin == std::string::npos) {
    return "sheet";
  }
  return out.substr(begin, out.find_last_not_of("._") - begin + 1U);
}

// `ironcalc_import.suite_name`: the suite a fixture sheet's case ids hang off.
std::string suite_name(const std::string& relative, std::string_view sheet) {
  const std::size_t slash = relative.rfind('/');
  const std::string file = slash == std::string::npos ? relative : relative.substr(slash + 1U);
  std::string dir = slash == std::string::npos ? "." : relative.substr(0, slash);
  for (std::size_t at = dir.find('/'); at != std::string::npos; at = dir.find('/', at + 2U)) {
    dir.replace(at, 1U, "__");
  }
  const std::size_t dot = file.rfind('.');
  return "ironcalc_" + scrub(dir) + "_" + scrub(file.substr(0, dot)) + "_" + scrub(sheet);
}

std::vector<std::uint8_t> read_bytes(const std::string& path) {
  std::ifstream in(path, std::ios::binary);
  return std::vector<std::uint8_t>(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
}

// The saving application `docProps/app.xml` records, or empty.
std::string package_application(const std::vector<std::uint8_t>& bytes) {
  io::ZipReader zip;
  if (!zip.open(io::ByteSpan{bytes.data(), bytes.size()})) {
    return {};
  }
  auto app = zip.read_entry("docProps/app.xml");
  if (!app) {
    return {};
  }
  const std::string xml(app.value().begin(), app.value().end());
  const std::size_t begin = xml.find("<Application>");
  const std::size_t end = xml.find("</Application>");
  if (begin == std::string::npos || end == std::string::npos || end < begin) {
    return {};
  }
  return xml.substr(begin + 13U, end - begin - 13U);
}

bool saved_by_excel(const std::string& application) {
  return application.find("Excel") != std::string::npos;
}

// A workbook-level profile for a file with no golden: Mac builds of Excel to
// the Mac profile, other Excel builds to the Windows one.
std::string profile_from_application(const std::string& application) {
  if (!saved_by_excel(application)) {
    return {};
  }
  return application.find("Macintosh") != std::string::npos ? "mac-365-en_US" : "win-365-en_US";
}

// Fixtures whose package records no Excel application. Their cached values
// were not computed by Excel, so they are not evidence of Excel's answer.
struct ExcludedFile {
  std::string_view relative;
  std::string_view reason;
};
constexpr ExcludedFile kNotSavedByExcel[] = {
    {"calc_tests/LOOKUP_AND_REFERENCE/XMATCH_arrays.xlsx", "written by a non-Excel spreadsheet"},
    {"calc_tests/STATISTICAL/GEOMEAN.xlsx", "no application record; errors cached as text"},
    {"calc_tests/array_in_scalar_repro.xlsx", "no application record"},
    {"docs/CHOOSE.xlsx", "written by a non-Excel spreadsheet"},
};

const ExcludedFile* excluded(const std::string& relative) {
  for (const ExcludedFile& file : kNotSavedByExcel) {
    if (file.relative == relative) {
      return &file;
    }
  }
  return nullptr;
}

// Which volatile families a formula calls.
struct VolatileCalls {
  bool random = false;
  bool now = false;
  bool today = false;
};

VolatileCalls volatile_calls(std::string_view formula) {
  std::string upper(formula);
  std::transform(upper.begin(), upper.end(), upper.begin(),
                 [](unsigned char ch) { return static_cast<char>(std::toupper(ch)); });
  static constexpr std::string_view kNames[] = {"RANDBETWEEN", "RANDARRAY", "RAND", "TODAY", "NOW"};
  VolatileCalls calls;
  for (std::size_t i = 0; i < std::size(kNames); ++i) {
    const std::string_view name = kNames[i];
    for (std::size_t at = upper.find(name); at != std::string::npos; at = upper.find(name, at + 1U)) {
      const bool word_start =
          at == 0U || (std::isalnum(static_cast<unsigned char>(upper[at - 1U])) == 0 && upper[at - 1U] != '_');
      std::size_t next = at + name.size();
      while (next < upper.size() && upper[next] == ' ') {
        ++next;
      }
      if (word_start && next < upper.size() && upper[next] == '(') {
        (i < 3U ? calls.random : i == 3U ? calls.today : calls.now) = true;
      }
    }
  }
  return calls;
}

// True when `formula` is exactly `=<name>()`, so its cached value is the
// clock reading itself.
bool is_bare_call(std::string_view formula, std::string_view name) {
  std::string compact;
  for (const char ch : formula) {
    if (ch != ' ') {
      compact.push_back(static_cast<char>(std::toupper(static_cast<unsigned char>(ch))));
    }
  }
  return compact == "=" + std::string(name) + "()";
}

using CellKey = std::tuple<std::size_t, std::uint32_t, std::uint32_t>;

// Every cell depending on `seeds` at any depth, the seeds included.
std::set<CellKey> dependent_closure(const Workbook& wb, const std::vector<CellKey>& seeds) {
  const eval::RecalcEngine& engine = wb.recalc_engine();
  const eval::DepGraph& graph = engine.dep_graph();
  std::set<CellKey> seen(seeds.begin(), seeds.end());
  std::vector<CellKey> queue(seeds.begin(), seeds.end());
  while (!queue.empty()) {
    const auto [sheet, row, col] = queue.back();
    queue.pop_back();
    const eval::CellNodeId node{static_cast<std::uint16_t>(sheet), row, col};
    std::vector<eval::CellNodeId> next = graph.dependents_of(node);
    const std::vector<eval::CellNodeId> compact = engine.compact_range_dependents_of(node);
    next.insert(next.end(), compact.begin(), compact.end());
    for (const eval::CellNodeId dep : next) {
      if (eval::is_range_node(dep)) {
        continue;
      }
      const CellKey key{dep.sheet_id, dep.row, dep.col};
      if (seen.insert(key).second) {
        queue.push_back(key);
      }
    }
  }
  return seen;
}

// The cached value as the golden record the comparator reads.
JsonValue expect_record(const Value& v) {
  if (v.is_number()) {
    return JsonValue::make_object(
        {{"kind", JsonValue::make_string("number")}, {"value", JsonValue::make_number(v.as_number())}});
  }
  if (v.is_text()) {
    return JsonValue::make_object(
        {{"kind", JsonValue::make_string("text")}, {"value", JsonValue::make_string(std::string(v.as_text()))}});
  }
  if (v.is_boolean()) {
    return JsonValue::make_object(
        {{"kind", JsonValue::make_string("bool")}, {"value", JsonValue::make_bool(v.as_boolean())}});
  }
  if (v.is_error()) {
    return JsonValue::make_object({{"kind", JsonValue::make_string("error")},
                                   {"code", JsonValue::make_string(std::string(display_name(v.as_error())))}});
  }
  return JsonValue::make_object({{"kind", JsonValue::make_string("blank")}});
}

std::vector<CorpusFile> corpus_files() {
  std::vector<CorpusFile> out;
  namespace fs = std::filesystem;
  const fs::path root(FORMULON_EXTERNAL_CORPUS_DIR);
  std::error_code ec;
  for (const fs::directory_entry& corpus : fs::directory_iterator(root, ec)) {
    const fs::path fixtures = corpus.path() / "fixtures";
    if (!fs::is_directory(fixtures, ec)) {
      continue;
    }
    std::vector<fs::path> files;
    for (const fs::directory_entry& entry : fs::recursive_directory_iterator(fixtures, ec)) {
      if (entry.is_regular_file() && entry.path().extension() == ".xlsx") {
        files.push_back(entry.path());
      }
    }
    std::sort(files.begin(), files.end());
    for (const fs::path& file : files) {
      for (const bool streaming : {false, true}) {
        out.push_back({corpus.path().filename().string(), file.lexically_relative(fixtures).generic_string(),
                       file.string(), streaming});
      }
    }
  }
  return out;
}

class ExternalCorpusLoaded : public ::testing::TestWithParam<CorpusFile> {};

TEST_P(ExternalCorpusLoaded, CachedValuesMatch) {
  const CorpusFile& file = GetParam();
  const std::vector<std::uint8_t> bytes = read_bytes(file.path);
  ASSERT_FALSE(bytes.empty()) << file.path;
  const std::string application = package_application(bytes);
  if (const ExcludedFile* skip = excluded(file.relative); skip != nullptr) {
    ASSERT_FALSE(saved_by_excel(application))
        << file.relative << " is listed as not saved by Excel, but " << application << " saved it";
    GTEST_SKIP() << skip->reason;
  }
  ASSERT_TRUE(saved_by_excel(application))
      << file.relative << ": saved by '" << application << "', not Excel; list it in kNotSavedByExcel";
  auto read = io::internal::ReadOoxmlWithSaxThresholdForTesting(io::ByteSpan{bytes.data(), bytes.size()},
                                                                file.streaming ? 0U : static_cast<std::size_t>(-1));
  ASSERT_TRUE(static_cast<bool>(read)) << file.relative << ": " << read.error().message;
  Workbook& wb = read.value().workbook;

  const RulesIndex index = rules_for(file);
  const auto rules_of = [&](std::size_t sheet) -> const SheetRules* {
    const auto it = index.find({file.relative, std::string(wb.sheet(sheet).name())});
    return it == index.end() ? nullptr : &it->second;
  };

  std::string profile_id;
  bool from_golden = false;
  for (std::size_t s = 0; s < wb.sheet_count() && !from_golden; ++s) {
    if (const SheetRules* rules = rules_of(s); rules != nullptr) {
      profile_id = rules->profile_id;
      from_golden = true;
    }
  }
  if (!from_golden) {
    profile_id = profile_from_application(application);
  }
  ExcelProfile profile;
  if (!profile_id.empty()) {
    ASSERT_TRUE(parse_excel_profile_id(profile_id, &profile)) << profile_id;
    wb.set_excel_profile(profile);
  }

  struct Cached {
    CellKey key;
    std::string formula;
    Value value;
  };
  std::vector<Cached> cached;
  std::vector<CellKey> random_seeds;
  std::vector<CellKey> now_seeds;
  std::vector<CellKey> today_seeds;
  // The saving instant, recovered from bare NOW() / TODAY() caches: the day
  // from either, the time only from a NOW() cached on a whole second.
  std::set<double> days;
  std::set<double> instants;
  bool now_unrecoverable = false;
  Arena keep;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const Sheet& sheet = wb.sheet(s);
    for (const CellAddress address : sheet.all_formula_cells()) {
      Sheet::CellRead read_cell;
      sheet.read_formula_cell(address.row, address.col, read_cell);
      const CellKey key{s, address.row, address.col};
      cached.push_back({key, std::string(read_cell.formula_text()), adopt_text_into(keep, read_cell.value())});
      const VolatileCalls calls = volatile_calls(read_cell.formula_text());
      if (calls.random) {
        random_seeds.push_back(key);
      } else if (calls.now) {
        now_seeds.push_back(key);
      } else if (calls.today) {
        today_seeds.push_back(key);
      }
      const Value& v = cached.back().value;
      if (!calls.random && v.is_number()) {
        const double serial = v.as_number();
        if (is_bare_call(read_cell.formula_text(), "TODAY") && serial == std::floor(serial)) {
          days.insert(serial);
        } else if (is_bare_call(read_cell.formula_text(), "NOW")) {
          days.insert(std::floor(serial));
          const double seconds = serial * 86400.0;
          if (std::fabs(seconds - std::round(seconds)) < 1e-3) {
            instants.insert(std::round(seconds));
          } else {
            now_unrecoverable = true;
          }
        }
      }
    }
  }

  const bool date_pinned = days.size() == 1U;
  const bool time_pinned = date_pinned && instants.size() == 1U && !now_unrecoverable &&
                           std::floor(*instants.begin() / 86400.0) == *days.begin();
  if (date_pinned) {
    date_time::CivilTime pin{date_time::ymd_from_serial(*days.begin(), wb.date1904()), {0U, 0U, 0U}};
    if (time_pinned) {
      const auto second_of_day = static_cast<unsigned>(*instants.begin() - *days.begin() * 86400.0);
      pin.time = {second_of_day / 3600U, (second_of_day % 3600U) / 60U, second_of_day % 60U};
    }
    wb.set_pinned_now(pin);
  }

  ASSERT_TRUE(static_cast<bool>(wb.recalc_full(eval::default_registry()))) << file.relative;

  // A volatile anchor's spill is volatile too.
  for (std::vector<CellKey>* seeds : {&random_seeds, &now_seeds, &today_seeds}) {
    const std::size_t anchors = seeds->size();
    for (std::size_t i = 0; i < anchors; ++i) {
      const auto [s, row, col] = (*seeds)[i];
      Arena scratch;
      std::vector<Value> cells;
      std::uint32_t rows = 0;
      std::uint32_t cols = 0;
      if (wb.sheet(s).read_spill_region_at_anchor(row, col, scratch, cells, &rows, &cols)) {
        for (std::uint32_t r = 0; r < rows; ++r) {
          for (std::uint32_t c = 0; c < cols; ++c) {
            if (r != 0U || c != 0U) {
              seeds->push_back({s, row + r, col + c});
            }
          }
        }
      }
    }
  }
  const std::set<CellKey> random_cells = dependent_closure(wb, random_seeds);
  const std::set<CellKey> now_cells = dependent_closure(wb, now_seeds);
  const std::set<CellKey> today_cells = dependent_closure(wb, today_seeds);

  std::size_t compared = 0;
  std::size_t no_cache = 0;
  std::size_t registry_skips = 0;
  std::size_t volatile_random = 0;
  std::size_t volatile_clock = 0;
  std::size_t clock_pinned = 0;
  std::ostringstream failures;
  std::size_t failed = 0;
  for (const Cached& cell : cached) {
    const auto [s, row, col] = cell.key;
    if (cell.value.is_blank()) {
      ++no_cache;
      continue;
    }
    if (random_cells.count(cell.key) != 0U) {
      ++volatile_random;
      continue;
    }
    const bool on_now = now_cells.count(cell.key) != 0U;
    const bool on_today = today_cells.count(cell.key) != 0U;
    if ((on_now && !time_pinned) || (on_today && !date_pinned)) {
      ++volatile_clock;
      continue;
    }
    if (on_now || on_today) {
      ++clock_pinned;
    }
    const std::string address = a1::encode_a1(row, col);
    if (oracle_skip_registry().count(suite_name(file.relative, wb.sheet(s).name()) + "." + address) != 0U) {
      ++registry_skips;
      continue;
    }
    const SheetRules* rules = rules_of(s);
    ++compared;
    const Value got = wb.sheet(s).resolve_cell_value(row, col);
    const std::string diff =
        compare_value(expect_record(cell.value), got, rules != nullptr ? rules->tol_abs : kDefaultAbsTolerance,
                      rules != nullptr ? rules->tol_rel : kDefaultRelTolerance, "");
    if (!diff.empty()) {
      ++failed;
      failures << "  " << file.relative << " | " << wb.sheet(s).name() << " | " << address << " | " << cell.formula
               << " | expected " << support::format_value(cell.value) << " | got " << support::format_value(got)
               << "\n";
    }
  }
  RecordProperty("compared", static_cast<int>(compared));
  RecordProperty("no_cached_value", static_cast<int>(no_cache));
  RecordProperty("registry_skips", static_cast<int>(registry_skips));
  RecordProperty("volatile_random", static_cast<int>(volatile_random));
  RecordProperty("volatile_clock", static_cast<int>(volatile_clock));
  RecordProperty("clock_pinned", static_cast<int>(clock_pinned));
  EXPECT_EQ(failed, 0U) << failed << " of " << compared << " compared cells differ (" << no_cache
                        << " without a cached value, " << registry_skips << " registry skips, " << volatile_random
                        << " random, " << volatile_clock << " clock, " << clock_pinned
                        << " compared on the pinned clock)\n"
                        << failures.str();
}

std::string ParamName(const ::testing::TestParamInfo<CorpusFile>& info) {
  std::string name = info.param.corpus + "_" + info.param.relative;
  const std::size_t dot = name.rfind(".xlsx");
  if (dot != std::string::npos) {
    name.erase(dot);
  }
  return fold_param_name(name + (info.param.streaming ? "_sax" : "_dom"));
}

INSTANTIATE_TEST_SUITE_P(Corpus, ExternalCorpusLoaded, ::testing::ValuesIn(corpus_files()), ParamName);

}  // namespace
}  // namespace formulon::tests::oracle
