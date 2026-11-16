// Corpus round trip: every .xlsx / .xlsb fixture under tests/fixtures/excel/ is
// loaded, saved as .xlsx and as .xlsb, and both outputs are reloaded. The
// reloaded model must keep the identities and counts of the first load and
// each output must pass the package validators. Each unedited save reports
// every part and relationship it counts as dropped (a drop_* failure naming
// the part path, or rel:<target>), so a part the model holds that a save
// counts as lost fails here rather than downstream.
//
// The corpus is sharded by top-level fixture directory (loose files form the
// "root" shard), a large directory in numbered chunks. Failures are "<relpath> <stage> <detail>"; the ones the
// committed manifest (fixtures/excel/corpus_roundtrip_manifest.txt) names are
// expected, and an exact entry whose failure did not occur fails the shard.
// Each unexpected failure also prints a "manifest-line: " candidate entry, so a
// commit adding fixtures can collect its lines with
//   formulon_unit_tests --gtest_filter='*CorpusRoundtrip*' | grep '^manifest-line: '
// and fill in the kind's reason.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <fstream>
#include <iterator>
#include <map>
#include <ostream>
#include <set>
#include <sstream>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/sparkline_groups.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/user_book_views.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "support/xlsb_package_validator.h"
#include "support/xlsx_package_validator.h"
#include "utils/structured_log.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

namespace fs = std::filesystem;

constexpr std::size_t kMinCorpusFiles = 1000;
constexpr const char* kRootShard = "root";
constexpr const char* kManifestName = "corpus_roundtrip_manifest.txt";

const fs::path& CorpusRoot() {
  static const fs::path root = fs::path(FORMULON_FIXTURES_DIR) / "excel";
  return root;
}

bool IsCorpusFile(const fs::path& path) {
  const std::string ext = path.extension().string();
  return ext == ".xlsx" || ext == ".xlsb";
}

/// The most files one shard checks: the slowest shard takes about 8 s in the Debug build.
constexpr std::size_t kMaxShardFiles = 60;

struct Shard {
  std::string name;
  std::vector<std::string> files;  // relative to the corpus root, '/'-separated
};

void PrintTo(const Shard& shard, std::ostream* os) {
  *os << shard.name;
}

std::string SafeName(const std::string& raw) {
  std::string out = raw;
  for (char& c : out) {
    const bool alnum = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
    if (!alnum) {
      c = '_';
    }
  }
  return out.empty() ? std::string("_") : out;
}

std::vector<Shard> ComputeShards() {
  std::map<std::string, std::vector<std::string>> by_group;
  std::error_code ec;
  if (!fs::is_directory(CorpusRoot(), ec)) {
    return {};
  }
  for (const auto& entry : fs::recursive_directory_iterator(CorpusRoot(), ec)) {
    if (!entry.is_regular_file() || !IsCorpusFile(entry.path())) {
      continue;
    }
    const std::string rel = fs::relative(entry.path(), CorpusRoot()).generic_string();
    const std::size_t slash = rel.find('/');
    by_group[slash == std::string::npos ? std::string() : rel.substr(0, slash)].push_back(rel);
  }
  std::vector<Shard> shards;
  std::set<std::string> used;
  for (auto& [group, files] : by_group) {
    std::sort(files.begin(), files.end());
    const std::string base = group.empty() ? std::string(kRootShard) : SafeName(group);
    // A large group runs as numbered chunks so that each test stays well inside the fast-tier timeout in Debug.
    const std::size_t chunks = (files.size() + kMaxShardFiles - 1U) / kMaxShardFiles;
    for (std::size_t c = 0; c < chunks; ++c) {
      std::string name = chunks == 1U ? base : base + "_" + std::to_string(c + 1U);
      while (!used.insert(name).second) {
        name += "_";
      }
      const auto first = files.begin() + static_cast<std::ptrdiff_t>(c * kMaxShardFiles);
      const auto last = files.begin() + static_cast<std::ptrdiff_t>(std::min(files.size(), (c + 1U) * kMaxShardFiles));
      shards.push_back(Shard{name, std::vector<std::string>(first, last)});
    }
  }
  return shards;
}

// ---- manifest ----

enum class Kind { kExcelRefuses, kMirrors, kSaveRefused, kKnownGap, kUnmodelledDropped, kDamaged };

struct Entry {
  std::string path;
  std::string stage;
  std::string detail;
  Kind kind = Kind::kKnownGap;
  std::string line;
  bool matched = false;
};

bool ParseKind(const std::string& s, Kind& kind) {
  static const std::pair<const char*, Kind> kKinds[] = {{"excel_refuses", Kind::kExcelRefuses},
                                                        {"mirrors_excel_refused_source", Kind::kMirrors},
                                                        {"save_refused", Kind::kSaveRefused},
                                                        {"known_gap", Kind::kKnownGap},
                                                        {"unmodelled_dropped", Kind::kUnmodelledDropped},
                                                        {"damaged", Kind::kDamaged}};
  for (const auto& [text, value] : kKinds) {
    if (s == text) {
      kind = value;
      return true;
    }
  }
  return false;
}

const std::set<std::string>& Stages() {
  static const std::set<std::string> stages = {"load",       "save_xlsx",     "save_xlsb",    "drop_xlsx",
                                               "drop_xlsb",  "reload_xlsx",   "reload_xlsb",  "model_xlsx",
                                               "model_xlsb", "validate_xlsx", "validate_xlsb"};
  return stages;
}

struct Manifest {
  std::vector<Entry> entries;
  std::vector<std::string> problems;  // malformed lines and kind/stage mismatches
};

Manifest LoadManifest() {
  Manifest manifest;
  std::ifstream in(CorpusRoot() / kManifestName);
  if (!in) {
    manifest.problems.push_back(std::string("cannot open ") + kManifestName);
    return manifest;
  }
  std::string raw;
  while (std::getline(in, raw)) {
    const std::string body = raw.substr(0, raw.find('#'));
    std::istringstream fields(body);
    Entry e;
    std::string kind;
    if (!(fields >> e.path)) {
      continue;
    }
    e.line = raw;
    if (!(fields >> e.stage >> e.detail >> kind) || !ParseKind(kind, e.kind)) {
      manifest.problems.push_back("malformed manifest line: " + raw);
      continue;
    }
    const bool file_level = e.kind == Kind::kExcelRefuses || e.kind == Kind::kDamaged;
    if (file_level != (e.stage == "*" && e.detail == "*")) {
      manifest.problems.push_back("stage/detail do not fit the kind: " + raw);
    } else if (!file_level && Stages().count(e.stage) == 0) {
      manifest.problems.push_back("unknown stage: " + raw);
    } else if (e.kind == Kind::kMirrors && e.stage.rfind("validate_", 0) != 0) {
      manifest.problems.push_back("mirrors_excel_refused_source needs a validate_ stage: " + raw);
    } else if (e.kind == Kind::kSaveRefused && e.stage.rfind("save_", 0) != 0) {
      manifest.problems.push_back("save_refused needs a save_ stage: " + raw);
    } else if ((e.kind == Kind::kUnmodelledDropped) != (e.stage.rfind("drop_", 0) == 0)) {
      manifest.problems.push_back("a drop_ stage takes unmodelled_dropped, and only it: " + raw);
    }
    manifest.entries.push_back(std::move(e));
  }
  return manifest;
}

// ---- model snapshot ----

using Model = std::vector<std::pair<std::string, std::string>>;  // field -> rendered value

std::size_t CountElements(std::string_view xml, std::string_view tag) {
  const std::string open = "<" + std::string(tag);
  std::size_t count = 0;
  for (std::size_t at = xml.find(open); at != std::string_view::npos; at = xml.find(open, at + open.size())) {
    const std::size_t next = at + open.size();
    if (next < xml.size() && (xml[next] == ' ' || xml[next] == '/' || xml[next] == '>' || xml[next] == '\t' ||
                              xml[next] == '\n' || xml[next] == '\r')) {
      ++count;
    }
  }
  return count;
}

// Number of records of `wanted` in a framed record block.
std::size_t CountRecords(const std::vector<std::uint8_t>& bytes, std::uint32_t wanted) {
  std::size_t count = 0;
  std::size_t at = 0;
  while (at < bytes.size()) {
    std::uint32_t type = bytes[at] & 0x7FU;
    if ((bytes[at++] & 0x80U) != 0U && at < bytes.size()) {
      type |= static_cast<std::uint32_t>(bytes[at++] & 0x7FU) << 7;
    }
    std::uint32_t size = 0;
    for (unsigned shift = 0; shift < 28U && at < bytes.size(); shift += 7U) {
      const std::uint8_t b = bytes[at++];
      size |= static_cast<std::uint32_t>(b & 0x7FU) << shift;
      if ((b & 0x80U) == 0U) {
        break;
      }
    }
    if (type == wanted) {
      ++count;
    }
    at += size;
  }
  return count;
}

template <typename Fn>
std::string PerSheet(const Workbook& wb, Fn&& fn) {
  std::string out;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    out += (i == 0 ? "" : ";") + fn(wb.sheet(i), i);
  }
  return out;
}

std::string Join(std::vector<std::string> items) {
  std::sort(items.begin(), items.end());
  std::string out;
  for (std::size_t i = 0; i < items.size(); ++i) {
    out += (i == 0 ? "" : ",") + items[i];
  }
  return out;
}

/// "on" when `<workbookProtection>` locks something or carries a password; a bare element locks nothing, and
/// `.xlsb` stores no record for it.
std::string WorkbookProtectionState(const std::string& xml) {
  pugi::xml_document doc;
  if (xml.empty() || !doc.load_buffer(xml.data(), xml.size())) {
    return "off";
  }
  for (const pugi::xml_attribute attr : doc.first_child().attributes()) {
    const std::string_view value = attr.value();
    if (value != "0" && value != "false") {
      return "on";
    }
  }
  return "off";
}

Model Snapshot(const Workbook& wb) {
  Model m;
  std::string names;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    names += (i == 0 ? "" : "|") + wb.sheet(i).name();
  }
  m.emplace_back("sheet_names", names);

  // A macro name is .xlsb-only, as in Excel (an .xlsx save leaves it out); xlsb_roundtrip_defined_names_test pins it.
  std::vector<std::string> defined;
  for (const DefinedName& d : wb.defined_names()) {
    if (d.macro_kind == NameMacroKind::kNone) {
      defined.push_back(d.name + "@" + std::to_string(d.local_sheet_id));
    }
  }
  m.emplace_back("defined_names", Join(defined));

  const std::size_t custom_views =
      io::xlsb::xml_book_view_count(wb) + CountRecords(wb.xlsb_user_book_views(), io::xlsb::kBrtUserBookView);
  m.emplace_back("custom_view_count", std::to_string(custom_views));
  m.emplace_back("custom_sheet_view_count", PerSheet(wb, [](const Sheet& s, std::size_t) {
                   std::size_t views = io::xlsb::xml_sheet_view_count(s);
                   const XlsbSheetTail& tail = s.xlsb_tail();
                   for (const std::vector<std::uint8_t>* slot :
                        {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
                     views += CountRecords(*slot, io::xlsb::kBrtBeginUserShView);
                   }
                   return std::to_string(views);
                 }));
  m.emplace_back("workbook_protection", WorkbookProtectionState(wb.workbook_protection_xml()));

  m.emplace_back("sheet_view_count", PerSheet(wb, [](const Sheet& s, std::size_t) {
                   return std::to_string(1U + CountElements(s.view().extra_views_xml, "sheetView"));
                 }));
  m.emplace_back("pivot_count",
                 PerSheet(wb, [](const Sheet& s, std::size_t) { return std::to_string(s.pivot_tables().size()); }));
  m.emplace_back("sheet_protection", PerSheet(wb, [](const Sheet& s, std::size_t) {
                   return std::string(s.protection().enabled ? "on" : "off");
                 }));
  m.emplace_back("sparkline_group_count", PerSheet(wb, [](const Sheet& s, std::size_t) {
                   return std::to_string(io::read_sparkline_groups(s.ext_lst_xml()).size());
                 }));
  m.emplace_back("table_names", PerSheet(wb, [&wb](const Sheet&, std::size_t i) {
                   std::vector<std::string> tables;
                   for (const TableMetadata& t : wb.tables()) {
                     if (t.sheet_index == i) {
                       tables.push_back(t.display_name);
                     }
                   }
                   return "[" + Join(tables) + "]";
                 }));
  return m;
}

// ---- pipeline ----

struct Failure {
  std::string path;
  std::string stage;
  std::string detail;
  std::string message;
};

bool ReadFile(const fs::path& path, std::vector<std::uint8_t>& out) {
  std::ifstream in(path, std::ios::binary);
  if (!in) {
    return false;
  }
  out.assign(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
  return true;
}

std::string Code(const Error& e) {
  return std::to_string(static_cast<int>(e.code));
}

io::ByteSpan Span(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

/// The value of string field `key` in one structured-log record, or "" when absent.
std::string RecordField(std::string_view record, std::string_view key) {
  const std::string open = "\"" + std::string(key) + "\":\"";
  const std::size_t at = record.find(open);
  if (at == std::string_view::npos) {
    return std::string();
  }
  std::string out;
  for (std::size_t i = at + open.size(); i < record.size() && record[i] != '"'; ++i) {
    if (record[i] == '\\' && i + 1U < record.size()) {
      ++i;
    }
    out.push_back(record[i]);
  }
  return out;
}

/// The drops one save logs: a part path, or the target of a relationship left out.
struct DropLog {
  std::vector<std::string> parts;
  std::vector<std::string> relationship_targets;
};

void CollectDrop(std::string_view record, void* user_data) {
  auto& log = *static_cast<DropLog*>(user_data);
  const std::string event = RecordField(record, "event");
  if (event == "ooxml_writer.passthrough_collision" || event == "xlsb.writer.passthrough_collision" ||
      event == "ooxml_writer.passthrough_dropped") {
    log.parts.push_back(RecordField(record, "path"));
  } else if (event.size() > 12U && event.compare(event.size() - 12U, 12U, "_rel_skipped") == 0) {
    log.relationship_targets.push_back(RecordField(record, "target"));
  }
}

/// Routes warnings to a `DropLog` for one save, restoring the default (off) afterwards.
class DropCapture {
 public:
  DropCapture() {
    set_structured_log_sink(CollectDrop, &log_);
    set_structured_log_min_level(StructuredLogLevel::kWarn);
  }
  ~DropCapture() {
    set_structured_log_min_level(StructuredLogLevel::kOff);
    set_structured_log_sink(nullptr);
  }
  DropCapture(const DropCapture&) = delete;
  DropCapture& operator=(const DropCapture&) = delete;
  const DropLog& log() const { return log_; }

 private:
  DropLog log_;
};

void AddDrops(const std::string& path, const std::string& tag, const io::WriteDiagnostics& diagnostics,
              const DropLog& log, std::vector<Failure>& out) {
  const std::string stage = "drop_" + tag;
  // A relationship is lost only with its target part.
  std::vector<std::string> relationships;
  for (const std::string& target : log.relationship_targets) {
    if (std::find(log.parts.begin(), log.parts.end(), target) != log.parts.end()) {
      relationships.push_back("rel:" + target);
    }
  }
  const std::vector<std::string>& counted_relationships = relationships;
  for (const std::vector<std::string>* items : {&log.parts, &counted_relationships}) {
    for (const std::string& item : *items) {
      out.push_back({path, stage, item, "counted as dropped by an unedited save"});
    }
  }
  if (diagnostics.dropped_part_count != log.parts.size() ||
      diagnostics.dropped_relationship_count != relationships.size()) {
    out.push_back({path, stage, "count_mismatch",
                   "counts " + std::to_string(diagnostics.dropped_part_count) + "/" +
                       std::to_string(diagnostics.dropped_relationship_count) + " but logged " +
                       std::to_string(log.parts.size()) + "/" + std::to_string(relationships.size())});
  }
}

// Loads `bytes` as xlsx or xlsb; on failure fills `fail` with the code and message.
bool LoadAny(const std::vector<std::uint8_t>& bytes, bool xlsb, Workbook& out, Error& error) {
  if (xlsb) {
    auto r = io::xlsb::read_xlsb(Span(bytes));
    if (!r) {
      error = r.error();
      return false;
    }
    out = std::move(r.value().workbook);
  } else {
    auto r = io::read_ooxml(Span(bytes));
    if (!r) {
      error = r.error();
      return false;
    }
    out = std::move(r.value().workbook);
  }
  return true;
}

void CompareModels(const std::string& path, const std::string& stage, const Model& want, const Model& got,
                   std::vector<Failure>& out) {
  for (std::size_t i = 0; i < want.size() && i < got.size(); ++i) {
    if (want[i].second != got[i].second) {
      out.push_back({path, stage, want[i].first,
                     want[i].first + " differs: '" + want[i].second + "' became '" + got[i].second + "'"});
    }
  }
}

void AddValidation(const std::string& path, const std::string& stage, const std::vector<std::string>& findings,
                   std::vector<Failure>& out) {
  std::map<std::string, std::string> by_rule;
  for (const std::string& f : findings) {
    const std::string rule = f.substr(0, f.find(": "));
    std::string& joined = by_rule[rule];
    joined += (joined.empty() ? "" : "; ") + f;
  }
  for (const auto& [rule, message] : by_rule) {
    out.push_back({path, stage, rule, message});
  }
}

std::vector<Failure> RunFile(const std::string& rel) {
  std::vector<Failure> out;
  std::vector<std::uint8_t> source;
  if (!ReadFile(CorpusRoot() / rel, source)) {
    out.push_back({rel, "load", "0", "cannot read the fixture"});
    return out;
  }
  Workbook wb = Workbook::create_empty();
  Error error;
  if (!LoadAny(source, fs::path(rel).extension() == ".xlsb", wb, error)) {
    out.push_back({rel, "load", Code(error), error.message});
    return out;
  }
  const Model want = Snapshot(wb);

  // One branch per output format; a failed save stops only its branch.
  for (const bool xlsb : {false, true}) {
    const std::string tag = xlsb ? "xlsb" : "xlsx";
    std::vector<std::uint8_t> saved;
    if (xlsb) {
      DropCapture capture;
      auto r = io::xlsb::write_xlsb_with_result(wb);
      if (!r) {
        out.push_back({rel, "save_" + tag, Code(r.error()), r.error().message});
        continue;
      }
      AddDrops(rel, tag, r.value().diagnostics, capture.log(), out);
      saved = std::move(r.value().bytes);
    } else {
      DropCapture capture;
      auto r = io::write_ooxml_with_result(wb);
      if (!r) {
        out.push_back({rel, "save_" + tag, Code(r.error()), r.error().message});
        continue;
      }
      AddDrops(rel, tag, r.value().diagnostics, capture.log(), out);
      saved = std::move(r.value().bytes);
    }

    Workbook reloaded = Workbook::create_empty();
    if (LoadAny(saved, xlsb, reloaded, error)) {
      CompareModels(rel, "model_" + tag, want, Snapshot(reloaded), out);
    } else {
      out.push_back({rel, "reload_" + tag, Code(error), error.message});
    }

    if (!xlsb) {
      AddValidation(rel, "validate_xlsx", test::ValidateXlsxPackage(Span(saved)), out);
    } else {
      AddValidation(rel, "validate_xlsb", test::ValidateXlsbPackage(Span(saved)), out);
    }
  }
  return out;
}

/// The manifest line that would accept `f`, with the kind its stage implies; the reason is the author's to write.
std::string SuggestedEntry(const Failure& f, bool source_refused) {
  const char* kind = "known_gap";
  if (f.stage.rfind("save_", 0) == 0) {
    kind = "save_refused";
  } else if (f.stage.rfind("drop_", 0) == 0) {
    kind = "unmodelled_dropped";
  } else if (source_refused && f.stage.rfind("validate_", 0) == 0) {
    kind = "mirrors_excel_refused_source";
  }
  return f.path + " " + f.stage + " " + f.detail + " " + kind + "  # <reason>";
}

class CorpusRoundtripShards : public ::testing::TestWithParam<Shard> {};

TEST_P(CorpusRoundtripShards, KeepsModelAndPassesValidators) {
  const Shard& shard = GetParam();
  Manifest manifest = LoadManifest();
  for (const std::string& problem : manifest.problems) {
    ADD_FAILURE() << problem;
  }
  const std::set<std::string> files(shard.files.begin(), shard.files.end());
  std::set<std::string> refused;
  for (const Entry& e : manifest.entries) {
    if (e.kind == Kind::kExcelRefuses) {
      refused.insert(e.path);
    }
  }
  for (const Entry& e : manifest.entries) {
    if (files.count(e.path) != 0 && e.kind == Kind::kMirrors && refused.count(e.path) == 0) {
      ADD_FAILURE() << "mirrors_excel_refused_source needs an excel_refuses line for the same path: " << e.line;
    }
  }

  for (const std::string& rel : shard.files) {
    for (const Failure& f : RunFile(rel)) {
      bool expected = false;
      for (Entry& e : manifest.entries) {
        if (e.path != rel) {
          continue;
        }
        if (e.kind == Kind::kDamaged) {
          expected = true;
        } else if (e.kind != Kind::kExcelRefuses && e.stage == f.stage && e.detail == f.detail) {
          e.matched = true;
          expected = true;
        }
      }
      if (!expected) {
        ADD_FAILURE() << f.path << " " << f.stage << " " << f.detail << ": " << f.message << "\n"
                      << "manifest-line: " << SuggestedEntry(f, refused.count(rel) != 0);
      }
    }
  }

  for (const Entry& e : manifest.entries) {
    const bool exact = e.kind != Kind::kExcelRefuses && e.kind != Kind::kDamaged;
    if (exact && files.count(e.path) != 0 && !e.matched) {
      ADD_FAILURE() << "no longer fails, remove the entry: " << e.line;
    }
  }
}

INSTANTIATE_TEST_SUITE_P(CorpusRoundtrip, CorpusRoundtripShards, ::testing::ValuesIn(ComputeShards()),
                         [](const ::testing::TestParamInfo<Shard>& param_info) { return param_info.param.name; });

TEST(CorpusRoundtripWalk, FindsTheWholeCorpus) {
  const std::vector<Shard> shards = ComputeShards();
  std::size_t total = 0;
  for (const Shard& s : shards) {
    total += s.files.size();
  }
  EXPECT_FALSE(shards.empty());
  EXPECT_GT(total, kMinCorpusFiles);
}

TEST(CorpusRoundtripManifest, EveryPathExistsAndEveryLineIsWellFormed) {
  const Manifest manifest = LoadManifest();
  for (const std::string& problem : manifest.problems) {
    ADD_FAILURE() << problem;
  }
  for (const Entry& e : manifest.entries) {
    EXPECT_TRUE(fs::is_regular_file(CorpusRoot() / e.path)) << "manifest path does not exist: " << e.line;
  }
}

}  // namespace
}  // namespace formulon
