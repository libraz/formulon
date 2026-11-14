// The notes and thread stubs the xlsb writer writes from the model (`xl/commentsN.bin`, its legacy drawing and the
// run fonts it names), held against Excel 16.115's own xlsb saves of the same sources
// (tests/fixtures/excel/xlsb_parity/notes and threads).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <iterator>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "utils/structured_log.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;

Bytes ReadFile(const std::string& path) {
  Bytes out;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(f);
  return out;
}

std::string Fixture(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/" + name;
}

/// The part a worksheet's relationship of `type_suffix` names, as package bytes; empty when there is none.
Bytes RelatedPart(const Bytes& package, std::size_t sheet_position, const std::string& type_suffix) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto rels = zip.read_entry("xl/worksheets/_rels/sheet" + std::to_string(sheet_position) + ".bin.rels");
  if (!rels) {
    return {};
  }
  const std::string xml(rels.value().begin(), rels.value().end());
  const std::size_t type = xml.find("relationships/" + type_suffix + "\"");
  if (type == std::string::npos) {
    return {};
  }
  const std::size_t open = xml.rfind('<', type);
  const std::size_t target = xml.find("Target=\"", open) + 8U;
  std::string path = xml.substr(target, xml.find('"', target) - target);
  path = path.rfind("../", 0) == 0 ? "xl/" + path.substr(3) : "xl/worksheets/" + path;
  auto part = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(part)) << path;
  return part ? std::move(part.value()) : Bytes{};
}

/// A part's records as `type:hex` lines, which keep a failure readable.
std::vector<std::string> RecordLines(const Bytes& part) {
  static const char kHex[] = "0123456789abcdef";
  std::vector<std::string> out;
  for (const FramedRecord& rec : split_records(part)) {
    std::string line = std::to_string(rec.type) + ":";
    for (std::size_t i = 0; i < rec.payload.size; ++i) {
      line.push_back(kHex[rec.payload.data[i] >> 4U]);
      line.push_back(kHex[rec.payload.data[i] & 0x0FU]);
    }
    out.push_back(std::move(line));
  }
  return out;
}

/// The BrtFont records of a package's styles part.
std::vector<std::string> FontLines(const Bytes& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto styles = zip.read_entry("xl/styles.bin");
  std::vector<std::string> out;
  if (styles) {
    for (const std::string& line : RecordLines(styles.value())) {
      if (line.rfind("43:", 0) == 0) {
        out.push_back(line);
      }
    }
  }
  return out;
}

/// The number of BrtLegacyDrawing records in a worksheet part.
std::size_t LegacyDrawings(const Bytes& package, std::size_t sheet_position) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet" + std::to_string(sheet_position) + ".bin");
  std::size_t count = 0;
  if (part) {
    for (const FramedRecord& rec : split_records(part.value())) {
      count += rec.type == kBrtLegacyDrawing ? 1U : 0U;
    }
  }
  return count;
}

/// The deferred-feature log records of the saves made while it lives, which name what a save left out.
class DeferredRecords {
 public:
  DeferredRecords() {
    set_structured_log_sink(&Collect, &records_);
    set_structured_log_min_level(StructuredLogLevel::kWarn);
  }
  ~DeferredRecords() {
    set_structured_log_min_level(StructuredLogLevel::kOff);
    set_structured_log_sink(nullptr);
  }
  DeferredRecords(const DeferredRecords&) = delete;
  DeferredRecords& operator=(const DeferredRecords&) = delete;
  const std::string& text() const noexcept { return records_; }

 private:
  static void Collect(std::string_view record, void* user) {
    if (record.find("xlsb.writer.deferred") != std::string_view::npos) {
      static_cast<std::string*>(user)->append(record);
    }
  }
  std::string records_;
};

struct Case {
  const char* source;
  const char* excel;
};

// A source without uids is the one Excel saw; `runs_fonts` starts from Excel's own .xlsx, where the runs were added.
const Case kNoteCases[] = {{"notes/base_src.xlsx", "notes/base.xlsb"},
                           {"notes/runs_fonts.xlsx", "notes/runs_fonts.xlsb"},
                           {"notes/shapes_src.xlsx", "notes/shapes.xlsb"},
                           {"notes/uid_two_sheets_src.xlsx", "notes/uid_two_sheets.xlsb"},
                           {"threads/base_src.xlsx", "threads/base.xlsb"}};

Workbook LoadSource(const char* source) {
  auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(source))));
  EXPECT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook wb = loaded ? std::move(loaded.value().workbook) : Workbook::create();
  // The sources were saved by Mac Excel in Japanese, which words the thread stubs.
  wb.set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

void ExpectCommentsPartMatchesExcel(const Case& c) {
  {
    SCOPED_TRACE(c.source);
    const Workbook wb = LoadSource(c.source);
    const DeferredRecords deferred;
    auto saved = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    const Bytes excel = ReadFile(Fixture(c.excel));
    for (std::size_t s = 1; s <= wb.sheet_count(); ++s) {
      SCOPED_TRACE(s);
      const Bytes theirs = RelatedPart(excel, s, "comments");
      ASSERT_FALSE(theirs.empty());
      EXPECT_EQ(RecordLines(RelatedPart(saved.value().bytes, s, "comments")), RecordLines(theirs));
      EXPECT_FALSE(RelatedPart(saved.value().bytes, s, "vmlDrawing").empty());
      EXPECT_EQ(LegacyDrawings(saved.value().bytes, s), 1U);
    }
    EXPECT_EQ(FontLines(saved.value().bytes), FontLines(excel));
    EXPECT_EQ(deferred.text().find("comments"), std::string::npos) << deferred.text();
  }
}

TEST(XlsbParityNotesTwins, CommentsPartMatchesExcel) {
  for (std::size_t i = 0; i + 1U < std::size(kNoteCases); ++i) {
    ExpectCommentsPartMatchesExcel(kNoteCases[i]);
  }
}

TEST(XlsbParityThreadsTwins, StubsMatchExcel) {
  // A thread's stub note sits among the notes in cell order, plain text under its `tc=` author.
  ExpectCommentsPartMatchesExcel(kNoteCases[std::size(kNoteCases) - 1U]);
}

TEST(XlsbParityNotesEdit, ExcelSaveWritesBackUnchanged) {
  for (const Case& c : kNoteCases) {
    SCOPED_TRACE(c.excel);
    const Bytes excel = ReadFile(Fixture(c.excel));
    auto loaded = read_xlsb(SpanOf(excel));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    Workbook& wb = loaded.value().workbook;
    wb.set_excel_profile(mac_365_ja_jp_profile());
    EXPECT_FALSE(wb.sheet(0).comments().empty());
    const DeferredRecords deferred;
    auto saved = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    for (std::size_t s = 1; s <= wb.sheet_count(); ++s) {
      SCOPED_TRACE(s);
      EXPECT_EQ(RecordLines(RelatedPart(saved.value().bytes, s, "comments")),
                RecordLines(RelatedPart(excel, s, "comments")));
      EXPECT_EQ(LegacyDrawings(saved.value().bytes, s), 1U);
    }
    EXPECT_EQ(FontLines(saved.value().bytes), FontLines(excel));
    EXPECT_EQ(deferred.text().find("comments"), std::string::npos) << deferred.text();
  }
}

TEST(XlsbParityNotesEdit, RunsAndFontsReadFromExcelsSave) {
  auto loaded = read_xlsb(SpanOf(ReadFile(Fixture("notes/runs_fonts.xlsb"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const CellComment& note = loaded.value().workbook.sheet(0).comments().at(0);
  EXPECT_EQ(note.text, "Alice: red italic note");
  ASSERT_EQ(note.runs.size(), 2U);
  EXPECT_EQ(note.runs[0].text, "Alice:");
  ASSERT_TRUE(note.runs[0].font.has_value());
  EXPECT_EQ(note.runs[0].font->name, "Tahoma");
  EXPECT_EQ(note.runs[1].text, " red italic note");
  ASSERT_TRUE(note.runs[1].font.has_value());
  EXPECT_EQ(note.runs[1].font->name, "Arial");
  EXPECT_EQ(note.uid, "{00000000-0006-0000-0000-000001000000}");
}

/// Each sheet's note uids in comment-list order, with the note's cell.
std::vector<std::vector<std::string>> NoteUids(const Workbook& wb) {
  std::vector<std::vector<std::string>> out(wb.sheet_count());
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const CellComment& note : wb.sheet(s).comments()) {
      out[s].push_back(std::to_string(note.row) + "," + std::to_string(note.col) + " " + note.uid);
    }
  }
  return out;
}

TEST(XlsbParityNotesEdit, UidlessNotesTakeExcelsUidOnEverySheet) {
  // Excel's uid for a note without one: the sheet's tab position and the note's 1-based ordinal, on any sheet,
  // whether or not earlier sheets have notes; a note's own uid is kept and still takes its ordinal (u3).
  for (const char* name : {"uid_u1", "uid_u2", "uid_u3", "uid_u4", "uid_u5"}) {
    SCOPED_TRACE(name);
    const Workbook wb = LoadSource((std::string("notes/") + name + "_src.xlsx").c_str());
    const DeferredRecords deferred;
    auto saved = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(deferred.text().find("comments"), std::string::npos) << deferred.text();
    auto ours = read_xlsb(SpanOf(saved.value().bytes));
    ASSERT_TRUE(static_cast<bool>(ours)) << ours.error().message;
    auto excel = read_xlsb(SpanOf(ReadFile(Fixture(std::string("notes/") + name + ".xlsb"))));
    ASSERT_TRUE(static_cast<bool>(excel)) << excel.error().message;
    EXPECT_EQ(NoteUids(ours.value().workbook), NoteUids(excel.value().workbook));
  }
}

TEST(XlsbParityThreadsEdit, ThreadsAndPersonsReadFromExcelsSave) {
  auto loaded = read_xlsb(SpanOf(ReadFile(Fixture("threads/base.xlsb"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;
  EXPECT_EQ(wb.sheet(0).comments().size(), 3U);
  EXPECT_FALSE(wb.sheet(0).threaded_comments().empty());
  EXPECT_FALSE(wb.persons().empty());
  const DeferredRecords deferred;
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_FALSE(RelatedPart(saved.value().bytes, 1, "threadedComment").empty());
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(saved.value().bytes))));
  EXPECT_TRUE(zip.has_entry("xl/persons/person.xml"));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
