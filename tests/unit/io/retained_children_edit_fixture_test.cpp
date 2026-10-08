//
// Row / column edits against Mac Excel 365 (16.113, ja-JP) saves under
// `tests/fixtures/excel/retained_children_edit/`. The `children` book carries
// every retained `<worksheet>` child that names cells -- protected ranges,
// three scenarios, a sheet-level sort state, a consolidation source, a custom
// view, cell watches, ignored errors and a range web-publish item; the
// `autofilter` book an AutoFilter with two criteria and a nested sort state,
// plus row and column breaks, which `.xlsb` retains as records. Excel saved
// each as-is and after each edit, in both formats. Applying the same edit to
// the base file and saving must leave those coordinates where Excel's own
// save does. The ranges inside one sqref are compared as a set: Excel
// reorders them on save, and their order carries nothing.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <functional>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& file) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/retained_children_edit/" + file;
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
    out.clear();
  }
  std::fclose(f);
  return out;
}

Workbook Load(const std::vector<std::uint8_t>& bytes, bool xlsb) {
  const ByteSpan span{bytes.data(), bytes.size()};
  if (xlsb) {
    auto r = xlsb::read_xlsb(span);
    EXPECT_TRUE(static_cast<bool>(r));
    return r ? std::move(r.value().workbook) : Workbook::create_empty();
  }
  auto r = read_ooxml(span);
  EXPECT_TRUE(static_cast<bool>(r));
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb, bool xlsb) {
  if (xlsb) {
    auto r = xlsb::write_xlsb(wb);
    EXPECT_TRUE(static_cast<bool>(r));
    return r ? std::move(r.value()) : std::vector<std::uint8_t>{};
  }
  auto r = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(r));
  return r ? std::move(r.value()) : std::vector<std::uint8_t>{};
}

std::string ReadEntry(const std::vector<std::uint8_t>& package, const std::string& name) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
  auto entry = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(entry)) << name;
  return entry ? std::string(entry.value().begin(), entry.value().end()) : std::string();
}

/// `text`'s space-separated tokens, sorted.
std::string SortedTokens(const std::string& text) {
  std::vector<std::string> tokens;
  std::size_t at = 0;
  while (at < text.size()) {
    const std::size_t end = std::min(text.find(' ', at), text.size());
    if (end > at) {
      tokens.push_back(text.substr(at, end - at));
    }
    at = end + 1U;
  }
  std::sort(tokens.begin(), tokens.end());
  std::string out;
  for (const std::string& t : tokens) {
    out += (out.empty() ? "" : " ") + t;
  }
  return out;
}

/// The coordinate attributes of the fixture's retained children, one string
/// per element in document order: `tag attr=value ...`.
std::vector<std::string> XlsxCoordinates(const std::vector<std::uint8_t>& package) {
  struct Probe {
    const char* tag;
    std::vector<const char*> attrs;
  };
  const std::vector<Probe> probes = {
      {"protectedRange", {"sqref"}},
      {"scenarios", {"current", "show", "sqref"}},
      {"scenario", {"name", "count"}},
      {"inputCells", {"r"}},
      {"sortState", {"ref"}},
      {"sortCondition", {"ref"}},
      {"dataRef", {"ref"}},
      {"customSheetView", {"topLeftCell"}},
      {"cellWatch", {"r"}},
      {"ignoredError", {"sqref"}},
      {"webPublishItems", {"count"}},
      {"webPublishItem", {"sourceRef"}},
      {"autoFilter", {"ref"}},
      {"filterColumn", {"colId"}},
      {"rowBreaks", {"count", "manualBreakCount"}},
      {"colBreaks", {"count", "manualBreakCount"}},
      {"brk", {"id"}},
  };
  const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.xml");
  std::vector<std::pair<std::size_t, std::string>> found;
  for (const Probe& probe : probes) {
    const std::string open = std::string("<") + probe.tag;
    for (std::size_t at = sheet.find(open); at != std::string::npos; at = sheet.find(open, at + 1U)) {
      const char next = sheet[at + open.size()];
      if (next != ' ' && next != '>' && next != '/') {
        continue;
      }
      const std::string tag = sheet.substr(at, sheet.find('>', at) - at);
      std::string line = probe.tag;
      for (const char* attr : probe.attrs) {
        const std::string key = std::string(" ") + attr + "=\"";
        const std::size_t v = tag.find(key);
        if (v != std::string::npos) {
          const std::size_t begin = v + key.size();
          line += std::string(" ") + attr + "=" + SortedTokens(tag.substr(begin, tag.find('"', begin) - begin));
        }
      }
      found.emplace_back(at, line);
    }
  }
  std::sort(found.begin(), found.end());
  std::vector<std::string> out;
  for (auto& [at, line] : found) {
    out.push_back(std::move(line));
  }
  return out;
}

std::uint32_t U32(const std::uint8_t* p) {
  return static_cast<std::uint32_t>(p[0]) | (static_cast<std::uint32_t>(p[1]) << 8) |
         (static_cast<std::uint32_t>(p[2]) << 16) | (static_cast<std::uint32_t>(p[3]) << 24);
}

std::string Hex(const std::uint8_t* p, std::size_t n) {
  std::string out;
  for (std::size_t i = 0; i < n; ++i) {
    char buf[4];
    std::snprintf(buf, sizeof(buf), " %02x", p[i]);
    out += buf;
  }
  return out;
}

/// The Sqrfx at `offset` as sorted `r0:r1/c0:c1` tokens, then the bytes after it.
std::string SqrfxAt(const xlsb::FramedRecord& rec, std::size_t offset) {
  const std::uint8_t* p = rec.payload.data;
  const std::uint32_t count = U32(p + offset);
  std::string ranges;
  for (std::uint32_t i = 0; i < count; ++i) {
    const std::uint8_t* r = p + offset + 4U + 16U * i;
    ranges += (i == 0 ? "" : " ") + std::to_string(U32(r)) + ":" + std::to_string(U32(r + 4)) + "/" +
              std::to_string(U32(r + 8)) + ":" + std::to_string(U32(r + 12));
  }
  const std::size_t rest = offset + 4U + 16U * count;
  return Hex(p, offset) + " [" + SortedTokens(ranges) + "]" + Hex(p + rest, rec.payload.size - rest);
}

/// The payloads of the records behind the same children, in stream order:
/// the AutoFilter block's range, columns and sort state, range protection,
/// the scenario block, the sheet sort state, the consolidation source, the
/// breaks, cell watches, ignored errors and the web-publish items.
std::vector<std::string> XlsbCoordinates(const std::vector<std::uint8_t>& package) {
  const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.bin");
  const std::vector<std::uint8_t> bytes(sheet.begin(), sheet.end());
  std::vector<std::string> out;
  for (const xlsb::FramedRecord& rec : xlsb::split_records(bytes)) {
    const std::uint16_t t = rec.type;
    if (t == 536U) {
      out.push_back("536" + SqrfxAt(rec, 2U));
    } else if (t == 649U) {
      out.push_back("649" + SqrfxAt(rec, 4U));
    } else if ((t >= 500U && t <= 504U) || (t >= 530U && t <= 533U) || t == 499U || t == 607U ||
               (t >= 554U && t <= 557U) || t == 161U || t == 162U || t == 163U || (t >= 392U && t <= 396U)) {
      out.push_back(std::to_string(t) + Hex(rec.payload.data, rec.payload.size));
    }
  }
  return out;
}

std::vector<std::string> Coordinates(const std::vector<std::uint8_t>& package, bool xlsb) {
  return xlsb ? XlsbCoordinates(package) : XlsxCoordinates(package);
}

struct EditCase {
  const char* state;
  std::function<Expected<void, Error>(Workbook&)> apply;
};

const std::vector<EditCase> kChildrenEdits = {
    {"1_insert_row4", [](Workbook& wb) { return wb.insert_rows(0, 3, 1); }},
    {"2_delete_row5", [](Workbook& wb) { return wb.delete_rows(0, 4, 1); }},
    {"3_delete_rows3_6", [](Workbook& wb) { return wb.delete_rows(0, 2, 4); }},
    {"4_delete_row9", [](Workbook& wb) { return wb.delete_rows(0, 8, 1); }},
    {"5_delete_rows2_10", [](Workbook& wb) { return wb.delete_rows(0, 1, 9); }},
    {"6_insert_colC", [](Workbook& wb) { return wb.insert_cols(0, 2, 1); }},
    {"7_delete_colB", [](Workbook& wb) { return wb.delete_cols(0, 1, 1); }},
    {"8_delete_colsB_D", [](Workbook& wb) { return wb.delete_cols(0, 1, 3); }},
};

const std::vector<EditCase> kAutoFilterEdits = {
    {"1_insert_row4", [](Workbook& wb) { return wb.insert_rows(0, 3, 1); }},
    {"2_delete_row6", [](Workbook& wb) { return wb.delete_rows(0, 5, 1); }},
    {"3_delete_colB", [](Workbook& wb) { return wb.delete_cols(0, 1, 1); }},
    {"4_insert_colB", [](Workbook& wb) { return wb.insert_cols(0, 1, 1); }},
    {"5_delete_colC", [](Workbook& wb) { return wb.delete_cols(0, 2, 1); }},
};

void CheckBook(const std::string& book, const std::vector<EditCase>& edits, bool xlsb) {
  const char* ext = xlsb ? ".xlsb" : ".xlsx";
  const std::vector<std::uint8_t> base = ReadFixture(book + "_0_base" + ext);
  const std::vector<std::string> base_coords = Coordinates(base, xlsb);
  ASSERT_GE(base_coords.size(), 8U) << book << ext;
  EXPECT_EQ(Coordinates(Save(Load(base, xlsb), xlsb), xlsb), base_coords) << book << ext;
  for (const EditCase& edit : edits) {
    Workbook wb = Load(base, xlsb);
    ASSERT_TRUE(static_cast<bool>(edit.apply(wb))) << edit.state;
    const std::vector<std::string> excel = Coordinates(ReadFixture(book + "_" + edit.state + ext), xlsb);
    EXPECT_NE(excel, base_coords) << book << "_" << edit.state << ext;
    EXPECT_EQ(Coordinates(Save(wb, xlsb), xlsb), excel) << book << "_" << edit.state << ext;
  }
}

TEST(RetainedChildrenEditFixture, ChildrenXlsx) {
  CheckBook("children", kChildrenEdits, false);
}

TEST(RetainedChildrenEditFixture, ChildrenXlsb) {
  CheckBook("children", kChildrenEdits, true);
}

TEST(RetainedChildrenEditFixture, AutoFilterXlsx) {
  CheckBook("autofilter", kAutoFilterEdits, false);
}

TEST(RetainedChildrenEditFixture, AutoFilterXlsb) {
  CheckBook("autofilter", kAutoFilterEdits, true);
}

// What the comparisons hold for the two rules no other structure follows,
// spelt out: an insert cuts an ignored-error range instead of stretching it,
// and a watch on a deleted cell keeps its address.
TEST(RetainedChildrenEditFixture, IgnoredErrorsCutAndWatchesStay) {
  Workbook inserted = Load(ReadFixture("children_0_base.xlsx"), false);
  ASSERT_TRUE(static_cast<bool>(inserted.insert_rows(0, 3, 1)));
  const std::vector<std::string> after_insert = XlsxCoordinates(Save(inserted, false));
  EXPECT_NE(std::find(after_insert.begin(), after_insert.end(), "ignoredError sqref=A2:A3 A5:A11 C5"),
            after_insert.end());

  Workbook deleted = Load(ReadFixture("children_0_base.xlsx"), false);
  ASSERT_TRUE(static_cast<bool>(deleted.delete_rows(0, 4, 1)));
  const std::vector<std::string> after_delete = XlsxCoordinates(Save(deleted, false));
  EXPECT_NE(std::find(after_delete.begin(), after_delete.end(), "cellWatch r=C5"), after_delete.end());
  EXPECT_NE(std::find(after_delete.begin(), after_delete.end(), "cellWatch r=E7"), after_delete.end());
}

}  // namespace
}  // namespace io
}  // namespace formulon
