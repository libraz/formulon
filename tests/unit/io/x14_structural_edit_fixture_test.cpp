//
// Row / column edits against Windows Excel 365 (16.0.20228, ja-JP) saves
// under `tests/fixtures/excel/win/x14_structural/`. Each book holds a
// DataBar with x14 settings (over the union `A1:A6 C1:C6`, or `A1:A6` as
// the control) and a sparkline at G1 over `E1:E6`; Excel saved it as-is
// and after each edit, in both formats. Applying the same edit to the base
// file and saving must move the legacy CF range, the x14 CF range and the
// sparkline's location and source exactly as Excel's own save does.
//
// One Excel behaviour is not reproduced: inserting row 2 also gives the new
// row a copy of the sparkline above it (G2 over `E2:E8`), as an insert
// inherits the formatting of the row above. That copy is spelt out below.

#include <cstdint>
#include <cstdio>
#include <functional>
#include <map>
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

/// `file` under tests/fixtures/excel/win/x14_structural/, or under
/// tests/fixtures/excel/ when `dir` is empty.
std::vector<std::uint8_t> ReadFixture(const std::string& file, const std::string& dir = "win/x14_structural/") {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + dir + file;
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

/// Text of every `<tag>` element in `xml`, in document order.
std::vector<std::string> Elements(const std::string& xml, const std::string& tag) {
  std::vector<std::string> out;
  const std::string open = "<" + tag + ">";
  const std::string close = "</" + tag + ">";
  for (std::size_t at = xml.find(open); at != std::string::npos; at = xml.find(open, at)) {
    at += open.size();
    const std::size_t end = xml.find(close, at);
    out.push_back(xml.substr(at, end - at));
  }
  return out;
}

/// The coordinates an .xlsx sheet part carries for the fixture's features:
/// the legacy CF sqref, then every `xm:sqref` / `xm:f` in the `<extLst>`.
std::vector<std::string> XlsxCoordinates(const std::vector<std::uint8_t>& package) {
  const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.xml");
  std::vector<std::string> out;
  const std::size_t cf = sheet.find("<conditionalFormatting sqref=\"");
  if (cf != std::string::npos) {
    const std::size_t at = cf + 30U;
    out.push_back("cf " + sheet.substr(at, sheet.find('"', at) - at));
  }
  const std::string ext = sheet.substr(sheet.rfind("<extLst>"));
  for (const std::string& s : Elements(ext, "xm:sqref")) {
    out.push_back("sqref " + s);
  }
  for (const std::string& f : Elements(ext, "xm:f")) {
    out.push_back("f " + f);
  }
  return out;
}

/// Payloads of the sheet records that carry the fixture's coordinates:
/// BrtBeginConditionalFormatting, the x14 CF block, and each sparkline.
std::vector<std::string> XlsbCoordinates(const std::vector<std::uint8_t>& package) {
  const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.bin");
  const std::vector<std::uint8_t> bytes(sheet.begin(), sheet.end());
  std::vector<std::string> out;
  for (const xlsb::FramedRecord& rec : xlsb::split_records(bytes)) {
    if (rec.type != 461U && rec.type != 1046U && rec.type != 1043U) {
      continue;
    }
    std::string hex = std::to_string(rec.type);
    for (std::size_t i = 0; i < rec.payload.size; ++i) {
      char buf[4];
      std::snprintf(buf, sizeof(buf), " %02x", rec.payload.data[i]);
      hex += buf;
    }
    out.push_back(hex);
  }
  return out;
}

struct EditCase {
  const char* state;
  std::function<Expected<void, Error>(Workbook&)> apply;
};

const EditCase kEdits[] = {
    {"1_insert_row2", [](Workbook& wb) { return wb.insert_rows(0, 1, 1); }},
    {"2_delete_colB", [](Workbook& wb) { return wb.delete_cols(0, 1, 1); }},
    {"3_insert_col1", [](Workbook& wb) { return wb.insert_cols(0, 0, 1); }},
};

std::vector<std::string> Coordinates(const std::vector<std::uint8_t>& package, bool xlsb) {
  return xlsb ? XlsbCoordinates(package) : XlsxCoordinates(package);
}

/// Excel's coordinates without the sparkline an inserted row inherits.
std::vector<std::string> WithoutInheritedSparkline(std::vector<std::string> coords, bool xlsb) {
  std::vector<std::string> out;
  for (const std::string& c : coords) {
    const bool inherited =
        xlsb ? c.rfind("1043 ", 0) == 0 && c.find(" 01 00 00 00 01 00 00 00 06 00 00 00") != std::string::npos
             : c == "sqref G2" || c == "f Sheet1!E2:E8";
    if (!inherited) {
      out.push_back(c);
    }
  }
  return out;
}

void CheckBook(const std::string& book, bool xlsb) {
  const char* ext = xlsb ? ".xlsb" : ".xlsx";
  const std::vector<std::uint8_t> base = ReadFixture(book + "_0_base" + ext);
  // The base file itself survives a load / save with the same coordinates.
  EXPECT_EQ(Coordinates(Save(Load(base, xlsb), xlsb), xlsb), Coordinates(base, xlsb)) << book << ext;
  for (const EditCase& edit : kEdits) {
    Workbook wb = Load(base, xlsb);
    ASSERT_TRUE(static_cast<bool>(edit.apply(wb))) << edit.state;
    std::vector<std::string> excel = Coordinates(ReadFixture(book + "_" + edit.state + ext), xlsb);
    if (std::string(edit.state) == "1_insert_row2") {
      const std::size_t before = excel.size();
      excel = WithoutInheritedSparkline(std::move(excel), xlsb);
      EXPECT_EQ(excel.size() + (xlsb ? 1U : 2U), before) << book << ext << " inherited sparkline not found";
    }
    // Legacy CF, x14 CF, sparkline: one record each, or four xlsx strings.
    EXPECT_EQ(excel.size(), xlsb ? 3U : 4U) << book << "_" << edit.state << ext;
    EXPECT_EQ(Coordinates(Save(wb, xlsb), xlsb), excel) << book << "_" << edit.state << ext;
  }
}

TEST(X14StructuralEditFixture, UnionDataBarXlsx) {
  CheckBook("union_databar", false);
}

// What the comparisons above hold, spelt out for the union book: the x14
// range mirrors the legacy union, and a delete that closes the gap between
// its two ranges joins them.
TEST(X14StructuralEditFixture, UnionJoinsAcrossDeletedColumn) {
  Workbook wb = Load(ReadFixture("union_databar_0_base.xlsx"), false);
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, 1, 1)));
  EXPECT_EQ(XlsxCoordinates(Save(wb, false)),
            (std::vector<std::string>{"cf A1:B6", "sqref A1:B6", "sqref F1", "f Sheet1!D1:D6"}));
}
TEST(X14StructuralEditFixture, UnionDataBarXlsb) {
  CheckBook("union_databar", true);
}
TEST(X14StructuralEditFixture, SingleDataBarXlsx) {
  CheckBook("single_databar", false);
}
TEST(X14StructuralEditFixture, SingleDataBarXlsb) {
  CheckBook("single_databar", true);
}

/// Record types of the saved sheet part, in order.
std::vector<std::uint16_t> XlsbSheetRecordTypes(const std::vector<std::uint8_t>& package) {
  const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.bin");
  const std::vector<std::uint8_t> bytes(sheet.begin(), sheet.end());
  std::vector<std::uint16_t> out;
  for (const xlsb::FramedRecord& rec : xlsb::split_records(bytes)) {
    out.push_back(rec.type);
  }
  return out;
}

std::size_t Count(const std::vector<std::uint16_t>& types, std::uint16_t type) {
  std::size_t n = 0;
  for (std::uint16_t t : types) {
    n += t == type ? 1U : 0U;
  }
  return n;
}

// Deleting a sparkline's cell removes it, its group and the group
// container; deleting every CF cell removes the x14 block with the rule.
// Excel's output for these was not captured: this pins that nothing stale
// is left behind and that the future-record brackets stay balanced.
TEST(X14StructuralEditFixture, XlsbDeletionsDropEmptiedRecords) {
  const std::vector<std::uint8_t> base = ReadFixture("union_databar_0_base.xlsb");
  Workbook wb = Load(base, true);
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, 6, 1)));
  std::vector<std::uint16_t> types = XlsbSheetRecordTypes(Save(wb, true));
  EXPECT_EQ(Count(types, 1043U) + Count(types, 1041U) + Count(types, 1058U), 0U);
  EXPECT_EQ(Count(types, 1046U), 1U);
  EXPECT_EQ(Count(types, 35U), Count(types, 36U));

  Workbook cleared = Load(base, true);
  ASSERT_TRUE(static_cast<bool>(cleared.delete_cols(0, 0, 3)));
  const std::vector<std::uint8_t> saved = Save(cleared, true);
  types = XlsbSheetRecordTypes(saved);
  EXPECT_EQ(Count(types, 461U) + Count(types, 1046U) + Count(types, 1135U), 0U);
  EXPECT_EQ(Count(types, 1043U), 1U);
  EXPECT_EQ(Count(types, 35U), Count(types, 36U));
  EXPECT_TRUE(Load(saved, true).sheet(0).conditional_formats().empty());
}

// Formulas inside retained x14 records, against Mac Excel 365 (16.112)
// saves: an x14-only expression rule over `A1:A6 C1:C6` reading
// `Sheet2!$A$1<A1`, and a DataBar over G1:G6 whose x14 minimum is the
// formula `$E$1`. In the .xlsb these are a PtgRef3d plus a PtgRefN
// anchored to the range's top-left cell (BrtBeginCFRule14) and a PtgRef
// (BrtCFVO14). Edits on either sheet must move them as Excel does.

std::vector<std::string> XlsxFormulaCoordinates(const std::vector<std::uint8_t>& package) {
  const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.xml");
  std::vector<std::string> out;
  for (const char* attr : {"<conditionalFormatting sqref=\"", "<cfvo type=\"formula\" val=\""}) {
    const std::string open(attr);
    for (std::size_t at = sheet.find(open); at != std::string::npos; at = sheet.find(open, at)) {
      at += open.size();
      out.push_back(open.substr(1, 4) + " " + sheet.substr(at, sheet.find('"', at) - at));
    }
  }
  const std::string ext = sheet.substr(sheet.rfind("<extLst>"));
  for (const char* tag : {"xm:f", "xm:sqref"}) {
    for (const std::string& text : Elements(ext, tag)) {
      out.push_back(std::string(tag) + " " + text);
    }
  }
  return out;
}

std::vector<std::string> XlsbFormulaCoordinates(const std::vector<std::uint8_t>& package) {
  const std::string sheet = ReadEntry(package, "xl/worksheets/sheet1.bin");
  const std::vector<std::uint8_t> bytes(sheet.begin(), sheet.end());
  std::vector<std::string> out;
  for (const xlsb::FramedRecord& rec : xlsb::split_records(bytes)) {
    if (rec.type != 1046U && rec.type != 1048U && rec.type != 1050U) {
      continue;
    }
    std::string hex = std::to_string(rec.type);
    for (std::size_t i = 0; i < rec.payload.size; ++i) {
      char buf[4];
      std::snprintf(buf, sizeof(buf), " %02x", rec.payload.data[i]);
      hex += buf;
    }
    out.push_back(hex);
  }
  return out;
}

void CheckFormulaBook(bool xlsb) {
  const auto coords = [xlsb](const std::vector<std::uint8_t>& package) {
    return xlsb ? XlsbFormulaCoordinates(package) : XlsxFormulaCoordinates(package);
  };
  const char* ext = xlsb ? ".xlsb" : ".xlsx";
  const std::vector<std::uint8_t> base = ReadFixture(std::string("x14_formula_edit_0_base") + ext, "");
  ASSERT_EQ(coords(base).size(), 6U);
  EXPECT_EQ(coords(Save(Load(base, xlsb), xlsb)), coords(base)) << ext;
  const EditCase edits[] = {
      {"1_s1_insert_row2", [](Workbook& wb) { return wb.insert_rows(0, 1, 1); }},
      {"2_s1_delete_colB", [](Workbook& wb) { return wb.delete_cols(0, 1, 1); }},
      {"3_s2_insert_row1", [](Workbook& wb) { return wb.insert_rows(1, 0, 1); }},
      {"4_s1_insert_row1", [](Workbook& wb) { return wb.insert_rows(0, 0, 1); }},
  };
  for (const EditCase& edit : edits) {
    Workbook wb = Load(base, xlsb);
    ASSERT_TRUE(static_cast<bool>(edit.apply(wb))) << edit.state;
    EXPECT_EQ(coords(Save(wb, xlsb)), coords(ReadFixture(std::string("x14_formula_edit_") + edit.state + ext, "")))
        << edit.state << ext;
  }
}

TEST(X14StructuralEditFixture, RuleFormulasXlsx) {
  CheckFormulaBook(false);
}
TEST(X14StructuralEditFixture, RuleFormulasXlsb) {
  CheckFormulaBook(true);
}

}  // namespace
}  // namespace io
}  // namespace formulon
