// Copyright 2026 libraz. Licensed under the MIT License.
//
// Integration tests for the OOXML writer empty-workbook slice. Each test
// round-trips the byte stream produced by `Workbook::save()` through miniz
// and/or pugixml to assert structural expectations: required parts are
// present, the XML is well-formed, and sheet-name mutations propagate
// through to workbook.xml.

#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <set>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "miniz.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

/// Extracts the entry at `name` from the ZIP in `archive_bytes`. Returns the
/// extracted bytes as a `std::string` (binary-safe) and asserts the entry
/// exists along the way. On failure the string is empty and an ASSERT fires
/// in the caller's scope.
std::string ExtractEntry(const std::vector<std::uint8_t>& archive_bytes, std::string_view name) {
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, archive_bytes.data(), archive_bytes.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return {};
  }
  const int index = mz_zip_reader_locate_file(&reader, std::string(name).c_str(), nullptr, 0);
  if (index < 0) {
    ADD_FAILURE() << "entry not found: " << name;
    mz_zip_reader_end(&reader);
    return {};
  }
  std::size_t extracted_size = 0;
  void* extracted = mz_zip_reader_extract_to_heap(&reader, static_cast<mz_uint>(index), &extracted_size, 0);
  if (extracted == nullptr) {
    ADD_FAILURE() << "extract_to_heap failed for: " << name;
    mz_zip_reader_end(&reader);
    return {};
  }
  std::string body(static_cast<const char*>(extracted), extracted_size);
  mz_free(extracted);
  mz_zip_reader_end(&reader);
  return body;
}

std::set<std::string> ListEntries(const std::vector<std::uint8_t>& archive_bytes) {
  std::set<std::string> names;
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, archive_bytes.data(), archive_bytes.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return names;
  }
  const mz_uint entry_count = mz_zip_reader_get_num_files(&reader);
  for (mz_uint i = 0; i < entry_count; ++i) {
    char name_buf[256] = {};
    const mz_uint len = mz_zip_reader_get_filename(&reader, i, name_buf, sizeof(name_buf));
    if (len > 0) {
      // miniz returns length including the trailing NUL.
      names.emplace(name_buf);
    }
  }
  mz_zip_reader_end(&reader);
  return names;
}

TEST(OoxmlMinimalWriter, RoundTripContainsAllRequiredParts) {
  Workbook wb = Workbook::create();
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result)) << "save() failed: " << result.error().message;
  const std::vector<std::uint8_t>& bytes = result.value();

  const std::set<std::string> entries = ListEntries(bytes);
  EXPECT_TRUE(entries.count("[Content_Types].xml") == 1) << "missing [Content_Types].xml";
  EXPECT_TRUE(entries.count("_rels/.rels") == 1) << "missing _rels/.rels";
  EXPECT_TRUE(entries.count("xl/workbook.xml") == 1) << "missing xl/workbook.xml";
  EXPECT_TRUE(entries.count("xl/_rels/workbook.xml.rels") == 1) << "missing xl/_rels/workbook.xml.rels";
  EXPECT_TRUE(entries.count("xl/worksheets/sheet1.xml") == 1) << "missing xl/worksheets/sheet1.xml";
  EXPECT_TRUE(entries.count("xl/styles.xml") == 1) << "missing xl/styles.xml";
}

TEST(OoxmlMinimalWriter, WorkbookXmlIsWellFormed) {
  Workbook wb = Workbook::create();
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));
  const std::string body = ExtractEntry(result.value(), "xl/workbook.xml");
  ASSERT_FALSE(body.empty());

  pugi::xml_document doc;
  pugi::xml_parse_result parse = doc.load_buffer(body.data(), body.size());
  ASSERT_TRUE(static_cast<bool>(parse)) << "workbook.xml parse failed: " << parse.description();

  pugi::xml_node root = doc.child("workbook");
  ASSERT_TRUE(static_cast<bool>(root));

  pugi::xml_node sheets = root.child("sheets");
  ASSERT_TRUE(static_cast<bool>(sheets));

  int sheet_count = 0;
  std::string first_name;
  for (pugi::xml_node sheet = sheets.child("sheet"); sheet; sheet = sheet.next_sibling("sheet")) {
    if (sheet_count == 0) {
      first_name = sheet.attribute("name").value();
    }
    ++sheet_count;
  }
  EXPECT_EQ(sheet_count, 1);
  EXPECT_EQ(first_name, "Sheet1");
}

TEST(OoxmlMinimalWriter, ContentTypesIsWellFormed) {
  Workbook wb = Workbook::create();
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));
  const std::string body = ExtractEntry(result.value(), "[Content_Types].xml");
  ASSERT_FALSE(body.empty());

  pugi::xml_document doc;
  pugi::xml_parse_result parse = doc.load_buffer(body.data(), body.size());
  ASSERT_TRUE(static_cast<bool>(parse)) << "[Content_Types].xml parse failed: " << parse.description();

  pugi::xml_node root = doc.child("Types");
  ASSERT_TRUE(static_cast<bool>(root));

  int default_count = 0;
  int override_count = 0;
  for (pugi::xml_node child = root.first_child(); child; child = child.next_sibling()) {
    const std::string_view name = child.name();
    if (name == "Default") {
      ++default_count;
    } else if (name == "Override") {
      ++override_count;
    }
  }
  EXPECT_EQ(default_count, 2);
  EXPECT_EQ(override_count, 3);
}

TEST(OoxmlMinimalWriter, SheetRenamePropagatesToWorkbookXml) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_name("\xE3\x82\xAB\xE3\x82\xB9\xE3\x82\xBF\xE3\x83\xA0");  // "カスタム" in UTF-8
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));

  const std::string body = ExtractEntry(result.value(), "xl/workbook.xml");
  ASSERT_FALSE(body.empty());

  pugi::xml_document doc;
  pugi::xml_parse_result parse = doc.load_buffer(body.data(), body.size());
  ASSERT_TRUE(static_cast<bool>(parse)) << "workbook.xml parse failed: " << parse.description();

  pugi::xml_node sheet = doc.child("workbook").child("sheets").child("sheet");
  ASSERT_TRUE(static_cast<bool>(sheet));
  // pugixml returns UTF-8 bytes verbatim; compare against the raw UTF-8.
  EXPECT_STREQ(sheet.attribute("name").value(), "\xE3\x82\xAB\xE3\x82\xB9\xE3\x82\xBF\xE3\x83\xA0");
}

// ---------------------------------------------------------------------------
// Cell-writer integration: round-trip through save() and pugixml.
// ---------------------------------------------------------------------------

TEST(CellRoundTrip, NumberSavesAndExtracts) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0U, 0U, Value::number(42.5));
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result)) << "save() failed: " << result.error().message;

  const std::string body = ExtractEntry(result.value(), "xl/worksheets/sheet1.xml");
  ASSERT_FALSE(body.empty());

  pugi::xml_document doc;
  pugi::xml_parse_result parse = doc.load_buffer(body.data(), body.size());
  ASSERT_TRUE(static_cast<bool>(parse)) << "sheet1.xml parse failed: " << parse.description();

  pugi::xml_node v = doc.child("worksheet").child("sheetData").child("row").child("c").child("v");
  ASSERT_TRUE(static_cast<bool>(v));
  EXPECT_STREQ(v.text().get(), "42.5");
}

TEST(CellRoundTrip, BlankOmitted) {
  Workbook wb = Workbook::create();
  // Default sheet: no cells written. sheetData should be empty (no <c>).
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));

  const std::string body = ExtractEntry(result.value(), "xl/worksheets/sheet1.xml");
  ASSERT_FALSE(body.empty());

  pugi::xml_document doc;
  pugi::xml_parse_result parse = doc.load_buffer(body.data(), body.size());
  ASSERT_TRUE(static_cast<bool>(parse));

  pugi::xml_node sheet_data = doc.child("worksheet").child("sheetData");
  // The node may or may not exist depending on self-closing vs explicit
  // form. Either way, no <c> children are allowed.
  if (sheet_data) {
    EXPECT_FALSE(static_cast<bool>(sheet_data.child("row")));
    EXPECT_FALSE(static_cast<bool>(sheet_data.child("c")));
  }
}

TEST(CellRoundTrip, UnicodeText) {
  Workbook wb = Workbook::create();
  // "日本" in UTF-8: E6 97 A5 E6 9C AC.
  wb.sheet(0).set_cell_value(0U, 0U, Value::text("\xE6\x97\xA5\xE6\x9C\xAC"));
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));

  const std::string body = ExtractEntry(result.value(), "xl/worksheets/sheet1.xml");
  ASSERT_NE(body.find("\xE6\x97\xA5\xE6\x9C\xAC"), std::string::npos) << body;
}

TEST(SpillRoundTrip, AnchorHasArrayType) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_formula(0U, 0U, "=SEQUENCE(3)");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(wb.sheet(0).commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));

  const std::string body = ExtractEntry(result.value(), "xl/worksheets/sheet1.xml");
  ASSERT_NE(body.find("<f t=\"array\">SEQUENCE(3)</f>"), std::string::npos) << body;
}

TEST(SpillRoundTrip, PhantomCellsAbsent) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_formula(0U, 0U, "=SEQUENCE(2,2)");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0)};
  ASSERT_TRUE(wb.sheet(0).commit_spill(0U, 0U, 2U, 2U, std::move(cells)));

  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));

  const std::string body = ExtractEntry(result.value(), "xl/worksheets/sheet1.xml");
  // The anchor must be present.
  EXPECT_NE(body.find("r=\"A1\""), std::string::npos) << body;
  // Phantoms must be absent from the worksheet XML.
  EXPECT_EQ(body.find("r=\"B1\""), std::string::npos) << body;
  EXPECT_EQ(body.find("r=\"A2\""), std::string::npos) << body;
  EXPECT_EQ(body.find("r=\"B2\""), std::string::npos) << body;
}

}  // namespace
}  // namespace formulon
