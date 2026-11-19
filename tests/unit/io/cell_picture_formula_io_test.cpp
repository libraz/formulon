// Pictures as formula values in a saved workbook: a load and save with no
// recalc keeps the rich-value parts and every cell's `vm` as Excel wrote them.

#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

std::vector<std::uint8_t> ReadPicture(std::string_view name) {
  std::vector<std::uint8_t> bytes =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/cell_pictures/" + std::string(name) + ".xlsx");
  EXPECT_FALSE(bytes.empty()) << name;
  return bytes;
}

/// The text of `part`, or nullopt when the package has none.
std::optional<std::string> Part(const std::vector<std::uint8_t>& package, std::string_view part) {
  ZipReader zip;
  if (!zip.open(test::span_of(package)) || !zip.has_entry(part)) {
    return std::nullopt;
  }
  auto bytes = zip.read_entry(part);
  if (!bytes) {
    return std::nullopt;
  }
  return std::string(bytes.value().begin(), bytes.value().end());
}

/// `vm` of every `<c>` in `sheet_part` that carries one, by A1 reference.
std::map<std::string, std::string> VmCells(const std::vector<std::uint8_t>& package, std::string_view sheet_part) {
  std::map<std::string, std::string> out;
  const std::string xml = Part(package, sheet_part).value_or("");
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(xml.data(), xml.size())) << sheet_part;
  for (pugi::xml_node row : doc.child("worksheet").child("sheetData").children("row")) {
    for (pugi::xml_node c : row.children("c")) {
      if (c.attribute("vm")) {
        out[c.attribute("r").value()] = c.attribute("vm").value();
      }
    }
  }
  return out;
}

TEST(CellPictureFormulaIo, XlsxRoundTripIsByteExactForPictureFormulaRefs) {
  // E2 and E4 hold pictures; G2 is `=E2`, cached as E2's picture.
  const std::vector<std::uint8_t> source = ReadPicture("picture_formula_refs");
  auto loaded = read_ooxml(test::span_of(source));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto saved = write_ooxml(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const std::vector<std::uint8_t>& out = saved.value();

  for (std::string_view part :
       {"xl/metadata.xml", "xl/richData/rdrichvalue.xml", "xl/richData/rdrichvaluestructure.xml"}) {
    SCOPED_TRACE(std::string(part));
    const std::optional<std::string> want = Part(source, part);
    ASSERT_TRUE(want.has_value());
    EXPECT_EQ(Part(out, part), want);
  }
  const std::map<std::string, std::string> want = VmCells(source, "xl/worksheets/sheet1.xml");
  EXPECT_EQ(want.count("G2"), 1U);
  EXPECT_EQ(VmCells(out, "xl/worksheets/sheet1.xml"), want);
}

}  // namespace
}  // namespace io
}  // namespace formulon
