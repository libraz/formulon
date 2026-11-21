// The metadata part as a model in both formats: Excel's `xl/metadata.xml` and
// `xl/metadata.bin` of one book parse to the same model, each serialises back
// to the other format's bytes and to its own, an attribute or flag the model
// does not carry stops a part reproducing, and the value-entry count survives
// a part the model refuses.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/rich_value_errors.h"
#include "io/xlsb/metadata_bin.h"
#include "io/zip_reader.h"
#include "support/roundtrip_symmetry.h"

namespace formulon {
namespace io {
namespace {

// Excel's xlsx / xlsb pairs of one book each, with the value entries their metadata parts hold.
struct Pair {
  const char* name;
  std::size_t value_entries;
};
constexpr Pair kPairs[] = {
    {"one_picture", 1U}, {"two_pictures", 2U}, {"picture_and_text", 1U}, {"seven_items", 3U}, {"numeric_field", 1U},
};

std::vector<std::uint8_t> PackagePart(const std::string& file, std::string_view part) {
  const std::vector<std::uint8_t> package =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_rich_records/" + file);
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package)))) << file;
  auto bytes = zip.read_entry(part);
  EXPECT_TRUE(static_cast<bool>(bytes)) << file << " " << part;
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

std::vector<std::uint8_t> MetadataXml(const Pair& pair) {
  return PackagePart(std::string(pair.name) + ".xlsx", "xl/metadata.xml");
}

std::vector<std::uint8_t> MetadataBin(const Pair& pair) {
  return PackagePart(std::string(pair.name) + ".xlsb", "xl/metadata.bin");
}

std::vector<std::uint8_t> Bytes(const std::string& text) {
  return std::vector<std::uint8_t>(text.begin(), text.end());
}

std::string Text(const std::vector<std::uint8_t>& bytes) {
  return std::string(bytes.begin(), bytes.end());
}

/// `bytes` with the one occurrence of `from` replaced by `to`.
std::vector<std::uint8_t> Replace(const std::vector<std::uint8_t>& bytes, const std::vector<std::uint8_t>& from,
                                  const std::vector<std::uint8_t>& to) {
  std::string text = Text(bytes);
  const std::string needle = Text(from);
  const std::size_t at = text.find(needle);
  EXPECT_NE(at, std::string::npos);
  EXPECT_EQ(text.find(needle, at + 1U), std::string::npos);
  if (at != std::string::npos) {
    text.replace(at, needle.size(), Text(to));
  }
  return Bytes(text);
}

TEST(MetadataModelCodec, ExcelPairsParseToOneModel) {
  for (const Pair& pair : kPairs) {
    SCOPED_TRACE(pair.name);
    const std::optional<MetadataModel> xml = parse_metadata_xml_model(MetadataXml(pair));
    const std::vector<std::uint8_t> bin_bytes = MetadataBin(pair);
    const std::optional<MetadataModel> bin = xlsb::parse_metadata_bin_model(test::span_of(bin_bytes));
    ASSERT_TRUE(xml.has_value());
    ASSERT_TRUE(bin.has_value());
    EXPECT_TRUE(*xml == *bin);
    EXPECT_EQ(xml->value_entries.size(), pair.value_entries);
  }
}

TEST(MetadataModelCodec, EachFormatReproducesItsOwnBytes) {
  for (const Pair& pair : kPairs) {
    SCOPED_TRACE(pair.name);
    const std::vector<std::uint8_t> xml_bytes = MetadataXml(pair);
    const std::vector<std::uint8_t> bin_bytes = MetadataBin(pair);
    const std::optional<MetadataModel> xml = parse_metadata_xml_model(xml_bytes);
    const std::optional<MetadataModel> bin = xlsb::parse_metadata_bin_model(test::span_of(bin_bytes));
    ASSERT_TRUE(xml.has_value());
    ASSERT_TRUE(bin.has_value());
    EXPECT_EQ(serialize_metadata_xml(*xml), Text(xml_bytes));
    EXPECT_EQ(xlsb::serialize_metadata_bin(*bin), bin_bytes);
  }
}

TEST(MetadataModelCodec, EachFormatSerialisesToTheOthersExcelBytes) {
  for (const Pair& pair : kPairs) {
    SCOPED_TRACE(pair.name);
    const std::vector<std::uint8_t> xml_bytes = MetadataXml(pair);
    const std::vector<std::uint8_t> bin_bytes = MetadataBin(pair);
    const std::optional<MetadataModel> xml = parse_metadata_xml_model(xml_bytes);
    const std::optional<MetadataModel> bin = xlsb::parse_metadata_bin_model(test::span_of(bin_bytes));
    ASSERT_TRUE(xml.has_value());
    ASSERT_TRUE(bin.has_value());
    EXPECT_EQ(xlsb::serialize_metadata_bin(*xml), bin_bytes);
    EXPECT_EQ(serialize_metadata_xml(*bin), Text(xml_bytes));

    const std::vector<std::uint8_t> as_bin = xlsb::serialize_metadata_bin(*xml);
    const std::optional<MetadataModel> back = xlsb::parse_metadata_bin_model(test::span_of(as_bin));
    ASSERT_TRUE(back.has_value());
    EXPECT_TRUE(*back == *xml);
    const std::optional<MetadataModel> back_xml = parse_metadata_xml_model(Bytes(serialize_metadata_xml(*bin)));
    ASSERT_TRUE(back_xml.has_value());
    EXPECT_TRUE(*back_xml == *bin);
  }
}

TEST(MetadataModelCodec, ExistingBuildersAreTheSerialisedModel) {
  for (const bool dynamic_arrays : {false, true}) {
    for (const std::size_t rich_values : {std::size_t{0}, std::size_t{1}, std::size_t{3}}) {
      for (const bool rich_first : {false, true}) {
        const MetadataModel model = metadata_model(dynamic_arrays, rich_values, rich_first);
        EXPECT_EQ(build_metadata_xml(dynamic_arrays, rich_values, rich_first), serialize_metadata_xml(model));
        EXPECT_EQ(xlsb::build_metadata_bin(dynamic_arrays, rich_values, rich_first),
                  xlsb::serialize_metadata_bin(model));
        const std::vector<std::uint8_t> bin = xlsb::serialize_metadata_bin(model);
        const std::optional<MetadataModel> parsed = xlsb::parse_metadata_bin_model(test::span_of(bin));
        ASSERT_TRUE(parsed.has_value());
        EXPECT_TRUE(*parsed == model);
      }
    }
  }
}

TEST(MetadataModelCodec, AnUnmodelledAttributeStopsTheXmlReproducing) {
  const std::vector<std::uint8_t> excel = MetadataXml(kPairs[0]);
  const std::vector<std::uint8_t> changed =
      Replace(excel, Bytes("minSupportedVersion=\"120000\""), Bytes("minSupportedVersion=\"120001\""));
  const std::optional<MetadataModel> model = parse_metadata_xml_model(changed);
  ASSERT_TRUE(model.has_value());
  EXPECT_NE(serialize_metadata_xml(*model), Text(changed));
}

TEST(MetadataModelCodec, AnUnmodelledFlagStopsTheBinReproducing) {
  const std::vector<std::uint8_t> excel = MetadataBin(kPairs[0]);
  // The XLRICHVALUE type record's first flag word, 0x986AC0B0.
  const std::vector<std::uint8_t> changed = Replace(excel, {0xB0, 0xC0, 0x6A, 0x98}, {0xB1, 0xC0, 0x6A, 0x98});
  const std::optional<MetadataModel> model = xlsb::parse_metadata_bin_model(test::span_of(changed));
  ASSERT_TRUE(model.has_value());
  EXPECT_NE(xlsb::serialize_metadata_bin(*model), changed);
}

TEST(MetadataModelCodec, AnUnknownTypeIsRefusedButItsValueEntriesCount) {
  for (const Pair& pair : kPairs) {
    SCOPED_TRACE(pair.name);
    const std::vector<std::uint8_t> changed =
        Replace(MetadataXml(pair), Bytes("<metadataTypes count=\"1\">"),
                Bytes("<metadataTypes count=\"2\"><metadataType name=\"XLFUTURE\" minSupportedVersion=\"120000\"/>"));
    EXPECT_FALSE(parse_metadata_xml_model(changed).has_value());
    EXPECT_FALSE(parse_metadata_xml(changed).consumable);
    EXPECT_EQ(metadata_xml_value_entry_count(changed), std::optional<std::size_t>(pair.value_entries));
  }
}

TEST(MetadataModelCodec, AnotherSetWordIsRefusedButItsValueEntriesCount) {
  for (const Pair& pair : kPairs) {
    SCOPED_TRACE(pair.name);
    // The first rich-value block: record 5003 (`8B 27`), an 8-byte payload, then the set word 0.
    const std::vector<std::uint8_t> changed =
        Replace(MetadataBin(pair), {0x8B, 0x27, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00},
                {0x8B, 0x27, 0x08, 0x01, 0x00, 0x00, 0x00, 0x00});
    EXPECT_FALSE(xlsb::parse_metadata_bin_model(test::span_of(changed)).has_value());
    EXPECT_EQ(xlsb::metadata_bin_value_entry_count(test::span_of(changed)),
              std::optional<std::size_t>(pair.value_entries));
  }
}

TEST(MetadataModelCodec, ExcelPartsCountTheirValueEntries) {
  for (const Pair& pair : kPairs) {
    SCOPED_TRACE(pair.name);
    const std::vector<std::uint8_t> bin = MetadataBin(pair);
    EXPECT_EQ(metadata_xml_value_entry_count(MetadataXml(pair)), std::optional<std::size_t>(pair.value_entries));
    EXPECT_EQ(xlsb::metadata_bin_value_entry_count(test::span_of(bin)), std::optional<std::size_t>(pair.value_entries));
  }
}

TEST(MetadataModelCodec, APartWithoutValueEntriesCountsZero) {
  EXPECT_EQ(metadata_xml_value_entry_count(Bytes(build_metadata_xml(true, 0U))), std::optional<std::size_t>(0U));
  const std::vector<std::uint8_t> bin = xlsb::build_metadata_bin(true, 0U);
  EXPECT_EQ(xlsb::metadata_bin_value_entry_count(test::span_of(bin)), std::optional<std::size_t>(0U));
}

TEST(MetadataModelCodec, UnreadablePartsCannotBeCounted) {
  EXPECT_FALSE(metadata_xml_value_entry_count(Bytes("<metadata")).has_value());
  const std::string two_lists =
      "<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<valueMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></valueMetadata>"
      "<valueMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></valueMetadata></metadata>";
  EXPECT_FALSE(metadata_xml_value_entry_count(Bytes(two_lists)).has_value());
  const std::vector<std::uint8_t> miscounted =
      Replace(MetadataXml(kPairs[3]), Bytes("<valueMetadata count=\"3\">"), Bytes("<valueMetadata count=\"4\">"));
  EXPECT_FALSE(metadata_xml_value_entry_count(miscounted).has_value());
  // seven_items: BrtBeginEsfmd (`D1 02`, 8 bytes) declaring 3 value entries, list 0.
  const std::vector<std::uint8_t> bin_miscounted =
      Replace(MetadataBin(kPairs[3]), {0xD1, 0x02, 0x08, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00},
              {0xD1, 0x02, 0x08, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00});
  EXPECT_FALSE(xlsb::metadata_bin_value_entry_count(test::span_of(bin_miscounted)).has_value());

  std::vector<std::uint8_t> cut = MetadataBin(kPairs[0]);
  cut.resize(cut.size() - 3U);
  EXPECT_FALSE(xlsb::metadata_bin_value_entry_count(test::span_of(cut)).has_value());
}

}  // namespace
}  // namespace io
}  // namespace formulon
