// XLSB writer round-trip tests: metadata parts and dynamic-array spills.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"
namespace formulon {
namespace io {
namespace xlsb {
namespace {

TEST(XlsbWriter, PassthroughPartsRoundTripVerbatim) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("S1"));
  s.set_cell_value(0U, 0U, Value::number(10.0));

  // Synthesize a passthrough part: pretend the original archive had
  // an `xl/theme/theme1.xml`. The bytes are arbitrary; the writer
  // copies them verbatim and the reader surfaces them again on
  // unknown_parts.
  std::vector<PassthroughPart> parts;
  PassthroughPart theme;
  theme.path = "xl/theme/theme1.xml";
  theme.content_type = "application/vnd.openxmlformats-officedocument.theme+xml";
  const std::string body = "<?xml version=\"1.0\"?><theme xmlns=\"x\"/>";
  theme.bytes.assign(body.begin(), body.end());
  parts.push_back(std::move(theme));
  PassthroughPart custom;
  custom.path = "docProps/custom.xml";
  custom.content_type = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
  custom.bytes = {'<', 'c', 'u', 's', 't', 'o', 'm', '/', '>'};
  parts.push_back(std::move(custom));
  PassthroughPart thumbnail;
  thumbnail.path = "docProps/thumbnail.jpeg";
  thumbnail.content_type = "image/jpeg";
  thumbnail.bytes = {0xffU, 0xd8U, 0xffU, 0xd9U};
  parts.push_back(std::move(thumbnail));
  wb.set_passthrough_parts(std::move(parts));
  wb.set_unknown_package_rels(
      {UnknownRelationship{"rId9", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
                           "docProps/thumbnail.jpeg", false}});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto package_rels_or = zip.read_entry("_rels/.rels");
  ASSERT_TRUE(static_cast<bool>(package_rels_or));
  const std::string package_rels(package_rels_or.value().begin(), package_rels_or.value().end());
  EXPECT_NE(package_rels.find("custom-properties"), std::string::npos);
  EXPECT_NE(package_rels.find("Target=\"docProps/custom.xml\""), std::string::npos);
  EXPECT_NE(package_rels.find("relationships/metadata/thumbnail"), std::string::npos);
  EXPECT_NE(package_rels.find("Target=\"docProps/thumbnail.jpeg\""), std::string::npos);

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;

  bool found_theme = false;
  for (const PassthroughPart& part : read_or.value().workbook.passthrough_parts()) {
    if (part.path == "xl/theme/theme1.xml") {
      found_theme = true;
      EXPECT_EQ(part.content_type, "application/vnd.openxmlformats-officedocument.theme+xml");
      const std::string round_tripped(part.bytes.begin(), part.bytes.end());
      EXPECT_EQ(round_tripped, body);
    }
  }
  EXPECT_TRUE(found_theme);
  ASSERT_EQ(read_or.value().workbook.unknown_package_rels().size(), 1U);
  EXPECT_EQ(read_or.value().workbook.unknown_package_rels()[0].target, "docProps/thumbnail.jpeg");
}

TEST(XlsbWriter, DropsXlsxMetadataAndItsWorkbookRelationship) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("S1"));
  sheet.set_cell_value(0U, 0U, Value::number(1.0));
  PassthroughPart metadata;
  metadata.path = "xl/metadata.xml";
  metadata.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
  metadata.bytes = {'<', 'm', '/', '>'};
  wb.set_passthrough_parts({metadata});
  wb.set_unknown_workbook_rels(
      {UnknownRelationship{"rId7", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
                           "xl/metadata.xml", false}});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  EXPECT_FALSE(zip.has_entry("xl/metadata.xml"));
  auto rels_or = zip.read_entry("xl/_rels/workbook.bin.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  EXPECT_EQ(rels.find("sheetMetadata"), std::string::npos);
}

TEST(XlsbWriter, EmitsDynamicArrayMetadataForSpillAnchors) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Spill"));
  sheet.set_cell_formula(0U, 0U, "=SEQUENCE(2)");
  sheet.set_cell_dynamic_array(0U, 0U, true);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 2U, 1U, {Value::number(1.0), Value::number(2.0)}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/metadata.bin"));

  auto metadata_or = zip.read_entry("xl/metadata.bin");
  ASSERT_TRUE(static_cast<bool>(metadata_or));
  ByteSpan metadata_cursor = SpanOf(metadata_or.value());
  auto metadata_record_or = read_record(metadata_cursor);
  ASSERT_TRUE(static_cast<bool>(metadata_record_or));
  EXPECT_EQ(metadata_record_or.value().type, 332U);  // BrtBeginMetadata

  auto content_types_or = zip.read_entry("[Content_Types].xml");
  ASSERT_TRUE(static_cast<bool>(content_types_or));
  const std::string content_types(content_types_or.value().begin(), content_types_or.value().end());
  EXPECT_NE(content_types.find("/xl/metadata.bin"), std::string::npos);
  EXPECT_NE(content_types.find("application/vnd.ms-excel.sheetMetadata"), std::string::npos);

  auto rels_or = zip.read_entry("xl/_rels/workbook.bin.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  EXPECT_NE(rels.find("relationships/sheetMetadata"), std::string::npos);
  EXPECT_NE(rels.find("Target=\"metadata.bin\""), std::string::npos);

  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  ByteSpan sheet_cursor = SpanOf(sheet_or.value());
  bool found_cell_metadata = false;
  while (sheet_cursor.size > 0U) {
    auto record_or = read_record(sheet_cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtCellMeta)) {
      continue;
    }
    ByteSpan payload = record_or.value().payload;
    auto index_or = read_u32(payload);
    ASSERT_TRUE(static_cast<bool>(index_or));
    EXPECT_EQ(index_or.value(), 1U);
    EXPECT_EQ(payload.size, 0U);
    found_cell_metadata = true;
  }
  EXPECT_TRUE(found_cell_metadata);
}

// ---------------------------------------------------------------------------
// A `BrtCellMeta` index names an entry of the metadata part that actually
// ships. The generated part declares one entry, so index 1 is right by
// construction; a retained passthrough part carries its own numbering, and
// its first entry may be rich-value or cube-function metadata rather than the
// dynamic-array one. An index that misses makes Excel repair the file, so an
// unidentifiable entry means no `BrtCellMeta` record at all.
// ---------------------------------------------------------------------------

namespace {

// Builds a metadata part declaring `type_names` in order, followed by one
// cell-metadata entry per name (entry N names type N). Mirrors the record
// framing Excel writes, so the writer's finder sees a realistic part.
std::vector<std::uint8_t> BuildMetadataPart(const std::vector<std::string>& type_names) {
  constexpr std::uint16_t kBrtBeginMetadata = 332;
  constexpr std::uint16_t kBrtEndMetadata = 333;
  constexpr std::uint16_t kBrtBeginEsmdtinfo = 334;
  constexpr std::uint16_t kBrtMdtinfo = 335;
  constexpr std::uint16_t kBrtEndEsmdtinfo = 336;
  constexpr std::uint16_t kBrtBeginEsfmd = 337;
  constexpr std::uint16_t kBrtEndEsfmd = 338;
  constexpr std::uint16_t kBrtMdb = 51;

  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_record(out, kBrtBeginMetadata, ByteSpan{});

  emit_u32(payload, static_cast<std::uint32_t>(type_names.size()));
  emit_record(out, kBrtBeginEsmdtinfo, payload);
  for (const std::string& name : type_names) {
    payload.clear();
    emit_u32(payload, 0xD86AC0B0U);
    emit_u32(payload, 0x0001D4C0U);
    emit_xlwidestring(payload, name);
    emit_record(out, kBrtMdtinfo, payload);
  }
  emit_record(out, kBrtEndEsmdtinfo, ByteSpan{});

  payload.clear();
  emit_u32(payload, static_cast<std::uint32_t>(type_names.size()));
  emit_u32(payload, 1U);
  emit_record(out, kBrtBeginEsfmd, payload);
  for (std::size_t i = 0; i < type_names.size(); ++i) {
    payload.clear();
    emit_u32(payload, 1U);                                 // One (type, id) pair.
    emit_u32(payload, static_cast<std::uint32_t>(i + 1));  // 1-based type ordinal.
    emit_u32(payload, 0U);
    emit_record(out, kBrtMdb, payload);
  }
  emit_record(out, kBrtEndEsfmd, ByteSpan{});
  emit_record(out, kBrtEndMetadata, ByteSpan{});
  return out;
}

Workbook SpillWorkbookWithRetainedMetadata(std::vector<std::uint8_t> metadata_bytes) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Spill"));
  sheet.set_cell_formula(0U, 0U, "=SEQUENCE(2)");
  sheet.set_cell_dynamic_array(0U, 0U, true);
  EXPECT_TRUE(sheet.commit_spill(0U, 0U, 2U, 1U, {Value::number(1.0), Value::number(2.0)}));
  PassthroughPart metadata;
  metadata.path = "xl/metadata.bin";
  metadata.content_type = "application/vnd.ms-excel.sheetMetadata";
  metadata.bytes = std::move(metadata_bytes);
  wb.set_passthrough_parts({metadata});
  return wb;
}

// Collects every `BrtCellMeta` index in a worksheet body, and reports whether
// the body carries an array-formula record at all.
struct SheetMetaScan {
  std::vector<std::uint32_t> cell_meta_indices;
  bool has_array_formula = false;
};

SheetMetaScan ScanSheetForCellMeta(const std::vector<std::uint8_t>& sheet_bytes) {
  SheetMetaScan scan;
  ByteSpan cursor = SpanOf(sheet_bytes);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    EXPECT_TRUE(static_cast<bool>(record_or));
    if (!record_or) {
      break;
    }
    if (record_or.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla)) {
      scan.has_array_formula = true;
      continue;
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtCellMeta)) {
      continue;
    }
    ByteSpan payload = record_or.value().payload;
    auto index_or = read_u32(payload);
    EXPECT_TRUE(static_cast<bool>(index_or));
    if (index_or) {
      scan.cell_meta_indices.push_back(index_or.value());
    }
  }
  return scan;
}

// Writes `wb`, then returns the scan of its first worksheet body.
SheetMetaScan WriteAndScanFirstSheet(const Workbook& wb) {
  auto bytes_or = write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or));
  if (!bytes_or) {
    return SheetMetaScan{};
  }
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(sheet_or));
  if (!sheet_or) {
    return SheetMetaScan{};
  }
  return ScanSheetForCellMeta(sheet_or.value());
}

}  // namespace

TEST(XlsbMetadataBin, FindsDynamicArrayEntryBehindOtherTypes) {
  const std::vector<std::uint8_t> part = BuildMetadataPart({"XLRICHVALUE", "XLDAPR", "XLMDX"});
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(part)), 2U);
}

TEST(XlsbMetadataBin, ReportsZeroWhenNoDynamicArrayTypeIsDeclared) {
  const std::vector<std::uint8_t> part = BuildMetadataPart({"XLRICHVALUE", "XLMDX"});
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(part)), 0U);
}

TEST(XlsbMetadataBin, GeneratedPartResolvesToIndexOne) {
  const std::vector<std::uint8_t> part = build_dynamic_array_metadata_bin();
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(part)), 1U);
}

TEST(XlsbMetadataBin, ReportsZeroForEveryTruncationOfAValidPart) {
  const std::vector<std::uint8_t> part = BuildMetadataPart({"XLRICHVALUE", "XLDAPR"});
  ASSERT_EQ(find_dynamic_array_cell_meta_index(SpanOf(part)), 2U);
  // Every strict prefix is either unreadable or missing the entry it would
  // have to name. None may produce an index.
  for (std::size_t length = 0; length < part.size(); ++length) {
    const ByteSpan prefix{part.data(), length};
    EXPECT_EQ(find_dynamic_array_cell_meta_index(prefix), 0U) << "prefix length " << length;
  }
}

TEST(XlsbMetadataBin, ReportsZeroForMalformedBytes) {
  EXPECT_EQ(find_dynamic_array_cell_meta_index(ByteSpan{}), 0U);
  // A record header promising more payload than the buffer holds.
  const std::vector<std::uint8_t> overrun = {0x4C, 0x02, 0x7F, 0x01, 0x02};
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(overrun)), 0U);
  // A type table whose entry payload ends before its name.
  std::vector<std::uint8_t> short_name;
  std::vector<std::uint8_t> payload;
  emit_u32(payload, 1U);
  emit_record(short_name, 334U, payload);
  payload.clear();
  emit_u32(payload, 0U);
  emit_record(short_name, 335U, payload);
  emit_record(short_name, 336U, ByteSpan{});
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(short_name)), 0U);
}

TEST(XlsbWriter, SpillAnchorNamesTheRetainedPartsDynamicArrayEntry) {
  const Workbook wb = SpillWorkbookWithRetainedMetadata(BuildMetadataPart({"XLRICHVALUE", "XLDAPR", "XLMDX"}));
  const SheetMetaScan scan = WriteAndScanFirstSheet(wb);
  ASSERT_EQ(scan.cell_meta_indices.size(), 1U);
  EXPECT_EQ(scan.cell_meta_indices[0], 2U) << "the anchor named the retained part's first entry instead of its "
                                              "dynamic-array entry";
  EXPECT_TRUE(scan.has_array_formula);
}

TEST(XlsbWriter, SpillAnchorEmitsNoCellMetaWhenRetainedPartHasNoDynamicArrayEntry) {
  const Workbook wb = SpillWorkbookWithRetainedMetadata(BuildMetadataPart({"XLRICHVALUE", "XLMDX"}));
  const SheetMetaScan scan = WriteAndScanFirstSheet(wb);
  EXPECT_TRUE(scan.cell_meta_indices.empty()) << "a BrtCellMeta index that names no dynamic-array entry";
  EXPECT_TRUE(scan.has_array_formula) << "the anchor must still be written as an array formula";
}

TEST(XlsbWriter, SpillAnchorEmitsNoCellMetaWhenRetainedPartIsMalformed) {
  const Workbook wb = SpillWorkbookWithRetainedMetadata({0x4C, 0x02, 0x7F, 0x01, 0x02});
  const SheetMetaScan scan = WriteAndScanFirstSheet(wb);
  EXPECT_TRUE(scan.cell_meta_indices.empty()) << "a dangling BrtCellMeta index from unreadable metadata bytes";
  EXPECT_TRUE(scan.has_array_formula);
}

TEST(XlsbWriter, RetainedMetadataPartShipsVerbatimAndIsNotRegenerated) {
  const std::vector<std::uint8_t> part = BuildMetadataPart({"XLRICHVALUE", "XLDAPR"});
  const Workbook wb = SpillWorkbookWithRetainedMetadata(part);
  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto shipped_or = zip.read_entry("xl/metadata.bin");
  ASSERT_TRUE(static_cast<bool>(shipped_or));
  EXPECT_EQ(shipped_or.value(), part) << "the generated part displaced the retained one";
  // The index the sheet carries must resolve inside the part that shipped.
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  const SheetMetaScan scan = ScanSheetForCellMeta(sheet_or.value());
  ASSERT_EQ(scan.cell_meta_indices.size(), 1U);
  EXPECT_EQ(scan.cell_meta_indices[0], find_dynamic_array_cell_meta_index(SpanOf(shipped_or.value())));
}

TEST(XlsbWriter, EmitsDynamicArrayMetadataForSingleCellArrayAnchors) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("SingleArray"));
  sheet.set_cell_formula(0U, 0U, "=IFS(TRUE,\"yes\")");
  sheet.set_cell_dynamic_array(0U, 0U, true);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 1U, 1U, {Value::text("yes")}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/metadata.bin"));

  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  ByteSpan cursor = SpanOf(sheet_or.value());
  bool found_cell_metadata = false;
  bool found_array_formula = false;
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    found_cell_metadata =
        found_cell_metadata || record_or.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtCellMeta);
    found_array_formula =
        found_array_formula || record_or.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla);
  }
  EXPECT_TRUE(found_cell_metadata);
  EXPECT_TRUE(found_array_formula);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
