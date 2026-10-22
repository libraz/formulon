// Slicers in XLSB: the workbook's slicer cache lists (SLICERCACHEIDS /
// TABLESLICERCACHEIDS) are kept from an .xlsb source and re-emitted with the
// relationship ids the writer assigns; the XML slicer parts of an .xlsx source
// have no binary form and are left out and counted as deferred.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "drawing/slicer_edit.h"
#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/slicer_records.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "table.h"
#include "unknown_relationship.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr const char* kCacheRel = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
constexpr const char* kCachePath = "xl/slicerCaches/slicerCache1.bin";
constexpr const char* kCacheCt = "application/vnd.ms-excel.slicerCache";

/// An FRTHeader carrying only fRelID, naming `id`.
std::vector<std::uint8_t> RelIdPayload(const std::string& id) {
  std::vector<std::uint8_t> p;
  emit_u32(p, 0x08U);
  emit_u16(p, static_cast<std::uint16_t>(id.size()));
  for (const char c : id) {
    emit_u16(p, static_cast<std::uint16_t>(c));
  }
  return p;
}

/// TABLESLICERCACHEIDS naming `id`.
std::vector<std::uint8_t> TableCacheIds(const std::string& id) {
  std::vector<std::uint8_t> out;
  const std::vector<std::uint8_t> version = {0x02, 0x0E, 0x00, 0x00};
  emit_record(out, 35U, version);
  emit_record(out, 2075U, ByteSpan{});
  emit_record(out, 2076U, RelIdPayload(id));
  emit_record(out, 36U, ByteSpan{});
  return out;
}

Workbook BookWithRetainedCache() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_value(0U, 0U, Value::number(1.0));
  wb.set_passthrough_parts({PassthroughPart(kCachePath, kCacheCt, std::vector<std::uint8_t>{0x01, 0x02})});
  wb.set_unknown_workbook_rels({UnknownRelationship{"rId42", kCacheRel, kCachePath, false}});
  wb.set_xlsb_book_frt(TableCacheIds("rId42"));
  return wb;
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb) {
  auto bytes = write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(bytes)) << (bytes ? "" : bytes.error().message);
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

/// The relationship ids the retained TableSlicerCacheID records name.
std::vector<std::string> CacheRelIds(const std::vector<std::uint8_t>& frt) {
  std::vector<std::string> out;
  for (const FramedRecord& rec : split_records(frt)) {
    if (rec.type != 2076U || rec.payload.size < 6U) {
      continue;
    }
    const std::size_t cch = static_cast<std::size_t>(rec.payload.data[4] | (rec.payload.data[5] << 8U));
    std::string id;
    for (std::size_t i = 0; i < cch && 7U + 2U * i < rec.payload.size; ++i) {
      id.push_back(static_cast<char>(rec.payload.data[6U + 2U * i]));
    }
    out.push_back(id);
  }
  return out;
}

std::string RelIdFor(const Workbook& wb, const std::string& target) {
  for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
    if (rel.target == target) {
      return rel.id;
    }
  }
  return {};
}

TEST(XlsbSlicerFidelity, CacheListFollowsTheWrittenRelationshipId) {
  const Workbook back = LoadXlsb(SaveXlsb(BookWithRetainedCache()));
  const std::string rid = RelIdFor(back, kCachePath);
  ASSERT_FALSE(rid.empty());
  EXPECT_NE(rid, "rId42");
  EXPECT_EQ(CacheRelIds(back.xlsb_book_frt()), std::vector<std::string>{rid});
  // A second save keeps it pointing at the cache part.
  const Workbook again = LoadXlsb(SaveXlsb(back));
  EXPECT_EQ(CacheRelIds(again.xlsb_book_frt()), std::vector<std::string>{RelIdFor(again, kCachePath)});
}

TEST(XlsbSlicerFidelity, CacheListWithoutItsPartIsDropped) {
  Workbook wb = BookWithRetainedCache();
  wb.set_passthrough_parts({});
  const Workbook back = LoadXlsb(SaveXlsb(wb));
  EXPECT_TRUE(back.xlsb_book_frt().empty());
}

TEST(XlsbSlicerFidelity, CacheListSitsBeforeEndBook) {
  const std::vector<std::uint8_t> bytes = SaveXlsb(BookWithRetainedCache());
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  auto body = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(body));
  const std::vector<FramedRecord> recs = split_records(body.value());
  ASSERT_GE(recs.size(), 5U);
  EXPECT_EQ(recs.back().type, 132U);
  EXPECT_EQ(recs[recs.size() - 2U].type, 36U);
  EXPECT_EQ(recs[recs.size() - 3U].type, 2076U);
  EXPECT_EQ(recs[recs.size() - 4U].type, 2075U);
}

Workbook BookWithXmlSlicer(bool slicer) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0U, 0U, Value::text("Product"));
  sheet.set_cell_value(0U, 1U, Value::text("Amount"));
  sheet.set_cell_value(1U, 0U, Value::text("Apple"));
  sheet.set_cell_value(1U, 1U, Value::number(3.0));
  TableMetadata table;
  table.sheet_index = 0;
  table.id = 1;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:B2";
  for (const char* name : {"Product", "Amount"}) {
    TableColumn column;
    column.id = static_cast<std::uint32_t>(table.columns.size() + 1U);
    column.name = name;
    table.columns.push_back(column);
  }
  wb.mutable_tables().push_back(table);
  if (slicer) {
    SlicerSpec spec;
    spec.fields = slicer_field::kAnchor;
    spec.from = AnchorPoint{4U, 0U, 0, 0};
    spec.to = AnchorPoint{12U, 2U, 0, 0};
    EXPECT_TRUE(static_cast<bool>(add_slicer(wb, 0, 0, 0, spec)));
  }
  return wb;
}

TEST(XlsbSlicerFidelity, XmlSlicerPartsAreLeftOutAndCounted) {
  auto with = write_xlsb_with_result(BookWithXmlSlicer(true));
  auto without = write_xlsb_with_result(BookWithXmlSlicer(false));
  ASSERT_TRUE(static_cast<bool>(with));
  ASSERT_TRUE(static_cast<bool>(without));
  // The slicer, and the drawing that holds its frame.
  EXPECT_EQ(with.value().diagnostics.deferred_feature_count, without.value().diagnostics.deferred_feature_count + 2U);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(with.value().bytes))));
  EXPECT_FALSE(static_cast<bool>(zip.read_entry("xl/slicers/slicer1.xml")));
  EXPECT_FALSE(static_cast<bool>(zip.read_entry("xl/slicerCaches/slicerCache1.xml")));
  auto types = zip.read_entry("[Content_Types].xml");
  ASSERT_TRUE(static_cast<bool>(types));
  EXPECT_EQ(std::string(types.value().begin(), types.value().end()).find("slicer"), std::string::npos);
}

TEST(XlsbSlicerFidelity, XmlSlicerPartPathsAreRecognised) {
  EXPECT_TRUE(is_xml_slicer_part("xl/slicers/slicer1.xml"));
  EXPECT_TRUE(is_xml_slicer_part("xl/slicerCaches/slicerCache2.xml"));
  EXPECT_FALSE(is_xml_slicer_part("xl/slicers/slicer1.bin"));
  EXPECT_FALSE(is_xml_slicer_part("xl/slicers/_rels/slicer1.xml.rels"));
  EXPECT_FALSE(is_xml_slicer_part("xl/worksheets/sheet1.xml"));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
