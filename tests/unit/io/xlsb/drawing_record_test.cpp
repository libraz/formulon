// BrtDrawing in a retained worksheet tail: reading the one relationship id,
// splicing the record into its grammar slot, and the drawing-part frame check.

#include "io/xlsb/drawing_record.h"

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;

void Append(Bytes& buf, std::uint16_t type, const Bytes& payload = {}) {
  emit_record(buf, type, payload);
}

Bytes WideString(const std::string& text) {
  Bytes payload;
  emit_xlwidestring(payload, text);
  return payload;
}

std::vector<std::uint16_t> Types(const Bytes& buf) {
  std::vector<std::uint16_t> out;
  for (const FramedRecord& rec : split_records(buf)) {
    out.push_back(rec.type);
  }
  return out;
}

/// The relationship ids of every BrtDrawing record in `buf`.
std::vector<std::string> DrawingIds(const Bytes& buf) {
  std::vector<std::string> out;
  for (const FramedRecord& rec : split_records(buf)) {
    if (rec.type != kBrtDrawing) {
      continue;
    }
    ByteSpan payload = rec.payload;
    auto id = read_xlwidestring(payload);
    out.push_back(id ? id.value() : std::string("<bad>"));
  }
  return out;
}

Bytes WithDrawingPayload(const Bytes& payload) {
  Bytes buf;
  Append(buf, kBrtPrintOptions, Bytes(2, 0));
  Append(buf, kBrtDrawing, payload);
  return buf;
}

TEST(XlsbDrawingRecord, SpliceOrder) {
  Bytes buf;
  Append(buf, kBrtPrintOptions, Bytes(2, 0));
  Append(buf, kBrtMargins, Bytes(48, 0));
  Append(buf, kBrtPageSetup, Bytes(34, 0));
  Append(buf, kBrtBeginHeaderFooter, Bytes(26, 0));
  Append(buf, kBrtEndHeaderFooter);
  Append(buf, kBrtLegacyDrawing, WideString("rId9"));
  Bytes count;
  emit_u32(count, 1U);
  Append(buf, kBrtBeginTableParts, count);
  Append(buf, kBrtTablePart, WideString("rId8"));
  Append(buf, kBrtEndTableParts);
  Append(buf, kFrtBegin, Bytes{0x02, 0x0E, 0x00, 0x00});
  Append(buf, kFrtEnd);

  ASSERT_TRUE(splice_drawing(buf, "rId7"));
  const std::vector<std::uint16_t> expected = {
      kBrtPrintOptions,  kBrtMargins,         kBrtPageSetup, kBrtBeginHeaderFooter, kBrtEndHeaderFooter, kBrtDrawing,
      kBrtLegacyDrawing, kBrtBeginTableParts, kBrtTablePart, kBrtEndTableParts,     kFrtBegin,           kFrtEnd};
  EXPECT_EQ(Types(buf), expected);
  EXPECT_EQ(DrawingIds(buf), std::vector<std::string>{"rId7"});
}

TEST(XlsbDrawingRecord, SpliceReplacesExisting) {
  Bytes buf;
  Append(buf, kBrtPrintOptions, Bytes(2, 0));
  Append(buf, kBrtDrawing, WideString("rIdOld"));
  Append(buf, kBrtLegacyDrawing, WideString("rId9"));
  Append(buf, kBrtDrawing, WideString("rIdStray"));

  ASSERT_TRUE(splice_drawing(buf, "rId3"));
  EXPECT_EQ(Types(buf), (std::vector<std::uint16_t>{kBrtPrintOptions, kBrtDrawing, kBrtLegacyDrawing}));
  EXPECT_EQ(DrawingIds(buf), std::vector<std::string>{"rId3"});

  ASSERT_TRUE(strip_drawing_records(buf));
  EXPECT_EQ(Types(buf), (std::vector<std::uint16_t>{kBrtPrintOptions, kBrtLegacyDrawing}));
}

TEST(XlsbDrawingRecord, SpliceIntoEmpty) {
  Bytes buf;
  ASSERT_TRUE(splice_drawing(buf, "rId1"));
  EXPECT_EQ(Types(buf), std::vector<std::uint16_t>{kBrtDrawing});
  const DrawingRecordScan scan = scan_drawing_records(buf);
  EXPECT_TRUE(scan.parsed);
  EXPECT_EQ(scan.count, 1U);
  EXPECT_EQ(scan.rel_id, std::optional<std::string>("rId1"));

  // With no record of the earlier slots the drawing leads the slot.
  Bytes tables;
  Append(tables, kBrtLegacyDrawing, WideString("rId2"));
  ASSERT_TRUE(splice_drawing(tables, "rId1"));
  EXPECT_EQ(Types(tables), (std::vector<std::uint16_t>{kBrtDrawing, kBrtLegacyDrawing}));
}

TEST(XlsbDrawingRecord, UnparseableBufferIsLeftAlone) {
  // Declares a five-byte payload and carries one byte.
  const Bytes broken = {0x01, 0x05, 0x00};
  Bytes buf = broken;
  EXPECT_FALSE(splice_drawing(buf, "rId1"));
  EXPECT_EQ(buf, broken);
  EXPECT_FALSE(strip_drawing_records(buf));
  EXPECT_EQ(buf, broken);
  const DrawingRecordScan scan = scan_drawing_records(buf);
  EXPECT_FALSE(scan.parsed);
  EXPECT_EQ(scan.count, 0U);
  EXPECT_FALSE(scan.rel_id.has_value());
}

TEST(XlsbDrawingRecord, ScanReadsOneRelId) {
  // The payload Excel writes for `rId2`.
  const DrawingRecordScan scan = scan_drawing_records(
      WithDrawingPayload(Bytes{0x04, 0x00, 0x00, 0x00, 0x72, 0x00, 0x49, 0x00, 0x64, 0x00, 0x32, 0x00}));
  EXPECT_TRUE(scan.parsed);
  EXPECT_EQ(scan.count, 1U);
  EXPECT_EQ(scan.rel_id, std::optional<std::string>("rId2"));

  // An empty id is still an id; matching it is the reader's business.
  const DrawingRecordScan empty = scan_drawing_records(WithDrawingPayload(WideString("")));
  EXPECT_EQ(empty.rel_id, std::optional<std::string>(""));

  const DrawingRecordScan none = scan_drawing_records(Bytes{});
  EXPECT_TRUE(none.parsed);
  EXPECT_EQ(none.count, 0U);
  EXPECT_FALSE(none.rel_id.has_value());
}

TEST(XlsbDrawingRecord, ScanRejectsTrailingBytes) {
  Bytes payload = WideString("rId2");
  payload.push_back(0x00);
  const DrawingRecordScan scan = scan_drawing_records(WithDrawingPayload(payload));
  EXPECT_TRUE(scan.parsed);
  EXPECT_EQ(scan.count, 1U);
  EXPECT_FALSE(scan.rel_id.has_value());
}

TEST(XlsbDrawingRecord, ScanRejectsNullAndOverlong) {
  const DrawingRecordScan null_id = scan_drawing_records(WithDrawingPayload(Bytes{0xFF, 0xFF, 0xFF, 0xFF}));
  EXPECT_EQ(null_id.count, 1U);
  EXPECT_FALSE(null_id.rel_id.has_value());

  // Claims ten code units and carries two.
  const DrawingRecordScan overlong =
      scan_drawing_records(WithDrawingPayload(Bytes{0x0A, 0x00, 0x00, 0x00, 0x72, 0x00, 0x49, 0x00}));
  EXPECT_EQ(overlong.count, 1U);
  EXPECT_FALSE(overlong.rel_id.has_value());

  const DrawingRecordScan no_length = scan_drawing_records(WithDrawingPayload(Bytes{}));
  EXPECT_EQ(no_length.count, 1U);
  EXPECT_FALSE(no_length.rel_id.has_value());
}

TEST(XlsbDrawingRecord, ScanCountsMany) {
  Bytes buf = WithDrawingPayload(WideString("rId1"));
  Append(buf, kBrtDrawing, WideString("rId2"));
  const DrawingRecordScan scan = scan_drawing_records(buf);
  EXPECT_TRUE(scan.parsed);
  EXPECT_EQ(scan.count, 2U);
  EXPECT_FALSE(scan.rel_id.has_value());
}

constexpr const char* kWsDrOpen =
    "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">";
constexpr const char* kAnchor = "<xdr:oneCellAnchor><xdr:from/><xdr:ext cx=\"1\" cy=\"1\"/>";
constexpr const char* kAnchorClose = "<xdr:clientData/></xdr:oneCellAnchor>";

std::string Drawing(const std::string& body) {
  return std::string(kWsDrOpen) + kAnchor + body + kAnchorClose + "</xdr:wsDr>";
}

std::string Frame(const char* uri) {
  return std::string("<xdr:graphicFrame><a:graphic><a:graphicData uri=\"") + uri +
         "\"/></a:graphic></xdr:graphicFrame>";
}

TEST(XlsbDrawingRecord, FramesWritable) {
  EXPECT_TRUE(drawing_frames_are_writable(Drawing("<xdr:pic><xdr:blipFill><a:blip/></xdr:blipFill></xdr:pic>")));
  EXPECT_TRUE(drawing_frames_are_writable(Drawing("<xdr:sp><xdr:spPr/></xdr:sp>")));
  EXPECT_TRUE(drawing_frames_are_writable(Drawing(Frame("http://schemas.openxmlformats.org/drawingml/2006/chart"))));
  EXPECT_TRUE(drawing_frames_are_writable(std::string(kWsDrOpen) + "</xdr:wsDr>"));

  // A slicer frame sits in the Choice of an AlternateContent.
  const std::string slicer =
      "<mc:AlternateContent><mc:Choice Requires=\"a14\">" +
      Frame("http://schemas.microsoft.com/office/drawing/2010/slicer") +
      "</mc:Choice><mc:Fallback><xdr:sp><xdr:spPr/></xdr:sp></mc:Fallback></mc:AlternateContent>";
  EXPECT_FALSE(drawing_frames_are_writable(Drawing(slicer)));
  const std::string chartex = "<mc:AlternateContent><mc:Choice Requires=\"cx1\">" +
                              Frame("http://schemas.microsoft.com/office/drawing/2014/chartex") +
                              "</mc:Choice></mc:AlternateContent>";
  EXPECT_FALSE(drawing_frames_are_writable(Drawing(chartex)));
  // One foreign frame among charts is enough.
  EXPECT_FALSE(drawing_frames_are_writable(Drawing(Frame("http://schemas.openxmlformats.org/drawingml/2006/chart") +
                                                   Frame("http://schemas.openxmlformats.org/drawingml/2006/table"))));

  EXPECT_FALSE(drawing_frames_are_writable("<xdr:wsDr><xdr:pic>"));
  EXPECT_FALSE(drawing_frames_are_writable(""));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
