// A model-owned drawing in XLSB output: the sheet links it with one BrtDrawing
// record and one drawing relationship, unless its part is missing or holds a
// frame the package cannot carry.

#include <cstddef>
#include <cstdint>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "chart/chart_edit.h"
#include "chart/chart_model.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "drawing/slicer_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "table.h"
#include "unknown_relationship.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kDrawingPath = "xl/drawings/drawing1.xml";
constexpr const char* kSheetRels = "xl/worksheets/_rels/sheet1.bin.rels";

Bytes Png(std::uint32_t w, std::uint32_t h) {
  Bytes b = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  for (std::uint32_t v : {w, h}) {
    for (int shift = 24; shift >= 0; shift -= 8) {
      b.push_back(static_cast<std::uint8_t>(v >> shift));
    }
  }
  b.insert(b.end(), {8, 6, 0, 0, 0, 0x1F, 0x15, 0xC4, 0x89});
  return b;
}

XlsbWriteResult Save(const Workbook& wb) {
  auto result = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value()) : XlsbWriteResult{};
}

/// The bytes of zip entry `path`; empty when absent.
Bytes Entry(const Bytes& package, const std::string& path) {
  ZipReader zip;
  if (!zip.open(SpanOf(package))) {
    return {};
  }
  auto entry = zip.read_entry(path);
  return entry ? std::move(entry.value()) : Bytes{};
}

bool HasEntry(const Bytes& package, const std::string& path) {
  ZipReader zip;
  return static_cast<bool>(zip.open(SpanOf(package))) && static_cast<bool>(zip.read_entry(path));
}

struct Rel {
  std::string id;
  std::string type;
  std::string target;
};

std::vector<Rel> SheetRels(const Bytes& package) {
  const Bytes xml = Entry(package, kSheetRels);
  std::vector<Rel> out;
  pugi::xml_document doc;
  if (xml.empty() || !doc.load_buffer(xml.data(), xml.size())) {
    return out;
  }
  for (const pugi::xml_node rel : doc.child("Relationships").children("Relationship")) {
    out.push_back(Rel{rel.attribute("Id").value(), rel.attribute("Type").value(), rel.attribute("Target").value()});
  }
  return out;
}

std::vector<Rel> DrawingRels(const std::vector<Rel>& rels) {
  std::vector<Rel> out;
  for (const Rel& rel : rels) {
    if (rel.type == kRelDrawing) {
      out.push_back(rel);
    }
  }
  return out;
}

std::vector<FramedRecord> SheetRecords(const Bytes& package, Bytes* body) {
  *body = Entry(package, "xl/worksheets/sheet1.bin");
  return split_records(*body);
}

/// Index of the first record of `type`, or `recs.size()`.
std::size_t IndexOf(const std::vector<FramedRecord>& recs, std::uint16_t type) {
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (recs[i].type == type) {
      return i;
    }
  }
  return recs.size();
}

/// The relationship ids of the sheet's BrtDrawing records.
std::vector<std::string> DrawingRecordIds(const std::vector<FramedRecord>& recs) {
  std::vector<std::string> out;
  for (const FramedRecord& rec : recs) {
    if (rec.type != kBrtDrawing) {
      continue;
    }
    ByteSpan payload = rec.payload;
    auto id = read_xlwidestring(payload);
    out.push_back(id ? id.value() : std::string("<bad>"));
  }
  return out;
}

/// Asserts the sheet links `kDrawingPath` through exactly one record and one
/// relationship naming the same id.
void ExpectLinked(const Bytes& package) {
  Bytes body;
  const std::vector<FramedRecord> recs = SheetRecords(package, &body);
  const std::vector<std::string> ids = DrawingRecordIds(recs);
  ASSERT_EQ(ids.size(), 1U);
  const std::vector<Rel> drawing_rels = DrawingRels(SheetRels(package));
  ASSERT_EQ(drawing_rels.size(), 1U);
  EXPECT_EQ(drawing_rels[0].id, ids[0]);
  EXPECT_EQ(drawing_rels[0].target, "../drawings/drawing1.xml");
  EXPECT_GT(IndexOf(recs, kBrtDrawing), IndexOf(recs, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheetData)));
}

void ExpectNotLinked(const Bytes& package) {
  Bytes body;
  const std::vector<FramedRecord> recs = SheetRecords(package, &body);
  ASSERT_FALSE(recs.empty());
  EXPECT_TRUE(DrawingRecordIds(recs).empty());
  EXPECT_TRUE(DrawingRels(SheetRels(package)).empty());
}

void FillSource(Sheet& sheet) {
  sheet.set_cell_value(0U, 0U, Value::text("Product"));
  sheet.set_cell_value(0U, 1U, Value::text("Amount"));
  sheet.set_cell_value(1U, 0U, Value::text("Apple"));
  sheet.set_cell_value(1U, 1U, Value::number(3.0));
  sheet.set_cell_value(2U, 0U, Value::text("Pear"));
  sheet.set_cell_value(2U, 1U, Value::number(5.0));
}

void AddTable(Workbook& wb) {
  TableMetadata table;
  table.sheet_index = 0;
  table.id = 1;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:B3";
  for (const char* name : {"Product", "Amount"}) {
    TableColumn column;
    column.id = static_cast<std::uint32_t>(table.columns.size() + 1U);
    column.name = name;
    table.columns.push_back(column);
  }
  wb.mutable_tables().push_back(table);
}

void InsertImage(Workbook& wb) {
  const Bytes png = Png(16, 8);
  ImageInsertOptions placement;
  placement.row = 5U;
  placement.col = 3U;
  ASSERT_TRUE(static_cast<bool>(insert_image(wb, 0, png.data(), png.size(), placement)));
}

TEST(XlsbDrawingWrite, NewImageAndChartLinked) {
  Workbook wb = Workbook::create();
  FillSource(wb.sheet(0));
  InsertImage(wb);
  ChartInsertOptions chart;
  chart.kind = ChartInsertKind::kColumn;
  chart.source = "A1:B3";
  ASSERT_TRUE(static_cast<bool>(insert_chart(wb, 0, chart)));
  ASSERT_EQ(wb.sheet(0).drawing_rel_target(), kDrawingPath);
  auto charts = list_chart_parts(wb, 0);
  ASSERT_TRUE(static_cast<bool>(charts));
  ASSERT_EQ(charts.value().size(), 1U);

  const XlsbWriteResult saved = Save(wb);
  EXPECT_TRUE(HasEntry(saved.bytes, kDrawingPath));
  EXPECT_TRUE(HasEntry(saved.bytes, charts.value()[0].part_path));
  ExpectLinked(saved.bytes);
}

TEST(XlsbDrawingWrite, RelIdsStayUnique) {
  Workbook wb = Workbook::create();
  FillSource(wb.sheet(0));
  AddTable(wb);
  Hyperlink link;
  link.row = link.last_row = 0U;
  link.col = link.last_col = 4U;
  link.target = "https://example.com/";
  wb.sheet(0).mutable_hyperlinks().push_back(link);
  InsertImage(wb);

  const XlsbWriteResult saved = Save(wb);
  const std::vector<Rel> rels = SheetRels(saved.bytes);
  std::set<std::string> ids;
  bool has_table = false;
  bool has_hyperlink = false;
  for (const Rel& rel : rels) {
    ids.insert(rel.id);
    has_table = has_table || rel.type == kRelTable;
    has_hyperlink = has_hyperlink || rel.type == kRelHyperlink;
  }
  EXPECT_EQ(ids.size(), rels.size());
  EXPECT_TRUE(has_table);
  EXPECT_TRUE(has_hyperlink);
  ExpectLinked(saved.bytes);
  // The drawing precedes the table parts.
  Bytes body;
  const std::vector<FramedRecord> recs = SheetRecords(saved.bytes, &body);
  EXPECT_LT(IndexOf(recs, kBrtDrawing), IndexOf(recs, kBrtBeginTableParts));
}

TEST(XlsbDrawingWrite, SlicerFrameIsNotLinked) {
  Workbook wb = Workbook::create();
  FillSource(wb.sheet(0));
  AddTable(wb);
  SlicerSpec spec;
  spec.fields = slicer_field::kAnchor;
  spec.from = AnchorPoint{4U, 0U, 0, 0};
  spec.to = AnchorPoint{12U, 2U, 0, 0};
  ASSERT_TRUE(static_cast<bool>(add_slicer(wb, 0, 0, 0, spec)));
  ASSERT_FALSE(wb.sheet(0).drawing_rel_target().empty());
  ExpectNotLinked(Save(wb).bytes);

  // The drawing alone counts as one deferred feature.
  constexpr const char* kSlicerDrawing =
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
      "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><xdr:oneCellAnchor><xdr:from/>"
      "<xdr:ext cx=\"1\" cy=\"1\"/><mc:AlternateContent><mc:Choice Requires=\"a14\"><xdr:graphicFrame><a:graphic>"
      "<a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\"/></a:graphic>"
      "</xdr:graphicFrame></mc:Choice></mc:AlternateContent><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>";
  const std::string xml(kSlicerDrawing);
  Workbook framed = Workbook::create();
  framed.set_passthrough_parts({PassthroughPart(kDrawingPath, std::string(kCtDrawing), Bytes(xml.begin(), xml.end()))});
  framed.sheet(0).set_drawing_rel_target(kDrawingPath);
  Workbook plain = Workbook::create();
  plain.set_passthrough_parts({PassthroughPart(kDrawingPath, std::string(kCtDrawing), Bytes(xml.begin(), xml.end()))});
  const XlsbWriteResult with = Save(framed);
  const XlsbWriteResult without = Save(plain);
  ExpectNotLinked(with.bytes);
  EXPECT_EQ(with.diagnostics.deferred_feature_count, without.diagnostics.deferred_feature_count + 1U);
  EXPECT_EQ(with.diagnostics.dropped_relationship_count, without.diagnostics.dropped_relationship_count);
}

TEST(XlsbDrawingWrite, MissingPartDropsLinkAndRetainedRecord) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0U, 0U, Value::number(1.0));
  sheet.set_drawing_rel_target(kDrawingPath);
  sheet.set_unknown_relationships({UnknownRelationship{"rId1", std::string(kRelDrawing), kDrawingPath, false}});
  XlsbSheetTail tail;
  Bytes payload;
  emit_xlwidestring(payload, "rId1");
  emit_record(tail.after_hyperlinks, kBrtDrawing, payload);
  sheet.set_xlsb_tail(std::move(tail));
  Workbook plain = Workbook::create();
  plain.sheet(0).set_cell_value(0U, 0U, Value::number(1.0));

  const XlsbWriteResult with = Save(wb);
  const XlsbWriteResult without = Save(plain);
  ExpectNotLinked(with.bytes);
  EXPECT_TRUE(SheetRels(with.bytes).empty());
  EXPECT_EQ(with.diagnostics.dropped_relationship_count, without.diagnostics.dropped_relationship_count + 1U);
  EXPECT_EQ(with.diagnostics.deferred_feature_count, without.diagnostics.deferred_feature_count);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
