// XLSB drawings through read_xlsb: a sheet linked to its drawing by one
// BrtDrawing record and one drawing relationship is lifted onto the model, so
// it lists, edits and saves as an OOXML drawing does; any other shape stays
// retained and refuses edits.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "chart/chart_edit.h"
#include "chart/chart_model.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "gtest/gtest.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/drawing_record.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "reader_test_helpers.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;
using reader_test_support::BuildZip;
using reader_test_support::PartFile;
using reader_test_support::SpanOf;

constexpr const char* kDrawingPath = "xl/drawings/drawing1.xml";
constexpr const char* kSheetPart = "xl/worksheets/sheet1.bin";
constexpr const char* kSheetRels = "xl/worksheets/_rels/sheet1.bin.rels";
constexpr const char* kContentTypes = "[Content_Types].xml";

Bytes ReadFileBytes(const std::string& path) {
  Bytes out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

Bytes FixtureBytes() {
  return ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/chart_xlsb.xlsb");
}

Workbook Read(const Bytes& package) {
  auto read = read_xlsb(SpanOf(package));
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

Bytes Write(const Workbook& wb) {
  auto written = write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(written)) << (written ? "" : written.error().message);
  return written ? std::move(written.value()) : Bytes{};
}

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

Expected<std::uint32_t, Error> InsertImage(Workbook& wb, std::size_t sheet) {
  const Bytes png = Png(16, 8);
  ImageInsertOptions placement;
  placement.row = 5U;
  placement.col = 3U;
  return insert_image(wb, sheet, png.data(), png.size(), placement);
}

ChartInsertOptions ColumnChart(const char* source) {
  ChartInsertOptions chart;
  chart.kind = ChartInsertKind::kColumn;
  chart.source = source;
  return chart;
}

template <typename T>
T Must(Expected<T, Error> got) {
  EXPECT_TRUE(static_cast<bool>(got)) << (got ? "" : got.error().message);
  return got ? std::move(got.value()) : T{};
}

void MustOk(const Expected<void, Error>& got) {
  EXPECT_TRUE(static_cast<bool>(got)) << (got ? "" : got.error().message);
}

bool HasDrawingRel(const Sheet& sheet) {
  for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
    if (rel.type == kRelDrawing) {
      return true;
    }
  }
  return false;
}

/// The sheet holds its drawing on the model, with no retained record or relationship left.
void ExpectLifted(const Sheet& sheet, const std::string& target) {
  EXPECT_EQ(sheet.drawing_rel_target(), target);
  EXPECT_FALSE(HasDrawingRel(sheet));
  EXPECT_EQ(scan_drawing_records(sheet.xlsb_tail().after_hyperlinks).count, 0U);
}

/// Zip entries in archive order.
std::vector<PartFile> Entries(const Bytes& package) {
  std::vector<PartFile> out;
  ZipReader zip;
  if (!zip.open(SpanOf(package))) {
    ADD_FAILURE() << "package does not open";
    return out;
  }
  for (const std::string& name : zip.list_entries()) {
    auto bytes = zip.read_entry(name);
    out.push_back(PartFile{name, bytes ? std::move(bytes.value()) : Bytes{}});
  }
  return out;
}

Bytes Entry(const Bytes& package, const std::string& path) {
  for (PartFile& part : Entries(package)) {
    if (part.path == path) {
      return std::move(part.body);
    }
  }
  return {};
}

struct StringWriter : pugi::xml_writer {
  std::string text;
  void write(const void* data, std::size_t size) override { text.append(static_cast<const char*>(data), size); }
};

/// `xml` after `edit` runs on its parsed document.
template <typename Edit>
Bytes EditXml(const Bytes& xml, Edit edit) {
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(xml.data(), xml.size()));
  edit(doc);
  StringWriter out;
  doc.save(out, "", pugi::format_raw);
  return Bytes(out.text.begin(), out.text.end());
}

/// `package` with entry `path` replaced by `edit(bytes)`, or dropped when `edit` is null.
template <typename Edit>
Bytes Repack(const Bytes& package, const std::string& path, Edit edit) {
  std::vector<PartFile> parts;
  bool found = false;
  for (PartFile& part : Entries(package)) {
    if (part.path == path) {
      found = true;
      part.body = edit(part.body);
    }
    parts.push_back(std::move(part));
  }
  EXPECT_TRUE(found) << path;
  return BuildZip(parts);
}

Bytes Drop(const Bytes& package, const std::string& path) {
  std::vector<PartFile> parts;
  for (PartFile& part : Entries(package)) {
    if (part.path != path) {
      parts.push_back(std::move(part));
    }
  }
  return BuildZip(parts);
}

/// The relationship ids of the sheet part's BrtDrawing records.
std::vector<std::string> DrawingRecordIds(const Bytes& package) {
  const Bytes body = Entry(package, kSheetPart);
  std::vector<std::string> out;
  for (const FramedRecord& rec : split_records(body)) {
    if (rec.type != kBrtDrawing) {
      continue;
    }
    ByteSpan payload = rec.payload;
    auto id = read_xlwidestring(payload);
    out.push_back(id ? id.value() : std::string("<bad>"));
  }
  return out;
}

/// The (id, target) of every drawing relationship in the sheet rels.
std::vector<std::pair<std::string, std::string>> DrawingRels(const Bytes& package) {
  const Bytes xml = Entry(package, kSheetRels);
  std::vector<std::pair<std::string, std::string>> out;
  pugi::xml_document doc;
  if (xml.empty() || !doc.load_buffer(xml.data(), xml.size())) {
    return out;
  }
  for (const pugi::xml_node rel : doc.child("Relationships").children("Relationship")) {
    if (std::string(rel.attribute("Type").value()) == kRelDrawing) {
      out.emplace_back(rel.attribute("Id").value(), rel.attribute("Target").value());
    }
  }
  return out;
}

void ExpectSameAnchor(const AnchorPoint& got, const AnchorPoint& want) {
  EXPECT_EQ(got.row, want.row);
  EXPECT_EQ(got.col, want.col);
  EXPECT_EQ(got.row_off, want.row_off);
  EXPECT_EQ(got.col_off, want.col_off);
}

void FillSource(Sheet& sheet) {
  sheet.set_cell_value(0U, 0U, Value::text("Product"));
  sheet.set_cell_value(0U, 1U, Value::text("Amount"));
  sheet.set_cell_value(1U, 0U, Value::text("Apple"));
  sheet.set_cell_value(1U, 1U, Value::number(3.0));
  sheet.set_cell_value(2U, 0U, Value::text("Pear"));
  sheet.set_cell_value(2U, 1U, Value::number(5.0));
}

TEST(XlsbDrawingRoundTrip, NewImageAndChartSurvive) {
  Workbook wb = Workbook::create();
  FillSource(wb.sheet(0));
  Must(InsertImage(wb, 0));
  Must(insert_chart(wb, 0, ColumnChart("A1:B3")));
  const std::vector<DrawingObject> objects = Must(list_drawing_objects(wb, 0));
  const std::vector<ChartPartRef> charts = Must(list_chart_parts(wb, 0));
  ASSERT_EQ(objects.size(), 2U);
  ASSERT_EQ(charts.size(), 1U);

  const Workbook back = Read(Write(wb));
  ExpectLifted(back.sheet(0), kDrawingPath);
  const std::vector<DrawingObject> got = Must(list_drawing_objects(back, 0));
  ASSERT_EQ(got.size(), objects.size());
  for (std::size_t i = 0; i < got.size(); ++i) {
    EXPECT_EQ(got[i].object_id, objects[i].object_id);
    EXPECT_EQ(got[i].kind, objects[i].kind);
    EXPECT_EQ(got[i].anchor_kind, objects[i].anchor_kind);
    ExpectSameAnchor(got[i].from, objects[i].from);
    ExpectSameAnchor(got[i].to, objects[i].to);
    EXPECT_EQ(got[i].cx, objects[i].cx);
    EXPECT_EQ(got[i].cy, objects[i].cy);
  }
  const std::vector<ChartPartRef> got_charts = Must(list_chart_parts(back, 0));
  ASSERT_EQ(got_charts.size(), 1U);
  EXPECT_EQ(got_charts[0].object_id, charts[0].object_id);
  EXPECT_EQ(got_charts[0].part_path, charts[0].part_path);
}

TEST(XlsbDrawingRoundTrip, ExcelFixtureIsLiftedAndEditable) {
  Workbook wb = Read(FixtureBytes());
  ExpectLifted(wb.sheet(0), kDrawingPath);
  const std::vector<ChartPartRef> charts = Must(list_chart_parts(wb, 0));
  ASSERT_EQ(charts.size(), 4U);
  EXPECT_EQ(Must(list_drawing_objects(wb, 0)).size(), 4U);

  const std::uint32_t image = Must(InsertImage(wb, 0));
  EXPECT_EQ(Must(list_drawing_objects(wb, 0)).size(), 5U);
  MustOk(remove_drawing(wb, 0, image));
  Must(insert_chart(wb, 0, ColumnChart("Sheet1!$A$1:$C$5")));
  ChartTitle title;
  title.kind = ChartTitleKind::kText;
  title.text = "Lifted";
  MustOk(set_chart_title(wb, 0, charts[0].object_id, title));
  EXPECT_EQ(Must(get_chart(wb, 0, charts[0].object_id)).title.text, "Lifted");
  EXPECT_EQ(wb.sheet(0).drawing_rel_target(), kDrawingPath);
}

TEST(XlsbDrawingRoundTrip, UneditedFixtureKeepsTheLink) {
  const Bytes saved = Write(Read(FixtureBytes()));
  const std::vector<std::string> ids = DrawingRecordIds(saved);
  ASSERT_EQ(ids.size(), 1U);
  const auto rels = DrawingRels(saved);
  ASSERT_EQ(rels.size(), 1U);
  EXPECT_EQ(rels[0].first, ids[0]);
  EXPECT_EQ(rels[0].second, "../drawings/drawing1.xml");

  const Workbook back = Read(saved);
  ExpectLifted(back.sheet(0), kDrawingPath);
  EXPECT_EQ(Must(list_chart_parts(back, 0)).size(), 4U);
}

TEST(XlsbDrawingRoundTrip, XlsbToXlsxKeepsCharts) {
  const Workbook wb = Read(FixtureBytes());
  auto xlsx = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const Bytes sheet_xml = Entry(xlsx.value(), "xl/worksheets/sheet1.xml");
  EXPECT_NE(std::string(sheet_xml.begin(), sheet_xml.end()).find("<drawing r:id="), std::string::npos);

  auto back = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
  EXPECT_FALSE(back.value().workbook.sheet(0).drawing_rel_target().empty());
  EXPECT_EQ(Must(list_chart_parts(back.value().workbook, 0)).size(), 4U);
}

TEST(XlsbDrawingRoundTrip, CopySheetCopiesLiftedChart) {
  Workbook wb = Read(FixtureBytes());
  const std::size_t copy = Must(wb.copy_sheet(0, 1, std::nullopt));
  ASSERT_EQ(copy, 1U);
  const std::string& source_drawing = wb.sheet(0).drawing_rel_target();
  const std::string& copy_drawing = wb.sheet(copy).drawing_rel_target();
  EXPECT_FALSE(copy_drawing.empty());
  EXPECT_NE(copy_drawing, source_drawing);
  const std::vector<ChartPartRef> source_charts = Must(list_chart_parts(wb, 0));
  const std::vector<ChartPartRef> copy_charts = Must(list_chart_parts(wb, copy));
  ASSERT_EQ(copy_charts.size(), source_charts.size());
  for (std::size_t i = 0; i < copy_charts.size(); ++i) {
    EXPECT_NE(copy_charts[i].part_path, source_charts[i].part_path);
  }
}

TEST(XlsbDrawingRoundTrip, InsertRowsMovesLiftedAnchor) {
  Workbook wb = Read(FixtureBytes());
  const std::vector<DrawingObject> before = Must(list_drawing_objects(wb, 0));
  ASSERT_FALSE(before.empty());
  MustOk(wb.insert_rows(0, 0, 2));
  const std::vector<DrawingObject> after = Must(list_drawing_objects(wb, 0));
  ASSERT_EQ(after.size(), before.size());
  for (std::size_t i = 0; i < after.size(); ++i) {
    ASSERT_EQ(before[i].anchor_kind, AnchorKind::kTwoCell);
    EXPECT_EQ(after[i].from.row, before[i].from.row + 2U);
    EXPECT_EQ(after[i].to.row, before[i].to.row + 2U);
  }
}

TEST(XlsbDrawingRoundTrip, UnliftableDrawingStaysRetained) {
  Workbook source = Workbook::create();
  FillSource(source.sheet(0));
  Must(InsertImage(source, 0));
  const Bytes base = Write(source);
  ExpectLifted(Read(base).sheet(0), kDrawingPath);
  const auto rels = DrawingRels(base);
  ASSERT_EQ(rels.size(), 1U);
  const std::string rid = rels[0].first;

  auto add_rel = [](pugi::xml_document& doc, const std::string& id) {
    pugi::xml_node rel = doc.child("Relationships").append_child("Relationship");
    rel.append_attribute("Id") = id.c_str();
    rel.append_attribute("Type") = std::string(kRelVmlDrawing).c_str();
    rel.append_attribute("Target") = "../drawings/vmlDrawing1.vml";
  };
  auto drawing_rel = [](pugi::xml_document& doc) {
    for (pugi::xml_node rel : doc.child("Relationships").children("Relationship")) {
      if (std::string(rel.attribute("Type").value()) == kRelDrawing) {
        return rel;
      }
    }
    return pugi::xml_node();
  };

  struct Variant {
    const char* name;
    Bytes package;
  };
  std::vector<Variant> variants;
  variants.push_back({"two records", Repack(base, kSheetPart, [](const Bytes& body) {
                        Bytes out;
                        for (const FramedRecord& rec : split_records(body)) {
                          append_record(out, body, rec);
                          if (rec.type == kBrtDrawing) {
                            append_record(out, body, rec);
                          }
                        }
                        return out;
                      })});
  variants.push_back({"wrong type", Repack(base, kSheetRels, [&](const Bytes& xml) {
                        return EditXml(xml, [&](pugi::xml_document& doc) {
                          drawing_rel(doc).attribute("Id") = "rIdMoved";
                          add_rel(doc, rid);
                        });
                      })});
  variants.push_back({"duplicate id", Repack(base, kSheetRels, [&](const Bytes& xml) {
                        return EditXml(xml, [&](pugi::xml_document& doc) { add_rel(doc, rid); });
                      })});
  variants.push_back({"missing part", Drop(base, kDrawingPath)});
  variants.push_back({"default typed", Repack(base, kContentTypes, [](const Bytes& xml) {
                        return EditXml(xml, [](pugi::xml_document& doc) {
                          pugi::xml_node types = doc.child("Types");
                          pugi::xml_node drawing = types.find_child_by_attribute(
                              "Override", "PartName", (std::string("/") + kDrawingPath).c_str());
                          EXPECT_TRUE(static_cast<bool>(drawing));
                          types.remove_child(drawing);
                        });
                      })});

  for (const Variant& variant : variants) {
    SCOPED_TRACE(variant.name);
    Workbook wb = Read(variant.package);
    const Sheet& sheet = wb.sheet(0);
    EXPECT_TRUE(sheet.drawing_rel_target().empty());
    EXPECT_TRUE(HasDrawingRel(sheet));
    auto refused = InsertImage(wb, 0);
    ASSERT_FALSE(static_cast<bool>(refused));
    EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoDrawingUnparseable) << refused.error().message;
  }
}

TEST(XlsbDrawingRoundTrip, LiftedDrawingAcceptsEdits) {
  Workbook wb = Read(FixtureBytes());
  const std::vector<ChartPartRef> charts = Must(list_chart_parts(wb, 0));
  ASSERT_EQ(charts.size(), 4U);
  const std::uint32_t image = Must(InsertImage(wb, 0));
  MustOk(remove_drawing(wb, 0, charts[0].object_id));
  const std::uint32_t added = Must(insert_chart(wb, 0, ColumnChart("Sheet1!$A$1:$C$5")));
  ChartTitle title;
  title.kind = ChartTitleKind::kText;
  title.text = "Edited";
  MustOk(set_chart_title(wb, 0, charts[1].object_id, title));
  const std::vector<DrawingObject> objects = Must(list_drawing_objects(wb, 0));
  ASSERT_EQ(objects.size(), 5U);

  const Bytes saved = Write(wb);
  ASSERT_EQ(DrawingRecordIds(saved).size(), 1U);
  const Workbook back = Read(saved);
  ExpectLifted(back.sheet(0), wb.sheet(0).drawing_rel_target());
  const std::vector<DrawingObject> got = Must(list_drawing_objects(back, 0));
  ASSERT_EQ(got.size(), objects.size());
  for (std::size_t i = 0; i < got.size(); ++i) {
    EXPECT_EQ(got[i].object_id, objects[i].object_id);
    EXPECT_EQ(got[i].kind, objects[i].kind);
  }
  std::vector<std::uint32_t> chart_ids;
  for (const ChartPartRef& ref : Must(list_chart_parts(back, 0))) {
    chart_ids.push_back(ref.object_id);
  }
  ASSERT_EQ(chart_ids.size(), 4U);
  EXPECT_EQ(std::count(chart_ids.begin(), chart_ids.end(), charts[0].object_id), 0);
  EXPECT_EQ(std::count(chart_ids.begin(), chart_ids.end(), added), 1);
  EXPECT_EQ(Must(get_chart(back, 0, charts[1].object_id)).title.text, "Edited");
  EXPECT_EQ(std::count_if(got.begin(), got.end(), [image](const DrawingObject& o) { return o.object_id == image; }), 1);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
