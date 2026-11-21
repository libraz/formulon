// Rich values the engine does not model survive an .xlsx or .xlsb load and
// save: the in-cell pictures Excel wrote, an overwritten picture dropping its
// value and its now-unreferenced media, a book the table cannot carry keeping
// its parts verbatim with a save diagnostic, an error-only book writing the
// same rich-value parts as the error table alone, and pictures crossing
// between the two formats.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/dynamic_array_formula.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/rich_value_errors.h"
#include "io/rich_values.h"
#include "io/xlsb/emission_plan.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/pivot_part_plan.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "passthrough_part.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "utils/a1_ref.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

constexpr std::string_view kMetadataPart = "xl/metadata.xml";
constexpr std::string_view kValuesPart = "xl/richData/rdrichvalue.xml";
constexpr std::string_view kStructuresPart = "xl/richData/rdrichvaluestructure.xml";
constexpr std::string_view kTypesPart = "xl/richData/rdRichValueTypes.xml";
constexpr std::string_view kRelPart = "xl/richData/richValueRel.xml";
constexpr std::string_view kRelRelsPart = "xl/richData/_rels/richValueRel.xml.rels";
constexpr std::string_view kImageRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

std::string FixturePath(std::string_view relative) {
  return std::string(FORMULON_FIXTURES_DIR) + "/" + std::string(relative);
}

std::vector<std::uint8_t> ReadFixture(std::string_view relative) {
  std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath(relative));
  EXPECT_FALSE(bytes.empty()) << relative;
  return bytes;
}

std::vector<std::uint8_t> ReadPicture(std::string_view name) {
  return ReadFixture("excel/cell_pictures/" + std::string(name) + ".xlsx");
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

std::vector<std::string> PartNames(const std::vector<std::uint8_t>& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  return zip.list_entries();
}

std::vector<std::string> MediaParts(const std::vector<std::uint8_t>& package) {
  std::vector<std::string> out;
  for (const std::string& name : PartNames(package)) {
    if (name.compare(0, 9, "xl/media/") == 0) {
      out.push_back(name);
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

void ExpectSamePart(const std::vector<std::uint8_t>& expected, const std::vector<std::uint8_t>& actual,
                    std::string_view part) {
  SCOPED_TRACE(std::string(part));
  const std::optional<std::string> want = Part(expected, part);
  const std::optional<std::string> got = Part(actual, part);
  ASSERT_EQ(got.has_value(), want.has_value());
  if (want.has_value()) {
    EXPECT_EQ(*got, *want);
  }
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto loaded = read_ooxml(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

OoxmlWriteResult Save(const Workbook& wb) {
  auto saved = write_ooxml_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : OoxmlWriteResult{};
}

/// `vm` of every `<c>` in worksheet part `sheet_part` that carries one, by A1 reference.
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

/// The `t` of every `<c>` in `sheet_part`, by A1 reference.
std::map<std::string, std::string> CellTypes(const std::vector<std::uint8_t>& package, std::string_view sheet_part) {
  std::map<std::string, std::string> out;
  const std::string xml = Part(package, sheet_part).value_or("");
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(xml.data(), xml.size())) << sheet_part;
  for (pugi::xml_node row : doc.child("worksheet").child("sheetData").children("row")) {
    for (pugi::xml_node c : row.children("c")) {
      out[c.attribute("r").value()] = c.attribute("t").value();
    }
  }
  return out;
}

/// (Id, Type, Target) of each relationship in the rels part `part`.
std::vector<std::vector<std::string>> Relationships(const std::vector<std::uint8_t>& package, std::string_view part) {
  std::vector<std::vector<std::string>> out;
  const std::string xml = Part(package, part).value_or("");
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(xml.data(), xml.size())) << part;
  for (pugi::xml_node r : doc.child("Relationships").children("Relationship")) {
    out.push_back({r.attribute("Id").value(), r.attribute("Type").value(), r.attribute("Target").value()});
  }
  std::sort(out.begin(), out.end());
  return out;
}

/// `package` with each part of `replaced` swapped for its new bytes.
std::vector<std::uint8_t> Repack(const std::vector<std::uint8_t>& package,
                                 const std::map<std::string, std::string>& replaced) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const std::string& name : zip.list_entries()) {
    const auto it = replaced.find(name);
    std::vector<std::uint8_t> bytes;
    if (it != replaced.end()) {
      bytes.assign(it->second.begin(), it->second.end());
    } else {
      auto read = zip.read_entry(name);
      EXPECT_TRUE(static_cast<bool>(read)) << name;
      if (read) {
        bytes = std::move(read.value());
      }
    }
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name.c_str(), bytes.data(), bytes.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << name;
  }
  void* archive = nullptr;
  std::size_t size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &size), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + size);
  mz_free(archive);
  mz_zip_writer_end(&writer);
  return out;
}

std::uint32_t Row(const char* a1) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  EXPECT_TRUE(a1::parse_a1_ref(a1, &row, &col)) << a1;
  return row;
}

std::uint32_t Col(const char* a1) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  EXPECT_TRUE(a1::parse_a1_ref(a1, &row, &col)) << a1;
  return col;
}

/// True when sheet 0 of `wb` caches a picture at `a1`.
bool HoldsPicture(const Workbook& wb, const char* a1) {
  const Cell* cell = wb.sheet(0).cell_at(Row(a1), Col(a1));
  return cell != nullptr && cell->cached_value.is_picture();
}

/// The picture-table entry of the picture sheet 0 of `wb` caches at `a1`, or null.
const PictureEntry* PictureAt(const Workbook& wb, const char* a1) {
  const Cell* cell = wb.sheet(0).cell_at(Row(a1), Col(a1));
  return cell != nullptr && cell->cached_value.is_picture()
             ? wb.picture_entry(cell->cached_value.as_picture().picture_id)
             : nullptr;
}

// ---------------------------------------------------------------------------
// In-cell pictures
// ---------------------------------------------------------------------------

TEST(RichValuePreserve, CellPicturesReadAsPictureValues) {
  const Workbook wb = Load(ReadPicture("cell_pictures"));
  const Sheet& sheet = wb.sheet(0);
  for (const char* cell : {"E2", "E4", "E6", "E8"}) {
    EXPECT_TRUE(HoldsPicture(wb, cell)) << cell;
    EXPECT_FALSE(sheet.rich_value_at(Row(cell), Col(cell)).has_value()) << cell;
  }
  EXPECT_FALSE(sheet.has_rich_values());
  const Cell* e2 = sheet.cell_at(Row("E2"), Col("E2"));
  ASSERT_NE(e2, nullptr);
  EXPECT_EQ(e2->cached_value.picture_origin(), 5U);
  const PictureEntry* e2_entry = PictureAt(wb, "E2");
  ASSERT_NE(e2_entry, nullptr);
  EXPECT_EQ(e2_entry->media_path, "xl/media/image1.png");
  EXPECT_FALSE(e2_entry->alt.has_value());
  const PictureEntry* e8 = PictureAt(wb, "E8");
  ASSERT_NE(e8, nullptr);
  EXPECT_EQ(e8->media_path, "xl/media/image2.png");
  EXPECT_EQ(e8->alt, std::optional<std::string>("Blue circle on a blue field"));
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    EXPECT_NE(part.path, kMetadataPart);
    EXPECT_NE(part.path, kValuesPart);
  }
}

TEST(RichValuePreserve, CellPicturesRoundTrip) {
  const std::vector<std::uint8_t> excel = ReadPicture("cell_pictures");
  const OoxmlWriteResult saved = Save(Load(excel));
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(saved.diagnostics.deferred_feature_count, 0U);

  const std::map<std::string, std::string> vm = VmCells(saved.bytes, "xl/worksheets/sheet1.xml");
  EXPECT_EQ(vm, (std::map<std::string, std::string>{{"E2", "1"}, {"E4", "1"}, {"E6", "2"}, {"E8", "3"}}));

  // The table numbers values, structures and richValueRel entries as Excel did.
  for (const std::string_view part : {kMetadataPart, kValuesPart, kStructuresPart, kTypesPart, kRelPart}) {
    ExpectSamePart(excel, saved.bytes, part);
  }
  EXPECT_EQ(Relationships(saved.bytes, kRelRelsPart),
            (std::vector<std::vector<std::string>>{{"rId1", std::string(kImageRel), "../media/image1.png"},
                                                   {"rId2", std::string(kImageRel), "../media/image2.png"}}));
  EXPECT_EQ(MediaParts(saved.bytes), (std::vector<std::string>{"xl/media/image1.png", "xl/media/image2.png"}));
  ExpectSamePart(excel, saved.bytes, "xl/media/image1.png");
  ExpectSamePart(excel, saved.bytes, "xl/media/image2.png");

  const std::string content_types = Part(saved.bytes, "[Content_Types].xml").value_or("");
  for (const char* override_part :
       {"<Override PartName=\"/xl/richData/richValueRel.xml\" "
        "ContentType=\"application/vnd.ms-excel.richvaluerel+xml\"/>",
        "<Override PartName=\"/xl/richData/rdrichvalue.xml\" "
        "ContentType=\"application/vnd.ms-excel.rdrichvalue+xml\"/>",
        "<Override PartName=\"/xl/richData/rdrichvaluestructure.xml\" "
        "ContentType=\"application/vnd.ms-excel.rdrichvaluestructure+xml\"/>",
        "<Override PartName=\"/xl/richData/rdRichValueTypes.xml\" "
        "ContentType=\"application/vnd.ms-excel.rdrichvaluetypes+xml\"/>",
        "<Override PartName=\"/xl/metadata.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml\"/>",
        "<Default Extension=\"png\" ContentType=\"image/png\"/>"}) {
    EXPECT_NE(content_types.find(override_part), std::string::npos) << override_part << "\n" << content_types;
  }
  bool rel_registered = false;
  for (const std::vector<std::string>& rel : Relationships(saved.bytes, "xl/_rels/workbook.xml.rels")) {
    rel_registered = rel_registered || (rel[1] == kRelRichValueRel && rel[2] == "richData/richValueRel.xml");
  }
  EXPECT_TRUE(rel_registered);

  // A second cycle writes the same parts again.
  const OoxmlWriteResult resaved = Save(Load(saved.bytes));
  for (const std::string_view part :
       {kMetadataPart, kValuesPart, kStructuresPart, kTypesPart, kRelPart, kRelRelsPart}) {
    ExpectSamePart(saved.bytes, resaved.bytes, part);
  }
  EXPECT_EQ(VmCells(resaved.bytes, "xl/worksheets/sheet1.xml"), vm);
}

TEST(RichValuePreserve, OverwritingAPictureDropsItsValueAndOrphanMedia) {
  const std::vector<std::uint8_t> excel = ReadPicture("cell_pictures");
  const WriteDiagnostics unedited = Save(Load(excel)).diagnostics;
  Workbook wb = Load(excel);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E2"), Col("E2"), Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E4"), Col("E4"), Value::number(2.0))));
  EXPECT_FALSE(HoldsPicture(wb, "E2"));
  EXPECT_FALSE(HoldsPicture(wb, "E4"));
  EXPECT_TRUE(HoldsPicture(wb, "E6"));

  const OoxmlWriteResult saved = Save(wb);
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(VmCells(saved.bytes, "xl/worksheets/sheet1.xml"),
            (std::map<std::string, std::string>{{"E6", "1"}, {"E8", "2"}}));
  const std::string values = Part(saved.bytes, kValuesPart).value_or("");
  EXPECT_NE(values.find("count=\"2\"><rv s=\"0\"><v>0</v><v>5</v></rv>"
                        "<rv s=\"1\"><v>0</v><v>5</v><v>Blue circle on a blue field</v></rv></rvData>"),
            std::string::npos)
      << values;
  EXPECT_NE(Part(saved.bytes, kRelPart).value_or("").find("<rel r:id=\"rId1\"/></richValueRels>"), std::string::npos);
  EXPECT_EQ(Relationships(saved.bytes, kRelRelsPart),
            (std::vector<std::vector<std::string>>{{"rId1", std::string(kImageRel), "../media/image2.png"}}));
  // image1.png was named only by the dropped value.
  EXPECT_EQ(MediaParts(saved.bytes), (std::vector<std::string>{"xl/media/image2.png"}));
  // Dropping the orphan is not a loss.
  EXPECT_EQ(saved.diagnostics.dropped_part_count, unedited.dropped_part_count);
  EXPECT_EQ(saved.diagnostics.dropped_relationship_count, unedited.dropped_relationship_count);
}

TEST(RichValuePreserve, ClearingEveryPictureDropsTheRichParts) {
  const std::vector<std::uint8_t> excel = ReadPicture("cell_pictures");
  const WriteDiagnostics control = Save(Load(excel)).diagnostics;
  Workbook wb = Load(excel);
  for (const char* cell : {"E2", "E4", "E6", "E8"}) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row(cell), Col(cell), Value::number(1.0)))) << cell;
  }
  const OoxmlWriteResult saved = Save(wb);
  ASSERT_FALSE(saved.bytes.empty());
  for (const std::string& name : PartNames(saved.bytes)) {
    EXPECT_NE(name.compare(0, 12, "xl/richData/"), 0) << name;
  }
  EXPECT_FALSE(Part(saved.bytes, kMetadataPart).has_value());
  EXPECT_TRUE(MediaParts(saved.bytes).empty());
  EXPECT_TRUE(VmCells(saved.bytes, "xl/worksheets/sheet1.xml").empty());
  for (const std::vector<std::string>& rel : Relationships(saved.bytes, "xl/_rels/workbook.xml.rels")) {
    EXPECT_NE(rel[1], kRelRichValueRel);
  }
  // The richValueRel link goes with its part on purpose; neither is a loss.
  EXPECT_EQ(saved.diagnostics.dropped_relationship_count, control.dropped_relationship_count);
  EXPECT_EQ(saved.diagnostics.dropped_part_count, control.dropped_part_count);
}

TEST(RichValuePreserve, AFormulaPictureLeavesWhenItsResultChanges) {
  // G2 is `=E2`, saved with E2's picture as its value (CalcOrigin 4).
  Workbook wb = Load(ReadPicture("picture_formula_refs"));
  ASSERT_TRUE(HoldsPicture(wb, "G2"));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E2"), Col("E2"), Value::number(7.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_FALSE(HoldsPicture(wb, "G2"));
  const OoxmlWriteResult saved = Save(wb);
  ASSERT_FALSE(saved.bytes.empty());
  const std::map<std::string, std::string> vm = VmCells(saved.bytes, "xl/worksheets/sheet1.xml");
  EXPECT_EQ(vm.count("E2"), 0U);
  EXPECT_EQ(vm.count("G2"), 0U);
  EXPECT_EQ(vm.count("E4"), 1U);
}

TEST(RichValuePreserve, ColumnEditsMoveAndDropPictures) {
  // Pictures in column E: an insert at D moves them to F, deleting F drops them.
  Workbook wb = Load(ReadPicture("cell_pictures"));
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, Col("D1"), 1U)));
  EXPECT_FALSE(HoldsPicture(wb, "E2"));
  EXPECT_TRUE(HoldsPicture(wb, "F2"));
  EXPECT_TRUE(HoldsPicture(wb, "F8"));
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, Col("F1"), 1U)));
  for (const char* cell : {"E2", "F2", "F4", "F6", "F8"}) {
    EXPECT_FALSE(HoldsPicture(wb, cell)) << cell;
  }
}

TEST(RichValuePreserve, RowEditsMoveAndDropPictures) {
  // Pictures at E2, E4, E6, E8: an insert above moves them to E3..E9, and
  // deleting row 3 then drops the first and moves the rest up.
  Workbook wb = Load(ReadPicture("cell_pictures"));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0U, 1U)));
  EXPECT_FALSE(HoldsPicture(wb, "E2"));
  EXPECT_TRUE(HoldsPicture(wb, "E3"));
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, Row("E3"), 1U)));
  EXPECT_FALSE(HoldsPicture(wb, "E3"));
  EXPECT_TRUE(HoldsPicture(wb, "E4"));
  std::size_t pictures = 0;
  for (const char* cell : {"E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "E10"}) {
    pictures += HoldsPicture(wb, cell) ? 1U : 0U;
  }
  EXPECT_EQ(pictures, 3U);
}

TEST(RichValuePreserve, ClearedPictureBookKeepsItsSingleMediaPart) {
  // Excel already dropped the cleared picture's media when it saved this book.
  const std::vector<std::uint8_t> excel = ReadPicture("picture_cleared_orphan_media");
  const OoxmlWriteResult saved = Save(Load(excel));
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(MediaParts(saved.bytes), (std::vector<std::string>{"xl/media/image1.png"}));
  ExpectSamePart(excel, saved.bytes, "xl/media/image1.png");
  // G7 is the formula `=E8`, whose value is the picture with CalcOrigin 4.
  EXPECT_EQ(VmCells(saved.bytes, "xl/worksheets/sheet1.xml"),
            (std::map<std::string, std::string>{{"E6", "1"}, {"E8", "3"}, {"G7", "2"}}));
  ExpectSamePart(excel, saved.bytes, kValuesPart);
  ExpectSamePart(excel, saved.bytes, kStructuresPart);
  ExpectSamePart(excel, saved.bytes, kRelPart);
  EXPECT_EQ(Relationships(saved.bytes, kRelRelsPart),
            (std::vector<std::vector<std::string>>{{"rId1", std::string(kImageRel), "../media/image1.png"}}));
}

// ---------------------------------------------------------------------------
// Fallback: a construct the table cannot carry keeps today's output.
// ---------------------------------------------------------------------------

TEST(RichValuePreserve, AnUnknownMetadataTypeKeepsEveryPartVerbatim) {
  const std::vector<std::uint8_t> excel = ReadPicture("cell_pictures");
  std::string metadata = Part(excel, kMetadataPart).value_or("");
  const std::string types = "<metadataTypes count=\"1\">";
  const std::size_t at = metadata.find(types);
  ASSERT_NE(at, std::string::npos);
  metadata.replace(at, types.size(),
                   "<metadataTypes count=\"2\"><metadataType name=\"XLFUTURE\" minSupportedVersion=\"120000\"/>");
  const std::vector<std::uint8_t> package = Repack(excel, {{std::string(kMetadataPart), metadata}});

  const Workbook wb = Load(package);
  EXPECT_TRUE(wb.sheet(0).has_rich_values());
  const OoxmlWriteResult saved = Save(wb);
  ASSERT_FALSE(saved.bytes.empty());
  // Each cell keeps the vm it was read with, so nothing is lost.
  EXPECT_EQ(saved.diagnostics.deferred_feature_count, 0U);
  EXPECT_EQ(VmCells(saved.bytes, "xl/worksheets/sheet1.xml"),
            (std::map<std::string, std::string>{{"E2", "1"}, {"E4", "1"}, {"E6", "2"}, {"E8", "3"}}));
  EXPECT_EQ(CellTypes(saved.bytes, "xl/worksheets/sheet1.xml").at("E2"), "e");
  for (const std::string_view part :
       {kMetadataPart, kValuesPart, kStructuresPart, kTypesPart, kRelPart, kRelRelsPart}) {
    ExpectSamePart(package, saved.bytes, part);
  }
  EXPECT_EQ(MediaParts(saved.bytes), (std::vector<std::string>{"xl/media/image1.png", "xl/media/image2.png"}));
  EXPECT_EQ(rich_value_fallback_reason(wb), kRichFallbackMetadataType);
}

TEST(RichValuePreserve, AnExternalPictureTargetKeepsEveryPartVerbatim) {
  const std::vector<std::uint8_t> excel = ReadPicture("cell_pictures");
  std::string rels = Part(excel, kRelRelsPart).value_or("");
  const std::string target = "Target=\"../media/image1.png\"";
  const std::size_t at = rels.find(target);
  ASSERT_NE(at, std::string::npos);
  rels.replace(at, target.size(), "Target=\"https://example.com/a.png\" TargetMode=\"External\"");
  const Workbook wb = Load(Repack(excel, {{std::string(kRelRelsPart), rels}}));
  EXPECT_TRUE(wb.sheet(0).has_rich_values());
  EXPECT_EQ(rich_value_fallback_reason(wb), kRichFallbackRelExternal);
  const OoxmlWriteResult saved = Save(wb);
  EXPECT_EQ(saved.diagnostics.deferred_feature_count, 0U);
  EXPECT_EQ(VmCells(saved.bytes, "xl/worksheets/sheet1.xml"),
            (std::map<std::string, std::string>{{"E2", "1"}, {"E4", "1"}, {"E6", "2"}, {"E8", "3"}}));
}

// ---------------------------------------------------------------------------
// A book holding only rich errors writes what the error table alone writes.
// ---------------------------------------------------------------------------

// Error-only books whose rich-value parts the save already reproduces byte for
// byte, paired with the Excel file holding those bytes.
constexpr std::pair<const char*, const char*> kFaithful[] = {
    {"s_value", "s_value"},
    {"s_grid", "s_grid"},
    {"c_lambda", "c_lambda"},
    {"combo", "combo"},
    {"multi", "multi"},
    {"notypes", "notypes_resaved"},
    {"notypes_pivot", "notypes_pivot_resaved"},
    {"c_calc", "c_calc"},
    {"e_spill", "e_spill"},
    {"e_edges", "e_edges"},
    {"pvpage_blocked", "pvpage_blocked"},
    {"pivot_blocked", "pivot_blocked"},
    {"b3_phantom", "b3_phantom"},
    {"b4_value", "b4_value"},
    {"b5_fn", "b5_fn"},
    {"v_value", "v_value"},
};

TEST(RichValuePreserve, ErrorOnlyBooksKeepExcelsBytes) {
  for (const auto& [sample, expected_sample] : kFaithful) {
    SCOPED_TRACE(sample);
    const std::vector<std::uint8_t> expected =
        ReadFixture("excel/rich_errors/" + std::string(expected_sample) + ".xlsx");
    const OoxmlWriteResult saved = Save(Load(ReadFixture("excel/rich_errors/" + std::string(sample) + ".xlsx")));
    for (const std::string_view part : {kMetadataPart, kValuesPart, kStructuresPart, kTypesPart}) {
      ExpectSamePart(expected, saved.bytes, part);
    }
  }
}

// Every rich-error book, against the parts and `vm`s the error table builds
// for the loaded model: that table and its serialisers are the save's when no
// cell holds an opaque value. A book the reader keeps verbatim keeps the
// file's own parts.
void ExpectErrorTableOutput(const std::string& relative) {
  SCOPED_TRACE(relative);
  const std::vector<std::uint8_t> excel = ReadFixture(relative);
  const Workbook wb = Load(excel);
  const OoxmlWriteResult saved = Save(wb);
  ASSERT_FALSE(saved.bytes.empty());
  if (keeps_foreign_rich_data(wb.passthrough_parts())) {
    for (const std::string_view part : {kMetadataPart, kValuesPart, kStructuresPart, kTypesPart}) {
      ExpectSamePart(excel, saved.bytes, part);
    }
    return;
  }
  const RichErrorTable table = build_rich_error_table(wb);
  const bool dynamic = has_dynamic_array_formula(wb);
  const std::optional<std::string> metadata = Part(saved.bytes, kMetadataPart);
  ASSERT_EQ(metadata.has_value(), dynamic || !table.empty());
  if (metadata.has_value()) {
    EXPECT_EQ(*metadata, build_metadata_xml(dynamic, table.values.size()));
  }
  ASSERT_EQ(Part(saved.bytes, kValuesPart).has_value(), !table.empty());
  if (!table.empty()) {
    EXPECT_EQ(Part(saved.bytes, kValuesPart).value_or(""), serialize_rdrichvalue(table));
    EXPECT_EQ(Part(saved.bytes, kStructuresPart).value_or(""), serialize_rdrichvaluestructure(table));
    EXPECT_EQ(Part(saved.bytes, kTypesPart).value_or(""), rich_value_types_xml());
  }
  EXPECT_FALSE(Part(saved.bytes, kRelPart).has_value());
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const std::string sheet_part = "xl/worksheets/sheet" + std::to_string(s + 1U) + ".xml";
    const std::map<std::string, std::string> types = CellTypes(saved.bytes, sheet_part);
    const std::map<std::string, std::string> vm = VmCells(saved.bytes, sheet_part);
    for (const auto& [ref, value] : vm) {
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      ASSERT_TRUE(a1::parse_a1_ref(ref, &row, &col)) << ref;
      EXPECT_EQ(value, std::to_string(rich_error_vm_at(table.vm_by_sheet.at(s), row, col))) << ref;
    }
    for (const auto& [key, expected_vm] : table.vm_by_sheet.at(s)) {
      const std::string ref =
          a1::encode_a1(static_cast<std::uint32_t>(key >> 32U), static_cast<std::uint32_t>(key & 0xFFFFFFFFU));
      if (types.count(ref) != 0U && types.at(ref) == "e") {
        EXPECT_EQ(vm.count(ref) != 0U ? vm.at(ref) : std::string(), std::to_string(expected_vm)) << ref;
      }
    }
  }
}

TEST(RichValuePreserve, ErrorOnlyBooksWriteTheErrorTable) {
  std::vector<std::string> books;
  for (const auto& entry : std::filesystem::directory_iterator(FixturePath("excel/rich_errors"))) {
    if (entry.path().extension() == ".xlsx") {
      books.push_back("excel/rich_errors/" + entry.path().filename().string());
    }
  }
  std::sort(books.begin(), books.end());
  EXPECT_GE(books.size(), 40U);
  books.push_back("cellxfs_compaction/mg2_s_merge.xlsx");
  books.push_back("excel/let_a1_names.xlsx");
  books.push_back("excel/let_a1_names2.xlsx");
  for (const std::string& book : books) {
    ExpectErrorTableOutput(book);
  }
}

// ---------------------------------------------------------------------------
// .xlsb: the same table over metadata.bin and the XML richData parts
// ---------------------------------------------------------------------------

constexpr std::string_view kMetadataBin = "xl/metadata.bin";
constexpr std::string_view kSheetBin = "xl/worksheets/sheet1.bin";
constexpr std::string_view kBookBinRels = "xl/_rels/workbook.bin.rels";

std::vector<std::uint8_t> ReadXlsbPicture(std::string_view name) {
  return ReadFixture("excel/cell_pictures/" + std::string(name) + ".xlsb");
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto loaded = xlsb::read_xlsb(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

xlsb::XlsbWriteResult SaveXlsb(const Workbook& wb) {
  auto saved = xlsb::write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : xlsb::XlsbWriteResult{};
}

using RecordList = std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>>;

RecordList RecordsOf(const std::vector<std::uint8_t>& bytes) {
  RecordList out;
  ByteSpan cursor = test::span_of(bytes);
  while (cursor.size > 0U) {
    auto record = xlsb::read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    out.emplace_back(record.value().type,
                     std::vector<std::uint8_t>(record.value().payload.data,
                                               record.value().payload.data + record.value().payload.size));
  }
  return out;
}

RecordList PartRecords(const std::vector<std::uint8_t>& package, std::string_view part) {
  const std::optional<std::string> text = Part(package, part);
  EXPECT_TRUE(text.has_value()) << part;
  const std::string bytes = text.value_or("");
  return RecordsOf(std::vector<std::uint8_t>(bytes.begin(), bytes.end()));
}

std::uint32_t U32(const std::vector<std::uint8_t>& payload) {
  return payload.size() < 4U
             ? 0U
             : static_cast<std::uint32_t>(payload[0]) | (static_cast<std::uint32_t>(payload[1]) << 8U) |
                   (static_cast<std::uint32_t>(payload[2]) << 16U) | (static_cast<std::uint32_t>(payload[3]) << 24U);
}

/// The `BrtValueMeta` index ahead of every cell record in worksheet part `sheet_part`, by A1 reference.
std::map<std::string, std::string> XlsbVmCells(const std::vector<std::uint8_t>& package, std::string_view sheet_part) {
  std::map<std::string, std::string> out;
  std::uint32_t row = 0;
  std::uint32_t vm = 0;
  for (const auto& [type, payload] : PartRecords(package, sheet_part)) {
    if (type == static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtRowHdr)) {
      row = U32(payload);
      vm = 0;
    } else if (type == static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtValueMeta)) {
      vm = U32(payload);
    } else if (type >= static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtCellBlank) &&
               type <= static_cast<std::uint16_t>(xlsb::XlsbRecordType::BrtFmlaError)) {
      if (vm != 0U) {
        out[a1::encode_a1(row, U32(payload))] = std::to_string(vm);
      }
      vm = 0;
    }
  }
  return out;
}

/// The parts a rich-value save regenerates, against `expected`'s.
// The relationships of `part` as sorted (id, type, target) triples: the writer regenerates the part, so its order
// and whitespace are its own, while each id must still name the same target.
std::vector<std::vector<std::string>> RelationshipTargets(const std::vector<std::uint8_t>& package,
                                                          std::string_view part) {
  std::vector<std::vector<std::string>> out = Relationships(package, part);
  std::sort(out.begin(), out.end());
  return out;
}

void ExpectSameXlsbRichParts(const std::vector<std::uint8_t>& expected, const std::vector<std::uint8_t>& actual) {
  for (const std::string_view part : {kValuesPart, kStructuresPart, kTypesPart, kRelPart}) {
    ExpectSamePart(expected, actual, part);
  }
  EXPECT_EQ(RelationshipTargets(actual, kRelRelsPart), RelationshipTargets(expected, kRelRelsPart));
  EXPECT_EQ(PartRecords(actual, kMetadataBin), PartRecords(expected, kMetadataBin));
}

TEST(RichValuePreserve, XlsbCellPicturesReadAsPictureValues) {
  const Workbook wb = LoadXlsb(ReadXlsbPicture("cell_pictures"));
  const Sheet& sheet = wb.sheet(0);
  for (const char* cell : {"E2", "E4", "E6", "E8"}) {
    EXPECT_TRUE(HoldsPicture(wb, cell)) << cell;
    EXPECT_FALSE(sheet.rich_value_at(Row(cell), Col(cell)).has_value()) << cell;
  }
  EXPECT_FALSE(sheet.has_rich_values());
  const Cell* e2 = sheet.cell_at(Row("E2"), Col("E2"));
  ASSERT_NE(e2, nullptr);
  EXPECT_EQ(e2->cached_value.picture_origin(), 5U);
  const PictureEntry* e2_entry = PictureAt(wb, "E2");
  ASSERT_NE(e2_entry, nullptr);
  EXPECT_EQ(e2_entry->media_path, "xl/media/image1.png");
  EXPECT_FALSE(e2_entry->alt.has_value());
  const PictureEntry* e8 = PictureAt(wb, "E8");
  ASSERT_NE(e8, nullptr);
  EXPECT_EQ(e8->media_path, "xl/media/image2.png");
  EXPECT_EQ(e8->alt, std::optional<std::string>("Blue circle on a blue field"));
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    EXPECT_NE(part.path, kMetadataBin);
    EXPECT_NE(part.path, kValuesPart);
  }
}

TEST(RichValuePreserve, XlsbCellPicturesRoundTrip) {
  const std::vector<std::uint8_t> excel = ReadXlsbPicture("cell_pictures");
  const xlsb::XlsbWriteResult saved = SaveXlsb(LoadXlsb(excel));
  ASSERT_FALSE(saved.bytes.empty());

  EXPECT_EQ(XlsbVmCells(saved.bytes, kSheetBin),
            (std::map<std::string, std::string>{{"E2", "1"}, {"E4", "1"}, {"E6", "2"}, {"E8", "3"}}));
  // The table numbers values, structures and richValueRel entries as Excel did.
  ExpectSameXlsbRichParts(excel, saved.bytes);
  EXPECT_EQ(Relationships(saved.bytes, kRelRelsPart),
            (std::vector<std::vector<std::string>>{{"rId1", std::string(kImageRel), "../media/image1.png"},
                                                   {"rId2", std::string(kImageRel), "../media/image2.png"}}));
  EXPECT_EQ(MediaParts(saved.bytes), (std::vector<std::string>{"xl/media/image1.png", "xl/media/image2.png"}));
  ExpectSamePart(excel, saved.bytes, "xl/media/image1.png");
  ExpectSamePart(excel, saved.bytes, "xl/media/image2.png");

  const std::string content_types = Part(saved.bytes, "[Content_Types].xml").value_or("");
  for (const char* override_part : {"<Override PartName=\"/xl/richData/richValueRel.xml\" "
                                    "ContentType=\"application/vnd.ms-excel.richvaluerel+xml\"/>",
                                    "<Override PartName=\"/xl/richData/rdrichvalue.xml\" "
                                    "ContentType=\"application/vnd.ms-excel.rdrichvalue+xml\"/>",
                                    "<Default Extension=\"png\" ContentType=\"image/png\"/>"}) {
    EXPECT_NE(content_types.find(override_part), std::string::npos) << override_part << "\n" << content_types;
  }
  bool rich_rel = false;
  bool metadata_rel = false;
  for (const std::vector<std::string>& rel : Relationships(saved.bytes, kBookBinRels)) {
    rich_rel = rich_rel || (rel[1] == kRelRichValueRel && rel[2] == "richData/richValueRel.xml");
    metadata_rel = metadata_rel || (rel[1] == kRelSheetMetadata && rel[2] == "metadata.bin");
  }
  EXPECT_TRUE(rich_rel);
  EXPECT_TRUE(metadata_rel);

  // A second cycle writes the same parts again.
  const xlsb::XlsbWriteResult resaved = SaveXlsb(LoadXlsb(saved.bytes));
  ExpectSameXlsbRichParts(saved.bytes, resaved.bytes);
  EXPECT_EQ(XlsbVmCells(resaved.bytes, kSheetBin), XlsbVmCells(saved.bytes, kSheetBin));
  EXPECT_EQ(resaved.diagnostics.deferred_feature_count, saved.diagnostics.deferred_feature_count);
}

TEST(RichValuePreserve, XlsbClearedPictureBookKeepsItsSingleMediaPart) {
  const std::vector<std::uint8_t> excel = ReadXlsbPicture("picture_cleared_orphan_media");
  const xlsb::XlsbWriteResult saved = SaveXlsb(LoadXlsb(excel));
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(MediaParts(saved.bytes), (std::vector<std::string>{"xl/media/image1.png"}));
  ExpectSamePart(excel, saved.bytes, "xl/media/image1.png");
  // G7 is the formula `=E8`, whose value is the picture with CalcOrigin 4.
  EXPECT_EQ(XlsbVmCells(saved.bytes, kSheetBin),
            (std::map<std::string, std::string>{{"E6", "1"}, {"E8", "3"}, {"G7", "2"}}));
  ExpectSameXlsbRichParts(excel, saved.bytes);
}

TEST(RichValuePreserve, XlsbOverwritingAPictureDropsItsValueAndOrphanMedia) {
  const std::vector<std::uint8_t> excel = ReadXlsbPicture("cell_pictures");
  const WriteDiagnostics unedited = SaveXlsb(LoadXlsb(excel)).diagnostics;
  Workbook wb = LoadXlsb(excel);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E2"), Col("E2"), Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E4"), Col("E4"), Value::number(2.0))));
  EXPECT_FALSE(HoldsPicture(wb, "E2"));
  EXPECT_TRUE(HoldsPicture(wb, "E6"));

  const xlsb::XlsbWriteResult saved = SaveXlsb(wb);
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(XlsbVmCells(saved.bytes, kSheetBin), (std::map<std::string, std::string>{{"E6", "1"}, {"E8", "2"}}));
  const std::string values = Part(saved.bytes, kValuesPart).value_or("");
  EXPECT_NE(values.find("count=\"2\"><rv s=\"0\"><v>0</v><v>5</v></rv>"
                        "<rv s=\"1\"><v>0</v><v>5</v><v>Blue circle on a blue field</v></rv></rvData>"),
            std::string::npos)
      << values;
  EXPECT_NE(Part(saved.bytes, kRelPart).value_or("").find("<rel r:id=\"rId1\"/></richValueRels>"), std::string::npos);
  EXPECT_EQ(Relationships(saved.bytes, kRelRelsPart),
            (std::vector<std::vector<std::string>>{{"rId1", std::string(kImageRel), "../media/image2.png"}}));
  // image1.png was named only by the dropped value.
  EXPECT_EQ(MediaParts(saved.bytes), (std::vector<std::string>{"xl/media/image2.png"}));
  EXPECT_EQ(saved.diagnostics.dropped_part_count, unedited.dropped_part_count);
  EXPECT_EQ(saved.diagnostics.dropped_relationship_count, unedited.dropped_relationship_count);
  EXPECT_EQ(saved.diagnostics.deferred_feature_count, unedited.deferred_feature_count);
}

TEST(RichValuePreserve, XlsbClearingEveryPictureDropsTheRichParts) {
  const std::vector<std::uint8_t> excel = ReadXlsbPicture("cell_pictures");
  const WriteDiagnostics control = SaveXlsb(LoadXlsb(excel)).diagnostics;
  Workbook wb = LoadXlsb(excel);
  for (const char* cell : {"E2", "E4", "E6", "E8"}) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row(cell), Col(cell), Value::number(1.0)))) << cell;
  }
  const xlsb::XlsbWriteResult saved = SaveXlsb(wb);
  ASSERT_FALSE(saved.bytes.empty());
  for (const std::string& name : PartNames(saved.bytes)) {
    EXPECT_NE(name.compare(0, 12, "xl/richData/"), 0) << name;
  }
  EXPECT_FALSE(Part(saved.bytes, kMetadataBin).has_value());
  EXPECT_TRUE(MediaParts(saved.bytes).empty());
  EXPECT_TRUE(XlsbVmCells(saved.bytes, kSheetBin).empty());
  for (const std::vector<std::string>& rel : Relationships(saved.bytes, kBookBinRels)) {
    EXPECT_NE(rel[1], kRelRichValueRel);
    EXPECT_NE(rel[1], kRelSheetMetadata);
  }
  // The richValueRel link goes with its part on purpose; neither is a loss.
  EXPECT_EQ(saved.diagnostics.dropped_relationship_count, control.dropped_relationship_count);
  EXPECT_EQ(saved.diagnostics.dropped_part_count, control.dropped_part_count);
}

TEST(RichValuePreserve, XlsbAnExternalPictureTargetKeepsEveryPartVerbatim) {
  const std::vector<std::uint8_t> excel = ReadXlsbPicture("cell_pictures");
  const WriteDiagnostics control = SaveXlsb(LoadXlsb(excel)).diagnostics;
  std::string rels = Part(excel, kRelRelsPart).value_or("");
  const std::string target = "Target=\"../media/image1.png\"";
  const std::size_t at = rels.find(target);
  ASSERT_NE(at, std::string::npos);
  rels.replace(at, target.size(), "Target=\"https://example.com/a.png\" TargetMode=\"External\"");
  const std::vector<std::uint8_t> package = Repack(excel, {{std::string(kRelRelsPart), rels}});

  const Workbook wb = LoadXlsb(package);
  EXPECT_TRUE(wb.sheet(0).has_rich_values());
  EXPECT_EQ(rich_value_fallback_reason(wb), kRichFallbackRelExternal);
  const xlsb::XlsbWriteResult saved = SaveXlsb(wb);
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(saved.diagnostics.deferred_feature_count, control.deferred_feature_count);
  EXPECT_EQ(XlsbVmCells(saved.bytes, kSheetBin),
            (std::map<std::string, std::string>{{"E2", "1"}, {"E4", "1"}, {"E6", "2"}, {"E8", "3"}}));
  ExpectSameXlsbRichParts(package, saved.bytes);
  EXPECT_EQ(MediaParts(saved.bytes), (std::vector<std::string>{"xl/media/image1.png", "xl/media/image2.png"}));
}

TEST(RichValuePreserve, PicturesCrossBetweenXlsxAndXlsb) {
  const std::vector<std::uint8_t> xlsx = ReadPicture("cell_pictures");
  const std::vector<std::uint8_t> xlsb_excel = ReadXlsbPicture("cell_pictures");
  const std::map<std::string, std::string> vm{{"E2", "1"}, {"E4", "1"}, {"E6", "2"}, {"E8", "3"}};

  // .xlsx to .xlsb: the rich-value parts Excel wrote for the same book.
  const xlsb::XlsbWriteResult to_xlsb = SaveXlsb(Load(xlsx));
  ASSERT_FALSE(to_xlsb.bytes.empty());
  EXPECT_EQ(XlsbVmCells(to_xlsb.bytes, kSheetBin), vm);
  ExpectSameXlsbRichParts(xlsb_excel, to_xlsb.bytes);
  EXPECT_EQ(MediaParts(to_xlsb.bytes), (std::vector<std::string>{"xl/media/image1.png", "xl/media/image2.png"}));
  const Workbook reread = LoadXlsb(to_xlsb.bytes);
  for (const char* cell : {"E2", "E4", "E6", "E8"}) {
    EXPECT_TRUE(HoldsPicture(reread, cell)) << cell;
  }
  EXPECT_FALSE(Part(to_xlsb.bytes, kMetadataPart).has_value());

  // .xlsb to .xlsx.
  const OoxmlWriteResult to_xlsx = Save(LoadXlsb(xlsb_excel));
  ASSERT_FALSE(to_xlsx.bytes.empty());
  EXPECT_EQ(VmCells(to_xlsx.bytes, "xl/worksheets/sheet1.xml"), vm);
  for (const std::string_view part : {kMetadataPart, kValuesPart, kStructuresPart, kTypesPart, kRelPart}) {
    ExpectSamePart(xlsx, to_xlsx.bytes, part);
  }
  EXPECT_EQ(RelationshipTargets(to_xlsx.bytes, kRelRelsPart), RelationshipTargets(xlsx, kRelRelsPart));
  EXPECT_EQ(MediaParts(to_xlsx.bytes), (std::vector<std::string>{"xl/media/image1.png", "xl/media/image2.png"}));
  EXPECT_FALSE(Part(to_xlsx.bytes, kMetadataBin).has_value());
}

// ---------------------------------------------------------------------------
// .xlsb books holding only rich errors write what the error table alone writes.
// ---------------------------------------------------------------------------

// A blocked pivot's anchor keeps its rich value only when its table part is
// written, as the .xlsb save's pivot plan decides.
std::vector<std::vector<bool>> WrittenPivots(const Workbook& wb) {
  std::unordered_set<std::string> taken;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    taken.insert(part.path);
  }
  xlsb::EmissionPlan plan;
  return xlsb::plan_pivot_parts(wb, taken, plan).written;
}

TEST(RichValuePreserve, XlsbErrorOnlyBooksWriteTheErrorTable) {
  std::vector<std::string> books;
  for (const auto& entry : std::filesystem::directory_iterator(FixturePath("excel/rich_errors_xlsb"))) {
    if (entry.path().extension() == ".xlsb") {
      books.push_back("excel/rich_errors_xlsb/" + entry.path().filename().string());
    }
  }
  std::sort(books.begin(), books.end());
  EXPECT_GE(books.size(), 30U);
  for (const std::string& book : books) {
    SCOPED_TRACE(book);
    const std::vector<std::uint8_t> excel = ReadFixture(book);
    const Workbook wb = LoadXlsb(excel);
    const xlsb::XlsbWriteResult saved = SaveXlsb(wb);
    ASSERT_FALSE(saved.bytes.empty());
    bool retained_metadata = false;
    for (const PassthroughPart& part : wb.passthrough_parts()) {
      retained_metadata = retained_metadata || part.path == kMetadataBin;
    }
    if (retained_metadata) {
      // The reader kept the file's numbering, so the save does too.
      for (const std::string_view part : {kValuesPart, kStructuresPart, kTypesPart}) {
        ExpectSamePart(excel, saved.bytes, part);
      }
      EXPECT_EQ(PartRecords(saved.bytes, kMetadataBin), PartRecords(excel, kMetadataBin));
      continue;
    }
    const std::vector<std::vector<bool>> written = WrittenPivots(wb);
    const RichErrorTable errors = build_rich_error_table(wb, &written);
    const RichErrorTable table = build_rich_value_table(wb, &written);
    ASSERT_EQ(table.values.size(), errors.values.size());
    EXPECT_TRUE(table.opaque.empty());
    EXPECT_EQ(serialize_rdrichvalue(table), serialize_rdrichvalue(errors));
    EXPECT_EQ(serialize_rdrichvaluestructure(table), serialize_rdrichvaluestructure(errors));
    EXPECT_EQ(table.vm_by_sheet, errors.vm_by_sheet);

    const bool dynamic = has_dynamic_array_formula(wb);
    ASSERT_EQ(Part(saved.bytes, kMetadataBin).has_value(), dynamic || !errors.empty());
    if (dynamic || !errors.empty()) {
      EXPECT_EQ(PartRecords(saved.bytes, kMetadataBin),
                RecordsOf(xlsb::build_metadata_bin(dynamic, errors.values.size())));
    }
    ASSERT_EQ(Part(saved.bytes, kValuesPart).has_value(), !errors.empty());
    if (!errors.empty()) {
      EXPECT_EQ(Part(saved.bytes, kValuesPart).value_or(""), serialize_rdrichvalue(errors));
      EXPECT_EQ(Part(saved.bytes, kStructuresPart).value_or(""), serialize_rdrichvaluestructure(errors));
      EXPECT_EQ(Part(saved.bytes, kTypesPart).value_or(""), rich_value_types_xml());
    }
    EXPECT_FALSE(Part(saved.bytes, kRelPart).has_value());
    for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
      const std::map<std::string, std::string> vm =
          XlsbVmCells(saved.bytes, "xl/worksheets/sheet" + std::to_string(s + 1U) + ".bin");
      for (const auto& [ref, value] : vm) {
        EXPECT_EQ(value, std::to_string(rich_error_vm_at(errors.vm_by_sheet.at(s), Row(ref.c_str()), Col(ref.c_str()))))
            << ref;
      }
    }
  }
}

// ---------------------------------------------------------------------------
// What a cached-value write does to a cell's opaque rich value
// ---------------------------------------------------------------------------

OpaqueRichValueKey ValueKey(std::string name, std::string type, std::string value) {
  OpaqueRichValueKey key;
  key.name = std::move(name);
  key.type = std::move(type);
  key.value = std::move(value);
  return key;
}

TEST(SheetRichValue, PictureWriteClearsOpaqueAndValueErrorKeeps) {
  // A linked data type and a `_localImage` the reader left opaque (here, one with a style on its origin key).
  OpaqueRichValue linked;
  linked.structure_type = "_linkedentity2";
  linked.keys.push_back(ValueKey("%cvb", "i", "1"));
  OpaqueRichValue image;
  image.structure_type = std::string(kPictureStructure);
  OpaqueRichValueKey rel = ValueKey(std::string(kImageKey), "i", "");
  rel.rel_target = "xl/media/image1.png";
  rel.rel_type = std::string(kImageRelType);
  image.keys.push_back(rel);
  OpaqueRichValueKey origin = ValueKey(std::string(kOriginKey), "i", "4");
  origin.style = "0";
  image.keys.push_back(origin);

  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  const auto holds = [&sheet](const OpaqueRichValue& value) {
    const std::optional<OpaqueRichValue> held = sheet.rich_value_at(0, 0);
    return held.has_value() && *held == value;
  };
  const auto place = [&sheet](const OpaqueRichValue& value) {
    sheet.set_cell_cached_value(0, 0, Value::error(ErrorCode::Value));
    sheet.set_rich_value(0, 0, value);
  };
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::error(ErrorCode::Value))));
  for (const OpaqueRichValue* value : {&linked, &image}) {
    SCOPED_TRACE(value->structure_type);
    place(*value);
    sheet.set_cell_cached_value(0, 0, Value::error(ErrorCode::Value));
    EXPECT_TRUE(holds(*value));

    place(*value);
    sheet.set_cell_cached_value(0, 0, Value::picture(PicturePayload{0U, 0U}, 4U));
    EXPECT_FALSE(sheet.rich_value_at(0, 0).has_value());

    place(*value);
    sheet.set_cell_cached_value(0, 0, Value::number(1.0));
    EXPECT_FALSE(sheet.rich_value_at(0, 0).has_value());

    place(*value);
    sheet.set_cell_cached_value(0, 0, Value::error(ErrorCode::NA));
    EXPECT_FALSE(sheet.rich_value_at(0, 0).has_value());
  }
}

}  // namespace
}  // namespace io
}  // namespace formulon
