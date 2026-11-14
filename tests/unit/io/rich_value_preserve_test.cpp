// Rich values the engine does not model survive an .xlsx load and save: the
// in-cell pictures Excel wrote, an overwritten picture dropping its value and
// its now-unreferenced media, a book the table cannot carry keeping its parts
// verbatim with a save diagnostic, and an error-only book writing the same
// rich-value parts as the error table alone.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/dynamic_array_formula.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/rich_value_errors.h"
#include "io/rich_values.h"
#include "io/zip_reader.h"
#include "miniz.h"
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

// ---------------------------------------------------------------------------
// In-cell pictures
// ---------------------------------------------------------------------------

TEST(RichValuePreserve, CellPicturesReadAsOpaqueValues) {
  const Workbook wb = Load(ReadPicture("cell_pictures"));
  const Sheet& sheet = wb.sheet(0);
  for (const char* cell : {"E2", "E4", "E6", "E8"}) {
    EXPECT_TRUE(sheet.rich_value_at(Row(cell), Col(cell)).has_value()) << cell;
  }
  const std::optional<OpaqueRichValue> e2 = sheet.rich_value_at(Row("E2"), Col("E2"));
  ASSERT_TRUE(e2.has_value());
  EXPECT_EQ(e2->structure_type, "_localImage");
  ASSERT_EQ(e2->keys.size(), 2U);
  EXPECT_EQ(e2->keys[0].name, "_rvRel:LocalImageIdentifier");
  EXPECT_EQ(e2->keys[0].rel_target, "xl/media/image1.png");
  EXPECT_EQ(e2->keys[0].rel_type, kImageRel);
  EXPECT_EQ(e2->keys[1].value, "5");
  const std::optional<OpaqueRichValue> e8 = sheet.rich_value_at(Row("E8"), Col("E8"));
  ASSERT_TRUE(e8.has_value());
  ASSERT_EQ(e8->keys.size(), 3U);
  EXPECT_EQ(e8->keys[0].rel_target, "xl/media/image2.png");
  EXPECT_EQ(e8->keys[2].name, "Text");
  EXPECT_EQ(e8->keys[2].type, "s");
  EXPECT_EQ(e8->keys[2].value, "Blue circle on a blue field");
  // The cells still read as their cached #VALUE!.
  EXPECT_EQ(sheet.cell_at(Row("E2"), Col("E2"))->cached_value.as_error(), ErrorCode::Value);
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
  EXPECT_FALSE(wb.sheet(0).rich_value_at(Row("E2"), Col("E2")).has_value());
  EXPECT_FALSE(wb.sheet(0).rich_value_at(Row("E4"), Col("E4")).has_value());
  EXPECT_TRUE(wb.sheet(0).rich_value_at(Row("E6"), Col("E6")).has_value());

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
  ASSERT_TRUE(wb.sheet(0).rich_value_at(Row("G2"), Col("G2")).has_value());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, Row("E2"), Col("E2"), Value::number(7.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_FALSE(wb.sheet(0).rich_value_at(Row("G2"), Col("G2")).has_value());
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
  EXPECT_FALSE(wb.sheet(0).rich_value_at(Row("E2"), Col("E2")).has_value());
  EXPECT_TRUE(wb.sheet(0).rich_value_at(Row("F2"), Col("F2")).has_value());
  EXPECT_TRUE(wb.sheet(0).rich_value_at(Row("F8"), Col("F8")).has_value());
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, Col("F1"), 1U)));
  EXPECT_FALSE(wb.sheet(0).has_rich_values());
}

TEST(RichValuePreserve, RowEditsMoveAndDropPictures) {
  // Pictures at E2, E4, E6, E8: an insert above moves them to E3..E9, and
  // deleting row 3 then drops the first and moves the rest up.
  Workbook wb = Load(ReadPicture("cell_pictures"));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0U, 1U)));
  EXPECT_FALSE(wb.sheet(0).rich_value_at(Row("E2"), Col("E2")).has_value());
  EXPECT_TRUE(wb.sheet(0).rich_value_at(Row("E3"), Col("E3")).has_value());
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, Row("E3"), 1U)));
  EXPECT_FALSE(wb.sheet(0).rich_value_at(Row("E3"), Col("E3")).has_value());
  EXPECT_TRUE(wb.sheet(0).rich_value_at(Row("E4"), Col("E4")).has_value());
  EXPECT_EQ(wb.sheet(0).rich_values().size(), 3U);
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
  EXPECT_FALSE(wb.sheet(0).has_rich_values());
  const OoxmlWriteResult saved = Save(wb);
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(saved.diagnostics.deferred_feature_count, 1U);
  EXPECT_TRUE(VmCells(saved.bytes, "xl/worksheets/sheet1.xml").empty());
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
  EXPECT_FALSE(wb.sheet(0).has_rich_values());
  EXPECT_EQ(rich_value_fallback_reason(wb), kRichFallbackRelExternal);
  EXPECT_EQ(Save(wb).diagnostics.deferred_feature_count, 1U);
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
// for the loaded model: that table and its serialisers are the .xlsb save's,
// which the opaque values do not reach. A book the reader keeps verbatim
// keeps the file's own parts.
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

}  // namespace
}  // namespace io
}  // namespace formulon
