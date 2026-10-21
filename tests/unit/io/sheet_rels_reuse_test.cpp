// Sheet relationship consumers share one parsed DOM so each OOXML rels
// payload is extracted from the ZIP archive only once.

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml/sheet_aux_rels_reader.h"
#include "io/ooxml_defs.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/error.h"

namespace formulon::io::ooxml {
namespace {

constexpr std::string_view kSheetRelsPath = "xl/worksheets/_rels/sheet1.xml.rels";
constexpr std::string_view kSheetDir = "xl/worksheets";
constexpr std::string_view kRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

std::vector<std::uint8_t> MakeZip(std::string_view path, std::string_view payload) {
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 0), MZ_TRUE);
  const std::string name(path);
  EXPECT_EQ(mz_zip_writer_add_mem(&writer, name.c_str(), payload.data(), payload.size(),
                                  static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
            MZ_TRUE);
  void* heap = nullptr;
  std::size_t heap_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&writer, &heap, &heap_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&writer), MZ_TRUE);
  std::vector<std::uint8_t> bytes(heap_size);
  std::memcpy(bytes.data(), heap, heap_size);
  mz_free(heap);
  return bytes;
}

std::string MixedSheetRels() {
  std::string xml =
      "<pkg:Relationships xmlns:pkg=\"" + std::string(kRelationshipsNamespace) + "\" xmlns:f=\"urn:foreign\">";
  // A foreign-namespace sibling must not be mistaken for a package relationship.
  xml += "<f:Relationship Id=\"foreign\" Type=\"" + std::string(kRelTable) + "\" Target=\"wrong.xml\"/>";
  xml += "<pkg:Relationship Id=\"tableB\" Type=\"" + std::string(kRelTable) + "\" Target=\"../tables/tableB.xml\"/>";
  xml += "<pkg:Relationship Id=\"pivotB\" Type=\"" + std::string(kRelPivotTable) +
         "\" Target=\"../pivotTables/pivotB.xml\"/>";
  xml += "<pkg:Relationship Id=\"link\" Type=\"" + std::string(kRelHyperlink) +
         "\" Target=\"https://example.test/a?x=1&amp;y=2\" TargetMode=\"External\"/>";
  xml +=
      "<pkg:Relationship Id=\"unknownExternal\" Type=\"urn:custom:external\" "
      "Target=\"custom://host/item\" TargetMode=\"External\"/>";
  xml += "<pkg:Relationship Id=\"comments\" Type=\"" + std::string(kRelComments) + "\" Target=\"../comments1.xml\"/>";
  xml += "<pkg:Relationship Id=\"vmlComment\" Type=\"" + std::string(kRelVmlDrawing) +
         "\" Target=\"../drawings/comments.vml\"/>";
  xml += "<pkg:Relationship Id=\"vmlHeader\" Type=\"" + std::string(kRelVmlDrawing) +
         "\" Target=\"../drawings/header.vml\"/>";
  xml += "<pkg:Relationship Id=\"tableA\" Type=\"" + std::string(kRelTable) + "\" Target=\"../tables/tableA.xml\"/>";
  xml += "<pkg:Relationship Id=\"pivotA\" Type=\"" + std::string(kRelPivotTable) +
         "\" Target=\"../pivotTables/pivotA.xml\"/>";
  xml += "<pkg:Relationship Id=\"unknownInternal\" Type=\"urn:custom:internal\" Target=\"../custom/item.xml\"/>";
  xml += "<pkg:Relationship Id=\"printer\" Type=\"" + std::string(kRelPrinterSettings) +
         "\" Target=\"../printerSettings/printer.bin\"/>";
  xml += "<pkg:Relationship Id=\"threaded\" Type=\"" + std::string(kRelThreadedComment) +
         "\" Target=\"../threadedComments/threadedComment1.xml\"/>";
  xml += "<pkg:Relationship Id=\"drawing\" Type=\"" + std::string(kRelDrawing) +
         "\" Target=\"../drawings/drawing1.xml\"/>";
  xml += "</pkg:Relationships>";
  return xml;
}

TEST(SheetRelsReuse, OneExtractionFeedsEveryCollectorAndPreservesOrderAndTargets) {
  const std::string xml = MixedSheetRels();
  const std::vector<std::uint8_t> package = MakeZip(kSheetRelsPath, xml);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()}, xml.size())));

  auto document_or = load_relationship_document(zip, kSheetRelsPath, "sheet rels", "ooxml_reader");
  ASSERT_TRUE(static_cast<bool>(document_or)) << document_or.error().message;
  pugi::xml_document document = document_or.take();

  auto tables_or = collect_sheet_table_targets(document, kSheetDir);
  ASSERT_TRUE(static_cast<bool>(tables_or)) << tables_or.error().message;
  EXPECT_EQ(tables_or.value(), (std::vector<std::string>{"xl/tables/tableB.xml", "xl/tables/tableA.xml"}));

  auto pivots_or = collect_sheet_pivot_table_targets(document, kSheetDir);
  ASSERT_TRUE(static_cast<bool>(pivots_or)) << pivots_or.error().message;
  EXPECT_EQ(pivots_or.value(), (std::vector<std::string>{"xl/pivotTables/pivotB.xml", "xl/pivotTables/pivotA.xml"}));

  // No body r:id is supplied: comment VML selection uses the documented
  // first-candidate fallback, retaining later VML relationships as unknown.
  auto aux_or = collect_sheet_aux_rels(document, kSheetDir, "");
  ASSERT_TRUE(static_cast<bool>(aux_or)) << aux_or.error().message;
  const SheetAuxRels& aux = aux_or.value();
  EXPECT_EQ(aux.hyperlink_rid_to_target.at("link"), "https://example.test/a?x=1&y=2");
  EXPECT_EQ(aux.comments_path, "xl/comments1.xml");
  EXPECT_EQ(aux.vml_path, "xl/drawings/comments.vml");
  EXPECT_EQ(aux.printer_settings_rid, "printer");
  EXPECT_EQ(aux.printer_settings_path, "xl/printerSettings/printer.bin");
  EXPECT_EQ(aux.threaded_comments_path, "xl/threadedComments/threadedComment1.xml");
  EXPECT_EQ(aux.drawing_path, "xl/drawings/drawing1.xml");
  ASSERT_EQ(aux.unknown_rels.size(), 3U);
  EXPECT_EQ(aux.unknown_rels[0].id, "unknownExternal");
  EXPECT_EQ(aux.unknown_rels[0].target, "custom://host/item");
  EXPECT_TRUE(aux.unknown_rels[0].target_external);
  EXPECT_EQ(aux.unknown_rels[1].id, "unknownInternal");
  EXPECT_EQ(aux.unknown_rels[1].target, "xl/custom/item.xml");
  EXPECT_FALSE(aux.unknown_rels[1].target_external);
  EXPECT_EQ(aux.unknown_rels[2].id, "vmlHeader");
  EXPECT_EQ(aux.unknown_rels[2].target, "xl/drawings/header.vml");
  EXPECT_FALSE(aux.unknown_rels[2].target_external);

  auto reread_or = zip.read_entry(kSheetRelsPath);
  ASSERT_FALSE(static_cast<bool>(reread_or));
  EXPECT_EQ(reread_or.error().code, FormulonErrorCode::kIoFileTooLarge);
}

TEST(SheetRelsReuse, DocumentLoaderKeepsMalformedAndMissingEnvelopeErrors) {
  const std::vector<std::uint8_t> malformed = MakeZip(kSheetRelsPath, "<Relationships>");
  ZipReader malformed_zip;
  ASSERT_TRUE(static_cast<bool>(malformed_zip.open(ByteSpan{malformed.data(), malformed.size()})));
  auto parse_or = load_relationship_document(malformed_zip, kSheetRelsPath, "sheet rels", "ooxml_reader");
  ASSERT_FALSE(static_cast<bool>(parse_or));
  EXPECT_EQ(parse_or.error().code, FormulonErrorCode::kIoXmlParse);
  EXPECT_EQ(parse_or.error().message, "sheet rels: pugixml parse failed");
  EXPECT_NE(parse_or.error().context.find("context=ooxml_reader part=xl/worksheets/_rels/sheet1.xml.rels desc="),
            std::string::npos);

  const std::vector<std::uint8_t> missing_root = MakeZip(kSheetRelsPath, "<wrongRoot/>");
  ZipReader missing_root_zip;
  ASSERT_TRUE(static_cast<bool>(missing_root_zip.open(ByteSpan{missing_root.data(), missing_root.size()})));
  auto envelope_or = load_relationship_document(missing_root_zip, kSheetRelsPath, "sheet rels", "ooxml_reader");
  ASSERT_FALSE(static_cast<bool>(envelope_or));
  EXPECT_EQ(envelope_or.error().code, FormulonErrorCode::kIoRelationshipBroken);
  EXPECT_EQ(envelope_or.error().message, "sheet rels: missing <Relationships>");
  EXPECT_EQ(envelope_or.error().context, "context=ooxml_reader part=xl/worksheets/_rels/sheet1.xml.rels");
}

}  // namespace
}  // namespace formulon::io::ooxml
