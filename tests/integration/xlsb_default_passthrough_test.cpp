// XLSB Default-content-type passthrough coverage. The package envelope is
// deliberately synthetic: the workbook/sheet bodies come from the real
// XLSB writer, while media/OLE/custom XML entries and their relationships
// are added as opaque package parts. This keeps the test focused on the
// residual-part classifier and exercises two complete load/save cycles.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <map>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/passthrough_part.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

io::ByteSpan SpanOf(const Bytes& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

Bytes ToBytes(std::string_view text) {
  return Bytes(text.begin(), text.end());
}

Bytes PngBytes() {
  return {0x89U, 'P', 'N', 'G', '\r', '\n', 0x1AU, '\n', 'f', 'o', 'r', 'm', 'u', 'l', 'o', 'n'};
}

Bytes OleBytes() {
  return {0x00U, 'O', 'L', 'E', '-', 'f', 'o', 'r', 'm', 'u', 'l', 'o', 'n', 0xFFU};
}

Bytes ReadPart(io::ZipReader& zip, std::string_view path) {
  auto bytes_or = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << "missing part: " << path;
  return bytes_or ? bytes_or.value() : Bytes{};
}

/// Rebuilds a ZIP while replacing selected entries and appending new ones.
/// The helper intentionally routes reads through ZipReader, matching the
/// production decompression and size policy used by the XLSB reader.
Bytes RewriteZip(const Bytes& source, const std::map<std::string, Bytes>& replacements,
                 const std::vector<std::pair<std::string, Bytes>>& additions) {
  io::ZipReader input;
  EXPECT_TRUE(static_cast<bool>(input.open(SpanOf(source))));
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  for (const std::string& path : input.list_entries()) {
    auto replacement = replacements.find(path);
    Bytes body = replacement == replacements.end() ? ReadPart(input, path) : replacement->second;
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, path.c_str(), body.data(), body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE)
        << "failed to rewrite " << path;
  }
  for (const auto& [path, body] : additions) {
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, path.c_str(), body.data(), body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE)
        << "failed to add " << path;
  }
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&writer), MZ_TRUE);
  Bytes out(static_cast<const std::uint8_t*>(archive), static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return out;
}

Bytes BaseXlsb() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_value(0U, 0U, Value::number(42.0));
  auto bytes_or = io::xlsb::write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;
  return bytes_or ? bytes_or.value() : Bytes{};
}

Bytes ContentTypesWithDefaults(io::ZipReader& zip, bool conflicting_defaults = false, bool duplicate_override = false,
                               bool conflicting_override = false) {
  const Bytes raw = ReadPart(zip, "[Content_Types].xml");
  std::string content_types(reinterpret_cast<const char*>(raw.data()), raw.size());
  const std::string old_bin =
      "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>";
  const std::string ole_bin = "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>";
  const std::size_t bin_pos = content_types.find(old_bin);
  EXPECT_NE(bin_pos, std::string::npos);
  if (bin_pos != std::string::npos) {
    content_types.replace(bin_pos, old_bin.size(), ole_bin);
  }
  const std::string old_xml = "<Default Extension=\"xml\" ContentType=\"application/xml\"/>";
  const std::string opaque_xml = "<Default Extension=\"xml\" ContentType=\"application/vnd.example.opaque+xml\"/>";
  const std::size_t xml_pos = content_types.find(old_xml);
  EXPECT_NE(xml_pos, std::string::npos);
  if (xml_pos != std::string::npos) {
    content_types.replace(xml_pos, old_xml.size(), opaque_xml);
  }
  const std::string workbook_override =
      "<Override PartName=\"/xl/workbook.bin\" "
      "ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>";
  const std::size_t close = content_types.rfind("</Types>");
  EXPECT_NE(close, std::string::npos);
  if (close != std::string::npos) {
    content_types.insert(close, "<Default Extension=\"PNG\" ContentType=\"image/png\"/>");
    content_types.insert(close, workbook_override);
    if (duplicate_override) {
      content_types.insert(close, workbook_override);
    }
    if (conflicting_defaults) {
      content_types.insert(close, "<Default Extension=\"PNG\" ContentType=\"image/gif\"/>");
    }
    if (conflicting_override) {
      content_types.insert(close,
                           "<Override PartName=\"/xl/workbook.bin\" ContentType=\"application/vnd.example.other\"/>");
    }
  }
  return ToBytes(content_types);
}

Bytes DefaultPackage(bool conflicting_defaults = false, bool add_orphan = false, bool add_unsafe = false,
                     bool duplicate_override = false, bool conflicting_override = false) {
  const Bytes base = BaseXlsb();
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(base))));
  std::map<std::string, Bytes> replacements;
  replacements.emplace("[Content_Types].xml",
                       ContentTypesWithDefaults(zip, conflicting_defaults, duplicate_override, conflicting_override));

  const Bytes rels = ReadPart(zip, "xl/_rels/workbook.bin.rels");
  std::string workbook_rels(reinterpret_cast<const char*>(rels.data()), rels.size());
  const std::size_t rels_close = workbook_rels.rfind("</Relationships>");
  EXPECT_NE(rels_close, std::string::npos);
  if (rels_close != std::string::npos) {
    workbook_rels.insert(rels_close,
                         "<Relationship Id=\"rIdOle\" "
                         "Type=\"http://schemas.microsoft.com/office/2006/relationships/oleObject\" "
                         "Target=\"embeddings/oleObject1.bin\"/>"
                         "<Relationship Id=\"rIdCustom\" "
                         "Type=\"http://schemas.example.test/customXml\" Target=\"customXml/item1.xml\"/>");
  }
  replacements.emplace("xl/_rels/workbook.bin.rels", ToBytes(workbook_rels));

  const std::string sheet_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rIdImage\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
      "Target=\"../media/image1.PNG\"/>"
      "</Relationships>";
  const std::vector<std::pair<std::string, Bytes>> additions = {
      {"xl/worksheets/_rels/sheet1.bin.rels", ToBytes(sheet_rels)},
      {"xl/media/image1.PNG", PngBytes()},
      {"xl/embeddings/oleObject1.bin", OleBytes()},
      {"xl/customXml/item1.xml", ToBytes("<custom>formulon</custom>")},
  };
  std::vector<std::pair<std::string, Bytes>> extras = additions;
  if (add_orphan) {
    extras.emplace_back("orphan.unknown", ToBytes("cannot-resolve"));
  }
  if (add_unsafe) {
    extras.emplace_back("../escape.bin", ToBytes("unsafe"));
  }
  return RewriteZip(base, replacements, extras);
}

std::string EffectiveContentType(io::ZipReader& zip, std::string_view path) {
  const Bytes raw = ReadPart(zip, "[Content_Types].xml");
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(raw.data(), raw.size()));
  const pugi::xml_node root = doc.child("Types");
  for (pugi::xml_node node = root.child("Override"); node; node = node.next_sibling("Override")) {
    std::string part_name = node.attribute("PartName").value();
    if (!part_name.empty() && part_name.front() == '/') {
      part_name.erase(0, 1);
    }
    if (part_name == path) {
      return node.attribute("ContentType").value();
    }
  }
  std::string extension;
  const std::size_t dot = path.find_last_of('.');
  if (dot != std::string_view::npos && dot + 1U < path.size()) {
    extension.assign(path.substr(dot + 1U));
    for (char& c : extension) {
      if (c >= 'A' && c <= 'Z') {
        c = static_cast<char>(c - 'A' + 'a');
      }
    }
  }
  for (pugi::xml_node node = root.child("Default"); node; node = node.next_sibling("Default")) {
    std::string candidate = node.attribute("Extension").value();
    for (char& c : candidate) {
      if (c >= 'A' && c <= 'Z') {
        c = static_cast<char>(c - 'A' + 'a');
      }
    }
    if (candidate == extension) {
      return node.attribute("ContentType").value();
    }
  }
  return {};
}

bool HasContentTypeOverride(io::ZipReader& zip, std::string_view path) {
  const Bytes raw = ReadPart(zip, "[Content_Types].xml");
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(raw.data(), raw.size()));
  const pugi::xml_node root = doc.child("Types");
  for (pugi::xml_node node = root.child("Override"); node; node = node.next_sibling("Override")) {
    std::string part_name = node.attribute("PartName").value();
    if (!part_name.empty() && part_name.front() == '/') {
      part_name.erase(0, 1);
    }
    if (part_name == path) {
      return true;
    }
  }
  return false;
}

std::vector<std::string> InventoryNames(io::ZipReader& zip) {
  std::vector<std::string> names = zip.list_entries();
  std::sort(names.begin(), names.end());
  return names;
}

std::vector<std::string> RelationshipSignatures(io::ZipReader& zip, std::string_view path) {
  const Bytes raw = ReadPart(zip, path);
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(raw.data(), raw.size()));
  const pugi::xml_node root = doc.child("Relationships");
  std::vector<std::string> signatures;
  for (pugi::xml_node node = root.child("Relationship"); node; node = node.next_sibling("Relationship")) {
    std::string signature = node.attribute("Type").value();
    signature.push_back('\0');
    signature.append(node.attribute("Target").value());
    signature.push_back('\0');
    signature.append(node.attribute("TargetMode").value());
    signatures.push_back(std::move(signature));
  }
  std::sort(signatures.begin(), signatures.end());
  return signatures;
}

const io::PassthroughPart* FindPart(const Workbook& wb, std::string_view path) {
  for (const io::PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return &part;
    }
  }
  return nullptr;
}

TEST(XlsbDefaultPassthrough, CapturesDefaultsAndRoundTripsTwoCycles) {
  const Bytes source = DefaultPackage();
  io::ZipReader source_zip;
  ASSERT_TRUE(static_cast<bool>(source_zip.open(SpanOf(source))));
  const std::vector<std::string> source_inventory = InventoryNames(source_zip);
  const std::vector<std::string> source_workbook_rels =
      RelationshipSignatures(source_zip, "xl/_rels/workbook.bin.rels");
  const std::vector<std::string> source_sheet_rels =
      RelationshipSignatures(source_zip, "xl/worksheets/_rels/sheet1.bin.rels");

  auto first_or = io::xlsb::read_xlsb(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(first_or)) << first_or.error().message << " | " << first_or.error().context;
  EXPECT_EQ(first_or.value().dropped_part_count, 0U);
  const Workbook& first = first_or.value().workbook;
  const io::PassthroughPart* image = FindPart(first, "xl/media/image1.PNG");
  const io::PassthroughPart* ole = FindPart(first, "xl/embeddings/oleObject1.bin");
  const io::PassthroughPart* custom = FindPart(first, "xl/customXml/item1.xml");
  ASSERT_NE(image, nullptr);
  ASSERT_NE(ole, nullptr);
  ASSERT_NE(custom, nullptr);
  EXPECT_TRUE(image->content_type.empty());
  EXPECT_TRUE(ole->content_type.empty());
  EXPECT_TRUE(custom->content_type.empty());
  EXPECT_EQ(image->bytes, PngBytes());
  EXPECT_EQ(ole->bytes, OleBytes());
  EXPECT_EQ(custom->bytes, ToBytes("<custom>formulon</custom>"));

  auto save1_or = io::xlsb::write_xlsb(first);
  ASSERT_TRUE(static_cast<bool>(save1_or)) << save1_or.error().message << " | " << save1_or.error().context;
  io::ZipReader zip1;
  ASSERT_TRUE(static_cast<bool>(zip1.open(SpanOf(save1_or.value()))));
  EXPECT_EQ(InventoryNames(zip1), source_inventory);
  EXPECT_EQ(ReadPart(zip1, "xl/media/image1.PNG"), PngBytes());
  EXPECT_EQ(ReadPart(zip1, "xl/embeddings/oleObject1.bin"), OleBytes());
  EXPECT_EQ(ReadPart(zip1, "xl/customXml/item1.xml"), ToBytes("<custom>formulon</custom>"));
  EXPECT_EQ(EffectiveContentType(zip1, "xl/media/image1.PNG"), "image/png");
  EXPECT_EQ(EffectiveContentType(zip1, "xl/embeddings/oleObject1.bin"), "application/vnd.ms-office.vbaProject");
  EXPECT_EQ(EffectiveContentType(zip1, "xl/customXml/item1.xml"), "application/vnd.example.opaque+xml");
  EXPECT_EQ(EffectiveContentType(zip1, "xl/workbook.bin"), "application/vnd.ms-excel.sheet.binary.macroEnabled.main");
  EXPECT_FALSE(HasContentTypeOverride(zip1, "[Content_Types].xml"));
  EXPECT_EQ(RelationshipSignatures(zip1, "xl/_rels/workbook.bin.rels"), source_workbook_rels);
  EXPECT_EQ(RelationshipSignatures(zip1, "xl/worksheets/_rels/sheet1.bin.rels"), source_sheet_rels);

  auto second_or = io::xlsb::read_xlsb(SpanOf(save1_or.value()));
  ASSERT_TRUE(static_cast<bool>(second_or)) << second_or.error().message << " | " << second_or.error().context;
  EXPECT_EQ(second_or.value().dropped_part_count, 0U);
  auto save2_or = io::xlsb::write_xlsb(second_or.value().workbook);
  ASSERT_TRUE(static_cast<bool>(save2_or)) << save2_or.error().message << " | " << save2_or.error().context;
  io::ZipReader zip2;
  ASSERT_TRUE(static_cast<bool>(zip2.open(SpanOf(save2_or.value()))));
  EXPECT_EQ(InventoryNames(zip2), source_inventory);
  EXPECT_EQ(ReadPart(zip2, "xl/media/image1.PNG"), PngBytes());
  EXPECT_EQ(ReadPart(zip2, "xl/embeddings/oleObject1.bin"), OleBytes());
  EXPECT_EQ(ReadPart(zip2, "xl/customXml/item1.xml"), ToBytes("<custom>formulon</custom>"));
  EXPECT_EQ(EffectiveContentType(zip2, "xl/media/image1.PNG"), "image/png");
  EXPECT_EQ(EffectiveContentType(zip2, "xl/embeddings/oleObject1.bin"), "application/vnd.ms-office.vbaProject");
  EXPECT_EQ(EffectiveContentType(zip2, "xl/customXml/item1.xml"), "application/vnd.example.opaque+xml");
  EXPECT_EQ(EffectiveContentType(zip2, "xl/workbook.bin"), "application/vnd.ms-excel.sheet.binary.macroEnabled.main");
  EXPECT_FALSE(HasContentTypeOverride(zip2, "[Content_Types].xml"));
  EXPECT_EQ(RelationshipSignatures(zip2, "xl/_rels/workbook.bin.rels"), source_workbook_rels);
  EXPECT_EQ(RelationshipSignatures(zip2, "xl/worksheets/_rels/sheet1.bin.rels"), source_sheet_rels);

  auto third_or = io::xlsb::read_xlsb(SpanOf(save2_or.value()));
  ASSERT_TRUE(static_cast<bool>(third_or)) << third_or.error().message << " | " << third_or.error().context;
  EXPECT_EQ(third_or.value().dropped_part_count, 0U);
  for (const std::string_view path :
       {"xl/media/image1.PNG", "xl/embeddings/oleObject1.bin", "xl/customXml/item1.xml"}) {
    const io::PassthroughPart* part = FindPart(third_or.value().workbook, path);
    ASSERT_NE(part, nullptr) << path;
    EXPECT_TRUE(part->content_type.empty());
  }
}

TEST(XlsbDefaultPassthrough, RejectsConflictingDefaults) {
  auto result = io::xlsb::read_xlsb(SpanOf(DefaultPackage(/*conflicting_defaults=*/true)));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(XlsbDefaultPassthrough, CoalescesDuplicateOverrides) {
  auto result = io::xlsb::read_xlsb(
      SpanOf(DefaultPackage(/*conflicting_defaults=*/false, /*add_orphan=*/false, /*add_unsafe=*/false,
                            /*duplicate_override=*/true)));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;
  EXPECT_EQ(result.value().dropped_part_count, 0U);
}

TEST(XlsbDefaultPassthrough, RejectsConflictingOverrides) {
  auto result = io::xlsb::read_xlsb(
      SpanOf(DefaultPackage(/*conflicting_defaults=*/false, /*add_orphan=*/false, /*add_unsafe=*/false,
                            /*duplicate_override=*/false, /*conflicting_override=*/true)));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(XlsbDefaultPassthrough, UnsafeResidualPathIsRejected) {
  auto result = io::xlsb::read_xlsb(SpanOf(DefaultPackage(/*conflicting_defaults=*/false, /*add_orphan=*/false,
                                                          /*add_unsafe=*/true)));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipSlip);
}

TEST(XlsbDefaultPassthrough, DirectoryMarkerIsNotAnUnresolvedPart) {
  const Bytes archive = RewriteZip(DefaultPackage(), {}, {{"xl/media/", Bytes{}}});
  auto result = io::xlsb::read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;
  EXPECT_EQ(result.value().dropped_part_count, 0U);
}

TEST(XlsbDefaultPassthrough, UnresolvedResidualIsDroppedAndDoesNotRoundTrip) {
  const Bytes source = DefaultPackage(/*conflicting_defaults=*/false, /*add_orphan=*/true);
  auto load_or = io::xlsb::read_xlsb(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(load_or)) << load_or.error().message;
  EXPECT_EQ(load_or.value().dropped_part_count, 1U);
  EXPECT_EQ(FindPart(load_or.value().workbook, "orphan.unknown"), nullptr);
  auto save_or = io::xlsb::write_xlsb(load_or.value().workbook);
  ASSERT_TRUE(static_cast<bool>(save_or));
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  EXPECT_FALSE(zip.has_entry("orphan.unknown"));
}

TEST(XlsbDefaultPassthrough, RejectsDuplicateCriticalPart) {
  const Bytes archive = RewriteZip(BaseXlsb(), {}, {{"xl/workbook.bin", ToBytes("duplicate")}});
  auto result = io::xlsb::read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipCorrupt);
  EXPECT_NE(result.error().context.find("duplicate_entry=xl/workbook.bin"), std::string::npos);
}

TEST(XlsbDefaultPassthrough, RejectsDuplicateModelledPart) {
  const Bytes archive = RewriteZip(BaseXlsb(), {}, {{"xl/worksheets/sheet1.bin", ToBytes("duplicate")}});
  auto result = io::xlsb::read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipCorrupt);
  EXPECT_NE(result.error().context.find("duplicate_entry=xl/worksheets/sheet1.bin"), std::string::npos);
}

TEST(XlsbDefaultPassthrough, RejectsDuplicatePassthroughPart) {
  const Bytes archive = RewriteZip(DefaultPackage(), {}, {{"xl/media/image1.PNG", PngBytes()}});
  auto result = io::xlsb::read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipCorrupt);
  EXPECT_NE(result.error().context.find("duplicate_entry=xl/media/image1.PNG"), std::string::npos);
}

TEST(XlsbDefaultPassthrough, WriterDropsGeneratedPathCollision) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.set_default_content_types({io::DefaultContentType{"png", "image/png"}});
  wb.set_passthrough_parts({io::PassthroughPart{"xl/workbook.bin", "", ToBytes("stale")},
                            io::PassthroughPart{"xl/media/image1.png", "", ToBytes("image")},
                            io::PassthroughPart{"xl/media/image1.png", "", ToBytes("duplicate")}});
  auto save_or = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(save_or)) << save_or.error().message << " | " << save_or.error().context;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  EXPECT_TRUE(zip.has_entry("xl/workbook.bin"));
  EXPECT_EQ(ReadPart(zip, "xl/media/image1.png"), ToBytes("image"));
}

}  // namespace
}  // namespace formulon
