//
// Unit tests for the element-text vs. attribute-value XML escaping split.

#include "io/xml_escape.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/passthrough_part.h"
#include "io/tables_reader.h"
#include "io/unknown_relationship.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "workbook.h"

namespace formulon::io {
namespace {

/// Escapes `value` into an attribute, parses the resulting element, and
/// returns what an attribute reader sees. Attribute readers take the
/// parser's output verbatim, so this composition is the round trip the
/// escaper has to be the inverse of.
std::string AttrRoundTrip(std::string_view value) {
  std::string xml = "<r name=\"";
  AppendXmlAttrEscaped(xml, value);
  xml += "\"/>";

  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(xml.data(), xml.size())) << xml;
  return doc.child("r").attribute("name").value();
}

TEST(XmlEscape, ElementTextEscapesCriticalCharsAndControls) {
  // LF and TAB are legal XML characters that element text preserves, so
  // they go through verbatim. CR would be folded into LF by line-end
  // normalisation and U+0001 is illegal outright, so both take the OOXML
  // escape.
  std::string out;
  AppendXmlEscaped(out, "a&b<c>\"d'e\n\tf\r\x01");
  EXPECT_EQ(out, "a&amp;b&lt;c&gt;&quot;d&apos;e\n\tf_x000D__x0001_");
}

TEST(XmlEscape, AttrEscapesCriticalCharsAndWhitespaceControls) {
  // Attribute-value normalisation would turn a literal LF / TAB / CR into a
  // space, so each becomes a character reference the parser restores. An
  // XML-illegal control has no attribute representation at all and is
  // replaced, the same way an invalid UTF-8 byte is.
  std::string out;
  AppendXmlAttrEscaped(out, "a&b<c>\"d'e\n\tf\r\x01");
  EXPECT_EQ(out, "a&amp;b&lt;c&gt;&quot;d&apos;e&#10;&#9;f&#13;\xEF\xBF\xBD");
}

TEST(XmlEscape, OoxmlControlAndLiteralEscapeRoundTrip) {
  const std::string original = "a\r\x01_x000D_\xF0\x9F\x98\x80";
  std::string escaped;
  AppendXmlEscaped(escaped, original);
  EXPECT_EQ(escaped, "a_x000D__x0001__x005F_x000D_\xF0\x9F\x98\x80");

  std::string unescaped;
  AppendOoxmlTextUnescaped(unescaped, escaped);
  EXPECT_EQ(unescaped, original);
}

TEST(XmlEscape, ReplacesInvalidUtf8BeforeWritingXml) {
  std::string escaped;
  AppendXmlEscaped(escaped, "a\xC3\x28\x80");
  EXPECT_EQ(escaped, "a\xEF\xBF\xBD(\xEF\xBF\xBD");

  pugi::xml_document doc;
  const std::string xml = "<t>" + escaped + "</t>";
  ASSERT_TRUE(doc.load_buffer(xml.data(), xml.size()));
  EXPECT_EQ(doc.child("t").text().get(), escaped);
}

TEST(XmlEscape, RichTextReaderDecodesOoxmlEscapes) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<si><t>a_x000D__x005F_x000D_</t></si>"));

  std::string text;
  EXPECT_EQ(append_rich_text(doc.child("si"), text), 1U);
  EXPECT_EQ(text, "a\r_x000D_");
}

TEST(XmlEscape, AttrEscapedNewlineSurvivesPugixmlRoundtrip) {
  // Attribute-value normalisation (mandatory for any conforming XML
  // parser, including pugixml) replaces a *literal* newline inside an
  // attribute value with a space. Only a character reference for it
  // survives intact -- this is the whole point of AppendXmlAttrEscaped.
  std::string attr_value;
  AppendXmlAttrEscaped(attr_value, "line one\nline two\ttabbed");

  std::string xml = "<r name=\"";
  xml += attr_value;
  xml += "\"/>";

  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml.c_str()));
  const std::string_view roundtripped = doc.child("r").attribute("name").value();
  EXPECT_EQ(roundtripped, "line one\nline two\ttabbed");
}

TEST(XmlEscape, PlainEscapeNewlineDoesNotSurviveAttributeRoundtrip) {
  // Contrast case: using the element-text escaper for an attribute value
  // loses the newline, because the literal character gets normalised
  // away by the XML parser instead of being preserved as `&#10;`.
  std::string attr_value;
  AppendXmlEscaped(attr_value, "line one\nline two");

  std::string xml = "<r name=\"";
  xml += attr_value;
  xml += "\"/>";

  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml.c_str()));
  const std::string_view roundtripped = doc.child("r").attribute("name").value();
  EXPECT_NE(roundtripped, "line one\nline two");
}

TEST(XmlEscape, AttrEscapeNonAsciiPassesThroughVerbatim) {
  std::string out;
  AppendXmlAttrEscaped(out, "\xE6\x97\xA5\xE6\x9C\xAC");  // "日本" in UTF-8.
  EXPECT_EQ(out, "\xE6\x97\xA5\xE6\x9C\xAC");
}

TEST(XmlEscape, AttrEscapeLeavesOoxmlSpellingLiteral) {
  // An `_xHHHH_`-shaped run is ordinary text in an attribute: nothing on the
  // read path decodes it, so doubling the underscore would leave `_x005F_`
  // in the value on reload and add another six bytes on every later save.
  std::string out;
  AppendXmlAttrEscaped(out, "Data_x0041_Q1");
  EXPECT_EQ(out, "Data_x0041_Q1");
}

TEST(XmlEscape, AttrEscapeIsInverseOfRawAttributeRead) {
  // Every string an XML 1.0 attribute can carry survives escape -> parse ->
  // raw read unchanged.
  const std::string_view cases[] = {
      "Data_x0041_Q1",
      "_x005F_",
      "_x000D__x0001_",
      "a&b<c>\"d'e",
      "tab\there\nnewline\rcarriage",
      "\xE6\x97\xA5\xE6\x9C\xAC",
      "\xF0\x9F\x98\x80",
      "  leading and trailing  ",
      "",
  };
  for (const std::string_view value : cases) {
    EXPECT_EQ(AttrRoundTrip(value), value) << "value=" << value;
  }
}

TEST(XmlEscape, AttrEscapeIsFixedPointForUnrepresentableBytes) {
  // A C0 control and an invalid UTF-8 byte have no attribute spelling, so
  // the first escape replaces them. The replacement must then be stable:
  // repeated save/load cycles may not keep rewriting the value.
  const std::string_view cases[] = {"a\x01\x0Bb", "a\xC3\x28\x80"};
  for (const std::string_view value : cases) {
    const std::string once = AttrRoundTrip(value);
    EXPECT_NE(once, value) << "value=" << value;
    EXPECT_EQ(AttrRoundTrip(once), once) << "value=" << value;
  }
}

TEST(XmlEscape, SheetNameWithOoxmlSpellingSurvivesThreeSaveLoadCycles) {
  // A sheet name that merely looks like an OOXML escape is a legal name.
  // Escaping the underscore would add six units per save, so this 21-unit
  // name would pass the 31-unit sheet-name limit within three cycles and
  // make the reader reject Formulon's own output as corrupt.
  const std::string kSheetName = "Sales_x0041_Q1_Report";

  Workbook current = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(current.rename_sheet(0U, kSheetName)));

  for (int cycle = 0; cycle < 3; ++cycle) {
    auto written = write_ooxml(current);
    ASSERT_TRUE(static_cast<bool>(written)) << "cycle " << cycle << ": " << written.error().message;
    const std::vector<std::uint8_t> package = std::move(written.value());

    auto read = read_ooxml(ByteSpan{package.data(), package.size()});
    ASSERT_TRUE(static_cast<bool>(read)) << "cycle " << cycle << ": " << read.error().message;
    current = std::move(read.value().workbook);
    ASSERT_EQ(current.sheet_count(), 1U);
    EXPECT_EQ(current.sheet(0U).name(), kSheetName) << "cycle " << cycle;
  }
}

TEST(XmlEscape, RelationshipTargetGetsNoOoxmlEscaping) {
  // `Relationship/@Target` is an `xsd:anyURI`. The rels reader takes it
  // straight off the parser, so an `_xHHHH_`-shaped run in a URL has to
  // reach the part verbatim: escaping the underscore would both corrupt the
  // URI and grow it by six bytes on every save.
  const std::string kTarget = "https://example.test/report_x0041_Q1?tab=_x000D_";

  Workbook current = Workbook::create();
  Hyperlink link;
  link.row = 0U;
  link.col = 0U;
  link.last_row = 0U;
  link.last_col = 0U;
  link.target = kTarget;
  current.sheet(0U).mutable_hyperlinks().push_back(link);

  for (int cycle = 0; cycle < 3; ++cycle) {
    auto written = write_ooxml(current);
    ASSERT_TRUE(static_cast<bool>(written)) << "cycle " << cycle << ": " << written.error().message;
    const std::vector<std::uint8_t> package = std::move(written.value());

    // The bytes actually written to the rels part carry the target as-is.
    ZipReader zip;
    ASSERT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
    auto rels = zip.read_entry("xl/worksheets/_rels/sheet1.xml.rels");
    ASSERT_TRUE(static_cast<bool>(rels)) << "cycle " << cycle << ": " << rels.error().message;
    const std::string rels_text(rels.value().begin(), rels.value().end());
    EXPECT_NE(rels_text.find("Target=\"" + kTarget + "\""), std::string::npos) << rels_text;
    EXPECT_EQ(rels_text.find("_x005F_"), std::string::npos) << rels_text;

    auto read = read_ooxml(ByteSpan{package.data(), package.size()});
    ASSERT_TRUE(static_cast<bool>(read)) << "cycle " << cycle << ": " << read.error().message;
    current = std::move(read.value().workbook);
    ASSERT_EQ(current.sheet(0U).hyperlinks().size(), 1U);
    EXPECT_EQ(current.sheet(0U).hyperlinks()[0].target, kTarget) << "cycle " << cycle;
  }
}

TEST(XmlEscape, XlsbPackageEnvelopeAttributesSurviveThreeRoundTrips) {
  // The `.xlsb` package envelope is XML: the writer hand-builds
  // `[Content_Types].xml` and the rels parts. Their attribute values must
  // follow the attribute rule like every other writer's, or a part path
  // gains `_x005F_` per save until the relationship no longer names a part
  // that exists -- at which point the writer drops it rather than growing
  // the string further, so the loss is silent and permanent.
  const std::string kPartPath = "xl/media/logo_x0041_.png";
  const std::string kRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  const std::string kHyperlinkTarget = "https://example.test/a_x0041_b?q=x\ty";

  Workbook current = Workbook::create();
  {
    PassthroughPart part;
    part.path = kPartPath;
    part.content_type = "image/png";
    part.bytes = {0x89, 0x50, 0x4E, 0x47};
    current.set_passthrough_parts({part});

    UnknownRelationship rel;
    rel.id = "rId100";
    rel.type = kRelType;
    rel.target = kPartPath;
    current.set_unknown_workbook_rels({rel});

    Hyperlink link;
    link.target = kHyperlinkTarget;
    current.sheet(0U).mutable_hyperlinks().push_back(link);
  }

  for (int cycle = 0; cycle < 3; ++cycle) {
    auto written = xlsb::write_xlsb_with_result(current);
    ASSERT_TRUE(static_cast<bool>(written)) << "cycle " << cycle << ": " << written.error().message;
    // A relationship whose target no longer names a package part is dropped
    // outright, so this counter -- not a byte comparison -- is what a broken
    // escaper trips first.
    EXPECT_EQ(written.value().diagnostics.dropped_relationship_count, 0U) << "cycle " << cycle;
    const std::vector<std::uint8_t> package = std::move(written.value().bytes);

    ZipReader zip;
    ASSERT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
    auto types = zip.read_entry("[Content_Types].xml");
    ASSERT_TRUE(static_cast<bool>(types)) << "cycle " << cycle << ": " << types.error().message;
    const std::string types_text(types.value().begin(), types.value().end());
    // The Override has to name the part under the path the archive actually
    // stores it at, or the part silently falls back to its Default type.
    EXPECT_NE(types_text.find("PartName=\"/" + kPartPath + "\""), std::string::npos) << types_text;
    EXPECT_EQ(types_text.find("_x005F_"), std::string::npos) << types_text;

    auto read = xlsb::read_xlsb(ByteSpan{package.data(), package.size()});
    ASSERT_TRUE(static_cast<bool>(read)) << "cycle " << cycle << ": " << read.error().message;
    current = std::move(read.value().workbook);

    // The writer also generates `xl/styles.bin`, which the reader captures
    // as a passthrough part of its own, so look the media part up by path.
    std::size_t media_parts = 0;
    for (const PassthroughPart& part : current.passthrough_parts()) {
      media_parts += static_cast<std::size_t>(part.path == kPartPath);
    }
    EXPECT_EQ(media_parts, 1U) << "cycle " << cycle;
    ASSERT_EQ(current.unknown_workbook_rels().size(), 1U) << "cycle " << cycle;
    EXPECT_EQ(current.unknown_workbook_rels()[0].target, kPartPath) << "cycle " << cycle;
    ASSERT_EQ(current.sheet(0U).hyperlinks().size(), 1U) << "cycle " << cycle;
    EXPECT_EQ(current.sheet(0U).hyperlinks()[0].target, kHyperlinkTarget) << "cycle " << cycle;
  }
}

TEST(XmlEscape, CapturedTableAttributeKeepsWhitespaceControls) {
  // Attributes the tables reader does not model are retained verbatim and
  // spliced back into the tag. They go through the shared attribute rule,
  // so TAB / LF / CR stay character references instead of collapsing to
  // spaces under attribute-value normalisation.
  std::string xml(kXmlDecl);
  xml.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"T1\" "
      "displayName=\"T1\" ref=\"A1:B2\" note=\"a&#9;b&#10;c&#13;d\" tag=\"lit_x0041_eral\" ctl=\"p&#1;q\">");
  xml.append("<tableColumns count=\"1\"><tableColumn id=\"1\" name=\"A\"/></tableColumns>");
  xml.append("</table>");

  auto table_or = read_table(std::vector<std::uint8_t>(xml.begin(), xml.end()), 0U);
  ASSERT_TRUE(static_cast<bool>(table_or)) << table_or.error().message;
  const std::string& extras = table_or.value().root_extra_attrs;
  EXPECT_NE(extras.find(" note=\"a&#9;b&#10;c&#13;d\""), std::string::npos) << extras;
  EXPECT_NE(extras.find(" tag=\"lit_x0041_eral\""), std::string::npos) << extras;
  // `&#1;` is not a legal XML 1.0 character reference, but the parser hands
  // the byte over anyway. Re-emitting it raw would produce a part no reader
  // can reopen, so it is replaced rather than carried through.
  EXPECT_NE(extras.find(" ctl=\"p\xEF\xBF\xBDq\""), std::string::npos) << extras;
  EXPECT_EQ(extras.find('\x01'), std::string::npos) << extras;

  // Re-emitting the captured run and reading it again is a fixed point.
  std::string rewritten(kXmlDecl);
  rewritten.append(
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"T1\" "
      "displayName=\"T1\" ref=\"A1:B2\"");
  rewritten.append(extras);
  rewritten.append("><tableColumns count=\"1\"><tableColumn id=\"1\" name=\"A\"/></tableColumns></table>");

  auto reread = read_table(std::vector<std::uint8_t>(rewritten.begin(), rewritten.end()), 0U);
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  EXPECT_EQ(reread.value().root_extra_attrs, extras);
}

}  // namespace
}  // namespace formulon::io
