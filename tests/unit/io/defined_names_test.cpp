//
// Unit tests for `formulon::io::read_defined_names`. The reader takes a
// pre-parsed `xl/workbook.xml` document, so each test builds the XML
// from a string literal, parses it via pugixml, then calls the reader.

#include "io/defined_names.h"

#include <cstdint>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "io/defined_names_internal.h"
#include "pugixml.hpp"
#include "utils/error.h"

namespace formulon {
namespace io {
namespace {

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

/// Loads `xml` into a pugixml document. The document is returned by
/// reference via the out-parameter so the test can keep it alive while
/// inspecting the reader's `string_view`-free output (the reader copies
/// every payload into the result, so doc lifetime here only needs to
/// span the call).
pugi::xml_parse_result LoadDoc(pugi::xml_document& doc, std::string_view xml) {
  return doc.load_buffer(xml.data(), xml.size(), pugi::parse_default, pugi::encoding_utf8);
}

// ---------------------------------------------------------------------------
// Empty / absent cases
// ---------------------------------------------------------------------------

TEST(DefinedNamesReader, NoDefinedNamesBlockYieldsEmptyVector) {
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_TRUE(static_cast<bool>(names_or)) << "read failed: " << names_or.error().message;
  EXPECT_TRUE(names_or.value().empty());
}

TEST(DefinedNamesReader, EmptyDefinedNamesBlockYieldsEmptyVector) {
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames></definedNames>");
  xml.append("  <sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_TRUE(static_cast<bool>(names_or));
  EXPECT_TRUE(names_or.value().empty());
}

// ---------------------------------------------------------------------------
// Happy paths
// ---------------------------------------------------------------------------

TEST(DefinedNamesReader, SingleWorkbookScopedName) {
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames>");
  xml.append("    <definedName name=\"Foo\">Sheet1!$A$1:$A$10</definedName>");
  xml.append("  </definedNames>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_TRUE(static_cast<bool>(names_or));
  ASSERT_EQ(names_or.value().size(), 1U);
  const DefinedName& dn = names_or.value()[0];
  EXPECT_EQ(dn.name, "Foo");
  EXPECT_EQ(dn.formula, "Sheet1!$A$1:$A$10");
  EXPECT_EQ(dn.local_sheet_id, -1);
  EXPECT_FALSE(dn.hidden);
  EXPECT_TRUE(dn.comment.empty());
}

TEST(DefinedNamesReader, SheetScopedNameCapturesLocalSheetId) {
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames>");
  xml.append("    <definedName name=\"Foo\" localSheetId=\"2\">Sheet3!$B$1</definedName>");
  xml.append("  </definedNames>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_TRUE(static_cast<bool>(names_or));
  ASSERT_EQ(names_or.value().size(), 1U);
  EXPECT_EQ(names_or.value()[0].local_sheet_id, 2);
}

TEST(DefinedNamesReader, HiddenAttributeAcceptsBothLexicalForms) {
  // Excel-style writers emit `hidden="1"`; pre-2003-style emit
  // `hidden="true"`. Both must yield true; absent / "false" / "0"
  // must yield false.
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames>");
  xml.append("    <definedName name=\"H1\" hidden=\"1\">Sheet1!$A$1</definedName>");
  xml.append("    <definedName name=\"H2\" hidden=\"true\">Sheet1!$A$2</definedName>");
  xml.append("    <definedName name=\"H3\" hidden=\"TRUE\">Sheet1!$A$3</definedName>");
  xml.append("    <definedName name=\"H4\" hidden=\"false\">Sheet1!$A$4</definedName>");
  xml.append("    <definedName name=\"H5\" hidden=\"0\">Sheet1!$A$5</definedName>");
  xml.append("    <definedName name=\"H6\">Sheet1!$A$6</definedName>");
  xml.append("  </definedNames>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_TRUE(static_cast<bool>(names_or));
  ASSERT_EQ(names_or.value().size(), 6U);
  EXPECT_TRUE(names_or.value()[0].hidden);
  EXPECT_TRUE(names_or.value()[1].hidden);
  EXPECT_TRUE(names_or.value()[2].hidden);
  EXPECT_FALSE(names_or.value()[3].hidden);
  EXPECT_FALSE(names_or.value()[4].hidden);
  EXPECT_FALSE(names_or.value()[5].hidden);
}

TEST(DefinedNamesReader, MultipleNamesPreserveDeclarationOrder) {
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames>");
  xml.append("    <definedName name=\"alpha\">Sheet1!$A$1</definedName>");
  xml.append("    <definedName name=\"beta\">Sheet1!$B$1</definedName>");
  xml.append("    <definedName name=\"gamma\">Sheet1!$C$1</definedName>");
  xml.append("  </definedNames>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_TRUE(static_cast<bool>(names_or));
  ASSERT_EQ(names_or.value().size(), 3U);
  EXPECT_EQ(names_or.value()[0].name, "alpha");
  EXPECT_EQ(names_or.value()[1].name, "beta");
  EXPECT_EQ(names_or.value()[2].name, "gamma");
}

TEST(DefinedNamesReader, FormulaWhitespaceIsTrimmed) {
  // pugixml exposes the element's child text verbatim, so a writer
  // that pretty-prints `<definedName>\n  =SUM(A1:A10)\n</definedName>`
  // hands us padded text. The reader must strip the framing whitespace
  // but preserve interior whitespace (here: the space inside the SUM
  // arglist).
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames>");
  xml.append("    <definedName name=\"Padded\">\n      =SUM(A1, A2, A3)\n    </definedName>");
  xml.append("  </definedNames>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_TRUE(static_cast<bool>(names_or));
  ASSERT_EQ(names_or.value().size(), 1U);
  EXPECT_EQ(names_or.value()[0].formula, "=SUM(A1, A2, A3)");
}

TEST(DefinedNamesReader, CommentAttributeRoundTrips) {
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames>");
  xml.append("    <definedName name=\"Sales\" comment=\"Quarterly revenue range\">Sheet1!$A$1:$D$1</definedName>");
  xml.append("  </definedNames>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_TRUE(static_cast<bool>(names_or));
  ASSERT_EQ(names_or.value().size(), 1U);
  EXPECT_EQ(names_or.value()[0].comment, "Quarterly revenue range");
}

TEST(DefinedNamesReader, JapaneseUnicodeNamePreserved) {
  // "日本語名" — capture the UTF-8 bytes verbatim. Pugixml parses
  // UTF-8 input and we hand bytes back unchanged.
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames>");
  xml.append(
      "    <definedName "
      "name=\"\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E\xE5\x90\x8D\">Sheet1!$A$1</definedName>");
  xml.append("  </definedNames>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_TRUE(static_cast<bool>(names_or));
  ASSERT_EQ(names_or.value().size(), 1U);
  EXPECT_EQ(names_or.value()[0].name, "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E\xE5\x90\x8D");
}

// ---------------------------------------------------------------------------
// Error paths
// ---------------------------------------------------------------------------

TEST(DefinedNamesReader, MissingNameAttributeIsCorruption) {
  // `<definedName>` without `name=` is malformed. Excel rejects the
  // workbook outright.
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames>");
  xml.append("    <definedName>Sheet1!$A$1</definedName>");
  xml.append("  </definedNames>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_FALSE(static_cast<bool>(names_or));
  EXPECT_EQ(names_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(DefinedNamesReader, EmptyNameAttributeIsCorruption) {
  std::string xml(kXmlDecl);
  xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("  <definedNames>");
  xml.append("    <definedName name=\"\">Sheet1!$A$1</definedName>");
  xml.append("  </definedNames>");
  xml.append("</workbook>");

  pugi::xml_document doc;
  ASSERT_TRUE(LoadDoc(doc, xml));

  auto names_or = read_defined_names(doc);
  ASSERT_FALSE(static_cast<bool>(names_or));
  EXPECT_EQ(names_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

}  // namespace
}  // namespace io
}  // namespace formulon
