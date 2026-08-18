//
// Unit tests for the typed node-attribute accessors in `io/xml_utils.h`.
// The legacy `parse_xml_*_attr` helpers (which take a `pugi::xml_attribute`
// directly) are exercised indirectly by the reader unit tests; this file
// targets the newer `attr_*(node, name, def)` API the readers are
// migrating onto.

#include "io/xml_utils.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

pugi::xml_document Load(const char* xml) {
  pugi::xml_document doc;
  pugi::xml_parse_result rc = doc.load_string(xml);
  EXPECT_TRUE(rc) << rc.description();
  return doc;
}

TEST(XmlUtilsAttr, StrPresentReturnsAttributeValue) {
  pugi::xml_document doc = Load(R"(<r name="hello"/>)");
  EXPECT_EQ(attr_str(doc.child("r"), "name"), std::string_view("hello"));
}

TEST(XmlUtilsAttr, StrMissingReturnsDefault) {
  pugi::xml_document doc = Load(R"(<r/>)");
  EXPECT_EQ(attr_str(doc.child("r"), "name"), std::string_view());
  EXPECT_EQ(attr_str(doc.child("r"), "name", "fallback"), std::string_view("fallback"));
}

TEST(XmlUtilsAttr, StrEmptyValueReturnsDefault) {
  // An empty `name=""` attribute is treated as absent: callers that
  // care about presence-without-value should still go through
  // `node.attribute(name)` directly.
  pugi::xml_document doc = Load(R"(<r name=""/>)");
  EXPECT_EQ(attr_str(doc.child("r"), "name", "fallback"), std::string_view("fallback"));
}

TEST(XmlUtilsAttr, U32Roundtrip) {
  pugi::xml_document doc = Load(R"(<r id="42" zero="0" big="4294967290"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_EQ(attr_u32(n, "id"), 42U);
  EXPECT_EQ(attr_u32(n, "zero"), 0U);
  EXPECT_EQ(attr_u32(n, "big"), 4294967290U);
  EXPECT_EQ(attr_u32(n, "missing"), 0U);
  EXPECT_EQ(attr_u32(n, "missing", 99U), 99U);
}

TEST(XmlUtilsAttr, U32MalformedFallsBackToDefault) {
  pugi::xml_document doc = Load(R"(<r empty="" junk="abc" neg="-1"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_EQ(attr_u32(n, "empty", 7U), 7U);
  EXPECT_EQ(attr_u32(n, "junk", 7U), 7U);
  // pugi's as_uint() rejects negative input by returning the supplied
  // default — this matches the legacy `parse_xml_u32_attr` behaviour.
  EXPECT_EQ(attr_u32(n, "neg", 7U), 7U);
}

TEST(XmlUtilsAttr, I32Roundtrip) {
  pugi::xml_document doc = Load(R"(<r a="42" b="-1" c="0" d="-2147483648"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_EQ(attr_i32(n, "a"), 42);
  EXPECT_EQ(attr_i32(n, "b"), -1);
  EXPECT_EQ(attr_i32(n, "c"), 0);
  EXPECT_EQ(attr_i32(n, "d"), INT32_MIN);
  EXPECT_EQ(attr_i32(n, "missing", -42), -42);
}

TEST(XmlUtilsAttr, F64Roundtrip) {
  pugi::xml_document doc = Load(R"(<r a="3.14" b="-2.5" c="0" d="1e3"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_DOUBLE_EQ(attr_f64(n, "a"), 3.14);
  EXPECT_DOUBLE_EQ(attr_f64(n, "b"), -2.5);
  EXPECT_DOUBLE_EQ(attr_f64(n, "c"), 0.0);
  EXPECT_DOUBLE_EQ(attr_f64(n, "d"), 1000.0);
  EXPECT_DOUBLE_EQ(attr_f64(n, "missing", 1.5), 1.5);
}

TEST(XmlUtilsAttr, BoolOneAndZero) {
  pugi::xml_document doc = Load(R"(<r a="1" b="0"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_TRUE(attr_bool(n, "a"));
  EXPECT_FALSE(attr_bool(n, "b"));
}

TEST(XmlUtilsAttr, BoolTrueFalseCaseInsensitive) {
  pugi::xml_document doc = Load(R"(<r a="true" b="True" c="TRUE" d="false" e="False" f="FALSE"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_TRUE(attr_bool(n, "a"));
  EXPECT_TRUE(attr_bool(n, "b"));
  EXPECT_TRUE(attr_bool(n, "c"));
  EXPECT_FALSE(attr_bool(n, "d"));
  EXPECT_FALSE(attr_bool(n, "e"));
  EXPECT_FALSE(attr_bool(n, "f"));
}

TEST(XmlUtilsAttr, BoolMissingUsesDefault) {
  pugi::xml_document doc = Load(R"(<r/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_FALSE(attr_bool(n, "absent"));
  EXPECT_TRUE(attr_bool(n, "absent", true));
}

TEST(XmlUtilsAttr, BoolUnknownTokensMapToFalse) {
  // Anything outside {1, 0, true(any case), false(any case)} maps to
  // false. Matches the legacy `parse_xml_bool_attr` lenient lexicon.
  pugi::xml_document doc = Load(R"(<r a="yes" b="2" c="trues" d="tru"/>)");
  pugi::xml_node n = doc.child("r");
  EXPECT_FALSE(attr_bool(n, "a"));
  EXPECT_FALSE(attr_bool(n, "b"));
  EXPECT_FALSE(attr_bool(n, "c"));
  EXPECT_FALSE(attr_bool(n, "d"));
}

TEST(XmlUtilsAttr, EmptyNodeReturnsDefaults) {
  pugi::xml_node empty;
  EXPECT_EQ(attr_str(empty, "x", "fb"), std::string_view("fb"));
  EXPECT_EQ(attr_u32(empty, "x", 11U), 11U);
  EXPECT_EQ(attr_i32(empty, "x", -3), -3);
  EXPECT_DOUBLE_EQ(attr_f64(empty, "x", 2.5), 2.5);
  EXPECT_TRUE(attr_bool(empty, "x", true));
  EXPECT_FALSE(attr_bool(empty, "x", false));
}

TEST(XmlUtilsAttr, RootExtraAttributesEscapeValuesForReEmission) {
  pugi::xml_document doc =
      Load(R"(<worksheet xmlns="urn:main" xmlns:r="urn:rel" custom="one &amp; two &quot;three&quot; &lt;four&gt;"/>)");

  EXPECT_EQ(capture_root_extra_ns_attrs(doc.document_element()),
            " custom=\"one &amp; two &quot;three&quot; &lt;four&gt;\"");
}

// ---------------------------------------------------------------------------
// Element-level passthrough
// ---------------------------------------------------------------------------
//
// A part the reader consumes drops out of the unknown-part sweep, so anything
// in it the model does not represent survives a save only if it was retained
// verbatim. These pin the retention format the writers re-emit.

TEST(XmlUtilsRawXml, SerialisesElementAndSubtreeWithoutIndentation) {
  pugi::xml_document doc = Load(R"(<ext uri="{X}">
      <inner a="1"><leaf/></inner>
    </ext>)");
  // No indentation and no added whitespace: the blob is re-emitted inside a
  // generated part, where pretty-printing would change the bytes Excel sees.
  EXPECT_EQ(raw_xml(doc.child("ext")), R"(<ext uri="{X}"><inner a="1"><leaf/></inner></ext>)");
}

TEST(XmlUtilsRawXml, EscapesMarkupInTextAndAttributes) {
  pugi::xml_document doc = Load(R"(<n v="a &amp; b">x &lt; y</n>)");
  EXPECT_EQ(raw_xml(doc.child("n")), R"(<n v="a &amp; b">x &lt; y</n>)");
}

TEST(XmlUtilsRawXml, AppendVariantDoesNotClearTheTarget) {
  pugi::xml_document doc = Load(R"(<a/>)");
  std::string out = "<pre/>";
  append_raw_xml(out, doc.child("a"));
  EXPECT_EQ(out, "<pre/><a/>");
}

TEST(XmlUtilsRawXml, EmptyNodeAppendsNothing) {
  pugi::xml_node empty;
  std::string out;
  append_raw_xml(out, empty);
  EXPECT_TRUE(out.empty());
}

TEST(XmlUtilsUnknownChildren, RetainsUnmodelledChildrenInDocumentOrder) {
  pugi::xml_document doc = Load(R"(<ws><known/><alpha a="1"/><other/><beta/></ws>)");
  std::string out;
  capture_unknown_children(doc.child("ws"), {"known", "other"}, out);
  // Document order matters: OOXML content models are ordered sequences, so
  // the blob is only valid re-emitted at the position it was read from.
  EXPECT_EQ(out, R"(<alpha a="1"/><beta/>)");
}

TEST(XmlUtilsUnknownChildren, SkipsTextCommentAndProcessingInstructionChildren) {
  pugi::xml_document doc = Load("<ws>text<!--c--><?pi?><alpha/></ws>");
  std::string out;
  capture_unknown_children(doc.child("ws"), {}, out);
  EXPECT_EQ(out, "<alpha/>");
}

TEST(XmlUtilsUnknownChildren, EverythingKnownProducesNothing) {
  pugi::xml_document doc = Load(R"(<ws><a/><b/></ws>)");
  std::string out;
  capture_unknown_children(doc.child("ws"), {"a", "b"}, out);
  EXPECT_TRUE(out.empty());
}

TEST(XmlUtilsUnknownChildren, VectorOverloadKeepsPerChildBoundaries) {
  pugi::xml_document doc = Load(R"(<styleSheet><fonts/><alpha/><beta><g/></beta></styleSheet>)");
  std::vector<std::string> out;
  capture_unknown_children(doc.child("styleSheet"), {"fonts"}, out);
  ASSERT_EQ(out.size(), 2U);
  EXPECT_EQ(out[0], "<alpha/>");
  EXPECT_EQ(out[1], "<beta><g/></beta>");
}

TEST(XmlUtilsUnknownChildren, VectorOverloadAppendsToExistingContent) {
  pugi::xml_document doc = Load(R"(<ws><alpha/></ws>)");
  std::vector<std::string> out{"<kept/>"};
  capture_unknown_children(doc.child("ws"), {}, out);
  ASSERT_EQ(out.size(), 2U);
  EXPECT_EQ(out[0], "<kept/>");
  EXPECT_EQ(out[1], "<alpha/>");
}

// ---------------------------------------------------------------------------
// Part loaders. The copying and in-place variants must be
// indistinguishable in what they produce — same tree, same whitespace
// policy, same error envelope — and differ only in whether the DOM owns
// its text or aliases the caller's buffer.
// ---------------------------------------------------------------------------

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

constexpr std::string_view kSamplePart =
    R"(<sst count="2"><si><t>alpha</t></si><si><t> </t></si><si><t>&amp;&lt;</t></si></sst>)";

TEST(XmlUtilsLoadPart, InPlaceProducesTheSameTreeAsTheCopyingLoader) {
  std::vector<std::uint8_t> copied = Bytes(kSamplePart);
  std::vector<std::uint8_t> aliased = Bytes(kSamplePart);

  pugi::xml_document copy_doc;
  ASSERT_TRUE(static_cast<bool>(load_xml_buffer(copy_doc, copied, "test", "part.xml")));
  pugi::xml_document inplace_doc;
  ASSERT_TRUE(static_cast<bool>(load_xml_buffer_inplace(inplace_doc, aliased, "test", "part.xml")));

  auto texts = [](const pugi::xml_document& doc) {
    std::vector<std::string> out;
    for (pugi::xml_node si = doc.child("sst").child("si"); si; si = si.next_sibling("si")) {
      out.emplace_back(si.child("t").text().get());
    }
    return out;
  };
  const std::vector<std::string> expected{"alpha", " ", "&<"};
  EXPECT_EQ(texts(copy_doc), expected);
  EXPECT_EQ(texts(inplace_doc), expected);
  EXPECT_EQ(attr_u32(copy_doc.child("sst"), "count", 0U), attr_u32(inplace_doc.child("sst"), "count", 0U));
}

TEST(XmlUtilsLoadPart, InPlaceTextAliasesTheCallerBuffer) {
  // The whole point of the in-place loader: the DOM's payloads live in
  // the caller's storage rather than in a second copy inside pugixml.
  // Any node text must therefore fall inside `bytes`, and the copying
  // loader's must not.
  std::vector<std::uint8_t> aliased = Bytes(kSamplePart);
  const char* const begin = reinterpret_cast<const char*>(aliased.data());
  const char* const end = begin + aliased.size();

  pugi::xml_document inplace_doc;
  ASSERT_TRUE(static_cast<bool>(load_xml_buffer_inplace(inplace_doc, aliased, "test", "part.xml")));
  const char* const inplace_text = inplace_doc.child("sst").child("si").child("t").text().get();
  EXPECT_GE(inplace_text, begin);
  EXPECT_LT(inplace_text, end);

  std::vector<std::uint8_t> copied = Bytes(kSamplePart);
  const char* const copy_begin = reinterpret_cast<const char*>(copied.data());
  pugi::xml_document copy_doc;
  ASSERT_TRUE(static_cast<bool>(load_xml_buffer(copy_doc, copied, "test", "part.xml")));
  const char* const copy_text = copy_doc.child("sst").child("si").child("t").text().get();
  EXPECT_TRUE(copy_text < copy_begin || copy_text >= copy_begin + copied.size());
}

TEST(XmlUtilsLoadPart, BothLoadersKeepWhitespaceOnlyLeafText) {
  // `parse_ws_pcdata_single` is what makes `<t> </t>` read back as a
  // space instead of an empty string; the in-place path must not lose it.
  std::vector<std::uint8_t> bytes = Bytes(R"(<si><t> </t></si>)");
  pugi::xml_document doc;
  ASSERT_TRUE(static_cast<bool>(load_xml_buffer_inplace(doc, bytes, "test", "part.xml")));
  EXPECT_EQ(std::string_view(doc.child("si").child("t").text().get()), std::string_view(" "));
}

TEST(XmlUtilsLoadPart, BothLoadersReportTheSameParseFailureEnvelope) {
  std::vector<std::uint8_t> copied = Bytes("<sst><si></sst>");
  std::vector<std::uint8_t> aliased = Bytes("<sst><si></sst>");

  pugi::xml_document copy_doc;
  auto copy_status = load_xml_buffer(copy_doc, copied, "sst_reader", "sharedStrings.xml");
  pugi::xml_document inplace_doc;
  auto inplace_status = load_xml_buffer_inplace(inplace_doc, aliased, "sst_reader", "sharedStrings.xml");

  ASSERT_FALSE(static_cast<bool>(copy_status));
  ASSERT_FALSE(static_cast<bool>(inplace_status));
  EXPECT_EQ(copy_status.error().code, FormulonErrorCode::kIoXmlParse);
  EXPECT_EQ(inplace_status.error().code, copy_status.error().code);
  EXPECT_EQ(inplace_status.error().message, copy_status.error().message);
  EXPECT_EQ(inplace_status.error().context, copy_status.error().context);
}

// ---------------------------------------------------------------------------
// `capture_unknown_attrs` and namespace bindings
//
// Excel writes `mc:Ignorable` and `xr:uid` on the pivot parts, and the
// model represents neither, so they travel as captured attributes. Re-
// emitting one without the binding for its prefix does not merely lose
// metadata: it produces XML no parser will accept, which is why the
// binding travels with the attribute rather than being left to the writer.
// ---------------------------------------------------------------------------

TEST(XmlUtilsCaptureAttrs, PrefixedAttributeCarriesItsNamespaceBinding) {
  pugi::xml_document doc = Load(
      "<root xmlns=\"urn:main\" xmlns:mc=\"urn:mce\" xmlns:xr=\"urn:rev\""
      " mc:Ignorable=\"xr\" xr:uid=\"{ABC}\" recordCount=\"4\"/>");

  std::vector<std::pair<std::string, std::string>> captured;
  capture_unknown_attrs(doc.child("root"), {"recordCount"}, captured);

  std::string rendered;
  append_raw_attrs(rendered, captured);
  EXPECT_NE(rendered.find("xmlns:mc=\"urn:mce\""), std::string::npos) << rendered;
  EXPECT_NE(rendered.find("xmlns:xr=\"urn:rev\""), std::string::npos) << rendered;
  EXPECT_NE(rendered.find("mc:Ignorable=\"xr\""), std::string::npos) << rendered;
  EXPECT_NE(rendered.find("xr:uid=\"{ABC}\""), std::string::npos) << rendered;
  EXPECT_EQ(rendered.find("recordCount"), std::string::npos) << rendered;

  // The real test: what the writer emits has to parse.
  pugi::xml_document reparsed;
  const std::string element = "<e" + rendered + "/>";
  EXPECT_TRUE(reparsed.load_string(element.c_str())) << element;
}

TEST(XmlUtilsCaptureAttrs, BindingIsFoundOnAnAncestor) {
  // Excel declares the prefixes once on the part root; a captured
  // attribute deeper in the tree still needs one emitted next to it.
  pugi::xml_document doc = Load("<root xmlns:xr=\"urn:rev\"><pivotField xr:uid=\"{DEF}\" compact=\"0\"/></root>");

  std::vector<std::pair<std::string, std::string>> captured;
  capture_unknown_attrs(doc.child("root").child("pivotField"), {}, captured);

  std::string rendered;
  append_raw_attrs(rendered, captured);
  EXPECT_NE(rendered.find("xmlns:xr=\"urn:rev\""), std::string::npos) << rendered;
  EXPECT_NE(rendered.find("compact=\"0\""), std::string::npos) << rendered;

  pugi::xml_document reparsed;
  const std::string element = "<e" + rendered + "/>";
  EXPECT_TRUE(reparsed.load_string(element.c_str())) << element;
}

TEST(XmlUtilsCaptureAttrs, UnprefixedAttributesAddNoBindings) {
  pugi::xml_document doc = Load("<root xmlns=\"urn:main\" compact=\"0\" outline=\"1\"/>");

  std::vector<std::pair<std::string, std::string>> captured;
  capture_unknown_attrs(doc.child("root"), {}, captured);

  ASSERT_EQ(captured.size(), 2U);
  std::string rendered;
  append_raw_attrs(rendered, captured);
  EXPECT_EQ(rendered.find("xmlns"), std::string::npos) << rendered;
}

TEST(XmlUtilsCaptureAttrs, UnbindablePrefixIsCapturedWithoutInventingABinding) {
  // Nothing declares `zz`. Emitting a made-up URI would be worse than the
  // loss it hides, so only the attribute travels -- and the document it
  // came from was already malformed.
  pugi::xml_document doc = Load("<root zz:thing=\"1\" xmlns:zz=\"urn:zz\"/>");
  pugi::xml_document orphan;
  orphan.load_string("<root/>");
  orphan.child("root").append_attribute("zz:thing") = "1";

  std::vector<std::pair<std::string, std::string>> captured;
  capture_unknown_attrs(orphan.child("root"), {}, captured);

  ASSERT_EQ(captured.size(), 1U);
  EXPECT_EQ(captured[0].first, "zz:thing");
}

}  // namespace
}  // namespace formulon::io
