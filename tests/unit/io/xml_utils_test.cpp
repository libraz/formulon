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

TEST(XmlUtilsAttr, U32StrictRejectsWhitespaceAndWrappedNegativeValues) {
  pugi::xml_document doc = Load("<r/>");
  pugi::xml_node node = doc.child("r");
  const std::vector<std::string> invalid = {"",           " ",
                                            " 0",         "0 ",
                                            "\t1",        " +1",
                                            " -0",        " -18446744073709551615",
                                            "+0",         "-0",
                                            "0junk",      "1.0",
                                            "4294967296", "18446744073709551616"};
  for (const std::string& value : invalid) {
    SCOPED_TRACE(value);
    node.append_attribute("index").set_value(value.c_str());
    EXPECT_FALSE(parse_xml_u32_attr_strict(node.attribute("index")).has_value());
    EXPECT_EQ(parse_xml_u32_attr(node.attribute("index"), 17U), 17U);
    EXPECT_EQ(attr_u32(node, "index", 17U), 17U);
    node.remove_attribute("index");
  }
  EXPECT_FALSE(parse_xml_u32_attr_strict(node.attribute("missing")).has_value());
}

TEST(XmlUtilsAttr, U32StrictPreservesDecimalBoundariesAndLeadingZeros) {
  pugi::xml_document doc = Load(R"(<r zero="0" padded="000001" max="4294967295"/>)");
  pugi::xml_node node = doc.child("r");
  EXPECT_EQ(parse_xml_u32_attr_strict(node.attribute("zero")), 0U);
  EXPECT_EQ(parse_xml_u32_attr_strict(node.attribute("padded")), 1U);
  EXPECT_EQ(parse_xml_u32_attr_strict(node.attribute("max")), UINT32_MAX);
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

TEST(XmlUtilsNamespaces, FindsDefaultAndNearestPrefixedBindings) {
  pugi::xml_document doc = Load(
      R"(<root xmlns="urn:default" xmlns:p="urn:root"><outer xmlns="urn:outer" xmlns:p="urn:outer-prefix"><inner/></outer></root>)");
  const pugi::xml_node inner = doc.child("root").child("outer").child("inner");

  EXPECT_EQ(find_namespace_binding(inner, ""), "urn:outer");
  EXPECT_EQ(find_namespace_binding(inner, "p"), "urn:outer-prefix");
  EXPECT_EQ(find_namespace_binding(doc.child("root"), ""), "urn:default");
  EXPECT_EQ(find_namespace_binding(doc.child("root"), "p"), "urn:root");
  EXPECT_TRUE(find_namespace_binding(inner, "missing").empty());
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

TEST(XmlUtilsRawXml, PcdataCarriageReturnsKeepCharacterReferencesAcrossTwoCycles) {
  const std::string expected_text = "one\rtwo\r\nthree & < _x0041_";
  const std::string xml =
      R"(<root name="attr&#13;value"><outer><text>one&#13;two&#13;&#10;three &amp; &lt; _x0041_</text></outer></root>)";
  pugi::xml_document first;
  ASSERT_TRUE(first.load_string(xml.c_str()));
  const std::string first_serialized = raw_xml(first.child("root"));
  EXPECT_EQ(first_serialized,
            R"(<root name="attr&#13;value"><outer><text>one&#13;two&#13;
three &amp; &lt; _x0041_</text></outer></root>)");
  EXPECT_EQ(first.child("root").child("outer").child("text").child_value(), expected_text);

  pugi::xml_document second;
  ASSERT_TRUE(second.load_string(first_serialized.c_str()));
  EXPECT_STREQ(second.child("root").attribute("name").value(), "attr\rvalue");
  EXPECT_EQ(second.child("root").child("outer").child("text").child_value(), expected_text);
  const std::string second_serialized = raw_xml(second.child("root"));
  EXPECT_EQ(second_serialized, first_serialized);

  pugi::xml_document third;
  ASSERT_TRUE(third.load_string(second_serialized.c_str()));
  EXPECT_EQ(third.child("root").child("outer").child("text").child_value(), expected_text);
  EXPECT_EQ(raw_xml(third.child("root")), first_serialized);

  pugi::xml_document ordinary = Load(R"(<plain id="7">a &amp; b &lt; c _x0041_</plain>)");
  EXPECT_EQ(raw_xml(ordinary.child("plain")), R"(<plain id="7">a &amp; b &lt; c _x0041_</plain>)");
}

TEST(XmlUtilsRawXml, LeavesCommentsCdataAndProcessingInstructionsUnchanged) {
  pugi::xml_document doc;
  pugi::xml_node root = doc.append_child("root");
  root.append_child(pugi::node_comment).set_value("comment\rtext");
  root.append_child(pugi::node_cdata).set_value("cdata\rtext");
  pugi::xml_node pi = root.append_child(pugi::node_pi);
  ASSERT_TRUE(pi.set_name("probe"));
  ASSERT_TRUE(pi.set_value("value\rtext"));

  EXPECT_EQ(raw_xml(root), "<root><!--comment\rtext--><![CDATA[cdata\rtext]]><?probe value\rtext?></root>");
}

TEST(XmlUtilsRawXml, MixedTextKeepsNonPcdataNodesAndPugiEscaping) {
  pugi::xml_document doc;
  pugi::xml_node root = doc.append_child("root");
  pugi::xml_node text = root.append_child("text");
  text.append_child(pugi::node_pcdata).set_value("A\rB \"_x0041_'");
  text.append_child(pugi::node_cdata).set_value("cdata\rtext");
  text.append_child(pugi::node_comment).set_value("comment\rtext");
  pugi::xml_node pi = text.append_child(pugi::node_pi);
  ASSERT_TRUE(pi.set_name("probe"));
  ASSERT_TRUE(pi.set_value("value\rtext"));

  EXPECT_EQ(
      raw_xml(root),
      "<root><text>A&#13;B \"_x0041_'<![CDATA[cdata\rtext]]><!--comment\rtext--><?probe value\rtext?></text></root>");
}

TEST(XmlUtilsRawXml, PCDATAAndDocumentRootsKeepBoundariesAndPrefixes) {
  pugi::xml_document doc;
  pugi::xml_node first = doc.append_child("first");
  first.text().set("A\rB");
  pugi::xml_node second = doc.append_child("second");
  second.text().set("C\rD");

  EXPECT_EQ(raw_xml(first.first_child()), "A&#13;B");
  EXPECT_EQ(raw_xml(doc), "<first>A&#13;B</first><second>C&#13;D</second>");

  std::string prefixed = "prefix\r";
  append_raw_xml(prefixed, first);
  EXPECT_EQ(prefixed, "prefix\r<first>A&#13;B</first>");
}

TEST(XmlUtilsRawXml, EmbeddedPCDATAKeepsCarriageReturnsAndEntities) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<root>A&amp;B&#13;_x0041_</root>", pugi::parse_default | pugi::parse_embed_pcdata));
  const pugi::xml_node root = doc.child("root");
  ASSERT_FALSE(root.first_child());
  EXPECT_STREQ(root.value(), "A&B\r_x0041_");

  const std::string serialized = raw_xml(root);
  EXPECT_EQ(serialized, "<root>A&amp;B&#13;_x0041_</root>");

  pugi::xml_document second;
  ASSERT_TRUE(second.load_string(serialized.c_str(), pugi::parse_default | pugi::parse_embed_pcdata));
  EXPECT_STREQ(second.child("root").value(), "A&B\r_x0041_");
}

TEST(XmlUtilsRawXml, LongTextAndSubtreeSerializationRemainStable) {
  pugi::xml_document doc;
  pugi::xml_node root = doc.append_child("root");
  pugi::xml_node target = root.append_child("target");
  const std::string text_payload = std::string(12U * 1024U, 'x') + "\r_end";
  const std::string expected_payload = std::string(12U * 1024U, 'x') + "&#13;_end";
  target.text().set(text_payload.c_str());
  root.append_child("sibling").text().set("outside\rsubtree");

  const std::string serialized = raw_xml(target);
  EXPECT_EQ(serialized, "<target>" + expected_payload + "</target>");
  EXPECT_EQ(raw_xml(root),
            "<root><target>" + expected_payload + "</target><sibling>outside&#13;subtree</sibling></root>");
}

TEST(XmlUtilsRawXml, DeepCarriageReturnSubtreeUsesIterativeFallback) {
  constexpr std::size_t kDepth = 5000U;
  pugi::xml_document doc;
  pugi::xml_node root = doc.append_child("level");
  pugi::xml_node leaf = root;
  for (std::size_t depth = 1U; depth < kDepth; ++depth) {
    leaf = leaf.append_child("level");
  }
  ASSERT_TRUE(leaf);
  leaf.append_child(pugi::node_pcdata).set_value("deep\rvalue");

  const std::string serialized = raw_xml(root);
  EXPECT_EQ(serialized.find("&#13;"), 7U * kDepth + 4U);
  EXPECT_EQ(serialized.find('\r'), std::string::npos);

  pugi::xml_document reparsed;
  ASSERT_TRUE(reparsed.load_string(serialized.c_str()));
  std::size_t observed_depth = 0;
  pugi::xml_node current = reparsed.document_element();
  while (current) {
    ++observed_depth;
    pugi::xml_node child = current.first_child();
    if (child.type() != pugi::node_element) {
      ASSERT_EQ(child.type(), pugi::node_pcdata);
      EXPECT_STREQ(child.value(), "deep\rvalue");
      break;
    }
    current = child;
  }
  EXPECT_EQ(observed_depth, kDepth);
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

std::string Num(double v) {
  std::string s;
  append_xml_number(s, v);
  return s;
}

TEST(XmlUtilsNumber, ExcelSpelling) {
  EXPECT_EQ(Num(0.0), "0");
  EXPECT_EQ(Num(-0.0), "0");
  EXPECT_EQ(Num(5E-324), "0");
  EXPECT_EQ(Num(0.1), "0.1");
  EXPECT_EQ(Num(1.0 / 3.0), "0.33333333333333331");
  EXPECT_EQ(Num(1000000000000000.0), "1000000000000000");
  EXPECT_EQ(Num(123456789012345.59), "123456789012345.59");
  EXPECT_EQ(Num(12345.678901234567), "12345.678901234567");
  EXPECT_EQ(Num(1428571428.5714285), "1428571428.5714285");
  EXPECT_EQ(Num(100.0), "100");
  EXPECT_EQ(Num(-2.5), "-2.5");
  EXPECT_EQ(Num(0.0123), "1.23E-2");
  EXPECT_EQ(Num(1E-3), "1E-3");
  EXPECT_EQ(Num(1E-4), "1E-4");
  EXPECT_EQ(Num(1E-5), "1.0000000000000001E-5");
  EXPECT_EQ(Num(1.5E-5), "1.5E-5");
  EXPECT_EQ(Num(1E-6), "9.9999999999999995E-7");
  EXPECT_EQ(Num(1.2345678901234499E-4), "1.2345678901234499E-4");
  EXPECT_EQ(Num(1E16), "1E+16");
  EXPECT_EQ(Num(1.2345678901234568E+17), "1.2345678901234568E+17");
  EXPECT_EQ(Num(-1E-20), "-9.9999999999999995E-21");
  EXPECT_EQ(Num(1E-300), "1E-300");
  EXPECT_EQ(Num(1.7976931348623157E+308), "1.7976931348623157E+308");
  EXPECT_EQ(Num(1E21), "1E+21");
}

}  // namespace
}  // namespace formulon::io
