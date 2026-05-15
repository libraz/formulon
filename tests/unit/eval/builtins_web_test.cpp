// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the Web-category built-ins:
//   * ENCODEURL   -- real impl (RFC 3986 percent-encoding, uppercase hex).
//   * FILTERXML   -- real impl (pugixml + XPath 1.0).
//   * WEBSERVICE  -- deterministic #VALUE! stub.
//   * PY          -- deterministic #NAME? stub.
//
// Stub arguments are evaluated for error-propagation side effects; the
// tests at the bottom of the file pin that behaviour.

#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry. Arenas
// are reset between calls to avoid cross-test contamination.
Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

// ---------------------------------------------------------------------------
// Registry pin
// ---------------------------------------------------------------------------

TEST(BuiltinsWebRegistry, NamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("ENCODEURL"), nullptr);
  EXPECT_NE(reg.lookup("FILTERXML"), nullptr);
  EXPECT_NE(reg.lookup("WEBSERVICE"), nullptr);
  EXPECT_NE(reg.lookup("PY"), nullptr);
}

// ---------------------------------------------------------------------------
// ENCODEURL
// ---------------------------------------------------------------------------

TEST(BuiltinsWebEncodeUrl, SpaceEncodesAsPercent20) {
  const Value v = EvalSource("=ENCODEURL(\"hello world\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello%20world");
}

TEST(BuiltinsWebEncodeUrl, QueryStringCharsAreEncoded) {
  const Value v = EvalSource("=ENCODEURL(\"a/b c?d=1&e=2\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a%2Fb%20c%3Fd%3D1%26e%3D2");
}

TEST(BuiltinsWebEncodeUrl, ExcelEncodesTilde) {
  const Value v = EvalSource("=ENCODEURL(\"ABCabc012-_.~\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ABCabc012-_.%7E");
}

TEST(BuiltinsWebEncodeUrl, UsesUppercaseHex) {
  // Each Japanese character takes three UTF-8 bytes; ENCODEURL emits them
  // as uppercase `%XX` sequences. This is the concrete uppercase pin.
  const Value v = EvalSource("=ENCODEURL(\"\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "%E6%97%A5%E6%9C%AC%E8%AA%9E");
}

TEST(BuiltinsWebEncodeUrl, BooleanCoercesToLiteral) {
  const Value v = EvalSource("=ENCODEURL(TRUE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "TRUE");
}

TEST(BuiltinsWebEncodeUrl, FalseBooleanCoercesToLiteral) {
  const Value v = EvalSource("=ENCODEURL(FALSE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "FALSE");
}

TEST(BuiltinsWebEncodeUrl, NumberCoercesToText) {
  const Value v = EvalSource("=ENCODEURL(123)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "123");
}

TEST(BuiltinsWebEncodeUrl, EmptyStringPassesThrough) {
  const Value v = EvalSource("=ENCODEURL(\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(BuiltinsWebEncodeUrl, PropagatesDivideByZero) {
  const Value v = EvalSource("=ENCODEURL(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsWebEncodeUrl, PercentIsEncoded) {
  // '%' itself is not in the unreserved set, so it encodes to "%25".
  const Value v = EvalSource("=ENCODEURL(\"100%\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "100%25");
}

TEST(BuiltinsWebEncodeUrl, ArityMismatchSurfacesValue) {
  const Value v = EvalSource("=ENCODEURL(\"a\",\"b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// FILTERXML
// ---------------------------------------------------------------------------

TEST(BuiltinsWebFilterXml, ReturnsSingleNodeText) {
  const Value v = EvalSource("=FILTERXML(\"<r><a>hello</a></r>\",\"//a\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "hello");
}

TEST(BuiltinsWebFilterXml, SingleNodeMatchStillReturnsScalar) {
  // Single match returns plain text (matches Mac Excel's anchor read-back
  // and pins the historical scalar contract for one-node node sets).
  const Value v = EvalSource("=FILTERXML(\"<r><a>only</a></r>\",\"//a\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "only");
}

TEST(BuiltinsWebFilterXml, MultiNodeSetReturnsArray) {
  // Mac Excel 365 spills the full node set into a vertical N x 1 region
  // via dynamic-array spill. Formulon now returns the spill payload as
  // a Value::Array; sheet contexts commit it through
  // dispatch_array_result, non-sheet contexts (here) see the Array
  // directly.
  const Value v = EvalSource("=FILTERXML(\"<r><a>1</a><a>2</a><a>3</a></r>\",\"//a\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  ASSERT_TRUE(v.as_array_cells()[1].is_number());
  ASSERT_TRUE(v.as_array_cells()[2].is_number());
  EXPECT_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_EQ(v.as_array_cells()[1].as_number(), 2.0);
  EXPECT_EQ(v.as_array_cells()[2].as_number(), 3.0);
}

TEST(BuiltinsWebFilterXml, MultiNodeSetSpillAnchorMatchesFirstNode) {
  // Anchor unwrap: the comparator-side projection landed in commit 779b028.
  // Mirror that here at the value layer so callers that intentionally
  // collapse an Array result (e.g., scalar consumers) see the first node.
  const Value v = EvalSource("=FILTERXML(\"<r><a>1</a><a>2</a></r>\",\"//a\")");
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_EQ(v.as_array_cells()[1].as_number(), 2.0);
}

TEST(BuiltinsWebFilterXml, AttributeAxisSpillsAcrossMatches) {
  // Attribute-axis multi-match: each `xpath_node` carries an attribute
  // rather than an element node. `node_text` reads the attribute's
  // value, and the spill shape is identical to element-axis multi-match.
  const Value v = EvalSource("=FILTERXML(\"<r><a id='x'/><a id='y'/></r>\",\"//a/@id\")");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "x");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "y");
}

TEST(BuiltinsWebFilterXml, EmptyNodeSetReturnsNotAvailable) {
  const Value v = EvalSource("=FILTERXML(\"<r><a>1</a></r>\",\"//b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsWebFilterXml, MalformedXmlReturnsValue) {
  const Value v = EvalSource("=FILTERXML(\"not xml\",\"//a\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsWebFilterXml, MalformedXPathReturnsValue) {
  // "///" is a syntax error in XPath 1.0.
  const Value v = EvalSource("=FILTERXML(\"<r/>\",\"///\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsWebFilterXml, NumericArgumentFailsXmlParse) {
  // Numbers coerce to text ("42"), which is not valid XML -> #VALUE!.
  const Value v = EvalSource("=FILTERXML(42,\"//a\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsWebFilterXml, BooleanArgumentFailsXmlParse) {
  // TRUE coerces to "TRUE" which is not valid XML -> #VALUE!.
  const Value v = EvalSource("=FILTERXML(TRUE,\"//a\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsWebFilterXml, AttributeNodeReturnsAttributeValue) {
  // `//@attr` is an attribute-node xpath; we return the attribute's value.
  const Value v = EvalSource("=FILTERXML(\"<r a=\"\"v\"\"/>\",\"//@a\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "v");
}

TEST(BuiltinsWebFilterXml, DefaultNamespaceCanBeQueriedWithLocalName) {
  const Value v = EvalSource("=FILTERXML(\"<r xmlns=\"\"urn:x\"\"><a>v</a></r>\",\"//*[local-name()=\"\"a\"\"]\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "v");
}

TEST(BuiltinsWebFilterXml, CdataReturnsTextContent) {
  const Value v = EvalSource("=FILTERXML(\"<r><![CDATA[a<b&c]]></r>\",\"//r\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a<b&c");
}

TEST(BuiltinsWebFilterXml, EntityReferencesAreDecoded) {
  const Value v = EvalSource("=FILTERXML(\"<r>a&amp;b&#10;&#x41;</r>\",\"//r\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "a&b\nA");
}

TEST(BuiltinsWebFilterXml, TextAxisOverMixedContentSpillsTextNodes) {
  const Value v = EvalSource("=FILTERXML(\"<r>a<b>B</b>c</r>\",\"//r/text()\")");
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_TRUE(v.as_array_cells()[0].is_error());
  EXPECT_EQ(v.as_array_cells()[0].as_error(), ErrorCode::Value);
  EXPECT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Value);
}

TEST(BuiltinsWebFilterXml, PropagatesErrorInFirstArg) {
  const Value v = EvalSource("=FILTERXML(1/0,\"//a\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsWebFilterXml, PropagatesErrorInSecondArg) {
  const Value v = EvalSource("=FILTERXML(\"<r/>\",1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsWebFilterXml, ArityMismatchSurfacesValue) {
  const Value v = EvalSource("=FILTERXML(\"<r/>\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// WEBSERVICE (stub)
// ---------------------------------------------------------------------------

TEST(BuiltinsWebWebService, AlwaysReturnsValue) {
  const Value v = EvalSource("=WEBSERVICE(\"http://example.com\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsWebWebService, PropagatesArgumentError) {
  // An error in the argument short-circuits before the stub body runs.
  const Value v = EvalSource("=WEBSERVICE(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsWebWebService, ArityMismatchSurfacesValue) {
  const Value v = EvalSource("=WEBSERVICE()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// PY (stub)
// ---------------------------------------------------------------------------

TEST(BuiltinsWebPy, AlwaysReturnsName) {
  const Value v = EvalSource("=PY(\"1+1\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsWebPy, PropagatesArgumentError) {
  const Value v = EvalSource("=PY(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
