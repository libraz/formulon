// Copyright 2026 libraz. Licensed under the MIT License.
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

TEST(BuiltinsWebEncodeUrl, UnreservedCharsPassThroughUnchanged) {
  // RFC 3986 unreserved: A-Z a-z 0-9 - _ . ~
  const Value v = EvalSource("=ENCODEURL(\"ABCabc012-_.~\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "ABCabc012-_.~");
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

TEST(BuiltinsWebFilterXml, MultiNodeSetReturnsFirst) {
  // Mac Excel 365 spills the full node set to adjacent cells via dynamic-
  // array spill, which Formulon does not yet implement at the cell level.
  // Returning a Value::Array here without spill plumbing would be worse
  // than the current scalar fallback, so we return the first node's text.
  // See TODO(filterxml-spill) in src/eval/builtins/web.cpp.
  const Value v = EvalSource("=FILTERXML(\"<r><a>1</a><a>2</a></r>\",\"//a\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1");
}

TEST(BuiltinsWebFilterXml, EmptyNodeSetReturnsNotAvailable) {
  const Value v = EvalSource("=FILTERXML(\"<r><a>1</a></r>\",\"//b\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
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
