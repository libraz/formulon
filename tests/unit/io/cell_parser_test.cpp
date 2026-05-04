// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for `formulon::io::parse_cell_element` and `parse_a1`. The
// tests construct `<c>` elements via pugixml `load_string` so they
// exercise the same code path the OOXML reader uses on real workbook
// archives.
//
// `parse_cell_element` takes a `std::deque<std::string>&` for inline
// text storage; each test owns its own deque on the stack so the
// `Value::text` views remain valid for the assertions below.

#include "io/cell_parser.h"

#include <cstdint>
#include <deque>
#include <string>
#include <utility>

#include "gtest/gtest.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace io {
namespace {

// `ASSERT_PARSE_OK(VAR, XML, DOC, STORAGE)` loads `XML` into `DOC`,
// parses its first child as a `<c>` element using `STORAGE` as the
// caller-owned text-storage buffer, and binds the resulting
// `ParsedCell` to `VAR` by const reference.
#define ASSERT_PARSE_OK(VAR, XML, DOC, STORAGE)                                                \
  ASSERT_TRUE((DOC).load_string((XML))) << "load_string failed for: " << (XML);                \
  auto VAR##_or = parse_cell_element((DOC).first_child(), (STORAGE));                          \
  ASSERT_TRUE(static_cast<bool>(VAR##_or))                                                     \
      << "parse error: " << (VAR##_or.has_value() ? std::string{} : VAR##_or.error().message); \
  const ParsedCell& VAR = VAR##_or.value()

// ---------------------------------------------------------------------------
// parse_a1 helper coverage
// ---------------------------------------------------------------------------

TEST(CellParser_A1, BasicAndUpperCorner) {
  auto a1 = parse_a1("A1");
  ASSERT_TRUE(static_cast<bool>(a1));
  EXPECT_EQ(a1.value().first, 0U);
  EXPECT_EQ(a1.value().second, 0U);

  auto z1 = parse_a1("Z1");
  ASSERT_TRUE(static_cast<bool>(z1));
  EXPECT_EQ(z1.value().first, 0U);
  EXPECT_EQ(z1.value().second, 25U);

  auto aa1 = parse_a1("AA1");
  ASSERT_TRUE(static_cast<bool>(aa1));
  EXPECT_EQ(aa1.value().first, 0U);
  EXPECT_EQ(aa1.value().second, 26U);

  auto ab1 = parse_a1("AB1");
  ASSERT_TRUE(static_cast<bool>(ab1));
  EXPECT_EQ(ab1.value().first, 0U);
  EXPECT_EQ(ab1.value().second, 27U);

  auto last = parse_a1("XFD1048576");
  ASSERT_TRUE(static_cast<bool>(last));
  EXPECT_EQ(last.value().first, 1048575U);
  EXPECT_EQ(last.value().second, 16383U);
}

TEST(CellParser_A1, RejectsInvalidShapes) {
  EXPECT_FALSE(static_cast<bool>(parse_a1("")));
  EXPECT_FALSE(static_cast<bool>(parse_a1("A")));          // column only
  EXPECT_FALSE(static_cast<bool>(parse_a1("1")));          // row only
  EXPECT_FALSE(static_cast<bool>(parse_a1("A1B")));        // trailing chars
  EXPECT_FALSE(static_cast<bool>(parse_a1("A1:B2")));      // range
  EXPECT_FALSE(static_cast<bool>(parse_a1("Sheet1!A1")));  // sheet-qualified
  EXPECT_FALSE(static_cast<bool>(parse_a1("$A$1")));       // absolute markers
  EXPECT_FALSE(static_cast<bool>(parse_a1("a1")));         // lowercase column
  EXPECT_FALSE(static_cast<bool>(parse_a1("XFE1")));       // out-of-range column
  EXPECT_FALSE(static_cast<bool>(parse_a1("ABCD1")));      // too many letters
}

TEST(CellParser_A1, RejectsRowOutOfRange) {
  // Excel max row is 1,048,576. parse_a1 rejects 1,048,577.
  auto over = parse_a1("A1048577");
  EXPECT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(over.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

// ---------------------------------------------------------------------------
// parse_cell_element coverage
// ---------------------------------------------------------------------------

TEST(CellParser, NumberLiteral) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"A1\"><v>42</v></c>", doc, storage);
  EXPECT_EQ(parsed.row, 0U);
  EXPECT_EQ(parsed.col, 0U);
  EXPECT_TRUE(parsed.formula.empty());
  ASSERT_TRUE(parsed.value.is_number());
  EXPECT_DOUBLE_EQ(parsed.value.as_number(), 42.0);
  EXPECT_FALSE(parsed.is_sst_index);
}

TEST(CellParser, NumberLiteralExplicitTypeN) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"B2\" t=\"n\"><v>3.14</v></c>", doc, storage);
  EXPECT_EQ(parsed.row, 1U);
  EXPECT_EQ(parsed.col, 1U);
  ASSERT_TRUE(parsed.value.is_number());
  EXPECT_DOUBLE_EQ(parsed.value.as_number(), 3.14);
}

TEST(CellParser, BooleanTrueFalse) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(t_cell, "<c r=\"B3\" t=\"b\"><v>1</v></c>", doc, storage);
  EXPECT_EQ(t_cell.row, 2U);
  EXPECT_EQ(t_cell.col, 1U);
  ASSERT_TRUE(t_cell.value.is_boolean());
  EXPECT_TRUE(t_cell.value.as_boolean());

  pugi::xml_document doc2;
  std::deque<std::string> storage2;
  ASSERT_PARSE_OK(f_cell, "<c r=\"C5\" t=\"b\"><v>0</v></c>", doc2, storage2);
  EXPECT_EQ(f_cell.row, 4U);
  EXPECT_EQ(f_cell.col, 2U);
  ASSERT_TRUE(f_cell.value.is_boolean());
  EXPECT_FALSE(f_cell.value.as_boolean());
}

TEST(CellParser, BooleanRejectsBadBody) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<c r=\"A1\" t=\"b\"><v>2</v></c>"));
  std::deque<std::string> storage;
  auto result = parse_cell_element(doc.first_child(), storage);
  EXPECT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CellParser, ErrorDiv0AndName) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(div0, "<c r=\"D7\" t=\"e\"><v>#DIV/0!</v></c>", doc, storage);
  EXPECT_EQ(div0.row, 6U);
  EXPECT_EQ(div0.col, 3U);
  ASSERT_TRUE(div0.value.is_error());
  EXPECT_EQ(div0.value.as_error(), ErrorCode::Div0);

  pugi::xml_document doc2;
  std::deque<std::string> storage2;
  ASSERT_PARSE_OK(name_err, "<c r=\"E9\" t=\"e\"><v>#NAME?</v></c>", doc2, storage2);
  EXPECT_EQ(name_err.row, 8U);
  EXPECT_EQ(name_err.col, 4U);
  ASSERT_TRUE(name_err.value.is_error());
  EXPECT_EQ(name_err.value.as_error(), ErrorCode::Name);
}

TEST(CellParser, ErrorRejectsUnknownDisplay) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<c r=\"A1\" t=\"e\"><v>#WAT?</v></c>"));
  std::deque<std::string> storage;
  auto result = parse_cell_element(doc.first_child(), storage);
  EXPECT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CellParser, InlineStringSimple) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"F11\" t=\"inlineStr\"><is><t>hello</t></is></c>", doc, storage);
  EXPECT_EQ(parsed.row, 10U);
  EXPECT_EQ(parsed.col, 5U);
  ASSERT_TRUE(parsed.value.is_text());
  EXPECT_EQ(parsed.value.as_text(), "hello");
  // Verifies the storage actually backs the view (lifetime contract).
  ASSERT_EQ(storage.size(), 1U);
  EXPECT_EQ(storage.front(), "hello");
}

TEST(CellParser, InlineStringRichTextConcatenated) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"G13\" t=\"inlineStr\"><is><r><t>foo</t></r><r><t>bar</t></r></is></c>", doc, storage);
  EXPECT_EQ(parsed.row, 12U);
  EXPECT_EQ(parsed.col, 6U);
  ASSERT_TRUE(parsed.value.is_text());
  EXPECT_EQ(parsed.value.as_text(), "foobar");
}

TEST(CellParser, SharedStringSurfacesIndex) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"H15\" t=\"s\"><v>3</v></c>", doc, storage);
  EXPECT_EQ(parsed.row, 14U);
  EXPECT_EQ(parsed.col, 7U);
  EXPECT_TRUE(parsed.is_sst_index);
  EXPECT_EQ(parsed.sst_index, 3U);
  ASSERT_TRUE(parsed.value.is_text());
  EXPECT_EQ(parsed.value.as_text(), "");
  // SST cells must not append to text_storage; their backing comes
  // from the shared-strings table at resolve time.
  EXPECT_EQ(storage.size(), 0U);
}

TEST(CellParser, FormulaWithCachedValue) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"I1\"><f>SUM(A1:A10)</f><v>55</v></c>", doc, storage);
  EXPECT_EQ(parsed.row, 0U);
  EXPECT_EQ(parsed.col, 8U);
  EXPECT_EQ(parsed.formula, "SUM(A1:A10)");
  ASSERT_TRUE(parsed.value.is_number());
  EXPECT_DOUBLE_EQ(parsed.value.as_number(), 55.0);
}

TEST(CellParser, FormulaWithoutCachedValue) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"J1\"><f>SUM(A1:A10)</f></c>", doc, storage);
  EXPECT_EQ(parsed.row, 0U);
  EXPECT_EQ(parsed.col, 9U);
  EXPECT_EQ(parsed.formula, "SUM(A1:A10)");
  EXPECT_TRUE(parsed.value.is_blank());
}

TEST(CellParser, FormulaStripsLeadingEquals) {
  // Defensive: most writers omit '=', but the parser strips it if present.
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"A1\"><f>=A2+1</f></c>", doc, storage);
  EXPECT_EQ(parsed.formula, "A2+1");
}

TEST(CellParser, EmptyCellIsBlank) {
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"A1\"></c>", doc, storage);
  EXPECT_EQ(parsed.row, 0U);
  EXPECT_EQ(parsed.col, 0U);
  EXPECT_TRUE(parsed.value.is_blank());
  EXPECT_TRUE(parsed.formula.empty());
}

TEST(CellParser, RejectsMissingRefAttribute) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<c><v>1</v></c>"));
  std::deque<std::string> storage;
  auto result = parse_cell_element(doc.first_child(), storage);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CellParser, RejectsInvalidRef) {
  // "A0" — row 0 is invalid in Excel.
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<c r=\"A0\"><v>1</v></c>"));
  std::deque<std::string> storage;
  auto result = parse_cell_element(doc.first_child(), storage);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CellParser, RejectsBogusType) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<c r=\"A1\" t=\"totally-bogus\"><v>x</v></c>"));
  std::deque<std::string> storage;
  auto result = parse_cell_element(doc.first_child(), storage);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CellParser, NumberWithUnparseableValue) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<c r=\"A1\"><v>not-a-number</v></c>"));
  std::deque<std::string> storage;
  auto result = parse_cell_element(doc.first_child(), storage);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CellParser, SharedStringIndexLargeIsExact) {
  // Regression: SST indices >= 2^24 + 1 used to be parsed via `double`
  // and silently rounded onto an even neighbour, surfacing the wrong
  // shared string. Verify the boundary value 16777217 (= 2^24 + 1)
  // round-trips exactly.
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"A1\" t=\"s\"><v>16777217</v></c>", doc, storage);
  EXPECT_TRUE(parsed.is_sst_index);
  EXPECT_EQ(parsed.sst_index, 16777217U);
}

TEST(CellParser, SharedStringIndexNegativeIsRejected) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<c r=\"A1\" t=\"s\"><v>-1</v></c>"));
  std::deque<std::string> storage;
  auto result = parse_cell_element(doc.first_child(), storage);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CellParser, SharedStringIndexOverflowIsRejected) {
  // 4294967296 (= 2^32) overflows uint32; must be rejected, not wrapped.
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string("<c r=\"A1\" t=\"s\"><v>4294967296</v></c>"));
  std::deque<std::string> storage;
  auto result = parse_cell_element(doc.first_child(), storage);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CellParser, LegacyStrTypeIsTextFromValue) {
  // t="str" is the legacy formula-result-as-string shape. The cached
  // text lives in <v> rather than <is>.
  pugi::xml_document doc;
  std::deque<std::string> storage;
  ASSERT_PARSE_OK(parsed, "<c r=\"A1\" t=\"str\"><f>UPPER(\"hi\")</f><v>HI</v></c>", doc, storage);
  EXPECT_EQ(parsed.formula, "UPPER(\"hi\")");
  ASSERT_TRUE(parsed.value.is_text());
  EXPECT_EQ(parsed.value.as_text(), "HI");
}

}  // namespace
}  // namespace io
}  // namespace formulon
