// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Round-trip tests for the `<mergeCells>` reader. The matching writer
// path lives in `ooxml_writer.cpp::BuildMergeCellsBlock`; the reader's
// public entry is `read_merges` in `sheet_reader.h`. Each test builds a
// minimal `<worksheet>` DOM in memory, runs `read_merges`, and checks
// the output against the expected `MergeRange` list.

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/sheet_reader.h"
#include "pugixml.hpp"
#include "sheet.h"

namespace formulon {
namespace io {
namespace {

pugi::xml_node ParseWorksheet(pugi::xml_document& doc, const std::string& body) {
  const std::string xml = "<worksheet>" + body + "</worksheet>";
  EXPECT_TRUE(doc.load_string(xml.c_str()));
  return doc.child("worksheet");
}

TEST(MergeRoundTrip, NoMergesYieldsEmptyVector) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc, "<sheetData/>");
  auto out = read_merges(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  EXPECT_TRUE(out.value().empty());
}

TEST(MergeRoundTrip, SingleMergeRectangle) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc,
                           "<mergeCells count=\"1\">"
                           "<mergeCell ref=\"A1:B2\"/>"
                           "</mergeCells>");
  auto out = read_merges(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  ASSERT_EQ(out.value().size(), 1U);
  const MergeRange& m = out.value()[0];
  EXPECT_EQ(m.first_row, 0U);
  EXPECT_EQ(m.first_col, 0U);
  EXPECT_EQ(m.last_row, 1U);
  EXPECT_EQ(m.last_col, 1U);
}

TEST(MergeRoundTrip, MultipleRanges) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc,
                           "<mergeCells count=\"3\">"
                           "<mergeCell ref=\"A1:A3\"/>"
                           "<mergeCell ref=\"B5:D5\"/>"
                           "<mergeCell ref=\"E10\"/>"
                           "</mergeCells>");
  auto out = read_merges(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  ASSERT_EQ(out.value().size(), 3U);
  EXPECT_EQ(out.value()[0].first_row, 0U);
  EXPECT_EQ(out.value()[0].last_row, 2U);
  EXPECT_EQ(out.value()[1].first_col, 1U);
  EXPECT_EQ(out.value()[1].last_col, 3U);
  EXPECT_EQ(out.value()[2].first_row, 9U);
  EXPECT_EQ(out.value()[2].first_col, 4U);
  EXPECT_EQ(out.value()[2].last_row, 9U);
  EXPECT_EQ(out.value()[2].last_col, 4U);
}

TEST(MergeRoundTrip, EmptyRefRejected) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc, "<mergeCells><mergeCell ref=\"\"/></mergeCells>");
  auto out = read_merges(ws);
  EXPECT_FALSE(static_cast<bool>(out));
  EXPECT_EQ(out.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(MergeRoundTrip, NormalisesReversedCorners) {
  pugi::xml_document doc;
  auto ws = ParseWorksheet(doc, "<mergeCells><mergeCell ref=\"C5:A1\"/></mergeCells>");
  auto out = read_merges(ws);
  ASSERT_TRUE(static_cast<bool>(out));
  ASSERT_EQ(out.value().size(), 1U);
  // Reader normalises corners so first <= last.
  EXPECT_EQ(out.value()[0].first_row, 0U);
  EXPECT_EQ(out.value()[0].first_col, 0U);
  EXPECT_EQ(out.value()[0].last_row, 4U);
  EXPECT_EQ(out.value()[0].last_col, 2U);
}

}  // namespace
}  // namespace io
}  // namespace formulon
