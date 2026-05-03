// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Round-trip tests for `xl/comments<N>.xml`. Builds a `CellComment`
// list, runs it through the writer, parses it back via the reader, and
// asserts the output matches the input.

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/comments_reader.h"
#include "io/comments_writer.h"
#include "sheet.h"

namespace formulon {
namespace io {
namespace {

std::vector<std::uint8_t> AsBytes(const std::string& s) {
  return std::vector<std::uint8_t>(s.begin(), s.end());
}

TEST(CommentsRoundTrip, EmptyListProducesEmptyString) {
  std::vector<CellComment> input;
  EXPECT_TRUE(write_comments(input).empty());
}

TEST(CommentsRoundTrip, SingleCommentRoundTrips) {
  std::vector<CellComment> input;
  CellComment c;
  c.row = 2;
  c.col = 3;
  c.author = "Alice";
  c.text = "Hello world";
  input.push_back(std::move(c));
  std::string xml = write_comments(input);
  ASSERT_FALSE(xml.empty());
  auto out_or = read_comments(AsBytes(xml));
  ASSERT_TRUE(static_cast<bool>(out_or)) << out_or.error().message;
  ASSERT_EQ(out_or.value().size(), 1U);
  const CellComment& got = out_or.value()[0];
  EXPECT_EQ(got.row, 2U);
  EXPECT_EQ(got.col, 3U);
  EXPECT_EQ(got.author, "Alice");
  EXPECT_EQ(got.text, "Hello world");
}

TEST(CommentsRoundTrip, MultipleAuthorsAreDeduplicated) {
  std::vector<CellComment> input;
  for (std::uint32_t i = 0; i < 4; ++i) {
    CellComment c;
    c.row = i;
    c.col = 0;
    c.author = (i % 2 == 0) ? "Alice" : "Bob";
    c.text = "comment " + std::to_string(i);
    input.push_back(std::move(c));
  }
  std::string xml = write_comments(input);
  auto out_or = read_comments(AsBytes(xml));
  ASSERT_TRUE(static_cast<bool>(out_or));
  ASSERT_EQ(out_or.value().size(), 4U);
  for (std::uint32_t i = 0; i < 4; ++i) {
    EXPECT_EQ(out_or.value()[i].row, i);
    EXPECT_EQ(out_or.value()[i].author, (i % 2 == 0) ? "Alice" : "Bob");
    EXPECT_EQ(out_or.value()[i].text, "comment " + std::to_string(i));
  }
}

TEST(CommentsRoundTrip, XmlEntityEscaping) {
  std::vector<CellComment> input;
  CellComment c;
  c.row = 0;
  c.col = 0;
  c.author = "<amp&aut>";
  c.text = "Use &amp; and \"quotes\" plus <tags>";
  input.push_back(std::move(c));
  std::string xml = write_comments(input);
  auto out_or = read_comments(AsBytes(xml));
  ASSERT_TRUE(static_cast<bool>(out_or));
  ASSERT_EQ(out_or.value().size(), 1U);
  EXPECT_EQ(out_or.value()[0].author, "<amp&aut>");
  EXPECT_EQ(out_or.value()[0].text, "Use &amp; and \"quotes\" plus <tags>");
}

TEST(CommentsRoundTrip, MultiRunRichTextConcatenated) {
  // Reader-only test: rich text with multiple runs collapses into a
  // single concatenated plain-text payload.
  const std::string xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
      "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<authors><author>Bob</author></authors>"
      "<commentList>"
      "<comment ref=\"A1\" authorId=\"0\">"
      "<text><r><t>foo</t></r><r><t>bar</t></r></text>"
      "</comment>"
      "</commentList></comments>";
  auto out_or = read_comments(AsBytes(xml));
  ASSERT_TRUE(static_cast<bool>(out_or));
  ASSERT_EQ(out_or.value().size(), 1U);
  EXPECT_EQ(out_or.value()[0].text, "foobar");
}

TEST(CommentsRoundTrip, MissingRootRejected) {
  const std::string xml = "<?xml version=\"1.0\"?><notcomments/>";
  auto out_or = read_comments(AsBytes(xml));
  EXPECT_FALSE(static_cast<bool>(out_or));
  EXPECT_EQ(out_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CommentsRoundTrip, VmlStubIsValidXml) {
  const std::string vml = write_vml_drawing_stub();
  EXPECT_FALSE(vml.empty());
  // Must contain the canonical VML namespace declaration so legacy
  // parsers recognise the document.
  EXPECT_NE(vml.find("urn:schemas-microsoft-com:vml"), std::string::npos);
}

}  // namespace
}  // namespace io
}  // namespace formulon
