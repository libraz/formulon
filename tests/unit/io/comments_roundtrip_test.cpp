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

std::string NoteShape(std::uint32_t row, std::uint32_t col, const char* style) {
  return std::string("<v:shape id=\"s\" style=\"") + style +
         "\"><x:ClientData ObjectType=\"Note\"><x:Anchor>1, 15, 0, 2, 3, 15, 4, 4</x:Anchor><x:Row>" +
         std::to_string(row) + "</x:Row><x:Column>" + std::to_string(col) + "</x:Column></x:ClientData></v:shape>";
}

TEST(CommentsRoundTrip, VmlNoteShapesMatchTheirCommentsOnce) {
  std::vector<CellComment> comments(3);
  comments[0].row = 0;
  comments[0].col = 0;
  comments[1].row = 4;
  comments[1].col = 2;
  comments[2].row = 4;
  comments[2].col = 2;  // a second comment at a cell keeps no shape
  const std::string vml = "<xml>" + NoteShape(4, 2, "visibility:visible") + NoteShape(0, 0, "visibility:hidden") +
                          "<v:shape id=\"c\"><x:ClientData ObjectType=\"Checkbox\"><x:Row>9</x:Row>"
                          "<x:Column>9</x:Column></x:ClientData></v:shape></xml>";
  apply_vml_note_shapes(vml, comments);
  EXPECT_FALSE(comments[0].visible);
  EXPECT_TRUE(comments[0].has_box);
  EXPECT_EQ(comments[0].vml_shape, NoteShape(0, 0, "visibility:hidden"));
  EXPECT_TRUE(comments[1].visible);
  EXPECT_EQ(comments[1].vml_shape, NoteShape(4, 2, "visibility:visible"));
  EXPECT_TRUE(comments[2].vml_shape.empty());
}

TEST(CommentsRoundTrip, VmlRunOfUnclosedShapeOpensIsScannedOnce) {
  // ~10 MB of opens ahead of one closed note; a rescan per open is quadratic.
  std::string vml;
  vml.reserve(10U * 1024U * 1024U + 256U);
  while (vml.size() < 10U * 1024U * 1024U) {
    vml.append("<v:shape ");
  }
  vml.append(NoteShape(1, 1, "visibility:visible"));
  std::vector<CellComment> comments(1);
  comments[0].row = 1;
  comments[0].col = 1;
  apply_vml_note_shapes(vml, comments);
  EXPECT_TRUE(comments[0].visible);
  EXPECT_EQ(comments[0].vml_shape, NoteShape(1, 1, "visibility:visible"));
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

TEST(CommentsRoundTrip, AuthorAndTextSurviveTwoOoxmlCycles) {
  const std::string payload = "literal _x0041_ & <猫> \rcontrol \x01 🙂";
  std::vector<CellComment> current = {CellComment(0, 0, payload, payload)};

  for (int cycle = 0; cycle < 2; ++cycle) {
    const std::string xml = write_comments(current);
    auto decoded = read_comments(AsBytes(xml));
    ASSERT_TRUE(static_cast<bool>(decoded)) << decoded.error().message;
    ASSERT_EQ(decoded.value().size(), 1U);
    current = std::move(decoded.value());
  }

  EXPECT_EQ(current[0].author, payload);
  EXPECT_EQ(current[0].text, payload);
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

TEST(CommentsRoundTrip, InvalidAuthorIdsDoNotSelectAnAuthor) {
  const std::string xml =
      "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<authors><author>Alice</author><author>Bob</author></authors>"
      "<commentList>"
      "<comment ref=\"A1\" authorId=\"0junk\"><text><t>suffix</t></text></comment>"
      "<comment ref=\"B1\" authorId=\"0.5\"><text><t>fraction</t></text></comment>"
      "<comment ref=\"C1\" authorId=\"-1\"><text><t>negative</t></text></comment>"
      "<comment ref=\"D1\" authorId=\"4294967296\"><text><t>overflow</t></text></comment>"
      "<comment ref=\"E1\"><text><t>missing</t></text></comment>"
      "<comment ref=\"F1\" authorId=\"2\"><text><t>out of range</t></text></comment>"
      "<comment ref=\"G1\" authorId=\"0\"><text><t>zero</t></text></comment>"
      "<comment ref=\"H1\" authorId=\"1\"><text><t>one</t></text></comment>"
      "<comment ref=\"I1\" authorId=\"01\"><text><t>leading zero</t></text></comment>"
      "</commentList></comments>";

  auto out_or = read_comments(AsBytes(xml));
  ASSERT_TRUE(static_cast<bool>(out_or)) << out_or.error().message;
  ASSERT_EQ(out_or.value().size(), 9U);

  const std::vector<std::string> expected_authors = {"", "", "", "", "", "", "Alice", "Bob", "Bob"};
  const std::vector<std::string> expected_text = {"suffix",       "fraction", "negative", "overflow",    "missing",
                                                  "out of range", "zero",     "one",      "leading zero"};
  for (std::size_t i = 0; i < out_or.value().size(); ++i) {
    EXPECT_EQ(out_or.value()[i].row, 0U);
    EXPECT_EQ(out_or.value()[i].col, i);
    EXPECT_EQ(out_or.value()[i].author, expected_authors[i]);
    EXPECT_EQ(out_or.value()[i].text, expected_text[i]);
  }
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

TEST(CommentsRoundTrip, KeptShapeOutsideDrawingIdBlockIsRenumbered) {
  const auto shape_with = [](const char* id, std::uint32_t row) {
    CommentShape shape;
    shape.row = row;
    shape.xml = std::string("<v:shape id=\"") + id +
                "\" type=\"#_x0000_t202\"><x:ClientData ObjectType=\"Note\"><x:Row>0</x:Row>"
                "<x:Column>0</x:Column></x:ClientData></v:shape>";
    return shape;
  };
  // Drawing 2 owns ids 2049..2071; 1025 belongs to drawing 1, and the repeated 2050 collides.
  const std::vector<CommentShape> shapes = {shape_with("_x0000_s1025", 0), shape_with("_x0000_s2050", 1),
                                            shape_with("_x0000_s2050", 2)};
  const std::string vml = write_vml_drawing(shapes, 2U);
  EXPECT_EQ(vml.find("_x0000_s1025"), std::string::npos) << vml;
  const auto count = [&](const std::string& needle) {
    std::size_t n = 0;
    for (std::size_t at = vml.find(needle); at != std::string::npos; at = vml.find(needle, at + 1U)) {
      ++n;
    }
    return n;
  };
  EXPECT_EQ(count("id=\"_x0000_s2050\""), 1U) << vml;
  EXPECT_EQ(count("<v:shape id=\"_x0000_s20"), 3U) << vml;
}

TEST(CommentsRoundTrip, FormControlsSurviveARegeneratedDrawing) {
  const std::string checkbox =
      "<v:shape id=\"_x0000_s1026\" type=\"#_x0000_t201\"><x:ClientData ObjectType=\"Checkbox\">"
      "<x:Row>4</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>";
  const std::string button_type = "<v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\"/>";
  const std::string source =
      "<xml><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout>"
      "<v:shapetype id=\"_x0000_t202\"><v:path/></v:shapetype>" +
      button_type + NoteShape(0, 0, "visibility:hidden") + checkbox + "</xml>";
  const std::string others = vml_non_note_shapes(source);
  EXPECT_NE(others.find(button_type), std::string::npos) << others;
  EXPECT_NE(others.find(checkbox), std::string::npos) << others;
  EXPECT_EQ(others.find("ObjectType=\"Note\""), std::string::npos) << others;
  EXPECT_EQ(others.find("_x0000_t202"), std::string::npos) << others;

  // Regenerated as drawing 2, the checkbox keeps id 1026, which the
  // worksheet's controls cite, and drawing 1's block stays mapped.
  CommentShape note;
  note.row = 0;
  note.col = 0;
  const std::string vml = write_vml_drawing({note}, 2U, others);
  EXPECT_NE(vml.find(checkbox), std::string::npos) << vml;
  EXPECT_NE(vml.find("<o:idmap v:ext=\"edit\" data=\"2,1\"/>"), std::string::npos) << vml;
  EXPECT_NE(vml.find("<v:shape id=\"_x0000_s2049\""), std::string::npos) << vml;

  // In its own block the checkbox's id is skipped by new note shapes.
  const std::string same_block = write_vml_drawing({note}, 1U, others);
  EXPECT_NE(same_block.find("<o:idmap v:ext=\"edit\" data=\"1\"/>"), std::string::npos) << same_block;
  EXPECT_NE(same_block.find("<v:shape id=\"_x0000_s1025\""), std::string::npos) << same_block;
  EXPECT_EQ(same_block.find("_x0000_s1026"), same_block.rfind("_x0000_s1026")) << same_block;
}

}  // namespace
}  // namespace io
}  // namespace formulon
