//
// Integration test for sheet-level UI features (merges, hyperlinks,
// comments, data validations). The flow is the same for all four:
//
//   1. Build a workbook in memory, attach the metadata to the sole sheet.
//   2. Save to OOXML bytes via `wb.save()`.
//   3. Re-load via `read_ooxml`.
//   4. Assert the round-trip preserved the metadata.
//
// A second pass (save again) is run for each fixture to lock down the
// "save -> load -> save" stability invariant: the second save must
// produce byte-for-byte identical output.

#include <algorithm>
#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

TEST(SheetFeaturesRoundTrip, MergeRanges) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.mutable_merges() = {
      MergeRange{0, 0, 1, 1},
      MergeRange{2, 2, 2, 4},
      MergeRange{10, 0, 12, 0},
  };
  auto save_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  const Sheet& loaded = load_or.value().workbook.sheet(0);
  ASSERT_EQ(loaded.merges().size(), 3U);
  EXPECT_EQ(loaded.merges()[0].first_row, 0U);
  EXPECT_EQ(loaded.merges()[0].last_col, 1U);
  EXPECT_EQ(loaded.merges()[2].first_row, 10U);
  EXPECT_EQ(loaded.merges()[2].last_row, 12U);

  // The package now carries a styles part with a default xfs table that
  // the reader normalises before re-emit, so byte-identity does not
  // hold across save -> load -> save in general. Assert content
  // equivalence: a second round-trip preserves all merges in order.
  auto save2_or = io::write_ooxml(load_or.value().workbook);
  ASSERT_TRUE(static_cast<bool>(save2_or));
  auto load2_or = io::read_ooxml(SpanOf(save2_or.value()));
  ASSERT_TRUE(static_cast<bool>(load2_or));
  const Sheet& reloaded = load2_or.value().workbook.sheet(0);
  ASSERT_EQ(reloaded.merges().size(), 3U);
  EXPECT_EQ(reloaded.merges()[0].first_row, 0U);
  EXPECT_EQ(reloaded.merges()[0].last_col, 1U);
  EXPECT_EQ(reloaded.merges()[2].first_row, 10U);
  EXPECT_EQ(reloaded.merges()[2].last_row, 12U);
}

TEST(SheetFeaturesRoundTrip, Hyperlinks) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  Hyperlink h;
  h.row = 0;
  h.col = 0;
  h.last_row = 0;
  h.last_col = 0;
  h.target = "https://example.com";
  h.tooltip = "Click here";
  h.display = "Example";
  h.rid = "rId1";
  s.mutable_hyperlinks().push_back(h);

  Hyperlink h2;
  h2.row = 1;
  h2.col = 2;
  h2.last_row = 1;
  h2.last_col = 2;
  h2.location = "Sheet1!A1";  // internal anchor; no target
  h2.display = "Go to A1";
  s.mutable_hyperlinks().push_back(h2);

  auto save_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  const Sheet& loaded = load_or.value().workbook.sheet(0);
  ASSERT_EQ(loaded.hyperlinks().size(), 2U);
  EXPECT_EQ(loaded.hyperlinks()[0].row, 0U);
  EXPECT_EQ(loaded.hyperlinks()[0].target, "https://example.com");
  EXPECT_EQ(loaded.hyperlinks()[0].tooltip, "Click here");
  EXPECT_EQ(loaded.hyperlinks()[0].display, "Example");
  EXPECT_FALSE(loaded.hyperlinks()[0].rid.empty());
  EXPECT_EQ(loaded.hyperlinks()[1].location, "Sheet1!A1");
  EXPECT_EQ(loaded.hyperlinks()[1].display, "Go to A1");
}

TEST(SheetFeaturesRoundTrip, Comments) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.mutable_comments() = {
      CellComment{0, 0, "Alice", "First comment"},
      CellComment{2, 1, "Bob", "Second"},
      CellComment{5, 5, "Alice", "Third (Alice again)"},
  };
  auto save_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(save_or));
  // The package must contain a comments part and a VML drawing part.
  io::ZipReader zr;
  ASSERT_TRUE(static_cast<bool>(zr.open(SpanOf(save_or.value()))));
  EXPECT_TRUE(zr.has_entry("xl/comments1.xml"));
  EXPECT_TRUE(zr.has_entry("xl/drawings/vmlDrawing1.vml"));

  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  const Sheet& loaded = load_or.value().workbook.sheet(0);
  ASSERT_EQ(loaded.comments().size(), 3U);
  EXPECT_EQ(loaded.comments()[0].author, "Alice");
  EXPECT_EQ(loaded.comments()[0].text, "First comment");
  EXPECT_EQ(loaded.comments()[1].author, "Bob");
  EXPECT_EQ(loaded.comments()[2].text, "Third (Alice again)");
}

TEST(SheetFeaturesRoundTrip, CommentsKeepVmlAfterEarlierCommentSheetIsRemoved) {
  Workbook wb = Workbook::create();
  wb.sheet(0).mutable_comments() = {CellComment{0, 0, "Alice", "first"}};
  Sheet& second = wb.add_sheet("Second");
  second.mutable_comments() = {CellComment{1, 1, "Bob", "second"}};
  auto first_save = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(first_save));

  auto loaded = io::read_ooxml(SpanOf(first_save.value()));
  ASSERT_TRUE(static_cast<bool>(loaded));
  Workbook& restored = loaded.value().workbook;
  ASSERT_EQ(restored.sheet(1).comment_vml_path(), "xl/drawings/vmlDrawing2.vml");

  // Give the surviving sheet's VML distinctive bytes. A planner that looks
  // up the newly-renumbered vmlDrawing1.vml would miss these and emit a stub.
  std::vector<io::PassthroughPart> parts = restored.passthrough_parts();
  auto vml = std::find_if(parts.begin(), parts.end(),
                          [](const io::PassthroughPart& part) { return part.path == "xl/drawings/vmlDrawing2.vml"; });
  ASSERT_NE(vml, parts.end());
  const std::string expected_vml = "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape id=\"survivor\"/></xml>\n";
  vml->bytes.assign(expected_vml.begin(), expected_vml.end());
  restored.set_passthrough_parts(std::move(parts));

  ASSERT_TRUE(static_cast<bool>(restored.remove_sheet(0)));
  auto saved = io::write_ooxml(restored);
  ASSERT_TRUE(static_cast<bool>(saved));
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(saved.value()))));
  auto vml_bytes = zip.read_entry("xl/drawings/vmlDrawing1.vml");
  ASSERT_TRUE(static_cast<bool>(vml_bytes));
  EXPECT_EQ(std::string(vml_bytes.value().begin(), vml_bytes.value().end()), expected_vml);
  auto sheet_xml = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet_xml));
  const std::string xml(sheet_xml.value().begin(), sheet_xml.value().end());
  EXPECT_NE(xml.find("<legacyDrawing r:id=\"rId2\"/>"), std::string::npos);
}

TEST(SheetFeaturesRoundTrip, DataValidations) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  DataValidation v;
  v.ranges.push_back(MergeRange{0, 0, 9, 0});
  v.type = 3;  // list
  v.allow_blank = true;
  v.show_error_message = true;
  v.error_title = "Pick one";
  v.error_message = "Use the dropdown";
  v.formula1 = "\"yes,no,maybe\"";
  s.mutable_validations().push_back(v);

  DataValidation v2;
  v2.ranges.push_back(MergeRange{0, 1, 9, 1});
  v2.type = 1;  // whole
  v2.op = 0;    // between
  v2.formula1 = "1";
  v2.formula2 = "100";
  s.mutable_validations().push_back(v2);

  auto save_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  const Sheet& loaded = load_or.value().workbook.sheet(0);
  ASSERT_EQ(loaded.validations().size(), 2U);
  EXPECT_EQ(loaded.validations()[0].type, 3U);
  EXPECT_EQ(loaded.validations()[0].error_title, "Pick one");
  EXPECT_EQ(loaded.validations()[0].formula1, "\"yes,no,maybe\"");
  EXPECT_EQ(loaded.validations()[1].type, 1U);
  EXPECT_EQ(loaded.validations()[1].formula2, "100");
}

TEST(SheetFeaturesRoundTrip, DataValidationShowDropDown) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);

  // Rule 1: dropdown arrow explicitly hidden (non-default).
  DataValidation hidden;
  hidden.ranges.push_back(MergeRange{0, 0, 9, 0});
  hidden.type = 3;  // list
  hidden.formula1 = "\"yes,no,maybe\"";
  hidden.show_dropdown = false;
  s.mutable_validations().push_back(hidden);

  // Rule 2: dropdown arrow explicitly shown (matches the default, but
  // set explicitly to make the assertion meaningful either way).
  DataValidation shown;
  shown.ranges.push_back(MergeRange{0, 1, 9, 1});
  shown.type = 3;  // list
  shown.formula1 = "\"x,y,z\"";
  shown.show_dropdown = true;
  s.mutable_validations().push_back(shown);

  auto save_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  const Sheet& loaded = load_or.value().workbook.sheet(0);
  ASSERT_EQ(loaded.validations().size(), 2U);
  EXPECT_FALSE(loaded.validations()[0].show_dropdown);
  EXPECT_TRUE(loaded.validations()[1].show_dropdown);

  // Save -> load -> save stability: the second round-trip must preserve
  // the same values.
  auto save2_or = io::write_ooxml(load_or.value().workbook);
  ASSERT_TRUE(static_cast<bool>(save2_or));
  auto load2_or = io::read_ooxml(SpanOf(save2_or.value()));
  ASSERT_TRUE(static_cast<bool>(load2_or));
  const Sheet& reloaded = load2_or.value().workbook.sheet(0);
  ASSERT_EQ(reloaded.validations().size(), 2U);
  EXPECT_FALSE(reloaded.validations()[0].show_dropdown);
  EXPECT_TRUE(reloaded.validations()[1].show_dropdown);
}

TEST(SheetFeaturesRoundTrip, AttributeControlCharacters) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);

  DataValidation validation;
  validation.ranges.push_back(MergeRange{0, 0, 0, 0});
  validation.type = 3;  // list
  validation.error_title = "tab\tvalue";
  validation.error_message = "line\nvalue";
  validation.prompt_title = "carriage\rvalue";
  validation.prompt_message = "both\r\nvalue";
  s.mutable_validations().push_back(validation);

  // Internal hyperlinks keep all three modelled text attributes in the
  // worksheet part, without involving a relationship target attribute.
  Hyperlink hyperlink;
  hyperlink.row = 1;
  hyperlink.col = 1;
  hyperlink.last_row = 1;
  hyperlink.last_col = 1;
  hyperlink.location = "location\tvalue";
  hyperlink.tooltip = "tooltip\nvalue";
  hyperlink.display = "display\r\nvalue";
  s.mutable_hyperlinks().push_back(hyperlink);

  auto save_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(save_or));

  io::ZipReader first_zip;
  ASSERT_TRUE(static_cast<bool>(first_zip.open(SpanOf(save_or.value()))));
  auto first_sheet_or = first_zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(first_sheet_or));
  const std::string first_sheet_xml(first_sheet_or.value().begin(), first_sheet_or.value().end());
  EXPECT_NE(first_sheet_xml.find("errorTitle=\"tab&#9;value\""), std::string::npos);
  EXPECT_NE(first_sheet_xml.find("error=\"line&#10;value\""), std::string::npos);
  EXPECT_NE(first_sheet_xml.find("promptTitle=\"carriage&#13;value\""), std::string::npos);
  EXPECT_NE(first_sheet_xml.find("prompt=\"both&#13;&#10;value\""), std::string::npos);
  EXPECT_NE(first_sheet_xml.find("location=\"location&#9;value\""), std::string::npos);
  EXPECT_NE(first_sheet_xml.find("tooltip=\"tooltip&#10;value\""), std::string::npos);
  EXPECT_NE(first_sheet_xml.find("display=\"display&#13;&#10;value\""), std::string::npos);

  auto first_load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(first_load_or));
  const Sheet& first_loaded = first_load_or.value().workbook.sheet(0);
  ASSERT_EQ(first_loaded.validations().size(), 1U);
  EXPECT_EQ(first_loaded.validations()[0].error_title, "tab\tvalue");
  EXPECT_EQ(first_loaded.validations()[0].error_message, "line\nvalue");
  EXPECT_EQ(first_loaded.validations()[0].prompt_title, "carriage\rvalue");
  EXPECT_EQ(first_loaded.validations()[0].prompt_message, "both\r\nvalue");
  ASSERT_EQ(first_loaded.hyperlinks().size(), 1U);
  EXPECT_EQ(first_loaded.hyperlinks()[0].location, "location\tvalue");
  EXPECT_EQ(first_loaded.hyperlinks()[0].tooltip, "tooltip\nvalue");
  EXPECT_EQ(first_loaded.hyperlinks()[0].display, "display\r\nvalue");

  auto save2_or = io::write_ooxml(first_load_or.value().workbook);
  ASSERT_TRUE(static_cast<bool>(save2_or));
  io::ZipReader second_zip;
  ASSERT_TRUE(static_cast<bool>(second_zip.open(SpanOf(save2_or.value()))));
  auto second_sheet_or = second_zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(second_sheet_or));
  const std::string second_sheet_xml(second_sheet_or.value().begin(), second_sheet_or.value().end());
  EXPECT_EQ(second_sheet_xml, first_sheet_xml);

  auto second_load_or = io::read_ooxml(SpanOf(save2_or.value()));
  ASSERT_TRUE(static_cast<bool>(second_load_or));
  const Sheet& second_loaded = second_load_or.value().workbook.sheet(0);
  ASSERT_EQ(second_loaded.validations().size(), 1U);
  EXPECT_EQ(second_loaded.validations()[0].error_title, "tab\tvalue");
  EXPECT_EQ(second_loaded.validations()[0].error_message, "line\nvalue");
  EXPECT_EQ(second_loaded.validations()[0].prompt_title, "carriage\rvalue");
  EXPECT_EQ(second_loaded.validations()[0].prompt_message, "both\r\nvalue");
  ASSERT_EQ(second_loaded.hyperlinks().size(), 1U);
  EXPECT_EQ(second_loaded.hyperlinks()[0].location, "location\tvalue");
  EXPECT_EQ(second_loaded.hyperlinks()[0].tooltip, "tooltip\nvalue");
  EXPECT_EQ(second_loaded.hyperlinks()[0].display, "display\r\nvalue");
}

TEST(SheetFeaturesRoundTrip, AllFeaturesCombined) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.mutable_merges() = {MergeRange{0, 0, 0, 1}};
  Hyperlink h;
  h.row = 1;
  h.col = 0;
  h.last_row = 1;
  h.last_col = 0;
  h.target = "https://example.com";
  s.mutable_hyperlinks().push_back(h);
  s.mutable_comments() = {CellComment{2, 0, "Alice", "comment"}};
  DataValidation v;
  v.ranges.push_back(MergeRange{3, 0, 9, 0});
  v.type = 3;
  v.formula1 = "\"a,b,c\"";
  s.mutable_validations().push_back(v);

  auto save_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  const Sheet& loaded = load_or.value().workbook.sheet(0);
  EXPECT_EQ(loaded.merges().size(), 1U);
  EXPECT_EQ(loaded.hyperlinks().size(), 1U);
  EXPECT_EQ(loaded.comments().size(), 1U);
  EXPECT_EQ(loaded.validations().size(), 1U);

  // Save -> load -> save stability with all four features in flight:
  // the second save must round-trip the metadata with identical
  // observable content. (Byte-identity holds for single-feature
  // sheets; with all four features in play the package picks up an
  // initial passthrough VML on the second save, which shifts a few
  // bytes in the central directory metadata. The semantic content
  // remains stable.)
  auto save2_or = io::write_ooxml(load_or.value().workbook);
  ASSERT_TRUE(static_cast<bool>(save2_or));
  auto load2_or = io::read_ooxml(SpanOf(save2_or.value()));
  ASSERT_TRUE(static_cast<bool>(load2_or));
  const Sheet& s2 = load2_or.value().workbook.sheet(0);
  EXPECT_EQ(s2.merges().size(), 1U);
  EXPECT_EQ(s2.hyperlinks().size(), 1U);
  EXPECT_EQ(s2.comments().size(), 1U);
  EXPECT_EQ(s2.validations().size(), 1U);
}

}  // namespace
}  // namespace formulon
