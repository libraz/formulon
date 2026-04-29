// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `formulon::io::read_styles`. The reader is a minimal
// validator at this slice — it only needs to parse the document and
// expose two coarse counts so the OOXML reader can stop surfacing
// `xl/styles.xml` as an unknown part. Full numFmt/font/fill expansion
// arrives with the formatter pipeline.

#include "io/styles_reader.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "utils/error.h"

namespace formulon {
namespace io {
namespace {

std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

TEST(StylesReader, EmptyStyleSheetSelfClosing) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read failed: " << result_or.error().message;
  EXPECT_EQ(result_or.value().cell_xfs_count, 0U);
  EXPECT_EQ(result_or.value().num_fmts_count, 0U);
}

TEST(StylesReader, EmptyStyleSheetWithChildren) {
  // The minimal styleSheet our writer emits today: empty fonts/fills/
  // borders/cellStyleXfs/cellXfs blocks. Both counts should be 0 when
  // the elements are empty.
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>\n");
  xml.append("  <fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>\n");
  xml.append("  <borders count=\"1\"><border/></borders>\n");
  xml.append(
      "  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n");
  xml.append("  <cellXfs count=\"0\"></cellXfs>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.value().cell_xfs_count, 0U);
  EXPECT_EQ(result_or.value().num_fmts_count, 0U);
}

TEST(StylesReader, CountsCellXfs) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <cellXfs count=\"3\">\n");
  xml.append("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n");
  xml.append("    <xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"1\"/>\n");
  xml.append("    <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"1\"/>\n");
  xml.append("  </cellXfs>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.value().cell_xfs_count, 3U);
  EXPECT_EQ(result_or.value().num_fmts_count, 0U);
}

TEST(StylesReader, CountsCustomNumFmts) {
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <numFmts count=\"2\">\n");
  xml.append("    <numFmt numFmtId=\"164\" formatCode=\"0.0000\"/>\n");
  xml.append("    <numFmt numFmtId=\"165\" formatCode=\"yyyy/mm/dd\"/>\n");
  xml.append("  </numFmts>\n");
  xml.append("  <cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.value().num_fmts_count, 2U);
  EXPECT_EQ(result_or.value().cell_xfs_count, 1U);
}

TEST(StylesReader, RejectsMissingStyleSheetRoot) {
  std::string xml(kXmlDecl);
  xml.append("<notAStyleSheet/>");
  auto result_or = read_styles(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(StylesReader, RejectsMalformedXml) {
  std::string xml = "<styleSheet><cellXfs>";  // unterminated
  auto result_or = read_styles(Bytes(xml));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoXmlParse);
}

TEST(StylesReader, IgnoresUnknownChildren) {
  // dxfs / tableStyles / extLst are common in Excel-emitted styleSheets
  // but not interesting to this slice — they must not perturb the
  // counted blocks.
  std::string xml(kXmlDecl);
  xml.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  xml.append("  <cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs>\n");
  xml.append("  <dxfs count=\"0\"/>\n");
  xml.append("  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\"/>\n");
  xml.append("  <extLst><ext uri=\"{...}\"/></extLst>\n");
  xml.append("</styleSheet>");

  auto result_or = read_styles(Bytes(xml));
  ASSERT_TRUE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.value().cell_xfs_count, 1U);
  EXPECT_EQ(result_or.value().num_fmts_count, 0U);
}

}  // namespace
}  // namespace io
}  // namespace formulon
