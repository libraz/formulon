// `<sheetFormatPr>` attributes survive an xlsx save -> load round trip.

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace {

TEST(SheetFormatPrRoundTrip, AllAttributesSurviveSaveAndLoad) {
  Workbook src = Workbook::create();
  SheetFormatDefaults& defaults = src.sheet(0).mutable_format_defaults();
  defaults.has_default_row_height = true;
  defaults.default_row_height = 15.75;
  defaults.has_default_col_width = true;
  defaults.default_col_width = 9.5;
  defaults.base_col_width = 10.0;
  defaults.custom_height = true;
  defaults.zero_height = true;
  defaults.thick_top = true;
  defaults.thick_bottom = true;
  defaults.outline_level_row = 2U;
  defaults.outline_level_col = 3U;

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << save_or.error().message;
  const std::vector<std::uint8_t>& bytes = save_or.value();

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{bytes.data(), bytes.size()})));
  auto entry_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(entry_or));
  const std::string xml(reinterpret_cast<const char*>(entry_or.value().data()), entry_or.value().size());
  EXPECT_NE(xml.find("customHeight=\"1\""), std::string::npos) << xml;

  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  const SheetFormatDefaults& out = read_or.value().workbook.sheet(0).format_defaults();
  EXPECT_TRUE(out.has_default_row_height);
  EXPECT_DOUBLE_EQ(out.default_row_height, 15.75);
  EXPECT_TRUE(out.has_default_col_width);
  EXPECT_DOUBLE_EQ(out.default_col_width, 9.5);
  EXPECT_DOUBLE_EQ(out.base_col_width, 10.0);
  EXPECT_TRUE(out.custom_height);
  EXPECT_TRUE(out.zero_height);
  EXPECT_TRUE(out.thick_top);
  EXPECT_TRUE(out.thick_bottom);
  EXPECT_EQ(out.outline_level_row, 2U);
  EXPECT_EQ(out.outline_level_col, 3U);
}

TEST(SheetFormatPrRoundTrip, CustomHeightAloneIsKept) {
  Workbook src = Workbook::create();
  SheetFormatDefaults& defaults = src.sheet(0).mutable_format_defaults();
  defaults.has_default_row_height = true;
  defaults.default_row_height = 15.75;
  defaults.custom_height = true;

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto read_or = io::read_ooxml(io::ByteSpan{save_or.value().data(), save_or.value().size()});
  ASSERT_TRUE(static_cast<bool>(read_or));
  const SheetFormatDefaults& out = read_or.value().workbook.sheet(0).format_defaults();
  EXPECT_TRUE(out.custom_height);
  EXPECT_FALSE(out.zero_height);
  EXPECT_FALSE(out.thick_top);
  EXPECT_EQ(out.outline_level_row, 0U);
}

}  // namespace
}  // namespace formulon
