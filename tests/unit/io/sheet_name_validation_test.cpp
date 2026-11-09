//
// Sheet-name validation at the file-read boundary.
//
// `sheet_by_name` / `sheet_index_by_name` resolve to the first match, so
// a workbook holding two sheets whose names fold together would answer
// every reference from one of them and compute a confident wrong number
// with no ambiguity signal a caller could observe. Both readers therefore
// validate the name as they append it, and a collision fails the load.
//
// The fixtures are produced by writing a workbook that was built through
// the name-unchecked `add_sheet` overload: that is the shape a crafted
// file has, and it exercises the real reader rather than a hand-rolled
// package.

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/format_detect.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return ByteSpan{bytes.data(), bytes.size()};
}

/// Builds a workbook whose two sheets differ only in case. `add_sheet`
/// is the name-unchecked overload, which is what lets the writer emit a
/// package the readers must now refuse.
Workbook BuildCollidingWorkbook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Data");
  wb.add_sheet("data");
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(2.0))));
  return wb;
}

Workbook BuildUnicodeCollidingWorkbook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("\xC3\x84");  // Ä
  wb.add_sheet("\xC3\xA4");  // ä
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(2.0))));
  return wb;
}

TEST(SheetNameValidation, OoxmlReaderRejectsCaseFoldedDuplicate) {
  const Workbook wb = BuildCollidingWorkbook();
  auto bytes_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;

  auto read_or = read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_FALSE(static_cast<bool>(read_or)) << "a duplicate sheet name must not load silently";
  EXPECT_EQ(read_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(SheetNameValidation, XlsbReaderRejectsCaseFoldedDuplicate) {
  const Workbook wb = BuildCollidingWorkbook();
  auto bytes_or = wb.save_as(WorkbookFormat::Xlsb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;

  auto read_or = xlsb::read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_FALSE(static_cast<bool>(read_or)) << "a duplicate sheet name must not load silently";
  EXPECT_EQ(read_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(SheetNameValidation, OoxmlReaderRejectsUnicodeSimpleFoldedDuplicate) {
  const Workbook wb = BuildUnicodeCollidingWorkbook();
  auto bytes_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;

  auto read_or = read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_FALSE(static_cast<bool>(read_or)) << "Ä and ä must not load as duplicate sheets";
  EXPECT_EQ(read_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(SheetNameValidation, XlsbReaderRejectsUnicodeSimpleFoldedDuplicate) {
  const Workbook wb = BuildUnicodeCollidingWorkbook();
  auto bytes_or = wb.save_as(WorkbookFormat::Xlsb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;

  auto read_or = xlsb::read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_FALSE(static_cast<bool>(read_or)) << "Ä and ä must not load as duplicate sheets";
  EXPECT_EQ(read_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(SheetNameValidation, DistinctUnicodeNamesLoadOnBothReaders) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("\xC3\x84");  // Ä
  wb.add_sheet("\xC3\x96");  // Ö
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(2.0))));

  auto xlsx_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(xlsx_or)) << xlsx_or.error().message;
  auto read_xlsx_or = read_ooxml(SpanOf(xlsx_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_xlsx_or)) << read_xlsx_or.error().message;
  EXPECT_EQ(read_xlsx_or.value().workbook.sheet_count(), 2U);

  auto xlsb_or = wb.save_as(WorkbookFormat::Xlsb);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message;
  auto read_xlsb_or = xlsb::read_xlsb(SpanOf(xlsb_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_xlsb_or)) << read_xlsb_or.error().message;
  EXPECT_EQ(read_xlsb_or.value().workbook.sheet_count(), 2U);
}

// The readers refuse case-only duplicates alone. Names that only the wider
// entry-time identity equates (width, kana type, prolonged-mark width) load
// side by side, as Excel files from a ja-JP UI can hold them.
TEST(SheetNameValidation, WidthAndKanaVariantsLoadOnBothReaders) {
  const std::vector<std::vector<std::string>> pairs = {
      {"Sheet1", "\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x91"},  // Ｓｈｅｅｔ１
      {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80", "\xE3\x83\x87\xE3\x83\xBC\xE3\x82\xBF"},  // ﾃﾞｰﾀ / データ
      {"\xE3\x82\xAB\xE3\x82\xBF", "\xE3\x81\x8B\xE3\x81\x9F"},                                      // カタ / かた
      {"\xE3\x82\xAC", "\xE3\x82\xAB\xE3\x82\x99"},  // ガ / カ + U+3099
  };
  for (const std::vector<std::string>& names : pairs) {
    SCOPED_TRACE(names[0] + " / " + names[1]);
    Workbook wb = Workbook::create_empty();
    wb.add_sheet(names[0]);
    wb.add_sheet(names[1]);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(2.0))));

    auto xlsx_or = wb.save();
    ASSERT_TRUE(static_cast<bool>(xlsx_or)) << xlsx_or.error().message;
    auto read_xlsx_or = read_ooxml(SpanOf(xlsx_or.value()));
    ASSERT_TRUE(static_cast<bool>(read_xlsx_or)) << read_xlsx_or.error().message;
    EXPECT_EQ(read_xlsx_or.value().workbook.sheet_count(), 2U);
    EXPECT_EQ(read_xlsx_or.value().workbook.sheet(1).name(), names[1]);

    auto xlsb_or = wb.save_as(WorkbookFormat::Xlsb);
    ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message;
    auto read_xlsb_or = xlsb::read_xlsb(SpanOf(xlsb_or.value()));
    ASSERT_TRUE(static_cast<bool>(read_xlsb_or)) << read_xlsb_or.error().message;
    EXPECT_EQ(read_xlsb_or.value().workbook.sheet_count(), 2U);
    EXPECT_EQ(read_xlsb_or.value().workbook.sheet(1).name(), names[1]);
  }
}

TEST(SheetNameValidation, DistinctNamesStillLoadOnBothReaders) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Data");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(2.0))));

  auto xlsx_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(xlsx_or)) << xlsx_or.error().message;
  auto read_xlsx_or = read_ooxml(SpanOf(xlsx_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_xlsx_or)) << read_xlsx_or.error().message;
  EXPECT_EQ(read_xlsx_or.value().workbook.sheet_count(), 2U);

  auto xlsb_or = wb.save_as(WorkbookFormat::Xlsb);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message;
  auto read_xlsb_or = xlsb::read_xlsb(SpanOf(xlsb_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_xlsb_or)) << read_xlsb_or.error().message;
  EXPECT_EQ(read_xlsb_or.value().workbook.sheet_count(), 2U);
}

}  // namespace
}  // namespace io
}  // namespace formulon
