// C ABI tests for pictures placed in cells: fm_sheet_get_cell_picture / fm_sheet_set_cell_picture.

#include <cstddef>
#include <cstdint>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr std::size_t kMaxPictureBytes = std::size_t{32} * 1024U * 1024U;

Bytes Png(std::uint32_t w, std::uint32_t h) {
  Bytes b = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  for (std::uint32_t v : {w, h}) {
    for (int shift = 24; shift >= 0; shift -= 8) {
      b.push_back(static_cast<std::uint8_t>(v >> shift));
    }
  }
  b.insert(b.end(), {8, 6, 0, 0, 0, 0x1F, 0x15, 0xC4, 0x89});
  return b;
}

Bytes Jpeg(std::uint32_t w, std::uint32_t h) {
  return {0xFF,
          0xD8,
          0xFF,
          0xC0,
          0x00,
          0x11,
          0x08,
          static_cast<std::uint8_t>(h >> 8),
          static_cast<std::uint8_t>(h),
          static_cast<std::uint8_t>(w >> 8),
          static_cast<std::uint8_t>(w),
          0x03,
          0x01,
          0x11,
          0x00,
          0x02,
          0x11,
          0x01,
          0x03,
          0x11,
          0x01,
          0xFF,
          0xD9};
}

Bytes Gif() {
  return {'G', 'I', 'F', '8', '9', 'a', 4, 0, 4, 0, 0, 0, 0, ';'};
}

fm_cell_picture_t Picture(const Bytes& bytes, const char* alt_text = nullptr) {
  fm_cell_picture_t picture = FM_STRUCT_INIT(fm_cell_picture_t);
  picture.bytes = bytes.data();
  picture.len = bytes.size();
  picture.alt_text = alt_text;
  return picture;
}

class FormulonCCellPicture : public ::testing::Test {
 protected:
  void SetUp() override { ASSERT_EQ(fm_workbook_create_default(&wb_.handle), 0); }

  fm_workbook_t* wb() { return wb_.handle; }

  fm_status_t Set(std::uint32_t row, std::uint32_t col, const Bytes& bytes, const char* alt_text = nullptr) {
    const fm_cell_picture_t picture = Picture(bytes, alt_text);
    return fm_sheet_set_cell_picture(wb(), 0, row, col, &picture);
  }

  bool Has(std::uint32_t row, std::uint32_t col) {
    std::int32_t present = -1;
    fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
    EXPECT_EQ(fm_sheet_get_cell_picture(wb(), 0, row, col, &present, &out), 0) << fm_get_last_error_message();
    return present == 1;
  }

  WorkbookGuard wb_;
};

TEST_F(FormulonCCellPicture, SetsAndReadsAPngWithAltText) {
  const Bytes png = Png(16, 8);
  ASSERT_EQ(Set(1, 2, png, "A red square"), 0) << fm_get_last_error_message();

  std::int32_t present = 0;
  fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
  ASSERT_EQ(fm_sheet_get_cell_picture(wb(), 0, 1, 2, &present, &out), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  ASSERT_EQ(out.len, png.size());
  EXPECT_EQ(Bytes(out.bytes, out.bytes + out.len), png);
  ASSERT_NE(out.alt_text, nullptr);
  EXPECT_STREQ(out.alt_text, "A red square");
  EXPECT_EQ(out.format, FM_IMAGE_FORMAT_PNG);
  EXPECT_EQ(out.width_px, 16U);
  EXPECT_EQ(out.height_px, 8U);
}

TEST_F(FormulonCCellPicture, SetsAndReadsAJpegWithoutAltText) {
  const Bytes jpeg = Jpeg(40, 24);
  ASSERT_EQ(Set(0, 0, jpeg), 0) << fm_get_last_error_message();

  std::int32_t present = 0;
  fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
  ASSERT_EQ(fm_sheet_get_cell_picture(wb(), 0, 0, 0, &present, &out), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  EXPECT_EQ(Bytes(out.bytes, out.bytes + out.len), jpeg);
  EXPECT_EQ(out.alt_text, nullptr);
  EXPECT_EQ(out.format, FM_IMAGE_FORMAT_JPEG);
  EXPECT_EQ(out.width_px, 40U);
  EXPECT_EQ(out.height_px, 24U);
}

TEST_F(FormulonCCellPicture, TheStoredPictureIsACopyAndTheReadOnlyFieldsOfInputAreIgnored) {
  Bytes png = Png(4, 4);
  const Bytes original = png;
  fm_cell_picture_t picture = Picture(png, "alt");
  picture.format = FM_IMAGE_FORMAT_GIF;
  picture.width_px = 999;
  picture.height_px = 999;
  ASSERT_EQ(fm_sheet_set_cell_picture(wb(), 0, 0, 0, &picture), 0) << fm_get_last_error_message();
  png.assign(png.size(), 0);

  std::int32_t present = 0;
  fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
  ASSERT_EQ(fm_sheet_get_cell_picture(wb(), 0, 0, 0, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(Bytes(out.bytes, out.bytes + out.len), original);
  EXPECT_EQ(out.format, FM_IMAGE_FORMAT_PNG);
  EXPECT_EQ(out.width_px, 4U);
  EXPECT_EQ(out.height_px, 4U);
}

TEST_F(FormulonCCellPicture, TheViewsSurviveUnrelatedReadsUntilTheNextPictureRead) {
  const Bytes png = Png(8, 8);
  ASSERT_EQ(Set(0, 0, png, "first"), 0);
  ASSERT_EQ(Set(1, 0, Jpeg(2, 2), "second"), 0);

  std::int32_t present = 0;
  fm_cell_picture_t first = FM_STRUCT_INIT(fm_cell_picture_t);
  ASSERT_EQ(fm_sheet_get_cell_picture(wb(), 0, 0, 0, &present, &first), 0);
  ASSERT_EQ(present, 1);

  fm_value_t value{};
  EXPECT_EQ(fm_sheet_get_value(wb(), 0, 5, 5, &value), 0);
  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  EXPECT_EQ(fm_sheet_get_hyperlink(wb(), 0, 5, 5, &present, &link), 0);
  EXPECT_EQ(Bytes(first.bytes, first.bytes + first.len), png);
  EXPECT_STREQ(first.alt_text, "first");

  fm_cell_picture_t second = FM_STRUCT_INIT(fm_cell_picture_t);
  ASSERT_EQ(fm_sheet_get_cell_picture(wb(), 0, 1, 0, &present, &second), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(second.alt_text, "second");
  EXPECT_EQ(second.format, FM_IMAGE_FORMAT_JPEG);
}

TEST_F(FormulonCCellPicture, ACellWithoutAPictureReadsAbsent) {
  std::int32_t present = 7;
  fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
  out.len = 5;
  out.format = 3;
  ASSERT_EQ(fm_sheet_get_cell_picture(wb(), 0, 3, 3, &present, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 0);
  EXPECT_EQ(out.struct_size, sizeof(fm_cell_picture_t));
  EXPECT_EQ(out.bytes, nullptr);
  EXPECT_EQ(out.len, 0U);
  EXPECT_EQ(out.alt_text, nullptr);
  EXPECT_EQ(out.format, 0);
  EXPECT_EQ(out.width_px, 0U);
  EXPECT_EQ(out.height_px, 0U);

  // A cell holding an ordinary error value is not a picture cell.
  ASSERT_EQ(fm_sheet_set_error(wb(), 0, 3, 3, 2), 0);
  EXPECT_FALSE(Has(3, 3));
}

TEST_F(FormulonCCellPicture, NullRemovesThePictureAndEmptiesTheCell) {
  ASSERT_EQ(Set(2, 2, Png(3, 3), "alt"), 0);
  ASSERT_TRUE(Has(2, 2));

  ASSERT_EQ(fm_sheet_set_cell_picture(wb(), 0, 2, 2, nullptr), 0) << fm_get_last_error_message();
  EXPECT_FALSE(Has(2, 2));
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb(), 0, 2, 2, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_BLANK);

  // A cell without a picture is left as it is.
  ASSERT_EQ(fm_sheet_set_number(wb(), 0, 4, 4, 12.5), 0);
  ASSERT_EQ(fm_sheet_set_cell_picture(wb(), 0, 4, 4, nullptr), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_value(wb(), 0, 4, 4, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(value.u.number, 12.5);
}

TEST_F(FormulonCCellPicture, ASecondPictureReplacesTheFirstAndAValueWriteDropsIt) {
  ASSERT_EQ(Set(0, 0, Png(2, 2), "one"), 0);
  ASSERT_EQ(Set(0, 0, Jpeg(6, 4)), 0);
  std::int32_t present = 0;
  fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
  ASSERT_EQ(fm_sheet_get_cell_picture(wb(), 0, 0, 0, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(out.format, FM_IMAGE_FORMAT_JPEG);
  EXPECT_EQ(out.alt_text, nullptr);

  ASSERT_EQ(fm_sheet_set_number(wb(), 0, 0, 0, 3.0), 0);
  EXPECT_FALSE(Has(0, 0));
}

TEST_F(FormulonCCellPicture, APictureReplacesAFormula) {
  ASSERT_EQ(fm_sheet_set_formula(wb(), 0, 0, 0, "=1+1"), 0);
  ASSERT_EQ(Set(0, 0, Png(2, 2)), 0) << fm_get_last_error_message();
  EXPECT_TRUE(Has(0, 0));
}

TEST_F(FormulonCCellPicture, AFormulaReadingAPictureCellGetsValueError) {
  ASSERT_EQ(Set(0, 0, Png(2, 2), "alt"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb(), 0, 0, 1, "=A1"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb(), 0, 0, 2, "=ISERROR(A1)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb()), 0) << fm_get_last_error_message();

  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_value(wb(), 0, 0, 0, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(value.u.error_code, 2);
  ASSERT_EQ(fm_sheet_get_value(wb(), 0, 0, 1, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(value.u.error_code, 2);
  ASSERT_EQ(fm_sheet_get_value(wb(), 0, 0, 2, &value), 0);
  EXPECT_EQ(value.kind, FM_VALUE_KIND_BOOL);
  EXPECT_EQ(value.u.boolean, 1);
}

TEST_F(FormulonCCellPicture, RefusesAnImageThatIsNotPngOrJpeg) {
  const std::uint8_t one = 0;
  fm_cell_picture_t empty = FM_STRUCT_INIT(fm_cell_picture_t);
  empty.bytes = &one;
  empty.len = 0;
  EXPECT_EQ(fm_sheet_set_cell_picture(wb(), 0, 0, 0, &empty), FM_STATUS_IO_IMAGE_UNSUPPORTED);
  EXPECT_EQ(Set(0, 0, Gif()), FM_STATUS_IO_IMAGE_UNSUPPORTED);
  EXPECT_EQ(Set(0, 0, Bytes{'n', 'o', 't', ' ', 'a', 'n', ' ', 'i', 'm', 'a', 'g', 'e'}),
            FM_STATUS_IO_IMAGE_UNSUPPORTED);
  EXPECT_FALSE(Has(0, 0));
}

TEST_F(FormulonCCellPicture, RefusesALengthOverTheCeilingWithoutReadingTheBytes) {
  const Bytes png = Png(2, 2);
  fm_cell_picture_t picture = Picture(png);
  picture.len = kMaxPictureBytes + 1U;
  EXPECT_EQ(fm_sheet_set_cell_picture(wb(), 0, 0, 0, &picture), FM_STATUS_SECURITY_RESOURCE_LIMIT);
  EXPECT_FALSE(Has(0, 0));
}

TEST_F(FormulonCCellPicture, RefusesEmptyAltTextAndMalformedUtf8) {
  const Bytes png = Png(2, 2);
  EXPECT_EQ(Set(0, 0, png, ""), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(Set(0, 0, png, "\xC3\x28"), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  EXPECT_FALSE(Has(0, 0));
}

TEST_F(FormulonCCellPicture, RefusalsLeaveAnExistingPictureAndTheNeighboursAlone) {
  const Bytes png = Png(5, 5);
  ASSERT_EQ(Set(0, 0, png, "keep"), 0);
  EXPECT_EQ(Set(0, 0, Gif(), "gif"), FM_STATUS_IO_IMAGE_UNSUPPORTED);
  EXPECT_EQ(Set(0, 0, png, ""), FM_STATUS_INVALID_ARGUMENT);

  std::int32_t present = 0;
  fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
  ASSERT_EQ(fm_sheet_get_cell_picture(wb(), 0, 0, 0, &present, &out), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(Bytes(out.bytes, out.bytes + out.len), png);
  EXPECT_STREQ(out.alt_text, "keep");
}

TEST_F(FormulonCCellPicture, ValidatesRecordSizeNullPointersAndAddresses) {
  const Bytes png = Png(2, 2);
  fm_cell_picture_t picture = Picture(png);

  picture.struct_size = 4;
  EXPECT_EQ(fm_sheet_set_cell_picture(wb(), 0, 0, 0, &picture), FM_STATUS_INVALID_ARGUMENT);
  picture = Picture(png);
  picture.bytes = nullptr;
  EXPECT_EQ(fm_sheet_set_cell_picture(wb(), 0, 0, 0, &picture), FM_STATUS_BINDING_NULL_POINTER);
  picture = Picture(png);
  EXPECT_EQ(fm_sheet_set_cell_picture(nullptr, 0, 0, 0, &picture), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_set_cell_picture(wb(), 9, 0, 0, &picture), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_cell_picture(wb(), 9, 0, 0, nullptr), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_cell_picture(wb(), 0, 1048576U, 0, &picture), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_set_cell_picture(wb(), 0, 0, 16384U, nullptr), FM_STATUS_INVALID_ARGUMENT);

  std::int32_t present = 0;
  fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
  EXPECT_EQ(fm_sheet_get_cell_picture(nullptr, 0, 0, 0, &present, &out), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_cell_picture(wb(), 0, 0, 0, nullptr, &out), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_cell_picture(wb(), 0, 0, 0, &present, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_cell_picture(wb(), 9, 0, 0, &present, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  out.struct_size = 4;
  EXPECT_EQ(fm_sheet_get_cell_picture(wb(), 0, 0, 0, &present, &out), FM_STATUS_INVALID_ARGUMENT);
  out = FM_STRUCT_INIT(fm_cell_picture_t);
  EXPECT_EQ(fm_sheet_get_cell_picture(wb(), 0, 1048576U, 0, &present, &out), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_FALSE(Has(0, 0));
}

TEST_F(FormulonCCellPicture, SurvivesASaveAndLoadInBothFormats) {
  const Bytes png = Png(12, 9);
  const Bytes jpeg = Jpeg(30, 20);
  ASSERT_EQ(Set(1, 1, png, "a chart"), 0);
  ASSERT_EQ(Set(2, 3, jpeg), 0);

  for (const std::int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    SCOPED_TRACE(format == FM_WORKBOOK_FORMAT_XLSX ? "xlsx" : "xlsb");
    BufferGuard saved;
    ASSERT_EQ(fm_workbook_save(wb(), format, &saved.data, &saved.len), 0) << fm_get_last_error_message();
    WorkbookGuard loaded;
    ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0) << fm_get_last_error_message();

    std::int32_t present = 0;
    fm_cell_picture_t out = FM_STRUCT_INIT(fm_cell_picture_t);
    ASSERT_EQ(fm_sheet_get_cell_picture(loaded.handle, 0, 1, 1, &present, &out), 0) << fm_get_last_error_message();
    ASSERT_EQ(present, 1);
    EXPECT_EQ(Bytes(out.bytes, out.bytes + out.len), png);
    ASSERT_NE(out.alt_text, nullptr);
    EXPECT_STREQ(out.alt_text, "a chart");
    EXPECT_EQ(out.format, FM_IMAGE_FORMAT_PNG);

    out = FM_STRUCT_INIT(fm_cell_picture_t);
    ASSERT_EQ(fm_sheet_get_cell_picture(loaded.handle, 0, 2, 3, &present, &out), 0) << fm_get_last_error_message();
    ASSERT_EQ(present, 1);
    EXPECT_EQ(Bytes(out.bytes, out.bytes + out.len), jpeg);
    EXPECT_EQ(out.alt_text, nullptr);
    EXPECT_EQ(out.format, FM_IMAGE_FORMAT_JPEG);

    fm_value_t value{};
    ASSERT_EQ(fm_sheet_get_value(loaded.handle, 0, 1, 1, &value), 0);
    EXPECT_EQ(value.kind, FM_VALUE_KIND_ERROR);
    EXPECT_EQ(value.u.error_code, 2);

    present = 7;
    ASSERT_EQ(fm_sheet_get_cell_picture(loaded.handle, 0, 0, 0, &present, &out), 0);
    EXPECT_EQ(present, 0);
  }
}

}  // namespace
