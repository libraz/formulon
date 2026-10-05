#include "drawing/image_header.h"

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"

namespace formulon {
namespace {

using Bytes = std::vector<std::uint8_t>;

Bytes Png(std::uint32_t w, std::uint32_t h) {
  Bytes b = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  for (std::uint32_t v : {w, h}) {
    b.push_back(static_cast<std::uint8_t>(v >> 24));
    b.push_back(static_cast<std::uint8_t>(v >> 16));
    b.push_back(static_cast<std::uint8_t>(v >> 8));
    b.push_back(static_cast<std::uint8_t>(v));
  }
  b.insert(b.end(), {8, 6, 0, 0, 0, 0, 0, 0, 0});
  return b;
}

ImageInfo Probe(const Bytes& b) {
  auto info = probe_image(b.data(), b.size());
  EXPECT_TRUE(static_cast<bool>(info));
  return info ? info.value() : ImageInfo{};
}

FormulonErrorCode ProbeError(const Bytes& b) {
  auto info = probe_image(b.data(), b.size());
  EXPECT_FALSE(static_cast<bool>(info));
  return info ? FormulonErrorCode::kOk : info.error().code;
}

TEST(ImageHeader, PngReadsIhdr) {
  const ImageInfo info = Probe(Png(640, 480));
  EXPECT_EQ(info.format, ImageFormat::kPng);
  EXPECT_EQ(info.px_width, 640U);
  EXPECT_EQ(info.px_height, 480U);
}

TEST(ImageHeader, JpegReadsBaselineFrameAfterApp0) {
  // SOI, APP0 (JFIF, length 16), SOF0: precision 8, height 0x01E0, width 0x0280.
  Bytes b = {0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 'J',  'F',  'I',  'F',  0,    1,    1,    0,    0,
             1,    0,    1,    0,    0,    0xFF, 0xC0, 0x00, 0x11, 0x08, 0x01, 0xE0, 0x02, 0x80, 0x03};
  const ImageInfo info = Probe(b);
  EXPECT_EQ(info.format, ImageFormat::kJpeg);
  EXPECT_EQ(info.px_width, 640U);
  EXPECT_EQ(info.px_height, 480U);
}

TEST(ImageHeader, JpegSkipsHuffmanTableAndFillBytes) {
  // A DHT segment (0xC4 is not a frame header) and a fill byte precede SOF2.
  Bytes b = {0xFF, 0xD8, 0xFF, 0xC4, 0x00, 0x04, 0x00, 0x00, 0xFF,
             0xFF, 0xC2, 0x00, 0x11, 0x08, 0x00, 0x20, 0x00, 0x10};
  const ImageInfo info = Probe(b);
  EXPECT_EQ(info.format, ImageFormat::kJpeg);
  EXPECT_EQ(info.px_width, 16U);
  EXPECT_EQ(info.px_height, 32U);
}

TEST(ImageHeader, GifReadsScreenDescriptor) {
  Bytes b = {'G', 'I', 'F', '8', '9', 'a', 0x2C, 0x01, 0x64, 0x00, 0x00, 0x00, 0x00};
  const ImageInfo info = Probe(b);
  EXPECT_EQ(info.format, ImageFormat::kGif);
  EXPECT_EQ(info.px_width, 300U);
  EXPECT_EQ(info.px_height, 100U);
}

TEST(ImageHeader, BmpInfoHeaderTopDown) {
  Bytes b(54, 0);
  b[0] = 'B';
  b[1] = 'M';
  b[14] = 40;    // BITMAPINFOHEADER
  b[18] = 0x20;  // width 32
  // height -16 (top-down), little endian
  b[22] = 0xF0;
  b[23] = 0xFF;
  b[24] = 0xFF;
  b[25] = 0xFF;
  const ImageInfo info = Probe(b);
  EXPECT_EQ(info.format, ImageFormat::kBmp);
  EXPECT_EQ(info.px_width, 32U);
  EXPECT_EQ(info.px_height, 16U);
}

TEST(ImageHeader, BmpCoreHeader) {
  Bytes b(26, 0);
  b[0] = 'B';
  b[1] = 'M';
  b[14] = 12;  // BITMAPCOREHEADER
  b[18] = 0x05;
  b[20] = 0x07;
  const ImageInfo info = Probe(b);
  EXPECT_EQ(info.format, ImageFormat::kBmp);
  EXPECT_EQ(info.px_width, 5U);
  EXPECT_EQ(info.px_height, 7U);
}

TEST(ImageHeader, TruncatedHeadersAreUnsupported) {
  Bytes png = Png(1, 1);
  png.resize(20);
  EXPECT_EQ(ProbeError(png), FormulonErrorCode::kIoImageUnsupported);
  EXPECT_EQ(ProbeError({0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x11, 0x08}), FormulonErrorCode::kIoImageUnsupported);
  EXPECT_EQ(ProbeError({'G', 'I', 'F', '8', '9', 'a', 1}), FormulonErrorCode::kIoImageUnsupported);
  EXPECT_EQ(ProbeError({'B', 'M', 0, 0}), FormulonErrorCode::kIoImageUnsupported);
}

TEST(ImageHeader, JpegWithoutFrameHeaderIsUnsupported) {
  // End of image, then start of scan, before any frame header.
  EXPECT_EQ(ProbeError({0xFF, 0xD8, 0xFF, 0xD9}), FormulonErrorCode::kIoImageUnsupported);
  EXPECT_EQ(ProbeError({0xFF, 0xD8, 0xFF, 0xDA, 0x00, 0x08, 0, 0, 0, 0, 0, 0}), FormulonErrorCode::kIoImageUnsupported);
}

TEST(ImageHeader, UnknownBytesAreUnsupported) {
  EXPECT_EQ(ProbeError({'h', 'e', 'l', 'l', 'o'}), FormulonErrorCode::kIoImageUnsupported);
  EXPECT_EQ(ProbeError({}), FormulonErrorCode::kIoImageUnsupported);
  EXPECT_FALSE(static_cast<bool>(probe_image(nullptr, 0)));
}

TEST(ImageHeader, ZeroDimensionIsUnsupported) {
  EXPECT_EQ(ProbeError(Png(0, 10)), FormulonErrorCode::kIoImageUnsupported);
}

TEST(ImageHeader, ExtensionsAndContentTypes) {
  EXPECT_STREQ(image_extension(ImageFormat::kPng), "png");
  EXPECT_STREQ(image_extension(ImageFormat::kJpeg), "jpeg");
  EXPECT_STREQ(image_content_type(ImageFormat::kJpeg), "image/jpeg");
  EXPECT_STREQ(image_content_type(ImageFormat::kGif), "image/gif");
  EXPECT_STREQ(image_extension(ImageFormat::kUnknown), "");
}

}  // namespace
}  // namespace formulon
