#include "drawing/image_header.h"

#include <cstring>

namespace formulon {
namespace {

std::uint32_t be16(const std::uint8_t* p) {
  return (static_cast<std::uint32_t>(p[0]) << 8) | p[1];
}
std::uint32_t be32(const std::uint8_t* p) {
  return (be16(p) << 16) | be16(p + 2);
}
std::uint32_t le16(const std::uint8_t* p) {
  return (static_cast<std::uint32_t>(p[1]) << 8) | p[0];
}
std::uint32_t le32(const std::uint8_t* p) {
  return (le16(p + 2) << 16) | le16(p);
}

// Walks the JPEG marker segments up to the first start-of-frame marker.
bool probe_jpeg(const std::uint8_t* b, std::size_t len, ImageInfo& out) {
  std::size_t pos = 2;
  while (pos + 1 < len) {
    if (b[pos] != 0xFF) {
      return false;
    }
    const std::uint8_t marker = b[pos + 1];
    if (marker == 0xFF) {
      ++pos;  // Fill byte.
      continue;
    }
    pos += 2;
    if (marker == 0x01 || (marker >= 0xD0 && marker <= 0xD8)) {
      continue;  // Standalone markers carry no length.
    }
    if (marker == 0xD9 || marker == 0xDA || pos + 2 > len) {
      return false;  // End of image or scan data before any frame header.
    }
    const std::uint32_t seg = be16(b + pos);
    if (seg < 2) {
      return false;
    }
    const bool sof = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
    if (sof) {
      if (seg < 7 || pos + 7 > len) {
        return false;
      }
      out.format = ImageFormat::kJpeg;
      out.px_height = be16(b + pos + 3);
      out.px_width = be16(b + pos + 5);
      return true;
    }
    pos += seg;
  }
  return false;
}

bool probe(const std::uint8_t* b, std::size_t len, ImageInfo& out) {
  static constexpr std::uint8_t kPngSig[8] = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A};
  if (len >= 24 && std::memcmp(b, kPngSig, 8) == 0 && std::memcmp(b + 12, "IHDR", 4) == 0) {
    out = {ImageFormat::kPng, be32(b + 16), be32(b + 20)};
    return true;
  }
  if (len >= 3 && b[0] == 0xFF && b[1] == 0xD8 && b[2] == 0xFF) {
    return probe_jpeg(b, len, out);
  }
  if (len >= 10 && (std::memcmp(b, "GIF87a", 6) == 0 || std::memcmp(b, "GIF89a", 6) == 0)) {
    out = {ImageFormat::kGif, le16(b + 6), le16(b + 8)};
    return true;
  }
  if (len >= 26 && b[0] == 'B' && b[1] == 'M') {
    const std::uint32_t dib = le32(b + 14);
    if (dib == 12) {
      out = {ImageFormat::kBmp, le16(b + 18), le16(b + 20)};
      return true;
    }
    if (dib >= 40) {
      // A negative height marks a top-down bitmap.
      const auto w = static_cast<std::int32_t>(le32(b + 18));
      const auto h = static_cast<std::int32_t>(le32(b + 22));
      if (w <= 0 || h == INT32_MIN) {
        return false;
      }
      out = {ImageFormat::kBmp, static_cast<std::uint32_t>(w), static_cast<std::uint32_t>(h < 0 ? -h : h)};
      return true;
    }
  }
  return false;
}

}  // namespace

Expected<ImageInfo, Error> probe_image(const std::uint8_t* bytes, std::size_t len) {
  ImageInfo info;
  if (bytes == nullptr || !probe(bytes, len, info) || info.px_width == 0 || info.px_height == 0) {
    return make_error(FormulonErrorCode::kIoImageUnsupported, "image: unrecognised or truncated image header");
  }
  return info;
}

const char* image_extension(ImageFormat format) {
  switch (format) {
    case ImageFormat::kPng:
      return "png";
    case ImageFormat::kJpeg:
      return "jpeg";
    case ImageFormat::kGif:
      return "gif";
    case ImageFormat::kBmp:
      return "bmp";
    case ImageFormat::kUnknown:
      break;
  }
  return "";
}

const char* image_content_type(ImageFormat format) {
  switch (format) {
    case ImageFormat::kPng:
      return "image/png";
    case ImageFormat::kJpeg:
      return "image/jpeg";
    case ImageFormat::kGif:
      return "image/gif";
    case ImageFormat::kBmp:
      return "image/bmp";
    case ImageFormat::kUnknown:
      break;
  }
  return "";
}

}  // namespace formulon
