//
// Image header sniffing: format and pixel size of PNG, JPEG, GIF and BMP bytes.
//
// Only the header is read (PNG IHDR, the first JPEG SOFn frame, the GIF
// logical screen descriptor, the BMP DIB header); the pixel data is never
// decoded, so a truncated body past the header still probes successfully.

#ifndef FORMULON_DRAWING_IMAGE_HEADER_H_
#define FORMULON_DRAWING_IMAGE_HEADER_H_

#include <cstddef>
#include <cstdint>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

/// Raster formats the drawing layer recognises.
enum class ImageFormat : std::int32_t {
  kUnknown = 0,
  kPng = 1,
  kJpeg = 2,
  kGif = 3,
  kBmp = 4,
};

/// Format and pixel size of an image.
struct ImageInfo {
  ImageFormat format = ImageFormat::kUnknown;
  std::uint32_t px_width = 0;
  std::uint32_t px_height = 0;
};

/// Sniffs the format and pixel size of `bytes`. Fails with
/// `kIoImageUnsupported` when the bytes are not one of the recognised
/// formats, the header is truncated, or it declares a zero dimension.
Expected<ImageInfo, Error> probe_image(const std::uint8_t* bytes, std::size_t len);

/// File extension Excel uses for `format`'s media part (`png`, `jpeg`, `gif`,
/// `bmp`), or an empty string for `kUnknown`.
const char* image_extension(ImageFormat format);

/// MIME content type registered for `format`'s extension, or an empty string
/// for `kUnknown`.
const char* image_content_type(ImageFormat format);

}  // namespace formulon

#endif  // FORMULON_DRAWING_IMAGE_HEADER_H_
