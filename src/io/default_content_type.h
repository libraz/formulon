//
// `DefaultContentType`: one `<Default Extension="..." ContentType="..."/>`
// entry from `[Content_Types].xml`. Excel declares binary and media
// parts (vbaProject.bin, images, VML, printer settings) by extension
// via `<Default>` rather than a per-part `<Override>`. The reader
// captures these so the writer can re-emit the matching `<Default>`
// registration for any Default-typed part that round-trips through
// `Workbook::passthrough_parts()`.
//
// Lives in its own lean header (mirroring `passthrough_part.h` and
// `unknown_relationship.h`) so both the OOXML reader (which produces
// these) and `Workbook` (which carries them through the round-trip) can
// include the type without dragging in the pugixml-backed content-types
// parser.

#ifndef FORMULON_IO_DEFAULT_CONTENT_TYPE_H_
#define FORMULON_IO_DEFAULT_CONTENT_TYPE_H_

#include <string>

namespace formulon {
namespace io {

/// One `<Default>` entry from `[Content_Types].xml`.
///
///   * `extension`    — the file extension the entry applies to, lower
///                      case with no leading dot (e.g. `"bin"`, `"png"`).
///   * `content_type` — the `ContentType=` attribute value verbatim.
struct DefaultContentType {
  std::string extension;
  std::string content_type;
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_DEFAULT_CONTENT_TYPE_H_
