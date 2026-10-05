//
// Load, edit and re-serialize one XML package part through a single pugixml document.
//
// A part the engine edits without modelling it (the theme, threaded comments,
// drawings) lives in `Workbook::passthrough_parts()` as raw bytes. Editing one
// is always the same three steps: parse the bytes with the standard part parse
// flags, mutate the DOM, and write it back as an XML declaration followed by
// the raw (unindented) root element. These functions own the first and last
// step so every such edit parses and serializes identically:
//
//   pugi::xml_document doc;
//   RETURN_IF_ERROR(io::ooxml::load_part_dom(wb, path, "theme", doc));
//   ... mutate doc ...
//   return io::ooxml::store_part_dom(wb, path, doc);
//
// Creating an absent part (and registering its relationship) stays with the
// caller, since the content type and the owning relationship differ per part.

#ifndef FORMULON_IO_OOXML_PART_DOM_H_
#define FORMULON_IO_OOXML_PART_DOM_H_

#include <string_view>

#include "pugixml.hpp"
#include "utils/expected.h"

namespace formulon {

class Workbook;
struct PassthroughPart;

namespace io {
namespace ooxml {

/// Returns the passthrough part stored at package path `path`, or NULL.
const PassthroughPart* find_passthrough_part(const Workbook& wb, std::string_view path);

/// Parses the passthrough part at `path` into `doc` with the standard part
/// parse flags. `reader_module` names the caller in a parse error's context.
/// Fails with `kInvalidArgument` when no such part exists and with
/// `kIoXmlParse` when its bytes are not well-formed XML.
Expected<void, Error> load_part_dom(const Workbook& wb, std::string_view path, std::string_view reader_module,
                                    pugi::xml_document& doc);

/// Serializes `doc` as the standalone UTF-8 XML declaration plus its root
/// element in raw form and replaces the bytes of the passthrough part at
/// `path`. Fails with `kInvalidArgument` when no such part exists.
Expected<void, Error> store_part_dom(Workbook& wb, std::string_view path, const pugi::xml_document& doc);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_PART_DOM_H_
