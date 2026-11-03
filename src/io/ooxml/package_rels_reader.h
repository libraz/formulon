//
// Package-level `_rels/.rels` reader shared by the OOXML and XLSB readers.
//
// Internal helper. Not exposed beyond `src/io/`.

#ifndef FORMULON_IO_OOXML_PACKAGE_RELS_READER_H_
#define FORMULON_IO_OOXML_PACKAGE_RELS_READER_H_

#include <cstdint>
#include <string_view>
#include <vector>

#include "unknown_relationship.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Collects the package-level relationships that are not one of the
/// well-known roots (office document, core / extended / custom properties),
/// in document order. A leading `/` is stripped from internal targets, which
/// must stay inside the package; external targets are kept verbatim.
/// `reader_module` is recorded in diagnostic contexts (e.g. `ooxml_reader`).
Expected<std::vector<UnknownRelationship>, Error> read_unknown_package_rels(const std::vector<std::uint8_t>& bytes,
                                                                            std::string_view reader_module);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_PACKAGE_RELS_READER_H_
