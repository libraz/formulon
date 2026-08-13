//
// External-link reader. Walks `<externalReferences>` in the parsed
// workbook document, joins each `<externalReference>` against the
// workbook rels (resolved external-link part paths) and the per-link
// rels file (which carries the actual remote URL), and produces one
// `ExternalLinkRecord` per entry in document order.
//
// The link's body part (`xl/externalLinks/externalLink<N>.xml`)
// continues to round-trip verbatim through `passthrough_parts()`; this
// reader only captures the metadata. Missing or unparseable parts
// produce an `kUnknown` record rather than failing the load — Excel
// itself tolerates partially-broken external link sections and we
// match that behaviour.
//
// Internal helper for the OOXML reader. Not exposed beyond `src/io/`.

#ifndef FORMULON_IO_OOXML_EXTERNAL_LINK_READER_H_
#define FORMULON_IO_OOXML_EXTERNAL_LINK_READER_H_

#include <string>
#include <vector>

#include "io/external_links.h"
#include "io/ooxml/workbook_rels_reader.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Output of `load_external_links`: parsed link records (in document
/// order) plus the paths of any per-link rels files the reader observed.
/// The caller marks those rels paths as consumed; the writer regenerates
/// them from the captured records, while the body parts themselves stay
/// in `passthrough_parts()`.
struct ExternalLinkLoadResult {
  std::vector<ExternalLinkRecord> records;
  std::vector<std::string> consumed_rels_paths;
};

/// Walks `<externalReferences>` in `wb_root` and joins each entry
/// against `wb_rels` to build per-link records. For each resolved body
/// part the helper classifies the link kind by peeking at the body's
/// root element and captures the target URL + TargetMode from the
/// per-link rels file. Missing parts and successfully-read malformed XML
/// remain failure-tolerant and produce `kUnknown`/partial records. A
/// `ZipReader::read_entry` failure is returned unchanged so callers do not
/// silently continue with a corrupted package.
Expected<ExternalLinkLoadResult, Error> load_external_links(const ZipReader& zip, const pugi::xml_node& wb_root,
                                                            const WorkbookRels& wb_rels);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_EXTERNAL_LINK_READER_H_
