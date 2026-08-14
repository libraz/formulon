//
// Workbook-level OOXML part builders: `[Content_Types].xml`,
// `_rels/.rels`, `xl/workbook.xml`, and `xl/_rels/workbook.xml.rels`.
// Internal to `src/io/ooxml/`; not part of the public API. The only
// caller is `src/io/ooxml_writer.cpp`'s orchestrator, which threads the
// results into the in-memory zip archive.

#ifndef FORMULON_IO_OOXML_WORKBOOK_XML_BUILDER_H_
#define FORMULON_IO_OOXML_WORKBOOK_XML_BUILDER_H_

#include <cstddef>
#include <string>

#include "io/ooxml/emission_plan.h"

namespace formulon {
class Workbook;
namespace io {

/// Builds the `[Content_Types].xml` part for the given workbook plus
/// emission plan. Emits the package-level `<Default>` entries, the
/// per-sheet / per-table / per-pivot / per-comments `<Override>`
/// entries, and any passthrough overrides for parts that carried an
/// explicit content type in the source archive.
std::string BuildContentTypes(const Workbook& wb, const EmissionPlan& plan);

/// Builds the package-level `_rels/.rels` part: the relationship to the
/// workbook part itself plus optional core / extended property
/// relationships when the corresponding passthrough parts survived.
///
/// A round-tripped relationship whose target part is no longer in the
/// package is omitted and bumps `diagnostics->dropped_relationship_count`.
/// `diagnostics` may be NULL, which discards the count.
std::string BuildPackageRels(const Workbook& wb, const EmissionPlan& plan, WriteDiagnostics* diagnostics);

/// Builds the `xl/workbook.xml` part: `<sheets>`, optional
/// `<externalReferences>`, `<definedNames>`, `<calcPr>`, and
/// `<pivotCaches>` in ECMA-376 element order.
std::string BuildWorkbookXml(const Workbook& wb, const EmissionPlan& plan);

/// Builds the `xl/_rels/workbook.xml.rels` part: per-sheet worksheet
/// relationships, the styles relationship, pivot-cache / external-link
/// relationships, and any round-tripped `UnknownRelationship` entries.
///
/// Drops entries whose target part is absent under the same rule, the
/// same counter and the same NULL policy as `BuildPackageRels`.
std::string BuildWorkbookRels(std::size_t sheet_count, const EmissionPlan& plan, const Workbook& wb,
                              WriteDiagnostics* diagnostics);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WORKBOOK_XML_BUILDER_H_
