// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
std::string BuildPackageRels(const EmissionPlan& plan);

/// Builds the `xl/workbook.xml` part: `<sheets>`, optional
/// `<externalReferences>`, `<definedNames>`, `<calcPr>`, and
/// `<pivotCaches>` in ECMA-376 element order.
std::string BuildWorkbookXml(const Workbook& wb, const EmissionPlan& plan);

/// Builds the `xl/_rels/workbook.xml.rels` part: per-sheet worksheet
/// relationships, the styles relationship, pivot-cache / external-link
/// relationships, and any round-tripped `UnknownRelationship` entries.
std::string BuildWorkbookRels(std::size_t sheet_count, const EmissionPlan& plan, const Workbook& wb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WORKBOOK_XML_BUILDER_H_
