// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// External-link metadata reader. Surfaces the cross-workbook references
// recorded in `<externalReferences>` (in `xl/workbook.xml`) joined with
// the relationship targets in `xl/_rels/workbook.xml.rels` and the
// per-link rels file (`xl/externalLinks/_rels/externalLink<N>.xml.rels`).
//
// The link's body part — `xl/externalLinks/externalLink<N>.xml`, which
// records cached cell values, sheet names and defined names from the
// remote workbook — is NOT parsed here; it round-trips verbatim through
// `Workbook::passthrough_parts()`. This module only exposes enough
// metadata for callers to enumerate the links and obtain the target
// URL / kind through the C ABI.
//
// Design references:
//   * ECMA-376 §18.14 (externalLink, externalBook, oleLink, ddeLink)
//   * backup/plans/04-xlsx-io.md (workbook.xml.rels handling)

#ifndef FORMULON_IO_EXTERNAL_LINKS_H_
#define FORMULON_IO_EXTERNAL_LINKS_H_

#include <cstdint>
#include <string>
#include <vector>

namespace formulon {
namespace io {

/// One entry in the workbook's `<externalReferences>` list, joined with
/// the resolved relationship metadata.
///
///   * `index`         — 1-based document order (matches the i-th
///                       `<externalReference>` in `xl/workbook.xml`).
///   * `rel_id`        — the `r:id` attribute on `<externalReference>`,
///                       matching a `<Relationship Id="..."/>` entry in
///                       `xl/_rels/workbook.xml.rels` whose Type is the
///                       externalLink relationship.
///   * `part_path`     — the package-relative path of the external link
///                       body part (e.g. `xl/externalLinks/externalLink1.xml`).
///                       Resolved from the workbook.xml.rels Target.
///   * `body_rel_id`   — the `r:id` attribute inside the body part
///                       (e.g. `<externalBook r:id="rId1"/>`), matching
///                       the `<Relationship Id="..."/>` entry in
///                       `xl/externalLinks/_rels/externalLink<N>.xml.rels`.
///                       Captured verbatim so the writer can re-emit the
///                       per-link rels file with the same id (the body
///                       part round-trips through `passthrough_parts()`
///                       so its inner reference must continue to match).
///   * `target`        — the actual remote workbook URL captured in the
///                       per-link rels file (e.g. `file:///path/book.xlsx`,
///                       `http://...`). Empty when the per-link rels file
///                       is absent or malformed.
///   * `target_external` — `true` when the per-link relationship was
///                         emitted with `TargetMode="External"` (the
///                         common case for cross-workbook links). `false`
///                         indicates an in-package target.
///   * `kind`          — the root element of the body part:
///                         `kExternalBook` for `<externalBook>` (most common),
///                         `kOleLink` / `kDdeLink` for legacy variants,
///                         `kUnknown` when the part is missing or unparseable.
struct ExternalLinkRecord {
  enum class Kind : std::uint8_t {
    kUnknown = 0,
    kExternalBook = 1,
    kOleLink = 2,
    kDdeLink = 3,
  };

  std::uint32_t index = 0;
  std::string rel_id;
  std::string part_path;
  std::string body_rel_id;
  std::string target;
  bool target_external = true;
  Kind kind = Kind::kUnknown;
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_EXTERNAL_LINKS_H_
