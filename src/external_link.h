//
// External-link metadata. Surfaces the cross-workbook references recorded
// in `<externalReferences>` (in `xl/workbook.xml`) joined with the
// relationship targets in `xl/_rels/workbook.xml.rels` and the per-link
// rels file (`xl/externalLinks/_rels/externalLink<N>.xml.rels`).
//
// The link's body part -- `xl/externalLinks/externalLink<N>.xml`, which
// records cached cell values, sheet names and defined names from the
// remote workbook -- is read into `ExternalLinkRecord::book` so that
// cross-workbook references can be evaluated against Excel's own cache
// (see `external_book.h`). A loaded body round-trips verbatim through
// `Workbook::passthrough_parts()` while the model still matches it.
//
// The engine also creates records: entering a formula that names a book
// with no link appends one (`Workbook::bind_external_books`), so every
// external book a formula mentions has a record to save it under.
//
// Design references:
//   * ECMA-376 §18.14 (externalLink, externalBook, oleLink, ddeLink)

#ifndef FORMULON_EXTERNAL_LINK_H_
#define FORMULON_EXTERNAL_LINK_H_

#include <cstdint>
#include <string>
#include <vector>

#include "external_book.h"

namespace formulon {

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
///   * `absolute_target` — the absolute path Excel records next to a relative
///                       `target`, through `<xxl21:alternateUrls>` /
///                       `<xxl21:absoluteUrl r:id>`. Empty when absent.
///   * `absolute_rel_id` — that relationship's id in the per-link rels file.
///   * `relative_target` — the relative path Excel records beside the target,
///                       through `<xxl21:alternateUrls>` /
///                       `<xxl21:relativeUrl r:id>`. Empty when absent.
///   * `relative_rel_id` — that relationship's id in the per-link rels file.
///   * `target_external` — `true` when the per-link relationship was
///                         emitted with `TargetMode="External"` (the
///                         common case for cross-workbook links). `false`
///                         indicates an in-package target.
///   * `target_path_missing` — `true` when the body target relationship is
///                         Excel's `xlPathMissing` kind: the source file
///                         could not be located, and `target` holds the bare
///                         name it was referenced by. Book kinds only.
///   * `kind`          — the root element of the body part:
///                         `kExternalBook` for `<externalBook>` (most common),
///                         `kOleLink` / `kDdeLink` for legacy variants,
///                         `kUnknown` when the part is missing or unparseable.
///   * `body_stale`    — `true` once a sheet name was added to `book` that
///                       the loaded body part does not list, so the part can
///                       no longer be written back verbatim.
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
  std::string absolute_target;
  std::string absolute_rel_id;
  std::string relative_target;
  std::string relative_rel_id;
  /// The `/`-separated path of a link the engine created, relative to the book's rule folder (below it) or to its
  /// parent (`../...`); the .xlsx writer spells the link with it. Empty otherwise.
  std::string created_relative;
  /// `created_relative` is the link's only target (the book is below the rule folder), not an alternate.
  bool created_relative_is_target = false;
  bool target_external = true;
  bool target_path_missing = false;
  Kind kind = Kind::kUnknown;
  bool body_stale = false;
  /// The supporting workbook's cached sheet names, defined names and
  /// cell values. Populated only for `kExternalBook`; an OLE or DDE link
  /// carries no such cache and leaves this empty, which reads as "no
  /// sheet, no name, nothing cached".
  ExternalBook book;
};

class Workbook;

/// The absolute folder and file of the book `rec` links, as a rule or
/// validation spells it by path: the link's primary target, resolved against
/// the folder the book is saved in when it is relative and used as is when
/// absolute. That folder is `wb.location()`'s directory,
/// `wb.loaded_abs_path()`, or the one both name; empty when the two
/// differ (the book was moved) or neither exists. Also empty when `rec` has no
/// target or a relative target with a `.` or `..` segment.
std::string linked_book_path(const Workbook& wb, const ExternalLinkRecord& rec);

}  // namespace formulon

#endif  // FORMULON_EXTERNAL_LINK_H_
