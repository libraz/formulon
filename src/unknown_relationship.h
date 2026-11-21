//
// `UnknownRelationship`: one entry from a package relationship part
// whose `Type` URI the reader does not recognise, captured so the
// writer can re-emit the relationship and keep the
// `<Override>`-listed part it points at reachable in the package
// graph.
//
// Lives in its own header (mirroring `passthrough_part.h`) so both
// the OOXML reader (which produces these) and `Workbook` (which
// carries them through the round-trip) can include the type without
// dragging the rest of the reader/writer surface in. Default-typed
// binary parts and recognised relationships (worksheets, styles,
// sharedStrings, pivotCacheDefinition, externalLink) are NOT
// represented here; only the unrecognised entries round-trip via
// this struct.

#ifndef FORMULON_UNKNOWN_RELATIONSHIP_H_
#define FORMULON_UNKNOWN_RELATIONSHIP_H_

#include <algorithm>
#include <charconv>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

namespace formulon {

/// One package `<Relationship>` entry the reader did not consume.
///
///   * `id`              — original `Id="rId..."` attribute as it
///                         appeared in the source rels file. Preserved
///                         for diagnostics; the writer mints fresh
///                         rIds when re-emitting to avoid collisions
///                         with the sheets / styles / pivot / external
///                         link numbering.
///   * `type`            — full `Type=` URI verbatim from the source
///                         `<Relationship>` element.
///   * `target`          — resolved package-relative path for internal
///                         entries (relative to `xl/` for workbook rels,
///                         package root for `_rels/.rels`); for entries
///                         with `TargetMode="External"` the raw attribute
///                         value is preserved unchanged.
///   * `target_external` — `true` when the source carried
///                         `TargetMode="External"` on the entry.
struct UnknownRelationship {
  std::string id;
  std::string type;
  std::string target;
  bool target_external = false;
};

/// `rId<n>` for the relationship number `n`.
inline std::string rel_id_text(std::size_t n) {
  return "rId" + std::to_string(n);
}

/// `rId<n>` with `n` one past the largest `rId` number in `rels`. The number is read at a
/// fixed width, so every target mints the same id; one that does not fit reads as 0.
inline std::string next_unknown_rel_id(const std::vector<UnknownRelationship>& rels) {
  std::uint32_t top = 0;
  for (const UnknownRelationship& rel : rels) {
    const std::string_view id(rel.id);
    std::uint32_t n = 0;
    if (id.size() > 3U && id.compare(0, 3, "rId") == 0 &&
        std::from_chars(id.data() + 3, id.data() + id.size(), n).ec == std::errc()) {
      top = std::max(top, n);
    }
  }
  return rel_id_text(top + 1U);
}

}  // namespace formulon

#endif  // FORMULON_UNKNOWN_RELATIONSHIP_H_
