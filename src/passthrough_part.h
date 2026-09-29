//
// `PassthroughPart`: one OOXML part the reader did not model, captured
// raw so the writer can re-emit it verbatim.
//
// Lives in its own header so both the OOXML reader (which produces
// these) and `Workbook` (which carries them through the round-trip)
// can include the type without dragging the rest of the reader/writer
// surface in. The writer (`io::write_ooxml`) consumes the same type
// off the workbook.
//
// Both `<Override>`-listed parts (theme, calcChain, custom XML, ...)
// and Default-typed binary/media parts (vbaProject.bin, xl/media/*
// images, drawings, VML companions, and their rels) round-trip through
// this type. Override-listed parts carry a non-empty `content_type` so
// the writer replicates the `<Override>`; Default-typed parts carry an
// empty `content_type` and rely on the round-tripped `<Default>`
// registration (see `DefaultContentType`).
//
// `retained_origin` / `model_fingerprint` are XLSB-only: the XLSB writer
// re-emits pivot and styles parts verbatim, so each carries a fingerprint
// of its model state at load and `write_xlsb` refuses a stale one (see
// `io/xlsb/retained_part_fingerprint.h`). Every other part leaves them at
// their defaults.

#ifndef FORMULON_PASSTHROUGH_PART_H_
#define FORMULON_PASSTHROUGH_PART_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "utils/index_sort.h"

namespace formulon {

/// One Override-listed part the reader did not consume.
///
///   * `path`         — package-relative path with no leading slash
///                      (e.g. `"xl/theme/theme1.xml"`).
///   * `content_type` — value of the `ContentType=` attribute on the
///                      `[Content_Types].xml` `<Override>` entry. Empty
///                      when the source archive carried the part under
///                      a Default extension (in which case the writer
///                      must NOT emit a per-part Override).
///   * `bytes`        — raw decompressed bytes from the source archive.
///                      The writer copies these straight into the
///                      output package.
struct PassthroughPart {
  PassthroughPart() = default;
  /// A part that carries no duplicated model state.
  PassthroughPart(std::string part_path, std::string part_content_type, std::vector<std::uint8_t> part_bytes)
      : path(std::move(part_path)), content_type(std::move(part_content_type)), bytes(std::move(part_bytes)) {}

  std::string path;
  std::string content_type;
  std::vector<std::uint8_t> bytes;

  /// Which piece of duplicated model state (if any) this part's bytes were
  /// derived from at load time.
  enum class RetainedOrigin : std::uint8_t {
    kNone,
    kPivotTable,
    kPivotCacheDefinition,
    kPivotCacheRecords,
    kStyles,
  };
  RetainedOrigin retained_origin = RetainedOrigin::kNone;
  /// Sheet index of the owning pivot table. Valid only for `kPivotTable`.
  std::size_t origin_sheet_index = 0;
  /// Index into that sheet's `pivot_tables()` at load time. Valid only for
  /// `kPivotTable`.
  std::size_t origin_pivot_index = 0;
  /// `PivotCache::cache_id()` of the owning cache. Valid only for
  /// `kPivotCacheDefinition` / `kPivotCacheRecords`.
  std::uint32_t origin_cache_id = 0;
  /// FNV-1a 64-bit fingerprint of the model state named by `retained_origin`
  /// as it stood right after load. `nullopt` when `retained_origin` is
  /// `kNone`.
  std::optional<std::uint64_t> model_fingerprint;
};

/// Orders captured parts by package path, giving both readers a stable
/// emission order callers and tests can compare against.
///
/// Routed through the shared index sort rather than `std::sort`: the element
/// owns two strings and a byte vector, so a direct sort inlines the whole
/// move-and-destroy machinery into a private copy of the sort body.
inline void sort_passthrough_parts(std::vector<PassthroughPart>& parts) {
  sort_by_index(parts, [](const PassthroughPart& lhs, const PassthroughPart& rhs) { return lhs.path < rhs.path; });
}

}  // namespace formulon

#endif  // FORMULON_PASSTHROUGH_PART_H_
