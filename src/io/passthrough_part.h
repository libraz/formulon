// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `PassthroughPart`: one Override-listed OOXML part the reader did not
// consume, captured raw so the writer can re-emit it verbatim.
//
// Lives in its own header so both the OOXML reader (which produces
// these) and `Workbook` (which carries them through the round-trip)
// can include the type without dragging the rest of the reader/writer
// surface in. The writer (`io::write_ooxml`) consumes the same type
// off the workbook.
//
// Default-typed binary parts (images, OLE objects, …) are NOT
// represented here; only Override-listed parts round-trip. See the
// reader docs for the v1 caveat.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.2 (package structure)
//   * backup/plans/26-implementation-plan.md (Phase 2.5)

#ifndef FORMULON_IO_PASSTHROUGH_PART_H_
#define FORMULON_IO_PASSTHROUGH_PART_H_

#include <cstdint>
#include <string>
#include <vector>

namespace formulon {
namespace io {

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
  std::string path;
  std::string content_type;
  std::vector<std::uint8_t> bytes;
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_PASSTHROUGH_PART_H_
