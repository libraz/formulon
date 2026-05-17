// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Pivot-cache rels target resolution. The cache definition part
// (`xl/pivotCache/pivotCacheDefinition<N>.xml`) carries its own rels
// file that points at the matching records part. This helper resolves
// that pairing in isolation so the orchestrator can attach the loaded
// records to the cache it just read.
//
// Internal helper for the OOXML reader. Not exposed beyond `src/io/`.

#ifndef FORMULON_IO_OOXML_PIVOT_TARGET_READER_H_
#define FORMULON_IO_OOXML_PIVOT_TARGET_READER_H_

#include <string>
#include <string_view>

#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Resolves the records-part target referenced from a
/// pivotCacheDefinition's own rels file (e.g.
/// `xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels`). Returns the
/// resolved path (relative to the package root) of the matching
/// `kRelPivotCacheRecords` entry, or an empty string when the rels file
/// is absent or carries no records relationship — both are valid OOXML
/// states (definition-only caches are uncommon but legal).
Expected<std::string, Error> load_pivot_cache_records_target(const ZipReader& zip, std::string_view definition_path);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_PIVOT_TARGET_READER_H_
