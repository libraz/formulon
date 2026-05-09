// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Reader for `xl/comments<N>.xml`. Each comments part carries a
// per-workbook author table plus a flat list of `<comment>` elements
// each anchored at a single cell. The reader resolves authors against
// the table inline and produces a list of `CellComment` records ready
// for assignment to the owning sheet.
//
// Rich-text fidelity: a `<comment>` body may contain `<r>` runs with
// per-run formatting. This reader concatenates every run's `<t>` text
// into a single plain-text payload (no per-run formatting is preserved).
// The writer re-emits the same plain text on save; rich-text editing in
// the engine is out of scope for this bundle.

#ifndef FORMULON_IO_COMMENTS_READER_H_
#define FORMULON_IO_COMMENTS_READER_H_

#include <cstdint>
#include <vector>

#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {

/// Decodes the bytes of an `xl/comments<N>.xml` part into a flat list
/// of `CellComment` records, in document order. Rich-text runs are
/// concatenated into the comment's plain-text payload.
///
/// Errors:
///   * `kIoXmlParse` on malformed XML.
///   * `kIoSheetCorrupt` on a missing root, missing `ref=`, or
///     unparseable A1 cell reference.
Expected<std::vector<CellComment>, Error> read_comments(const std::vector<std::uint8_t>& bytes);

}  // namespace formulon::io

#endif  // FORMULON_IO_COMMENTS_READER_H_
