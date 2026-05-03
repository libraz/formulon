// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared-strings (`xl/sharedStrings.xml`) reader. The OOXML SST is a flat
// list of `<si>` (string item) entries. Each entry is either a single
// `<t>` text run or a sequence of `<r>` rich-text runs whose `<t>`
// payloads concatenate into a single plain-text string. Cells reference
// SST entries by 0-based position via `<c t="s"><v>INDEX</v></c>`.
//
// This reader produces a flat `std::vector` of `string_view`s into
// caller-supplied `text_storage` so the resulting strings share lifetime
// with the rest of the OOXML reader's owned text payloads (the same
// `OoxmlReadResult::text_storage` deque used for inline strings). Rich-
// text formatting (`<r><rPr>`) is intentionally dropped: this layer is
// plain-text only. Phonetic guides (`<rPh>`) ARE preserved through a
// parallel `phonetic_for_entries` vector so PHONETIC() can surface the
// IME-typed kana attached to a cell's source `<si>` entry.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.4 (Reader pipeline)
//   * backup/plans/26-implementation-plan.md (Phase 2.3)

#ifndef FORMULON_IO_SST_READER_H_
#define FORMULON_IO_SST_READER_H_

#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// In-memory representation of `xl/sharedStrings.xml`.
///
/// `entries[i]` is the resolved plain-text payload of the `i`-th `<si>`
/// in the source document (preserving document order — Excel uses these
/// indices verbatim). Each `string_view` is non-owning and aliases an
/// entry inside the `text_storage` argument supplied to
/// `read_shared_strings`. The caller therefore controls text lifetime:
/// keep `text_storage` alive at least as long as `entries` views are
/// dereferenced.
struct SharedStringTable {
  std::vector<std::string_view> entries;
  /// `phonetic_for_entries[i]` is the concatenated kana from every
  /// `<rPh>` annotation on the `i`-th `<si>` (in document order across
  /// blocks). Empty `string_view{}` when the entry carries no `<rPh>`.
  /// Each view aliases an entry inside `text_storage`, with the same
  /// lifetime contract as `entries`. The vector is held parallel to
  /// `entries`: `phonetic_for_entries.size() == entries.size()` is an
  /// invariant maintained by `read_shared_strings`.
  std::vector<std::string_view> phonetic_for_entries;
};

/// Parses an OOXML shared-strings part.
///
/// Behaviour:
///   * Empty `<sst/>` (no `<si>` children) is valid and yields zero
///     entries.
///   * Each `<si>` is reduced to a single plain-text string by
///     concatenating all descendant `<t>` payloads in document order.
///     Rich-text runs (`<r><t>...</t></r>`) are walked transparently and
///     their `<t>` payloads are appended to `entries[i]`.
///   * Phonetic guides (`<rPh>`) are NOT folded into `entries`. Instead,
///     each `<rPh>` block's `<t>` descendants are concatenated in
///     document order and the resulting kana string is stored at
///     `phonetic_for_entries[i]` (or empty when the `<si>` has no
///     `<rPh>`).
///   * The decoded payload is appended to `text_storage` (a pointer-
///     stable container), and a `string_view` to that entry is added to
///     `entries` so subsequent appends do not invalidate earlier views.
///   * `xml:space="preserve"` is honoured implicitly: pugixml's default
///     parser keeps internal whitespace in element text, so the raw
///     `text().get()` payload is taken without trimming.
///
/// Errors:
///   * `kIoXmlParse` — pugixml could not parse the document.
///   * `kIoSheetCorrupt` — an `<si>` carried no resolvable `<t>` text
///     (no `<t>` direct child and no `<r><t>` runs). This catches
///     truncated SST entries that would otherwise silently turn into
///     empty strings and lose data on round-trip.
Expected<SharedStringTable, Error> read_shared_strings(const std::vector<std::uint8_t>& sst_bytes,
                                                       std::deque<std::string>& text_storage);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SST_READER_H_
