// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// MS-XLSB shared-string table builder + emitter.
//
// `SstBuilder` interns text payloads in insertion order so the cell
// writer can replace `Value::Text` cells with a `BrtCellIsst` record
// pointing at the SST index. `emit_sst` packages the interned strings
// as a sequence of `BrtBeginSst | BrtSSTItem* | BrtEndSst` records.
//
// The writer emits plain (non-rich) entries only — every `BrtSSTItem`
// is `flags=0x00` followed by an `XLWideString`. This is the minimum
// shape `read_sst` recognises, and matches the reader's expectations.
//
// Design references:
//   * [MS-XLSB] §2.4.293 (BrtSSTItem) and §2.4.290 (BrtBeginSst)

#ifndef FORMULON_IO_XLSB_SST_WRITER_H_
#define FORMULON_IO_XLSB_SST_WRITER_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Interns text payloads for `xl/sharedStrings.bin`.
///
/// Identical strings dedupe to the same index. The hash table only
/// holds keys (the strings themselves live in `entries_`), so callers
/// can observe insertion order via `entries()`.
class SstBuilder {
 public:
  /// Interns `text` and returns the assigned 0-based index. The first
  /// time a string is seen the index equals the prior `size()`.
  std::uint32_t intern(std::string_view text);

  /// Number of distinct strings interned so far. Equals
  /// `entries().size()`.
  std::uint32_t size() const noexcept { return static_cast<std::uint32_t>(entries_.size()); }

  /// Returns `true` when no string has been interned yet. Callers
  /// (notably `write_xlsb`) gate emission of the SST part on this
  /// predicate so a workbook with no text cells produces no SST
  /// part / Override.
  bool empty() const noexcept { return entries_.empty(); }

  /// Read-only access to the interned strings in insertion order.
  /// Stable across subsequent `intern` calls (entries are append-only).
  const std::vector<std::string>& entries() const noexcept { return entries_; }

 private:
  std::vector<std::string> entries_;
  std::unordered_map<std::string, std::uint32_t> index_;
};

/// Serialises `sst` as the body of `xl/sharedStrings.bin`.
///
/// Layout:
///   * `BrtBeginSst` payload — `(u32 cstTotal, u32 cstUnique)`. We
///     emit `cstTotal == cstUnique == sst.size()` because the writer
///     does not track multiplicity (every cell that referenced the
///     same string was redirected to the same SST index by the
///     interner).
///   * One `BrtSSTItem` per entry: `(u8 flags=0x00, XLWideString)`.
///   * `BrtEndSst` (empty payload).
///
/// The returned bytes are a complete XLSB part body ready to be
/// stored as `xl/sharedStrings.bin`. Returns no errors today; the
/// `Expected` shape is preserved for forward compatibility (rich-
/// text + phonetic-guide emission may surface allocation failures
/// in the future).
Expected<std::vector<std::uint8_t>, Error> emit_sst(const SstBuilder& sst);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SST_WRITER_H_
