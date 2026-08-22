//
// MS-XLSB shared-string table builder + emitter.
//
// `SstBuilder` interns text payloads in insertion order so the cell
// writer can replace `Value::Text` cells with a `BrtCellIsst` record
// pointing at the SST index. `emit_sst` packages the interned strings
// as a sequence of `BrtBeginSst | BrtSSTItem* | BrtEndSst` records.
//
// Interning is keyed on the text AND its phonetic guide, matching the
// OOXML shared-strings writer: two cells reading the same kanji with
// different furigana are different `<si>` entries, and merging them
// would move one cell's reading onto the other.
//
// The writer emits plain (non-rich) entries only: `flags` carries the
// phonetic bit when the entry has a guide and is otherwise `0x00`, and
// rich-format runs are never produced. Both shapes are what `read_sst`
// recognises.
//
// Design references:
//   * [MS-XLSB] §2.4.293 (BrtSSTItem), §2.4.290 (BrtBeginSst),
//     §2.5.87 (RichStr)

#ifndef FORMULON_IO_XLSB_SST_WRITER_H_
#define FORMULON_IO_XLSB_SST_WRITER_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "phonetic.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// One interned shared string: the text plus the phonetic guide that
/// travels with it.
struct SstEntry {
  std::string text;
  std::vector<PhoneticRun> phonetic;
};

/// Interns text payloads for `xl/sharedStrings.bin`.
///
/// Entries with identical text and identical phonetic guides dedupe to
/// the same index. The hash table only holds keys (the payloads
/// themselves live in `entries_`), so callers can observe insertion
/// order via `entries()`.
class SstBuilder {
 public:
  /// Interns `text` carrying `phonetic` and returns the assigned 0-based
  /// index. The first time a payload is seen the index equals the prior
  /// `size()`.
  std::uint32_t intern(std::string_view text, const std::vector<PhoneticRun>& phonetic);

  /// Number of distinct payloads interned so far. Equals
  /// `entries().size()`.
  std::uint32_t size() const noexcept { return static_cast<std::uint32_t>(entries_.size()); }

  /// Returns `true` when nothing has been interned yet. Callers
  /// (notably `write_xlsb`) gate emission of the SST part on this
  /// predicate so a workbook with no text cells produces no SST
  /// part / Override.
  bool empty() const noexcept { return entries_.empty(); }

  /// Read-only access to the interned payloads in insertion order.
  /// Stable across subsequent `intern` calls (entries are append-only).
  const std::vector<SstEntry>& entries() const noexcept { return entries_; }

 private:
  std::vector<SstEntry> entries_;
  std::unordered_map<std::string, std::uint32_t> index_;
};

/// Serialises `sst` as the body of `xl/sharedStrings.bin`.
///
/// Layout:
///   * `BrtBeginSst` payload — `(u32 cstTotal, u32 cstUnique)`. We
///     emit `cstTotal == cstUnique == sst.size()` because the writer
///     does not track multiplicity (every cell that referenced the
///     same payload was redirected to the same SST index by the
///     interner).
///   * One `BrtSSTItem` per entry: `(u8 flags, XLWideString)`, followed
///     by the phonetic tail when the entry carries a guide.
///   * `BrtEndSst` (empty payload).
///
/// The phonetic tail stores the kana once, concatenated across every
/// run, then `(u32 count)` and one `(u16 ichFirst, u16 ichMom, u16
/// cchMom)` per run: where this run's kana starts inside the
/// concatenation, which surface-text offset it reads, and how many
/// surface characters it covers. A run's kana ends where the next one's
/// starts. The closing `(u16 ifnt, u16 flags)` is the phonetic font and
/// the annotation type / alignment — emitted as font 0 with Excel's own
/// defaults, since `PhoneticRun` models the reading rather than how it
/// is rendered, and the OOXML writer likewise emits no `<phoneticPr>`.
///
/// The returned bytes are a complete XLSB part body ready to be
/// stored as `xl/sharedStrings.bin`. Returns no errors today; the
/// `Expected` shape is preserved for forward compatibility (rich-text
/// emission may surface allocation failures in the future).
Expected<std::vector<std::uint8_t>, Error> emit_sst(const SstBuilder& sst);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SST_WRITER_H_
