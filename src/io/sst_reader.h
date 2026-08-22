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
// kana attached to a cell's source `<si>` entry, span offsets included.

#ifndef FORMULON_IO_SST_READER_H_
#define FORMULON_IO_SST_READER_H_

#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <vector>

#include "phonetic.h"
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
  /// `phonetic_for_entries[i]` holds one `PhoneticRun` per `<rPh>` block
  /// on the `i`-th `<si>`, in document order, and is empty when the entry
  /// carries no `<rPh>`. Unlike `entries` these runs own their kana, so
  /// they do not depend on `text_storage`. The vector is held parallel to
  /// `entries`: `phonetic_for_entries.size() == entries.size()` is an
  /// invariant maintained by `read_shared_strings`.
  std::vector<std::vector<PhoneticRun>> phonetic_for_entries;
  /// `phonetic_props_for_entries[i]` is the `<phoneticPr>` sibling of those
  /// runs, defaulted when the entry carries none. Held parallel to
  /// `entries` under the same invariant as `phonetic_for_entries`.
  std::vector<PhoneticProperties> phonetic_props_for_entries;
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
///     each `<rPh>` block becomes one run in `phonetic_for_entries[i]`,
///     carrying its `sb`/`eb` span and the concatenation of its `<t>`
///     descendants (empty when the `<si>` has no `<rPh>`). A missing or
///     unparsable `sb`/`eb` reads as `0`, which degrades to an insertion
///     at the head of the string rather than dropping the kana.
///   * The decoded payload is appended to `text_storage` (a pointer-
///     stable container), and a `string_view` to that entry is added to
///     `entries` so subsequent appends do not invalidate earlier views.
///   * `xml:space="preserve"` is honoured implicitly: pugixml's default
///     parser keeps internal whitespace in element text, so the raw
///     `text().get()` payload is taken without trimming.
///
/// `sst_bytes` is a sink: the shared-string table is the one part whose
/// size scales with the workbook's total distinct text, so it is parsed
/// in place rather than copied into pugixml. Pass an rvalue and treat
/// the buffer as consumed on return — the decoded payloads have already
/// been copied into `text_storage` by then.
///
/// Errors:
///   * `kIoXmlParse` — pugixml could not parse the document.
///   * `kIoSheetCorrupt` — an `<si>` carried no resolvable `<t>` text
///     (no `<t>` direct child and no `<r><t>` runs). This catches
///     truncated SST entries that would otherwise silently turn into
///     empty strings and lose data on round-trip.
Expected<SharedStringTable, Error> read_shared_strings(std::vector<std::uint8_t> sst_bytes,
                                                       std::deque<std::string>& text_storage);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SST_READER_H_
