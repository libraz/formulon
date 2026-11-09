// Worksheet-name and collection-name identity helpers.
//
// Worksheet names, and the names of every other named collection of a
// workbook, follow one identity rule split by direction:
//   * A name being written (add, rename, copy, insert) collides with every
//     other name it equals under the wide fold, the mac-365-en_US text `=`
//     collation (case, width, kana type and voicing composition fold).
//   * A file reader rejects only names equal under the narrow fold, Unicode
//     simple case folding, so a loaded workbook may hold wide-fold twins.
//   * A typed name resolves to an exact match, else a narrow-fold match, else
//     the one name it equals under the wide fold; two wide candidates resolve
//     to nothing.
// The rule does not depend on the workbook's profile. It is separate from the
// ASCII-only string helpers, which serve function names, logical literals and
// other byte-oriented contracts.

#ifndef FORMULON_SHEET_NAME_H_
#define FORMULON_SHEET_NAME_H_

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace formulon {
namespace sheet_names {

/// Excel measures the sheet-name limit in UTF-16 code units (its internal
/// string representation), not UTF-8 bytes; a supplementary-plane character
/// counts as two units.
inline constexpr std::uint32_t kMaxNameUnits = 31U;

/// Returns true when `lhs` and `rhs` have the same worksheet identity under
/// Unicode simple case folding (Unicode CaseFolding C/S mappings, no
/// normalization or full/expanding mappings).
///
/// The comparison decodes both strings as strict UTF-8 and returns false for
/// malformed input, including when both malformed inputs have identical
/// bytes. Structural worksheet-name validation rejects malformed UTF-8 before
/// names enter a workbook, so the exact-byte fast path is useful for valid
/// ASCII and for callers that already own the validation contract.
bool equal(std::string_view lhs, std::string_view rhs) noexcept;

/// Returns the 0-based index of the first entry of `names` that is `equal` to
/// `sheet` (the narrow fold only; a typed name goes through `resolve`), or
/// `kNotFound`.
inline constexpr std::size_t kNotFound = static_cast<std::size_t>(-1);
std::size_t find_index(const std::vector<std::string>& names, std::string_view sheet) noexcept;

/// Returns true when `lhs` and `rhs` compare equal under the mac-365-en_US
/// text `=` collation at compare strength, whatever the workbook's profile.
/// Malformed UTF-8 on either side is never equal.
bool wide_equal(std::string_view lhs, std::string_view rhs) noexcept;

/// How strictly two names must agree.
enum class Fold : std::uint8_t {
  kExact,   ///< Same bytes.
  kNarrow,  ///< Same bytes, or `equal`.
  kWide,    ///< `wide_equal`.
};

/// Returns true when `lhs` and `rhs` agree under `fold`.
bool same(std::string_view lhs, std::string_view rhs, Fold fold) noexcept;

/// Whether entry `index` of the collection behind `context` answers to the
/// typed name under `fold`.
using MatchFn = bool (*)(const void* context, std::size_t index, Fold fold);

/// Resolves a typed name over `count` entries, where `matches(context, i, fold)`
/// says whether entry `i` answers to the name under `fold`: the first exact
/// match, else the first narrow match, else the single wide match. Returns
/// `kNotFound` when nothing matches or two or more entries match only wide.
std::size_t resolve_by(std::size_t count, MatchFn matches, const void* context);

namespace detail {

template <typename Matches>
bool invoke_matches(const void* context, std::size_t index, Fold fold) {
  return (*static_cast<const Matches*>(context))(index, fold);
}

template <typename Named>
struct NamedQuery {
  const std::vector<Named>* items;
  std::string_view query;
};

template <typename Named>
bool named_matches(const void* context, std::size_t index, Fold fold) {
  const NamedQuery<Named>& q = *static_cast<const NamedQuery<Named>*>(context);
  return same((*q.items)[index].name, q.query, fold);
}

template <typename Table>
bool table_matches(const void* context, std::size_t index, Fold fold) {
  const NamedQuery<Table>& q = *static_cast<const NamedQuery<Table>*>(context);
  const Table& table = (*q.items)[index];
  return same(table.name, q.query, fold) || same(table.display_name, q.query, fold);
}

}  // namespace detail

/// `resolve_by` with `matches(i, fold)` a callable.
template <typename Matches>
std::size_t resolve_by(std::size_t count, const Matches& matches) {
  return resolve_by(count, &detail::invoke_matches<Matches>, &matches);
}

/// `resolve_by` over a list of names.
std::size_t resolve(const std::vector<std::string>& names, std::string_view query);

/// `resolve_by` over items answering to their `name` member (table columns,
/// styles, protected ranges).
template <typename Named>
std::size_t resolve_named(const std::vector<Named>& items, std::string_view query) {
  const detail::NamedQuery<Named> q{&items, query};
  return resolve_by(items.size(), &detail::named_matches<Named>, &q);
}

/// `resolve_by` over tables, each answering to its `name` and its `display_name`.
template <typename Table>
std::size_t resolve_table(const std::vector<Table>& tables, std::string_view query) {
  const detail::NamedQuery<Table> q{&tables, query};
  return resolve_by(tables.size(), &detail::table_matches<Table>, &q);
}

/// A copy of the stored names of one collection (a workbook's sheets, its
/// tables) taken for a rewrite or a formatting pass, so a spelling a formula
/// carries is matched the way it resolves.
class NameList {
 public:
  explicit NameList(std::vector<std::string> names) noexcept : names_(std::move(names)) {}

  /// The index `spelling` resolves to (`resolve_by`), or `kNotFound`.
  std::size_t resolve(std::string_view spelling) const noexcept;

  /// Whether `spelling` resolves to an entry spelled exactly `stored`. With a
  /// null `list` (no name set at hand), whether the two agree under the narrow fold.
  static bool denotes(const NameList* list, std::string_view spelling, std::string_view stored) noexcept;

  /// Whether spellings `a` and `b` are identical or resolve to the same
  /// entry. With a null `list`, whether they agree under the narrow fold.
  static bool same_entry(const NameList* list, std::string_view a, std::string_view b) noexcept;

 private:
  std::vector<std::string> names_;
};

/// The tables' display names, and each name that differs from its display
/// name, as a `NameList` a table spelling resolves among.
template <typename Table>
NameList table_name_list(const std::vector<Table>& tables) {
  std::vector<std::string> names;
  names.reserve(tables.size());
  for (const Table& table : tables) {
    names.emplace_back(table.display_name);
    if (table.name != table.display_name) {
      names.emplace_back(table.name);
    }
  }
  return NameList(std::move(names));
}

/// Hash key with `wide_equal(a, b) == (*wide_key(a) == *wide_key(b))` for
/// well-formed names; nullopt for malformed UTF-8.
std::optional<std::string> wide_key(std::string_view name);

/// Returns true when writing `candidate` collides with `existing`: the wide
/// fold, plus the narrow fold a file reader rejects.
bool write_collides(std::string_view existing, std::string_view candidate) noexcept;

/// Returns true when writing `candidate` at `self_index` collides with any
/// other entry of `names`; pass `kNotFound` (or `names.size()`) for a new name.
bool write_collides(const std::vector<std::string>& names, std::string_view candidate, std::size_t self_index);

/// Returns true when `name` is well-formed UTF-8. Unicode scalar values are
/// decoded strictly: overlong encodings, surrogate code points, truncated
/// sequences, stray continuation bytes, and values above U+10FFFF are
/// rejected.
bool valid_utf8(std::string_view name) noexcept;

}  // namespace sheet_names

namespace collection_names {

/// Returns true when `lhs` and `rhs` agree under the narrow fold a file
/// reader rejects duplicates by: the worksheet-name fold, plus byte identity
/// so a name a file stored as malformed UTF-8 still matches itself.
inline bool equal(std::string_view lhs, std::string_view rhs) noexcept {
  return lhs == rhs || sheet_names::equal(lhs, rhs);
}

/// Hash key with `equal(a, b) == (fold_key(a) == fold_key(b))`: a well-formed
/// name maps to its folded UTF-8 spelling, a malformed one to its own bytes
/// (never well-formed, so the two kinds cannot collide).
std::string fold_key(std::string_view name);

}  // namespace collection_names
}  // namespace formulon

#endif  // FORMULON_SHEET_NAME_H_
