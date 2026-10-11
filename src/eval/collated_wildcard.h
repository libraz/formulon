//
// Excel's wildcard dialect under the locale's `=` character equivalence:
// each code point is compared by its comparison collation class, so case,
// width and kana type fold one character at a time, while expansions and
// contractions never span characters. `*` matches any run, `?` one
// character, and `~` makes the next character literal. This is the model
// SEARCH and the PivotTable caption predicates (begins / ends / contains)
// share, measured on Mac Excel 365.

#ifndef FORMULON_EVAL_COLLATED_WILDCARD_H_
#define FORMULON_EVAL_COLLATED_WILDCARD_H_

#include <cstddef>
#include <memory>
#include <string_view>

#include "excel_profile.h"

namespace formulon {
namespace eval {

/// A pattern parsed once, with the per-code-point collation classes kept
/// across calls, for matching many texts against one pattern. Not
/// thread-safe: the class cache and scratch buffers mutate under `const`.
class CollatedWildcard {
 public:
  CollatedWildcard(std::string_view pattern, ExcelProfile profile);
  ~CollatedWildcard();
  CollatedWildcard(CollatedWildcard&&) noexcept;
  CollatedWildcard& operator=(CollatedWildcard&&) noexcept;
  CollatedWildcard(const CollatedWildcard&) = delete;
  CollatedWildcard& operator=(const CollatedWildcard&) = delete;

  /// Same contract as `collated_wildcard_match`.
  bool match(std::string_view text) const;

  /// Same contract as `collated_wildcard_find`.
  std::size_t find(std::string_view text, std::size_t start_byte, bool sbcs_question) const;

 private:
  struct Impl;
  std::unique_ptr<Impl> impl_;
};

/// Whether `pattern` matches the whole of `text` under `profile`.
bool collated_wildcard_match(std::string_view pattern, std::string_view text, ExcelProfile profile);

/// Byte offset of the first position at or after `start_byte` where
/// `pattern` matches a prefix of the rest of `text`, or
/// `std::string_view::npos`. Under `sbcs_question`, `?` takes only a
/// single-byte character.
std::size_t collated_wildcard_find(std::string_view pattern, std::string_view text, std::size_t start_byte,
                                   ExcelProfile profile, bool sbcs_question);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COLLATED_WILDCARD_H_
