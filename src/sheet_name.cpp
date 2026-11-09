// Worksheet-name UTF-8 decoding, the narrow and wide name folds, and the
// by-name resolver.
//
// @size-budget: 4 KB

#include "sheet_name.h"

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#include "eval/collation.h"
#include "excel_profile.h"
#include "sheet_name_casefold_data.h"
#include "utils/strings.h"
#include "utils/text_ops.h"

namespace formulon {
namespace sheet_names {
namespace {

// Decodes one scalar at `*offset`, applying the strict UTF-8 constraints from
// RFC 3629. On failure, the offset is left at the malformed lead byte so the
// caller can terminate without reading past the input.
bool decode_one(std::string_view text, std::size_t* offset, std::uint32_t* scalar) noexcept {
  const std::size_t start = *offset;
  if (start >= text.size()) {
    return false;
  }
  const Utf8DecodeResult decoded = decode_first_utf8_codepoint(text.substr(start));
  if (!decoded.valid) {
    return false;
  }
  *scalar = decoded.codepoint;
  *offset = start + decoded.byte_len;
  return true;
}

// The wide fold is the en-US collation for every profile: no measured locale folds wider.
constexpr ExcelProfile kNameIdentityProfile = mac_365_en_us_profile();

}  // namespace

bool valid_utf8(std::string_view name) noexcept {
  std::size_t offset = 0;
  std::uint32_t scalar = 0;
  while (offset < name.size()) {
    if (!decode_one(name, &offset, &scalar)) {
      return false;
    }
  }
  return true;
}

bool equal(std::string_view lhs, std::string_view rhs) noexcept {
  // ASCII prefix: a one-byte scalar folds by A-Z -> a-z only; the first non-ASCII byte on either side hands over to
  // the scalar loop at the same offset.
  const std::size_t common = lhs.size() < rhs.size() ? lhs.size() : rhs.size();
  std::size_t lhs_offset = 0;
  while (lhs_offset < common) {
    unsigned char a = static_cast<unsigned char>(lhs[lhs_offset]);
    unsigned char b = static_cast<unsigned char>(rhs[lhs_offset]);
    if (((a | b) & 0x80U) != 0U) {
      break;
    }
    a = static_cast<unsigned char>(strings::ascii_to_lower(static_cast<char>(a)));
    b = static_cast<unsigned char>(strings::ascii_to_lower(static_cast<char>(b)));
    if (a != b) {
      return false;
    }
    ++lhs_offset;
  }
  std::size_t rhs_offset = lhs_offset;
  while (lhs_offset < lhs.size() && rhs_offset < rhs.size()) {
    std::uint32_t lhs_scalar = 0;
    std::uint32_t rhs_scalar = 0;
    if (!decode_one(lhs, &lhs_offset, &lhs_scalar) || !decode_one(rhs, &rhs_offset, &rhs_scalar)) {
      return false;
    }
    if (detail::fold_scalar(lhs_scalar) != detail::fold_scalar(rhs_scalar)) {
      return false;
    }
  }
  return lhs_offset == lhs.size() && rhs_offset == rhs.size();
}

std::size_t find_index(const std::vector<std::string>& names, std::string_view sheet) noexcept {
  for (std::size_t i = 0; i < names.size(); ++i) {
    if (equal(names[i], sheet)) {
      return i;
    }
  }
  return kNotFound;
}

bool wide_equal(std::string_view lhs, std::string_view rhs) noexcept {
  if (!valid_utf8(lhs) || !valid_utf8(rhs)) {
    return false;
  }
  return eval::collate(lhs, rhs, kNameIdentityProfile, eval::CollationStrength::kCompare) == 0;
}

bool same(std::string_view lhs, std::string_view rhs, Fold fold) noexcept {
  switch (fold) {
    case Fold::kExact:
      return lhs == rhs;
    case Fold::kNarrow:
      return lhs == rhs || equal(lhs, rhs);
    case Fold::kWide:
      return wide_equal(lhs, rhs);
  }
  return false;
}

std::size_t resolve_by(std::size_t count, MatchFn matches, const void* context) {
  for (const Fold fold : {Fold::kExact, Fold::kNarrow}) {
    for (std::size_t i = 0; i < count; ++i) {
      if (matches(context, i, fold)) {
        return i;
      }
    }
  }
  std::size_t found = kNotFound;
  for (std::size_t i = 0; i < count; ++i) {
    if (matches(context, i, Fold::kWide)) {
      if (found != kNotFound) {
        return kNotFound;
      }
      found = i;
    }
  }
  return found;
}

std::size_t resolve(const std::vector<std::string>& names, std::string_view query) {
  const detail::NamedQuery<std::string> q{&names, query};
  return resolve_by(
      names.size(),
      [](const void* context, std::size_t index, Fold fold) {
        const auto& list = *static_cast<const detail::NamedQuery<std::string>*>(context);
        return same((*list.items)[index], list.query, fold);
      },
      &q);
}

std::size_t NameList::resolve(std::string_view spelling) const noexcept {
  return sheet_names::resolve(names_, spelling);
}

bool NameList::denotes(const NameList* list, std::string_view spelling, std::string_view stored) noexcept {
  if (list == nullptr) {
    return same(spelling, stored, Fold::kNarrow);
  }
  const std::size_t index = list->resolve(spelling);
  return index != kNotFound && list->names_[index] == stored;
}

bool NameList::same_entry(const NameList* list, std::string_view a, std::string_view b) noexcept {
  if (a == b) {
    return true;
  }
  if (list == nullptr) {
    return same(a, b, Fold::kNarrow);
  }
  const std::size_t index = list->resolve(a);
  return index != kNotFound && index == list->resolve(b);
}

std::optional<std::string> wide_key(std::string_view name) {
  if (!valid_utf8(name)) {
    return std::nullopt;
  }
  return eval::collation_key(name, kNameIdentityProfile, eval::CollationStrength::kCompare);
}

bool write_collides(std::string_view existing, std::string_view candidate) noexcept {
  return same(existing, candidate, Fold::kNarrow) || wide_equal(existing, candidate);
}

bool write_collides(const std::vector<std::string>& names, std::string_view candidate, std::size_t self_index) {
  for (std::size_t i = 0; i < names.size(); ++i) {
    if (i != self_index && write_collides(names[i], candidate)) {
      return true;
    }
  }
  return false;
}

}  // namespace sheet_names

namespace collection_names {

std::string fold_key(std::string_view name) {
  std::string key;
  key.reserve(name.size());
  std::size_t offset = 0;
  while (offset < name.size()) {
    const char c = name[offset];
    // Same per-scalar fold as `sheet_names::equal`, ASCII fast path included.
    if ((static_cast<unsigned char>(c) & 0x80U) == 0U) {
      key.push_back(strings::ascii_to_lower(c));
      ++offset;
      continue;
    }
    std::uint32_t scalar = 0;
    if (!sheet_names::decode_one(name, &offset, &scalar)) {
      return std::string(name);
    }
    key += encode_utf8_codepoint(sheet_names::detail::fold_scalar(scalar));
  }
  return key;
}

}  // namespace collection_names
}  // namespace formulon
