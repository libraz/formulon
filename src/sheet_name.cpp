// Worksheet-name UTF-8 decoding and Unicode simple-fold identity.

#include "sheet_name.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "sheet_name_casefold_data.h"
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
  if (lhs == rhs) {
    return valid_utf8(lhs);
  }
  std::size_t lhs_offset = 0;
  std::size_t rhs_offset = 0;
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

}  // namespace sheet_names
}  // namespace formulon
