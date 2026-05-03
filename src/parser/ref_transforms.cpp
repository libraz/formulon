// Copyright 2026 libraz. Licensed under the MIT License.

#include "parser/ref_transforms.h"

#include <cstdint>
#include <optional>
#include <string_view>

#include "parser/reference.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {

bool sheet_name_needs_quoting(std::string_view name) noexcept {
  if (name.empty()) {
    return true;
  }
  for (char c : name) {
    const bool ok = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '.';
    if (!ok) {
      return true;
    }
  }
  return false;
}

SheetRenameTransform::SheetRenameTransform(std::string_view old_name, std::string_view new_name) noexcept
    : old_name_(old_name), new_name_(new_name), new_name_needs_quotes_(sheet_name_needs_quoting(new_name)) {}

std::optional<std::string_view> SheetRenameTransform::remap_sheet(std::string_view sheet) const noexcept {
  if (sheet.empty()) {
    return std::nullopt;
  }
  if (!strings::case_insensitive_eq(sheet, old_name_)) {
    return std::nullopt;
  }
  return new_name_;
}

std::optional<Reference> SheetRenameTransform::apply(const Reference& ref) const {
  std::optional<std::string_view> remapped = remap_sheet(ref.sheet);
  if (!remapped.has_value()) {
    // Local reference (no sheet) or non-matching sheet: pass through
    // unchanged. The walker treats an unchanged Reference as identity.
    return ref;
  }
  Reference out = ref;
  out.sheet = *remapped;
  out.sheet_quoted = new_name_needs_quotes_;
  return out;
}

std::optional<std::string_view> SheetRenameTransform::transform_external_sheet(std::uint32_t /*book_id*/,
                                                                               std::string_view sheet) const {
  return remap_sheet(sheet);
}

}  // namespace parser
}  // namespace formulon
