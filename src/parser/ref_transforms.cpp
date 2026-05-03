// Copyright 2026 libraz. Licensed under the MIT License.

#include "parser/ref_transforms.h"

#include <cstdint>
#include <optional>
#include <string_view>

#include "parser/reference.h"
#include "sheet.h"
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

RowColShiftTransform::RowColShiftTransform(std::string_view target_sheet, RowColAxis axis, RowColEdit edit,
                                           std::uint32_t index, std::uint32_t count, bool local_means_target) noexcept
    : target_sheet_(target_sheet),
      axis_(axis),
      edit_(edit),
      index_(index),
      count_(count),
      local_means_target_(local_means_target) {}

bool RowColShiftTransform::sheet_in_scope(std::string_view sheet) const noexcept {
  if (sheet.empty()) {
    // Reference is local to its formula's owning sheet. The caller
    // controls whether that owning sheet is the target via the
    // `local_means_target` constructor flag.
    return local_means_target_;
  }
  if (target_sheet_.empty()) {
    // Local-scope transform with no named target: a sheet-qualified
    // reference always falls outside the local scope.
    return false;
  }
  return strings::case_insensitive_eq(sheet, target_sheet_);
}

std::optional<std::uint32_t> RowColShiftTransform::shift_axis(std::uint32_t coord, std::uint32_t bound) const noexcept {
  if (edit_ == RowColEdit::kInsert) {
    if (coord < index_) {
      return coord;
    }
    const std::uint64_t shifted = static_cast<std::uint64_t>(coord) + count_;
    if (shifted >= bound) {
      return std::nullopt;
    }
    return static_cast<std::uint32_t>(shifted);
  }
  // Delete.
  if (coord < index_) {
    return coord;
  }
  if (coord < index_ + count_) {
    return std::nullopt;
  }
  return coord - count_;
}

std::optional<Reference> RowColShiftTransform::apply(const Reference& ref) const {
  if (!sheet_in_scope(ref.sheet)) {
    return ref;
  }
  Reference out = ref;
  if (axis_ == RowColAxis::kRow) {
    // Whole-column references are anchored along the column axis only;
    // their row coordinate is meaningless and should not gate a #REF!
    // collapse on a row edit.
    if (ref.is_full_col) {
      return out;
    }
    const std::optional<std::uint32_t> shifted = shift_axis(ref.row, Sheet::kMaxRows);
    if (!shifted.has_value()) {
      return std::nullopt;
    }
    out.row = *shifted;
    return out;
  }
  // Column axis.
  if (ref.is_full_row) {
    return out;
  }
  const std::optional<std::uint32_t> shifted = shift_axis(ref.col, Sheet::kMaxCols);
  if (!shifted.has_value()) {
    return std::nullopt;
  }
  out.col = *shifted;
  return out;
}

}  // namespace parser
}  // namespace formulon
