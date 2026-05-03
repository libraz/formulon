// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Concrete `RefTransform` implementations used for whole-workbook
// edits. The relative-shift transform lives next to the walker in
// `ast_shift.cpp` because it has no external dependencies; the
// identity-aware transforms here pull in additional context (sheet name
// matching, optional new sheet name to emit) and would bloat the walker
// header otherwise.

#ifndef FORMULON_PARSER_REF_TRANSFORMS_H_
#define FORMULON_PARSER_REF_TRANSFORMS_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#include "parser/ast_shift.h"
#include "parser/reference.h"

namespace formulon {
namespace parser {

/// Renames every reference that names `old_name` to `new_name`.
///
/// Matching is ASCII-case-insensitive (Excel's sheet-name comparison
/// rule). Quoting of the new sheet name is recomputed from the new bytes:
/// if any byte falls outside `[A-Za-z0-9_.]`, or the name is empty, the
/// `Reference.sheet_quoted` flag is set so `format_a1` round-trips with
/// the canonical quoted form.
///
/// External references store their sheet name in a separate slot from the
/// inner cell's `Reference.sheet`. The external-sheet hook
/// (`transform_external_sheet`) reports the new name when the slot
/// matches; the walker is responsible for interning it into the AST
/// arena.
///
/// `old_name` and `new_name` must outlive the transform.
class SheetRenameTransform final : public RefTransform {
 public:
  SheetRenameTransform(std::string_view old_name, std::string_view new_name) noexcept;

  /// Pure mapping from old to new for a single sheet field. Public so
  /// callers that want to manipulate refs outside the AST walker (for
  /// example, table-range rebinding) can reuse the logic.
  std::optional<std::string_view> remap_sheet(std::string_view sheet) const noexcept;

  std::optional<Reference> apply(const Reference& ref) const override;
  std::optional<std::string_view> transform_external_sheet(std::uint32_t book_id,
                                                           std::string_view sheet) const override;

 private:
  std::string_view old_name_;
  std::string_view new_name_;
  // Pre-computed quoting decision for the new name. Computing this once at
  // construction time keeps `apply` allocation-free.
  bool new_name_needs_quotes_;
};

/// Returns true iff `name` must be wrapped in single quotes when written
/// into a sheet-qualified A1 reference. Excel quotes any name containing a
/// byte outside `[A-Za-z0-9_.]`, plus the empty name. UTF-8 bytes (high
/// bit set) are conservatively treated as needing quoting because the
/// tokenizer's bare-Ident rule rejects them at the start of a sheet name.
bool sheet_name_needs_quoting(std::string_view name) noexcept;

/// Direction of a row/column structural edit.
enum class RowColEdit : std::uint8_t {
  kInsert,  ///< `count` rows/cols inserted starting at `index`.
  kDelete,  ///< `count` rows/cols deleted starting at `index`.
};

/// Axis on which a row/column structural edit operates.
enum class RowColAxis : std::uint8_t {
  kRow,
  kCol,
};

/// Reference rewriter for row / column insert and delete operations.
///
/// The transform applies to references whose sheet field matches
/// `target_sheet` case-insensitively. References with an empty sheet
/// field are local to the formula's owning sheet; whether they fall in
/// scope depends on the formula's location, which is information the
/// per-Reference walker does not carry. The caller therefore runs the
/// transform once per sheet, toggling `local_means_target` when the
/// formula being rewritten lives on the target sheet itself.
///
/// Insert: every coordinate at or past `index` shifts by `count`.
/// Coordinates that would land past the sheet bound (`Sheet::kMaxRows` /
/// `kMaxCols`) collapse to `#REF!` via `std::nullopt`.
///
/// Delete: coordinates strictly less than `index` are unchanged; coords
/// in `[index, index + count)` collapse to `#REF!`; coords at or past
/// `index + count` shift down by `count`.
///
/// Range endpoint clamping (Excel's "shrink the range" behaviour for
/// deletes that catch one endpoint but not the other) is a follow-up
/// enhancement; the simple per-endpoint policy here matches Excel for
/// every range whose endpoints are both inside or both outside the
/// affected interval, and falls back to `#REF!` for the clamp case.
class RowColShiftTransform final : public RefTransform {
 public:
  RowColShiftTransform(std::string_view target_sheet, RowColAxis axis, RowColEdit edit, std::uint32_t index,
                       std::uint32_t count, bool local_means_target = false) noexcept;

  std::optional<Reference> apply(const Reference& ref) const override;

 private:
  // Returns the rewritten coordinate, or std::nullopt for #REF!.
  std::optional<std::uint32_t> shift_axis(std::uint32_t coord, std::uint32_t bound) const noexcept;

  // True iff `ref.sheet` falls within the transform's scope.
  bool sheet_in_scope(std::string_view sheet) const noexcept;

  std::string_view target_sheet_;
  RowColAxis axis_;
  RowColEdit edit_;
  std::uint32_t index_;
  std::uint32_t count_;
  bool local_means_target_;
};

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_REF_TRANSFORMS_H_
