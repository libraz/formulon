//
// Concrete `RefTransform` implementations used for whole-workbook
// edits. The relative-shift transform lives next to the walker in
// `ast_shift.cpp` because it has no external dependencies; the
// identity-aware transforms here pull in additional context (sheet name
// matching, optional new sheet name to emit) and would bloat the walker
// header otherwise.

#ifndef FORMULON_PARSER_REF_TRANSFORMS_H_
#define FORMULON_PARSER_REF_TRANSFORMS_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "parser/ast_shift.h"
#include "parser/reference.h"

namespace formulon {
namespace parser {

/// Renames every reference that names `old_name` to `new_name`.
///
/// Matching is ASCII-case-insensitive (Excel's sheet-name comparison
/// rule). Quoting of the new sheet name is recomputed from the new bytes:
/// if it cannot be represented unambiguously in bare A1 syntax, the
/// `Reference.sheet_quoted` flag is set so `format_a1` round-trips with
/// the canonical quoted form.
///
/// 3-D references store their sheet endpoints in a separate span slot. The
/// dedicated `apply_ref3d_span` hook maps those workbook-local endpoint names.
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
  std::optional<Ref3DSheetSpan> apply_ref3d_span(std::string_view begin, std::string_view end) const override;

 private:
  std::string_view old_name_;
  std::string_view new_name_;
  // Pre-computed quoting decision for the new name. Computing this once at
  // construction time keeps `apply` allocation-free.
  bool new_name_needs_quotes_;
};

/// Rewrites workbook-local references after removing one sheet.
///
/// `pre_removal_sheet_order` is the workbook's sheet order before the removal;
/// the transform copies the string views so callers may pass a temporary
/// vector, but the referenced sheet-name bytes must remain alive for the
/// duration of `shift_refs`. `removed_index` is an index into that order.
/// Ordinary qualified references naming the removed sheet collapse to
/// `#REF!`; local, unqualified references and references to other sheets are
/// left untouched.
///
/// For a 3-D span, removing a middle sheet leaves the textual endpoints
/// unchanged. Removing an endpoint moves that endpoint one sheet inward in
/// the original direction (including reverse spans). A span whose only sheet
/// is removed collapses to `#REF!`. An unresolved endpoint is preserved when
/// the removed sheet is not resolved by either endpoint; if the other endpoint
/// resolves to the removed sheet, the unresolved counterpart collapses to
/// `#REF!` because the inward direction cannot be determined.
class SheetRemovalTransform final : public RefTransform {
 public:
  SheetRemovalTransform(const std::vector<std::string_view>& pre_removal_sheet_order, std::uint32_t removed_index);

  std::optional<Reference> apply(const Reference& ref) const override;
  std::optional<Ref3DSheetSpan> apply_ref3d_span(std::string_view begin, std::string_view end) const override;

 private:
  static constexpr std::size_t kInvalidSheetIndex = static_cast<std::size_t>(-1);

  std::size_t find_sheet(std::string_view sheet) const noexcept;
  bool is_removed(std::size_t index) const noexcept;

  std::vector<std::string_view> pre_removal_sheet_order_;
  std::size_t removed_index_;
};

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
/// `index + count` shift down by `count`. For a range whose endpoints are
/// both in this transform's scope, a deletion that catches one endpoint
/// shrinks the surviving interval just as Excel does; the whole range becomes
/// `#REF!` only when every coordinate in it is deleted.
class RowColShiftTransform final : public RefTransform {
 public:
  RowColShiftTransform(std::string_view target_sheet, RowColAxis axis, RowColEdit edit, std::uint32_t index,
                       std::uint32_t count, bool local_means_target = false) noexcept;

  std::optional<Reference> apply(const Reference& ref) const override;
  std::optional<std::pair<Reference, Reference>> apply_range(const Reference& lhs, const Reference& rhs) const override;
  bool preserves_ref3d_coordinates() const noexcept override { return true; }

 private:
  // Returns the rewritten coordinate, or std::nullopt for #REF!.
  std::optional<std::uint32_t> shift_axis(std::uint32_t coord, std::uint32_t bound) const noexcept;

  // Rewrites an inclusive interval on the edited axis. Unlike shift_axis,
  // this clamps an interval around a deletion to its surviving coordinates.
  std::optional<std::pair<std::uint32_t, std::uint32_t>> shift_interval(std::uint32_t first,
                                                                        std::uint32_t last) const noexcept;

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
