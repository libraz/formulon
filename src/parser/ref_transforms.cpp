
#include "parser/ref_transforms.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "parser/reference.h"
#include "sheet_name.h"
#include "utils/a1_ref.h"

namespace formulon {
namespace parser {

SheetRenameTransform::SheetRenameTransform(std::string_view old_name, std::string_view new_name,
                                           const sheet_names::NameList* sheet_list) noexcept
    : old_name_(old_name),
      new_name_(new_name),
      sheet_list_(sheet_list),
      new_name_needs_quotes_(local_sheet_needs_quoting_a1(new_name)) {}

std::optional<std::string_view> SheetRenameTransform::remap_sheet(std::string_view sheet) const noexcept {
  if (sheet.empty()) {
    return std::nullopt;
  }
  if (!sheet_names::NameList::denotes(sheet_list_, sheet, old_name_)) {
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

std::optional<RefTransform::Ref3DSheetSpan> SheetRenameTransform::apply_ref3d_span(std::string_view begin,
                                                                                   std::string_view end) const {
  const std::string_view mapped_begin = remap_sheet(begin).value_or(begin);
  const std::string_view mapped_end = remap_sheet(end).value_or(end);
  return Ref3DSheetSpan{mapped_begin, mapped_end};
}

std::optional<std::string_view> SheetRenameTransform::apply_name_sheet(std::string_view sheet) const {
  return remap_sheet(sheet).value_or(sheet);
}

SheetCopyTransform::SheetCopyTransform(std::string_view source, std::string_view copy,
                                       std::vector<std::pair<std::string, std::string>> table_map,
                                       const sheet_names::NameList* sheet_list, const sheet_names::NameList* table_list)
    : rename_(source, copy, sheet_list),
      source_(source),
      copy_(copy),
      table_map_(std::move(table_map)),
      sheet_list_(sheet_list),
      table_list_(table_list) {}

std::optional<Reference> SheetCopyTransform::apply(const Reference& ref) const {
  return rename_.apply(ref);
}

std::optional<RefTransform::Ref3DSheetSpan> SheetCopyTransform::apply_ref3d_span(std::string_view begin,
                                                                                 std::string_view end) const {
  const bool begin_is_source = sheet_names::NameList::denotes(sheet_list_, begin, source_);
  const bool end_is_source = sheet_names::NameList::denotes(sheet_list_, end, source_);
  if (begin_is_source && end_is_source) {
    return Ref3DSheetSpan{copy_, copy_, /*collapse=*/true};
  }
  if (begin_is_source || end_is_source) {
    // A span anchored on the source at one end only has no image on the copy.
    return std::nullopt;
  }
  return Ref3DSheetSpan{begin, end};
}

std::optional<std::string_view> SheetCopyTransform::apply_name_sheet(std::string_view sheet) const {
  return rename_.apply_name_sheet(sheet);
}

std::optional<StructuredRefTarget> SheetCopyTransform::apply_structured_ref(std::string_view table,
                                                                            std::string_view /*column*/) const {
  for (const std::pair<std::string, std::string>& entry : table_map_) {
    if (sheet_names::NameList::denotes(table_list_, table, entry.first)) {
      return StructuredRefTarget{entry.second, {}};
    }
  }
  return StructuredRefTarget{table, {}};
}

SheetRemovalTransform::SheetRemovalTransform(const std::vector<std::string_view>& pre_removal_sheet_order,
                                             std::uint32_t removed_index, std::vector<std::string> rebound_names)
    : pre_removal_sheet_order_(pre_removal_sheet_order),
      removed_index_(removed_index),
      rebound_names_(std::move(rebound_names)) {}

std::size_t SheetRemovalTransform::find_sheet(std::string_view sheet) const noexcept {
  const std::size_t index =
      sheet_names::resolve_by(pre_removal_sheet_order_.size(), [&](std::size_t i, sheet_names::Fold fold) {
        return sheet_names::same(pre_removal_sheet_order_[i], sheet, fold);
      });
  return index == sheet_names::kNotFound ? kInvalidSheetIndex : index;
}

bool SheetRemovalTransform::is_removed(std::size_t index) const noexcept {
  return removed_index_ < pre_removal_sheet_order_.size() && index == removed_index_;
}

std::optional<Reference> SheetRemovalTransform::apply(const Reference& ref) const {
  if (ref.sheet.empty()) {
    return ref;
  }
  const std::size_t sheet_index = find_sheet(ref.sheet);
  if (is_removed(sheet_index)) {
    return std::nullopt;
  }
  return ref;
}

std::optional<std::string_view> SheetRemovalTransform::apply_name_sheet(std::string_view sheet) const {
  if (is_removed(find_sheet(sheet))) {
    return std::nullopt;
  }
  return sheet;
}

bool SheetRemovalTransform::rebinds_name_to_book(std::string_view sheet, std::string_view name) const noexcept {
  if (!is_removed(find_sheet(sheet))) {
    return false;
  }
  return sheet_names::resolve(rebound_names_, name) != sheet_names::kNotFound;
}

std::optional<RefTransform::Ref3DSheetSpan> SheetRemovalTransform::apply_ref3d_span(std::string_view begin,
                                                                                    std::string_view end) const {
  if (removed_index_ >= pre_removal_sheet_order_.size()) {
    // An invalid removal index cannot describe a safe structural edit.
    return std::nullopt;
  }
  const std::size_t begin_index = find_sheet(begin);
  const std::size_t end_index = find_sheet(end);
  const bool begin_removed = is_removed(begin_index);
  const bool end_removed = is_removed(end_index);
  if (!begin_removed && !end_removed) {
    // An unresolved endpoint is not itself evidence that this deletion
    // affects the reference. Preserve existing unresolved spans, as well as
    // resolved spans outside the removed sheet, byte-for-byte.
    return Ref3DSheetSpan{begin, end};
  }
  if (begin_removed && end_removed) {
    // The span is degenerate (the only endpoint names the removed sheet), so
    // no sheet remains to carry the reference.
    return std::nullopt;
  }
  if ((begin_removed && end_index == kInvalidSheetIndex) || (end_removed && begin_index == kInvalidSheetIndex)) {
    // Once one endpoint resolves to the deleted sheet, the other endpoint
    // must resolve too: without it there is no safe direction for the
    // inward step. Only this affected unresolved case collapses to #REF!.
    return std::nullopt;
  }

  const bool forward = begin_index < end_index;
  if (begin_removed) {
    // Move the first endpoint toward the surviving end.
    const std::size_t inward = forward ? begin_index + 1U : begin_index - 1U;
    if (inward >= pre_removal_sheet_order_.size() || inward == removed_index_) {
      return std::nullopt;
    }
    return Ref3DSheetSpan{pre_removal_sheet_order_[inward], end, /*collapse=*/inward == end_index};
  }

  // Move the last endpoint toward the surviving begin. The two endpoint
  // names cannot be equal here because begin_removed is false.
  const std::size_t inward = forward ? end_index - 1U : end_index + 1U;
  if (inward >= pre_removal_sheet_order_.size() || inward == removed_index_) {
    return std::nullopt;
  }
  return Ref3DSheetSpan{begin, pre_removal_sheet_order_[inward], /*collapse=*/inward == begin_index};
}

SheetMoveTransform::SheetMoveTransform(const std::vector<std::string_view>& pre_move_sheet_order, std::uint32_t from,
                                       std::uint32_t to)
    : pre_move_sheet_order_(pre_move_sheet_order.begin(), pre_move_sheet_order.end()), from_(from), to_(to) {}

std::size_t SheetMoveTransform::find_sheet(std::string_view sheet) const noexcept {
  return sheet_names::resolve_by(pre_move_sheet_order_.size(), [&](std::size_t i, sheet_names::Fold fold) {
    return sheet_names::same(pre_move_sheet_order_[i], sheet, fold);
  });
}

std::size_t SheetMoveTransform::post_move_index(std::size_t index) const noexcept {
  if (index == from_) {
    return to_;
  }
  if (from_ < to_ && index > from_ && index <= to_) {
    return index - 1U;
  }
  if (from_ > to_ && index >= to_ && index < from_) {
    return index + 1U;
  }
  return index;
}

std::optional<RefTransform::Ref3DSheetSpan> SheetMoveTransform::apply_ref3d_span(std::string_view begin,
                                                                                 std::string_view end) const {
  const std::size_t begin_index = find_sheet(begin);
  const std::size_t end_index = find_sheet(end);
  if (begin_index == sheet_names::kNotFound || end_index == sheet_names::kNotFound) {
    return Ref3DSheetSpan{begin, end};
  }
  if (begin_index == end_index) {
    return Ref3DSheetSpan{begin, end, /*collapse=*/begin_index == from_};
  }
  const std::size_t first = std::min(begin_index, end_index);
  const std::size_t last = std::max(begin_index, end_index);
  std::size_t new_first = first;
  std::size_t new_last = last;
  if (first == from_ && post_move_index(first) > post_move_index(last)) {
    new_first = first + 1U;
  } else if (last == from_ && post_move_index(last) < post_move_index(first)) {
    new_last = last - 1U;
  } else {
    return Ref3DSheetSpan{begin, end};
  }
  return Ref3DSheetSpan{pre_move_sheet_order_[new_first], pre_move_sheet_order_[new_last], new_first == new_last};
}

RowColShiftTransform::RowColShiftTransform(std::string_view target_sheet, RowColAxis axis, RowColEdit edit,
                                           std::uint32_t index, std::uint32_t count, bool local_means_target,
                                           const sheet_names::NameList* sheet_list) noexcept
    : target_sheet_(target_sheet),
      axis_(axis),
      edit_(edit),
      index_(index),
      count_(count),
      local_means_target_(local_means_target),
      sheet_list_(sheet_list) {}

RowColShiftTransform::RowColShiftTransform(std::string_view target_sheet, RowColAxis axis, RowColEdit edit,
                                           std::uint32_t index, std::uint32_t count, Band band, bool local_means_target,
                                           const sheet_names::NameList* sheet_list) noexcept
    : RowColShiftTransform(target_sheet, axis, edit, index, count, local_means_target, sheet_list) {
  band_ = band;
}

bool RowColShiftTransform::inside_band(std::uint32_t first, std::uint32_t last) const noexcept {
  return !band_.has_value() || (band_->first <= std::min(first, last) && std::max(first, last) <= band_->last);
}

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
  return sheet_names::NameList::denotes(sheet_list_, sheet, target_sheet_);
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
  const std::uint64_t delete_end = static_cast<std::uint64_t>(index_) + count_;
  if (static_cast<std::uint64_t>(coord) < delete_end) {
    return std::nullopt;
  }
  return coord - count_;
}

std::optional<std::pair<std::uint32_t, std::uint32_t>> RowColShiftTransform::shift_interval(
    std::uint32_t first, std::uint32_t last) const noexcept {
  const std::uint32_t bound = a1::axis_extent(axis_ == RowColAxis::kRow);
  const std::uint32_t low = std::min(first, last);
  const std::uint32_t high = std::max(first, last);
  const std::uint64_t delete_begin = index_;
  const std::uint64_t delete_end = delete_begin + count_;

  // An insert, or an interval wholly before or wholly after a deletion, has no
  // surviving gap to clamp; endpoint mapping gives the ordinary structural shift.
  if (edit_ == RowColEdit::kInsert || count_ == 0 || static_cast<std::uint64_t>(high) < delete_begin ||
      static_cast<std::uint64_t>(low) >= delete_end) {
    const std::optional<std::uint32_t> shifted_first = shift_axis(first, bound);
    if (!shifted_first.has_value()) {
      return std::nullopt;
    }
    const std::optional<std::uint32_t> shifted_last = shift_axis(last, bound);
    if (!shifted_last.has_value()) {
      return std::nullopt;
    }
    return std::make_pair(*shifted_first, *shifted_last);
  }

  // The deletion intersects the interval. Keep the first survivor from the
  // original prefix, if any, otherwise the first coordinate after the
  // deleted span. Do the symmetric calculation for the final survivor.
  const bool has_prefix = static_cast<std::uint64_t>(low) < delete_begin;
  const bool has_suffix = static_cast<std::uint64_t>(high) >= delete_end;
  if (!has_prefix && !has_suffix) {
    return std::nullopt;  // every coordinate in the interval was deleted
  }

  const std::uint64_t survivor_low = has_prefix ? low : delete_end - count_;
  const std::uint64_t survivor_high = has_suffix ? static_cast<std::uint64_t>(high) - count_ : delete_begin - 1U;
  if (survivor_low >= bound || survivor_high >= bound || survivor_low > survivor_high) {
    return std::nullopt;
  }

  const std::uint32_t out_low = static_cast<std::uint32_t>(survivor_low);
  const std::uint32_t out_high = static_cast<std::uint32_t>(survivor_high);
  if (first <= last) {
    return std::make_pair(out_low, out_high);
  }
  return std::make_pair(out_high, out_low);
}

std::optional<Reference> RowColShiftTransform::apply(const Reference& ref) const {
  if (!sheet_in_scope(ref.sheet)) {
    return ref;
  }
  Reference out = ref;
  // A band edit leaves whole columns and rows, and cells outside the band, alone.
  if (band_.has_value() &&
      (ref.is_full_col || ref.is_full_row ||
       !inside_band(axis_ == RowColAxis::kRow ? ref.col : ref.row, axis_ == RowColAxis::kRow ? ref.col : ref.row))) {
    return out;
  }
  if (axis_ == RowColAxis::kRow) {
    // Whole-column references are anchored along the column axis only;
    // their row coordinate is meaningless and should not gate a #REF!
    // collapse on a row edit.
    if (ref.is_full_col) {
      return out;
    }
    const std::optional<std::uint32_t> shifted = shift_axis(ref.row, a1::kMaxRows);
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
  const std::optional<std::uint32_t> shifted = shift_axis(ref.col, a1::kMaxCols);
  if (!shifted.has_value()) {
    return std::nullopt;
  }
  out.col = *shifted;
  return out;
}

std::optional<std::pair<Reference, Reference>> RowColShiftTransform::apply_range(const Reference& lhs,
                                                                                 const Reference& rhs) const {
  // A range-level clamp is meaningful only for a single sheet and a
  // well-formed pair of matching whole-axis flags. If either endpoint is
  // outside scope (including a cross-sheet range), retain the base contract:
  // transform each endpoint independently and let a deleted endpoint poison
  // the range.
  if (!sheet_names::NameList::same_entry(sheet_list_, lhs.sheet, rhs.sheet) || !sheet_in_scope(lhs.sheet) ||
      lhs.is_full_col != rhs.is_full_col || lhs.is_full_row != rhs.is_full_row) {
    return RefTransform::apply_range(lhs, rhs);
  }
  if (band_.has_value()) {
    if (lhs.is_full_col || lhs.is_full_row) {
      return std::make_pair(lhs, rhs);
    }
    const bool row_axis = axis_ == RowColAxis::kRow;
    if (!inside_band(row_axis ? lhs.col : lhs.row, row_axis ? rhs.col : rhs.row)) {
      return apply_range_across_band(lhs, rhs);
    }
  }
  // A whole-column range is independent of row edits, and a whole-row range
  // is independent of column edits. Endpoint application is equivalent and
  // keeps this hook conservative for malformed mixed ranges.
  if ((axis_ == RowColAxis::kRow && lhs.is_full_col) || (axis_ == RowColAxis::kCol && lhs.is_full_row)) {
    return RefTransform::apply_range(lhs, rhs);
  }

  if (edit_ == RowColEdit::kInsert) {
    return RefTransform::apply_range(lhs, rhs);
  }

  const std::uint32_t first = axis_ == RowColAxis::kRow ? lhs.row : lhs.col;
  const std::uint32_t last = axis_ == RowColAxis::kRow ? rhs.row : rhs.col;
  const std::optional<std::pair<std::uint32_t, std::uint32_t>> shifted = shift_interval(first, last);
  if (!shifted.has_value()) {
    return std::nullopt;
  }

  Reference out_lhs = lhs;
  Reference out_rhs = rhs;
  if (axis_ == RowColAxis::kRow) {
    out_lhs.row = shifted->first;
    out_rhs.row = shifted->second;
  } else {
    out_lhs.col = shifted->first;
    out_rhs.col = shifted->second;
  }
  return std::make_pair(out_lhs, out_rhs);
}

std::pair<Reference, Reference> RowColShiftTransform::apply_range_across_band(const Reference& lhs,
                                                                              const Reference& rhs) const {
  const bool row_axis = axis_ == RowColAxis::kRow;
  const std::uint32_t axial_low = std::min(row_axis ? lhs.row : lhs.col, row_axis ? rhs.row : rhs.col);
  const std::uint32_t axial_high = std::max(row_axis ? lhs.row : lhs.col, row_axis ? rhs.row : rhs.col);
  const std::uint32_t perp_low = std::min(row_axis ? lhs.col : lhs.row, row_axis ? rhs.col : rhs.row);
  const std::uint32_t perp_high = std::max(row_axis ? lhs.col : lhs.row, row_axis ? rhs.col : rhs.row);
  const std::uint64_t delete_end = static_cast<std::uint64_t>(index_) + count_;
  // A delete whose rows (columns) hold the range's whole axial extent cuts the
  // band's part off a range the band reaches at exactly one perpendicular end.
  const bool covers_low = band_->first <= perp_low && perp_low <= band_->last;
  const bool covers_high = band_->first <= perp_high && perp_high <= band_->last;
  if (edit_ != RowColEdit::kDelete || axial_low < index_ || axial_high >= delete_end || covers_low == covers_high) {
    return std::make_pair(lhs, rhs);
  }
  const std::uint32_t kept_low = covers_low ? band_->last + 1U : perp_low;
  const std::uint32_t kept_high = covers_low ? perp_high : band_->first - 1U;
  Reference out_lhs = lhs;
  Reference out_rhs = rhs;
  std::uint32_t& lhs_perp = row_axis ? out_lhs.col : out_lhs.row;
  std::uint32_t& rhs_perp = row_axis ? out_rhs.col : out_rhs.row;
  const bool lhs_is_low = (row_axis ? lhs.col : lhs.row) <= (row_axis ? rhs.col : rhs.row);
  lhs_perp = lhs_is_low ? kept_low : kept_high;
  rhs_perp = lhs_is_low ? kept_high : kept_low;
  return std::make_pair(out_lhs, out_rhs);
}

}  // namespace parser
}  // namespace formulon
