//
// Save-time plan for the differential-format (dxf) table: which records a save
// writes, in which order, and the saved index each reference takes. The model
// is never touched. Emitters write model indices; the writer then rewrites
// every written part through one plan, reference by reference.
//
// Excel saves one record per user, so a model dxf named by several users is
// written once per user. Users register in this order and the records are
// written reversed:
// styles-part users; then each sheet's other users, the active sheet first and
// the rest in tab order, its worksheet part before its pivots before its
// tables; then CF users by sheet in tab order, saved block order and
// priority ascending.

#ifndef FORMULON_IO_DXF_SAVE_PLAN_H_
#define FORMULON_IO_DXF_SAVE_PLAN_H_

#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <vector>

namespace formulon::io {

/// The holder kind of one dxf reference.
enum class DxfUseKind : std::uint8_t {
  kConditionalFormat,  ///< A conditional-format rule.
  kTable,              ///< A table's own dxf attribute.
  kTableColumn,        ///< A table column's dxf attribute.
  kTableStyleElement,  ///< A custom table or pivot style element.
  kColorFilter,        ///< An autoFilter colour filter.
  kSortCondition,      ///< A sort condition.
  kPivotFormat,        ///< A pivot table `<format>`.
};

/// The kind of part a reference is written in, in per-sheet registration order.
enum class DxfPartType : std::uint8_t {
  kStyles,     ///< The styles part.
  kWorksheet,  ///< A worksheet body.
  kPivot,      ///< A pivot table part.
  kTable,      ///< A table part.
};

/// One written part.
struct DxfPartKey {
  std::int32_t sheet = -1;  ///< Tab index; -1 for the styles part or a part no sheet relates.
  DxfPartType type = DxfPartType::kStyles;
  std::uint32_t ordinal = 0;  ///< Ordinal of the pivot / table among its sheet's; 0 otherwise.

  friend bool operator==(const DxfPartKey& lhs, const DxfPartKey& rhs) {
    return lhs.sheet == rhs.sheet && lhs.type == rhs.type && lhs.ordinal == rhs.ordinal;
  }
};

/// One reference to a model dxf that a save will write.
struct DxfUse {
  std::uint32_t model_index = 0;
  DxfUseKind kind = DxfUseKind::kConditionalFormat;
  DxfPartKey part;
  std::uint32_t position = 0;  ///< Document position within the part.
  /// A CF rule's priority; unset ranks after every set one.
  std::uint32_t priority = std::numeric_limits<std::uint32_t>::max();
  /// False for a use that takes a record but has no index in the written part
  /// (an xlsx x14-only CF rule's inline dxf).
  bool written = true;
  /// A CF rule's saved block (`<conditionalFormatting>`) ordinal within its sheet.
  std::uint32_t block = 0;
};

/// The dxf references a save writes; `blocked` asks for the identity plan.
/// The uses of one part appear in document order.
struct DxfUseSet {
  std::vector<DxfUse> uses;
  bool blocked = false;
};

/// The records a save writes and the saved index of each reference.
class DxfSavePlan {
 public:
  /// The model dxf of each saved record, in write order; a model dxf may repeat.
  const std::vector<std::uint32_t>& order() const { return order_; }

  /// Number of records the save writes.
  std::size_t saved_count() const { return order_.size(); }

  /// True when every model dxf is written once, at its own index.
  bool identity() const;

  /// Saved index of the `nth` written reference to `model_index` in `part`.
  /// A reference no use describes takes the first record of its model dxf;
  /// nullopt when the model dxf has no record (unused or out of range).
  std::optional<std::uint32_t> saved(const DxfPartKey& part, std::uint32_t model_index, std::uint32_t nth) const;

 private:
  friend DxfSavePlan build_dxf_save_plan(std::size_t model_count, const std::vector<DxfUse>& uses, bool blocked,
                                         std::int32_t active_sheet);
  friend DxfSavePlan identity_dxf_plan(std::size_t model_count);
  friend class DxfPartCursor;

  struct Entry {
    DxfPartKey part;
    std::uint32_t model_index = 0;
    std::uint32_t nth = 0;
    std::uint32_t saved = 0;
  };

  std::vector<std::uint32_t> order_;
  std::vector<std::optional<std::uint32_t>> first_;  ///< Indexed by model index.
  std::vector<Entry> entries_;                       ///< Sorted by (part, model, nth).
  bool identity_ = false;
};

/// The saved indices of one part's references, handed out in document order.
class DxfPartCursor {
 public:
  DxfPartCursor(const DxfSavePlan& plan, DxfPartKey part) : plan_(&plan), part_(part) {}

  /// Saved index of the part's next reference, which names `model_index`.
  std::optional<std::uint32_t> next(std::uint32_t model_index);

 private:
  const DxfSavePlan* plan_;
  DxfPartKey part_;
  std::vector<std::uint32_t> seen_;  ///< References so far, by model index; empty until the first.
};

/// The plan writing one record per use of a model dxf in `uses`, in Excel's
/// order. A use naming no model dxf gets no
/// record. `blocked` yields the identity plan; an empty `uses` without it
/// yields an empty table. `active_sheet` is the tab index of the active sheet.
DxfSavePlan build_dxf_save_plan(std::size_t model_count, const std::vector<DxfUse>& uses, bool blocked,
                                std::int32_t active_sheet);

/// The plan writing every model dxf at its own index.
DxfSavePlan identity_dxf_plan(std::size_t model_count);

}  // namespace formulon::io

#endif  // FORMULON_IO_DXF_SAVE_PLAN_H_
