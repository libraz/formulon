//
// Save-time plan for the differential-format (dxf) table: which model dxfs a
// save writes, in which order, and the saved index each reference takes. The
// model is never touched; every emitter of either format consults one plan,
// and a reference whose model index the plan does not map is omitted.

#ifndef FORMULON_IO_DXF_SAVE_PLAN_H_
#define FORMULON_IO_DXF_SAVE_PLAN_H_

#include <cstddef>
#include <cstdint>
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

/// One reference to a model dxf that a save will write.
struct DxfUse {
  std::uint32_t model_index = 0;
  DxfUseKind kind = DxfUseKind::kConditionalFormat;
  std::int32_t sheet = -1;     ///< Tab index, or -1 for the styles part.
  std::uint32_t part = 0;      ///< Ordinal of the table / pivot / autoFilter within the sheet.
  std::uint32_t position = 0;  ///< Document position within that part.
};

/// The dxf references a save writes; `blocked` asks for the identity plan.
struct DxfUseSet {
  std::vector<DxfUse> uses;
  bool blocked = false;
};

/// The dxfs a save writes and the saved index of each model dxf.
class DxfSavePlan {
 public:
  /// Model indices in write order.
  const std::vector<std::uint32_t>& order() const { return order_; }

  /// Saved index of `model_index`; nullopt when it is unused or not a model dxf.
  std::optional<std::uint32_t> map(std::uint32_t model_index) const;

  /// Number of dxfs the save writes.
  std::size_t saved_count() const { return order_.size(); }

  /// True when every model dxf is written at its own index.
  bool identity() const;

 private:
  friend DxfSavePlan build_dxf_save_plan(std::size_t model_count, const std::vector<DxfUse>& uses, bool blocked);
  friend DxfSavePlan identity_dxf_plan(std::size_t model_count);

  std::vector<std::uint32_t> order_;
  std::vector<std::optional<std::uint32_t>> saved_;  ///< Indexed by model index.
};

/// The plan writing exactly the model dxfs `uses` names, in ascending model
/// order. A use naming no model dxf is ignored. `blocked` yields the identity
/// plan; an empty `uses` without it yields an empty table.
DxfSavePlan build_dxf_save_plan(std::size_t model_count, const std::vector<DxfUse>& uses, bool blocked);

/// The plan writing every model dxf at its own index.
DxfSavePlan identity_dxf_plan(std::size_t model_count);

}  // namespace formulon::io

#endif  // FORMULON_IO_DXF_SAVE_PLAN_H_
