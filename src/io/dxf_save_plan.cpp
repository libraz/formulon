#include "io/dxf_save_plan.h"

#include <limits>

#include "utils/index_sort.h"

// @size-budget: 12 KB

namespace formulon::io {
namespace {

/// Three-way order of two part keys.
int ComparePart(const DxfPartKey& lhs, const DxfPartKey& rhs) {
  if (lhs.sheet != rhs.sheet) {
    return lhs.sheet < rhs.sheet ? -1 : 1;
  }
  if (lhs.type != rhs.type) {
    return lhs.type < rhs.type ? -1 : 1;
  }
  if (lhs.ordinal != rhs.ordinal) {
    return lhs.ordinal < rhs.ordinal ? -1 : 1;
  }
  return 0;
}

/// Three-way order of (part, model index).
int ComparePartModel(const DxfPartKey& lhs_part, std::uint32_t lhs_model, const DxfPartKey& rhs_part,
                     std::uint32_t rhs_model) {
  const int part = ComparePart(lhs_part, rhs_part);
  if (part != 0) {
    return part;
  }
  return lhs_model == rhs_model ? 0 : (lhs_model < rhs_model ? -1 : 1);
}

constexpr std::int64_t kLast = std::numeric_limits<std::int64_t>::max();

/// A sheet's place in the registration of non-CF users: the active sheet,
/// then tab order; a part no sheet relates comes last.
std::int64_t SheetRank(std::int32_t sheet, std::int32_t active_sheet) {
  if (sheet < 0) {
    return kLast;
  }
  return sheet == active_sheet ? -1 : sheet;
}

/// Registration key, compared field by field: group (styles-part users, other
/// non-CF users, CF), sheet, CF block, CF priority, part type, part ordinal,
/// document position.
struct RegistrationKey {
  static constexpr int kFields = 7;
  std::int64_t field[kFields] = {};

  bool operator<(const RegistrationKey& rhs) const {
    for (int i = 0; i < kFields; ++i) {
      if (field[i] != rhs.field[i]) {
        return field[i] < rhs.field[i];
      }
    }
    return false;
  }
};

RegistrationKey KeyOf(const DxfUse& use, std::int32_t active_sheet) {
  RegistrationKey key;
  std::int64_t group = 1;
  std::int64_t sheet = SheetRank(use.part.sheet, active_sheet);
  std::int64_t block = 0;
  std::int64_t priority = 0;
  if (use.kind == DxfUseKind::kConditionalFormat) {
    group = 2;
    sheet = use.part.sheet < 0 ? kLast : use.part.sheet;
    block = use.block;
    priority = use.priority;
  } else if (use.part.type == DxfPartType::kStyles) {
    group = 0;
    sheet = 0;
  }
  key.field[0] = group;
  key.field[1] = sheet;
  key.field[2] = block;
  key.field[3] = priority;
  key.field[4] = static_cast<std::int64_t>(use.part.type);
  key.field[5] = use.part.ordinal;
  key.field[6] = use.position;
  return key;
}

}  // namespace

bool DxfSavePlan::identity() const {
  return identity_;
}

std::optional<std::uint32_t> DxfSavePlan::saved(const DxfPartKey& part, std::uint32_t model_index,
                                                std::uint32_t nth) const {
  std::size_t low = 0;
  std::size_t high = entries_.size();
  while (low < high) {
    const std::size_t mid = low + (high - low) / 2U;
    const Entry& entry = entries_[mid];
    const int order = ComparePartModel(entry.part, entry.model_index, part, model_index);
    if (order < 0 || (order == 0 && entry.nth < nth)) {
      low = mid + 1U;
    } else {
      high = mid;
    }
  }
  if (low < entries_.size() && entries_[low].part == part && entries_[low].model_index == model_index &&
      entries_[low].nth == nth) {
    return entries_[low].saved;
  }
  if (model_index >= first_.size()) {
    return std::nullopt;
  }
  return first_[model_index];
}

std::optional<std::uint32_t> DxfPartCursor::next(std::uint32_t model_index) {
  if (model_index >= plan_->first_.size()) {
    return std::nullopt;
  }
  if (seen_.empty()) {
    seen_.assign(plan_->first_.size(), 0U);
  }
  return plan_->saved(part_, model_index, seen_[model_index]++);
}

DxfSavePlan build_dxf_save_plan(std::size_t model_count, const std::vector<DxfUse>& uses, bool blocked,
                                std::int32_t active_sheet) {
  if (blocked) {
    return identity_dxf_plan(model_count);
  }
  const auto use_count = static_cast<std::uint32_t>(uses.size());
  // Uses grouped by (part, model); the sort is stable, so each group keeps document order.
  const auto part_model_less = [&uses](std::uint32_t lhs, std::uint32_t rhs) {
    return ComparePartModel(uses[lhs].part, uses[lhs].model_index, uses[rhs].part, uses[rhs].model_index) < 0;
  };
  std::vector<std::uint32_t> by_part;
  sorted_index_order(by_part, use_count, make_index_less(part_model_less));
  // The n-th written reference of each model dxf within its part.
  std::vector<std::uint32_t> nth(uses.size(), 0U);
  std::uint32_t written = 0;
  for (std::size_t k = 0; k < by_part.size(); ++k) {
    const DxfUse& use = uses[by_part[k]];
    if (k == 0U || ComparePartModel(uses[by_part[k - 1U]].part, uses[by_part[k - 1U]].model_index, use.part,
                                    use.model_index) != 0) {
      written = 0;
    }
    if (use.written) {
      nth[by_part[k]] = written++;
    }
  }
  std::vector<RegistrationKey> keys;
  keys.reserve(uses.size());
  for (const DxfUse& use : uses) {
    keys.push_back(KeyOf(use, active_sheet));
  }
  const auto registration_less = [&keys](std::uint32_t lhs, std::uint32_t rhs) { return keys[lhs] < keys[rhs]; };
  std::vector<std::uint32_t> registered;
  sorted_index_order(registered, use_count, make_index_less(registration_less));
  // Registration index of each use. CF uses of one sheet are contiguous in
  // registration order, and those naming one model dxf share its record.
  std::vector<std::uint32_t> record_of(uses.size(), 0U);
  std::vector<std::uint32_t> record_model;
  std::vector<std::int64_t> cf_sheet(model_count, std::numeric_limits<std::int64_t>::min());
  std::vector<std::uint32_t> cf_record(model_count, 0U);
  for (const std::uint32_t i : registered) {
    const DxfUse& use = uses[i];
    if (use.model_index >= model_count) {
      continue;
    }
    if (use.kind == DxfUseKind::kConditionalFormat) {
      if (cf_sheet[use.model_index] == use.part.sheet) {
        record_of[i] = cf_record[use.model_index];
        continue;
      }
      cf_sheet[use.model_index] = use.part.sheet;
      cf_record[use.model_index] = static_cast<std::uint32_t>(record_model.size());
    }
    record_of[i] = static_cast<std::uint32_t>(record_model.size());
    record_model.push_back(use.model_index);
  }
  // Records are written in reverse registration order.
  const auto count = static_cast<std::uint32_t>(record_model.size());
  DxfSavePlan plan;
  plan.order_.assign(record_model.rbegin(), record_model.rend());
  plan.first_.assign(model_count, std::nullopt);
  for (std::size_t i = 0; i < uses.size(); ++i) {
    if (uses[i].model_index < model_count && !plan.first_[uses[i].model_index].has_value()) {
      plan.first_[uses[i].model_index] = count - 1U - record_of[i];
    }
  }
  // In (part, model, nth) order, as `saved` searches them.
  for (const std::uint32_t i : by_part) {
    if (uses[i].model_index < model_count && uses[i].written) {
      plan.entries_.push_back(DxfSavePlan::Entry{uses[i].part, uses[i].model_index, nth[i], count - 1U - record_of[i]});
    }
  }
  plan.identity_ = plan.order_.size() == model_count;
  for (std::size_t i = 0; plan.identity_ && i < plan.order_.size(); ++i) {
    plan.identity_ = plan.order_[i] == i;
  }
  return plan;
}

DxfSavePlan identity_dxf_plan(std::size_t model_count) {
  DxfSavePlan plan;
  plan.order_.reserve(model_count);
  plan.first_.reserve(model_count);
  for (std::size_t i = 0; i < model_count; ++i) {
    plan.order_.push_back(static_cast<std::uint32_t>(i));
    plan.first_.emplace_back(static_cast<std::uint32_t>(i));
  }
  plan.identity_ = true;
  return plan;
}

}  // namespace formulon::io
