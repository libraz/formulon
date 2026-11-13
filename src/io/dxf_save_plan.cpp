#include "io/dxf_save_plan.h"

namespace formulon::io {

std::optional<std::uint32_t> DxfSavePlan::map(std::uint32_t model_index) const {
  if (model_index >= saved_.size()) {
    return std::nullopt;
  }
  return saved_[model_index];
}

bool DxfSavePlan::identity() const {
  if (order_.size() != saved_.size()) {
    return false;
  }
  for (std::size_t i = 0; i < order_.size(); ++i) {
    if (order_[i] != i) {
      return false;
    }
  }
  return true;
}

DxfSavePlan build_dxf_save_plan(std::size_t model_count, const std::vector<DxfUse>& uses, bool blocked) {
  if (blocked) {
    return identity_dxf_plan(model_count);
  }
  std::vector<bool> used(model_count, false);
  for (const DxfUse& use : uses) {
    if (use.model_index < model_count) {
      used[use.model_index] = true;
    }
  }
  DxfSavePlan plan;
  plan.saved_.assign(model_count, std::nullopt);
  for (std::size_t i = 0; i < model_count; ++i) {
    if (used[i]) {
      plan.saved_[i] = static_cast<std::uint32_t>(plan.order_.size());
      plan.order_.push_back(static_cast<std::uint32_t>(i));
    }
  }
  return plan;
}

DxfSavePlan identity_dxf_plan(std::size_t model_count) {
  DxfSavePlan plan;
  plan.order_.reserve(model_count);
  plan.saved_.reserve(model_count);
  for (std::size_t i = 0; i < model_count; ++i) {
    plan.order_.push_back(static_cast<std::uint32_t>(i));
    plan.saved_.emplace_back(static_cast<std::uint32_t>(i));
  }
  return plan;
}

}  // namespace formulon::io
