// @size-budget: 1 KB
//
// The thread-local error-origin slot. See `error_origin.h`.

#include "eval/error_origin.h"

#include <cstdint>

#include "sheet_metadata.h"
#include "value.h"

namespace formulon::eval {

namespace {

struct OriginSlot {
  ErrorOrigin origin = ErrorOrigin::kNone;
  CellErrorDetail detail;
};

thread_local OriginSlot g_origin_slot;

}  // namespace

void note_error_origin(CellErrorDetail detail) noexcept {
  g_origin_slot.origin = ErrorOrigin::kMeasured;
  g_origin_slot.detail = detail;
}

void note_calc_sub_type(std::uint32_t sub_type) noexcept {
  note_error_origin(CellErrorDetail{ErrorCode::Calc, CellErrorDetail::Shape::kSubType, sub_type, 0U});
}

void note_unknown_error_origin() noexcept {
  g_origin_slot.origin = ErrorOrigin::kUnknown;
}

void note_spill_refused_here() noexcept {
  g_origin_slot.origin = ErrorOrigin::kRefusedHere;
}

void forget_error_origin() noexcept {
  g_origin_slot.origin = ErrorOrigin::kNone;
}

ErrorOrigin error_origin() noexcept {
  return g_origin_slot.origin;
}

CellErrorDetail noted_error_detail() noexcept {
  return g_origin_slot.detail;
}

ErrorOriginSnapshot snapshot_error_origin() noexcept {
  return ErrorOriginSnapshot{g_origin_slot.origin, g_origin_slot.detail};
}

void restore_error_origin(const ErrorOriginSnapshot& snapshot) noexcept {
  g_origin_slot.origin = snapshot.origin;
  g_origin_slot.detail = snapshot.detail;
}

ErrorOriginScope::ErrorOriginScope() noexcept
    : saved_origin_(g_origin_slot.origin), saved_detail_(g_origin_slot.detail) {
  g_origin_slot.origin = ErrorOrigin::kNone;
}

ErrorOriginScope::~ErrorOriginScope() noexcept {
  g_origin_slot.origin = saved_origin_;
  g_origin_slot.detail = saved_detail_;
}

}  // namespace formulon::eval
