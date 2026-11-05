//
// Translation of the C chart patch structs into the C++ patch types, keeping
// the first validation failure (`PatchReader`).

#ifndef FORMULON_C_API_PARTS_CHART_PATCH_READER_H_
#define FORMULON_C_API_PARTS_CHART_PATCH_READER_H_

#include <cstdint>
#include <optional>
#include <string>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "chart/chart_edit.h"
#include "utils/error.h"

namespace formulon {
namespace c_api {
namespace parts {

/// Translates patch fields, keeping the first validation failure.
class PatchReader {
 public:
  explicit PatchReader(const char* api) : api_(api) {}

  fm_status_t status() const { return rc_; }

  template <typename T>
  void field(std::int32_t op, T value, ChartPatchField<T>& out, const char* name) {
    if (!op_ok(op, name)) {
      return;
    }
    if (op == FM_CHART_PATCH_CLEAR) {
      out = ChartPatchField<T>::Clear();
    } else if (op == FM_CHART_PATCH_SET) {
      out = ChartPatchField<T>::Set(std::move(value));
    }
  }

  template <typename E>
  void enum_field(std::int32_t op, std::int32_t value, E max, ChartPatchField<E>& out, const char* name) {
    if (op == FM_CHART_PATCH_SET && !enum_ok(value, max, name)) {
      return;
    }
    field(op, static_cast<E>(value), out, name);
  }

  // A removable boolean; a `SET` value other than 0/1 is rejected.
  void bool_field(std::int32_t op, std::int32_t value, ChartPatchField<bool>& out, const char* name) {
    if (op == FM_CHART_PATCH_SET && !enum_ok(value, 1, name)) {
      return;
    }
    field(op, value == 1, out, name);
  }

  void text_field(std::int32_t op, const char* value, ChartPatchField<std::string>& out, const char* name) {
    if (text_ok(value, name)) {
      field(op, std::string(value != nullptr ? value : ""), out, name);
    }
  }

  // Records a malformed-UTF-8 `value` as the failure; NULL passes.
  bool text_ok(const char* value, const char* name) {
    if (rc_ == 0 && value != nullptr) {
      rc_ = check_utf8(value, api_, "field", name);
    }
    return rc_ == 0;
  }

  template <typename T>
  void optional(std::int32_t op, T value, std::optional<T>& out, const char* name) {
    if (!op_ok(op, name)) {
      return;
    }
    if (op == FM_CHART_PATCH_CLEAR) {
      fail(name, "cannot be cleared");
    } else if (op == FM_CHART_PATCH_SET) {
      out = std::move(value);
    }
  }

  void flag(std::int32_t op, std::int32_t value, std::optional<bool>& out, const char* name) {
    optional(op, value != 0, out, name);
  }

  template <typename E>
  void enum_optional(std::int32_t op, std::int32_t value, E max, std::optional<E>& out, const char* name) {
    if (op == FM_CHART_PATCH_SET && !enum_ok(value, max, name)) {
      return;
    }
    optional(op, static_cast<E>(value), out, name);
  }

  bool enum_ok(std::int64_t value, std::int64_t max, const char* name) {
    if (rc_ == 0) {
      rc_ = check_enum_domain(value, max, api_, name);
    }
    return rc_ == 0;
  }

  template <typename E>
  bool enum_ok(std::int64_t value, E max, const char* name) {
    return enum_ok(value, static_cast<std::int64_t>(max), name);
  }

  void fail(const char* name, const char* why) {
    if (rc_ == 0) {
      rc_ =
          set_binding_error(FormulonErrorCode::kInvalidArgument, (std::string(api_) + ": " + name + " " + why).c_str());
    }
  }

 private:
  bool op_ok(std::int32_t op, const char* name) {
    return enum_ok(op, static_cast<std::int64_t>(FM_CHART_PATCH_SET), name);
  }

  const char* api_;
  fm_status_t rc_ = 0;
};

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_CHART_PATCH_READER_H_
