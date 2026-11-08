//
// Translation of the C chart patch structs into the C++ patch types, keeping
// the first validation failure (`PatchReader`).

#ifndef FORMULON_C_API_PARTS_CHART_PATCH_READER_H_
#define FORMULON_C_API_PARTS_CHART_PATCH_READER_H_

#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "chart/chart_edit.h"
#include "utils/error.h"

namespace formulon {
namespace c_api {
namespace parts {

/// Why `value` cannot be a chart or shape colour (None, Rgb or Theme with a tint in [-1, 1] and an
/// index below 12); NULL when it can.
inline const char* color_spec_problem(const fm_color_spec_t& value) {
  constexpr std::uint32_t kThemeCount = 12U;
  if (value.kind == FM_COLOR_KIND_THEME) {
    if (value.theme >= kThemeCount) {
      return "theme index must be below 12";
    }
    if (!(value.tint >= -1.0 && value.tint <= 1.0)) {
      return "tint must be within [-1, 1]";
    }
    return nullptr;
  }
  if (value.kind != FM_COLOR_KIND_NONE && value.kind != FM_COLOR_KIND_RGB) {
    return "must be a None, Rgb or Theme colour";
  }
  return nullptr;
}

/// Translates patch fields, keeping the first validation failure.
class PatchReader {
 public:
  fm_status_t status() const { return rc_; }

  template <typename T>
  void field(std::int32_t op, T value, ChartPatchField<T>& out, const char* name) {
    if (!op_ok(op, name)) {
      return;
    }
    if (op == FM_CHART_PATCH_OP_CLEAR) {
      out = ChartPatchField<T>::Clear();
    } else if (op == FM_CHART_PATCH_OP_SET) {
      out = ChartPatchField<T>::Set(std::move(value));
    }
  }

  template <typename E>
  void enum_field(std::int32_t op, std::int32_t value, E max, ChartPatchField<E>& out, const char* name) {
    if (op == FM_CHART_PATCH_OP_SET && !enum_ok(value, max, name)) {
      return;
    }
    field(op, static_cast<E>(value), out, name);
  }

  // An enum whose C order differs from the model's: `from_c` maps a value
  // already inside `[min, max]`.
  template <typename E>
  void enum_field(std::int32_t op, std::int32_t value, std::int64_t min, std::int64_t max, E (*from_c)(std::int32_t),
                  ChartPatchField<E>& out, const char* name) {
    if (op == FM_CHART_PATCH_OP_SET && !enum_ok(value, min, max, name)) {
      return;
    }
    field(op, from_c(value), out, name);
  }

  // A removable boolean; a `SET` value other than 0/1 is rejected.
  void bool_field(std::int32_t op, std::int32_t value, ChartPatchField<bool>& out, const char* name) {
    if (op == FM_CHART_PATCH_OP_SET && !enum_ok(value, 1, name)) {
      return;
    }
    field(op, value == 1, out, name);
  }

  // A colour patch: a `SET` value must be None, Rgb or a valid Theme colour.
  void color_field(std::int32_t op, const fm_color_spec_t& value, ChartPatchField<formulon::ColorSpec>& out,
                   const char* name) {
    if (op == FM_CHART_PATCH_OP_SET) {
      if (const char* why = color_spec_problem(value); why != nullptr) {
        fail(name, why);
        return;
      }
    }
    field(op, color_spec_from_c(value), out, name);
  }

  // A closed set the model keeps as its OOXML token; `tokens` is indexed by
  // the C ordinal.
  template <std::size_t N>
  void token_field(std::int32_t op, std::int32_t value, const std::array<std::string_view, N>& tokens,
                   ChartPatchField<std::string>& out, const char* name) {
    if (op == FM_CHART_PATCH_OP_SET && !enum_ok(value, 0, static_cast<std::int64_t>(N) - 1, name)) {
      return;
    }
    field(op, op == FM_CHART_PATCH_OP_SET ? std::string(tokens[static_cast<std::size_t>(value)]) : std::string(), out,
          name);
  }

  void text_field(std::int32_t op, const char* value, ChartPatchField<std::string>& out, const char* name) {
    if (text_ok(value, name)) {
      field(op, std::string(value != nullptr ? value : ""), out, name);
    }
  }

  // Records a malformed-UTF-8 `value` as the failure; NULL passes.
  bool text_ok(const char* value, const char* name) {
    if (rc_ == 0 && value != nullptr) {
      rc_ = check_utf8(value, "field", name);
    }
    return rc_ == 0;
  }

  template <typename T>
  void optional(std::int32_t op, T value, std::optional<T>& out, const char* name) {
    if (!op_ok(op, name)) {
      return;
    }
    if (op == FM_CHART_PATCH_OP_CLEAR) {
      fail(name, "cannot be cleared");
    } else if (op == FM_CHART_PATCH_OP_SET) {
      out = std::move(value);
    }
  }

  void flag(std::int32_t op, std::int32_t value, std::optional<bool>& out, const char* name) {
    optional(op, value != 0, out, name);
  }

  template <typename E>
  void enum_optional(std::int32_t op, std::int32_t value, E max, std::optional<E>& out, const char* name) {
    if (op == FM_CHART_PATCH_OP_SET && !enum_ok(value, max, name)) {
      return;
    }
    optional(op, static_cast<E>(value), out, name);
  }

  template <typename E>
  void enum_optional(std::int32_t op, std::int32_t value, std::int64_t min, std::int64_t max, E (*from_c)(std::int32_t),
                     std::optional<E>& out, const char* name) {
    if (op == FM_CHART_PATCH_OP_SET && !enum_ok(value, min, max, name)) {
      return;
    }
    optional(op, from_c(value), out, name);
  }

  bool enum_ok(std::int64_t value, std::int64_t min, std::int64_t max, const char* name) {
    if (rc_ == 0) {
      rc_ = check_enum_domain(value, min, max, name);
    }
    return rc_ == 0;
  }

  bool enum_ok(std::int64_t value, std::int64_t max, const char* name) { return enum_ok(value, 0, max, name); }

  template <typename E>
  bool enum_ok(std::int64_t value, E max, const char* name) {
    return enum_ok(value, static_cast<std::int64_t>(max), name);
  }

  void fail(const char* name, const char* why) {
    if (rc_ == 0) {
      rc_ = invalid_argument((std::string(name) + " " + why).c_str());
    }
  }

  /// Records `rc` (a status another check produced) as the failure when it is the first.
  bool check(fm_status_t rc) {
    if (rc_ == 0) {
      rc_ = rc;
    }
    return rc_ == 0;
  }

 private:
  bool op_ok(std::int32_t op, const char* name) {
    return enum_ok(op, static_cast<std::int64_t>(FM_CHART_PATCH_OP_SET), name);
  }

  fm_status_t rc_ = 0;
};

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_CHART_PATCH_READER_H_
