//
// Translation of the C chart patch structs into the C++ patch types, keeping
// the first validation failure (`PatchReader`), and of the model's paint
// records into the C ones; the chart and shape parts share both.

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
inline const char* color_spec_problem(const fm_color_t& value) {
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
    if (op == FM_PATCH_OP_CLEAR) {
      out = ChartPatchField<T>::Clear();
    } else if (op == FM_PATCH_OP_SET) {
      out = ChartPatchField<T>::Set(std::move(value));
    }
  }

  template <typename E>
  void enum_field(std::int32_t op, std::int32_t value, E max, ChartPatchField<E>& out, const char* name) {
    if (op == FM_PATCH_OP_SET && !enum_ok(value, max, name)) {
      return;
    }
    field(op, static_cast<E>(value), out, name);
  }

  // An enum whose C order differs from the model's: `from_c` maps a value
  // already inside `[min, max]`.
  template <typename E>
  void enum_field(std::int32_t op, std::int32_t value, std::int64_t min, std::int64_t max, E (*from_c)(std::int32_t),
                  ChartPatchField<E>& out, const char* name) {
    if (op == FM_PATCH_OP_SET && !enum_ok(value, min, max, name)) {
      return;
    }
    field(op, from_c(value), out, name);
  }

  // A removable boolean; a `SET` value other than 0/1 is rejected.
  void bool_field(std::int32_t op, std::int32_t value, ChartPatchField<bool>& out, const char* name) {
    if (op == FM_PATCH_OP_SET && !enum_ok(value, 1, name)) {
      return;
    }
    field(op, value == 1, out, name);
  }

  // A colour patch: a `SET` value must be None, Rgb or a valid Theme colour.
  void color_field(std::int32_t op, const fm_color_t& value, ChartPatchField<formulon::ColorSpec>& out,
                   const char* name) {
    if (op == FM_PATCH_OP_SET) {
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
    if (op == FM_PATCH_OP_SET && !enum_ok(value, 0, static_cast<std::int64_t>(N) - 1, name)) {
      return;
    }
    field(op, op == FM_PATCH_OP_SET ? std::string(tokens[static_cast<std::size_t>(value)]) : std::string(), out, name);
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
    if (op == FM_PATCH_OP_CLEAR) {
      fail(name, "cannot be cleared");
    } else if (op == FM_PATCH_OP_SET) {
      out = std::move(value);
    }
  }

  void flag(std::int32_t op, std::int32_t value, std::optional<bool>& out, const char* name) {
    optional(op, value != 0, out, name);
  }

  template <typename E>
  void enum_optional(std::int32_t op, std::int32_t value, E max, std::optional<E>& out, const char* name) {
    if (op == FM_PATCH_OP_SET && !enum_ok(value, max, name)) {
      return;
    }
    optional(op, static_cast<E>(value), out, name);
  }

  template <typename E>
  void enum_optional(std::int32_t op, std::int32_t value, std::int64_t min, std::int64_t max, E (*from_c)(std::int32_t),
                     std::optional<E>& out, const char* name) {
    if (op == FM_PATCH_OP_SET && !enum_ok(value, min, max, name)) {
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
    return enum_ok(op, static_cast<std::int64_t>(FM_PATCH_OP_SET), name);
  }

  fm_status_t rc_ = 0;
};

/// The C colour of a model paint colour; None when absent.
inline fm_color_t paint_color_to_c(const std::optional<formulon::ColorSpec>& color) {
  return color_spec_to_c(color ? *color : formulon::ColorSpec{});
}

/// The transparency percentage of an alpha in thousandths of a percent; 0 without one.
inline double paint_transparency_of(const std::optional<std::uint32_t>& alpha) {
  return alpha ? 100.0 - static_cast<double>(*alpha) / 1000.0 : 0.0;
}

/// A model fill as the C record reports it.
inline fm_chart_fill_t chart_fill_to_c(const formulon::ChartFill& f) {
  fm_chart_fill_t out{};
  out.color = paint_color_to_c(f.color);
  out.has_transparency_pct = f.alpha ? 1 : 0;
  out.transparency_pct = paint_transparency_of(f.alpha);
  out.none = f.none ? 1 : 0;
  out.other = f.other ? 1 : 0;
  return out;
}

/// A model outline as the C record reports it.
inline fm_chart_line_t chart_line_to_c(const formulon::ChartLine& l) {
  constexpr double kEmuPerPt = 12700.0;
  fm_chart_line_t out{};
  out.color = paint_color_to_c(l.color);
  out.has_transparency_pct = l.alpha ? 1 : 0;
  out.transparency_pct = paint_transparency_of(l.alpha);
  out.has_width_pt = l.width_emu ? 1 : 0;
  out.width_pt = l.width_emu ? static_cast<double>(*l.width_emu) / kEmuPerPt : 0.0;
  out.has_dash = l.dash ? 1 : 0;
  out.dash = l.dash ? static_cast<std::int32_t>(*l.dash) : 0;
  out.none = l.none ? 1 : 0;
  out.other = l.other ? 1 : 0;
  return out;
}

/// A C fill patch as the model's; shared by the chart and shape edits.
inline formulon::ChartFillPatch read_fill(PatchReader& r, const fm_chart_fill_patch_t& p) {
  formulon::ChartFillPatch out;
  r.color_field(p.color_op, p.color, out.color, "fill.color_op");
  r.field(p.transparency_pct_op, p.transparency_pct, out.transparency_pct, "fill.transparency_pct_op");
  // CLEAR drops an explicit no-fill as SET 0 does, so a patch clearing every member stays valid.
  if (p.none_op == FM_PATCH_OP_CLEAR) {
    out.none = false;
  } else {
    r.flag(p.none_op, p.none, out.none, "fill.none_op");
  }
  return out;
}

/// A C outline patch as the model's; shared by the chart and shape edits.
inline formulon::ChartLinePatch read_line(PatchReader& r, const fm_chart_line_patch_t& p) {
  formulon::ChartLinePatch out;
  r.color_field(p.color_op, p.color, out.color, "line.color_op");
  r.field(p.transparency_pct_op, p.transparency_pct, out.transparency_pct, "line.transparency_pct_op");
  r.field(p.width_pt_op, p.width_pt, out.width_pt, "line.width_pt_op");
  r.enum_field(p.dash_op, p.dash, formulon::ChartLineDash::kOther, out.dash, "line.dash");
  // CLEAR drops an explicit no-line as SET 0 does, so a patch clearing every member stays valid.
  if (p.none_op == FM_PATCH_OP_CLEAR) {
    out.none = false;
  } else {
    r.flag(p.none_op, p.none, out.none, "line.none_op");
  }
  return out;
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_CHART_PATCH_READER_H_
