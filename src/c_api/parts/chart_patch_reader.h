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
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "chart/chart_edit.h"
#include "drawing/drawing_color.h"
#include "utils/error.h"
#include "utils/resource_budget.h"

namespace formulon {
namespace c_api {
namespace parts {

/// Why `value` cannot be a chart or shape colour to set (Rgb, or Theme with a tint in [-1, 1] and an
/// index below 12); NULL when it can. Only `CLEAR` removes a colour.
inline const char* color_spec_problem(const fm_color_t& value) {
  constexpr std::uint32_t kThemeCount = 12U;
  if (value.kind == FM_COLOR_KIND_THEME) {
    if (value.theme >= kThemeCount) {
      return "theme must be in 0..11";
    }
    if (!(value.tint >= -1.0 && value.tint <= 1.0)) {
      return "tint must be in -1..1";
    }
    return nullptr;
  }
  if (value.kind != FM_COLOR_KIND_RGB) {
    return "must be an Rgb or Theme colour; CLEAR removes one";
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

  // A paint colour patch: a `SET` value is converted by `drawing_color_from_c`.
  void drawing_color_field(std::int32_t op, const fm_drawing_color_t& value,
                           ChartPatchField<formulon::DrawingColor>& out, const char* name);

  // A line end patch: a `SET` value must lie inside its enumerations.
  void line_end_field(std::int32_t op, const fm_chart_line_end_t& value, ChartPatchField<formulon::ChartLineEnd>& out,
                      const char* name) {
    if (op == FM_PATCH_OP_SET && !line_end_ok(value, name)) {
      return;
    }
    field(op, line_end_from_c(value), out, name);
  }

  // Whether `value` is a valid line end, recording the failure under `name`.
  bool line_end_ok(const fm_chart_line_end_t& value, const char* name) {
    const std::string at = public_name(name);
    return enum_ok(value.type, FM_CHART_LINE_END_TYPE_ARROW, (at + ".type").c_str()) &&
           enum_ok(value.has_width, 1, (at + ".has_width").c_str()) &&
           (value.has_width == 0 || enum_ok(value.width, FM_CHART_LINE_END_SIZE_LARGE, (at + ".width").c_str())) &&
           enum_ok(value.has_length, 1, (at + ".has_length").c_str()) &&
           (value.has_length == 0 || enum_ok(value.length, FM_CHART_LINE_END_SIZE_LARGE, (at + ".length").c_str()));
  }

  /// The model line end of a validated C one.
  static formulon::ChartLineEnd line_end_from_c(const fm_chart_line_end_t& value) {
    formulon::ChartLineEnd out;
    out.type = static_cast<formulon::ChartLineEndType>(value.type);
    if (value.has_width != 0) {
      out.width = static_cast<formulon::ChartLineEndSize>(value.width);
    }
    if (value.has_length != 0) {
      out.length = static_cast<formulon::ChartLineEndSize>(value.length);
    }
    return out;
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
      field(op, text_or_empty(value), out, name);
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
      fail(name, "must not be null");
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
      rc_ = check_enum_domain(value, min, max, public_name(name).c_str());
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
      rc_ = invalid_argument((public_name(name) + " " + why).c_str());
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
  /// The C field `name` as a binding sees it: without the patch-op suffix.
  static std::string public_name(const std::string& name) {
    std::string out(name);
    constexpr std::string_view kOpSuffix = "_op";
    if (out.size() > kOpSuffix.size() && out.compare(out.size() - kOpSuffix.size(), kOpSuffix.size(), kOpSuffix) == 0) {
      out.resize(out.size() - kOpSuffix.size());
    }
    return out;
  }

  bool op_ok(std::int32_t op, const char* name) {
    if (op >= 0 && op <= static_cast<std::int32_t>(FM_PATCH_OP_SET)) {
      return rc_ == 0;
    }
    if (rc_ == 0) {
      rc_ = invalid_argument((std::string(name) + " is not a patch op").c_str(),
                             std::string(name) + "=" + std::to_string(op));
    }
    return false;
  }

  fm_status_t rc_ = 0;
};

/// A C paint colour as the model's, recording the first failure in `r`: the base must be None, Rgb or a valid Theme
/// colour and each transform carry a name and a 0 / 1 `has_value`. The engine checks the rest.
inline formulon::DrawingColor drawing_color_from_c(PatchReader& r, const fm_drawing_color_t& value,
                                                   const std::string& name) {
  formulon::DrawingColor out;
  if (const char* why = color_spec_problem(value.base); why != nullptr) {
    r.fail(name.c_str(), why);
    return out;
  }
  out.base = color_spec_from_c(value.base);
  if (value.transform_count == 0U) {
    return out;
  }
  if (value.transforms == nullptr) {
    r.check(null_field(name + ".transforms"));
    return out;
  }
  if (value.transform_count > formulon::drawing_color::kMaxTransforms) {
    r.check(set_binding_error(
        formulon::FormulonErrorCode::kSecurityResourceLimit, "transform_count is above the limit",
        formulon::limit_context("color_transforms", value.transform_count, formulon::drawing_color::kMaxTransforms)));
    return out;
  }
  out.transforms.reserve(value.transform_count);
  for (std::uint32_t i = 0; i < value.transform_count; ++i) {
    const fm_color_transform_t& t = value.transforms[i];
    if (t.name == nullptr) {
      r.check(null_field(name + ".transforms.name"));
      return out;
    }
    if (!r.text_ok(t.name, (name + ".transforms.name").c_str()) ||
        !r.enum_ok(t.has_value, 1, (name + ".transforms.has_value").c_str())) {
      return out;
    }
    out.transforms.push_back(
        formulon::ColorTransform{t.name, t.has_value != 0 ? std::optional<std::int32_t>(t.value) : std::nullopt});
  }
  return out;
}

inline void PatchReader::drawing_color_field(std::int32_t op, const fm_drawing_color_t& value,
                                             ChartPatchField<formulon::DrawingColor>& out, const char* name) {
  field(op, op == FM_PATCH_OP_SET ? drawing_color_from_c(*this, value, public_name(name)) : formulon::DrawingColor{},
        out, name);
}

/// The C view of a model paint colour; base None when absent. `sink` owns what the view borrows:
/// `sink.text(s)` returns a stable copy of a transform name, and `sink.transforms(color, view)` a stable array.
template <typename Sink>
fm_drawing_color_t drawing_color_to_c(const std::optional<formulon::DrawingColor>& color, Sink& sink) {
  fm_drawing_color_t out{};
  if (!color) {
    return out;
  }
  out.base = color_spec_to_c(color->base);
  if (!color->transforms.empty()) {
    std::vector<fm_color_transform_t> view;
    view.reserve(color->transforms.size());
    for (const formulon::ColorTransform& t : color->transforms) {
      view.push_back(fm_color_transform_t{sink.text(t.name), t.value ? 1 : 0, t.value.value_or(0)});
    }
    out.transform_count = static_cast<std::uint32_t>(view.size());
    out.transforms = sink.transforms(*color, std::move(view));
  }
  if (color->resolved_argb) {
    out.has_resolved = 1;
    out.resolved = *color->resolved_argb;
  }
  if (color->resolved_alpha) {
    constexpr double kOpaque = 100000.0;
    out.has_resolved_transparency_pct = 1;
    out.resolved_transparency_pct = (kOpaque - static_cast<double>(*color->resolved_alpha)) / 1000.0;
  }
  return out;
}

/// A model line end as the C record reports it.
inline fm_chart_line_end_t line_end_to_c(const formulon::ChartLineEnd& end) {
  fm_chart_line_end_t out{};
  out.type = static_cast<std::int32_t>(end.type);
  out.has_width = end.width ? 1 : 0;
  out.width = end.width ? static_cast<std::int32_t>(*end.width) : 0;
  out.has_length = end.length ? 1 : 0;
  out.length = end.length ? static_cast<std::int32_t>(*end.length) : 0;
  return out;
}

/// The transparency percentage of an alpha in thousandths of a percent; 0 without one.
inline double paint_transparency_of(const std::optional<std::uint32_t>& alpha) {
  return alpha ? 100.0 - static_cast<double>(*alpha) / 1000.0 : 0.0;
}

/// A model fill as the C record reports it; `sink` as `drawing_color_to_c`.
template <typename Sink>
fm_chart_fill_t chart_fill_to_c(const formulon::ChartFill& f, Sink& sink) {
  fm_chart_fill_t out{};
  out.color = drawing_color_to_c(f.color, sink);
  out.has_transparency_pct = f.alpha ? 1 : 0;
  out.transparency_pct = paint_transparency_of(f.alpha);
  out.none = f.none ? 1 : 0;
  out.other = f.other ? 1 : 0;
  return out;
}

/// A model outline as the C record reports it; `sink` as `drawing_color_to_c`.
template <typename Sink>
fm_chart_line_t chart_line_to_c(const formulon::ChartLine& l, Sink& sink) {
  constexpr double kEmuPerPt = 12700.0;
  fm_chart_line_t out{};
  out.color = drawing_color_to_c(l.color, sink);
  out.has_transparency_pct = l.alpha ? 1 : 0;
  out.transparency_pct = paint_transparency_of(l.alpha);
  out.has_width_pt = l.width_emu ? 1 : 0;
  out.width_pt = l.width_emu ? static_cast<double>(*l.width_emu) / kEmuPerPt : 0.0;
  out.has_dash = l.dash ? 1 : 0;
  out.dash = l.dash ? static_cast<std::int32_t>(*l.dash) : 0;
  out.none = l.none ? 1 : 0;
  out.other = l.other ? 1 : 0;
  out.has_head_end = l.head_end ? 1 : 0;
  out.head_end = l.head_end ? line_end_to_c(*l.head_end) : fm_chart_line_end_t{};
  out.has_tail_end = l.tail_end ? 1 : 0;
  out.tail_end = l.tail_end ? line_end_to_c(*l.tail_end) : fm_chart_line_end_t{};
  return out;
}

/// A C fill patch as the model's; shared by the chart and shape edits.
inline formulon::ChartFillPatch read_fill(PatchReader& r, const fm_chart_fill_patch_t& p) {
  formulon::ChartFillPatch out;
  r.drawing_color_field(p.color_op, p.color, out.color, "fill.color_op");
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
  r.drawing_color_field(p.color_op, p.color, out.color, "line.color_op");
  r.field(p.transparency_pct_op, p.transparency_pct, out.transparency_pct, "line.transparency_pct_op");
  r.field(p.width_pt_op, p.width_pt, out.width_pt, "line.width_pt_op");
  r.enum_field(p.dash_op, p.dash, formulon::ChartLineDash::kSysDashDotDot, out.dash, "line.dash");
  r.line_end_field(p.head_end_op, p.head_end, out.head_end, "line.head_end_op");
  r.line_end_field(p.tail_end_op, p.tail_end, out.tail_end, "line.tail_end_op");
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
