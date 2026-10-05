//
// C ABI - workbook theme, colour resolution and effective cell style.

#include "theme.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "color_resolve.h"
#include "sheet.h"
#include "style_resolve.h"
#include "styles.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_finite;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::TextStore;

namespace {

static_assert(sizeof(fm_theme_colors::argb) / sizeof(fm_theme_colors::argb[0]) == formulon::kThemeColorCount,
              "fm_theme_colors must carry every scheme colour");
static_assert(FM_COLOR_CONTEXT_FONT == static_cast<int>(formulon::ColorContext::kFont), "colour context drift");
static_assert(FM_COLOR_CONTEXT_FILL_FG == static_cast<int>(formulon::ColorContext::kFillForeground),
              "colour context drift");
static_assert(FM_COLOR_CONTEXT_FILL_BG == static_cast<int>(formulon::ColorContext::kFillBackground),
              "colour context drift");
static_assert(FM_COLOR_CONTEXT_BORDER == static_cast<int>(formulon::ColorContext::kBorder), "colour context drift");
static_assert(FM_COLOR_RESOLUTION_EXACT == static_cast<int>(formulon::ColorResolution::kExact),
              "colour resolution drift");
static_assert(FM_COLOR_RESOLUTION_DEFAULT_THEME == static_cast<int>(formulon::ColorResolution::kDefaultTheme),
              "colour resolution drift");
static_assert(FM_COLOR_RESOLUTION_INDEX_OUT_OF_RANGE == static_cast<int>(formulon::ColorResolution::kIndexOutOfRange),
              "colour resolution drift");
static_assert(FM_COLOR_RESOLUTION_THEME_UNPARSEABLE == static_cast<int>(formulon::ColorResolution::kThemeUnparseable),
              "colour resolution drift");
static_assert(FM_COLOR_RESOLUTION_AUTO_CONTEXT == static_cast<int>(formulon::ColorResolution::kAutoContext),
              "colour resolution drift");

}  // namespace

extern "C" fm_status_t fm_workbook_get_theme_colors(const fm_workbook_t* wb, fm_theme_colors* out,
                                                    int32_t* out_source) {
  clear_last_error();
  if (wb == nullptr || out == nullptr || out_source == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_get_theme_colors: NULL argument");
  }
  const formulon::LoadedTheme loaded = wb->workbook().load_theme();
  for (std::size_t i = 0; i < formulon::kThemeColorCount; ++i) {
    out->argb[i] = loaded.theme.colors[i];
  }
  *out_source = static_cast<int32_t>(loaded.source);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_theme_colors(fm_workbook_t* wb, const fm_theme_colors* colors) {
  clear_last_error();
  if (wb == nullptr || colors == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_theme_colors: NULL argument");
  }
  formulon::ThemeColors model{};
  for (std::size_t i = 0; i < formulon::kThemeColorCount; ++i) {
    model[i] = colors->argb[i];
  }
  auto r = wb->workbook().set_theme_colors(model);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_get_theme_fonts(const fm_workbook_t* wb, fm_theme_fonts* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_get_theme_fonts: NULL argument");
  }
  formulon::ThemeFonts fonts = wb->workbook().load_theme().theme.fonts;
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  store.emplace_back(std::move(fonts.major_latin));
  out->major_latin = store.back().c_str();
  store.emplace_back(std::move(fonts.major_ea));
  out->major_east_asian = store.back().c_str();
  store.emplace_back(std::move(fonts.minor_latin));
  out->minor_latin = store.back().c_str();
  store.emplace_back(std::move(fonts.minor_ea));
  out->minor_east_asian = store.back().c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_set_theme_fonts(fm_workbook_t* wb, const fm_theme_fonts* fonts) {
  clear_last_error();
  if (wb == nullptr || fonts == nullptr || fonts->major_latin == nullptr || fonts->major_east_asian == nullptr ||
      fonts->minor_latin == nullptr || fonts->minor_east_asian == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_theme_fonts: NULL argument");
  }
  formulon::ThemeFonts model;
  model.major_latin = fonts->major_latin;
  model.major_ea = fonts->major_east_asian;
  model.minor_latin = fonts->minor_latin;
  model.minor_ea = fonts->minor_east_asian;
  auto r = wb->workbook().set_theme_fonts(model);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_resolve_color(const fm_workbook_t* wb, fm_color_spec spec, int32_t context,
                                                 uint32_t* out_argb, int32_t* out_resolution) {
  static constexpr const char* kFn = "fm_workbook_resolve_color";
  clear_last_error();
  if (wb == nullptr || out_argb == nullptr || out_resolution == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_resolve_color: NULL argument");
  }
  if (auto rc = check_enum_domain(context, static_cast<std::int64_t>(formulon::ColorContext::kBorder), kFn, "context");
      rc != 0) {
    return rc;
  }
  if (auto rc =
          check_enum_domain(spec.kind, static_cast<std::int64_t>(formulon::ColorSpec::Kind::kAuto), kFn, "spec.kind");
      rc != 0) {
    return rc;
  }
  formulon::ColorSpec model;
  model.kind = static_cast<formulon::ColorSpec::Kind>(spec.kind);
  model.rgb = spec.rgb;
  model.theme = spec.theme;
  model.indexed = spec.indexed;
  if (model.kind == formulon::ColorSpec::Kind::kTheme) {
    if (auto rc = check_finite(spec.tint, kFn, "spec.tint"); rc != 0) {
      return rc;
    }
    model.tint = spec.tint;
  }
  const formulon::ResolvedColor resolved =
      formulon::resolve_color(wb->workbook(), model, static_cast<formulon::ColorContext>(context));
  *out_argb = resolved.argb;
  *out_resolution = static_cast<int32_t>(resolved.resolution);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_effective_style(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                    uint32_t col, fm_effective_style* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_effective_style: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_effective_style"); rc != 0) {
    return rc;
  }
  if (!formulon::Sheet::coord_in_grid(row, col)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_get_effective_style: cell coordinate out of range",
                             "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  const formulon::Workbook& workbook = wb->workbook();
  formulon::EffectiveStyle style = formulon::effective_style(workbook, workbook.sheet(sheet_index), row, col);
  *out = fm_effective_style{};
  out->xf_index = style.xf_index;
  out->source = static_cast<int32_t>(style.source);
  out->font_index = style.font_index;
  out->fill_index = style.fill_index;
  out->border_index = style.border_index;
  out->font_argb = style.font_color.argb;
  out->font_resolution = static_cast<int32_t>(style.font_color.resolution);
  out->fill_fg_argb = style.fill_foreground.argb;
  out->fill_fg_resolution = static_cast<int32_t>(style.fill_foreground.resolution);
  out->fill_bg_argb = style.fill_background.argb;
  out->fill_bg_resolution = static_cast<int32_t>(style.fill_background.resolution);
  for (std::size_t i = 0; i < style.border_colors.size(); ++i) {
    out->border_argb[i] = style.border_colors[i].argb;
    out->border_resolution[i] = static_cast<int32_t>(style.border_colors[i].resolution);
  }
  out->locked = style.locked ? 1 : 0;
  out->hidden = style.hidden ? 1 : 0;
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  store.emplace_back(std::move(style.num_fmt_code));
  out->num_fmt_code = store.back().c_str();
  return 0;
}
