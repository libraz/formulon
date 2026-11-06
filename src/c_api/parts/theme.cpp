//
// C ABI - workbook theme, colour resolution and effective cell style.

#include "theme.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "color_resolve.h"
#include "sheet.h"
#include "style_resolve.h"
#include "styles.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_finite;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::TextStore;

namespace {

static_assert(sizeof(fm_theme_colors_t::argb) / sizeof(fm_theme_colors_t::argb[0]) == formulon::kThemeColorCount,
              "fm_theme_colors_t must carry every scheme colour");
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

extern "C" fm_status_t fm_workbook_get_theme_colors(const fm_workbook_t* wb, fm_theme_colors_t* out,
                                                    int32_t* out_source) {
  clear_last_error();
  if (wb == nullptr || out == nullptr || out_source == nullptr) {
    return null_argument();
  }
  const formulon::LoadedTheme loaded = wb->workbook().load_theme();
  for (std::size_t i = 0; i < formulon::kThemeColorCount; ++i) {
    out->argb[i] = loaded.theme.colors[i];
  }
  *out_source = static_cast<int32_t>(loaded.source);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_theme_colors(fm_workbook_t* wb, const fm_theme_colors_t* colors) {
  clear_last_error();
  if (wb == nullptr || colors == nullptr) {
    return null_argument();
  }
  formulon::ThemeColors model{};
  for (std::size_t i = 0; i < formulon::kThemeColorCount; ++i) {
    model[i] = colors->argb[i];
  }
  return finish_status(wb->workbook().set_theme_colors(model));
}

extern "C" fm_status_t fm_workbook_get_theme_fonts(const fm_workbook_t* wb, fm_theme_fonts_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  formulon::ThemeFonts fonts = wb->workbook().load_theme().theme.fonts;
  TextStore& store = fresh_read_scratch(wb);
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

extern "C" fm_status_t fm_workbook_set_theme_fonts(fm_workbook_t* wb, const fm_theme_fonts_t* fonts) {
  clear_last_error();
  if (wb == nullptr || fonts == nullptr || fonts->major_latin == nullptr || fonts->major_east_asian == nullptr ||
      fonts->minor_latin == nullptr || fonts->minor_east_asian == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({fonts->major_latin, fonts->major_east_asian, fonts->minor_latin, fonts->minor_east_asian});
      rc != 0) {
    return rc;
  }
  formulon::ThemeFonts model;
  model.major_latin = fonts->major_latin;
  model.major_ea = fonts->major_east_asian;
  model.minor_latin = fonts->minor_latin;
  model.minor_ea = fonts->minor_east_asian;
  return finish_status(wb->workbook().set_theme_fonts(model));
}

extern "C" fm_status_t fm_workbook_get_theme_script_font_count(const fm_workbook_t* wb, int32_t major,
                                                               size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  const formulon::LoadedTheme loaded = wb->workbook().load_theme();
  *out_count = loaded.source != formulon::ThemeSource::kPart
                   ? formulon::kOfficeScriptCount
                   : (major != 0 ? loaded.theme.major_scripts : loaded.theme.minor_scripts).size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_theme_script_font_at(const fm_workbook_t* wb, int32_t major, size_t idx,
                                                            const char** out_script, const char** out_typeface) {
  clear_last_error();
  if (wb == nullptr || out_script == nullptr || out_typeface == nullptr) {
    return null_argument();
  }
  const formulon::LoadedTheme loaded = wb->workbook().load_theme();
  const bool from_part = loaded.source == formulon::ThemeSource::kPart;
  const std::vector<formulon::ThemeScriptFont>& scripts =
      major != 0 ? loaded.theme.major_scripts : loaded.theme.minor_scripts;
  if (auto rc = check_index(idx, from_part ? scripts.size() : formulon::kOfficeScriptCount, "index"); rc != 0) {
    return rc;
  }
  TextStore& store = fresh_read_scratch(wb);
  store.emplace_back(from_part ? scripts[idx].script : formulon::office_script(idx));
  *out_script = store.back().c_str();
  store.emplace_back(from_part ? scripts[idx].typeface : formulon::office_script_face(idx, major != 0));
  *out_typeface = store.back().c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_theme_script_font(const fm_workbook_t* wb, int32_t major, const char* script,
                                                         int32_t* out_present, const char** out_typeface) {
  clear_last_error();
  if (wb == nullptr || script == nullptr || out_present == nullptr || out_typeface == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({script}); rc != 0) {
    return rc;
  }
  *out_present = 0;
  *out_typeface = nullptr;
  const formulon::LoadedTheme loaded = wb->workbook().load_theme();
  const bool from_part = loaded.source == formulon::ThemeSource::kPart;
  const std::vector<formulon::ThemeScriptFont>& scripts =
      major != 0 ? loaded.theme.major_scripts : loaded.theme.minor_scripts;
  const std::size_t count = from_part ? scripts.size() : formulon::kOfficeScriptCount;
  for (std::size_t i = 0; i < count; ++i) {
    if ((from_part ? std::string_view(scripts[i].script) : formulon::office_script(i)) == script) {
      TextStore& store = fresh_read_scratch(wb);
      store.emplace_back(from_part ? scripts[i].typeface : formulon::office_script_face(i, major != 0));
      *out_typeface = store.back().c_str();
      *out_present = 1;
      return 0;
    }
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_theme_script_font(fm_workbook_t* wb, int32_t major, const char* script,
                                                         const char* typeface) {
  clear_last_error();
  if (wb == nullptr || script == nullptr || typeface == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({script, typeface}); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().set_theme_script_font(major != 0, script, typeface));
}

extern "C" fm_status_t fm_workbook_clear_theme(fm_workbook_t* wb) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  return finish_status(wb->workbook().reset_theme());
}

extern "C" fm_status_t fm_workbook_resolve_color(const fm_workbook_t* wb, const fm_color_spec_t* spec, int32_t context,
                                                 uint32_t* out_argb, int32_t* out_resolution) {
  clear_last_error();
  if (wb == nullptr || spec == nullptr || out_argb == nullptr || out_resolution == nullptr) {
    return null_argument();
  }
  if (auto rc = check_enum_domain(context, static_cast<std::int64_t>(formulon::ColorContext::kBorder), "context");
      rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec->kind, static_cast<std::int64_t>(formulon::ColorSpec::Kind::kAuto), "spec.kind");
      rc != 0) {
    return rc;
  }
  formulon::ColorSpec model;
  model.kind = static_cast<formulon::ColorSpec::Kind>(spec->kind);
  model.rgb = spec->rgb;
  model.theme = spec->theme;
  model.indexed = spec->indexed;
  if (model.kind == formulon::ColorSpec::Kind::kTheme) {
    if (auto rc = check_finite(spec->tint, "spec.tint"); rc != 0) {
      return rc;
    }
    model.tint = spec->tint;
  }
  const formulon::ResolvedColor resolved =
      formulon::resolve_color(wb->workbook(), model, static_cast<formulon::ColorContext>(context));
  *out_argb = resolved.argb;
  *out_resolution = static_cast<int32_t>(resolved.resolution);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_effective_style(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                    uint32_t col, fm_effective_style_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  const formulon::Workbook& workbook = wb->workbook();
  formulon::EffectiveStyle style = formulon::effective_style(workbook, workbook.sheet(sheet_index), row, col);
  *out = fm_effective_style_t{};
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
  TextStore& store = fresh_read_scratch(wb);
  store.emplace_back(std::move(style.num_fmt_code));
  out->num_fmt_code = store.back().c_str();
  return 0;
}
