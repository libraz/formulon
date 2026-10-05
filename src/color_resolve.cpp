#include "color_resolve.h"

#include <algorithm>
#include <array>
#include <cmath>

#include "workbook.h"

namespace formulon {
namespace {

constexpr int kHlsMax = 240;
constexpr int kRgbMax = 255;

constexpr std::uint32_t kBlack = 0xFF000000U;
constexpr std::uint32_t kWhite = 0xFFFFFFFFU;

constexpr std::array<std::uint32_t, 64> kDefaultPalette = {
    0xFF000000U, 0xFFFFFFFFU, 0xFFFF0000U, 0xFF00FF00U, 0xFF0000FFU, 0xFFFFFF00U, 0xFFFF00FFU, 0xFF00FFFFU,
    0xFF000000U, 0xFFFFFFFFU, 0xFFFF0000U, 0xFF00FF00U, 0xFF0000FFU, 0xFFFFFF00U, 0xFFFF00FFU, 0xFF00FFFFU,
    0xFF800000U, 0xFF008000U, 0xFF000080U, 0xFF808000U, 0xFF800080U, 0xFF008080U, 0xFFC0C0C0U, 0xFF808080U,
    0xFF9999FFU, 0xFF993366U, 0xFFFFFFCCU, 0xFFCCFFFFU, 0xFF660066U, 0xFFFF8080U, 0xFF0066CCU, 0xFFCCCCFFU,
    0xFF000080U, 0xFFFF00FFU, 0xFFFFFF00U, 0xFF00FFFFU, 0xFF800080U, 0xFF800000U, 0xFF008080U, 0xFF0000FFU,
    0xFF00CCFFU, 0xFFCCFFFFU, 0xFFCCFFCCU, 0xFFFFFF99U, 0xFF99CCFFU, 0xFFFF99CCU, 0xFFCC99FFU, 0xFFFFCC99U,
    0xFF3366FFU, 0xFF33CCCCU, 0xFF99CC00U, 0xFFFFCC00U, 0xFFFF9900U, 0xFFFF6600U, 0xFF666699U, 0xFF969696U,
    0xFF003366U, 0xFF339966U, 0xFF003300U, 0xFF333300U, 0xFF993300U, 0xFF993366U, 0xFF333399U, 0xFF333333U};

/// `spec.theme` order (0=lt1, 1=dk1, 2=lt2, 3=dk2) to scheme order.
constexpr std::array<std::size_t, kThemeColorCount> kThemeToScheme = {1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11};

int hue_to_rgb(int n1, int n2, int hue) {
  hue %= kHlsMax;
  if (hue < 0) {
    hue += kHlsMax;
  }
  if (hue < kHlsMax / 6) {
    return n1 + (((n2 - n1) * hue + (kHlsMax / 12)) / (kHlsMax / 6));
  }
  if (hue < kHlsMax / 2) {
    return n2;
  }
  if (hue < (kHlsMax * 2) / 3) {
    return n1 + (((n2 - n1) * (((kHlsMax * 2) / 3) - hue) + (kHlsMax / 12)) / (kHlsMax / 6));
  }
  return n1;
}

std::uint32_t rgb_pack(int r, int g, int b) {
  return 0xFF000000U | (static_cast<std::uint32_t>(r) << 16U) | (static_cast<std::uint32_t>(g) << 8U) |
         static_cast<std::uint32_t>(b);
}

ResolvedColor auto_color(ColorContext context) {
  const bool fill = context == ColorContext::kFillForeground || context == ColorContext::kFillBackground;
  return {fill ? kWhite : kBlack, ColorResolution::kAutoContext};
}

}  // namespace

std::uint32_t default_indexed_color(std::uint32_t index) {
  return index < kDefaultPalette.size() ? kDefaultPalette[index] : kBlack;
}

std::uint32_t apply_tint(std::uint32_t argb, double tint) {
  if (tint == 0.0) {
    return argb | 0xFF000000U;
  }
  const int r = static_cast<int>((argb >> 16U) & 0xFFU);
  const int g = static_cast<int>((argb >> 8U) & 0xFFU);
  const int b = static_cast<int>(argb & 0xFFU);
  const int c_max = std::max({r, g, b});
  const int c_min = std::min({r, g, b});
  const int lum = (((c_max + c_min) * kHlsMax) + kRgbMax) / (2 * kRgbMax);
  int sat = 0;
  int hue = 0;
  if (c_max != c_min) {
    if (lum <= kHlsMax / 2) {
      sat = (((c_max - c_min) * kHlsMax) + ((c_max + c_min) / 2)) / (c_max + c_min);
    } else {
      sat = (((c_max - c_min) * kHlsMax) + ((2 * kRgbMax - c_max - c_min) / 2)) / (2 * kRgbMax - c_max - c_min);
    }
    const int span = c_max - c_min;
    const int r_delta = (((c_max - r) * (kHlsMax / 6)) + (span / 2)) / span;
    const int g_delta = (((c_max - g) * (kHlsMax / 6)) + (span / 2)) / span;
    const int b_delta = (((c_max - b) * (kHlsMax / 6)) + (span / 2)) / span;
    if (r == c_max) {
      hue = b_delta - g_delta;
    } else if (g == c_max) {
      hue = (kHlsMax / 3) + r_delta - b_delta;
    } else {
      hue = ((2 * kHlsMax) / 3) + g_delta - r_delta;
    }
    if (hue < 0) {
      hue += kHlsMax;
    }
    if (hue > kHlsMax) {
      hue -= kHlsMax;
    }
  }
  const double lum_d = static_cast<double>(lum);
  // Each product is its own statement: a fused multiply-add would shift the
  // value just below an integer and floor it one step low.
  const double keep = tint < 0.0 ? 1.0 + tint : 1.0 - tint;
  const double kept = lum_d * keep;
  const double boost_base = kHlsMax * keep;
  const double boost = kHlsMax - boost_base;
  const double scaled = tint < 0.0 ? kept : kept + boost;
  const int new_lum = std::clamp(static_cast<int>(std::floor(scaled)), 0, kHlsMax);
  if (sat == 0) {
    const int v = (new_lum * kRgbMax + kHlsMax / 2) / kHlsMax;
    return rgb_pack(v, v, v);
  }
  int magic2 = 0;
  if (new_lum <= kHlsMax / 2) {
    magic2 = (new_lum * (kHlsMax + sat) + (kHlsMax / 2)) / kHlsMax;
  } else {
    magic2 = new_lum + sat - ((new_lum * sat) + (kHlsMax / 2)) / kHlsMax;
  }
  const int magic1 = 2 * new_lum - magic2;
  const auto channel = [&](int h) { return (hue_to_rgb(magic1, magic2, h) * kRgbMax + kHlsMax / 2) / kHlsMax; };
  return rgb_pack(channel(hue + kHlsMax / 3), channel(hue), channel(hue - kHlsMax / 3));
}

ResolvedColor resolve_color(const ColorSpec& spec, const Theme& theme, ThemeSource source,
                            const IndexedPalette& palette, ColorContext context) {
  switch (spec.kind) {
    case ColorSpec::Kind::kRgb:
      return {spec.rgb, ColorResolution::kExact};
    case ColorSpec::Kind::kTheme: {
      if (spec.theme >= kThemeColorCount) {
        return {kBlack, ColorResolution::kIndexOutOfRange};
      }
      ColorResolution resolution = ColorResolution::kExact;
      if (source == ThemeSource::kDefault) {
        resolution = ColorResolution::kDefaultTheme;
      } else if (source == ThemeSource::kUnparseable) {
        resolution = ColorResolution::kThemeUnparseable;
      }
      return {apply_tint(theme.colors[kThemeToScheme[spec.theme]], spec.tint), resolution};
    }
    case ColorSpec::Kind::kIndexed:
      if (spec.indexed == 64U || spec.indexed == 65U) {
        // System foreground / background: windowText in text and lines, window in fills.
        const bool fill = context == ColorContext::kFillForeground || context == ColorContext::kFillBackground;
        const std::uint32_t argb = (spec.indexed == 65U || fill) ? kWhite : kBlack;
        return {argb, ColorResolution::kAutoContext};
      }
      if (spec.indexed < palette.size()) {
        return {palette[spec.indexed] | 0xFF000000U, ColorResolution::kExact};
      }
      if (spec.indexed < kDefaultPalette.size()) {
        return {kDefaultPalette[spec.indexed], ColorResolution::kExact};
      }
      return {kBlack, ColorResolution::kIndexOutOfRange};
    case ColorSpec::Kind::kAuto:
    case ColorSpec::Kind::kNone:
    default:
      return auto_color(context);
  }
}

ResolvedColor resolve_color(const Workbook& wb, const ColorSpec& spec, ColorContext context) {
  if (spec.kind != ColorSpec::Kind::kTheme) {
    return resolve_color(spec, default_theme(), ThemeSource::kDefault, wb.styles().indexed_colors, context);
  }
  const LoadedTheme loaded = load_theme(wb);
  return resolve_color(spec, loaded.theme, loaded.source, wb.styles().indexed_colors, context);
}

}  // namespace formulon
