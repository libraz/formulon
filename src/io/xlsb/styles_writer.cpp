
#include "io/xlsb/styles_writer.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <tuple>
#include <utility>
#include <vector>

#include "cell_style_save_plan.h"
#include "io/dxf_save_plan.h"
#include "io/xlsb/brt_color.h"
#include "io/xlsb/dxf_props.h"
#include "io/xlsb/dxf_record_slots.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/xf_flags.h"
#include "utils/index_sort.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// The structural record ids live in the XLSB record-name table.  Keeping them
// here, rather than pretending every style-only record is understood by the
// sheet reader, avoids widening its public record enum unnecessarily.
constexpr std::uint16_t kBrtBeginStyleSheet = 278;
constexpr std::uint16_t kBrtEndStyleSheet = 279;
constexpr std::uint16_t kBrtBeginFills = 603;
constexpr std::uint16_t kBrtEndFills = 604;
constexpr std::uint16_t kBrtBeginFonts = 611;
constexpr std::uint16_t kBrtEndFonts = 612;
constexpr std::uint16_t kBrtBeginBorders = 613;
constexpr std::uint16_t kBrtEndBorders = 614;
constexpr std::uint16_t kBrtBeginFmts = 615;
constexpr std::uint16_t kBrtEndFmts = 616;
constexpr std::uint16_t kBrtBeginStyles = 619;
constexpr std::uint16_t kBrtEndStyles = 620;
constexpr std::uint16_t kBrtStyle = 48;

std::uint16_t ClampIndex(std::uint32_t value) {
  return static_cast<std::uint16_t>(std::min<std::uint32_t>(value, std::numeric_limits<std::uint16_t>::max()));
}

std::uint16_t FontHeightTwips(double points) {
  const double twips = std::round(points * 20.0);
  if (!std::isfinite(twips)) {
    return 220U;
  }
  return static_cast<std::uint16_t>(std::clamp(twips, 20.0, 8191.0));
}

void EmitColor(std::vector<std::uint8_t>& payload, std::uint32_t argb, const ColorSpec& spec) {
  const auto bytes = encode_brt_color(argb, spec);
  payload.insert(payload.end(), bytes.begin(), bytes.end());
}

void EmitFont(std::vector<std::uint8_t>& out, const FontRecord& font) {
  std::vector<std::uint8_t> payload;
  emit_u16(payload, FontHeightTwips(font.size));
  // Excel mirrors a bold weight in grbit bit 0.
  std::uint16_t flags = font.bold ? 0x0001U : 0U;
  if (font.italic)
    flags |= 0x0002U;
  if (font.strike)
    flags |= 0x0008U;
  if (font.outline)
    flags |= 0x0010U;
  if (font.shadow)
    flags |= 0x0020U;
  if (font.condense)
    flags |= 0x0040U;
  if (font.extend)
    flags |= 0x0080U;
  emit_u16(payload, flags);
  emit_u16(payload, font.bold ? 700U : 400U);
  emit_u16(payload, font.vert_align == 1U ? 1U : (font.vert_align == 2U ? 2U : 0U));
  emit_u8(payload, kDxfUnderlineWire[std::min<std::uint8_t>(font.underline, 4U)]);
  emit_u8(payload, font.has_family ? font.family : 0U);
  emit_u8(payload, font.has_charset ? font.charset : 0U);
  emit_u8(payload, 0U);
  if (font.color.kind == ColorSpec::Kind::kNone && font.color_argb == 0xFF000000U) {
    // A font without `<color>` is saved as automatic with a zero RGB.
    ColorSpec automatic;
    automatic.kind = ColorSpec::Kind::kAuto;
    EmitColor(payload, 0U, automatic);
  } else {
    EmitColor(payload, font.color_argb, font.color);
  }
  emit_u8(payload, font.scheme);  // bFontScheme shares the shared model's ordinals
  emit_xlwidestring(payload, font.name.empty() ? std::string_view("Calibri") : std::string_view(font.name));
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtFont), payload);
}

/// Gradient type, then degree, left, right, top, bottom: the order BrtFill
/// and XFPropGradient share.
void EmitGradientGeometry(std::vector<std::uint8_t>& payload, const GradientFill& gradient) {
  emit_u32(payload, gradient.path ? 1U : 0U);
  emit_double(payload, gradient.path ? 0.0 : gradient.degree);
  for (double edge : {gradient.left, gradient.right, gradient.top, gradient.bottom}) {
    emit_double(payload, gradient.path ? edge : 0.0);
  }
}

void EmitFill(std::vector<std::uint8_t>& out, const FillRecord& fill) {
  std::vector<std::uint8_t> payload;
  if (fill.has_gradient) {
    // fls 0x28: the pattern colours are ignored, the gradient tail is live.
    emit_u32(payload, kBrtFillGradient);
    EmitColor(payload, fill.fg_argb, fill.fg);
    EmitColor(payload, fill.bg_argb, fill.bg);
    EmitGradientGeometry(payload, fill.gradient);
    const std::size_t stops = fill.gradient.stops.size();
    emit_u32(payload, static_cast<std::uint32_t>(stops));
    for (std::size_t i = 0; i < stops; ++i) {
      // GradientStop (2.5.68): BrtColor, then Xnum position.
      const GradientStop& stop = fill.gradient.stops[i];
      EmitColor(payload, stop.argb, stop.color);
      emit_double(payload, stop.position);
    }
    emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtFill), payload);
    return;
  }
  emit_u32(payload, fill.pattern <= 18U ? fill.pattern : 0U);
  EmitColor(payload, fill.fg_argb, fill.fg);
  EmitColor(payload, fill.bg_argb, fill.bg);
  emit_u32(payload, 0U);  // non-gradient fill
  for (int i = 0; i < 5; ++i) {
    emit_double(payload, 0.0);
  }
  emit_u32(payload, 0U);  // no gradient stops
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtFill), payload);
}

void EmitBorderSide(std::vector<std::uint8_t>& payload, const BorderSide& side) {
  emit_u8(payload, std::min<std::uint8_t>(side.style, 13U));
  emit_u8(payload, 0U);
  EmitColor(payload, side.color_argb, side.color);
}

void EmitBorder(std::vector<std::uint8_t>& out, const BorderRecord& border) {
  std::vector<std::uint8_t> payload;
  emit_u8(payload, static_cast<std::uint8_t>((border.diagonal_down ? 1U : 0U) | (border.diagonal_up ? 2U : 0U)));
  EmitBorderSide(payload, border.top);
  EmitBorderSide(payload, border.bottom);
  EmitBorderSide(payload, border.left);
  EmitBorderSide(payload, border.right);
  EmitBorderSide(payload, border.diagonal);
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtBorder), payload);
}

// `parent` is the written ixfeParent, absent for a style xf.
void EmitXf(std::vector<std::uint8_t>& out, const CellXf& xf, std::optional<std::uint32_t> parent) {
  std::vector<std::uint8_t> payload;
  emit_u16(payload, parent.has_value() ? ClampIndex(*parent) : kXfNoParent);
  emit_u16(payload, xf.num_fmt_id);
  emit_u16(payload, ClampIndex(xf.font_index));
  emit_u16(payload, ClampIndex(xf.fill_index));
  emit_u16(payload, ClampIndex(xf.border_index));
  // `trot` and `indent` are u8 fields; the OOXML reader already rejects an
  // out-of-range `textRotation` / `indent`, so the clamp only guards an xf
  // built in memory.
  emit_u8(payload, static_cast<std::uint8_t>(std::min<std::uint32_t>(xf.text_rotation, 255U)));
  emit_u8(payload, static_cast<std::uint8_t>(std::min<std::uint32_t>(xf.indent, 255U)));
  // Every field of the flags word is unconditional on the wire, so the
  // effective value is emitted whether or not the source `<xf>` spelled it
  // out: `CellXf`'s `has_*` bits record how OOXML wrote a value, not
  // whether the value applies.
  std::uint16_t flags = static_cast<std::uint16_t>(
      (static_cast<std::uint16_t>(xf.horizontal_align) & kXfHorizontalAlignMask) |
      ((static_cast<std::uint16_t>(xf.vertical_align) << kXfVerticalAlignShift) & kXfVerticalAlignMask) |
      ((std::min<std::uint32_t>(xf.reading_order, 3U) << kXfReadingOrderShift) & kXfReadingOrderMask));
  if (xf.wrap_text)
    flags |= kXfWrapText;
  if (xf.justify_last_line)
    flags |= kXfJustifyLastLine;
  if (xf.shrink_to_fit)
    flags |= kXfShrinkToFit;
  // An absent `<protection>` means the schema defaults apply, so a locked,
  // non-hidden cell is what the wire has to say.
  if (!xf.has_protection || xf.locked)
    flags |= kXfLocked;
  if (xf.has_protection && xf.hidden)
    flags |= kXfHidden;
  if (xf.quote_prefix)
    flags |= kXfQuotePrefix;
  emit_u16(payload, flags);
  std::uint16_t apply = 0U;
  if (xf.apply_number_format)
    apply |= kXfApplyNumberFormat;
  if (xf.apply_font)
    apply |= kXfApplyFont;
  if (xf.apply_alignment)
    apply |= kXfApplyAlignment;
  if (xf.apply_border)
    apply |= kXfApplyBorder;
  if (xf.apply_fill)
    apply |= kXfApplyFill;
  if (xf.apply_protection)
    apply |= kXfApplyProtection;
  emit_u16(payload, apply);
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtXF), payload);
}

// Returns `entries`, or a one-element default-initialised collection staged in
// `storage` when `entries` is empty; `storage` must outlive the result. The
// default is built in place rather than taken as a parameter, because binding
// a caller's temporary to a reference parameter of a reference-returning
// function is what -Wdangling-reference reports, whether or not that reference
// can actually escape.
/// One XFProp of a BrtDXF: u16 type, u16 total size, then `data`.
void AppendProp(std::vector<std::uint8_t>& props, std::uint16_t& count, std::uint16_t type,
                const std::vector<std::uint8_t>& data) {
  emit_u16(props, type);
  emit_u16(props, static_cast<std::uint16_t>(4U + data.size()));
  props.insert(props.end(), data.begin(), data.end());
  ++count;
}

/// A wide string with a u16 length rather than XLWideString's u32 one.
std::vector<std::uint8_t> ShortWideString(std::string_view text) {
  std::vector<std::uint8_t> out;
  emit_xlwidestring(out, text);
  // Drop the high half of the u32 prefix.
  out.erase(out.begin() + 2, out.begin() + 4);
  return out;
}

std::vector<std::uint8_t> ColorProp(std::uint32_t argb, const ColorSpec& spec) {
  std::vector<std::uint8_t> data;
  EmitColor(data, argb, spec);
  return data;
}

/// A dxf as Excel writes it: fill, font, number format, border, each only
/// where the dxf states it. Alignment and protection are left out: Excel
/// discards them from a CF dxf, so they have no property.
void EmitDxf(std::vector<std::uint8_t>& out, const DifferentialFormat& dxf) {
  std::vector<std::uint8_t> props;
  std::uint16_t count = 0;
  if (dxf.has_fill && dxf.fill.has_gradient) {
    // XFProps forbids a pattern (type 0) beside a gradient (types 3 / 4).
    const GradientFill& gradient = dxf.fill.gradient;
    std::vector<std::uint8_t> geometry;
    EmitGradientGeometry(geometry, gradient);
    AppendProp(props, count, kDxfPropFillGradient, geometry);
    for (std::size_t i = 0; i < gradient.stops.size(); ++i) {
      const GradientStop& stop = gradient.stops[i];
      std::vector<std::uint8_t> data;
      emit_u16(data, 0U);
      emit_double(data, stop.position);
      EmitColor(data, stop.argb, stop.color);
      AppendProp(props, count, kDxfPropGradientStop, data);
    }
  } else if (dxf.has_fill) {
    const FillRecord& fill = dxf.fill;
    AppendProp(props, count, kDxfPropFillPattern, {static_cast<std::uint8_t>(fill.pattern <= 18U ? fill.pattern : 0U)});
    if (fill.fg.kind != ColorSpec::Kind::kNone) {
      AppendProp(props, count, kDxfPropFillFg, ColorProp(fill.fg_argb, fill.fg));
    }
    if (fill.bg.kind != ColorSpec::Kind::kNone) {
      AppendProp(props, count, kDxfPropFillBg, ColorProp(fill.bg_argb, fill.bg));
    }
  }
  if (dxf.has_font) {
    const FontRecord& font = dxf.font;
    // 0xFF000000 is the "automatic" sentinel an unstated colour carries.
    if (font.color.kind != ColorSpec::Kind::kNone || font.color_argb != 0xFF000000U) {
      AppendProp(props, count, kDxfPropFontColor, ColorProp(font.color_argb, font.color));
    }
    if (!font.name.empty()) {
      AppendProp(props, count, kDxfPropFontName, ShortWideString(font.name));
    }
    // A stated toggle is written even when off: a dxf `val="0"` clears it.
    if (font.has_bold) {
      std::vector<std::uint8_t> weight;
      emit_u16(weight, font.bold ? 700U : 400U);
      AppendProp(props, count, kDxfPropFontWeight, weight);
    }
    if (font.underline != 0U) {
      AppendProp(props, count, kDxfPropUnderline,
                 {kDxfUnderlineWire[std::min<std::uint8_t>(font.underline, 4U)], 0x00});
    }
    if (font.vert_align == 1U || font.vert_align == 2U) {
      AppendProp(props, count, kDxfPropVertAlign, {font.vert_align, 0x00});
    }
    const std::tuple<std::uint16_t, bool, bool> toggles[] = {
        {kDxfPropItalic, font.has_italic, font.italic},       {kDxfPropStrike, font.has_strike, font.strike},
        {kDxfPropOutline, font.has_outline, font.outline},    {kDxfPropShadow, font.has_shadow, font.shadow},
        {kDxfPropCondense, font.has_condense, font.condense}, {kDxfPropExtend, font.has_extend, font.extend}};
    for (const auto& [type, stated, on] : toggles) {
      if (stated) {
        AppendProp(props, count, type, {static_cast<std::uint8_t>(on ? 1U : 0U)});
      }
    }
    if (font.has_size) {
      std::vector<std::uint8_t> twips;
      emit_u32(twips, FontHeightTwips(font.size));
      AppendProp(props, count, kDxfPropFontSize, twips);
    }
  }
  if (dxf.has_num_fmt) {
    std::vector<std::uint8_t> id;
    emit_u16(id, dxf.num_fmt_id);
    AppendProp(props, count, kDxfPropNumFmtId, id);
    AppendProp(props, count, kDxfPropNumFmtCode, ShortWideString(dxf.num_fmt_code));
  }
  if (dxf.has_border) {
    const BorderRecord& border = dxf.border;
    const std::pair<std::uint16_t, const BorderSide*> sides[] = {{kDxfPropBorderTop, &border.top},
                                                                 {kDxfPropBorderBottom, &border.bottom},
                                                                 {kDxfPropBorderLeft, &border.left},
                                                                 {kDxfPropBorderRight, &border.right},
                                                                 {kDxfPropBorderDiagonal, &border.diagonal}};
    for (const auto& [type, side] : sides) {
      if (side->style == 0U) {
        continue;
      }
      ColorSpec auto_color;
      auto_color.kind = ColorSpec::Kind::kAuto;
      const bool stated = side->color.kind != ColorSpec::Kind::kNone;
      std::vector<std::uint8_t> data = ColorProp(stated ? side->color_argb : 0U, stated ? side->color : auto_color);
      emit_u16(data, std::min<std::uint8_t>(side->style, 13U));
      AppendProp(props, count, type, data);
    }
    if (border.diagonal_down) {
      AppendProp(props, count, kDxfPropDiagonalDown, {0x01});
    }
    if (border.diagonal_up) {
      AppendProp(props, count, kDxfPropDiagonalUp, {0x01});
    }
  }
  std::vector<std::uint8_t> payload;
  emit_u16(payload, 0x8000U);
  emit_u16(payload, 0U);
  emit_u16(payload, count);
  payload.insert(payload.end(), props.begin(), props.end());
  emit_record(out, kBrtDxf, payload);
}

template <typename T>
const std::vector<T>& Normalized(const std::vector<T>& entries, std::vector<T>& storage) {
  if (!entries.empty())
    return entries;
  storage.assign(1U, T{});
  return storage;
}

void EmitCount(std::vector<std::uint8_t>& out, std::uint16_t type, std::size_t count) {
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(count));
  emit_record(out, type, payload);
}

void EmitStyle(std::vector<std::uint8_t>& out, const CellStyleRecord& style) {
  std::vector<std::uint8_t> payload;
  emit_u32(payload, style.xf_id);
  std::uint16_t flags = static_cast<std::uint16_t>((style.builtin_id != CellStyleRecord::kBuiltinIdNone ? 1U : 0U) |
                                                   (style.hidden ? 2U : 0U) | (style.custom_builtin ? 4U : 0U));
  emit_u16(payload, flags);
  emit_u8(payload,
          style.builtin_id == CellStyleRecord::kBuiltinIdNone ? 0U : static_cast<std::uint8_t>(style.builtin_id));
  emit_u8(payload, static_cast<std::uint8_t>(style.i_level));
  emit_xlwidestring(payload, style.name);
  emit_record(out, kBrtStyle, payload);
}

/// BrtACBegin payload of a custom table style's revision-uid wrapper.
constexpr std::array<std::uint8_t, 6> kTableStyleUidBeginPayload = {0x01, 0x00, 0x01, 0x11, 0x00, 0x80};

/// The 16-byte uid of a custom table style: its `xr9:uid`, else the one Excel assigns, u32s 0, 0x11, 0xFFFFFFFF and
/// the style's position among those written.
std::array<std::uint8_t, kGuidBytes> TableStyleUid(const CustomTableStyle& style, std::uint32_t position) {
  std::array<std::uint8_t, kGuidBytes> guid{};
  constexpr std::string_view kKey = " xr9:uid=\"";
  const std::size_t at = style.extra_attrs.find(kKey);
  const std::size_t start = at == std::string::npos ? at : at + kKey.size();
  const std::size_t end = start == std::string::npos ? start : style.extra_attrs.find('"', start);
  if (end != std::string::npos && ParseGuid(std::string_view(style.extra_attrs).substr(start, end - start), guid)) {
    return guid;
  }
  guid = {};
  guid[4] = 0x11U;
  for (std::size_t b = 0; b < 4U; ++b) {
    guid[8U + b] = 0xFFU;
    guid[12U + b] = static_cast<std::uint8_t>((position >> (8U * b)) & 0xFFU);
  }
  return guid;
}

/// The custom table styles block, which Excel always writes: the defaults in effect, then each writable style behind
/// its revision-uid wrapper with its elements in type order. An element's dxf is its saved index for the `nth`
/// reference to that model dxf in the styles part, counted in model element order as the uses were registered.
void EmitTableStyles(std::vector<std::uint8_t>& out, const StylesTable& table, const DxfSavePlan* plan) {
  const TableStylesSection& section = table.table_styles;
  std::vector<const CustomTableStyle*> written;
  for (const CustomTableStyle& style : section.styles) {
    if (table_style_is_written(style)) {
      written.push_back(&style);
    }
  }
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(written.size()));
  emit_xlwidestring(payload, section.effective_table_default());
  emit_xlwidestring(payload, section.effective_pivot_default());
  emit_record(out, kBrtBeginTableStyles, payload);
  std::vector<std::uint32_t> nth(table.dxfs.size(), 0U);
  const DxfPartKey styles_part{-1, DxfPartType::kStyles, 0U};
  for (std::size_t i = 0; i < written.size(); ++i) {
    const CustomTableStyle& style = *written[i];
    const auto guid = TableStyleUid(style, static_cast<std::uint32_t>(i));
    emit_record(out, kAcBegin, ByteSpan{kTableStyleUidBeginPayload.data(), kTableStyleUidBeginPayload.size()});
    emit_record(out, kUid, ByteSpan{guid.data(), guid.size()});
    emit_record(out, kAcEnd, ByteSpan{});
    payload.clear();
    emit_u16(payload, static_cast<std::uint16_t>((style.pivot ? kTableStyleIsPivot : 0U) |
                                                 (style.table ? kTableStyleIsTable : 0U)));
    emit_u32(payload, static_cast<std::uint32_t>(style.elements.size()));
    emit_xlwidestring(payload, style.name);
    emit_record(out, kBrtBeginTableStyle, payload);
    // (type, size, saved dxf) per element, in model order, then written in type order.
    std::vector<std::array<std::uint32_t, 3>> elements;
    for (const TableStyleElement& element : style.elements) {
      const std::uint32_t model = *element.dxf_id;
      const std::uint32_t saved =
          plan != nullptr && model < nth.size() ? plan->saved(styles_part, model, nth[model]++).value_or(model) : model;
      elements.push_back({table_style_element_type(element.type), element.size, saved});
    }
    sort_by_index(elements, [](const auto& a, const auto& b) { return a[0] < b[0]; });
    for (const auto& element : elements) {
      payload.clear();
      for (const std::uint32_t field : element) {
        emit_u32(payload, field);
      }
      emit_record(out, kBrtTableStyleElement, payload);
    }
    emit_record(out, kBrtEndTableStyle, ByteSpan{});
  }
  emit_record(out, kBrtEndTableStyles, ByteSpan{});
}

}  // namespace

void emit_font_record(std::vector<std::uint8_t>& out, const FontRecord& font) {
  EmitFont(out, font);
}

std::uint32_t table_style_element_type(std::string_view type) {
  const auto found = std::find(kTableStyleElementTypes.begin(), kTableStyleElementTypes.end(), type);
  return static_cast<std::uint32_t>(found - kTableStyleElementTypes.begin());
}

bool table_style_is_written(const CustomTableStyle& style) {
  if (style.pivot == style.table) {
    return false;
  }
  return std::all_of(style.elements.begin(), style.elements.end(), [](const TableStyleElement& element) {
    return element.dxf_id.has_value() && table_style_element_type(element.type) < kTableStyleElementTypes.size();
  });
}

std::vector<std::uint8_t> write_styles_bin(const StylesTable& table, const DxfSavePlan* plan,
                                           const CellStyleSavePlan* style_plan, const FontSavePlan* font_plan,
                                           const std::vector<FontRecord>* appended_fonts) {
  std::vector<std::uint8_t> out;
  out.reserve(256U + table.fonts.size() * 40U + table.fills.size() * 64U + table.borders.size() * 56U +
              table.cell_xfs.size() * 20U);
  const std::vector<std::uint8_t> empty;
  emit_record(out, kBrtBeginStyleSheet, empty);

  std::vector<NumFmtRecord> custom_formats;
  custom_formats.reserve(table.num_fmts.size());
  for (const NumFmtRecord& fmt : table.num_fmts) {
    if (fmt.id >= 164U && fmt.format_string_index < table.num_fmt_strings.size())
      custom_formats.push_back(fmt);
  }
  EmitCount(out, kBrtBeginFmts, custom_formats.size());
  for (const NumFmtRecord& fmt : custom_formats) {
    std::vector<std::uint8_t> payload;
    emit_u16(payload, fmt.id);
    emit_xlwidestring(payload, table.num_fmt_strings[fmt.format_string_index]);
    emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtFmt), payload);
  }
  emit_record(out, kBrtEndFmts, empty);

  std::vector<FontRecord> font_default;
  const std::vector<FontRecord>& fonts =
      Normalized(font_plan != nullptr ? font_plan->fonts() : table.fonts, font_default);
  const std::size_t appended = appended_fonts != nullptr ? appended_fonts->size() : 0U;
  EmitCount(out, kBrtBeginFonts, fonts.size() + appended);
  for (const FontRecord& font : fonts)
    EmitFont(out, font);
  for (std::size_t i = 0; i < appended; ++i)
    EmitFont(out, (*appended_fonts)[i]);
  emit_record(out, kBrtEndFonts, empty);

  std::vector<FillRecord> fill_default;
  const std::vector<FillRecord>& fills = Normalized(table.fills, fill_default);
  EmitCount(out, kBrtBeginFills, fills.size());
  for (const FillRecord& fill : fills)
    EmitFill(out, fill);
  emit_record(out, kBrtEndFills, empty);

  std::vector<BorderRecord> border_default;
  const std::vector<BorderRecord>& borders = Normalized(table.borders, border_default);
  EmitCount(out, kBrtBeginBorders, borders.size());
  for (const BorderRecord& border : borders)
    EmitBorder(out, border);
  emit_record(out, kBrtEndBorders, empty);

  std::vector<CellXf> style_default;
  const std::vector<CellXf>& style_xfs = Normalized(table.cell_style_xfs, style_default);
  if (style_plan != nullptr) {
    EmitCount(out, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginCellStyleXFs),
              style_plan->style_xf_order().size());
    for (const std::uint32_t x : style_plan->style_xf_order())
      EmitXf(out, style_xfs[x], std::nullopt);
  } else {
    EmitCount(out, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginCellStyleXFs), style_xfs.size());
    for (const CellXf& xf : style_xfs)
      EmitXf(out, xf, std::nullopt);
  }
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtEndCellStyleXFs), empty);

  std::vector<CellXf> cell_default;
  const std::vector<CellXf>& cell_xfs = Normalized(table.cell_xfs, cell_default);
  if (style_plan != nullptr && !table.cell_xfs.empty()) {
    EmitCount(out, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginCellXFs), style_plan->cell_xf_order().size());
    for (const std::uint32_t c : style_plan->cell_xf_order())
      EmitXf(out, cell_xfs[c], style_plan->cell_xf_parent(cell_xfs[c].xf_id));
  } else {
    EmitCount(out, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginCellXFs), cell_xfs.size());
    for (const CellXf& xf : cell_xfs)
      EmitXf(out, xf, style_plan != nullptr ? style_plan->cell_xf_parent(xf.xf_id) : xf.xf_id);
  }
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtEndCellXFs), empty);

  // The collection is mandatory. A plan always holds Normal; without one, a
  // book with no named styles gets a synthetic Normal on style-XF 0, which is
  // always present because it was normalised above.
  if (style_plan != nullptr) {
    EmitCount(out, kBrtBeginStyles, style_plan->cell_styles().size());
    for (const CellStyleRecord& style : style_plan->cell_styles())
      EmitStyle(out, style);
  } else if (table.cell_styles.empty()) {
    EmitCount(out, kBrtBeginStyles, 1U);
    std::vector<std::uint8_t> payload;
    emit_u32(payload, 0U);
    emit_u16(payload, 0U);
    emit_u8(payload, 0U);
    emit_u8(payload, 0U);
    emit_xlwidestring(payload, "Normal");
    emit_record(out, kBrtStyle, payload);
  } else {
    EmitCount(out, kBrtBeginStyles, table.cell_styles.size());
    for (const CellStyleRecord& style : table.cell_styles)
      EmitStyle(out, style);
  }
  emit_record(out, kBrtEndStyles, empty);
  const DxfSavePlan identity = plan == nullptr ? identity_dxf_plan(table.dxfs.size()) : DxfSavePlan();
  const std::vector<std::uint32_t>& order = plan == nullptr ? identity.order() : plan->order();
  // Excel writes the dxf list even when it is empty.
  std::vector<std::uint8_t> dxf_count;
  emit_u32(dxf_count, static_cast<std::uint32_t>(order.size()));
  emit_record(out, kBrtBeginDxfs, dxf_count);
  for (const std::uint32_t index : order) {
    EmitDxf(out, table.dxfs[index]);
  }
  emit_record(out, kBrtEndDxfs, empty);
  EmitTableStyles(out, table, plan);
  emit_record(out, kBrtEndStyleSheet, empty);
  return out;
}

std::optional<std::size_t> styles_bin_dxf_count(const std::vector<std::uint8_t>& raw) {
  const std::vector<FramedRecord> recs = split_records(raw);
  if (recs.empty() && !raw.empty()) {
    return std::nullopt;
  }
  std::optional<std::size_t> begin;
  std::size_t count = 0;
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (recs[i].type == kBrtBeginDxfs) {
      if (begin.has_value() || recs[i].payload.size < 4U) {
        return std::nullopt;
      }
      begin = i;
    } else if (recs[i].type == kBrtDxf) {
      if (!begin.has_value() || i != *begin + 1U + count) {
        return std::nullopt;
      }
      ++count;
    }
  }
  return count;
}

Expected<std::vector<std::uint8_t>, Error> compact_styles_bin(const std::vector<std::uint8_t>& raw,
                                                              const DxfSavePlan& plan) {
  if (!styles_bin_dxf_count(raw).has_value()) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb styles part's dxf list is not one block",
                      "context=xlsb_styles_dxfs");
  }
  auto mapped = rewrite_dxf_slots(raw, plan, DxfPartKey{-1, DxfPartType::kStyles, 0U});
  if (!mapped) {
    return std::move(mapped.error());
  }
  const std::vector<std::uint8_t>& src = mapped.value();
  const std::vector<FramedRecord> recs = split_records(src);
  std::vector<const FramedRecord*> dxfs;
  for (const FramedRecord& rec : recs) {
    if (rec.type == kBrtDxf) {
      dxfs.push_back(&rec);
    }
  }
  for (const std::uint32_t index : plan.order()) {
    if (index >= dxfs.size()) {
      return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb styles part holds fewer dxfs than the plan",
                        "context=xlsb_styles_dxfs");
    }
  }
  std::vector<std::uint8_t> out;
  out.reserve(src.size());
  for (const FramedRecord& rec : recs) {
    if (rec.type == kBrtDxf) {
      continue;
    }
    append_record(out, src, rec);
    if (rec.type == kBrtBeginDxfs) {
      // The count is the payload's first u32; the record's framing bytes are kept.
      const std::size_t payload_at = static_cast<std::size_t>(rec.payload.data - src.data());
      store_u32(out.data() + out.size() - (rec.end - payload_at), static_cast<std::uint32_t>(plan.saved_count()));
      for (const std::uint32_t index : plan.order()) {
        append_record(out, src, *dxfs[index]);
      }
    }
  }
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
