
#include "io/xlsb/styles_writer.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <limits>
#include <string_view>
#include <vector>

#include "io/xlsb/brt_color.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/xf_flags.h"

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
constexpr std::uint16_t kBrtBeginDxfs = 505;
constexpr std::uint16_t kBrtEndDxfs = 506;
constexpr std::uint16_t kBrtDxf = 507;

// BrtDXF property types, measured against Excel-saved dxfs (each checked
// against the <dxf> Excel wrote for the same book).
constexpr std::uint16_t kPropFillPattern = 0;
constexpr std::uint16_t kPropFillFg = 1;
constexpr std::uint16_t kPropFillBg = 2;
constexpr std::uint16_t kPropFontColor = 5;
constexpr std::uint16_t kPropBorderTop = 6;
constexpr std::uint16_t kPropBorderBottom = 7;
constexpr std::uint16_t kPropBorderLeft = 8;
constexpr std::uint16_t kPropBorderRight = 9;
constexpr std::uint16_t kPropBorderDiagonal = 10;
constexpr std::uint16_t kPropDiagonalUp = 13;
constexpr std::uint16_t kPropDiagonalDown = 14;
constexpr std::uint16_t kPropFontName = 24;
constexpr std::uint16_t kPropFontWeight = 25;
constexpr std::uint16_t kPropUnderline = 26;
constexpr std::uint16_t kPropVertAlign = 27;
constexpr std::uint16_t kPropItalic = 28;
constexpr std::uint16_t kPropStrike = 29;
constexpr std::uint16_t kPropFontSize = 36;
constexpr std::uint16_t kPropNumFmtCode = 38;
constexpr std::uint16_t kPropNumFmtId = 41;
constexpr std::uint8_t kUnderlineMap[] = {0U, 1U, 2U, 0x21U, 0x22U};

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
  std::uint16_t flags = 0U;
  if (font.italic)
    flags |= 0x0002U;
  if (font.strike)
    flags |= 0x0008U;
  emit_u16(payload, flags);
  emit_u16(payload, font.bold ? 700U : 400U);
  emit_u16(payload, font.vert_align == 1U ? 1U : (font.vert_align == 2U ? 2U : 0U));
  emit_u8(payload, kUnderlineMap[std::min<std::uint8_t>(font.underline, 4U)]);
  emit_u8(payload, font.has_family ? font.family : 0U);
  emit_u8(payload, font.has_charset ? font.charset : 0U);
  emit_u8(payload, 0U);
  EmitColor(payload, font.color_argb, font.color);
  emit_u8(payload, font.scheme);  // bFontScheme shares the shared model's ordinals
  emit_xlwidestring(payload, font.name.empty() ? std::string_view("Calibri") : std::string_view(font.name));
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtFont), payload);
}

void EmitFill(std::vector<std::uint8_t>& out, const FillRecord& fill) {
  std::vector<std::uint8_t> payload;
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

void EmitXf(std::vector<std::uint8_t>& out, const CellXf& xf, bool is_style_xf) {
  std::vector<std::uint8_t> payload;
  emit_u16(payload, is_style_xf ? kXfNoParent : ClampIndex(xf.xf_id));
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
  std::vector<std::uint8_t> wide;
  emit_xlwidestring(wide, text);
  std::vector<std::uint8_t> out;
  emit_u16(out, static_cast<std::uint16_t>((wide.size() - 4U) / 2U));
  out.insert(out.end(), wide.begin() + 4, wide.end());
  return out;
}

std::vector<std::uint8_t> ColorProp(std::uint32_t argb, const ColorSpec& spec) {
  std::vector<std::uint8_t> data;
  EmitColor(data, argb, spec);
  return data;
}

/// A dxf as Excel writes it: fill, font, number format, border, each only
/// where the dxf states it. Alignment and protection are left out: Excel
/// discards them from a CF dxf, so they have no measured property.
void EmitDxf(std::vector<std::uint8_t>& out, const DifferentialFormat& dxf) {
  std::vector<std::uint8_t> props;
  std::uint16_t count = 0;
  if (dxf.has_fill) {
    const FillRecord& fill = dxf.fill;
    AppendProp(props, count, kPropFillPattern, {static_cast<std::uint8_t>(fill.pattern <= 18U ? fill.pattern : 0U)});
    if (fill.fg.kind != ColorSpec::Kind::kNone) {
      AppendProp(props, count, kPropFillFg, ColorProp(fill.fg_argb, fill.fg));
    }
    if (fill.bg.kind != ColorSpec::Kind::kNone) {
      AppendProp(props, count, kPropFillBg, ColorProp(fill.bg_argb, fill.bg));
    }
  }
  if (dxf.has_font) {
    const FontRecord& font = dxf.font;
    // 0xFF000000 is the "automatic" sentinel an unstated colour carries.
    if (font.color.kind != ColorSpec::Kind::kNone || font.color_argb != 0xFF000000U) {
      AppendProp(props, count, kPropFontColor, ColorProp(font.color_argb, font.color));
    }
    if (!font.name.empty()) {
      AppendProp(props, count, kPropFontName, ShortWideString(font.name));
    }
    if (font.has_bold && font.bold) {
      AppendProp(props, count, kPropFontWeight, {0xBC, 0x02});  // 700
    }
    if (font.underline != 0U) {
      AppendProp(props, count, kPropUnderline, {kUnderlineMap[std::min<std::uint8_t>(font.underline, 4U)], 0x00});
    }
    if (font.vert_align == 1U || font.vert_align == 2U) {
      AppendProp(props, count, kPropVertAlign, {font.vert_align, 0x00});
    }
    if (font.has_italic && font.italic) {
      AppendProp(props, count, kPropItalic, {0x01});
    }
    if (font.has_strike && font.strike) {
      AppendProp(props, count, kPropStrike, {0x01});
    }
    if (font.has_size) {
      std::vector<std::uint8_t> twips;
      emit_u32(twips, FontHeightTwips(font.size));
      AppendProp(props, count, kPropFontSize, twips);
    }
  }
  if (dxf.has_num_fmt) {
    std::vector<std::uint8_t> id;
    emit_u16(id, dxf.num_fmt_id);
    AppendProp(props, count, kPropNumFmtId, id);
    AppendProp(props, count, kPropNumFmtCode, ShortWideString(dxf.num_fmt_code));
  }
  if (dxf.has_border) {
    const BorderRecord& border = dxf.border;
    const std::pair<std::uint16_t, const BorderSide*> sides[] = {{kPropBorderTop, &border.top},
                                                                 {kPropBorderBottom, &border.bottom},
                                                                 {kPropBorderLeft, &border.left},
                                                                 {kPropBorderRight, &border.right},
                                                                 {kPropBorderDiagonal, &border.diagonal}};
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
      AppendProp(props, count, kPropDiagonalDown, {0x01});
    }
    if (border.diagonal_up) {
      AppendProp(props, count, kPropDiagonalUp, {0x01});
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

}  // namespace

std::vector<std::uint8_t> write_styles_bin(const StylesTable& table) {
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
  const std::vector<FontRecord>& fonts = Normalized(table.fonts, font_default);
  EmitCount(out, kBrtBeginFonts, fonts.size());
  for (const FontRecord& font : fonts)
    EmitFont(out, font);
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
  EmitCount(out, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginCellStyleXFs), style_xfs.size());
  for (const CellXf& xf : style_xfs)
    EmitXf(out, xf, true);
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtEndCellStyleXFs), empty);

  std::vector<CellXf> cell_default;
  const std::vector<CellXf>& cell_xfs = Normalized(table.cell_xfs, cell_default);
  EmitCount(out, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginCellXFs), cell_xfs.size());
  for (const CellXf& xf : cell_xfs)
    EmitXf(out, xf, false);
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtEndCellXFs), empty);

  // The collection is mandatory.  A synthetic Normal style is sufficient for
  // workbooks that do not carry named OOXML styles; style-XF 0 is always
  // present because it was normalised above.
  const std::size_t style_count = std::max<std::size_t>(table.cell_styles.size(), 1U);
  EmitCount(out, kBrtBeginStyles, style_count);
  if (table.cell_styles.empty()) {
    std::vector<std::uint8_t> payload;
    emit_u32(payload, 0U);
    emit_u16(payload, 0U);
    emit_u8(payload, 0U);
    emit_u8(payload, 0U);
    emit_xlwidestring(payload, "Normal");
    emit_record(out, kBrtStyle, payload);
  } else {
    for (const CellStyleRecord& style : table.cell_styles) {
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
  }
  emit_record(out, kBrtEndStyles, empty);
  if (!table.dxfs.empty()) {
    std::vector<std::uint8_t> dxf_count;
    emit_u32(dxf_count, static_cast<std::uint32_t>(table.dxfs.size()));
    emit_record(out, kBrtBeginDxfs, dxf_count);
    for (const DifferentialFormat& dxf : table.dxfs) {
      EmitDxf(out, dxf);
    }
    emit_record(out, kBrtEndDxfs, empty);
  }
  emit_record(out, kBrtEndStyleSheet, empty);
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
