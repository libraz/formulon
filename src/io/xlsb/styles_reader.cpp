//
// Implementation of the `xl/styles.bin` reader. See
// `io/xlsb/styles_reader.h` for the contract.
//
// Record shapes below were derived from a real Excel-365-produced
// `xl/styles.bin` (byte-level verification, not transcribed from a
// third-party summary):
//
//   BrtFmt   (44): u16 ifmt, u32 cch, cch x UTF-16LE code units.
//   BrtXF    (47): 8 x u16 — [ixfeParent, numFmtId, fontId, fillId,
//                  borderId, trot|indent, flags, xfGrbitAtr]; the two
//                  bit fields are laid out in `io/xlsb/xf_flags.h`.
//                  Appears once per entry in both the `<cellStyleXfs>`
//                  block (bracketed by BrtBeginCellStyleXFs /
//                  BrtEndCellStyleXFs) and the `<cellXfs>` block
//                  (bracketed by BrtBeginCellXFs / BrtEndCellXFs); which
//                  block a given BrtXF belongs to is tracked by the most
//                  recently seen Begin marker.
//   BrtColor (—):  8 bytes — flags(fValidRGB in bit 0, XColorType in bits
//                  1..7), palette/theme index, i16 tint scaled to
//                  +-32767, then red, green, blue, alpha.
//   BrtFont  (43): u16 dyHeight (twips), u16 grbit, u16 bls (weight),
//                  u16 sss (vertical alignment), u8 uls (underline),
//                  u8 bFamily, u8 bCharSet, u8 unused, BrtColor,
//                  u8 bFontScheme, XLWideString name.
//   BrtFill  (45): u32 fls (pattern), BrtColor foreground, BrtColor
//                  background, then the gradient tail (u32 type,
//                  5 x double, u32 stop count).
//   BrtBorder(46): u8 diagonal flags (bit 0 down, bit 1 up), then five
//                  sides in the order top, bottom, left, right,
//                  diagonal, each u8 style + u8 unused + BrtColor.
//
// Every one of these is decoded into the shared `io::StylesTable`, so a
// `.xlsb`-sourced workbook hands its consumers the same font / fill /
// border / alignment attributes an `.xlsx`-sourced one does. The record
// layouts are symmetric with `io/xlsb/styles_writer.cpp`, which emits them.

#include "io/xlsb/styles_reader.h"

#include <cstdint>
#include <string>
#include <utility>

#include "io/xlsb/record.h"
#include "io/xlsb/xf_flags.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// Which `<cellXfs>`-family table is currently being populated. `BrtXF`
// records outside both brackets (should not happen in a well-formed
// part, but bounds-checked input requires a defined behaviour) are
// dropped rather than misattributed.
enum class XfTarget { kNone, kCellStyleXfs, kCellXfs };

/// The palette slot Excel writes for "no colour of my own". Paired with
/// the automatic `XColorType` it is how a `BrtColor` field spells an
/// absent `<color>` element, which OOXML expresses by omitting the
/// element altogether.
constexpr std::uint8_t kAutomaticPaletteIndex = 0x40U;
/// The `<bgColor>` half of the same idiom: a fill that specifies neither
/// colour carries the system-foreground / system-background pair.
constexpr std::uint8_t kBackgroundPaletteIndex = 0x41U;

/// Decodes an 8-byte `BrtColor` into the shared `(argb, ColorSpec)` pair.
///
/// `XColorType` (bits 1..7 of the flags byte) selects how the entry is to
/// be read: `0` automatic, `1` legacy palette index, `2` literal RGB, `3`
/// theme index with a tint. The RGB triple is present whatever the
/// selector says, so `argb` is filled for every selector and stays a
/// compatibility fallback for the non-RGB ones, exactly as the OOXML
/// reader treats `<color rgb="...">` alongside `theme` / `indexed`.
///
/// The automatic selector on the automatic palette slot is the encoding
/// of an absent `<color>`, and resolves to `kNone` with `unset_argb` --
/// the same state the OOXML reader leaves behind when the element is
/// missing -- so the two containers describe such a record identically.
/// An automatic selector on any other slot is a genuine `auto="1"`.
///
/// The writer does not produce that spelling: a `kNone` spec still
/// carries an ARGB the OOXML writer would serialise as `rgb="..."`, so
/// `styles_writer.cpp` emits it as literal RGB rather than as the absent
/// idiom. Reading Excel's spelling is therefore a widening, not a
/// round-trip pair.
Expected<void, Error> DecodeColor(ByteSpan& p, std::uint32_t unset_argb, std::uint32_t& argb, ColorSpec& spec) {
  auto flags_or = read_u8(p);
  auto index_or = read_u8(p);
  auto tint_or = read_u16(p);
  auto red_or = read_u8(p);
  auto green_or = read_u8(p);
  auto blue_or = read_u8(p);
  auto alpha_or = read_u8(p);
  if (!flags_or || !index_or || !tint_or || !red_or || !green_or || !blue_or || !alpha_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtColor truncated",
                      "context=xlsb_styles_reader");
  }
  argb = (static_cast<std::uint32_t>(alpha_or.value()) << 24U) | (static_cast<std::uint32_t>(red_or.value()) << 16U) |
         (static_cast<std::uint32_t>(green_or.value()) << 8U) | static_cast<std::uint32_t>(blue_or.value());
  switch (static_cast<std::uint32_t>(flags_or.value() >> 1U)) {
    case 0U:
      if (index_or.value() == kAutomaticPaletteIndex) {
        spec.kind = ColorSpec::Kind::kNone;
        argb = unset_argb;
      } else {
        spec.kind = ColorSpec::Kind::kAuto;
      }
      break;
    case 1U:
      spec.kind = ColorSpec::Kind::kIndexed;
      spec.indexed = index_or.value();
      break;
    case 2U:
      spec.kind = ColorSpec::Kind::kRgb;
      spec.rgb = argb;
      break;
    case 3U:
      spec.kind = ColorSpec::Kind::kTheme;
      spec.theme = index_or.value();
      spec.tint = static_cast<double>(static_cast<std::int16_t>(tint_or.value())) / 32767.0;
      break;
    default:
      // A selector outside the four defined values names no colour this
      // model can reproduce; leave the spec unset so the writer emits no
      // `<color>` element rather than an invented one.
      spec.kind = ColorSpec::Kind::kNone;
      break;
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> DecodeFont(ByteSpan payload, StylesTable& table) {
  ByteSpan p = payload;
  auto height_or = read_u16(p);
  auto grbit_or = read_u16(p);
  auto weight_or = read_u16(p);
  auto vert_align_or = read_u16(p);
  auto underline_or = read_u8(p);
  auto family_or = read_u8(p);
  auto charset_or = read_u8(p);
  auto unused_or = read_u8(p);
  if (!height_or || !grbit_or || !weight_or || !vert_align_or || !underline_or || !family_or || !charset_or ||
      !unused_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtFont fields truncated",
                      "context=xlsb_styles_reader");
  }
  FontRecord rec;
  rec.size = static_cast<double>(height_or.value()) / 20.0;
  // `grbit` carries italic in bit 1 and strikethrough in bit 3. Weight is
  // its own field (400 normal, 700 bold), and Excel sets grbit bit 0
  // alongside a bold weight, so bold is read from `bls` and the redundant
  // flag bit ignored. The presence bits follow the OOXML reader's rule --
  // Excel writes `<b/>` only when the toggle is on -- so a font read from
  // an `.xlsb` converts to the same `<font>` element as its `.xlsx` twin.
  rec.bold = weight_or.value() >= 550U;
  rec.has_bold = rec.bold;
  rec.italic = (grbit_or.value() & 0x0002U) != 0U;
  rec.has_italic = rec.italic;
  rec.strike = (grbit_or.value() & 0x0008U) != 0U;
  rec.has_strike = rec.strike;
  switch (underline_or.value()) {
    case 0x01U:
      rec.underline = 1U;  // single
      break;
    case 0x02U:
      rec.underline = 2U;  // double
      break;
    case 0x21U:
      rec.underline = 3U;  // singleAccounting
      break;
    case 0x22U:
      rec.underline = 4U;  // doubleAccounting
      break;
    default:
      rec.underline = 0U;
      break;
  }
  switch (vert_align_or.value()) {
    case 1U:
      rec.vert_align = 1U;  // superscript
      break;
    case 2U:
      rec.vert_align = 2U;  // subscript
      break;
    default:
      rec.vert_align = 0U;  // baseline
      break;
  }
  // `bFamily` / `bCharSet` are mandatory fields with no presence bit of
  // their own. `0` is both "not specified" and a legal value; the OOXML
  // side distinguishes the two by element presence, and Excel omits the
  // element for `0`, so treat `0` as absent to keep the two paths equal.
  rec.has_family = family_or.value() != 0U;
  rec.family = family_or.value();
  rec.has_charset = charset_or.value() != 0U;
  rec.charset = charset_or.value();
  if (auto color = DecodeColor(p, /*unset_argb=*/rec.color_argb, rec.color_argb, rec.color); !color) {
    return color.error();
  }
  auto scheme_or = read_u8(p);
  if (!scheme_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtFont scheme truncated",
                      "context=xlsb_styles_reader");
  }
  // `bFontScheme` shares the shared model's ordinals (0=none, 1=major,
  // 2=minor). Anything else is a value this build cannot name, and
  // writing back a link Excel would not resolve is worse than dropping it.
  rec.scheme = scheme_or.value() <= 2U ? scheme_or.value() : 0U;
  auto name_or = read_xlwidestring(p);
  if (!name_or) {
    return name_or.error();
  }
  rec.name = std::move(name_or.value());
  table.fonts.push_back(std::move(rec));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> DecodeFill(ByteSpan payload, StylesTable& table) {
  ByteSpan p = payload;
  auto pattern_or = read_u32(p);
  if (!pattern_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtFill pattern truncated",
                      "context=xlsb_styles_reader");
  }
  FillRecord rec;
  // `fls` beyond the 19 OOXML pattern values names a gradient fill, whose
  // stops the shared model does not carry; it degrades to "no pattern"
  // the same way an unrecognised `patternType` does on the OOXML side.
  rec.pattern = pattern_or.value() <= 18U ? static_cast<std::uint8_t>(pattern_or.value()) : std::uint8_t{0};
  // A `none` fill carries the placeholder indexed 64 / 65 pair on the wire
  // but no `<fgColor>` / `<bgColor>` in OOXML. Leaving the colours unset
  // keeps the converted `<fill>` identical to the one Excel writes.
  if (rec.pattern != 0U) {
    if (auto fg = DecodeColor(p, /*unset_argb=*/0U, rec.fg_argb, rec.fg); !fg) {
      return fg.error();
    }
    if (auto bg = DecodeColor(p, /*unset_argb=*/0U, rec.bg_argb, rec.bg); !bg) {
      return bg.error();
    }
    // The system-foreground / system-background palette pair is how a
    // patterned fill says it chose neither colour; Excel writes no
    // `<fgColor>` / `<bgColor>` for it. A fill the user did colour keeps
    // its own foreground, so the pair is unambiguous.
    if (rec.fg.kind == ColorSpec::Kind::kIndexed && rec.fg.indexed == kAutomaticPaletteIndex &&
        rec.bg.kind == ColorSpec::Kind::kIndexed && rec.bg.indexed == kBackgroundPaletteIndex) {
      rec.fg = ColorSpec{};
      rec.bg = ColorSpec{};
      rec.fg_argb = 0U;
      rec.bg_argb = 0U;
    }
  }
  table.fills.push_back(rec);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> DecodeBorderSide(ByteSpan& p, BorderSide& side) {
  auto style_or = read_u8(p);
  auto unused_or = read_u8(p);
  if (!style_or || !unused_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtBorder side truncated",
                      "context=xlsb_styles_reader");
  }
  // Styles above the 14 OOXML ordinals have no name to serialise to, so
  // they read as `none`, matching `ParseBorderStyle`'s unknown-string case.
  side.style = style_or.value() <= 13U ? style_or.value() : std::uint8_t{0};
  std::uint32_t argb = 0;
  ColorSpec spec;
  if (auto color = DecodeColor(p, /*unset_argb=*/0U, argb, spec); !color) {
    return color.error();
  }
  // A `none` side has no `<color>` child in OOXML; the wire always carries
  // one, so drop it rather than emit a colour Excel never wrote.
  if (side.style != 0U) {
    side.color_argb = argb;
    side.color = spec;
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> DecodeBorder(ByteSpan payload, StylesTable& table) {
  ByteSpan p = payload;
  auto flags_or = read_u8(p);
  if (!flags_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtBorder flags truncated",
                      "context=xlsb_styles_reader");
  }
  BorderRecord rec;
  rec.diagonal_down = (flags_or.value() & 0x01U) != 0U;
  rec.diagonal_up = (flags_or.value() & 0x02U) != 0U;
  for (BorderSide* side : {&rec.top, &rec.bottom, &rec.left, &rec.right, &rec.diagonal}) {
    if (auto decoded = DecodeBorderSide(p, *side); !decoded) {
      return decoded.error();
    }
  }
  table.borders.push_back(rec);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> DecodeFmt(ByteSpan payload, StylesTable& table) {
  ByteSpan p = payload;
  auto ifmt_or = read_u16(p);
  if (!ifmt_or) {
    return ifmt_or.error();
  }
  auto name_or = read_xlwidestring(p);
  if (!name_or) {
    return name_or.error();
  }
  NumFmtRecord rec;
  rec.id = ifmt_or.value();
  rec.format_string_index = static_cast<std::uint32_t>(table.num_fmt_strings.size());
  table.num_fmt_strings.push_back(std::move(name_or.value()));
  table.num_fmts.push_back(rec);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> DecodeXf(ByteSpan payload, XfTarget target, StylesTable& table) {
  ByteSpan p = payload;
  // 8 x u16: [ixfeParent, numFmtId, fontId, fillId, borderId,
  // trot|indent, flags, xfGrbitAtr]. Field positions are in
  // `io/xlsb/xf_flags.h`, shared with the writer.
  auto parent_or = read_u16(p);
  if (!parent_or) {
    return parent_or.error();
  }
  auto num_fmt_id_or = read_u16(p);
  if (!num_fmt_id_or) {
    return num_fmt_id_or.error();
  }
  auto font_id_or = read_u16(p);
  if (!font_id_or) {
    return font_id_or.error();
  }
  auto fill_id_or = read_u16(p);
  if (!fill_id_or) {
    return fill_id_or.error();
  }
  auto border_id_or = read_u16(p);
  if (!border_id_or) {
    return border_id_or.error();
  }
  auto text_rotation_or = read_u8(p);
  if (!text_rotation_or) {
    return text_rotation_or.error();
  }
  auto indent_or = read_u8(p);
  if (!indent_or) {
    return indent_or.error();
  }
  auto flags_or = read_u16(p);
  if (!flags_or) {
    return flags_or.error();
  }
  auto apply_or = read_u16(p);
  if (!apply_or) {
    return apply_or.error();
  }
  const std::uint16_t flags = flags_or.value();
  const std::uint16_t apply = apply_or.value();
  CellXf xf;
  xf.num_fmt_id = num_fmt_id_or.value();
  xf.font_index = font_id_or.value();
  xf.fill_index = fill_id_or.value();
  xf.border_index = border_id_or.value();
  // A `BrtXF` states every field unconditionally, so the OOXML presence
  // bits are derived from the value rather than read: a field holding its
  // schema default is what Excel writes as an omitted attribute, and the
  // `Has*` predicates in `io/styles_reader.h` apply the same rule to the
  // four attributes whose default is indistinguishable from their value.
  // Reproducing it is what makes a converted `.xlsx` carry the same `<xf>`
  // Excel's own `.xlsx` export of the workbook does.
  xf.horizontal_align = static_cast<std::uint8_t>(flags & kXfHorizontalAlignMask);
  xf.vertical_align = static_cast<std::uint8_t>((flags & kXfVerticalAlignMask) >> kXfVerticalAlignShift);
  xf.wrap_text = (flags & kXfWrapText) != 0U;
  xf.justify_last_line = (flags & kXfJustifyLastLine) != 0U;
  xf.shrink_to_fit = (flags & kXfShrinkToFit) != 0U;
  xf.has_shrink_to_fit = xf.shrink_to_fit;
  xf.reading_order = (flags & kXfReadingOrderMask) >> kXfReadingOrderShift;
  xf.has_reading_order = xf.reading_order != 0U;
  xf.text_rotation = text_rotation_or.value();
  xf.has_text_rotation = xf.text_rotation != 0U;
  xf.indent = indent_or.value();
  xf.has_indent = xf.indent != 0U;
  xf.quote_prefix = (flags & kXfQuotePrefix) != 0U;
  xf.locked = (flags & kXfLocked) != 0U;
  xf.hidden = (flags & kXfHidden) != 0U;
  // `<protection>` is absent exactly when both halves hold their schema
  // default (locked, not hidden); anything else has to be spelled out.
  xf.has_protection = !xf.locked || xf.hidden;
  xf.apply_number_format = (apply & kXfApplyNumberFormat) != 0U;
  xf.apply_font = (apply & kXfApplyFont) != 0U;
  xf.apply_fill = (apply & kXfApplyFill) != 0U;
  xf.apply_border = (apply & kXfApplyBorder) != 0U;
  xf.apply_alignment = (apply & kXfApplyAlignment) != 0U;
  xf.apply_protection = (apply & kXfApplyProtection) != 0U;
  switch (target) {
    case XfTarget::kCellStyleXfs:
      // `ixfeParent` on a named-style xf names no record; OOXML omits
      // `xfId` there, which the model spells as the default 0.
      table.cell_style_xfs.push_back(xf);
      break;
    case XfTarget::kCellXfs:
      if (parent_or.value() != kXfNoParent) {
        xf.xf_id = parent_or.value();
      }
      table.cell_xfs.push_back(xf);
      break;
    case XfTarget::kNone:
      // A `BrtXF` outside any bracket has no destination table; drop it
      // rather than guessing. Real Excel output never hits this path.
      break;
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<StylesTable, Error> read_styles_bin(ByteSpan bytes) {
  StylesTable table;
  XfTarget xf_target = XfTarget::kNone;
  ByteSpan cursor = bytes;
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    switch (static_cast<XlsbRecordType>(rec.type)) {
      case XlsbRecordType::BrtFmt: {
        if (auto r = DecodeFmt(rec.payload, table); !r) {
          return r.error();
        }
        break;
      }
      case XlsbRecordType::BrtFont: {
        if (auto r = DecodeFont(rec.payload, table); !r) {
          return r.error();
        }
        break;
      }
      case XlsbRecordType::BrtFill: {
        if (auto r = DecodeFill(rec.payload, table); !r) {
          return r.error();
        }
        break;
      }
      case XlsbRecordType::BrtBorder: {
        if (auto r = DecodeBorder(rec.payload, table); !r) {
          return r.error();
        }
        break;
      }
      case XlsbRecordType::BrtBeginCellStyleXFs:
        xf_target = XfTarget::kCellStyleXfs;
        break;
      case XlsbRecordType::BrtEndCellStyleXFs:
        xf_target = XfTarget::kNone;
        break;
      case XlsbRecordType::BrtBeginCellXFs:
        xf_target = XfTarget::kCellXfs;
        break;
      case XlsbRecordType::BrtEndCellXFs:
        xf_target = XfTarget::kNone;
        break;
      case XlsbRecordType::BrtXF: {
        if (auto r = DecodeXf(rec.payload, xf_target, table); !r) {
          return r.error();
        }
        break;
      }
      default:
        break;
    }
  }
  // Empty-document contract mirrors `io::read_styles`: every consumer
  // indexes `cell_xfs[xf_index]` / `fonts[font_index]` / etc. without a
  // bounds check once `xf_index` itself has been validated, so a part
  // that carried none of the relevant records still needs a resolvable
  // default row in each vector.
  if (table.fonts.empty()) {
    table.fonts.push_back(FontRecord{});
  }
  if (table.fills.empty()) {
    table.fills.push_back(FillRecord{});
  }
  if (table.borders.empty()) {
    table.borders.push_back(BorderRecord{});
  }
  if (table.cell_xfs.empty()) {
    table.cell_xfs.push_back(CellXf{});
  }
  return table;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
