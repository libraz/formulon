//
// Implementation of the styles writer plus the single owning copy of
// the built-in number-format table. `builtin_num_fmt(id)` is exported
// here (declared in `styles_reader.h`) so reader and writer share one
// `.rodata` definition; duplicating the table across translation units
// would inflate the WASM binary unnecessarily.

#include "io/styles_writer.h"

#include <array>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>

#include "io/styles_reader.h"
#include "io/xml_escape.h"

namespace formulon {
namespace io {
namespace {

// Built-in Excel number-format ids. The id space is sparse (Excel only
// defines a subset of 0..163; the remaining slots are reserved). Empty
// strings indicate "not a documented built-in" and behave the same way
// as an out-of-range id.
//
// Source: ECMA-376 Part 1, §18.8.30 (numFmt) and §18.8.31 (numFmts).
constexpr std::array<const char*, 164> kBuiltinNumFmts = {
    /*  0 */ "General",
    /*  1 */ "0",
    /*  2 */ "0.00",
    /*  3 */ "#,##0",
    /*  4 */ "#,##0.00",
    /*  5 */ "",
    /*  6 */ "",
    /*  7 */ "",
    /*  8 */ "",
    /*  9 */ "0%",
    /* 10 */ "0.00%",
    /* 11 */ "0.00E+00",
    /* 12 */ "# ?/?",
    /* 13 */ "# ?\?/?\?",
    /* 14 */ "mm-dd-yy",
    /* 15 */ "d-mmm-yy",
    /* 16 */ "d-mmm",
    /* 17 */ "mmm-yy",
    /* 18 */ "h:mm AM/PM",
    /* 19 */ "h:mm:ss AM/PM",
    /* 20 */ "h:mm",
    /* 21 */ "h:mm:ss",
    /* 22 */ "m/d/yy h:mm",
    /* 23 */ "",
    /* 24 */ "",
    /* 25 */ "",
    /* 26 */ "",
    /* 27 */ "",
    /* 28 */ "",
    /* 29 */ "",
    /* 30 */ "",
    /* 31 */ "",
    /* 32 */ "",
    /* 33 */ "",
    /* 34 */ "",
    /* 35 */ "",
    /* 36 */ "",
    /* 37 */ "#,##0 ;(#,##0)",
    /* 38 */ "#,##0 ;[Red](#,##0)",
    /* 39 */ "#,##0.00;(#,##0.00)",
    /* 40 */ "#,##0.00;[Red](#,##0.00)",
    /* 41 */ "",
    /* 42 */ "",
    /* 43 */ "",
    /* 44 */ "",
    /* 45 */ "mm:ss",
    /* 46 */ "[h]:mm:ss",
    /* 47 */ "mmss.0",
    /* 48 */ "##0.0E+0",
    /* 49 */ "@",
    /* 50 */ "",
    /* 51 */ "",
    /* 52 */ "",
    /* 53 */ "",
    /* 54 */ "",
    /* 55 */ "",
    /* 56 */ "",
    /* 57 */ "",
    /* 58 */ "",
    /* 59 */ "",
    /* 60 */ "",
    /* 61 */ "",
    /* 62 */ "",
    /* 63 */ "",
    /* 64 */ "",
    /* 65 */ "",
    /* 66 */ "",
    /* 67 */ "",
    /* 68 */ "",
    /* 69 */ "",
    /* 70 */ "",
    /* 71 */ "",
    /* 72 */ "",
    /* 73 */ "",
    /* 74 */ "",
    /* 75 */ "",
    /* 76 */ "",
    /* 77 */ "",
    /* 78 */ "",
    /* 79 */ "",
    /* 80 */ "",
    /* 81 */ "",
    /* 82 */ "",
    /* 83 */ "",
    /* 84 */ "",
    /* 85 */ "",
    /* 86 */ "",
    /* 87 */ "",
    /* 88 */ "",
    /* 89 */ "",
    /* 90 */ "",
    /* 91 */ "",
    /* 92 */ "",
    /* 93 */ "",
    /* 94 */ "",
    /* 95 */ "",
    /* 96 */ "",
    /* 97 */ "",
    /* 98 */ "",
    /* 99 */ "",
    /*100 */ "",
    /*101 */ "",
    /*102 */ "",
    /*103 */ "",
    /*104 */ "",
    /*105 */ "",
    /*106 */ "",
    /*107 */ "",
    /*108 */ "",
    /*109 */ "",
    /*110 */ "",
    /*111 */ "",
    /*112 */ "",
    /*113 */ "",
    /*114 */ "",
    /*115 */ "",
    /*116 */ "",
    /*117 */ "",
    /*118 */ "",
    /*119 */ "",
    /*120 */ "",
    /*121 */ "",
    /*122 */ "",
    /*123 */ "",
    /*124 */ "",
    /*125 */ "",
    /*126 */ "",
    /*127 */ "",
    /*128 */ "",
    /*129 */ "",
    /*130 */ "",
    /*131 */ "",
    /*132 */ "",
    /*133 */ "",
    /*134 */ "",
    /*135 */ "",
    /*136 */ "",
    /*137 */ "",
    /*138 */ "",
    /*139 */ "",
    /*140 */ "",
    /*141 */ "",
    /*142 */ "",
    /*143 */ "",
    /*144 */ "",
    /*145 */ "",
    /*146 */ "",
    /*147 */ "",
    /*148 */ "",
    /*149 */ "",
    /*150 */ "",
    /*151 */ "",
    /*152 */ "",
    /*153 */ "",
    /*154 */ "",
    /*155 */ "",
    /*156 */ "",
    /*157 */ "",
    /*158 */ "",
    /*159 */ "",
    /*160 */ "",
    /*161 */ "",
    /*162 */ "",
    /*163 */ ""};

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr std::string_view kXmlNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

void AppendUint(std::string& out, std::uint64_t v) {
  char buf[24];
  std::snprintf(buf, sizeof(buf), "%llu", static_cast<unsigned long long>(v));
  out.append(buf);
}

void AppendDouble(std::string& out, double v) {
  char buf[32];
  // %.17g is round-trip safe under IEEE 754, so a font size such as 10.5
  // survives a recalc-save unchanged. Matches the row-height / column-
  // width writers in the worksheet builder.
  std::snprintf(buf, sizeof(buf), "%.17g", v);
  out.append(buf);
}

void AppendArgb(std::string& out, std::uint32_t argb) {
  char buf[12];
  std::snprintf(buf, sizeof(buf), "%08X", argb);
  out.append(buf);
}

/// Emits a self-closing colour element (`<color>`, `<fgColor>`, or
/// `<bgColor>`) from a colour spec, using `tag` as the element name.
///
/// When the spec carries no explicit kind (`kNone`) — e.g. a record built
/// programmatically through the bindings that set only the resolved
/// `*_argb` field — falls back to emitting `rgb="<fallback>"`, preserving
/// the legacy writer behaviour for callers that never populate a spec.
void AppendColor(std::string& out, const char* tag, const ColorSpec& spec, std::uint32_t fallback_argb) {
  out.push_back('<');
  out.append(tag);
  switch (spec.kind) {
    case ColorSpec::Kind::kTheme:
      out.append(" theme=\"");
      AppendUint(out, spec.theme);
      out.append("\"");
      if (spec.tint != 0.0) {
        out.append(" tint=\"");
        AppendDouble(out, spec.tint);
        out.append("\"");
      }
      break;
    case ColorSpec::Kind::kIndexed:
      out.append(" indexed=\"");
      AppendUint(out, spec.indexed);
      out.append("\"");
      break;
    case ColorSpec::Kind::kAuto:
      out.append(" auto=\"1\"");
      break;
    case ColorSpec::Kind::kRgb:
      out.append(" rgb=\"");
      AppendArgb(out, spec.rgb);
      out.append("\"");
      break;
    case ColorSpec::Kind::kNone:
    default:
      out.append(" rgb=\"");
      AppendArgb(out, fallback_argb);
      out.append("\"");
      break;
  }
  out.append("/>");
}

/// True when a colour location should emit a `<color>` element: either
/// the source carried an explicit spec, or a non-zero resolved value was
/// set programmatically. Used by fills and borders, where an all-zero
/// colour means "no colour set" and no element is emitted.
bool HasColor(const ColorSpec& spec, std::uint32_t argb) {
  return spec.kind != ColorSpec::Kind::kNone || argb != 0U;
}

const char* HorizontalAlignName(std::uint8_t v) {
  switch (v) {
    case 1:
      return "left";
    case 2:
      return "center";
    case 3:
      return "right";
    case 4:
      return "fill";
    case 5:
      return "justify";
    case 6:
      return "centerContinuous";
    case 7:
      return "distributed";
    default:
      return nullptr;  // 0 = general; omit attribute.
  }
}

const char* VerticalAlignName(std::uint8_t v) {
  switch (v) {
    case 0:
      return "top";
    case 1:
      return "center";
    case 3:
      return "justify";
    case 4:
      return "distributed";
    case 2:
    default:
      return nullptr;  // 2 = bottom (default); omit attribute.
  }
}

const char* UnderlineName(std::uint8_t v) {
  switch (v) {
    case 1:
      return "single";
    case 2:
      return "double";
    case 3:
      return "singleAccounting";
    case 4:
      return "doubleAccounting";
    default:
      return nullptr;
  }
}

const char* VertAlignName(std::uint8_t v) {
  switch (v) {
    case 1:
      return "superscript";
    case 2:
      return "subscript";
    default:
      return nullptr;  // 0 = baseline (default); emit no element.
  }
}

void AppendVertAlign(std::string& out, std::uint8_t v) {
  if (const char* name = VertAlignName(v); name != nullptr) {
    out.append("<vertAlign val=\"");
    out.append(name);
    out.append("\"/>");
  }
}

void AppendFontFamilyCharset(std::string& out, const FontRecord& f) {
  if (f.has_family) {
    out.append("<family val=\"");
    AppendUint(out, f.family);
    out.append("\"/>");
  }
  if (f.has_charset) {
    out.append("<charset val=\"");
    AppendUint(out, f.charset);
    out.append("\"/>");
  }
}

const char* FillPatternName(std::uint8_t v) {
  switch (v) {
    case 0:
      return "none";
    case 1:
      return "solid";
    case 2:
      return "mediumGray";
    case 3:
      return "darkGray";
    case 4:
      return "lightGray";
    case 5:
      return "darkHorizontal";
    case 6:
      return "darkVertical";
    case 7:
      return "darkDown";
    case 8:
      return "darkUp";
    case 9:
      return "darkGrid";
    case 10:
      return "darkTrellis";
    case 11:
      return "lightHorizontal";
    case 12:
      return "lightVertical";
    case 13:
      return "lightDown";
    case 14:
      return "lightUp";
    case 15:
      return "lightGrid";
    case 16:
      return "lightTrellis";
    case 17:
      return "gray125";
    case 18:
      return "gray0625";
    default:
      return "none";
  }
}

const char* BorderStyleName(std::uint8_t v) {
  switch (v) {
    case 1:
      return "thin";
    case 2:
      return "medium";
    case 3:
      return "dashed";
    case 4:
      return "dotted";
    case 5:
      return "thick";
    case 6:
      return "double";
    case 7:
      return "hair";
    case 8:
      return "mediumDashed";
    case 9:
      return "dashDot";
    case 10:
      return "mediumDashDot";
    case 11:
      return "dashDotDot";
    case 12:
      return "mediumDashDotDot";
    case 13:
      return "slantDashDot";
    default:
      return nullptr;  // 0 = none; emit no style attribute.
  }
}

void AppendBorderSide(std::string& out, const char* tag, const BorderSide& side) {
  out.append("<");
  out.append(tag);
  const char* style = BorderStyleName(side.style);
  if (style != nullptr) {
    out.append(" style=\"");
    out.append(style);
    out.append("\"");
  }
  if (HasColor(side.color, side.color_argb)) {
    out.append(">");
    AppendColor(out, "color", side.color, side.color_argb);
    out.append("</");
    out.append(tag);
    out.append(">");
  } else {
    out.append("/>");
  }
}

// Inline style fragments (used by both `<dxf>` records and, for fills,
// the section writers). Declared ahead of the section writers that reuse
// them; defined below alongside the other fragment emitters.
void AppendFillFragment(std::string& out, const FillRecord& fill);

void AppendFontToggle(std::string& out, std::string_view name, bool was_explicit, bool value) {
  if (!was_explicit && !value) {
    return;
  }
  out.push_back('<');
  out.append(name);
  if (!value) {
    out.append(" val=\"0\"");
  }
  out.append("/>");
}

void AppendNumFmts(std::string& out, const StylesTable& table) {
  // Count emitted entries: only custom ids (>= 164) survive. Excel
  // rejects packages that redeclare built-in ids.
  std::size_t emit_count = 0;
  for (const NumFmtRecord& n : table.num_fmts) {
    if (n.id >= 164U) {
      ++emit_count;
    }
  }
  if (emit_count == 0) {
    return;
  }
  out.append("  <numFmts count=\"");
  AppendUint(out, emit_count);
  out.append("\">\n");
  for (const NumFmtRecord& n : table.num_fmts) {
    if (n.id < 164U) {
      continue;
    }
    out.append("    <numFmt numFmtId=\"");
    AppendUint(out, n.id);
    out.append("\" formatCode=\"");
    if (n.format_string_index < table.num_fmt_strings.size()) {
      AppendXmlAttrEscaped(out, table.num_fmt_strings[n.format_string_index]);
    }
    out.append("\"/>\n");
  }
  out.append("  </numFmts>\n");
}

void AppendFonts(std::string& out, const StylesTable& table) {
  // Fonts vector always carries at least one entry (the default).
  // Reader guarantees this; writer preserves it.
  const std::size_t count = table.fonts.empty() ? std::size_t{1} : table.fonts.size();
  out.append("  <fonts count=\"");
  AppendUint(out, count);
  out.append("\">\n");
  if (table.fonts.empty()) {
    out.append("    <font><sz val=\"11\"/><name val=\"Calibri\"/></font>\n");
  } else {
    for (const FontRecord& f : table.fonts) {
      out.append("    <font>");
      AppendFontToggle(out, "b", f.has_bold, f.bold);
      AppendFontToggle(out, "i", f.has_italic, f.italic);
      AppendFontToggle(out, "strike", f.has_strike, f.strike);
      if (const char* uname = UnderlineName(f.underline); uname != nullptr) {
        out.append("<u val=\"");
        out.append(uname);
        out.append("\"/>");
      }
      AppendVertAlign(out, f.vert_align);
      out.append("<sz val=\"");
      AppendDouble(out, f.size);
      out.append("\"/>");
      AppendColor(out, "color", f.color, f.color_argb);
      if (!f.name.empty()) {
        out.append("<name val=\"");
        AppendXmlAttrEscaped(out, f.name);
        out.append("\"/>");
      } else {
        out.append("<name val=\"Calibri\"/>");
      }
      AppendFontFamilyCharset(out, f);
      out.append("</font>\n");
    }
  }
  out.append("  </fonts>\n");
}

void AppendFills(std::string& out, const StylesTable& table) {
  const std::size_t count = table.fills.empty() ? std::size_t{1} : table.fills.size();
  out.append("  <fills count=\"");
  AppendUint(out, count);
  out.append("\">\n");
  if (table.fills.empty()) {
    out.append("    <fill><patternFill patternType=\"none\"/></fill>\n");
  } else {
    for (const FillRecord& fill : table.fills) {
      out.append("    ");
      AppendFillFragment(out, fill);
      out.append("\n");
    }
  }
  out.append("  </fills>\n");
}

void AppendBorders(std::string& out, const StylesTable& table) {
  const std::size_t count = table.borders.empty() ? std::size_t{1} : table.borders.size();
  out.append("  <borders count=\"");
  AppendUint(out, count);
  out.append("\">\n");
  if (table.borders.empty()) {
    out.append("    <border/>\n");
  } else {
    for (const BorderRecord& b : table.borders) {
      out.append("    <border");
      if (b.diagonal_up) {
        out.append(" diagonalUp=\"1\"");
      }
      if (b.diagonal_down) {
        out.append(" diagonalDown=\"1\"");
      }
      out.append(">");
      AppendBorderSide(out, "left", b.left);
      AppendBorderSide(out, "right", b.right);
      AppendBorderSide(out, "top", b.top);
      AppendBorderSide(out, "bottom", b.bottom);
      AppendBorderSide(out, "diagonal", b.diagonal);
      out.append("</border>\n");
    }
  }
  out.append("  </borders>\n");
}

void AppendXfBody(std::string& out, const CellXf& xf, bool emit_xf_id) {
  out.append("    <xf numFmtId=\"");
  AppendUint(out, xf.num_fmt_id);
  out.append("\" fontId=\"");
  AppendUint(out, xf.font_index);
  out.append("\" fillId=\"");
  AppendUint(out, xf.fill_index);
  out.append("\" borderId=\"");
  AppendUint(out, xf.border_index);
  out.append("\"");
  if (emit_xf_id) {
    out.append(" xfId=\"");
    AppendUint(out, xf.xf_id);
    out.append("\"");
  }
  auto append_apply = [&out](const char* name, bool value) {
    if (value) {
      out.append(" ");
      out.append(name);
      out.append("=\"1\"");
    }
  };
  append_apply("applyNumberFormat", xf.apply_number_format);
  append_apply("applyFont", xf.apply_font);
  append_apply("applyFill", xf.apply_fill);
  append_apply("applyBorder", xf.apply_border);
  append_apply("applyAlignment", xf.apply_alignment);
  append_apply("applyProtection", xf.apply_protection);
  append_apply("quotePrefix", xf.quote_prefix);
  const char* halign = HorizontalAlignName(xf.horizontal_align);
  const char* valign = VerticalAlignName(xf.vertical_align);
  const bool has_alignment = halign != nullptr || valign != nullptr || xf.wrap_text || xf.justify_last_line;
  if (!has_alignment && !xf.has_protection) {
    out.append("/>\n");
    return;
  }
  out.append(">");
  if (has_alignment) {
    out.append("<alignment");
    if (halign != nullptr) {
      out.append(" horizontal=\"");
      out.append(halign);
      out.append("\"");
    }
    if (valign != nullptr) {
      out.append(" vertical=\"");
      out.append(valign);
      out.append("\"");
    }
    if (xf.wrap_text) {
      out.append(" wrapText=\"1\"");
    }
    if (xf.justify_last_line) {
      out.append(" justifyLastLine=\"1\"");
    }
    out.append("/>");
  }
  if (xf.has_protection) {
    // Child order in CT_Xf is alignment then protection. Both defaults
    // (locked=1, hidden=0) are emitted explicitly so the element survives
    // a re-read even when it matches the schema default.
    out.append("<protection locked=\"");
    out.append(xf.locked ? "1" : "0");
    out.append("\" hidden=\"");
    out.append(xf.hidden ? "1" : "0");
    out.append("\"/>");
  }
  out.append("</xf>\n");
}

void AppendCellStyleXfs(std::string& out, const StylesTable& table) {
  if (table.cell_style_xfs.empty()) {
    return;
  }
  out.append("  <cellStyleXfs count=\"");
  AppendUint(out, table.cell_style_xfs.size());
  out.append("\">\n");
  for (const CellXf& xf : table.cell_style_xfs) {
    AppendXfBody(out, xf, /*emit_xf_id=*/false);
  }
  out.append("  </cellStyleXfs>\n");
}

void AppendCellXfs(std::string& out, const StylesTable& table) {
  const std::size_t count = table.cell_xfs.empty() ? std::size_t{1} : table.cell_xfs.size();
  out.append("  <cellXfs count=\"");
  AppendUint(out, count);
  out.append("\">\n");
  if (table.cell_xfs.empty()) {
    out.append("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\n");
  } else {
    for (const CellXf& xf : table.cell_xfs) {
      AppendXfBody(out, xf, /*emit_xf_id=*/true);
    }
  }
  out.append("  </cellXfs>\n");
}

void AppendCellStyles(std::string& out, const StylesTable& table) {
  if (table.cell_styles.empty()) {
    return;
  }
  out.append("  <cellStyles count=\"");
  AppendUint(out, table.cell_styles.size());
  out.append("\">\n");
  for (const CellStyleRecord& cs : table.cell_styles) {
    out.append("    <cellStyle name=\"");
    AppendXmlAttrEscaped(out, cs.name);
    out.append("\" xfId=\"");
    AppendUint(out, cs.xf_id);
    out.append("\"");
    if (cs.builtin_id != CellStyleRecord::kBuiltinIdNone) {
      out.append(" builtinId=\"");
      AppendUint(out, cs.builtin_id);
      out.append("\"");
    }
    if (cs.i_level != 0U) {
      out.append(" iLevel=\"");
      AppendUint(out, cs.i_level);
      out.append("\"");
    }
    if (cs.hidden) {
      out.append(" hidden=\"1\"");
    }
    if (cs.custom_builtin) {
      out.append(" customBuiltin=\"1\"");
    }
    out.append("/>\n");
  }
  out.append("  </cellStyles>\n");
}

void AppendFontFragment(std::string& out, const FontRecord& f) {
  out.append("<font>");
  AppendFontToggle(out, "b", f.has_bold, f.bold);
  AppendFontToggle(out, "i", f.has_italic, f.italic);
  AppendFontToggle(out, "strike", f.has_strike, f.strike);
  if (const char* uname = UnderlineName(f.underline); uname != nullptr) {
    out.append("<u val=\"");
    out.append(uname);
    out.append("\"/>");
  }
  AppendVertAlign(out, f.vert_align);
  out.append("<sz val=\"");
  AppendDouble(out, f.size);
  out.append("\"/>");
  AppendColor(out, "color", f.color, f.color_argb);
  if (!f.name.empty()) {
    out.append("<name val=\"");
    AppendXmlAttrEscaped(out, f.name);
    out.append("\"/>");
  }
  AppendFontFamilyCharset(out, f);
  out.append("</font>");
}

void AppendFillFragment(std::string& out, const FillRecord& fill) {
  out.append("<fill><patternFill patternType=\"");
  out.append(FillPatternName(fill.pattern));
  out.append("\"");
  const bool has_fg = HasColor(fill.fg, fill.fg_argb);
  const bool has_bg = HasColor(fill.bg, fill.bg_argb);
  if (has_fg || has_bg) {
    out.append(">");
    if (has_fg) {
      AppendColor(out, "fgColor", fill.fg, fill.fg_argb);
    }
    if (has_bg) {
      AppendColor(out, "bgColor", fill.bg, fill.bg_argb);
    }
    out.append("</patternFill>");
  } else {
    out.append("/>");
  }
  out.append("</fill>");
}

void AppendBorderFragment(std::string& out, const BorderRecord& b) {
  out.append("<border");
  if (b.diagonal_up) {
    out.append(" diagonalUp=\"1\"");
  }
  if (b.diagonal_down) {
    out.append(" diagonalDown=\"1\"");
  }
  out.append(">");
  AppendBorderSide(out, "left", b.left);
  AppendBorderSide(out, "right", b.right);
  AppendBorderSide(out, "top", b.top);
  AppendBorderSide(out, "bottom", b.bottom);
  AppendBorderSide(out, "diagonal", b.diagonal);
  out.append("</border>");
}

void AppendDxfs(std::string& out, const StylesTable& table) {
  if (table.dxfs.empty()) {
    return;
  }
  out.append("  <dxfs count=\"");
  AppendUint(out, table.dxfs.size());
  out.append("\">\n");
  for (const DifferentialFormat& dxf : table.dxfs) {
    out.append("    <dxf>");
    if (dxf.has_font) {
      AppendFontFragment(out, dxf.font);
    }
    if (dxf.has_num_fmt) {
      out.append("<numFmt numFmtId=\"");
      AppendUint(out, dxf.num_fmt_id);
      out.append("\" formatCode=\"");
      AppendXmlAttrEscaped(out, dxf.num_fmt_code);
      out.append("\"/>");
    }
    if (dxf.has_fill) {
      AppendFillFragment(out, dxf.fill);
    }
    // CT_Dxf child order: font, numFmt, fill, alignment, border, protection.
    // The captured `<alignment>` / `<protection>` are re-emitted verbatim
    // at their schema positions.
    out.append(dxf.alignment_xml);
    if (dxf.has_border) {
      AppendBorderFragment(out, dxf.border);
    }
    out.append(dxf.protection_xml);
    out.append("</dxf>\n");
  }
  out.append("  </dxfs>\n");
}

}  // namespace

const char* builtin_num_fmt(std::uint16_t id) {
  if (id >= kBuiltinNumFmts.size()) {
    return "";
  }
  return kBuiltinNumFmts[id];
}

std::string write_styles(const StylesTable& table) {
  std::string out;
  out.reserve(1024);
  out.append(kXmlDecl);
  out.append("<styleSheet xmlns=\"");
  out.append(kXmlNs);
  out.append("\"");
  out.append(table.root_extra_attrs);
  out.append(">\n");
  AppendNumFmts(out, table);
  AppendFonts(out, table);
  AppendFills(out, table);
  AppendBorders(out, table);
  AppendCellStyleXfs(out, table);
  AppendCellXfs(out, table);
  AppendCellStyles(out, table);
  AppendDxfs(out, table);
  if (!table.table_styles_xml.empty()) {
    out.append("  ");
    out.append(table.table_styles_xml);
    out.push_back('\n');
  }
  if (!table.colors_xml.empty()) {
    out.append("  ");
    out.append(table.colors_xml);
    out.push_back('\n');
  }
  for (const std::string& raw : table.unknown_top_level_xml) {
    out.append("  ");
    out.append(raw);
    out.push_back('\n');
  }
  if (!table.ext_lst_xml.empty()) {
    out.append("  ");
    out.append(table.ext_lst_xml);
    out.push_back('\n');
  }
  out.append("</styleSheet>\n");
  return out;
}

}  // namespace io
}  // namespace formulon
