//
// The single owning copy of the built-in number-format table.
// `builtin_num_fmt(id)` is declared in `styles.h` so the styles reader and
// writer share one `.rodata` definition; duplicating the table across
// translation units would inflate the WASM binary unnecessarily.

#include "styles.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

namespace formulon {
namespace {

/// Copies the alignment attributes (and their presence bits) from `from`.
void copy_alignment(CellXf& to, const CellXf& from) {
  to.horizontal_align = from.horizontal_align;
  to.vertical_align = from.vertical_align;
  to.wrap_text = from.wrap_text;
  to.justify_last_line = from.justify_last_line;
  to.has_alignment = from.has_alignment;
  to.has_horizontal_align = from.has_horizontal_align;
  to.has_vertical_align = from.has_vertical_align;
  to.has_wrap_text = from.has_wrap_text;
  to.has_justify_last_line = from.has_justify_last_line;
  to.has_text_rotation = from.has_text_rotation;
  to.text_rotation = from.text_rotation;
  to.has_indent = from.has_indent;
  to.indent = from.indent;
  to.has_relative_indent = from.has_relative_indent;
  to.relative_indent = from.relative_indent;
  to.has_shrink_to_fit = from.has_shrink_to_fit;
  to.shrink_to_fit = from.shrink_to_fit;
  to.has_reading_order = from.has_reading_order;
  to.reading_order = from.reading_order;
}

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

}  // namespace

Expected<void, Error> remove_cell_style(StylesTable& styles, std::string_view name) {
  std::size_t pos = styles.cell_styles.size();
  for (std::size_t i = 0; i < styles.cell_styles.size(); ++i) {
    if (styles.cell_styles[i].name == name) {
      pos = i;
      break;
    }
  }
  if (pos == styles.cell_styles.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "remove_cell_style: no such cell style",
                      "name=" + std::string(name));
  }
  if (styles.cell_styles[pos].builtin_id == 0U) {
    return make_error(FormulonErrorCode::kInvalidArgument, "remove_cell_style: the Normal style cannot be removed",
                      "name=" + std::string(name));
  }
  const std::uint32_t removed_xf = styles.cell_styles[pos].xf_id;
  styles.cell_styles.erase(styles.cell_styles.begin() + static_cast<std::ptrdiff_t>(pos));
  if (removed_xf == 0U) {
    return Expected<void, Error>::Ok();
  }
  for (const CellStyleRecord& other : styles.cell_styles) {
    if (other.xf_id == removed_xf) {
      return Expected<void, Error>::Ok();
    }
  }
  const CellXf normal = styles.cell_style_xfs.empty() ? CellXf{} : styles.cell_style_xfs[0];
  for (CellXf& xf : styles.cell_xfs) {
    if (xf.xf_id != removed_xf) {
      continue;
    }
    xf.xf_id = 0U;
    if (!xf.apply_number_format) {
      xf.num_fmt_id = normal.num_fmt_id;
    }
    if (!xf.apply_font) {
      xf.font_index = normal.font_index;
    }
    if (!xf.apply_fill) {
      xf.fill_index = normal.fill_index;
    }
    if (!xf.apply_border) {
      xf.border_index = normal.border_index;
    }
    if (!xf.apply_alignment) {
      copy_alignment(xf, normal);
    }
    if (!xf.apply_protection) {
      xf.has_protection = normal.has_protection;
      xf.locked = normal.locked;
      xf.hidden = normal.hidden;
    }
  }
  return Expected<void, Error>::Ok();
}

const char* builtin_num_fmt(std::uint16_t id) {
  if (id >= kBuiltinNumFmts.size()) {
    return "";
  }
  return kBuiltinNumFmts[id];
}

std::optional<std::string_view> effective_num_fmt(const formulon::StylesTable& styles, std::uint16_t id) {
  for (const formulon::NumFmtRecord& record : styles.num_fmts) {
    if (record.id == id && record.format_string_index < styles.num_fmt_strings.size()) {
      return styles.num_fmt_strings[record.format_string_index];
    }
  }
  if (id < 164U) {
    const char* builtin = formulon::builtin_num_fmt(id);
    if (builtin != nullptr && builtin[0] != '\0') {
      return std::string_view(builtin);
    }
  }
  return std::nullopt;
}

}  // namespace formulon
