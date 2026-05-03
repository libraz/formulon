// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `xl/styles.xml` reader. Decodes the OOXML styles part into a flat
// in-memory `StylesTable` carrying every record kind the engine needs
// to round-trip a workbook: fonts, fills, borders, custom number-format
// strings, and the per-cell xf table that ties them together.
//
// Number-format strings are interned in a single `std::vector<std::string>`
// owned by `StylesTable`; `NumFmtRecord` carries only an integer index
// into that vector. The 164 built-in Excel format ids (0..163) are
// resolved through `builtin_num_fmt(id)` (see below) and never appear
// in `num_fmt_strings`.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.4 (Reader pipeline)
//   * backup/plans/05-style-runtime.md (Style record schema)

#ifndef FORMULON_IO_STYLES_READER_H_
#define FORMULON_IO_STYLES_READER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// One `<font>` entry inside `<fonts>`.
///
/// `name` is intentionally a `std::string` because the font-name domain
/// is open-set (system fonts, custom fonts, locale-specific variants);
/// the per-record overhead is acceptable. `color_argb` carries the raw
/// AARRGGBB packed colour. The sentinel `0xFF000000U` is preserved
/// verbatim and treated as "automatic" by the consumer.
struct FontRecord {
  std::string name;
  double size = 11.0;
  bool bold = false;
  bool italic = false;
  bool strike = false;
  /// 0=none, 1=single, 2=double, 3=singleAccounting, 4=doubleAccounting.
  std::uint8_t underline = 0;
  std::uint32_t color_argb = 0xFF000000U;
};

/// One `<fill>` entry inside `<fills>`.
///
/// `pattern` is the OOXML pattern index. `0=none`, `1=solid`,
/// `2..18` = the standard pattern set. `fg_argb` / `bg_argb` are the
/// AARRGGBB foreground / background colours.
struct FillRecord {
  std::uint8_t pattern = 0;
  std::uint32_t fg_argb = 0;
  std::uint32_t bg_argb = 0;
};

/// One side of a `<border>` (left/right/top/bottom/diagonal).
struct BorderSide {
  /// 0=none, 1=thin, 2=medium, 3=dashed, ..., 13=slantDashDot
  /// (OOXML border style ordinal).
  std::uint8_t style = 0;
  std::uint32_t color_argb = 0;
};

/// One `<border>` entry inside `<borders>`.
struct BorderRecord {
  BorderSide left;
  BorderSide right;
  BorderSide top;
  BorderSide bottom;
  BorderSide diagonal;
  bool diagonal_up = false;
  bool diagonal_down = false;
};

/// One `<numFmt>` entry inside `<numFmts>`. The format string itself is
/// stored once in `StylesTable::num_fmt_strings`; this record only
/// carries the id and an index into that vector.
struct NumFmtRecord {
  std::uint16_t id = 0;
  std::uint32_t format_string_index = 0;
};

/// One `<xf>` entry inside `<cellXfs>`. Indexes into the parallel
/// `fonts` / `fills` / `borders` vectors (`0` for the default record).
/// `num_fmt_id` references either a built-in (`0..163`) or a custom
/// entry registered in `num_fmts`.
struct CellXf {
  std::uint32_t font_index = 0;
  std::uint32_t fill_index = 0;
  std::uint32_t border_index = 0;
  std::uint16_t num_fmt_id = 0;
  /// 0=general, 1=left, 2=center, 3=right, 4=fill, 5=justify,
  /// 6=centerContinuous, 7=distributed.
  std::uint8_t horizontal_align = 0;
  /// 0=top, 1=center, 2=bottom (default), 3=justify, 4=distributed.
  std::uint8_t vertical_align = 0;
  bool wrap_text = false;
};

/// Flat in-memory representation of the parsed `xl/styles.xml`.
///
/// Empty-document semantics: a `<styleSheet/>` (or any document missing
/// every section) yields a table whose `fonts` / `fills` / `borders`
/// each contain a single default record so `xf_index = 0` always
/// resolves, and whose `cell_xfs` contains a single default `CellXf{}`.
/// Callers may therefore index `cell_xfs[xf_index]` without further
/// bounds-checking once they have validated `xf_index < cell_xfs.size()`.
struct StylesTable {
  std::vector<FontRecord> fonts;
  std::vector<FillRecord> fills;
  std::vector<BorderRecord> borders;
  std::vector<NumFmtRecord> num_fmts;
  /// Interned format strings. Each `NumFmtRecord::format_string_index`
  /// points into this vector.
  std::vector<std::string> num_fmt_strings;
  std::vector<CellXf> cell_xfs;
};

/// Parses an OOXML styles part.
///
/// Behaviour:
///   * Empty `<styleSheet/>` (no children) yields a table whose `fonts`,
///     `fills`, `borders`, and `cell_xfs` each contain a single default
///     entry so `xf_index = 0` is always resolvable.
///   * Sections present but empty (`<fonts count="0"/>`) similarly fall
///     back to the single-default-record shape.
///   * Children other than the recognised set (`numFmts`, `fonts`,
///     `fills`, `borders`, `cellXfs`) are accepted but ignored —
///     forward compatibility for `<dxfs>`, `<tableStyles>`, `<extLst>`,
///     `<cellStyleXfs>`, and so on.
///   * Unknown attributes inside a recognised element are ignored.
///
/// Errors:
///   * `kIoXmlParse` — pugixml could not parse the document.
///   * `kIoContentTypeInvalid` — the document parses but its root is
///     not `<styleSheet>`.
Expected<StylesTable, Error> read_styles(const std::vector<std::uint8_t>& styles_bytes);

/// Returns the format string for a built-in Excel number-format id
/// (0..163). For ids outside the documented range returns an empty
/// string. The returned pointer is a static `.rodata` view with
/// program lifetime; the canonical source for this table lives in one
/// translation unit (the styles writer) and is exposed here so reader
/// and writer share a single definition.
const char* builtin_num_fmt(std::uint16_t id);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_STYLES_READER_H_
