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

#ifndef FORMULON_IO_STYLES_READER_H_
#define FORMULON_IO_STYLES_READER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// Original specification of an OOXML `<color>` element, preserved so a
/// theme / indexed / auto colour survives a read-modify-write cycle
/// instead of collapsing to a resolved RGB. The parallel `*_argb` field
/// on the owning record still carries the resolved AARRGGBB value that
/// the evaluator and bindings consume; this struct exists purely for
/// verbatim OOXML round-tripping.
struct ColorSpec {
  /// How the source `<color>` element expressed its value.
  enum class Kind : std::uint8_t {
    kNone = 0,  ///< No `<color>` element, or none of the recognised attributes.
    kRgb,       ///< `rgb="AARRGGBB"` (or 6-hex `RRGGBB`).
    kTheme,     ///< `theme="N"` with optional `tint`.
    kIndexed,   ///< `indexed="N"` legacy palette index.
    kAuto,      ///< `auto="1"` system foreground / background.
  };
  Kind kind = Kind::kNone;
  std::uint32_t rgb = 0;      ///< AARRGGBB when `kind == kRgb`.
  std::uint32_t theme = 0;    ///< Theme index when `kind == kTheme`.
  double tint = 0.0;          ///< Theme tint (-1..1) when `kind == kTheme`.
  std::uint32_t indexed = 0;  ///< Palette index when `kind == kIndexed`.
};

/// One `<font>` entry inside `<fonts>`.
///
/// `name` is intentionally a `std::string` because the font-name domain
/// is open-set (system fonts, custom fonts, locale-specific variants);
/// the per-record overhead is acceptable. `color_argb` carries the raw
/// AARRGGBB packed colour. The sentinel `0xFF000000U` is preserved
/// verbatim and treated as "automatic" by the consumer. `color`
/// preserves the original `<color>` specification (theme / indexed /
/// auto) for round-tripping; `color_argb` remains the resolved value.
struct FontRecord {
  std::string name;
  double size = 11.0;
  /// Presence bits preserve an explicit `val="0"` in differential fonts,
  /// where absence means "leave the source formatting unchanged".
  bool has_bold = false;
  bool bold = false;
  bool has_italic = false;
  bool italic = false;
  bool has_strike = false;
  bool strike = false;
  /// 0=none, 1=single, 2=double, 3=singleAccounting, 4=doubleAccounting.
  std::uint8_t underline = 0;
  /// `<vertAlign>` run: 0=baseline (no element), 1=superscript,
  /// 2=subscript. Baseline emits no element on write.
  std::uint8_t vert_align = 0;
  /// `<family>` font-family class (0..5). `has_family` distinguishes an
  /// explicit `<family val="0"/>` from an absent element.
  bool has_family = false;
  std::uint8_t family = 0;
  /// `<charset>` codepage id (e.g. 128 = Shift_JIS, relevant to ja-JP
  /// font substitution). `has_charset` distinguishes an explicit value
  /// from an absent element.
  bool has_charset = false;
  std::uint8_t charset = 0;
  std::uint32_t color_argb = 0xFF000000U;
  ColorSpec color;
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
  /// Original `<fgColor>` / `<bgColor>` specifications (theme / indexed /
  /// auto), preserved for round-tripping. The `*_argb` fields remain the
  /// resolved values.
  ColorSpec fg;
  ColorSpec bg;
};

/// One side of a `<border>` (left/right/top/bottom/diagonal).
struct BorderSide {
  /// 0=none, 1=thin, 2=medium, 3=dashed, ..., 13=slantDashDot
  /// (OOXML border style ordinal).
  std::uint8_t style = 0;
  std::uint32_t color_argb = 0;
  /// Original `<color>` specification (theme / indexed / auto), preserved
  /// for round-tripping. `color_argb` remains the resolved value.
  ColorSpec color;
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
  /// Parent named-style record in `<cellStyleXfs>` for a `<cellXfs>` entry.
  std::uint32_t xf_id = 0;
  bool apply_number_format = false;
  bool apply_font = false;
  bool apply_fill = false;
  bool apply_border = false;
  bool apply_alignment = false;
  bool apply_protection = false;
  bool quote_prefix = false;
  /// 0=general, 1=left, 2=center, 3=right, 4=fill, 5=justify,
  /// 6=centerContinuous, 7=distributed.
  std::uint8_t horizontal_align = 0;
  /// 0=top, 1=center, 2=bottom (default), 3=justify, 4=distributed.
  std::uint8_t vertical_align = 0;
  bool wrap_text = false;
  /// OOXML `justifyLastLine`; retained verbatim so imports round-trip.
  bool justify_last_line = false;
  /// Whether a `<protection>` child was present on the source `<xf>`.
  /// The OOXML schema default is `locked=true`, `hidden=false`; a common
  /// form ("protected sheet, a few input cells unlocked") sets
  /// `locked="0"` on those cells. Preserving the element verbatim keeps
  /// that form from collapsing back to all-locked on save. On write the
  /// `<protection>` element is emitted only when `has_protection` is set.
  bool has_protection = false;
  bool locked = true;
  bool hidden = false;
};

/// One `<dxf>` differential-format record. Unlike `<xf>`, a dxf stores
/// inline optional style fragments instead of indexes into the global
/// font/fill/border tables.
struct DifferentialFormat {
  bool has_font = false;
  FontRecord font;
  bool has_fill = false;
  FillRecord fill;
  bool has_border = false;
  BorderRecord border;
  bool has_num_fmt = false;
  std::uint16_t num_fmt_id = 0;
  std::string num_fmt_code;
  /// Raw `<alignment>` / `<protection>` child elements, captured verbatim
  /// so a dxf that carries them (rare in conditional formats, but valid)
  /// survives a read -> write round trip. Empty when the child is absent.
  /// The writer re-emits them at their CT_Dxf positions (alignment after
  /// fill, protection after border).
  std::string alignment_xml;
  std::string protection_xml;
};

/// One `<cellStyle>` entry inside `<cellStyles>`. Defines a named cell
/// style ("Normal", "Heading 1", custom user names, etc.) that points
/// at a record in the parallel `<cellStyleXfs>` table via `xf_id`.
///
/// `builtin_id` is the OOXML built-in style ordinal (`0..47`); the
/// sentinel `kBuiltinIdNone` indicates the style is custom and the
/// attribute should be omitted on write. `i_level` is the outline level
/// for built-in heading styles (0 for everything else).
struct CellStyleRecord {
  /// Sentinel for "no `builtinId` attribute" — picked above the OOXML
  /// 0..47 built-in range so it cannot collide with a real id.
  static constexpr std::uint32_t kBuiltinIdNone = 0xFFFFFFFFU;

  std::string name;
  std::uint32_t xf_id = 0;
  std::uint32_t builtin_id = kBuiltinIdNone;
  std::uint32_t i_level = 0;
  bool hidden = false;
  bool custom_builtin = false;
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
  /// `<cellStyleXfs>` records — the named-style xf table. Empty when
  /// the workbook does not use named cell styles. Same shape as
  /// `cell_xfs`; the two tables are kept distinct because OOXML
  /// indexes them independently.
  std::vector<CellXf> cell_style_xfs;
  /// `<cellStyles>` records — named cell styles ("Normal", "Heading 1",
  /// custom user styles). Each entry's `xf_id` indexes into
  /// `cell_style_xfs`. Empty when the workbook does not declare any
  /// named cell styles.
  std::vector<CellStyleRecord> cell_styles;
  /// `<dxfs>` records referenced by conditional-format `dxfId`.
  std::vector<DifferentialFormat> dxfs;
  /// Namespace declarations and compatibility attributes from the
  /// `<styleSheet>` root that raw extension fragments depend on.
  std::string root_extra_attrs;
  /// Unmodelled top-level style sections retained verbatim. Their schema
  /// positions are fixed by `write_styles`, so a style-table edit does not
  /// discard custom palettes, table/pivot style defaults, or extensions.
  std::string colors_xml;
  std::string table_styles_xml;
  std::string ext_lst_xml;
  /// Complete top-level `<styleSheet>` children which the style model does
  /// not recognise. Preserved in source order to avoid silent deletion of
  /// extension data during a read-modify-write cycle.
  std::vector<std::string> unknown_top_level_xml;
};

/// Parses an OOXML styles part.
///
/// Behaviour:
///   * Empty `<styleSheet/>` (no children) yields a table whose `fonts`,
///     `fills`, `borders`, and `cell_xfs` each contain a single default
///     entry so `xf_index = 0` is always resolvable.
///   * Sections present but empty (`<fonts count="0"/>`) similarly fall
///     back to the single-default-record shape.
///   * `<colors>`, `<tableStyles>`, and `<extLst>` are retained as raw XML;
///     other unrecognised children are accepted but ignored.
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
