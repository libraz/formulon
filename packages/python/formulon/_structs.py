"""WASM32 struct marshalling for the Formulon C ABI.

Emscripten lowers every by-value struct parameter (and every struct
out-parameter) of the C ABI to a pointer into linear memory. The Python
wrapper therefore allocates a scratch block, writes each field at its
``wasm32`` offset, and passes the pointer. This module centralises those
layouts so the offsets live in exactly one place.

Layout rules (wasm32 / ILP32):
  * pointers are 4 bytes, 4-aligned;
  * ``int32`` / ``uint32`` are 4 bytes, 4-aligned;
  * ``uint16`` is 2 bytes, 2-aligned;
  * ``uint8`` is 1 byte, 1-aligned;
  * ``int64`` is 8 bytes, 8-aligned;
  * ``double`` is 8 bytes, 8-aligned;
  * a struct's size is rounded up to its largest member alignment.

Each :class:`Struct` computes offsets once at import time. ``pack`` writes
a freshly-zeroed block; ``unpack`` reads every scalar field back into a
plain dict. Pointer fields are surfaced as raw integer offsets -- the
caller decodes C strings via ``LIB.read_cstr`` and frees buffers it owns.
"""

from __future__ import annotations

import struct
from typing import Dict, List, Optional, Tuple

# (size, alignment) for every primitive the C ABI structs use.
PTR = ("ptr", 4, 4)
U32 = ("u32", 4, 4)
I32 = ("i32", 4, 4)
I64 = ("i64", 8, 8)
U16 = ("u16", 2, 2)
U8 = ("u8", 1, 1)
F64 = ("f64", 8, 8)
# Opaque 16-byte, 8-aligned blob: an inline ``fm_value_t``. It is not
# decoded by :meth:`Struct.unpack`; callers read it via Value._from_wasm
# against ``ptr + offset``.
VALUE_BLOB = ("blob16", 16, 8)
# Opaque 12-byte, 4-aligned blob: an inline ``fm_cfvo_t`` (``uint8_t type;
# uint8_t _pad[3]; int32_t gte; const char* value;``). Not decoded by
# :meth:`Struct.unpack`; callers who need the sub-fields read them at
# ``ptr + offset`` with their own struct.unpack_from calls.
CFVO_BLOB = ("blob_cfvo", 12, 4)
# Opaque 16-byte, 4-aligned blob: an inline ``fm_merge_range``. Pack and
# unpack it with :data:`MERGE_RANGE` against ``ptr + offset``.
MERGE_RANGE_BLOB = ("blob_merge_range", 16, 4)
# Opaque 4-byte, 1-aligned blob: an inline ``fm_cf_color_t`` (four
# ``uint8_t`` channels). Not decoded by :meth:`Struct.unpack`.
CF_COLOR_BLOB = ("blob_cf_color", 4, 1)

_FMT = {
    "ptr": "<I",
    "u32": "<I",
    "i32": "<i",
    "i64": "<q",
    "u16": "<H",
    "u8": "<B",
    "f64": "<d",
}


def nested(layout: "Struct") -> Tuple:
    """Field type for a struct embedded by value in another struct."""
    return ("nested", layout.size, layout.align, layout)


def array(elem: Tuple, count: int) -> Tuple:
    """Field type for an inline ``elem[count]`` array of scalars."""
    return ("array", elem[1] * count, elem[2], (_FMT[elem[0]], count))


class Struct:
    """A wasm32 POD layout: ordered (name, ctype) fields with offsets."""

    __slots__ = ("name", "fields", "offsets", "size", "align", "_extra")

    def __init__(self, name: str, fields: List[Tuple[str, Tuple]]) -> None:
        self.name = name
        self.fields = fields
        self.offsets: Dict[str, Tuple[str, int]] = {}
        self._extra: Dict[str, object] = {}
        off = 0
        max_align = 1
        for fname, ctype in fields:
            kind, sz, align = ctype[:3]
            off = (off + align - 1) // align * align
            self.offsets[fname] = (kind, off)
            if len(ctype) > 3:
                self._extra[fname] = ctype[3]
            off += sz
            max_align = max(max_align, align)
        self.align = max_align
        self.size = (off + max_align - 1) // max_align * max_align

    def pack_into(self, raw: bytearray, base: int, values: Dict[str, object]) -> None:
        """Write ``values`` into ``raw`` at ``base``; a nested struct takes a dict."""
        for fname, value in values.items():
            kind, off = self.offsets[fname]
            if kind == "nested":
                self._extra[fname].pack_into(raw, base + off, value)
            else:
                struct.pack_into(_FMT[kind], raw, base + off, value)

    def pack(self, lib, ptr: int, values: Dict[str, object]) -> None:
        """Pack supplied fields into a zeroed block and write it once."""
        raw = bytearray(self.size)
        self.pack_into(raw, 0, values)
        lib.write_bytes(ptr, raw)

    def unpack_from(self, raw: bytes, base: int = 0) -> Dict[str, object]:
        """Decode every scalar, nested-struct and array field of ``raw`` at ``base``."""
        out: Dict[str, object] = {}
        for fname, (kind, off) in self.offsets.items():
            if kind == "nested":
                out[fname] = self._extra[fname].unpack_from(raw, base + off)
            elif kind == "array":
                fmt, count = self._extra[fname]
                out[fname] = list(struct.unpack_from("<" + fmt[1] * count, raw, base + off))
            elif kind in _FMT:
                out[fname] = struct.unpack_from(_FMT[kind], raw, base + off)[0]
        return out

    def unpack(self, lib, ptr: int) -> Dict[str, int]:
        """Read every scalar field from ``[ptr, ptr+size)`` into a dict.

        Opaque blob fields (any ``kind`` not in ``_FMT``, e.g. an inline
        ``fm_value_t`` or ``fm_cfvo_t``) are skipped; callers decode those
        separately against ``ptr + self.offsets[name][1]``. Nested structs
        and inline arrays decode to a dict and a list.
        """
        return self.unpack_from(lib.read_bytes(ptr, self.size))


def zero_struct(lib, layout: Struct, ptr: int) -> None:
    """Re-zero an already-allocated scratch block sized for ``layout``.

    Lets a collection reader hoist one scratch block out of its loop and
    still hand each ABI call a cleared block, without paying a WASM
    ``malloc``/``free`` pair per item.
    """
    lib.write_bytes(ptr, b"\x00" * layout.size)


# ---------------------------------------------------------------------------
# Struct layouts (mirror src/c_api/formulon_c.h field-for-field)
# ---------------------------------------------------------------------------

MERGE_RANGE = Struct(
    "fm_merge_range",
    [("first_row", U32), ("first_col", U32), ("last_row", U32), ("last_col", U32)],
)

# fm_cf_cell_range_t shares fm_merge_range's shape.
CF_CELL_RANGE = Struct(
    "fm_cf_cell_range_t",
    [("first_row", U32), ("first_col", U32), ("last_row", U32), ("last_col", U32)],
)

HYPERLINK = Struct(
    "fm_hyperlink",
    [
        ("row", U32),
        ("col", U32),
        ("last_row", U32),
        ("last_col", U32),
        ("target", PTR),
        ("location", PTR),
        ("display", PTR),
        ("tooltip", PTR),
    ],
)

COMMENT = Struct(
    "fm_comment",
    [("row", U32), ("col", U32), ("author", PTR), ("text", PTR)],
)

WORKBOOK_PROTECTION = Struct(
    "fm_workbook_protection_t",
    [
        ("lock_structure", I32),
        ("lock_windows", I32),
        ("lock_revision", I32),
        ("workbook_spin_count", U32),
        ("revisions_spin_count", U32),
        ("workbook_password", PTR),
        ("workbook_algorithm_name", PTR),
        ("workbook_hash_value", PTR),
        ("workbook_salt_value", PTR),
        ("revisions_password", PTR),
        ("revisions_algorithm_name", PTR),
        ("revisions_hash_value", PTR),
        ("revisions_salt_value", PTR),
    ],
)

PROTECTED_RANGE = Struct(
    "fm_protected_range_t",
    [
        ("spin_count", U32),
        ("name", PTR),
        ("sqref", PTR),
        ("password", PTR),
        ("algorithm_name", PTR),
        ("hash_value", PTR),
        ("salt_value", PTR),
        ("security_descriptor", PTR),
    ],
)

DATA_VALIDATION = Struct(
    "fm_data_validation",
    [
        ("ranges", PTR),
        ("range_count", U32),
        ("type", U8),
        ("op", U8),
        ("error_style", U8),
        ("allow_blank", I32),
        ("show_input_message", I32),
        ("show_error_message", I32),
        ("show_dropdown", I32),
        ("formula1", PTR),
        ("formula2", PTR),
        ("error_title", PTR),
        ("error_message", PTR),
        ("prompt_title", PTR),
        ("prompt_message", PTR),
    ],
)

SHEET_PROTECTION = Struct(
    "fm_sheet_protection_t",
    [
        ("enabled", I32),
        ("algorithm_name", PTR),
        ("hash_value", PTR),
        ("salt_value", PTR),
        ("spin_count", U32),
        ("legacy_password", PTR),
        ("sheet", I32),
        ("objects", I32),
        ("scenarios", I32),
        ("format_cells", I32),
        ("format_columns", I32),
        ("format_rows", I32),
        ("insert_columns", I32),
        ("insert_rows", I32),
        ("insert_hyperlinks", I32),
        ("delete_columns", I32),
        ("delete_rows", I32),
        ("select_locked_cells", I32),
        ("select_unlocked_cells", I32),
        ("sort", I32),
        ("auto_filter", I32),
        ("pivot_tables", I32),
    ],
)

VIEWPORT = Struct(
    "fm_viewport",
    [
        ("sheet", U32),
        ("first_row", U32),
        ("last_row", U32),
        ("first_col", U32),
        ("last_col", U32),
    ],
)

CF_COLOR = Struct(
    "fm_cf_color_t",
    [("r", U8), ("g", U8), ("b", U8), ("a", U8)],
)

CFVO = Struct(
    "fm_cfvo_t",
    [("type", U8), ("_pad", ("blob_pad3", 3, 1)), ("gte", I32), ("value", PTR)],
)

CF_MATCH = Struct(
    "fm_cf_match_t",
    [
        ("kind", I32),
        ("priority", I32),
        ("dxf_id_engaged", I32),
        ("dxf_id", U32),
        ("color_r", U8),
        ("color_g", U8),
        ("color_b", U8),
        ("color_a", U8),
        ("bar_length_pct", F64),
        ("bar_axis_position_pct", F64),
        ("bar_is_negative", I32),
        ("bar_fill_r", U8),
        ("bar_fill_g", U8),
        ("bar_fill_b", U8),
        ("bar_fill_a", U8),
        ("bar_border_engaged", I32),
        ("bar_border_r", U8),
        ("bar_border_g", U8),
        ("bar_border_b", U8),
        ("bar_border_a", U8),
        ("bar_gradient", I32),
        ("icon_set_name", I32),
        ("icon_index", U8),
        ("bar_direction", U8),
    ],
)

CF_RULE = Struct(
    "fm_cf_rule_t",
    [
        ("id", PTR),
        ("type", U8),
        ("op", U8),
        ("time_period", U8),
        ("_pad0", U8),
        ("priority", I32),
        ("stop_if_true", I32),
        ("dxf_id_engaged", I32),
        ("dxf_id", U32),
        ("sqref", PTR),
        ("sqref_count", U32),
        ("formula1", PTR),
        ("formula2", PTR),
        ("op_engaged", I32),
        ("rank_engaged", I32),
        ("rank", I32),
        ("percent", I32),
        ("bottom", I32),
        ("above_average", I32),
        ("equal_average", I32),
        ("std_dev_engaged", I32),
        ("std_dev", F64),
        ("text", PTR),
        ("time_period_engaged", I32),
        ("color_scale_thresholds", PTR),
        ("color_scale_colors", PTR),
        ("color_scale_count", U32),
        ("data_bar_engaged", I32),
        ("data_bar_min", CFVO_BLOB),
        ("data_bar_max", CFVO_BLOB),
        ("data_bar_fill", CF_COLOR_BLOB),
        ("data_bar_show_value", I32),
        ("data_bar_min_length_pct", U8),
        ("data_bar_max_length_pct", U8),
        ("icon_set_engaged", I32),
        ("icon_set_name", U8),
        ("icon_set_thresholds", PTR),
        ("icon_set_threshold_count", U32),
        ("icon_set_reverse", I32),
        ("icon_set_show_value", I32),
        ("icon_set_percent", I32),
        # Data-bar attributes carried as engaged-flag / value pairs: an
        # unengaged flag means "use the model default" rather than "zero".
        ("data_bar_gradient_engaged", I32),
        ("data_bar_gradient", I32),
        ("data_bar_axis_position_engaged", I32),
        ("data_bar_axis_position", U8),
        ("data_bar_negative_fill_engaged", I32),
        ("data_bar_negative_fill", CF_COLOR_BLOB),
        ("data_bar_border_engaged", I32),
        ("data_bar_border", CF_COLOR_BLOB),
        ("data_bar_negative_border_engaged", I32),
        ("data_bar_negative_border", CF_COLOR_BLOB),
        ("data_bar_axis_color_engaged", I32),
        ("data_bar_axis_color", CF_COLOR_BLOB),
        ("icon_set_floor_engaged", I32),
        ("icon_set_floor", CFVO_BLOB),
        ("data_bar_direction", U8),
        ("icon_set_icons", PTR),
        ("icon_set_icon_count", U32),
    ],
)

CF_ICON = Struct("fm_cf_icon_t", [("icon_set", PTR), ("icon_id", I32)])

CELL_NODE = Struct(
    "fm_cell_node_t",
    [("sheet", U32), ("row", U32), ("col", U32)],
)

# One ``<rPh>`` block. Passed both ways: as an element of the array
# ``fm_workbook_set_cell_phonetic_runs`` takes, and as the out-parameter of
# ``fm_workbook_get_cell_phonetic_run``.
PHONETIC_RUN = Struct(
    "fm_phonetic_run_t",
    [("sb", U32), ("eb", U32), ("text", PTR)],
)

# fm_pivot_cell_t embeds an fm_value_t (16 bytes, 8-aligned) inline. The
# ``value`` field is an opaque blob; decode it with Value._from_wasm
# against ``cell_ptr + PIVOT_CELL_VALUE_OFFSET``.
PIVOT_CELL = Struct(
    "fm_pivot_cell_t",
    [
        ("row", U32),
        ("col", U32),
        ("value", VALUE_BLOB),
        ("kind", I32),
        ("depth", U32),
        ("field_name", PTR),
        ("number_format", PTR),
    ],
)
PIVOT_CELL_VALUE_OFFSET = PIVOT_CELL.offsets["value"][1]

PIVOT_FIELD_SPEC = Struct(
    "fm_pivot_field_spec_t",
    [
        ("source_name", PTR),
        ("custom_name", PTR),
        ("axis", I32),
        ("subtotal_top", I32),
        ("number_format", PTR),
    ],
)

PIVOT_DATA_FIELD_SPEC = Struct(
    "fm_pivot_data_field_spec_t",
    [
        ("name", PTR),
        ("field_index", U32),
        ("aggregation", I32),
        ("number_format", PTR),
        ("show_as", I32),
        ("show_as_base_field", I32),
        ("show_as_base_item", I32),
    ],
)

PIVOT_FIELD_INFO = Struct(
    "fm_pivot_field_info_t",
    [
        ("source_name", PTR),
        ("custom_name", PTR),
        ("axis", I32),
        ("subtotal_top", I32),
        ("number_format", PTR),
        ("cache_field_index", I32),
        ("sort_order", I32),
        ("sort_by_field", PTR),
        ("has_date_group", I32),
        ("date_granularity", I32),
        ("date_calendar", I32),
        ("date_start_year", I32),
        ("date_end_year", I32),
        ("date_interval_days", U32),
        ("date_start_serial", F64),
        ("date_end_serial", F64),
        ("default_subtotal", I32),
        ("subtotal_fns", U32),
    ],
)

PIVOT_ITEM_INFO = Struct(
    "fm_pivot_item_info_t",
    [("name", PTR), ("cache_index", I32), ("visible", I32)],
)

TABLE_INFO = Struct(
    "fm_table_info_t",
    [
        ("style_name", PTR),
        ("style_present", I32),
        ("show_first_column", I32),
        ("show_last_column", I32),
        ("show_row_stripes", I32),
        ("show_column_stripes", I32),
        ("header_row_count", U32),
        ("totals_row_count", U32),
        ("column_count", U32),
    ],
)

TABLE_COLUMN_INFO = Struct(
    "fm_table_column_info_t",
    [
        ("name", PTR),
        ("totals_row_function", PTR),
        ("totals_row_label", PTR),
        ("totals_row_formula", PTR),
    ],
)

TABLE_STYLE_ELEMENT = Struct(
    "fm_table_style_element_t",
    [("type", PTR), ("size", U32), ("has_dxf_id", I32), ("dxf_id", U32)],
)

TABLE_STYLE = Struct(
    "fm_table_style_t",
    [("name", PTR), ("pivot", I32), ("table", I32), ("elements", PTR), ("element_count", U32)],
)

PIVOT_STYLE_INFO = Struct(
    "fm_pivot_style_info_t",
    [
        ("name", PTR),
        ("present", I32),
        ("show_row_headers", I32),
        ("show_col_headers", I32),
        ("show_row_stripes", I32),
        ("show_col_stripes", I32),
        ("show_last_column", I32),
    ],
)

PIVOT_FILTER_SPEC = Struct(
    "fm_pivot_filter_spec_t",
    [
        ("axis", I32),
        ("field_name", PTR),
        ("type", I32),
        ("value_kind", I32),
        ("value_int", I32),
        ("value_double", F64),
        ("value_text", PTR),
        ("value_high_kind", I32),
        ("value_high_int", I32),
        ("value_high_double", F64),
        ("data_field_index", U32),
    ],
)

READ_DIAGNOSTICS = Struct(
    "fm_read_diagnostics_t",
    [
        ("undecoded_formula_count", U32),
        ("undecoded_defined_name_count", U32),
        ("undecoded_part_count", U32),
        ("skipped_feature_count", U32),
        ("unknown_content_type_count", U32),
    ],
)

SAVE_DIAGNOSTICS = Struct(
    "fm_save_diagnostics_t",
    [
        ("downgraded_formula_count", U32),
        ("deferred_feature_count", U32),
        ("dropped_part_count", U32),
        ("dropped_relationship_count", U32),
        ("renumbered_part_count", U32),
    ],
)

SPILL_INFO = Struct(
    "fm_spill_info_t",
    [
        ("anchor_row", U32),
        ("anchor_col", U32),
        ("rows", U32),
        ("cols", U32),
        ("engaged", I32),
    ],
)

FUNCTION_METADATA = Struct(
    "fm_function_metadata_t",
    [
        ("canonical_name", PTR),
        ("min_arity", U32),
        ("max_arity", U32),
        ("availability", I32),
        ("signature_template", PTR),
        ("description", PTR),
    ],
)

LOCALE_FACTS = Struct(
    "fm_locale_facts_t",
    [
        ("decimal_separator", PTR),
        ("group_separator", PTR),
        ("list_separator", PTR),
        ("array_column_separator", PTR),
        ("array_row_separator", PTR),
        ("true_name", PTR),
        ("false_name", PTR),
        ("date_order", I32),
        ("currency_symbol", PTR),
        ("currency_suffix", I32),
        ("currency_space", I32),
        ("currency_default_decimals", I32),
        ("measured", I32),
        ("collation_measured", I32),
    ],
)

CIVIL_TIME = Struct(
    "fm_civil_time_t",
    [
        ("year", I32),
        ("month", I32),
        ("day", I32),
        ("hour", I32),
        ("minute", I32),
        ("second", I32),
    ],
)

SHEET_VIEW = Struct(
    "fm_sheet_view_t",
    [
        ("zoom_scale", U32),
        ("freeze_rows", U32),
        ("freeze_cols", U32),
        ("tab_hidden", I32),
        ("visibility", I32),
        ("show_grid_lines", I32),
        ("show_row_col_headers", I32),
        ("show_zeros", I32),
        ("right_to_left", I32),
        ("tab_selected", I32),
        ("view_mode", PTR),
        ("top_left_row", U32),
        ("top_left_col", U32),
        ("pane_top_left_row", U32),
        ("pane_top_left_col", U32),
        ("split_x", F64),
        ("split_y", F64),
        ("show_formulas", I32),
        ("show_outline_symbols", I32),
    ],
)

COLUMN_LAYOUT = Struct(
    "fm_column_layout_t",
    [
        ("first", U32),
        ("last", U32),
        ("width", F64),
        ("hidden", I32),
        ("outline_level", U8),
        ("collapsed", U8),
        ("has_width", I32),
        ("has_style", I32),
        ("style_xf", U32),
    ],
)

ROW_LAYOUT = Struct(
    "fm_row_layout_t",
    [
        ("row", U32),
        ("height", F64),
        ("hidden", I32),
        ("outline_level", U8),
        ("collapsed", U8),
        ("has_style", I32),
        ("style_xf", U32),
        ("has_height", I32),
        ("custom_height", I32),
    ],
)

SHEET_FORMAT_DEFAULTS = Struct(
    "fm_sheet_format_defaults",
    [
        ("default_col_width", F64),
        ("default_row_height", F64),
        ("base_col_width", F64),
        ("has_default_col_width", I32),
        ("has_default_row_height", I32),
        ("custom_height", I32),
        ("zero_height", I32),
        ("effective_default_col_width", F64),
        ("effective_default_row_height", F64),
    ],
)

RECT_PT = Struct(
    "fm_rect_pt",
    [("x", F64), ("y", F64), ("width", F64), ("height", F64)],
)

WIDTH_MODEL = Struct(
    "fm_width_model",
    [
        ("points_per_char", F64),
        ("padding_pt", F64),
        ("normal_font_size", F64),
        ("calibrated", I32),
        ("_pad", I32),
        ("normal_font_name", PTR),
        ("platform", PTR),
    ],
)

PAPER_INFO = Struct(
    "fm_paper_info",
    [("width_pt", F64), ("height_pt", F64), ("landscape", I32), ("known", I32)],
)

MARGINS_PT = Struct(
    "fm_margins_pt",
    [
        ("left", F64),
        ("right", F64),
        ("top", F64),
        ("bottom", F64),
        ("header", F64),
        ("footer", F64),
    ],
)

PRINT_TITLES = Struct(
    "fm_print_titles",
    [
        ("has_rows", I32),
        ("first_row", U32),
        ("last_row", U32),
        ("has_cols", I32),
        ("first_col", U32),
        ("last_col", U32),
    ],
)

PAGE_LAYOUT = Struct(
    "fm_page_layout",
    [
        ("area_index", U32),
        ("first_row", U32),
        ("last_row", U32),
        ("first_col", U32),
        ("last_col", U32),
        ("origin_x_pt", F64),
        ("origin_y_pt", F64),
        ("width_pt", F64),
        ("height_pt", F64),
    ],
)

# Every optional alignment attribute has an explicit presence flag so zero /
# false values can be round-tripped.
CELL_XF = Struct(
    "fm_cell_xf",
    [
        ("font_index", U32),
        ("fill_index", U32),
        ("border_index", U32),
        ("num_fmt_id", U16),
        ("horizontal_align", U8),
        ("vertical_align", U8),
        ("wrap_text", I32),
        ("justify_last_line", I32),
        ("xf_id", U32),
        ("has_alignment", I32),
        ("has_text_rotation", I32),
        ("text_rotation", U32),
        ("has_indent", I32),
        ("indent", U32),
        ("has_relative_indent", I32),
        ("relative_indent", I32),
        ("has_shrink_to_fit", I32),
        ("shrink_to_fit", I32),
        ("has_reading_order", I32),
        ("reading_order", U32),
        ("has_horizontal_align", I32),
        ("has_vertical_align", I32),
        ("has_wrap_text", I32),
        ("has_justify_last_line", I32),
        ("apply_number_format", I32),
        ("apply_font", I32),
        ("apply_fill", I32),
        ("apply_border", I32),
        ("apply_alignment", I32),
        ("apply_protection", I32),
        ("quote_prefix", I32),
        ("has_protection", I32),
        ("locked", I32),
        ("hidden", I32),
    ],
)

COLOR_SPEC = Struct(
    "fm_color_spec",
    [("tint", F64), ("rgb", U32), ("theme", U32), ("indexed", U32), ("kind", U8)],
)
# Inline ``fm_color_spec``. Nested rather than flattened into its owner: it
# is 8-aligned, so its trailing padding is part of the C layout and a flat
# field list would place the following member too early.
COLOR_SPEC_BLOB = ("blob_color_spec", COLOR_SPEC.size, 8)

FONT_RECORD = Struct(
    "fm_font_record",
    [
        ("name", PTR),
        ("size", F64),
        ("color_argb", U32),
        ("bold", I32),
        ("italic", I32),
        ("strike", I32),
        ("has_bold", I32),
        ("has_italic", I32),
        ("has_strike", I32),
        ("has_family", I32),
        ("has_charset", I32),
        ("has_size", I32),
        ("outline", I32),
        ("shadow", I32),
        ("condense", I32),
        ("extend", I32),
        ("has_outline", I32),
        ("has_shadow", I32),
        ("has_condense", I32),
        ("has_extend", I32),
        ("underline", U8),
        ("vert_align", U8),
        ("family", U8),
        ("charset", U8),
        ("scheme", U8),
        ("color", COLOR_SPEC_BLOB),
    ],
)

GRADIENT_STOP = Struct(
    "fm_gradient_stop",
    [("position", F64), ("color_argb", U32), ("color", COLOR_SPEC_BLOB)],
)

FILL_RECORD = Struct(
    "fm_fill_record",
    [
        ("pattern", U8),
        ("fg_argb", U32),
        ("bg_argb", U32),
        ("fg", COLOR_SPEC_BLOB),
        ("bg", COLOR_SPEC_BLOB),
        ("has_gradient", I32),
        ("gradient_type", I32),
        ("gradient_degree", F64),
        ("gradient_left", F64),
        ("gradient_right", F64),
        ("gradient_top", F64),
        ("gradient_bottom", F64),
        ("gradient_stops", PTR),
        # ``size_t`` is 4 bytes on wasm32, so the count shares PTR's layout.
        ("gradient_stop_count", PTR),
    ],
)

BORDER_SIDE = Struct(
    "fm_border_side",
    [("style", U8), ("color_argb", U32), ("color", COLOR_SPEC_BLOB)],
)
BORDER_SIDE_BLOB = ("blob_border_side", BORDER_SIDE.size, 8)

BORDER_RECORD = Struct(
    "fm_border_record",
    [
        ("left", BORDER_SIDE_BLOB),
        ("right", BORDER_SIDE_BLOB),
        ("top", BORDER_SIDE_BLOB),
        ("bottom", BORDER_SIDE_BLOB),
        ("diagonal", BORDER_SIDE_BLOB),
        ("diagonal_up", I32),
        ("diagonal_down", I32),
    ],
)

# Inline members of ``fm_dxf_record``. They deliberately remain opaque here:
# callers marshal them through the standalone layouts above at their recorded
# offsets.
FONT_RECORD_BLOB = ("blob_font_record", FONT_RECORD.size, 8)
FILL_RECORD_BLOB = ("blob_fill_record", FILL_RECORD.size, 8)
BORDER_RECORD_BLOB = ("blob_border_record", BORDER_RECORD.size, 8)

DXF_RECORD = Struct(
    "fm_dxf_record",
    [
        ("font_engaged", I32),
        ("font", FONT_RECORD_BLOB),
        ("fill_engaged", I32),
        ("fill", FILL_RECORD_BLOB),
        ("border_engaged", I32),
        ("border", BORDER_RECORD_BLOB),
        ("num_fmt_engaged", I32),
        ("num_fmt_id", U16),
        ("num_fmt_code", PTR),
        ("alignment_xml", PTR),
        ("protection_xml", PTR),
    ],
)

# ``size_t`` is 4 bytes on wasm32, so the count fields share PTR's layout.
STYLES_BATCH = Struct(
    "fm_styles_batch",
    [
        ("fonts", PTR),
        ("font_count", PTR),
        ("font_indices", PTR),
        ("fills", PTR),
        ("fill_count", PTR),
        ("fill_indices", PTR),
        ("borders", PTR),
        ("border_count", PTR),
        ("border_indices", PTR),
        ("cell_xfs", PTR),
        ("cell_xf_count", PTR),
        ("cell_xf_indices", PTR),
        ("num_fmt_codes", PTR),
        ("num_fmt_count", PTR),
        ("num_fmt_ids", PTR),
    ],
)

CELL_STYLE_RECORD = Struct(
    "fm_cell_style_record_t",
    [
        ("name", PTR),
        ("xf_id", U32),
        ("builtin_id", U32),
        ("i_level", U32),
        ("hidden", I32),
        ("custom_builtin", I32),
    ],
)

THEME_COLORS = Struct("fm_theme_colors", [("argb", ("blob_u32x12", 48, 4))])

THEME_FONTS = Struct(
    "fm_theme_fonts",
    [
        ("major_latin", PTR),
        ("major_east_asian", PTR),
        ("minor_latin", PTR),
        ("minor_east_asian", PTR),
    ],
)

EFFECTIVE_STYLE = Struct(
    "fm_effective_style",
    [
        ("xf_index", U32),
        ("source", I32),
        ("font_index", U32),
        ("fill_index", U32),
        ("border_index", U32),
        ("font_argb", U32),
        ("font_resolution", I32),
        ("fill_fg_argb", U32),
        ("fill_fg_resolution", I32),
        ("fill_bg_argb", U32),
        ("fill_bg_resolution", I32),
        ("border_argb", ("blob_u32x5", 20, 4)),
        ("border_resolution", ("blob_i32x5", 20, 4)),
        ("locked", I32),
        ("hidden", I32),
        ("num_fmt_code", PTR),
    ],
)

EXTERNAL_LINK_RECORD = Struct(
    "fm_external_link_record_t",
    [
        ("index", U32),
        ("rel_id", PTR),
        ("part_path", PTR),
        ("target", PTR),
        ("target_external", I32),
        ("kind", U32),
    ],
)

PAGE_BREAK = Struct(
    "fm_page_break_t",
    [
        ("id", U32),
        ("min", U32),
        ("max", U32),
        ("manual", I32),
    ],
)

# The three print patch structs pair every value field with an
# ``*_engaged`` flag: on the write path the flag selects which attributes
# the engine touches, on the read path it reports whether the attribute is
# stated in the XML at all. Field order must match the C declaration
# exactly -- ``Struct`` derives offsets from declaration order, so a
# reordering here silently mis-decodes rather than failing.
PAGE_SETUP = Struct(
    "fm_page_setup_t",
    [
        ("orientation_engaged", I32),
        ("orientation", U32),
        ("paper_size_engaged", I32),
        ("paper_size", U32),
        ("scale_engaged", I32),
        ("scale", U32),
        ("fit_to_width_engaged", I32),
        ("fit_to_width", U32),
        ("fit_to_height_engaged", I32),
        ("fit_to_height", U32),
        ("fit_to_page_engaged", I32),
        ("fit_to_page", I32),
    ],
)

# `double` is 8-aligned, so each `I32` flag is followed by four bytes of
# padding. `Struct` inserts it from the declared alignments.
PAGE_MARGINS = Struct(
    "fm_page_margins_t",
    [
        ("left_engaged", I32),
        ("left", F64),
        ("right_engaged", I32),
        ("right", F64),
        ("top_engaged", I32),
        ("top", F64),
        ("bottom_engaged", I32),
        ("bottom", F64),
        ("header_engaged", I32),
        ("header", F64),
        ("footer_engaged", I32),
        ("footer", F64),
    ],
)

PRINT_OPTIONS = Struct(
    "fm_print_options_t",
    [
        ("grid_lines_engaged", I32),
        ("grid_lines", I32),
        ("headings_engaged", I32),
        ("headings", I32),
        ("horizontal_centered_engaged", I32),
        ("horizontal_centered", I32),
        ("vertical_centered_engaged", I32),
        ("vertical_centered", I32),
    ],
)

HEADER_FOOTER = Struct(
    "fm_header_footer_t",
    [
        ("odd_header", PTR),
        ("odd_footer", PTR),
        ("even_header", PTR),
        ("even_footer", PTR),
        ("first_header", PTR),
        ("first_footer", PTR),
        ("different_odd_even_engaged", I32),
        ("different_odd_even", I32),
        ("different_first_engaged", I32),
        ("different_first", I32),
        ("scale_with_doc_engaged", I32),
        ("scale_with_doc", I32),
        ("align_with_margins_engaged", I32),
        ("align_with_margins", I32),
    ],
)

CELL_STYLE_BUILTIN_ID_NONE = 0xFFFFFFFF


def write_str_field(
    lib,
    struct_ptr: int,
    layout: Struct,
    field: str,
    value: Optional[str],
    scope,
    *,
    allow_empty: bool = False,
) -> None:
    """Encode ``value`` into WASM memory and store its pointer in a field.

    ``scope`` is the caller's ``_ArgScope``: it owns the buffer and does the
    string-to-pointer mapping. ``None`` leaves the field as the NULL pointer
    (already zeroed by :meth:`Struct.pack`). ``""`` does too, unless
    ``allow_empty`` is set -- needed for a field the C ABI itself
    distinguishes NULL from a zero-length string on, such as
    ``fm_phonetic_run_t.text``: a NULL run text is rejected
    (``kBindingNullPointer``), but Excel legitimately produces empty-reading
    runs (see ``src/io/xlsb/sst_reader.cpp``'s out-of-range kana slices), so
    ``get_phonetic_runs()``'s own output could not always round-trip back
    through ``set_phonetic_runs()``.
    """
    ptr = scope.opt_cstr(value) if allow_empty else scope.nonempty_cstr(value)
    if ptr == 0:
        return
    kind, off = layout.offsets[field]
    lib.write_bytes(struct_ptr + off, struct.pack("<I", ptr))


# ---------------------------------------------------------------------------
# AutoFilter, validation outcome, threaded comments and persons
# ---------------------------------------------------------------------------

DATE_GROUP_ITEM = Struct(
    "fm_date_group_item",
    [
        ("year", U16),
        ("month", U8),
        ("day", U8),
        ("hour", U8),
        ("minute", U8),
        ("second", U8),
        ("grouping", U8),
    ],
)

FILTER_COLUMN = Struct(
    "fm_filter_column",
    [
        ("col_id", U32),
        ("hidden_button", I32),
        ("show_button", I32),
        ("kind", I32),
        ("filter_blank", I32),
        ("values", PTR),
        ("value_count", U32),
        ("date_groups", PTR),
        ("date_group_count", U32),
        ("custom_and", I32),
        ("custom_count", I32),
        ("op1", I32),
        ("val1", PTR),
        ("op2", I32),
        ("val2", PTR),
        ("top", I32),
        ("percent", I32),
        ("has_filter_val", I32),
        ("top_val", F64),
        ("filter_val", F64),
        ("dynamic_type", I32),
        ("has_dyn_val", I32),
        ("has_dyn_max_val", I32),
        ("dyn_val", F64),
        ("dyn_max_val", F64),
        ("val_iso", PTR),
        ("max_val_iso", PTR),
        ("dxf_id", U32),
        ("cell_color", I32),
        ("icon_set", I32),
        ("icon_id", I32),
        ("has_icon_id", I32),
    ],
)

SORT_CONDITION = Struct(
    "fm_sort_condition",
    [
        ("ref", MERGE_RANGE_BLOB),
        ("descending", I32),
        ("sort_by", I32),
        ("custom_list", PTR),
        ("dxf_id", U32),
        ("has_dxf_id", I32),
        ("icon_set", I32),
        ("icon_id", I32),
        ("has_icon_id", I32),
    ],
)

AUTO_FILTER = Struct(
    "fm_auto_filter",
    [
        ("range", MERGE_RANGE_BLOB),
        ("columns", PTR),
        ("column_count", U32),
        ("has_sort", I32),
        ("sort_ref", MERGE_RANGE_BLOB),
        ("column_sort", I32),
        ("case_sensitive", I32),
        ("sort_method", I32),
        ("conditions", PTR),
        ("condition_count", U32),
    ],
)

VALIDATION_OUTCOME = Struct(
    "fm_validation_outcome",
    [
        ("has_rule", I32),
        ("valid", I32),
        ("rule_index", U32),
        ("error_style", U8),
        ("_pad", ("blob_pad3", 3, 1)),
    ],
)

MENTION = Struct(
    "fm_mention",
    [("person_id", PTR), ("mention_id", PTR), ("start", U32), ("length", U32)],
)

THREADED_COMMENT = Struct(
    "fm_threaded_comment",
    [
        ("id", PTR),
        ("row", U32),
        ("col", U32),
        ("person_id", PTR),
        ("created", PTR),
        ("text", PTR),
        ("parent_id", PTR),
        ("done", I32),
        ("mentions", PTR),
        ("mention_count", U32),
    ],
)

PERSON = Struct(
    "fm_person",
    [("id", PTR), ("display_name", PTR), ("user_id", PTR), ("provider_id", PTR)],
)


# ---------------------------------------------------------------------------
# Drawing images
# ---------------------------------------------------------------------------

IMAGE_INFO = Struct(
    "fm_image_info",
    [("format", I32), ("px_width", U32), ("px_height", U32)],
)

DRAWING_OBJECT = Struct(
    "fm_drawing_object",
    [
        ("object_id", U32),
        ("kind", I32),
        ("anchor_kind", I32),
        ("edit_as", I32),
        ("from_row", U32),
        ("from_col", U32),
        ("from_row_off", I64),
        ("from_col_off", I64),
        ("to_row", U32),
        ("to_col", U32),
        ("to_row_off", I64),
        ("to_col_off", I64),
        ("cx", I64),
        ("cy", I64),
        ("image_format", I32),
        ("name", PTR),
        ("descr", PTR),
        ("media_path", PTR),
    ],
)

IMAGE_INSERT = Struct(
    "fm_image_insert",
    [
        ("name", PTR),
        ("descr", PTR),
        ("anchor_kind", I32),
        ("edit_as", I32),
        ("row", U32),
        ("col", U32),
        ("row_off_emu", I64),
        ("col_off_emu", I64),
        ("width_emu", I64),
        ("height_emu", I64),
    ],
)

DRAWING_ANCHOR = Struct(
    "fm_drawing_anchor",
    [
        ("anchor_kind", I32),
        ("edit_as", I32),
        ("row", U32),
        ("col", U32),
        ("row_off_emu", I64),
        ("col_off_emu", I64),
        ("width_emu", I64),
        ("height_emu", I64),
    ],
)


SHAPE = Struct(
    "fm_shape",
    [
        ("fields", U32),
        ("object_id", U32),
        ("kind", I32),
        ("anchor_kind", I32),
        ("edit_as", I32),
        ("from_row", U32),
        ("from_col", U32),
        ("to_row", U32),
        ("to_col", U32),
        ("fill_kind", I32),
        ("fill_rgb", U32),
        ("line_kind", I32),
        ("line_rgb", U32),
        ("rotation", I32),
        ("hidden", I32),
        ("text_box", I32),
        ("from_row_off", I64),
        ("from_col_off", I64),
        ("to_row_off", I64),
        ("to_col_off", I64),
        ("cx", I64),
        ("cy", I64),
        ("line_width_emu", I64),
        ("name", PTR),
        ("prst", PTR),
        ("text", PTR),
    ],
)

DATA_TABLE = Struct(
    "fm_data_table",
    [
        ("first_row", U32),
        ("first_col", U32),
        ("last_row", U32),
        ("last_col", U32),
        ("row_input_row", U32),
        ("row_input_col", U32),
        ("col_input_row", U32),
        ("col_input_col", U32),
        ("has_row_input", I32),
        ("has_col_input", I32),
        ("two_dimensional", I32),
        ("row_oriented", I32),
        ("row_input_deleted", I32),
        ("col_input_deleted", I32),
    ],
)

SLICER = Struct(
    "fm_slicer",
    [
        ("fields", U32),
        ("object_id", U32),
        ("column_count", U32),
        ("sort_order", I32),
        ("cross_filter", I32),
        ("anchor_kind", I32),
        ("from_row", U32),
        ("from_col", U32),
        ("to_row", U32),
        ("to_col", U32),
        ("from_row_off", I64),
        ("from_col_off", I64),
        ("to_row_off", I64),
        ("to_col_off", I64),
        ("row_height", I64),
        ("name", PTR),
        ("caption", PTR),
        ("cache_name", PTR),
        ("table_name", PTR),
        ("column_name", PTR),
        ("style", PTR),
    ],
)

SPARKLINE = Struct("fm_sparkline", [("formula", PTR), ("location", PTR)])

SPARKLINE_GROUP = Struct(
    "fm_sparkline_group",
    [
        ("color_series", COLOR_SPEC_BLOB),
        ("color_negative", COLOR_SPEC_BLOB),
        ("color_axis", COLOR_SPEC_BLOB),
        ("color_markers", COLOR_SPEC_BLOB),
        ("color_first", COLOR_SPEC_BLOB),
        ("color_last", COLOR_SPEC_BLOB),
        ("color_high", COLOR_SPEC_BLOB),
        ("color_low", COLOR_SPEC_BLOB),
        ("manual_min", F64),
        ("manual_max", F64),
        ("line_weight", F64),
        ("type", I32),
        ("markers", I32),
        ("high", I32),
        ("low", I32),
        ("first", I32),
        ("last", I32),
        ("negative", I32),
        ("display_x_axis", I32),
        ("display_hidden", I32),
        ("min_axis_type", I32),
        ("max_axis_type", I32),
        ("has_manual_min", I32),
        ("has_manual_max", I32),
        ("date_axis", I32),
        ("empty_cells_as", I32),
        ("date_axis_ref", PTR),
        ("sparkline_count", U32),
    ],
)


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------

CHART_TITLE = Struct("fm_chart_title_t", [("kind", I32), ("text", PTR), ("formula", PTR)])

CHART_FILL = Struct(
    "fm_chart_fill_t",
    [("has_color", I32), ("color", U32), ("transparency_pct", F64)],
)

CHART_LINE = Struct(
    "fm_chart_line_t",
    [
        ("has_color", I32),
        ("color", U32),
        ("transparency_pct", F64),
        ("width_pt", F64),
        ("has_dash", I32),
        ("dash", I32),
        ("none", I32),
    ],
)

CHART_MARKER = Struct(
    "fm_chart_marker_t",
    [
        ("symbol", I32),
        ("has_size", I32),
        ("size", U32),
        ("fill", nested(CHART_FILL)),
        ("line", nested(CHART_LINE)),
    ],
)

CHART_LABELS = Struct(
    "fm_chart_labels_t",
    [
        ("show_value", I32),
        ("show_category", I32),
        ("show_series", I32),
        ("show_percent", I32),
        ("show_legend_key", I32),
        ("separator", PTR),
        ("has_position", I32),
        ("position", I32),
        ("number_format", PTR),
        ("source_linked", I32),
        ("has_font_color", I32),
        ("font_color", U32),
        ("font_size_pt", F64),
        ("deleted", I32),
    ],
)

CHART_INFO = Struct(
    "fm_chart_info_t",
    [
        ("object_id", U32),
        ("name", PTR),
        ("descr", PTR),
        ("is_chartex", I32),
        ("is_pivot_chart", I32),
        ("on_chartsheet", I32),
        ("anchor_kind", I32),
        ("edit_as", I32),
        ("from_row", U32),
        ("from_col", U32),
        ("from_row_off", I64),
        ("from_col_off", I64),
        ("to_row", U32),
        ("to_col", U32),
        ("to_row_off", I64),
        ("to_col_off", I64),
        ("cx", I64),
        ("cy", I64),
        ("title", nested(CHART_TITLE)),
        ("auto_title_deleted", I32),
        ("legend_position", I32),
        ("legend_overlay", I32),
        ("legend_hidden_count", U32),
        ("blanks_as", I32),
        ("plot_vis_only", I32),
        ("date1904", I32),
        ("has_style", I32),
        ("style", I32),
        ("plot_count", U32),
        ("axis_count", U32),
        ("hidden_category_count", U32),
        ("source_range", PTR),
        ("source_direction", I32),
        ("source_header_row", I32),
        ("source_header_column", I32),
    ],
)

CHART_PLOT = Struct(
    "fm_chart_plot_t",
    [
        ("type", I32),
        ("chartex_layout", PTR),
        ("bar_dir", I32),
        ("has_grouping", I32),
        ("grouping", I32),
        ("vary_colors", I32),
        ("has_gap_width", I32),
        ("gap_width", U32),
        ("has_overlap", I32),
        ("overlap", I32),
        ("has_hole_size", I32),
        ("hole_size", U32),
        ("has_first_slice_angle", I32),
        ("first_slice_angle", U32),
        ("axis_id_count", U32),
        ("series_count", U32),
    ],
)

CHART_DATA_SOURCE = Struct(
    "fm_chart_data_source_t",
    [("kind", I32), ("formula", PTR), ("literal_count", U32)],
)

CHART_SERIES = Struct(
    "fm_chart_series_t",
    [
        ("idx", U32),
        ("order", U32),
        ("filtered", I32),
        ("name", nested(CHART_DATA_SOURCE)),
        ("categories", nested(CHART_DATA_SOURCE)),
        ("values", nested(CHART_DATA_SOURCE)),
        ("x", nested(CHART_DATA_SOURCE)),
        ("y", nested(CHART_DATA_SOURCE)),
        ("size", nested(CHART_DATA_SOURCE)),
        ("fill", nested(CHART_FILL)),
        ("line", nested(CHART_LINE)),
        ("has_marker", I32),
        ("marker", nested(CHART_MARKER)),
        ("smooth", I32),
        ("point_count", U32),
        ("has_labels", I32),
        ("labels", nested(CHART_LABELS)),
    ],
)

CHART_LITERAL = Struct(
    "fm_chart_literal_t",
    [("idx", U32), ("is_text", I32), ("number", F64), ("text", PTR)],
)

CHART_POINT = Struct(
    "fm_chart_point_t",
    [
        ("idx", U32),
        ("fill", nested(CHART_FILL)),
        ("line", nested(CHART_LINE)),
        ("has_marker", I32),
        ("marker", nested(CHART_MARKER)),
    ],
)

CHART_AXIS = Struct(
    "fm_chart_axis_t",
    [
        ("id", U32),
        ("kind", I32),
        ("position", I32),
        ("deleted", I32),
        ("cross_axis", U32),
        ("crosses", I32),
        ("crosses_at", F64),
        ("reversed", I32),
        ("log_base", F64),
        ("min", F64),
        ("max", F64),
        ("major_unit", F64),
        ("minor_unit", F64),
        ("number_format", PTR),
        ("source_linked", I32),
        ("title", nested(CHART_TITLE)),
        ("major_gridlines", I32),
        ("minor_gridlines", I32),
        ("tick_label_pos", PTR),
    ],
)

CHART_PREVIEW_OVERRIDES = Struct(
    "fm_chart_preview_overrides",
    [
        ("source", PTR),
        ("direction_op", I32),
        ("direction", I32),
        ("header_row_op", I32),
        ("header_row", I32),
        ("header_column_op", I32),
        ("header_column", I32),
        ("blanks_as_op", I32),
        ("blanks_as", I32),
        ("plot_vis_only_op", I32),
        ("plot_vis_only", I32),
    ],
)

CHART_DATA_INFO = Struct(
    "fm_chart_data_info_t",
    [
        ("is_xy", I32),
        ("point_count", U32),
        ("series_count", U32),
        ("category_label_count", U32),
        ("category_value_count", U32),
        ("category_level_count", U32),
        ("category_format_code", PTR),
        ("categories_generated", I32),
        ("title_text", PTR),
        ("axis_title_count", U32),
        ("unresolved_count", U32),
    ],
)

CHART_DATA_SERIES = Struct(
    "fm_chart_data_series_t",
    [
        ("plot_index", U32),
        ("idx", U32),
        ("name", PTR),
        ("format_code", PTR),
        ("filtered", I32),
        ("point_count", U32),
        ("has_sizes", I32),
        ("range_count", array(U32, 6)),
        ("x_format_code", PTR),
    ],
)

CHART_DATA_POINTS = Struct(
    "fm_chart_data_points_t",
    [("count", U32), ("values", PTR), ("states", PTR), ("x", PTR), ("sizes", PTR)],
)

CHART_SHEET_RANGE = Struct(
    "fm_chart_sheet_range_t",
    [("sheet", U32), ("top", U32), ("left", U32), ("bottom", U32), ("right", U32)],
)

CHART_AXIS_TITLE_TEXT = Struct(
    "fm_chart_axis_title_text_t",
    [("axis_id", U32), ("text", PTR)],
)

CHART_UNRESOLVED = Struct(
    "fm_chart_unresolved_t",
    [("series_idx", U32), ("source", PTR), ("axis_id", U32)],
)

CHART_INSERT = Struct(
    "fm_chart_insert",
    [
        ("kind", I32),
        ("grouping", I32),
        ("source", PTR),
        ("direction", I32),
        ("header_row", I32),
        ("header_column", I32),
        ("title", PTR),
        ("legend", I32),
        ("placement", nested(IMAGE_INSERT)),
    ],
)

CHART_FILL_PATCH = Struct(
    "fm_chart_fill_patch",
    [
        ("color_op", I32),
        ("color", U32),
        ("transparency_op", I32),
        ("transparency_pct", F64),
    ],
)

CHART_LINE_PATCH = Struct(
    "fm_chart_line_patch",
    [
        ("color_op", I32),
        ("color", U32),
        ("transparency_op", I32),
        ("transparency_pct", F64),
        ("width_op", I32),
        ("width_pt", F64),
        ("dash_op", I32),
        ("dash", I32),
        ("none_op", I32),
        ("none", I32),
    ],
)

CHART_MARKER_PATCH = Struct(
    "fm_chart_marker_patch",
    [
        ("symbol_op", I32),
        ("symbol", I32),
        ("size_op", I32),
        ("size", F64),
        ("fill", nested(CHART_FILL_PATCH)),
        ("line", nested(CHART_LINE_PATCH)),
    ],
)

CHART_LABELS_PATCH = Struct(
    "fm_chart_labels_patch",
    [
        ("show_value_op", I32),
        ("show_value", I32),
        ("show_category_op", I32),
        ("show_category", I32),
        ("show_series_op", I32),
        ("show_series", I32),
        ("show_percent_op", I32),
        ("show_percent", I32),
        ("show_legend_key_op", I32),
        ("show_legend_key", I32),
        ("separator_op", I32),
        ("separator", PTR),
        ("position_op", I32),
        ("position", I32),
        ("number_format_op", I32),
        ("number_format", PTR),
        ("font_color_op", I32),
        ("font_color", U32),
        ("font_size_op", I32),
        ("font_size_pt", F64),
    ],
)

CHART_SERIES_PATCH = Struct(
    "fm_chart_series_patch",
    [
        ("name", PTR),
        ("categories", PTR),
        ("values", PTR),
        ("x", PTR),
        ("y", PTR),
        ("size", PTR),
        ("fill", nested(CHART_FILL_PATCH)),
        ("line", nested(CHART_LINE_PATCH)),
        ("marker_op", I32),
        ("marker", nested(CHART_MARKER_PATCH)),
        ("smooth_op", I32),
        ("smooth", I32),
        ("labels_op", I32),
        ("labels", nested(CHART_LABELS_PATCH)),
    ],
)

CHART_POINT_PATCH = Struct(
    "fm_chart_point_patch",
    [
        ("fill", nested(CHART_FILL_PATCH)),
        ("line", nested(CHART_LINE_PATCH)),
        ("marker_op", I32),
        ("marker", nested(CHART_MARKER_PATCH)),
    ],
)

CHART_PLOT_PATCH = Struct(
    "fm_chart_plot_patch",
    [
        ("gap_width_op", I32),
        ("gap_width", U32),
        ("overlap_op", I32),
        ("overlap", I32),
        ("first_slice_angle_op", I32),
        ("first_slice_angle", U32),
        ("vary_colors_op", I32),
        ("vary_colors", I32),
        ("hole_size_op", I32),
        ("hole_size", U32),
    ],
)

CHART_AXIS_PATCH = Struct(
    "fm_chart_axis_patch",
    [
        ("visible_op", I32),
        ("visible", I32),
        ("title_op", I32),
        ("title", nested(CHART_TITLE)),
        ("min_op", I32),
        ("min", F64),
        ("max_op", I32),
        ("max", F64),
        ("major_unit_op", I32),
        ("major_unit", F64),
        ("number_format_op", I32),
        ("number_format", PTR),
        ("major_gridlines_op", I32),
        ("major_gridlines", I32),
        ("reversed_op", I32),
        ("reversed", I32),
        ("log_base_op", I32),
        ("log_base", F64),
        ("crosses_op", I32),
        ("crosses", I32),
        ("crosses_at", F64),
        ("tick_label_pos_op", I32),
        ("tick_label_pos", PTR),
    ],
)

CHART_LEGEND_PATCH = Struct(
    "fm_chart_legend_patch",
    [
        ("position", I32),
        ("overlay", I32),
        ("hidden_series_op", I32),
        ("hidden_series", PTR),
        ("hidden_series_count", U32),
    ],
)
