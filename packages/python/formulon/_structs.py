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
# Opaque 4-byte, 1-aligned blob: an inline ``fm_cf_color_t`` (four
# ``uint8_t`` channels). Not decoded by :meth:`Struct.unpack`.
CF_COLOR_BLOB = ("blob_cf_color", 4, 1)

_FMT = {
    "ptr": "<I",
    "u32": "<I",
    "i32": "<i",
    "u16": "<H",
    "u8": "<B",
    "f64": "<d",
}


class Struct:
    """A wasm32 POD layout: ordered (name, ctype) fields with offsets."""

    __slots__ = ("name", "fields", "offsets", "size")

    def __init__(self, name: str, fields: List[Tuple[str, Tuple]]) -> None:
        self.name = name
        self.fields = fields
        self.offsets: Dict[str, Tuple[str, int]] = {}
        off = 0
        max_align = 1
        for fname, (kind, sz, align) in fields:
            off = (off + align - 1) // align * align
            self.offsets[fname] = (kind, off)
            off += sz
            max_align = max(max_align, align)
        self.size = (off + max_align - 1) // max_align * max_align

    def pack(self, lib, ptr: int, values: Dict[str, object]) -> None:
        """Zero ``[ptr, ptr+size)`` then write each supplied field."""
        lib.write_bytes(ptr, b"\x00" * self.size)
        for fname, value in values.items():
            kind, off = self.offsets[fname]
            buf = struct.pack(_FMT[kind], value)
            lib.write_bytes(ptr + off, buf)

    def unpack(self, lib, ptr: int) -> Dict[str, int]:
        """Read every scalar field from ``[ptr, ptr+size)`` into a dict.

        Opaque blob fields (any ``kind`` not in ``_FMT``, e.g. an inline
        ``fm_value_t`` or ``fm_cfvo_t``) are skipped; callers decode those
        separately against ``ptr + self.offsets[name][1]``.
        """
        raw = lib.read_bytes(ptr, self.size)
        out: Dict[str, int] = {}
        for fname, (kind, off) in self.offsets.items():
            if kind not in _FMT:
                continue
            out[fname] = struct.unpack_from(_FMT[kind], raw, off)[0]
        return out


def alloc_struct(lib, layout: Struct) -> int:
    """Allocate a zeroed scratch block sized for ``layout``."""
    ptr = lib.alloc(layout.size)
    lib.write_bytes(ptr, b"\x00" * layout.size)
    return ptr


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
    ],
)

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
        ("has_style", I32),
        ("style_xf", U32),
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
        ("underline", U8),
        ("vert_align", U8),
        ("family", U8),
        ("charset", U8),
        ("scheme", U8),
        ("color", COLOR_SPEC_BLOB),
    ],
)

FILL_RECORD = Struct(
    "fm_fill_record",
    [
        ("pattern", U8),
        ("fg_argb", U32),
        ("bg_argb", U32),
        ("fg", COLOR_SPEC_BLOB),
        ("bg", COLOR_SPEC_BLOB),
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


def write_str_field(lib, struct_ptr: int, layout: Struct, field: str, value: Optional[str], owned: List[int]) -> None:
    """Encode ``value`` into WASM memory and store its pointer in a field.

    ``None`` / empty leaves the field as the NULL pointer (already zeroed
    by :meth:`Struct.pack`). Every buffer allocated here is appended to
    ``owned`` so the caller can free them after the WASM call returns.
    """
    if value is None or value == "":
        return
    ptr, _ = lib.alloc_utf8(value)
    owned.append(ptr)
    kind, off = layout.offsets[field]
    lib.write_bytes(struct_ptr + off, struct.pack("<I", ptr))
