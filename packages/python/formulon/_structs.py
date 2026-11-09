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
# Opaque 12-byte, 4-aligned blob: an inline ``fm_cfvo_t`` (``int32_t type;
# int32_t gte; const char* value;``). Not decoded by
# :meth:`Struct.unpack`; callers who need the sub-fields read them at
# ``ptr + offset`` with their own struct.unpack_from calls.
CFVO_BLOB = ("blob_cfvo", 12, 4)
# Opaque 16-byte, 4-aligned blob: an inline ``fm_cell_rect_t``. Pack and
# unpack it with :data:`CELL_RECT` against ``ptr + offset``.
CELL_RECT_BLOB = ("blob_cell_rect", 16, 4)

_FMT = {
    "ptr": "<I",
    "u32": "<I",
    "i32": "<i",
    "i64": "<q",
    "u16": "<H",
    "u8": "<B",
    "f64": "<d",
}


# Records: structs passed by pointer whose first member is ``uint32_t
# struct_size``. :class:`Struct` prepends the field and every packer and
# zeroing helper fills it with the layout's wasm32 size, so a caller never
# sets it. The set matches ``src/c_api/parts/struct_v1_sizes.h``; value types
# (array elements, colour, style and chart parts) are not in it.
RECORD_NAMES = frozenset(
    (
        "fm_array_formula_info_t",
        "fm_auto_filter_t",
        "fm_cell_node_info_t",
        "fm_cf_match_info_t",
        "fm_cf_rule_t",
        "fm_chart_axis_info_t",
        "fm_chart_axis_patch_t",
        "fm_chart_axis_title_text_info_t",
        "fm_chart_data_info_t",
        "fm_chart_data_points_info_t",
        "fm_chart_data_series_info_t",
        "fm_chart_display_patch_t",
        "fm_chart_info_t",
        "fm_chart_insert_spec_t",
        "fm_chart_legend_patch_t",
        "fm_chart_legend_t",
        "fm_chart_literal_info_t",
        "fm_chart_plot_info_t",
        "fm_chart_plot_patch_t",
        "fm_chart_point_info_t",
        "fm_chart_point_patch_t",
        "fm_chart_preview_overrides_spec_t",
        "fm_chart_series_info_t",
        "fm_chart_series_patch_t",
        "fm_chart_unresolved_info_t",
        "fm_column_break_t",
        "fm_column_layout_info_t",
        "fm_comment_info_t",
        "fm_data_table_t",
        "fm_cf_rule_patch_t",
        "fm_data_validation_patch_t",
        "fm_data_validation_t",
        "fm_defined_name_t",
        "fm_drawing_object_info_t",
        "fm_drawing_patch_t",
        "fm_effective_style_info_t",
        "fm_external_link_info_t",
        "fm_function_metadata_info_t",
        "fm_header_footer_patch_t",
        "fm_header_footer_t",
        "fm_hyperlink_t",
        "fm_image_info_t",
        "fm_iterative_settings_info_t",
        "fm_iterative_settings_patch_t",
        "fm_locale_facts_info_t",
        "fm_margins_pt_info_t",
        "fm_page_layout_info_t",
        "fm_page_margins_patch_t",
        "fm_page_margins_t",
        "fm_page_setup_patch_t",
        "fm_page_setup_t",
        "fm_paper_info_t",
        "fm_parallel_recalc_stats_info_t",
        "fm_parse_error_info_t",
        "fm_parse_reference_spec_t",
        "fm_parsed_reference_info_t",
        "fm_person_t",
        "fm_drawing_placement_spec_t",
        "fm_pivot_cache_options_patch_t",
        "fm_pivot_cache_options_t",
        "fm_pivot_calculated_item_info_t",
        "fm_pivot_cell_info_t",
        "fm_pivot_data_field_t",
        "fm_pivot_date_group_spec_t",
        "fm_pivot_field_info_t",
        "fm_pivot_field_patch_t",
        "fm_pivot_field_spec_t",
        "fm_pivot_filter_t",
        "fm_pivot_grand_totals_patch_t",
        "fm_pivot_item_info_t",
        "fm_pivot_item_patch_t",
        "fm_pivot_options_patch_t",
        "fm_pivot_options_t",
        "fm_pivot_spec_t",
        "fm_pivot_style_options_patch_t",
        "fm_pivot_style_options_t",
        "fm_pivot_worksheet_source_t",
        "fm_print_options_patch_t",
        "fm_print_options_t",
        "fm_print_titles_t",
        "fm_protected_range_patch_t",
        "fm_protected_range_t",
        "fm_read_diagnostics_info_t",
        "fm_rect_pt_info_t",
        "fm_row_break_t",
        "fm_row_layout_info_t",
        "fm_save_diagnostics_info_t",
        "fm_shape_info_t",
        "fm_shape_patch_t",
        "fm_shape_spec_t",
        "fm_sheet_format_defaults_patch_t",
        "fm_sheet_format_defaults_t",
        "fm_sheet_protection_patch_t",
        "fm_sheet_protection_t",
        "fm_sheet_view_info_t",
        "fm_sheet_view_patch_t",
        "fm_slicer_info_t",
        "fm_slicer_patch_t",
        "fm_slicer_spec_t",
        "fm_sparkline_group_patch_t",
        "fm_sparkline_group_t",
        "fm_spill_info_t",
        "fm_styles_batch_spec_t",
        "fm_table_column_info_t",
        "fm_table_info_t",
        "fm_table_patch_t",
        "fm_table_spec_t",
        "fm_table_style_defaults_patch_t",
        "fm_table_style_options_patch_t",
        "fm_table_style_options_t",
        "fm_table_style_t",
        "fm_theme_colors_t",
        "fm_theme_fonts_t",
        "fm_threaded_comment_patch_t",
        "fm_threaded_comment_t",
        "fm_validation_outcome_info_t",
        "fm_width_model_info_t",
        "fm_workbook_protection_t",
    )
)
STRUCT_SIZE_FIELD = "struct_size"


def nested(layout: "Struct") -> Tuple:
    """Field type for a struct embedded by value in another struct."""
    return ("nested", layout.size, layout.align, layout)


def array(elem: Tuple, count: int) -> Tuple:
    """Field type for an inline ``elem[count]`` array of scalars."""
    return ("array", elem[1] * count, elem[2], (_FMT[elem[0]], count))


class Struct:
    """A wasm32 POD layout: ordered (name, ctype) fields with offsets."""

    __slots__ = ("name", "fields", "offsets", "size", "align", "is_record", "_extra")

    def __init__(self, name: str, fields: List[Tuple[str, Tuple]]) -> None:
        self.name = name
        self.is_record = name in RECORD_NAMES
        if self.is_record:
            fields = [(STRUCT_SIZE_FIELD, U32)] + list(fields)
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
        """Write ``values`` into ``raw`` at ``base``; a nested struct takes a dict.

        A record's ``struct_size`` is filled first, so ``values`` need not carry
        it (a test may pass one to build a deliberately wrong size).
        """
        if self.is_record:
            struct.pack_into("<I", raw, base, self.size)
        for fname, value in values.items():
            kind, off = self.offsets[fname]
            if kind == "nested":
                self._extra[fname].pack_into(raw, base + off, value)
            else:
                struct.pack_into(_FMT[kind], raw, base + off, value)

    def blank(self) -> bytearray:
        """A zeroed block of ``size`` bytes with a record's ``struct_size`` set."""
        raw = bytearray(self.size)
        if self.is_record:
            struct.pack_into("<I", raw, 0, self.size)
        return raw

    def pack(self, lib, ptr: int, values: Dict[str, object]) -> None:
        """Pack supplied fields into a zeroed block and write it once."""
        raw = self.blank()
        self.pack_into(raw, 0, values)
        lib.write_bytes(ptr, raw)

    def unpack_from(self, raw: bytes, base: int = 0) -> Dict[str, object]:
        """Decode every scalar, nested-struct and array field of ``raw`` at ``base``."""
        out: Dict[str, object] = {}
        for fname, (kind, off) in self.offsets.items():
            if self.is_record and fname == STRUCT_SIZE_FIELD:
                continue
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
    ``malloc``/``free`` pair per item. A record keeps its ``struct_size``.
    """
    lib.write_bytes(ptr, layout.blank())


# ---------------------------------------------------------------------------
# Struct layouts (mirror src/c_api/formulon_c.h field-for-field)
# ---------------------------------------------------------------------------

CELL_RECT = Struct(
    "fm_cell_rect_t",
    [("first_row", U32), ("first_col", U32), ("last_row", U32), ("last_col", U32)],
)

SHEET_RECT = Struct(
    "fm_sheet_rect_t",
    [("sheet_index", U32), ("rect", nested(CELL_RECT))],
)

COLOR_SPEC = Struct(
    "fm_color_t",
    [("tint", F64), ("argb", U32), ("theme", U32), ("indexed", U32), ("kind", I32)],
)
# Inline ``fm_color_t``, packed and decoded as a nested dict. Nested
# rather than flattened into its owner: it is 8-aligned, so its trailing
# padding is part of the C layout.
COLOR_SPEC_BLOB = nested(COLOR_SPEC)

HYPERLINK = Struct(
    "fm_hyperlink_t",
    [
        ("rect", nested(CELL_RECT)),
        ("target", PTR),
        ("location", PTR),
        ("display", PTR),
        ("tooltip", PTR),
    ],
)

COMMENT = Struct(
    "fm_comment_info_t",
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
        ("workbook_legacy_password_hash", PTR),
        ("workbook_algorithm_name", PTR),
        ("workbook_hash_value", PTR),
        ("workbook_salt_value", PTR),
        ("revisions_legacy_password_hash", PTR),
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
        ("ranges", PTR),
        ("range_count", U32),
        ("legacy_password_hash", PTR),
        ("algorithm_name", PTR),
        ("hash_value", PTR),
        ("salt_value", PTR),
        ("security_descriptor", PTR),
    ],
)

DATA_VALIDATION = Struct(
    "fm_data_validation_t",
    [
        ("id", PTR),
        ("ranges", PTR),
        ("range_count", U32),
        ("type", I32),
        ("op", I32),
        ("error_style", I32),
        ("allow_blank", I32),
        ("show_input_message", I32),
        ("show_error_message", I32),
        ("show_in_cell_dropdown", I32),
        ("formula1", PTR),
        ("formula2", PTR),
        ("error_title", PTR),
        ("error_message", PTR),
        ("prompt_title", PTR),
        ("prompt_message", PTR),
    ],
)

DATA_VALIDATION_PATCH = Struct(
    "fm_data_validation_patch_t",
    [
        ("ranges_op", I32),
        ("ranges", PTR),
        ("range_count", U32),
        ("type_op", I32),
        ("type", I32),
        ("op_op", I32),
        ("op", I32),
        ("error_style_op", I32),
        ("error_style", I32),
        ("allow_blank_op", I32),
        ("allow_blank", I32),
        ("show_input_message_op", I32),
        ("show_input_message", I32),
        ("show_error_message_op", I32),
        ("show_error_message", I32),
        ("show_in_cell_dropdown_op", I32),
        ("show_in_cell_dropdown", I32),
        ("formula1_op", I32),
        ("formula1", PTR),
        ("formula2_op", I32),
        ("formula2", PTR),
        ("error_title_op", I32),
        ("error_title", PTR),
        ("error_message_op", I32),
        ("error_message", PTR),
        ("prompt_title_op", I32),
        ("prompt_title", PTR),
        ("prompt_message_op", I32),
        ("prompt_message", PTR),
    ],
)

# The ``allow_*`` flags of ``fm_sheet_protection_t`` in field order. The two
# select flags are ``fm_tristate_t``; the others are 0 / 1.
SHEET_PROTECTION_ALLOW_FLAGS = (
    "allow_edit_objects",
    "allow_edit_scenarios",
    "allow_format_cells",
    "allow_format_columns",
    "allow_format_rows",
    "allow_insert_columns",
    "allow_insert_rows",
    "allow_insert_hyperlinks",
    "allow_delete_columns",
    "allow_delete_rows",
    "allow_select_locked_cells",
    "allow_select_unlocked_cells",
    "allow_sort",
    "allow_auto_filter",
    "allow_pivot_tables",
)
SHEET_PROTECTION_TEXT_FIELDS = ("algorithm_name", "hash_value", "salt_value", "legacy_password_hash")

SHEET_PROTECTION = Struct(
    "fm_sheet_protection_t",
    [
        ("algorithm_name", PTR),
        ("hash_value", PTR),
        ("salt_value", PTR),
        ("spin_count", U32),
        ("legacy_password_hash", PTR),
    ]
    + [(name, I32) for name in SHEET_PROTECTION_ALLOW_FLAGS],
)

SHEET_PROTECTION_PATCH = Struct(
    "fm_sheet_protection_patch_t",
    [
        ("algorithm_name_op", I32),
        ("algorithm_name", PTR),
        ("hash_value_op", I32),
        ("hash_value", PTR),
        ("salt_value_op", I32),
        ("salt_value", PTR),
        ("spin_count_op", I32),
        ("spin_count", U32),
        ("legacy_password_hash_op", I32),
        ("legacy_password_hash", PTR),
    ]
    + [field for name in SHEET_PROTECTION_ALLOW_FLAGS for field in ((name + "_op", I32), (name, I32))],
)

CFVO = Struct(
    "fm_cfvo_t",
    [
        ("type", I32),
        ("gte", I32),
        ("value", PTR),
    ],
)

CF_MATCH = Struct(
    "fm_cf_match_info_t",
    [
        ("kind", I32),
        ("priority", I32),
        ("has_dxf_index", I32),
        ("dxf_index", U32),
        ("color", COLOR_SPEC_BLOB),
        ("bar_length_pct", F64),
        ("bar_axis_position_pct", F64),
        ("bar_is_negative", I32),
        ("bar_fill", COLOR_SPEC_BLOB),
        ("bar_border", COLOR_SPEC_BLOB),
        ("bar_gradient", I32),
        ("icon_set_name", I32),
        ("icon_index", U8),
        ("bar_direction", I32),
    ],
)

CF_RULE = Struct(
    "fm_cf_rule_t",
    [
        ("id", PTR),
        ("type", I32),
        ("op", I32),
        ("time_period", I32),
        ("priority", I32),
        ("stop_if_true", I32),
        ("has_dxf_index", I32),
        ("dxf_index", U32),
        ("ranges", PTR),
        ("range_count", U32),
        ("formula1", PTR),
        ("formula2", PTR),
        ("has_op", I32),
        ("has_rank", I32),
        ("rank", I32),
        ("percent", I32),
        ("bottom", I32),
        ("above_average", I32),
        ("equal_average", I32),
        ("has_std_dev", I32),
        ("std_dev", F64),
        ("text", PTR),
        ("has_time_period", I32),
        ("color_scale_thresholds", PTR),
        ("color_scale_colors", PTR),
        ("color_scale_count", U32),
        ("has_data_bar", I32),
        ("data_bar_min", CFVO_BLOB),
        ("data_bar_max", CFVO_BLOB),
        ("data_bar_fill", COLOR_SPEC_BLOB),
        ("data_bar_show_value", I32),
        ("data_bar_min_length_pct", U8),
        ("data_bar_max_length_pct", U8),
        ("has_icon_set", I32),
        ("icon_set_name", I32),
        ("icon_set_thresholds", PTR),
        ("icon_set_threshold_count", U32),
        ("icon_set_reverse", I32),
        ("icon_set_show_value", I32),
        ("icon_set_percent", I32),
        # Data-bar attributes carried as has-flag / value pairs: a clear
        # flag means "use the model default" rather than "zero".
        ("has_data_bar_gradient", I32),
        ("data_bar_gradient", I32),
        ("has_data_bar_axis_position", I32),
        ("data_bar_axis_position", I32),
        ("data_bar_negative_fill", COLOR_SPEC_BLOB),
        ("data_bar_border", COLOR_SPEC_BLOB),
        ("data_bar_negative_border", COLOR_SPEC_BLOB),
        ("data_bar_axis_color", COLOR_SPEC_BLOB),
        ("has_icon_set_floor", I32),
        ("icon_set_floor", CFVO_BLOB),
        ("data_bar_direction", I32),
        ("icon_set_icons", PTR),
        ("icon_set_icon_count", U32),
    ],
)

CF_ICON = Struct("fm_cf_icon_t", [("icon_set", I32), ("icon_index", I32)])

CF_RULE_PATCH = Struct(
    "fm_cf_rule_patch_t",
    [
        ("priority_op", I32),
        ("priority", I32),
        ("stop_if_true_op", I32),
        ("stop_if_true", I32),
        ("dxf_index_op", I32),
        ("dxf_index", U32),
        ("ranges_op", I32),
        ("ranges", PTR),
        ("range_count", U32),
        ("formula1_op", I32),
        ("formula1", PTR),
        ("formula2_op", I32),
        ("formula2", PTR),
        ("op_op", I32),
        ("op", I32),
        ("rank_op", I32),
        ("rank", I32),
        ("percent_op", I32),
        ("percent", I32),
        ("bottom_op", I32),
        ("bottom", I32),
        ("above_average_op", I32),
        ("above_average", I32),
        ("equal_average_op", I32),
        ("equal_average", I32),
        ("std_dev_op", I32),
        ("std_dev", F64),
        ("text_op", I32),
        ("text", PTR),
        ("time_period_op", I32),
        ("time_period", I32),
        ("color_scale_thresholds_op", I32),
        ("color_scale_thresholds", PTR),
        ("color_scale_thresholds_count", U32),
        ("color_scale_colors_op", I32),
        ("color_scale_colors", PTR),
        ("color_scale_colors_count", U32),
        ("data_bar_min_op", I32),
        ("data_bar_min", nested(CFVO)),
        ("data_bar_max_op", I32),
        ("data_bar_max", nested(CFVO)),
        ("data_bar_fill_op", I32),
        ("data_bar_fill", COLOR_SPEC_BLOB),
        ("data_bar_show_value_op", I32),
        ("data_bar_show_value", I32),
        ("data_bar_min_length_pct_op", I32),
        ("data_bar_min_length_pct", U32),
        ("data_bar_max_length_pct_op", I32),
        ("data_bar_max_length_pct", U32),
        ("data_bar_gradient_op", I32),
        ("data_bar_gradient", I32),
        ("data_bar_axis_position_op", I32),
        ("data_bar_axis_position", I32),
        ("data_bar_negative_fill_op", I32),
        ("data_bar_negative_fill", COLOR_SPEC_BLOB),
        ("data_bar_border_op", I32),
        ("data_bar_border", COLOR_SPEC_BLOB),
        ("data_bar_negative_border_op", I32),
        ("data_bar_negative_border", COLOR_SPEC_BLOB),
        ("data_bar_axis_color_op", I32),
        ("data_bar_axis_color", COLOR_SPEC_BLOB),
        ("data_bar_direction_op", I32),
        ("data_bar_direction", I32),
        ("icon_set_name_op", I32),
        ("icon_set_name", I32),
        ("icon_set_thresholds_op", I32),
        ("icon_set_thresholds", PTR),
        ("icon_set_thresholds_count", U32),
        ("icon_set_reverse_op", I32),
        ("icon_set_reverse", I32),
        ("icon_set_show_value_op", I32),
        ("icon_set_show_value", I32),
        ("icon_set_percent_op", I32),
        ("icon_set_percent", I32),
        ("icon_set_floor_op", I32),
        ("icon_set_floor", nested(CFVO)),
        ("icon_set_icons_op", I32),
        ("icon_set_icons", PTR),
        ("icon_set_icons_count", U32),
    ],
)

CELL_NODE = Struct(
    "fm_cell_node_info_t",
    [("sheet_index", U32), ("row", U32), ("col", U32)],
)

PARSE_ERROR = Struct(
    "fm_parse_error_info_t",
    [("code", I32), ("start", U32), ("end", U32), ("line", U32), ("column", U32)],
)

# One ``<rPh>`` block. Passed both ways: as an element of the array
# ``fm_sheet_set_cell_phonetic_runs`` takes, and as the out-parameter of
# ``fm_sheet_get_cell_phonetic_run_at``.
PHONETIC_RUN = Struct(
    "fm_phonetic_run_t",
    [("sb", U32), ("eb", U32), ("text", PTR)],
)

# fm_pivot_cell_info_t embeds an fm_value_t (16 bytes, 8-aligned) inline. The
# ``value`` field is an opaque blob; decode it with Value._from_wasm
# against ``cell_ptr + PIVOT_CELL_VALUE_OFFSET``.
PIVOT_CELL = Struct(
    "fm_pivot_cell_info_t",
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
    "fm_pivot_data_field_t",
    [
        ("name", PTR),
        ("field_index", U32),
        ("aggregation", I32),
        ("number_format", PTR),
        ("show_as", I32),
        ("has_show_as_base_field", I32),
        ("show_as_base_field", U32),
        ("has_show_as_base_item", I32),
        ("show_as_base_item", U32),
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
        ("has_cache_field_index", I32),
        ("cache_field_index", U32),
        ("sort_order", I32),
        ("sort_by_field", PTR),
        ("date_periods", U32),
        ("date_calendar", I32),
        ("has_group_base_index", I32),
        ("group_base_index", U32),
        ("default_subtotal", I32),
        ("subtotal_functions", U32),
        ("insert_blank_row", I32),
        ("fill_down_labels", I32),
        ("show_all", I32),
        ("show_detail", I32),
    ],
)

PIVOT_ITEM_INFO = Struct(
    "fm_pivot_item_info_t",
    [
        ("name", PTR),
        ("has_cache_index", I32),
        ("cache_index", U32),
        ("visible", I32),
        ("caption", PTR),
        ("show_detail", I32),
    ],
)

TABLE_INFO = Struct(
    "fm_table_info_t",
    [
        ("name", PTR),
        ("range", nested(CELL_RECT)),
        ("style_name", PTR),
        ("show_first_column", I32),
        ("show_last_column", I32),
        ("show_row_stripes", I32),
        ("show_column_stripes", I32),
        ("header_row_count", U32),
        ("totals_row_count", U32),
        ("column_count", U32),
    ],
)

TABLE_SPEC = Struct(
    "fm_table_spec_t",
    [
        ("range", nested(CELL_RECT)),
        ("name", PTR),
        ("column_names", PTR),
        ("column_count", U32),
        ("style_name", PTR),
        ("header_row", I32),
        ("totals_row", I32),
        ("show_first_column", I32),
        ("show_last_column", I32),
        ("show_row_stripes", I32),
        ("show_column_stripes", I32),
    ],
)

TABLE_STYLE_OPTIONS = Struct(
    "fm_table_style_options_t",
    [
        ("style_name", PTR),
        ("show_first_column", I32),
        ("show_last_column", I32),
        ("show_row_stripes", I32),
        ("show_column_stripes", I32),
    ],
)

TABLE_STYLE_OPTIONS_PATCH = Struct(
    "fm_table_style_options_patch_t",
    [
        ("style_name_op", I32),
        ("style_name", PTR),
        ("show_first_column_op", I32),
        ("show_first_column", I32),
        ("show_last_column_op", I32),
        ("show_last_column", I32),
        ("show_row_stripes_op", I32),
        ("show_row_stripes", I32),
        ("show_column_stripes_op", I32),
        ("show_column_stripes", I32),
    ],
)

TABLE_COLUMN_INFO = Struct(
    "fm_table_column_info_t",
    [
        ("name", PTR),
        ("totals_row_function", I32),
        ("totals_row_label", PTR),
        ("totals_row_formula", PTR),
    ],
)

TABLE_STYLE_ELEMENT = Struct(
    "fm_table_style_element_t",
    [("type", I32), ("size", U32), ("has_dxf_index", I32), ("dxf_index", U32)],
)

TABLE_STYLE = Struct(
    "fm_table_style_t",
    [("name", PTR), ("pivot", I32), ("table", I32), ("elements", PTR), ("element_count", U32)],
)

PIVOT_STYLE_OPTIONS = Struct(
    "fm_pivot_style_options_t",
    [
        ("style_name", PTR),
        ("show_row_headers", I32),
        ("show_column_headers", I32),
        ("show_row_stripes", I32),
        ("show_column_stripes", I32),
        ("show_last_column", I32),
    ],
)

PIVOT_SPEC = Struct(
    "fm_pivot_spec_t",
    [("name", PTR), ("cache_id", U32), ("anchor_row", U32), ("anchor_col", U32)],
)

PIVOT_WORKSHEET_SOURCE = Struct(
    "fm_pivot_worksheet_source_t",
    [("ref", PTR), ("sheet", PTR), ("name", PTR)],
)

PIVOT_STYLE_OPTIONS_PATCH = Struct(
    "fm_pivot_style_options_patch_t",
    [
        ("style_name_op", I32),
        ("style_name", PTR),
        ("show_row_headers_op", I32),
        ("show_row_headers", I32),
        ("show_column_headers_op", I32),
        ("show_column_headers", I32),
        ("show_row_stripes_op", I32),
        ("show_row_stripes", I32),
        ("show_column_stripes_op", I32),
        ("show_column_stripes", I32),
        ("show_last_column_op", I32),
        ("show_last_column", I32),
    ],
)

PIVOT_FIELD_PATCH = Struct(
    "fm_pivot_field_patch_t",
    [
        ("custom_name_op", I32),
        ("custom_name", PTR),
        ("number_format_op", I32),
        ("number_format", PTR),
        ("default_subtotal_op", I32),
        ("default_subtotal", I32),
        ("insert_blank_row_op", I32),
        ("insert_blank_row", I32),
        ("fill_down_labels_op", I32),
        ("fill_down_labels", I32),
        ("show_all_op", I32),
        ("show_all", I32),
        ("subtotal_top_op", I32),
        ("subtotal_top", I32),
        ("sort_order_op", I32),
        ("sort_order", I32),
        ("sort_by_field_op", I32),
        ("sort_by_field", PTR),
        ("show_detail_op", I32),
        ("show_detail", I32),
    ],
)

PIVOT_ITEM_PATCH = Struct(
    "fm_pivot_item_patch_t",
    [
        ("visible_op", I32),
        ("visible", I32),
        ("show_detail_op", I32),
        ("show_detail", I32),
        ("caption_op", I32),
        ("caption", PTR),
    ],
)

PIVOT_DATE_GROUP_SPEC = Struct(
    "fm_pivot_date_group_spec_t",
    [
        ("periods", U32),
        ("calendar", I32),
        ("interval_days", U32),
        ("has_start", I32),
        ("start", F64),
        ("has_end", I32),
        ("end", F64),
    ],
)

PIVOT_FILTER_SPEC = Struct(
    "fm_pivot_filter_t",
    [
        ("field_index", U32),
        ("type", I32),
        ("value_kind", I32),
        ("value_int", I32),
        ("value_double", F64),
        ("value_text", PTR),
        ("value_high_kind", I32),
        ("value_high_int", I32),
        ("value_high_double", F64),
        ("data_field_index", U32),
        ("value_high_text", PTR),
    ],
)

PIVOT_OPTIONS = Struct(
    "fm_pivot_options_t",
    [
        ("show_error", I32),
        ("error_caption", PTR),
        ("show_missing", I32),
        ("missing_caption", PTR),
        ("show_headers", I32),
        ("show_drill", I32),
        ("show_data_tips", I32),
        ("grid_drop_zones", I32),
        ("show_empty_row", I32),
        ("show_empty_column", I32),
        ("show_data_drop_down", I32),
        ("print_drill", I32),
        ("field_print_titles", I32),
        ("item_print_titles", I32),
        ("merge_item", I32),
        ("indent", U32),
        ("page_over_then_down", I32),
        ("page_wrap", U32),
        ("multiple_field_filters", I32),
        ("preserve_formatting", I32),
        ("custom_list_sort", I32),
        ("alt_text", PTR),
        ("alt_text_summary", PTR),
        ("grand_total_caption", PTR),
        ("auto_formatting", I32),
        ("enable_drill", I32),
    ],
)

PIVOT_CACHE_OPTIONS = Struct(
    "fm_pivot_cache_options_t",
    [("save_data", I32), ("has_missing_items_limit", I32), ("missing_items_limit", U32), ("refresh_on_load", I32)],
)

PIVOT_CALCULATED_ITEM = Struct(
    "fm_pivot_calculated_item_info_t",
    [("field_index", U32), ("name", PTR), ("formula", PTR), ("scoped", I32)],
)

READ_DIAGNOSTICS = Struct(
    "fm_read_diagnostics_info_t",
    [
        ("undecoded_formula_count", U32),
        ("undecoded_defined_name_count", U32),
        ("undecoded_part_count", U32),
        ("skipped_feature_count", U32),
        ("unknown_content_type_count", U32),
    ],
)

SAVE_DIAGNOSTICS = Struct(
    "fm_save_diagnostics_info_t",
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
        ("rect", nested(CELL_RECT)),
        ("blocked", I32),
    ],
)

PARSE_REFERENCE_SPEC = Struct(
    "fm_parse_reference_spec_t",
    [("r1c1", I32), ("has_base", I32), ("base_row", U32), ("base_col", U32)],
)

REFERENCE = Struct(
    "fm_parsed_reference_info_t",
    [
        ("sheet_name", PTR),
        ("rect", nested(CELL_RECT)),
    ],
)

ARRAY_FORMULA = Struct(
    "fm_array_formula_info_t",
    [
        ("rect", nested(CELL_RECT)),
        ("formula", PTR),
    ],
)

FUNCTION_METADATA = Struct(
    "fm_function_metadata_info_t",
    [
        ("canonical_name", PTR),
        ("min_arity", U32),
        ("has_max_arity", I32),
        ("max_arity", U32),
        ("availability", I32),
        ("signature_template", PTR),
        ("description", PTR),
    ],
)

LOCALE_FACTS = Struct(
    "fm_locale_facts_info_t",
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

ITERATIVE_SETTINGS_PATCH = Struct(
    "fm_iterative_settings_patch_t",
    [
        ("enabled_op", I32),
        ("enabled", I32),
        ("max_iterations_op", I32),
        ("max_iterations", U32),
        ("max_change_op", I32),
        ("max_change", F64),
    ],
)

DEFINED_NAME = Struct(
    "fm_defined_name_t",
    [
        ("name", PTR),
        ("formula", PTR),
        ("local_sheet_id", I32),
        ("comment", PTR),
        ("hidden", I32),
    ],
)

SHEET_VIEW = Struct(
    "fm_sheet_view_info_t",
    [
        ("zoom_pct", U32),
        ("freeze_rows", U32),
        ("freeze_cols", U32),
        ("show_grid_lines", I32),
        ("show_row_column_headers", I32),
        ("show_zeros", I32),
        ("right_to_left", I32),
        ("tab_selected", I32),
        ("view_mode", I32),
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
    "fm_column_layout_info_t",
    [
        ("first", U32),
        ("last", U32),
        ("width_chars", F64),
        ("hidden", I32),
        ("outline_level", U8),
        ("collapsed", I32),
        ("has_width_chars", I32),
        ("has_xf_index", I32),
        ("xf_index", U32),
    ],
)

ROW_LAYOUT = Struct(
    "fm_row_layout_info_t",
    [
        ("row", U32),
        ("height_pt", F64),
        ("hidden", I32),
        ("outline_level", U8),
        ("collapsed", I32),
        ("has_xf_index", I32),
        ("xf_index", U32),
        ("has_height_pt", I32),
        ("custom_height", I32),
    ],
)

SHEET_FORMAT_DEFAULTS = Struct(
    "fm_sheet_format_defaults_t",
    [
        ("default_column_width_chars", F64),
        ("default_row_height_pt", F64),
        ("base_column_width_chars", F64),
        ("has_default_column_width_chars", I32),
        ("has_default_row_height_pt", I32),
        ("custom_height", I32),
        ("zero_height", I32),
        ("effective_default_column_width_chars", F64),
        ("effective_default_row_height_pt", F64),
    ],
)

RECT_PT = Struct(
    "fm_rect_pt_info_t",
    [("x", F64), ("y", F64), ("width_pt", F64), ("height_pt", F64)],
)

WIDTH_MODEL = Struct(
    "fm_width_model_info_t",
    [
        ("points_per_char", F64),
        ("padding_pt", F64),
        ("normal_font_size_pt", F64),
        ("calibrated", I32),
        ("platform", I32),
        ("normal_font_name", PTR),
    ],
)

PAPER_INFO = Struct(
    "fm_paper_info_t",
    [("width_pt", F64), ("height_pt", F64), ("landscape", I32), ("known", I32)],
)

MARGINS_PT = Struct(
    "fm_margins_pt_info_t",
    [
        ("left_in", F64),
        ("right_in", F64),
        ("top_in", F64),
        ("bottom_in", F64),
        ("header_in", F64),
        ("footer_in", F64),
    ],
)

PRINT_TITLES = Struct(
    "fm_print_titles_t",
    [
        ("has_rows", I32),
        ("first_row", U32),
        ("last_row", U32),
        ("has_columns", I32),
        ("first_col", U32),
        ("last_col", U32),
    ],
)

PAGE_LAYOUT = Struct(
    "fm_page_layout_info_t",
    [
        ("area_index", U32),
        ("rect", nested(CELL_RECT)),
        ("origin_x_pt", F64),
        ("origin_y_pt", F64),
        ("width_pt", F64),
        ("height_pt", F64),
    ],
)

# Every optional alignment attribute has an explicit presence flag so zero /
# false values can be round-tripped.
CELL_XF = Struct(
    "fm_cell_xf_t",
    [
        ("font_index", U32),
        ("fill_index", U32),
        ("border_index", U32),
        ("num_fmt_id", U16),
        ("horizontal", I32),
        ("vertical", I32),
        ("wrap_text", I32),
        ("justify_last_line", I32),
        ("style_xf_index", U32),
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
        ("reading_order", I32),
        ("has_horizontal", I32),
        ("has_vertical", I32),
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

FONT_RECORD = Struct(
    "fm_font_t",
    [
        ("name", PTR),
        ("size_pt", F64),
        ("bold", I32),
        ("italic", I32),
        ("strike", I32),
        ("has_bold", I32),
        ("has_italic", I32),
        ("has_strike", I32),
        ("has_family", I32),
        ("has_charset", I32),
        ("has_size_pt", I32),
        ("outline", I32),
        ("shadow", I32),
        ("condense", I32),
        ("extend", I32),
        ("has_outline", I32),
        ("has_shadow", I32),
        ("has_condense", I32),
        ("has_extend", I32),
        ("underline", I32),
        ("vert_align", I32),
        ("family", I32),
        ("charset", U8),
        ("scheme", I32),
        ("color", COLOR_SPEC_BLOB),
    ],
)

GRADIENT_STOP = Struct(
    "fm_gradient_stop_t",
    [("position", F64), ("color", COLOR_SPEC_BLOB)],
)

FILL_RECORD = Struct(
    "fm_fill_t",
    [
        ("pattern", I32),
        ("foreground", COLOR_SPEC_BLOB),
        ("background", COLOR_SPEC_BLOB),
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
    "fm_border_side_t",
    [("style", I32), ("color", COLOR_SPEC_BLOB)],
)
BORDER_SIDE_BLOB = ("blob_border_side", BORDER_SIDE.size, 8)

BORDER_RECORD = Struct(
    "fm_border_t",
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

# Inline members of ``fm_dxf_t``. They deliberately remain opaque here:
# callers marshal them through the standalone layouts above at their recorded
# offsets.
FONT_RECORD_BLOB = ("blob_font_record", FONT_RECORD.size, 8)
FILL_RECORD_BLOB = ("blob_fill_record", FILL_RECORD.size, 8)
BORDER_RECORD_BLOB = ("blob_border_record", BORDER_RECORD.size, 8)

DXF_RECORD = Struct(
    "fm_dxf_t",
    [
        ("has_font", I32),
        ("font", FONT_RECORD_BLOB),
        ("has_fill", I32),
        ("fill", FILL_RECORD_BLOB),
        ("has_border", I32),
        ("border", BORDER_RECORD_BLOB),
        ("number_format", PTR),
        ("alignment_xml", PTR),
        ("protection_xml", PTR),
    ],
)

CELL_STYLE_RECORD = Struct(
    "fm_cell_style_t",
    [
        ("name", PTR),
        ("style_xf_index", U32),
        ("has_builtin_id", I32),
        ("builtin_id", U32),
        ("i_level", U32),
        ("hidden", I32),
        ("custom_builtin", I32),
    ],
)

THEME_COLORS = Struct("fm_theme_colors_t", [("argb", ("blob_u32x12", 48, 4))])

THEME_FONTS = Struct(
    "fm_theme_fonts_t",
    [
        ("major_latin", PTR),
        ("major_east_asian", PTR),
        ("minor_latin", PTR),
        ("minor_east_asian", PTR),
    ],
)

EFFECTIVE_STYLE = Struct(
    "fm_effective_style_info_t",
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
    "fm_external_link_info_t",
    [
        ("book_number", U32),
        ("rel_id", PTR),
        ("part_path", PTR),
        ("target", PTR),
        ("target_external", I32),
        ("kind", I32),
    ],
)

ROW_BREAK = Struct(
    "fm_row_break_t",
    [
        ("row", U32),
        ("first", U32),
        ("last", U32),
        ("manual", I32),
    ],
)

COLUMN_BREAK = Struct(
    "fm_column_break_t",
    [
        ("col", U32),
        ("first", U32),
        ("last", U32),
        ("manual", I32),
    ],
)

# The three print patch structs pair every value field with an
# ``has_*`` flag: on the write path the flag selects which attributes
# the engine touches, on the read path it reports whether the attribute is
# stated in the XML at all. Field order must match the C declaration
# exactly -- ``Struct`` derives offsets from declaration order, so a
# reordering here silently mis-decodes rather than failing.
PAGE_SETUP = Struct(
    "fm_page_setup_t",
    [
        ("has_orientation", I32),
        ("orientation", I32),
        ("has_paper_size", I32),
        ("paper_size", U32),
        ("has_scale_pct", I32),
        ("scale_pct", U32),
        ("has_fit_to_width", I32),
        ("fit_to_width", U32),
        ("has_fit_to_height", I32),
        ("fit_to_height", U32),
        ("has_fit_to_page", I32),
        ("fit_to_page", I32),
    ],
)

# `double` is 8-aligned, so each `I32` flag is followed by four bytes of
# padding. `Struct` inserts it from the declared alignments.
PAGE_MARGINS = Struct(
    "fm_page_margins_t",
    [
        ("has_left_in", I32),
        ("left_in", F64),
        ("has_right_in", I32),
        ("right_in", F64),
        ("has_top_in", I32),
        ("top_in", F64),
        ("has_bottom_in", I32),
        ("bottom_in", F64),
        ("has_header_in", I32),
        ("header_in", F64),
        ("has_footer_in", I32),
        ("footer_in", F64),
    ],
)

PRINT_OPTIONS = Struct(
    "fm_print_options_t",
    [
        ("has_grid_lines", I32),
        ("grid_lines", I32),
        ("has_headings", I32),
        ("headings", I32),
        ("has_horizontal_centered", I32),
        ("horizontal_centered", I32),
        ("has_vertical_centered", I32),
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
        ("has_different_odd_even", I32),
        ("different_odd_even", I32),
        ("has_different_first", I32),
        ("different_first", I32),
        ("has_scale_with_doc", I32),
        ("scale_with_doc", I32),
        ("has_align_with_margins", I32),
        ("align_with_margins", I32),
    ],
)


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
    ``get_cell_phonetic_run()``'s output could not always round-trip back
    through ``set_cell_phonetic_runs()``.
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
    "fm_date_group_item_t",
    [
        ("year", U16),
        ("month", U8),
        ("day", U8),
        ("hour", U8),
        ("minute", U8),
        ("second", U8),
        ("grouping", I32),
    ],
)

# One custom condition of a filter column; an array element, not a record.
FILTER_CONDITION = Struct("fm_filter_condition_t", [("op", I32), ("value", PTR)])

FILTER_COLUMN = Struct(
    "fm_filter_column_t",
    [
        ("column_offset", U32),
        ("hidden_button", I32),
        ("show_button", I32),
        ("kind", I32),
        ("filter_blank", I32),
        ("values", PTR),
        ("value_count", U32),
        ("date_groups", PTR),
        ("date_group_count", U32),
        ("custom_and", I32),
        ("conditions", PTR),
        ("condition_count", U32),
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
        ("has_dxf_index", I32),
        ("dxf_index", U32),
        ("cell_color", I32),
        ("icon_set", I32),
        ("icon_index", I32),
        ("has_icon_index", I32),
    ],
)

SORT_CONDITION = Struct(
    "fm_sort_condition_t",
    [
        ("ref", CELL_RECT_BLOB),
        ("descending", I32),
        ("sort_by", I32),
        ("custom_list", PTR),
        ("dxf_index", U32),
        ("has_dxf_index", I32),
        ("icon_set", I32),
        ("icon_index", I32),
        ("has_icon_index", I32),
    ],
)

AUTO_FILTER = Struct(
    "fm_auto_filter_t",
    [
        ("range", CELL_RECT_BLOB),
        ("columns", PTR),
        ("column_count", U32),
        ("has_sort", I32),
        ("sort_ref", CELL_RECT_BLOB),
        ("column_sort", I32),
        ("case_sensitive", I32),
        ("sort_method", I32),
        ("conditions", PTR),
        ("condition_count", U32),
    ],
)

VALIDATION_OUTCOME = Struct(
    "fm_validation_outcome_info_t",
    [
        ("has_rule", I32),
        ("valid", I32),
        ("error_style", I32),
        ("rule_id", PTR),
    ],
)

MENTION = Struct(
    "fm_mention_t",
    [("person_id", PTR), ("mention_id", PTR), ("start", U32), ("length", U32)],
)

THREADED_COMMENT = Struct(
    "fm_threaded_comment_t",
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
        ("mentions_count", U32),
    ],
)

# ``text`` NULL keeps the text; ``mentions_engaged`` 0 keeps the mentions;
# ``done`` -1 keeps the resolved state, 0 / 1 sets it.
THREADED_COMMENT_PATCH = Struct(
    "fm_threaded_comment_patch_t",
    [
        ("text_op", I32),
        ("text", PTR),
        ("mentions_op", I32),
        ("mentions", PTR),
        ("mentions_count", U32),
        ("done_op", I32),
        ("done", I32),
    ],
)

PERSON = Struct(
    "fm_person_t",
    [("id", PTR), ("display_name", PTR), ("user_id", PTR), ("provider_id", PTR)],
)


# ---------------------------------------------------------------------------
# Drawing images
# ---------------------------------------------------------------------------

IMAGE_INFO = Struct(
    "fm_image_info_t",
    [("format", I32), ("width_px", U32), ("height_px", U32)],
)

DRAWING_ANCHOR = Struct(
    "fm_drawing_anchor_t",
    [
        ("kind", I32),
        ("edit_as", I32),
        ("from_row", U32),
        ("from_col", U32),
        ("from_row_offset_emu", I64),
        ("from_col_offset_emu", I64),
        ("to_row", U32),
        ("to_col", U32),
        ("to_row_offset_emu", I64),
        ("to_col_offset_emu", I64),
        ("width_emu", I64),
        ("height_emu", I64),
    ],
)

DRAWING_OBJECT = Struct(
    "fm_drawing_object_info_t",
    [
        ("object_id", U32),
        ("kind", I32),
        ("hidden", I32),
        ("anchor", nested(DRAWING_ANCHOR)),
        ("name", PTR),
        ("description", PTR),
        ("media_path", PTR),
    ],
)

DRAWING_PLACEMENT_SPEC = Struct(
    "fm_drawing_placement_spec_t",
    [
        ("name", PTR),
        ("description", PTR),
        ("hidden", I32),
        ("anchor", nested(DRAWING_ANCHOR)),
    ],
)

DRAWING_ANCHOR_PATCH = Struct(
    "fm_drawing_anchor_patch_t",
    [
        ("kind_op", I32),
        ("kind", I32),
        ("edit_as_op", I32),
        ("edit_as", I32),
        ("from_row_op", I32),
        ("from_row", U32),
        ("from_col_op", I32),
        ("from_col", U32),
        ("from_row_offset_emu_op", I32),
        ("from_row_offset_emu", I64),
        ("from_col_offset_emu_op", I32),
        ("from_col_offset_emu", I64),
        ("to_row_op", I32),
        ("to_row", U32),
        ("to_col_op", I32),
        ("to_col", U32),
        ("to_row_offset_emu_op", I32),
        ("to_row_offset_emu", I64),
        ("to_col_offset_emu_op", I32),
        ("to_col_offset_emu", I64),
        ("width_emu_op", I32),
        ("width_emu", I64),
        ("height_emu_op", I32),
        ("height_emu", I64),
    ],
)

DRAWING_PATCH = Struct(
    "fm_drawing_patch_t",
    [
        ("name_op", I32),
        ("name", PTR),
        ("description_op", I32),
        ("description", PTR),
        ("hidden_op", I32),
        ("hidden", I32),
        ("anchor", nested(DRAWING_ANCHOR_PATCH)),
    ],
)


SHAPE = Struct(
    "fm_shape_info_t",
    [
        ("object_id", U32),
        ("kind", I32),
        ("fill_kind", I32),
        ("fill_color", COLOR_SPEC_BLOB),
        ("line_kind", I32),
        ("line_color", COLOR_SPEC_BLOB),
        ("rotation_deg", F64),
        ("text_box", I32),
        ("line_width_pt", F64),
        ("preset_geometry", PTR),
        ("text", PTR),
    ],
)

SHAPE_SPEC = Struct(
    "fm_shape_spec_t",
    [
        ("placement", nested(DRAWING_PLACEMENT_SPEC)),
        ("fill_kind", I32),
        ("fill_color", COLOR_SPEC_BLOB),
        ("line_kind", I32),
        ("line_color", COLOR_SPEC_BLOB),
        ("line_width_pt", F64),
        ("rotation_deg", F64),
        ("text_box", I32),
        ("preset_geometry", PTR),
        ("text", PTR),
    ],
)

DATA_TABLE = Struct(
    "fm_data_table_t",
    [
        ("rect", nested(CELL_RECT)),
        ("row_input_row", U32),
        ("row_input_col", U32),
        ("column_input_row", U32),
        ("column_input_col", U32),
        ("has_row_input", I32),
        ("has_column_input", I32),
        ("two_dimensional", I32),
        ("row_oriented", I32),
        ("row_input_deleted", I32),
        ("column_input_deleted", I32),
    ],
)

SLICER = Struct(
    "fm_slicer_info_t",
    [
        ("object_id", U32),
        ("column_count", U32),
        ("sort_order", I32),
        ("cross_filter", I32),
        ("row_height_emu", I64),
        ("caption", PTR),
        ("cache_name", PTR),
        ("table_name", PTR),
        ("column_name", PTR),
        ("style", PTR),
    ],
)

SLICER_SPEC = Struct(
    "fm_slicer_spec_t",
    [
        ("placement", nested(DRAWING_PLACEMENT_SPEC)),
        ("column_count", U32),
        ("sort_order", I32),
        ("cross_filter", I32),
        ("row_height_emu", I64),
        ("caption", PTR),
        ("style", PTR),
    ],
)

SPARKLINE = Struct("fm_sparkline_t", [("formula", PTR), ("location", PTR)])

SPARKLINE_GROUP = Struct(
    "fm_sparkline_group_t",
    [
        ("id", PTR),
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
    [
        ("color", COLOR_SPEC_BLOB),
        ("has_transparency_pct", I32),
        ("transparency_pct", F64),
    ],
)

CHART_LINE = Struct(
    "fm_chart_line_t",
    [
        ("color", COLOR_SPEC_BLOB),
        ("has_transparency_pct", I32),
        ("transparency_pct", F64),
        ("has_width_pt", I32),
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
        ("has_size_pt", I32),
        ("size_pt", U32),
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
        ("font_color", COLOR_SPEC_BLOB),
        ("has_font_size_pt", I32),
        ("font_size_pt", F64),
        ("deleted", I32),
    ],
)

CHART_INFO = Struct(
    "fm_chart_info_t",
    [
        ("object_id", U32),
        ("is_extended", I32),
        ("is_pivot_chart", I32),
        ("on_chartsheet", I32),
        ("title", nested(CHART_TITLE)),
        ("auto_title_deleted", I32),
        ("legend_position", I32),
        ("legend_overlay", I32),
        ("legend_hidden_count", U32),
        ("blanks_as", I32),
        ("plot_visible_only", I32),
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
        ("na_as_blank", I32),
    ],
)

CHART_PLOT = Struct(
    "fm_chart_plot_info_t",
    [
        ("type", I32),
        ("extended_layout", I32),
        ("bar_direction", I32),
        ("has_grouping", I32),
        ("grouping", I32),
        ("vary_colors", I32),
        ("has_gap_width_pct", I32),
        ("gap_width_pct", U32),
        ("has_overlap", I32),
        ("overlap", I32),
        ("has_hole_size", I32),
        ("hole_size", U32),
        ("has_first_slice_angle_deg", I32),
        ("first_slice_angle_deg", U32),
        ("axis_id_count", U32),
        ("series_count", U32),
    ],
)

CHART_DATA_SOURCE = Struct(
    "fm_chart_data_source_t",
    [("kind", I32), ("formula", PTR), ("literal_count", U32)],
)

CHART_SERIES = Struct(
    "fm_chart_series_info_t",
    [
        ("series_id", U32),
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
    "fm_chart_literal_info_t",
    [("idx", U32), ("is_text", I32), ("number", F64), ("text", PTR)],
)

CHART_POINT = Struct(
    "fm_chart_point_info_t",
    [
        ("idx", U32),
        ("fill", nested(CHART_FILL)),
        ("line", nested(CHART_LINE)),
        ("has_marker", I32),
        ("marker", nested(CHART_MARKER)),
    ],
)

CHART_AXIS = Struct(
    "fm_chart_axis_info_t",
    [
        ("id", U32),
        ("kind", I32),
        ("position", I32),
        ("visible", I32),
        ("cross_axis", U32),
        ("crosses", I32),
        ("has_crosses_at", I32),
        ("crosses_at", F64),
        ("reversed", I32),
        ("has_log_base", I32),
        ("log_base", F64),
        ("has_min", I32),
        ("min", F64),
        ("has_max", I32),
        ("max", F64),
        ("has_major_unit", I32),
        ("major_unit", F64),
        ("has_minor_unit", I32),
        ("minor_unit", F64),
        ("number_format", PTR),
        ("source_linked", I32),
        ("title", nested(CHART_TITLE)),
        ("major_gridlines", I32),
        ("minor_gridlines", I32),
        ("tick_label_position", I32),
    ],
)

CHART_PREVIEW_OVERRIDES = Struct(
    "fm_chart_preview_overrides_spec_t",
    [
        ("source_op", I32),
        ("source", PTR),
        ("direction_op", I32),
        ("direction", I32),
        ("header_row_op", I32),
        ("header_row", I32),
        ("header_column_op", I32),
        ("header_column", I32),
        ("blanks_as_op", I32),
        ("blanks_as", I32),
        ("plot_visible_only_op", I32),
        ("plot_visible_only", I32),
        ("na_as_blank_op", I32),
        ("na_as_blank", I32),
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
    "fm_chart_data_series_info_t",
    [
        ("plot_index", U32),
        ("series_id", U32),
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
    "fm_chart_data_points_info_t",
    [("count", U32), ("values", PTR), ("states", PTR), ("x", PTR), ("sizes", PTR)],
)

CHART_AXIS_TITLE_TEXT = Struct(
    "fm_chart_axis_title_text_info_t",
    [("axis_id", U32), ("text", PTR)],
)

CHART_UNRESOLVED = Struct(
    "fm_chart_unresolved_info_t",
    [("series_id", U32), ("source", I32), ("axis_id", U32)],
)

CHART_INSERT = Struct(
    "fm_chart_insert_spec_t",
    [
        ("type", I32),
        ("bar_direction", I32),
        ("grouping", I32),
        ("source", PTR),
        ("direction", I32),
        ("header_row", I32),
        ("header_column", I32),
        ("has_title", I32),
        ("title", nested(CHART_TITLE)),
        ("legend", I32),
        ("placement", nested(DRAWING_PLACEMENT_SPEC)),
    ],
)

CHART_FILL_PATCH = Struct(
    "fm_chart_fill_patch_t",
    [
        ("color_op", I32),
        ("color", COLOR_SPEC_BLOB),
        ("transparency_pct_op", I32),
        ("transparency_pct", F64),
    ],
)

CHART_LINE_PATCH = Struct(
    "fm_chart_line_patch_t",
    [
        ("color_op", I32),
        ("color", COLOR_SPEC_BLOB),
        ("transparency_pct_op", I32),
        ("transparency_pct", F64),
        ("width_pt_op", I32),
        ("width_pt", F64),
        ("dash_op", I32),
        ("dash", I32),
        ("none_op", I32),
        ("none", I32),
    ],
)

CHART_MARKER_PATCH = Struct(
    "fm_chart_marker_patch_t",
    [
        ("symbol_op", I32),
        ("symbol", I32),
        ("size_pt_op", I32),
        ("size_pt", F64),
        ("fill", nested(CHART_FILL_PATCH)),
        ("line", nested(CHART_LINE_PATCH)),
    ],
)

CHART_LABELS_PATCH = Struct(
    "fm_chart_labels_patch_t",
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
        ("font_color", COLOR_SPEC_BLOB),
        ("font_size_pt_op", I32),
        ("font_size_pt", F64),
        ("source_linked_op", I32),
        ("source_linked", I32),
        ("deleted_op", I32),
        ("deleted", I32),
    ],
)

# ``fm_chart_data_source_spec_t``: one series input to write. Embedded by value
# in the series patch; not a record.
CHART_DATA_SOURCE_SPEC = Struct(
    "fm_chart_data_source_spec_t",
    [("kind", I32), ("formula", PTR), ("literal", PTR), ("literal_count", U32)],
)

CHART_SERIES_PATCH = Struct(
    "fm_chart_series_patch_t",
    [
        ("name_op", I32),
        ("name", nested(CHART_DATA_SOURCE_SPEC)),
        ("categories_op", I32),
        ("categories", nested(CHART_DATA_SOURCE_SPEC)),
        ("values_op", I32),
        ("values", nested(CHART_DATA_SOURCE_SPEC)),
        ("x_op", I32),
        ("x", nested(CHART_DATA_SOURCE_SPEC)),
        ("y_op", I32),
        ("y", nested(CHART_DATA_SOURCE_SPEC)),
        ("size_op", I32),
        ("size", nested(CHART_DATA_SOURCE_SPEC)),
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
    "fm_chart_point_patch_t",
    [
        ("fill", nested(CHART_FILL_PATCH)),
        ("line", nested(CHART_LINE_PATCH)),
        ("marker_op", I32),
        ("marker", nested(CHART_MARKER_PATCH)),
    ],
)

CHART_PLOT_PATCH = Struct(
    "fm_chart_plot_patch_t",
    [
        ("gap_width_pct_op", I32),
        ("gap_width_pct", U32),
        ("overlap_op", I32),
        ("overlap", I32),
        ("first_slice_angle_deg_op", I32),
        ("first_slice_angle_deg", U32),
        ("vary_colors_op", I32),
        ("vary_colors", I32),
        ("hole_size_op", I32),
        ("hole_size", U32),
    ],
)

CHART_AXIS_PATCH = Struct(
    "fm_chart_axis_patch_t",
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
        ("crosses_at_op", I32),
        ("crosses_at", F64),
        ("tick_label_position_op", I32),
        ("tick_label_position", I32),
        ("minor_unit_op", I32),
        ("minor_unit", F64),
        ("minor_gridlines_op", I32),
        ("minor_gridlines", I32),
        ("position_op", I32),
        ("position", I32),
        ("source_linked_op", I32),
        ("source_linked", I32),
    ],
)

CHART_LEGEND = Struct(
    "fm_chart_legend_t",
    [
        ("position", I32),
        ("overlay", I32),
        ("hidden_entries", PTR),
        ("hidden_entries_count", U32),
    ],
)

CHART_LEGEND_PATCH = Struct(
    "fm_chart_legend_patch_t",
    [
        ("position_op", I32),
        ("position", I32),
        ("overlay_op", I32),
        ("overlay", I32),
        ("hidden_entries_op", I32),
        ("hidden_entries", PTR),
        ("hidden_entries_count", U32),
    ],
)

CHART_DISPLAY_PATCH = Struct(
    "fm_chart_display_patch_t",
    [
        ("blanks_as_op", I32),
        ("blanks_as", I32),
        ("plot_visible_only_op", I32),
        ("plot_visible_only", I32),
        ("na_as_blank_op", I32),
        ("na_as_blank", I32),
        ("style_op", I32),
        ("style", I32),
    ],
)

# Patch records (an op beside every field) and the iterative-settings read record.

HEADER_FOOTER_PATCH = Struct(
    "fm_header_footer_patch_t",
    [
        ("odd_header_op", I32),
        ("odd_header", PTR),
        ("odd_footer_op", I32),
        ("odd_footer", PTR),
        ("even_header_op", I32),
        ("even_header", PTR),
        ("even_footer_op", I32),
        ("even_footer", PTR),
        ("first_header_op", I32),
        ("first_header", PTR),
        ("first_footer_op", I32),
        ("first_footer", PTR),
        ("different_odd_even_op", I32),
        ("different_odd_even", I32),
        ("different_first_op", I32),
        ("different_first", I32),
        ("scale_with_doc_op", I32),
        ("scale_with_doc", I32),
        ("align_with_margins_op", I32),
        ("align_with_margins", I32),
    ],
)

ITERATIVE_SETTINGS = Struct(
    "fm_iterative_settings_info_t",
    [
        ("enabled", I32),
        ("max_iterations", U32),
        ("max_change", F64),
    ],
)

PAGE_MARGINS_PATCH = Struct(
    "fm_page_margins_patch_t",
    [
        ("left_in_op", I32),
        ("left_in", F64),
        ("right_in_op", I32),
        ("right_in", F64),
        ("top_in_op", I32),
        ("top_in", F64),
        ("bottom_in_op", I32),
        ("bottom_in", F64),
        ("header_in_op", I32),
        ("header_in", F64),
        ("footer_in_op", I32),
        ("footer_in", F64),
    ],
)

PAGE_SETUP_PATCH = Struct(
    "fm_page_setup_patch_t",
    [
        ("orientation_op", I32),
        ("orientation", I32),
        ("paper_size_op", I32),
        ("paper_size", U32),
        ("scale_pct_op", I32),
        ("scale_pct", U32),
        ("fit_to_width_op", I32),
        ("fit_to_width", U32),
        ("fit_to_height_op", I32),
        ("fit_to_height", U32),
        ("fit_to_page_op", I32),
        ("fit_to_page", I32),
    ],
)

PIVOT_CACHE_OPTIONS_PATCH = Struct(
    "fm_pivot_cache_options_patch_t",
    [
        ("save_data_op", I32),
        ("save_data", I32),
        ("missing_items_limit_op", I32),
        ("missing_items_limit", U32),
        ("refresh_on_load_op", I32),
        ("refresh_on_load", I32),
    ],
)

PIVOT_GRAND_TOTALS_PATCH = Struct(
    "fm_pivot_grand_totals_patch_t",
    [
        ("rows_op", I32),
        ("rows", I32),
        ("columns_op", I32),
        ("columns", I32),
    ],
)

PIVOT_OPTIONS_PATCH = Struct(
    "fm_pivot_options_patch_t",
    [
        ("show_error_op", I32),
        ("show_error", I32),
        ("error_caption_op", I32),
        ("error_caption", PTR),
        ("show_missing_op", I32),
        ("show_missing", I32),
        ("missing_caption_op", I32),
        ("missing_caption", PTR),
        ("show_headers_op", I32),
        ("show_headers", I32),
        ("show_drill_op", I32),
        ("show_drill", I32),
        ("show_data_tips_op", I32),
        ("show_data_tips", I32),
        ("grid_drop_zones_op", I32),
        ("grid_drop_zones", I32),
        ("show_empty_row_op", I32),
        ("show_empty_row", I32),
        ("show_empty_column_op", I32),
        ("show_empty_column", I32),
        ("show_data_drop_down_op", I32),
        ("show_data_drop_down", I32),
        ("print_drill_op", I32),
        ("print_drill", I32),
        ("field_print_titles_op", I32),
        ("field_print_titles", I32),
        ("item_print_titles_op", I32),
        ("item_print_titles", I32),
        ("merge_item_op", I32),
        ("merge_item", I32),
        ("indent_op", I32),
        ("indent", U32),
        ("page_over_then_down_op", I32),
        ("page_over_then_down", I32),
        ("page_wrap_op", I32),
        ("page_wrap", U32),
        ("multiple_field_filters_op", I32),
        ("multiple_field_filters", I32),
        ("preserve_formatting_op", I32),
        ("preserve_formatting", I32),
        ("custom_list_sort_op", I32),
        ("custom_list_sort", I32),
        ("alt_text_op", I32),
        ("alt_text", PTR),
        ("alt_text_summary_op", I32),
        ("alt_text_summary", PTR),
        ("grand_total_caption_op", I32),
        ("grand_total_caption", PTR),
        ("auto_formatting_op", I32),
        ("auto_formatting", I32),
        ("enable_drill_op", I32),
        ("enable_drill", I32),
    ],
)

PRINT_OPTIONS_PATCH = Struct(
    "fm_print_options_patch_t",
    [
        ("grid_lines_op", I32),
        ("grid_lines", I32),
        ("headings_op", I32),
        ("headings", I32),
        ("horizontal_centered_op", I32),
        ("horizontal_centered", I32),
        ("vertical_centered_op", I32),
        ("vertical_centered", I32),
    ],
)

PROTECTED_RANGE_PATCH = Struct(
    "fm_protected_range_patch_t",
    [
        ("spin_count_op", I32),
        ("spin_count", U32),
        ("name_op", I32),
        ("name", PTR),
        ("ranges_op", I32),
        ("ranges", PTR),
        ("range_count", U32),
        ("legacy_password_hash_op", I32),
        ("legacy_password_hash", PTR),
        ("algorithm_name_op", I32),
        ("algorithm_name", PTR),
        ("hash_value_op", I32),
        ("hash_value", PTR),
        ("salt_value_op", I32),
        ("salt_value", PTR),
        ("security_descriptor_op", I32),
        ("security_descriptor", PTR),
    ],
)

SHAPE_PATCH = Struct(
    "fm_shape_patch_t",
    [
        ("fill_kind_op", I32),
        ("fill_kind", I32),
        ("fill_color_op", I32),
        ("fill_color", COLOR_SPEC_BLOB),
        ("line_kind_op", I32),
        ("line_kind", I32),
        ("line_color_op", I32),
        ("line_color", COLOR_SPEC_BLOB),
        ("line_width_pt_op", I32),
        ("line_width_pt", F64),
        ("rotation_deg_op", I32),
        ("rotation_deg", F64),
        ("text_box_op", I32),
        ("text_box", I32),
        ("preset_geometry_op", I32),
        ("preset_geometry", PTR),
        ("text_op", I32),
        ("text", PTR),
    ],
)

SHEET_FORMAT_DEFAULTS_PATCH = Struct(
    "fm_sheet_format_defaults_patch_t",
    [
        ("default_column_width_chars_op", I32),
        ("default_column_width_chars", F64),
        ("default_row_height_pt_op", I32),
        ("default_row_height_pt", F64),
        ("base_column_width_chars_op", I32),
        ("base_column_width_chars", F64),
        ("custom_height_op", I32),
        ("custom_height", I32),
        ("zero_height_op", I32),
        ("zero_height", I32),
    ],
)

SHEET_VIEW_PATCH = Struct(
    "fm_sheet_view_patch_t",
    [
        ("zoom_pct_op", I32),
        ("zoom_pct", U32),
        ("freeze_rows_op", I32),
        ("freeze_rows", U32),
        ("freeze_cols_op", I32),
        ("freeze_cols", U32),
        ("show_grid_lines_op", I32),
        ("show_grid_lines", I32),
        ("show_row_column_headers_op", I32),
        ("show_row_column_headers", I32),
        ("show_zeros_op", I32),
        ("show_zeros", I32),
        ("right_to_left_op", I32),
        ("right_to_left", I32),
        ("tab_selected_op", I32),
        ("tab_selected", I32),
        ("view_mode_op", I32),
        ("view_mode", I32),
        ("top_left_row_op", I32),
        ("top_left_row", U32),
        ("top_left_col_op", I32),
        ("top_left_col", U32),
        ("pane_top_left_row_op", I32),
        ("pane_top_left_row", U32),
        ("pane_top_left_col_op", I32),
        ("pane_top_left_col", U32),
        ("split_x_op", I32),
        ("split_x", F64),
        ("split_y_op", I32),
        ("split_y", F64),
        ("show_formulas_op", I32),
        ("show_formulas", I32),
        ("show_outline_symbols_op", I32),
        ("show_outline_symbols", I32),
    ],
)

SLICER_PATCH = Struct(
    "fm_slicer_patch_t",
    [
        ("column_count_op", I32),
        ("column_count", U32),
        ("sort_order_op", I32),
        ("sort_order", I32),
        ("cross_filter_op", I32),
        ("cross_filter", I32),
        ("row_height_emu_op", I32),
        ("row_height_emu", I64),
        ("caption_op", I32),
        ("caption", PTR),
        ("style_op", I32),
        ("style", PTR),
    ],
)

SPARKLINE_GROUP_PATCH = Struct(
    "fm_sparkline_group_patch_t",
    [
        ("color_series_op", I32),
        ("color_series", COLOR_SPEC_BLOB),
        ("color_negative_op", I32),
        ("color_negative", COLOR_SPEC_BLOB),
        ("color_axis_op", I32),
        ("color_axis", COLOR_SPEC_BLOB),
        ("color_markers_op", I32),
        ("color_markers", COLOR_SPEC_BLOB),
        ("color_first_op", I32),
        ("color_first", COLOR_SPEC_BLOB),
        ("color_last_op", I32),
        ("color_last", COLOR_SPEC_BLOB),
        ("color_high_op", I32),
        ("color_high", COLOR_SPEC_BLOB),
        ("color_low_op", I32),
        ("color_low", COLOR_SPEC_BLOB),
        ("manual_min_op", I32),
        ("manual_min", F64),
        ("manual_max_op", I32),
        ("manual_max", F64),
        ("line_weight_op", I32),
        ("line_weight", F64),
        ("type_op", I32),
        ("type", I32),
        ("markers_op", I32),
        ("markers", I32),
        ("high_op", I32),
        ("high", I32),
        ("low_op", I32),
        ("low", I32),
        ("first_op", I32),
        ("first", I32),
        ("last_op", I32),
        ("last", I32),
        ("negative_op", I32),
        ("negative", I32),
        ("display_x_axis_op", I32),
        ("display_x_axis", I32),
        ("display_hidden_op", I32),
        ("display_hidden", I32),
        ("min_axis_type_op", I32),
        ("min_axis_type", I32),
        ("max_axis_type_op", I32),
        ("max_axis_type", I32),
        ("date_axis_op", I32),
        ("date_axis", I32),
        ("empty_cells_as_op", I32),
        ("empty_cells_as", I32),
        ("date_axis_ref_op", I32),
        ("date_axis_ref", PTR),
        ("items_op", I32),
        ("items", PTR),
        ("items_count", U32),
    ],
)

TABLE_PATCH = Struct(
    "fm_table_patch_t",
    [
        ("range_op", I32),
        ("range", nested(CELL_RECT)),
        ("header_row_op", I32),
        ("header_row", I32),
        ("totals_row_op", I32),
        ("totals_row", I32),
    ],
)

TABLE_STYLE_DEFAULTS_PATCH = Struct(
    "fm_table_style_defaults_patch_t",
    [
        ("table_style_op", I32),
        ("table_style", PTR),
        ("pivot_style_op", I32),
        ("pivot_style", PTR),
    ],
)
