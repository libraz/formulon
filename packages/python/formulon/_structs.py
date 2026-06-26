# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

        Opaque ``blob16`` fields (inline ``fm_value_t``) are skipped; read
        them separately via ``Value._from_wasm(ptr + self.offsets[name][1])``.
        """
        raw = lib.read_bytes(ptr, self.size)
        out: Dict[str, int] = {}
        for fname, (kind, off) in self.offsets.items():
            if kind == "blob16":
                continue
            out[fname] = struct.unpack_from(_FMT[kind], raw, off)[0]
        return out


def alloc_struct(lib, layout: Struct) -> int:
    """Allocate a zeroed scratch block sized for ``layout``."""
    ptr = lib.alloc(layout.size)
    lib.write_bytes(ptr, b"\x00" * layout.size)
    return ptr


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
    ],
)

CELL_NODE = Struct(
    "fm_cell_node_t",
    [("sheet", U32), ("row", U32), ("col", U32)],
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

SHEET_VIEW = Struct(
    "fm_sheet_view_t",
    [
        ("zoom_scale", U32),
        ("freeze_rows", U32),
        ("freeze_cols", U32),
        ("tab_hidden", I32),
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
    ],
)

ROW_LAYOUT = Struct(
    "fm_row_layout_t",
    [
        ("row", U32),
        ("height", F64),
        ("hidden", I32),
        ("outline_level", U8),
    ],
)

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
    ],
)

FONT_RECORD = Struct(
    "fm_font_record",
    [
        ("name", PTR),
        ("size", F64),
        ("color_argb", U32),
        ("bold", I32),
        ("italic", I32),
        ("strike", I32),
        ("underline", U8),
    ],
)

FILL_RECORD = Struct(
    "fm_fill_record",
    [("pattern", U8), ("fg_argb", U32), ("bg_argb", U32)],
)

BORDER_RECORD = Struct(
    "fm_border_record",
    [
        ("left_style", U8),
        ("left_color_argb", U32),
        ("right_style", U8),
        ("right_color_argb", U32),
        ("top_style", U8),
        ("top_color_argb", U32),
        ("bottom_style", U8),
        ("bottom_color_argb", U32),
        ("diagonal_style", U8),
        ("diagonal_color_argb", U32),
        ("diagonal_up", I32),
        ("diagonal_down", I32),
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

CELL_STYLE_BUILTIN_ID_NONE = 0xFFFFFFFF


def write_str_field(lib, struct_ptr: int, layout: Struct, field: str,
                    value: Optional[str], owned: List[int]) -> None:
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
