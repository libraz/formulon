"""Helper: copy a read-back cell XF record into an input record."""

from __future__ import annotations

from formulon import CellXf, CellXfInput


def input_of(xf: CellXf) -> CellXfInput:
    """Copy a read-back record into an input; an unstated attribute stays ``None``."""
    return CellXfInput(
        font_index=xf.font_index,
        fill_index=xf.fill_index,
        border_index=xf.border_index,
        num_fmt_id=xf.num_fmt_id,
        horizontal=xf.horizontal,
        vertical=xf.vertical,
        wrap_text=xf.wrap_text,
        justify_last_line=xf.justify_last_line,
        locked=xf.locked,
        hidden=xf.hidden,
        text_rotation=xf.text_rotation,
        indent=xf.indent,
        relative_indent=xf.relative_indent,
        shrink_to_fit=xf.shrink_to_fit,
        reading_order=xf.reading_order,
        style_xf_index=xf.style_xf_index,
        apply_number_format=xf.apply_number_format,
        apply_font=xf.apply_font,
        apply_fill=xf.apply_fill,
        apply_border=xf.apply_border,
        apply_alignment=xf.apply_alignment,
        apply_protection=xf.apply_protection,
        quote_prefix=xf.quote_prefix,
    )
