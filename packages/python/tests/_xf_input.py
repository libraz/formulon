"""Helper: copy a read-back cell XF record into an input record."""

from __future__ import annotations

from formulon import CellXf, CellXfInput, HorizontalAlign, VerticalAlign


def input_of(xf: CellXf) -> CellXfInput:
    """Copy a read-back record into an input; an unstated attribute is a default with its ``has_*`` off."""
    return CellXfInput(
        font_index=xf.font_index,
        fill_index=xf.fill_index,
        border_index=xf.border_index,
        num_fmt_id=xf.num_fmt_id,
        horizontal_align=HorizontalAlign.GENERAL if xf.horizontal_align is None else xf.horizontal_align,
        vertical_align=VerticalAlign.BOTTOM if xf.vertical_align is None else xf.vertical_align,
        wrap_text=bool(xf.wrap_text),
        justify_last_line=xf.justify_last_line,
        has_horizontal_align=xf.horizontal_align is not None,
        has_vertical_align=xf.vertical_align is not None,
        has_wrap_text=xf.wrap_text is not None,
        has_justify_last_line=xf.justify_last_line is not None,
        has_alignment=any(
            v is not None
            for v in (
                xf.horizontal_align,
                xf.vertical_align,
                xf.wrap_text,
                xf.justify_last_line,
                xf.text_rotation,
                xf.indent,
                xf.shrink_to_fit,
                xf.reading_order,
                xf.relative_indent,
            )
        ),
        has_protection=xf.locked is not None or xf.hidden is not None,
        locked=xf.locked,
        hidden=xf.hidden,
        text_rotation=xf.text_rotation,
        indent=xf.indent,
        relative_indent=xf.relative_indent,
        shrink_to_fit=xf.shrink_to_fit,
        reading_order=xf.reading_order,
        xf_id=xf.xf_id,
        apply_number_format=xf.apply_number_format,
        apply_font=xf.apply_font,
        apply_fill=xf.apply_fill,
        apply_border=xf.apply_border,
        apply_alignment=xf.apply_alignment,
        apply_protection=xf.apply_protection,
        quote_prefix=xf.quote_prefix,
    )
