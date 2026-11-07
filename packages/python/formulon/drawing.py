"""Handles on the anchored objects of a sheet's drawing: :class:`Drawing` and its kinds.

A :class:`~formulon.Sheet` hands out one handle per drawing object id. Each is
an instance of the subclass its kind selects -- :class:`Picture`,
:class:`Chart`, :class:`Shape` (shapes, connectors and groups) or
:class:`Slicer` (a table slicer's frame) -- and a :class:`Drawing` itself for
any other object. A handle names its object by id: once the object is removed
every member but :attr:`~Drawing.sheet`, :attr:`~Drawing.object_id` and
:attr:`~Drawing.kind` fails with :class:`~formulon.NotFoundError`, and a
:meth:`~formulon.Sheet.restore_drawing` that puts the same id back makes the
handle address it again.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Dict, Literal, Optional, Sequence, Tuple, Type, Union, cast

from . import _structs as S
from ._c import LIB, FormulonError, _sint, _uint
from ._status import StatusCode
from .workbook import (
    _CHART_TITLE_NONE,
    _DISPLAY_PATCH_SPEC,
    _PLOT_PATCH_SPEC,
    _PREVIEW_OVERRIDE_SPEC,
    UNCHANGED,
    AnchorEditAs,
    AnchorKind,
    ChartAxisCrosses,
    ChartAxisPatch,
    ChartBlanksAs,
    ChartData,
    ChartDisplayPatch,
    ChartFillPatch,
    ChartGrouping,
    ChartInsertKind,
    ChartLabelsPatch,
    ChartLegendPosition,
    ChartLinePatch,
    ChartMarkerPatch,
    ChartPlotPatch,
    ChartPointPatch,
    ChartPreviewOverridesInput,
    ChartSeriesDirection,
    ChartSeriesPatch,
    ChartTitle,
    DrawingObjectKind,
    ImageFormat,
    ShapePaint,
    SlicerCrossFilter,
    SlicerSortOrder,
    _ArgScope,
    _check,
    _encode_axis_patch,
    _encode_patch,
    _encode_point_patch,
    _encode_series_patch,
    _enum,
    _header_mode,
    _id_array,
    _image_anchor_fields,
    _install_workbook_locks,
    _nullable_enum,
    _pack_shape,
    _pack_slicer,
    _read_chart,
    _read_chart_data,
    _read_count,
    _title_fields,
)

if TYPE_CHECKING:
    from .sheet import Sheet
    from .workbook import (
        ChartProperties,
        DrawingProperties,
        PictureProperties,
        ShapeProperties,
        SlicerProperties,
    )

_SHAPE_KINDS = frozenset({DrawingObjectKind.SHAPE, DrawingObjectKind.CONNECTOR, DrawingObjectKind.GROUP})


def _given(arguments: Dict[str, Any], patch: type) -> Dict[str, Any]:
    """The keyword arguments among the keys of the patch ``TypedDict`` that are not :data:`UNCHANGED`."""
    return {key: arguments[key] for key in patch.__annotations__ if arguments[key] is not UNCHANGED}


def _removed() -> FormulonError:
    """The error a handle raises when the engine no longer has its object."""
    return FormulonError(StatusCode.NOT_FOUND, _diagnostic_override=("the drawing was removed", ""))


def _raw_kind(sheet: Sheet, object_id: int) -> Optional[int]:
    """The raw kind of drawing object ``object_id`` of ``sheet``, or ``None`` when its drawing has none."""
    return sheet._workbook._lookup(
        LIB.fm_sheet_get_drawing_object,
        (_uint(sheet._index(), "sheet_index"), _uint(object_id, "object_id")),
        S.DRAWING_OBJECT,
        lambda d: d["kind"],
    )


def _has_slicer(sheet: Sheet, object_id: int) -> bool:
    """Whether drawing object ``object_id`` of ``sheet`` is the frame of a table slicer."""
    found = sheet._workbook._lookup(
        LIB.fm_sheet_get_slicer,
        (_uint(sheet._index(), "sheet_index"), _uint(object_id, "object_id")),
        S.SLICER,
        lambda d: True,
    )
    return bool(found)


def _drawing_for(sheet: Sheet, object_id: int) -> Drawing:
    """Return the handle of drawing object ``object_id`` of ``sheet``, of the class its kind selects.

    A picture is a :class:`Picture`, a chart a :class:`Chart`, a shape,
    connector or group a :class:`Shape`, a table slicer's frame a
    :class:`Slicer`, and any other object a :class:`Drawing`. The sheet keeps
    one handle per object id in ``sheet._drawings``: a handle made earlier
    for the id is returned again, so every lookup of one object yields the
    same handle, and a new one is recorded there.

    Raises:
      NotFoundError: when the sheet's drawing has no object ``object_id``, or
        the sheet was removed.
    """
    raw = _raw_kind(sheet, object_id)
    if raw is None:
        raise FormulonError(
            StatusCode.NOT_FOUND, _diagnostic_override=(f"the sheet's drawing has no object {object_id}", "")
        )
    handles: Dict[int, Drawing] = sheet._drawings
    handle = handles.get(object_id)
    if handle is None:
        kind = _enum(DrawingObjectKind, raw)
        cls: Type[Drawing]
        if kind == DrawingObjectKind.PICTURE:
            cls = Picture
        elif kind == DrawingObjectKind.CHART:
            cls = Chart
        elif kind in _SHAPE_KINDS:
            cls = Shape
        elif _has_slicer(sheet, object_id):
            cls = Slicer
        else:
            cls = Drawing
        handle = cls(sheet, object_id, kind)
        handles[object_id] = handle
    return handle


class Drawing:
    """One anchored object of a sheet's drawing, named by its object id.

    Get it from :meth:`Sheet.drawing` or
    :meth:`Sheet.list_drawings`; each is an
    instance of the subclass its kind selects (:class:`Picture`,
    :class:`Chart`, :class:`Shape`, :class:`Slicer`), and a :class:`Drawing`
    itself for any other object. The sheet keeps one handle per object id, so
    every lookup returns the same object (compare with ``is``). Once the object
    is removed, every member but :attr:`sheet`, :attr:`object_id` and
    :attr:`kind` fails with :class:`~formulon.NotFoundError`; a
    :meth:`~formulon.Sheet.restore_drawing` that puts the same id back makes
    the handle address it again. Remove an object with
    :meth:`Sheet.remove_drawing`.
    """

    __slots__ = ("_sheet", "_object_id", "_kind")

    def __init__(self, sheet: Sheet, object_id: int, kind: DrawingObjectKind) -> None:
        self._sheet = sheet
        self._object_id = object_id
        self._kind = kind

    @property
    def sheet(self) -> Sheet:
        """The sheet whose drawing holds the object."""
        return self._sheet

    @property
    def object_id(self) -> int:
        """The object's id in the sheet's drawing."""
        return self._object_id

    @property
    def kind(self) -> DrawingObjectKind:
        """The object's kind."""
        return self._kind

    def _at(self) -> Tuple[int, int, int]:
        """``(workbook handle, sheet index, object id)`` for a raw call.

        Raises ``NotFoundError`` once the sheet or the object is removed; the raw
        calls refuse a missing object id as an invalid argument instead.
        """
        if _raw_kind(self._sheet, self._object_id) is None:
            raise _removed()
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        return h, _uint(sheet, "sheet_index"), _uint(self._object_id, "object_id")

    def __repr__(self) -> str:
        return f"{type(self).__name__}(object_id={self._object_id})"

    def get_drawing_properties(self) -> DrawingProperties:
        """Return the object's kind, anchor, name and description.

        Raises:
          NotFoundError: when the object or its sheet was removed.
        """
        from .workbook import _drawing_object_from

        _, sheet, object_id = self._at()
        found = self._sheet._workbook._lookup(
            LIB.fm_sheet_get_drawing_object, (sheet, object_id), S.DRAWING_OBJECT, _drawing_object_from
        )
        if found is None:
            raise _removed()
        return cast("DrawingProperties", found)


def _rename(handle: Drawing, name: str) -> None:
    """Rename the picture or chart ``handle``."""
    h, sheet, object_id = handle._at()
    with _ArgScope() as scope:
        name_ptr = scope.cstr(name)
        _check(LIB.fm_sheet_set_drawing_name(h, sheet, object_id, name_ptr))


def _set_anchor(handle: Drawing, anchor: Tuple[Any, ...]) -> None:
    """Move and resize the picture or chart ``handle`` to the :func:`_image_anchor_fields` arguments ``anchor``."""
    h, sheet, object_id = handle._at()
    with _ArgScope() as scope:
        ptr = scope.struct_array(S.DRAWING_ANCHOR, 1)
        S.DRAWING_ANCHOR.pack(LIB, ptr, _image_anchor_fields(*anchor))
        _check(LIB.fm_sheet_set_drawing_anchor(h, sheet, object_id, ptr))


def _set_z_order(handle: Drawing, z_order_index: int) -> None:
    """Move the picture or chart ``handle`` to item ``z_order_index`` of the stacking order."""
    h, sheet, object_id = handle._at()
    _check(LIB.fm_sheet_set_drawing_z_order(h, sheet, object_id, _uint(z_order_index, "z_order_index")))


def _snapshot(handle: Drawing) -> bytes:
    """Capture the picture or chart ``handle`` as opaque bytes for ``Sheet.restore_drawing``."""
    h, sheet, object_id = handle._at()
    with _ArgScope() as scope:
        bytes_ptr = scope.out_ptr()
        len_ptr = scope.out_ptr()
        data_ptr = 0
        try:
            _check(LIB.fm_sheet_snapshot_drawing(h, sheet, object_id, bytes_ptr, len_ptr))
            data_ptr = LIB.read_u32(bytes_ptr)
            return LIB.read_bytes(data_ptr, LIB.read_u32(len_ptr))
        finally:
            if data_ptr:
                LIB.fm_buffer_destroy(data_ptr)


class Picture(Drawing):
    """A picture of a sheet's drawing; also has the members of :class:`Drawing`.

    Add one with :meth:`Sheet.add_picture`.
    """

    __slots__ = ()

    def get_properties(self) -> PictureProperties:
        """Return the picture's format, pixel size and a copy of its media bytes.

        Raises:
          NotFoundError: when the picture or its sheet was removed, or the
            picture's media part is missing.
        """
        from .workbook import PictureProperties

        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            bytes_ptr = scope.out_ptr()
            len_ptr = scope.out_ptr()
            info = scope.struct(S.IMAGE_INFO)
            _check(LIB.fm_sheet_get_image(h, sheet, object_id, present, bytes_ptr, len_ptr, info))
            if not LIB.read_u32(present):
                raise FormulonError(
                    StatusCode.NOT_FOUND, _diagnostic_override=("the picture's media part is missing", "")
                )
            d = S.IMAGE_INFO.unpack(LIB, info)
            return PictureProperties(
                format=_enum(ImageFormat, d["format"]),
                px_width=d["px_width"],
                px_height=d["px_height"],
                bytes=LIB.read_bytes(LIB.read_u32(bytes_ptr), LIB.read_u32(len_ptr)),
            )

    def rename(self, name: str) -> None:
        """Rename the picture. Duplicate names are allowed; ``name`` must be non-empty."""
        _rename(self, name)

    def set_anchor(
        self,
        *,
        anchor_kind: AnchorKind = AnchorKind.ONE_CELL,
        edit_as: AnchorEditAs = AnchorEditAs.TWO_CELL,
        row: int = 0,
        col: int = 0,
        row_offset_emu: int = 0,
        col_offset_emu: int = 0,
        width_emu: int = 0,
        height_emu: int = 0,
    ) -> None:
        """Move and resize the picture, keeping its id and every other element.

        The placement is in the quantities :meth:`get_drawing_properties`
        reports: ``(row, col)`` and the offsets are the top-left marker of the
        anchor box, ``width_emu`` and ``height_emu`` the object's own unrotated
        size; ``0`` keeps the current size. ``edit_as`` is
        honoured for a two-cell anchor only.
        """
        _set_anchor(self, (anchor_kind, edit_as, row, col, row_offset_emu, col_offset_emu, width_emu, height_emu))

    def set_z_order(self, z_order_index: int) -> None:
        """Make the picture item ``z_order_index`` of :meth:`Sheet.list_drawings`.

        That list's order is the stacking order; ``0`` is the back.
        """
        _set_z_order(self, z_order_index)

    def snapshot(self) -> bytes:
        """Capture the picture as opaque bytes for :meth:`Sheet.restore_drawing`.

        The capture holds the picture's list position, element and every part it reaches.
        """
        return _snapshot(self)


class Chart(Drawing):
    """A chart of a sheet's drawing; also has the members of :class:`Drawing`.

    Add one with :meth:`Sheet.add_chart`.
    """

    __slots__ = ()

    def get_properties(self) -> ChartProperties:
        """Return the typed view of the chart: plots, series, axes, legend and source.

        Raises:
          NotFoundError: when the chart or its sheet was removed.
          FormulonError: when the chart part does not parse.
        """
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            out = scope.out_ptr()
            _check(LIB.fm_sheet_get_chart(h, sheet, object_id, present, out))
            if not LIB.read_u32(present):
                raise _removed()
            chart = LIB.read_u32(out)
        try:
            return cast("ChartProperties", _read_chart(chart))
        finally:
            LIB.fm_chart_destroy(chart)

    def rename(self, name: str) -> None:
        """Rename the chart. Duplicate names are allowed; ``name`` must be non-empty."""
        _rename(self, name)

    def set_anchor(
        self,
        *,
        anchor_kind: AnchorKind = AnchorKind.ONE_CELL,
        edit_as: AnchorEditAs = AnchorEditAs.TWO_CELL,
        row: int = 0,
        col: int = 0,
        row_offset_emu: int = 0,
        col_offset_emu: int = 0,
        width_emu: int = 0,
        height_emu: int = 0,
    ) -> None:
        """Move and resize the chart frame, keeping its id and every other element.

        The placement is in the quantities :meth:`get_drawing_properties`
        reports: ``(row, col)`` and the offsets are the top-left marker of the
        anchor box, ``width_emu`` and ``height_emu`` the object's own unrotated
        size; ``0`` keeps the current size. ``edit_as`` is
        honoured for a two-cell anchor only.
        """
        _set_anchor(self, (anchor_kind, edit_as, row, col, row_offset_emu, col_offset_emu, width_emu, height_emu))

    def set_z_order(self, z_order_index: int) -> None:
        """Make the chart item ``z_order_index`` of :meth:`Sheet.list_drawings`.

        That list's order is the stacking order; ``0`` is the back.
        """
        _set_z_order(self, z_order_index)

    def snapshot(self) -> bytes:
        """Capture the chart as opaque bytes for :meth:`Sheet.restore_drawing`.

        The capture holds the chart's list position, element and every part it reaches.
        """
        return _snapshot(self)

    def get_data(self) -> ChartData:
        """Resolve the chart's series against the current cell values by Excel's rules.

        Raises:
          NotFoundError: when the chart or its sheet was removed.
        """
        h, sheet, object_id = self._at()
        data = _read_count(LIB.fm_sheet_get_chart_data, h, sheet, object_id)
        try:
            return _read_chart_data(data)
        finally:
            LIB.fm_chart_data_destroy(data)

    def get_preview(self, overrides: Optional[ChartPreviewOverridesInput] = None) -> ChartData:
        """Resolve the data of a copy of the chart with ``overrides`` applied; the workbook is unchanged.

        Series are rebuilt as :meth:`set_source` does: over the overridden
        ``source`` (an unset direction or header is auto), else over the
        chart's derived range.

        Raises:
          NotFoundError: when the chart or its sheet was removed.
        """
        h, sheet, object_id = self._at()
        given = {k: v for k, v in vars(overrides).items() if v is not None} if overrides is not None else {}
        with _ArgScope() as scope:
            fields = _encode_patch(_PREVIEW_OVERRIDE_SPEC, given, ("source",), "preview override", scope)
            if "source" in given:
                fields["source"] = scope.cstr(given["source"])
            ptr = scope.struct(S.CHART_PREVIEW_OVERRIDES)
            S.CHART_PREVIEW_OVERRIDES.pack(LIB, ptr, fields)
            data = _read_count(LIB.fm_sheet_get_chart_preview, h, sheet, object_id, ptr)
        try:
            return _read_chart_data(data)
        finally:
            if data:
                LIB.fm_chart_data_destroy(data)

    def set_type(self, plot: int, kind: ChartInsertKind, grouping: ChartGrouping = ChartGrouping.CLUSTERED) -> None:
        """Replace plot group ``plot`` with ``kind``, keeping each series' references, idx, order and explicit colors.

        ``grouping`` applies to column, bar, line and area plots.
        """
        h, sheet, object_id = self._at()
        _check(
            LIB.fm_sheet_set_chart_type(
                h, sheet, object_id, _uint(plot, "plot"), _sint(kind, "kind"), _sint(grouping, "grouping")
            )
        )

    def set_title(self, title: Union[str, None, ChartTitle]) -> None:
        """Set the chart title to ``title``; ``None`` removes it.

        A :class:`~formulon.ChartTitle` sets that kind, so
        ``ChartTitle(ChartTitleKind.AUTO, "", "")`` returns the title to
        Excel's automatic one.
        """
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_TITLE, 1)
            fields = {"kind": _sint(_CHART_TITLE_NONE, "kind")} if title is None else _title_fields(title, scope)
            S.CHART_TITLE.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_set_chart_title(h, sheet, object_id, ptr))

    def set_legend(
        self,
        legend: None = UNCHANGED,
        /,
        *,
        position: ChartLegendPosition = UNCHANGED,
        overlay: bool = UNCHANGED,
        hidden_series: Sequence[int] = UNCHANGED,
    ) -> None:
        """Update the chart legend; an omitted argument keeps its value.

        ``set_legend(None)`` removes the legend. Without a legend, one is
        created and an omitted argument takes Excel's default
        (:attr:`ChartLegendPosition.RIGHT`, no overlay, no hidden entries).
        ``hidden_series`` replaces the hidden legend entries (series ``c:idx``
        values); ``[]`` shows them all.

        Raises:
          TypeError: when ``position``, ``overlay`` or ``hidden_series`` is
            ``None``, when the first argument is not ``None``, or when it is
            ``None`` together with a keyword argument.
        """
        given = {
            key: value
            for key, value in (("position", position), ("overlay", overlay), ("hidden_series", hidden_series))
            if value is not UNCHANGED
        }
        if legend is not UNCHANGED:
            if legend is not None:
                raise TypeError(f"formulon: the legend argument removes the legend and must be None, got {legend!r}")
            if given:
                raise TypeError("formulon: removing the legend takes no keyword arguments")
        for key, value in given.items():
            if value is None:
                raise TypeError(f"formulon: {key} cannot be None")
        if overlay is not UNCHANGED and not isinstance(overlay, bool):
            raise TypeError(f"formulon: overlay must be a bool, got {type(overlay).__name__}")
        if legend is UNCHANGED and not given:
            return
        h, sheet, object_id = self._at()
        fields: Dict[str, object]
        if legend is None:
            fields = {"position": _sint(_nullable_enum(None, "position"), "position"), "overlay": 0}
        else:
            current = self.get_properties().legend
            if position is UNCHANGED:
                position = current.position if current is not None else ChartLegendPosition.RIGHT
            if overlay is UNCHANGED:
                overlay = current.overlay if current is not None else False
            fields = {
                "position": _sint(_nullable_enum(position, "position"), "position"),
                "overlay": 1 if overlay else 0,
            }
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_LEGEND_PATCH, 1)
            if hidden_series is not UNCHANGED:
                fields["hidden_series_op"] = 2
                fields["hidden_series"] = _id_array(hidden_series, "hidden_series", scope)
                fields["hidden_series_count"] = _uint(len(hidden_series), "hidden_series_count")
            S.CHART_LEGEND_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_set_chart_legend(h, sheet, object_id, ptr))

    def set_axis(
        self,
        axis_id: int,
        *,
        visible: Optional[bool] = UNCHANGED,
        title: Union[str, ChartTitle, None] = UNCHANGED,
        min: Optional[float] = UNCHANGED,
        max: Optional[float] = UNCHANGED,
        major_unit: Optional[float] = UNCHANGED,
        number_format: Optional[str] = UNCHANGED,
        major_gridlines: Optional[bool] = UNCHANGED,
        reversed: Optional[bool] = UNCHANGED,
        log_base: Optional[float] = UNCHANGED,
        crosses: Optional[ChartAxisCrosses] = UNCHANGED,
        crosses_at: Optional[float] = UNCHANGED,
        tick_label_position: Optional[Literal["nextTo", "low", "high", "none"]] = UNCHANGED,
    ) -> None:
        """Update axis ``axis_id``; an omitted argument keeps its value, ``None`` removes the element.

        ``title`` takes the forms of :meth:`set_title`. ``None`` on ``title``
        removes the axis title, on ``log_base`` makes the axis linear and on
        ``tick_label_position`` (``nextTo``, ``low``, ``high``, ``none``)
        restores the default. ``crosses_at`` is read when ``crosses`` is
        ``AT``.
        """
        patch = _given(dict(locals()), ChartAxisPatch)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_AXIS_PATCH, 1)
            fields = _encode_axis_patch(patch, scope)
            S.CHART_AXIS_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_set_chart_axis(h, sheet, object_id, _uint(axis_id, "axis_id"), ptr))

    def set_display(
        self,
        *,
        blanks_as: Optional[ChartBlanksAs] = UNCHANGED,
        plot_visible_only: Optional[bool] = UNCHANGED,
        na_as_blank: Optional[bool] = UNCHANGED,
    ) -> None:
        """Update how blanks plot, whether hidden cells are skipped and whether ``#N/A`` plots as blank.

        An omitted argument keeps its value, ``None`` removes the element.
        """
        patch = _given(dict(locals()), ChartDisplayPatch)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            fields = _encode_patch(_DISPLAY_PATCH_SPEC, patch, (), "display patch", scope)
            ptr = scope.struct(S.CHART_DISPLAY_PATCH)
            S.CHART_DISPLAY_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_set_chart_display(h, sheet, object_id, ptr))

    def set_series(
        self,
        plot: int,
        series_index: int,
        *,
        name: Optional[str] = UNCHANGED,
        categories: Optional[str] = UNCHANGED,
        values: Optional[str] = UNCHANGED,
        x: Optional[str] = UNCHANGED,
        y: Optional[str] = UNCHANGED,
        size: Optional[str] = UNCHANGED,
        fill: Optional[ChartFillPatch] = UNCHANGED,
        line: Optional[ChartLinePatch] = UNCHANGED,
        marker: Optional[ChartMarkerPatch] = UNCHANGED,
        smooth: Optional[bool] = UNCHANGED,
        labels: Optional[ChartLabelsPatch] = UNCHANGED,
    ) -> None:
        """Update the series with ``c:idx`` ``series_index`` of plot group ``plot``; an omitted argument keeps its value.

        A reference is its new text; ``""`` removes ``name``, ``categories`` or
        ``x`` (an empty ``values``, ``y`` or ``size`` is rejected). ``None`` on
        ``marker`` or ``labels`` removes it. A changed reference refreshes its cache.
        """
        patch = _given(dict(locals()), ChartSeriesPatch)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_SERIES_PATCH, 1)
            fields = _encode_series_patch(patch, scope)
            S.CHART_SERIES_PATCH.pack(LIB, ptr, fields)
            _check(
                LIB.fm_sheet_set_chart_series(
                    h, sheet, object_id, _uint(plot, "plot"), _uint(series_index, "series_index"), ptr
                )
            )

    def set_point(
        self,
        plot: int,
        series_index: int,
        point_index: int,
        *,
        fill: Optional[ChartFillPatch] = UNCHANGED,
        line: Optional[ChartLinePatch] = UNCHANGED,
        marker: Optional[ChartMarkerPatch] = UNCHANGED,
    ) -> None:
        """Update point ``point_index`` of a series; an omitted argument keeps its value.

        A call with no keyword argument removes the point's formatting.
        """
        patch = _given(dict(locals()), ChartPointPatch)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_POINT_PATCH, 1)
            fields = _encode_point_patch(patch, scope)
            S.CHART_POINT_PATCH.pack(LIB, ptr, fields)
            _check(
                LIB.fm_sheet_set_chart_point(
                    h,
                    sheet,
                    object_id,
                    _uint(plot, "plot"),
                    _uint(series_index, "series_index"),
                    _uint(point_index, "point_index"),
                    ptr,
                )
            )

    def set_plot(
        self,
        plot: int,
        *,
        gap_width: Optional[int] = UNCHANGED,
        overlap: Optional[int] = UNCHANGED,
        first_slice_angle: Optional[int] = UNCHANGED,
        vary_colors: Optional[bool] = UNCHANGED,
        hole_size: Optional[int] = UNCHANGED,
    ) -> None:
        """Update plot group ``plot``; an omitted argument keeps its value, ``None`` removes the element.

        An argument the plot type does not have is rejected. ``hole_size`` is
        10..90 (doughnut only).
        """
        patch = _given(dict(locals()), ChartPlotPatch)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_PLOT_PATCH, 1)
            fields = _encode_patch(_PLOT_PATCH_SPEC, patch, (), "plot patch", scope)
            S.CHART_PLOT_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_set_chart_plot(h, sheet, object_id, _uint(plot, "plot"), ptr))

    def set_filter(self, hidden_series: Sequence[int], hidden_categories: Sequence[int]) -> None:
        """Replace the chart filter: the listed series (by ``c:idx``) and category positions are hidden."""
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            _check(
                LIB.fm_sheet_set_chart_filter(
                    h,
                    sheet,
                    object_id,
                    _id_array(hidden_series, "hidden_series", scope),
                    _uint(len(hidden_series), "hidden_series_count"),
                    _id_array(hidden_categories, "hidden_categories", scope),
                    _uint(len(hidden_categories), "hidden_category_count"),
                )
            )

    def add_series(
        self,
        plot: int,
        *,
        name: Optional[str] = UNCHANGED,
        categories: Optional[str] = UNCHANGED,
        values: Optional[str] = UNCHANGED,
        x: Optional[str] = UNCHANGED,
        y: Optional[str] = UNCHANGED,
        size: Optional[str] = UNCHANGED,
        fill: Optional[ChartFillPatch] = UNCHANGED,
        line: Optional[ChartLinePatch] = UNCHANGED,
        marker: Optional[ChartMarkerPatch] = UNCHANGED,
        smooth: Optional[bool] = UNCHANGED,
        labels: Optional[ChartLabelsPatch] = UNCHANGED,
    ) -> int:
        """Append a series to plot group ``plot`` and return its ``c:idx``.

        ``values`` is required (``y`` for a scatter plot); the arguments read
        as in :meth:`set_series`.
        """
        patch = _given(dict(locals()), ChartSeriesPatch)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = scope.struct_array(S.CHART_SERIES_PATCH, 1)
            fields = _encode_series_patch(patch, scope)
            S.CHART_SERIES_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_add_chart_series(h, sheet, object_id, _uint(plot, "plot"), ptr, out))
            return LIB.read_u32(out)

    def remove_series(self, plot: int, series_index: int) -> None:
        """Remove the series with ``c:idx`` ``series_index`` from plot group ``plot``."""
        h, sheet, object_id = self._at()
        _check(
            LIB.fm_sheet_remove_chart_series(
                h, sheet, object_id, _uint(plot, "plot"), _uint(series_index, "series_index")
            )
        )

    def set_source(
        self,
        source: str,
        *,
        direction: ChartSeriesDirection = ChartSeriesDirection.AUTO,
        header_row: Optional[bool] = None,
        header_column: Optional[bool] = None,
    ) -> None:
        """Rebuild every series from ``source`` by the insert splitting rules.

        The chart kind and each series' formatting (by idx) are kept and the
        filter is cleared. ``header_row`` / ``header_column`` ``None`` follow
        Excel's detection.
        """
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            _check(
                LIB.fm_sheet_set_chart_source(
                    h,
                    sheet,
                    object_id,
                    scope.cstr(source),
                    _sint(direction, "direction"),
                    _sint(_header_mode(header_row), "header_row"),
                    _sint(_header_mode(header_column), "header_column"),
                )
            )


class Shape(Drawing):
    """A shape, connector or group of a sheet's drawing; also has the members of :class:`Drawing`.

    Add one with :meth:`Sheet.add_shape`;
    removing a group removes its children.
    """

    __slots__ = ()

    def get_properties(self) -> ShapeProperties:
        """Return the shape's anchor, geometry, text and paint.

        Raises:
          NotFoundError: when the shape or its sheet was removed.
        """
        from .workbook import _shape_from

        _, sheet, object_id = self._at()
        found = self._sheet._workbook._lookup(LIB.fm_sheet_get_shape, (sheet, object_id), S.SHAPE, _shape_from)
        if found is None:
            raise _removed()
        return cast("ShapeProperties", found)

    def update(
        self,
        *,
        anchor_kind: Optional[AnchorKind] = None,
        edit_as: AnchorEditAs = AnchorEditAs.TWO_CELL,
        from_row: int = 0,
        from_col: int = 0,
        from_row_offset_emu: int = 0,
        from_col_offset_emu: int = 0,
        to_row: int = 0,
        to_col: int = 0,
        to_row_offset_emu: int = 0,
        to_col_offset_emu: int = 0,
        width_emu: int = 0,
        height_emu: int = 0,
        name: Optional[str] = None,
        preset_geometry: Optional[str] = None,
        text: Optional[str] = None,
        text_box: Optional[bool] = None,
        fill_kind: Optional[ShapePaint] = None,
        fill_rgb: int = 0,
        line_kind: Optional[ShapePaint] = None,
        line_rgb: int = 0,
        line_width_emu: Optional[int] = None,
        rotation: Optional[int] = None,
        hidden: Optional[bool] = None,
    ) -> None:
        """Apply the arguments that are not ``None`` to the shape, the name and anchor included; the rest of its XML is kept.

        The anchor is applied as a whole when ``anchor_kind`` is given, and its
        kind cannot change. Text applies to a shape only and keeps the first
        paragraph's properties and first run's formatting.
        """
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = _pack_shape(
                scope, name, preset_geometry, text, text_box, anchor_kind, edit_as, from_row, from_col, from_row_offset_emu,
                from_col_offset_emu, to_row, to_col, to_row_offset_emu, to_col_offset_emu, width_emu, height_emu, fill_kind, fill_rgb, line_kind,
                line_rgb, line_width_emu, rotation, hidden,
            )  # fmt: skip
            _check(LIB.fm_sheet_update_shape(h, sheet, object_id, ptr))


class Slicer(Drawing):
    """A table slicer, through its drawing frame; also has the members of :class:`Drawing`.

    Add one with :meth:`Sheet.add_slicer`.
    Removing its table or sheet removes the slicer.
    """

    __slots__ = ()

    def get_properties(self) -> SlicerProperties:
        """Return the slicer's caption, cache, table column, layout and anchor.

        Raises:
          NotFoundError: when the slicer or its sheet was removed.
        """
        from .workbook import _slicer_from

        _, sheet, object_id = self._at()
        found = self._sheet._workbook._lookup(LIB.fm_sheet_get_slicer, (sheet, object_id), S.SLICER, _slicer_from)
        if found is None:
            raise _removed()
        return cast("SlicerProperties", found)

    def update(
        self,
        *,
        from_row: Optional[int] = None,
        from_col: Optional[int] = None,
        to_row: Optional[int] = None,
        to_col: Optional[int] = None,
        from_row_offset_emu: int = 0,
        from_col_offset_emu: int = 0,
        to_row_offset_emu: int = 0,
        to_col_offset_emu: int = 0,
        caption: Optional[str] = None,
        style: Optional[str] = None,
        column_count: Optional[int] = None,
        row_height_emu: Optional[int] = None,
        sort_order: Optional[SlicerSortOrder] = None,
        cross_filter: Optional[SlicerCrossFilter] = None,
    ) -> None:
        """Apply the arguments that are not ``None`` to the slicer, the anchor included; the rest of its XML is kept.

        An empty ``style`` removes the style. The four ``from_*`` / ``to_*``
        cells, given together, move the slicer's two-cell frame (0-based cells
        plus EMU offsets); the EMU offsets apply only with them.
        """
        corners = (from_row, from_col, to_row, to_col)
        anchor: Optional[Tuple[int, ...]] = None
        if any(c is not None for c in corners):
            if any(c is None for c in corners):
                raise TypeError("formulon: from_row, from_col, to_row and to_col move the frame together")
            anchor = (
                *cast(Tuple[int, int, int, int], corners),
                from_row_offset_emu, from_col_offset_emu, to_row_offset_emu, to_col_offset_emu,
            )  # fmt: skip
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = _pack_slicer(
                scope, None, caption, style, column_count, row_height_emu, anchor, sort_order, cross_filter
            )
            _check(LIB.fm_sheet_update_slicer(h, sheet, object_id, ptr))


for _cls in (Drawing, Picture, Chart, Shape, Slicer):
    _install_workbook_locks(_cls)
del _cls
