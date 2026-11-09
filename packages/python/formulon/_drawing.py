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

from typing import TYPE_CHECKING, Any, Dict, Optional, Sequence, Tuple, Type, Union, cast

from . import _structs as S
from ._c import LIB, FormulonError, _NoCopy, _sint, _uint
from ._status import StatusCode
from ._workbook import (
    _CHART_TITLE_NONE,
    _DISPLAY_PATCH_SPEC,
    _PATCH_SET,
    _PLOT_PATCH_SPEC,
    _PREVIEW_OVERRIDE_SPEC,
    UNCHANGED,
    BlanksAs,
    ChartAxisCrosses,
    ChartAxisKind,
    ChartAxisPatch,
    ChartAxisPosition,
    ChartAxisRef,
    ChartBarDirection,
    ChartData,
    ChartDataSource,
    ChartDisplayPatch,
    ChartFill,
    ChartFillPatch,
    ChartFilter,
    ChartGrouping,
    ChartLabelsPatch,
    ChartLegend,
    ChartLegendPosition,
    ChartLine,
    ChartLinePatch,
    ChartMarkerPatch,
    ChartPlotPatch,
    ChartPlotType,
    ChartPoint,
    ChartPointPatch,
    ChartSeriesDirection,
    ChartSeriesInput,
    ChartSeriesPatch,
    ChartSource,
    ChartSourceInput,
    ChartTickLabelPosition,
    ChartTitle,
    DrawingAnchorInput,
    DrawingObjectKind,
    ImageFormat,
    SheetRect,
    SlicerCrossFilter,
    SlicerSortOrder,
    Unchanged,
    _ArgScope,
    _chart_range_text,
    _chart_source_fields,
    _check,
    _encode_axis_patch,
    _encode_patch,
    _encode_point_patch,
    _encode_series_patch,
    _enum,
    _id_array,
    _install_workbook_locks,
    _pack_drawing_patch,
    _pack_shape_patch,
    _pack_slicer_patch,
    _read_chart,
    _read_chart_data,
    _resolve_chart_axis,
    _title_fields,
)

if TYPE_CHECKING:
    from ._sheet import Sheet
    from ._workbook import (
        ChartProperties,
        DrawingProperties,
        PictureProperties,
        ShapeProperties,
        SlicerProperties,
    )

_PREVIEW_OVERRIDE_NAMES = (
    "range",
    "direction",
    "header_row",
    "header_column",
    "blanks_as",
    "plot_visible_only",
    "na_as_blank",
)
_COMMON_UPDATE_NAMES = ("name", "description", "hidden", "anchor")
# The common keys whose ``None`` clears the field.
_CLEARABLE_COMMON_NAMES = ("description",)
_SHAPE_UPDATE_NAMES = (
    "preset_geometry",
    "text",
    "text_box",
    "fill",
    "line",
    "rotation_deg",
)
# The shape keys whose ``None`` clears the field to the style's paint.
_CLEARABLE_SHAPE_NAMES = ("fill", "line")
_SLICER_UPDATE_NAMES = ("caption", "style_name", "column_count", "row_height_pt", "sort_order", "cross_filter")
_SHAPE_KINDS = frozenset({DrawingObjectKind.SHAPE, DrawingObjectKind.CONNECTOR, DrawingObjectKind.GROUP})


def _given(arguments: Dict[str, Any], patch: type) -> Dict[str, Any]:
    """The keyword arguments among the keys of the patch ``TypedDict`` that are not :data:`UNCHANGED`."""
    return {key: arguments[key] for key in patch.__annotations__ if arguments[key] is not UNCHANGED}


def _present(
    arguments: Dict[str, Any], names: Sequence[str], absent: Sequence[str] = (), clearable: Sequence[str] = ()
) -> Dict[str, Any]:
    """The arguments among ``names`` that are not :data:`UNCHANGED`.

    A ``None`` among ``absent`` means "keep" and is dropped; a ``None`` among ``clearable`` is kept (it clears);
    any other ``None`` is a ``TypeError``.
    """
    out: Dict[str, Any] = {}
    for key in names:
        value = arguments[key]
        if value is UNCHANGED or (value is None and key in absent):
            continue
        if value is None and key not in clearable:
            raise TypeError(f"formulon: {key} cannot be None")
        out[key] = value
    return out


def _removed() -> FormulonError:
    """The error a handle raises when the engine no longer has its object."""
    return FormulonError(StatusCode.NOT_FOUND, "the drawing was removed", "")


def _raw_kind(sheet: Sheet, object_id: int) -> Optional[int]:
    """The raw kind of drawing object ``object_id`` of ``sheet``, or ``None`` when its drawing has none."""
    return sheet._workbook._lookup(
        LIB.fm_sheet_get_drawing_object,
        (_uint(sheet._index(), "sheet_index", role="sheet"), _uint(object_id, "object_id")),
        S.DRAWING_OBJECT,
        lambda d: d["kind"],
    )


def _has_slicer(sheet: Sheet, object_id: int) -> bool:
    """Whether drawing object ``object_id`` of ``sheet`` is the frame of a table slicer."""
    found = sheet._workbook._lookup(
        LIB.fm_sheet_get_slicer,
        (_uint(sheet._index(), "sheet_index", role="sheet"), _uint(object_id, "object_id")),
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
        raise FormulonError(StatusCode.NOT_FOUND, f"the sheet's drawing has no object {object_id}", "")
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
        handle = cls._new(sheet, object_id, kind)
        handles[object_id] = handle
    return handle


def _fill_patch_of(fill: Optional[ChartFill]) -> Optional[ChartFillPatch]:
    """The patch that makes a fill equal ``fill``: ``None`` clears it, an absent member is cleared, ``other`` keeps it."""
    if fill is None:
        return None
    if fill.other:
        return {}
    return {"color": fill.color, "transparency_pct": fill.transparency_pct, "none": fill.none}


def _line_patch_of(line: Optional[ChartLine]) -> Optional[ChartLinePatch]:
    """The patch that makes an outline equal ``line``: ``None`` clears it, an absent member is cleared, ``other`` keeps it."""
    if line is None:
        return None
    if line.other:
        return {}
    return {
        "color": line.color,
        "transparency_pct": line.transparency_pct,
        "width_pt": line.width_pt,
        "dash": line.dash,
        "none": line.none,
    }


class Drawing(_NoCopy):
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

    A subclass reads its record with ``get_properties()``; a plain
    :class:`Drawing` reads the common fields with ``get_drawing_properties()``.
    """

    __slots__ = ("_sheet", "_object_id", "_kind")

    def __init__(self, *args: object, **kwargs: object) -> None:
        raise TypeError("obtain a drawing object from Sheet.drawing(), Sheet.list_drawings() or an add_* method")

    @classmethod
    def _new(cls, sheet: Sheet, object_id: int, kind: DrawingObjectKind) -> "Drawing":
        handle = object.__new__(cls)
        handle._sheet = sheet
        handle._object_id = object_id
        handle._kind = kind
        return handle

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

    @property
    def name(self) -> str:
        """The object's name, read from the drawing on each access.

        Raises:
          NotFoundError: when the object or its sheet was removed.
        """
        return self.get_drawing_properties().name

    def _at(self) -> Tuple[int, int, int]:
        """``(workbook handle, sheet index, object id)`` for a raw call.

        The raw calls raise ``NotFoundError`` once the object is removed.
        """
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        return h, _uint(sheet, "sheet_index", role="sheet"), _uint(self._object_id, "object_id")

    def __repr__(self) -> str:
        return f"{type(self).__name__}(object_id={self._object_id})"

    def get_drawing_properties(self) -> DrawingProperties:
        """Return the object's kind, name, description, visibility and anchor.

        Raises:
          NotFoundError: when the object or its sheet was removed.
        """
        from ._workbook import _drawing_object_from

        _, sheet, object_id = self._at()
        found = self._sheet._workbook._lookup(
            LIB.fm_sheet_get_drawing_object, (sheet, object_id), S.DRAWING_OBJECT, _drawing_object_from
        )
        if found is None:
            raise _removed()
        return cast("DrawingProperties", found)

    def update(
        self,
        *,
        name: Union[str, Unchanged] = UNCHANGED,
        description: Union[str, None, Unchanged] = UNCHANGED,
        hidden: Union[bool, Unchanged] = UNCHANGED,
        anchor: Union[DrawingAnchorInput, Unchanged] = UNCHANGED,
    ) -> None:
        """Apply the arguments that are given to the object's name, description, visibility and anchor.

        An argument left at :data:`UNCHANGED` keeps its value. ``anchor`` merges into the current anchor
        field by field (see :class:`~formulon.DrawingAnchorInput`): an omitted field keeps its value and an
        omitted ``kind`` keeps the object's kind. ``description=None`` removes the description. A slicer's name
        cannot change. Duplicate names are allowed.

        Raises:
          TypeError: when an argument other than ``description`` is ``None`` or has the wrong type.
          FormulonError: with ``PRECONDITION_FAILED`` for a slicer ``name``.
        """
        given = _present(dict(locals()), _COMMON_UPDATE_NAMES, clearable=_CLEARABLE_COMMON_NAMES)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = _common_patch(scope, given)
            if ptr:
                _check(LIB.fm_sheet_update_drawing(h, sheet, object_id, ptr))


def _common_patch(scope: _ArgScope, given: Dict[str, Any]) -> int:
    """An ``fm_drawing_patch_t`` of the common keys in ``given``, or ``0`` when it has none."""
    if not any(key in given for key in _COMMON_UPDATE_NAMES):
        return 0
    clear = [key for key in _CLEARABLE_COMMON_NAMES if key in given and given[key] is None]
    return _pack_drawing_patch(
        scope, given.get("name"), given.get("description"), given.get("hidden"), given.get("anchor"), clear=clear
    )


class _ZOrdered(Drawing):
    """The drawing kinds with a stacking position: pictures, charts, shapes and slicers."""

    __slots__ = ()

    def set_z_order(self, z_order_index: int) -> None:
        """Make the object item ``z_order_index`` of :meth:`Sheet.list_drawings`.

        That list's order is the stacking order; ``0`` is the back.
        """
        h, sheet, object_id = self._at()
        _check(
            LIB.fm_sheet_set_drawing_z_order(
                h, sheet, object_id, _uint(z_order_index, "z_order_index", role="position")
            )
        )


class _Snapshotable(_ZOrdered):
    """The drawing kinds :meth:`Sheet.restore_drawing` can bring back: pictures, charts and shapes."""

    __slots__ = ()

    def snapshot(self) -> bytes:
        """Capture the object as opaque bytes for :meth:`Sheet.restore_drawing`.

        The capture holds the object's list position, element and every part it reaches.
        """
        h, sheet, object_id = self._at()
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


class Picture(_Snapshotable):
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
        from ._workbook import PictureProperties

        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            bytes_ptr = scope.out_ptr()
            len_ptr = scope.out_ptr()
            info = scope.struct(S.IMAGE_INFO)
            _check(LIB.fm_sheet_get_picture(h, sheet, object_id, present, bytes_ptr, len_ptr, info))
            if not LIB.read_u32(present):
                raise FormulonError(StatusCode.NOT_FOUND, "the picture's media part is missing", "")
            d = S.IMAGE_INFO.unpack(LIB, info)
            return PictureProperties(
                format=_enum(ImageFormat, d["format"]),
                width_px=d["width_px"],
                height_px=d["height_px"],
                bytes=LIB.read_bytes(LIB.read_u32(bytes_ptr), LIB.read_u32(len_ptr)),
            )


class Chart(_Snapshotable):
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
            return cast("ChartProperties", _read_chart(chart, self._sheet._workbook, sheet))
        finally:
            LIB.fm_chart_destroy(chart)

    def get_data(self) -> ChartData:
        """Resolve the chart's series against the current cell values by Excel's rules.

        Raises:
          NotFoundError: when the chart or its sheet was removed.
        """
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            out = scope.out_ptr()
            _check(LIB.fm_sheet_get_chart_data(h, sheet, object_id, present, out))
            if LIB.read_i32(present) == 0:
                raise _removed()
            data = LIB.read_u32(out)
        try:
            return _read_chart_data(data)
        finally:
            LIB.fm_chart_data_destroy(data)

    def get_preview(
        self,
        *,
        range: Union[SheetRect, str, Unchanged] = UNCHANGED,
        direction: Union[ChartSeriesDirection, Unchanged] = UNCHANGED,
        header_row: Union[bool, Unchanged] = UNCHANGED,
        header_column: Union[bool, Unchanged] = UNCHANGED,
        blanks_as: Union[BlanksAs, Unchanged] = UNCHANGED,
        plot_visible_only: Union[bool, Unchanged] = UNCHANGED,
        na_as_blank: Union[bool, Unchanged] = UNCHANGED,
    ) -> ChartData:
        """Resolve the data of a copy of the chart with the given overrides applied; the workbook is unchanged.

        An argument left at :data:`UNCHANGED` does not override: the chart's
        stored value is used. Series are rebuilt as :meth:`set_source` does:
        over the overridden ``range`` (a :class:`SheetRect` or an A1 range; an
        unset direction or header is auto), else over the chart's derived range.

        Raises:
          TypeError: when an argument is ``None``.
          NotFoundError: when the chart or its sheet was removed.
        """
        given = _present(dict(locals()), _PREVIEW_OVERRIDE_NAMES)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            fields = _encode_patch(_PREVIEW_OVERRIDE_SPEC, given, ("range",), "preview override", scope)
            if "range" in given:
                fields["source_op"] = _PATCH_SET
                fields["source"] = scope.cstr(_chart_range_text(self._sheet._workbook, given["range"], sheet, "range"))
            ptr = scope.struct(S.CHART_PREVIEW_OVERRIDES)
            S.CHART_PREVIEW_OVERRIDES.pack(LIB, ptr, fields)
            present = scope.out_ptr()
            out = scope.out_ptr()
            _check(LIB.fm_sheet_get_chart_preview(h, sheet, object_id, ptr, present, out))
            if LIB.read_i32(present) == 0:
                raise _removed()
            data = LIB.read_u32(out)
        try:
            return _read_chart_data(data)
        finally:
            if data:
                LIB.fm_chart_data_destroy(data)

    def set_type(
        self,
        plot: int,
        type: ChartPlotType,
        *,
        bar_direction: Optional[ChartBarDirection] = None,
        grouping: ChartGrouping = ChartGrouping.CLUSTERED,
    ) -> None:
        """Replace plot group ``plot`` with a plot of ``type``, keeping each series' references, idx, order and explicit colors.

        ``bar_direction`` (default ``COLUMN``) applies to ``BAR`` only; ``grouping``
        applies to bar, line and area plots.

        Raises:
          TypeError: for a ``bar_direction`` on a type other than ``BAR``.
          InvalidArgumentError: for a ``type`` that cannot be written.
        """
        if bar_direction is not None and _enum(ChartPlotType, type) != ChartPlotType.BAR:
            raise TypeError("formulon: bar_direction applies to a BAR chart only")
        h, sheet, object_id = self._at()
        _check(
            LIB.fm_sheet_set_chart_type(
                h,
                sheet,
                object_id,
                _uint(plot, "plot"),
                _sint(type, "type"),
                _sint(ChartBarDirection.COLUMN if bar_direction is None else bar_direction, "bar_direction"),
                _sint(grouping, "grouping"),
            )
        )

    def set_title(self, title: Union[str, None, ChartTitle]) -> None:
        """Set the chart title to ``title``; ``None`` removes it.

        A :class:`~formulon.ChartTitle` sets that kind, so
        ``ChartTitle(ChartTitleKind.AUTO)`` returns the title to
        Excel's automatic one.

        Raises:
          InvalidArgumentError: when ``title`` is an empty string; remove the
            title with ``None``.
        """
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_TITLE, 1)
            fields = {"kind": _sint(_CHART_TITLE_NONE, "kind")} if title is None else _title_fields(title, scope)
            S.CHART_TITLE.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_set_chart_title(h, sheet, object_id, ptr))

    def set_legend(self, legend: Optional[ChartLegend]) -> None:
        """Replace the legend with ``legend`` (the record :meth:`get_properties` returns), or remove it with ``None``.

        Removing the legend of a chart that has none does nothing.
        ``legend.hidden_entries`` holds series ``c:idx`` values.

        Raises:
          TypeError: when ``legend`` is neither a :class:`ChartLegend` nor ``None``.
          InvalidArgumentError: when ``legend.position`` is not a :class:`ChartLegendPosition`.
        """
        if legend is not None and not isinstance(legend, ChartLegend):
            raise TypeError(f"formulon: legend must be a ChartLegend or None, got {type(legend).__name__}")
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = 0
            if legend is not None:
                ptr = scope.struct_array(S.CHART_LEGEND, 1)
                S.CHART_LEGEND.pack(
                    LIB,
                    ptr,
                    {
                        "position": _sint(legend.position, "position"),
                        "overlay": 1 if legend.overlay else 0,
                        "hidden_entries": _id_array(legend.hidden_entries, "hidden_entries", scope),
                        "hidden_entries_count": _uint(len(legend.hidden_entries), "hidden_entries_count"),
                    },
                )
            _check(LIB.fm_sheet_set_chart_legend(h, sheet, object_id, ptr))

    def update_legend(
        self,
        *,
        position: Union[ChartLegendPosition, Unchanged] = UNCHANGED,
        overlay: Union[bool, Unchanged] = UNCHANGED,
        hidden_entries: Union[Sequence[int], Unchanged] = UNCHANGED,
    ) -> None:
        """Update the chart legend with the arguments given, creating it when the chart has none.

        An omitted argument keeps the legend's value, or takes Excel's default
        (:attr:`ChartLegendPosition.RIGHT`, no overlay, no hidden entries) when
        the chart has no legend. No argument has a cleared form;
        ``hidden_entries`` replaces the hidden legend entries (series ``c:idx``
        values) and ``[]`` shows them all.

        Raises:
          TypeError: when an argument is ``None`` or ``overlay`` is not a ``bool``.
        """
        given = {
            key: value
            for key, value in (("position", position), ("overlay", overlay), ("hidden_entries", hidden_entries))
            if value is not UNCHANGED
        }
        for key, value in given.items():
            if value is None:
                raise TypeError(f"formulon: {key} cannot be None")
        if overlay is not UNCHANGED and not isinstance(overlay, bool):
            raise TypeError(f"formulon: overlay must be a bool, got {type(overlay).__name__}")
        if not given:
            return
        h, sheet, object_id = self._at()
        fields: Dict[str, object] = {}
        if position is not UNCHANGED:
            fields["position_op"] = _PATCH_SET
            fields["position"] = _sint(position, "position")
        if overlay is not UNCHANGED:
            fields["overlay_op"] = _PATCH_SET
            fields["overlay"] = 1 if overlay else 0
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_LEGEND_PATCH, 1)
            if hidden_entries is not UNCHANGED:
                fields["hidden_entries_op"] = _PATCH_SET
                fields["hidden_entries"] = _id_array(hidden_entries, "hidden_entries", scope)
                fields["hidden_entries_count"] = _uint(len(hidden_entries), "hidden_entries_count")
            S.CHART_LEGEND_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_update_chart_legend(h, sheet, object_id, ptr))

    def update_axis(
        self,
        axis: ChartAxisRef,
        *,
        visible: Union[bool, None, Unchanged] = UNCHANGED,
        title: Union[str, ChartTitle, None, Unchanged] = UNCHANGED,
        min: Union[float, None, Unchanged] = UNCHANGED,
        max: Union[float, None, Unchanged] = UNCHANGED,
        major_unit: Union[float, None, Unchanged] = UNCHANGED,
        number_format: Union[str, None, Unchanged] = UNCHANGED,
        major_gridlines: Union[bool, None, Unchanged] = UNCHANGED,
        reversed: Union[bool, None, Unchanged] = UNCHANGED,
        log_base: Union[float, None, Unchanged] = UNCHANGED,
        crosses: Union[ChartAxisCrosses, None, Unchanged] = UNCHANGED,
        crosses_at: Union[float, None, Unchanged] = UNCHANGED,
        tick_label_position: Union[ChartTickLabelPosition, None, Unchanged] = UNCHANGED,
        minor_unit: Union[float, None, Unchanged] = UNCHANGED,
        minor_gridlines: Union[bool, Unchanged] = UNCHANGED,
        position: Union[ChartAxisPosition, Unchanged] = UNCHANGED,
        source_linked: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update an axis; an omitted argument keeps its value, ``None`` removes the element.

        ``axis`` is an axis id, a :class:`ChartAxisKind` (that kind's axis of the first plot group) or a
        :class:`ChartAxisKey` (the axis of that kind of plot group ``plot``).

        ``title`` takes the forms of :meth:`set_title`. ``None`` on ``title``
        removes the axis title, on ``log_base`` makes the axis linear, on
        ``minor_unit`` makes the unit automatic and on ``tick_label_position``
        restores the default. ``crosses_at`` is read when ``crosses`` is
        ``AT``. ``minor_gridlines``, ``position`` and ``source_linked`` reject ``None``.

        Raises:
          TypeError: when ``minor_gridlines``, ``position`` or ``source_linked`` is ``None``, or ``axis`` has the
            wrong type.
          NotFoundError: when the plot group has no axis of the kind ``axis`` names.
        """
        patch = _given(dict(locals()), ChartAxisPatch)
        for key in ("minor_gridlines", "position", "source_linked"):
            if key in patch and patch[key] is None:
                raise TypeError(f"formulon: {key} cannot be None")
        if isinstance(axis, int) and not isinstance(axis, (bool, ChartAxisKind)):
            axis_id = axis
        else:
            axis_id = _resolve_chart_axis(axis, self.get_properties())
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_AXIS_PATCH, 1)
            fields = _encode_axis_patch(patch, scope)
            S.CHART_AXIS_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_update_chart_axis(h, sheet, object_id, _uint(axis_id, "axis_id"), ptr))

    def update_display(
        self,
        *,
        blanks_as: Union[BlanksAs, None, Unchanged] = UNCHANGED,
        plot_visible_only: Union[bool, None, Unchanged] = UNCHANGED,
        na_as_blank: Union[bool, None, Unchanged] = UNCHANGED,
        style: Union[int, None, Unchanged] = UNCHANGED,
    ) -> None:
        """Update how blanks plot, whether hidden cells are skipped, whether ``#N/A`` plots as blank and the chart style.

        ``style`` is 1..48. An omitted argument keeps its value, ``None`` removes the element.
        """
        patch = _given(dict(locals()), ChartDisplayPatch)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            fields = _encode_patch(_DISPLAY_PATCH_SPEC, patch, (), "display patch", scope)
            ptr = scope.struct(S.CHART_DISPLAY_PATCH)
            S.CHART_DISPLAY_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_update_chart_display(h, sheet, object_id, ptr))

    def update_series(
        self,
        plot: int,
        series_id: int,
        *,
        name: Union[str, ChartDataSource, None, Unchanged] = UNCHANGED,
        categories: Union[str, ChartDataSource, None, Unchanged] = UNCHANGED,
        values: Union[str, ChartDataSource, None, Unchanged] = UNCHANGED,
        x: Union[str, ChartDataSource, None, Unchanged] = UNCHANGED,
        y: Union[str, ChartDataSource, None, Unchanged] = UNCHANGED,
        size: Union[str, ChartDataSource, None, Unchanged] = UNCHANGED,
        fill: Union[ChartFillPatch, None, Unchanged] = UNCHANGED,
        line: Union[ChartLinePatch, None, Unchanged] = UNCHANGED,
        marker: Union[ChartMarkerPatch, None, Unchanged] = UNCHANGED,
        smooth: Union[bool, None, Unchanged] = UNCHANGED,
        labels: Union[ChartLabelsPatch, None, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the series with ``c:idx`` ``series_id`` of plot group ``plot``; an omitted argument keeps its value.

        A reference is its new text or a :class:`ChartDataSource` with a reference; ``""`` removes ``name``,
        ``categories`` or ``x`` (an empty ``values``, ``y`` or ``size`` is rejected). ``None`` on
        ``marker`` or ``labels`` removes it. A changed reference refreshes its cache.
        """
        patch = _given(dict(locals()), ChartSeriesPatch)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_SERIES_PATCH, 1)
            fields = _encode_series_patch(patch, scope)
            S.CHART_SERIES_PATCH.pack(LIB, ptr, fields)
            _check(
                LIB.fm_sheet_update_chart_series(
                    h, sheet, object_id, _uint(plot, "plot"), _uint(series_id, "series_id"), ptr
                )
            )

    def set_point(self, plot: int, series_id: int, point_index: int, value: Optional[ChartPoint]) -> None:
        """Replace the formatting of data point ``point_index`` with ``value``, or remove it with ``None``.

        ``value`` is a :class:`ChartPoint` as :meth:`get_properties` reads it;
        its ``index`` is ignored. Removing a point that has no formatting does
        nothing.

        Raises:
          TypeError: when ``value`` is neither a :class:`ChartPoint` nor ``None``.
        """
        if value is None:
            self._patch_point(plot, series_id, point_index, None)
            return
        if not isinstance(value, ChartPoint):
            raise TypeError(f"formulon: value must be a ChartPoint or None, got {type(value).__name__}")
        # An absent member is written as ``None``, which clears whatever the point has there; a marker is
        # cleared only where the point has one, as plots without markers reject any marker key.
        if value.marker is not None:
            marker: Union[ChartMarkerPatch, None, Unchanged] = {
                "symbol": value.marker.symbol,
                "size_pt": value.marker.size_pt,
                "fill": _fill_patch_of(value.marker.fill),
                "line": _line_patch_of(value.marker.line),
            }
        else:
            marker = UNCHANGED
            plots = self.get_properties().plots
            for series in plots[plot].series if 0 <= plot < len(plots) else ():
                if series.series_id == series_id:
                    if any(p.index == point_index and p.marker is not None for p in series.points):
                        marker = None
                    break
        self.update_point(
            plot,
            series_id,
            point_index,
            fill=_fill_patch_of(value.fill),
            line=_line_patch_of(value.line),
            marker=marker,
        )

    def update_point(
        self,
        plot: int,
        series_id: int,
        point_index: int,
        *,
        fill: Union[ChartFillPatch, None, Unchanged] = UNCHANGED,
        line: Union[ChartLinePatch, None, Unchanged] = UNCHANGED,
        marker: Union[ChartMarkerPatch, None, Unchanged] = UNCHANGED,
    ) -> None:
        """Update point ``point_index`` of a series; an omitted argument keeps its value.

        A call with no keyword argument changes nothing; ``None`` on a keyword removes that member. Remove the
        point's formatting with :meth:`set_point` and ``None``.
        """
        self._patch_point(plot, series_id, point_index, _given(dict(locals()), ChartPointPatch))

    def _patch_point(self, plot: int, series_id: int, point_index: int, patch: Optional[Dict[str, Any]]) -> None:
        """Apply a point patch; ``None`` clears every member."""
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_POINT_PATCH, 1)
            fields = _encode_point_patch(patch, scope)
            S.CHART_POINT_PATCH.pack(LIB, ptr, fields)
            _check(
                LIB.fm_sheet_update_chart_point(
                    h,
                    sheet,
                    object_id,
                    _uint(plot, "plot"),
                    _uint(series_id, "series_id"),
                    _uint(point_index, "point_index", role="position"),
                    ptr,
                )
            )

    def update_plot(
        self,
        plot: int,
        *,
        gap_width_pct: Union[int, None, Unchanged] = UNCHANGED,
        overlap_pct: Union[int, None, Unchanged] = UNCHANGED,
        first_slice_angle_deg: Union[int, None, Unchanged] = UNCHANGED,
        vary_colors: Union[bool, None, Unchanged] = UNCHANGED,
        hole_size_pct: Union[int, None, Unchanged] = UNCHANGED,
    ) -> None:
        """Update plot group ``plot``; an omitted argument keeps its value, ``None`` removes the element.

        An argument the plot type does not have is rejected. ``hole_size_pct`` is
        10..90 (doughnut only).
        """
        patch = _given(dict(locals()), ChartPlotPatch)
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_PLOT_PATCH, 1)
            fields = _encode_patch(_PLOT_PATCH_SPEC, patch, (), "plot patch", scope)
            S.CHART_PLOT_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_update_chart_plot(h, sheet, object_id, _uint(plot, "plot"), ptr))

    def set_filter(self, filter: ChartFilter) -> None:
        """Replace the chart filter: the series (by ``c:idx``) and category positions ``filter`` lists are hidden.

        The :class:`ChartFilter` :meth:`get_properties` returns passes back as is.

        Raises:
          TypeError: when ``filter`` is not a :class:`ChartFilter`.
        """
        if not isinstance(filter, ChartFilter):
            raise TypeError(f"formulon: filter must be a ChartFilter, got {type(filter).__name__}")
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            _check(
                LIB.fm_sheet_set_chart_filter(
                    h,
                    sheet,
                    object_id,
                    _id_array(filter.hidden_series, "hidden_series", scope),
                    _uint(len(filter.hidden_series), "hidden_series_count"),
                    _id_array(filter.hidden_categories, "hidden_categories", scope),
                    _uint(len(filter.hidden_categories), "hidden_category_count"),
                )
            )

    def add_series(self, plot: int, series: ChartSeriesInput) -> int:
        """Append ``series`` to plot group ``plot`` and return its series id (``c:idx``).

        ``series.values`` is required (``y`` for a scatter plot); the other fields read as in
        :meth:`update_series`, and a field left ``None`` is not written.
        """
        if not isinstance(series, ChartSeriesInput):
            raise TypeError(f"Chart.add_series: series must be a ChartSeriesInput, got {type(series).__name__}")
        patch = {
            key: getattr(series, key) for key in ChartSeriesPatch.__annotations__ if getattr(series, key) is not None
        }
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = scope.struct_array(S.CHART_SERIES_PATCH, 1)
            fields = _encode_series_patch(patch, scope)
            S.CHART_SERIES_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_add_chart_series(h, sheet, object_id, _uint(plot, "plot"), ptr, out))
            return LIB.read_u32(out)

    def remove_series(self, plot: int, series_id: int) -> None:
        """Remove the series with ``c:idx`` ``series_id`` from plot group ``plot``."""
        h, sheet, object_id = self._at()
        _check(
            LIB.fm_sheet_remove_chart_series(h, sheet, object_id, _uint(plot, "plot"), _uint(series_id, "series_id"))
        )

    def set_source(self, source: Union[ChartSourceInput, ChartSource]) -> None:
        """Rebuild every series from ``source`` by the insert splitting rules.

        The chart kind and each series' formatting (by idx) are kept and the
        filter is cleared. ``source.range`` is a :class:`SheetRect` or an A1
        range (on the chart's sheet when unqualified); ``header_row`` /
        ``header_column`` ``None`` follow Excel's detection. The
        :class:`ChartSource` :meth:`get_properties` returns passes back as is.

        Raises:
          TypeError: when ``source`` is not one of the two record types.
        """
        h, sheet, object_id = self._at()
        text, direction, header_row, header_column = _chart_source_fields(self._sheet._workbook, sheet, source)
        with _ArgScope() as scope:
            _check(
                LIB.fm_sheet_set_chart_source(
                    h,
                    sheet,
                    object_id,
                    scope.cstr(text),
                    _sint(direction, "direction"),
                    _sint(header_row, "header_row"),
                    _sint(header_column, "header_column"),
                )
            )


class Shape(_Snapshotable):
    """A shape, connector or group of a sheet's drawing; also has the members of :class:`Drawing`.

    Add one with :meth:`Sheet.add_shape`;
    removing a group removes its children.
    """

    __slots__ = ()

    def get_properties(self) -> ShapeProperties:
        """Return the shape's geometry, text and paint.

        Raises:
          NotFoundError: when the shape or its sheet was removed.
        """
        from ._workbook import _shape_from

        _, sheet, object_id = self._at()
        found = self._sheet._workbook._lookup(LIB.fm_sheet_get_shape, (sheet, object_id), S.SHAPE, _shape_from)
        if found is None:
            raise _removed()
        return cast("ShapeProperties", found)

    def update(
        self,
        *,
        name: Union[str, Unchanged] = UNCHANGED,
        description: Union[str, None, Unchanged] = UNCHANGED,
        hidden: Union[bool, Unchanged] = UNCHANGED,
        anchor: Union[DrawingAnchorInput, Unchanged] = UNCHANGED,
        preset_geometry: Union[str, Unchanged] = UNCHANGED,
        text: Union[str, Unchanged] = UNCHANGED,
        text_box: Union[bool, Unchanged] = UNCHANGED,
        fill: Union[ChartFill, ChartFillPatch, None, Unchanged] = UNCHANGED,
        line: Union[ChartLine, ChartLinePatch, None, Unchanged] = UNCHANGED,
        rotation_deg: Union[float, Unchanged] = UNCHANGED,
    ) -> None:
        """Apply the arguments that are given to the shape; the rest of its XML is kept.

        An argument left at :data:`UNCHANGED` keeps its value. ``name``, ``description``, ``hidden`` and
        ``anchor`` are those of :meth:`Drawing.update` and apply first; the other arguments are the shape's
        own. ``fill`` and ``line`` are :class:`ChartFill` / :class:`ChartLine` records or patches; ``None``
        removes the shape's own paint so the one its style names applies, and a record that reads ``other``
        keeps the stored paint. Text applies to a shape only and keeps the first paragraph's
        properties and first run's formatting. Every argument is checked before anything changes, and the
        call changes everything or nothing.

        Raises:
          TypeError: when an argument other than ``description`` is ``None`` or has the wrong type.
        """
        arguments = dict(locals())
        clear = [key for key in _CLEARABLE_SHAPE_NAMES if arguments[key] is None]
        given = _present(
            arguments,
            _COMMON_UPDATE_NAMES + _SHAPE_UPDATE_NAMES,
            clearable=_CLEARABLE_COMMON_NAMES + _CLEARABLE_SHAPE_NAMES,
        )
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            common = _common_patch(scope, given)
            kind_ptr = 0
            if clear or any(key in given for key in _SHAPE_UPDATE_NAMES):
                kind_ptr = _pack_shape_patch(scope, *(given.get(key) for key in _SHAPE_UPDATE_NAMES), clear=clear)
            if common or kind_ptr:
                _check(LIB.fm_sheet_update_shape(h, sheet, object_id, common, kind_ptr))


class Slicer(_ZOrdered):
    """A table slicer, through its drawing frame; also has the members of :class:`Drawing`.

    Add one with :meth:`Sheet.add_slicer`.
    Removing its table or sheet removes the slicer.

    A slicer has no ``snapshot``: a snapshot would have to capture parts
    outside the drawing (the slicer part, its cache and extension data), and
    its name is the key that ties the frame to the slicer part, so
    :meth:`update` refuses to change it.
    """

    __slots__ = ()

    def get_properties(self) -> SlicerProperties:
        """Return the slicer's caption, cache, table column and layout.

        Raises:
          NotFoundError: when the slicer or its sheet was removed.
        """
        from ._workbook import _slicer_from

        _, sheet, object_id = self._at()
        found = self._sheet._workbook._lookup(LIB.fm_sheet_get_slicer, (sheet, object_id), S.SLICER, _slicer_from)
        if found is None:
            raise _removed()
        return cast("SlicerProperties", found)

    def update(
        self,
        *,
        name: Union[str, Unchanged] = UNCHANGED,
        description: Union[str, None, Unchanged] = UNCHANGED,
        hidden: Union[bool, Unchanged] = UNCHANGED,
        anchor: Union[DrawingAnchorInput, Unchanged] = UNCHANGED,
        caption: Union[str, Unchanged] = UNCHANGED,
        style_name: Union[str, None, Unchanged] = UNCHANGED,
        column_count: Union[int, Unchanged] = UNCHANGED,
        row_height_pt: Union[float, Unchanged] = UNCHANGED,
        sort_order: Union[SlicerSortOrder, Unchanged] = UNCHANGED,
        cross_filter: Union[SlicerCrossFilter, Unchanged] = UNCHANGED,
    ) -> None:
        """Apply the arguments that are given to the slicer; the rest of its XML is kept.

        An argument left at :data:`UNCHANGED` keeps its value; ``style_name=None`` removes the style and
        ``description=None`` the description.
        ``description``, ``hidden`` and ``anchor`` are those of :meth:`Drawing.update` and apply first; the
        frame stays two-cell. A ``name`` is refused: it keys the frame to the slicer part. Every argument is
        checked before anything changes.

        Raises:
          TypeError: when an argument other than ``style_name`` or ``description`` is ``None`` or has the wrong type.
          FormulonError: with ``PRECONDITION_FAILED`` when ``name`` is given.
        """
        arguments = dict(locals())
        clear = ["style_name"] if style_name is None else []
        given = _present(
            arguments, _COMMON_UPDATE_NAMES + _SLICER_UPDATE_NAMES, clearable=_CLEARABLE_COMMON_NAMES + ("style_name",)
        )
        h, sheet, object_id = self._at()
        with _ArgScope() as scope:
            common = _common_patch(scope, given)
            kind_ptr = 0
            if clear or any(key in given for key in _SLICER_UPDATE_NAMES):
                kind_ptr = _pack_slicer_patch(scope, *(given.get(key) for key in _SLICER_UPDATE_NAMES), clear=clear)
            if common or kind_ptr:
                _check(LIB.fm_sheet_update_slicer(h, sheet, object_id, common, kind_ptr))


for _cls in (Drawing, _ZOrdered, _Snapshotable, Picture, Chart, Shape, Slicer):
    _install_workbook_locks(_cls)
del _cls
