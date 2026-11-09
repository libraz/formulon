"""Helper: build chart inserts from a kind name and a source range."""

from __future__ import annotations

from typing import Any, Tuple

from formulon import (
    ChartBarDirection,
    ChartInput,
    ChartPlotType,
)

Kind = Tuple[ChartPlotType, "ChartBarDirection | None"]

COLUMN: Kind = (ChartPlotType.BAR, ChartBarDirection.COLUMN)
BAR: Kind = (ChartPlotType.BAR, ChartBarDirection.BAR)
LINE: Kind = (ChartPlotType.LINE, None)
AREA: Kind = (ChartPlotType.AREA, None)
PIE: Kind = (ChartPlotType.PIE, None)
SCATTER: Kind = (ChartPlotType.SCATTER, None)
DOUGHNUT: Kind = (ChartPlotType.DOUGHNUT, None)


def insert(kind: Kind, source: Any, *, anchor: Any = None, **options: Any) -> ChartInput:
    """The insert record for ``kind`` over ``source`` placed at ``anchor``."""
    plot, direction = kind
    return ChartInput(type=plot, range=source, bar_direction=direction, anchor=anchor, **options)


def add_chart(ws: Any, kind: Kind, source: Any, **options: Any) -> Any:
    """Insert a chart of ``kind`` on ``ws`` and return its handle."""
    return ws.add_chart(insert(kind, source, **options))


def set_type(chart: Any, plot_index: int, kind: Kind, **options: Any) -> None:
    """Change plot ``plot_index`` of ``chart`` to ``kind``."""
    plot, direction = kind
    chart.set_type(plot_index, plot, bar_direction=direction, **options)
