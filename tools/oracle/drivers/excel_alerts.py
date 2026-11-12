"""Excel's alert switch, turned off for a capture and always turned back on.

``display_alerts`` is application-wide on Mac Excel and outlives the process
that set it, so a driver that leaves it off silences the user's own dialogs
afterwards. Every driver goes through :func:`alerts_off`; it is the only code
allowed to assign the property (``tools/dev/check_display_alerts.py``).
"""

from __future__ import annotations

from contextlib import contextmanager
from typing import Any, Iterator


@contextmanager
def alerts_off(app: Any) -> Iterator[Any]:
    """Turns ``app``'s alerts off for the block and back on when it exits."""

    app.display_alerts = False
    try:
        yield app
    finally:
        try:
            app.display_alerts = True
        except Exception:  # the app may already have quit or been killed
            pass
