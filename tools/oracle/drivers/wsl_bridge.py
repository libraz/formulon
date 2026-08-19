#!/usr/bin/env python3
"""WSL2 -> Windows Excel bridge driver.

Implements :class:`OracleDriver` by spawning ``python.exe -m
tools.oracle.drivers.windows_excel --serve`` once and ferrying
newline-delimited JSON requests to it over stdin/stdout. Excel is
opened a single time on the Windows side and reused across every
``run_suite`` call; without this, generating 90+ suites paid one Excel
cold-start each (~5-15s) and dominated wall-clock time.

Used automatically when :func:`tools.oracle.drivers.select_driver` sees:

  - ``target.driver == 'windows_excel'`` AND host is WSL2.

Native Windows hosts use :class:`WindowsExcelOracle` directly; native
macOS hosts have no use for this module.

## Why subprocess instead of in-process COM?

WSL2's Linux Python cannot load the Windows pywin32 COM bridge -- the
two ABIs are incompatible. The only way to drive Excel from a WSL2
shell is to invoke Windows-side ``python.exe`` (which can be mounted as
``/mnt/c/...``) and shuttle data over a serializable channel.

## Configuration

  - ``target['win_python']``: absolute Windows-side path to
    ``python.exe``. Required. The ``make oracle-setup`` preflight is
    expected to discover a usable interpreter and write it back to
    ``targets.yaml``; until that lands the field must be filled in
    manually.
"""

from __future__ import annotations

import json
import platform
import subprocess
import sys
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional

from .base import CaseResult, EnvironmentInfo, OracleDriver


def _ensure_wsl2() -> None:
    """Refuses to start unless the host is a WSL2 kernel.

    Plain Linux without the Microsoft kernel marker would have no
    ``python.exe`` to invoke and no ``wslpath`` to translate paths.
    Surfacing the failure here keeps error messages local to the
    bridge instead of bubbling up as a confused subprocess error.
    """

    if platform.system() != "Linux":
        raise RuntimeError(f"WSL bridge requires Linux/WSL2 host, got {platform.system()}")
    try:
        proc = Path("/proc/version").read_text(encoding="utf-8").lower()
    except OSError as exc:
        raise RuntimeError("cannot read /proc/version (WSL2 detection failed)") from exc
    if "microsoft" not in proc:
        raise RuntimeError("WSL bridge requires WSL2 (no 'microsoft' in /proc/version)")


class WSLBridgeOracle(OracleDriver):
    """Drives a Windows-side :class:`WindowsExcelOracle` through a persistent subprocess.

    Lifetime is bounded by the surrounding ``with`` block: ``__enter__``
    spawns ``python.exe ... --serve``, waits for the ``ready`` line
    (which also carries the cached environment), and ``__exit__``
    sends a ``shutdown`` command and reaps the process. All
    :meth:`probe_environment`, :meth:`run_suite`, and
    :meth:`run_workbook_case` calls flow through the same long-lived
    stdio pipe, so Excel cold-start is paid once.
    """

    def __init__(self, *, win_python: str, visible: bool = False) -> None:
        _ensure_wsl2()
        if not win_python:
            raise RuntimeError(
                "WSL bridge requires `win_python` in targets.yaml; run `make oracle-setup` to discover it."
            )
        self._win_python = win_python
        self._visible = visible
        self._proc: Optional[subprocess.Popen] = None
        self._stderr_buf: List[str] = []
        self._stderr_thread: Optional[threading.Thread] = None
        self._cached_env: Optional[Dict[str, Any]] = None

    def __enter__(self) -> "WSLBridgeOracle":
        # `-X utf8=1` enables Python's UTF-8 mode on the Windows side
        # without relying on env-var inheritance. Why we need this:
        #
        #   - The Windows console default code page is locale-bound:
        #     CP932 (ja-JP), CP1252 (de-DE/fr-FR), GBK (zh-CN), and so
        #     on. Without UTF-8 mode the subprocess writes
        #     traceback/error text in that code page; decoding it as
        #     utf-8 on the WSL side produces mojibake (e.g. the COM
        #     error "例外が発生しました。" comes back as "??O...").
        #
        #   - WSL2 does NOT forward arbitrary env vars to Windows .exe
        #     processes; only names listed in $WSLENV cross. So setting
        #     `PYTHONUTF8=1` via `env=` would silently be lost --
        #     measured: subprocess sees an empty PYTHONUTF8 and
        #     sys.stdout.encoding stays at cp932.
        #
        #   - The `-X utf8=1` flag is a command-line equivalent of
        #     `PYTHONUTF8=1` (Python 3.7+ "UTF-8 mode"). Because it's
        #     an argv element it bypasses the WSLENV gate entirely and
        #     reliably forces sys.stdout/stderr to utf-8 across every
        #     Windows locale.
        cmd = [
            self._win_python,
            "-X",
            "utf8=1",
            "-m",
            "tools.oracle.drivers.windows_excel",
            "--serve",
        ]
        if self._visible:
            cmd.append("--visible")
        self._proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            encoding="utf-8",
            errors="replace",
            bufsize=1,  # line-buffered text mode
        )
        # Drain stderr in a background thread so a chatty subprocess
        # never blocks on a full pipe. The buffer is consulted only when
        # we surface an error.
        self._stderr_thread = threading.Thread(target=self._drain_stderr, daemon=True)
        self._stderr_thread.start()

        # Wait for the ready line. If the subprocess fails to start
        # (Excel activation prompt, missing xlwings, COM hang, ...),
        # _read_line raises with the captured stderr.
        ready = self._read_line()
        if ready.get("type") != "ready":
            raise RuntimeError(f"windows_excel did not announce ready: {ready!r}")
        self._cached_env = ready.get("environment") or {}
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        proc = self._proc
        if proc is None:
            return
        try:
            if proc.stdin and not proc.stdin.closed:
                try:
                    proc.stdin.write(json.dumps({"version": 1, "command": "shutdown"}) + "\n")
                    proc.stdin.flush()
                except (BrokenPipeError, OSError):
                    pass
                try:
                    proc.stdin.close()
                except OSError:
                    pass
            try:
                proc.wait(timeout=30)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait(timeout=5)
        finally:
            self._proc = None

    def _drain_stderr(self) -> None:
        """Buffers the Windows side's stderr and echoes it as it arrives.

        The buffer alone only ever reaches an operator through
        `_stderr_dump`, which runs on subprocess death or a `type: error`
        reply -- so a warning from a run that *succeeded* was invisible.
        That is how every print suite came to paginate against the host's
        default network printer without a word: the driver's
        "Microsoft Print to PDF is unavailable" notice went into this
        buffer and stayed there. A capture's warnings belong on the
        operator's terminal, so the line is echoed too.
        """

        assert self._proc is not None and self._proc.stderr is not None
        for line in self._proc.stderr:
            self._stderr_buf.append(line)
            print(line.rstrip("\n"), file=sys.stderr, flush=True)

    def _stderr_dump(self) -> str:
        return "".join(self._stderr_buf).strip()

    def _read_line(self) -> Dict[str, Any]:
        assert self._proc is not None and self._proc.stdout is not None
        line = self._proc.stdout.readline()
        if not line:
            rc = self._proc.poll()
            raise RuntimeError(
                f"windows_excel subprocess closed unexpectedly (rc={rc}):\nstderr: {self._stderr_dump()}"
            )
        return json.loads(line)

    def _invoke(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """Runs one wire-protocol round-trip against the long-lived server.

        Sends one JSON line on stdin, reads one JSON line on stdout.
        Subprocess death or a ``type: error`` response is surfaced as
        ``RuntimeError`` carrying any captured stderr so the operator
        can diagnose Office activation / COM issues.
        """

        assert self._proc is not None, "use as context manager"
        if self._proc.stdin is None:
            raise RuntimeError("bridge stdin is not open")
        try:
            self._proc.stdin.write(json.dumps(payload, ensure_ascii=False) + "\n")
            self._proc.stdin.flush()
        except (BrokenPipeError, OSError) as exc:
            raise RuntimeError(f"windows_excel bridge stdin closed: {exc}\nstderr: {self._stderr_dump()}") from exc
        resp = self._read_line()
        if resp.get("type") == "error":
            raise RuntimeError(
                f"windows_excel server error: {resp.get('error', 'unknown')}\nstderr: {self._stderr_dump()}"
            )
        return resp

    def probe_environment(self) -> EnvironmentInfo:
        # The server already sent the environment in its ready line. We
        # still issue a round-trip so callers that probe explicitly get
        # the same shape as the legacy path, but the data itself is the
        # cached snapshot -- Excel's locale doesn't shift mid-run.
        env = (self._cached_env or {}) if self._cached_env else {}
        if not env:
            out = self._invoke({"version": 1, "command": "probe_environment"})
            env = out.get("environment") or {}
        return EnvironmentInfo(
            excel_version=env.get("excel_version", ""),
            excel_locale=env.get("excel_locale", ""),
            date1904=bool(env.get("date1904", False)),
            iterative=bool(env.get("iterative", False)),
        )

    def run_suite(
        self,
        suite_name: str,
        cases: List[Dict[str, Any]],
        *,
        date1904: bool = False,
        iterative: bool = False,
    ) -> List[CaseResult]:
        out = self._invoke(
            {
                "version": 1,
                "command": "run_suite",
                "suite_name": suite_name,
                "date1904": date1904,
                "iterative": iterative,
                "cases": cases,
            }
        )
        return [
            CaseResult(
                id=r["id"],
                kind=r["kind"],
                value=r.get("value"),
                error_code=r.get("error_code"),
                array_shape=r.get("array_shape"),
            )
            for r in out["results"]
        ]

    def run_workbook_case(self, case: Dict[str, Any]) -> Dict[str, Any]:
        out = self._invoke(
            {
                "version": 1,
                "command": "run_workbook_case",
                "case": case,
            }
        )
        expect = out.get("expect")
        if not isinstance(expect, dict):
            raise RuntimeError(f"windows_excel bridge returned malformed run_workbook_case response: {out!r}")
        return expect
