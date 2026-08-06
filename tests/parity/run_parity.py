#!/usr/bin/env python3
"""Cross-language parity gate: CLI vs npm vs Python wheel.

Runs every fixture in ``fixtures.json`` through every available channel
and asserts the channels agree at IEEE-754-bit granularity for numbers
and at byte-identical UTF-8 for text/errors/booleans.

Skip semantics:
  A missing channel (no native CLI binary, no node, no installed Python
  wheel) is reported but does not fail the gate. Fewer than two active
  channels produces CTest's conventional skip code (77), not a passing
  parity result.

Usage:
  python3 tests/parity/run_parity.py [--fixtures FILE] [--verbose]

Exit codes:
  0   -- at least two channels agreed with each other and every fixture
         expectation.
  1   -- channel evaluation failed, channels disagreed, or a channel did
         not satisfy a fixture expectation.
  2   -- usage error (bad fixture file, missing argument, ...).
  77  -- fewer than two channels were available; parity was skipped.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import struct
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

# ---------------------------------------------------------------------------
# Paths and configuration
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent.parent
DEFAULT_FIXTURES = SCRIPT_DIR / "fixtures.json"

CLI_PATH = REPO_ROOT / "build" / "bin" / "formulon_cli"
NPM_DIST_JS = REPO_ROOT / "packages" / "npm" / "dist" / "formulon.js"
PYTHON_PKG_DIR = REPO_ROOT / "packages" / "python"


# Mirror of formulon::ErrorCode (src/value.h). The C ABI exposes only the
# integer ordinal in fm_value_t.u.error_code; the CLI stringifies via
# ErrorCode::display_name. The runner reproduces that mapping so npm and
# Python channels can be normalized to the CLI's flat shape.
ERROR_CODE_TO_NAME: Dict[int, str] = {
    0: "#NULL!",
    1: "#DIV/0!",
    2: "#VALUE!",
    3: "#REF!",
    4: "#NAME?",
    5: "#NUM!",
    6: "#N/A",
    7: "#GETTING_DATA",
    8: "#SPILL!",
    9: "#CALC!",
    10: "#FIELD!",
    11: "#BLOCKED!",
    12: "#CONNECT!",
    13: "#EXTERNAL!",
    14: "#BUSY!",
    15: "#PYTHON!",
    16: "#UNKNOWN!",
}

# fm_value_kind_t mirror (src/c_api/formulon_c.h).
KIND_BLANK = 0
KIND_NUMBER = 1
KIND_BOOL = 2
KIND_TEXT = 3
KIND_ERROR = 4
SKIP_RETURN_CODE = 77


# ---------------------------------------------------------------------------
# Normalization
# ---------------------------------------------------------------------------


def _canonicalize_number(x: float) -> float:
    """Round ``x`` through Excel's General-format printer.

    The CLI's ``--json`` output goes through ``format_double``
    (``src/utils/double_format.cpp``), which renders values as
    ``%.15g`` (Excel's General format: 15 significant digits). Re-parsing
    that decimal back into a ``double`` drops up to ~3 trailing mantissa
    bits relative to the engine's internal value. The npm and Python
    channels can return the engine's full-precision ``double``, so a
    naive bit comparison would always disagree with the CLI on values
    like ``PI()`` or ``1/3``.

    To keep all three channels apples-to-apples we apply the same
    %.15g-then-reparse transform on every channel's numeric output
    before comparing. Engine-level disagreements on ``double`` values
    (one channel returning ``1.0`` while another returns ``2.0``)
    survive the rounding intact; only sub-displayable mantissa noise
    is collapsed. This mirrors the contract of ``format_double``: it is
    an Excel-General printer, not a full-precision serializer.
    """
    if x != x:  # NaN: %.15g would render "nan"; keep NaN visible.
        return float("nan")
    if x == 0.0:
        # Collapse signed zero so the CLI's "0" matches the +0.0 / -0.0
        # bit patterns from npm / Python uniformly.
        return 0.0
    # %.15g matches the CLI's `%.15g` exactly. The integer fast path
    # in format_double does not round, but those values already
    # round-trip through `%.15g` losslessly (any int < 1e16 fits in
    # 15 significant digits), so a single pipeline suffices.
    return float(f"{x:.15g}")


def number_bits(x: float) -> str:
    """Return the IEEE-754 big-endian hex representation of ``x``.

    Numbers are first canonicalized through ``_canonicalize_number`` to
    align with the CLI's lossy ``--json`` decimal output. See that
    helper's docstring for the rationale.
    """
    return struct.pack(">d", _canonicalize_number(float(x))).hex()


def make_number_record(x: float) -> Dict[str, Any]:
    return {"kind": "number", "bits": number_bits(x)}


def make_bool_record(x: bool) -> Dict[str, Any]:
    return {"kind": "bool", "value": bool(x)}


def make_text_record(s: str) -> Dict[str, Any]:
    return {"kind": "text", "value": s}


def make_error_record(name: str) -> Dict[str, Any]:
    return {"kind": "error", "value": name}


def make_blank_record() -> Dict[str, Any]:
    return {"kind": "blank"}


# ---------------------------------------------------------------------------
# Channels
# ---------------------------------------------------------------------------


@dataclass
class ChannelResult:
    """Outcome of evaluating one fixture on one channel."""

    record: Optional[Dict[str, Any]]  # normalized, None on call failure.
    raw: Any  # unprocessed payload, for diagnostics.
    error: Optional[str]  # human-readable failure reason.

    @property
    def ok(self) -> bool:
        return self.record is not None and self.error is None


class Channel:
    """Abstract base. Concrete channels evaluate a formula on demand."""

    name: str = "channel"

    def available(self) -> bool:
        return False

    def availability_reason(self) -> str:
        """Human-readable explanation of why available() is False."""
        return "not configured"

    def evaluate(self, formula: str) -> ChannelResult:
        raise NotImplementedError


# -- CLI channel ------------------------------------------------------------


class CliChannel(Channel):
    name = "cli"

    def __init__(self, cli_path: Path) -> None:
        self.cli_path = cli_path

    def available(self) -> bool:
        return self.cli_path.is_file() and os.access(self.cli_path, os.X_OK)

    def availability_reason(self) -> str:
        if not self.cli_path.exists():
            return f"missing: {self.cli_path}"
        if not self.cli_path.is_file():
            return f"not a regular file: {self.cli_path}"
        if not os.access(self.cli_path, os.X_OK):
            return f"not executable: {self.cli_path}"
        return "available"

    def evaluate(self, formula: str) -> ChannelResult:
        try:
            proc = subprocess.run(
                [str(self.cli_path), "eval", "--json", formula],
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
            )
        except (OSError, subprocess.SubprocessError) as exc:
            return ChannelResult(None, None, f"subprocess failure: {exc}")
        if proc.returncode != 0:
            return ChannelResult(
                None,
                {"stdout": proc.stdout, "stderr": proc.stderr, "rc": proc.returncode},
                f"cli exit {proc.returncode}: {proc.stderr.strip()}",
            )
        stdout = proc.stdout.strip()
        try:
            payload = json.loads(stdout)
        except json.JSONDecodeError as exc:
            return ChannelResult(None, stdout, f"non-JSON cli output: {exc}")
        return ChannelResult(_normalize_cli(payload), payload, None)


def _normalize_cli(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Translate the CLI's ``{"kind": str, "value": ...}`` JSON into a record."""
    kind = payload.get("kind")
    value = payload.get("value")
    if kind == "blank":
        return make_blank_record()
    if kind == "number":
        # The CLI stringifies via format_double, so the JSON carries a
        # JS-Number-compatible decimal. Re-parse to a Python float and
        # encode as IEEE-754 bits to align with the npm/Python channels.
        return make_number_record(float(value))
    if kind == "bool":
        return make_bool_record(bool(value))
    if kind == "text":
        return make_text_record(str(value))
    if kind == "error":
        return make_error_record(str(value))
    # Reserved kinds (array/ref/lambda) carry "value":null; surface as-is.
    return {"kind": str(kind), "value": value}


# -- npm channel ------------------------------------------------------------


# This script is written to a temp file and run via `node <script>`.
# It imports the staged module, calls evalFormula, and prints exactly one
# JSON line containing a *raw* payload. Numbers are emitted as the
# big-endian IEEE-754 hex string (Buffer.writeDoubleBE) so the runner
# never has to round-trip through JS's decimal printer; the Python side
# decodes the bits back to a double and then applies the same
# %.15g canonicalization the CLI's `format_double` uses, keeping the
# three channels apples-to-apples.
NPM_RUNNER_SCRIPT = r"""
import factory from {dist_url};
const ERROR_CODE_TO_NAME = {error_table};
function payloadOf(value) {{
  switch (value.kind) {{
    case 0: return {{ kind: 'blank' }};
    case 1: {{
      const buf = Buffer.alloc(8);
      buf.writeDoubleBE(value.number, 0);
      return {{ kind: 'number', bits: buf.toString('hex') }};
    }}
    case 2: return {{ kind: 'bool', value: !!value.boolean }};
    case 3: return {{ kind: 'text', value: String(value.text) }};
    case 4: {{
      const name = ERROR_CODE_TO_NAME[value.errorCode] || '#UNKNOWN!';
      return {{ kind: 'error', value: name }};
    }}
    default: return {{ kind: String(value.kind), value: null }};
  }}
}}
const formula = process.argv[process.argv.length - 1];
const M = await factory();
const r = M.evalFormula(formula);
if (!r.status.ok) {{
  process.stderr.write('npm-channel: status not ok: ' + JSON.stringify(r.status) + '\n');
  process.exit(2);
}}
process.stdout.write(JSON.stringify(payloadOf(r.value)) + '\n');
"""


class NpmChannel(Channel):
    name = "npm"

    def __init__(self, dist_js: Path) -> None:
        self.dist_js = dist_js
        self._node = shutil.which("node")
        # Reuse a single staged runner script across calls.
        self._runner: Optional[Path] = None

    def __del__(self) -> None:  # pragma: no cover -- GC ordering best-effort
        try:
            if self._runner is not None and self._runner.exists():
                self._runner.unlink()
        except OSError:
            pass

    def available(self) -> bool:
        return self._node is not None and self.dist_js.is_file()

    def availability_reason(self) -> str:
        if self._node is None:
            return "node not found in PATH"
        if not self.dist_js.exists():
            return f"missing: {self.dist_js}"
        return "available"

    def _ensure_runner(self) -> Path:
        if self._runner is not None and self._runner.exists():
            return self._runner
        # The `from {dist_url}` substitution must produce a JS string
        # literal containing a file:// URL so Node's ESM loader resolves
        # the staged module unambiguously regardless of CWD.
        dist_url = json.dumps(self.dist_js.resolve().as_uri())
        error_table = json.dumps({str(k): v for k, v in ERROR_CODE_TO_NAME.items()})
        body = NPM_RUNNER_SCRIPT.format(dist_url=dist_url, error_table=error_table)
        # Write to a NamedTemporaryFile and keep it for the lifetime of
        # the channel. Node v25 cannot `--input-type=module --eval` on
        # macOS (the worker rejects it), so a real .mjs file is the only
        # portable invocation form.
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".mjs", delete=False, encoding="utf-8")
        try:
            tmp.write(body)
        finally:
            tmp.close()
        self._runner = Path(tmp.name)
        return self._runner

    def evaluate(self, formula: str) -> ChannelResult:
        if self._node is None:
            return ChannelResult(None, None, "node not in PATH")
        runner = self._ensure_runner()
        try:
            proc = subprocess.run(
                [self._node, str(runner), formula],
                capture_output=True,
                text=True,
                timeout=60,
                check=False,
            )
        except (OSError, subprocess.SubprocessError) as exc:
            return ChannelResult(None, None, f"subprocess failure: {exc}")
        if proc.returncode != 0:
            return ChannelResult(
                None,
                {"stdout": proc.stdout, "stderr": proc.stderr, "rc": proc.returncode},
                f"npm exit {proc.returncode}: {proc.stderr.strip()}",
            )
        stdout = proc.stdout.strip()
        try:
            payload = json.loads(stdout)
        except json.JSONDecodeError as exc:
            return ChannelResult(None, stdout, f"non-JSON npm output: {exc}")
        return ChannelResult(_normalize_npm(payload), payload, None)


def _normalize_npm(payload: Dict[str, Any]) -> Dict[str, Any]:
    """Translate the npm runner's raw payload into a canonicalized record."""
    kind = payload.get("kind")
    if kind == "number":
        # The runner emits big-endian IEEE-754 hex; decode then route
        # through the same canonicalization used for the CLI / Python
        # so all three channels see identical bits at %.15g precision.
        bits = payload["bits"]
        x = struct.unpack(">d", bytes.fromhex(bits))[0]
        return make_number_record(x)
    if kind == "bool":
        return make_bool_record(bool(payload["value"]))
    if kind == "text":
        return make_text_record(str(payload["value"]))
    if kind == "error":
        return make_error_record(str(payload["value"]))
    if kind == "blank":
        return make_blank_record()
    return {"kind": str(kind), "value": payload.get("value")}


# -- Python channel ---------------------------------------------------------


class PythonChannel(Channel):
    name = "python"

    def __init__(self, pkg_dir: Path) -> None:
        self.pkg_dir = pkg_dir
        self._mod = None
        self._reason = "not yet probed"
        self._probe()

    def _probe(self) -> None:
        # Strategy 1: module already importable (installed wheel).
        # We catch `Exception` rather than `ImportError` because the
        # `formulon` package's `__init__` eagerly resolves
        # `libformulon.so` via ctypes; a missing/unbuilt shared library
        # raises `OSError`, not `ImportError`. Treat any failure to
        # bring the module up as "channel unavailable" so the harness
        # stays skip-aware (per the README contract) instead of
        # aborting the whole parity run.
        try:
            import formulon  # type: ignore[import-not-found]

            self._mod = formulon
            self._reason = "available (installed)"
            return
        except Exception as exc:  # noqa: BLE001 -- channel-boundary skip
            self._reason = f"installed import failed: {exc}"
        # Strategy 2: source-tree fallback. Add packages/python to sys.path.
        if self.pkg_dir.is_dir():
            sys.path.insert(0, str(self.pkg_dir))
            try:
                import formulon  # type: ignore[import-not-found]

                self._mod = formulon
                self._reason = f"available (source tree: {self.pkg_dir})"
                return
            except Exception as exc:  # noqa: BLE001 -- channel-boundary skip
                self._reason = f"source-tree import failed: {exc}"
                return
        # Only overwrite reason here if the installed-strategy probe
        # didn't already record an informative message.
        if self._reason == "not yet probed":
            self._reason = "not installed and source tree missing"

    def available(self) -> bool:
        return self._mod is not None

    def availability_reason(self) -> str:
        return self._reason

    def evaluate(self, formula: str) -> ChannelResult:
        if self._mod is None:
            return ChannelResult(None, None, self._reason)
        # Catching everything keeps a host-side FormulonError (e.g. parser
        # crash) from aborting the whole run. We surface it as a channel
        # failure on this fixture only.
        try:
            value = self._mod.eval_formula(formula)
        except Exception as exc:  # noqa: BLE001 -- channel boundary
            return ChannelResult(None, None, f"python eval_formula raised: {exc}")
        return ChannelResult(_normalize_python(value), value, None)


def _normalize_python(value: Any) -> Dict[str, Any]:
    """Translate ``formulon.Value`` into a record."""
    kind_int = int(value.kind)
    if kind_int == KIND_BLANK:
        return make_blank_record()
    if kind_int == KIND_NUMBER:
        return make_number_record(float(value.number))
    if kind_int == KIND_BOOL:
        return make_bool_record(bool(value.boolean))
    if kind_int == KIND_TEXT:
        return make_text_record(str(value.text))
    if kind_int == KIND_ERROR:
        code = int(value.error_code) if value.error_code is not None else -1
        return make_error_record(ERROR_CODE_TO_NAME.get(code, "#UNKNOWN!"))
    return {"kind": str(kind_int), "value": None}


# ---------------------------------------------------------------------------
# Runner core
# ---------------------------------------------------------------------------


@dataclass
class Divergence:
    fixture_id: str
    formula: str
    expect: Dict[str, Any]
    triplet: List[Tuple[str, ChannelResult]]
    baseline: str  # name of the channel chosen as canonical reference


def load_fixtures(path: Path) -> List[Dict[str, Any]]:
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, list):
        raise ValueError(f"fixtures file is not a list: {path}")
    for entry in data:
        if "id" not in entry or "formula" not in entry:
            raise ValueError(f"fixture missing id/formula: {entry!r}")
    return data


def records_match(a: Dict[str, Any], b: Dict[str, Any]) -> bool:
    return a == b


def record_matches_expect(record: Dict[str, Any], expect: Dict[str, Any]) -> bool:
    """Return whether a normalized channel record satisfies a fixture hint."""
    if "kind" in expect and record.get("kind") != expect["kind"]:
        return False
    if "value" not in expect:
        return True
    if expect.get("kind") == "number":
        return record.get("bits") == number_bits(float(expect["value"]))
    return record.get("value") == expect["value"]


def run(fixtures: List[Dict[str, Any]], channels: List[Channel], verbose: bool) -> int:
    active = [c for c in channels if c.available()]

    print(f"parity: {len(fixtures)} fixtures, {len(channels)} channels declared")
    for ch in channels:
        status = "active" if ch.available() else "skipped"
        print(f"  channel {ch.name:<7} {status:<7} -- {ch.availability_reason()}")

    if len(active) < 2:
        print(
            f"parity: only {len(active)} channel(s) active; at least 2 required for a meaningful parity check. SKIPPED."
        )
        return SKIP_RETURN_CODE

    divergences: List[Divergence] = []
    channel_failures: List[Tuple[str, str, str]] = []  # (fixture_id, channel, error)
    expectation_failures: List[Tuple[str, str, Dict[str, Any], Dict[str, Any]]] = []

    for entry in fixtures:
        fid = entry["id"]
        formula = entry["formula"]
        expect = entry.get("expect", {})

        triplet: List[Tuple[str, ChannelResult]] = []
        for ch in active:
            result = ch.evaluate(formula)
            triplet.append((ch.name, result))
            if not result.ok and result.error:
                channel_failures.append((fid, ch.name, result.error))
            elif result.ok and not record_matches_expect(result.record, expect):
                expectation_failures.append((fid, ch.name, expect, result.record))

        # Choose the first channel that produced a record as the baseline.
        baseline_name: Optional[str] = None
        baseline_record: Optional[Dict[str, Any]] = None
        for name, result in triplet:
            if result.ok:
                baseline_name = name
                baseline_record = result.record
                break

        if baseline_record is None:
            # Every active channel failed to produce a record. That's not a
            # parity divergence; channel_failures already captured the
            # individual errors.
            if verbose:
                print(f"  fixture {fid:<24} all channels failed to evaluate")
            continue

        mismatch = False
        for name, result in triplet:
            if not result.ok:
                continue  # captured in channel_failures.
            if not records_match(baseline_record, result.record):
                mismatch = True
                break

        if mismatch:
            divergences.append(
                Divergence(
                    fixture_id=fid,
                    formula=formula,
                    expect=expect,
                    triplet=triplet,
                    baseline=baseline_name,
                )
            )
            if verbose:
                print(f"  fixture {fid:<24} DIVERGENCE")
        else:
            if verbose:
                print(f"  fixture {fid:<24} ok")

    print()
    print(
        f"parity summary: fixtures={len(fixtures)} channels={len(active)} "
        f"divergences={len(divergences)} channel_failures={len(channel_failures)} "
        f"expectation_failures={len(expectation_failures)}"
    )

    for fid, channel, err in channel_failures:
        print(f"  channel-failure {channel:<7} on {fid}: {err}")

    for fid, channel, expect, actual in expectation_failures:
        print(f"  expectation-failure {channel:<7} on {fid}: expected {expect}, got {actual}")

    for div in divergences:
        print()
        print(f"DIVERGENCE on fixture '{div.fixture_id}': {div.formula}")
        print(f"  expect:   {div.expect}")
        print(f"  baseline: {div.baseline}")
        for name, result in div.triplet:
            if result.ok:
                print(f"  {name:<7} -> {result.record}")
            else:
                print(f"  {name:<7} !! {result.error}")

    if divergences or channel_failures or expectation_failures:
        return 1
    return 0


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Cross-channel parity gate for Formulon (CLI/npm/python).")
    parser.add_argument(
        "--fixtures",
        type=Path,
        default=DEFAULT_FIXTURES,
        help=f"path to fixtures JSON (default: {DEFAULT_FIXTURES})",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="print per-fixture status during the run",
    )
    args = parser.parse_args(argv)

    if not args.fixtures.is_file():
        print(f"parity: fixtures file not found: {args.fixtures}", file=sys.stderr)
        return 2

    try:
        fixtures = load_fixtures(args.fixtures)
    except (ValueError, json.JSONDecodeError) as exc:
        print(f"parity: cannot load fixtures: {exc}", file=sys.stderr)
        return 2

    channels: List[Channel] = [
        CliChannel(CLI_PATH),
        NpmChannel(NPM_DIST_JS),
        PythonChannel(PYTHON_PKG_DIR),
    ]
    return run(fixtures, channels, verbose=args.verbose)


if __name__ == "__main__":
    sys.exit(main())
