#!/usr/bin/env python3
# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
"""Stage libformulon into the Python package's _lib/ directory.

Run via ``make python-package``. Stdlib only -- no third-party deps so
this works in any clean Python 3.9+ environment.

Inputs (defaults; override via command-line flags):
  --build-dir build-py
  --config    Release

Steps:
  1. Configure CMake with FM_BUILD_C_API_SHARED=ON.
  2. Build the ``formulon`` shared-library target.
  3. Copy the resulting libformulon.{dylib,so,dll} into
     packages/python/formulon/_lib/.
"""

from __future__ import annotations

import argparse
import platform
import shutil
import subprocess
import sys
from pathlib import Path

# packages/python/scripts/stage.py -> repo root is three levels up.
REPO_ROOT = Path(__file__).resolve().parent.parent.parent.parent
PKG_LIB_DIR = REPO_ROOT / "packages" / "python" / "formulon" / "_lib"


def _parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--build-dir",
        default="build-py",
        help="CMake build directory (relative to repo root). Default: build-py",
    )
    p.add_argument(
        "--config",
        default="Release",
        choices=["Release", "Debug", "RelWithDebInfo", "MinSizeRel"],
        help="CMake build configuration. Default: Release",
    )
    return p.parse_args(argv)


def _run(cmd: list[str], cwd: Path | None = None) -> None:
    """Run ``cmd`` and abort with a clear message on failure."""
    print(f"+ {' '.join(cmd)}", flush=True)
    result = subprocess.run(cmd, cwd=cwd, check=False)
    if result.returncode != 0:
        print(
            f"stage.py: command failed (exit {result.returncode}): {' '.join(cmd)}",
            file=sys.stderr,
        )
        sys.exit(result.returncode)


def _resolve_built_lib(build_dir: Path, config: str) -> Path:
    """Locate the freshly built shared library inside ``build_dir``.

    Different generators put the artefact in different locations:
      * Single-config (Make / Ninja, default on macOS / Linux):
        ``build_dir/libformulon.{dylib,so}``.
      * Multi-config (Visual Studio, Xcode):
        ``build_dir/<Config>/formulon.dll`` (Windows) or
        ``build_dir/<Config>/libformulon.{dylib,so}``.
    """
    system = platform.system()
    if system == "Darwin":
        names = ["libformulon.dylib"]
    elif system == "Windows":
        names = ["formulon.dll", "libformulon.dll"]
    else:
        names = ["libformulon.so"]

    candidates: list[Path] = []
    for name in names:
        candidates.append(build_dir / name)
        candidates.append(build_dir / config / name)

    for c in candidates:
        if c.is_file():
            return c

    raise FileNotFoundError(
        f"stage.py: could not locate built shared library. Looked in:\n"
        + "\n".join(f"  {c}" for c in candidates)
    )


def _windows_import_libs(build_dir: Path, config: str) -> list[Path]:
    """On Windows, the .lib import library should ship next to the .dll."""
    libs: list[Path] = []
    for sub in (build_dir, build_dir / config):
        for name in ("formulon.lib", "libformulon.lib"):
            p = sub / name
            if p.is_file():
                libs.append(p)
    return libs


def main(argv: list[str]) -> int:
    args = _parse_args(argv)
    build_dir = (REPO_ROOT / args.build_dir).resolve()

    # 1. Configure.
    _run(
        [
            "cmake",
            "-B",
            str(build_dir),
            "-S",
            str(REPO_ROOT),
            "-DFM_BUILD_C_API_SHARED=ON",
            "-DFM_BUILD_TESTING=OFF",
            "-DFM_BUILD_CLI=OFF",
            f"-DCMAKE_BUILD_TYPE={args.config}",
        ]
    )

    # 2. Build the shared library only.
    _run(
        [
            "cmake",
            "--build",
            str(build_dir),
            "--target",
            "formulon",
            "--config",
            args.config,
            "--parallel",
        ]
    )

    # 3. Copy artefacts into _lib/.
    PKG_LIB_DIR.mkdir(parents=True, exist_ok=True)

    src = _resolve_built_lib(build_dir, args.config)
    dst = PKG_LIB_DIR / src.name
    shutil.copy2(src, dst)
    print(f"staged libformulon -> {dst} ({dst.stat().st_size} bytes)")

    if platform.system() == "Windows":
        for lib in _windows_import_libs(build_dir, args.config):
            shutil.copy2(lib, PKG_LIB_DIR / lib.name)
            print(f"staged {lib.name} -> {PKG_LIB_DIR / lib.name}")

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
