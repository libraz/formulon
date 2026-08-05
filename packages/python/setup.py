"""setuptools glue for the Formulon Python wheel.

The wheel ships a single platform-agnostic ``formulon_capi.wasm`` plus
pure-Python source under ``formulon/``. ``wasmtime`` does the heavy
lifting at runtime, so the distribution itself is ``py3-none-any`` --
one wheel works on every interpreter and OS that has a wasmtime build.

Override::

    FORMULON_WHEEL_PURE=0   # build a platform-tagged wheel (legacy)

is intentionally NOT supported: the binding has no native components.
"""

from __future__ import annotations

from setuptools import setup

cmdclass = {}

try:
    from wheel.bdist_wheel import bdist_wheel as _bdist_wheel
except ModuleNotFoundError:
    _bdist_wheel = None

if _bdist_wheel is not None:

    class bdist_wheel(_bdist_wheel):
        def finalize_options(self) -> None:
            super().finalize_options()
            # Pure-Python: no platform tag. wasmtime is a runtime
            # dependency listed in pyproject.toml and pip resolves the
            # right binary wheel for the user's OS / arch.
            self.root_is_pure = True

        def get_tag(self) -> tuple[str, str, str]:
            return "py3", "none", "any"

    cmdclass["bdist_wheel"] = bdist_wheel


setup(cmdclass=cmdclass)
