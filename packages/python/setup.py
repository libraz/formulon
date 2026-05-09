# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
"""setuptools glue for platform-tagged Formulon wheels.

The Python code is pure ctypes, but the wheel bundles a native shared library.
Mark the distribution as non-pure so wheel filenames carry a platform tag
instead of py3-none-any.
"""

from __future__ import annotations

import os

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
            self.root_is_pure = False
            self.python_tag = "py3"
            plat_name = os.environ.get("FORMULON_PYTHON_PLAT_NAME")
            if plat_name:
                self.plat_name = plat_name

        def get_tag(self) -> tuple[str, str, str]:
            _, _, plat_name = super().get_tag()
            return "py3", "none", plat_name

    cmdclass["bdist_wheel"] = bdist_wheel


setup(cmdclass=cmdclass)
