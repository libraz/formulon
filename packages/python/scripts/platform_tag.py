#!/usr/bin/env python3
# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
"""Print the local setuptools platform tag for host-targeted wheel builds."""

from __future__ import annotations

import sysconfig


def main() -> int:
    tag = sysconfig.get_platform().replace("-", "_").replace(".", "_")
    print(tag)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
