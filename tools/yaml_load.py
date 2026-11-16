"""The one YAML loader of the repository tools.

PyYAML's safe constructor over libyaml's C scanner and parser when PyYAML was
built with them, and over the pure-Python ones otherwise. Both build the same
objects; the C path parses the oracle case and divergence files several times
faster. PyYAML is imported on the first load, so a tool that treats it as
optional can import this module and still see the ImportError where it calls.
"""

from __future__ import annotations

from typing import IO, Any, Union


def load_yaml(stream: Union[str, bytes, IO[str], IO[bytes]]) -> Any:
    """Parses one YAML document the way `yaml.safe_load` does."""
    import yaml

    return yaml.load(stream, Loader=getattr(yaml, "CSafeLoader", yaml.SafeLoader))
