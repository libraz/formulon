"""Host-injected function-metadata provider seam.

The calculation engine ships no human-readable function documentation:
:meth:`Workbook.function_metadata` returns only structural data (name,
arity, availability) and leaves ``signature_template`` / ``description``
as ``None``. That content is a display concern, so a host supplies its own
metadata document and merges it over the engine result at display time.

This module holds the pure, side-effect-free merge helper and the provider
type shapes. The document contract lives in
``docs/function-metadata-schema.md``. The metadata is display-only: it never
affects formula parsing or evaluation, and the localized names here are not
used for formula-language input localization.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Mapping, Optional, TypedDict, Union

from .workbook import FunctionMetadata

__all__ = [
    "FunctionMetadataEntry",
    "FunctionMetadataLocalized",
    "FunctionMetadataProvider",
    "MergedFunctionMetadata",
    "merge_function_metadata",
]


class FunctionMetadataLocalized(TypedDict, total=False):
    """Per-locale display overrides inside a :class:`FunctionMetadataEntry`."""

    signature: str
    description: str


class FunctionMetadataEntry(TypedDict, total=False):
    """One host-injected metadata entry.

    Keyed by canonical UPPERCASE function name inside a
    :data:`FunctionMetadataProvider`. Every field is optional.
    """

    #: Default (locale-agnostic) signature template.
    signature: str
    #: Default (locale-agnostic) description.
    description: str
    #: Map of BCP-47 locale tag -> localized display name.
    aliases: Dict[str, str]
    #: Map of BCP-47 locale tag -> per-locale signature/description overrides.
    localized: Dict[str, FunctionMetadataLocalized]


#: A whole host metadata document's ``functions`` map: canonical UPPERCASE
#: function name -> :class:`FunctionMetadataEntry`.
FunctionMetadataProvider = Dict[str, FunctionMetadataEntry]


@dataclass(frozen=True)
class MergedFunctionMetadata:
    """:class:`FunctionMetadata` with a resolved localized display name.

    Returned by :func:`merge_function_metadata` when a provider entry is
    supplied; carries the engine's structural fields plus the merged
    ``signature_template`` / ``description`` and the ``localized_name``.
    """

    name: str
    min_arity: int
    max_arity: Optional[int]
    availability: int
    signature_template: Optional[str]
    description: Optional[str]
    #: ``entry.aliases[locale]`` when present, else the canonical ``name``.
    localized_name: str


def _first(*candidates: Optional[str]) -> Optional[str]:
    """Return the first non-``None`` candidate, or ``None``."""
    for value in candidates:
        if value is not None:
            return value
    return None


def merge_function_metadata(
    base: FunctionMetadata,
    entry: Optional[Mapping[str, object]],
    locale: str,
) -> Union[FunctionMetadata, MergedFunctionMetadata]:
    """Merge a host-supplied ``entry`` over the engine result ``base``.

    Field precedence (first non-``None`` wins):

    - ``signature_template``: ``entry["localized"][locale]["signature"]`` ->
      ``entry["signature"]`` -> ``base.signature_template``
    - ``description``: ``entry["localized"][locale]["description"]`` ->
      ``entry["description"]`` -> ``base.description``
    - ``localized_name``: ``entry["aliases"][locale]`` -> ``base.name``

    Args:
      base: a :class:`FunctionMetadata` from
        :meth:`Workbook.function_metadata`.
      entry: the provider's ``functions[NAME]`` mapping, or ``None`` to
        leave ``base`` unchanged (``signature_template`` / ``description``
        stay ``None``).
      locale: a BCP-47 display locale tag (e.g. ``"fr-FR"``) matching the
        keys in ``aliases`` / ``localized``. Independent of the numeric
        locale code passed to :meth:`Workbook.function_metadata`.

    Returns:
      ``base`` verbatim when ``entry`` is ``None``; otherwise a
      :class:`MergedFunctionMetadata`.
    """
    if entry is None:
        return base

    localized_map = entry.get("localized") or {}
    localized = localized_map.get(locale) or {}
    signature = _first(
        localized.get("signature"),
        entry.get("signature"),  # type: ignore[arg-type]
        base.signature_template,
    )
    description = _first(
        localized.get("description"),
        entry.get("description"),  # type: ignore[arg-type]
        base.description,
    )
    aliases = entry.get("aliases") or {}
    localized_name = aliases.get(locale, base.name)
    return MergedFunctionMetadata(
        name=base.name,
        min_arity=base.min_arity,
        max_arity=base.max_arity,
        availability=base.availability,
        signature_template=signature,
        description=description,
        localized_name=localized_name,
    )
