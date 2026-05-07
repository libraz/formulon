#!/usr/bin/env python3
"""Locale detection helpers shared across oracle drivers.

Excel's COM / AppleScript bridge exposes the host's regional setting as a
phone-style country code (``Application.International(xlCountryCode)``,
constant 1). Locale-sensitive surfaces of the driver — the
``EnvironmentInfo.excel_locale`` field, the localized error-token names
returned by the ``Range.Text`` property — depend on that code.

This module centralises:

  - :data:`COUNTRY_CODE_TO_BCP47` — the int-keyed map from
    ``xlCountryCode`` to BCP-47 locale strings. Used both by the driver
    (for environment probes) and by ``cli.py`` (for the contributor probe
    that derives a target name).
  - :data:`LOCALIZED_ERROR_TO_CANONICAL` — the localized form of every
    Excel error token mapped back to its canonical English form. The
    driver normalises through this map so non-English locales surface
    the same ``#WERT!`` / ``#VALEUR!`` cells as ``#VALUE!`` in the golden
    JSON.
  - :func:`detect_locale_from_app` — small helper that probes a live
    xlwings ``App`` for its country code and returns BCP-47 (or ``None``
    when the probe fails / the code is unknown).

Keeping the maps here means a fix to a single mistranslated error token
or a missing country code is a one-line change that flows through both
the Mac and Windows drivers.
"""

from __future__ import annotations

from typing import Any, Dict, Optional


# ``Application.International(xlCountryCode)`` returns a phone-style
# country code rather than a locale identifier. The mapping below covers
# every ``status: wanted`` slot in targets.yaml plus the most common
# variants we expect contributors to bring. Codes we don't know fall
# through to ``None`` at the call site.
#
# Source: Microsoft "xlApplicationInternational" enumeration docs,
# cross-checked against contributor reports.
COUNTRY_CODE_TO_BCP47: Dict[int, str] = {
    1: "en-US",
    2: "en-CA",
    7: "ru-RU",
    27: "en-ZA",
    31: "nl-NL",
    32: "nl-BE",
    33: "fr-FR",
    34: "es-ES",
    36: "hu-HU",
    39: "it-IT",
    41: "de-CH",
    44: "en-GB",
    45: "da-DK",
    46: "sv-SE",
    47: "no-NO",
    48: "pl-PL",
    49: "de-DE",
    52: "es-MX",
    55: "pt-BR",
    58: "es-VE",
    60: "ms-MY",
    61: "en-AU",
    62: "id-ID",
    64: "en-NZ",
    65: "en-SG",
    66: "th-TH",
    81: "ja-JP",
    82: "ko-KR",
    84: "vi-VN",
    86: "zh-CN",
    90: "tr-TR",
    91: "hi-IN",
    351: "pt-PT",
    358: "fi-FI",
    420: "cs-CZ",
    886: "zh-TW",
    966: "ar-SA",
    972: "he-IL",
}


# Excel localises the seven legacy error tokens for some locales. The
# COM ``Range.Text`` property returns the localised form, so when the
# CVErr / .value paths fail we have to normalise through this table to
# write a canonical English token into the golden JSON.
#
# Coverage rationale: every locale listed as ``wanted`` in targets.yaml
# (de-DE, fr-FR, zh-CN, ko-KR, th-TH) plus the closely-related Latin
# locales contributors are likely to bring (es-ES, pt-PT/BR, nl-NL,
# it-IT). Identity entries are not needed because the canonical forms
# are already accepted by the caller via ``_ERR_DISPLAY_NAMES``.
#
# Cell of error semantics:
#   - #DIV/0!, #NULL! tend to render the symbol unchanged across most
#     Latin locales but get a leading punctuation mark in es-ES.
#   - #N/A renders as #NV in de-DE and #N/D in pt-PT/BR/it-IT.
#   - The CJK locales (ja-JP, ko-KR, zh-CN, th-TH) retain the English
#     tokens for the error names; their localisation is in the function
#     names and TEXT format codes, not in the error sentinels. We still
#     keep identity entries documented in case the locale flips.
LOCALIZED_ERROR_TO_CANONICAL: Dict[str, str] = {
    # de-DE
    "#WERT!": "#VALUE!",
    "#NV": "#N/A",
    "#ZAHL!": "#NUM!",
    "#BEZUG!": "#REF!",
    # fr-FR
    "#VALEUR!": "#VALUE!",
    "#NOM?": "#NAME?",
    "#NOMBRE!": "#NUM!",
    "#NUL!": "#NULL!",
    # nl-NL
    "#WAARDE!": "#VALUE!",
    "#N/B": "#N/A",
    "#NAAM?": "#NAME?",
    "#GETAL!": "#NUM!",
    "#VERW!": "#REF!",
    "#LEEG!": "#NULL!",
    "#DEEL/0!": "#DIV/0!",
    # es-ES (and most es-* variants)
    "#¡VALOR!": "#VALUE!",
    "#¡DIV/0!": "#DIV/0!",
    "#¿NOMBRE?": "#NAME?",
    "#¡NUM!": "#NUM!",
    "#¡REF!": "#REF!",
    "#¡NULO!": "#NULL!",
    # pt-PT / pt-BR / it-IT (each picks a subset of these)
    "#VALOR!": "#VALUE!",
    "#N/D": "#N/A",
    "#NOME?": "#NAME?",
    "#NULO!": "#NULL!",
    "#NÚM!": "#NUM!",  # pt: NÚM!
}


def detect_locale_from_app(app: Any) -> Optional[str]:
    """Returns a BCP-47 locale string for `app`, or ``None`` on failure.

    Probes ``Application.International(xlCountryCode)`` on the live
    xlwings ``App`` instance. xlwings exposes the COM/appscript
    ``International`` accessor as ``app.api.International`` on Windows
    and ``app.api.international`` on Mac (the appscript bridge
    lower-cases property names); we try both spellings before giving
    up.

    Returns ``None`` when:
      - the bridge raises (e.g., Excel is not yet attached),
      - the call returns a non-integer,
      - the integer is not in :data:`COUNTRY_CODE_TO_BCP47`.

    Callers should treat ``None`` as "fall back to the target's declared
    locale" rather than aborting; the goldens' value is still useful
    even if we can't pin the locale string.
    """

    cc: Optional[int] = None
    try:
        api = app.api
    except Exception:
        return None
    for attr in ("International", "international"):
        try:
            accessor = getattr(api, attr)
        except Exception:
            continue
        try:
            v = accessor(1) if callable(accessor) else accessor[1]
        except Exception:
            continue
        try:
            cc = int(v)
            break
        except (TypeError, ValueError):
            continue
    if cc is None:
        return None
    return COUNTRY_CODE_TO_BCP47.get(cc)


def normalise_error_token(text: str) -> Optional[str]:
    """Maps a localised Excel error token to its canonical English form.

    Returns ``None`` if `text` is neither a known canonical token nor a
    known localised one. The caller decides whether to treat that as a
    non-error (typical) or a parse failure (rare, surfaces as
    ``#UNKNOWN!`` in the golden).
    """

    if not text:
        return None
    if text in LOCALIZED_ERROR_TO_CANONICAL:
        return LOCALIZED_ERROR_TO_CANONICAL[text]
    return None
