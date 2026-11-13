#!/usr/bin/env python3
"""Regenerates the data section of src/builtin_style_names.cpp.

Reads tests/fixtures/cell_style_save/builtin_style_names.json (14 locales x 54
built-in cell style names) and rewrites the text between the
`BEGIN GENERATED` / `END GENERATED` markers of src/builtin_style_names.cpp.

Per locale row: the 30 non-accent names (ids 0..28 and 53) as one NUL-separated
blob with uint16 offsets, and the template of the 24 accent names (ids 29..52):
accent word, separator before the number, percent-prefix spacing and dash, and
the accent word used after the percent prefix. Locales with identical rows are
aliased. The script fails unless the template reproduces all 756 strings.

Run: python3 -I tools/dev/gen_builtin_style_names.py
"""

import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "tests/fixtures/cell_style_save/builtin_style_names.json"
TARGET = ROOT / "src/builtin_style_names.cpp"
BEGIN = "// BEGIN GENERATED (tools/dev/gen_builtin_style_names.py)"
END = "// END GENERATED"

# Locale order of the ExcelLocale enum.
ORDER = [
    "ja_JP",
    "en_US",
    "de_DE",
    "fr_FR",
    "zh_CN",
    "ko_KR",
    "th_TH",
    "es_ES",
    "es_MX",
    "pt_BR",
    "ru_RU",
    "zh_TW",
    "it_IT",
    "nl_NL",
]
SEPS = ["", " ", "\u00a0"]
DASHES = [" - ", " \u2014 "]
ACCENT_FIRST, ACCENT_LAST, EXPLANATORY = 29, 52, 53
PLAIN_IDS = list(range(ACCENT_FIRST)) + [EXPLANATORY]


def synth(t, idx):
    word, pword, sep, psep, dash = t
    n, k = idx // 4 + 1, idx % 4
    if k == 0:
        return f"{word}{SEPS[sep]}{n}"
    return f"{20 * k}{SEPS[psep]}%{DASHES[dash]}{pword}{SEPS[sep]}{n}"


def derive(names):
    first = names[ACCENT_FIRST]
    assert first.endswith("1"), first
    stem = first[:-1]
    sep = 0
    for i in (1, 2):
        if stem.endswith(SEPS[i]):
            sep = i
    word = stem[: len(stem) - len(SEPS[sep])]
    m = re.fullmatch(r"20(\s|\u00a0)?%( - | \u2014 )(.*)" + re.escape(SEPS[sep]) + "1", names[ACCENT_FIRST + 1])
    assert m, names[ACCENT_FIRST + 1]
    psep = SEPS.index(m.group(1) or "")
    return (word, m.group(3), sep, psep, DASHES.index(m.group(2)))


def cstr(s):
    return '"' + s.replace("\\", "\\\\").replace('"', '\\"') + '"'


def main():
    data = json.loads(FIXTURE.read_text(encoding="utf-8"))["locales"]
    assert list(data) == ORDER, list(data)
    rows, remap, checked = [], [], 0
    for loc in ORDER:
        names = data[loc]["names"]
        assert len(names) == 54
        t = derive(names)
        for idx in range(24):
            got = synth(t, idx)
            assert got == names[ACCENT_FIRST + idx], (loc, idx, got, names[ACCENT_FIRST + idx])
        checked += 24 + len(PLAIN_IDS)
        row = ([names[i] for i in PLAIN_IDS], t)
        if row not in rows:
            rows.append(row)
        remap.append(rows.index(row))
        print(f"{loc}: word={t[0]!r} pct_word={t[1]!r} sep={t[2]} pct_sep={t[3]} dash={t[4]}")
    assert checked == 756, checked

    out = [BEGIN, "// clang-format off", "const char kNameBlob[] ="]
    offsets, pos = [], 0
    for ri, (plain, _) in enumerate(rows):
        offs = []
        for s in plain:
            offs.append(pos)
            pos += len(s.encode("utf-8")) + 1
            out.append(f'    {cstr(s)} "\\0"')
        offsets.append(offs + [pos])
    out[-1] += ";"
    assert pos < 65536
    out.append("")
    out.append("const std::uint16_t kNameOffsets[%d][%d] = {" % (len(rows), len(PLAIN_IDS) + 1))
    for offs in offsets:
        out.append("    {" + ", ".join(map(str, offs)) + "},")
    out.append("};")
    out.append("")
    out.append("const AccentTemplate kAccentTemplates[%d] = {" % len(rows))
    for ri, (_, t) in enumerate(rows):
        out.append(f"    {{{cstr(t[0])}, {cstr(t[1])}, {t[2]}, {t[3]}, {t[4]}}},")
    out.append("};")
    out.append("")
    out.append("// ExcelLocale -> row; locales with identical rows share one.")
    out.append("const std::uint8_t kLocaleRow[%d] = {%s};" % (len(ORDER), ", ".join(map(str, remap))))
    out.append("// clang-format on")
    out.append(END)

    text = TARGET.read_text(encoding="utf-8")
    a, b = text.index(BEGIN), text.index(END) + len(END)
    TARGET.write_text(text[:a] + "\n".join(out) + text[b:], encoding="utf-8")
    print(f"ok: {checked} strings reproduced, {len(rows)} rows, blob {pos} B")


if __name__ == "__main__":
    sys.exit(main())
