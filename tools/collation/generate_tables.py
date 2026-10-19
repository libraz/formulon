#!/usr/bin/env python3
"""Generates the text collation tables from the Mac Excel measurements.

Inputs are the TSVs that tools/collation/capture.py writes under data/ (its
docstring documents them), data/unihan_total_strokes.tsv, and the code-page
cells of tools/dbcs/generate_tables.py, so a Han code here is the packed
lead/trail the runtime DBCS grids return.

Each locale is a base table (W from en_US, Z from zh_CN) plus a derived
delta: P remap segments, single-weight overrides, expansion changes and a Han
rule. Code points outside `EXACT_RANGES`, and on Z the ideographs outside
GB 2312 level 1, are approximated. `Collator` is the reference model that
src/eval/collation.cpp must match key for key; the probe-matrix cells it gets
wrong are written to data/fixtures/excused_<loc>.tsv.

Usage:
    python3 tools/collation/generate_tables.py          # rewrite the outputs
    python3 tools/collation/generate_tables.py --check  # exit 1 on drift
"""

from __future__ import annotations

import argparse
import bisect
import importlib.util
import shutil
import subprocess
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
DATA = REPO_ROOT / "tools" / "collation" / "data"
FIXTURES = DATA / "fixtures"
HEADER_PATH = REPO_ROOT / "src" / "eval" / "collation_table.h"
SOURCE_PATH = REPO_ROOT / "src" / "eval" / "collation_table.cpp"
DBCS_GENERATOR = REPO_ROOT / "tools" / "dbcs" / "generate_tables.py"
REGENERATION_GUIDANCE = "Regenerate with: python3 tools/collation/generate_tables.py"

# ExcelLocale enum order (src/excel_profile.h).
LOCALES = (
    "ja_JP", "en_US", "de_DE", "fr_FR", "zh_CN", "ko_KR", "th_TH",
    "es_ES", "es_MX", "pt_BR", "ru_RU", "zh_TW", "it_IT", "nl_NL",
)  # fmt: skip
Z_LOCALES = ("zh_CN", "zh_TW")
BASE_LOCALE = {"W": "en_US", "Z": "zh_CN"}

# Code points whose measured weights the tables reproduce; the rest are approximated.
EXACT_RANGES = (
    (0x0001, 0x04FF),  # Latin through Extended-B, IPA, combining marks, Greek, Cyrillic
    (0x0E00, 0x0E7F),  # Thai
    (0x1100, 0x11FF),  # Hangul jamo
    (0x1E00, 0x1EFF),  # Latin Extended Additional
    (0x2000, 0x21FF),  # general punctuation through arrows
    (0x2460, 0x24FF),  # enclosed alphanumerics
    (0x25A0, 0x26FF),  # geometric shapes, miscellaneous symbols
    (0x3000, 0x30FF),  # CJK symbols, hiragana, katakana
    (0x3130, 0x318F),  # Hangul compatibility jamo
    (0x3200, 0x32FF),  # enclosed CJK letters and months
    (0x4E00, 0x9FFF),  # CJK unified ideographs
    (0xAC00, 0xD7A3),  # Hangul syllables
    (0xF900, 0xFAFF),  # CJK compatibility ideographs
    (0xFF00, 0xFFEF),  # halfwidth and fullwidth forms
)
URO = (0x4E00, 0x9FFF)
COMBINING_MARKS = (0x0300, 0x036F)  # a base and a following mark compose canonically
SYLLABLES = (0xAC00, 0xD7A3)
COMPAT_IDEOGRAPHS = (0xF900, 0xFAFF)

FULL, MID, WORD, SEC, BASE = range(5)
CLASS_NAMES = ("full", "mid", "word", "sec", "base")
LEVELS = "=SDC"  # expansion level codes 0..3

# Locale rule bits; must match kRule* in collation_table.h.
RULE_KANA_MARKS = 1  # ja prolonged-sound and iteration marks
RULE_THAI_REORDER = 2  # Thai prefix vowel after its consonant (W)
RULE_HANGUL_JAMO = 4  # Hangul syllables decompose into jamo (W)
RULE_IGNORE_UNLISTED = 8  # code points outside the exact ranges are ignorable (Z)
RULE_REPEAT_MARKS = 16  # zh repeat marks take the preceding character (Z)

# Repeat-mark kinds; must match CollationRepeatKind in collation_table.h.
REPEAT_SAME, REPEAT_VOICED, REPEAT_VOWEL = range(3)

# Han rule kinds; must match CollationHanKind in collation_table.h.
HAN_BASE, HAN_CODEPAGE, HAN_STROKES, HAN_READINGS = range(4)
HAN_KIND_NAMES = ("kBase", "kCodepage", "kStrokes", "kReadings")
# Han ordering per locale: (kind, code page, first row, last row).
HAN_RULES = {
    "ja_JP": (HAN_CODEPAGE, "jis0208", 16, 84),
    "zh_CN": (HAN_CODEPAGE, "gb2312", 16, 55),
    "ko_KR": (HAN_READINGS, "ksx1001", 42, 93),
    "zh_TW": (HAN_STROKES, None, 0, 0),
}
CODEPAGE_ENUM = {None: "kNone", "jis0208": "kJis0208", "gb2312": "kGb2312", "ksx1001": "kKsX1001"}
# Above this many Unihan disagreements the zh_TW stroke order is encoded as measured.
STROKE_OVERRIDE_LIMIT = 500

MAX_WEIGHT = 253 * 254 + 253  # two base-254 digits
ALIGN_GAP_LIMIT = 64  # largest run of unused locale P values kept to avoid a remap segment
SIZE_AIM_BYTES = 22 * 1024  # leaves room for the collation code inside the feature budget
TABLE_BUDGET_KB = 25  # @size-budget annotation of the generated table
SIZE_BUDGET_BYTES = 28 * 1024  # Brotli growth allowed for the whole collation feature

COMPARE, DYNAMIC, EXACT = "compare", "dynamic", "exact"


def in_range(cp: int, rng: tuple[int, int]) -> bool:
    return rng[0] <= cp <= rng[1]


def is_exact(cp: int) -> bool:
    return any(lo <= cp <= hi for lo, hi in EXACT_RANGES)


def out_of_range_index(cp: int) -> int:
    """Rank of `cp` among the BMP code points outside EXACT_RANGES."""
    return cp - sum(max(0, min(hi + 1, cp) - lo) for lo, hi in EXACT_RANGES)


def fail(message: str):
    raise SystemExit(f"generate_tables: {message}")


# ---------------------------------------------------------------- inputs


def read_tsv(path: Path) -> list[list[str]]:
    lines = path.read_text(encoding="utf-8").split("\n")
    return [ln.split("\t") for ln in lines[1:] if ln]


def dec(text: str) -> str:
    return "".join(chr(int(t[2:], 16)) for t in text.split())


def enc(s: str) -> str:
    return " ".join(f"U+{ord(ch):04X}" for ch in s)


Weight = tuple  # (cls, p, s, d, c)


@dataclass
class Measured:
    """One locale's measured BMP order as per-character rank weights."""

    tag: str
    weights: dict[int, Weight]
    expanding: dict[int, tuple[tuple[int, ...], int]]  # cp -> (targets, level)
    order: list[int]


def load_measured(tag: str) -> Measured:
    weights: dict[int, Weight] = {}
    counters: dict[int, tuple[int, int, int, int]] = {}
    order, xs = [], set()
    for cp_s, cls_s, rel in read_tsv(DATA / f"order_{tag}.tsv"):
        cp = int(cp_s[2:], 16)
        cls = CLASS_NAMES.index(cls_s)
        order.append(cp)
        if rel == "X":
            xs.add(cp)
            continue
        if cls in (FULL, MID):
            weights[cp] = (cls, 0, 0, 0, 0)
            continue
        p, s, d, c = counters.get(cls, (0, 0, 0, 0))
        if rel == "P":
            p, s, d, c = p + 1, 0, 0, 0
        elif rel == "S":
            s, d, c = s + 1, 0, 0
        elif rel == "D":
            d, c = d + 1, 0
        elif rel == "C":
            c += 1
        counters[cls] = (p, s, d, c)
        weights[cp] = (cls, p, s, d, c)
    expanding = {}
    for cp_s, to, level in read_tsv(DATA / f"expansion_{tag}.tsv"):
        expanding[int(cp_s[2:], 16)] = (tuple(ord(ch) for ch in dec(to)), LEVELS.index(level))
    if set(expanding) != xs:
        fail(f"{tag}: X rows and the expansion table disagree")
    return Measured(tag, weights, expanding, order)


def load_strokes() -> dict[int, int]:
    return {int(cp[2:], 16): int(n) for cp, n in read_tsv(DATA / "unihan_total_strokes.tsv")}


_DBCS = None


def codepage_cells(name: str, first_row: int, last_row: int) -> list[tuple[int, int]]:
    """(packed code, cp) of the unified-ideograph cells in rows first..last, in code order."""
    global _DBCS
    if _DBCS is None:
        spec = importlib.util.spec_from_file_location("formulon_dbcs_generator_for_collation", DBCS_GENERATOR)
        if spec is None or spec.loader is None:
            fail(f"cannot load {DBCS_GENERATOR}")
        _DBCS = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = _DBCS
        spec.loader.exec_module(_DBCS)
    table = next(t for t in _DBCS.TABLES if t.name == name)
    out = []
    for (lead, trail), cp in zip(table.shape.positions(), _DBCS.cells_for(table)):
        if first_row <= lead <= last_row and in_range(cp, URO):
            out.append(((lead << 8) | trail, cp))
    return out


# ---------------------------------------------------------------- locale spaces


@dataclass
class HanRule:
    kind: int = HAN_BASE
    codepage: str | None = None
    first_row: int = 0
    last_row: int = 0
    listed: list[tuple[int, int]] = field(default_factory=list)  # codepage: (code, p) in p order
    groups: dict[int, int] = field(default_factory=dict)  # strokes: cp -> group
    group_p: list[int] = field(default_factory=list)  # strokes: p in (group, cp) order
    readings: list[tuple[int, int]] = field(default_factory=list)  # readings: (syllable or 0, cells)
    # readings: cp -> (syllable, rank) for ideographs outside the code page
    unlisted: dict[int, tuple[int, int]] = field(default_factory=dict)
    aliases: dict[int, int] = field(default_factory=dict)  # readings: cp -> equal ideograph of its reading

    def positions(self) -> dict[int, int]:
        """cp -> locale P of every ideograph the rule places as a single primary."""
        if self.kind == HAN_CODEPAGE:
            code_cp = dict(codepage_cells(self.codepage, self.first_row, self.last_row))
            return {code_cp[code]: p for code, p in self.listed}
        if self.kind == HAN_STROKES:
            order = sorted(self.groups, key=lambda cp: (self.groups[cp], cp))
            return dict(zip(order, self.group_p))
        return {}


@dataclass
class Space:
    """A locale's target weights over the code points it carries as single elements."""

    weights: dict[int, Weight]
    han: HanRule
    approximated: set[int]
    computed: set[int]  # ko ideographs placed by their reading
    unihan_misses: int = 0


def canonical_alias(cp: int) -> int | None:
    if not in_range(cp, COMPAT_IDEOGRAPHS):
        return None
    d = unicodedata.normalize("NFD", chr(cp))
    if len(d) == 1 and d != chr(cp) and in_range(ord(d), URO):
        return ord(d)
    return None


def alias_holds(m: Measured, cp: int) -> bool:
    t = canonical_alias(cp)
    return t is not None and cp in m.weights and t in m.weights and m.weights[cp] == m.weights[t]


def dense_p(m: Measured, kept: set[int]) -> dict[int, Weight]:
    """Measured weights of `kept`, P re-ranked densely from 1 per class."""
    out = {}
    for cls in (WORD, BASE):
        ps = sorted({m.weights[cp][1] for cp in kept if m.weights[cp][0] == cls})
        rank = {p: i + 1 for i, p in enumerate(ps)}
        for cp in kept:
            w = m.weights[cp]
            if w[0] == cls:
                out[cp] = (cls, rank[w[1]], *w[2:])
    for cp in kept:
        out.setdefault(cp, m.weights[cp])
    return out


def codepage_rule(tag: str, w: dict[int, Weight]) -> HanRule:
    """Code-page ideographs in measured order; a tie after the first falls back to an override."""
    kind, page, first, last = HAN_RULES[tag]
    rule = HanRule(kind, page, first, last)
    entries = sorted((w[cp][1], code) for code, cp in codepage_cells(page, first, last) if cp in w)
    code_cp = dict(codepage_cells(page, first, last))
    seen = set()
    for p, code in entries:
        if w[code_cp[code]][:1] == (BASE,) and w[code_cp[code]][2:] == (0, 0, 0) and p not in seen:
            seen.add(p)
            rule.listed.append((code, p))
    return rule


def strokes_rule(w: dict[int, Weight], strokes: dict[int, int]) -> tuple[HanRule, int]:
    """zh_TW: measured stroke groups, code point order inside a group.

    Returns the rule and how many ideographs the Unihan (kTotalStrokes, cp) order misplaces.
    """
    heads = sorted((v[1], cp) for cp, v in w.items() if in_range(cp, URO) and v[0] == BASE and v[2:] == (0, 0, 0))
    seq, seen = [], set()
    for p, cp in heads:
        if p not in seen:
            seen.add(p)
            seq.append((p, cp))
    rule = HanRule(HAN_STROKES)
    group = 0
    for i, (_, cp) in enumerate(seq):
        if i == 0 or cp < seq[i - 1][1]:
            group += 1
        rule.groups[cp] = group
    rule.group_p = [p for p, _ in seq]
    by_unihan = sorted(rule.groups, key=lambda cp: (strokes.get(cp, 99), cp))
    rank = {cp: i for i, cp in enumerate(by_unihan)}
    return rule, len(seq) - longest_increasing([rank[cp] for _, cp in seq])


def longest_increasing(values: list[int]) -> int:
    tails: list[int] = []
    for v in values:
        i = bisect.bisect_left(tails, v)
        if i == len(tails):
            tails.append(v)
        else:
            tails[i] = v
    return len(tails)


def readings_rule(m: Measured) -> tuple[HanRule, set[int]]:
    """ko: reading runs over the KS ideograph cells in code order, plus (cp, reading) outside KS.

    An ideograph in the run that directly follows a syllable in the measured order reads
    as that syllable (the measured P is unusable there: the capture's primary test
    cannot see a break while Hangul sorts before Latin). Its rank inside the reading is
    its KS position, then code point order after the KS ones.
    """
    kind, page, first, last = HAN_RULES["ko_KR"]
    hanja: dict[int, tuple[int, int]] = {}
    aliases: dict[int, int] = {}
    syllable, rank, head = 0, 0, 0
    for cp in m.order:
        if in_range(cp, SYLLABLES):
            syllable, rank, head = cp, 0, 0
        elif in_range(cp, URO) and syllable:
            if head and m.weights[cp] == m.weights[head]:
                aliases[cp] = head  # measured equal to the previous ideograph of the reading
                continue
            rank, head = rank + 1, cp
            hanja[cp] = (syllable, rank)
        elif not in_range(cp, COMPAT_IDEOGRAPHS):
            syllable = 0
    rule = HanRule(kind, page, first, last)
    cells = codepage_cells(page, first, last)
    for _, cp in cells:
        syl = hanja[cp][0] if cp in hanja else 0
        if rule.readings and rule.readings[-1][0] == syl:
            rule.readings[-1] = (syl, rule.readings[-1][1] + 1)
        else:
            rule.readings.append((syl, 1))
    ks_count: Counter[int] = Counter()
    for _, cp in cells:
        if cp in hanja:
            ks_count[hanja[cp][0]] += 1
            if hanja[cp][1] != ks_count[hanja[cp][0]]:
                fail(f"ko_KR: reading {chr(hanja[cp][0])} is not in KS order")
    ks = {cp for _, cp in cells}
    rule.unlisted = {cp: (syl, rank - ks_count[syl]) for cp, (syl, rank) in sorted(hanja.items()) if cp not in ks}
    rule.aliases = aliases
    return rule, set(hanja) | set(aliases)


def locale_space(tag: str, m: Measured, han_rank_size: int, strokes: dict[int, int]) -> Space:
    base = "Z" if tag in Z_LOCALES else "W"
    kept = set()
    for cp in m.weights:
        if not is_exact(cp) or alias_holds(m, cp):
            continue
        if base == "W" and in_range(cp, SYLLABLES):
            continue
        kept.add(cp)
    computed, approximated = set(), set()
    if HAN_RULES.get(tag, (HAN_BASE,))[0] == HAN_READINGS:
        han, computed = readings_rule(m)
        kept -= computed
    if tag == BASE_LOCALE["Z"]:
        listed = {cp for _, cp in codepage_cells(*HAN_RULES[tag][1:])}
        approximated = {cp for cp in kept if in_range(cp, URO) and m.weights[cp][0] == BASE and cp not in listed}
        kept -= approximated
        # Compatibility ideographs equal to an approximated one share its position.
        approximated |= {cp for cp in m.weights if alias_holds(m, cp) and canonical_alias(cp) in approximated}
    w = dense_p(m, kept)
    unihan_misses = 0
    kind = HAN_RULES.get(tag, (HAN_BASE,))[0]
    if kind == HAN_CODEPAGE:
        han = codepage_rule(tag, w)
        if tag == BASE_LOCALE["Z"]:
            # The approximated ideographs form one block, in W order, after the last listed one.
            rest_p = max(p for _, p in han.listed) + 1
            w = {
                cp: (v[0], v[1] + han_rank_size if v[0] == BASE and v[1] >= rest_p else v[1], *v[2:])
                for cp, v in w.items()
            }
            han = codepage_rule(tag, w)
    elif kind == HAN_STROKES:
        han, unihan_misses = strokes_rule(w, strokes)
        if unihan_misses <= STROKE_OVERRIDE_LIMIT:
            fail("zh_TW: Unihan order is close enough; encode it as Unihan plus overrides instead")
    elif kind == HAN_BASE:
        han = HanRule()
    return Space(w, han, approximated, computed, unihan_misses)


# ---------------------------------------------------------------- tables


@dataclass
class BaseTable:
    name: str
    elements: dict[int, Weight]
    expansions: dict[int, tuple[tuple[int, ...], int]]
    han_rest_p: int
    out_p: int  # W: base P of the first out-of-range code point; 0 on Z
    repeat_marks: dict[int, int] = field(default_factory=dict)  # Z: cp -> repeat kind


@dataclass
class LocaleTable:
    tag: str
    base: str
    rules: int
    remap: list[tuple[int, int]]  # (base P start, locale P start)
    overrides: dict[int, Weight]
    expansions: dict[int, tuple[tuple[int, ...], int]]
    han: HanRule


@dataclass
class Tables:
    bases: dict[str, BaseTable]
    han_rank: dict[int, int]  # unified ideograph -> rank in the W (radical-stroke) order
    locales: dict[str, LocaleTable]
    compositions: dict[tuple[int, int], int]  # (base, combining mark) -> composed character


def build_compositions() -> dict[tuple[int, int], int]:
    """Canonical compositions of a base and a combining mark whose result is reproduced exactly."""
    out = {}
    for cp in range(0x10000):
        if not is_exact(cp):
            continue
        decomposition = unicodedata.decomposition(chr(cp))
        if decomposition.startswith("<"):
            continue
        parts = [int(x, 16) for x in decomposition.split()]
        if len(parts) == 2 and in_range(parts[1], COMBINING_MARKS):
            if unicodedata.normalize("NFC", chr(parts[0]) + chr(parts[1])) == chr(cp):
                out[(parts[0], parts[1])] = cp
    return out


def base_expansions(m: Measured) -> dict[int, tuple[tuple[int, ...], int]]:
    out = {cp: e for cp, e in m.expanding.items() if is_exact(cp)}
    for cp in range(COMPAT_IDEOGRAPHS[0], COMPAT_IDEOGRAPHS[1] + 1):
        if alias_holds(m, cp):
            out[cp] = ((canonical_alias(cp),), 0)
    return out


def build_han_rank(en: Measured) -> dict[int, int]:
    space = locale_space("en_US", en, 0, {})
    w = space.weights
    uro = [cp for cp in w if in_range(cp, URO)]
    if len(uro) != URO[1] - URO[0] + 1 or any(w[cp][0] != BASE or w[cp][2:] != (0, 0, 0) for cp in uro):
        fail("en_US: every unified ideograph must be a primary head")
    lo = min(w[cp][1] for cp in uro)
    return {cp: w[cp][1] - lo for cp in uro}


def build_base(name: str, m: Measured, space: Space) -> BaseTable:
    w = space.weights
    listed = set(space.han.positions())
    if name == "W":
        elements = {cp: v for cp, v in w.items() if not in_range(cp, URO)}
        han_rest_p = min(w[cp][1] for cp in w if in_range(cp, URO))
        out_p = max(v[1] for v in w.values() if v[0] == BASE) + 1
    else:
        elements = {cp: v for cp, v in w.items() if cp not in listed}
        han_rest_p = max(p for _, p in space.han.listed) + 1
        out_p = 0
    table = BaseTable(name, elements, base_expansions(m), han_rest_p, out_p)
    if name == "Z":
        table.repeat_marks = repeat_marks(m)
    return table


def repeat_marks(m: Measured) -> dict[int, int]:
    """zh: the characters sharing the last primary of the order; after another character
    each stands for it (the vowel for a prolonged-sound mark, voiced for a voiced
    iteration mark) plus a secondary difference."""
    last = max(w[1] for w in m.weights.values() if w[0] == BASE)
    out = {}
    for cp, w in m.weights.items():
        if w[0] == BASE and w[1] == last and is_exact(cp):
            name = unicodedata.name(chr(cp), "")
            kind = REPEAT_VOWEL if "PROLONGED SOUND MARK" in name else REPEAT_SAME
            out[cp] = REPEAT_VOICED if "VOICED ITERATION MARK" in name else kind
    return out


def remap_of(segments: list[tuple[int, int]], p: int) -> int:
    i = bisect.bisect_right(segments, (p, float("inf"))) - 1
    b, loc = segments[i]
    return loc + (p - b)


def first_diff(a: tuple, b: tuple) -> int:
    return next((i for i in range(len(a)) if a[i] != b[i]), len(a))


def derive_locale(tag: str, m: Measured, space: Space, base: BaseTable, han_rank: dict[int, int]) -> LocaleTable:
    """Remap segments, overrides and expansion changes reproducing `space` on top of `base`."""
    rules = RULE_IGNORE_UNLISTED | RULE_REPEAT_MARKS if base.name == "Z" else RULE_THAI_REORDER | RULE_HANGUL_JAMO
    if tag == "ja_JP":
        rules |= RULE_KANA_MARKS

    def candidate(cp: int):
        if cp in base.expansions:
            return None
        if cp in base.elements:
            return base.elements[cp]
        if in_range(cp, URO):
            return (BASE, base.han_rest_p + han_rank[cp], 0, 0, 0)
        return None

    # The base locale keeps its own spacing, including the block left for approximated ideographs.
    limit = float("inf") if tag == BASE_LOCALE[base.name] else ALIGN_GAP_LIMIT
    w, listed_p, segments = align_primaries(space, base, candidate, limit)
    han = space.han
    if base.name == "W":
        top = max(v[1] for v in w.values() if v[0] == BASE)
        if remap_of(segments, base.out_p) != top + 1:
            segments.append((base.out_p, top + 1))

    overrides: dict[int, Weight] = {}
    assigned: dict[int, Weight] = {}
    for cp, v in w.items():
        if cp in listed_p:
            assigned[cp] = (BASE, listed_p[cp], 0, 0, 0)
            continue
        c = candidate(cp)
        if c is not None and c[0] == BASE:
            c = (BASE, remap_of(segments, c[1]), *c[2:])
        if c is None or c[0] != v[0] or (c[0] in (BASE, WORD) and c[1] != v[1]):
            overrides[cp] = v
        else:
            assigned[cp] = c
    groups: dict[tuple, list[int]] = defaultdict(list)
    for cp, v in w.items():
        groups[(v[0], v[1]) if v[0] in (BASE, WORD) else (v[0],)].append(cp)
    for key, members in groups.items():
        if key[0] in (FULL, MID):
            continue
        members.sort(key=lambda cp: (w[cp][2:], cp))
        if not fit_group(members, w, assigned, overrides):
            for cp in members:
                if cp not in listed_p:
                    overrides[cp] = assigned[cp] = w[cp]
    expansions = {}
    for cp, e in m.expanding.items():
        if is_exact(cp) and base.expansions.get(cp) != e:
            expansions[cp] = e
    for cp in range(COMPAT_IDEOGRAPHS[0], COMPAT_IDEOGRAPHS[1] + 1):
        if alias_holds(m, cp) and cp not in base.expansions:
            expansions[cp] = ((canonical_alias(cp),), 0)
    for cp, target in space.han.aliases.items():
        expansions[cp] = ((target,), 0)
    return LocaleTable(tag, base.name, rules, segments, overrides, expansions, han)


def align_primaries(
    space: Space, base: BaseTable, candidate, gap_limit: float
) -> tuple[dict[int, Weight], dict[int, int], list]:
    """Locale P values that keep the base P spacing, so moved characters leave gaps
    instead of shifting their neighbours; returns weights, Han rule P and remap segments.

    Walking the locale's primaries in order, a primary whose members mostly come from
    base P b takes b + offset; the offset changes only when that would not ascend or
    would skip more than `gap_limit` values.
    """
    w = space.weights
    members: dict[int, list[int]] = defaultdict(list)
    for cp, v in w.items():
        if v[0] == BASE:
            members[v[1]].append(cp)
    listed = space.han.positions()
    value, offsets = {}, {}
    last, off = 0, None
    for p in sorted(members):
        votes = Counter(c[1] for cp in members[p] if cp not in listed and (c := candidate(cp)) and c[0] == BASE)
        b = votes.most_common(1)[0][0] if votes else None
        if b is not None and b not in offsets:
            if off is None or not last < b + off <= last + gap_limit:
                off = last + 1 - b
            offsets[b] = off
            value[p] = b + off
        else:
            value[p] = last + 1
        last = value[p]
    segments = [(0, 0)]
    for b in sorted(offsets):
        if remap_of(segments, b) != b + offsets[b]:
            segments.append((b, b + offsets[b]))
    out = {cp: (v[0], value[v[1]], *v[2:]) if v[0] == BASE else v for cp, v in w.items()}
    han = space.han
    if han.kind == HAN_CODEPAGE:
        han.listed = [(code, value[p]) for code, p in han.listed]
    elif han.kind == HAN_STROKES:
        han.group_p = [value[p] for p in han.group_p]
    return out, han.positions(), segments


def fit_group(members: list[int], w: dict, assigned: dict, overrides: dict) -> bool:
    """Gives each override of a sorted group a sub-primary tuple consistent with its neighbours."""
    prev = None
    for i, cp in enumerate(members):
        sub = w[cp][2:]
        if cp not in overrides:
            tup = assigned[cp][2:]
            if prev is not None:
                psub, ptup = prev
                if first_diff(psub, sub) != first_diff(ptup, tup) or (psub != sub and not ptup < tup):
                    return False
            prev = (sub, tup)
            continue
        if prev is None:
            tup = sub
        elif prev[0] == sub:
            tup = prev[1]
        else:
            level = first_diff(prev[0], sub)
            tup = prev[1][:level] + (prev[1][level] + 1,) + (0,) * (2 - level)
        nxt = next((x for x in members[i + 1 :] if x not in overrides), None)
        if nxt is not None:
            nsub, ntup = w[nxt][2:], assigned[nxt][2:]
            if first_diff(sub, nsub) != first_diff(tup, ntup) or (sub != nsub and not tup < ntup):
                return False
        overrides[cp] = assigned[cp] = (w[cp][0], w[cp][1], *tup)
        prev = (sub, tup)
    return True


@dataclass
class Build:
    tables: Tables
    measured: dict[str, Measured]
    spaces: dict[str, Space]


def build() -> Build:
    measured = {tag: load_measured(tag) for tag in LOCALES}
    strokes = load_strokes()
    han_rank = build_han_rank(measured["en_US"])
    rank_size = max(han_rank.values()) + 1
    spaces = {tag: locale_space(tag, measured[tag], rank_size, strokes) for tag in LOCALES}
    bases = {name: build_base(name, measured[tag], spaces[tag]) for name, tag in BASE_LOCALE.items()}
    locales = {}
    for tag in LOCALES:
        base = bases["Z" if tag in Z_LOCALES else "W"]
        locales[tag] = derive_locale(tag, measured[tag], spaces[tag], base, han_rank)
    for name, tag in BASE_LOCALE.items():
        t = locales[tag]
        if t.overrides or t.expansions or len(t.remap) > 1 + (name == "W"):
            fail(f"{tag}: the base locale must reproduce its own table unchanged")
    tables = Tables(bases, han_rank, locales, build_compositions())
    check_ranges(tables)
    return Build(tables, measured, spaces)


def check_ranges(t: Tables) -> None:
    """Every weight must fit two base-254 digits; never truncate."""
    values = []
    for b in t.bases.values():
        values += [x for v in b.elements.values() for x in v[1:]] + [b.han_rest_p, b.out_p]
        values.append(b.han_rest_p + max(t.han_rank.values()))
        if b.name == "W":
            values.append(b.out_p + out_of_range_index(0xFFFF))
    for loc in t.locales.values():
        base = t.bases[loc.base]
        values += [x for v in loc.overrides.values() for x in v[1:]]
        values += [p for _, p in loc.han.listed] + loc.han.group_p
        values += [remap_of(loc.remap, v[1]) for v in base.elements.values() if v[0] == BASE]
        values += [remap_of(loc.remap, base.han_rest_p + r) for r in set(t.han_rank.values())]
        if loc.base == "W":
            values += [remap_of(loc.remap, base.out_p + i) for i in (0, out_of_range_index(0xFFFF))]
        # A ko ideograph adds its rank inside a reading to the last jamo's S.
        ranks = Counter(syl for syl, _ in loc.han.unlisted.values())
        ranks.update({syl: n for syl, n in loc.han.readings if syl})
        if ranks:
            values.append(max(base.elements[j][2] for j in range(0x1161, 0x11C3)) + max(ranks.values()))
    bad = [v for v in values if not 0 <= v <= MAX_WEIGHT]
    if bad:
        fail(f"weight {bad[0]} does not fit two base-254 digits")


# ---------------------------------------------------------------- encoding


def zigzag(n: int) -> int:
    return 2 * n if n >= 0 else -2 * n - 1


def unzigzag(n: int) -> int:
    return n >> 1 if n % 2 == 0 else -((n + 1) >> 1)


def leb128(values: list[int]) -> bytes:
    out = bytearray()
    for n in values:
        if n < 0:
            fail(f"negative varint {n}")
        while n >= 0x80:
            out.append(n & 0x7F | 0x80)
            n >>= 7
        out.append(n)
    return bytes(out)


def unleb128(data: bytes) -> list[int]:
    out, n, shift = [], 0, 0
    for b in data:
        n |= (b & 0x7F) << shift
        shift += 7
        if not b & 0x80:
            out.append(n)
            n, shift = 0, 0
    return out


@dataclass
class Encoded:
    """Named varint columns plus the record count, as one C++ table struct."""

    count: int
    columns: dict[str, bytes]


def encode_runs(entries: dict[int, Weight]) -> tuple[Encoded, Encoded]:
    """Element runs plus the sparse D/C table.

    A run is consecutive code points with one class and S, and P stepping by 0 or 1.
    """
    cps = sorted(entries)
    cols = {k: [] for k in ("gap", "length", "p", "s", "cls_step")}
    prev_end, prev_p, i, count = 0, 0, 0, 0
    while i < len(cps):
        cp = cps[i]
        cls, p, s = entries[cp][:3]
        j, step = i + 1, None
        while j < len(cps) and cps[j] == cps[j - 1] + 1 and entries[cps[j]][0] == cls and entries[cps[j]][2] == s:
            st = entries[cps[j]][1] - entries[cps[j - 1]][1]
            if step is None and st in (0, 1):
                step = st
            elif st != step:
                break
            j += 1
        step = step or 0
        n = j - i
        cols["gap"].append(cp - prev_end)
        cols["length"].append(n - 1)
        cols["p"].append(zigzag(p - prev_p))
        cols["s"].append(s)
        cols["cls_step"].append(cls | step << 3)
        prev_end, prev_p = cp + n, p + (n - 1) * step
        count += 1
        i = j
    tert = {"gap": [], "d": [], "c": []}
    prev = 0
    for cp in cps:
        d, c = entries[cp][3:]
        if d or c:
            tert["gap"].append(cp - prev)
            tert["d"].append(d)
            tert["c"].append(c)
            prev = cp
    return (
        Encoded(count, {k: leb128(v) for k, v in cols.items()}),
        Encoded(len(tert["gap"]), {k: leb128(v) for k, v in tert.items()}),
    )


def decode_runs(runs: Encoded, tert: Encoded) -> dict[int, Weight]:
    cols = {k: unleb128(v) for k, v in runs.columns.items()}
    out = {}
    prev_end, prev_p = 0, 0
    for i in range(runs.count):
        first = prev_end + cols["gap"][i]
        n = cols["length"][i] + 1
        p = prev_p + unzigzag(cols["p"][i])
        cls, step = cols["cls_step"][i] & 7, cols["cls_step"][i] >> 3
        for k in range(n):
            out[first + k] = (cls, p + k * step, cols["s"][i], 0, 0)
        prev_end, prev_p = first + n, p + (n - 1) * step
    t = {k: unleb128(v) for k, v in tert.columns.items()}
    cp = 0
    for i in range(tert.count):
        cp += t["gap"][i]
        out[cp] = out[cp][:3] + (t["d"][i], t["c"][i])
    return out


def encode_expansions(exps: dict[int, tuple[tuple[int, ...], int]]) -> Encoded:
    cols = {"gap": [], "head": [], "targets": []}
    prev, first = 0, 0
    for cp in sorted(exps):
        targets, level = exps[cp]
        cols["gap"].append(cp - prev)
        cols["head"].append(level | (len(targets) - 1) << 2)
        last = first
        for t in targets:
            cols["targets"].append(zigzag(t - last))
            last = t
        prev, first = cp, targets[0]
    return Encoded(len(exps), {k: leb128(v) for k, v in cols.items()})


def decode_expansions(e: Encoded) -> dict[int, tuple[tuple[int, ...], int]]:
    cols = {k: unleb128(v) for k, v in e.columns.items()}
    out, cp, k, first = {}, 0, 0, 0
    for i in range(e.count):
        cp += cols["gap"][i]
        level, n = cols["head"][i] & 3, (cols["head"][i] >> 2) + 1
        targets, last = [], first
        for _ in range(n):
            last += unzigzag(cols["targets"][k])
            targets.append(last)
            k += 1
        out[cp] = (tuple(targets), level)
        first = targets[0]
    return out


def encode_pairs(pairs: list[tuple[int, int]], names: tuple[str, str], modes: str) -> Encoded:
    """Two-column records. Mode per column: u = delta, z = zigzag delta, n = count - 1, a = value."""
    cols = {names[0]: [], names[1]: []}
    prev = [0, 0]
    for pair in pairs:
        for k in (0, 1):
            v = pair[k]
            if modes[k] in "na":
                cols[names[k]].append(v - (modes[k] == "n"))
                continue
            delta = v - prev[k]
            cols[names[k]].append(zigzag(delta) if modes[k] == "z" else delta)
            prev[k] = v
    return Encoded(len(pairs), {k: leb128(v) for k, v in cols.items()})


def decode_pairs(e: Encoded, names: tuple[str, str], modes: str) -> list[tuple[int, int]]:
    cols = [unleb128(e.columns[n]) for n in names]
    out, prev = [], [0, 0]
    for i in range(e.count):
        row = []
        for k in (0, 1):
            if modes[k] in "na":
                row.append(cols[k][i] + (modes[k] == "n"))
                continue
            prev[k] += unzigzag(cols[k][i]) if modes[k] == "z" else cols[k][i]
            row.append(prev[k])
        out.append(tuple(row))
    return out


REMAP = (("base_p", "offset"), "uz")
READING_RUNS = (("syllable", "length"), "zn")
UNLISTED = (("syllable", "count"), "un")


def encode_step_runs(pairs: list[tuple[int, int]]) -> Encoded:
    """(slot, p) pairs in p order as runs where both advance by 1.

    Per run: slot = zigzag(first slot - previous run's last slot - 1), p = first p -
    previous run's last p (both initially -1 and 0), length = pairs - 1.
    """
    cols = {"slot": [], "p": [], "length": []}
    last_slot, last_p, i = -1, 0, 0
    while i < len(pairs):
        j = i + 1
        while j < len(pairs) and pairs[j] == (pairs[j - 1][0] + 1, pairs[j - 1][1] + 1):
            j += 1
        slot, p = pairs[i]
        cols["slot"].append(zigzag(slot - last_slot - 1))
        cols["p"].append(p - last_p)
        cols["length"].append(j - i - 1)
        last_slot, last_p, i = pairs[j - 1][0], pairs[j - 1][1], j
    return Encoded(len(cols["slot"]), {k: leb128(v) for k, v in cols.items()})


def decode_step_runs(e: Encoded) -> list[tuple[int, int]]:
    cols = {k: unleb128(v) for k, v in e.columns.items()}
    out, last_slot, last_p = [], -1, 0
    for i in range(e.count):
        slot, p = last_slot + 1 + unzigzag(cols["slot"][i]), last_p + cols["p"][i]
        out += [(slot + k, p + k) for k in range(cols["length"][i] + 1)]
        last_slot, last_p = out[-1]
    return out


def code_slot(code: int, first_row: int) -> int:
    """Index of a 94x94 grid cell counted from the first row of a Han rule."""
    return ((code >> 8) - first_row) * 94 + (code & 0xFF) - 1


def slot_code(slot: int, first_row: int) -> int:
    return (first_row + slot // 94) << 8 | (slot % 94 + 1)


def encode_unlisted(unlisted: dict[int, tuple[int, int]]) -> tuple[Encoded, Encoded]:
    """ko ideographs outside the code page, grouped by reading, in rank order inside a group."""
    groups: dict[int, list[tuple[int, int]]] = defaultdict(list)
    for cp, (syl, rank) in unlisted.items():
        groups[syl].append((rank, cp))
    cps = []
    for syl in sorted(groups):
        ranked = sorted(groups[syl])
        if [r for r, _ in ranked] != list(range(1, len(ranked) + 1)):
            fail("ko_KR: ranks outside the code page must be 1..n per reading")
        last = URO[0]
        for _, cp in ranked:
            cps.append(zigzag(cp - last))
            last = cp
    heads = encode_pairs([(syl, len(groups[syl])) for syl in sorted(groups)], *UNLISTED)
    return heads, Encoded(len(cps), {"cp": leb128(cps)})


def decode_unlisted(heads: Encoded, cps: Encoded) -> dict[int, tuple[int, int]]:
    deltas = iter(unleb128(cps.columns["cp"]))
    out = {}
    for syl, n in decode_pairs(heads, *UNLISTED):
        last = URO[0]
        for rank in range(1, n + 1):
            last += unzigzag(next(deltas))
            out[last] = (syl, rank)
    return out


REPEAT_MARKS = (("cp", "kind"), "ua")


def encode_remap(segments: list[tuple[int, int]]) -> Encoded:
    # Offsets are stored relative to the base start so identity segments are zero.
    return encode_pairs([(base_p, loc_p - base_p) for base_p, loc_p in segments], *REMAP)


def decode_remap(e: Encoded) -> list[tuple[int, int]]:
    return [(b, b + off) for b, off in decode_pairs(e, *REMAP)]


def encode_strokes(rule: HanRule) -> tuple[Encoded, Encoded]:
    cps = sorted(rule.groups)
    cols = {"gap": [], "length": [], "group": []}
    last_group = max(rule.groups.values())
    prev_end, prev_group, i = URO[0], 0, 0
    while i < len(cps):
        j = i + 1
        while j < len(cps) and cps[j] == cps[j - 1] + 1 and rule.groups[cps[j]] == rule.groups[cps[i]]:
            j += 1
        group = rule.groups[cps[i]]
        cols["gap"].append(cps[i] - prev_end)
        cols["length"].append(j - i - 1)
        # The last group (ideographs ordered after every stroke count) is 0; others chain.
        if group == last_group:
            cols["group"].append(0)
        else:
            cols["group"].append(zigzag(group - prev_group) + 1)
            prev_group = group
        prev_end, i = cps[j - 1] + 1, j
    runs = {"p": [], "length": []}
    last, i, ps = 0, 0, rule.group_p
    while i < len(ps):
        j = i + 1
        while j < len(ps) and ps[j] == ps[j - 1] + 1:
            j += 1
        runs["p"].append(ps[i] - last)
        runs["length"].append(j - i - 1)
        last, i = ps[j - 1], j
    return (
        Encoded(len(cols["gap"]), {k: leb128(v) for k, v in cols.items()}),
        Encoded(len(runs["p"]), {k: leb128(v) for k, v in runs.items()}),
    )


def decode_strokes(runs: Encoded, ps: Encoded) -> tuple[dict[int, int], list[int]]:
    cols = {k: unleb128(v) for k, v in runs.columns.items()}
    groups, prev_end, chained, last = {}, URO[0], 0, []
    for i in range(runs.count):
        first = prev_end + cols["gap"][i]
        if cols["group"][i]:
            chained += unzigzag(cols["group"][i] - 1)
        for k in range(cols["length"][i] + 1):
            if cols["group"][i]:
                groups[first + k] = chained
            else:
                last.append(first + k)
        prev_end = first + cols["length"][i] + 1
    top = max(groups.values(), default=0) + 1
    groups.update({cp: top for cp in last})
    out, last = [], 0
    for d, n in zip(unleb128(ps.columns["p"]), unleb128(ps.columns["length"])):
        out += [last + d + k for k in range(n + 1)]
        last = out[-1]
    return groups, out


def encode_compositions(comps: dict[tuple[int, int], int]) -> Encoded:
    """Records in composed order: gap = cp - previous cp, base = zigzag(base - previous base), mark = mark - U+0300."""
    cols = {"gap": [], "base": [], "mark": []}
    prev_cp, prev_base = 0, 0
    for cp, (base, mark) in sorted((cp, pair) for pair, cp in comps.items()):
        cols["gap"].append(cp - prev_cp)
        cols["base"].append(zigzag(base - prev_base))
        cols["mark"].append(mark - COMBINING_MARKS[0])
        prev_cp, prev_base = cp, base
    return Encoded(len(comps), {k: leb128(v) for k, v in cols.items()})


def decode_compositions(e: Encoded) -> dict[tuple[int, int], int]:
    cols = {k: unleb128(v) for k, v in e.columns.items()}
    out, cp, base = {}, 0, 0
    for i in range(e.count):
        cp += cols["gap"][i]
        base += unzigzag(cols["base"][i])
        out[(base, COMBINING_MARKS[0] + cols["mark"][i])] = cp
    return out


def encode_han_rank(rank: dict[int, int]) -> Encoded:
    cps = sorted(rank)
    cols = {"gap": [], "length": [], "rank": [], "step": []}
    prev_end, prev_rank, i = URO[0], 0, 0
    while i < len(cps):
        j, step = i + 1, None
        while j < len(cps) and cps[j] == cps[j - 1] + 1:
            st = rank[cps[j]] - rank[cps[j - 1]]
            if step is None and st in (0, 1):
                step = st
            elif st != step:
                break
            j += 1
        step = step or 0
        cols["gap"].append(cps[i] - prev_end)
        cols["length"].append(j - i - 1)
        cols["rank"].append(zigzag(rank[cps[i]] - prev_rank))
        cols["step"].append(step)
        prev_end, prev_rank, i = cps[j - 1] + 1, rank[cps[i]] + (j - i - 1) * step, j
    return Encoded(len(cols["gap"]), {k: leb128(v) for k, v in cols.items()})


def decode_han_rank(e: Encoded) -> dict[int, int]:
    cols = {k: unleb128(v) for k, v in e.columns.items()}
    out, prev_end, rank = {}, URO[0], 0
    for i in range(e.count):
        first = prev_end + cols["gap"][i]
        rank += unzigzag(cols["rank"][i])
        n, step = cols["length"][i] + 1, cols["step"][i]
        for k in range(n):
            out[first + k] = rank + k * step
        prev_end, rank = first + n, rank + (n - 1) * step
    return out


@dataclass
class EncodedTables:
    bases: dict[str, dict[str, Encoded]]
    han_rank: Encoded
    locales: dict[str, dict[str, Encoded]]
    compositions: Encoded


def encode(t: Tables) -> EncodedTables:
    bases = {}
    for name, b in t.bases.items():
        runs, tert = encode_runs(b.elements)
        bases[name] = {
            "elements": runs,
            "tertiary": tert,
            "expansions": encode_expansions(b.expansions),
            "repeat_marks": encode_pairs(sorted(b.repeat_marks.items()), *REPEAT_MARKS),
        }
    locales = {}
    for tag, loc in t.locales.items():
        runs, tert = encode_runs(loc.overrides)
        e = {
            "remap": encode_remap(loc.remap),
            "overrides": runs,
            "override_tertiary": tert,
            "expansions": encode_expansions(loc.expansions),
        }
        h = loc.han
        if h.kind == HAN_CODEPAGE:
            e["han_list"] = encode_step_runs([(code_slot(code, h.first_row), p) for code, p in h.listed])
        elif h.kind == HAN_STROKES:
            e["stroke_runs"], e["stroke_p"] = encode_strokes(h)
        elif h.kind == HAN_READINGS:
            e["reading_runs"] = encode_pairs(h.readings, *READING_RUNS)
            e["unlisted"], e["unlisted_cp"] = encode_unlisted(h.unlisted)
        locales[tag] = e
    return EncodedTables(bases, encode_han_rank(t.han_rank), locales, encode_compositions(t.compositions))


def decode(e: EncodedTables, t: Tables) -> Tables:
    """Rebuilds the tables from the encoded columns; scalar fields come from `t`."""
    bases = {}
    for name, b in t.bases.items():
        eb = e.bases[name]
        bases[name] = BaseTable(
            name,
            decode_runs(eb["elements"], eb["tertiary"]),
            decode_expansions(eb["expansions"]),
            b.han_rest_p,
            b.out_p,
            dict(decode_pairs(eb["repeat_marks"], *REPEAT_MARKS)),
        )
    locales = {}
    for tag, loc in t.locales.items():
        el = e.locales[tag]
        h = HanRule(loc.han.kind, loc.han.codepage, loc.han.first_row, loc.han.last_row)
        if h.kind == HAN_CODEPAGE:
            h.listed = [(slot_code(slot, h.first_row), p) for slot, p in decode_step_runs(el["han_list"])]
        elif h.kind == HAN_STROKES:
            h.groups, h.group_p = decode_strokes(el["stroke_runs"], el["stroke_p"])
        elif h.kind == HAN_READINGS:
            h.readings = decode_pairs(el["reading_runs"], *READING_RUNS)
            h.unlisted = decode_unlisted(el["unlisted"], el["unlisted_cp"])
        locales[tag] = LocaleTable(
            tag,
            loc.base,
            loc.rules,
            decode_remap(el["remap"]),
            decode_runs(el["overrides"], el["override_tertiary"]),
            decode_expansions(el["expansions"]),
            h,
        )
    return Tables(bases, decode_han_rank(e.han_rank), locales, decode_compositions(e.compositions))


# ---------------------------------------------------------------- reference model


def kana_vowel(cp: int) -> int | None:
    """Vowel kana (same script) a following prolonged-sound mark stands for, if any."""
    name = unicodedata.name(chr(cp), "")
    for prefix, vowels in (
        ("HIRAGANA LETTER ", "あいうえお"),
        ("KATAKANA LETTER ", "アイウエオ"),
        ("HALFWIDTH KATAKANA LETTER ", "アイウエオ"),
    ):
        if name.startswith(prefix):
            last = name.rsplit(" ", 1)[1][-1]
            return ord(vowels["AIUEO".index(last)]) if last in "AIUEO" else None
    return None


def is_hiragana(cp: int) -> bool:
    return unicodedata.name(chr(cp), "").startswith("HIRAGANA LETTER ")


def is_katakana(cp: int) -> bool:
    name = unicodedata.name(chr(cp), "")
    return name.startswith("KATAKANA LETTER ") or name.startswith("HALFWIDTH KATAKANA LETTER ")


def voiced(cp: int, mark: int = 0x3099) -> int | None:
    """Fullwidth kana `cp` (halfwidth widened first) with a voiced or semi-voiced mark composed."""
    base = unicodedata.normalize("NFKC", chr(cp))
    base = unicodedata.normalize("NFD", base)[0]
    composed = unicodedata.normalize("NFC", base + chr(mark))
    return ord(composed) if len(composed) == 1 else None


PROLONGED = (0x30FC, 0xFF70)
HIRA_REPEAT, HIRA_REPEAT_VOICED = 0x309D, 0x309E
KATA_REPEAT, KATA_REPEAT_VOICED = 0x30FD, 0x30FE
HALFWIDTH_KANA = (0xFF66, 0xFF9D)
HALFWIDTH_VOICING = {0xFF9E: 0x3099, 0xFF9F: 0x309A}


@dataclass
class Token:
    cp: int
    pos: int  # code point index in the original string
    c_from: int | None = None  # halfwidth voicing: the halfwidth base whose C the composed element takes
    repeat: bool = False  # zh repeat mark: a secondary marker follows the element


class Collator:
    """Reference key builder for one locale; strengths are COMPARE, DYNAMIC and EXACT."""

    def __init__(self, tables: Tables, tag: str):
        self.loc = tables.locales[tag]
        self.base = tables.bases[self.loc.base]
        self.han_rank = tables.han_rank
        self.compositions = tables.compositions
        h = self.loc.han
        self.han_p = {}
        self.readings = {}
        if h.kind == HAN_CODEPAGE:
            code_cp = dict(codepage_cells(h.codepage, h.first_row, h.last_row))
            self.han_p = {code_cp[code]: p for code, p in h.listed}
        elif h.kind == HAN_STROKES:
            order = sorted(h.groups, key=lambda cp: (h.groups[cp], cp))
            self.han_p = dict(zip(order, h.group_p))
        elif h.kind == HAN_READINGS:
            cells = [cp for _, cp in codepage_cells(h.codepage, h.first_row, h.last_row)]
            counts: Counter[int] = Counter()
            i = 0
            for syl, n in h.readings:
                for cp in cells[i : i + n]:
                    if syl:
                        counts[syl] += 1
                        self.readings[cp] = (syl, counts[syl])
                i += n
            for cp, (syl, rank) in h.unlisted.items():
                self.readings[cp] = (syl, counts[syl] + rank)
        self.cache: dict[tuple[int, bool], list[Weight]] = {}

    def remap(self, w: Weight) -> Weight:
        if w[0] != BASE:
            return w
        return (BASE, remap_of(self.loc.remap, w[1]), *w[2:])

    def single(self, cp: int) -> Weight | None:
        """Weight of a code point that is one element, before expansions."""
        if cp in self.loc.overrides:
            return self.loc.overrides[cp]
        if cp in self.base.elements:
            return self.remap(self.base.elements[cp])
        if in_range(cp, URO):
            if cp in self.han_p:
                return (BASE, self.han_p[cp], 0, 0, 0)
            if cp not in self.readings:
                return self.remap((BASE, self.base.han_rest_p + self.han_rank[cp], 0, 0, 0))
        return None

    def elements(self, cp: int, expand: bool) -> list[Weight]:
        """Elements of one code point; a (MARK, level) entry is an expansion's level marker.

        FULL and MID come back as a single pseudo-element so the caller can see a trailing MID.
        """
        key = (cp, expand)
        if key in self.cache:
            return self.cache[key]
        exp = self.loc.expansions.get(cp)
        if exp is None and cp not in self.loc.overrides:
            exp = self.base.expansions.get(cp)
        if exp is not None:
            targets, level = exp
            if not expand and len(targets) > 1 and level == 0:
                level = 3  # no expansion: it still differs from its targets at C
            out = []
            for t in targets:
                out += self.elements(t, expand)
            if level:
                out.append((MARK, level))
        elif (w := self.single(cp)) is not None:
            out = [w]
        elif cp in self.readings:
            syl, rank = self.readings[cp]
            out = list(self.elements(syl, expand))
            last = out[-1]
            out[-1] = (last[0], last[1], last[2] + rank, last[3], last[4])
        elif in_range(cp, SYLLABLES) and self.loc.rules & RULE_HANGUL_JAMO:
            out = [w for j in hangul_jamo(cp) for w in self.elements(j, expand)]
        elif self.loc.rules & RULE_IGNORE_UNLISTED:
            out = [(FULL, 0, 0, 0, 0)]
        else:
            out = [self.remap((BASE, self.base.out_p + out_of_range_index(cp), 0, 0, 0))]
        self.cache[key] = out
        return out

    def halfwidth_composed(self, voiced_weight: Weight, half: int) -> Weight:
        """Halfwidth voicing: the voiced kana's P, S, D; C moved by the halfwidth base's C offset from its fullwidth form."""
        full = ord(unicodedata.normalize("NFKC", chr(half)))
        (h,), (f,) = self.elements(half, True), self.elements(full, True)
        return voiced_weight[:4] + (voiced_weight[4] + h[4] - f[4],)

    def tokens(self, s: str, strength: str) -> list[Token]:
        toks = [Token(ord(ch), i) for i, ch in enumerate(s)]
        # Halfwidth kana + halfwidth (semi-)voiced mark compose into one element.
        if strength != EXACT:
            out, i = [], 0
            while i < len(toks):
                t = toks[i]
                if (
                    i + 1 < len(toks)
                    and in_range(t.cp, HALFWIDTH_KANA)
                    and toks[i + 1].cp in HALFWIDTH_VOICING
                    and (v := voiced(t.cp, HALFWIDTH_VOICING[toks[i + 1].cp])) is not None
                ):
                    out.append(Token(v, t.pos, t.cp))
                    i += 2
                    continue
                out.append(t)
                i += 1
            toks = out
            # A base and a following combining mark compose, repeatedly.
            out = []
            for t in toks:
                c = self.compositions.get((out[-1].cp, t.cp)) if out and out[-1].c_from is None else None
                if c is not None:
                    out[-1] = Token(c, out[-1].pos)
                    continue
                out.append(t)
            toks = out
        # ja: prolonged-sound and iteration marks take the preceding kana.
        if self.loc.rules & RULE_KANA_MARKS:
            for i in range(1, len(toks)):
                prev, t = toks[i - 1].cp, toks[i]
                sub = None
                if t.cp in PROLONGED:
                    sub = kana_vowel(prev)
                elif t.cp in (HIRA_REPEAT, HIRA_REPEAT_VOICED) and is_hiragana(prev):
                    sub = prev if t.cp == HIRA_REPEAT else voiced(prev)
                elif t.cp in (KATA_REPEAT, KATA_REPEAT_VOICED) and is_katakana(prev):
                    sub = prev if t.cp == KATA_REPEAT else voiced(prev)
                if sub is not None:
                    toks[i] = Token(sub, t.pos)
        # zh repeat marks take the preceding character, plus a secondary marker.
        if self.loc.rules & RULE_REPEAT_MARKS:
            for i in range(1, len(toks)):
                kind = self.base.repeat_marks.get(toks[i].cp)
                if kind is None:
                    continue
                prev = toks[i - 1].cp
                sub = prev
                if kind == REPEAT_VOWEL:
                    sub = kana_vowel(prev) or prev
                elif kind == REPEAT_VOICED:
                    sub = voiced(prev) or prev
                toks[i] = Token(sub, toks[i].pos, repeat=True)
        # A Thai prefix vowel sorts after the consonant that follows it.
        if self.loc.rules & RULE_THAI_REORDER:
            i = 0
            while i + 1 < len(toks):
                if 0x0E40 <= toks[i].cp <= 0x0E44 and 0x0E01 <= toks[i + 1].cp <= 0x0E2E:
                    toks[i], toks[i + 1] = toks[i + 1], toks[i]
                    i += 2
                else:
                    i += 1
        return toks

    def key(self, s: str, strength: str) -> bytes:
        """Collation key: parts joined by 0x01, each weight as two bytes (v // 254 + 2, v % 254 + 2).

        1. P of every base element; a word symbol after the last P/S/marker element adds P 0.
        2. S of base and sec elements; an S-level marker (expansion level S, zh repeat) adds 0.
        3. D of base and sec elements; a D-level marker adds 0.
        4. C + 1 of base and sec elements; a C-level marker adds 0. Omitted for DYNAMIC.
        5. Per word symbol: original code point index, P, S, D, and C unless DYNAMIC.
        6. 0 when the last code point is a MID character, else empty.
        EXACT turns off halfwidth voicing and expansion, keeping an expansion to two or more characters
        at level `=` apart from its targets with a C-level marker.
        """
        expand = strength != EXACT
        toks = self.tokens(s, strength)
        per_token = []
        for t in toks:
            els = self.elements(t.cp, expand)
            if t.c_from is not None:
                els = [self.halfwidth_composed(els[0], t.c_from)]
            if t.repeat:
                els = els + [(MARK, 1)]
            per_token.append(els)
        # Word symbols after the last element with a P or S also weigh as the lowest primary.
        content = [i for i, els in enumerate(per_token) if any(e[0] in (BASE, SEC, MARK) for e in els)]
        tail_from = content[-1] + 1 if content else len(per_token)
        parts: list[list[int]] = [[], [], [], [], [], []]
        for i, (t, els) in enumerate(zip(toks, per_token)):
            for e in els:
                cls = e[0]
                if cls == BASE:
                    parts[0].append(e[1])
                if cls in (BASE, SEC):
                    parts[1].append(e[2])
                    parts[2].append(e[3])
                    parts[3].append(e[4] + 1)
                elif cls == WORD:
                    if i >= tail_from:
                        parts[0].append(0)
                    parts[4] += [t.pos, e[1], e[2], e[3]] + ([e[4]] if strength != DYNAMIC else [])
                elif cls == MARK:
                    parts[e[1]].append(0)
        if per_token and len(per_token[-1]) == 1 and per_token[-1][0][0] == MID:
            parts[5].append(0)
        if strength == DYNAMIC:
            del parts[3]
        return b"\x01".join(b"".join(weight_bytes(v) for v in part) for part in parts)


MARK = 5  # pseudo-class of an expansion's level marker


def weight_bytes(v: int) -> bytes:
    if not 0 <= v <= MAX_WEIGHT:
        fail(f"weight {v} out of range")
    return bytes((v // 254 + 2, v % 254 + 2))


def hangul_jamo(cp: int) -> list[int]:
    s = cp - SYLLABLES[0]
    out = [0x1100 + s // 588, 0x1161 + (s % 588) // 28]
    if s % 28:
        out.append(0x11A7 + s % 28)
    return out


def sign(a: bytes, b: bytes) -> str:
    return "-" if a < b else "0" if a == b else "+"


# ---------------------------------------------------------------- validation


def matrix_misses(tables: Tables, tag: str) -> list[tuple[str, str, str, str, str]]:
    """(a, b, strength, expected, model) for every probe-matrix cell the model gets wrong."""
    strings = [dec(r[1]) for r in read_tsv(FIXTURES / "probe_strings.tsv")]
    rows = read_tsv(FIXTURES / f"matrix_{tag}.tsv")
    col = Collator(tables, tag)
    out = []
    for strength, k in ((COMPARE, 1), (DYNAMIC, 2)):
        keys = [col.key(s, strength) for s in strings]
        for i, a in enumerate(strings):
            for j, b in enumerate(strings):
                want, got = rows[i][k][j], sign(keys[i], keys[j])
                if want != got:
                    out.append((enc(a), enc(b), strength, want, got))
    return out


def order_check(tables: Tables, b: Build, tag: str) -> tuple[int, int]:
    """(approximated code points, order failures) for the single-character round trip.

    The measured order restricted to the exactly reproduced code points must come back
    from the single-character keys: each pair of neighbours in the same class compares
    as measured (equal for `=`, ascending otherwise; C neighbours are equal under SORT).
    """
    m, space = b.measured[tag], b.spaces[tag]
    excluded = {cp for cp in m.order if not is_exact(cp)} | space.approximated
    col = Collator(tables, tag)
    fails = 0
    for strength in (COMPARE, DYNAMIC):
        prev: dict[int, tuple[Weight, bytes]] = {}
        for cp in m.order:
            if cp in excluded or cp in m.expanding:
                continue
            w = m.weights[cp]
            k = col.key(chr(cp), strength)
            if w[0] in prev:
                pw, pk = prev[w[0]]
                level = first_diff(pw[1:], w[1:])
                want = "0" if level == 4 or (level == 3 and strength == DYNAMIC) else "-"
                if w[0] in (FULL, MID):
                    want = "0"
                if sign(pk, k) != want and not (in_range(cp, SYLLABLES) or cp in space.computed):
                    fails += 1
            prev[w[0]] = (w, k)
    return len(excluded), fails


# ---------------------------------------------------------------- C++ rendering

HEADER = """\
// GENERATED by tools/collation/generate_tables.py; do not edit.
//
// Text collation tables from Mac Excel. Each locale is a base table
// (W for the Windows-style locales, Z for zh_CN/zh_TW) plus a delta. A character's
// weights come from, in order: the locale expansions, the locale overrides, the
// base expansions, the base elements, the locale Han rule (unified ideographs),
// the W radical-stroke rank shifted through the remap, Hangul jamo decomposition,
// and last the out-of-range rule. Every table is a set of LEB128 varint columns
// with one value per record; "zigzag" columns hold signed deltas.

#ifndef FORMULON_EVAL_COLLATION_TABLE_H_
#define FORMULON_EVAL_COLLATION_TABLE_H_

#include <cstddef>
#include <cstdint>

#include "excel_locale.h"

namespace formulon {
namespace eval {
namespace collation_detail {

/// Element class: kFull ignored everywhere, kMid ignored except as a trailing
/// mark, kWord zh word-sort symbol (word part only), kSec secondary weight only.
enum class CollationClass : std::uint8_t { kFull = 0, kMid = 1, kWord = 2, kSec = 3, kBase = 4 };

/// Expansion level: how an expanding character differs from its targets.
enum class CollationLevel : std::uint8_t { kEqual = 0, kSecondary = 1, kDynamic = 2, kCompareOnly = 3 };

enum class CollationHanKind : std::uint8_t { kBase = 0, kCodepage = 1, kStrokes = 2, kReadings = 3 };

/// Locale rule bits.
constexpr std::uint8_t kRuleKanaMarks = 1;       // ja prolonged-sound and iteration marks
constexpr std::uint8_t kRuleThaiReorder = 2;     // Thai prefix vowel after its consonant
constexpr std::uint8_t kRuleHangulJamo = 4;      // Hangul syllables decompose into jamo
constexpr std::uint8_t kRuleIgnoreUnlisted = 8;  // out-of-range code points are ignorable
constexpr std::uint8_t kRuleRepeatMarks = 16;    // zh repeat marks take the preceding character

/// What a zh repeat mark stands for after another character: the character itself,
/// its voiced form, or (prolonged-sound mark) the vowel of a preceding kana.
enum class CollationRepeatKind : std::uint8_t { kSame = 0, kVoiced = 1, kVowel = 2 };

/// Largest weight a key can hold (two base-254 digits).
constexpr std::uint16_t kMaxCollationWeight = @MAX_WEIGHT@;

/// One LEB128 varint column.
struct CollationColumn {
  const std::uint8_t* data;
  std::uint32_t size;
};

/// Element runs in code-point order. Per run: gap = first cp - previous run end
/// (initially 0); length = cp count - 1; p = zigzag(first P - previous run's last
/// P) (initially 0); s; cls_step = class | step << 3, where P advances by step
/// (0 or 1) per code point. D and C are 0 unless listed in the paired sparse table.
struct CollationRunTable {
  std::uint32_t count;
  CollationColumn gap, length, p, s, cls_step;
};

/// Sparse tertiary weights: gap = cp - previous cp (initially 0); d; c.
struct CollationTertiaryTable {
  std::uint32_t count;
  CollationColumn gap, d, c;
};

/// Expansions: gap = cp - previous cp (initially 0); head = level | (target count - 1) << 2;
/// targets holds zigzag(target - previous target) per target, the first relative to the
/// previous record's first target (initially 0).
struct CollationExpansionTable {
  std::uint32_t count;
  CollationColumn gap, head, targets;
};

/// P remap segments: base_p = segment start - previous start (initially 0);
/// offset = zigzag(d - previous d) with d = locale P - base P at the segment start
/// (initially 0). A base P maps through the last segment starting at or below it.
struct CollationRemapTable {
  std::uint32_t count;
  CollationColumn base_p, offset;
};

/// Code-page ideographs in P order, as runs where both the cell and P advance by 1;
/// cell = (lead - first_row) * 94 + trail - 1 from lookup_unicode_to_dbcs. Per run:
/// slot = zigzag(first cell - previous run's last cell - 1) (initially -1),
/// p = first P - previous run's last P (initially 0), length = cells - 1.
struct CollationCodepageList {
  std::uint32_t count;
  CollationColumn slot, p, length;
};

/// zh_TW stroke groups: runs over unified ideographs, gap = first cp - previous run
/// end (initially U+4E00), length = count - 1, group = 0 for the last group (after every
/// other) or zigzag(group - previous nonzero run's group) + 1 (initially 0). The listed
/// ideographs sort by (group, cp); their locale P in that order come as runs advancing
/// by 1: p = first P - previous run's last P (initially 0), p_length = count - 1.
struct CollationStrokeTable {
  std::uint32_t run_count;
  CollationColumn gap, length, group;
  std::uint32_t p_count;
  CollationColumn p, p_length;
};

/// ko readings. Runs over the code page's unified-ideograph cells in code order:
/// syllable = zigzag(reading syllable - previous), 0 for no reading; length = cells - 1.
/// Ideographs outside the code page, grouped by reading in syllable order: per group
/// syllable = syllable - previous (initially 0) and count = ideographs - 1; then cp
/// holds every group's ideographs in rank order as zigzag(cp - previous in the group,
/// the first relative to U+4E00). A code-page cell ranks by its position among the
/// reading's cells (from 1); the i-th ideograph of a group (from 1) ranks cells + i.
struct CollationReadingTable {
  std::uint32_t run_count;
  CollationColumn syllable, length;
  std::uint32_t group_count;
  CollationColumn group_syllable, group_size;
  std::uint32_t unlisted_count;
  CollationColumn unlisted_cp;
};

struct CollationHanRule {
  CollationHanKind kind;
  DbcsCodepage codepage;
  std::uint8_t first_row;
  std::uint8_t last_row;
  CollationCodepageList list;
  CollationStrokeTable strokes;
  CollationReadingTable readings;
};

/// Z repeat marks: cp = cp - previous (initially 0); kind = CollationRepeatKind.
struct CollationRepeatTable {
  std::uint32_t count;
  CollationColumn cp, kind;
};

struct CollationBaseTable {
  CollationRunTable elements;
  CollationTertiaryTable tertiary;
  CollationExpansionTable expansions;
  CollationRepeatTable repeat_marks;
  /// Base P of W radical-stroke rank 0.
  std::uint16_t han_rest_p;
  /// W only: base P of the first out-of-range code point (they follow in code-point order).
  std::uint16_t out_of_range_p;
};

struct CollationLocaleTable {
  const CollationBaseTable* base;
  std::uint8_t rules;
  CollationRemapTable remap;
  CollationRunTable overrides;
  CollationTertiaryTable override_tertiary;
  CollationExpansionTable expansions;
  CollationHanRule han;
};

struct CollationRange {
  std::uint16_t first;
  std::uint16_t last;
};

/// Code points reproduced exactly; the out-of-range rank of a code point is
/// cp minus the number of in-range code points below it.
extern const CollationRange kCollationExactRanges[@RANGE_COUNT@];

/// W radical-stroke rank of each unified ideograph: runs with gap = first cp - previous
/// run end (initially U+4E00), length = count - 1, rank = zigzag(first rank - previous
/// run's last rank), step = rank advance per code point (0 or 1).
struct CollationRankTable {
  std::uint32_t count;
  CollationColumn gap, length, rank, step;
};
extern const CollationRankTable kCollationHanRank;

/// Canonical compositions of a base and a combining mark (U+0300..U+036F), in
/// composed order: gap = cp - previous cp (initially 0); base = zigzag(base -
/// previous base) (initially 0); mark = mark - U+0300.
struct CollationCompositionTable {
  std::uint32_t count;
  CollationColumn gap, base, mark;
};
extern const CollationCompositionTable kCollationCompositions;

/// Indexed by ExcelLocale.
extern const CollationLocaleTable kCollationLocales[@LOCALE_COUNT@];

}  // namespace collation_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COLLATION_TABLE_H_
"""


def render_header() -> str:
    return (
        HEADER.replace("@MAX_WEIGHT@", str(MAX_WEIGHT))
        .replace("@RANGE_COUNT@", str(len(EXACT_RANGES)))
        .replace("@LOCALE_COUNT@", str(len(LOCALES)))
    )


class Emitter:
    """Collects byte arrays, sharing identical ones, and renders struct initializers."""

    def __init__(self) -> None:
        self.arrays: dict[bytes, str] = {}
        self.order: list[tuple[str, bytes]] = []

    def column(self, data: bytes, name: str) -> str:
        if not data:
            return "{nullptr, 0}"
        if data not in self.arrays:
            self.arrays[data] = name
            self.order.append((name, data))
        return f"{{{self.arrays[data]}, {len(data)}}}"

    def table(self, e: Encoded | None, prefix: str, names: tuple[str, ...]) -> str:
        if e is None:
            return "{" + ", ".join(["0"] + ["{nullptr, 0}"] * len(names)) + "}"
        cols = [self.column(e.columns[n], f"{prefix}{camel(n)}") for n in names]
        return "{" + ", ".join([str(e.count)] + cols) + "}"

    def payload_bytes(self) -> bytes:
        return b"".join(data for _, data in self.order)

    def render_arrays(self) -> str:
        out = []
        for name, data in self.order:
            out.append(f"constexpr std::uint8_t {name}[] = {{\n")
            for start in range(0, len(data), 19):
                out.append("    " + ", ".join(f"0x{b:02X}" for b in data[start : start + 19]) + ",\n")
            out.append("};\n")
        return "".join(out)


def camel(name: str) -> str:
    return "".join(part.capitalize() for part in name.split("_"))


RUN_COLUMNS = ("gap", "length", "p", "s", "cls_step")
TERT_COLUMNS = ("gap", "d", "c")
EXP_COLUMNS = ("gap", "head", "targets")


def render_source(t: Tables, e: EncodedTables) -> tuple[str, bytes]:
    em = Emitter()
    base_inits = {}
    for name, b in t.bases.items():
        eb = e.bases[name]
        p = f"k{name}"
        base_inits[name] = (
            "{"
            + ", ".join(
                [
                    em.table(eb["elements"], f"{p}Elements", RUN_COLUMNS),
                    em.table(eb["tertiary"], f"{p}Tertiary", TERT_COLUMNS),
                    em.table(eb["expansions"], f"{p}Expansions", EXP_COLUMNS),
                    em.table(eb["repeat_marks"], f"{p}RepeatMarks", ("cp", "kind")),
                    str(b.han_rest_p),
                    str(b.out_p),
                ]
            )
            + "}"
        )
    rank_init = em.table(e.han_rank, "kHanRank", ("gap", "length", "rank", "step"))
    composition_init = em.table(e.compositions, "kCompositions", ("gap", "base", "mark"))
    locale_inits = []
    for tag in LOCALES:
        loc, el = t.locales[tag], e.locales[tag]
        p = "k" + camel(tag.lower())
        h = loc.han
        han = (
            "{"
            + ", ".join(
                [
                    f"CollationHanKind::{HAN_KIND_NAMES[h.kind]}",
                    f"DbcsCodepage::{CODEPAGE_ENUM[h.codepage]}",
                    str(h.first_row),
                    str(h.last_row),
                    em.table(el.get("han_list"), f"{p}HanList", ("slot", "p", "length")),
                    "{"
                    + ", ".join(
                        [
                            em.table(el.get("stroke_runs"), f"{p}Stroke", ("gap", "length", "group"))[1:-1],
                            em.table(el.get("stroke_p"), f"{p}StrokeP", ("p", "length"))[1:-1],
                        ]
                    )
                    + "}",
                    "{"
                    + ", ".join(
                        [
                            em.table(el.get("reading_runs"), f"{p}Reading", ("syllable", "length"))[1:-1],
                            em.table(el.get("unlisted"), f"{p}Unlisted", ("syllable", "count"))[1:-1],
                            em.table(el.get("unlisted_cp"), f"{p}Unlisted", ("cp",))[1:-1],
                        ]
                    )
                    + "}",
                ]
            )
            + "}"
        )
        locale_inits.append(
            f"    // {tag}\n    {{"
            + ", ".join(
                [
                    f"&kCollationBase{loc.base}",
                    str(loc.rules),
                    em.table(el["remap"], f"{p}Remap", ("base_p", "offset")),
                    em.table(el["overrides"], f"{p}Overrides", RUN_COLUMNS),
                    em.table(el["override_tertiary"], f"{p}OverrideTertiary", TERT_COLUMNS),
                    em.table(el["expansions"], f"{p}Expansions", EXP_COLUMNS),
                    han,
                ]
            )
            + "},\n"
        )
    payload = em.payload_bytes()
    ranges = ", ".join(f"{{0x{lo:04X}, 0x{hi:04X}}}" for lo, hi in EXACT_RANGES)
    src = (
        "// GENERATED by tools/collation/generate_tables.py from tools/collation/data; do not edit.\n"
        f"// {len(payload)} encoded bytes.\n"
        f"// @size-budget: {TABLE_BUDGET_KB} KB\n"
        "\n"
        "// clang-format off\n"
        "#include <cstdint>\n"
        "\n"
        '#include "eval/collation_table.h"\n'
        "\n"
        "namespace formulon {\n"
        "namespace eval {\n"
        "namespace collation_detail {\n"
        "namespace {\n"
        "\n" + em.render_arrays() + "\n"
        "}  // namespace\n"
        "\n"
        f"const CollationRange kCollationExactRanges[{len(EXACT_RANGES)}] = {{{ranges}}};\n"
        "\n"
        f"const CollationRankTable kCollationHanRank = {rank_init};\n"
        "\n"
        f"const CollationCompositionTable kCollationCompositions = {composition_init};\n"
        "\n"
        "namespace {\n"
        "\n"
        + "".join(f"const CollationBaseTable kCollationBase{name} = {init};\n" for name, init in base_inits.items())
        + "\n"
        "}  // namespace\n"
        "\n"
        f"const CollationLocaleTable kCollationLocales[{len(LOCALES)}] = {{\n" + "".join(locale_inits) + "};\n"
        "\n"
        "}  // namespace collation_detail\n"
        "}  // namespace eval\n"
        "}  // namespace formulon\n"
        "// clang-format on\n"
    )
    return src, payload


def render_excused(misses: list[tuple[str, str, str, str, str]]) -> str:
    lines = ["a\tb\tstrength\texpected\tmodel"] + ["\t".join(r) for r in misses]
    return "\n".join(lines) + "\n"


def brotli_size(data: bytes) -> int | None:
    if shutil.which("brotli") is None:
        return None
    out = subprocess.run(["brotli", "-c", "-q", "11"], input=data, capture_output=True, check=True).stdout
    return len(out)


# ---------------------------------------------------------------- driver


@dataclass
class Generated:
    files: dict[Path, str]
    payload: bytes
    report: list[str]
    out_of_range: dict[str, int]
    tables: Tables


def generate() -> Generated:
    b = build()
    encoded = encode(b.tables)
    tables = decode(encoded, b.tables)
    if not tables_equal(tables, b.tables):
        fail("encoded tables do not round-trip")
    source, payload = render_source(tables, encoded)
    files = {HEADER_PATH: render_header(), SOURCE_PATH: source}
    report, out_of_range = [], {}
    for tag in LOCALES:
        misses = matrix_misses(tables, tag)
        files[FIXTURES / f"excused_{tag}.tsv"] = render_excused(misses)
        count, fails = order_check(tables, b, tag)
        out_of_range[tag] = count
        loc = tables.locales[tag]
        split = Counter(r[2] for r in misses)
        report.append(
            f"{tag}: out-of-range {count}, order failures {fails}, excused {len(misses)}"
            f" (compare {split[COMPARE]}, dynamic {split[DYNAMIC]}), remap {len(loc.remap)},"
            f" overrides {len(loc.overrides)}, expansions {len(loc.expansions)}"
        )
    report.append(f"zh_TW Unihan (kTotalStrokes, cp) misplaces {b.spaces['zh_TW'].unihan_misses} ideographs")
    return Generated(files, payload, report, out_of_range, tables)


def tables_equal(a: Tables, b: Tables) -> bool:
    if a.han_rank != b.han_rank or a.compositions != b.compositions:
        return False
    for name in a.bases:
        x, y = a.bases[name], b.bases[name]
        if (x.elements, x.expansions, x.repeat_marks) != (y.elements, y.expansions, y.repeat_marks):
            return False
    for tag in a.locales:
        x, y = a.locales[tag], b.locales[tag]
        if (x.remap, x.overrides, x.expansions) != (y.remap, y.overrides, y.expansions):
            return False
        hx, hy = x.han, y.han
        if (hx.listed, hx.groups, hx.group_p, hx.readings, hx.unlisted) != (
            hy.listed,
            hy.groups,
            hy.group_p,
            hy.readings,
            hy.unlisted,
        ):
            return False
    return True


def check(paths: dict[Path, Path] | None = None, generated: Generated | None = None) -> int:
    """Compares tracked files with the generator output; 1 on drift, 2 when unreadable.

    `paths` maps a canonical output path to the file actually compared, so a test can
    point the check at a temporary fixture.
    """
    paths = paths or {}
    expected = (generated or generate()).files
    drifted = []
    for canonical, text in expected.items():
        try:
            actual = paths.get(canonical, canonical).read_text(encoding="utf-8")
        except OSError as exc:
            print(f"cannot check generated tables: {exc}", file=sys.stderr)
            return 2
        if actual != text:
            drifted.append(canonical)
    for path in drifted:
        print(f"generated table drift detected: {path}", file=sys.stderr)
    if drifted:
        print(REGENERATION_GUIDANCE, file=sys.stderr)
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n", 1)[0])
    parser.add_argument("--check", action="store_true", help="compare the outputs with the tracked files")
    args = parser.parse_args(argv)
    if args.check:
        return check()
    g = generate()
    for path, text in g.files.items():
        path.write_text(text, encoding="utf-8")
    for line in g.report:
        print(line)
    size = brotli_size(g.payload)
    if size is None:
        print(f"table payload {len(g.payload)} bytes (brotli not on PATH)")
        return 0
    print(
        f"table payload {len(g.payload)} bytes, brotli {size} bytes (aim {SIZE_AIM_BYTES}, feature budget {SIZE_BUDGET_BYTES})"
    )
    if size > SIZE_BUDGET_BYTES:
        print("table payload alone exceeds the feature budget", file=sys.stderr)
        return 1
    if size > SIZE_AIM_BYTES:
        print("warning: table payload leaves less than the planned room for code", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
