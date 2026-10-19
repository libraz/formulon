#!/usr/bin/env python3
"""Capture Mac Excel's text collation into the TSV tables under data/.

Every observation is an Excel formula result; no weight is written by hand.
The BMP order comes from sorting `UNICHAR(SEQUENCE(...))` once in ascending
and once in descending code-point input: SORT is stable, so a pair whose
order flips between the two runs is equal under SORT's (dynamic) strength.

Files (strings are space-separated `U+XXXX`; the empty string is an empty field):
  order_<loc>.tsv            cp, cls, rel -- BMP in collation order. cls is one of
                             full (ignored everywhere), mid (ignored except as a
                             trailing mark), word (zh word-sort symbol), sec
                             (secondary weight only), base (has a primary).
                             rel compares with the previous row of the same cls:
                             P primary, S secondary, D dynamic-visible tertiary,
                             C tertiary visible only to `=`, `=` equal.
  expansion_<loc>.tsv        cp, expands_to, level -- compatibility/ligature
                             candidates whose difference from their expansion is
                             below primary (level S, D, C or =).
  fixtures/probe_strings.tsv id, string
  fixtures/matrix_<loc>.tsv  a, compare, dynamic -- one row per probe string; the
                             columns hold one sign character (-, 0, +) per probe
                             string b, for `=`/`<` and for SORT respectively.
  fixtures/eq_<loc>.tsv      a, b, then one 0/1/n column per function probe.
  fixtures/tail_<loc>.tsv    form, c, c2, x, y, sign -- trailing-ignorable probes.
  meta.tsv                   locale, Excel version and a locale fingerprint.

Usage:
  capture.py capture-all             switch Excel through every locale (macOS)
  capture.py capture --locale ja_JP  capture with Excel already in that locale
  capture.py --validate              offline consistency check of data/
"""

from __future__ import annotations

import argparse
import itertools
import subprocess
import sys
import time
import unicodedata
from pathlib import Path

DATA = Path(__file__).resolve().parent / "data"
FIX = DATA / "fixtures"

# (tag, AppleLanguages, AppleLocale); ja_JP is the host default and clears both.
LOCALES = [
    ("ja_JP", "", ""),
    ("en_US", "en-US", "en_US"),
    ("de_DE", "de-DE", "de_DE"),
    ("fr_FR", "fr-FR", "fr_FR"),
    ("es_ES", "es-ES", "es_ES"),
    ("es_MX", "es-MX", "es_MX"),
    ("pt_BR", "pt-BR", "pt_BR"),
    ("it_IT", "it-IT", "it_IT"),
    ("nl_NL", "nl-NL", "nl_NL"),
    ("ru_RU", "ru-RU", "ru_RU"),
    ("th_TH", "th-TH", "th_TH"),
    ("ko_KR", "ko-KR", "ko_KR"),
    ("zh_CN", "zh-Hans-CN", "zh_CN"),
    ("zh_TW", "zh-Hant-TW", "zh_TW"),
]
Z_LOCALES = ("zh_CN", "zh_TW")
# Locales that share one order; every other locale has its own.
SAME_ORDER = [("en_US", "de_DE", "fr_FR", "pt_BR", "it_IT", "nl_NL"), ("es_ES", "es_MX")]
CLASSES = ("full", "mid", "word", "sec", "base")
# Affixes that place a lone character in a context where its class can be compared.
AFFIX = {"base": ("", ""), "sec": ('"a"&', ""), "word": ('"a"&', '&"b"')}
ACUTE = "UNICHAR(225)"

BMP_ASC = "VSTACK(SEQUENCE(55295),SEQUENCE(8190,1,57344))"
BMP_DESC = "VSTACK(SEQUENCE(8190,1,65533,-1),SEQUENCE(55295,1,55295,-1))"


# ---------------------------------------------------------------- encoding


def enc(s: str) -> str:
    return " ".join(f"U+{ord(ch):04X}" for ch in s)


def dec(field: str) -> str:
    return "".join(chr(int(t[2:], 16)) for t in field.split())


def read_tsv(path: Path) -> list[list[str]]:
    lines = path.read_text(encoding="utf-8").split("\n")
    return [ln.split("\t") for ln in lines[1:] if ln]


def write_tsv(path: Path, header: list[str], rows) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    out = ["\t".join(header)] + ["\t".join(str(c) for c in r) for r in rows]
    path.write_text("\n".join(out) + "\n", encoding="utf-8")


# ---------------------------------------------------------------- probe strings

_BASE_STRINGS = [
    " ",
    "!",
    "#",
    "$",
    "%",
    "&",
    "'",
    "(",
    "*",
    "+",
    ",",
    "-",
    ".",
    "/",
    ":",
    ";",
    "<",
    "=",
    ">",
    "?",
    "@",
    "[",
    "\\",
    "]",
    "^",
    "_",
    "`",
    "{",
    "|",
    "}",
    "~",
    "0",
    "1",
    "9",
    "10",
    "2",
    "a",
    "A",
    "b",
    "B",
    "z",
    "Z",
    "ab",
    "a-b",
    "a b",
    "a'b",
    "aB",
    "Ab",
    "ae",
    "ss",
    "oe",
    "à",
    "á",
    "â",
    "ä",
    "Ä",
    "å",
    "æ",
    "Æ",
    "ç",
    "è",
    "é",
    "ê",
    "ë",
    "í",
    "ñ",
    "Ñ",
    "ò",
    "ó",
    "ö",
    "ø",
    "ù",
    "ü",
    "ý",
    "ÿ",
    "ß",
    "œ",
    "þ",
    "ð",
    "ł",
    "ş",
    "Ａ",
    "ａ",
    "１",
    "＃",
    "　",
    "あ",
    "ア",
    "ｱ",
    "ぁ",
    "ァ",
    "い",
    "イ",
    "ｲ",
    "か",
    "が",
    "カ",
    "ガ",
    "ｶﾞ",
    "ｶ",
    "つ",
    "っ",
    "ー",
    "かー",
    "かあ",
    "かぁ",
    "ゝ",
    "ヽ",
    "ん",
    "ン",
    "ヴ",
    "ゔ",
    "亜",
    "一",
    "日",
    "本",
    "中",
    "々",
    "、",
    "。",
    "「",
    "α",
    "β",
    "Ω",
    "а",
    "б",
    "е",
    "ё",
    "Ё",
    "ж",
    "я",
    "가",
    "나",
    "ㄱ",
    "ก",
    "ข",
    "เก",
    "กา",
    "๑",
    "ä1",
    "二",
    "三",
    "人",
    "大",
    "阿",
    "愛",
    "安",
    "王",
    "山",
    "水",
    "木",
    "火",
    "金",
    "土",
    "東",
    "京",
    "国",
    "學",
    "学",
    "生",
    "漢",
    "字",
    "家",
    "乙",
    "丁",
    "龍",
    "嗎",
    "각",
    "간",
    "하",
    "힣",
    "ㅏ",
    "n",
    "nz",
    "o",
    "oa",
    "c",
    "ch",
    "cz",
    "d",
    "l",
    "ll",
    "lz",
    "i",
    "ij",
    "ĳ",
    "iz",
    "y",
    "č",
    "ő",
    "ā",
    "ı",
    "e",
    "ez",
    "u",
    "uz",
    "s",
    "sz",
    "t",
    "ез",
    "ёа",
    "ь",
    "ъ",
    "і",
    "ά",
    "ς",
    "σ",
    "ค",
    "ฤ",
    "ๆ",
    "่",
    "ก่",
    "กะ",
    "แก",
    "ใก",
    "ไก",
    "€",
    "£",
    "¥",
    "°",
    "½",
    "²",
    "™",
    "©",
    "→",
    "★",
    "♪",
    "Ⅰ",
    "①",
    "㈱",
    "ﾟ",
    "゛",
]
_CONTEXT_STRINGS = [
    # Hangul jamo decomposition and ko hanja readings.
    "가z",
    "갂a",
    "가나",
    "가ㄱ",
    "각",
    "갂",
    "각ㄱ",
    "家z",
    "가a",
    "刻",
    "刻z",
    "각a",
    # Spaces, ideographic space, control characters, word-sort symbols.
    "a　b",
    "a\x01b",
    "a\x01",
    "\x01",
    "co-op",
    "coop",
    "abc",
    # Combining marks.
    "á",
    "áb",
    "áb",
    # Thai prefix vowels before consonants and non-consonants.
    "เ",
    "เa",
    "เ1",
    "ก",
    # French accents, German umlaut, sharp s inside words.
    "cote",
    "coté",
    "côte",
    "côté",
    "aea",
    "äz",
    "straße",
    "strasse",
    "strasze",
    # Japanese prolonged-sound and iteration marks after hiragana, katakana and others.
    "かゝ",
    "かゞ",
    "かヽ",
    "かヾ",
    "かか",
    "かが",
    "がゝ",
    "がか",
    "カゝ",
    "カゞ",
    "カヽ",
    "カヾ",
    "カー",
    "カカ",
    "カガ",
    "カア",
    "んー",
    "aー",
    "中ー",
    "ンー",
]


def probe_strings() -> list[str]:
    seen: dict[str, None] = {}
    for s in _BASE_STRINGS + _CONTEXT_STRINGS:
        seen.setdefault(s, None)
    return list(seen)


EQ_PAIRS = [
    ("ß", "ss"),
    ("ĳ", "ij"),
    ("①", "1"),
    ("²", "2"),
    ("Ａ", "A"),
    ("ａ", "A"),
    ("ｱ", "ア"),
    ("ぁ", "あ"),
    ("ア", "あ"),
    ("ヴ", "ゔ"),
    ("かー", "かあ"),
    ("Ⅰ", "i"),
    ("Ä", "ä"),
    ("A", "a"),
    ("ｶﾞ", "ガ"),
    ("１", "1"),
    ("　", " "),
    ("ゝ", "ヽ"),
    ("ё", "Ё"),
    ("ς", "σ"),
    ("ä", "a"),
    ("æ", "ae"),
    ("ss", "ß"),
    ("เก", "กา"),
    ("かゝ", "かか"),
    ("ガ", "ｶﾞ"),
]
EQ_TEMPLATES = {
    "eq": '=IF("{a}"="{b}",1,0)',
    "match0": '=IFERROR(MATCH("{a}",{{"zz";"{b}"}},0),0)',
    "vlookupF": '=IFERROR(VLOOKUP("{a}",{{"zz",1;"{b}",2}},2,FALSE),0)',
    "hlookupF": '=IFERROR(HLOOKUP("{a}",{{"zz","{b}";1,2}},2,FALSE),0)',
    "xlookup": '=IFERROR(XLOOKUP("{a}",{{"zz";"{b}"}},{{1;2}}),0)',
    "xmatch0": '=IFERROR(XMATCH("{a}",{{"zz";"{b}"}}),0)',
    "xmatchbin": '=IFERROR(XMATCH("{a}",{{"0";"{b}"}},0,2),0)',
    "countif": '=COUNTIF(B1:B2,"{a}")',
    "countifeq": '=COUNTIF(B1:B2,"={a}")',
    "sumif": '=SUMIF(B1:B2,"{a}",C1:C2)',
    "unique": '=ROWS(UNIQUE({{"{a}";"{b}"}}))',
    "switch": '=SWITCH("{a}","{b}",1,0)',
    "groupby": '=ROWS(GROUPBY({{"{a}";"{b}"}},{{1;2}},SUM,0,0))',
    "filter": '=IFERROR(ROWS(FILTER({{"{b}"}},{{"{b}"}}="{a}")),0)',
    "exact": '=IF(EXACT("{a}","{b}"),1,0)',
    "dcount": "=DCOUNTA(B1:B2,1,E1:E2)",
}

TAIL_X = ["a", "A", "ア", "あ", "ss", "ß", "1", "-", " ", "é", "é", "中", "가", "ก", "ー", ""]
# Preferred trailing-ignorable representatives; the first 8 present in the locale's mid
# (or, for zh, full) class are used.
TAIL_PREFERRED = [0x01, 0x1F, 0x7F, 0xAD, 0x200B, 0x200C, 0x200D, 0x2060, 0xFEFF, 0x0640]


def expansion_candidates() -> list[tuple[int, str]]:
    extra = {0xDF: "ss", 0x1E9E: "SS", 0xE6: "ae", 0xC6: "AE", 0x153: "oe", 0x152: "OE", 0xFE: "th", 0xDE: "TH"}
    out = []
    for cp in list(range(1, 0xD800)) + list(range(0xE000, 0xFFFE)):
        if cp in extra:
            out.append((cp, extra[cp]))
            continue
        k = unicodedata.normalize("NFKC", chr(cp))
        num = unicodedata.numeric(chr(cp), None)
        if len(k) >= 2:
            out.append((cp, k))
        elif num is not None and num >= 10 and num == int(num):
            out.append((cp, str(int(num))))
    return out


# ---------------------------------------------------------------- Excel side


class Excel:
    def __init__(self) -> None:
        import xlwings as xw  # only the capture path needs Excel

        self.xw = xw
        self.app = xw.apps.active if xw.apps.count else xw.App(visible=False)
        for attempt in range(4):
            try:
                self.book = self.app.books.add()
                break
            except Exception:  # Excel refuses `make new workbook` for a while after launch
                if attempt == 3:
                    raise
                time.sleep(10)

    def close(self) -> None:
        self.book.close()

    def sheet(self):
        return self.book.sheets.add()

    def spill(self, sh, formula: str, rows: int, cols: int, at: str = "A1", text=False):
        cell = sh.range(at)
        cell.formula2 = formula
        self.app.calculate()
        for _ in range(10):
            if cell.formula2:
                break
            time.sleep(1)
        else:
            raise SystemExit(f"Excel rejected the formula: {formula[:120]}")
        r0, c0 = cell.row, cell.column
        for _ in range(3):
            v = sh.range((r0, c0), (r0 + rows - 1, c0 + cols - 1)).options(ndim=2).value
            bad = [
                (i, r)
                for i, r in enumerate(v)
                if len(r) != cols or (not text and any(c is None or isinstance(c, str) for c in r))
            ]
            if len(v) == rows and not bad:
                return v
            time.sleep(2)
        raise SystemExit(f"spill of {formula[:120]} gave {len(v)} rows; first bad {bad[:3]}")

    def column(self, sh, values, col: int = 1) -> None:
        if values:
            sh.range((1, col)).options(transpose=True).value = list(values)


def _flag(v) -> int:
    return int(round(float(v)))


def capture_bmp(x: Excel, exclude: list[int] = ()):
    """Sorted BMP with per-character class probes, and the descending-input sort.

    Formulas set through AppleScript are silently dropped past roughly 255 characters,
    so the input code points are spilled into column M first and referenced as M1#.
    """
    total = 63453 - len(exclude)
    sh = x.sheet()
    keep = "NOT(ISERROR(UNICHAR(n)))"
    if exclude:
        x.column(sh, exclude, 11)
        keep += f"*ISNA(XMATCH(n,K1:K{len(exclude)}))"
    out = []
    for seq in (BMP_ASC, BMP_DESC):
        x.spill(sh, f"=LET(n,{seq},FILTER(n,{keep}))", total, 1, at="M1")
        out.append(
            x.spill(
                sh,
                '=LET(s,SORT(UNICHAR(M1#)),HSTACK(UNICODE(s),IF("a"&s="a",1,0),'
                'IF("a"&s&"b"="ab",1,0),IF("a"&s&"z">"aya",1,0),IF("a"&s&"a"<"ab",1,0),'
                'IF("a"&s&"b"<"ab"&UNICHAR(769),1,0)))',
                total,
                6,
            )
        )
    asc, desc = out
    assert asc[-1][0] is not None and desc[-1][0] is not None, "BMP spill shorter than expected"
    sh.delete()
    rows = [[_flag(c) for c in r] for r in asc]
    return rows, [_flag(r[0]) for r in desc]


def classify(row) -> str:
    _, a_eq, ab_eq, anc1, anc2, wordlt = row
    if a_eq and ab_eq:
        return "full"
    if ab_eq:
        return "mid"
    if anc1 and anc2:
        return "word" if wordlt else "sec"
    return "base"


def _groups(asc: list[int], desc: list[int]) -> tuple[dict[int, int], list[int]]:
    pos_d = {cp: i for i, cp in enumerate(desc)}
    gid, out = 0, {}
    for i, cp in enumerate(asc):
        if i and pos_d[cp] > pos_d[asc[i - 1]]:
            gid += 1
        out[cp] = gid
    # The mirror rule on the descending list must give the same partition.
    pos_a = {cp: i for i, cp in enumerate(asc)}
    bad = []
    for i in range(1, len(desc)):
        same = pos_a[desc[i]] < pos_a[desc[i - 1]]
        if same != (out[desc[i]] == out[desc[i - 1]]):
            bad += [desc[i - 1], desc[i]]
    return out, bad


def dynamic_groups(asc: list[int], desc: list[int]) -> dict[int, int]:
    """Group id per code point: a run whose order flips between the two stable sorts."""
    out, bad = _groups(asc, desc)
    if bad:
        raise SystemExit(f"SORT has no consistent partition around U+{bad[0]:04X}")
    return out


def _windows(desc: list[int], bad: list[int], margin: int = 12) -> list[list[int]]:
    pos = {cp: i for i, cp in enumerate(desc)}
    spans = sorted(pos[c] for c in bad)
    out, lo, hi = [], spans[0], spans[0]
    for p in spans[1:]:
        if p - hi > 2 * margin:
            out.append(desc[max(0, lo - margin) : hi + margin + 1])
            lo = p
        hi = p
    out.append(desc[max(0, lo - margin) : hi + margin + 1])
    return out


_PAIR_SORT = (
    "=LET(c,{c},MAKEARRAY({n},{n},LAMBDA(i,j,LET(a,INDEX(c,i),b,INDEX(c,j),"
    "f,INDEX(SORT(HSTACK(VSTACK(a,b),{{1;2}})),1,2),g,INDEX(SORT(HSTACK(VSTACK(b,a),{{2;1}})),1,2),"
    "IF(i=j,0,IF(f<>g,0,IF(f=1,-1,1)))))))"
)


def pair_sort_signs(x: Excel, sh, values: str, n: int, at: str) -> list[list[int]]:
    """SORT-strength sign of every ordered pair of `values`, from two-element stable sorts."""
    return [[_flag(v) for v in r] for r in x.spill(sh, _PAIR_SORT.format(c=values, n=n), n, n, at=at)]


def sort_relation(x: Excel, cps: list[int]) -> dict[tuple[int, int], int]:
    sh = x.sheet()
    x.column(sh, cps)
    n = len(cps)
    v = pair_sort_signs(x, sh, f"UNICHAR(A1:A{n})", n, "C1")
    sh.delete()
    return {(a, b): v[i][j] for i, a in enumerate(cps) for j, b in enumerate(cps)}


def intransitive_set(rel: dict[tuple[int, int], int], cps: list[int]) -> list[int]:
    """Smallest set of code points whose removal leaves `rel` a weak order."""

    def violations(keep):
        bad = set()
        for a, b, c in itertools.permutations(keep, 3):
            ab, bc, ac = rel[(a, b)], rel[(b, c)], rel[(a, c)]
            if ab <= 0 and bc <= 0 and ac != min(ab, bc):
                bad.add((a, b, c))
        return bad

    viol = violations(cps)
    involved = sorted({c for t in viol for c in t})
    for size in range(5):
        for drop in itertools.combinations(involved, size):
            if all(set(t) & set(drop) for t in viol):
                return list(drop)
    # Too many for an exhaustive search: drop the character in the most violations first.
    drop, left = [], set(viol)
    while left:
        worst = max(involved, key=lambda c: (sum(c in t for t in left), -c))
        drop.append(worst)
        left = {t for t in left if worst not in t}
    return drop


def bmp_order(x: Excel):
    """Classified BMP rows, and the two sorts with intransitive characters left out."""
    rows, desc = capture_bmp(x)
    flags = {r[0]: r for r in rows}
    full_asc = [r[0] for r in rows]
    anomalies: list[int] = []
    asc = full_asc
    while True:
        _, bad = _groups(asc, desc)
        if not bad:
            break
        found = []
        for w in _windows(desc, bad):
            w = [c for c in w if c not in anomalies]
            found += intransitive_set(sort_relation(x, w), w)
        if not found:
            raise SystemExit("SORT disagrees without an intransitive triple")
        anomalies = sorted(set(anomalies) | set(found))
        sub, desc = capture_bmp(x, anomalies)
        asc = [r[0] for r in sub]
    # Anomalies keep their place after the preceding ordinary character of the full sort.
    after: dict[int, list[int]] = {}
    prev = None
    for cp in full_asc:
        if cp in anomalies:
            after.setdefault(prev, []).append(cp)
        else:
            prev = cp
    merged = after.get(None, [])
    for cp in asc:
        merged += [cp] + after.get(cp, [])
    return merged, flags, desc, anomalies


def adjacent_tests(x: Excel, cps: list[int], cls: str):
    """eq, primary-break, secondary-break and `<` between consecutive cps."""
    if len(cps) < 2:
        return []
    pre, post = AFFIX[cls]
    sh = x.sheet()
    x.column(sh, cps)
    n = len(cps)
    v = x.spill(
        sh,
        f"=LET(c,{pre}UNICHAR(A1:A{n}){post},p,DROP(c,-1),s,DROP(c,1),"
        f'HSTACK(IF(p=s,1,0),IF(p&"z"<s&"a",1,0),IF(p&{ACUTE}<s&"a",1,0),IF(p<s,1,0)))',
        n - 1,
        4,
        at="C1",
    )
    sh.delete()
    return [[_flag(c) for c in r] for r in v]


def cp_pairs(x: Excel, pairs: list[tuple[int, int]], cls: str):
    """(a<b, a=b) for single code points under the class affixes."""
    if not pairs:
        return []
    pre, post = AFFIX[cls]
    sh = x.sheet()
    x.column(sh, [a for a, _ in pairs], 1)
    x.column(sh, [b for _, b in pairs], 2)
    n = len(pairs)
    v = x.spill(
        sh,
        f"=LET(a,{pre}UNICHAR(A1:A{n}){post},b,{pre}UNICHAR(B1:B{n}){post},HSTACK(IF(a<b,1,0),IF(a=b,1,0)))",
        n,
        2,
        at="D1",
    )
    sh.delete()
    return [[_flag(c) for c in r] for r in v]


# "_" separates the code points: a space is a digit-group separator in fr and ru.
_STR = 'LAMBDA(x,IF(x="","",CONCAT(UNICHAR(--TEXTSPLIT(x&"","_")))))'


def str_pairs(x: Excel, pairs: list[tuple[str, str]], tests: list[str]):
    """Evaluate `tests` (expressions over a and b) for arbitrary string pairs."""
    sh = x.sheet()

    def to_cell(s):
        return "_".join(str(ord(ch)) for ch in s) if s else None

    x.column(sh, [to_cell(a) for a, _ in pairs], 1)
    x.column(sh, [to_cell(b) for _, b in pairs], 2)
    n = len(pairs)
    x.spill(sh, f"=MAP(A1:A{n},{_STR})", n, 1, at="D1", text=True)
    x.spill(sh, f"=MAP(B1:B{n},{_STR})", n, 1, at="E1", text=True)
    cols = [
        x.spill(sh, f"=LET(a,D1#,b,E1#,IF({t},1,0))", n, 1, at=sh.range((1, 7 + k)).address)
        for k, t in enumerate(tests)
    ]
    sh.delete()
    return [[_flag(col[i][0]) for col in cols] for i in range(n)]


def capture_order(x: Excel, expanding: set[int]):
    """Rows of order_<loc>.tsv. Expanding characters take their weights from the
    expansion table, so they are marked `X` and skipped when ranking their neighbours."""
    asc, flags, desc, anomalies = bmp_order(x)
    cls = {cp: classify(flags[cp]) for cp in asc}
    odd = set(anomalies)
    group = dynamic_groups([c for c in asc if c not in odd], desc)
    for cp in anomalies:
        group[cp] = ("odd", cp)
    out = []  # (cp, cls, rel) in final order
    for c in ("full", "mid"):
        out += [(cp, c, "=") for cp in asc if cls[cp] == c]
    for c in ("word", "sec", "base"):
        members: dict[int, list[int]] = {}
        follows: dict[int | None, list[int]] = {}
        prev = None
        for cp in asc:
            if cls[cp] != c:
                continue
            if cp in expanding:
                follows.setdefault(prev, []).append(cp)
            else:
                members.setdefault(group[cp], []).append(cp)
                prev = cp
        pairs = [(a, b) for m in members.values() if len(m) > 1 for a in m for b in m if a != b]
        res = dict(zip(pairs, cp_pairs(x, pairs, c)))
        seq, starts = [], set()
        for m in members.values():
            rank = {a: sum(res[(b, a)][0] for b in m if b != a) for a in m}
            ordered = sorted(m, key=lambda a: (rank[a], a))
            for a in m:
                for b in m:
                    if a < b and (
                        res[(a, b)][1] != res[(b, a)][1] or res[(a, b)][0] + res[(b, a)][0] + res[(a, b)][1] != 1
                    ):
                        raise SystemExit(f"inconsistent pair U+{a:04X} U+{b:04X}")
            for j in range(1, len(ordered)):
                p, s = ordered[j - 1], ordered[j]
                if not (res[(p, s)][1] or res[(p, s)][0]):
                    raise SystemExit(f"rank order broken at U+{p:04X} U+{s:04X}")
            starts.add(len(seq))
            seq += ordered
        adj = adjacent_tests(x, seq, c)
        rows = [(e, c, "X") for e in follows.get(None, [])]
        for i, a in enumerate(seq):
            if i == 0:
                rel = "S" if c == "sec" else "P"
            elif i not in starts:
                rel = "=" if res[(seq[i - 1], a)][1] else "C"
            else:
                eq, pbreak, sbreak, _ = adj[i - 1]
                if eq and not ({seq[i - 1], a} & odd):
                    raise SystemExit(f"U+{seq[i - 1]:04X} = U+{a:04X} but SORT keeps them apart")
                if eq:
                    rel = "="
                elif c == "word" or (c == "base" and pbreak):
                    rel = "P"
                else:
                    rel = "S" if sbreak else "D"
            rows.append((a, c, rel))
            rows += [(e, c, "X") for e in follows.get(a, [])]
        out += rows
    return out, anomalies


def capture_expansions(x: Excel):
    cands = expansion_candidates()
    # The last test is SORT-strength equality; XMATCH needs a per-row lookup, hence MAP.
    tests = [
        "a=b",
        'a&"z"<b&"a"',
        'b&"z"<a&"a"',
        f'a&{ACUTE}<b&"a"',
        f'b&{ACUTE}<a&"a"',
        "MAP(a,b,LAMBDA(u,v,ISNUMBER(XMATCH(u,VSTACK(v)))))",
    ]
    res = str_pairs(x, [(chr(cp), d) for cp, d in cands], tests)
    out = []
    for (cp, d), (eq, p1, p2, s1, s2, dyn) in zip(cands, res):
        if p1 or p2:
            continue
        level = "=" if eq else "S" if (s1 or s2) else "C" if dyn else "D"
        out.append((cp, d, level))
    return out


def capture_tail(x: Excel, order):
    cls = {cp: c for cp, c, _ in order}
    want = "mid" if any(c == "mid" for c in cls.values()) else "full"
    reps = [cp for cp in TAIL_PREFERRED if cls.get(cp) == want][:8]
    reps += [cp for cp, c, _ in order if c == want and cp not in reps][: 8 - len(reps)]
    rows, pairs = [], []
    for k, c in enumerate(reps):
        c2 = reps[(k + 1) % len(reps)]
        for xs in TAIL_X:
            for ys in TAIL_X:
                for form, a, b in (
                    ("x|y+c", xs, ys + chr(c)),
                    ("x+c|y", xs + chr(c), ys),
                    ("x+c|y+c2", xs + chr(c), ys + chr(c2)),
                ):
                    rows.append((form, c, c2, xs, ys))
                    pairs.append((a, b))
    res = str_pairs(x, pairs, ["a<b", "a=b"])

    def sign(lt, eq):
        return "-" if lt else "0" if eq else "+"

    return [(f, f"U+{c:04X}", f"U+{c2:04X}", enc(xs), enc(ys), sign(*r)) for (f, c, c2, xs, ys), r in zip(rows, res)]


def capture_matrix(x: Excel, strings: list[str]):
    n = len(strings)
    sh = x.sheet()
    sh.range((1, 1), (n, 1)).formula = [["=" + "&".join(f"UNICHAR({ord(ch)})" for ch in s)] for s in strings]
    got = sh.range((1, 1), (n, 1)).options(ndim=2).value
    if [str(r[0]) for r in got] != strings:
        raise SystemExit("probe strings did not round-trip through Excel")
    m = x.spill(sh, f"=LET(c,A1:A{n},r,TRANSPOSE(c),IF(c<r,-1,IF(c=r,0,1)))", n, n, at="C1")
    dyn = pair_sort_signs(x, sh, f"A1:A{n}", n, sh.range((n + 3, 3)).address)
    sh.delete()
    sym = {-1: "-", 0: "0", 1: "+"}
    return [(i, "".join(sym[_flag(v)] for v in m[i]), "".join(sym[v] for v in dyn[i])) for i in range(n)]


def capture_eq(x: Excel):
    sh = x.sheet()
    out = []
    for a, b in EQ_PAIRS:
        sh.range("B1").formula = '="h"'
        sh.range("B2").formula = f'="{b}"'
        sh.range("C1").value = 1
        sh.range("C2").value = 2
        sh.range("E1").formula = '="h"'
        sh.range("E2").formula = f'="{a}"'
        row = [enc(a), enc(b)]
        for t in EQ_TEMPLATES.values():
            c = sh.range("G1")
            c.formula2 = t.format(a=a, b=b)
            v = c.value
            row.append("n" if v is None or isinstance(v, str) else _flag(v))
        out.append(row)
    sh.delete()
    return out


FINGERPRINT = {
    "month": '=TEXT(DATE(2024,1,5),"mmmm")',
    "decimal": '=1.5&""',
    "han": '=IF("阿"<"一",1,0)',
    "kana_greek": '=IF("あ"<"α",1,0)',
    "enye": '=IF("ñ">"nz",1,0)',
    "cyr_latin": '=IF("а"<"a",1,0)',
}


def fingerprint(x: Excel) -> list[str]:
    sh = x.sheet()
    out = []
    for f in FINGERPRINT.values():
        c = sh.range("A1")
        c.formula2 = f
        v = c.value
        out.append(str(_flag(v)) if isinstance(v, float) else str(v))
    sh.delete()
    return out


def capture(tag: str) -> None:
    x = Excel()
    try:
        t0 = time.time()
        fp = fingerprint(x)
        meta = upsert_meta(tag, [str(x.app.version)] + fp)
        print(f"[{tag}] fingerprint {fp}", flush=True)
        expansions = capture_expansions(x)
        write_tsv(
            DATA / f"expansion_{tag}.tsv",
            ["cp", "expands_to", "level"],
            [(f"U+{cp:04X}", enc(d), lv) for cp, d, lv in expansions],
        )
        order, anomalies = capture_order(x, {cp for cp, _, _ in expansions})
        write_tsv(DATA / f"order_{tag}.tsv", ["cp", "cls", "rel"], [(f"U+{cp:04X}", c, r) for cp, c, r in order])
        write_tsv(DATA / f"anomalies_{tag}.tsv", ["cp"], [(f"U+{cp:04X}",) for cp in anomalies])
        print(
            f"[{tag}] order {len(order)} rows, anomalies {[hex(c) for c in anomalies]} {time.time() - t0:.0f}s",
            flush=True,
        )
        write_tsv(FIX / f"tail_{tag}.tsv", ["form", "c", "c2", "x", "y", "sign"], capture_tail(x, order))
        strings = probe_strings()
        write_tsv(FIX / "probe_strings.tsv", ["id", "string"], [(i, enc(s)) for i, s in enumerate(strings)])
        write_tsv(FIX / f"matrix_{tag}.tsv", ["a", "compare", "dynamic"], capture_matrix(x, strings))
        write_tsv(FIX / f"eq_{tag}.tsv", ["a", "b", *EQ_TEMPLATES], capture_eq(x))
        print(f"[{tag}] done {time.time() - t0:.0f}s; meta rows {len(meta)}", flush=True)
    finally:
        x.close()


def upsert_meta(tag: str, values: list[str]):
    path = DATA / "meta.tsv"
    rows = {r[0]: r for r in read_tsv(path)} if path.exists() else {}
    rows[tag] = [tag, *values]
    order = [t for t, _, _ in LOCALES]
    write_tsv(path, ["locale", "excel_version", *FINGERPRINT], [rows[t] for t in order if t in rows])
    return rows


# ---------------------------------------------------------------- locale switching


def _switch(lang: str, loc: str) -> None:
    subprocess.run(["osascript", "-e", 'tell application "Microsoft Excel" to quit saving no'], capture_output=True)
    for _ in range(30):
        if subprocess.run(["pgrep", "-x", "Microsoft Excel"], capture_output=True).returncode:
            break
        time.sleep(1)
    subprocess.run(["pkill", "-x", "Microsoft Excel"], capture_output=True)
    time.sleep(2)
    dom = "com.microsoft.Excel"
    if lang:
        subprocess.run(["defaults", "write", dom, "AppleLanguages", "-array", lang], check=True)
        subprocess.run(["defaults", "write", dom, "AppleLocale", loc], check=True)
    else:
        subprocess.run(["defaults", "delete", dom, "AppleLanguages"], capture_output=True)
        subprocess.run(["defaults", "delete", dom, "AppleLocale"], capture_output=True)
    subprocess.run(["open", "-a", "Microsoft Excel"], check=True)
    time.sleep(25)


def capture_all(only: list[str] | None) -> None:
    import xlwings as xw

    for app in xw.apps if xw.apps.count else []:
        for b in app.books:
            used = [sh.used_range for sh in b.sheets]
            if b.fullname != b.name or any(u.address != "$A$1" or u.value is not None for u in used):
                raise SystemExit(f"Excel has open work ({b.name}); close it before switching")
    try:
        for tag, lang, loc in LOCALES:
            if only and tag not in only:
                continue
            for attempt in range(2):
                _switch(lang, loc)
                try:
                    capture(tag)
                    break
                except Exception as exc:  # a fresh Excel sometimes refuses new workbooks
                    if attempt or "make" not in str(exc):
                        raise
                    print(f"[{tag}] relaunching Excel after: {exc}", flush=True)
    finally:
        _switch("", "")
        print("ja_JP restored", flush=True)


# ---------------------------------------------------------------- validation


def single_keys(order_rows, dynamic: bool) -> dict[int, tuple]:
    """Sort key of each one-character string, from the order table."""
    keys, ctr = {}, {}
    for cp_s, c, rel in order_rows:
        cp = int(cp_s[2:], 16)
        if c in ("full", "mid"):
            keys[cp] = ((), (), (), (), (), 1 if c == "mid" else 0)
            continue
        if rel == "X":
            continue
        p, s, d, k = ctr.get(c, (0, 0, 0, 0))
        if rel == "P":
            p, s, d, k = p + 1, 0, 0, 0
        elif rel == "S":
            s, d, k = s + 1, 0, 0
        elif rel == "D":
            d, k = d + 1, 0
        elif rel == "C":
            k += 1
        ctr[c] = (p, s, d, k)
        k_eff = () if dynamic else (k,)
        if c == "base":
            keys[cp] = ((p,), (s,), (d,), k_eff, (), 0)
        elif c == "sec":
            keys[cp] = ((), (s,), (d,), k_eff, (), 0)
        else:
            keys[cp] = ((), (), (), (), ((0, p, s, d) + k_eff,), 0)
    return keys


def validate(only: list[str] | None) -> int:
    bad_total = 0
    strings = [dec(r[1]) for r in read_tsv(FIX / "probe_strings.tsv")]
    n = len(strings)
    exp_by_loc = {}
    tags = [t for t, _, _ in LOCALES if not only or t in only]
    for tag in tags:
        order = read_tsv(DATA / f"order_{tag}.tsv")
        odd = {int(r[0][2:], 16) for r in read_tsv(DATA / f"anomalies_{tag}.tsv")}
        counts = {c: 0 for c in CLASSES}
        rels = {r: 0 for r in "PSDCX="}
        for _, c, rel in order:
            counts[c] += 1
            rels[rel] += 1
        cps = [int(r[0][2:], 16) for r in order]
        problems = []
        if len(set(cps)) != len(cps) or len(cps) != 63453:
            problems.append(f"order has {len(cps)} rows / {len(set(cps))} unique")
        mat = read_tsv(FIX / f"matrix_{tag}.tsv")
        for strength, col in (("compare", 1), ("dynamic", 2)):
            keys = single_keys(order, strength == "dynamic")
            for i, a in enumerate(strings):
                row = mat[i][col]
                if len(row) != n:
                    problems.append(f"matrix row {i} has {len(row)} cells")
                    break
                for j, b in enumerate(strings):
                    got = row[j]
                    rev = mat[j][col][i]
                    if {"-": "+", "+": "-", "0": "0"}[got] != rev:
                        problems.append(f"{strength} antisymmetry {enc(a)} | {enc(b)}")
                    if len(a) == 1 and len(b) == 1 and ord(a) in keys and ord(b) in keys and not {ord(a), ord(b)} & odd:
                        ka, kb = keys[ord(a)], keys[ord(b)]
                        want = "-" if ka < kb else "0" if ka == kb else "+"
                        if want != got:
                            problems.append(f"{strength} {enc(a)} vs {enc(b)}: order {want}, matrix {got}")
        exp = read_tsv(DATA / f"expansion_{tag}.tsv")
        exp_by_loc[tag] = {(r[0], r[1]): r[2] for r in exp}
        index = {s: i for i, s in enumerate(strings)}
        for cp_s, to, level in exp:
            a, b = dec(cp_s), dec(to)
            if level not in ("=", "S", "D", "C"):
                problems.append(f"expansion level {level}")
            if a in index and b in index:
                cmp_ = mat[index[a]][1][index[b]]
                dyn = mat[index[a]][2][index[b]]
                if (cmp_ == "0") != (level == "=") or (dyn == "0") != (level in ("=", "C")):
                    problems.append(f"expansion {cp_s}->{to} {level} vs matrix {cmp_}{dyn}")
        tail = read_tsv(FIX / f"tail_{tag}.tsv")
        if len(tail) != 8 * len(TAIL_X) ** 2 * 3:
            problems.append(f"tail has {len(tail)} rows")
        eq = read_tsv(FIX / f"eq_{tag}.tsv")
        if len(eq) != len(EQ_PAIRS) or any(len(r) != 2 + len(EQ_TEMPLATES) for r in eq):
            problems.append("eq table shape")
        bad_total += len(problems)
        print(f"{tag}: cls {counts} rel {rels} expansions {len(exp)} tail {len(tail)} contradictions {len(problems)}")
        for p in problems[:10]:
            print("   ", p)
    for tag in tags:
        base = "zh_CN" if tag in Z_LOCALES else "en_US"
        if base not in exp_by_loc:
            continue
        diff = set(exp_by_loc[tag].items()) ^ set(exp_by_loc[base].items())
        print(f"expansions {tag} vs {base}: {len(diff)} differing rows")

    # Switching check: locales whose orders coincide (outside the RTL scripts, which vary
    # between Excel launches) must be exactly the groups SAME_ORDER lists.
    def rtl(cp):
        return 0x590 <= cp <= 0x8FF or 0xFB1D <= cp <= 0xFDFF or 0xFE70 <= cp <= 0xFEFF or cp in (0x200E, 0x200F)

    sig = {}
    for tag in tags:
        rows = read_tsv(DATA / f"order_{tag}.tsv")
        sig[tag] = hash(tuple((r[0], r[1], r[2]) for r in rows if not rtl(int(r[0][2:], 16))))
    groups = sorted(sorted(t for t in tags if sig[t] == v) for v in set(sig.values()))
    want = sorted(sorted(t for t in g if t in tags) for g in SAME_ORDER if any(t in tags for t in g))
    want += sorted([t] for t in tags if not any(t in g for g in SAME_ORDER))
    if sorted(groups) != sorted(want):
        print(f"order groups {groups} differ from expected {sorted(want)}")
        bad_total += 1
    print("TOTAL contradictions", bad_total)
    return 1 if bad_total else 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("command", nargs="?", choices=["capture", "capture-all"])
    ap.add_argument("--locale", action="append")
    ap.add_argument("--validate", action="store_true")
    args = ap.parse_args()
    if args.validate:
        return validate(args.locale)
    if args.command == "capture":
        if not args.locale or len(args.locale) != 1:
            ap.error("capture needs exactly one --locale")
        capture(args.locale[0])
    elif args.command == "capture-all":
        capture_all(args.locale)
    else:
        ap.error("nothing to do")
    return 0


if __name__ == "__main__":
    sys.exit(main())
