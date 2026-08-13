#!/usr/bin/env python3
"""Harvests MS-XLSB function ids for native Excel built-ins from Excel itself.

`src/io/xlsb/func_id_table.cpp` maps a 16-bit function id to a name in both
directions: the reader resolves `PtgFunc(id)` / `PtgFuncVar(id)` to a name and
the writer does the reverse. A wrong id is therefore a silent substitution of
a different function, so no id in that table may come from recollection. This
tool derives ids from bytes Excel itself wrote.

Pipeline:

  build   Drives Excel 365 (macOS, via the oracle driver's xlwings stack),
          types one short probe formula per function per probed arity, and
          saves the workbook in the Excel Binary Workbook format.
  decode  Walks `xl/worksheets/sheet1.bin`, locates every formula cell, and
          parses its whole Ptg stream token by token. The parse must land
          exactly on the end of the stream and finish on a `PtgFunc` /
          `PtgFuncVar` token -- the outermost call of a probe formula is the
          last token of its postfix stream. A stream that does not parse
          cleanly is reported as unresolved rather than guessed at.
  emit    Prints the `func_id_table.cpp` rows for the decoded ids.

Every function is probed at its minimum arity and, when it takes optional
arguments, again at its maximum. The two observations must agree on the id
and on the token kind; a function seen as fixed-arity (`PtgFunc`) at two
different argument counts is contradictory and is reported instead of
emitted, because the reader derives a fixed-arity call's argument count from
the table's `arg_min` alone.

The probe formulas use literal arguments (and a small helper block for the
few functions that need a range) so each Ptg stream stays short: the
AppleEvent formula setter silently truncates long formulas.

Usage::

    tools/oracle/.venv/bin/python tools/dev/xlsb_func_id_harvest.py build \\
        --out tests/fixtures/excel/xlsb_func_ids.xlsb
    tools/oracle/.venv/bin/python tools/dev/xlsb_func_id_harvest.py decode \\
        tests/fixtures/excel/xlsb_func_ids.xlsb
    tools/oracle/.venv/bin/python tools/dev/xlsb_func_id_harvest.py emit \\
        tests/fixtures/excel/xlsb_func_ids.xlsb

`emit` exits 1 when a decoded id would displace an existing table row. Such a
row is withheld from the output and reported as a `// collision:` comment
instead: `DBCS` is the standing case, because Excel encodes it with the id
already mapped to `JIS` (the two names are one function) and the table maps
an id to exactly one name.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile
from typing import Dict, Iterator, List, NamedTuple, Optional, Tuple

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Date serials used by the financial probes (1900 date system):
#   43831 = 2020-01-01, 43922 = 2020-04-01, 44013 = 2020-07-01,
#   44197 = 2021-01-01.
ISSUE = 43831
SETTLEMENT = 43922
FIRST_COUPON = 44013
MATURITY = 44197

# Helper cells the range-taking probes read. Written before the formulas.
HELPERS: Dict[str, float] = {
    "D1": -100.0,
    "D2": 120.0,
    "E1": float(ISSUE),
    "E2": float(MATURITY),
    "F1": 0.09,
    "F2": 0.11,
    "F3": 0.10,
}


class Probe(NamedTuple):
    """One probe cell: a function, a formula, and the formula's arity.

    `arity` is stated rather than counted out of the formula text: it is
    what a fixed-arity (`PtgFunc`) row's `arg_min` becomes, and the reader
    pops exactly that many operands for such a call.
    """

    name: str
    formula: str
    arity: int


# One entry per probe cell, in row order starting at PROBE_ROW_0. The row
# index is stable even if Excel rejects a formula, so a rejected probe
# leaves a hole rather than shifting its neighbours. Functions with
# optional arguments appear twice: once at the minimum arity, once at the
# maximum.
PROBES: List[Probe] = [
    Probe("ACCRINT", f"=ACCRINT({ISSUE},{FIRST_COUPON},{SETTLEMENT},0.1,1000,2)", 6),
    Probe("ACCRINT", f"=ACCRINT({ISSUE},{FIRST_COUPON},{SETTLEMENT},0.1,1000,2,0,TRUE)", 8),
    Probe("ACCRINTM", f"=ACCRINTM({ISSUE},{MATURITY},0.1,1000)", 4),
    Probe("ACCRINTM", f"=ACCRINTM({ISSUE},{MATURITY},0.1,1000,0)", 5),
    Probe("AMORDEGRC", f"=AMORDEGRC(2400,{ISSUE},{MATURITY},300,1,0.15)", 6),
    Probe("AMORDEGRC", f"=AMORDEGRC(2400,{ISSUE},{MATURITY},300,1,0.15,1)", 7),
    Probe("AMORLINC", f"=AMORLINC(2400,{ISSUE},{MATURITY},300,1,0.15)", 6),
    Probe("AMORLINC", f"=AMORLINC(2400,{ISSUE},{MATURITY},300,1,0.15,1)", 7),
    Probe("BESSELI", "=BESSELI(1.5,1)", 2),
    Probe("BESSELJ", "=BESSELJ(1.9,2)", 2),
    Probe("BESSELK", "=BESSELK(1.5,1)", 2),
    Probe("BESSELY", "=BESSELY(2.5,1)", 2),
    Probe("BIN2DEC", "=BIN2DEC(1100100)", 1),
    Probe("BIN2HEX", "=BIN2HEX(11111011)", 1),
    Probe("BIN2HEX", "=BIN2HEX(11111011,4)", 2),
    Probe("BIN2OCT", "=BIN2OCT(1001)", 1),
    Probe("BIN2OCT", "=BIN2OCT(1001,3)", 2),
    Probe("COMPLEX", "=COMPLEX(3,4)", 2),
    Probe("COMPLEX", '=COMPLEX(3,4,"j")', 3),
    Probe("CONVERT", '=CONVERT(1,"lbm","kg")', 3),
    Probe("COUPDAYBS", f"=COUPDAYBS({SETTLEMENT},{MATURITY},2)", 3),
    Probe("COUPDAYBS", f"=COUPDAYBS({SETTLEMENT},{MATURITY},2,0)", 4),
    Probe("COUPDAYS", f"=COUPDAYS({SETTLEMENT},{MATURITY},2)", 3),
    Probe("COUPDAYS", f"=COUPDAYS({SETTLEMENT},{MATURITY},2,0)", 4),
    Probe("COUPDAYSNC", f"=COUPDAYSNC({SETTLEMENT},{MATURITY},2)", 3),
    Probe("COUPDAYSNC", f"=COUPDAYSNC({SETTLEMENT},{MATURITY},2,0)", 4),
    Probe("COUPNCD", f"=COUPNCD({SETTLEMENT},{MATURITY},2)", 3),
    Probe("COUPNCD", f"=COUPNCD({SETTLEMENT},{MATURITY},2,0)", 4),
    Probe("COUPNUM", f"=COUPNUM({SETTLEMENT},{MATURITY},2)", 3),
    Probe("COUPNUM", f"=COUPNUM({SETTLEMENT},{MATURITY},2,0)", 4),
    Probe("COUPPCD", f"=COUPPCD({SETTLEMENT},{MATURITY},2)", 3),
    Probe("COUPPCD", f"=COUPPCD({SETTLEMENT},{MATURITY},2,0)", 4),
    Probe("CUMIPMT", "=CUMIPMT(0.01,360,125000,13,24,0)", 6),
    Probe("CUMPRINC", "=CUMPRINC(0.01,360,125000,13,24,0)", 6),
    Probe("DBCS", '=DBCS("ABC")', 1),
    Probe("DEC2BIN", "=DEC2BIN(9)", 1),
    Probe("DEC2BIN", "=DEC2BIN(9,4)", 2),
    Probe("DEC2HEX", "=DEC2HEX(100)", 1),
    Probe("DEC2HEX", "=DEC2HEX(100,4)", 2),
    Probe("DEC2OCT", "=DEC2OCT(58)", 1),
    Probe("DEC2OCT", "=DEC2OCT(58,3)", 2),
    Probe("DELTA", "=DELTA(5)", 1),
    Probe("DELTA", "=DELTA(5,4)", 2),
    Probe("DISC", f"=DISC({SETTLEMENT},{MATURITY},97.975,100)", 4),
    Probe("DISC", f"=DISC({SETTLEMENT},{MATURITY},97.975,100,0)", 5),
    Probe("DOLLARDE", "=DOLLARDE(1.02,16)", 2),
    Probe("DOLLARFR", "=DOLLARFR(1.125,16)", 2),
    Probe("DURATION", f"=DURATION({SETTLEMENT},{MATURITY},0.08,0.09,2)", 5),
    Probe("DURATION", f"=DURATION({SETTLEMENT},{MATURITY},0.08,0.09,2,0)", 6),
    Probe("EFFECT", "=EFFECT(0.0525,4)", 2),
    Probe("ERF", "=ERF(0.745)", 1),
    Probe("ERF", "=ERF(0.745,1)", 2),
    Probe("ERFC", "=ERFC(1)", 1),
    Probe("FACTDOUBLE", "=FACTDOUBLE(6)", 1),
    Probe("FVSCHEDULE", "=FVSCHEDULE(1,F1:F3)", 2),
    Probe("GCD", "=GCD(24,36)", 2),
    Probe("GCD", "=GCD(24,36,60)", 3),
    Probe("GESTEP", "=GESTEP(5)", 1),
    Probe("GESTEP", "=GESTEP(5,4)", 2),
    Probe("HEX2BIN", '=HEX2BIN("F")', 1),
    Probe("HEX2BIN", '=HEX2BIN("F",8)', 2),
    Probe("HEX2DEC", '=HEX2DEC("A5")', 1),
    Probe("HEX2OCT", '=HEX2OCT("F")', 1),
    Probe("HEX2OCT", '=HEX2OCT("F",3)', 2),
    Probe("IMABS", '=IMABS("5+12i")', 1),
    Probe("IMAGINARY", '=IMAGINARY("3+4i")', 1),
    Probe("IMARGUMENT", '=IMARGUMENT("3+4i")', 1),
    Probe("IMCONJUGATE", '=IMCONJUGATE("3+4i")', 1),
    Probe("IMCOS", '=IMCOS("1+i")', 1),
    Probe("IMDIV", '=IMDIV("-238+240i","10+24i")', 2),
    Probe("IMEXP", '=IMEXP("1+i")', 1),
    Probe("IMLN", '=IMLN("3+4i")', 1),
    Probe("IMLOG10", '=IMLOG10("3+4i")', 1),
    Probe("IMLOG2", '=IMLOG2("3+4i")', 1),
    Probe("IMPOWER", '=IMPOWER("2+3i",3)', 2),
    Probe("IMPRODUCT", '=IMPRODUCT("3+4i","5-3i")', 2),
    Probe("IMPRODUCT", '=IMPRODUCT("3+4i","5-3i","1+i")', 3),
    Probe("IMREAL", '=IMREAL("6-9i")', 1),
    Probe("IMSIN", '=IMSIN("3+4i")', 1),
    Probe("IMSQRT", '=IMSQRT("1+i")', 1),
    Probe("IMSUB", '=IMSUB("13+4i","5+3i")', 2),
    Probe("IMSUM", '=IMSUM("3+4i","5-3i")', 2),
    Probe("IMSUM", '=IMSUM("3+4i","5-3i","1+i")', 3),
    Probe("INTRATE", f"=INTRATE({SETTLEMENT},{MATURITY},1000000,1014420)", 4),
    Probe("INTRATE", f"=INTRATE({SETTLEMENT},{MATURITY},1000000,1014420,0)", 5),
    Probe("ISEVEN", "=ISEVEN(2)", 1),
    Probe("ISO.CEILING", "=ISO.CEILING(4.3)", 1),
    Probe("ISO.CEILING", "=ISO.CEILING(4.3,1)", 2),
    Probe("ISODD", "=ISODD(3)", 1),
    Probe("LCM", "=LCM(5,2)", 2),
    Probe("LCM", "=LCM(5,2,3)", 3),
    Probe("MDURATION", f"=MDURATION({SETTLEMENT},{MATURITY},0.08,0.09,2)", 5),
    Probe("MDURATION", f"=MDURATION({SETTLEMENT},{MATURITY},0.08,0.09,2,0)", 6),
    Probe("MROUND", "=MROUND(10,3)", 2),
    Probe("MULTINOMIAL", "=MULTINOMIAL(2,3)", 2),
    Probe("MULTINOMIAL", "=MULTINOMIAL(2,3,4)", 3),
    Probe("NOMINAL", "=NOMINAL(0.053543,4)", 2),
    Probe("OCT2BIN", "=OCT2BIN(3)", 1),
    Probe("OCT2BIN", "=OCT2BIN(3,3)", 2),
    Probe("OCT2DEC", "=OCT2DEC(54)", 1),
    Probe("OCT2HEX", "=OCT2HEX(100)", 1),
    Probe("OCT2HEX", "=OCT2HEX(100,4)", 2),
    Probe("ODDFPRICE", f"=ODDFPRICE({SETTLEMENT},{MATURITY},{ISSUE},{FIRST_COUPON},0.0785,0.0625,100,2)", 8),
    Probe("ODDFPRICE", f"=ODDFPRICE({SETTLEMENT},{MATURITY},{ISSUE},{FIRST_COUPON},0.0785,0.0625,100,2,0)", 9),
    Probe("ODDFYIELD", f"=ODDFYIELD({SETTLEMENT},{MATURITY},{ISSUE},{FIRST_COUPON},0.0575,84.5,100,2)", 8),
    Probe("ODDFYIELD", f"=ODDFYIELD({SETTLEMENT},{MATURITY},{ISSUE},{FIRST_COUPON},0.0575,84.5,100,2,0)", 9),
    Probe("ODDLPRICE", f"=ODDLPRICE({SETTLEMENT},{MATURITY},{ISSUE},0.0375,0.0405,100,2)", 7),
    Probe("ODDLPRICE", f"=ODDLPRICE({SETTLEMENT},{MATURITY},{ISSUE},0.0375,0.0405,100,2,0)", 8),
    Probe("ODDLYIELD", f"=ODDLYIELD({SETTLEMENT},{MATURITY},{ISSUE},0.0375,99.875,100,2)", 7),
    Probe("ODDLYIELD", f"=ODDLYIELD({SETTLEMENT},{MATURITY},{ISSUE},0.0375,99.875,100,2,0)", 8),
    Probe("PRICE", f"=PRICE({SETTLEMENT},{MATURITY},0.0575,0.065,100,2)", 6),
    Probe("PRICE", f"=PRICE({SETTLEMENT},{MATURITY},0.0575,0.065,100,2,0)", 7),
    Probe("PRICEDISC", f"=PRICEDISC({SETTLEMENT},{MATURITY},0.0525,100)", 4),
    Probe("PRICEDISC", f"=PRICEDISC({SETTLEMENT},{MATURITY},0.0525,100,0)", 5),
    Probe("PRICEMAT", f"=PRICEMAT({SETTLEMENT},{MATURITY},{ISSUE},0.061,0.061)", 5),
    Probe("PRICEMAT", f"=PRICEMAT({SETTLEMENT},{MATURITY},{ISSUE},0.061,0.061,0)", 6),
    Probe("QUOTIENT", "=QUOTIENT(5,2)", 2),
    Probe("RANDBETWEEN", "=RANDBETWEEN(1,100)", 2),
    Probe("RECEIVED", f"=RECEIVED({SETTLEMENT},{MATURITY},1000000,0.0575)", 4),
    Probe("RECEIVED", f"=RECEIVED({SETTLEMENT},{MATURITY},1000000,0.0575,0)", 5),
    Probe("SERIESSUM", "=SERIESSUM(1,0,1,F1:F3)", 4),
    Probe("SQRTPI", "=SQRTPI(1)", 1),
    Probe("TBILLEQ", f"=TBILLEQ({SETTLEMENT},{FIRST_COUPON},0.0914)", 3),
    Probe("TBILLPRICE", f"=TBILLPRICE({SETTLEMENT},{FIRST_COUPON},0.09)", 3),
    Probe("TBILLYIELD", f"=TBILLYIELD({SETTLEMENT},{FIRST_COUPON},98.45)", 3),
    Probe("WEEKNUM", f"=WEEKNUM({SETTLEMENT})", 1),
    Probe("WEEKNUM", f"=WEEKNUM({SETTLEMENT},1)", 2),
    Probe("XIRR", "=XIRR(D1:D2,E1:E2)", 2),
    Probe("XIRR", "=XIRR(D1:D2,E1:E2,0.1)", 3),
    Probe("XNPV", "=XNPV(0.09,D1:D2,E1:E2)", 3),
    Probe("YEARFRAC", f"=YEARFRAC({ISSUE},{MATURITY})", 2),
    Probe("YEARFRAC", f"=YEARFRAC({ISSUE},{MATURITY},0)", 3),
    Probe("YIELD", f"=YIELD({SETTLEMENT},{MATURITY},0.0575,95.04287,100,2)", 6),
    Probe("YIELD", f"=YIELD({SETTLEMENT},{MATURITY},0.0575,95.04287,100,2,0)", 7),
    Probe("YIELDDISC", f"=YIELDDISC({SETTLEMENT},{MATURITY},99.795,100)", 4),
    Probe("YIELDDISC", f"=YIELDDISC({SETTLEMENT},{MATURITY},99.795,100,0)", 5),
    Probe("YIELDMAT", f"=YIELDMAT({SETTLEMENT},{MATURITY},{ISSUE},0.061,99.98)", 5),
    Probe("YIELDMAT", f"=YIELDMAT({SETTLEMENT},{MATURITY},{ISSUE},0.061,99.98,0)", 6),
]

# Column A carries the function name as a label, column B the probe formula.
LABEL_COL = 0
FORMULA_COL = 1
PROBE_ROW_0 = 4  # zero-based; rows 0-3 hold the helper block.

# `PtgFuncVar` id 255 is the future-function sentinel: the real callee was
# pushed as a `PtgName` operand, so no classic id exists for it.
FUTURE_FUNC_SENTINEL = 255

BRT_ROW_HDR = 0
BRT_FMLA_STRING = 8
BRT_FMLA_NUM = 9
BRT_FMLA_BOOL = 10
BRT_FMLA_ERROR = 11
BRT_FMLA_TYPES = {BRT_FMLA_STRING, BRT_FMLA_NUM, BRT_FMLA_BOOL, BRT_FMLA_ERROR}


# ---------------------------------------------------------------------------
# build
# ---------------------------------------------------------------------------


def build(out_path: str, *, visible: bool) -> None:
    """Types every probe into a fresh workbook and saves it as `.xlsb`."""

    sys.path.insert(0, os.path.join(REPO_ROOT, "tools", "oracle"))

    import xlwings as xw  # noqa: PLC0415
    from drivers.macos_excel import _assign_formula  # noqa: PLC0415

    out_path = os.path.abspath(out_path)
    if os.path.exists(out_path):
        os.remove(out_path)

    app = xw.App(visible=visible, add_book=False)
    app.display_alerts = False
    rejected: List[Tuple[str, str, str]] = []
    try:
        app.calculation = "manual"
        book = app.books.add()
        try:
            sheet = book.sheets[0]
            for addr, value in HELPERS.items():
                sheet.range(addr).value = value
            for i, probe in enumerate(PROBES):
                row = PROBE_ROW_0 + i + 1  # xlwings rows are 1-based
                sheet.range((row, LABEL_COL + 1)).value = probe.name
                cell = sheet.range((row, FORMULA_COL + 1))
                try:
                    cell.number_format = "General"
                    _assign_formula(cell, probe.formula, context=f"{probe.name} probe")
                except Exception as exc:  # noqa: BLE001 - reported, not raised
                    rejected.append((probe.name, probe.formula, f"{type(exc).__name__}: {exc}"[:200]))
                    cell.clear_contents()
            app.calculate()
            book.save(out_path)
        finally:
            try:
                book.close()
            except Exception:  # noqa: BLE001
                pass
    finally:
        try:
            app.quit()
        except Exception:  # noqa: BLE001
            pass

    if rejected:
        print("Excel rejected these probes (no id harvested):", file=sys.stderr)
        for name, formula, why in rejected:
            print(f"  {name} {formula}: {why}", file=sys.stderr)
    print(f"wrote {out_path}")


# ---------------------------------------------------------------------------
# decode
# ---------------------------------------------------------------------------


class Record:
    """One framed MS-XLSB record with the payload's absolute file offset."""

    __slots__ = ("type", "offset", "size")

    def __init__(self, type_: int, offset: int, size: int) -> None:
        self.type = type_
        self.offset = offset
        self.size = size


def iter_records(buf: bytes) -> Iterator[Record]:
    """Yields every record in a `.bin` part.

    Framing is a variable-length record type (up to 2 bytes) followed by a
    variable-length payload size (up to 4 bytes), each byte contributing 7
    bits little-endian with the high bit as the continuation flag -- the same
    shape `src/io/xlsb/record.cpp` reads.
    """

    i = 0
    n = len(buf)
    while i < n:
        b = buf[i]
        i += 1
        rec_type = b & 0x7F
        if b & 0x80:
            b = buf[i]
            i += 1
            rec_type |= (b & 0x7F) << 7
        size = 0
        for k in range(4):
            b = buf[i]
            i += 1
            size |= (b & 0x7F) << (7 * k)
            if not (b & 0x80):
                break
        yield Record(rec_type, i, size)
        i += size


def _u16(buf: bytes, at: int) -> int:
    return buf[at] | (buf[at + 1] << 8)


def _u32(buf: bytes, at: int) -> int:
    return buf[at] | (buf[at + 1] << 8) | (buf[at + 2] << 16) | (buf[at + 3] << 24)


def _rgce_span(buf: bytes, rec: Record) -> Optional[Tuple[int, int]]:
    """Returns the `(offset, length)` of a formula record's Ptg stream.

    Every `BrtFmla*` record is an 8-byte cell header, then the cached result
    (8 bytes for Num, an XLWideString for String, 1 byte for Bool / Error),
    then a `u16` flag word, then `CellParsedFormula` (`u32 cce` + `cce`
    bytes). Mirrors the reader in `src/io/xlsb/reader.cpp`.
    """

    p = rec.offset + 8  # cell header: u32 col + u24 iStyleRef + u8 fPhShow
    end = rec.offset + rec.size
    if rec.type == BRT_FMLA_NUM:
        p += 8
    elif rec.type == BRT_FMLA_STRING:
        if p + 4 > end:
            return None
        cch = _u32(buf, p)
        p += 4 + 2 * cch
    else:  # Bool / Error carry a single result byte.
        p += 1
    p += 2  # grbitFlags
    if p + 4 > end:
        return None
    cce = _u32(buf, p)
    p += 4
    if cce == 0 or p + cce > end:
        return None
    return p, cce


# Fixed on-the-wire length of every Ptg the probe formulas can produce,
# keyed by the token's class-independent low 5 bits. Lengths include the
# opcode byte and mirror the reader's own consumption in
# `src/io/xlsb/ptg_reader.cpp`; `PtgStr` (0x17) and `PtgAttr` (0x19) are
# variable-length and handled separately.
_PTG_FIXED_LEN = {
    0x01: 5,  # Exp
    0x02: 5,  # Tbl
    0x03: 1,  # Add
    0x04: 1,  # Sub
    0x05: 1,  # Mul
    0x06: 1,  # Div
    0x07: 1,  # Power
    0x08: 1,  # Concat
    0x09: 1,  # Lt
    0x0A: 1,  # Le
    0x0B: 1,  # Eq
    0x0C: 1,  # Ge
    0x0D: 1,  # Gt
    0x0E: 1,  # Ne
    0x0F: 1,  # Isect
    0x10: 1,  # Union
    0x11: 1,  # Range
    0x12: 1,  # Uplus
    0x13: 1,  # Uminus
    0x14: 1,  # Percent
    0x15: 1,  # Paren
    0x16: 1,  # MissArg
    0x1C: 2,  # Err
    0x1D: 2,  # Bool
    0x1E: 3,  # Int
    0x1F: 9,  # Num
}
# Class-marked tokens (reference-class base bytes 0x20..0x3D). The value and
# array classes set 0x20 / 0x40 over the base, so masking the class bits off
# folds all three spellings back onto one row.
_PTG_CLASSED_LEN = {
    0x20: 15,  # Array (opcode + 14 reserved bytes; data lives in rgcb)
    0x21: 3,  # Func: u16 id
    0x22: 4,  # FuncVar: u8 cparams + u16 id
    0x23: 5,  # Name: u32 ilbl
    0x24: 7,  # Ref: u32 row + u16 col
    0x25: 13,  # Area: u32 row1 + u32 row2 + u16 col1 + u16 col2
    0x26: 7,  # MemArea: u32 reserved + u16 cce
    0x27: 7,  # MemErr
    0x28: 7,  # MemNoMem
    0x29: 3,  # MemFunc: u16 cce
    0x2A: 7,  # RefErr
    0x2B: 13,  # AreaErr
    0x2C: 7,  # RefN
    0x2D: 13,  # AreaN
    0x39: 7,  # NameX: u16 ixti + u32 ilbl
    0x3A: 9,  # Ref3d: u16 ixti + u32 row + u16 col
    0x3B: 15,  # Area3d: u16 ixti + RgceArea
    0x3C: 9,  # RefErr3d
    0x3D: 15,  # AreaErr3d
}

_PTG_ATTR = 0x19
_PTG_STR = 0x17
_PTG_ATTR_CHOOSE = 0x04
_PTG_FUNC_BASE = 0x21
_PTG_FUNC_VAR_BASE = 0x22


def _ptg_base_byte(op: int) -> int:
    """Folds a class-marked Ptg opcode onto its reference-class base byte."""

    return (op & 0x1F) | 0x20


def _ptg_token_length(buf: bytes, at: int, end: int) -> Optional[int]:
    """Returns the byte length of the Ptg token starting at `at`, or None."""

    op = buf[at]
    if op == _PTG_STR:
        if at + 3 > end:
            return None
        return 3 + 2 * _u16(buf, at + 1)
    if op == _PTG_ATTR:
        if at + 2 > end:
            return None
        if buf[at + 1] == _PTG_ATTR_CHOOSE:
            if at + 4 > end:
                return None
            return 4 + 4 * (_u16(buf, at + 2) + 1)
        return 4
    if op < 0x20:
        return _PTG_FIXED_LEN.get(op)
    return _PTG_CLASSED_LEN.get(_ptg_base_byte(op))


def _final_func_token(buf: bytes, rgce_at: int, cce: int) -> Optional[Dict[str, object]]:
    """Parses a whole Ptg stream and returns its closing call token.

    A probe formula is a single call, so its postfix stream ends with the
    call token. The stream is walked from the front and the parse must land
    exactly on `rgce_at + cce`; a stream that desynchronises (or ends on
    anything other than `PtgFunc` / `PtgFuncVar`) yields None rather than a
    guess read backwards off the tail.
    """

    end = rgce_at + cce
    at = rgce_at
    last: Optional[Dict[str, object]] = None
    while at < end:
        length = _ptg_token_length(buf, at, end)
        if length is None or length <= 0 or at + length > end:
            return None
        op = buf[at]
        base = _ptg_base_byte(op) if op >= 0x20 else op
        if base == _PTG_FUNC_BASE:
            last = {
                "token": "PtgFunc",
                "ptg": op,
                "id": _u16(buf, at + 1),
                "cparams": None,
                "token_offset": at,
            }
        elif base == _PTG_FUNC_VAR_BASE:
            last = {
                "token": "PtgFuncVar",
                "ptg": op,
                "id": _u16(buf, at + 2),
                "cparams": buf[at + 1],
                "token_offset": at,
            }
        else:
            last = None
        at += length
    if at != end:
        return None
    return last


def decode(xlsb_path: str, part: str = "xl/worksheets/sheet1.bin") -> List[Dict[str, object]]:
    """Decodes the trailing call token of every formula cell in `part`."""

    with zipfile.ZipFile(xlsb_path) as zf:
        buf = zf.read(part)

    out: List[Dict[str, object]] = []
    row = -1
    for rec in iter_records(buf):
        if rec.type == BRT_ROW_HDR:
            row = _u32(buf, rec.offset)
            continue
        if rec.type not in BRT_FMLA_TYPES:
            continue
        col = _u32(buf, rec.offset)
        span = _rgce_span(buf, rec)
        if span is None:
            continue
        token = _final_func_token(buf, span[0], span[1])
        if token is None:
            continue
        entry: Dict[str, object] = {"row": row, "col": col, "part": part}
        entry.update(token)
        out.append(entry)
    return out


def _a1(row: int, col: int) -> str:
    letters = ""
    c = col + 1
    while c:
        c, rem = divmod(c - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return f"{letters}{row + 1}"


def harvest(xlsb_path: str) -> Tuple[List[Dict[str, object]], List[Tuple[str, str]]]:
    """Joins the decoded tokens back onto the probe list by cell row.

    Returns `(resolved, unresolved)` where each `resolved` row carries the id
    plus every cell and byte offset it was observed at, and `unresolved`
    pairs a function name with the reason no usable id was obtained.
    Observations of the same function must agree on the id and the token
    kind, and a fixed-arity function must not appear at two arities.
    """

    by_cell = {(e["row"], e["col"]): e for e in decode(xlsb_path)}
    observations: Dict[str, List[Dict[str, object]]] = {}
    unresolved: List[Tuple[str, str]] = []
    order: List[str] = []
    for i, probe in enumerate(PROBES):
        row = PROBE_ROW_0 + i
        if probe.name not in observations:
            observations[probe.name] = []
            order.append(probe.name)
        entry = by_cell.get((row, FORMULA_COL))
        if entry is None:
            continue
        observations[probe.name].append(
            {
                "formula": probe.formula,
                "arity": probe.arity,
                "cell": _a1(row, FORMULA_COL),
                "id": entry["id"],
                "token": entry["token"],
                "ptg": entry["ptg"],
                "cparams": entry["cparams"],
                "byte_offset": entry["token_offset"],
            }
        )

    resolved: List[Dict[str, object]] = []
    for name in order:
        seen = observations[name]
        if not seen:
            unresolved.append((name, "no decodable formula cell in the saved workbook"))
            continue
        ids = {int(o["id"]) for o in seen}
        if len(ids) != 1:
            unresolved.append((name, f"probes disagree on the id: {sorted(ids)}"))
            continue
        func_id = ids.pop()
        if func_id == FUTURE_FUNC_SENTINEL:
            unresolved.append((name, "Excel encoded it as a future function (PtgFuncVar id 255)"))
            continue
        kinds = {str(o["token"]) for o in seen}
        if len(kinds) != 1:
            unresolved.append((name, f"probes disagree on the token kind: {sorted(kinds)}"))
            continue
        kind = kinds.pop()
        arities = sorted({int(o["arity"]) for o in seen})
        if kind == "PtgFunc" and len(arities) != 1:
            unresolved.append((name, f"fixed-arity token observed at several arities: {arities}"))
            continue
        for obs in seen:
            if obs["cparams"] is not None and int(obs["cparams"]) != int(obs["arity"]):
                unresolved.append((name, f"cparams {obs['cparams']} does not match probe arity {obs['arity']}"))
                break
        else:
            resolved.append(
                {
                    "name": name,
                    "id": func_id,
                    "token": kind,
                    "variadic": kind == "PtgFuncVar",
                    "arities": arities,
                    "observations": seen,
                }
            )
    return resolved, unresolved


# ---------------------------------------------------------------------------
# emit
# ---------------------------------------------------------------------------

_VARIADIC_SENTINEL = "kVariadicMax"


def registry_arity() -> Dict[str, Tuple[int, Optional[int]]]:
    """Reads `(min_args, max_args)` off the engine's own builtin registrations.

    The XLSB table's `arg_min` / `arg_max` are the Excel function's argument
    bounds; the engine's `FunctionDef` rows are the repo's existing statement
    of exactly that, so they are the source rather than a second hand-written
    list. `max_args == None` means variadic (`kVariadic`). Two registration
    shapes are in use: the flat `{"NAME", impl, min, max, ...}` rows and the
    nested `{"NAME", {&Impl_, min, max}}` family tables.
    """

    flat = re.compile(r'\{"([A-Z0-9_.]+)",\s*(\d+)u?,\s*(kVariadic|\d+)u?,')
    nested = re.compile(r'\{"([A-Z0-9_.]+)",\s*\{&\w+,\s*(\d+)u?,\s*(kVariadic|\d+)u?\}')
    out: Dict[str, Tuple[int, Optional[int]]] = {}
    eval_dir = os.path.join(REPO_ROOT, "src", "eval")
    for dirpath, _dirnames, filenames in os.walk(eval_dir):
        for fn in filenames:
            if not fn.endswith((".cpp", ".h")):
                continue
            with open(os.path.join(dirpath, fn), encoding="utf-8") as fh:
                text = fh.read()
            for pattern in (flat, nested):
                for name, lo, hi in pattern.findall(text):
                    out.setdefault(name, (int(lo), None if hi == "kVariadic" else int(hi)))
    return out


def emit_rows(resolved: List[Dict[str, object]]) -> List[str]:
    """Formats decoded ids as `func_id_table.cpp` initialiser rows.

    A fixed-arity (`PtgFunc`) row's `arg_min` is the probe arity Excel
    accepted, because that is exactly what the reader pops for such a call.
    A variadic row takes its bounds from the engine's own registration where
    one exists, and otherwise from the arities the probes exercised.
    """

    arity = registry_arity()
    rows: List[str] = []
    for entry in sorted(resolved, key=lambda e: e["id"]):
        name = str(entry["name"])
        arities = [int(a) for a in entry["arities"]]
        if entry["variadic"]:
            lo, hi = arity.get(name, (min(arities), max(arities)))
            hi_text = _VARIADIC_SENTINEL if hi is None else str(hi)
        else:
            lo, hi_text = arities[0], str(arities[0])
        rows.append(f'    {{{entry["id"]}, "{name}", {lo}, {hi_text}, {str(bool(entry["variadic"])).lower()}}},')
    return rows


def current_table() -> Dict[int, str]:
    """Parses the ids already present in `src/io/xlsb/func_id_table.cpp`."""

    path = os.path.join(REPO_ROOT, "src", "io", "xlsb", "func_id_table.cpp")
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    return {int(i): n for i, n in re.findall(r'\{(\d+),\s*"([A-Z0-9_.]+)"', text)}


def _collides(entry: Dict[str, object], table: Dict[int, str]) -> bool:
    """True when `entry` cannot be added without displacing an existing row."""

    held = table.get(int(entry["id"]))
    return (held is not None and held != str(entry["name"])) or (
        str(entry["name"]) in table.values() and held != str(entry["name"])
    )


def collisions(resolved: List[Dict[str, object]]) -> List[str]:
    """Reports decoded ids that clash with the table or with each other."""

    table = current_table()
    out: List[str] = []
    seen: Dict[int, str] = {}
    for entry in resolved:
        func_id = int(entry["id"])
        name = str(entry["name"])
        held = table.get(func_id)
        if held is not None and held != name:
            out.append(f"id {func_id} decoded for {name} is already mapped to {held}")
        if name in table.values() and table.get(func_id) != name:
            out.append(f"{name} is already in the table under a different id")
        if func_id in seen and seen[func_id] != name:
            out.append(f"id {func_id} decoded for both {seen[func_id]} and {name}")
        seen[func_id] = name
    return out


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)

    p_build = sub.add_parser("build", help="drive Excel and save the probe workbook as .xlsb")
    p_build.add_argument("--out", required=True)
    p_build.add_argument("--visible", action="store_true", help="show the Excel window while driving it")

    p_decode = sub.add_parser("decode", help="decode ids out of a probe .xlsb")
    p_decode.add_argument("xlsb")
    p_decode.add_argument("--json", action="store_true")

    p_emit = sub.add_parser("emit", help="print func_id_table.cpp rows for a probe .xlsb")
    p_emit.add_argument("xlsb")

    args = parser.parse_args(argv)

    if args.command == "build":
        build(args.out, visible=args.visible)
        return 0

    resolved, unresolved = harvest(args.xlsb)
    clashes = collisions(resolved)
    if args.command == "decode":
        if getattr(args, "json", False):
            print(json.dumps({"resolved": resolved, "unresolved": unresolved, "collisions": clashes}, indent=2))
        else:
            print(f"{'function':<14} {'id':>5}  {'token':<11} {'arities':<8} cells (byte offset)")
            for entry in resolved:
                cells = ", ".join(f"{o['cell']}@{o['byte_offset']}" for o in entry["observations"])
                arities = ",".join(str(a) for a in entry["arities"])
                print(f"{entry['name']:<14} {entry['id']:>5}  {entry['token']:<11} {arities:<8} {cells}")
            for name, why in unresolved:
                print(f"UNRESOLVED {name}: {why}")
            for clash in clashes:
                print(f"COLLISION {clash}")
        return 0

    # A colliding name is withheld from the emitted rows: overwriting an id
    # the table already maps to a different name would make the reader
    # resolve that id to the wrong function.
    colliding = {str(e["name"]) for e in resolved if _collides(e, current_table())}
    for row in emit_rows([e for e in resolved if str(e["name"]) not in colliding]):
        print(row)
    for name, why in unresolved:
        print(f"// unresolved: {name} -- {why}")
    for clash in clashes:
        print(f"// collision: {clash}")
    return 1 if clashes else 0


if __name__ == "__main__":
    raise SystemExit(main())
