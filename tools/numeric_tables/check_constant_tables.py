#!/usr/bin/env python3
"""Recompute the engine's embedded numeric tables from their definitions.

Every checked table or scalar is a mathematically defined quantity converted to
a double by a stated rule. The definition is evaluated here in exact arithmetic
(Fraction for rational values, 70-digit Decimal for logarithms and pi) and the
literals found in the C++ source must be bit-identical doubles.

Conversion rules:
  nearest    - correctly rounded to the nearest double.
  trunc(N)   - exact value truncated toward zero to N significant decimal
               digits, then parsed as a double.

Usage: check_constant_tables.py [--root DIR] [--table NAME ...] [--list]
Exit status is non-zero if any table is missing or differs.
"""

from __future__ import annotations

import argparse
import re
import struct
import sys
from dataclasses import dataclass
from decimal import ROUND_DOWN, Decimal, localcontext
from fractions import Fraction
from math import comb, factorial
from pathlib import Path
from typing import Callable, List, Optional, Sequence, Tuple

PREC = 70  # working precision (decimal digits) for logarithms, pi, roots

# ---------------------------------------------------------------------------
# Exact arithmetic
# ---------------------------------------------------------------------------


def bernoulli(n_max: int) -> List[Fraction]:
    """Bernoulli numbers B_0..B_{n_max} (B_1 = -1/2)."""
    b = [Fraction(1)]
    for m in range(1, n_max + 1):
        s = sum(comb(m + 1, k) * b[k] for k in range(m))
        b.append(-s / (m + 1))
    return b


def dec_pi() -> Decimal:
    """pi = 16 atan(1/5) - 4 atan(1/239) (Machin), at PREC digits."""

    def atan_inv(n: int) -> Decimal:
        x = Decimal(1) / n
        x2 = x * x
        term, total, k = x, x, 1
        eps = Decimal(10) ** -(PREC + 5)
        while abs(term) > eps:
            term = -term * x2
            k += 2
            total += term / k
        return total

    with localcontext() as ctx:
        ctx.prec = PREC + 10
        return 16 * atan_inv(5) - 4 * atan_inv(239)


def dec_ln(x) -> Decimal:
    with localcontext() as ctx:
        ctx.prec = PREC + 10
        return Decimal(x).ln()


def dec_sqrt(x) -> Decimal:
    with localcontext() as ctx:
        ctx.prec = PREC + 10
        return Decimal(x).sqrt()


def dec_euler_gamma() -> Decimal:
    """gamma = H_n - ln n - 1/(2n) + sum_k B_2k / (2k n^2k), n = 1000."""
    n = 1000
    h = sum(Fraction(1, i) for i in range(1, n + 1))
    b = bernoulli(40)
    corr = sum(b[2 * k] / (2 * k * Fraction(n) ** (2 * k)) for k in range(1, 21))
    exact = h - Fraction(1, 2 * n) + corr
    with localcontext() as ctx:
        ctx.prec = PREC + 10
        return Decimal(exact.numerator) / Decimal(exact.denominator) - Decimal(n).ln()


def rn(v) -> float:
    """Correctly rounded double of a Fraction, Decimal, int or float."""
    if isinstance(v, Fraction):
        return float(v)  # int / int true division is correctly rounded
    if isinstance(v, Decimal):
        return float(str(v))
    return float(v)


def trunc_sig(v: Decimal, digits: int) -> float:
    """Truncate toward zero to `digits` significant digits, then parse."""
    if v == 0:
        return 0.0
    with localcontext() as ctx:
        ctx.prec = PREC + 10
        quantum = Decimal(1).scaleb(v.adjusted() - digits + 1)
        return float(str(v.quantize(quantum, rounding=ROUND_DOWN)))


# --- truncated power series over Fraction ----------------------------------


def ps_mul(a: Sequence[Fraction], b: Sequence[Fraction], n: int) -> List[Fraction]:
    out = [Fraction(0)] * n
    for i, ai in enumerate(a[:n]):
        if ai == 0:
            continue
        for j, bj in enumerate(b[: n - i]):
            out[i + j] += ai * bj
    return out


def ps_inv(a: Sequence[Fraction], n: int) -> List[Fraction]:
    """1 / a for a[0] != 0."""
    out = [Fraction(0)] * n
    out[0] = 1 / a[0]
    for k in range(1, n):
        s = sum(a[j] * out[k - j] for j in range(1, min(k, len(a) - 1) + 1))
        out[k] = -s / a[0]
    return out


def ps_sqrt(a: Sequence[Fraction], n: int) -> List[Fraction]:
    """sqrt(a) for a[0] == 1."""
    out = [Fraction(0)] * n
    out[0] = Fraction(1)
    for k in range(1, n):
        s = sum(out[j] * out[k - j] for j in range(1, k))
        ak = a[k] if k < len(a) else Fraction(0)
        out[k] = (ak - s) / 2
    return out


# ---------------------------------------------------------------------------
# Table definitions
# ---------------------------------------------------------------------------


def gamma_stirling() -> List[float]:
    """2 c_n, n = 15..0, where Gamma(x) ~ sqrt(2 pi / x) (x/e)^x sum c_n x^-n.

    sum c_n t^n = exp(sum_k B_2k / (2k (2k-1)) t^(2k-1)).
    """
    n = 16
    b = bernoulli(2 * n)
    ell = [Fraction(0)] * n
    for k in range(1, n):
        if 2 * k - 1 < n:
            ell[2 * k - 1] = b[2 * k] / (2 * k * (2 * k - 1))
    e = [Fraction(1)] + [Fraction(0)] * (n - 1)
    for m in range(1, n):
        e[m] = sum(k * ell[k] * e[m - k] for k in range(1, m + 1)) / m
    return [rn(2 * c) for c in reversed(e)]


def recip_factorial() -> List[float]:
    """1/k! for k = 21..0."""
    f = [1]
    for k in range(1, 22):
        f.append(f[-1] * k)
    return [rn(Fraction(1, f[k])) for k in range(21, -1, -1)]


def half_int_series() -> List[float]:
    """2^k / (2k-1)!! for k = 21..1, then 0."""

    def dfact(m: int) -> int:
        r = 1
        while m > 1:
            r *= m
            m -= 2
        return r

    out = [rn(Fraction(2**k, dfact(2 * k - 1))) for k in range(21, 0, -1)]
    return out + [0.0]


def stirling_poly(digits: int) -> List[float]:
    """B_2k / (2k (2k-1)) for k = 10..1; k = 10 truncated to `digits` digits."""
    b = bernoulli(20)
    vals = [b[2 * k] / (2 * k * (2 * k - 1)) for k in range(10, 0, -1)]
    first = vals[0]
    dec = Decimal(first.numerator) / Decimal(first.denominator)
    return [trunc_sig(dec, digits)] + [rn(v) for v in vals[1:]]


def stirling_small(digits: int) -> List[float]:
    """stirlerr(n) = ln n! - (n + 1/2) ln n + n - ln sqrt(2 pi), n = 1..6.

    Each value truncated to `digits` significant digits.
    """
    out = []
    ln_2pi_half = dec_ln(2 * dec_pi()) / 2
    fact = 1
    with localcontext() as ctx:
        ctx.prec = PREC + 10
        for n in range(1, 7):
            fact *= n
            v = dec_ln(fact) - (Decimal(n) + Decimal("0.5")) * dec_ln(n) + n - ln_2pi_half
            out.append(trunc_sig(v, digits))
    return out


def temme_seed() -> List[float]:
    """Coefficients of eta^k, k = 29..0, in c0(eta) = 1/(lambda-1) - 1/eta.

    eta^2 / 2 = lambda - 1 - ln(lambda), sign(eta) = sign(lambda - 1).
    With u = lambda - 1: eta = u sqrt(h(u)), h = 2 (u - ln(1+u)) / u^2; u(eta)
    follows by Lagrange inversion and c0 = (1 / (u/eta) - 1) / eta.
    """
    n = 31
    # h(u) = 2 sum_{j>=2} (-1)^j u^(j-2) / j
    h = [Fraction(2 * (-1) ** j, j) for j in range(2, n + 3)]
    g = ps_sqrt(h, n + 1)  # eta = u g(u)
    r = ps_inv(g, n + 1)  # u / eta(u) as a series in u
    # [eta^m] u = (1/m) [u^(m-1)] r^m
    u_coef = [Fraction(0), Fraction(1)]
    power = r[:]
    for m in range(2, n + 1):
        power = ps_mul(power, r, n + 1)
        u_coef.append(power[m - 1] / m)
    # u/eta as series in eta: u_coef shifted down one degree.
    ratio = [u_coef[m + 1] for m in range(n)]
    inv_ratio = ps_inv(ratio, n)  # eta / u
    # 1/u - 1/eta = (eta/u - 1) / eta: the coefficient of eta^k is inv_ratio[k + 1].
    return [rn(inv_ratio[k + 1]) for k in range(29, -1, -1)]


def small_b_poly(i: int) -> List[float]:
    """Coefficients in b, highest degree first, of a_j(b) with j = 9 - i.

    a_j(b) = [s^(2j)] (sinh(s/2) / (s/2))^(b-1)
           = [t^j] exp((b-1) L(t)),  L = sum_k B_2k t^k / (2k (2k)!),  t = s^2.
    """
    n = 10
    j = 9 - i
    b = bernoulli(2 * n)
    ell = [Fraction(0)] + [b[2 * k] / (2 * k * factorial(2 * k)) for k in range(1, n)]
    # [t^j] L^m / m! are the coefficients of a_j in powers of beta = b - 1.
    power = [Fraction(1)] + [Fraction(0)] * (n - 1)
    beta = []
    for m in range(j + 1):
        beta.append(power[j] / factorial(m))
        power = ps_mul(power, ell, n)
    asc = [Fraction(0)] * (j + 1)
    for m, cm in enumerate(beta):
        for r in range(m + 1):
            asc[r] += cm * comb(m, r) * (-1) ** (m - r)
    return [rn(asc[d]) for d in range(j, -1, -1)]


# ---------------------------------------------------------------------------
# Source parsing
# ---------------------------------------------------------------------------

_NUM = r"[+-]?(?:0[xX][0-9a-fA-F]*\.?[0-9a-fA-F]*[pP][+-]?\d+|(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)"
_NUM_RE = re.compile(_NUM)


def strip_comments(text: str) -> str:
    text = re.sub(r"/\*.*?\*/", " ", text, flags=re.S)
    return re.sub(r"//[^\n]*", "", text)


def parse_literal(tok: str) -> float:
    tok = tok.strip()
    if not _NUM_RE.fullmatch(tok):
        raise ValueError(f"not a numeric literal: {tok!r}")
    return float.fromhex(tok) if "0x" in tok.lower() else float(tok)


def find_array(text: str, symbol: str) -> Optional[List[float]]:
    m = re.search(r"\b" + re.escape(symbol) + r"\s*\[\s*\w*\s*\]\s*=\s*\{(.*?)\}\s*;", text, flags=re.S)
    if not m:
        return None
    items = [t for t in m.group(1).split(",") if t.strip()]
    return [parse_literal(t) for t in items]


def find_scalar(text: str, symbol: str) -> Optional[float]:
    m = re.search(r"\bconstexpr\s+double\s+" + re.escape(symbol) + r"\s*=\s*(" + _NUM + r")\s*;", text)
    return parse_literal(m.group(1)) if m else None


def find_inline(text: str, literal: str) -> Optional[float]:
    return parse_literal(literal) if re.search(r"(?<![\w.])" + re.escape(literal) + r"(?![\w.])", text) else None


# ---------------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------------


def _ordered(x: float) -> int:
    (i,) = struct.unpack("<q", struct.pack("<d", x))
    return i if i >= 0 else -(i & 0x7FFFFFFFFFFFFFFF)


def ulp_diff(a: float, b: float) -> int:
    return abs(_ordered(a) - _ordered(b))


@dataclass
class Check:
    name: str
    file: str
    symbol: str
    kind: str  # "array" | "scalar" | "inline"
    expected: Callable[[], List[float]]
    rule: str


def _scalar(name, file, symbol, fn, rule="nearest") -> Check:
    return Check(name, file, symbol, "scalar", lambda: [fn()], rule)


def build_checks() -> List[Check]:
    pi = dec_pi
    gam, beta, special = "stats/incomplete_gamma.cpp", "stats/incomplete_beta.cpp", "stats/special_functions.cpp"
    dist = "builtins/distributions.cpp"
    two_pi = lambda: 2 * pi()  # noqa: E731
    ln_sqrt_2pi = lambda: rn(dec_ln(two_pi()) / 2)  # noqa: E731
    inv_sqrt_pi = lambda: rn(1 / dec_sqrt(pi()))  # noqa: E731
    checks = [
        Check(
            "kGammaStirling",
            dist,
            "kGammaStirling",
            "array",
            gamma_stirling,
            "2 c_n, n=15..0 (Stirling series of Gamma); nearest",
        ),
        Check("kRecipFactorial", gam, "kRecipFactorial", "array", recip_factorial, "1/k!, k=21..0; nearest"),
        Check(
            "kHalfIntSeries", gam, "kHalfIntSeries", "array", half_int_series, "2^k/(2k-1)!!, k=21..1, then 0; nearest"
        ),
        Check(
            "kStirlingPoly",
            gam,
            "kStirlingPoly",
            "array",
            lambda: stirling_poly(STIRLING_POLY_DIGITS),
            f"B_2k/(2k(2k-1)), k=10..1; k=10 trunc({STIRLING_POLY_DIGITS}), rest nearest",
        ),
        Check(
            "kStirlingSmall",
            gam,
            "kStirlingSmall",
            "array",
            lambda: stirling_small(STIRLING_SMALL_DIGITS),
            f"stirlerr(n), n=1..6; trunc({STIRLING_SMALL_DIGITS})",
        ),
        Check("kTemmeSeed", gam, "kTemmeSeed", "array", temme_seed, "[eta^k] (1/(lambda-1) - 1/eta), k=29..0; nearest"),
        _scalar("kInvSqrtPi (gamma)", gam, "kInvSqrtPi", inv_sqrt_pi),
        _scalar("kPhi0", gam, "kPhi0", lambda: rn(1 / dec_sqrt(two_pi()))),
        _scalar("kTwoPi", gam, "kTwoPi", lambda: rn(two_pi())),
        _scalar("kNegLn2", gam, "kNegLn2", lambda: rn(-dec_ln(2))),
        _scalar("kCfTol", gam, "kCfTol", lambda: 1e-16, "nearest of 1e-16"),
        _scalar("kLnSqrt2Pi (gamma)", gam, "kLnSqrt2Pi", ln_sqrt_2pi),
        _scalar("kEps (gamma)", gam, "kEps", lambda: 2.0**-52),
        _scalar("kLnSqrt2Pi (beta)", beta, "kLnSqrt2Pi", ln_sqrt_2pi),
        _scalar("kEps (beta)", beta, "kEps", lambda: 2.0**-52),
        _scalar("kLog07", beta, "kLog07", lambda: rn(dec_ln(Decimal("0.7")))),
        _scalar("kLog09", beta, "kLog09", lambda: rn(dec_ln(Decimal("0.9")))),
        _scalar("kRescale", beta, "kRescale", lambda: rn(970 * dec_ln(2))),
        _scalar("kTwoOverSqrtPi", beta, "kTwoOverSqrtPi", lambda: rn(2 / dec_sqrt(pi()))),
        _scalar("kInvSqrtPi (beta)", beta, "kInvSqrtPi", inv_sqrt_pi),
        _scalar("kQuarterSqrtPi", beta, "kQuarterSqrtPi", lambda: rn(dec_sqrt(pi()) / 4)),
        _scalar("kTwoPowNeg1p5", beta, "kTwoPowNeg1p5", lambda: rn(dec_sqrt(2) / 4)),
        _scalar("kEps (inverse)", "stats/inverse_distributions.cpp", "kEps", lambda: 2.0**-52),
        _scalar("kHalfEps", "stats/inverse_distributions.cpp", "kHalfEps", lambda: 2.0**-53),
        _scalar("kPi (inverse)", "stats/inverse_distributions.cpp", "kPi", lambda: rn(pi())),
        _scalar("kEulerGamma", dist, "kEulerGamma", lambda: rn(dec_euler_gamma())),
        _scalar("kSqrtPi", dist, "kSqrtPi", lambda: rn(dec_sqrt(pi()))),
        _scalar("kSqrtHalfPi", dist, "kSqrtHalfPi", lambda: rn(dec_sqrt(pi() / 2))),
        _scalar("kLogSqrt2Pi (dist)", dist, "kLogSqrt2Pi", ln_sqrt_2pi),
    ]
    for i in range(9):
        checks.append(
            Check(
                f"kBgratPoly{i}",
                beta,
                f"kBgratPoly{i}",
                "array",
                lambda i=i: small_b_poly(i),
                f"[s^{2 * (9 - i)}] (sinh(s/2)/(s/2))^(b-1) as a polynomial in b; nearest",
            )
        )
    for lit, fn in (
        ("0.7071067811865476", lambda: rn(dec_sqrt(Decimal("0.5")))),
        ("1.4142135623730951", lambda: rn(dec_sqrt(2))),
    ):
        checks.append(Check(f"literal {lit}", special, lit, "inline", lambda fn=fn: [fn()], "nearest"))
    return checks


# Truncation widths of the two truncated tables. kStirlingPoly[0] is reproduced
# by 14 and by 15 digits; 15 is used so both tables share one rule.
STIRLING_POLY_DIGITS = 15
STIRLING_SMALL_DIGITS = 15


def run_check(chk: Check, root: Path) -> Tuple[bool, str]:
    path = root / "src" / "eval" / chk.file
    if not path.is_file():
        return False, f"{chk.name}: source file missing: {chk.file}"
    text = strip_comments(path.read_text(encoding="utf-8"))
    expected = chk.expected()
    if chk.kind == "array":
        got = find_array(text, chk.symbol)
    elif chk.kind == "scalar":
        v = find_scalar(text, chk.symbol)
        got = None if v is None else [v]
    else:
        v = find_inline(text, chk.symbol)
        got = None if v is None else [v]
    if got is None:
        return False, f"{chk.name}: symbol not found in {chk.file}"
    if len(got) != len(expected):
        return False, f"{chk.name}: length {len(got)} in source, {len(expected)} expected"
    bad = [(i, g, e) for i, (g, e) in enumerate(zip(got, expected)) if struct.pack("<d", g) != struct.pack("<d", e)]
    if bad:
        lines = [f"{chk.name}: {len(bad)}/{len(got)} elements differ ({chk.file})"]
        for i, g, e in bad:
            lines.append(f"    [{i}] source {g.hex()}  expected {e.hex()}  ulp diff {ulp_diff(g, e)}")
        return False, "\n".join(lines)
    return True, f"ok  {chk.name:<22} {len(got):>2} element(s)  {chk.rule}"


def main(argv: Optional[Sequence[str]] = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[2])
    ap.add_argument("--table", action="append", help="check only this table (repeatable)")
    ap.add_argument("--list", action="store_true", help="list table names and exit")
    args = ap.parse_args(argv)
    checks = build_checks()
    if args.list:
        for c in checks:
            print(f"{c.name}\t{c.file}\t{c.rule}")
        return 0
    if args.table:
        wanted = set(args.table)
        unknown = wanted - {c.name for c in checks}
        if unknown:
            print("unknown table(s): " + ", ".join(sorted(unknown)), file=sys.stderr)
            return 2
        checks = [c for c in checks if c.name in wanted]
    failed = 0
    for chk in checks:
        ok, msg = run_check(chk, args.root)
        print(msg)
        failed += not ok
    print(f"{len(checks) - failed}/{len(checks)} tables match")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
