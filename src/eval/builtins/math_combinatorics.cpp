// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's combinatorial, numeral-system, precise-
// rounding, and miscellaneous scalar math built-in functions:
//   ARABIC, ROMAN, BASE, DECIMAL, CEILING.PRECISE, FLOOR.PRECISE,
//   ISO.CEILING, COMBIN, COMBINA, FACT, FACTDOUBLE, GCD, LCM,
//   MULTINOMIAL, SQRTPI.
//
// Each impl follows the same recipe as the rest of the builtin catalog:
// coerce arguments via `eval/coerce.h`, propagate the left-most coercion
// error, and return a `Value`.

#include "eval/builtins/math_combinatorics.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <numeric>
#include <string>
#include <string_view>
#include <utility>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Local PI constant used by SQRTPI. Defined locally rather than shared to
// keep this translation unit self-contained, matching the pattern used in
// the test files for `kPi`.
constexpr double kPi = 3.14159265358979323846;

// Factorial table for n = 0..170 (double precision). 170! is the largest
// factorial representable as a finite double; 171! overflows to +Inf.
// Initialised lazily on first use via a function-local static, so
// translation-unit ordering does not matter.
inline const double* factorial_table() {
  static const auto kTable = [] {
    std::array<double, 171> table{};
    table[0] = 1.0;
    for (std::size_t i = 1; i <= 170; ++i) {
      table[i] = table[i - 1] * static_cast<double>(i);
    }
    return table;
  }();
  return kTable.data();
}

// Returns n! as a finite double, or an infinity if n > 170.
inline double factorial_lookup(std::uint32_t n) {
  if (n > 170) {
    return std::numeric_limits<double>::infinity();
  }
  return factorial_table()[n];
}

// Returns true iff `x` is safely representable as a non-negative integer
// after truncation (i.e. Excel-style INT(x) for non-negative inputs). NaN
// and infinities are rejected.
inline bool try_truncate_nonneg(double x, std::uint64_t max, std::uint64_t* out) {
  if (std::isnan(x) || std::isinf(x) || x < 0.0) {
    return false;
  }
  const double t = std::trunc(x);
  if (t > static_cast<double>(max)) {
    return false;
  }
  *out = static_cast<std::uint64_t>(t);
  return true;
}

// ---------------------------------------------------------------------------
// FACT / FACTDOUBLE
// ---------------------------------------------------------------------------

// FACT(n) - n! for non-negative integer n <= 170. Fractional input is
// truncated toward zero. Negative n or n > 170 yields #NUM!.
Value Fact(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  if (std::isnan(x) || std::isinf(x) || x < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double t = std::trunc(x);
  if (t > 170.0) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(factorial_lookup(static_cast<std::uint32_t>(t)));
}

// FACTDOUBLE(n) - double factorial n!! = n*(n-2)*(n-4)*...*1 or *2. By
// Excel convention, `FACTDOUBLE(0) = 1` and `FACTDOUBLE(-1) = 1`; every
// other negative value yields #NUM!. Overflow to +Inf yields #NUM!.
Value FactDouble(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  if (std::isnan(x) || std::isinf(x)) {
    return Value::error(ErrorCode::Num);
  }
  const double t = std::trunc(x);
  if (t < -1.0) {
    return Value::error(ErrorCode::Num);
  }
  if (t <= 0.0) {
    // Both 0 and -1 map to the empty product, which is 1.
    return Value::number(1.0);
  }
  double result = 1.0;
  for (double k = t; k >= 1.0; k -= 2.0) {
    result *= k;
    if (std::isinf(result)) {
      return Value::error(ErrorCode::Num);
    }
  }
  return Value::number(result);
}

// ---------------------------------------------------------------------------
// COMBIN / COMBINA / MULTINOMIAL
// ---------------------------------------------------------------------------

// Computes nCk exactly as a double using the factorial table when n <= 170,
// and log-gamma otherwise. Returns +Inf on overflow; the caller must map
// that to #NUM!.
inline double combin_exact(std::uint64_t n, std::uint64_t k) {
  if (k > n) {
    return 0.0;
  }
  if (k > n - k) {
    k = n - k;
  }
  if (k == 0) {
    return 1.0;
  }
  if (n <= 170) {
    // Direct factorial ratio is safe: denominator factorials both fit.
    return factorial_lookup(static_cast<std::uint32_t>(n)) /
           (factorial_lookup(static_cast<std::uint32_t>(k)) * factorial_lookup(static_cast<std::uint32_t>(n - k)));
  }
  // Fallback via log-gamma for very large n.
  const double log_combin = std::lgamma(static_cast<double>(n) + 1.0) - std::lgamma(static_cast<double>(k) + 1.0) -
                            std::lgamma(static_cast<double>(n - k) + 1.0);
  return std::exp(log_combin);
}

// COMBIN(n, k) - n choose k. Fractional inputs truncated toward zero.
// Negative n, negative k, or k > n yields #NUM!. Overflow yields #NUM!.
Value Combin(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto n_v = coerce_to_number(args[0]);
  if (!n_v) {
    return Value::error(n_v.error());
  }
  auto k_v = coerce_to_number(args[1]);
  if (!k_v) {
    return Value::error(k_v.error());
  }
  std::uint64_t n = 0;
  std::uint64_t k = 0;
  if (!try_truncate_nonneg(n_v.value(), static_cast<std::uint64_t>(1) << 53u, &n) ||
      !try_truncate_nonneg(k_v.value(), static_cast<std::uint64_t>(1) << 53u, &k)) {
    return Value::error(ErrorCode::Num);
  }
  if (k > n) {
    return Value::error(ErrorCode::Num);
  }
  const double r = combin_exact(n, k);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// COMBINA(n, k) - multichoose = C(n+k-1, k). Same error conditions as
// COMBIN (after the k <= n check — which does NOT apply to COMBINA;
// COMBINA allows k > n since order-with-repetition has no such cap).
Value CombinA(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto n_v = coerce_to_number(args[0]);
  if (!n_v) {
    return Value::error(n_v.error());
  }
  auto k_v = coerce_to_number(args[1]);
  if (!k_v) {
    return Value::error(k_v.error());
  }
  std::uint64_t n = 0;
  std::uint64_t k = 0;
  if (!try_truncate_nonneg(n_v.value(), static_cast<std::uint64_t>(1) << 53u, &n) ||
      !try_truncate_nonneg(k_v.value(), static_cast<std::uint64_t>(1) << 53u, &k)) {
    return Value::error(ErrorCode::Num);
  }
  // Excel quirk: `COMBINA(0, 0) = 1`; `COMBINA(0, k>0) = 0`.
  if (n == 0 && k == 0) {
    return Value::number(1.0);
  }
  if (n == 0) {
    return Value::error(ErrorCode::Num);
  }
  // Guard against overflow on n + k - 1.
  if (n > (static_cast<std::uint64_t>(1) << 52u) || k > (static_cast<std::uint64_t>(1) << 52u)) {
    return Value::error(ErrorCode::Num);
  }
  const std::uint64_t upper = n + k - 1;
  const double r = combin_exact(upper, k);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// PERMUT(n, k) - number of k-permutations of n distinct items =
// n! / (n-k)! = n * (n-1) * ... * (n-k+1). Both arguments floor to
// non-negative integer; `k > n` yields `#NUM!`, as does overflow.
// Edge cases: `PERMUT(n, 0) = 1` for any n >= 0.
Value Permut(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto n_v = coerce_to_number(args[0]);
  if (!n_v) {
    return Value::error(n_v.error());
  }
  auto k_v = coerce_to_number(args[1]);
  if (!k_v) {
    return Value::error(k_v.error());
  }
  std::uint64_t n = 0;
  std::uint64_t k = 0;
  if (!try_truncate_nonneg(n_v.value(), static_cast<std::uint64_t>(1) << 53u, &n) ||
      !try_truncate_nonneg(k_v.value(), static_cast<std::uint64_t>(1) << 53u, &k)) {
    return Value::error(ErrorCode::Num);
  }
  if (k > n) {
    return Value::error(ErrorCode::Num);
  }
  // Multiply incrementally in double precision; bail out as soon as the
  // running product overflows to infinity.
  double result = 1.0;
  for (std::uint64_t i = 0; i < k; ++i) {
    result *= static_cast<double>(n - i);
    if (std::isinf(result) || std::isnan(result)) {
      return Value::error(ErrorCode::Num);
    }
  }
  return Value::number(result);
}

// PERMUTATIONA(n, k) - number of k-permutations of n items with repetition =
// n^k. Both arguments floor to non-negative integer. `PERMUTATIONA(0, 0) = 1`
// by Excel convention; `PERMUTATIONA(0, k>0) = 0`. Overflow yields `#NUM!`.
Value PermutationA(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto n_v = coerce_to_number(args[0]);
  if (!n_v) {
    return Value::error(n_v.error());
  }
  auto k_v = coerce_to_number(args[1]);
  if (!k_v) {
    return Value::error(k_v.error());
  }
  std::uint64_t n = 0;
  std::uint64_t k = 0;
  if (!try_truncate_nonneg(n_v.value(), static_cast<std::uint64_t>(1) << 53u, &n) ||
      !try_truncate_nonneg(k_v.value(), static_cast<std::uint64_t>(1) << 53u, &k)) {
    return Value::error(ErrorCode::Num);
  }
  if (n == 0 && k == 0) {
    return Value::number(1.0);
  }
  if (n == 0) {
    return Value::number(0.0);
  }
  const double r = std::pow(static_cast<double>(n), static_cast<double>(k));
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// MULTINOMIAL(a1, a2, ...) - multinomial coefficient = (sum(a_i))! / prod(a_i!).
// Each argument truncated to non-negative integer; negative -> #NUM!.
// Overflow -> #NUM!.
Value Multinomial(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  // Accumulate sum first while validating each argument.
  std::uint64_t total = 0;
  // We also maintain the running result multiplicatively using
  // multinomial(n, k) = multinomial(n-1, k_last_minus_1) * C(n, k_last);
  // equivalently: result = (sum!) / prod(k_i!) which we compute incrementally
  // via `result *= C(total_so_far, next_k)` (Pascal's rule for multinomials).
  double result = 1.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    std::uint64_t k = 0;
    if (!try_truncate_nonneg(coerced.value(), static_cast<std::uint64_t>(1) << 52u, &k)) {
      return Value::error(ErrorCode::Num);
    }
    total += k;
    if (total > (static_cast<std::uint64_t>(1) << 52u)) {
      return Value::error(ErrorCode::Num);
    }
    // multinomial step: multiply by C(total, k).
    const double step = combin_exact(total, k);
    result *= step;
    if (std::isnan(result) || std::isinf(result)) {
      return Value::error(ErrorCode::Num);
    }
  }
  return Value::number(result);
}

// ---------------------------------------------------------------------------
// GCD / LCM
// ---------------------------------------------------------------------------

// GCD(a1, a2, ...) - greatest common divisor. All args truncated to
// non-negative integers. Negative -> #NUM!. All zero -> 0.
Value Gcd(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::uint64_t g = 0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    std::uint64_t k = 0;
    if (!try_truncate_nonneg(coerced.value(), static_cast<std::uint64_t>(1) << 53u, &k)) {
      return Value::error(ErrorCode::Num);
    }
    g = std::gcd(g, k);
  }
  return Value::number(static_cast<double>(g));
}

// LCM(a1, a2, ...) - least common multiple. All args truncated to
// non-negative integers. Any zero -> 0 (Excel quirk). Overflow -> #NUM!.
Value Lcm(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  std::uint64_t l = 1;
  bool any_zero = false;
  for (std::uint32_t i = 0; i < arity; ++i) {
    auto coerced = coerce_to_number(args[i]);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    std::uint64_t k = 0;
    if (!try_truncate_nonneg(coerced.value(), static_cast<std::uint64_t>(1) << 53u, &k)) {
      return Value::error(ErrorCode::Num);
    }
    if (k == 0) {
      any_zero = true;
      continue;
    }
    const std::uint64_t g = std::gcd(l, k);
    // Safe multiplication: l / g * k, with overflow check.
    const std::uint64_t lhs = l / g;
    // Overflow if `lhs * k` exceeds 2^53 (the double-precision integer limit).
    const std::uint64_t kLimit = static_cast<std::uint64_t>(1) << 53u;
    if (k != 0 && lhs > kLimit / k) {
      return Value::error(ErrorCode::Num);
    }
    l = lhs * k;
  }
  if (any_zero) {
    return Value::number(0.0);
  }
  return Value::number(static_cast<double>(l));
}

// ---------------------------------------------------------------------------
// ARABIC / ROMAN
// ---------------------------------------------------------------------------

// Maps a single Roman character to its value, or 0 if invalid.
inline int roman_char_value(char c) {
  switch (c) {
    case 'I':
    case 'i':
      return 1;
    case 'V':
    case 'v':
      return 5;
    case 'X':
    case 'x':
      return 10;
    case 'L':
    case 'l':
      return 50;
    case 'C':
    case 'c':
      return 100;
    case 'D':
    case 'd':
      return 500;
    case 'M':
    case 'm':
      return 1000;
    default:
      return 0;
  }
}

// ARABIC(text) - Roman numeral string -> integer. Accepts modern
// subtractive forms (MCMXCIX = 1999) as well as additive forms (MDCCCCLXXXXVIIII).
// Optional leading '-' produces a negative result. Empty / whitespace-only
// input yields 0. Any non-Roman / non-whitespace character yields #VALUE!.
Value Arabic(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& s = text.value();
  // Strip surrounding ASCII whitespace (Excel is tolerant of padding).
  std::size_t lo = 0;
  std::size_t hi = s.size();
  while (lo < hi && (s[lo] == ' ' || s[lo] == '\t')) {
    ++lo;
  }
  while (hi > lo && (s[hi - 1] == ' ' || s[hi - 1] == '\t')) {
    --hi;
  }
  if (lo == hi) {
    return Value::number(0.0);
  }
  bool negative = false;
  if (s[lo] == '-') {
    negative = true;
    ++lo;
  }
  if (lo == hi) {
    // A lone '-' is not a valid Roman numeral.
    return Value::error(ErrorCode::Value);
  }
  // Left-to-right subtractive evaluation: if the current char's value is
  // less than the next char's, subtract; otherwise add. Non-Roman char in
  // the middle -> #VALUE!.
  long total = 0;
  for (std::size_t i = lo; i < hi; ++i) {
    const int cur = roman_char_value(s[i]);
    if (cur == 0) {
      return Value::error(ErrorCode::Value);
    }
    const int next = (i + 1 < hi) ? roman_char_value(s[i + 1]) : 0;
    if (next == 0 && i + 1 < hi) {
      return Value::error(ErrorCode::Value);
    }
    if (cur < next) {
      total -= cur;
    } else {
      total += cur;
    }
  }
  return Value::number(static_cast<double>(negative ? -total : total));
}

// Converts `n` (in 0..3999) to a Roman numeral under the requested `form`.
//
// Excel's ROMAN supports five forms (0..4) of increasing concision. Form 0
// is the strict classical subtractive form; forms 1..4 progressively allow
// more aggressive subtractive abbreviations. Microsoft documents the
// reference outputs at https://support.microsoft.com/en-us/office/roman-function:
//
//   ROMAN(499, 0..4)  = CDXCIX, LDVLIV, XDIX, VDIV, ID
//   ROMAN(1999, 0..4) = MCMXCIX, MLMVLIV, MXMIX, MVMIV, MIM
//
// The implementation is greedy over a per-form ordered table of
// (value, glyph) pairs. Each higher form simply prepends additional
// subtractive pairs to the form-0 table.
inline std::string roman_render(int n, int form) {
  // Base (form 0) subtractive pair table, value-descending.
  struct Pair {
    int value;
    const char* glyph;
  };
  static constexpr Pair kForm0[] = {
      {1000, "M"}, {900, "CM"}, {500, "D"}, {400, "CD"}, {100, "C"}, {90, "XC"}, {50, "L"},
      {40, "XL"},  {10, "X"},   {9, "IX"},  {5, "V"},    {4, "IV"},  {1, "I"},
  };
  // Form 1 additions: allows V/L/C as the subtracted glyph in one-step
  // pairings (VL=45, LD=450, LM=950, VC=95).
  static constexpr Pair kForm1Adds[] = {
      {950, "LM"},
      {450, "LD"},
      {95, "VC"},
      {45, "VL"},
  };
  // Form 2 additions: allows X as the subtracted glyph two decades down
  // (XM=990, XD=490, IL=49).
  static constexpr Pair kForm2Adds[] = {
      {990, "XM"},
      {490, "XD"},
      {49, "IL"},
  };
  // Form 3 additions: allows V two decades down (VM=995, VD=495).
  static constexpr Pair kForm3Adds[] = {
      {995, "VM"},
      {495, "VD"},
  };
  // Form 4 additions: allows I three decades down (IM=999, ID=499, IC=99).
  static constexpr Pair kForm4Adds[] = {
      {999, "IM"},
      {499, "ID"},
      {99, "IC"},
  };

  // Build a single ordered table by merging all enabled additions into the
  // form-0 table. The count is bounded and small enough to use a fixed
  // static buffer.
  Pair table[sizeof(kForm0) / sizeof(kForm0[0]) + 16];
  std::size_t count = 0;
  auto append = [&](const Pair* src, std::size_t n_src) {
    for (std::size_t i = 0; i < n_src; ++i) {
      table[count++] = src[i];
    }
  };
  append(kForm0, sizeof(kForm0) / sizeof(kForm0[0]));
  if (form >= 1) {
    append(kForm1Adds, sizeof(kForm1Adds) / sizeof(kForm1Adds[0]));
  }
  if (form >= 2) {
    append(kForm2Adds, sizeof(kForm2Adds) / sizeof(kForm2Adds[0]));
  }
  if (form >= 3) {
    append(kForm3Adds, sizeof(kForm3Adds) / sizeof(kForm3Adds[0]));
  }
  if (form >= 4) {
    append(kForm4Adds, sizeof(kForm4Adds) / sizeof(kForm4Adds[0]));
  }
  // Sort value-descending so the greedy consumption picks the longest
  // subtractive pair first.
  std::sort(table, table + count, [](const Pair& a, const Pair& b) { return a.value > b.value; });

  std::string out;
  int remaining = n;
  for (std::size_t i = 0; i < count && remaining > 0; ++i) {
    while (remaining >= table[i].value) {
      out.append(table[i].glyph);
      remaining -= table[i].value;
    }
  }
  return out;
}

// ROMAN(num, [form]) - integer in [0, 3999] -> Roman string. `form` in
// 0..4 controls abbreviation level. Out-of-range num or form yields #VALUE!.
// Fractional input truncates toward zero. Excel's documented output for 0
// is the empty string.
Value Roman(const Value* args, std::uint32_t arity, Arena& arena) {
  auto n_v = coerce_to_number(args[0]);
  if (!n_v) {
    return Value::error(n_v.error());
  }
  int form = 0;
  if (arity >= 2) {
    auto f_v = coerce_to_number(args[1]);
    if (!f_v) {
      return Value::error(f_v.error());
    }
    // Excel also accepts TRUE / FALSE for form: 0 / 1 under coerce_to_number.
    const double f = std::trunc(f_v.value());
    if (std::isnan(f) || std::isinf(f) || f < 0.0 || f > 4.0) {
      return Value::error(ErrorCode::Value);
    }
    form = static_cast<int>(f);
  }
  const double raw = n_v.value();
  if (std::isnan(raw) || std::isinf(raw)) {
    return Value::error(ErrorCode::Value);
  }
  const double t = std::trunc(raw);
  if (t < 0.0 || t > 3999.0) {
    return Value::error(ErrorCode::Value);
  }
  const int n = static_cast<int>(t);
  const std::string out = roman_render(n, form);
  return Value::text(arena.intern(out));
}

// ---------------------------------------------------------------------------
// BASE / DECIMAL
// ---------------------------------------------------------------------------

// BASE(num, radix, [min_len]) - non-negative integer -> string in `radix`
// (2..36). Pads with leading '0' to `min_len`. num must fit in 2^53-1;
// larger or negative num yields #NUM!. radix out of range -> #NUM!. min_len
// must be in 0..255; out-of-range -> #VALUE!.
Value Base(const Value* args, std::uint32_t arity, Arena& arena) {
  auto num_v = coerce_to_number(args[0]);
  if (!num_v) {
    return Value::error(num_v.error());
  }
  auto radix_v = coerce_to_number(args[1]);
  if (!radix_v) {
    return Value::error(radix_v.error());
  }
  int min_len = 0;
  if (arity >= 3) {
    auto len_v = coerce_to_number(args[2]);
    if (!len_v) {
      return Value::error(len_v.error());
    }
    const double lf = std::trunc(len_v.value());
    if (std::isnan(lf) || std::isinf(lf) || lf < 0.0 || lf > 255.0) {
      return Value::error(ErrorCode::Value);
    }
    min_len = static_cast<int>(lf);
  }
  const double rf = std::trunc(radix_v.value());
  if (std::isnan(rf) || std::isinf(rf) || rf < 2.0 || rf > 36.0) {
    return Value::error(ErrorCode::Num);
  }
  const int radix = static_cast<int>(rf);
  const double nf = std::trunc(num_v.value());
  if (std::isnan(nf) || std::isinf(nf) || nf < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double kMax = static_cast<double>((static_cast<std::uint64_t>(1) << 53u) - 1);
  if (nf > kMax) {
    return Value::error(ErrorCode::Num);
  }
  std::uint64_t n = static_cast<std::uint64_t>(nf);
  std::string digits;
  if (n == 0) {
    digits.push_back('0');
  } else {
    while (n > 0) {
      const int d = static_cast<int>(n % static_cast<std::uint64_t>(radix));
      n /= static_cast<std::uint64_t>(radix);
      digits.push_back(static_cast<char>(d < 10 ? ('0' + d) : ('A' + (d - 10))));
    }
    // Reverse to most-significant-first.
    for (std::size_t i = 0, j = digits.size() - 1; i < j; ++i, --j) {
      std::swap(digits[i], digits[j]);
    }
  }
  if (static_cast<int>(digits.size()) < min_len) {
    digits.insert(digits.begin(), static_cast<std::size_t>(min_len) - digits.size(), '0');
  }
  return Value::text(arena.intern(digits));
}

// DECIMAL(text, radix) - string in `radix` (2..36) -> integer. Case-
// insensitive. Empty / whitespace-only text yields 0 (matches Mac Excel 365;
// the empty product of digits contributes nothing). Digits out of range for
// the radix yield #NUM!. Overflow (> 2^53-1) yields #NUM!.
Value Decimal(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto radix_v = coerce_to_number(args[1]);
  if (!radix_v) {
    return Value::error(radix_v.error());
  }
  const double rf = std::trunc(radix_v.value());
  if (std::isnan(rf) || std::isinf(rf) || rf < 2.0 || rf > 36.0) {
    return Value::error(ErrorCode::Num);
  }
  const int radix = static_cast<int>(rf);
  const std::string& s = text.value();
  // Trim surrounding ASCII whitespace.
  std::size_t lo = 0;
  std::size_t hi = s.size();
  while (lo < hi && (s[lo] == ' ' || s[lo] == '\t')) {
    ++lo;
  }
  while (hi > lo && (s[hi - 1] == ' ' || s[hi - 1] == '\t')) {
    --hi;
  }
  if (lo == hi) {
    // Excel 365 returns 0 for empty / whitespace-only input.
    return Value::number(0.0);
  }
  std::uint64_t acc = 0;
  const std::uint64_t kLimit = (static_cast<std::uint64_t>(1) << 53u) - 1;
  for (std::size_t i = lo; i < hi; ++i) {
    char c = s[i];
    int d = -1;
    if (c >= '0' && c <= '9') {
      d = c - '0';
    } else if (c >= 'A' && c <= 'Z') {
      d = c - 'A' + 10;
    } else if (c >= 'a' && c <= 'z') {
      d = c - 'a' + 10;
    }
    if (d < 0 || d >= radix) {
      return Value::error(ErrorCode::Num);
    }
    // Overflow guard: acc * radix + d must stay within 2^53 - 1.
    if (acc > (kLimit - static_cast<std::uint64_t>(d)) / static_cast<std::uint64_t>(radix)) {
      return Value::error(ErrorCode::Num);
    }
    acc = acc * static_cast<std::uint64_t>(radix) + static_cast<std::uint64_t>(d);
  }
  return Value::number(static_cast<double>(acc));
}

// ---------------------------------------------------------------------------
// CEILING.PRECISE / FLOOR.PRECISE / ISO.CEILING
// ---------------------------------------------------------------------------

// Shared helper: round to nearest multiple of |sig| in the given direction.
// `up = true`  -> ceil (toward +infinity).
// `up = false` -> floor (toward -infinity).
inline Value precise_rounding(const Value* args, std::uint32_t arity, bool up) {
  auto num_v = coerce_to_number(args[0]);
  if (!num_v) {
    return Value::error(num_v.error());
  }
  double sig = 1.0;
  if (arity >= 2) {
    auto s_v = coerce_to_number(args[1]);
    if (!s_v) {
      return Value::error(s_v.error());
    }
    sig = s_v.value();
  }
  const double n = num_v.value();
  if (sig == 0.0) {
    return Value::number(0.0);
  }
  const double abs_s = std::fabs(sig);
  const double scaled = n / abs_s;
  const double rounded = up ? std::ceil(scaled) : std::floor(scaled);
  const double r = rounded * abs_s;
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

// CEILING.PRECISE(num, [sig]) - round up toward +infinity to the nearest
// multiple of |sig|. sig defaults to 1. Sign of sig is ignored.
Value CeilingPrecise(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return precise_rounding(args, arity, /*up=*/true);
}

// FLOOR.PRECISE(num, [sig]) - round down toward -infinity to the nearest
// multiple of |sig|.
Value FloorPrecise(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return precise_rounding(args, arity, /*up=*/false);
}

// ISO.CEILING(num, [sig]) - alias of CEILING.PRECISE per the ISO/IEC 29500
// definition Excel uses.
Value IsoCeiling(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  return precise_rounding(args, arity, /*up=*/true);
}

// ---------------------------------------------------------------------------
// SQRTPI
// ---------------------------------------------------------------------------

// SQRTPI(num) - sqrt(num * PI). Negative num -> #NUM!.
Value SqrtPi(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  if (x < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  const double r = std::sqrt(x * kPi);
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

}  // namespace

void register_math_combinatorics_builtins(FunctionRegistry& registry) {
  // Factorials / combinatorics.
  registry.register_function(FunctionDef{"FACT", 1u, 1u, &Fact});
  registry.register_function(FunctionDef{"FACTDOUBLE", 1u, 1u, &FactDouble});
  registry.register_function(FunctionDef{"COMBIN", 2u, 2u, &Combin});
  registry.register_function(FunctionDef{"COMBINA", 2u, 2u, &CombinA});
  registry.register_function(FunctionDef{"PERMUT", 2u, 2u, &Permut});
  registry.register_function(FunctionDef{"PERMUTATIONA", 2u, 2u, &PermutationA});

  // Variadic combinatorial aggregators. Range-aware: `=MULTINOMIAL(A1:A3)`
  // expands the rectangle into scalar args before the impl runs. Unlike the
  // SUM family, these do NOT set `range_filter_numeric_only`: Excel surfaces
  // #VALUE! on a Text cell sourced from a range, matching the direct
  // coercion path.
  {
    FunctionDef def{"MULTINOMIAL", 1u, kVariadic, &Multinomial};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"GCD", 1u, kVariadic, &Gcd};
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  {
    FunctionDef def{"LCM", 1u, kVariadic, &Lcm};
    def.accepts_ranges = true;
    registry.register_function(def);
  }

  // Numeral-system conversions.
  registry.register_function(FunctionDef{"ARABIC", 1u, 1u, &Arabic});
  registry.register_function(FunctionDef{"ROMAN", 1u, 2u, &Roman});
  registry.register_function(FunctionDef{"BASE", 2u, 3u, &Base});
  registry.register_function(FunctionDef{"DECIMAL", 2u, 2u, &Decimal});

  // Precise / ISO rounding.
  registry.register_function(FunctionDef{"CEILING.PRECISE", 1u, 2u, &CeilingPrecise});
  registry.register_function(FunctionDef{"FLOOR.PRECISE", 1u, 2u, &FloorPrecise});
  registry.register_function(FunctionDef{"ISO.CEILING", 1u, 2u, &IsoCeiling});

  // Miscellaneous.
  registry.register_function(FunctionDef{"SQRTPI", 1u, 1u, &SqrtPi});
}

}  // namespace eval
}  // namespace formulon
