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
#include <utility>

#include "eval/builtins/numeric_helpers.h"
#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using builtins_detail::kPi;
using builtins_detail::snap_to_integer;
using builtins_detail::to_finite_value;

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

inline Expected<double, ErrorCode> read_number_arg(const Value* args, std::uint32_t index) {
  return builtins_detail::read_required_number(args, index);
}

Expected<std::uint64_t, ErrorCode> read_nonneg_uint_arg(const Value* args, std::uint32_t index, std::uint64_t max) {
  auto x = read_number_arg(args, index);
  if (!x) {
    return x.error();
  }
  std::uint64_t out = 0;
  if (!try_truncate_nonneg(x.value(), max, &out)) {
    return ErrorCode::Num;
  }
  return out;
}

struct UIntPair {
  std::uint64_t first;
  std::uint64_t second;
};

Expected<UIntPair, ErrorCode> read_nonneg_uint_pair(const Value* args, std::uint64_t max) {
  auto first = read_nonneg_uint_arg(args, 0, max);
  if (!first) {
    return first.error();
  }
  auto second = read_nonneg_uint_arg(args, 1, max);
  if (!second) {
    return second.error();
  }
  return UIntPair{first.value(), second.value()};
}

// ---------------------------------------------------------------------------
// FACT / FACTDOUBLE
// ---------------------------------------------------------------------------

// FACT(n) - n! for non-negative integer n <= 170. Fractional input is
// truncated toward zero. Negative n or n > 170 yields #NUM!.
Value Fact(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto n = read_nonneg_uint_arg(args, 0, 170u);
  if (!n) {
    return Value::error(n.error());
  }
  return Value::number(factorial_lookup(static_cast<std::uint32_t>(n.value())));
}

// FACTDOUBLE(n) - double factorial n!! = n*(n-2)*(n-4)*...*1 or *2. By
// Excel convention, `FACTDOUBLE(0) = 1` and `FACTDOUBLE(-1) = 1`; every
// other negative value yields #NUM!. Overflow to +Inf yields #NUM!.
Value FactDouble(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto x = read_number_arg(args, 0);
  if (!x) {
    return Value::error(x.error());
  }
  const double t = std::trunc(x.value());
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
  auto pair = read_nonneg_uint_pair(args, static_cast<std::uint64_t>(1) << 53u);
  if (!pair) {
    return Value::error(pair.error());
  }
  const std::uint64_t n = pair.value().first;
  const std::uint64_t k = pair.value().second;
  if (k > n) {
    return Value::error(ErrorCode::Num);
  }
  return to_finite_value(combin_exact(n, k));
}

// COMBINA(n, k) - multichoose = C(n+k-1, k). Same error conditions as
// COMBIN (after the k <= n check — which does NOT apply to COMBINA;
// COMBINA allows k > n since order-with-repetition has no such cap).
Value CombinA(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto pair = read_nonneg_uint_pair(args, static_cast<std::uint64_t>(1) << 53u);
  if (!pair) {
    return Value::error(pair.error());
  }
  const std::uint64_t n = pair.value().first;
  const std::uint64_t k = pair.value().second;
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
  return to_finite_value(combin_exact(upper, k));
}

// PERMUT(n, k) - number of k-permutations of n distinct items =
// n! / (n-k)! = n * (n-1) * ... * (n-k+1). Both arguments floor to
// non-negative integer; `k > n` yields `#NUM!`, as does overflow.
// Edge cases: `PERMUT(n, 0) = 1` for any n >= 0.
Value Permut(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto pair = read_nonneg_uint_pair(args, static_cast<std::uint64_t>(1) << 53u);
  if (!pair) {
    return Value::error(pair.error());
  }
  const std::uint64_t n = pair.value().first;
  const std::uint64_t k = pair.value().second;
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
  auto pair = read_nonneg_uint_pair(args, static_cast<std::uint64_t>(1) << 53u);
  if (!pair) {
    return Value::error(pair.error());
  }
  const std::uint64_t n = pair.value().first;
  const std::uint64_t k = pair.value().second;
  if (n == 0 && k == 0) {
    return Value::number(1.0);
  }
  if (n == 0) {
    return Value::number(0.0);
  }
  return to_finite_value(std::pow(static_cast<double>(n), static_cast<double>(k)));
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
    auto k_e = read_nonneg_uint_arg(args, i, static_cast<std::uint64_t>(1) << 52u);
    if (!k_e) {
      return Value::error(k_e.error());
    }
    const std::uint64_t k = k_e.value();
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
// The implementation is greedy over one value-descending table of
// (value, glyph) pairs; each pair records the lowest form that admits it,
// so a form is a subsequence of that single table.
inline std::string roman_render(int n, int form) {
  struct Pair {
    int value;
    const char* glyph;
    int min_form;  // Lowest `form` that admits this pairing.
  };
  // The full value-descending pair table, form 0 first and every
  // higher-form addition slotted into its ordered position. A form selects
  // a subsequence of this one table, so no per-call merge or sort is needed:
  //
  //   * form 1 adds V/L/C as the subtracted glyph in one-step pairings
  //     (VL=45, LD=450, LM=950, VC=95);
  //   * form 2 adds X and I two decades up (XM=990, XD=490, IC=99, IL=49) --
  //     Mac Excel 365 places IC/IL here, not in form 4 as the Microsoft docs'
  //     499/1999 examples might suggest; the oracle fixture is authoritative;
  //   * form 3 adds V two decades up (VM=995, VD=495);
  //   * form 4 adds I three decades up (IM=999, ID=499).
  static constexpr Pair kPairs[] = {
      {1000, "M", 0}, {999, "IM", 4}, {995, "VM", 3}, {990, "XM", 2}, {950, "LM", 1}, {900, "CM", 0}, {500, "D", 0},
      {499, "ID", 4}, {495, "VD", 3}, {490, "XD", 2}, {450, "LD", 1}, {400, "CD", 0}, {100, "C", 0},  {99, "IC", 2},
      {95, "VC", 1},  {90, "XC", 0},  {50, "L", 0},   {49, "IL", 2},  {45, "VL", 1},  {40, "XL", 0},  {10, "X", 0},
      {9, "IX", 0},   {5, "V", 0},    {4, "IV", 0},   {1, "I", 0},
  };

  std::string out;
  int remaining = n;
  for (const Pair& pair : kPairs) {
    if (remaining <= 0) {
      break;
    }
    if (pair.min_form > form) {
      continue;
    }
    while (remaining >= pair.value) {
      out.append(pair.glyph);
      remaining -= pair.value;
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
    // TRUE / FALSE map to the two ends of the scale (Classic / Most simplified),
    // not 1 / 0 as `coerce_to_number` would give.
    if (args[1].is_boolean()) {
      form = args[1].as_boolean() ? 0 : 4;
    } else {
      auto f_v = coerce_to_number(args[1]);
      if (!f_v) {
        return Value::error(f_v.error());
      }
      const double f = std::trunc(f_v.value());
      if (std::isnan(f) || std::isinf(f) || f < 0.0 || f > 4.0) {
        return Value::error(ErrorCode::Value);
      }
      form = static_cast<int>(f);
    }
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
  const double scaled = snap_to_integer(n / abs_s);
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
  auto x = read_number_arg(args, 0);
  if (!x) {
    return Value::error(x.error());
  }
  if (x.value() < 0.0) {
    return Value::error(ErrorCode::Num);
  }
  return to_finite_value(std::sqrt(x.value() * kPi));
}

}  // namespace

void register_math_combinatorics_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"FACT", 1u, 1u, &Fact},
      {"FACTDOUBLE", 1u, 1u, &FactDouble},
      {"COMBIN", 2u, 2u, &Combin},
      {"COMBINA", 2u, 2u, &CombinA},
      {"PERMUT", 2u, 2u, &Permut},
      {"PERMUTATIONA", 2u, 2u, &PermutationA},
      {"ARABIC", 1u, 1u, &Arabic},
      {"ROMAN", 1u, 2u, &Roman},
      {"BASE", 2u, 3u, &Base},
      {"DECIMAL", 2u, 2u, &Decimal},
      {"CEILING.PRECISE", 1u, 2u, &CeilingPrecise},
      {"FLOOR.PRECISE", 1u, 2u, &FloorPrecise},
      {"ISO.CEILING", 1u, 2u, &IsoCeiling},
      {"SQRTPI", 1u, 1u, &SqrtPi},
      {"MULTINOMIAL", 1u, kVariadic, &Multinomial, true, true},
      {"GCD", 1u, kVariadic, &Gcd, true, true, false, false, false, FunctionDef::BlankScalarPolicy::RejectAnyScalar,
       ErrorCode::Value},
      {"LCM", 1u, kVariadic, &Lcm, true, true, false, false, false, FunctionDef::BlankScalarPolicy::RejectAnyScalar,
       ErrorCode::Value},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
