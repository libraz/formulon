//
// Implementation of the shared decimal parser. See `double_parse.h` for the
// contract.
//
// The scan collects significant digits and a decimal exponent, then hands
// them to double-conversion's `Strtod`, which rounds correctly. `Strtod`
// looks at no more than 780 significant digits; when more nonzero digits
// follow, a trailing sticky `1` keeps the rounding direction, the same
// device `Strtod` uses when it cuts a longer buffer itself.

#include "utils/double_parse.h"

#include <cfloat>
#include <cmath>
#include <cstdint>
#include <cstring>

#include "double-conversion/strtod.h"

namespace formulon {
namespace {

constexpr int kMaxDigits = 780;
// Exponent digits beyond this magnitude cannot change the result: every
// value is already 0 or ±inf, whatever the mantissa length.
constexpr int kExponentCap = 100000;

bool is_c_space(char c) noexcept {
  return c == ' ' || (c >= '\t' && c <= '\r');
}

bool is_digit(char c) noexcept {
  return c >= '0' && c <= '9';
}

// The sign, the leading significant digits (at most `max_digits`, leading zeros
// dropped), the decimal exponent of the last kept digit and whether any
// dropped digit was nonzero, for the longest decimal prefix of a string.
struct ScannedDecimal {
  std::size_t consumed = 0;
  bool negative = false;
  char digits[kMaxDigits + 1];
  int count = 0;
  int exponent = 0;
  bool dropped_nonzero = false;
};

ScannedDecimal scan_decimal(std::string_view text, int max_digits) noexcept {
  ScannedDecimal out;
  const std::size_t n = text.size();
  std::size_t i = 0;
  while (i < n && is_c_space(text[i])) {
    ++i;
  }
  if (i < n && (text[i] == '+' || text[i] == '-')) {
    out.negative = text[i] == '-';
    ++i;
  }
  bool mantissa_digit = false;
  for (; i < n && is_digit(text[i]); ++i) {
    mantissa_digit = true;
    if (out.count == 0 && text[i] == '0') {
      continue;
    }
    if (out.count < max_digits) {
      out.digits[out.count++] = text[i];
    } else {
      ++out.exponent;
      out.dropped_nonzero = out.dropped_nonzero || text[i] != '0';
    }
  }
  if (i < n && text[i] == '.' && (mantissa_digit || (i + 1 < n && is_digit(text[i + 1])))) {
    for (++i; i < n && is_digit(text[i]); ++i) {
      mantissa_digit = true;
      if (out.count == 0 && text[i] == '0') {
        --out.exponent;
      } else if (out.count < max_digits) {
        out.digits[out.count++] = text[i];
        --out.exponent;
      } else {
        out.dropped_nonzero = out.dropped_nonzero || text[i] != '0';
      }
    }
  }
  if (!mantissa_digit) {
    return out;
  }
  if (i < n && (text[i] == 'e' || text[i] == 'E')) {
    std::size_t j = i + 1;
    bool exponent_negative = false;
    if (j < n && (text[j] == '+' || text[j] == '-')) {
      exponent_negative = text[j] == '-';
      ++j;
    }
    if (j < n && is_digit(text[j])) {
      int e = 0;
      for (; j < n && is_digit(text[j]); ++j) {
        if (e < kExponentCap) {
          e = e * 10 + (text[j] - '0');
        }
      }
      out.exponent += exponent_negative ? -e : e;
      i = j;
    }
  }
  out.consumed = i;
  return out;
}

// Excel's decimal scaling works in a 64-bit significand, `mant * 2^exp`.
struct Wide {
  std::uint64_t mant;
  int exp;
};

// A power of ten as the nearest 64-bit significand, `mant * 2^exp`.
struct PowerOfTen {
  std::uint64_t mant;
  int exp;
};

// 10^1..10^16, then 10^32, 10^48, ..., 10^256 (index 16 + q - 2 for 10^(16q)).
constexpr PowerOfTen kPositivePowers[] = {
    {0xa000000000000000ULL, -60}, {0xc800000000000000ULL, -57}, {0xfa00000000000000ULL, -54},
    {0x9c40000000000000ULL, -50}, {0xc350000000000000ULL, -47}, {0xf424000000000000ULL, -44},
    {0x9896800000000000ULL, -40}, {0xbebc200000000000ULL, -37}, {0xee6b280000000000ULL, -34},
    {0x9502f90000000000ULL, -30}, {0xba43b74000000000ULL, -27}, {0xe8d4a51000000000ULL, -24},
    {0x9184e72a00000000ULL, -20}, {0xb5e620f480000000ULL, -17}, {0xe35fa931a0000000ULL, -14},
    {0x8e1bc9bf04000000ULL, -10}, {0x9dc5ada82b70b59eULL, 43},  {0xaf298d050e4395d7ULL, 96},
    {0xc2781f49ffcfa6d5ULL, 149}, {0xd7e77a8f87daf7fcULL, 202}, {0xefb3ab16c59b14a3ULL, 255},
    {0x850fadc09923329eULL, 309}, {0x93ba47c980e98ce0ULL, 362}, {0xa402b9c5a8d3a6e7ULL, 415},
    {0xb616a12b7fe617aaULL, 468}, {0xca28a291859bbf93ULL, 521}, {0xe070f78d3927556bULL, 574},
    {0xf92e0c3537826146ULL, 627}, {0x8a5296ffe33cc930ULL, 681}, {0x9991a6f3d6bf1766ULL, 734},
    {0xaa7eebfb9df9de8eULL, 787},
};

constexpr PowerOfTen kNegativePowers[] = {
    {0xcccccccccccccccdULL, -67},  {0xa3d70a3d70a3d70aULL, -70},  {0x83126e978d4fdf3bULL, -73},
    {0xd1b71758e219652cULL, -77},  {0xa7c5ac471b478423ULL, -80},  {0x8637bd05af6c69b6ULL, -83},
    {0xd6bf94d5e57a42bcULL, -87},  {0xabcc77118461cefdULL, -90},  {0x89705f4136b4a597ULL, -93},
    {0xdbe6fecebdedd5bfULL, -97},  {0xafebff0bcb24aaffULL, -100}, {0x8cbccc096f5088ccULL, -103},
    {0xe12e13424bb40e13ULL, -107}, {0xb424dc35095cd80fULL, -110}, {0x901d7cf73ab0acd9ULL, -113},
    {0xe69594bec44de15bULL, -117}, {0xcfb11ead453994baULL, -170}, {0xbb127c53b17ec159ULL, -223},
    {0xa87fea27a539e9a5ULL, -276}, {0x97c560ba6b0919a6ULL, -329}, {0x88b402f7fd75539bULL, -382},
    {0xf64335bcf065d37dULL, -436}, {0xddd0467c64bce4a1ULL, -489}, {0xc7caba6e7c5382c9ULL, -542},
    {0xb3f4e093db73a093ULL, -595}, {0xa21727db38cb0030ULL, -648}, {0x91ff83775423cc06ULL, -701},
    {0x8380dea93da4bc60ULL, -754}, {0xece53cec4a314ebeULL, -808}, {0xd5605fcdcf32e1d7ULL, -861},
    {0xc0314325637a193aULL, -914},
};

int bit_length(std::uint64_t v) noexcept {
  int n = 0;
  for (; v != 0U; v >>= 1U) {
    ++n;
  }
  return n;
}

// Full 128-bit product of two 64-bit values from 32-bit halves.
void multiply_64(std::uint64_t a, std::uint64_t b, std::uint64_t* hi, std::uint64_t* lo) noexcept {
  const std::uint64_t a0 = a & 0xffffffffULL;
  const std::uint64_t a1 = a >> 32U;
  const std::uint64_t b0 = b & 0xffffffffULL;
  const std::uint64_t b1 = b >> 32U;
  const std::uint64_t p00 = a0 * b0;
  const std::uint64_t p01 = a0 * b1;
  const std::uint64_t p10 = a1 * b0;
  const std::uint64_t p11 = a1 * b1;
  const std::uint64_t mid = (p00 >> 32U) + (p01 & 0xffffffffULL) + (p10 & 0xffffffffULL);
  *lo = (p00 & 0xffffffffULL) | (mid << 32U);
  *hi = p11 + (p01 >> 32U) + (p10 >> 32U) + (mid >> 32U);
}

// `x * entry` rounded to a 64-bit significand, ties to even.
Wide scale_by(Wide x, const PowerOfTen& entry) noexcept {
  std::uint64_t hi = 0;
  std::uint64_t lo = 0;
  multiply_64(x.mant, entry.mant, &hi, &lo);
  const int exp = x.exp + entry.exp;
  if (hi == 0U) {
    return Wide{lo, exp};
  }
  const int drop = bit_length(hi);
  std::uint64_t mant = 0;
  std::uint64_t rest = 0;
  std::uint64_t half = 0;
  if (drop == 64) {
    mant = hi;
    rest = lo;
    half = 1ULL << 63U;
  } else {
    mant = (hi << (64 - drop)) | (lo >> drop);
    rest = lo & ((1ULL << drop) - 1U);
    half = 1ULL << (drop - 1);
  }
  int shift = drop;
  if (rest > half || (rest == half && (mant & 1U) != 0U)) {
    ++mant;
    if (mant == 0U) {
      mant = 1ULL << 63U;
      ++shift;
    }
  }
  return Wide{mant, exp + shift};
}

const PowerOfTen& power_of_ten(int sign, int k) noexcept {
  const int index = k <= 16 ? k - 1 : 16 + k / 16 - 2;
  return sign > 0 ? kPositivePowers[index] : kNegativePowers[index];
}

}  // namespace

double excel_scale_decimal(std::uint64_t n, long long e) noexcept {
  if (n == 0U) {
    return 0.0;
  }
  Wide x{n, 0};
  if (e != 0) {
    const int sign = e > 0 ? 1 : -1;
    const long long magnitude = e > 0 ? e : -e;
    const long long q = magnitude / 16;
    const int r = static_cast<int>(magnitude % 16);
    // Past 10^528 the value is far beyond the double range in either direction.
    if (q > 32) {
      return sign > 0 ? HUGE_VAL : 0.0;
    }
    if (r != 0) {
      x = scale_by(x, power_of_ten(sign, r));
    }
    if (q >= 1 && q <= 16) {
      x = scale_by(x, power_of_ten(sign, static_cast<int>(16 * q)));
    } else if (q > 16) {
      x = scale_by(x, power_of_ten(sign, static_cast<int>(16 * (q - 16))));
      x = scale_by(x, power_of_ten(sign, 256));
    }
  }
  const int bits = bit_length(x.mant);
  if (bits - 1 + x.exp < -1022) {
    return 0.0;
  }
  std::uint64_t mant = x.mant;
  int exp = x.exp;
  if (bits > 53) {
    const int drop = bits - 53;
    const std::uint64_t rest = mant & ((1ULL << drop) - 1U);
    const std::uint64_t half = 1ULL << (drop - 1);
    mant >>= drop;
    exp += drop;
    if (rest > half || (rest == half && (mant & 1U) != 0U)) {
      ++mant;
      if (mant == (1ULL << 53U)) {
        mant >>= 1U;
        ++exp;
      }
    }
  }
  if (bit_length(mant) - 1 + exp >= 1024) {
    return HUGE_VAL;
  }
  return std::ldexp(static_cast<double>(mant), exp);
}

ParsedDouble parse_double_prefix(std::string_view text) noexcept {
  ParsedDouble out;
  ScannedDecimal scan = scan_decimal(text, kMaxDigits);
  if (scan.consumed == 0) {
    return out;
  }
  out.consumed = scan.consumed;
  if (scan.count == 0) {
    out.value = scan.negative ? -0.0 : 0.0;
    return out;
  }
  if (scan.dropped_nonzero) {
    scan.digits[scan.count++] = '1';
    --scan.exponent;
  }
  const double magnitude =
      double_conversion::Strtod(double_conversion::Vector<const char>(scan.digits, scan.count), scan.exponent);
  out.value = scan.negative ? -magnitude : magnitude;
  out.out_of_range = std::isinf(magnitude) || magnitude < DBL_MIN;
  return out;
}

namespace {

constexpr int kExcelDigits = 15;

void store_excel_decimal(const ScannedDecimal& scan, double* out) noexcept {
  std::uint64_t n = 0;
  for (int i = 0; i < scan.count; ++i) {
    n = n * 10U + static_cast<std::uint64_t>(scan.digits[i] - '0');
  }
  const double magnitude = excel_scale_decimal(n, scan.exponent);
  *out = scan.negative ? -magnitude : magnitude;
}

}  // namespace

bool parse_excel_decimal(std::string_view text, double* out) noexcept {
  const ScannedDecimal scan = scan_decimal(text, kExcelDigits);
  if (scan.consumed == 0 || scan.consumed != text.size()) {
    return false;
  }
  store_excel_decimal(scan, out);
  return true;
}

bool parse_excel_literal(std::string_view text, double* out) noexcept {
  const ScannedDecimal scan = scan_decimal(text, kExcelDigits);
  if (scan.consumed == 0 || scan.consumed != text.size()) {
    return false;
  }
  if (scan.count > 0 && !excel_leading_exponent_in_range(static_cast<long long>(scan.exponent) + scan.count - 1)) {
    return false;
  }
  store_excel_decimal(scan, out);
  return true;
}

bool parse_double_exact(std::string_view text, double* out) noexcept {
  const ParsedDouble parsed = parse_double_prefix(text);
  if (parsed.consumed == 0 || parsed.consumed != text.size()) {
    return false;
  }
  *out = parsed.value;
  return true;
}

}  // namespace formulon

#if defined(FORMULON_WASM)
// Target of the WASM link's `--wrap=strtod` (cmake/FormulonWasm.cmake): the
// only remaining strtod callers are inside pugixml, and routing them here
// keeps libc's scanner out of the module.
extern "C" double __wrap_strtod(const char* str, char** endptr) {  // NOLINT(bugprone-reserved-identifier)
  const formulon::ParsedDouble parsed = formulon::parse_double_prefix(std::string_view(str, std::strlen(str)));
  if (endptr != nullptr) {
    *endptr = const_cast<char*>(str + parsed.consumed);  // NOLINT(cppcoreguidelines-pro-type-const-cast)
  }
  return parsed.value;
}
#endif
