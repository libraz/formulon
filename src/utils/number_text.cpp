//
// Implementation of the printf-free number formatters. See `number_text.h`.
//
// Digits come from double-conversion, which rounds exact decimal ties up
// where printf rounds them to even. An exact tie is detected with integer
// arithmetic on the binary significand and resolved here; every other value
// is correctly rounded by the library already.

#include "utils/number_text.h"

#include <cmath>
#include <cstdarg>
#include <cstdint>
#include <cstring>

#include "double-conversion/double-conversion.h"

namespace formulon {
namespace {

using double_conversion::DoubleToStringConverter;

// A double has at most 767 significant digits and 1074 fractional places;
// requests beyond these only add zeros, which the layout pads itself.
constexpr int kMaxSignificant = 800;
constexpr int kMaxFraction = 1100;
// 309 integer digits + kMaxFraction places + one tie digit + NUL.
constexpr int kDigitBufferSize = 1500;

/// Bounded writer with snprintf's truncate-and-count behaviour.
class Out {
 public:
  Out(char* buf, std::size_t cap) : buf_(buf), cap_(cap) {}

  void put(char c) {
    if (n_ + 1 < cap_) {
      buf_[n_] = c;
    }
    ++n_;
  }

  void put_n(char c, int count) {
    for (int i = 0; i < count; ++i) {
      put(c);
    }
  }

  int finish() {
    if (cap_ > 0) {
      buf_[n_ < cap_ ? n_ : cap_ - 1] = '\0';
    }
    return static_cast<int>(n_);
  }

 private:
  char* buf_;
  std::size_t cap_;
  std::size_t n_ = 0;
};

/// Decimal digits of a value: 0.D1D2... x 10^point. Digits past `len` are zero.
struct Digits {
  char d[kDigitBufferSize];
  int len = 0;
  int point = 0;

  char at(int idx) const { return (idx >= 0 && idx < len) ? d[idx] : '0'; }
};

/// Splits |v| into m * 2^e (m a 53-bit integer) and returns its trailing zero
/// bit count and odd part.
void decompose(double v, int* z, std::uint64_t* odd, int* e) {
  std::uint64_t bits;
  std::memcpy(&bits, &v, sizeof(bits));
  const int biased = static_cast<int>((bits >> 52) & 0x7FFU);
  std::uint64_t m = bits & ((std::uint64_t{1} << 52) - 1U);
  if (biased == 0) {
    *e = -1074;
  } else {
    m |= std::uint64_t{1} << 52;
    *e = biased - 1075;
  }
  int tz = 0;
  while ((m & 1U) == 0U) {
    m >>= 1U;
    ++tz;
  }
  *z = tz;
  *odd = m;
}

/// True iff |v| is exactly halfway between two multiples of 10^-t, i.e.
/// 2 * |v| * 10^t is an odd integer. `v` must be finite and nonzero.
bool is_tie_at(double v, int t) {
  int z;
  int e;
  std::uint64_t odd;
  decompose(v, &z, &odd, &e);
  if (t >= 0) {
    return z + e + t + 1 == 0;
  }
  const int s = -t;
  if (z + e - s + 1 != 0 || s > 22) {
    return false;
  }
  std::uint64_t p5 = 1;
  for (int i = 0; i < s; ++i) {
    p5 *= 5U;
  }
  return odd % p5 == 0U;
}

/// Rounds up in place by one unit in the last of the first `keep` digits.
void increment(Digits* g, int keep) {
  int i = keep - 1;
  while (i >= 0 && g->d[i] == '9') {
    --i;
  }
  if (i < 0) {
    g->d[0] = '1';
    g->len = 1;
    ++g->point;
    return;
  }
  ++g->d[i];
  g->len = i + 1;
}

/// Drops digits past `keep`, rounding half to even when the dropped part is
/// exactly one half (the caller has established that it is).
void round_tie_even(Digits* g, int keep) {
  g->len = keep;
  if (keep > 0 && ((g->d[keep - 1] - '0') & 1) != 0) {
    increment(g, keep);
  }
}

void to_ascii(double v, DoubleToStringConverter::DtoaMode mode, int requested, Digits* g) {
  bool sign;
  DoubleToStringConverter::DoubleToAscii(v, mode, requested, g->d, kDigitBufferSize, &sign, &g->len, &g->point);
}

/// Digits of |v| rounded to `frac` places after the point (printf `%f`).
void fixed_digits(double v, int frac, Digits* g) {
  if (v == 0.0) {
    g->len = 0;
    g->point = 0;
    return;
  }
  if (frac > kMaxFraction) {
    frac = kMaxFraction;
  }
  if (is_tie_at(v, frac)) {
    // The exact expansion ends in '5' one place further: generate it whole.
    to_ascii(v, DoubleToStringConverter::FIXED, frac + 1, g);
    round_tie_even(g, g->len - 1);
    return;
  }
  to_ascii(v, DoubleToStringConverter::FIXED, frac, g);
}

/// Digits of |v| rounded to `n` significant digits (printf `%e` / `%g`).
void precision_digits(double v, int n, Digits* g) {
  if (v == 0.0) {
    g->d[0] = '0';
    g->len = 1;
    g->point = 1;
    return;
  }
  if (n > kMaxSignificant) {
    n = kMaxSignificant;
  }
  // One extra digit settles the rounding unless that digit is a '5'.
  to_ascii(v, DoubleToStringConverter::PRECISION, n + 1, g);
  if (g->len <= n) {
    return;
  }
  if (g->d[n] != '5') {
    if (g->d[n] > '5') {
      increment(g, n);
    } else {
      g->len = n;
    }
    return;
  }
  if (is_tie_at(v, n - g->point)) {
    round_tie_even(g, n);
    return;
  }
  to_ascii(v, DoubleToStringConverter::PRECISION, n, g);
}

void put_special(Out* out, double v) {
  if (std::signbit(v)) {
    out->put('-');
  }
  const char* text = std::isnan(v) ? "nan" : "inf";
  while (*text != '\0') {
    out->put(*text++);
  }
}

/// Writes the digits as a fixed-point number with exactly `frac` places.
void put_fixed(Out* out, const Digits& g, int frac) {
  if (g.point > 0) {
    for (int i = 0; i < g.point; ++i) {
      out->put(g.at(i));
    }
  } else {
    out->put('0');
  }
  if (frac > 0) {
    out->put('.');
    for (int j = 0; j < frac; ++j) {
      out->put(g.at(g.point + j));
    }
  }
}

/// Writes `D.DDD` + `e+XX` using the first `frac + 1` digits.
void put_exponential(Out* out, const Digits& g, int frac) {
  out->put(g.at(0));
  if (frac > 0) {
    out->put('.');
    for (int j = 1; j <= frac; ++j) {
      out->put(g.at(j));
    }
  }
  const int exp = g.point - 1;
  const int mag = exp < 0 ? -exp : exp;
  out->put('e');
  out->put(exp < 0 ? '-' : '+');
  if (mag < 10) {
    out->put('0');
  }
  char tmp[8];
  int n = 0;
  for (int m = mag; m > 0; m /= 10) {
    tmp[n++] = static_cast<char>('0' + m % 10);
  }
  while (n > 0) {
    out->put(tmp[--n]);
  }
  if (mag == 0) {
    out->put('0');
  }
}

int format_unsigned_base(char* buf, std::size_t cap, unsigned long long v, int min_width, bool negative, unsigned base,
                         bool upper) {
  char tmp[24];
  int n = 0;
  do {
    const unsigned digit = static_cast<unsigned>(v % base);
    tmp[n++] = static_cast<char>(digit < 10U ? '0' + digit : (upper ? 'A' : 'a') + (digit - 10U));
    v /= base;
  } while (v != 0U);
  Out out(buf, cap);
  if (negative) {
    out.put('-');
  }
  const int used = n + (negative ? 1 : 0);
  if (min_width > used) {
    out.put_n('0', min_width - used);
  }
  while (n > 0) {
    out.put(tmp[--n]);
  }
  return out.finish();
}

}  // namespace

int format_fixed(char* buf, std::size_t cap, double v, int frac_digits) {
  Out out(buf, cap);
  if (!std::isfinite(v)) {
    put_special(&out, v);
    return out.finish();
  }
  if (frac_digits < 0) {
    frac_digits = 0;
  }
  if (std::signbit(v)) {
    out.put('-');
  }
  Digits g;
  fixed_digits(std::fabs(v), frac_digits, &g);
  put_fixed(&out, g, frac_digits);
  return out.finish();
}

std::string fixed_string(double v, int frac_digits) {
  std::string out;
  out.resize(static_cast<std::size_t>(format_fixed(nullptr, 0, v, frac_digits)));
  format_fixed(&out[0], out.size() + 1, v, frac_digits);
  return out;
}

int format_exponential(char* buf, std::size_t cap, double v, int frac_digits) {
  Out out(buf, cap);
  if (!std::isfinite(v)) {
    put_special(&out, v);
    return out.finish();
  }
  if (frac_digits < 0) {
    frac_digits = 0;
  }
  if (std::signbit(v)) {
    out.put('-');
  }
  Digits g;
  precision_digits(std::fabs(v), frac_digits + 1, &g);
  put_exponential(&out, g, frac_digits);
  return out.finish();
}

int format_general(char* buf, std::size_t cap, double v, int precision) {
  Out out(buf, cap);
  if (!std::isfinite(v)) {
    put_special(&out, v);
    return out.finish();
  }
  if (precision < 1) {
    precision = 1;
  }
  if (std::signbit(v)) {
    out.put('-');
  }
  Digits g;
  precision_digits(std::fabs(v), precision, &g);
  // Significant digits without trailing zeros: `%g` drops them.
  int sig = g.len < precision ? g.len : precision;
  while (sig > 1 && g.d[sig - 1] == '0') {
    --sig;
  }
  const int x = g.point - 1;
  if (precision > x && x >= -4) {
    const int frac = sig > g.point ? sig - g.point : 0;
    put_fixed(&out, g, frac);
  } else {
    put_exponential(&out, g, sig - 1);
  }
  return out.finish();
}

int format_signed(char* buf, std::size_t cap, long long v, int min_width) {
  const bool negative = v < 0;
  const unsigned long long mag =
      negative ? 0ULL - static_cast<unsigned long long>(v) : static_cast<unsigned long long>(v);
  return format_unsigned_base(buf, cap, mag, min_width, negative, 10U, false);
}

int format_unsigned(char* buf, std::size_t cap, unsigned long long v, int min_width) {
  return format_unsigned_base(buf, cap, v, min_width, false, 10U, false);
}

int format_hex(char* buf, std::size_t cap, unsigned long long v, int min_width, bool upper) {
  return format_unsigned_base(buf, cap, v, min_width, false, 16U, upper);
}

}  // namespace formulon

#if defined(FORMULON_WASM)
// Target of the WASM link's `--wrap=snprintf` (cmake/FormulonWasm.cmake). After
// the engine moved to number_text, pugixml's XPath number-to-string is the
// only snprintf caller in the module; it uses exactly "%.*e" and "%.*g". A new
// snprintf call in engine code would silently land here and fail with -1:
// engine code must use the formatters above instead.
extern "C" int __wrap_snprintf(char* buf, std::size_t cap, const char* fmt,
                               ...) {  // NOLINT(bugprone-reserved-identifier)
  int result = -1;
  const bool is_e = std::strcmp(fmt, "%.*e") == 0;
  if (is_e || std::strcmp(fmt, "%.*g") == 0) {
    va_list args;
    va_start(args, fmt);
    const int precision = va_arg(args, int);
    const double value = va_arg(args, double);
    va_end(args);
    result = is_e ? formulon::format_exponential(buf, cap, value, precision)
                  : formulon::format_general(buf, cap, value, precision);
  } else if (cap > 0) {
    buf[0] = '\0';
  }
  return result;
}
#endif
