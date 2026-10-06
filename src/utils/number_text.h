//
// snprintf-compatible number formatting that does not link libc's printf.
//
// Each function writes NUL-terminated text into `buf` (capacity `cap`) and
// returns the length the full text has, like `snprintf`: a result >= `cap`
// means the text was truncated. Output is byte-identical to the printf
// conversion named on each declaration, including the exact-tie rounding
// (ties to even on the exact binary value) that double-conversion alone
// does not give.
//
// Engine code must use these instead of `snprintf`: on the WASM target the
// only remaining `snprintf` caller is pugixml, routed to `__wrap_snprintf`
// in number_text.cpp.

#ifndef FORMULON_UTILS_NUMBER_TEXT_H_
#define FORMULON_UTILS_NUMBER_TEXT_H_

#include <cstddef>
#include <string>

namespace formulon {

/// Equivalent of `snprintf(buf, cap, "%.*f", frac_digits, v)`.
int format_fixed(char* buf, std::size_t cap, double v, int frac_digits);

/// `format_fixed` into a string; `fixed_string(v, 6)` replaces `std::to_string(double)`,
/// which libc++ implements through `snprintf("%f")`.
std::string fixed_string(double v, int frac_digits);

/// Equivalent of `snprintf(buf, cap, "%.*e", frac_digits, v)`.
int format_exponential(char* buf, std::size_t cap, double v, int frac_digits);

/// Equivalent of `snprintf(buf, cap, "%.*g", precision, v)`; a precision of 0
/// is treated as 1.
int format_general(char* buf, std::size_t cap, double v, int precision);

/// Equivalent of `"%d"` / `"%lld"` (min_width 0) or `"%0<min_width>d"`.
int format_signed(char* buf, std::size_t cap, long long v, int min_width = 0);

/// Equivalent of `"%u"` / `"%llu"` (min_width 0) or `"%0<min_width>u"`.
int format_unsigned(char* buf, std::size_t cap, unsigned long long v, int min_width = 0);

/// Equivalent of `"%0<min_width>X"` (`upper`) or `"%0<min_width>x"`.
int format_hex(char* buf, std::size_t cap, unsigned long long v, int min_width, bool upper);

}  // namespace formulon

#endif  // FORMULON_UTILS_NUMBER_TEXT_H_
