//
// Internal header — include only from `src/eval/builtins/complex_num*`.
//
// The complex-number value type shared by the complex-number built-ins, and
// the parsing / formatting of its text form ("a+bi" / "a+bj").

#ifndef FORMULON_EVAL_BUILTINS_COMPLEX_NUM_TEXT_H_
#define FORMULON_EVAL_BUILTINS_COMPLEX_NUM_TEXT_H_

#include <string>

#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace complex_detail {

/// A parsed complex number: (real, imag, suffix). `suffix` is always 'i'
/// or 'j'; `'i'` is the default when the input had no imaginary part.
struct Complex {
  double re;
  double im;
  char suffix;  // 'i' or 'j'
};

/// Front-door parser: accepts Text / Number / Bool / Blank / Error and
/// propagates the Excel error code.
Expected<Complex, ErrorCode> parse_complex_value(const Value& v);

/// Formats a complex triple into Excel's canonical string form.
std::string format_complex(double re, double im, char suffix);

/// Interns and returns a Text value for the complex triple (#NUM! when a part
/// is not finite).
Value text_complex(Complex z, Arena& arena);

}  // namespace complex_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_COMPLEX_NUM_TEXT_H_
