//
// Shared definition of `read_int_arg` for the text builtin family. Hoisted
// out of `text.cpp` into its own TU because the DBCS family (`text_dbcs.cpp`)
// and the modern TEXTBEFORE/TEXTAFTER family (`text_modern.cpp`) both
// reference it; keeping a single definition here avoids ODR violations.

#include "eval/builtins/text_detail.h"

#include <cmath>
#include <limits>
#include <utility>

#include "eval/coerce.h"
#include "eval/utf8_length.h"

namespace formulon {
namespace eval {
namespace text_detail {

Expected<int, ErrorCode> read_int_arg(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  // Converting a double outside `int`'s range is undefined, and the two
  // architectures disagree on what they produce: x86-64 yields INT_MIN
  // while WASM's `--enable-nontrapping-float-to-int` saturates to
  // INT_MAX. A count like `1E+15` (a routine LEFT/MID/RIGHT/REPT/
  // SUBSTITUTE/REPLACE argument, not an attack input) would therefore
  // read as a huge negative on one target and a huge positive on the
  // other, so `LEFT("text",1E+15)` returns `#VALUE!` on native and
  // `"text"` on WASM. Every caller already treats "past the end of the
  // text" the same way it treats "the whole text", so saturating before
  // the cast carries the same meaning as the real magnitude and keeps
  // the result identical across targets. Mirrors `read_digits` in
  // `eval/builtins/math.cpp`.
  const double truncated = std::trunc(d);
  constexpr double kIntMax = 2147483647.0;
  constexpr double kIntMin = -2147483648.0;
  if (truncated >= kIntMax) {
    return std::numeric_limits<int>::max();
  }
  if (truncated <= kIntMin) {
    return std::numeric_limits<int>::min();
  }
  return static_cast<int>(truncated);
}

Expected<int, ErrorCode> read_optional_int_arg(const Value* args, std::uint32_t arity, std::uint32_t index,
                                               int default_value) {
  if (arity <= index) {
    return default_value;
  }
  return read_int_arg(args[index]);
}

bool read_search_args(const Value* args, std::uint32_t arity, SearchUnit unit, SearchArgs* out, Value* out_result) {
  auto needle = coerce_to_text(args[0]);
  if (!needle) {
    *out_result = Value::error(needle.error());
    return false;
  }
  auto haystack = coerce_to_text(args[1]);
  if (!haystack) {
    *out_result = Value::error(haystack.error());
    return false;
  }
  auto parsed = read_optional_int_arg(args, arity, 2u, 1);
  if (!parsed) {
    *out_result = Value::error(parsed.error());
    return false;
  }
  const int start = parsed.value();
  const std::uint64_t total = (unit == SearchUnit::DbcsByte)
                                  ? bytes_in_jajp(haystack.value())
                                  : static_cast<std::uint64_t>(utf16_units_in(haystack.value()));
  if (start < 1 || static_cast<std::uint64_t>(start) > total + 1) {
    *out_result = Value::error(ErrorCode::Value);
    return false;
  }
  if (needle.value().empty()) {
    *out_result = Value::number(static_cast<double>(start));
    return false;
  }
  out->needle = std::move(needle.value());
  out->haystack = std::move(haystack.value());
  out->start = start;
  return true;
}

}  // namespace text_detail
}  // namespace eval
}  // namespace formulon
