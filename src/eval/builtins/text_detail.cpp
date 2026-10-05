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
#include "eval/wildcard.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

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

std::size_t find_folded(const std::string& haystack, const std::string& needle, std::size_t start_byte,
                        SearchUnit unit) {
  const std::string lowered_haystack = to_lower_ascii(haystack);
  const std::string lowered_needle = to_lower_ascii(needle);
  // Fast path: a pattern with no `*`, `?` or `~` is a plain substring search.
  // A bare `~` still needs the wildcard path because `~?` / `~*` unescape.
  if (lowered_needle.find_first_of("*?~") == std::string::npos) {
    return lowered_haystack.find(lowered_needle, start_byte);
  }
  // The wildcard matchers report offsets relative to the scanned suffix.
  const std::string_view suffix = std::string_view(lowered_haystack).substr(start_byte);
  const std::size_t rel =
      unit == SearchUnit::DbcsByte ? wildcard_find_dbcs(lowered_needle, suffix) : wildcard_find(lowered_needle, suffix);
  if (rel == std::string_view::npos) {
    return std::string::npos;
  }
  return start_byte + rel;
}

Expected<TextWindowArgs, ErrorCode> read_text_window_args(const Value* args, std::uint32_t arity) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return text.error();
  }
  auto start = read_int_arg(args[1]);
  if (!start) {
    return start.error();
  }
  auto count = read_int_arg(args[2]);
  if (!count) {
    return count.error();
  }
  std::string new_text;
  if (arity > 3) {
    auto coerced = coerce_to_text(args[3]);
    if (!coerced) {
      return coerced.error();
    }
    new_text = std::move(coerced.value());
  }
  if (start.value() < 1 || count.value() < 0) {
    return ErrorCode::Value;
  }
  return TextWindowArgs{std::move(text.value()), start.value(), count.value(), std::move(new_text)};
}

}  // namespace text_detail
}  // namespace eval
}  // namespace formulon
