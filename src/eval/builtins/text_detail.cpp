//
// Shared definition of `read_int_arg` for the text builtin family. Hoisted
// out of `text.cpp` into its own TU because the DBCS family (`text_dbcs.cpp`)
// and the modern TEXTBEFORE/TEXTAFTER family (`text_modern.cpp`) both
// reference it; keeping a single definition here avoids ODR violations.

#include "eval/builtins/text_detail.h"

#include <cmath>
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
  return static_cast<int>(std::trunc(d));
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
