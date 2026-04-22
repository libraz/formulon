// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the scalar coercion helpers declared in `coerce.h`.

#include "eval/coerce.h"

#include <cmath>
#include <cstdlib>
#include <cstring>
#include <string>
#include <string_view>

#include "utils/double_format.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

Expected<double, ErrorCode> coerce_to_number(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Number: {
      const double d = v.as_number();
      if (std::isnan(d) || std::isinf(d)) {
        return ErrorCode::Num;
      }
      return d;
    }
    case ValueKind::Bool:
      return v.as_boolean() ? 1.0 : 0.0;
    case ValueKind::Blank:
      return 0.0;
    case ValueKind::Text: {
      const std::string_view trimmed = strings::trim(v.as_text());
      if (trimmed.empty()) {
        // Excel treats an empty / whitespace-only text as 0 in arithmetic
        // contexts (matches "=\"\"+1" behaviour observed in 365).
        return 0.0;
      }
      // Defensive copy into a NUL-terminated stack buffer; std::strtod
      // requires a C string. 64 bytes is generous for any IEEE-754 double
      // including subnormals. Longer inputs are forwarded to the heap path.
      char stack_buf[64];
      char* heap_buf = nullptr;
      const std::size_t n = trimmed.size();
      char* buf = stack_buf;
      if (n + 1 > sizeof(stack_buf)) {
        heap_buf = static_cast<char*>(std::malloc(n + 1));
        if (heap_buf == nullptr) {
          return ErrorCode::Value;
        }
        buf = heap_buf;
      }
      std::memcpy(buf, trimmed.data(), n);
      buf[n] = '\0';
      char* end_ptr = nullptr;
      const double parsed = std::strtod(buf, &end_ptr);
      const bool ok = end_ptr == buf + n;
      if (heap_buf != nullptr) {
        std::free(heap_buf);
      }
      if (!ok) {
        return ErrorCode::Value;
      }
      if (std::isnan(parsed) || std::isinf(parsed)) {
        return ErrorCode::Num;
      }
      return parsed;
    }
    case ValueKind::Error:
      return v.as_error();
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

Expected<std::string, ErrorCode> coerce_to_text(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Number: {
      std::string out;
      format_double(out, v.as_number());
      return out;
    }
    case ValueKind::Bool:
      return std::string(v.as_boolean() ? "TRUE" : "FALSE");
    case ValueKind::Blank:
      return std::string();
    case ValueKind::Text:
      return std::string(v.as_text());
    case ValueKind::Error:
      return v.as_error();
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

Expected<bool, ErrorCode> coerce_to_bool(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Bool:
      return v.as_boolean();
    case ValueKind::Number: {
      const double d = v.as_number();
      if (std::isnan(d) || std::isinf(d)) {
        return ErrorCode::Num;
      }
      return d != 0.0;
    }
    case ValueKind::Blank:
      return false;
    case ValueKind::Text: {
      // Excel accepts "TRUE" / "FALSE" case-insensitively in boolean
      // contexts; trim whitespace defensively to match the strtod path.
      const std::string_view trimmed = strings::trim(v.as_text());
      if (strings::case_insensitive_eq(trimmed, std::string_view("TRUE"))) {
        return true;
      }
      if (strings::case_insensitive_eq(trimmed, std::string_view("FALSE"))) {
        return false;
      }
      return ErrorCode::Value;
    }
    case ValueKind::Error:
      return v.as_error();
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

}  // namespace eval
}  // namespace formulon
