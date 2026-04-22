// Copyright 2026 libraz. Licensed under the MIT License.
//
// Out-of-line members of the scalar `Value`. See `value.h` for the
// class contract.

#include "value.h"

#include <string>

#include "utils/expected.h"

namespace formulon {

double Value::as_number() const {
  FM_CHECK(kind_ == ValueKind::Number, "Value::as_number() on non-Number");
  return data_.number;
}

bool Value::as_boolean() const {
  FM_CHECK(kind_ == ValueKind::Bool, "Value::as_boolean() on non-Bool");
  return data_.boolean;
}

ErrorCode Value::as_error() const {
  FM_CHECK(kind_ == ValueKind::Error, "Value::as_error() on non-Error");
  return data_.error;
}

std::string_view Value::as_text() const {
  FM_CHECK(kind_ == ValueKind::Text, "Value::as_text() on non-Text");
  return data_.text;
}

std::string Value::debug_to_string() const {
  switch (kind_) {
    case ValueKind::Blank:
      return "Blank";
    case ValueKind::Number:
      return "Number(" + std::to_string(data_.number) + ")";
    case ValueKind::Bool:
      return std::string("Bool(") + (data_.boolean ? "true" : "false") + ")";
    case ValueKind::Error:
      return std::string("Error(") + display_name(data_.error) + ")";
    case ValueKind::Text: {
      // Quote the payload and escape internal `"` as `""` and backslash as
      // `\\` so the debug output is unambiguous when the text contains
      // either character.
      std::string out;
      out.reserve(data_.text.size() + 8);
      out.append("Text(\"");
      for (char c : data_.text) {
        if (c == '"') {
          out.append("\"\"");
        } else if (c == '\\') {
          out.append("\\\\");
        } else {
          out.push_back(c);
        }
      }
      out.append("\")");
      return out;
    }
    case ValueKind::Array:
      return "Array(<unimplemented>)";
    case ValueKind::Ref:
      return "Ref(<unimplemented>)";
    case ValueKind::Lambda:
      return "Lambda(<unimplemented>)";
  }
  return "Value(<invalid kind>)";
}

bool operator==(const Value& a, const Value& b) noexcept {
  if (a.kind_ != b.kind_) {
    return false;
  }
  switch (a.kind_) {
    case ValueKind::Blank:
      return true;
    case ValueKind::Number:
      // Use operator== on double so NaN != NaN, matching Excel's #NUM!
      // propagation model. The evaluator layer is responsible for turning
      // observed NaN into #NUM! where appropriate.
      return a.data_.number == b.data_.number;
    case ValueKind::Bool:
      return a.data_.boolean == b.data_.boolean;
    case ValueKind::Error:
      return a.data_.error == b.data_.error;
    case ValueKind::Text:
      return a.data_.text == b.data_.text;
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      // Not yet reachable: these kinds have no factory yet, so no
      // `Value` ever carries them. Treat as equal (same unimplemented
      // state) to keep `operator==` total.
      return true;
  }
  return false;
}

}  // namespace formulon
