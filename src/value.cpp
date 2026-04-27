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

const ArrayValue* Value::as_array() const {
  FM_CHECK(kind_ == ValueKind::Array, "Value::as_array() on non-Array");
  return data_.array;
}

std::uint32_t Value::as_array_rows() const {
  FM_CHECK(kind_ == ValueKind::Array, "Value::as_array_rows() on non-Array");
  return data_.array->rows;
}

std::uint32_t Value::as_array_cols() const {
  FM_CHECK(kind_ == ValueKind::Array, "Value::as_array_cols() on non-Array");
  return data_.array->cols;
}

const Value* Value::as_array_cells() const {
  FM_CHECK(kind_ == ValueKind::Array, "Value::as_array_cells() on non-Array");
  return data_.array->cells;
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
    case ValueKind::Array: {
      // Show only the shape: cell contents could be arbitrarily long, and
      // the shape is the load-bearing identifier for debugging shape /
      // broadcasting bugs.
      std::string out;
      out.reserve(24);
      out.append("Array(");
      out.append(std::to_string(data_.array != nullptr ? data_.array->rows : 0));
      out.push_back('x');
      out.append(std::to_string(data_.array != nullptr ? data_.array->cols : 0));
      out.push_back(')');
      return out;
    }
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
    case ValueKind::Array: {
      const ArrayValue* lhs = a.data_.array;
      const ArrayValue* rhs = b.data_.array;
      if (lhs == rhs) {
        return true;
      }
      if (lhs == nullptr || rhs == nullptr) {
        return false;
      }
      if (lhs->rows != rhs->rows || lhs->cols != rhs->cols) {
        return false;
      }
      const std::size_t n = static_cast<std::size_t>(lhs->rows) * static_cast<std::size_t>(lhs->cols);
      for (std::size_t i = 0; i < n; ++i) {
        if (!(lhs->cells[i] == rhs->cells[i])) {
          return false;
        }
      }
      return true;
    }
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
