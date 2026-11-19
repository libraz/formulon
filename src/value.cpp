//
// Out-of-line members of the scalar `Value`. See `value.h` for the
// class contract.

#include "value.h"

#include <cassert>
#include <string>

#include "eval/lambda_value.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/number_text.h"

namespace formulon {

Value adopt_text_into(Arena& arena, Value value) noexcept {
  if (!value.is_text()) {
    return value;
  }
  return Value::text(arena.intern(value.as_text()));
}

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

const eval::LambdaValue* Value::as_lambda() const {
  FM_CHECK(kind_ == ValueKind::Lambda, "Value::as_lambda() on non-Lambda");
  return data_.lambda;
}

RefForm Value::ref_form() const {
  FM_CHECK(kind_ == ValueKind::Ref, "Value::ref_form() on non-Ref");
  return static_cast<RefForm>(aux_ & ~(kRefOutOfLine | kRefDynamic));
}

RefView Value::as_ref() const {
  FM_CHECK(kind_ == ValueKind::Ref, "Value::as_ref() on non-Ref");
  const RefForm form = static_cast<RefForm>(aux_ & ~(kRefOutOfLine | kRefDynamic));
  if ((aux_ & kRefOutOfLine) == 0) {
    return RefView{&data_.area, nullptr, 1, 0, 0, 0, form, nullptr};
  }
  const RefValue* rv = data_.ref;
  return RefView{rv->areas, rv->element_books, rv->area_count, rv->rows, rv->cols, rv->book, form, rv->element_values};
}

PicturePayload Value::as_picture() const {
  FM_CHECK(kind_ == ValueKind::Picture, "Value::as_picture() on non-Picture");
  return data_.picture;
}

std::uint8_t Value::picture_origin() const {
  FM_CHECK(kind_ == ValueKind::Picture, "Value::picture_origin() on non-Picture");
  return aux_;
}

Value contain_escape(Value v) noexcept {
  assert(!escapes_boundary(v) && "a reference value cannot leave the evaluator");
  return escapes_boundary(v) ? Value::error(ErrorCode::Value) : v;
}

Value make_ref(Arena& arena, RefForm form, std::uint16_t book, const RefArea* areas, std::uint32_t area_count,
               std::uint32_t rows, std::uint32_t cols, const std::uint16_t* element_books,
               const Value* element_values) {
  const bool single = form == RefForm::kArea || form == RefForm::kSpan3D;
  FM_CHECK(!single || area_count == 1, "make_ref: kArea / kSpan3D takes exactly one area");
  if (single && book == 0) {
    return Value::ref_area(areas[0]);
  }
  RefValue* rv = arena.create<RefValue>();
  RefArea* copy = arena.create_array<RefArea>(area_count);
  if (rv == nullptr || (copy == nullptr && area_count != 0)) {
    return Value::error(ErrorCode::Calc);
  }
  for (std::uint32_t i = 0; i < area_count; ++i) {
    copy[i] = areas[i];
  }
  const std::uint16_t* books = nullptr;
  if (element_books != nullptr) {
    const std::size_t n = static_cast<std::size_t>(rows) * cols;
    std::uint16_t* b = arena.create_array<std::uint16_t>(n);
    if (b == nullptr && n != 0) {
      return Value::error(ErrorCode::Calc);
    }
    for (std::size_t i = 0; i < n; ++i) {
      b[i] = element_books[i];
    }
    books = b;
  }
  const Value* values = nullptr;
  if (element_values != nullptr) {
    Value* v = arena.create_array<Value>(area_count);
    if (v == nullptr && area_count != 0) {
      return Value::error(ErrorCode::Calc);
    }
    for (std::uint32_t i = 0; i < area_count; ++i) {
      v[i] = element_values[i];
    }
    values = v;
  }
  const bool span = single && areas[0].sheet_first != areas[0].sheet_last;
  rv->element_values = values;
  rv->areas = copy;
  rv->element_books = books;
  rv->area_count = area_count;
  rv->rows = rows;
  rv->cols = cols;
  rv->book = book;
  rv->form = single ? (span ? RefForm::kSpan3D : RefForm::kArea) : form;
  return Value::ref_out_of_line(rv);
}

std::string Value::debug_to_string() const {
  switch (kind_) {
    case ValueKind::Blank:
      return "Blank";
    case ValueKind::Number:
      return "Number(" + fixed_string(data_.number, 6) + ")";
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
    case ValueKind::Lambda: {
      // Show only the parameter count: the body is an arbitrary AST
      // sub-tree and the captured environment is unbounded in size, so
      // neither belongs in a debug-string. The arity is the load-bearing
      // identifier for closure-shape bugs (arity-mismatch failures, the
      // capture-vs-call boundary, etc.).
      std::string out;
      out.reserve(20);
      out.append("Lambda(");
      out.append(std::to_string(data_.lambda != nullptr ? data_.lambda->param_count : 0));
      out.append(" params)");
      return out;
    }
    case ValueKind::Picture:
      return "picture(" + std::to_string(data_.picture.picture_id) + "," + std::to_string(data_.picture.media_id) +
             "," + std::to_string(aux_) + ")";
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
    case ValueKind::Lambda:
      // Lambda equality is identity-only: closures with identical AST,
      // params, and captured env may still represent distinct expression
      // sites and Excel's surface contract does not give a meaningful
      // value-level equality on closures. Pointer equality is total,
      // cheap, and matches how `==` is consumed today (cache reuse,
      // memo-table hits) where two values reach the comparator only when
      // they came from the same allocation site.
      return a.data_.lambda == b.data_.lambda;
    case ValueKind::Picture:
      return a.data_.picture.picture_id == b.data_.picture.picture_id &&
             a.data_.picture.media_id == b.data_.picture.media_id && a.aux_ == b.aux_;
    case ValueKind::Ref:
      // Not yet reachable: this kind has no factory yet, so no `Value`
      // ever carries it. Treat as equal (same unimplemented state) to
      // keep `operator==` total.
      return true;
  }
  return false;
}

}  // namespace formulon
