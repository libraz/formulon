// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "c_api/internal/value_marshal.h"

#include <cstdint>
#include <string>
#include <type_traits>

#include "c_api/formulon_c.h"
#include "value.h"

namespace formulon {
namespace c_api {
namespace internal {

// Triple-identity check: FmValueTag <-> FM_VAL_* <-> ValueKind. The chain is
// load-bearing because bindings rely on it to `static_cast` between the three
// types without translation tables.
static_assert(static_cast<std::int32_t>(FmValueTag::Blank) == FM_VAL_BLANK,
              "FmValueTag::Blank must match FM_VAL_BLANK");
static_assert(static_cast<std::int32_t>(FmValueTag::Number) == FM_VAL_NUMBER,
              "FmValueTag::Number must match FM_VAL_NUMBER");
static_assert(static_cast<std::int32_t>(FmValueTag::Bool) == FM_VAL_BOOL, "FmValueTag::Bool must match FM_VAL_BOOL");
static_assert(static_cast<std::int32_t>(FmValueTag::Text) == FM_VAL_TEXT, "FmValueTag::Text must match FM_VAL_TEXT");
static_assert(static_cast<std::int32_t>(FmValueTag::Error) == FM_VAL_ERROR,
              "FmValueTag::Error must match FM_VAL_ERROR");
static_assert(static_cast<std::int32_t>(FmValueTag::Array) == FM_VAL_ARRAY,
              "FmValueTag::Array must match FM_VAL_ARRAY");
static_assert(static_cast<std::int32_t>(FmValueTag::Ref) == FM_VAL_REF, "FmValueTag::Ref must match FM_VAL_REF");
static_assert(static_cast<std::int32_t>(FmValueTag::Lambda) == FM_VAL_LAMBDA,
              "FmValueTag::Lambda must match FM_VAL_LAMBDA");

// And the other half of the identity: FM_VAL_* matches ValueKind ordinals.
static_assert(FM_VAL_BLANK == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Blank),
              "FM_VAL_BLANK must match ValueKind::Blank");
static_assert(FM_VAL_NUMBER == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Number),
              "FM_VAL_NUMBER must match ValueKind::Number");
static_assert(FM_VAL_BOOL == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Bool),
              "FM_VAL_BOOL must match ValueKind::Bool");
static_assert(FM_VAL_TEXT == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Text),
              "FM_VAL_TEXT must match ValueKind::Text");
static_assert(FM_VAL_ERROR == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Error),
              "FM_VAL_ERROR must match ValueKind::Error");
static_assert(FM_VAL_ARRAY == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Array),
              "FM_VAL_ARRAY must match ValueKind::Array");
static_assert(FM_VAL_REF == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Ref),
              "FM_VAL_REF must match ValueKind::Ref");
static_assert(FM_VAL_LAMBDA == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Lambda),
              "FM_VAL_LAMBDA must match ValueKind::Lambda");

FmValueTag value_tag(const Value& v) noexcept {
  // Both sides are validated as identical by the static_asserts above, so
  // a single cast is safe and total over every ValueKind enumerator.
  return static_cast<FmValueTag>(static_cast<std::underlying_type_t<ValueKind>>(v.kind()));
}

std::int32_t error_to_fm_code(ErrorCode ec) noexcept {
  // Re-export of formulon::ooxml_code: single source of truth lives in
  // value.h's kErrorTable. The return type is widened to int32_t to match
  // the fm_value_t::u.error_code field width at the C ABI.
  return static_cast<std::int32_t>(formulon::ooxml_code(ec));
}

const char* error_to_fm_text(ErrorCode ec) noexcept {
  // Re-export of formulon::display_name. The returned pointer is a static
  // string literal with program lifetime.
  return formulon::display_name(ec);
}

ArrayShape inspect_array(const Value& v) noexcept {
  if (v.kind() != ValueKind::Array) {
    return ArrayShape{0, 0, true};
  }
  const ArrayValue* a = v.as_array();
  if (a == nullptr) {
    return ArrayShape{0, 0, true};
  }
  const std::uint32_t rows = a->rows;
  const std::uint32_t cols = a->cols;
  return ArrayShape{rows, cols, rows == 0 || cols == 0};
}

double value_as_number_or_zero(const Value& v) noexcept {
  return v.is_number() ? v.as_number() : 0.0;
}

bool value_as_bool_or_false(const Value& v) noexcept {
  return v.is_boolean() && v.as_boolean();
}

const std::string* value_as_text_or_null(const Value& v, std::string& storage) noexcept {
  // Value::Text holds a non-owning string_view borrowed from arena
  // storage. Bindings need a NUL-terminated std::string for JS engines
  // (Napi::String, embind val) and have to copy in every case, so this
  // helper folds the copy into `storage` and returns a pointer-or-null
  // discriminator. Returning a string_view would not let callers tell
  // "no text" from "empty text".
  if (!v.is_text()) {
    return nullptr;
  }
  const std::string_view sv = v.as_text();
  storage.assign(sv.data(), sv.size());
  return &storage;
}

}  // namespace internal
}  // namespace c_api
}  // namespace formulon
