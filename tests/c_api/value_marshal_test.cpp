// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for the shared marshalling leaves in
// `src/c_api/internal/value_marshal.{h,cpp}`. The helpers are pure
// inspections over `formulon::Value` and `formulon::ErrorCode`; this file
// exercises every variant and every Excel error code that the bindings
// reach for.

#include "c_api/internal/value_marshal.h"

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace c_api {
namespace internal {
namespace {

// Compile-time parity check between FmValueTag and the C ABI tag. The
// production code already asserts this; we restate it here so a deliberate
// edit that ignored the parity is caught at test-build time.
static_assert(static_cast<std::int32_t>(FmValueTag::Blank) == FM_VAL_BLANK, "Blank parity");
static_assert(static_cast<std::int32_t>(FmValueTag::Number) == FM_VAL_NUMBER, "Number parity");
static_assert(static_cast<std::int32_t>(FmValueTag::Bool) == FM_VAL_BOOL, "Bool parity");
static_assert(static_cast<std::int32_t>(FmValueTag::Text) == FM_VAL_TEXT, "Text parity");
static_assert(static_cast<std::int32_t>(FmValueTag::Error) == FM_VAL_ERROR, "Error parity");
static_assert(static_cast<std::int32_t>(FmValueTag::Array) == FM_VAL_ARRAY, "Array parity");
static_assert(static_cast<std::int32_t>(FmValueTag::Ref) == FM_VAL_REF, "Ref parity");
static_assert(static_cast<std::int32_t>(FmValueTag::Lambda) == FM_VAL_LAMBDA, "Lambda parity");

// Build a (rows x cols) ArrayValue inside `arena` populated with sequential
// Number cells. Used to exercise inspect_array on shaped payloads without
// pulling in the array-test helpers.
const ArrayValue* MakeNumberedArray(Arena* arena, std::uint32_t rows, std::uint32_t cols) {
  const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  Value* cells = arena->create_array<Value>(n);
  for (std::size_t i = 0; i < n; ++i) {
    cells[i] = Value::number(static_cast<double>(i));
  }
  ArrayValue* arr = arena->create<ArrayValue>();
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = cells;
  return arr;
}

// ---------------------------------------------------------------------------
// value_tag: one assertion per ValueKind
// ---------------------------------------------------------------------------

TEST(ValueMarshalTag, BlankMapsToBlank) {
  EXPECT_EQ(FmValueTag::Blank, value_tag(Value::blank()));
}

TEST(ValueMarshalTag, NumberMapsToNumber) {
  EXPECT_EQ(FmValueTag::Number, value_tag(Value::number(42.0)));
}

TEST(ValueMarshalTag, BoolMapsToBool) {
  EXPECT_EQ(FmValueTag::Bool, value_tag(Value::boolean(true)));
  EXPECT_EQ(FmValueTag::Bool, value_tag(Value::boolean(false)));
}

TEST(ValueMarshalTag, TextMapsToText) {
  const std::string s = "hello";
  EXPECT_EQ(FmValueTag::Text, value_tag(Value::text(s)));
}

TEST(ValueMarshalTag, ErrorMapsToError) {
  EXPECT_EQ(FmValueTag::Error, value_tag(Value::error(ErrorCode::Div0)));
}

TEST(ValueMarshalTag, ArrayMapsToArray) {
  Arena arena;
  const ArrayValue* arr = MakeNumberedArray(&arena, 2, 2);
  EXPECT_EQ(FmValueTag::Array, value_tag(Value::array(arr)));
}

TEST(ValueMarshalTag, RefSlotIsUnreachableButTotalOverEnum) {
  // ValueKind::Ref has no factory, so we cannot construct a Ref-kind Value
  // directly. Re-verify the underlying numeric identity instead — this is
  // the same property the static_asserts in value_marshal.cpp lock in.
  EXPECT_EQ(static_cast<std::int32_t>(FmValueTag::Ref),
            static_cast<std::int32_t>(static_cast<std::uint8_t>(ValueKind::Ref)));
}

TEST(ValueMarshalTag, LambdaMatchesValueKindOrdinal) {
  // Lambda construction requires a parser arena + LambdaValue payload that
  // pulls in more dependencies than this leaf test should touch. Verify the
  // ordinal identity directly — bindings rely only on the static cast.
  EXPECT_EQ(static_cast<std::int32_t>(FmValueTag::Lambda),
            static_cast<std::int32_t>(static_cast<std::uint8_t>(ValueKind::Lambda)));
}

// ---------------------------------------------------------------------------
// error_to_fm_code / error_to_fm_text: spot-check 5 representative codes
// ---------------------------------------------------------------------------

TEST(ValueMarshalError, Div0RoundTrip) {
  EXPECT_EQ(7, error_to_fm_code(ErrorCode::Div0));
  EXPECT_STREQ("#DIV/0!", error_to_fm_text(ErrorCode::Div0));
}

TEST(ValueMarshalError, NaRoundTrip) {
  EXPECT_EQ(42, error_to_fm_code(ErrorCode::NA));
  EXPECT_STREQ("#N/A", error_to_fm_text(ErrorCode::NA));
}

TEST(ValueMarshalError, ValueRoundTrip) {
  EXPECT_EQ(15, error_to_fm_code(ErrorCode::Value));
  EXPECT_STREQ("#VALUE!", error_to_fm_text(ErrorCode::Value));
}

TEST(ValueMarshalError, RefRoundTrip) {
  EXPECT_EQ(23, error_to_fm_code(ErrorCode::Ref));
  EXPECT_STREQ("#REF!", error_to_fm_text(ErrorCode::Ref));
}

TEST(ValueMarshalError, NameRoundTrip) {
  EXPECT_EQ(29, error_to_fm_code(ErrorCode::Name));
  EXPECT_STREQ("#NAME?", error_to_fm_text(ErrorCode::Name));
}

TEST(ValueMarshalError, MatchesValueHelpers) {
  // The bindings should be free to migrate from the value.h symbols to
  // the c_api re-exports and back. Lock the identity here so a divergence
  // would surface as a test failure rather than a silent drift.
  for (auto ec :
       {ErrorCode::Null, ErrorCode::Div0, ErrorCode::Value, ErrorCode::Ref, ErrorCode::Name, ErrorCode::Num,
        ErrorCode::NA, ErrorCode::GettingData, ErrorCode::Spill, ErrorCode::Calc, ErrorCode::Field, ErrorCode::Blocked,
        ErrorCode::Connect, ErrorCode::External, ErrorCode::Busy, ErrorCode::Python, ErrorCode::Unknown}) {
    EXPECT_EQ(static_cast<std::int32_t>(formulon::ooxml_code(ec)), error_to_fm_code(ec));
    EXPECT_STREQ(formulon::display_name(ec), error_to_fm_text(ec));
  }
}

// ---------------------------------------------------------------------------
// inspect_array
// ---------------------------------------------------------------------------

TEST(ValueMarshalArray, NonArrayReportsEmpty) {
  ArrayShape s = inspect_array(Value::number(1.0));
  EXPECT_EQ(0u, s.rows);
  EXPECT_EQ(0u, s.cols);
  EXPECT_TRUE(s.empty);
}

TEST(ValueMarshalArray, ZeroByZeroIsEmpty) {
  Arena arena;
  ArrayValue* a = arena.create<ArrayValue>();
  a->rows = 0;
  a->cols = 0;
  a->cells = nullptr;
  ArrayShape s = inspect_array(Value::array(a));
  EXPECT_EQ(0u, s.rows);
  EXPECT_EQ(0u, s.cols);
  EXPECT_TRUE(s.empty);
}

TEST(ValueMarshalArray, OneByOneNotEmpty) {
  Arena arena;
  const ArrayValue* a = MakeNumberedArray(&arena, 1, 1);
  ArrayShape s = inspect_array(Value::array(a));
  EXPECT_EQ(1u, s.rows);
  EXPECT_EQ(1u, s.cols);
  EXPECT_FALSE(s.empty);
}

TEST(ValueMarshalArray, FiveByThreeReportsShape) {
  Arena arena;
  const ArrayValue* a = MakeNumberedArray(&arena, 5, 3);
  ArrayShape s = inspect_array(Value::array(a));
  EXPECT_EQ(5u, s.rows);
  EXPECT_EQ(3u, s.cols);
  EXPECT_FALSE(s.empty);
}

// ---------------------------------------------------------------------------
// value_as_number_or_zero
// ---------------------------------------------------------------------------

TEST(ValueMarshalScalar, NumberHappyPath) {
  EXPECT_DOUBLE_EQ(3.14, value_as_number_or_zero(Value::number(3.14)));
  EXPECT_DOUBLE_EQ(0.0, value_as_number_or_zero(Value::number(0.0)));
  EXPECT_DOUBLE_EQ(-1.5, value_as_number_or_zero(Value::number(-1.5)));
}

TEST(ValueMarshalScalar, NumberMissReturnsZero) {
  EXPECT_DOUBLE_EQ(0.0, value_as_number_or_zero(Value::blank()));
  EXPECT_DOUBLE_EQ(0.0, value_as_number_or_zero(Value::boolean(true)));
  EXPECT_DOUBLE_EQ(0.0, value_as_number_or_zero(Value::error(ErrorCode::Div0)));
  const std::string s = "1.5";
  EXPECT_DOUBLE_EQ(0.0, value_as_number_or_zero(Value::text(s)));
}

// ---------------------------------------------------------------------------
// value_as_bool_or_false
// ---------------------------------------------------------------------------

TEST(ValueMarshalScalar, BoolHappyPath) {
  EXPECT_TRUE(value_as_bool_or_false(Value::boolean(true)));
  EXPECT_FALSE(value_as_bool_or_false(Value::boolean(false)));
}

TEST(ValueMarshalScalar, BoolMissReturnsFalse) {
  EXPECT_FALSE(value_as_bool_or_false(Value::blank()));
  EXPECT_FALSE(value_as_bool_or_false(Value::number(1.0)));
  EXPECT_FALSE(value_as_bool_or_false(Value::number(0.0)));
  EXPECT_FALSE(value_as_bool_or_false(Value::error(ErrorCode::Value)));
  const std::string s = "TRUE";
  EXPECT_FALSE(value_as_bool_or_false(Value::text(s)));
}

// ---------------------------------------------------------------------------
// value_as_text_or_null
// ---------------------------------------------------------------------------

TEST(ValueMarshalScalar, TextHappyPath) {
  std::string storage;
  const std::string source = "hello world";
  const std::string* out = value_as_text_or_null(Value::text(source), storage);
  ASSERT_NE(nullptr, out);
  EXPECT_EQ(&storage, out);
  EXPECT_EQ("hello world", *out);
}

TEST(ValueMarshalScalar, TextEmptyDistinguishesFromMiss) {
  std::string storage = "stale";
  const std::string empty_source;
  const std::string* out = value_as_text_or_null(Value::text(empty_source), storage);
  ASSERT_NE(nullptr, out);
  // Empty text is "present and empty", not "absent".
  EXPECT_TRUE(out->empty());
}

TEST(ValueMarshalScalar, TextMissReturnsNull) {
  std::string storage = "untouched";
  EXPECT_EQ(nullptr, value_as_text_or_null(Value::blank(), storage));
  EXPECT_EQ(nullptr, value_as_text_or_null(Value::number(0.0), storage));
  EXPECT_EQ(nullptr, value_as_text_or_null(Value::boolean(true), storage));
  EXPECT_EQ(nullptr, value_as_text_or_null(Value::error(ErrorCode::NA), storage));
  // Storage is left untouched on a miss.
  EXPECT_EQ("untouched", storage);
}

}  // namespace
}  // namespace internal
}  // namespace c_api
}  // namespace formulon
