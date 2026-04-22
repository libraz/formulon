// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the M2.1 scalar `Value`. The non-scalar variants
// (Text/Array/Ref/Lambda) are intentionally out of scope here.

#include "value.h"

#include <cmath>
#include <cstdint>
#include <limits>
#include <string>

#include "gtest/gtest.h"

namespace formulon {
namespace {

constexpr ErrorCode kAllErrorCodes[] = {
    ErrorCode::Null,   ErrorCode::Div0,    ErrorCode::Value,       ErrorCode::Ref,      ErrorCode::Name,
    ErrorCode::Num,    ErrorCode::NA,      ErrorCode::GettingData, ErrorCode::Spill,    ErrorCode::Calc,
    ErrorCode::Field,  ErrorCode::Blocked, ErrorCode::Connect,     ErrorCode::External, ErrorCode::Busy,
    ErrorCode::Python, ErrorCode::Unknown,
};

TEST(ValueTest, DefaultIsBlank) {
  Value v = Value::blank();
  EXPECT_EQ(ValueKind::Blank, v.kind());
  EXPECT_TRUE(v.is_blank());
  EXPECT_FALSE(v.is_number());
  EXPECT_FALSE(v.is_boolean());
  EXPECT_FALSE(v.is_error());
  EXPECT_FALSE(v.is_text());
  EXPECT_FALSE(v.is_array());
  EXPECT_FALSE(v.is_ref());
  EXPECT_FALSE(v.is_lambda());
}

TEST(ValueTest, NumberFactoryRoundTrip) {
  const double samples[] = {
      0.0,
      -0.0,
      1.0,
      -1.0,
      42.0,
      1e300,
      -1e300,
      std::numeric_limits<double>::min(),
      std::numeric_limits<double>::max(),
      std::numeric_limits<double>::infinity(),
      -std::numeric_limits<double>::infinity(),
  };
  for (double d : samples) {
    Value v = Value::number(d);
    EXPECT_EQ(ValueKind::Number, v.kind());
    EXPECT_TRUE(v.is_number());
    EXPECT_EQ(d, v.as_number()) << "d=" << d;
  }

  // NaN round-trips but is not equal to itself; test separately below.
  Value nan_v = Value::number(std::numeric_limits<double>::quiet_NaN());
  EXPECT_TRUE(nan_v.is_number());
  EXPECT_TRUE(std::isnan(nan_v.as_number()));
}

TEST(ValueTest, NumberPreservesNegativeZero) {
  Value neg = Value::number(-0.0);
  Value pos = Value::number(0.0);
  EXPECT_TRUE(std::signbit(neg.as_number()));
  EXPECT_FALSE(std::signbit(pos.as_number()));
  // IEEE-754: -0.0 == 0.0.
  EXPECT_EQ(neg, pos);
}

TEST(ValueTest, NaNIsNotEqualToItself) {
  Value a = Value::number(std::numeric_limits<double>::quiet_NaN());
  Value b = Value::number(std::numeric_limits<double>::quiet_NaN());
  EXPECT_NE(a, b) << "NaN must not compare equal to itself";
  EXPECT_FALSE(a == b);
}

TEST(ValueTest, BooleanFactoryRoundTrip) {
  Value t = Value::boolean(true);
  EXPECT_EQ(ValueKind::Bool, t.kind());
  EXPECT_TRUE(t.is_boolean());
  EXPECT_TRUE(t.as_boolean());

  Value f = Value::boolean(false);
  EXPECT_EQ(ValueKind::Bool, f.kind());
  EXPECT_TRUE(f.is_boolean());
  EXPECT_FALSE(f.as_boolean());
}

TEST(ValueTest, ErrorFactoryRoundTrip) {
  for (ErrorCode c : kAllErrorCodes) {
    Value v = Value::error(c);
    EXPECT_EQ(ValueKind::Error, v.kind()) << "code=" << static_cast<int>(c);
    EXPECT_TRUE(v.is_error());
    EXPECT_EQ(c, v.as_error());
  }
}

TEST(ValueTest, ValueEqualityWithinSameKind) {
  EXPECT_EQ(Value::blank(), Value::blank());
  EXPECT_EQ(Value::number(3.14), Value::number(3.14));
  EXPECT_NE(Value::number(3.14), Value::number(2.71));
  EXPECT_EQ(Value::boolean(true), Value::boolean(true));
  EXPECT_NE(Value::boolean(true), Value::boolean(false));
  EXPECT_EQ(Value::error(ErrorCode::Div0), Value::error(ErrorCode::Div0));
  EXPECT_NE(Value::error(ErrorCode::Div0), Value::error(ErrorCode::Value));
}

TEST(ValueTest, ValueInequalityAcrossKinds) {
  EXPECT_NE(Value::blank(), Value::number(0.0));
  EXPECT_NE(Value::number(0.0), Value::boolean(false));
  EXPECT_NE(Value::number(1.0), Value::boolean(true));
  EXPECT_NE(Value::blank(), Value::boolean(false));
  EXPECT_NE(Value::number(0.0), Value::error(ErrorCode::Null));
  EXPECT_NE(Value::boolean(true), Value::error(ErrorCode::Value));
  EXPECT_NE(Value::blank(), Value::error(ErrorCode::NA));
}

TEST(ValueTest, DebugToStringFormats) {
  EXPECT_EQ("Blank", Value::blank().debug_to_string());

  // Exact number formatting is not Excel-display-exact at M2.1; we only
  // check the wrapper shape and that the payload appears.
  const std::string num = Value::number(42.0).debug_to_string();
  EXPECT_EQ(0u, num.find("Number("));
  EXPECT_EQ(num.size() - 1, num.rfind(")"));
  EXPECT_NE(std::string::npos, num.find("42"));

  EXPECT_EQ("Bool(true)", Value::boolean(true).debug_to_string());
  EXPECT_EQ("Bool(false)", Value::boolean(false).debug_to_string());
  EXPECT_EQ("Error(#DIV/0!)", Value::error(ErrorCode::Div0).debug_to_string());
  EXPECT_EQ("Error(#N/A)", Value::error(ErrorCode::NA).debug_to_string());
  EXPECT_EQ("Error(#NAME?)", Value::error(ErrorCode::Name).debug_to_string());
}

TEST(ValueTest, OoxmlCodeMatchesSpec) {
  // Authoritative table from backup/plans/02-calc-engine.md §2.1.
  struct Expected {
    ErrorCode code;
    std::int32_t wire;
  };
  const Expected table[] = {
      {ErrorCode::Null, 0},     {ErrorCode::Div0, 7},     {ErrorCode::Value, 15},    {ErrorCode::Ref, 23},
      {ErrorCode::Name, 29},    {ErrorCode::Num, 36},     {ErrorCode::NA, 42},       {ErrorCode::GettingData, 43},
      {ErrorCode::Spill, 14},   {ErrorCode::Calc, 13},    {ErrorCode::Unknown, 9},   {ErrorCode::Field, 10},
      {ErrorCode::Connect, 11}, {ErrorCode::Blocked, 12}, {ErrorCode::External, 19}, {ErrorCode::Busy, 16},
      {ErrorCode::Python, 17},
  };
  for (const auto& row : table) {
    EXPECT_EQ(row.wire, ooxml_code(row.code)) << "code=" << static_cast<int>(row.code);
  }
}

TEST(ValueTest, OoxmlCodeIsConstexpr) {
  static_assert(ooxml_code(ErrorCode::Null) == 0, "Null wire code");
  static_assert(ooxml_code(ErrorCode::Div0) == 7, "Div0 wire code");
  static_assert(ooxml_code(ErrorCode::Value) == 15, "Value wire code");
  static_assert(ooxml_code(ErrorCode::NA) == 42, "NA wire code");
  static_assert(ooxml_code(ErrorCode::GettingData) == 43, "GettingData wire code");
  static_assert(ooxml_code(ErrorCode::Python) == 17, "Python wire code");
  SUCCEED();
}

TEST(ValueTest, DisplayNameMatchesSpec) {
  struct Expected {
    ErrorCode code;
    const char* text;
  };
  const Expected table[] = {
      {ErrorCode::Null, "#NULL!"},       {ErrorCode::Div0, "#DIV/0!"},
      {ErrorCode::Value, "#VALUE!"},     {ErrorCode::Ref, "#REF!"},
      {ErrorCode::Name, "#NAME?"},       {ErrorCode::Num, "#NUM!"},
      {ErrorCode::NA, "#N/A"},           {ErrorCode::GettingData, "#GETTING_DATA"},
      {ErrorCode::Spill, "#SPILL!"},     {ErrorCode::Calc, "#CALC!"},
      {ErrorCode::Field, "#FIELD!"},     {ErrorCode::Blocked, "#BLOCKED!"},
      {ErrorCode::Connect, "#CONNECT!"}, {ErrorCode::External, "#EXTERNAL!"},
      {ErrorCode::Busy, "#BUSY!"},       {ErrorCode::Python, "#PYTHON!"},
      {ErrorCode::Unknown, "#UNKNOWN!"},
  };
  for (const auto& row : table) {
    EXPECT_STREQ(row.text, display_name(row.code)) << "code=" << static_cast<int>(row.code);
  }
}

TEST(ValueTest, SizeofValueIsReasonable) {
  // The tag (1 byte) plus an 8-byte double payload should pack into 16 bytes
  // on every platform Formulon targets. Enforcing this in a test (in
  // addition to the static_assert in value.h) gives a louder failure.
  EXPECT_LE(sizeof(Value), static_cast<std::size_t>(16));
  EXPECT_TRUE(std::is_trivially_copyable<Value>::value);
}

TEST(ValueTest, AllErrorCodesHaveWireCodeAndDisplayName) {
  // Sanity: every `ErrorCode` listed in kAllErrorCodes has both a
  // non-negative wire code and a non-empty display string.
  for (ErrorCode c : kAllErrorCodes) {
    EXPECT_GE(ooxml_code(c), 0) << "code=" << static_cast<int>(c);
    const char* name = display_name(c);
    ASSERT_NE(nullptr, name);
    EXPECT_GT(std::string(name).size(), 0u) << "code=" << static_cast<int>(c);
    EXPECT_EQ('#', name[0]) << "code=" << static_cast<int>(c);
  }
}

#if GTEST_HAS_DEATH_TEST
TEST(ValueDeathTest, AsNumberOnBlankAborts) {
  Value v = Value::blank();
  EXPECT_DEATH(v.as_number(), ".*");
}

TEST(ValueDeathTest, AsBooleanOnNumberAborts) {
  Value v = Value::number(1.0);
  EXPECT_DEATH(v.as_boolean(), ".*");
}

TEST(ValueDeathTest, AsErrorOnBooleanAborts) {
  Value v = Value::boolean(true);
  EXPECT_DEATH(v.as_error(), ".*");
}

TEST(ValueDeathTest, AsNumberOnErrorAborts) {
  Value v = Value::error(ErrorCode::Div0);
  EXPECT_DEATH(v.as_number(), ".*");
}
#endif  // GTEST_HAS_DEATH_TEST

}  // namespace
}  // namespace formulon
