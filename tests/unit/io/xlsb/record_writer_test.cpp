// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Symmetry tests for the MS-XLSB record-emission primitives. Each
// case round-trips the writer output through the reader (`io/xlsb/
// record.{h,cpp}`) and checks that the decoded `(type, payload)` /
// string / RkNumber matches the original. This is the core
// correctness invariant for Bundle 4.2: any byte sequence the writer
// produces must round-trip through the reader without loss.

#include "io/xlsb/record_writer.h"

#include <cmath>
#include <cstdint>
#include <cstring>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/record.h"
#include "io/zip_reader.h"
#include "utils/error.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return ByteSpan{v.data(), v.size()};
}

// ---------------------------------------------------------------------------
// Record framing: type + size + payload round-trip.
// ---------------------------------------------------------------------------

struct FramingCase {
  std::uint16_t type;
  std::vector<std::uint8_t> payload;
  std::size_t expected_header_bytes;  // for sanity-checking minimal-length encoding
};

class RecordFramingRoundTrip : public ::testing::TestWithParam<FramingCase> {};

TEST_P(RecordFramingRoundTrip, WriterReaderSymmetry) {
  const FramingCase& tc = GetParam();
  std::vector<std::uint8_t> bytes;
  emit_record(bytes, tc.type, ByteSpan{tc.payload.data(), tc.payload.size()});

  ByteSpan cursor = SpanOf(bytes);
  auto rec = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(rec)) << rec.error().message;
  EXPECT_EQ(rec.value().type, tc.type);
  ASSERT_EQ(rec.value().payload.size, tc.payload.size());
  if (!tc.payload.empty()) {
    EXPECT_EQ(0, std::memcmp(rec.value().payload.data, tc.payload.data(), tc.payload.size()));
  }
  EXPECT_EQ(cursor.size, 0U);

  // Sanity: header length matches the minimum encoding the writer
  // should pick. payload must come right after the header.
  EXPECT_EQ(bytes.size() - tc.payload.size(), tc.expected_header_bytes);
}

INSTANTIATE_TEST_SUITE_P(XlsbRecordWriter, RecordFramingRoundTrip,
                         ::testing::Values(
                             // 1-byte type, 1-byte size, empty payload.
                             FramingCase{5U, {}, 2U},
                             // 1-byte type, 1-byte size, small payload.
                             FramingCase{1U, {0xAA, 0xBB, 0xCC, 0xDD, 0xEE}, 2U},
                             // 2-byte type (>= 0x80), 1-byte size.
                             FramingCase{130U, {}, 3U},
                             // 2-byte type (== max 14-bit), 1-byte size.
                             FramingCase{16383U, {}, 3U},
                             // 1-byte type, 2-byte size (>= 0x80).
                             FramingCase{1U, std::vector<std::uint8_t>(128, 0xCC), 3U},
                             // 1-byte type, 3-byte size (>= 0x4000).
                             FramingCase{1U, std::vector<std::uint8_t>(0x4000, 0xDE), 4U},
                             // 1-byte type, 4-byte size (>= 0x200000).
                             FramingCase{1U, std::vector<std::uint8_t>(0x200000, 0xEE), 5U},
                             // 2-byte type + 3-byte size combination.
                             FramingCase{130U, std::vector<std::uint8_t>(0x4000, 0xAB), 5U},
                             // BrtCellRk (type=2): 4-byte payload, classic small record.
                             FramingCase{2U, {0x42, 0x42, 0x42, 0x42}, 2U},
                             // BrtBundleSh (type=156): 2-byte type, real-world record.
                             FramingCase{156U, {0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00}, 3U}));

// ---------------------------------------------------------------------------
// XLWideString: ASCII, Unicode BMP, and surrogate-pair round-trip.
// ---------------------------------------------------------------------------

TEST(XlsbRecordWriter, XLWideStringEmpty) {
  std::vector<std::uint8_t> bytes;
  emit_xlwidestring(bytes, std::string_view{""});
  ByteSpan cursor = SpanOf(bytes);
  auto s = read_xlwidestring(cursor);
  ASSERT_TRUE(static_cast<bool>(s)) << s.error().message;
  EXPECT_EQ(s.value(), "");
  EXPECT_EQ(cursor.size, 0U);
  EXPECT_EQ(bytes.size(), 4U);  // length prefix only
}

TEST(XlsbRecordWriter, XLWideStringAscii) {
  std::vector<std::uint8_t> bytes;
  emit_xlwidestring(bytes, std::string_view{"hello"});
  ByteSpan cursor = SpanOf(bytes);
  auto s = read_xlwidestring(cursor);
  ASSERT_TRUE(static_cast<bool>(s));
  EXPECT_EQ(s.value(), "hello");
  EXPECT_EQ(bytes.size(), 4U + 5U * 2U);
}

TEST(XlsbRecordWriter, XLWideStringBmpJapanese) {
  // U+65E5 U+672C ("日本"): both BMP, two UTF-16 code units.
  const std::string jp = "\xE6\x97\xA5\xE6\x9C\xAC";
  std::vector<std::uint8_t> bytes;
  emit_xlwidestring(bytes, jp);
  ByteSpan cursor = SpanOf(bytes);
  auto s = read_xlwidestring(cursor);
  ASSERT_TRUE(static_cast<bool>(s));
  EXPECT_EQ(s.value(), jp);
  // 4-byte length + 2 code units * 2 bytes.
  EXPECT_EQ(bytes.size(), 4U + 4U);
}

TEST(XlsbRecordWriter, XLWideStringSurrogatePair) {
  // U+1F600 GRINNING FACE: outside BMP, encoded as surrogate pair.
  const std::string emoji = "\xF0\x9F\x98\x80";
  std::vector<std::uint8_t> bytes;
  emit_xlwidestring(bytes, emoji);
  ByteSpan cursor = SpanOf(bytes);
  auto s = read_xlwidestring(cursor);
  ASSERT_TRUE(static_cast<bool>(s));
  EXPECT_EQ(s.value(), emoji);
  // 4-byte length + 1 codepoint expanded to 2 surrogates * 2 bytes = 4.
  EXPECT_EQ(bytes.size(), 4U + 4U);
}

// ---------------------------------------------------------------------------
// XLNullableWideString: sentinel + present forms.
// ---------------------------------------------------------------------------

TEST(XlsbRecordWriter, XLNullableWideStringSentinel) {
  std::vector<std::uint8_t> bytes;
  emit_xlnullablewidestring(bytes, std::nullopt);
  ASSERT_EQ(bytes.size(), 4U);
  EXPECT_EQ(bytes[0], 0xFF);
  EXPECT_EQ(bytes[1], 0xFF);
  EXPECT_EQ(bytes[2], 0xFF);
  EXPECT_EQ(bytes[3], 0xFF);

  ByteSpan cursor = SpanOf(bytes);
  auto s = read_xlnullablewidestring(cursor);
  ASSERT_TRUE(static_cast<bool>(s));
  EXPECT_EQ(s.value(), "");
  EXPECT_EQ(cursor.size, 0U);
}

TEST(XlsbRecordWriter, XLNullableWideStringPresentEmpty) {
  std::vector<std::uint8_t> bytes;
  emit_xlnullablewidestring(bytes, std::optional<std::string_view>{""});
  // Present empty: 4-byte length zero, no body.
  ASSERT_EQ(bytes.size(), 4U);
  EXPECT_EQ(bytes[0], 0x00);
  EXPECT_EQ(bytes[1], 0x00);
  EXPECT_EQ(bytes[2], 0x00);
  EXPECT_EQ(bytes[3], 0x00);

  ByteSpan cursor = SpanOf(bytes);
  auto s = read_xlnullablewidestring(cursor);
  ASSERT_TRUE(static_cast<bool>(s));
  EXPECT_EQ(s.value(), "");
}

TEST(XlsbRecordWriter, XLNullableWideStringPresentNonEmpty) {
  std::vector<std::uint8_t> bytes;
  emit_xlnullablewidestring(bytes, std::optional<std::string_view>{"rId1"});
  ByteSpan cursor = SpanOf(bytes);
  auto s = read_xlnullablewidestring(cursor);
  ASSERT_TRUE(static_cast<bool>(s));
  EXPECT_EQ(s.value(), "rId1");
}

// ---------------------------------------------------------------------------
// RkNumber round-trip cases.
// ---------------------------------------------------------------------------

struct RkCase {
  double value;
  bool round_trips;
};

class RkRoundTrip : public ::testing::TestWithParam<RkCase> {};

TEST_P(RkRoundTrip, ValueRoundTripsExactlyWhenPredicateSaysSo) {
  const RkCase& tc = GetParam();
  EXPECT_EQ(rk_round_trips_value(tc.value), tc.round_trips);

  std::vector<std::uint8_t> bytes;
  emit_rk_number(bytes, tc.value);
  ASSERT_EQ(bytes.size(), 4U);

  ByteSpan cursor = SpanOf(bytes);
  auto rk = read_u32(cursor);
  ASSERT_TRUE(static_cast<bool>(rk));
  EXPECT_EQ(cursor.size, 0U);
  const double decoded = decode_rk_number(rk.value());

  if (tc.round_trips) {
    // For NaN we cannot use ==; here we only flag finite cases as
    // round-tripping, so plain equality is sufficient.
    EXPECT_EQ(decoded, tc.value) << "value=" << tc.value << " rk=0x" << std::hex << rk.value();
  } else {
    // The predicate said this value would lose precision; just sanity-
    // check that the encoding succeeded and produced *some* double.
    (void)decoded;
  }
}

INSTANTIATE_TEST_SUITE_P(XlsbRecordWriter, RkRoundTrip,
                         ::testing::Values(RkCase{0.0, true},                                 // exact integer zero
                                           RkCase{1.0, true},                                 // small positive integer
                                           RkCase{-1.0, true},                                // small negative integer
                                           RkCase{42.0, true},                                // mid-range integer
                                           RkCase{static_cast<double>((1 << 29) - 1), true},  // upper bound
                                           RkCase{-static_cast<double>(1 << 29), true},       // lower bound
                                           RkCase{0.5, true},      // IEEE-form: low 34 bits zero
                                           RkCase{123.45, true},   // x100 form
                                           RkCase{-123.45, true},  // negative x100 form
                                           // Slightly outside x100 representable range: pick a
                                           // double whose lower 34 bits are non-zero so neither
                                           // integer nor IEEE-exact form applies.
                                           RkCase{1.0 / 3.0, false}));

TEST(XlsbRecordWriter, RkNumberNegativeZeroPicksIeeeForm) {
  // -0.0 must NOT take the integer encoding (which loses the sign).
  // The predicate already rejects it; here we additionally confirm
  // the encoded form preserves the sign of zero through the IEEE
  // path even though the predicate said the value won't round-trip
  // (the upper 32 bits of -0.0 are 0x80000000, lower 32 bits are 0,
  // so the IEEE form is exact for -0.0 specifically).
  std::vector<std::uint8_t> bytes;
  emit_rk_number(bytes, -0.0);
  ASSERT_EQ(bytes.size(), 4U);
  ByteSpan cursor = SpanOf(bytes);
  auto rk = read_u32(cursor);
  ASSERT_TRUE(static_cast<bool>(rk));
  // Low two bits clear means IEEE form was selected (fInt=0, fX100=0).
  EXPECT_EQ(rk.value() & 0x3U, 0U);
  const double decoded = decode_rk_number(rk.value());
  EXPECT_TRUE(std::signbit(decoded));
  EXPECT_EQ(decoded, 0.0);
}

TEST(XlsbRecordWriter, RkNumberLargeIntegerFallsBackToIeee) {
  // Beyond the 30-bit signed range. (1 << 30) is exactly representable
  // as IEEE 754 double with low 34 bits zero — confirm the encoding
  // still round-trips even though the integer form doesn't apply.
  const double v = static_cast<double>(1LL << 30);
  EXPECT_TRUE(rk_round_trips_value(v));
  std::vector<std::uint8_t> bytes;
  emit_rk_number(bytes, v);
  ByteSpan cursor = SpanOf(bytes);
  auto rk = read_u32(cursor);
  ASSERT_TRUE(static_cast<bool>(rk));
  // fInt=0 because we fell back to IEEE form.
  EXPECT_EQ(rk.value() & 0x2U, 0U);
  EXPECT_EQ(decode_rk_number(rk.value()), v);
}

// ---------------------------------------------------------------------------
// Smoke: integer scalar emitters compose into record bodies cleanly.
// ---------------------------------------------------------------------------

TEST(XlsbRecordWriter, U8U16U32EmitInLittleEndianOrder) {
  std::vector<std::uint8_t> bytes;
  emit_u8(bytes, 0xAB);
  emit_u16(bytes, 0x1234);
  emit_u32(bytes, 0xDEADBEEFU);

  ByteSpan cursor = SpanOf(bytes);
  auto u8 = read_u8(cursor);
  auto u16 = read_u16(cursor);
  auto u32 = read_u32(cursor);
  ASSERT_TRUE(static_cast<bool>(u8));
  ASSERT_TRUE(static_cast<bool>(u16));
  ASSERT_TRUE(static_cast<bool>(u32));
  EXPECT_EQ(u8.value(), 0xABU);
  EXPECT_EQ(u16.value(), 0x1234U);
  EXPECT_EQ(u32.value(), 0xDEADBEEFU);
  EXPECT_EQ(cursor.size, 0U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
