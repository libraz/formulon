// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the MS-XLSB record framing primitives.

#include "io/xlsb/record.h"

#include <cmath>
#include <cstdint>
#include <cstring>
#include <vector>

#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "utils/error.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return ByteSpan{v.data(), v.size()};
}

TEST(XlsbRecord, ReadsOneByteRecordType) {
  // Record-type = 0x05 (BrtCellReal), payload-size = 0x00, no payload.
  const std::vector<std::uint8_t> bytes = {0x05, 0x00};
  ByteSpan cursor = SpanOf(bytes);
  auto rec = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(rec)) << rec.error().message;
  EXPECT_EQ(rec.value().type, 5U);
  EXPECT_EQ(rec.value().payload.size, 0U);
  // Cursor consumed both bytes.
  EXPECT_EQ(cursor.size, 0U);
}

TEST(XlsbRecord, ReadsTwoByteRecordType) {
  // Record-type encoded as two bytes: 0x82, 0x01 -> (0x02 | (0x01 << 7))
  // = 0x82 = 130 (BrtEndSheet). The first byte's MSB signals the
  // continuation; only 7 bits per byte contribute.
  const std::vector<std::uint8_t> bytes = {0x82, 0x01, 0x00};
  ByteSpan cursor = SpanOf(bytes);
  auto rec = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(rec)) << rec.error().message;
  EXPECT_EQ(rec.value().type, 130U);
  EXPECT_EQ(rec.value().payload.size, 0U);
  EXPECT_EQ(cursor.size, 0U);
}

TEST(XlsbRecord, ReadsOneByteSize) {
  // Type 0x01, size 0x05, then 5 bytes of payload.
  const std::vector<std::uint8_t> bytes = {0x01, 0x05, 0xAA, 0xBB, 0xCC, 0xDD, 0xEE};
  ByteSpan cursor = SpanOf(bytes);
  auto rec = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(rec)) << rec.error().message;
  EXPECT_EQ(rec.value().type, 1U);
  ASSERT_EQ(rec.value().payload.size, 5U);
  EXPECT_EQ(rec.value().payload.data[0], 0xAA);
  EXPECT_EQ(rec.value().payload.data[4], 0xEE);
  EXPECT_EQ(cursor.size, 0U);
}

TEST(XlsbRecord, ReadsTwoByteSize) {
  // Size = 0x80 (128) encoded as 0x80, 0x01 -> (0x00 | (0x01 << 7)).
  std::vector<std::uint8_t> bytes;
  bytes.push_back(0x01);  // type
  bytes.push_back(0x80);  // size byte 1 (continuation set)
  bytes.push_back(0x01);  // size byte 2
  bytes.resize(bytes.size() + 128, 0xCC);
  ByteSpan cursor = SpanOf(bytes);
  auto rec = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(rec)) << rec.error().message;
  EXPECT_EQ(rec.value().payload.size, 128U);
  EXPECT_EQ(cursor.size, 0U);
}

TEST(XlsbRecord, ReadsThreeByteSize) {
  // Size = 0x4000 (16384) encoded as 0x80, 0x80, 0x01.
  std::vector<std::uint8_t> bytes;
  bytes.push_back(0x01);
  bytes.push_back(0x80);
  bytes.push_back(0x80);
  bytes.push_back(0x01);
  bytes.resize(bytes.size() + 16384, 0xDE);
  ByteSpan cursor = SpanOf(bytes);
  auto rec = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(rec)) << rec.error().message;
  EXPECT_EQ(rec.value().payload.size, 16384U);
  EXPECT_EQ(cursor.size, 0U);
}

TEST(XlsbRecord, ReadsFourByteSize) {
  // Size = 0x200000 (2097152) encoded as 0x80, 0x80, 0x80, 0x01.
  std::vector<std::uint8_t> bytes;
  bytes.push_back(0x01);
  bytes.push_back(0x80);
  bytes.push_back(0x80);
  bytes.push_back(0x80);
  bytes.push_back(0x01);
  bytes.resize(bytes.size() + 0x200000, 0xEE);
  ByteSpan cursor = SpanOf(bytes);
  auto rec = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(rec)) << rec.error().message;
  EXPECT_EQ(rec.value().payload.size, 0x200000U);
}

TEST(XlsbRecord, TruncatedHeaderReturnsKIoXlsbRecordTruncated) {
  // Type byte present, no size byte.
  const std::vector<std::uint8_t> bytes = {0x01};
  ByteSpan cursor = SpanOf(bytes);
  auto rec = read_record(cursor);
  ASSERT_FALSE(static_cast<bool>(rec));
  EXPECT_EQ(rec.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}

TEST(XlsbRecord, TruncatedPayloadReturnsKIoXlsbRecordTruncated) {
  // Type 0x01, size 0x10 (16), but only 4 bytes of payload follow.
  const std::vector<std::uint8_t> bytes = {0x01, 0x10, 0xAA, 0xBB, 0xCC, 0xDD};
  ByteSpan cursor = SpanOf(bytes);
  auto rec = read_record(cursor);
  ASSERT_FALSE(static_cast<bool>(rec));
  EXPECT_EQ(rec.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}

TEST(XlsbRecord, ReadU8U16U32AdvanceCursor) {
  const std::vector<std::uint8_t> bytes = {0x12, 0x34, 0x12, 0x78, 0x56, 0x34, 0x12};
  ByteSpan cursor = SpanOf(bytes);

  auto u8 = read_u8(cursor);
  ASSERT_TRUE(static_cast<bool>(u8));
  EXPECT_EQ(u8.value(), 0x12U);

  auto u16 = read_u16(cursor);
  ASSERT_TRUE(static_cast<bool>(u16));
  EXPECT_EQ(u16.value(), 0x1234U);

  auto u32 = read_u32(cursor);
  ASSERT_TRUE(static_cast<bool>(u32));
  EXPECT_EQ(u32.value(), 0x12345678U);

  EXPECT_EQ(cursor.size, 0U);
}

TEST(XlsbRecord, ReadU8OnEmptyCursorReturnsTruncated) {
  ByteSpan cursor{};
  auto u8 = read_u8(cursor);
  ASSERT_FALSE(static_cast<bool>(u8));
  EXPECT_EQ(u8.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}

TEST(XlsbRecord, RkNumberDecodesIntegerForm) {
  // fInt=1, fX100=0, payload=42 (signed). Encoded value:
  //   payload << 2 | 0x02 (fInt) = (42 << 2) | 2 = 0xAA.
  const std::uint32_t rk = (42U << 2) | 0x02U;
  const double v = decode_rk_number(rk);
  EXPECT_EQ(v, 42.0);
}

TEST(XlsbRecord, RkNumberDecodesIntegerNegative) {
  // fInt=1, payload=-1 (signed) which sign-extends to a 30-bit ones
  // pattern.
  const std::int32_t signed_neg = -7;
  const std::uint32_t rk = (static_cast<std::uint32_t>(signed_neg) << 2) | 0x02U;
  const double v = decode_rk_number(rk);
  EXPECT_EQ(v, -7.0);
}

TEST(XlsbRecord, RkNumberDecodesIntegerWithX100) {
  // fInt=1, fX100=1, payload=12345. Result: 123.45.
  const std::uint32_t rk = (12345U << 2) | 0x03U;
  const double v = decode_rk_number(rk);
  EXPECT_NEAR(v, 123.45, 1e-9);
}

TEST(XlsbRecord, RkNumberDecodesScaledDoubleForm) {
  // For fInt=0, the upper 30 bits of the RK value are taken to be the
  // upper 30 bits of an IEEE 754 double (low 34 bits zero). 0.5 has
  // bit pattern 0x3FE0000000000000; the upper 32 bits are 0x3FE00000,
  // so the RK should be that with the low 2 bits cleared.
  double half = 0.5;
  std::uint64_t bits = 0;
  std::memcpy(&bits, &half, sizeof(half));
  // Verify the low 34 bits are zero so the round-trip is exact.
  EXPECT_EQ(bits & 0x3FFFFFFFFULL, 0ULL);
  const std::uint32_t rk = static_cast<std::uint32_t>(bits >> 32) & 0xFFFFFFFCU;
  const double v = decode_rk_number(rk);
  EXPECT_EQ(v, 0.5);
}

TEST(XlsbRecord, XLNullableWideStringHandlesSentinel) {
  // The sentinel 0xFFFFFFFF means "null" -> empty string.
  std::vector<std::uint8_t> bytes = {0xFF, 0xFF, 0xFF, 0xFF};
  ByteSpan cursor = SpanOf(bytes);
  auto s = read_xlnullablewidestring(cursor);
  ASSERT_TRUE(static_cast<bool>(s));
  EXPECT_TRUE(s.value().empty());
  EXPECT_EQ(cursor.size, 0U);
}

TEST(XlsbRecord, XLWideStringDecodesAsciiPayload) {
  // 5-character ASCII string "hello": cch=5, then 5 LE u16 code units.
  std::vector<std::uint8_t> bytes = {
      0x05, 0x00, 0x00, 0x00,  // cch
      'h',  0x00, 'e',  0x00, 'l', 0x00, 'l', 0x00, 'o', 0x00,
  };
  ByteSpan cursor = SpanOf(bytes);
  auto s = read_xlwidestring(cursor);
  ASSERT_TRUE(static_cast<bool>(s)) << s.error().message;
  EXPECT_EQ(s.value(), "hello");
  EXPECT_EQ(cursor.size, 0U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
