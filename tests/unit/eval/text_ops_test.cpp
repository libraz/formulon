// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the pure UTF-8 / UTF-16 helpers in `text_ops.{h,cpp}`.
// These cover the two converters (`utf16_to_byte_offset`, `utf16_substring`)
// and the ASCII case-folding helpers in isolation, without going through
// the parser or the function registry.

#include "eval/text_ops.h"

#include <cstdint>
#include <string>
#include <string_view>

#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace {

// "あいう": three BMP codepoints, each 3 UTF-8 bytes / 1 UTF-16 unit.
constexpr const char kAiu[] = "\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86";
// "🎉": one supplementary codepoint, 4 UTF-8 bytes / 2 UTF-16 units.
constexpr const char kEmojiPopper[] = "\xF0\x9F\x8E\x89";
// "🎊": one supplementary codepoint, 4 UTF-8 bytes / 2 UTF-16 units.
constexpr const char kEmojiConfetti[] = "\xF0\x9F\x8E\x8A";

TEST(TextOpsUtf16ToByteOffset, EmptyStringReturnsZero) {
  EXPECT_EQ(utf16_to_byte_offset(std::string_view{}, 0u), 0u);
  EXPECT_EQ(utf16_to_byte_offset(std::string_view{}, 5u), 0u);
}

TEST(TextOpsUtf16ToByteOffset, AsciiOffsets) {
  const std::string_view s = "hello";
  EXPECT_EQ(utf16_to_byte_offset(s, 0u), 0u);
  EXPECT_EQ(utf16_to_byte_offset(s, 1u), 1u);
  EXPECT_EQ(utf16_to_byte_offset(s, 5u), 5u);
  // Beyond end clamps to text size.
  EXPECT_EQ(utf16_to_byte_offset(s, 99u), 5u);
}

TEST(TextOpsUtf16ToByteOffset, BmpOffsets) {
  // Each codepoint is 3 bytes / 1 unit.
  EXPECT_EQ(utf16_to_byte_offset(kAiu, 0u), 0u);
  EXPECT_EQ(utf16_to_byte_offset(kAiu, 1u), 3u);
  EXPECT_EQ(utf16_to_byte_offset(kAiu, 2u), 6u);
  EXPECT_EQ(utf16_to_byte_offset(kAiu, 3u), 9u);
}

TEST(TextOpsUtf16ToByteOffset, SupplementaryOffsetsRoundUpAtMidpoint) {
  // Single emoji = 4 bytes / 2 units.
  EXPECT_EQ(utf16_to_byte_offset(kEmojiPopper, 0u), 0u);
  // Asking for unit-1 splits the surrogate pair: rounding-up returns the
  // byte position past the entire codepoint.
  EXPECT_EQ(utf16_to_byte_offset(kEmojiPopper, 1u), 4u);
  EXPECT_EQ(utf16_to_byte_offset(kEmojiPopper, 2u), 4u);
}

TEST(TextOpsUtf16Substring, EmptyText) {
  EXPECT_EQ(utf16_substring(std::string_view{}, 0u, 5u), "");
}

TEST(TextOpsUtf16Substring, AsciiBasic) {
  EXPECT_EQ(utf16_substring("hello", 0u, 3u), "hel");
  EXPECT_EQ(utf16_substring("hello", 2u, 2u), "ll");
  EXPECT_EQ(utf16_substring("hello", 4u, 10u), "o");
}

TEST(TextOpsUtf16Substring, ZeroLength) {
  EXPECT_EQ(utf16_substring("hello", 1u, 0u), "");
}

TEST(TextOpsUtf16Substring, StartBeyondEndReturnsEmpty) {
  EXPECT_EQ(utf16_substring("hello", 99u, 3u), "");
}

TEST(TextOpsUtf16Substring, BmpSlicePreservesBytes) {
  // Take the middle codepoint of "あいう" -> "い" (3 bytes).
  EXPECT_EQ(utf16_substring(kAiu, 1u, 1u), "\xE3\x81\x84");
}

TEST(TextOpsUtf16Substring, SupplementaryWholeCodepoint) {
  // Asking for 2 units of a single emoji yields the whole codepoint.
  EXPECT_EQ(utf16_substring(kEmojiPopper, 0u, 2u), kEmojiPopper);
}

TEST(TextOpsUtf16Substring, SupplementaryMidpointRoundsUp) {
  // Two-emoji string. Asking for length=1 starting at unit 0 splits the
  // first emoji's surrogate pair. Rounding-up captures the entire first
  // emoji (4 bytes).
  std::string two_emojis(kEmojiPopper);
  two_emojis += kEmojiConfetti;
  EXPECT_EQ(utf16_substring(two_emojis, 0u, 1u), kEmojiPopper);
}

TEST(TextOpsCaseFold, AsciiLower) {
  EXPECT_EQ(to_lower_ascii("Hello, WORLD!"), "hello, world!");
  EXPECT_EQ(to_lower_ascii(""), "");
  EXPECT_EQ(to_lower_ascii("abc"), "abc");
}

TEST(TextOpsCaseFold, AsciiUpper) {
  EXPECT_EQ(to_upper_ascii("Hello, world!"), "HELLO, WORLD!");
  EXPECT_EQ(to_upper_ascii(""), "");
  EXPECT_EQ(to_upper_ascii("XYZ"), "XYZ");
}

TEST(TextOpsCaseFold, NonAsciiBytesUnchanged) {
  // ASCII case folding leaves multi-byte UTF-8 sequences alone. "café":
  // 'c','a','f' fold to 'C','A','F' but the 'é' (0xC3 0xA9) is preserved
  // verbatim.
  const std::string folded = to_upper_ascii("caf\xC3\xA9");
  EXPECT_EQ(folded, "CAF\xC3\xA9");
  // And the lowercase round-trip on the same input only flips ASCII.
  const std::string lowered = to_lower_ascii("CAF\xC3\xA9");
  EXPECT_EQ(lowered, "caf\xC3\xA9");
}

TEST(TextOpsCaseFold, BoundaryBytesAroundAlpha) {
  // Bytes adjacent to 'A' (0x40 '@'), 'Z' (0x5B '['), 'a' (0x60 '`'),
  // 'z' (0x7B '{') must NOT be affected.
  const std::string s = "@A[Z`a{z";
  EXPECT_EQ(to_upper_ascii(s), "@A[Z`A{Z");
  EXPECT_EQ(to_lower_ascii(s), "@a[z`a{z");
}

// ---------------------------------------------------------------------------
// encode_utf8_codepoint
// ---------------------------------------------------------------------------

TEST(TextOpsEncodeUtf8, AsciiOneByte) {
  EXPECT_EQ(encode_utf8_codepoint(0x41u), "A");
  EXPECT_EQ(encode_utf8_codepoint(0x00u), std::string("\x00", 1));
  EXPECT_EQ(encode_utf8_codepoint(0x7Fu), "\x7F");
}

TEST(TextOpsEncodeUtf8, TwoByteRange) {
  // U+00A9 = (c) -> 0xC2 0xA9.
  EXPECT_EQ(encode_utf8_codepoint(0x00A9u), "\xC2\xA9");
  // U+07FF -> 0xDF 0xBF (largest 2-byte).
  EXPECT_EQ(encode_utf8_codepoint(0x07FFu), "\xDF\xBF");
}

TEST(TextOpsEncodeUtf8, ThreeByteRange) {
  // U+3042 "あ" -> 0xE3 0x81 0x82.
  EXPECT_EQ(encode_utf8_codepoint(0x3042u), "\xE3\x81\x82");
  // U+FFFD replacement char -> 0xEF 0xBF 0xBD.
  EXPECT_EQ(encode_utf8_codepoint(0xFFFDu), "\xEF\xBF\xBD");
}

TEST(TextOpsEncodeUtf8, FourByteSupplementary) {
  // U+1F600 "😀" -> 0xF0 0x9F 0x98 0x80.
  EXPECT_EQ(encode_utf8_codepoint(0x1F600u), "\xF0\x9F\x98\x80");
  // U+10FFFF (max valid codepoint) -> 0xF4 0x8F 0xBF 0xBF.
  EXPECT_EQ(encode_utf8_codepoint(0x10FFFFu), "\xF4\x8F\xBF\xBF");
}

TEST(TextOpsEncodeUtf8, SurrogateAndOutOfRangeReturnEmpty) {
  EXPECT_EQ(encode_utf8_codepoint(0xD800u), "");
  EXPECT_EQ(encode_utf8_codepoint(0xDFFFu), "");
  EXPECT_EQ(encode_utf8_codepoint(0x110000u), "");
  EXPECT_EQ(encode_utf8_codepoint(0xFFFFFFFFu), "");
}

// ---------------------------------------------------------------------------
// decode_first_utf8_codepoint
// ---------------------------------------------------------------------------

TEST(TextOpsDecodeUtf8, EmptyIsInvalid) {
  const auto r = decode_first_utf8_codepoint(std::string_view{});
  EXPECT_FALSE(r.valid);
  EXPECT_EQ(r.codepoint, 0u);
  EXPECT_EQ(r.byte_len, 0u);
}

TEST(TextOpsDecodeUtf8, AsciiOneByte) {
  const auto r = decode_first_utf8_codepoint("Abc");
  EXPECT_TRUE(r.valid);
  EXPECT_EQ(r.codepoint, 0x41u);
  EXPECT_EQ(r.byte_len, 1u);
}

TEST(TextOpsDecodeUtf8, TwoByteCopyright) {
  const auto r = decode_first_utf8_codepoint("\xC2\xA9");
  EXPECT_TRUE(r.valid);
  EXPECT_EQ(r.codepoint, 0x00A9u);
  EXPECT_EQ(r.byte_len, 2u);
}

TEST(TextOpsDecodeUtf8, ThreeByteHiragana) {
  const auto r = decode_first_utf8_codepoint("\xE3\x81\x82");
  EXPECT_TRUE(r.valid);
  EXPECT_EQ(r.codepoint, 0x3042u);
  EXPECT_EQ(r.byte_len, 3u);
}

TEST(TextOpsDecodeUtf8, FourByteSupplementary) {
  const auto r = decode_first_utf8_codepoint("\xF0\x9F\x98\x80");
  EXPECT_TRUE(r.valid);
  EXPECT_EQ(r.codepoint, 0x1F600u);
  EXPECT_EQ(r.byte_len, 4u);
}

TEST(TextOpsDecodeUtf8, OnlyFirstCodepointReturned) {
  // Two emojis: first decode returns the first one and reports byte_len=4.
  const auto r = decode_first_utf8_codepoint("\xF0\x9F\x8E\x89\xF0\x9F\x8E\x8A");
  EXPECT_TRUE(r.valid);
  EXPECT_EQ(r.codepoint, 0x1F389u);
  EXPECT_EQ(r.byte_len, 4u);
}

TEST(TextOpsDecodeUtf8, MalformedLeadingByte) {
  // A continuation-style byte (0x80) in lead position is invalid.
  const auto r = decode_first_utf8_codepoint("\x80hello");
  EXPECT_FALSE(r.valid);
}

TEST(TextOpsDecodeUtf8, TruncatedSequence) {
  // A 3-byte lead with only one continuation byte present.
  const auto r = decode_first_utf8_codepoint("\xE3\x81");
  EXPECT_FALSE(r.valid);
}

TEST(TextOpsDecodeUtf8, BadContinuationByte) {
  // 3-byte lead, but the second supposed continuation has the wrong tag.
  const auto r = decode_first_utf8_codepoint("\xE3\x81\x20");
  EXPECT_FALSE(r.valid);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
