// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the standalone UTF-16 unit counter used by `LEN` and
// related text functions.

#include "eval/utf8_length.h"

#include <cstdint>
#include <string_view>

#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace {

TEST(Utf8Length, EmptyString) {
  EXPECT_EQ(utf16_units_in(std::string_view{}), 0u);
}

TEST(Utf8Length, AsciiOnly) {
  EXPECT_EQ(utf16_units_in("hello"), 5u);
  EXPECT_EQ(utf16_units_in("a"), 1u);
  EXPECT_EQ(utf16_units_in("123 ABC"), 7u);
}

TEST(Utf8Length, MultibyteBmp) {
  // "あいう" = 3 BMP codepoints, each 3 bytes in UTF-8 but 1 UTF-16 unit.
  EXPECT_EQ(utf16_units_in("\xE3\x81\x82\xE3\x81\x84\xE3\x81\x86"), 3u);
}

TEST(Utf8Length, SupplementaryPlane) {
  // "🎉" U+1F389 = 4 bytes in UTF-8, 2 UTF-16 units (surrogate pair).
  EXPECT_EQ(utf16_units_in("\xF0\x9F\x8E\x89"), 2u);
}

TEST(Utf8Length, MixedAsciiBmpSupplementary) {
  // "a" + "あ" + "🎉" + "z" = 1 + 1 + 2 + 1 = 5 units.
  EXPECT_EQ(utf16_units_in("a\xE3\x81\x82\xF0\x9F\x8E\x89z"), 5u);
}

TEST(Utf8Length, MalformedLeadingByteCountsAsOne) {
  // 0xC0 with no continuation bytes is malformed; the helper must still
  // make forward progress. We expect each malformed byte to count as one.
  const char malformed[] = {static_cast<char>(0xC0), 'a', '\0'};
  EXPECT_EQ(utf16_units_in(std::string_view(malformed, 2)), 2u);
}

TEST(Utf8Length, TruncatedMultibyteCountsAsOne) {
  // 0xE3 wants two continuation bytes; provide only one then EOF.
  const char truncated[] = {static_cast<char>(0xE3), static_cast<char>(0x81), '\0'};
  // The leading byte is treated as one unit and we then fall off the end.
  EXPECT_EQ(utf16_units_in(std::string_view(truncated, 2)), 2u);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
