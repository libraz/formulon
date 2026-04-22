// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for utils/strings.h.

#include "utils/strings.h"

#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"

namespace formulon {
namespace strings {
namespace {

TEST(StringsTrim, EmptyInputReturnsEmpty) {
  EXPECT_EQ("", trim(""));
  EXPECT_EQ("", ltrim(""));
  EXPECT_EQ("", rtrim(""));
}

TEST(StringsTrim, AllWhitespaceCollapsesToEmpty) {
  EXPECT_EQ("", trim("  \t\r\n\v\f  "));
  EXPECT_EQ("", ltrim(" \t\n"));
  EXPECT_EQ("", rtrim(" \t\n"));
}

TEST(StringsTrim, NoWhitespacePassesThrough) {
  EXPECT_EQ("hello", trim("hello"));
  EXPECT_EQ("hello", ltrim("hello"));
  EXPECT_EQ("hello", rtrim("hello"));
}

TEST(StringsTrim, InternalWhitespacePreserved) {
  EXPECT_EQ("a b\tc", trim("  a b\tc  "));
  EXPECT_EQ("a b\tc  ", ltrim("  a b\tc  "));
  EXPECT_EQ("  a b\tc", rtrim("  a b\tc  "));
}

TEST(StringsSplit, CharDelimiterPreservesEmptyFields) {
  const std::vector<std::string_view> expected = {"", "", "a", ""};
  EXPECT_EQ(expected, split(std::string_view(",,a,"), ','));
}

TEST(StringsSplit, CharDelimiterOnEmptyInputReturnsOneEmpty) {
  const std::vector<std::string_view> expected = {""};
  EXPECT_EQ(expected, split(std::string_view(""), ','));
}

TEST(StringsSplit, CharDelimiterOnlyDelimiters) {
  const std::vector<std::string_view> expected = {"", "", "", ""};
  EXPECT_EQ(expected, split(std::string_view(",,,"), ','));
}

TEST(StringsSplit, StringViewDelimiterSpansMultipleBytes) {
  const std::vector<std::string_view> expected = {"alpha", "beta", "gamma"};
  EXPECT_EQ(expected, split(std::string_view("alpha--beta--gamma"), std::string_view("--")));
}

TEST(StringsSplit, StringViewDelimiterWithTrailingEmpty) {
  const std::vector<std::string_view> expected = {"x", "y", ""};
  EXPECT_EQ(expected, split(std::string_view("x--y--"), std::string_view("--")));
}

TEST(StringsSplit, EmptyDelimiterYieldsInputAsSingleton) {
  const std::vector<std::string_view> expected = {"abc"};
  EXPECT_EQ(expected, split(std::string_view("abc"), std::string_view("")));
}

TEST(StringsJoin, VectorOfStringsBasic) {
  const std::vector<std::string> parts = {"a", "b", "c"};
  EXPECT_EQ("a,b,c", join(parts, ","));
}

TEST(StringsJoin, VectorOfStringViewsBasic) {
  const std::vector<std::string_view> parts = {"alpha", "beta"};
  EXPECT_EQ("alpha::beta", join(parts, "::"));
}

TEST(StringsJoin, EmptyVectorsYieldEmptyString) {
  const std::vector<std::string> empty_s;
  const std::vector<std::string_view> empty_sv;
  EXPECT_EQ("", join(empty_s, ","));
  EXPECT_EQ("", join(empty_sv, ","));
}

TEST(StringsJoin, SingleElementNoSeparator) {
  const std::vector<std::string> parts = {"only"};
  EXPECT_EQ("only", join(parts, ","));
}

TEST(StringsJoin, EmptySeparatorConcatenates) {
  const std::vector<std::string> parts = {"ab", "cd", "ef"};
  EXPECT_EQ("abcdef", join(parts, ""));
}

TEST(StringsCaseInsensitiveEq, MatchesAsciiCaseVariants) {
  EXPECT_TRUE(case_insensitive_eq("Hello", "HELLO"));
  EXPECT_TRUE(case_insensitive_eq("hELLo", "HelLO"));
  EXPECT_TRUE(case_insensitive_eq("", ""));
  EXPECT_TRUE(case_insensitive_eq("SUM", "sum"));
}

TEST(StringsCaseInsensitiveEq, LengthMismatchFailsFast) {
  EXPECT_FALSE(case_insensitive_eq("abc", "abcd"));
  EXPECT_FALSE(case_insensitive_eq("abcd", "abc"));
}

TEST(StringsCaseInsensitiveEq, NonAsciiIsComparedByteForByte) {
  // ASCII letters fold, but the 0xC3 / 0x9F bytes that encode ß do not match
  // 0x53 ('S') / 0x53 ('S'). So "straße" != "STRASSE".
  static constexpr char kStrasze[] = {'s', 't', 'r', 'a', '\xc3', '\x9f', 'e', '\0'};
  EXPECT_FALSE(case_insensitive_eq(kStrasze, "STRASSE"));
  // Same Japanese bytes compare equal to themselves via the verbatim path.
  EXPECT_TRUE(case_insensitive_eq("\xe3\x81\x82\xe3\x81\x84", "\xe3\x81\x82\xe3\x81\x84"));
}

TEST(StringsStartsEndsWith, PrefixLongerThanHaystackReturnsFalse) {
  EXPECT_FALSE(starts_with("ab", "abc"));
  EXPECT_FALSE(ends_with("ab", "xab"));
}

TEST(StringsStartsEndsWith, BasicMatches) {
  EXPECT_TRUE(starts_with("formula=SUM(A1)", "formula="));
  EXPECT_TRUE(ends_with("workbook.xlsx", ".xlsx"));
  EXPECT_FALSE(starts_with("workbook.xlsx", "Workbook"));
  EXPECT_FALSE(ends_with("workbook.xlsx", ".xlsb"));
}

TEST(StringsStartsEndsWith, EmptyPrefixOrSuffixAlwaysMatches) {
  EXPECT_TRUE(starts_with("anything", ""));
  EXPECT_TRUE(ends_with("anything", ""));
  EXPECT_TRUE(starts_with("", ""));
  EXPECT_TRUE(ends_with("", ""));
}

TEST(StringsAsciiCase, ConstexprHelpersFoldOnlyAsciiLetters) {
  static_assert(ascii_to_lower('A') == 'a', "A -> a");
  static_assert(ascii_to_lower('a') == 'a', "a -> a");
  static_assert(ascii_to_lower('7') == '7', "digits unchanged");
  static_assert(ascii_to_upper('z') == 'Z', "z -> Z");
  static_assert(ascii_to_upper('Z') == 'Z', "Z -> Z");
  static_assert(ascii_to_upper('!') == '!', "punct unchanged");
  SUCCEED();
}

TEST(StringsAsciiCase, ToAsciiLowerAndUpperPassThroughNonAscii) {
  EXPECT_EQ("hello, world!", to_ascii_lower("Hello, World!"));
  EXPECT_EQ("HELLO, WORLD!", to_ascii_upper("Hello, World!"));

  // Japanese bytes (UTF-8) are preserved verbatim because every byte has the
  // high bit set and is therefore outside the ASCII letter ranges.
  const std::string ja = "\xe3\x81\x82\xe3\x81\x84";  // "あい"
  EXPECT_EQ(ja, to_ascii_lower(ja));
  EXPECT_EQ(ja, to_ascii_upper(ja));
}

}  // namespace
}  // namespace strings
}  // namespace formulon
