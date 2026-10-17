//
// Tests for the simple Unicode case mapping behind UPPER, LOWER and PROPER.

#include "utils/case_map.h"

#include <string>

#include "gtest/gtest.h"

namespace formulon {
namespace {

TEST(CaseMap, UpperCoversLatinGreekCyrillicAndFullwidth) {
  EXPECT_EQ(case_map_to_upper("abc \xC3\xA9"), "ABC \xC3\x89");                          // é -> É
  EXPECT_EQ(case_map_to_upper("\xCE\xB1\xCF\x83\xCF\x82"), "\xCE\x91\xCE\xA3\xCE\xA3");  // ασς -> ΑΣΣ
  EXPECT_EQ(case_map_to_upper("\xD1\x89"), "\xD0\xA9");                                  // щ -> Щ
  EXPECT_EQ(case_map_to_upper("\xEF\xBD\x81"), "\xEF\xBC\xA1");                          // ａ -> Ａ
  EXPECT_EQ(case_map_to_upper("\xC4\xB1"), "I");                                         // ı -> I
}

TEST(CaseMap, MappingsNeverChangeLength) {
  for (const char* same : {"\xC3\x9F", "\xEF\xAC\x81", "\xC5\x89", "\xC7\xB0", "\xE1\xBE\xB3"}) {
    EXPECT_EQ(case_map_to_upper(same), same);
    EXPECT_EQ(case_map_to_lower(same), same);
  }
}

TEST(CaseMap, TitlecaseDigraphsGoToFullUppercase) {
  EXPECT_EQ(case_map_to_upper("\xC7\x85"), "\xC7\x84");  // ǅ -> Ǆ
  EXPECT_EQ(case_map_to_lower("\xC7\x85"), "\xC7\x86");  // ǅ -> ǆ
  EXPECT_EQ(case_map_to_proper("\xC7\x86"
                               "a"),
            "\xC7\x84"
            "a");
}

TEST(CaseMap, DottedCapitalILowersToAscii) {
  EXPECT_EQ(case_map_to_lower("\xC4\xB0"), "i");
  EXPECT_EQ(case_map_to_proper("\xC4\xB0"), "I");
}

TEST(CaseMap, FinalSigmaFollowsTheNextCharacter) {
  const std::string sigma = "\xCE\xA3";
  const std::string alpha = "\xCE\x91";
  EXPECT_EQ(case_map_to_lower(sigma), "\xCF\x82");
  EXPECT_EQ(case_map_to_lower(sigma + alpha), "\xCF\x83\xCE\xB1");
  EXPECT_EQ(case_map_to_lower(alpha + sigma), "\xCE\xB1\xCF\x82");
  EXPECT_EQ(case_map_to_lower(alpha + sigma + "."), "\xCE\xB1\xCF\x82.");
  EXPECT_EQ(case_map_to_lower(alpha + sigma + " " + alpha), "\xCE\xB1\xCF\x82 \xCE\xB1");
  EXPECT_EQ(case_map_to_lower(alpha + sigma + "("), "\xCE\xB1\xCF\x83(");
  EXPECT_EQ(case_map_to_lower(alpha + sigma + "1"),
            "\xCE\xB1\xCF\x83"
            "1");
  EXPECT_EQ(case_map_to_lower(alpha + sigma + sigma), "\xCE\xB1\xCF\x83\xCF\x82");
  EXPECT_EQ(case_map_lower_isolated(0x03A3U), 0x03C2U);
}

TEST(CaseMap, ProperTreatsNonAsciiLettersAsWordCharacters) {
  EXPECT_EQ(case_map_to_proper("\xC3\xA9"
                               "a a"
                               "\xC3\xA9"),
            "\xC3\x89"
            "a A\xC3\xA9");
  EXPECT_EQ(case_map_to_proper("a1b"), "A1B");
  EXPECT_EQ(case_map_to_proper("a'b"), "A'B");
  EXPECT_EQ(case_map_to_proper("\xE6\x97\xA5"
                               "A"),
            "\xE6\x97\xA5"
            "a");  // 日A: an ideograph is a word character
  EXPECT_EQ(case_map_to_proper("\xCE\xAC"
                               "\xCE\x91"),
            "\xCE\x86"
            "\xCE\xB1");  // άΑ -> Άα keeps the accent
}

TEST(CaseMap, InvalidBytesPassThrough) {
  EXPECT_EQ(case_map_to_upper("a\xFF"
                              "b"),
            "A\xFF"
            "B");
  EXPECT_EQ(case_map_to_lower("A\xFF"
                              "B"),
            "a\xFF"
            "b");
}

}  // namespace
}  // namespace formulon
