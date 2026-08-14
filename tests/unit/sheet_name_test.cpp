// Tests for worksheet identity's strict UTF-8 Unicode simple fold.

#include "sheet_name.h"

#include <string>

#include "gtest/gtest.h"

namespace formulon {
namespace sheet_names {
namespace {

TEST(SheetNameIdentity, FoldsAsciiAndLatin) {
  EXPECT_TRUE(equal("Data", "data"));
  EXPECT_TRUE(equal("\xC3\x84", "\xC3\xA4"));  // Ä / ä
}

TEST(SheetNameIdentity, FoldsSigmaVariants) {
  EXPECT_TRUE(equal("\xCE\xA3", "\xCF\x83"));  // Σ / σ
  EXPECT_TRUE(equal("\xCE\xA3", "\xCF\x82"));  // Σ / ς
  EXPECT_TRUE(equal("\xCF\x83", "\xCF\x82"));  // σ / ς
}

TEST(SheetNameIdentity, FoldsSupplementaryDeseretPair) {
  EXPECT_TRUE(equal("\xF0\x90\x90\x80", "\xF0\x90\x90\xA8"));  // U+10400 / U+10428
}

TEST(SheetNameIdentity, LeavesNonCaseDifferencesDistinct) {
  EXPECT_FALSE(equal("\xC3\x9F", "ss"));         // ß / ss: no expansion
  EXPECT_FALSE(equal("\xC3\x85", "A\xCC\x8A"));  // Å / A + combining ring
  EXPECT_FALSE(equal("\xC4\xB0", "i"));          // İ / i: no full fold
  EXPECT_FALSE(equal("\xC4\xB1", "I"));          // ı / I: no locale tailoring
  EXPECT_TRUE(equal("\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E", "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E"));
}

TEST(SheetNameIdentity, RejectsMalformedUtf8) {
  for (const std::string malformed : {"\xC3", "\xE0\x80\x80", "\xED\xA0\x80", "\xF4\x90\x80\x80", "\x80"}) {
    EXPECT_FALSE(valid_utf8(malformed));
    EXPECT_FALSE(equal(malformed, malformed));
  }
  EXPECT_FALSE(equal("\xC3", "\xC3\xA4"));
}

}  // namespace
}  // namespace sheet_names
}  // namespace formulon
