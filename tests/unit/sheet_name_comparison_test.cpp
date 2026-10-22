// Edge cases for worksheet-name folding while comparing UTF-8 byte streams.

#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "sheet_name.h"

namespace formulon {
namespace sheet_names {
namespace {

TEST(SheetNameComparison, FoldsAcrossDifferentUtf8ByteWidths) {
  const std::string lhs = "\xE2\x84\xAA\xF0\x90\x90\x80\xCE\xA3";  // Kelvin, Deseret capital, sigma.
  const std::string rhs = "k\xF0\x90\x90\xA8\xCF\x82";             // k, Deseret small, final sigma.

  ASSERT_NE(lhs.size(), rhs.size());
  EXPECT_TRUE(equal(lhs, rhs));
}

TEST(SheetNameComparison, ComparesBytesAfterEmbeddedNul) {
  std::string lhs = "A";
  lhs.push_back('\0');
  lhs += "Sheet";
  std::string rhs = "a";
  rhs.push_back('\0');
  rhs += "sHEET";

  EXPECT_TRUE(equal(lhs, rhs));
  rhs.back() = 'x';
  EXPECT_FALSE(equal(lhs, rhs));
}

TEST(SheetNameComparison, HandlesEmptyInputs) {
  EXPECT_TRUE(equal("", ""));
  EXPECT_FALSE(equal("", "a"));
  EXPECT_FALSE(equal("a", ""));
}

TEST(SheetNameComparison, RejectsMalformedPrefixAndSuffix) {
  const std::vector<std::string> malformed_encodings = {
      "\xC0\x80",          // Overlong 2-byte sequence.
      "\xE0\x80\x80",      // Overlong 3-byte sequence.
      "\xF0\x80\x80\x80",  // Overlong 4-byte sequence.
      "\xED\xA0\x80",      // UTF-8 encoded surrogate.
      "\xF4\x90\x80\x80",  // Above U+10FFFF.
      "\xE2\x82",          // Truncated sequence.
  };

  for (const std::string& malformed : malformed_encodings) {
    const std::string malformed_prefix = malformed + "Sheet";
    const std::string folded_prefix = malformed + "sheet";
    const std::string malformed_suffix = std::string("Sheet") + malformed;
    const std::string folded_suffix = std::string("sheet") + malformed;

    EXPECT_FALSE(equal(malformed_prefix, malformed_prefix));
    EXPECT_FALSE(equal(malformed_prefix, folded_prefix));
    EXPECT_FALSE(equal(malformed_suffix, malformed_suffix));
    EXPECT_FALSE(equal(malformed_suffix, folded_suffix));
  }
}

}  // namespace
}  // namespace sheet_names
}  // namespace formulon
