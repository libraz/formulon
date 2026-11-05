// Tests for worksheet identity's strict UTF-8 Unicode simple fold.

#include "sheet_name.h"

#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "sheet_name_casefold_data.h"
#include "utils/text_ops.h"

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
  for (const std::string malformed :
       {"\xC0\x80", "\xC3", "\xE0\x80\x80", "\xED\xA0\x80", "\xF4\x90\x80\x80", "\xF5\x80\x80\x80", "\x80"}) {
    EXPECT_FALSE(valid_utf8(malformed));
    EXPECT_FALSE(equal(malformed, malformed));
  }
  EXPECT_FALSE(equal("\xC3", "\xC3\xA4"));
}

TEST(SheetNameIdentity, AcceptsMaximumUnicodeScalar) {
  EXPECT_TRUE(valid_utf8("\xF4\x8F\xBF\xBF"));
  EXPECT_TRUE(equal("\xF4\x8F\xBF\xBF", "\xF4\x8F\xBF\xBF"));
}

// Scalar-by-scalar reference: strict decode, simple fold, no byte-level shortcuts.
bool reference_equal(const std::string& lhs, const std::string& rhs) {
  std::size_t l = 0;
  std::size_t r = 0;
  while (l < lhs.size() && r < rhs.size()) {
    const Utf8DecodeResult a = decode_first_utf8_codepoint(std::string_view(lhs).substr(l));
    const Utf8DecodeResult b = decode_first_utf8_codepoint(std::string_view(rhs).substr(r));
    if (!a.valid || !b.valid || detail::fold_scalar(a.codepoint) != detail::fold_scalar(b.codepoint)) {
      return false;
    }
    l += a.byte_len;
    r += b.byte_len;
  }
  return l == lhs.size() && r == rhs.size();
}

TEST(SheetNameIdentity, AsciiFastPathMatchesScalarFold) {
  for (int a = 0; a < 128; ++a) {
    for (int b = 0; b < 128; ++b) {
      const std::string x(1, static_cast<char>(a));
      const std::string y(1, static_cast<char>(b));
      EXPECT_EQ(equal(x, y), reference_equal(x, y)) << a << "," << b;
    }
  }
  const std::vector<std::string> corpus = {"",
                                           "Sheet1",
                                           "sheet1",
                                           "SHEET1",
                                           "Sheet12",
                                           "Sheet",
                                           "Sheet1\xC3\xA4",
                                           "sheet1\xC3\x84",
                                           "Sheet1\xFF",
                                           "sheet1\xFF",
                                           "\xE2\x84\xAA",
                                           "K",
                                           "k",
                                           "\xE2\x84\xAA"
                                           "elvin",
                                           "Kelvin",
                                           "kelvin",
                                           "a\xC3\x84"
                                           "b",
                                           "A\xC3\xA4"
                                           "B",
                                           "a\xC3\x84",
                                           "a\xC3",
                                           "A\xC3",
                                           "\xC3\xA4",
                                           "\xC3\x84",
                                           "Data \xE6\x97\xA5",
                                           "data \xE6\x97\xA5",
                                           "DATA \xE6\x97\xA5\xE6",
                                           "Z[",
                                           "z{",
                                           "@`",
                                           "`@"};
  for (const std::string& x : corpus) {
    for (const std::string& y : corpus) {
      EXPECT_EQ(equal(x, y), reference_equal(x, y)) << x << " vs " << y;
    }
  }
}

}  // namespace
}  // namespace sheet_names
}  // namespace formulon
