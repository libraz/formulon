//
// Tests for `derived_guid`, the deterministic GUID a sheet copy assigns in
// place of a source GUID.

#include "utils/derived_guid.h"

#include <cctype>
#include <string>

#include "gtest/gtest.h"
#include "threaded_comment.h"

namespace formulon {
namespace {

constexpr const char* kSource = "{6B4D1F2A-0C3E-4A5B-9D8E-7F6A5B4C3D2E}";

TEST(DerivedGuid, IsDeterministic) {
  EXPECT_EQ(derived_guid(kSource, "Src (2)"), derived_guid(kSource, "Src (2)"));
}

TEST(DerivedGuid, MatchesFnv1a128Reference) {
  // FNV-1a 128 over source, NUL, salt; version and variant bits then forced.
  EXPECT_EQ(derived_guid(kSource, "Src (2)"), "{539FD18E-7427-4055-8BF9-34BBAE2E89A3}");
}

TEST(DerivedGuid, BracedInputKeepsBracedUppercaseFormat) {
  const std::string guid = derived_guid(kSource, "Src (2)");
  ASSERT_EQ(guid.size(), 38U);
  EXPECT_EQ(guid.front(), '{');
  EXPECT_EQ(guid.back(), '}');
  for (std::size_t i = 1; i + 1 < guid.size(); ++i) {
    if (i == 9U || i == 14U || i == 19U || i == 24U) {
      EXPECT_EQ(guid[i], '-') << guid;
    } else {
      EXPECT_TRUE(std::isxdigit(static_cast<unsigned char>(guid[i])) != 0 &&
                  std::islower(static_cast<unsigned char>(guid[i])) == 0)
          << guid;
    }
  }
  // RFC 4122 version 4 and variant 10xx.
  EXPECT_EQ(guid[15], '4') << guid;
  EXPECT_NE(std::string("89AB").find(guid[20]), std::string::npos) << guid;
}

TEST(DerivedGuid, UnbracedInputKeepsUnbracedFormat) {
  const std::string guid = derived_guid("6B4D1F2A-0C3E-4A5B-9D8E-7F6A5B4C3D2E", "Src (2)");
  ASSERT_EQ(guid.size(), 36U);
  EXPECT_NE(guid.front(), '{');
  EXPECT_EQ(guid[8], '-');
  EXPECT_EQ(guid[23], '-');
}

TEST(DerivedGuid, AcceptedAsThreadedCommentGuid) {
  EXPECT_TRUE(is_threaded_comment_guid(derived_guid(kSource, "Src (2)")));
  EXPECT_TRUE(is_threaded_comment_guid(derived_guid(kSource, "")));
}

TEST(DerivedGuid, DifferentSaltGivesDifferentValue) {
  EXPECT_NE(derived_guid(kSource, "Src (2)"), derived_guid(kSource, "Src (3)"));
}

TEST(DerivedGuid, DifferentSourceGivesDifferentValue) {
  EXPECT_NE(derived_guid(kSource, "Src (2)"), derived_guid("{6B4D1F2A-0C3E-4A5B-9D8E-7F6A5B4C3D2F}", "Src (2)"));
}

TEST(DerivedGuid, SeparatorKeepsSourceAndSaltApart) {
  EXPECT_NE(derived_guid("{AB", "C}"), derived_guid("{A", "BC}"));
}

TEST(DerivedGuid, DiffersFromSource) {
  EXPECT_NE(derived_guid(kSource, "Src (2)"), kSource);
}

}  // namespace
}  // namespace formulon
