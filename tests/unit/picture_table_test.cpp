// Unit tests for the picture table and for picture equality and order.

#include "picture_table.h"

#include <cstdint>
#include <optional>
#include <string>

#include "eval/eval_profile_scope.h"
#include "eval/picture_compare.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "value.h"

namespace formulon {
namespace {

using eval::pictures_equal;
using eval::pictures_order;
using eval::ScopedEvalProfile;
using eval::ScopedPictureTable;

PicturePayload payload_of(const PictureTable& table, std::uint32_t picture_id) {
  return PicturePayload{picture_id, table.entry(picture_id)->media_id};
}

TEST(PictureTable, AddMediaAtShiftsRanksAtAndAbovePosition) {
  PictureTable table;
  const std::uint32_t m0 = table.add_media_at(0, MediaKey{1, 10});
  const std::uint32_t m1 = table.add_media_at(0, MediaKey{2, 20});
  EXPECT_EQ(0U, m0);
  EXPECT_EQ(1U, m1);
  EXPECT_EQ(0U, table.media_rank(m1));
  EXPECT_EQ(1U, table.media_rank(m0));

  // Appending at the end leaves every earlier rank alone.
  const std::uint32_t m2 = table.add_media_at(2, MediaKey{3, 30});
  EXPECT_EQ(0U, table.media_rank(m1));
  EXPECT_EQ(1U, table.media_rank(m0));
  EXPECT_EQ(2U, table.media_rank(m2));

  const std::uint32_t m3 = table.add_media_at(1, MediaKey{4, 40});
  EXPECT_EQ(0U, table.media_rank(m1));
  EXPECT_EQ(1U, table.media_rank(m3));
  EXPECT_EQ(2U, table.media_rank(m0));
  EXPECT_EQ(3U, table.media_rank(m2));
  EXPECT_EQ(4U, table.media_count());
}

TEST(PictureTable, AddEntryAndLookup) {
  PictureTable table;
  const std::uint32_t media = table.add_media_at(0, MediaKey{3, 7}, "xl/media/image1.png");
  EXPECT_EQ(0U, table.size());
  const std::uint32_t plain = table.add_entry(media, std::nullopt);
  const std::uint32_t titled = table.add_entry(media, std::string("Blue circle"));
  EXPECT_EQ(0U, plain);
  EXPECT_EQ(1U, titled);
  EXPECT_EQ(2U, table.size());

  const PictureEntry* e = table.entry(titled);
  ASSERT_NE(nullptr, e);
  EXPECT_EQ(media, e->media_id);
  EXPECT_EQ("xl/media/image1.png", e->media_path);
  ASSERT_TRUE(e->alt.has_value());
  EXPECT_EQ("Blue circle", *e->alt);
  ASSERT_NE(nullptr, table.entry(plain));
  EXPECT_FALSE(table.entry(plain)->alt.has_value());
  EXPECT_EQ(nullptr, table.entry(2));
}

TEST(PictureTable, OrderComparesUnsignedBytes) {
  const std::uint8_t png[] = {0x89, 0x50};
  const std::uint8_t jpeg[] = {0xFF, 0xD8};
  // The greater byte sorts first, compared unsigned.
  EXPECT_EQ(-1, compare_media_bytes(jpeg, 2, png, 2));
  EXPECT_EQ(1, compare_media_bytes(png, 2, jpeg, 2));
  const std::uint8_t low[] = {0x01, 0x7F};
  const std::uint8_t high[] = {0x01, 0x80};
  EXPECT_EQ(-1, compare_media_bytes(high, 2, low, 2));
  EXPECT_EQ(0, compare_media_bytes(png, 2, png, 2));
}

TEST(PictureTable, OrderPutsTheLongerOfAPrefixPairFirst) {
  const std::uint8_t longer[] = {0x10, 0x20, 0x00};
  const std::uint8_t shorter[] = {0x10, 0x20};
  EXPECT_EQ(-1, compare_media_bytes(longer, 3, shorter, 2));
  EXPECT_EQ(1, compare_media_bytes(shorter, 2, longer, 3));
  EXPECT_EQ(0, compare_media_bytes(shorter, 0, longer, 0));
}

TEST(PictureTableCompare, EqualityFoldsAltCaseAndKeysOnMedia) {
  PictureTable table;
  const std::uint32_t m0 = table.add_media_at(0, MediaKey{1, 1});
  const std::uint32_t m1 = table.add_media_at(1, MediaKey{2, 2});
  const PicturePayload title = payload_of(table, table.add_entry(m0, std::string("Picture")));
  const PicturePayload upper = payload_of(table, table.add_entry(m0, std::string("PICTURE")));
  const PicturePayload other = payload_of(table, table.add_entry(m0, std::string("Other")));
  const PicturePayload other_media = payload_of(table, table.add_entry(m1, std::string("Picture")));

  ScopedEvalProfile profile(mac_365_ja_jp_profile());
  ScopedPictureTable scope(&table);
  EXPECT_TRUE(pictures_equal(title, title));
  EXPECT_TRUE(pictures_equal(title, upper));
  EXPECT_FALSE(pictures_equal(title, other));
  EXPECT_FALSE(pictures_equal(title, other_media));
}

TEST(PictureTableCompare, AbsentAndPresentAltDiffer) {
  PictureTable table;
  const std::uint32_t m0 = table.add_media_at(0, MediaKey{1, 1});
  const PicturePayload plain = payload_of(table, table.add_entry(m0, std::nullopt));
  const PicturePayload plain2 = payload_of(table, table.add_entry(m0, std::nullopt));
  const PicturePayload titled = payload_of(table, table.add_entry(m0, std::string("a")));

  ScopedEvalProfile profile(mac_365_ja_jp_profile());
  ScopedPictureTable scope(&table);
  EXPECT_TRUE(pictures_equal(plain, plain2));
  EXPECT_FALSE(pictures_equal(plain, titled));
  EXPECT_FALSE(pictures_equal(titled, plain));
}

TEST(PictureTableCompare, OrderByRankThenAltPresenceThenAlt) {
  PictureTable table;
  const std::uint32_t second = table.add_media_at(0, MediaKey{1, 1});
  const std::uint32_t first = table.add_media_at(0, MediaKey{2, 2});
  const PicturePayload second_plain = payload_of(table, table.add_entry(second, std::nullopt));
  const PicturePayload first_plain = payload_of(table, table.add_entry(first, std::nullopt));
  const PicturePayload first_b = payload_of(table, table.add_entry(first, std::string("b")));
  const PicturePayload first_a = payload_of(table, table.add_entry(first, std::string("a")));

  ScopedEvalProfile profile(mac_365_ja_jp_profile());
  ScopedPictureTable scope(&table);
  // Media rank first, whatever the alts.
  EXPECT_EQ(-1, pictures_order(first_plain, second_plain));
  EXPECT_EQ(1, pictures_order(second_plain, first_a));
  // A present alt before an absent one.
  EXPECT_EQ(-1, pictures_order(first_b, first_plain));
  EXPECT_EQ(1, pictures_order(first_plain, first_a));
  // Two present alts in collation order.
  EXPECT_EQ(-1, pictures_order(first_a, first_b));
  EXPECT_EQ(1, pictures_order(first_b, first_a));
  EXPECT_EQ(0, pictures_order(first_a, first_a));
}

TEST(PictureTableCompare, NullInnerScopeKeepsTheOuterTable) {
  PictureTable outer;
  PictureTable inner;
  EXPECT_EQ(nullptr, eval::current_picture_table());
  {
    ScopedPictureTable a(&outer);
    EXPECT_EQ(&outer, eval::current_picture_table());
    {
      ScopedPictureTable keep(nullptr);
      EXPECT_EQ(&outer, eval::current_picture_table());
    }
    {
      ScopedPictureTable replace(&inner);
      EXPECT_EQ(&inner, eval::current_picture_table());
    }
    EXPECT_EQ(&outer, eval::current_picture_table());
  }
  EXPECT_EQ(nullptr, eval::current_picture_table());
}

TEST(PictureTableCompare, UnknownIdFallsBackToPictureId) {
  PictureTable table;
  const std::uint32_t m0 = table.add_media_at(0, MediaKey{1, 1});
  const PicturePayload known = payload_of(table, table.add_entry(m0, std::nullopt));
  const PicturePayload unknown{7, m0};
  const PicturePayload unknown_higher{9, m0};

  ScopedEvalProfile profile(mac_365_ja_jp_profile());
  ScopedPictureTable scope(&table);
  EXPECT_FALSE(pictures_equal(known, unknown));
  EXPECT_TRUE(pictures_equal(unknown, unknown));
  EXPECT_EQ(-1, pictures_order(known, unknown));
  EXPECT_EQ(1, pictures_order(unknown_higher, unknown));
}

TEST(PictureTableCompare, NoScopeFallsBackToPictureId) {
  ASSERT_EQ(nullptr, eval::current_picture_table());
  const PicturePayload a{1, 0};
  const PicturePayload b{2, 0};
  (void)b;
  EXPECT_TRUE(pictures_equal(a, a));
  EXPECT_EQ(0, pictures_order(a, a));
#if defined(NDEBUG)
  EXPECT_FALSE(pictures_equal(a, b));
  EXPECT_EQ(-1, pictures_order(a, b));
  EXPECT_EQ(1, pictures_order(b, a));
#elif GTEST_HAS_DEATH_TEST
  // A Debug build flags two different ids compared with no table published.
  EXPECT_DEATH(pictures_equal(a, b), "ScopedPictureTable");
  EXPECT_DEATH(pictures_order(a, b), "ScopedPictureTable");
#endif
}

}  // namespace
}  // namespace formulon
