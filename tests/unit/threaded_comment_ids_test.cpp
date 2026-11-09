#include <cctype>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "threaded_comment.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::string_view kAlice = "{11111111-AAAA-4BBB-8CCC-DDDDDDDDDDDD}";

bool IsBracedUpperGuid(std::string_view s) {
  if (!is_threaded_comment_guid(s)) {
    return false;
  }
  for (const char ch : s) {
    if (std::islower(static_cast<unsigned char>(ch)) != 0) {
      return false;
    }
  }
  return true;
}

Person Named(std::string display_name) {
  Person p;
  p.display_name = std::move(display_name);
  return p;
}

TEST(PersonIdGeneration, GeneratesBracedUpperGuidWhenIdEmpty) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.add_person(Named("Alice Smith"))));
  ASSERT_EQ(wb.persons().size(), 1U);
  const std::string first = wb.persons().back().id;
  EXPECT_TRUE(IsBracedUpperGuid(first)) << first;

  ASSERT_TRUE(static_cast<bool>(wb.add_person(Named("Bob Jones"))));
  ASSERT_EQ(wb.persons().size(), 2U);
  EXPECT_TRUE(IsBracedUpperGuid(wb.persons().back().id));
  EXPECT_NE(wb.persons().back().id, first);
}

TEST(PersonIdGeneration, PinnedClockMakesIdDeterministic) {
  const date_time::CivilTime pinned{{2026, 11, 9}, {8, 5, 3}};
  Workbook a = Workbook::create();
  a.set_pinned_now(pinned);
  Workbook b = Workbook::create();
  b.set_pinned_now(pinned);
  ASSERT_TRUE(static_cast<bool>(a.add_person(Named("Alice Smith"))));
  ASSERT_TRUE(static_cast<bool>(b.add_person(Named("Alice Smith"))));
  ASSERT_EQ(a.persons().size(), 1U);
  ASSERT_EQ(b.persons().size(), 1U);
  EXPECT_TRUE(IsBracedUpperGuid(a.persons().back().id));
  EXPECT_EQ(a.persons().back().id, b.persons().back().id);
}

TEST(MentionIdGeneration, EmptyMentionIdGenerated) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.add_person(Person{std::string(kAlice), "Alice Smith", "", "None"})));
  ThreadedComment c;
  c.row = 0;
  c.col = 0;
  c.person_id = std::string(kAlice);
  c.created = "2026-10-05T10:00:00.00";
  c.text = "@Alice Smith hi";
  Mention m;
  m.person_id = std::string(kAlice);
  m.start = 0;
  m.length = 12;
  c.mentions.push_back(m);
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, c)));
  ASSERT_EQ(wb.sheet(0).threaded_comments().size(), 1U);
  const std::vector<Mention>& stored = wb.sheet(0).threaded_comments()[0].mentions;
  ASSERT_EQ(stored.size(), 1U);
  EXPECT_TRUE(IsBracedUpperGuid(stored[0].mention_id)) << stored[0].mention_id;
}

}  // namespace
}  // namespace formulon
