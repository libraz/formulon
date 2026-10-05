//
// Stable C ABI tests for threaded comments and the workbook person list:
// record shapes, CRUD, the note/thread exclusivity, and a save and reload.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_mention) == (kWasm32 ? 16U : 24U), "fm_mention ABI layout changed");
static_assert(sizeof(fm_threaded_comment) == (kWasm32 ? 40U : 72U), "fm_threaded_comment ABI layout changed");
static_assert(sizeof(fm_person) == 4U * sizeof(void*), "fm_person ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kThreadedCommentInvalid =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kThreadedCommentInvalid);

constexpr const char* kAlice = "{0A1B2C3D-0000-4000-8000-000000000001}";
constexpr const char* kBob = "{0A1B2C3D-0000-4000-8000-000000000002}";
constexpr const char* kThread = "{6F1E0D2C-1111-4111-8111-000000000001}";
constexpr const char* kReply = "{6F1E0D2C-1111-4111-8111-000000000002}";
constexpr const char* kMention = "{6F1E0D2C-1111-4111-8111-0000000000AA}";
constexpr const char* kCreated = "2026-01-02T03:04:05.67";

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

void SaveAndReload(fm_workbook_t* wb, WorkbookGuard* out) {
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, &bytes, &len), 0) << fm_last_error_message();
  const fm_status_t rc = fm_workbook_load(bytes, len, &out->handle);
  fm_buffer_free(bytes);
  ASSERT_EQ(rc, 0) << fm_last_error_message();
}

std::string Str(const char* text) {
  return text != nullptr ? std::string(text) : std::string();
}

/// Copy of one comment record whose views may die with the next read.
struct CommentCopy {
  std::string id, person_id, created, text, parent_id;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  int32_t done = 0;
  struct MentionCopy {
    std::string person_id, mention_id;
    std::uint32_t start = 0;
    std::uint32_t length = 0;
  };
  std::vector<MentionCopy> mentions;
};

std::vector<CommentCopy> ReadComments(const fm_workbook_t* wb) {
  std::vector<CommentCopy> out;
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_threaded_comment_count(wb, 0, &count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    fm_threaded_comment c{};
    EXPECT_EQ(fm_sheet_threaded_comment_at(wb, 0, i, &c), 0) << fm_last_error_message();
    CommentCopy copy;
    copy.id = Str(c.id);
    copy.person_id = Str(c.person_id);
    copy.created = Str(c.created);
    copy.text = Str(c.text);
    copy.parent_id = Str(c.parent_id);
    copy.row = c.row;
    copy.col = c.col;
    copy.done = c.done;
    for (std::uint32_t m = 0; m < c.mention_count; ++m) {
      copy.mentions.push_back(
          {Str(c.mentions[m].person_id), Str(c.mentions[m].mention_id), c.mentions[m].start, c.mentions[m].length});
    }
    out.push_back(copy);
  }
  return out;
}

std::vector<std::string> PersonNames(const fm_workbook_t* wb) {
  std::vector<std::string> names;
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_person_count(wb, &count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    fm_person p{};
    EXPECT_EQ(fm_workbook_person_at(wb, i, &p), 0);
    names.push_back(Str(p.display_name));
  }
  return names;
}

void AddPeople(fm_workbook_t* wb) {
  const fm_person alice{kAlice, "Alice", "alice@example.com", "None"};
  const fm_person bob{kBob, "Bob", "bob@example.com", "None"};
  ASSERT_EQ(fm_workbook_add_person(wb, &alice), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_add_person(wb, &bob), 0) << fm_last_error_message();
}

fm_threaded_comment Comment(const char* id, std::uint32_t row, std::uint32_t col, const char* person, const char* text,
                            const char* parent) {
  fm_threaded_comment c{};
  c.id = id;
  c.row = row;
  c.col = col;
  c.person_id = person;
  c.created = kCreated;
  c.text = text;
  c.parent_id = parent;
  return c;
}

/// A thread on B2 by Alice mentioning Bob, and Bob's reply.
void AddThread(fm_workbook_t* wb) {
  const fm_mention mention{kBob, kMention, 0, 4};
  fm_threaded_comment opening = Comment(kThread, 1, 1, kAlice, "@Bob please check", nullptr);
  opening.mentions = &mention;
  opening.mention_count = 1;
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb, 0, &opening), 0) << fm_last_error_message();
  // A reply takes the thread's anchor, whatever coordinates it carries.
  const fm_threaded_comment reply = Comment(kReply, 7, 7, kBob, "Done", kThread);
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb, 0, &reply), 0) << fm_last_error_message();
}

TEST(FormulonCApiThreadedComments, PersonsAddListAndRemove) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  AddPeople(wb.handle);
  EXPECT_EQ(PersonNames(wb.handle), (std::vector<std::string>{"Alice", "Bob"}));
  fm_person p{};
  ASSERT_EQ(fm_workbook_person_at(wb.handle, 1, &p), 0);
  EXPECT_STREQ(p.id, kBob);
  EXPECT_STREQ(p.user_id, "bob@example.com");
  EXPECT_STREQ(p.provider_id, "None");
  EXPECT_EQ(fm_workbook_person_at(wb.handle, 2, &p), kInvalidArgument);

  const fm_person duplicate{kAlice, "Again", nullptr, nullptr};
  EXPECT_EQ(fm_workbook_add_person(wb.handle, &duplicate), kThreadedCommentInvalid);
  const fm_person lowercase{"{0a1b2c3d-0000-4000-8000-000000000009}", "Low", nullptr, nullptr};
  EXPECT_EQ(fm_workbook_add_person(wb.handle, &lowercase), kThreadedCommentInvalid);
  EXPECT_EQ(fm_workbook_add_person(wb.handle, nullptr), kBindingNullPointer);

  ASSERT_EQ(fm_workbook_remove_person(wb.handle, kAlice), 0);
  EXPECT_EQ(PersonNames(wb.handle), (std::vector<std::string>{"Bob"}));
  EXPECT_EQ(fm_workbook_remove_person(wb.handle, kAlice), kThreadedCommentInvalid);
}

TEST(FormulonCApiThreadedComments, ThreadRepliesAndMentionsRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  AddPeople(wb.handle);
  AddThread(wb.handle);
  ASSERT_EQ(fm_sheet_set_thread_resolved(wb.handle, 0, kThread, 1), 0) << fm_last_error_message();
  EXPECT_EQ(fm_sheet_set_thread_resolved(wb.handle, 0, kReply, 1), kThreadedCommentInvalid);

  const fm_mention edited_mention{kAlice, kMention, 5, 6};
  ASSERT_EQ(fm_sheet_edit_threaded_comment(wb.handle, 0, kReply, "Done @Alice", &edited_mention, 1), 0)
      << fm_last_error_message();

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  EXPECT_EQ(PersonNames(reloaded.handle), (std::vector<std::string>{"Alice", "Bob"}));
  const std::vector<CommentCopy> comments = ReadComments(reloaded.handle);
  ASSERT_EQ(comments.size(), 2U);

  const CommentCopy& opening = comments[0];
  EXPECT_EQ(opening.id, kThread);
  EXPECT_EQ(opening.row, 1U);
  EXPECT_EQ(opening.col, 1U);
  EXPECT_EQ(opening.person_id, kAlice);
  EXPECT_EQ(opening.created, kCreated);
  EXPECT_EQ(opening.text, "@Bob please check");
  EXPECT_EQ(opening.parent_id, "");
  EXPECT_EQ(opening.done, 1);
  ASSERT_EQ(opening.mentions.size(), 1U);
  EXPECT_EQ(opening.mentions[0].person_id, kBob);
  EXPECT_EQ(opening.mentions[0].mention_id, kMention);
  EXPECT_EQ(opening.mentions[0].start, 0U);
  EXPECT_EQ(opening.mentions[0].length, 4U);

  const CommentCopy& reply = comments[1];
  EXPECT_EQ(reply.id, kReply);
  EXPECT_EQ(reply.parent_id, kThread);
  EXPECT_EQ(reply.row, 1U);
  EXPECT_EQ(reply.col, 1U);
  EXPECT_EQ(reply.text, "Done @Alice");
  ASSERT_EQ(reply.mentions.size(), 1U);
  EXPECT_EQ(reply.mentions[0].person_id, kAlice);
  EXPECT_EQ(reply.mentions[0].start, 5U);

  // The legacy stub Excel stores for the thread is not listed as a note.
  std::uint32_t notes = 9;
  ASSERT_EQ(fm_sheet_get_comment_count(reloaded.handle, 0, &notes), 0);
  EXPECT_EQ(notes, 0U);
}

TEST(FormulonCApiThreadedComments, RemovingTheOpeningCommentRemovesTheThread) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  AddPeople(wb.handle);
  AddThread(wb.handle);
  EXPECT_EQ(fm_workbook_remove_person(wb.handle, kBob), kThreadedCommentInvalid);

  ASSERT_EQ(fm_sheet_remove_threaded_comment(wb.handle, 0, kReply), 0);
  EXPECT_EQ(ReadComments(wb.handle).size(), 1U);
  const fm_threaded_comment reply = Comment(kReply, 0, 0, kBob, "Again", kThread);
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &reply), 0) << fm_last_error_message();
  EXPECT_EQ(ReadComments(wb.handle).size(), 2U);
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, nullptr), kBindingNullPointer);
  ASSERT_EQ(fm_sheet_remove_threaded_comment(wb.handle, 0, kThread), 0);
  EXPECT_TRUE(ReadComments(wb.handle).empty());
  EXPECT_EQ(fm_sheet_remove_threaded_comment(wb.handle, 0, kThread), kThreadedCommentInvalid);
  ASSERT_EQ(fm_workbook_remove_person(wb.handle, kBob), 0);
}

TEST(FormulonCApiThreadedComments, CellCarriesNoteOrThreadNotBoth) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  AddPeople(wb.handle);
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 4, 2, "Alice", "a note"), 0);
  const fm_threaded_comment on_note = Comment(kThread, 4, 2, kAlice, "hi", nullptr);
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &on_note), kThreadedCommentInvalid);

  AddThread(wb.handle);
  EXPECT_EQ(fm_sheet_set_comment(wb.handle, 0, 1, 1, "Alice", "a note"), kThreadedCommentInvalid);
  const fm_threaded_comment second = Comment("{6F1E0D2C-1111-4111-8111-000000000003}", 1, 1, kAlice, "x", nullptr);
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &second), kThreadedCommentInvalid);
}

TEST(FormulonCApiThreadedComments, RejectsMalformedInput) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  AddPeople(wb.handle);
  fm_threaded_comment c = Comment(kThread, 0, 0, kAlice, "text", nullptr);
  c.created = "2026-01-02 03:04:05";
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c), kThreadedCommentInvalid);
  c.created = kCreated;
  c.person_id = "{0A1B2C3D-0000-4000-8000-0000000000FF}";
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c), kThreadedCommentInvalid);
  c.person_id = kAlice;
  c.mention_count = 1;
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c), kInvalidArgument);
  const fm_mention beyond{kBob, kMention, 2, 10};
  c.mentions = &beyond;
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c), kThreadedCommentInvalid);
  c.mentions = nullptr;
  c.mention_count = 0;
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 3, &c), kInvalidArgument);
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c), 0) << fm_last_error_message();
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c), kThreadedCommentInvalid);

  EXPECT_EQ(fm_sheet_edit_threaded_comment(wb.handle, 0, nullptr, "x", nullptr, 0), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_edit_threaded_comment(wb.handle, 0, kThread, "x", nullptr, 2), kInvalidArgument);
  EXPECT_EQ(fm_sheet_edit_threaded_comment(wb.handle, 0, kReply, "x", nullptr, 0), kThreadedCommentInvalid);
  EXPECT_EQ(fm_sheet_set_thread_resolved(wb.handle, 0, nullptr, 1), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_remove_threaded_comment(wb.handle, 0, nullptr), kBindingNullPointer);

  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_threaded_comment_count(wb.handle, 5, &count), kInvalidArgument);
  fm_threaded_comment out{};
  EXPECT_EQ(fm_sheet_threaded_comment_at(wb.handle, 0, 1, &out), kInvalidArgument);
}

}  // namespace
