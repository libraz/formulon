//
// Stable C ABI tests for threaded comments and the workbook person list:
// record shapes, CRUD, the note/thread exclusivity, and a save and reload.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_mention_t) == (kWasm32 ? 16U : 24U), "fm_mention_t ABI layout changed");
static_assert(sizeof(fm_threaded_comment_t) == (kWasm32 ? 44U : 80U), "fm_threaded_comment_t ABI layout changed");
static_assert(sizeof(fm_person_t) == (kWasm32 ? 20U : 40U), "fm_person_t ABI layout changed");
static_assert(sizeof(fm_threaded_comment_patch_t) == (kWasm32 ? 32U : 48U),
              "fm_threaded_comment_patch_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kAlreadyExists = static_cast<fm_status_t>(formulon::FormulonErrorCode::kAlreadyExists);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);

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
  ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  const fm_status_t rc = fm_workbook_load(bytes, len, &out->handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();
}

/// A patch that keeps everything; callers SET the fields they test.
fm_threaded_comment_patch_t KeepAll() {
  return FM_STRUCT_INIT(fm_threaded_comment_patch_t);
}

fm_threaded_comment_patch_t SetDone(int32_t done) {
  fm_threaded_comment_patch_t patch = KeepAll();
  patch.done_op = FM_PATCH_OP_SET;
  patch.done = done;
  return patch;
}

fm_threaded_comment_patch_t SetText(const char* text) {
  fm_threaded_comment_patch_t patch = KeepAll();
  patch.text_op = FM_PATCH_OP_SET;
  patch.text = text;
  return patch;
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
    std::uint32_t end = 0;
  };
  std::vector<MentionCopy> mentions;
};

std::vector<CommentCopy> ReadComments(const fm_workbook_t* wb) {
  std::vector<CommentCopy> out;
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_threaded_comment_count(wb, 0, &count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    fm_threaded_comment_t c = FM_STRUCT_INIT(fm_threaded_comment_t);
    EXPECT_EQ(fm_sheet_get_threaded_comment_at(wb, 0, i, &c), 0) << fm_get_last_error_message();
    CommentCopy copy;
    copy.id = Str(c.id);
    copy.person_id = Str(c.person_id);
    copy.created = Str(c.created);
    copy.text = Str(c.text);
    copy.parent_id = Str(c.parent_id);
    copy.row = c.row;
    copy.col = c.col;
    copy.done = c.done;
    for (std::uint32_t m = 0; m < c.mentions_count; ++m) {
      copy.mentions.push_back(
          {Str(c.mentions[m].person_id), Str(c.mentions[m].mention_id), c.mentions[m].start, c.mentions[m].end});
    }
    out.push_back(copy);
  }
  return out;
}

std::vector<std::string> PersonNames(const fm_workbook_t* wb) {
  std::vector<std::string> names;
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_get_person_count(wb, &count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    fm_person_t p = FM_STRUCT_INIT(fm_person_t);
    EXPECT_EQ(fm_workbook_get_person_at(wb, i, &p), 0);
    names.push_back(Str(p.display_name));
  }
  return names;
}

fm_person_t MakePerson(const char* id, const char* display_name, const char* user_id, const char* provider_id) {
  fm_person_t person = FM_STRUCT_INIT(fm_person_t);
  person.id = id;
  person.display_name = display_name;
  person.user_id = user_id;
  person.provider_id = provider_id;
  return person;
}

void AddPeople(fm_workbook_t* wb) {
  const fm_person_t alice = MakePerson(kAlice, "Alice", "alice@example.com", "None");
  const fm_person_t bob = MakePerson(kBob, "Bob", "bob@example.com", "None");
  const char* person_id = nullptr;
  ASSERT_EQ(fm_workbook_add_person(wb, &alice, &person_id), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_add_person(wb, &bob, &person_id), 0) << fm_get_last_error_message();
}

fm_threaded_comment_t Comment(const char* id, std::uint32_t row, std::uint32_t col, const char* person,
                              const char* text, const char* parent) {
  fm_threaded_comment_t c = FM_STRUCT_INIT(fm_threaded_comment_t);
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
  const fm_mention_t mention{kBob, kMention, 0, 4};
  fm_threaded_comment_t opening = Comment(kThread, 1, 1, kAlice, "@Bob please check", nullptr);
  opening.mentions = &mention;
  opening.mentions_count = 1;
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb, 0, &opening, ignored_id()), 0) << fm_get_last_error_message();
  // A reply takes the thread's anchor, whatever coordinates it carries.
  const fm_threaded_comment_t reply = Comment(kReply, 7, 7, kBob, "Done", kThread);
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb, 0, &reply, ignored_id()), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiThreadedComments, PersonsAddListAndRemove) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  EXPECT_EQ(PersonNames(wb.handle), (std::vector<std::string>{"Alice", "Bob"}));
  fm_person_t p = FM_STRUCT_INIT(fm_person_t);
  ASSERT_EQ(fm_workbook_get_person_at(wb.handle, 1, &p), 0);
  EXPECT_STREQ(p.id, kBob);
  EXPECT_STREQ(p.user_id, "bob@example.com");
  EXPECT_STREQ(p.provider_id, "None");
  EXPECT_EQ(fm_workbook_get_person_at(wb.handle, 2, &p), FM_STATUS_INDEX_OUT_OF_RANGE);

  const fm_person_t duplicate = MakePerson(kAlice, "Again", nullptr, nullptr);
  const char* added_id = nullptr;
  EXPECT_EQ(fm_workbook_add_person(wb.handle, &duplicate, &added_id), kAlreadyExists);
  EXPECT_EQ(added_id, nullptr);
  // A GUID is taken in either case, stored uppercase and matched across case.
  const fm_person_t lowercase = MakePerson("{0a1b2c3d-0000-4000-8000-00000000000c}", "Low", nullptr, nullptr);
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &lowercase, &added_id), 0) << fm_get_last_error_message();
  EXPECT_STREQ(added_id, "{0A1B2C3D-0000-4000-8000-00000000000C}");
  int32_t present = 0;
  ASSERT_EQ(fm_workbook_get_person(wb.handle, "{0A1B2C3D-0000-4000-8000-00000000000C}", &present, &p), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(p.id, "{0A1B2C3D-0000-4000-8000-00000000000C}");
  EXPECT_EQ(fm_workbook_get_person(wb.handle, "{0a1b2c3d-0000-4000-8000-00000000000c}", &present, &p), 0);
  EXPECT_EQ(present, 1);
  ASSERT_EQ(fm_workbook_remove_person(wb.handle, "{0a1b2c3d-0000-4000-8000-00000000000C}"), 0);
  const fm_person_t malformed = MakePerson("{0A1B2C3D-0000-4000-8000-00000000000Z}", "Bad", nullptr, nullptr);
  EXPECT_EQ(fm_workbook_add_person(wb.handle, &malformed, &added_id), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_workbook_add_person(wb.handle, nullptr, &added_id), kBindingNullPointer);

  ASSERT_EQ(fm_workbook_remove_person(wb.handle, kAlice), 0);
  EXPECT_EQ(PersonNames(wb.handle), (std::vector<std::string>{"Bob"}));
  EXPECT_EQ(fm_workbook_remove_person(wb.handle, kAlice), FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiThreadedComments, ThreadRepliesAndMentionsRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  AddThread(wb.handle);
  const fm_threaded_comment_patch_t resolve = SetDone(1);
  ASSERT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &resolve), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kReply, &resolve), kInvalidArgument);
  // A comment id matches across letter case.
  fm_threaded_comment_t found = FM_STRUCT_INIT(fm_threaded_comment_t);
  int32_t found_present = 0;
  ASSERT_EQ(
      fm_sheet_get_threaded_comment(wb.handle, 0, "{6f1e0d2c-1111-4111-8111-000000000001}", &found_present, &found), 0);
  EXPECT_EQ(found_present, 1);
  EXPECT_STREQ(found.id, kThread);

  const fm_mention_t edited_mention{kAlice, kMention, 5, 11};
  fm_threaded_comment_patch_t edit = SetText("Done @Alice");
  edit.mentions_op = FM_PATCH_OP_SET;
  edit.mentions = &edited_mention;
  edit.mentions_count = 1;
  ASSERT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kReply, &edit), 0) << fm_get_last_error_message();

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
  EXPECT_EQ(opening.parent_id, "");  // read NULL
  EXPECT_EQ(opening.done, 1);
  ASSERT_EQ(opening.mentions.size(), 1U);
  EXPECT_EQ(opening.mentions[0].person_id, kBob);
  EXPECT_EQ(opening.mentions[0].mention_id, kMention);
  EXPECT_EQ(opening.mentions[0].start, 0U);
  EXPECT_EQ(opening.mentions[0].end, 4U);

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
  std::size_t notes = 9;
  ASSERT_EQ(fm_sheet_get_comment_count(reloaded.handle, 0, &notes), 0);
  EXPECT_EQ(notes, 0U);
}

TEST(FormulonCApiThreadedComments, RemovingTheOpeningCommentRemovesTheThread) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  AddThread(wb.handle);
  EXPECT_EQ(fm_workbook_remove_person(wb.handle, kBob), kPreconditionFailed);

  ASSERT_EQ(fm_sheet_remove_threaded_comment(wb.handle, 0, kReply), 0);
  EXPECT_EQ(ReadComments(wb.handle).size(), 1U);
  const fm_threaded_comment_t reply = Comment(kReply, 0, 0, kBob, "Again", kThread);
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &reply, ignored_id()), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadComments(wb.handle).size(), 2U);
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, nullptr, ignored_id()), kBindingNullPointer);
  ASSERT_EQ(fm_sheet_remove_threaded_comment(wb.handle, 0, kThread), 0);
  EXPECT_TRUE(ReadComments(wb.handle).empty());
  EXPECT_EQ(fm_sheet_remove_threaded_comment(wb.handle, 0, kThread), FM_STATUS_NOT_FOUND);
  ASSERT_EQ(fm_workbook_remove_person(wb.handle, kBob), 0);
}

TEST(FormulonCApiThreadedComments, CellCarriesNoteOrThreadNotBoth) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  const fm_comment_t note{sizeof(fm_comment_t), 0U, 0U, "Alice", "a note"};
  ASSERT_EQ(fm_sheet_set_comment(wb.handle, 0, 4, 2, &note), 0);
  const fm_threaded_comment_t on_note = Comment(kThread, 4, 2, kAlice, "hi", nullptr);
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &on_note, ignored_id()), kAlreadyExists);

  AddThread(wb.handle);
  EXPECT_EQ(fm_sheet_set_comment(wb.handle, 0, 1, 1, &note), kAlreadyExists);
  const fm_threaded_comment_t second = Comment("{6F1E0D2C-1111-4111-8111-000000000003}", 1, 1, kAlice, "x", nullptr);
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &second, ignored_id()), kAlreadyExists);
}

TEST(FormulonCApiThreadedComments, RejectsMalformedInput) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  fm_threaded_comment_t c = Comment(kThread, 0, 0, kAlice, "text", nullptr);
  c.created = "2026-01-02 03:04:05";
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), kInvalidArgument);
  c.created = kCreated;
  c.person_id = "{0A1B2C3D-0000-4000-8000-0000000000FF}";
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), FM_STATUS_NOT_FOUND);
  c.person_id = kAlice;
  c.mentions_count = 1;
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), kInvalidArgument);
  const fm_mention_t beyond{kBob, kMention, 2, 10};
  c.mentions = &beyond;
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), kInvalidArgument);
  c.mentions = nullptr;
  c.mentions_count = 0;
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 3, &c, ignored_id()), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), kAlreadyExists);

  fm_threaded_comment_patch_t patch = SetText("x");
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, nullptr, &patch), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, nullptr), kBindingNullPointer);
  patch.mentions_op = FM_PATCH_OP_SET;
  patch.mentions_count = 2;
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &patch), kInvalidArgument);
  patch.mentions_op = FM_PATCH_OP_KEEP;
  patch.mentions_count = 0;
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kReply, &patch), FM_STATUS_NOT_FOUND);
  patch.done_op = FM_PATCH_OP_SET;
  patch.done = 2;
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &patch), kInvalidArgument);
  EXPECT_EQ(fm_sheet_remove_threaded_comment(wb.handle, 0, nullptr), kBindingNullPointer);

  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_threaded_comment_count(wb.handle, 5, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  fm_threaded_comment_t out = FM_STRUCT_INIT(fm_threaded_comment_t);
  EXPECT_EQ(fm_sheet_get_threaded_comment_at(wb.handle, 0, 1, &out), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiThreadedComments, UpdatePatchKeepsWhatItDoesNotEngage) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  AddThread(wb.handle);

  // `done` alone keeps the text and mentions.
  const fm_threaded_comment_patch_t resolve = SetDone(1);
  ASSERT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &resolve), 0) << fm_get_last_error_message();
  std::vector<CommentCopy> comments = ReadComments(wb.handle);
  ASSERT_EQ(comments.size(), 2U);
  EXPECT_EQ(comments[0].done, 1);
  EXPECT_EQ(comments[0].text, "@Bob please check");
  ASSERT_EQ(comments[0].mentions.size(), 1U);

  // Text alone keeps the mentions, which must still fit the new text.
  const fm_threaded_comment_patch_t retext = SetText("@Bob look again");
  ASSERT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &retext), 0) << fm_get_last_error_message();
  comments = ReadComments(wb.handle);
  EXPECT_EQ(comments[0].text, "@Bob look again");
  ASSERT_EQ(comments[0].mentions.size(), 1U);
  EXPECT_EQ(comments[0].mentions[0].person_id, kBob);
  EXPECT_EQ(comments[0].done, 1);
  const fm_threaded_comment_patch_t too_short = SetText("@B");
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &too_short), kInvalidArgument);

  // SET mentions alone replace the list and keep the text; an empty list clears it.
  fm_threaded_comment_patch_t unmention = KeepAll();
  unmention.mentions_op = FM_PATCH_OP_SET;
  ASSERT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &unmention), 0) << fm_get_last_error_message();
  comments = ReadComments(wb.handle);
  EXPECT_EQ(comments[0].text, "@Bob look again");
  EXPECT_TRUE(comments[0].mentions.empty());

  // `done` on a reply fails as a whole: the text in the same patch does not land.
  fm_threaded_comment_patch_t reply_done = SetText("changed");
  reply_done.done_op = FM_PATCH_OP_SET;
  reply_done.done = 0;
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kReply, &reply_done), kInvalidArgument);
  comments = ReadComments(wb.handle);
  EXPECT_EQ(comments[1].text, "Done");

  const fm_threaded_comment_patch_t reopen = SetDone(0);
  ASSERT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &reopen), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadComments(wb.handle)[0].done, 0);

  // An unknown id fails even when the patch sets nothing.
  const fm_threaded_comment_patch_t nothing = KeepAll();
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, "{6F1E0D2C-1111-4111-8111-0000000000FF}", &nothing),
            FM_STATUS_NOT_FOUND);
}

// No threaded-comment field can be cleared, an op outside 0..2 is refused,
// and SET with a NULL text is refused; each refusal leaves the comment as it was.
TEST(FormulonCApiThreadedComments, UpdateRejectsClearUnknownOpsAndNullSetText) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  AddThread(wb.handle);
  for (const int32_t raw_op : {static_cast<int32_t>(FM_PATCH_OP_CLEAR), 3}) {
    const auto op = static_cast<fm_patch_op_t>(raw_op);
    fm_threaded_comment_patch_t text = SetText("changed");
    text.text_op = op;
    EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &text), kInvalidArgument) << raw_op;
    fm_threaded_comment_patch_t mentions = SetText("changed");
    mentions.mentions_op = op;
    EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &mentions), kInvalidArgument) << raw_op;
    fm_threaded_comment_patch_t done = SetText("changed");
    done.done_op = op;
    done.done = 1;
    EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &done), kInvalidArgument) << raw_op;
  }
  const fm_threaded_comment_patch_t null_text = SetText(nullptr);
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &null_text), kInvalidArgument);

  const std::vector<CommentCopy> comments = ReadComments(wb.handle);
  ASSERT_EQ(comments.size(), 2U);
  EXPECT_EQ(comments[0].text, "@Bob please check");
  EXPECT_EQ(comments[0].done, 0);
  EXPECT_EQ(comments[0].mentions.size(), 1U);

  // An all-KEEP patch on a known id is a successful no-op.
  const fm_threaded_comment_patch_t keep = KeepAll();
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &keep), 0);
  EXPECT_EQ(ReadComments(wb.handle)[0].text, "@Bob please check");
}

// Optional strings read NULL when absent: an opening comment's parent id and a
// person without user or provider id. A comment added without a time is
// stamped, so its created string is present.
TEST(FormulonCApiThreadedComments, AbsentOptionalStringsReadNull) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_person_t bare = MakePerson(kAlice, "Alice", nullptr, "");
  const char* added_id = nullptr;
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &bare, &added_id), 0) << fm_get_last_error_message();
  fm_person_t p = FM_STRUCT_INIT(fm_person_t);
  ASSERT_EQ(fm_workbook_get_person_at(wb.handle, 0, &p), 0);
  EXPECT_STREQ(p.display_name, "Alice");
  EXPECT_EQ(p.user_id, nullptr);
  EXPECT_EQ(p.provider_id, nullptr);

  fm_threaded_comment_t c = Comment(kThread, 0, 0, kAlice, "text", nullptr);
  c.created = nullptr;
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), 0) << fm_get_last_error_message();
  fm_threaded_comment_t out = FM_STRUCT_INIT(fm_threaded_comment_t);
  ASSERT_EQ(fm_sheet_get_threaded_comment_at(wb.handle, 0, 0, &out), 0);
  EXPECT_STREQ(out.text, "text");
  EXPECT_EQ(out.parent_id, nullptr);
  EXPECT_NE(out.created, nullptr);
}

}  // namespace

TEST(FormulonCApiThreadedComments, AddReturnsTheGivenOrGeneratedId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  const char* id = nullptr;
  const fm_threaded_comment_t given = Comment(kThread, 1, 1, kAlice, "hello", nullptr);
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &given, &id), 0) << fm_get_last_error_message();
  EXPECT_STREQ(id, kThread);

  const fm_threaded_comment_t generated = Comment(nullptr, 2, 2, kAlice, "again", nullptr);
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &generated, &id), 0) << fm_get_last_error_message();
  ASSERT_NE(id, nullptr);
  const std::string generated_id = id;
  EXPECT_EQ(generated_id.size(), 38U);
  EXPECT_EQ(generated_id.front(), '{');
  EXPECT_EQ(ReadComments(wb.handle).at(1).id, generated_id);

  fm_threaded_comment_t read = FM_STRUCT_INIT(fm_threaded_comment_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_threaded_comment(wb.handle, 0, generated_id.c_str(), &present, &read), 0);
  EXPECT_EQ(present, 1);
}

TEST(FormulonCApiThreadedComments, AddWithoutCreatedStampsThePinnedClock) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  const fm_civil_time_t pinned{2026, 3, 4, 5, 6, 7};
  ASSERT_EQ(fm_workbook_set_pinned_now(wb.handle, &pinned), 0);
  fm_threaded_comment_t c = Comment(kThread, 1, 1, kAlice, "hello", nullptr);
  c.created = nullptr;
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadComments(wb.handle).at(0).created.rfind("2026-03-04T05:06:07", 0), 0U);
}

TEST(FormulonCApiThreadedComments, GeneratedIdsAreReproducibleWithAPinnedClock) {
  std::string ids[2];
  for (std::string& out : ids) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    AddPeople(wb.handle);
    const fm_civil_time_t pinned{2026, 3, 4, 5, 6, 7};
    ASSERT_EQ(fm_workbook_set_pinned_now(wb.handle, &pinned), 0);
    const fm_threaded_comment_t c = Comment(nullptr, 1, 1, kAlice, "hello", nullptr);
    const char* id = nullptr;
    ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, &id), 0) << fm_get_last_error_message();
    out = id != nullptr ? id : "";
  }
  EXPECT_FALSE(ids[0].empty());
  EXPECT_EQ(ids[0], ids[1]);
}

TEST(FormulonCApiThreadedComments, UnknownParentAndUnknownPersonAreNotFound) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  AddThread(wb.handle);
  const fm_threaded_comment_t orphan =
      Comment(nullptr, 0, 0, kAlice, "reply", "{6F1E0D2C-1111-4111-8111-0000000000EE}");
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &orphan, ignored_id()), FM_STATUS_NOT_FOUND);
  const fm_threaded_comment_t stranger =
      Comment(nullptr, 5, 5, "{0A1B2C3D-0000-4000-8000-0000000000FF}", "hi", nullptr);
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &stranger, ignored_id()), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(ReadComments(wb.handle).size(), 2U);
}

TEST(FormulonCApiGeneratedIds, PersonNullIdGenerated) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_person_t person = MakePerson(nullptr, "Alice", nullptr, nullptr);
  const char* id = nullptr;
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &person, &id), 0) << fm_get_last_error_message();
  ASSERT_NE(id, nullptr);
  const std::string generated = id;
  ASSERT_EQ(generated.size(), 38U);
  EXPECT_EQ(generated.front(), '{');
  EXPECT_EQ(generated.back(), '}');
  EXPECT_EQ(generated.find_first_of("abcdef"), std::string::npos);
  fm_person_t read = FM_STRUCT_INIT(fm_person_t);
  int32_t present = 0;
  ASSERT_EQ(fm_workbook_get_person(wb.handle, generated.c_str(), &present, &read), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(read.display_name, "Alice");
}

TEST(FormulonCApiGeneratedIds, PersonEmptyIdGenerated) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_civil_time_t pinned{2026, 3, 4, 5, 6, 7};
  ASSERT_EQ(fm_workbook_set_pinned_now(wb.handle, &pinned), 0);
  const fm_person_t first = MakePerson("", "Alice", nullptr, nullptr);
  const fm_person_t second = MakePerson("", "Bob", nullptr, nullptr);
  const char* id = nullptr;
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &first, &id), 0) << fm_get_last_error_message();
  const std::string first_id = id != nullptr ? id : "";
  ASSERT_EQ(fm_workbook_add_person(wb.handle, &second, &id), 0) << fm_get_last_error_message();
  const std::string second_id = id != nullptr ? id : "";
  EXPECT_EQ(first_id.size(), 38U);
  EXPECT_EQ(second_id.size(), 38U);
  EXPECT_NE(first_id, second_id);
  EXPECT_EQ(PersonNames(wb.handle), (std::vector<std::string>{"Alice", "Bob"}));
}

TEST(FormulonCApiGeneratedIds, MentionNullIdGenerated) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  const fm_mention_t mention{kBob, nullptr, 0, 4};
  fm_threaded_comment_t c = Comment(kThread, 1, 1, kAlice, "@Bob hi", nullptr);
  c.mentions = &mention;
  c.mentions_count = 1;
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), 0) << fm_get_last_error_message();
  const std::vector<CommentCopy> comments = ReadComments(wb.handle);
  ASSERT_EQ(comments.size(), 1U);
  ASSERT_EQ(comments[0].mentions.size(), 1U);
  EXPECT_EQ(comments[0].mentions[0].mention_id.size(), 38U);
  EXPECT_EQ(comments[0].mentions[0].mention_id.front(), '{');
}

TEST(FormulonCApiMentionSpan, StartEndRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  const std::string text = "Hello @Bob, see this";
  const fm_mention_t mention{kBob, kMention, 6, 10};
  fm_threaded_comment_t c = Comment(kThread, 1, 1, kAlice, text.c_str(), nullptr);
  c.mentions = &mention;
  c.mentions_count = 1;
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), 0) << fm_get_last_error_message();

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  const std::vector<CommentCopy> comments = ReadComments(reloaded.handle);
  ASSERT_EQ(comments.size(), 1U);
  ASSERT_EQ(comments[0].mentions.size(), 1U);
  const CommentCopy::MentionCopy& read = comments[0].mentions[0];
  EXPECT_EQ(read.start, 6U);
  EXPECT_EQ(read.end, 10U);
  EXPECT_EQ(comments[0].text.substr(read.start, read.end - read.start), "@Bob");
}

TEST(FormulonCApiMentionSpan, EndBeforeStartInvalid) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  const fm_mention_t inverted{kBob, kMention, 4, 0};
  fm_threaded_comment_t c = Comment(kThread, 1, 1, kAlice, "@Bob hi", nullptr);
  c.mentions = &inverted;
  c.mentions_count = 1;
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), kInvalidArgument);
  EXPECT_TRUE(ReadComments(wb.handle).empty());

  c.mentions = nullptr;
  c.mentions_count = 0;
  ASSERT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &c, ignored_id()), 0) << fm_get_last_error_message();
  fm_threaded_comment_patch_t patch = KeepAll();
  patch.mentions_op = FM_PATCH_OP_SET;
  patch.mentions = &inverted;
  patch.mentions_count = 1;
  EXPECT_EQ(fm_sheet_update_threaded_comment(wb.handle, 0, kThread, &patch), kInvalidArgument);
}

TEST(FormulonCApiThreadParentEmpty, AddWithEmptyParentInvalid) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddPeople(wb.handle);
  const fm_threaded_comment_t empty_parent = Comment(kThread, 1, 1, kAlice, "hello", "");
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &empty_parent, ignored_id()), kInvalidArgument);
  EXPECT_TRUE(ReadComments(wb.handle).empty());
  const fm_threaded_comment_t opening = Comment(kThread, 1, 1, kAlice, "hello", nullptr);
  EXPECT_EQ(fm_sheet_add_threaded_comment(wb.handle, 0, &opening, ignored_id()), 0) << fm_get_last_error_message();
}
