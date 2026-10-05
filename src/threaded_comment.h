//
// Threaded comments: threads, replies, mentions and the person list of a workbook.
//
// A thread is the top-level `ThreadedComment` of a cell (empty `parent_id`);
// its replies name the top-level id in `parent_id` and share its anchor.
// A sheet stores its comments as one flat list in file order, each thread
// followed by its replies. Ids and timestamps come from the caller: the
// engine never mints a GUID or reads the clock, so a save is deterministic.
//
// Excel also stores every thread as a legacy note ("stub") whose author is
// `tc={thread id}` so older readers can show it. The stub is not modelled:
// the reader drops it and the writer regenerates it from the thread.

#ifndef FORMULON_THREADED_COMMENT_H_
#define FORMULON_THREADED_COMMENT_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

namespace formulon {

/// One `@person` mention inside a threaded comment's text.
struct Mention {
  std::string person_id;     ///< `Person::id` of the mentioned person.
  std::string mention_id;    ///< Brace-wrapped uppercase GUID of the mention itself.
  std::uint32_t start = 0;   ///< Offset into the text, in UTF-16 code units.
  std::uint32_t length = 0;  ///< Length of the mention, in UTF-16 code units.
};

/// One threaded comment: a thread's opening comment or one of its replies.
struct ThreadedComment {
  std::string id;  ///< Brace-wrapped uppercase GUID.
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string person_id;  ///< `Person::id` of the author.
  std::string created;    ///< `YYYY-MM-DDTHH:MM:SS.ss` (Excel's `dT`).
  std::string text;
  std::string parent_id;  ///< Empty for a thread's opening comment.
  bool done = false;      ///< Resolved state; meaningful on the opening comment only.
  std::vector<Mention> mentions;
};

/// One entry of the workbook's person list (`xl/persons/person.xml`).
struct Person {
  std::string id;  ///< Brace-wrapped uppercase GUID.
  std::string display_name;
  std::string user_id;
  std::string provider_id;  ///< `AD`, `Windows Live`, `PeoplePicker` or `None`.
};

/// True for a brace-wrapped uppercase GUID, `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`.
bool is_threaded_comment_guid(std::string_view s) noexcept;

/// True for Excel's threaded-comment timestamp form `YYYY-MM-DDTHH:MM:SS.ss`.
bool is_threaded_comment_timestamp(std::string_view s) noexcept;

/// Index of the comment with `id` in `list`, or `list.size()` when absent.
std::size_t find_threaded_comment(const std::vector<ThreadedComment>& list, std::string_view id) noexcept;

/// Index of the thread (opening comment) anchored at `row`/`col`, or
/// `list.size()` when the cell has none.
std::size_t find_thread_at(const std::vector<ThreadedComment>& list, std::uint32_t row, std::uint32_t col) noexcept;

/// Text of the legacy note Excel stores for the thread opened at
/// `list[root]`: a fixed preamble followed by the opening comment and every
/// reply in list order.
std::string threaded_comment_stub_text(const std::vector<ThreadedComment>& list, std::size_t root);

}  // namespace formulon

#endif  // FORMULON_THREADED_COMMENT_H_
