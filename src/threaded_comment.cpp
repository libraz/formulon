#include "threaded_comment.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/utf8_length.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::string_view kStubPreamble =
    "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it "
    "will get removed if the file is opened in a newer version of Excel. Learn more: "
    "https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ";

bool IsDigit(char c) noexcept {
  return c >= '0' && c <= '9';
}

bool IsUpperHex(char c) noexcept {
  return IsDigit(c) || (c >= 'A' && c <= 'F');
}

// Out of line so the string building is emitted once rather than at every validation site.
[[gnu::noinline]] Error Invalid(std::string_view message, std::string_view value) {
  std::string detail = "value=";
  detail.append(value);
  return make_error(FormulonErrorCode::kThreadedCommentInvalid, std::string(message), std::move(detail));
}

bool HasPerson(const std::vector<Person>& persons, std::string_view id) noexcept {
  for (const Person& p : persons) {
    if (p.id == id) {
      return true;
    }
  }
  return false;
}

Expected<void, Error> ValidateText(std::string_view text, const std::vector<Mention>& mentions,
                                   const std::vector<Person>& persons) {
  const std::uint64_t units = utf16_units_in(text);
  for (const Mention& m : mentions) {
    if (!is_threaded_comment_guid(m.mention_id)) {
      return Invalid("threaded comment: mention id is not a GUID", m.mention_id);
    }
    if (!HasPerson(persons, m.person_id)) {
      return Invalid("threaded comment: mention names an unknown person", m.person_id);
    }
    if (static_cast<std::uint64_t>(m.start) + m.length > units) {
      return Invalid("threaded comment: mention range lies outside the text", m.mention_id);
    }
  }
  return Expected<void, Error>::Ok();
}

Expected<Sheet*, Error> SheetAt(Workbook& wb, std::size_t sheet) {
  if (sheet >= wb.sheet_count()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "threaded comment: sheet index out of range",
                      "sheet=" + std::to_string(sheet));
  }
  return &wb.sheet(sheet);
}

}  // namespace

bool is_threaded_comment_guid(std::string_view s) noexcept {
  if (s.size() != 38U || s.front() != '{' || s.back() != '}') {
    return false;
  }
  for (std::size_t i = 1; i + 1 < s.size(); ++i) {
    const bool dash = i == 9U || i == 14U || i == 19U || i == 24U;
    if (dash ? s[i] != '-' : !IsUpperHex(s[i])) {
      return false;
    }
  }
  return true;
}

bool is_threaded_comment_timestamp(std::string_view s) noexcept {
  constexpr std::string_view kShape = "dddd-dd-ddTdd:dd:dd.dd";
  if (s.size() != kShape.size()) {
    return false;
  }
  for (std::size_t i = 0; i < s.size(); ++i) {
    if (kShape[i] == 'd' ? !IsDigit(s[i]) : s[i] != kShape[i]) {
      return false;
    }
  }
  return true;
}

std::size_t find_threaded_comment(const std::vector<ThreadedComment>& list, std::string_view id) noexcept {
  for (std::size_t i = 0; i < list.size(); ++i) {
    if (list[i].id == id) {
      return i;
    }
  }
  return list.size();
}

std::size_t find_thread_at(const std::vector<ThreadedComment>& list, std::uint32_t row, std::uint32_t col) noexcept {
  for (std::size_t i = 0; i < list.size(); ++i) {
    if (list[i].parent_id.empty() && list[i].row == row && list[i].col == col) {
      return i;
    }
  }
  return list.size();
}

std::string threaded_comment_stub_text(const std::vector<ThreadedComment>& list, std::size_t root) {
  std::string out(kStubPreamble);
  out.append(list[root].text);
  for (const ThreadedComment& c : list) {
    if (c.parent_id == list[root].id) {
      out.append("\nReply:\n    ");
      out.append(c.text);
    }
  }
  return out;
}

Expected<void, Error> Workbook::add_person(Person person) {
  if (!is_threaded_comment_guid(person.id)) {
    return Invalid("add_person: id is not a GUID", person.id);
  }
  if (HasPerson(persons_, person.id)) {
    return Invalid("add_person: id already listed", person.id);
  }
  if (person.display_name.empty()) {
    return Invalid("add_person: empty display name", person.id);
  }
  persons_.push_back(std::move(person));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_person(std::string_view id) {
  std::size_t index = persons_.size();
  for (std::size_t i = 0; i < persons_.size(); ++i) {
    if (persons_[i].id == id) {
      index = i;
      break;
    }
  }
  if (index == persons_.size()) {
    return Invalid("remove_person: no such person", id);
  }
  for (std::size_t s = 0; s < sheet_count(); ++s) {
    for (const ThreadedComment& c : sheet(s).threaded_comments()) {
      bool referenced = c.person_id == id;
      for (const Mention& m : c.mentions) {
        referenced = referenced || m.person_id == id;
      }
      if (referenced) {
        return Invalid("remove_person: a threaded comment still names this person", id);
      }
    }
  }
  persons_.erase(persons_.begin() + static_cast<std::ptrdiff_t>(index));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::add_threaded_comment(std::size_t sheet_index, ThreadedComment comment) {
  ASSIGN_OR_RETURN(Sheet * target, SheetAt(*this, sheet_index));
  std::vector<ThreadedComment>& list = target->mutable_threaded_comments();
  if (!is_threaded_comment_guid(comment.id)) {
    return Invalid("add_threaded_comment: id is not a GUID", comment.id);
  }
  if (find_threaded_comment(list, comment.id) != list.size()) {
    return Invalid("add_threaded_comment: id already used on this sheet", comment.id);
  }
  if (!HasPerson(persons_, comment.person_id)) {
    return Invalid("add_threaded_comment: author is not in the person list", comment.person_id);
  }
  if (!is_threaded_comment_timestamp(comment.created)) {
    return Invalid("add_threaded_comment: created is not YYYY-MM-DDTHH:MM:SS.ss", comment.created);
  }
  RETURN_IF_ERROR(ValidateText(comment.text, comment.mentions, persons_));

  std::size_t insert_at = list.size();
  if (comment.parent_id.empty()) {
    if (comment.row >= Sheet::kMaxRows || comment.col >= Sheet::kMaxCols) {
      return make_error(FormulonErrorCode::kInvalidArgument, "add_threaded_comment: coordinate out of grid",
                        "row=" + std::to_string(comment.row) + " col=" + std::to_string(comment.col));
    }
    if (find_thread_at(list, comment.row, comment.col) != list.size()) {
      return Invalid("add_threaded_comment: cell already has a thread", comment.id);
    }
    for (const CellComment& note : target->comments()) {
      if (note.row == comment.row && note.col == comment.col) {
        return Invalid("add_threaded_comment: cell already has a note", comment.id);
      }
    }
  } else {
    const std::size_t root = find_threaded_comment(list, comment.parent_id);
    if (root == list.size() || !list[root].parent_id.empty()) {
      return Invalid("add_threaded_comment: parent is not a thread on this sheet", comment.parent_id);
    }
    comment.row = list[root].row;
    comment.col = list[root].col;
    comment.done = false;
    insert_at = root + 1U;
    for (std::size_t i = insert_at; i < list.size(); ++i) {
      if (list[i].parent_id == comment.parent_id) {
        insert_at = i + 1U;
      }
    }
  }
  list.insert(list.begin() + static_cast<std::ptrdiff_t>(insert_at), std::move(comment));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::edit_threaded_comment(std::size_t sheet_index, std::string_view id, std::string text,
                                                      std::vector<Mention> mentions) {
  ASSIGN_OR_RETURN(Sheet * target, SheetAt(*this, sheet_index));
  std::vector<ThreadedComment>& list = target->mutable_threaded_comments();
  const std::size_t index = find_threaded_comment(list, id);
  if (index == list.size()) {
    return Invalid("edit_threaded_comment: no such comment", id);
  }
  RETURN_IF_ERROR(ValidateText(text, mentions, persons_));
  list[index].text = std::move(text);
  list[index].mentions = std::move(mentions);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_thread_resolved(std::size_t sheet_index, std::string_view thread_id, bool done) {
  ASSIGN_OR_RETURN(Sheet * target, SheetAt(*this, sheet_index));
  std::vector<ThreadedComment>& list = target->mutable_threaded_comments();
  const std::size_t index = find_threaded_comment(list, thread_id);
  if (index == list.size() || !list[index].parent_id.empty()) {
    return Invalid("set_thread_resolved: no thread with this id", thread_id);
  }
  list[index].done = done;
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_threaded_comment(std::size_t sheet_index, std::string_view id) {
  ASSIGN_OR_RETURN(Sheet * target, SheetAt(*this, sheet_index));
  std::vector<ThreadedComment>& list = target->mutable_threaded_comments();
  const std::size_t index = find_threaded_comment(list, id);
  if (index == list.size()) {
    return Invalid("remove_threaded_comment: no such comment", id);
  }
  if (!list[index].parent_id.empty()) {
    list.erase(list.begin() + static_cast<std::ptrdiff_t>(index));
    return Expected<void, Error>::Ok();
  }
  const std::string root_id = list[index].id;
  std::vector<ThreadedComment> kept;
  kept.reserve(list.size());
  for (ThreadedComment& c : list) {
    if (c.id != root_id && c.parent_id != root_id) {
      kept.push_back(std::move(c));
    }
  }
  list = std::move(kept);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
