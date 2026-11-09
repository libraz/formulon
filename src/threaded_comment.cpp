#include "threaded_comment.h"

#include <chrono>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "sheet.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "utils/utf8_length.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::string_view kStubPreamble =
    "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits to it "
    "will get removed if the file is opened in a newer version of Excel. Learn more: "
    "https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    ";

bool IsHex(char c) noexcept {
  const char upper = strings::ascii_to_upper(c);
  return strings::is_ascii_digit(c) || (upper >= 'A' && upper <= 'F');
}

// The uppercase spelling Excel writes; `id` is a well-formed GUID in either case.
std::string CanonicalGuid(std::string_view id) {
  std::string out(id);
  for (char& c : out) {
    c = strings::ascii_to_upper(c);
  }
  return out;
}

// Out of line so the string building is emitted once rather than at every validation site.
[[gnu::noinline]] Error Refuse(FormulonErrorCode code, std::string_view message, std::string_view value) {
  std::string detail = "value=";
  detail.append(value);
  return make_error(code, std::string(message), std::move(detail));
}

Error Invalid(std::string_view message, std::string_view value) {
  return Refuse(FormulonErrorCode::kInvalidArgument, message, value);
}

// The listed person whose id is `id` (ASCII case-insensitive), or null.
const Person* FindPerson(const std::vector<Person>& persons, std::string_view id) noexcept {
  for (const Person& p : persons) {
    if (strings::case_insensitive_eq(p.id, id)) {
      return &p;
    }
  }
  return nullptr;
}

// Rewrites `id` to the listed spelling of its person; `kNotFound` when nobody is listed under it.
Expected<void, Error> BindPerson(const std::vector<Person>& persons, std::string& id, std::string_view what) {
  const Person* person = FindPerson(persons, id);
  if (person == nullptr) {
    return Refuse(FormulonErrorCode::kNotFound, what, id);
  }
  id = person->id;
  return Expected<void, Error>::Ok();
}

// splitmix64 finaliser: a bijective 64-bit mix.
std::uint64_t Mix(std::uint64_t x) noexcept {
  x += 0x9E3779B97F4A7C15ULL;
  x = (x ^ (x >> 30U)) * 0xBF58476D1CE4E5B9ULL;
  x = (x ^ (x >> 27U)) * 0x94D049BB133111EBULL;
  return x ^ (x >> 31U);
}

std::uint64_t ClockSeed(const date_time::CivilTime& t) noexcept {
  std::uint64_t seed = static_cast<std::uint64_t>(static_cast<std::int64_t>(t.date.y));
  for (const unsigned field : {t.date.m, t.date.d, t.time.h, t.time.m, t.time.s}) {
    seed = Mix(seed ^ field);
  }
  return seed;
}

// `{XXXXXXXX-XXXX-4XXX-YXXX-XXXXXXXXXXXX}` (a version-4 layout) from two 64-bit words.
std::string FormatGuid(std::uint64_t high, std::uint64_t low) {
  high = (high & ~0xF000ULL) | 0x4000ULL;
  low = (low & ~(0xC000ULL << 48U)) | (0x8000ULL << 48U);
  constexpr char kHex[] = "0123456789ABCDEF";
  std::string out = "{";
  for (int nibble = 0; nibble < 32; ++nibble) {
    if (nibble == 8 || nibble == 12 || nibble == 16 || nibble == 20) {
      out.push_back('-');
    }
    const std::uint64_t word = nibble < 16 ? high : low;
    out.push_back(kHex[(word >> (60 - 4 * (nibble % 16))) & 0xFU]);
  }
  out.push_back('}');
  return out;
}

// `YYYY-MM-DDTHH:MM:SS.00`, the `dT` form.
std::string FormatStamp(const date_time::CivilTime& t) {
  const auto two = [](std::string& out, unsigned value) {
    out.push_back(static_cast<char>('0' + value / 10U % 10U));
    out.push_back(static_cast<char>('0' + value % 10U));
  };
  const unsigned year = static_cast<unsigned>(t.date.y);
  std::string out;
  two(out, year / 100U);
  two(out, year);
  out.push_back('-');
  two(out, t.date.m);
  out.push_back('-');
  two(out, t.date.d);
  out.push_back('T');
  two(out, t.time.h);
  out.push_back(':');
  two(out, t.time.m);
  out.push_back(':');
  two(out, t.time.s);
  out.append(".00");
  return out;
}

std::string BeginStubText(const ThreadedComment& root) {
  std::string out(kStubPreamble);
  out.append(root.text);
  return out;
}

void AppendStubReply(std::string& out, const ThreadedComment& reply) {
  out.append("\nReply:\n    ");
  out.append(reply.text);
}

Expected<void, Error> ValidateText(std::string_view text, std::vector<Mention>& mentions,
                                   const std::vector<Person>& persons) {
  const std::uint64_t units = utf16_units_in(text);
  for (Mention& m : mentions) {
    if (!is_threaded_comment_guid(m.mention_id)) {
      return Invalid("threaded comment: mention id is not a GUID", m.mention_id);
    }
    m.mention_id = CanonicalGuid(m.mention_id);
    RETURN_IF_ERROR(BindPerson(persons, m.person_id, "threaded comment: mention names an unknown person"));
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
    if (dash ? s[i] != '-' : !IsHex(s[i])) {
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
    if (kShape[i] == 'd' ? !strings::is_ascii_digit(s[i]) : s[i] != kShape[i]) {
      return false;
    }
  }
  return true;
}

std::size_t find_threaded_comment(const std::vector<ThreadedComment>& list, std::string_view id) noexcept {
  for (std::size_t i = 0; i < list.size(); ++i) {
    if (strings::case_insensitive_eq(list[i].id, id)) {
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
  std::vector<std::size_t> reply_indices;
  for (std::size_t i = 0; i < list.size(); ++i) {
    if (strings::case_insensitive_eq(list[i].parent_id, list[root].id)) {
      reply_indices.push_back(i);
    }
  }
  return threaded_comment_stub_text(list, root, reply_indices);
}

std::string threaded_comment_stub_text(const std::vector<ThreadedComment>& list, std::size_t root,
                                       const std::vector<std::size_t>& reply_indices) {
  std::string out = BeginStubText(list[root]);
  for (const std::size_t index : reply_indices) {
    AppendStubReply(out, list[index]);
  }
  return out;
}

Expected<void, Error> Workbook::add_person(Person person) {
  if (!is_threaded_comment_guid(person.id)) {
    return Invalid("id is not a GUID", person.id);
  }
  person.id = CanonicalGuid(person.id);
  if (FindPerson(persons_, person.id) != nullptr) {
    return Refuse(FormulonErrorCode::kAlreadyExists, "id already listed", person.id);
  }
  if (person.display_name.empty()) {
    return Invalid("empty display name", person.id);
  }
  persons_.push_back(std::move(person));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_person(std::string_view id) {
  const Person* person = FindPerson(persons_, id);
  if (person == nullptr) {
    return Refuse(FormulonErrorCode::kNotFound, "no such person", id);
  }
  const std::size_t index = static_cast<std::size_t>(person - persons_.data());
  for (std::size_t s = 0; s < sheet_count(); ++s) {
    for (const ThreadedComment& c : sheet(s).threaded_comments()) {
      bool referenced = strings::case_insensitive_eq(c.person_id, id);
      for (const Mention& m : c.mentions) {
        referenced = referenced || strings::case_insensitive_eq(m.person_id, id);
      }
      if (referenced) {
        return Refuse(FormulonErrorCode::kPreconditionFailed, "a threaded comment still names this person", id);
      }
    }
  }
  persons_.erase(persons_.begin() + static_cast<std::ptrdiff_t>(index));
  return Expected<void, Error>::Ok();
}

std::string Workbook::next_threaded_comment_guid() {
  if (!guid_seed_.has_value()) {
    // A pinned clock makes the sequence reproducible; otherwise the host clock and this workbook's address seed it.
    guid_seed_ = pinned_now_.has_value()
                     ? ClockSeed(*pinned_now_)
                     : Mix(static_cast<std::uint64_t>(std::chrono::system_clock::now().time_since_epoch().count()) ^
                           static_cast<std::uint64_t>(reinterpret_cast<std::uintptr_t>(this)));
  }
  const std::uint64_t high = Mix(*guid_seed_ ^ ++guid_sequence_);
  const std::uint64_t low = Mix(high ^ *guid_seed_);
  return FormatGuid(high, low);
}

Expected<std::string, Error> Workbook::add_threaded_comment(std::size_t sheet_index, ThreadedComment comment) {
  ASSIGN_OR_RETURN(Sheet * target, SheetAt(*this, sheet_index));
  std::vector<ThreadedComment>& list = target->mutable_threaded_comments();
  if (comment.id.empty()) {
    do {
      comment.id = next_threaded_comment_guid();
    } while (find_threaded_comment(list, comment.id) != list.size());
  }
  if (!is_threaded_comment_guid(comment.id)) {
    return Invalid("id is not a GUID", comment.id);
  }
  comment.id = CanonicalGuid(comment.id);
  if (find_threaded_comment(list, comment.id) != list.size()) {
    return Refuse(FormulonErrorCode::kAlreadyExists, "id already used on this sheet", comment.id);
  }
  RETURN_IF_ERROR(BindPerson(persons_, comment.person_id, "author is not in the person list"));
  if (comment.created.empty()) {
    comment.created = FormatStamp(pinned_now_.has_value() ? *pinned_now_ : date_time::host_utc_civil_time());
  } else if (comment.created.back() == 'Z') {
    comment.created.pop_back();
  }
  if (!is_threaded_comment_timestamp(comment.created)) {
    return Invalid("created is not YYYY-MM-DDTHH:MM:SS.ss", comment.created);
  }
  RETURN_IF_ERROR(ValidateText(comment.text, comment.mentions, persons_));

  std::size_t insert_at = list.size();
  if (comment.parent_id.empty()) {
    if (!Sheet::coord_in_grid(comment.row, comment.col)) {
      return make_error(FormulonErrorCode::kInvalidArgument, "coordinate out of grid",
                        "row=" + std::to_string(comment.row) + " col=" + std::to_string(comment.col));
    }
    if (find_thread_at(list, comment.row, comment.col) != list.size()) {
      return Refuse(FormulonErrorCode::kAlreadyExists, "cell already has a thread", comment.id);
    }
    for (const CellComment& note : target->comments()) {
      if (note.row == comment.row && note.col == comment.col) {
        return Refuse(FormulonErrorCode::kAlreadyExists, "cell already has a note", comment.id);
      }
    }
  } else {
    const std::size_t root = find_threaded_comment(list, comment.parent_id);
    if (root == list.size() || !list[root].parent_id.empty()) {
      return Invalid("parent is not a thread on this sheet", comment.parent_id);
    }
    comment.parent_id = list[root].id;
    comment.row = list[root].row;
    comment.col = list[root].col;
    comment.done = false;
    insert_at = root + 1U;
    for (std::size_t i = insert_at; i < list.size(); ++i) {
      if (strings::case_insensitive_eq(list[i].parent_id, comment.parent_id)) {
        insert_at = i + 1U;
      }
    }
  }
  std::string id = comment.id;
  list.insert(list.begin() + static_cast<std::ptrdiff_t>(insert_at), std::move(comment));
  return id;
}

Expected<void, Error> Workbook::edit_threaded_comment(std::size_t sheet_index, std::string_view id, std::string text,
                                                      std::vector<Mention> mentions) {
  ASSIGN_OR_RETURN(Sheet * target, SheetAt(*this, sheet_index));
  std::vector<ThreadedComment>& list = target->mutable_threaded_comments();
  const std::size_t index = find_threaded_comment(list, id);
  if (index == list.size()) {
    return Refuse(FormulonErrorCode::kNotFound, "no such comment", id);
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
  if (index == list.size()) {
    return Refuse(FormulonErrorCode::kNotFound, "no comment with this id", thread_id);
  }
  if (!list[index].parent_id.empty()) {
    return Invalid("the id names a reply, not a thread", thread_id);
  }
  list[index].done = done;
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_threaded_comment(std::size_t sheet_index, std::string_view id) {
  ASSIGN_OR_RETURN(Sheet * target, SheetAt(*this, sheet_index));
  std::vector<ThreadedComment>& list = target->mutable_threaded_comments();
  const std::size_t index = find_threaded_comment(list, id);
  if (index == list.size()) {
    return Refuse(FormulonErrorCode::kNotFound, "no such comment", id);
  }
  if (!list[index].parent_id.empty()) {
    list.erase(list.begin() + static_cast<std::ptrdiff_t>(index));
    return Expected<void, Error>::Ok();
  }
  const std::string root_id = list[index].id;
  std::vector<ThreadedComment> kept;
  kept.reserve(list.size());
  for (ThreadedComment& c : list) {
    if (!strings::case_insensitive_eq(c.id, root_id) && !strings::case_insensitive_eq(c.parent_id, root_id)) {
      kept.push_back(std::move(c));
    }
  }
  list = std::move(kept);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
