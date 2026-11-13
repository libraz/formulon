#include "io/threaded_comments_io.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/cell_parser.h"
#include "io/ooxml_defs.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "threaded_comment.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon::io {
namespace {

constexpr std::string_view kThreadedNs = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments";
constexpr std::string_view kStubAuthorPrefix = "tc=";

void AppendRootOpen(std::string& out, std::string_view root) {
  out.append(kXmlDecl);
  out.push_back('<');
  out.append(root);
  append_xml_attr(out, "xmlns", kThreadedNs);
  append_xml_attr(out, "xmlns:x", kNsSpreadsheetMl);
  out.append(">\n");
}

std::uint64_t CellCoordinateKey(std::uint32_t row, std::uint32_t col) noexcept {
  return (static_cast<std::uint64_t>(row) << 32U) | col;
}

struct LegacyStubIndexes {
  std::unordered_set<std::uint64_t> root_coordinates;
  std::unordered_map<std::string_view, std::vector<std::size_t>> replies_by_parent_id;

  explicit LegacyStubIndexes(const std::vector<ThreadedComment>& threads) {
    root_coordinates.reserve(threads.size());
    replies_by_parent_id.reserve(threads.size());
    for (std::size_t i = 0; i < threads.size(); ++i) {
      const ThreadedComment& thread = threads[i];
      if (thread.parent_id.empty()) {
        root_coordinates.insert(CellCoordinateKey(thread.row, thread.col));
      }
      // Empty parent ids are indexed too, matching parent ids verbatim like the per-root scan.
      replies_by_parent_id[std::string_view(thread.parent_id)].push_back(i);
    }
    for (auto& entry : replies_by_parent_id) {
      sort_replies_for_save(threads, entry.second);
    }
  }
};

// The order Excel writes a sheet's threaded comments in: each thread opener in list order followed by its replies
// in save order, then any reply whose opener is missing, in list order.
std::vector<std::size_t> threads_in_save_order(const std::vector<ThreadedComment>& threads) {
  const LegacyStubIndexes indexes(threads);
  std::vector<bool> placed(threads.size(), false);
  std::vector<std::size_t> order;
  order.reserve(threads.size());
  for (std::size_t i = 0; i < threads.size(); ++i) {
    if (!threads[i].parent_id.empty()) {
      continue;
    }
    order.push_back(i);
    placed[i] = true;
    const auto replies = indexes.replies_by_parent_id.find(std::string_view(threads[i].id));
    if (replies == indexes.replies_by_parent_id.end()) {
      continue;
    }
    for (const std::size_t reply : replies->second) {
      if (!placed[reply]) {
        order.push_back(reply);
        placed[reply] = true;
      }
    }
  }
  for (std::size_t i = 0; i < threads.size(); ++i) {
    if (!placed[i]) {
      order.push_back(i);
    }
  }
  return order;
}

// Mention offsets in the part count UTF-16 units over the text as stored, where every line break is CRLF; the model
// text keeps a bare LF. Both walks advance by UTF-16 units; a LF not preceded by CR is one unit in the model and two
// in the file.
std::uint32_t Utf16Width(unsigned char lead) {
  return lead >= 0xF0U ? 2U : 1U;
}

bool IsBareLf(std::string_view text, std::size_t i) {
  return text[i] == '\n' && (i == 0 || text[i - 1U] != '\r');
}

std::uint32_t FileOffset(std::string_view text, std::uint32_t offset) {
  std::uint32_t model = 0;
  std::uint32_t file = 0;
  for (std::size_t i = 0; i < text.size(); ++i) {
    const auto byte = static_cast<unsigned char>(text[i]);
    if ((byte & 0xC0U) == 0x80U) {
      continue;
    }
    const std::uint32_t width = Utf16Width(byte);
    if (offset < model + width) {
      return file + (offset - model);
    }
    model += width;
    file += IsBareLf(text, i) ? 2U : width;
  }
  return file + (offset - model);
}

// The inverse; an offset between the CR and LF of a stored break lands before the break.
std::uint32_t ModelOffset(std::string_view text, std::uint32_t offset) {
  std::uint32_t model = 0;
  std::uint32_t file = 0;
  for (std::size_t i = 0; i < text.size(); ++i) {
    const auto byte = static_cast<unsigned char>(text[i]);
    if ((byte & 0xC0U) == 0x80U) {
      continue;
    }
    const std::uint32_t width = Utf16Width(byte);
    const std::uint32_t stored = IsBareLf(text, i) ? 2U : width;
    if (offset < file + stored) {
      return model + (stored == width ? offset - file : 0U);
    }
    model += width;
    file += stored;
  }
  return model + (offset - file);
}

}  // namespace

Expected<std::vector<ThreadedComment>, Error> read_threaded_comments(const std::vector<std::uint8_t>& bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, bytes, "threaded_comments_reader", "threadedComments part"));
  const pugi::xml_node root = doc.child("ThreadedComments");
  if (!root) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "threadedComments part: missing <ThreadedComments> root",
                      "context=threaded_comments_reader");
  }
  std::vector<ThreadedComment> out;
  for (pugi::xml_node node = root.child("threadedComment"); node; node = node.next_sibling("threadedComment")) {
    auto rc = parse_a1(node.attribute("ref").value());
    if (!rc) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "threadedComment: ref missing or unparseable",
                        "context=threaded_comments_reader ref=" + std::string(node.attribute("ref").value()));
    }
    ThreadedComment c;
    c.row = rc.value().first;
    c.col = rc.value().second;
    c.id = node.attribute("id").value();
    c.person_id = node.attribute("personId").value();
    c.created = node.attribute("dT").value();
    if (!c.created.empty() && c.created.back() == 'Z') {
      c.created.pop_back();
    }
    if (!is_threaded_comment_timestamp(c.created)) {
      c.created.clear();
    }
    c.parent_id = node.attribute("parentId").value();
    c.done = attr_bool(node, "done");
    AppendOoxmlTextUnescaped(c.text, node.child("text").text().get());
    for (pugi::xml_node m = node.child("mentions").child("mention"); m; m = m.next_sibling("mention")) {
      Mention mention;
      mention.person_id = m.attribute("mentionpersonId").value();
      mention.mention_id = m.attribute("mentionId").value();
      const std::uint32_t start = attr_u32(m, "startIndex");
      const std::uint32_t length = attr_u32(m, "length");
      mention.start = ModelOffset(c.text, start);
      mention.length = length > UINT32_MAX - start ? length : ModelOffset(c.text, start + length) - mention.start;
      c.mentions.push_back(std::move(mention));
    }
    out.push_back(std::move(c));
  }
  return out;
}

Expected<std::vector<Person>, Error> read_persons(const std::vector<std::uint8_t>& bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, bytes, "persons_reader", "persons part"));
  const pugi::xml_node root = doc.child("personList");
  if (!root) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "persons part: missing <personList> root",
                      "context=persons_reader");
  }
  std::vector<Person> out;
  for (pugi::xml_node node = root.child("person"); node; node = node.next_sibling("person")) {
    Person p;
    p.id = node.attribute("id").value();
    p.display_name = node.attribute("displayName").value();
    p.user_id = node.attribute("userId").value();
    p.provider_id = node.attribute("providerId").value();
    out.push_back(std::move(p));
  }
  return out;
}

std::string write_threaded_comments(const std::vector<ThreadedComment>& comments) {
  if (comments.empty()) {
    return {};
  }
  std::string out;
  out.reserve(256 + comments.size() * 256);
  AppendRootOpen(out, "ThreadedComments");
  for (const std::size_t index : threads_in_save_order(comments)) {
    const ThreadedComment& c = comments[index];
    out.append("  <threadedComment");
    append_xml_attr(out, "ref", a1::encode_a1(c.row, c.col));
    if (!c.created.empty()) {
      append_xml_attr(out, "dT", c.created);
    }
    append_xml_attr(out, "personId", c.person_id);
    append_xml_attr(out, "id", c.id);
    if (!c.parent_id.empty()) {
      append_xml_attr(out, "parentId", c.parent_id);
    } else if (c.done) {
      out.append(" done=\"1\"");
    }
    out.append("><text>");
    AppendXmlEscapedCrlf(out, c.text);
    out.append("</text>");
    if (!c.mentions.empty()) {
      out.append("<mentions>");
      for (const Mention& m : c.mentions) {
        out.append("<mention");
        append_xml_attr(out, "mentionpersonId", m.person_id);
        append_xml_attr(out, "mentionId", m.mention_id);
        const std::uint32_t start = FileOffset(c.text, m.start);
        append_xml_attr_uint(out, "startIndex", start);
        append_xml_attr_uint(out, "length", FileOffset(c.text, m.start + m.length) - start);
        out.append("/>");
      }
      out.append("</mentions>");
    }
    out.append("</threadedComment>\n");
  }
  out.append("</ThreadedComments>\n");
  return out;
}

std::string write_persons(const std::vector<Person>& persons) {
  if (persons.empty()) {
    return {};
  }
  std::string out;
  out.reserve(256 + persons.size() * 192);
  AppendRootOpen(out, "personList");
  for (const Person& p : persons) {
    out.append("  <person");
    append_xml_attr(out, "displayName", p.display_name);
    append_xml_attr(out, "id", p.id);
    if (!p.user_id.empty()) {
      append_xml_attr(out, "userId", p.user_id);
    }
    if (!p.provider_id.empty()) {
      append_xml_attr(out, "providerId", p.provider_id);
    }
    out.append("/>\n");
  }
  out.append("</personList>\n");
  return out;
}

bool is_thread_stub_author(std::string_view author) {
  return author.substr(0, kStubAuthorPrefix.size()) == kStubAuthorPrefix;
}

void drop_thread_stubs(std::vector<CellComment>& notes, const std::vector<ThreadedComment>& threads) {
  if (threads.empty()) {
    return;
  }
  std::vector<CellComment> kept;
  kept.reserve(notes.size());
  std::unordered_map<std::string_view, std::size_t> first_index_by_id;
  first_index_by_id.reserve(threads.size());
  for (std::size_t i = 0; i < threads.size(); ++i) {
    first_index_by_id.try_emplace(std::string_view(threads[i].id), i);
  }
  for (CellComment& note : notes) {
    const std::string_view author(note.author);
    bool stub = false;
    if (is_thread_stub_author(author)) {
      const auto found = first_index_by_id.find(author.substr(kStubAuthorPrefix.size()));
      if (found != first_index_by_id.end()) {
        const ThreadedComment& root = threads[found->second];
        stub = root.parent_id.empty() && root.row == note.row && root.col == note.col;
      }
    }
    if (!stub) {
      kept.push_back(std::move(note));
    }
  }
  notes = std::move(kept);
}

std::vector<CellComment> legacy_comments_with_stubs(const Sheet& sheet, ExcelLocale locale) {
  const std::vector<ThreadedComment>& threads = sheet.threaded_comments();
  std::vector<CellComment> out;
  out.reserve(sheet.comments().size() + threads.size());
  const LegacyStubIndexes indexes(threads);
  for (const CellComment& note : sheet.comments()) {
    if (indexes.root_coordinates.find(CellCoordinateKey(note.row, note.col)) == indexes.root_coordinates.end()) {
      out.push_back(note);
    }
  }
  const std::vector<std::size_t> no_replies;
  for (std::size_t i = 0; i < threads.size(); ++i) {
    const ThreadedComment& t = threads[i];
    if (t.parent_id.empty()) {
      const auto replies = indexes.replies_by_parent_id.find(std::string_view(t.id));
      const std::vector<std::size_t>& reply_indices =
          replies == indexes.replies_by_parent_id.end() ? no_replies : replies->second;
      out.emplace_back(t.row, t.col, std::string(kStubAuthorPrefix) + t.id,
                       threaded_comment_stub_text(threads, i, reply_indices, locale), t.id);
    }
  }
  return out;
}

}  // namespace formulon::io
