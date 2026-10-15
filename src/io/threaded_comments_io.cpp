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
constexpr std::string_view kMainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
constexpr std::string_view kStubAuthorPrefix = "tc=";

void AppendRootOpen(std::string& out, std::string_view root) {
  out.append(kXmlDecl);
  out.push_back('<');
  out.append(root);
  append_xml_attr(out, "xmlns", kThreadedNs);
  append_xml_attr(out, "xmlns:x", kMainNs);
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
  }
};

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
    c.parent_id = node.attribute("parentId").value();
    c.done = attr_bool(node, "done");
    AppendOoxmlTextUnescaped(c.text, node.child("text").text().get());
    for (pugi::xml_node m = node.child("mentions").child("mention"); m; m = m.next_sibling("mention")) {
      Mention mention;
      mention.person_id = m.attribute("mentionpersonId").value();
      mention.mention_id = m.attribute("mentionId").value();
      mention.start = attr_u32(m, "startIndex");
      mention.length = attr_u32(m, "length");
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
  for (const ThreadedComment& c : comments) {
    out.append("  <threadedComment");
    append_xml_attr(out, "ref", a1::encode_a1(c.row, c.col));
    append_xml_attr(out, "dT", c.created);
    append_xml_attr(out, "personId", c.person_id);
    append_xml_attr(out, "id", c.id);
    if (!c.parent_id.empty()) {
      append_xml_attr(out, "parentId", c.parent_id);
    } else if (c.done) {
      out.append(" done=\"1\"");
    }
    out.append("><text>");
    AppendXmlEscaped(out, c.text);
    out.append("</text>");
    if (!c.mentions.empty()) {
      out.append("<mentions>");
      for (const Mention& m : c.mentions) {
        out.append("<mention");
        append_xml_attr(out, "mentionpersonId", m.person_id);
        append_xml_attr(out, "mentionId", m.mention_id);
        append_xml_attr_uint(out, "startIndex", m.start);
        append_xml_attr_uint(out, "length", m.length);
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
    if (author.substr(0, kStubAuthorPrefix.size()) == kStubAuthorPrefix) {
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

std::vector<CellComment> legacy_comments_with_stubs(const Sheet& sheet) {
  const std::vector<ThreadedComment>& threads = sheet.threaded_comments();
  std::vector<CellComment> out;
  out.reserve(sheet.comments().size() + threads.size());
  const LegacyStubIndexes indexes(threads);
  for (const CellComment& note : sheet.comments()) {
    if (indexes.root_coordinates.find(CellCoordinateKey(note.row, note.col)) == indexes.root_coordinates.end()) {
      out.push_back(note);
    }
  }
  for (std::size_t i = 0; i < threads.size(); ++i) {
    const ThreadedComment& t = threads[i];
    if (t.parent_id.empty()) {
      const auto replies = indexes.replies_by_parent_id.find(std::string_view(t.id));
      const std::vector<std::size_t> no_replies;
      const std::vector<std::size_t>& reply_indices =
          replies == indexes.replies_by_parent_id.end() ? no_replies : replies->second;
      out.emplace_back(t.row, t.col, std::string(kStubAuthorPrefix) + t.id,
                       threaded_comment_stub_text(threads, i, reply_indices), t.id);
    }
  }
  return out;
}

}  // namespace formulon::io
