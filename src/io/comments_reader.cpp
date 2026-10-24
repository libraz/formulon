//
// Implementation of the comments-part reader. See comments_reader.h
// for the contract; the format is a small fragment of OOXML so the
// implementation is intentionally short and self-contained.

#include "io/comments_reader.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "io/cell_parser.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon::io {

Expected<std::vector<CellComment>, Error> read_comments(const std::vector<std::uint8_t>& bytes,
                                                        std::vector<std::string>* authors_out) {
  std::vector<CellComment> out;
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, bytes, "comments_reader", "comments part"));
  pugi::xml_node root = doc.child("comments");
  if (!root) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "comments part: missing <comments> root",
                      "context=comments_reader");
  }

  // Author table: `<authors><author>Alice</author>...</authors>`. The
  // table is referenced by `<comment authorId="N">` indices.
  std::vector<std::string> authors;
  if (pugi::xml_node a = root.child("authors"); a) {
    for (pugi::xml_node au = a.child("author"); au; au = au.next_sibling("author")) {
      std::string author;
      AppendOoxmlTextUnescaped(author, au.text().get());
      authors.push_back(std::move(author));
    }
  }

  if (authors_out != nullptr) {
    *authors_out = authors;
  }
  pugi::xml_node list = root.child("commentList");
  if (!list) {
    // Empty author table with no <commentList> is technically legal;
    // surface the empty list rather than failing.
    return out;
  }
  for (pugi::xml_node c = list.child("comment"); c; c = c.next_sibling("comment")) {
    const std::string_view ref = c.attribute("ref").value();
    if (ref.empty()) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "comment: missing/empty ref", "context=comments_reader");
    }
    auto rc = parse_a1(ref);
    if (!rc) {
      std::string ctx("context=comments_reader ref=");
      ctx.append(ref);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "comment: ref unparseable", std::move(ctx));
    }
    CellComment cc;
    cc.row = rc.value().first;
    cc.col = rc.value().second;
    const auto author_id = parse_xml_u32_attr_strict(c.attribute("authorId"));
    if (author_id && static_cast<std::size_t>(*author_id) < authors.size()) {
      cc.author = authors[static_cast<std::size_t>(*author_id)];
    }
    cc.uid = c.attribute("xr:uid").value();
    if (pugi::xml_node text = c.child("text"); text) {
      // Comment XML never carries `<rPh>` in practice (kana phonetic
      // guides only attach to shared-string entries), so the unified
      // walker's rPh-skip is a no-op for this path.
      (void)append_rich_text(text, cc.text);
    }
    out.push_back(std::move(cc));
  }
  return out;
}

namespace {

/// The text between `<tag ...>` and `</tag>` inside `block`, or empty.
std::string_view ElementText(std::string_view block, std::string_view open_tag, std::string_view close_tag) {
  const std::size_t open = block.find(open_tag);
  if (open == std::string_view::npos) {
    return {};
  }
  const std::size_t start = block.find('>', open);
  const std::size_t end = block.find(close_tag, open);
  if (start == std::string_view::npos || end == std::string_view::npos || end < start) {
    return {};
  }
  return block.substr(start + 1U, end - start - 1U);
}

std::string_view Trim(std::string_view text) {
  while (!text.empty() &&
         (text.front() == ' ' || text.front() == '\t' || text.front() == '\r' || text.front() == '\n')) {
    text.remove_prefix(1U);
  }
  while (!text.empty() && (text.back() == ' ' || text.back() == '\t' || text.back() == '\r' || text.back() == '\n')) {
    text.remove_suffix(1U);
  }
  return text;
}

bool ParseInt(std::string_view text, std::int64_t& out) {
  text = Trim(text);
  if (text.empty() || text.size() > 10U) {
    return false;
  }
  const bool negative = text.front() == '-';
  if (negative) {
    text.remove_prefix(1U);
  }
  if (text.empty()) {
    return false;
  }
  std::int64_t value = 0;
  for (const char ch : text) {
    if (ch < '0' || ch > '9') {
      return false;
    }
    value = value * 10 + (ch - '0');
  }
  out = negative ? -value : value;
  return true;
}

}  // namespace

void apply_vml_note_shapes(std::string_view vml, std::vector<CellComment>& comments) {
  constexpr std::string_view kShapeOpen = "<v:shape ";
  constexpr std::string_view kShapeClose = "</v:shape>";
  // The first comment at a cell takes its shape.
  std::unordered_map<std::uint64_t, std::size_t> by_cell;
  by_cell.reserve(comments.size());
  for (std::size_t i = 0; i < comments.size(); ++i) {
    by_cell.emplace((static_cast<std::uint64_t>(comments[i].row) << 32U) | comments[i].col, i);
  }
  // Each byte is scanned once: a shape is the last open before its close,
  // and the search resumes after that close.
  for (std::size_t at = vml.find(kShapeOpen); at != std::string_view::npos;) {
    const std::size_t end = vml.find(kShapeClose, at);
    if (end == std::string_view::npos) {
      break;
    }
    for (std::size_t inner = vml.find(kShapeOpen, at + 1U); inner < end; inner = vml.find(kShapeOpen, at + 1U)) {
      at = inner;
    }
    const std::size_t shape_at = at;
    at = vml.find(kShapeOpen, end + kShapeClose.size());
    const std::string_view shape = vml.substr(shape_at, end - shape_at);
    const std::size_t tag_end = shape.find('>');
    const std::string_view client = ElementText(shape, "<x:ClientData", "</x:ClientData>");
    if (tag_end == std::string_view::npos || shape.find("ObjectType=\"Note\"") == std::string_view::npos) {
      continue;
    }
    std::int64_t row = 0;
    std::int64_t col = 0;
    if (!ParseInt(ElementText(client, "<x:Row", "</x:Row>"), row) ||
        !ParseInt(ElementText(client, "<x:Column", "</x:Column>"), col) || row < 0 || col < 0) {
      continue;
    }
    if (row > std::numeric_limits<std::uint32_t>::max() || col > std::numeric_limits<std::uint32_t>::max()) {
      continue;
    }
    const auto found = by_cell.find((static_cast<std::uint64_t>(row) << 32U) | static_cast<std::uint64_t>(col));
    if (found != by_cell.end()) {
      CellComment& comment = comments[found->second];
      comment.vml_shape = std::string(vml.substr(shape_at, end + kShapeClose.size() - shape_at));
      comment.visible = shape.substr(0, tag_end).find("visibility:visible") != std::string_view::npos ||
                        client.find("<x:Visible") != std::string_view::npos;
      // Eight comma-separated integers; anything else leaves the default box.
      std::string_view anchor = ElementText(client, "<x:Anchor", "</x:Anchor>");
      std::array<std::int32_t, 8> box{};
      bool ok = !anchor.empty();
      for (std::size_t k = 0; k < box.size() && ok; ++k) {
        const std::size_t comma = anchor.find(',');
        std::int64_t value = 0;
        ok = ParseInt(anchor.substr(0, comma), value) && (k + 1U == box.size() || comma != std::string_view::npos);
        box[k] = static_cast<std::int32_t>(value);
        anchor = comma == std::string_view::npos ? std::string_view() : anchor.substr(comma + 1U);
      }
      if (ok) {
        box[0] -= static_cast<std::int32_t>(comment.col);
        box[2] -= static_cast<std::int32_t>(comment.row);
        box[4] -= static_cast<std::int32_t>(comment.col);
        box[6] -= static_cast<std::int32_t>(comment.row);
        comment.has_box = true;
        comment.box = box;
      }
    }
  }
}

std::string vml_non_note_shapes(std::string_view vml) {
  constexpr std::string_view kTypeOpen = "<v:shapetype";
  constexpr std::string_view kTypeClose = "</v:shapetype>";
  constexpr std::string_view kShapeOpen = "<v:shape ";
  constexpr std::string_view kShapeClose = "</v:shape>";
  std::string out;
  for (std::size_t at = vml.find("<v:shape"); at != std::string_view::npos; at = vml.find("<v:shape", at)) {
    const bool is_type = vml.compare(at, kTypeOpen.size(), kTypeOpen) == 0;
    if (!is_type && vml.compare(at, kShapeOpen.size(), kShapeOpen) != 0) {
      at += kShapeOpen.size() - 1U;
      continue;
    }
    const std::size_t tag_end = vml.find('>', at);
    if (tag_end == std::string_view::npos) {
      break;
    }
    std::size_t end = tag_end + 1U;
    if (vml[tag_end - 1U] != '/') {
      const std::string_view close = is_type ? kTypeClose : kShapeClose;
      const std::size_t close_at = vml.find(close, tag_end);
      if (close_at == std::string_view::npos) {
        break;
      }
      end = close_at + close.size();
    }
    const std::string_view element = vml.substr(at, end - at);
    const std::string_view open_tag = vml.substr(at, tag_end - at);
    const bool keep = is_type ? open_tag.find("id=\"_x0000_t202\"") == std::string_view::npos
                              : element.find("ObjectType=\"Note\"") == std::string_view::npos;
    if (keep) {
      out.append(element);
      out.push_back('\n');
    }
    at = end;
  }
  return out;
}

}  // namespace formulon::io
