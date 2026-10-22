//
// Implementation of `comments_writer.h`. Format details live in the
// header; this TU is self-contained so the writer side has minimal
// linkage requirements.

#include "io/comments_writer.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/ooxml/cell_ref_writer.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "sheet.h"

namespace formulon::io {
std::string write_comments(const std::vector<CellComment>& comments, const std::vector<std::string>& known_authors) {
  if (comments.empty()) {
    return {};
  }

  // Build a stable author table: the source's authors, then first-occurrence order.
  std::vector<std::string> authors;
  std::unordered_map<std::string, std::uint32_t> author_index;
  const auto add_author = [&](const std::string& author) {
    if (author_index.find(author) == author_index.end()) {
      author_index.emplace(author, static_cast<std::uint32_t>(authors.size()));
      authors.push_back(author);
    }
  };
  for (const std::string& a : known_authors) {
    add_author(a);
  }
  for (const CellComment& c : comments) {
    add_author(c.author);
  }

  std::string out;
  out.reserve(256 + comments.size() * 96 + authors.size() * 32);
  out.append(kXmlDecl);
  bool any_uid = false;
  for (const CellComment& c : comments) {
    any_uid = any_uid || !c.uid.empty();
  }
  out.append("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
  if (any_uid) {
    out.append(
        " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"xr\""
        " xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\"");
  }
  out.append(">\n");
  out.append("  <authors>\n");
  for (const std::string& a : authors) {
    out.append("    <author>");
    AppendXmlEscaped(out, a);
    out.append("</author>\n");
  }
  out.append("  </authors>\n");
  out.append("  <commentList>\n");
  for (const CellComment& c : comments) {
    out.append("    <comment ref=\"");
    AppendCellRefForRef(out, c.row, c.col);
    out.append("\" authorId=\"");
    out.append(std::to_string(author_index[c.author]));
    out.push_back('"');
    if (!c.uid.empty()) {
      append_xml_attr(out, "xr:uid", c.uid);
    }
    out.append(">\n");
    out.append("      <text><r><t xml:space=\"preserve\">");
    AppendXmlEscaped(out, c.text);
    out.append("</t></r></text>\n");
    out.append("    </comment>\n");
  }
  out.append("  </commentList>\n");
  out.append("</comments>\n");
  return out;
}

namespace {

/// The note box in absolute cells: the stored one clamped onto the grid, else
/// Excel's default (two columns wide, four rows tall, one column right of the
/// cell and starting a row above it).
std::array<std::int64_t, 8> NoteBox(const CommentShape& shape) {
  const std::uint32_t row = shape.row;
  const std::uint32_t col = shape.col;
  std::array<std::int64_t, 8> box = {std::min(col + 1U, Sheet::kMaxCols - 1U),
                                     15,
                                     row == 0U ? 0U : row - 1U,
                                     row == 0U ? 2 : 10,
                                     std::min(col + 3U, Sheet::kMaxCols - 1U),
                                     15,
                                     0,
                                     4};
  box[6] = std::min<std::int64_t>(box[2] + 4, Sheet::kMaxRows - 1U);
  if (shape.has_box) {
    // A stored box is relative to the cell; edits may have pushed it off the grid.
    const std::array<std::int32_t, 8>& rel = shape.box;
    const std::int64_t col_cap = Sheet::kMaxCols - 1U;
    const std::int64_t row_cap = Sheet::kMaxRows - 1U;
    box = {std::clamp<std::int64_t>(col + std::int64_t{rel[0]}, 0, col_cap), rel[1],
           std::clamp<std::int64_t>(row + std::int64_t{rel[2]}, 0, row_cap), rel[3],
           std::clamp<std::int64_t>(col + std::int64_t{rel[4]}, 0, col_cap), rel[5],
           std::clamp<std::int64_t>(row + std::int64_t{rel[6]}, 0, row_cap), rel[7]};
  }
  return box;
}

/// Replaces the text inside the first `<tag ...>...</tag>` of `xml`; no-op when absent.
void ReplaceInner(std::string& xml, std::string_view tag, std::string_view value) {
  const std::string open = "<" + std::string(tag);
  const std::string close = "</" + std::string(tag) + ">";
  const std::size_t at = xml.find(open);
  const std::size_t start = at == std::string::npos ? at : xml.find('>', at);
  const std::size_t end = start == std::string::npos ? start : xml.find(close, start);
  if (end != std::string::npos) {
    xml.replace(start + 1U, end - start - 1U, value);
  }
}

/// Sets the note's visibility in the `<v:shape>` style and `<x:Visible>` marker.
void PatchVisibility(std::string& xml, bool visible) {
  const std::size_t tag_end = xml.find('>');
  const std::size_t style = xml.find(" style=");
  if (tag_end != std::string::npos && style < tag_end && style + 7U < xml.size()) {
    const char quote_char = xml[style + 7U];
    const std::size_t value = style + 8U;
    const std::size_t quote = xml.find(quote_char, value);
    const std::size_t vis = xml.find("visibility:", value);
    const std::string_view word = visible ? "visible" : "hidden";
    if (vis < quote) {
      const std::size_t from = vis + 11U;
      xml.replace(from, xml.find_first_of(std::string(";") + quote_char, from) - from, word);
    } else if (!visible && quote != std::string::npos) {
      xml.insert(quote, std::string(";visibility:") + std::string(word));
    }
  } else if (!visible && tag_end != std::string::npos) {
    xml.insert(tag_end, " style=\"visibility:hidden\"");
  }
  const std::size_t marker = xml.find("<x:Visible");
  const bool has_marker = marker != std::string::npos && marker + 10U < xml.size() &&
                          (xml[marker + 10U] == '/' || xml[marker + 10U] == '>' || xml[marker + 10U] == ' ');
  if (has_marker && !visible) {
    const std::size_t close = xml.find('>', marker);
    const bool self_closed = close != std::string::npos && xml[close - 1U] == '/';
    const std::size_t pair = self_closed ? close : xml.find("</x:Visible>", marker);
    if (pair != std::string::npos) {
      xml.erase(marker, (self_closed ? close + 1U : pair + 12U) - marker);
    }
  } else if (!has_marker && visible) {
    const std::size_t at = xml.rfind("</x:ClientData>");
    if (at != std::string::npos) {
      xml.insert(at, "<x:Visible/>");
    }
  }
}

/// A source note `<v:shape>` with its anchor, visibility and cell set to `shape`'s.
std::string PatchNoteShape(const CommentShape& shape) {
  std::string xml = shape.xml;
  const auto box = NoteBox(shape);
  std::string anchor;
  for (std::size_t k = 0; k < box.size(); ++k) {
    anchor.append(std::to_string(box[k])).append(k + 1U < box.size() ? ", " : "");
  }
  ReplaceInner(xml, "x:Anchor", anchor);
  ReplaceInner(xml, "x:Row", std::to_string(shape.row));
  ReplaceInner(xml, "x:Column", std::to_string(shape.col));
  PatchVisibility(xml, shape.visible);
  return xml;
}

/// The numeric part of a `_x0000_s<N>` shape id in the opening tag of `xml`, or 0.
std::uint64_t ShapeNumber(const std::string& xml) {
  const std::size_t at = xml.find(" id=\"_x0000_s");
  if (at == std::string::npos || at > xml.find('>')) {
    return 0;
  }
  return std::strtoull(xml.c_str() + at + 13U, nullptr, 10);
}

}  // namespace

std::string write_vml_drawing(const std::vector<CommentShape>& shapes, std::uint32_t drawing_id) {
  std::string out;
  out.reserve(512 + shapes.size() * 640);
  out.append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\n");
  out.append(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
  out.append(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n");
  out.append(" <o:shapelayout v:ext=\"edit\">\n");
  out.append("  <o:idmap v:ext=\"edit\" data=\"");
  out.append(std::to_string(drawing_id));
  out.append("\"/>\n");
  out.append(" </o:shapelayout>\n");
  out.append(" <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\n");
  out.append("  path=\"m,l,21600r21600,l21600,xe\">\n");
  out.append("  <v:stroke joinstyle=\"miter\"/>\n");
  out.append("  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n");
  out.append(" </v:shapetype>\n");
  // A kept shape keeps its id only inside this drawing's block and only once.
  const std::uint64_t block_first = static_cast<std::uint64_t>(drawing_id) * 1024U + 1U;
  std::unordered_set<std::uint64_t> used;
  std::vector<bool> renumber(shapes.size(), false);
  for (std::size_t i = 0; i < shapes.size(); ++i) {
    const std::uint64_t number = ShapeNumber(shapes[i].xml);
    if (number == 0U) {
      continue;
    }
    renumber[i] = number < block_first || number >= block_first + 1023U || !used.insert(number).second;
  }
  std::uint64_t next_shape = block_first;
  for (std::size_t i = 0; i < shapes.size(); ++i) {
    const std::uint32_t row = shapes[i].row;
    const std::uint32_t col = shapes[i].col;
    if (!shapes[i].xml.empty()) {
      std::string xml = PatchNoteShape(shapes[i]);
      if (renumber[i]) {
        while (used.count(next_shape) != 0U) {
          ++next_shape;
        }
        const std::size_t at = xml.find(" id=\"_x0000_s") + 13U;
        xml.replace(at, xml.find('"', at) - at, std::to_string(next_shape++));
      }
      out.append(" ").append(xml).append("\n");
      continue;
    }
    const auto box = NoteBox(shapes[i]);
    while (used.count(next_shape) != 0U) {
      ++next_shape;
    }
    out.append(" <v:shape id=\"_x0000_s");
    out.append(std::to_string(next_shape++));
    out.append("\" type=\"#_x0000_t202\" style=\"position:absolute;width:108pt;height:59.25pt;z-index:");
    out.append(std::to_string(i + 1U));
    out.append(shapes[i].visible ? ";visibility:visible" : ";visibility:hidden");
    out.append("\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\n");
    out.append("  <v:fill color2=\"#ffffe1\"/>\n");
    out.append("  <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\n");
    out.append("  <v:path o:connecttype=\"none\"/>\n");
    out.append("  <v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox>\n");
    out.append("  <x:ClientData ObjectType=\"Note\">\n");
    out.append("   <x:MoveWithCells/>\n");
    out.append("   <x:SizeWithCells/>\n");
    out.append("   <x:Anchor>");
    for (std::size_t k = 0; k < box.size(); ++k) {
      out.append(std::to_string(box[k])).append(k + 1U < box.size() ? ", " : "</x:Anchor>\n");
    }
    out.append("   <x:AutoFill>False</x:AutoFill>\n");
    out.append("   <x:Row>");
    out.append(std::to_string(row));
    out.append("</x:Row>\n");
    out.append("   <x:Column>");
    out.append(std::to_string(col));
    out.append("</x:Column>\n");
    if (shapes[i].visible) {
      out.append("   <x:Visible/>\n");
    }
    out.append("  </x:ClientData>\n");
    out.append(" </v:shape>\n");
  }
  out.append("</xml>\n");
  return out;
}

std::string write_vml_drawing_stub() {
  return write_vml_drawing({}, 1U);
}

}  // namespace formulon::io
