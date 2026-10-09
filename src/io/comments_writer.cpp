//
// Implementation of `comments_writer.h`. Format details live in the
// header; this TU is self-contained so the writer side has minimal
// linkage requirements.

#include "io/comments_writer.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "io/ooxml/cell_ref_writer.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "sheet.h"

namespace formulon::io {
std::string write_comments(const std::vector<CellComment>& comments) {
  if (comments.empty()) {
    return {};
  }

  // Build a stable author table: first-occurrence order.
  std::vector<std::string> authors;
  std::unordered_map<std::string, std::uint32_t> author_index;
  for (const CellComment& c : comments) {
    if (author_index.find(c.author) == author_index.end()) {
      author_index.emplace(c.author, static_cast<std::uint32_t>(authors.size()));
      authors.push_back(c.author);
    }
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

std::string write_vml_drawing(const std::vector<std::pair<std::uint32_t, std::uint32_t>>& anchors,
                              std::uint32_t drawing_id) {
  std::string out;
  out.reserve(512 + anchors.size() * 640);
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
  const std::uint64_t first_shape = static_cast<std::uint64_t>(drawing_id) * 1024U + 1U;
  for (std::size_t i = 0; i < anchors.size(); ++i) {
    const std::uint32_t row = anchors[i].first;
    const std::uint32_t col = anchors[i].second;
    // Excel's default note box: two columns wide, four rows tall, offset
    // one column right of the cell and starting a row above it.
    const std::uint32_t left = std::min(col + 1U, Sheet::kMaxCols - 1U);
    const std::uint32_t right = std::min(col + 3U, Sheet::kMaxCols - 1U);
    const std::uint32_t top = row == 0U ? 0U : row - 1U;
    const std::uint32_t bottom = std::min(top + 4U, Sheet::kMaxRows - 1U);
    out.append(" <v:shape id=\"_x0000_s");
    out.append(std::to_string(first_shape + i));
    out.append("\" type=\"#_x0000_t202\" style=\"position:absolute;width:108pt;height:59.25pt;z-index:");
    out.append(std::to_string(i + 1U));
    out.append(";visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\">\n");
    out.append("  <v:fill color2=\"#ffffe1\"/>\n");
    out.append("  <v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>\n");
    out.append("  <v:path o:connecttype=\"none\"/>\n");
    out.append("  <v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox>\n");
    out.append("  <x:ClientData ObjectType=\"Note\">\n");
    out.append("   <x:MoveWithCells/>\n");
    out.append("   <x:SizeWithCells/>\n");
    out.append("   <x:Anchor>");
    out.append(std::to_string(left)).append(", 15, ");
    out.append(std::to_string(top)).append(row == 0U ? ", 2, " : ", 10, ");
    out.append(std::to_string(right)).append(", 15, ");
    out.append(std::to_string(bottom)).append(", 4</x:Anchor>\n");
    out.append("   <x:AutoFill>False</x:AutoFill>\n");
    out.append("   <x:Row>");
    out.append(std::to_string(row));
    out.append("</x:Row>\n");
    out.append("   <x:Column>");
    out.append(std::to_string(col));
    out.append("</x:Column>\n");
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
