// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `comments_writer.h`. Format details live in the
// header; this TU is self-contained so the writer side has minimal
// linkage requirements.

#include "io/comments_writer.h"

#include <cstdint>
#include <string>
#include <unordered_map>
#include <vector>

#include "io/xml_escape.h"
#include "sheet.h"

namespace formulon::io {
namespace {

/// Bijective base-26 column letters, matching the cell-parser's `parse_a1`.
void AppendColumnLetters(std::string& out, std::uint32_t col) {
  char buf[4];
  std::uint32_t i = 0;
  std::uint32_t v = col + 1;
  while (v > 0 && i < 4) {
    const std::uint32_t rem = (v - 1) % 26U;
    buf[i++] = static_cast<char>('A' + rem);
    v = (v - 1) / 26U;
  }
  while (i > 0) {
    out.push_back(buf[--i]);
  }
}

/// Appends an A1 cell reference (`A1`, `XFD1048576`, ...).
void AppendCellRef(std::string& out, std::uint32_t row, std::uint32_t col) {
  AppendColumnLetters(out, col);
  out.append(std::to_string(row + 1));
}

}  // namespace

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
  out.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n");
  out.append(
      "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
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
    AppendCellRef(out, c.row, c.col);
    out.append("\" authorId=\"");
    out.append(std::to_string(author_index[c.author]));
    out.append("\">\n");
    out.append("      <text><r><t xml:space=\"preserve\">");
    AppendXmlEscaped(out, c.text);
    out.append("</t></r></text>\n");
    out.append("    </comment>\n");
  }
  out.append("  </commentList>\n");
  out.append("</comments>\n");
  return out;
}

std::string write_vml_drawing_stub() {
  // Minimal but well-formed VML payload. The shape has no visible
  // styling — Excel regenerates per-comment shapes on first edit. The
  // namespaces match the canonical Excel emission so older parsers
  // recognise the document.
  std::string out;
  out.reserve(384);
  out.append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\n");
  out.append("     xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
  out.append("     xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n");
  out.append("  <o:shapelayout v:ext=\"edit\">\n");
  out.append("    <o:idmap v:ext=\"edit\" data=\"1\"/>\n");
  out.append("  </o:shapelayout>\n");
  out.append("  <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\n");
  out.append("               path=\"m,l,21600r21600,l21600,xe\">\n");
  out.append("    <v:stroke joinstyle=\"miter\"/>\n");
  out.append("    <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n");
  out.append("  </v:shapetype>\n");
  out.append("</xml>\n");
  return out;
}

}  // namespace formulon::io
