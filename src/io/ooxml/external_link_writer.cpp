//
// External-link body, rels and `[N]` numbering for the OOXML writer. See
// the header for the contract.

#include "io/ooxml/external_link_writer.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "external_book.h"
#include "external_link.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml_defs.h"
#include "io/stored_cell_error.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "parser/ast.h"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"
#include "utils/expected.h"  // FM_CHECK
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

/// Root and `<externalBook>` start tags as Excel writes them, so the
/// `xxl21:` alternate-URL element resolves without further declarations.
constexpr std::string_view kExternalLinkOpen =
    "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14 xxl21\" "
    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
    "xmlns:xxl21=\"http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021\">"
    "<externalBook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"";

std::string_view TargetRelId(const ExternalLinkRecord& rec) {
  return rec.body_rel_id.empty() ? std::string_view("rId1") : std::string_view(rec.body_rel_id);
}

std::string_view AbsoluteRelId(const ExternalLinkRecord& rec) {
  if (!rec.absolute_rel_id.empty()) {
    return rec.absolute_rel_id;
  }
  return TargetRelId(rec) == "rId2" ? std::string_view("rId3") : std::string_view("rId2");
}

void AppendAbsoluteCell(std::string& out, std::uint32_t row, std::uint32_t col) {
  out.push_back('$');
  a1::append_column_letters(out, col);
  out.push_back('$');
  out.append(std::to_string(row + 1U));
}

/// `='Sheet'!$A$1` / `='Sheet'!$A$1:$B$2`, the form the reader resolves;
/// a name the cache could not resolve is `#REF!`.
std::string RefersTo(const ExternalBook& book, const ExternalBookName& name) {
  if (!name.resolvable || name.sheet >= book.sheet_names.size()) {
    return "#REF!";
  }
  std::string out("='");
  for (const char c : book.sheet_names[name.sheet]) {
    if (c == '\'') {
      out.push_back('\'');
    }
    out.push_back(c);
  }
  out.append("'!");
  AppendAbsoluteCell(out, name.row, name.col);
  if (name.is_range) {
    out.push_back(':');
    AppendAbsoluteCell(out, name.row_end, name.col_end);
  }
  return out;
}

struct CachedCell {
  std::uint32_t row;
  std::uint32_t col;
  const ExternalCell* cell;
};

/// Cached cells of `sheet` in row-major order. Unpacks the key layout
/// `ExternalBook::cell_key` documents (16-bit sheet, 21-bit row, 14-bit column).
std::vector<CachedCell> CellsOfSheet(const ExternalBook& book, std::uint32_t sheet) {
  constexpr unsigned kColBits = 14U;
  constexpr unsigned kRowBits = 21U;
  std::vector<CachedCell> out;
  for (const std::uint64_t key : book.sorted_cell_keys()) {
    if (static_cast<std::uint32_t>(key >> (kColBits + kRowBits)) != sheet) {
      continue;
    }
    const auto row = static_cast<std::uint32_t>((key >> kColBits) & ((std::uint64_t{1} << kRowBits) - 1U));
    const auto col = static_cast<std::uint32_t>(key & ((std::uint64_t{1} << kColBits) - 1U));
    FM_CHECK(ExternalBook::cell_key(sheet, row, col) == key, "external cell key layout changed");
    out.push_back(CachedCell{row, col, &book.cells.at(key)});
  }
  return out;
}

void AppendCachedCell(std::string& out, const CachedCell& c) {
  const Value value = c.cell->resolved();
  out.append("<cell r=\"");
  out.append(a1::encode_a1(c.row, c.col));
  out.push_back('"');
  if (value.is_blank()) {
    out.append("/>");
    return;
  }
  if (value.is_text()) {
    out.append(" t=\"str\"><v xml:space=\"preserve\">");
    AppendXmlEscaped(out, value.as_text());
  } else if (value.is_boolean()) {
    out.append(" t=\"b\"><v>");
    out.push_back(value.as_boolean() ? '1' : '0');
  } else if (value.is_error()) {
    out.append(" t=\"e\"><v>");
    out.append(display_name(stored_cell_error(value.as_error())));
  } else {
    out.append("><v>");
    append_xml_number(out, value.is_number() ? value.as_number() : 0.0);
  }
  out.append("</v></cell>");
}

void AppendSheetData(std::string& out, const ExternalBook& book, std::uint32_t sheet) {
  out.append("<sheetData sheetId=\"");
  out.append(std::to_string(sheet));
  out.append("\">");
  const std::vector<CachedCell> cells = CellsOfSheet(book, sheet);
  for (std::size_t i = 0; i < cells.size();) {
    const std::uint32_t row = cells[i].row;
    out.append("<row r=\"");
    out.append(std::to_string(row + 1U));
    out.append("\">");
    for (; i < cells.size() && cells[i].row == row; ++i) {
      AppendCachedCell(out, cells[i]);
    }
    out.append("</row>");
  }
  out.append("</sheetData>");
}

bool ContainsExternalRef(const parser::AstNode& root) {
  std::vector<const parser::AstNode*> pending{&root};
  while (!pending.empty()) {
    const parser::AstNode& node = *pending.back();
    pending.pop_back();
    if (node.kind() == parser::NodeKind::ExternalRef) {
      return true;
    }
    for (const parser::AstNode* child : parser::child_nodes(node)) {
      pending.push_back(child);
    }
  }
  return false;
}

}  // namespace

ExternalLinkOrdinals::ExternalLinkOrdinals(const Workbook& wb, std::vector<std::uint32_t> written_indices)
    : links_(wb.external_book_indexer()), written_indices_(std::move(written_indices)) {}

std::uint32_t ExternalLinkOrdinals::OrdinalOfIndex(std::uint32_t index) const noexcept {
  const auto it = std::find(written_indices_.begin(), written_indices_.end(), index);
  if (it == written_indices_.end()) {
    return 0U;
  }
  return static_cast<std::uint32_t>(it - written_indices_.begin()) + 1U;
}

std::uint32_t ExternalLinkOrdinals::Ordinal(const void* ctx, std::string_view path, std::string_view book) {
  const auto* self = static_cast<const ExternalLinkOrdinals*>(ctx);
  if (self->links_.index == nullptr) {
    return 0U;
  }
  return self->OrdinalOfIndex(self->links_.index(self->links_.ctx, path, book));
}

std::uint32_t ExternalLinkOrdinals::QualifierOrdinal(const void* ctx, std::string_view qualifier) {
  const auto* self = static_cast<const ExternalLinkOrdinals*>(ctx);
  if (self->links_.qualifier_index == nullptr) {
    return 0U;
  }
  return self->OrdinalOfIndex(self->links_.qualifier_index(self->links_.ctx, qualifier));
}

std::string BuildExternalLinkXml(const ExternalLinkRecord& rec) {
  const ExternalBook& book = rec.book;
  std::string out;
  out.reserve(512U + book.sheet_names.size() * 48U + book.names.size() * 64U + book.cells.size() * 40U);
  out.append(kXmlDecl);
  out.append(kExternalLinkOpen);
  AppendXmlAttrEscaped(out, TargetRelId(rec));
  out.append("\">");
  if (!rec.absolute_target.empty()) {
    out.append("<xxl21:alternateUrls><xxl21:absoluteUrl r:id=\"");
    AppendXmlAttrEscaped(out, AbsoluteRelId(rec));
    out.append("\"/></xxl21:alternateUrls>");
  }
  if (!book.sheet_names.empty()) {
    out.append("<sheetNames>");
    for (const std::string& sheet : book.sheet_names) {
      out.append("<sheetName val=\"");
      AppendXmlAttrEscaped(out, sheet);
      out.append("\"/>");
    }
    out.append("</sheetNames>");
  }
  if (!book.names.empty()) {
    out.append("<definedNames>");
    for (const ExternalBookName& name : book.names) {
      out.append("<definedName name=\"");
      AppendXmlAttrEscaped(out, name.name);
      out.push_back('"');
      if (name.exists) {
        out.append(" refersTo=\"");
        AppendXmlAttrEscaped(out, RefersTo(book, name));
        out.push_back('"');
      }
      if (name.scope_sheet != ExternalBook::kNoSheet) {
        out.append(" sheetId=\"");
        out.append(std::to_string(name.scope_sheet));
        out.push_back('"');
      }
      out.append("/>");
    }
    out.append("</definedNames>");
  }
  bool any_data = false;
  for (std::uint32_t s = 0; s < static_cast<std::uint32_t>(book.sheet_names.size()); ++s) {
    if (!book.sheet_has_data(s)) {
      continue;
    }
    if (!any_data) {
      out.append("<sheetDataSet>");
      any_data = true;
    }
    AppendSheetData(out, book, s);
  }
  if (any_data) {
    out.append("</sheetDataSet>");
  }
  out.append("</externalBook></externalLink>");
  return out;
}

std::string BuildExternalLinkRels(const ExternalLinkRecord& rec) {
  if (rec.target.empty() && rec.absolute_target.empty()) {
    return {};
  }
  std::string_view type = kRelExternalLinkPath;
  switch (rec.kind) {
    case ExternalLinkRecord::Kind::kOleLink:
      type = kRelOleLink;
      break;
    case ExternalLinkRecord::Kind::kDdeLink:
      type = kRelDdeLink;
      break;
    case ExternalLinkRecord::Kind::kExternalBook:
    case ExternalLinkRecord::Kind::kUnknown:
    default:
      break;
  }
  std::string out;
  out.reserve(384);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  // Absolute path first, the order Excel writes.
  if (!rec.absolute_target.empty()) {
    AppendRelationship(out, AbsoluteRelId(rec), kRelExternalLinkPath, rec.absolute_target, /*target_external=*/true,
                       /*escape_target=*/true);
  }
  if (!rec.target.empty()) {
    AppendRelationship(out, TargetRelId(rec), type, rec.target, rec.target_external, /*escape_target=*/true);
  }
  out.append("</Relationships>\n");
  return out;
}

bool NeedsStorageSpelling(const parser::AstNode& root, std::string_view storage) {
  return storage != parser::format_formula(root) || parser::formula_needs_storage_requote(root) ||
         ContainsExternalRef(root);
}

}  // namespace io
}  // namespace formulon
