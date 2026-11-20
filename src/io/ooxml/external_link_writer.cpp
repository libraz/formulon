//
// External-link body, rels and `[N]` numbering for the OOXML writer. See
// the header for the contract.

#include "io/ooxml/external_link_writer.h"

#include <algorithm>
#include <cmath>
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
#include "parser/ast_format.h"
#include "parser/external_book_spelling.h"
#include "parser/formula_prefix.h"
#include "parser/literal_spelling.h"
#include "sheet_name.h"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"
#include "utils/expected.h"  // FM_CHECK
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

/// Appends the root and `<externalBook>` start tags as Excel writes them, so
/// the `xxl21:` alternate-URL element resolves without further declarations.
void AppendExternalLinkOpen(std::string& out) {
  out.append("<externalLink");
  append_xml_ns_decl(out, "", kNsSpreadsheetMl);
  append_xml_ns_decl(out, "mc", kNsMarkupCompat);
  out.append(" mc:Ignorable=\"x14 xxl21\"");
  append_xml_ns_decl(out, "x14", kNsX14);
  append_xml_ns_decl(out, "xxl21", kNsXxl21);
  out.append("><externalBook");
  append_xml_ns_decl(out, "r", kNsOfficeRelationships);
  out.append(" r:id=\"");
}

std::string_view TargetRelId(const ExternalLinkRecord& rec) {
  return rec.body_rel_id.empty() ? std::string_view("rId1") : std::string_view(rec.body_rel_id);
}

std::string_view AbsoluteRelId(const ExternalLinkRecord& rec) {
  if (!rec.absolute_rel_id.empty()) {
    return rec.absolute_rel_id;
  }
  return TargetRelId(rec) == "rId2" ? std::string_view("rId3") : std::string_view("rId2");
}

/// The loaded relative-URL id when it clashes with neither sibling, else the first free of rId3, rId2, rId1.
std::string_view RelativeRelId(const ExternalLinkRecord& rec) {
  const std::string_view target = TargetRelId(rec);
  const std::string_view absolute = rec.absolute_target.empty() ? std::string_view() : AbsoluteRelId(rec);
  if (!rec.relative_rel_id.empty() && rec.relative_rel_id != target && rec.relative_rel_id != absolute) {
    return rec.relative_rel_id;
  }
  for (const std::string_view candidate :
       {std::string_view("rId3"), std::string_view("rId2"), std::string_view("rId1")}) {
    if (candidate != target && candidate != absolute) {
      return candidate;
    }
  }
  return "rId3";
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
  std::string out("=");
  parser::append_sheet_name_quoted_if(book.sheet_names[name.sheet], /*quote=*/true, out);
  out.push_back('!');
  AppendAbsoluteCell(out, name.row, name.col);
  if (name.is_range) {
    out.push_back(':');
    AppendAbsoluteCell(out, name.row_end, name.col_end);
  }
  return out;
}

constexpr unsigned kExternalColBits = 14U;
constexpr unsigned kExternalRowBits = 21U;

std::uint32_t CachedCellSheet(std::uint64_t key) {
  return static_cast<std::uint32_t>(key >> (kExternalColBits + kExternalRowBits));
}

std::uint32_t CachedCellRow(std::uint64_t key) {
  return static_cast<std::uint32_t>((key >> kExternalColBits) & ((std::uint64_t{1} << kExternalRowBits) - 1U));
}

std::uint32_t CachedCellColumn(std::uint64_t key) {
  return static_cast<std::uint32_t>(key & ((std::uint64_t{1} << kExternalColBits) - 1U));
}

void AppendCachedCell(std::string& out, const ExternalBook& book, std::uint32_t sheet, std::uint64_t key) {
  const std::uint32_t row = CachedCellRow(key);
  const std::uint32_t col = CachedCellColumn(key);
  FM_CHECK(ExternalBook::cell_key(sheet, row, col) == key, "external cell key layout changed");
  const Value value = book.cells.at(key).resolved();
  out.append("<cell r=\"");
  out.append(a1::encode_a1(row, col));
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
  } else if (value.is_error() || (value.is_number() && !std::isfinite(value.as_number()))) {
    out.append(" t=\"e\"><v>");
    const ErrorCode error = value.is_error() ? stored_cell_error(value.as_error()) : ErrorCode::Num;
    out.append(display_name(error));
  } else {
    out.append("><v>");
    append_xml_number(out, value.is_number() ? value.as_number() : 0.0);
  }
  out.append("</v></cell>");
}

void AppendSheetData(std::string& out, const ExternalBook& book, std::uint32_t sheet,
                     const std::vector<std::uint64_t>& sorted_keys, std::size_t begin, std::size_t end) {
  out.append("<sheetData sheetId=\"");
  out.append(std::to_string(sheet));
  out.append("\"");
  if (book.sheet_saves_refresh_error(sheet)) {
    out.append(" refreshError=\"1\"");
  }
  // Excel self-closes a sheet with no cached cells.
  if (begin == end) {
    out.append("/>");
    return;
  }
  out.append(">");
  for (std::size_t i = begin; i < end;) {
    const std::uint32_t row = CachedCellRow(sorted_keys[i]);
    out.append("<row r=\"");
    out.append(std::to_string(row + 1U));
    out.append("\">");
    for (; i < end && CachedCellRow(sorted_keys[i]) == row; ++i) {
      AppendCachedCell(out, book, sheet, sorted_keys[i]);
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

std::size_t ExternalLinkOrdinals::SheetPosition(const void* ctx, std::string_view path, std::string_view book,
                                                std::string_view sheet) {
  const auto* self = static_cast<const ExternalLinkOrdinals*>(ctx);
  if (self->links_.sheet_position == nullptr) {
    return parser::kUnknownSheet;
  }
  return self->links_.sheet_position(self->links_.ctx, path, book, sheet);
}

std::string BuildExternalLinkXml(const ExternalLinkRecord& rec) {
  const ExternalBook& book = rec.book;
  std::string out;
  out.reserve(512U + book.sheet_names.size() * 48U + book.names.size() * 64U + book.cells.size() * 40U);
  out.append(kXmlDecl);
  AppendExternalLinkOpen(out);
  AppendXmlAttrEscaped(out, TargetRelId(rec));
  out.append("\">");
  if (!rec.absolute_target.empty() || !rec.relative_target.empty()) {
    out.append("<xxl21:alternateUrls>");
    if (!rec.absolute_target.empty()) {
      out.append("<xxl21:absoluteUrl r:id=\"");
      AppendXmlAttrEscaped(out, AbsoluteRelId(rec));
      out.append("\"/>");
    }
    if (!rec.relative_target.empty()) {
      out.append("<xxl21:relativeUrl r:id=\"");
      AppendXmlAttrEscaped(out, RelativeRelId(rec));
      out.append("\"/>");
    }
    out.append("</xxl21:alternateUrls>");
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
        append_xml_attr(out, "refersTo", RefersTo(book, name));
      }
      if (name.scope_sheet != ExternalBook::kNoSheet) {
        append_xml_attr_uint(out, "sheetId", name.scope_sheet);
      }
      out.append("/>");
    }
    out.append("</definedNames>");
  }
  bool any_data = false;
  std::vector<std::uint64_t> sorted_cell_keys;
  bool has_cached_sheet = false;
  for (std::uint32_t s = 0; s < static_cast<std::uint32_t>(book.sheet_names.size()); ++s) {
    if (book.sheet_has_data(s)) {
      has_cached_sheet = true;
      break;
    }
  }
  if (has_cached_sheet && !book.cells.empty()) {
    sorted_cell_keys = book.sorted_cell_keys();
  }
  std::size_t next_key = 0U;
  for (std::uint32_t s = 0; s < static_cast<std::uint32_t>(book.sheet_names.size()); ++s) {
    while (next_key < sorted_cell_keys.size() && CachedCellSheet(sorted_cell_keys[next_key]) < s) {
      ++next_key;
    }
    const std::size_t begin = next_key;
    while (next_key < sorted_cell_keys.size() && CachedCellSheet(sorted_cell_keys[next_key]) == s) {
      ++next_key;
    }
    if (!any_data) {
      out.append("<sheetDataSet>");
      any_data = true;
    }
    // A sheet without cached data writes none of the cells its cache may still hold.
    AppendSheetData(out, book, s, sorted_cell_keys, begin, book.sheet_has_data(s) ? next_key : begin);
  }
  if (any_data) {
    out.append("</sheetDataSet>");
  }
  out.append("</externalBook></externalLink>");
  return out;
}

std::string BuildExternalLinkRels(const ExternalLinkRecord& rec) {
  if (rec.target.empty() && rec.absolute_target.empty() && rec.relative_target.empty()) {
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
      type = rec.target_path_missing ? kRelExternalLinkPathMissing : kRelExternalLinkPath;
      break;
  }
  std::string out;
  out.reserve(384);
  out.append(kXmlDecl);
  BeginRelationships(out);
  // Relative path, then absolute, then the target: the order Excel writes.
  if (!rec.relative_target.empty()) {
    AppendRelationship(out, RelativeRelId(rec), kRelExternalLinkPath, rec.relative_target, /*target_external=*/true,
                       /*escape_target=*/true);
  }
  if (!rec.absolute_target.empty()) {
    AppendRelationship(out, AbsoluteRelId(rec), kRelExternalLinkPath, rec.absolute_target, /*target_external=*/true,
                       /*escape_target=*/true);
  }
  if (!rec.target.empty()) {
    // A missing-path target is stored decoded and re-encoded with only a space and `%` escaped, as Excel does.
    const std::string target =
        rec.target_path_missing ? parser::percent_encode_target(parser::percent_decode(rec.target)) : rec.target;
    AppendRelationship(out, TargetRelId(rec), type, target, rec.target_external, /*escape_target=*/true);
  }
  EndRelationships(out);
  return out;
}

bool NeedsStorageSpelling(const parser::AstNode& root, std::string_view storage) {
  return storage != parser::format_formula(root) || parser::formula_needs_storage_requote(root) ||
         ContainsExternalRef(root);
}

namespace {

bool HiddenBuiltinIsBare(const void* ctx, std::string_view name) {
  const auto& scope = *static_cast<const HiddenBuiltinScope*>(ctx);
  if (scope.host < 0 || !strings::case_insensitive_starts_with(name, parser::kNamePrefix)) {
    return false;
  }
  for (const DefinedName& entry : *scope.names) {
    if (entry.local_sheet_id == scope.host && sheet_names::equal(entry.name, name)) {
      return true;
    }
  }
  return false;
}

}  // namespace

parser::ExternalBookIndexer WithBareHiddenBuiltins(const parser::ExternalBookIndexer& base,
                                                   const HiddenBuiltinScope& scope) {
  parser::ExternalBookIndexer out = base;
  out.bare_self_book = &HiddenBuiltinIsBare;
  out.bare_ctx = &scope;
  return out;
}

}  // namespace io
}  // namespace formulon
