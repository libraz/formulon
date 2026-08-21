
#include "io/ooxml/external_link_reader.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/a1_ref.h"
#include "io/external_book.h"
#include "io/external_links.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/workbook_rels_reader.h"
#include "io/ooxml_defs.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "value.h"

namespace formulon {
namespace io {
namespace ooxml {
namespace {

/// Maps an Excel error display name (`"#N/A"`) to its `ErrorCode`,
/// scanning the same `kErrorTable` the writer formats from so the two
/// directions cannot drift. Unknown spellings become `#N/A`, which is
/// what a cached cell Excel could not classify already shows.
ErrorCode ErrorFromDisplay(std::string_view text) {
  for (std::size_t i = 0; i < std::size(kErrorTable); ++i) {
    if (text == kErrorTable[i].display_name) {
      return static_cast<ErrorCode>(i);
    }
  }
  return ErrorCode::NA;
}

/// Parses the `refersTo` of an external `<definedName>` into the
/// rectangle it names.
///
/// The grammar accepted here is the one Excel writes for a name that
/// points at cells: `='Sheet'!$A$1` or `='Sheet'!$A$1:$B$2`, with the
/// sheet quoted or not and the `$` anchors optional. Anything else — a
/// constant, an expression, a name chained onto a further workbook — is
/// rejected, because resolving it would mean evaluating a formula in the
/// supporting workbook's namespace against cells this cache does not
/// hold. `out->resolvable` stays false and the reference reads `#REF!`.
bool ParseExternalRefersTo(std::string_view text, const ExternalBook& book, ExternalBookName* out) {
  if (!text.empty() && text.front() == '=') {
    text.remove_prefix(1);
  }
  // Sheet qualifier, up to the last `!` so a quoted name containing one
  // cannot truncate the reference.
  const std::size_t bang = text.rfind('!');
  if (bang == std::string_view::npos) {
    return false;
  }
  std::string_view sheet = text.substr(0, bang);
  std::string_view tail = text.substr(bang + 1);
  if (sheet.size() >= 2U && sheet.front() == '\'' && sheet.back() == '\'') {
    sheet = sheet.substr(1, sheet.size() - 2U);
  }
  const std::uint32_t sheet_index = book.sheet_index(sheet);
  if (sheet_index == ExternalBook::kNoSheet) {
    return false;
  }

  // `$A$1` / `$A$1:$B$2`. The anchors carry no meaning for a reference
  // into another workbook (nothing here is ever shifted), so they are
  // stripped rather than recorded.
  std::string plain;
  plain.reserve(tail.size());
  for (const char c : tail) {
    if (c != '$') {
      plain.push_back(c);
    }
  }
  const std::size_t colon = plain.find(':');
  const std::string_view first =
      colon == std::string::npos ? std::string_view(plain) : std::string_view(plain).substr(0, colon);
  const std::string_view last =
      colon == std::string::npos ? std::string_view(plain) : std::string_view(plain).substr(colon + 1);
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t row_end = 0;
  std::uint32_t col_end = 0;
  if (!parse_a1_ref(first, &row, &col) || !parse_a1_ref(last, &row_end, &col_end)) {
    return false;
  }
  out->sheet = sheet_index;
  out->row = row;
  out->col = col;
  out->row_end = row_end;
  out->col_end = col_end;
  out->is_range = colon != std::string::npos;
  out->resolvable = true;
  return true;
}

/// Decodes one `<externalBook>` into the shared cache model.
///
/// Sheet names come first because `<definedNames>` resolves its sheet
/// qualifier against them, and `<sheetDataSet>` keys its `sheetId` by
/// position in the same list.
void DecodeExternalBook(const pugi::xml_node& book_node, ExternalBook* out) {
  for (pugi::xml_node sheet = book_node.child("sheetNames").child("sheetName"); sheet;
       sheet = sheet.next_sibling("sheetName")) {
    out->sheet_names.emplace_back(sheet.attribute("val").value());
  }
  for (pugi::xml_node name = book_node.child("definedNames").child("definedName"); name;
       name = name.next_sibling("definedName")) {
    ExternalBookName entry;
    entry.name = name.attribute("name").value();
    if (entry.name.empty()) {
      continue;
    }
    // A name that fails to resolve is still recorded: it is the
    // difference between `#REF!` (the name exists, its target is not
    // reachable) and `#NAME?` (no such name), and Excel distinguishes
    // the two.
    ParseExternalRefersTo(name.attribute("refersTo").value(), *out, &entry);
    out->names.push_back(std::move(entry));
  }
  for (pugi::xml_node data = book_node.child("sheetDataSet").child("sheetData"); data;
       data = data.next_sibling("sheetData")) {
    const std::uint32_t sheet_id = data.attribute("sheetId").as_uint();
    if (sheet_id >= out->sheet_names.size()) {
      continue;
    }
    for (pugi::xml_node row = data.child("row"); row; row = row.next_sibling("row")) {
      for (pugi::xml_node cell = row.child("cell"); cell; cell = cell.next_sibling("cell")) {
        std::uint32_t cell_row = 0;
        std::uint32_t cell_col = 0;
        if (!parse_a1_ref(cell.attribute("r").value(), &cell_row, &cell_col)) {
          continue;
        }
        const std::string_view type = cell.attribute("t").value();
        const std::string_view raw = cell.child("v").text().get();
        ExternalCell out_cell;
        if (type == "str") {
          out_cell.value = Value::text({});
          out_cell.text = raw;
        } else if (type == "b") {
          out_cell.value = Value::boolean(raw == "1");
        } else if (type == "e") {
          out_cell.value = Value::error(ErrorFromDisplay(raw));
        } else {
          // No `t` (or an unrecognised one) is the numeric default.
          out_cell.value = Value::number(cell.child("v").text().as_double());
        }
        out->cells.emplace(ExternalBook::cell_key(sheet_id, cell_row, cell_col), std::move(out_cell));
      }
    }
  }
}

}  // namespace

Expected<ExternalLinkLoadResult, Error> load_external_links(const ZipReader& zip, const pugi::xml_node& wb_root,
                                                            const WorkbookRels& wb_rels) {
  ExternalLinkLoadResult out;
  pugi::xml_node refs_node = wb_root.child("externalReferences");
  if (!refs_node) {
    return out;
  }
  std::uint32_t index = 1;
  for (pugi::xml_node ref = refs_node.child("externalReference"); ref;
       ref = ref.next_sibling("externalReference"), ++index) {
    std::string rid = relationship_ref_id(ref);
    if (rid.empty()) {
      continue;
    }
    auto it = wb_rels.external_link_paths_by_rid.find(rid);
    if (it == wb_rels.external_link_paths_by_rid.end()) {
      continue;
    }
    ExternalLinkRecord rec;
    rec.index = index;
    rec.rel_id = std::move(rid);
    rec.part_path = it->second;
    rec.kind = ExternalLinkRecord::Kind::kUnknown;

    // Body part — detect kind and capture the inner r:id reference.
    if (zip.has_entry(rec.part_path)) {
      auto body_or = zip.read_entry(rec.part_path);
      if (!body_or) {
        return body_or.error();
      }
      const std::vector<std::uint8_t>& body_bytes = body_or.value();
      pugi::xml_document body_doc;
      pugi::xml_parse_result body_parse =
          body_doc.load_buffer(body_bytes.data(), body_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
      if (body_parse) {
        pugi::xml_node link_root = body_doc.child("externalLink");
        if (link_root) {
          if (pugi::xml_node book = link_root.child("externalBook"); book) {
            rec.kind = ExternalLinkRecord::Kind::kExternalBook;
            rec.body_rel_id = relationship_ref_id(book);
            DecodeExternalBook(book, &rec.book);
          } else if (pugi::xml_node ole = link_root.child("oleLink"); ole) {
            rec.kind = ExternalLinkRecord::Kind::kOleLink;
            rec.body_rel_id = relationship_ref_id(ole);
          } else if (link_root.child("ddeLink")) {
            rec.kind = ExternalLinkRecord::Kind::kDdeLink;
            // ddeLink carries its connection metadata inline; no inner r:id.
          }
        }
      }
    }

    // Per-link rels — capture target URL + target_mode for round-trip.
    const std::string body_rels_path = rels_path_for_part(rec.part_path);
    if (zip.has_entry(body_rels_path)) {
      auto rels_or = zip.read_entry(body_rels_path);
      if (!rels_or) {
        return rels_or.error();
      }
      const std::vector<std::uint8_t>& rels_bytes = rels_or.value();
      pugi::xml_document rels_doc;
      pugi::xml_parse_result rels_parse =
          rels_doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
      if (rels_parse) {
        pugi::xml_node rels_root = rels_doc.child("Relationships");
        if (rels_root) {
          // Pick the relationship whose Id matches the body's inner
          // r:id when available; otherwise take the first link-typed
          // relationship as a best-effort fallback.
          for (pugi::xml_node rel = rels_root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
            const std::string_view type = rel.attribute("Type").value();
            if (type != kRelExternalLinkPath && type != kRelOleLink && type != kRelDdeLink) {
              continue;
            }
            const std::string_view rel_id = rel.attribute("Id").value();
            const bool id_match = !rec.body_rel_id.empty() && rel_id == rec.body_rel_id;
            if (rec.target.empty() || id_match) {
              rec.target = rel.attribute("Target").value();
              // `TargetMode` is optional and defaults to `Internal`, so an
              // absent attribute means an in-package target. Reading it as
              // external instead would let the next save emit
              // `TargetMode="External"` for a relationship the source file
              // never declared that way. Matches every other TargetMode
              // read site in the reader family.
              rec.target_external = rel.attribute("TargetMode").value() == std::string_view("External");
              if (rec.body_rel_id.empty()) {
                rec.body_rel_id = rel_id;
              }
              if (id_match) {
                break;
              }
            }
          }
        }
      }
      out.consumed_rels_paths.push_back(body_rels_path);
    }
    out.records.push_back(std::move(rec));
  }
  return out;
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
