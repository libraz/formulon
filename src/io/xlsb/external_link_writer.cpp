//
// Implementation of the XLSB external-link part writer. See
// `io/xlsb/external_link_writer.h`.

#include "io/xlsb/external_link_writer.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "external_book.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// Records of the part `XlsbRecordType` does not list, because the reader
// skips them. Payloads below are the ones Excel 365 writes.
constexpr std::uint16_t kBrtFrtBegin = 37;
constexpr std::uint16_t kBrtFrtEnd = 38;
constexpr std::uint16_t kBrtAlternateUrls = 5108;
constexpr std::uint16_t kBrtExternNameBits = 586;

constexpr std::uint8_t kFrtVersionPayload[] = {0x01, 0x00, 0x06, 0x11, 0x00, 0x80};
constexpr std::uint8_t kNameBitsPayload[] = {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};

constexpr std::uint8_t kPtgRef3d = 0x3A;
constexpr std::uint8_t kPtgArea3d = 0x3B;
constexpr std::uint8_t kPtgErr = 0x1C;

struct LinkRelIds {
  std::string target;
  std::string absolute;  // Empty without an absolute URL.
};

/// The rel ids the part names its target and absolute URL by: the loaded
/// ones, else `rId1` / `rId2` as Excel numbers them.
LinkRelIds RelIdsOf(const ExternalLinkRecord& link) {
  LinkRelIds ids;
  ids.target = link.body_rel_id.empty() ? std::string("rId1") : link.body_rel_id;
  if (!link.absolute_target.empty()) {
    ids.absolute = link.absolute_rel_id;
    if (ids.absolute.empty() || ids.absolute == ids.target) {
      ids.absolute = ids.target == "rId2" ? "rId1" : "rId2";
    }
  }
  return ids;
}

/// A name's stored target: an absolute cell or rectangle on one sheet of the
/// supporting workbook, in the part's 16-bit-row layout, `#REF!` when the
/// cache cannot resolve it, and no body (cce 0) for a name the book does not
/// declare.
void EmitNameFormula(std::vector<std::uint8_t>& p, const ExternalBookName* name) {
  constexpr std::uint32_t kMaxField = 0xFFFFU;
  if (name == nullptr || !name->exists) {
    emit_u32(p, 0U);  // cce
    return;
  }
  const bool encodable =
      name->resolvable && name->sheet <= kMaxField && name->row_end <= kMaxField && name->col_end <= kMaxField;
  if (!encodable) {
    emit_u32(p, 2U);  // cce
    emit_u8(p, kPtgErr);
    emit_u8(p, static_cast<std::uint8_t>(ooxml_code(ErrorCode::Ref)));
    return;
  }
  const auto field = [&p](std::uint32_t v) { emit_u16(p, static_cast<std::uint16_t>(v)); };
  if (!name->is_range) {
    emit_u32(p, 9U);  // cce
    emit_u8(p, kPtgRef3d);
    field(name->sheet);
    field(name->sheet);
    field(name->row);
    field(name->col);
    return;
  }
  emit_u32(p, 13U);  // cce
  emit_u8(p, kPtgArea3d);
  field(name->sheet);
  field(name->sheet);
  field(name->row);
  field(name->row_end);
  field(name->col);
  field(name->col_end);
}

/// The cached name `tables.names[i]` stands for: the `i`-th cache entry, or
/// none for a name appended because a formula names it (the book declares
/// no such name).
const ExternalBookName* CachedName(const ExternalBook& book, std::size_t i) {
  return i < book.names.size() ? &book.names[i] : nullptr;
}

/// One cached sheet: its table, then a row header before each run of cells
/// sharing a row. `[first, last)` are the sheet's cache keys in ascending order.
void EmitSheetTable(std::vector<std::uint8_t>& out, const ExternalBook& book, std::uint32_t sheet,
                    const std::uint64_t* first, const std::uint64_t* last) {
  std::vector<std::uint8_t> p;
  emit_u32(p, sheet);
  emit_u8(p, 0);
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginExternTable), p);
  const std::uint64_t base = ExternalBook::cell_key(sheet, 0, 0);
  const std::uint64_t row_unit = ExternalBook::cell_key(0, 1, 0);
  bool row_open = false;
  std::uint32_t current_row = 0;
  for (const std::uint64_t* it = first; it != last; ++it) {
    const std::uint64_t key = *it;
    const ExternalCell& cell = book.cells.at(key);
    const Value value = cell.resolved();
    if (value.is_blank()) {
      continue;
    }
    const auto row = static_cast<std::uint32_t>((key - base) / row_unit);
    const auto col = static_cast<std::uint32_t>((key - base) % row_unit);
    if (!row_open || row != current_row) {
      p.clear();
      emit_u32(p, row);
      emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtExternRowHdr), p);
      row_open = true;
      current_row = row;
    }
    p.clear();
    emit_u32(p, col);
    XlsbRecordType type = XlsbRecordType::BrtExternCellReal;
    if (value.is_number()) {
      emit_double(p, value.as_number());
    } else if (value.is_boolean()) {
      type = XlsbRecordType::BrtExternCellBool;
      emit_u8(p, value.as_boolean() ? 1U : 0U);
    } else if (value.is_error()) {
      type = XlsbRecordType::BrtExternCellError;
      emit_u8(p, static_cast<std::uint8_t>(ooxml_code(value.as_error())));
    } else if (value.is_text()) {
      type = XlsbRecordType::BrtExternCellString;
      emit_xlwidestring(p, cell.text);
    } else {
      continue;  // No other kind is cached.
    }
    emit_record(out, static_cast<std::uint16_t>(type), p);
  }
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtEndExternTable), ByteSpan{});
}

}  // namespace

std::vector<const ExternalLinkRecord*> written_external_links(const Workbook& wb) {
  // An OLE or DDE link is no supporting workbook; a formula naming one
  // keeps its cached value.
  return wb.external_links_by_index(/*skip_ole_dde=*/true);
}

std::string external_link_part_path(std::size_t position) {
  return "xl/externalLinks/externalLink" + std::to_string(position) + ".bin";
}

std::vector<std::uint8_t> build_external_link_bin(const ExternalLinkRecord& link, const XlsbLinkTables& tables) {
  const LinkRelIds ids = RelIdsOf(link);
  const ExternalBook& book = link.book;
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> p;

  emit_u16(p, 0);  // sbt: a workbook
  emit_xlnullablewidestring(p, std::string_view(ids.target));
  emit_xlnullablewidestring(p, std::nullopt);
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginExternalBook), p);

  if (!ids.absolute.empty()) {
    emit_record(out, kBrtFrtBegin, ByteSpan{kFrtVersionPayload, sizeof(kFrtVersionPayload)});
    p.clear();
    emit_u32(p, 0);
    emit_u32(p, 0);
    emit_xlwidestring(p, ids.absolute);
    emit_u32(p, 0);
    emit_record(out, kBrtAlternateUrls, p);
    emit_record(out, kBrtFrtEnd, ByteSpan{});
  }

  p.clear();
  emit_u32(p, static_cast<std::uint32_t>(tables.sheet_names.size()));
  for (const std::string& sheet : tables.sheet_names) {
    emit_xlwidestring(p, sheet);
  }
  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtSupTabs), p);

  for (std::size_t i = 0; i < tables.names.size(); ++i) {
    p.clear();
    emit_xlwidestring(p, tables.names[i]);
    emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtExternNameStart), p);
    p.clear();
    EmitNameFormula(p, CachedName(book, i));
    emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtExternNameFmla), p);
    emit_record(out, kBrtExternNameBits, ByteSpan{kNameBitsPayload, sizeof(kNameBitsPayload)});
    emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtExternNameEnd), ByteSpan{});
  }

  const std::vector<std::uint64_t> keys = book.sorted_cell_keys();
  for (std::uint32_t sheet = 0; sheet < tables.sheet_names.size(); ++sheet) {
    if (!book.sheet_has_data(sheet)) {
      continue;
    }
    const auto lo = std::lower_bound(keys.begin(), keys.end(), ExternalBook::cell_key(sheet, 0, 0));
    const auto hi = std::lower_bound(lo, keys.end(), ExternalBook::cell_key(sheet + 1U, 0, 0));
    EmitSheetTable(out, book, sheet, keys.data() + (lo - keys.begin()), keys.data() + (hi - keys.begin()));
  }

  emit_record(out, static_cast<std::uint16_t>(XlsbRecordType::BrtEndExternalBook), ByteSpan{});
  return out;
}

std::string build_external_link_rels(const ExternalLinkRecord& link) {
  const LinkRelIds ids = RelIdsOf(link);
  std::string out;
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  // Excel lists the absolute URL first.
  if (!ids.absolute.empty()) {
    AppendRelationship(out, ids.absolute, kRelExternalLinkPath, link.absolute_target, /*target_external=*/true,
                       /*escape_target=*/true);
  }
  AppendRelationship(out, ids.target, kRelExternalLinkPath, link.target, link.target_external,
                     /*escape_target=*/true);
  out.append("</Relationships>\n");
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
