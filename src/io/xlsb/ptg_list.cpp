#include "io/xlsb/ptg_list.h"

#include <algorithm>
#include <string>
#include <string_view>
#include <unordered_set>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "parser/structured_ref_payload.h"
#include "sheet_name.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::size_t kPtgListSize = 14;

// PtgRowType ([MS-XLSB] 2.5.98.73).
constexpr std::uint8_t kRowsData = 0x00;
constexpr std::uint8_t kRowsAll = 0x01;
constexpr std::uint8_t kRowsHeaders = 0x02;
constexpr std::uint8_t kRowsData2 = 0x04;
constexpr std::uint8_t kRowsDataHeaders = 0x06;
constexpr std::uint8_t kRowsTotals = 0x08;
constexpr std::uint8_t kRowsDataTotals = 0x0C;
constexpr std::uint8_t kRowsCurrent = 0x10;

// PtgList flags: columns (bits 0-1), rowType (2-6), the two space bits (7-8), an unused bit (9, written clear,
// ignored on read), type (10-11), invalid (12) and nonresident (13).
constexpr std::uint16_t kColumnsMask = 0x0003U;
constexpr unsigned kRowTypeShift = 2U;
constexpr std::uint16_t kRowTypeMask = 0x1FU;
constexpr unsigned kTypeShift = 10U;
constexpr std::uint16_t kInvalidBit = 0x1000U;
constexpr std::uint16_t kNonresidentBit = 0x2000U;

Error Unsupported(const std::string& message) {
  return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, message, "context=xlsb_ptg_list");
}

/// A column name with the structured-reference escape (`'`) before each of
/// the characters the bracket syntax reserves.
std::string EscapeColumn(std::string_view name) {
  std::string out;
  out.reserve(name.size());
  for (const char c : name) {
    if (c == '[' || c == ']' || c == '#' || c == '\'') {
      out.push_back('\'');
    }
    out.push_back(c);
  }
  return out;
}

const char* SelectorText(std::uint8_t row_type) {
  switch (row_type) {
    case kRowsAll:
      return "#All";
    case kRowsHeaders:
      return "#Headers";
    case kRowsTotals:
      return "#Totals";
    case kRowsCurrent:
      return "#This Row";
    default:
      return nullptr;
  }
}

const TableMetadata* FindTable(const std::vector<TableMetadata>* tables, std::string_view name, std::size_t* index) {
  if (tables == nullptr) {
    return nullptr;
  }
  const std::size_t found = sheet_names::resolve_table(*tables, name);
  if (found == sheet_names::kNotFound) {
    return nullptr;
  }
  *index = found;
  return &(*tables)[found];
}

const XlsbRetainedTable* FindRetainedTable(const SheetRangeTable& ranges, std::string_view name) {
  const std::size_t found = sheet_names::resolve_table(ranges.retained_tables, name);
  return found == sheet_names::kNotFound ? nullptr : &ranges.retained_tables[found];
}

std::optional<std::uint8_t> RowTypeFor(std::uint8_t selectors) {
  using Spec = parser::StructuredRefSpecifiers;
  switch (selectors) {
    case 0:
    case Spec::kData:
      return kRowsData;
    case Spec::kAll:
      return kRowsAll;
    case Spec::kHeaders:
      return kRowsHeaders;
    case Spec::kHeaders | Spec::kData:
      return kRowsDataHeaders;
    case Spec::kTotals:
      return kRowsTotals;
    case Spec::kData | Spec::kTotals:
      return kRowsDataTotals;
    case Spec::kThisRow:
      return kRowsCurrent;
    default:
      return std::nullopt;
  }
}

std::optional<std::uint16_t> ColumnIndex(const TableMetadata& table, const std::string& name) {
  const std::size_t found = sheet_names::resolve_named(table.columns, name);
  if (found == sheet_names::kNotFound) {
    return std::nullopt;
  }
  return static_cast<std::uint16_t>(found);
}

}  // namespace

bool structured_ref_is_single_cell(const parser::AstNode& node) {
  if (node.as_structured_ref_modifier() == parser::StructuredRefModifier::At) {
    return true;
  }
  using Spec = parser::StructuredRefSpecifiers;
  if (node.as_structured_ref_modifier() != parser::StructuredRefModifier::None) {
    return false;
  }
  const auto payload = parser::parse_structured_ref_payload_text(node.as_structured_ref_column());
  if (!payload || payload->columns.size() != 1U) {
    return false;
  }
  return payload->specifiers == Spec::kThisRow || payload->specifiers == Spec::kHeaders ||
         payload->specifiers == Spec::kTotals;
}

Expected<parser::AstNode*, Error> decode_ptg_list(ByteSpan& cursor, Arena& arena,
                                                  const std::vector<TableMetadata>* tables) {
  if (cursor.size < kPtgListSize) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "PtgList payload truncated", "context=xlsb_ptg_list");
  }
  cursor.data += 2;
  cursor.size -= 2;
  ASSIGN_OR_RETURN(const std::uint16_t ixti, read_u16(cursor));
  ASSIGN_OR_RETURN(const std::uint16_t flags, read_u16(cursor));
  ASSIGN_OR_RETURN(const std::uint32_t list_index, read_u32(cursor));
  ASSIGN_OR_RETURN(const std::uint16_t col_first, read_u16(cursor));
  ASSIGN_OR_RETURN(const std::uint16_t col_last, read_u16(cursor));
  (void)ixti;
  if ((flags & (kInvalidBit | kNonresidentBit)) != 0U) {
    return Unsupported("PtgList is invalid or names another workbook's table");
  }
  // A table read from the file carries the id its PtgLists name.
  const TableMetadata* table = nullptr;
  if (tables != nullptr) {
    for (const TableMetadata& candidate : *tables) {
      if (candidate.id == list_index) {
        table = &candidate;
        break;
      }
    }
  }
  if (table == nullptr) {
    return Unsupported("PtgList names no table of this workbook: " + std::to_string(list_index));
  }
  const std::uint16_t columns = flags & kColumnsMask;
  const std::uint8_t row_type = static_cast<std::uint8_t>((flags >> kRowTypeShift) & kRowTypeMask);
  if (columns == 3U || (columns != 0U && (col_first > col_last || col_last >= table->columns.size()))) {
    return Unsupported("PtgList column slice is outside its table");
  }
  std::vector<std::string> specs;
  switch (row_type) {
    case kRowsData:
    case kRowsData2:
      break;
    case kRowsDataHeaders:
      specs = {"#Headers", "#Data"};
      break;
    case kRowsDataTotals:
      specs = {"#Data", "#Totals"};
      break;
    default:
      if (const char* text = SelectorText(row_type); text != nullptr) {
        specs = {text};
        break;
      }
      return Unsupported("PtgList has an unknown row type");
  }
  std::string cols;
  if (columns == 1U || (columns == 2U && col_first == col_last)) {
    cols = "[" + EscapeColumn(table->columns[col_first].name) + "]";
  } else if (columns == 2U) {
    cols =
        "[" + EscapeColumn(table->columns[col_first].name) + "]:[" + EscapeColumn(table->columns[col_last].name) + "]";
  }
  std::string payload;
  if (specs.empty()) {
    // `Table1[Col]` spells a lone column without its brackets.
    payload =
        columns == 1U || (columns == 2U && col_first == col_last) ? EscapeColumn(table->columns[col_first].name) : cols;
  } else if (specs.size() == 1U && cols.empty()) {
    payload = specs.front();
  } else {
    for (const std::string& spec : specs) {
      payload.append(payload.empty() ? "[" : ",[").append(spec).append("]");
    }
    if (!cols.empty()) {
      payload.append(",").append(cols);
    }
  }
  parser::AstNode* node =
      parser::make_structured_ref(arena, table->display_name, payload, parser::StructuredRefModifier::None);
  if (node == nullptr) {
    return make_error(FormulonErrorCode::kOutOfMemory, "arena exhausted (PtgList)", "context=xlsb_ptg_list");
  }
  return node;
}

std::uint32_t xlsb_table_id(const std::vector<TableMetadata>& tables, std::size_t index,
                            const std::unordered_set<std::uint32_t>& reserved) {
  const std::uint32_t own = tables[index].id;
  if (own != 0U && own != 0xFFFFFFFFU && reserved.count(own) == 0U &&
      std::none_of(tables.begin(), tables.begin() + static_cast<std::ptrdiff_t>(index),
                   [own](const TableMetadata& t) { return t.id == own; })) {
    return own;
  }
  std::unordered_set<std::uint32_t> taken;
  std::unordered_set<std::uint32_t> held = reserved;
  for (const TableMetadata& t : tables) {
    held.insert(t.id);
  }
  std::uint32_t next = 1;
  for (std::size_t i = 0; i <= index && i < tables.size(); ++i) {
    std::uint32_t id = tables[i].id;
    if (id == 0U || id == 0xFFFFFFFFU || reserved.count(id) != 0U || !taken.insert(id).second) {
      while (held.count(next) != 0U || taken.count(next) != 0U) {
        ++next;
      }
      id = next;
      taken.insert(id);
    }
    if (i == index) {
      return id;
    }
  }
  return 0U;
}

std::optional<std::int32_t> structured_ref_sheet(const parser::AstNode& node, const SheetRangeTable& ranges) {
  std::size_t index = 0;
  if (const TableMetadata* table = FindTable(ranges.tables, node.as_structured_ref_table(), &index)) {
    return static_cast<std::int32_t>(table->sheet_index);
  }
  if (const XlsbRetainedTable* retained = FindRetainedTable(ranges, node.as_structured_ref_table());
      retained != nullptr && retained->sheet_index) {
    return static_cast<std::int32_t>(*retained->sheet_index);
  }
  return std::nullopt;
}

Expected<void, Error> encode_ptg_list(std::vector<std::uint8_t>& out, const parser::AstNode& node, std::uint8_t cls,
                                      const SheetRangeTable& ranges) {
  std::size_t index = 0;
  const TableMetadata* table = FindTable(ranges.tables, node.as_structured_ref_table(), &index);
  // A table kept verbatim resolves through its retained identity and columns.
  TableMetadata retained_view;
  std::uint32_t table_id = 0;
  if (table != nullptr) {
    table_id = xlsb_table_id(*ranges.tables, index, ranges.reserved_table_ids);
  } else if (const XlsbRetainedTable* retained = FindRetainedTable(ranges, node.as_structured_ref_table());
             retained != nullptr && retained->sheet_index) {
    retained_view.sheet_index = *retained->sheet_index;
    for (const std::string& caption : retained->columns) {
      retained_view.columns.emplace_back();
      retained_view.columns.back().name = caption;
    }
    table = &retained_view;
    table_id = retained->id;
  }
  if (table == nullptr) {
    return Unsupported("structured reference names no table of this workbook");
  }
  std::optional<parser::StructuredRefPayload> parsed;
  if (node.as_structured_ref_modifier() == parser::StructuredRefModifier::None) {
    parsed = parser::parse_structured_ref_payload_text(node.as_structured_ref_column());
  }
  if (!parsed) {
    return Unsupported("structured reference payload has no PtgList form");
  }
  const parser::StructuredRefPayload& payload = *parsed;
  const std::optional<std::uint8_t> row_type = RowTypeFor(payload.specifiers);
  if (!row_type) {
    return Unsupported("structured reference selectors have no PtgRowType");
  }
  std::uint16_t col_first = 0;
  std::uint16_t col_last = 0;
  std::uint16_t columns = 0;
  if (!payload.columns.empty()) {
    const std::optional<std::uint16_t> first = ColumnIndex(*table, payload.columns.front());
    const std::optional<std::uint16_t> last = ColumnIndex(*table, payload.columns.back());
    if (!first || !last) {
      return Unsupported("structured reference names an unknown column");
    }
    col_first = std::min(*first, *last);
    col_last = std::max(*first, *last);
    columns = payload.columns.size() == 1U ? 1U : 2U;
  }
  const XtiEntry want{0U, static_cast<std::int32_t>(table->sheet_index), static_cast<std::int32_t>(table->sheet_index)};
  const auto xti = std::find(ranges.xti.begin(), ranges.xti.end(), want);
  if (xti == ranges.xti.end()) {
    return Unsupported("structured reference's table sheet has no ExternSheet entry");
  }
  const std::uint16_t type = cls == 0x20U ? 0U : cls == 0x60U ? 2U : 1U;
  std::uint16_t flags = columns;
  flags |= static_cast<std::uint16_t>(*row_type << kRowTypeShift);
  flags |= static_cast<std::uint16_t>(type << kTypeShift);
  emit_u8(out, kPtgList);
  emit_u8(out, kEptgList);
  emit_u16(out, static_cast<std::uint16_t>(xti - ranges.xti.begin()));
  emit_u16(out, flags);
  emit_u32(out, table_id);
  emit_u16(out, col_first);
  emit_u16(out, col_last);
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
