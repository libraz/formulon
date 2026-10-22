#include "io/xlsb/ptg_list.h"

#include <algorithm>
#include <string>
#include <string_view>
#include <unordered_set>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
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

// Selector bits of a parsed bracket payload.
constexpr std::uint8_t kSelAll = 0x01;
constexpr std::uint8_t kSelData = 0x02;
constexpr std::uint8_t kSelHeaders = 0x04;
constexpr std::uint8_t kSelTotals = 0x08;
constexpr std::uint8_t kSelThisRow = 0x10;

// PtgList flag fields.
constexpr std::uint16_t kColumnsMask = 0x0003U;
constexpr unsigned kRowTypeShift = 2U;
constexpr std::uint16_t kRowTypeMask = 0x1FU;
// Excel sets the "unused" bit on the PtgLists it writes ([MS-XLSB] 3.3.11).
constexpr std::uint16_t kUnusedBit = 0x0200U;
constexpr unsigned kTypeShift = 10U;
constexpr std::uint16_t kInvalidBit = 0x1000U;
constexpr std::uint16_t kNonresidentBit = 0x2000U;

Error Unsupported(const std::string& message) {
  return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, message, "context=xlsb_ptg_list");
}

bool EqualsIgnoreCase(std::string_view a, std::string_view b) {
  return a.size() == b.size() && std::equal(a.begin(), a.end(), b.begin(), [](char x, char y) {
           auto lower = [](char c) { return c >= 'A' && c <= 'Z' ? static_cast<char>(c - 'A' + 'a') : c; };
           return lower(x) == lower(y);
         });
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
  for (std::size_t i = 0; i < tables->size(); ++i) {
    const TableMetadata& t = (*tables)[i];
    if (EqualsIgnoreCase(t.display_name, name) || EqualsIgnoreCase(t.name, name)) {
      *index = i;
      return &t;
    }
  }
  return nullptr;
}

/// The parsed bracket payload: selector bits and the column slice.
struct Payload {
  std::uint8_t selectors = 0;
  std::vector<std::string> columns;  // unescaped; one, or two for a range
};

/// Reads one `[...]` item starting at `text[*pos] == '['`; unescapes `'`.
/// `selector` is set when the item opens with an unescaped `#`.
bool ReadBracket(std::string_view text, std::size_t* pos, std::string* out, bool* selector) {
  std::size_t i = *pos + 1U;
  out->clear();
  *selector = i < text.size() && text[i] == '#';
  while (i < text.size()) {
    const char c = text[i];
    if (c == '\'' && i + 1U < text.size()) {
      out->push_back(text[i + 1U]);
      i += 2U;
      continue;
    }
    if (c == ']') {
      *pos = i + 1U;
      return true;
    }
    if (c == '[') {
      return false;
    }
    out->push_back(c);
    ++i;
  }
  return false;
}

std::string Unescape(std::string_view text) {
  std::string out;
  for (std::size_t i = 0; i < text.size(); ++i) {
    if (text[i] == '\'' && i + 1U < text.size()) {
      ++i;
    }
    out.push_back(text[i]);
  }
  return out;
}

bool AddSelector(std::string_view item, std::uint8_t* selectors) {
  struct Name {
    std::string_view text;
    std::uint8_t bit;
  };
  constexpr Name kNames[] = {{"#All", kSelAll},
                             {"#Data", kSelData},
                             {"#Headers", kSelHeaders},
                             {"#Totals", kSelTotals},
                             {"#This Row", kSelThisRow}};
  for (const Name& name : kNames) {
    if (EqualsIgnoreCase(item, name.text)) {
      *selectors |= name.bit;
      return true;
    }
  }
  return false;
}

std::string_view Trim(std::string_view s) {
  while (!s.empty() && s.front() == ' ') {
    s.remove_prefix(1);
  }
  while (!s.empty() && s.back() == ' ') {
    s.remove_suffix(1);
  }
  return s;
}

/// Parses `[...]` items: selectors separated by `,`, then at most one column
/// or `[A]:[B]` range.
bool ParseBracketList(std::string_view text, Payload* out) {
  std::size_t pos = 0;
  bool range_pending = false;
  while (pos < text.size()) {
    if (text[pos] == ' ' || text[pos] == ',') {
      ++pos;
      continue;
    }
    if (text[pos] == ':') {
      if (out->columns.size() != 1U || range_pending) {
        return false;
      }
      range_pending = true;
      ++pos;
      continue;
    }
    if (text[pos] != '[') {
      return false;
    }
    std::string item;
    bool selector = false;
    if (!ReadBracket(text, &pos, &item, &selector)) {
      return false;
    }
    if (selector) {
      if (!out->columns.empty() || !AddSelector(Trim(item), &out->selectors)) {
        return false;
      }
      continue;
    }
    if (!out->columns.empty() && !range_pending) {
      return false;
    }
    out->columns.push_back(item);
    range_pending = false;
  }
  return !range_pending;
}

bool ParsePayload(std::string_view text, Payload* out) {
  text = Trim(text);
  if (text.empty()) {
    return true;
  }
  if (text.front() == '@') {
    out->selectors |= kSelThisRow;
    text = Trim(text.substr(1));
    if (text.empty()) {
      return true;
    }
    if (text.front() != '[') {
      out->columns.push_back(Unescape(text));
      return true;
    }
    return ParseBracketList(text, out) && (out->selectors & ~kSelThisRow) == 0U;
  }
  if (text.front() == '[') {
    return ParseBracketList(text, out);
  }
  if (text.front() == '#') {
    return AddSelector(text, &out->selectors);
  }
  out->columns.push_back(Unescape(text));
  return true;
}

std::optional<std::uint8_t> RowTypeFor(std::uint8_t selectors) {
  switch (selectors) {
    case 0:
    case kSelData:
      return kRowsData;
    case kSelAll:
      return kRowsAll;
    case kSelHeaders:
      return kRowsHeaders;
    case kSelHeaders | kSelData:
      return kRowsDataHeaders;
    case kSelTotals:
      return kRowsTotals;
    case kSelData | kSelTotals:
      return kRowsDataTotals;
    case kSelThisRow:
      return kRowsCurrent;
    default:
      return std::nullopt;
  }
}

std::optional<std::uint16_t> ColumnIndex(const TableMetadata& table, const std::string& name) {
  for (std::size_t i = 0; i < table.columns.size(); ++i) {
    if (EqualsIgnoreCase(table.columns[i].name, name)) {
      return static_cast<std::uint16_t>(i);
    }
  }
  return std::nullopt;
}

}  // namespace

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

std::optional<std::int32_t> structured_ref_sheet(const parser::AstNode& node,
                                                 const std::vector<TableMetadata>* tables) {
  std::size_t index = 0;
  const TableMetadata* table = FindTable(tables, node.as_structured_ref_table(), &index);
  if (table == nullptr) {
    return std::nullopt;
  }
  return static_cast<std::int32_t>(table->sheet_index);
}

Expected<void, Error> encode_ptg_list(std::vector<std::uint8_t>& out, const parser::AstNode& node, std::uint8_t cls,
                                      const SheetRangeTable& ranges) {
  std::size_t index = 0;
  const TableMetadata* table = FindTable(ranges.tables, node.as_structured_ref_table(), &index);
  if (table == nullptr) {
    return Unsupported("structured reference names no table of this workbook");
  }
  Payload payload;
  if (node.as_structured_ref_modifier() != parser::StructuredRefModifier::None ||
      !ParsePayload(node.as_structured_ref_column(), &payload)) {
    return Unsupported("structured reference payload has no PtgList form");
  }
  const std::optional<std::uint8_t> row_type = RowTypeFor(payload.selectors);
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
  flags |= kUnusedBit;
  flags |= static_cast<std::uint16_t>(type << kTypeShift);
  emit_u8(out, kPtgList);
  emit_u8(out, kEptgList);
  emit_u16(out, static_cast<std::uint16_t>(xti - ranges.xti.begin()));
  emit_u16(out, flags);
  emit_u32(out, xlsb_table_id(*ranges.tables, index, ranges.reserved_table_ids));
  emit_u16(out, col_first);
  emit_u16(out, col_last);
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
