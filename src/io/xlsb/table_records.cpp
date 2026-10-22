#include "io/xlsb/table_records.h"

#include <algorithm>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "auto_filter.h"
#include "io/auto_filter_xml.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/auto_filter_records.h"
#include "io/xlsb/cf_x14_link.h"
#include "io/xlsb/ptg_list.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "parser/ast_format.h"
#include "parser/parser.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/structured_log.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

enum : std::uint16_t {
  kBrtFRTBegin = 35,
  kBrtFRTEnd = 36,
  kBrtACBegin = 37,
  kBrtACEnd = 38,
  kBrtBeginAFilter = 161,
  kBrtBeginList = 343,
  kBrtEndList = 344,
  kBrtBeginListCols = 345,
  kBrtEndListCols = 346,
  kBrtBeginListCol = 347,
  kBrtEndListCol = 348,
  kBrtListCCFmla = 351,
  kBrtListTrFmla = 352,
  kBrtTableStyleClient = 513,
  kBrtBeginSortState = 530,
  kBrtBeginListParts = 660,
  kBrtListPart = 661,
  kBrtEndListParts = 662,
  kBrtList14 = 1111,
  kBrtUid = 3072,
};

// ACUID ([MS-XLSB] 2.4.2 BrtACBegin, 2.4.857 BrtUid): the revision uid OOXML
// spells `xr:uid` / `xr3:uid`. The BrtACBegin payload is the one Excel writes.
constexpr std::uint8_t kAcBeginPayload[] = {0x01, 0x00, 0x00, 0x10, 0x00, 0x80};
constexpr std::string_view kMcNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";
constexpr std::string_view kXrNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
constexpr std::string_view kXr3Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3";

constexpr std::uint32_t kNoDxf = 0xFFFFFFFFU;

// BrtBeginList flag bits ([MS-XLSB] 2.4.100).
constexpr std::uint32_t kShownTotalRow = 0x01U;
constexpr std::uint32_t kSingleCell = 0x02U;
constexpr std::uint32_t kForceInsertToBeVisible = 0x04U;
constexpr std::uint32_t kInsertRowInsCells = 0x08U;
constexpr std::uint32_t kPublished = 0x10U;

// ListType ([MS-XLSB] 2.5.89).
constexpr std::uint32_t kLtRange = 0U;
constexpr std::uint32_t kLtXml = 2U;
constexpr std::uint32_t kLtExtData = 3U;

// BrtListCCFmla / BrtListTrFmla fArray.
constexpr std::uint8_t kFmlaArray = 0x02U;

// `<x14:table>` (alt text) extension of a table, the OOXML twin of BrtList14.
constexpr std::string_view kTableAltTextExtUri = "{504A1905-F514-4f6f-8877-14C23A59335A}";
constexpr std::string_view kX14Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

// ListTotalRowFunction ([MS-XLSB] 2.5.88) indexed by value.
constexpr std::string_view kTotalsFunctions[] = {"",    "average", "count",  "countNums", "max",
                                                 "min", "sum",     "stdDev", "var",       "custom"};
constexpr std::uint32_t kIltaCustom = 9U;

// BrtBeginList's six dxf ids, in record order, and the attributes holding them.
// Record order; `kTableDxfSchemaOrder` lists them in CT_Table attribute order.
constexpr std::string_view kTableDxfAttrs[] = {"headerRowDxfId",       "dataDxfId",
                                               "totalsRowDxfId",       "tableBorderDxfId",
                                               "headerRowBorderDxfId", "totalsRowBorderDxfId"};
constexpr std::size_t kTableDxfSchemaOrder[] = {0, 1, 2, 4, 3, 5};
constexpr std::string_view kTableStyleAttrs[] = {"headerRowCellStyle", "dataCellStyle", "totalsRowCellStyle"};
constexpr std::string_view kColumnDxfAttrs[] = {"headerRowDxfId", "dataDxfId", "totalsRowDxfId"};

using Attrs = std::vector<std::pair<std::string, std::string>>;

Error Unsupported(std::uint16_t type) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord, "xlsb table part carries a record outside the model",
                    "context=xlsb_table record_id=" + std::to_string(type));
}

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, "context=xlsb_table");
}

/// Splits model attribute text (` name="value"` pairs) into unescaped pairs.
Attrs ParseAttrs(std::string_view text) {
  Attrs out;
  if (text.empty()) {
    return out;
  }
  std::string doc_text("<x");
  doc_text.append(text);
  doc_text.append("/>");
  pugi::xml_document doc;
  if (!doc.load_buffer(doc_text.data(), doc_text.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return out;
  }
  for (pugi::xml_attribute attr : doc.first_child().attributes()) {
    out.emplace_back(attr.name(), attr.value());
  }
  return out;
}

const std::string* FindAttr(const Attrs& attrs, std::string_view name) {
  for (const auto& [key, value] : attrs) {
    if (key == name) {
      return &value;
    }
  }
  return nullptr;
}

std::uint32_t DxfAttr(const Attrs& attrs, std::string_view name) {
  const std::string* value = FindAttr(attrs, name);
  if (value == nullptr || value->empty()) {
    return kNoDxf;
  }
  std::uint32_t parsed = 0;
  for (const char c : *value) {
    if (c < '0' || c > '9' || parsed > 0x0FFFFFFFU) {
      return kNoDxf;
    }
    parsed = parsed * 10U + static_cast<std::uint32_t>(c - '0');
  }
  return parsed;
}

bool BoolAttr(const Attrs& attrs, std::string_view name) {
  const std::string* value = FindAttr(attrs, name);
  return value != nullptr && parse_xml_bool(*value);
}

std::optional<std::string_view> NullableAttr(const Attrs& attrs, std::string_view name) {
  const std::string* value = FindAttr(attrs, name);
  if (value == nullptr) {
    return std::nullopt;
  }
  return std::string_view(*value);
}

void AppendDxf(std::string& out, std::string_view name, std::uint32_t dxf) {
  if (dxf != kNoDxf) {
    append_xml_attr_uint(out, name, dxf);
  }
}

void AppendText(std::string& out, std::string_view name, const std::string& value) {
  if (!value.empty()) {
    append_xml_attr(out, name, value);
  }
}

/// Splits a ListParsedFormula off `payload` (after its flag byte) as bytes.
Expected<std::vector<std::uint8_t>, Error> ReadFormulaBytes(ByteSpan payload) {
  ASSIGN_OR_RETURN(const std::uint8_t flags, read_u8(payload));
  if ((flags & kFmlaArray) != 0U) {
    // An array calculated column has no model spelling.
    return Unsupported(kBrtListCCFmla);
  }
  return std::vector<std::uint8_t>(payload.data, payload.data + payload.size);
}

/// `<sortState>` text for a decoded SORTSTATE block.
std::string SortStateXml(const SortState& sort) {
  AutoFilter carrier;
  carrier.sort = sort;
  const std::string wrapped = serialize_auto_filter(carrier);
  pugi::xml_document doc;
  if (!doc.load_buffer(wrapped.data(), wrapped.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return {};
  }
  return raw_xml(doc.first_child().child("sortState"));
}

/// The model of a table's `<sortState>` text, or nullopt when it has none.
std::optional<SortState> SortStateFromXml(const std::string& xml) {
  const AutoFilter carrier = auto_filter_from_xml("<autoFilter ref=\"A1\">" + xml + "</autoFilter>");
  if (carrier.is_opaque() || !carrier.sort || !carrier.columns.empty()) {
    return std::nullopt;
  }
  return carrier.sort;
}

/// `<extLst>` text for a BrtList14 alt text.
std::string AltTextExtLst(const std::string& alt_text, const std::string& summary) {
  std::string out("<extLst><ext");
  append_xml_attr(out, "uri", kTableAltTextExtUri);
  append_xml_attr(out, "xmlns:x14", kX14Namespace);
  out.append("><x14:table");
  AppendText(out, "altText", alt_text);
  AppendText(out, "altTextSummary", summary);
  out.append("/></ext></extLst>");
  return out;
}

/// The alt text pair of an `<extLst>` that holds nothing else, else nullopt.
std::optional<std::pair<std::string, std::string>> AltTextFromExtLst(const std::string& xml) {
  pugi::xml_document doc;
  if (!doc.load_buffer(xml.data(), xml.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return std::nullopt;
  }
  pugi::xml_node ext = doc.first_child().first_child();
  if (!ext || ext.next_sibling() || std::string_view(ext.attribute("uri").value()) != kTableAltTextExtUri) {
    return std::nullopt;
  }
  pugi::xml_node table = ext.first_child();
  if (!table || table.next_sibling() || table.first_child()) {
    return std::nullopt;
  }
  return std::make_pair(std::string(table.attribute("altText").value()),
                        std::string(table.attribute("altTextSummary").value()));
}

/// `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` of a [MS-DTYP] GUID.
std::string GuidText(const std::uint8_t* b) {
  constexpr char kHex[] = "0123456789ABCDEF";
  // Data1, Data2 and Data3 are little-endian; Data4 is a byte array.
  constexpr int kOrder[] = {3, 2, 1, 0, -1, 5, 4, -1, 7, 6, -1, 8, 9, -1, 10, 11, 12, 13, 14, 15};
  std::string out("{");
  for (const int i : kOrder) {
    if (i < 0) {
      out.push_back('-');
      continue;
    }
    out.push_back(kHex[b[i] >> 4U]);
    out.push_back(kHex[b[i] & 0x0FU]);
  }
  out.push_back('}');
  return out;
}

std::optional<std::vector<std::uint8_t>> ParseGuid(std::string_view text) {
  constexpr int kOrder[] = {3, 2, 1, 0, -1, 5, 4, -1, 7, 6, -1, 8, 9, -1, 10, 11, 12, 13, 14, 15};
  if (text.size() != 38U || text.front() != '{' || text.back() != '}') {
    return std::nullopt;
  }
  auto nibble = [](char c) -> int {
    if (c >= '0' && c <= '9') {
      return c - '0';
    }
    if (c >= 'a' && c <= 'f') {
      return c - 'a' + 10;
    }
    if (c >= 'A' && c <= 'F') {
      return c - 'A' + 10;
    }
    return -1;
  };
  std::vector<std::uint8_t> out(16U, 0U);
  std::size_t pos = 1;
  for (const int i : kOrder) {
    if (i < 0) {
      if (text[pos++] != '-') {
        return std::nullopt;
      }
      continue;
    }
    const int hi = nibble(text[pos]);
    const int lo = nibble(text[pos + 1U]);
    if (hi < 0 || lo < 0) {
      return std::nullopt;
    }
    out[static_cast<std::size_t>(i)] = static_cast<std::uint8_t>((hi << 4) | lo);
    pos += 2U;
  }
  return out;
}

/// Reads the rest of an ACUID block after its BrtACBegin: one BrtUid, then
/// BrtACEnd.
Expected<std::string, Error> ReadAcuid(ByteSpan& cursor) {
  ASSIGN_OR_RETURN(const XlsbRecord uid, read_record(cursor));
  ASSIGN_OR_RETURN(const XlsbRecord end, read_record(cursor));
  if (uid.type != kBrtUid || uid.payload.size != 16U || end.type != kBrtACEnd) {
    return Unsupported(uid.type);
  }
  return GuidText(uid.payload.data);
}

/// Emits an ACUID block for `guid`; nothing when it is not a GUID.
void EmitAcuid(std::vector<std::uint8_t>& out, const std::string* guid) {
  if (guid == nullptr) {
    return;
  }
  const std::optional<std::vector<std::uint8_t>> bytes = ParseGuid(*guid);
  if (!bytes) {
    return;
  }
  emit_record(out, kBrtACBegin, ByteSpan{kAcBeginPayload, sizeof(kAcBeginPayload)});
  emit_record(out, kBrtUid, *bytes);
  emit_record(out, kBrtACEnd, ByteSpan{});
}

Expected<void, Error> ReadBeginList(ByteSpan payload, TableMetadata* table) {
  ASSIGN_OR_RETURN(const MergeRange rfx, read_rfx(payload));
  if (rfx.first_row > rfx.last_row || rfx.first_col > rfx.last_col || rfx.last_row >= Sheet::kMaxRows ||
      rfx.last_col >= Sheet::kMaxCols) {
    return Corrupt("xlsb BrtBeginList range is inverted or outside the grid");
  }
  ASSIGN_OR_RETURN(const std::uint32_t lt, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint32_t id, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint32_t crw_header, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint32_t crw_totals, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint32_t flags, read_u32(payload));
  std::uint32_t dxfs[6] = {};
  for (std::uint32_t& dxf : dxfs) {
    ASSIGN_OR_RETURN(const std::uint32_t value, read_u32(payload));
    dxf = value;
  }
  ASSIGN_OR_RETURN(const std::uint32_t conn_id, read_u32(payload));
  ASSIGN_OR_RETURN(auto name, read_xlnullablewidestring(payload));
  ASSIGN_OR_RETURN(auto display_name, read_xlnullablewidestring(payload));
  ASSIGN_OR_RETURN(auto comment, read_xlnullablewidestring(payload));
  std::string styles[3];
  for (std::string& style : styles) {
    ASSIGN_OR_RETURN(auto value, read_xlnullablewidestring(payload));
    style = std::move(value);
  }
  if ((flags & kSingleCell) != 0U || (lt != kLtRange && lt != kLtXml && lt != kLtExtData)) {
    // Single-cell XML-map tables live in their own part type.
    return Unsupported(kBrtBeginList);
  }
  std::string ref;
  if (!append_a1_rectangle(ref, rfx) || display_name.empty()) {
    return Corrupt("xlsb BrtBeginList has no display name or an unprintable range");
  }
  table->id = id;
  table->ref = std::move(ref);
  table->name = name.empty() ? display_name : std::move(name);
  table->display_name = std::move(display_name);
  table->header_row = crw_header != 0U;
  table->totals_row = crw_totals != 0U;
  // Absent `totalsRowShown` reads as true, so only a cleared bit is spelled.
  table->totals_row_shown = (flags & kShownTotalRow) != 0U ? std::nullopt : std::optional<bool>(false);
  std::string& extra = table->root_extra_attrs;
  AppendText(extra, "comment", comment);
  if (lt == kLtXml) {
    append_xml_attr(extra, "tableType", "xml");
  } else if (lt == kLtExtData) {
    append_xml_attr(extra, "tableType", "queryTable");
  }
  if ((flags & kForceInsertToBeVisible) != 0U) {
    append_xml_attr(extra, "insertRow", "1");
  }
  if ((flags & kInsertRowInsCells) != 0U) {
    append_xml_attr(extra, "insertRowShift", "1");
  }
  if ((flags & kPublished) != 0U) {
    append_xml_attr(extra, "published", "1");
  }
  for (const std::size_t i : kTableDxfSchemaOrder) {
    AppendDxf(extra, kTableDxfAttrs[i], dxfs[i]);
  }
  for (std::size_t i = 0; i < 3U; ++i) {
    AppendText(extra, kTableStyleAttrs[i], styles[i]);
  }
  if (conn_id != 0U) {
    append_xml_attr_uint(extra, "connectionId", conn_id);
  }
  return Expected<void, Error>::Ok();
}

Expected<TableColumn, Error> ReadBeginListCol(ByteSpan payload) {
  TableColumn col;
  ASSIGN_OR_RETURN(const std::uint32_t id, read_u32(payload));
  ASSIGN_OR_RETURN(const std::uint32_t ilta, read_u32(payload));
  std::uint32_t dxfs[3] = {};
  for (std::uint32_t& dxf : dxfs) {
    ASSIGN_OR_RETURN(const std::uint32_t value, read_u32(payload));
    dxf = value;
  }
  ASSIGN_OR_RETURN(const std::uint32_t qsi_field, read_u32(payload));
  ASSIGN_OR_RETURN(auto unique_name, read_xlnullablewidestring(payload));
  ASSIGN_OR_RETURN(auto caption, read_xlnullablewidestring(payload));
  ASSIGN_OR_RETURN(auto total, read_xlnullablewidestring(payload));
  std::string styles[3];
  for (std::string& style : styles) {
    ASSIGN_OR_RETURN(auto value, read_xlnullablewidestring(payload));
    style = std::move(value);
  }
  if (ilta >= std::size(kTotalsFunctions)) {
    return Corrupt("xlsb BrtBeginListCol has an unknown totals function");
  }
  col.id = id;
  col.name = std::move(caption);
  col.totals_label = std::move(total);
  col.totals_function = std::string(kTotalsFunctions[ilta]);
  std::string& extra = col.extra_attrs;
  AppendText(extra, "uniqueName", unique_name);
  if (qsi_field != 0U) {
    append_xml_attr_uint(extra, "queryTableFieldId", qsi_field);
  }
  for (std::size_t i = 0; i < 3U; ++i) {
    AppendDxf(extra, kColumnDxfAttrs[i], dxfs[i]);
  }
  for (std::size_t i = 0; i < 3U; ++i) {
    AppendText(extra, kTableStyleAttrs[i], styles[i]);
  }
  return col;
}

void ReadStyleClient(ByteSpan payload, TableStyleInfo* info) {
  auto flags = read_u16(payload);
  auto name = read_xlnullablewidestring(payload);
  info->present = true;
  if (flags) {
    info->show_first_column = (flags.value() & 0x01U) != 0U;
    info->show_last_column = (flags.value() & 0x02U) != 0U;
    info->show_row_stripes = (flags.value() & 0x04U) != 0U;
    info->show_column_stripes = (flags.value() & 0x08U) != 0U;
  }
  if (name) {
    info->name = std::move(name.value());
  }
}

/// Decodes one ListParsedFormula to formula text without a leading `=`.
Expected<std::string, Error> DecodeListFormula(const std::vector<std::uint8_t>& bytes,
                                               const FeatureFormulaReadContext& ctx,
                                               const std::vector<TableMetadata>* tables) {
  ByteSpan cursor{bytes.data(), bytes.size()};
  ASSIGN_OR_RETURN(const std::uint32_t cce, read_u32(cursor));
  if (cce == 0U || cce > cursor.size) {
    return Corrupt("xlsb table formula rgce is empty or truncated");
  }
  const ByteSpan rgce{cursor.data, cce};
  cursor.data += cce;
  cursor.size -= cce;
  ASSIGN_OR_RETURN(const std::uint32_t cb, read_u32(cursor));
  if (cb > cursor.size) {
    return Corrupt("xlsb table formula rgcb is truncated");
  }
  const ByteSpan rgcb{cursor.data, cb};
  Arena arena(/*initial_chunk_bytes=*/4096, kMaxLoadArenaBytes);
  ASSIGN_OR_RETURN(parser::AstNode * ast,
                   decode_ptgs(rgce, rgcb, arena, ctx.sheet_names, ctx.name_table, ctx.sheet_ranges, ctx.external_books,
                               static_cast<std::int32_t>(ctx.sheet_index), std::nullopt, nullptr, tables));
  return parser::format_formula(*ast);
}

/// Encodes `text` as a ListParsedFormula, or nullopt when it cannot be.
std::optional<std::vector<std::uint8_t>> EncodeListFormula(const std::string& text,
                                                           const FeatureFormulaWriteContext& ctx) {
  Arena arena;
  parser::Parser parser(text, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    return std::nullopt;
  }
  auto encoded = encode_ptgs(*root, ctx.sheet_names, ctx.sheet_ranges, ctx.name_table, PtgRootClass::kValue);
  if (!encoded || encoded.value().rgce.empty()) {
    return std::nullopt;
  }
  std::vector<std::uint8_t> out;
  emit_u8(out, 0U);
  emit_u32(out, static_cast<std::uint32_t>(encoded.value().rgce.size()));
  out.insert(out.end(), encoded.value().rgce.begin(), encoded.value().rgce.end());
  emit_u32(out, static_cast<std::uint32_t>(encoded.value().rgcb.size()));
  out.insert(out.end(), encoded.value().rgcb.begin(), encoded.value().rgcb.end());
  return out;
}

std::uint32_t IltaFor(const std::string& function) {
  for (std::uint32_t i = 0; i < std::size(kTotalsFunctions); ++i) {
    if (kTotalsFunctions[i] == function) {
      return i;
    }
  }
  return 0U;
}

void EmitBeginList(std::vector<std::uint8_t>& out, const TableMetadata& table, const MergeRange& rfx,
                   std::uint32_t id) {
  const Attrs attrs = ParseAttrs(table.root_extra_attrs);
  const std::string* type = FindAttr(attrs, "tableType");
  std::uint32_t lt = kLtRange;
  if (type != nullptr && *type == "xml") {
    lt = kLtXml;
  } else if (type != nullptr && *type == "queryTable") {
    lt = kLtExtData;
  }
  std::uint32_t flags = 0;
  flags |= table.totals_row_shown.value_or(true) ? kShownTotalRow : 0U;
  flags |= BoolAttr(attrs, "insertRow") ? kForceInsertToBeVisible : 0U;
  flags |= BoolAttr(attrs, "insertRowShift") ? kInsertRowInsCells : 0U;
  flags |= BoolAttr(attrs, "published") ? kPublished : 0U;
  std::vector<std::uint8_t> payload;
  emit_rfx(payload, rfx);
  emit_u32(payload, lt);
  emit_u32(payload, id);
  emit_u32(payload, table.header_row ? 1U : 0U);
  emit_u32(payload, table.totals_row ? 1U : 0U);
  emit_u32(payload, flags);
  for (const std::string_view name : kTableDxfAttrs) {
    emit_u32(payload, DxfAttr(attrs, name));
  }
  // dwConnID must be zero except on an XML table.
  const std::uint32_t conn_id = DxfAttr(attrs, "connectionId");
  emit_u32(payload, lt == kLtXml && conn_id != kNoDxf ? conn_id : 0U);
  emit_xlnullablewidestring(payload, table.name.empty() || table.name == table.display_name
                                         ? std::nullopt
                                         : std::optional<std::string_view>(table.name));
  emit_xlnullablewidestring(payload, std::string_view(table.display_name));
  const std::string* comment = FindAttr(attrs, "comment");
  emit_xlnullablewidestring(payload, std::string_view(comment != nullptr ? *comment : std::string()));
  for (const std::string_view name : kTableStyleAttrs) {
    emit_xlnullablewidestring(payload, NullableAttr(attrs, name));
  }
  emit_record(out, kBrtBeginList, payload);
}

std::uint32_t EmitListCol(std::vector<std::uint8_t>& out, const TableColumn& col,
                          const FeatureFormulaWriteContext& ctx) {
  const Attrs attrs = ParseAttrs(col.extra_attrs);
  EmitAcuid(out, FindAttr(attrs, "xr3:uid"));
  const std::uint32_t ilta = IltaFor(col.totals_function);
  std::uint32_t lost = (!col.totals_function.empty() && ilta == 0U && col.totals_function != "none") ? 1U : 0U;
  std::vector<std::uint8_t> payload;
  emit_u32(payload, col.id);
  emit_u32(payload, ilta);
  for (const std::string_view name : kColumnDxfAttrs) {
    emit_u32(payload, DxfAttr(attrs, name));
  }
  const std::uint32_t qsi = DxfAttr(attrs, "queryTableFieldId");
  emit_u32(payload, qsi == kNoDxf ? 0U : qsi);
  emit_xlnullablewidestring(payload, NullableAttr(attrs, "uniqueName"));
  emit_xlnullablewidestring(payload, std::string_view(col.name));
  emit_xlnullablewidestring(payload, col.totals_label.empty() || ilta == kIltaCustom
                                         ? std::nullopt
                                         : std::optional<std::string_view>(col.totals_label));
  for (const std::string_view name : kTableStyleAttrs) {
    emit_xlnullablewidestring(payload, NullableAttr(attrs, name));
  }
  emit_record(out, kBrtBeginListCol, payload);
  if (!col.calculated_column_formula.empty()) {
    if (auto bytes = EncodeListFormula(col.calculated_column_formula, ctx)) {
      emit_record(out, kBrtListCCFmla, *bytes);
    } else {
      ++lost;
    }
  }
  if (ilta == kIltaCustom && !col.totals_row_formula.empty()) {
    if (auto bytes = EncodeListFormula(col.totals_row_formula, ctx)) {
      emit_record(out, kBrtListTrFmla, *bytes);
    } else {
      ++lost;
    }
  }
  lost += col.trailing_children_xml.empty() ? 0U : 1U;
  emit_record(out, kBrtEndListCol, ByteSpan{});
  return lost;
}

void EmitStyleClient(std::vector<std::uint8_t>& out, const TableStyleInfo& info) {
  std::vector<std::uint8_t> payload;
  std::uint16_t flags = 0;
  flags |= info.show_first_column ? 0x01U : 0U;
  flags |= info.show_last_column ? 0x02U : 0U;
  flags |= info.show_row_stripes ? 0x04U : 0U;
  flags |= info.show_column_stripes ? 0x08U : 0U;
  emit_u16(payload, flags);
  emit_xlnullablewidestring(payload, info.name.empty() ? std::nullopt : std::optional<std::string_view>(info.name));
  emit_record(out, kBrtTableStyleClient, payload);
}

/// The LISTPARTS block of a tail slot: record index range and its ids.
struct ListPartsBlock {
  std::size_t begin = 0;  // index of BrtBeginListParts
  std::size_t end = 0;    // index of BrtEndListParts
  std::vector<std::string> rel_ids;
};

std::optional<ListPartsBlock> FindListParts(const std::vector<FramedRecord>& recs) {
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (recs[i].type != kBrtBeginListParts) {
      continue;
    }
    ListPartsBlock block;
    block.begin = i;
    for (std::size_t j = i + 1U; j < recs.size(); ++j) {
      if (recs[j].type == kBrtEndListParts) {
        block.end = j;
        return block;
      }
      if (recs[j].type != kBrtListPart) {
        return std::nullopt;
      }
      ByteSpan payload = recs[j].payload;
      auto id = read_xlnullablewidestring(payload);
      if (!id || id.value().empty()) {
        return std::nullopt;
      }
      block.rel_ids.push_back(std::move(id.value()));
    }
    return std::nullopt;
  }
  return std::nullopt;
}

/// Records the worksheet grammar places after the hyperlinks and before
/// LISTPARTS (print setup through web publishing items).
bool PrecedesListParts(std::uint16_t type) {
  switch (type) {
    case 477:  // BrtPrintOptions
    case 476:  // BrtMargins
    case 478:  // BrtPageSetup
    case 479:  // BrtBeginHeaderFooter
    case 480:  // BrtEndHeaderFooter
    case 392:  // BrtBeginRwBrk
    case 393:  // BrtEndRwBrk
    case 394:  // BrtBeginColBrk
    case 395:  // BrtEndColBrk
    case 396:  // BrtBrk
    case 625:  // BrtBigName
    case 605:  // BrtBeginCellWatches
    case 606:  // BrtEndCellWatches
    case 607:  // BrtCellWatch
    case 648:  // BrtBeginCellIgnoreECs
    case 649:  // BrtCellIgnoreEC
    case 650:  // BrtEndCellIgnoreECs
    case 589:  // BrtCellSmartTagProperty
    case 590:  // BrtBeginCellSmartTag
    case 591:  // BrtEndCellSmartTag
    case 592:  // BrtBeginCellSmartTags
    case 593:  // BrtEndCellSmartTags
    case 594:  // BrtBeginSmartTags
    case 595:  // BrtEndSmartTags
    case 550:  // BrtDrawing
    case 551:  // BrtLegacyDrawing
    case 552:  // BrtLegacyDrawingHF
    case 562:  // BrtBkHim
    case 638:  // BrtBeginOleObjects
    case 639:  // BrtOleObject
    case 640:  // BrtEndOleObjects
    case 643:  // BrtBeginActiveXControls
    case 644:  // BrtActiveX
    case 645:  // BrtEndActiveXControls
    case 554:  // BrtBeginWebPubItems
    case 555:  // BrtEndWebPubItems
    case 556:  // BrtBeginWebPubItem
    case 557:  // BrtEndWebPubItem
      return true;
    default:
      return false;
  }
}

std::vector<std::uint8_t> ListPartRecord(const std::string& rel_id) {
  std::vector<std::uint8_t> payload;
  emit_xlnullablewidestring(payload, std::string_view(rel_id));
  std::vector<std::uint8_t> out;
  emit_record(out, kBrtListPart, payload);
  return out;
}

/// Removes the BrtListPart records naming `rel_ids` from the sheet's LISTPARTS
/// block (the whole block once none is left) and the relationships behind
/// them from the sheet.
void ReleaseSheetTables(Sheet& sheet, const std::vector<std::string>& rel_ids) {
  auto released = [&rel_ids](const std::string& id) {
    return std::find(rel_ids.begin(), rel_ids.end(), id) != rel_ids.end();
  };
  XlsbSheetTail tail = sheet.xlsb_tail();
  const std::vector<FramedRecord> recs = split_records(tail.after_hyperlinks);
  if (const std::optional<ListPartsBlock> block = FindListParts(recs)) {
    std::vector<std::string> kept_ids;
    for (const std::string& id : block->rel_ids) {
      if (!released(id)) {
        kept_ids.push_back(id);
      }
    }
    std::vector<std::uint8_t> kept;
    kept.reserve(tail.after_hyperlinks.size());
    for (std::size_t i = 0; i < recs.size(); ++i) {
      if (i < block->begin || i > block->end) {
        append_record(kept, tail.after_hyperlinks, recs[i]);
      } else if (i == block->begin && !kept_ids.empty()) {
        std::vector<std::uint8_t> payload;
        emit_u32(payload, static_cast<std::uint32_t>(kept_ids.size()));
        emit_record(kept, kBrtBeginListParts, payload);
        for (const std::string& id : kept_ids) {
          const std::vector<std::uint8_t> rec = ListPartRecord(id);
          kept.insert(kept.end(), rec.begin(), rec.end());
        }
        emit_record(kept, kBrtEndListParts, ByteSpan{});
      }
    }
    tail.after_hyperlinks = std::move(kept);
    sheet.set_xlsb_tail(std::move(tail));
  }
  std::vector<UnknownRelationship> rels = sheet.unknown_relationships();
  rels.erase(std::remove_if(rels.begin(), rels.end(),
                            [&released](const UnknownRelationship& rel) { return released(rel.id); }),
             rels.end());
  sheet.set_unknown_relationships(std::move(rels));
}

}  // namespace

Expected<TableMetadata, Error> read_table_bin(ByteSpan part, std::size_t sheet_index, TableFormulaBytes* formulas) {
  TableMetadata table;
  table.sheet_index = sheet_index;
  bool begun = false;
  bool ended = false;
  bool style_seen = false;
  std::optional<std::uint32_t> declared_columns;
  bool in_columns = false;
  bool in_column = false;
  bool in_frt = false;
  // The ACUID uid of the record that follows, and whether any was seen.
  std::optional<std::string> uid;
  std::optional<std::string> root_uid;
  bool any_uid = false;
  while (part.size > 0U) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(part));
    if (ended) {
      return Unsupported(rec.type);
    }
    if (rec.type == kBrtACBegin && !uid && !in_frt) {
      ASSIGN_OR_RETURN(auto text, ReadAcuid(part));
      uid = std::move(text);
      any_uid = true;
      continue;
    }
    if (uid && rec.type != kBrtBeginList && rec.type != kBrtBeginAFilter && rec.type != kBrtBeginListCol) {
      return Unsupported(kBrtACBegin);
    }
    if (!begun) {
      if (rec.type != kBrtBeginList) {
        return Unsupported(rec.type);
      }
      RETURN_IF_ERROR(ReadBeginList(rec.payload, &table));
      root_uid = std::move(uid);
      uid.reset();
      begun = true;
      continue;
    }
    if (in_frt) {
      if (rec.type == kBrtList14) {
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(const std::uint32_t frt_header, read_u32(payload));
        ASSIGN_OR_RETURN(auto alt_text, read_xlnullablewidestring(payload));
        ASSIGN_OR_RETURN(auto summary, read_xlnullablewidestring(payload));
        if (frt_header != 0U || !table.ext_lst_xml.empty()) {
          return Unsupported(rec.type);
        }
        table.ext_lst_xml = AltTextExtLst(alt_text, summary);
      } else if (rec.type == kBrtFRTEnd) {
        in_frt = false;
      } else {
        return Unsupported(rec.type);
      }
      continue;
    }
    switch (rec.type) {
      case kBrtBeginAFilter: {
        if (declared_columns || !table.auto_filter_xml.empty()) {
          return Corrupt("xlsb table AutoFilter is out of place");
        }
        ASSIGN_OR_RETURN(AutoFilter filter, read_auto_filter_block(part, rec.payload));
        // Excel nests the sort inside the AutoFilter; the model keeps a
        // table's `<sortState>` beside it, as `.xlsx` does.
        if (filter.sort) {
          table.sort_state_xml = SortStateXml(*filter.sort);
          filter.sort.reset();
        }
        if (uid) {
          filter.extra_attrs.insert(filter.extra_attrs.begin(), {"xr:uid", std::move(*uid)});
          uid.reset();
        }
        table.auto_filter_xml.set(std::move(filter));
        break;
      }
      case kBrtBeginSortState: {
        if (declared_columns || !table.sort_state_xml.empty()) {
          return Corrupt("xlsb table sort state is out of place");
        }
        ASSIGN_OR_RETURN(SortState sort, read_sort_state_block(part, rec.payload));
        table.sort_state_xml = SortStateXml(sort);
        break;
      }
      case kBrtBeginListCols: {
        if (declared_columns) {
          return Corrupt("xlsb table declares its columns twice");
        }
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(const std::uint32_t count, read_u32(payload));
        declared_columns = count;
        in_columns = true;
        break;
      }
      case kBrtBeginListCol: {
        if (!in_columns || in_column) {
          return Corrupt("xlsb table column outside its collection");
        }
        ASSIGN_OR_RETURN(TableColumn col, ReadBeginListCol(rec.payload));
        if (uid) {
          std::string attrs;
          append_xml_attr(attrs, "xr3:uid", *uid);
          col.extra_attrs.insert(0, attrs);
          uid.reset();
        }
        table.columns.push_back(std::move(col));
        formulas->calculated.emplace_back();
        formulas->totals.emplace_back();
        in_column = true;
        break;
      }
      case kBrtListCCFmla: {
        if (!in_column || !formulas->calculated.back().empty()) {
          return Corrupt("xlsb calculated column formula outside a column");
        }
        ASSIGN_OR_RETURN(auto bytes, ReadFormulaBytes(rec.payload));
        formulas->calculated.back() = std::move(bytes);
        break;
      }
      case kBrtListTrFmla: {
        if (!in_column || !formulas->totals.back().empty()) {
          return Corrupt("xlsb totals row formula outside a column");
        }
        ASSIGN_OR_RETURN(auto bytes, ReadFormulaBytes(rec.payload));
        formulas->totals.back() = std::move(bytes);
        break;
      }
      case kBrtEndListCol:
        if (!in_column) {
          return Corrupt("xlsb table column end without a start");
        }
        in_column = false;
        break;
      case kBrtEndListCols:
        if (!in_columns || in_column) {
          return Corrupt("xlsb table column collection end out of place");
        }
        in_columns = false;
        break;
      case kBrtTableStyleClient:
        if (!declared_columns || in_columns || style_seen) {
          return Corrupt("xlsb table style is out of place");
        }
        ReadStyleClient(rec.payload, &table.style_info);
        style_seen = true;
        break;
      case kBrtFRTBegin:
        if (!style_seen) {
          return Unsupported(rec.type);
        }
        in_frt = true;
        break;
      case kBrtEndList:
        ended = true;
        break;
      default:
        return Unsupported(rec.type);
    }
  }
  if (!ended || !style_seen || !declared_columns || *declared_columns != table.columns.size() ||
      table.columns.empty()) {
    return Corrupt("xlsb table part is incomplete or its column count disagrees");
  }
  if (any_uid) {
    // The namespaces Excel declares on a table root that carries uids.
    std::string prefix;
    append_xml_attr(prefix, "xmlns:mc", kMcNamespace);
    append_xml_attr(prefix, "mc:Ignorable", "xr xr3");
    append_xml_attr(prefix, "xmlns:xr", kXrNamespace);
    append_xml_attr(prefix, "xmlns:xr3", kXr3Namespace);
    if (root_uid) {
      append_xml_attr(prefix, "xr:uid", *root_uid);
    }
    table.root_extra_attrs.insert(0, prefix);
  }
  return table;
}

Expected<void, Error> decode_table_formulas(TableMetadata& table, const TableFormulaBytes& formulas,
                                            const FeatureFormulaReadContext& ctx,
                                            const std::vector<TableMetadata>* tables) {
  for (std::size_t i = 0; i < table.columns.size() && i < formulas.calculated.size(); ++i) {
    if (!formulas.calculated[i].empty()) {
      ASSIGN_OR_RETURN(auto text, DecodeListFormula(formulas.calculated[i], ctx, tables));
      table.columns[i].calculated_column_formula = std::move(text);
    }
    if (i < formulas.totals.size() && !formulas.totals[i].empty()) {
      ASSIGN_OR_RETURN(auto text, DecodeListFormula(formulas.totals[i], ctx, tables));
      table.columns[i].totals_row_formula = std::move(text);
    }
  }
  return Expected<void, Error>::Ok();
}

std::vector<std::uint8_t> write_table_bin(const TableMetadata& table, std::uint32_t id,
                                          const FeatureFormulaWriteContext& ctx, std::uint32_t* lost) {
  std::vector<std::uint8_t> out;
  const std::optional<MergeRange> rfx = parse_a1_rectangle(table.ref);
  if (!rfx || table.columns.empty() || table.display_name.empty()) {
    ++*lost;
    return out;
  }
  EmitAcuid(out, FindAttr(ParseAttrs(table.root_extra_attrs), "xr:uid"));
  EmitBeginList(out, table, *rfx, id);
  std::optional<SortState> sort;
  if (!table.sort_state_xml.empty()) {
    sort = SortStateFromXml(table.sort_state_xml);
    *lost += sort ? 0U : 1U;
  }
  if (const AutoFilter* filter = table.auto_filter_xml.get(); filter != nullptr && !filter->is_opaque()) {
    AutoFilter written = *filter;
    if (sort) {
      written.sort = std::move(sort);
      sort.reset();
    }
    for (const auto& [name, value] : written.extra_attrs) {
      if (name == "xr:uid") {
        EmitAcuid(out, &value);
      }
    }
    *lost += emit_auto_filter_block(out, written);
  } else if (filter != nullptr) {
    ++*lost;
  }
  if (sort) {
    *lost += emit_sort_state_block(out, *sort);
  }
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(table.columns.size()));
  emit_record(out, kBrtBeginListCols, payload);
  for (const TableColumn& col : table.columns) {
    *lost += EmitListCol(out, col, ctx);
  }
  emit_record(out, kBrtEndListCols, ByteSpan{});
  EmitStyleClient(out, table.style_info);
  if (!table.ext_lst_xml.empty()) {
    if (auto alt = AltTextFromExtLst(table.ext_lst_xml)) {
      emit_record(out, kBrtFRTBegin, ByteSpan{kFrtVersion.data(), kFrtVersion.size()});
      payload.clear();
      emit_u32(payload, 0U);
      emit_xlnullablewidestring(payload, std::string_view(alt->first));
      emit_xlnullablewidestring(payload, std::string_view(alt->second));
      emit_record(out, kBrtList14, payload);
      emit_record(out, kBrtFRTEnd, ByteSpan{});
    } else {
      ++*lost;
    }
  }
  emit_record(out, kBrtEndList, ByteSpan{});
  return out;
}

void read_xlsb_table_parts(const ZipReader& zip, std::size_t sheet_index, const std::vector<UnknownRelationship>& rels,
                           XlsbTableParts& out) {
  for (const UnknownRelationship& rel : rels) {
    if (rel.type != kRelTable || rel.target_external || !zip.has_entry(rel.target)) {
      continue;
    }
    auto bytes = zip.read_entry(rel.target);
    if (!bytes) {
      continue;
    }
    XlsbTableParts::Part part;
    auto table = read_table_bin(ByteSpan{bytes.value().data(), bytes.value().size()}, sheet_index, &part.formulas);
    if (!table) {
      StructuredLog("xlsb.table.retained_raw").field("part", rel.target).field("reason", table.error().message).warn();
      continue;
    }
    part.sheet_index = sheet_index;
    part.rel_id = rel.id;
    part.path = rel.target;
    part.table = std::move(table.value());
    out.parts.push_back(std::move(part));
  }
}

std::vector<TableMetadata> xlsb_table_directory(const XlsbTableParts& parts) {
  std::vector<TableMetadata> out;
  out.reserve(parts.parts.size());
  for (const XlsbTableParts::Part& part : parts.parts) {
    out.push_back(part.table);
  }
  return out;
}

void load_xlsb_tables(Workbook& wb, XlsbTableParts parts, const std::vector<std::string>& sheet_names,
                      const std::vector<XlsbName>& name_table, const std::vector<XlsbSheetRange>& sheet_ranges,
                      const XlsbExternalBooks& external_books, std::unordered_set<std::string>& consumed_parts) {
  // Formulas decode against every readable table, as the cells did.
  const std::vector<TableMetadata> directory = xlsb_table_directory(parts);
  std::vector<TableMetadata> tables;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const std::optional<ListPartsBlock> block = FindListParts(split_records(wb.sheet(i).xlsb_tail().after_hyperlinks));
    if (!block) {
      continue;
    }
    const FeatureFormulaReadContext ctx{sheet_names, name_table, sheet_ranges, external_books, i};
    std::vector<std::string> modelled;
    for (const std::string& rel_id : block->rel_ids) {
      auto it = std::find_if(parts.parts.begin(), parts.parts.end(), [&](const XlsbTableParts::Part& part) {
        return part.sheet_index == i && part.rel_id == rel_id;
      });
      if (it == parts.parts.end()) {
        continue;
      }
      if (auto status = decode_table_formulas(it->table, it->formulas, ctx, &directory); !status) {
        StructuredLog("xlsb.table.retained_raw").field("part", it->path).field("reason", status.error().message).warn();
        continue;
      }
      tables.push_back(std::move(it->table));
      consumed_parts.insert(it->path);
      modelled.push_back(rel_id);
    }
    if (!modelled.empty()) {
      ReleaseSheetTables(wb.sheet(i), modelled);
    }
  }
  wb.set_tables(std::move(tables));
}

std::unordered_set<std::uint32_t> xlsb_retained_table_ids(const Workbook& wb) {
  std::unordered_set<std::uint32_t> ids;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path.compare(0, 10, "xl/tables/") != 0 || part.path.size() < 4U ||
        part.path.compare(part.path.size() - 4U, 4U, ".bin") != 0) {
      continue;
    }
    for (const FramedRecord& rec : split_records(part.bytes)) {
      if (rec.type != kBrtBeginList) {
        continue;
      }
      ByteSpan payload = rec.payload;
      // rfx (four u32) and the list type precede idList.
      if (payload.size >= 24U) {
        payload = ByteSpan{payload.data + 20, payload.size - 20U};
        if (auto id = read_u32(payload)) {
          ids.insert(id.value());
        }
      }
      break;
    }
  }
  return ids;
}

XlsbTableAllocator::XlsbTableAllocator(std::unordered_set<std::string> taken_paths)
    : taken_paths_(std::move(taken_paths)) {}

std::string XlsbTableAllocator::path_for(std::uint32_t id) {
  std::string path = "xl/tables/table" + std::to_string(id) + ".bin";
  std::uint32_t n = 1;
  while (taken_paths_.count(path) != 0U) {
    path = "xl/tables/table" + std::to_string(id) + "_" + std::to_string(n++) + ".bin";
  }
  taken_paths_.insert(path);
  return path;
}

std::vector<XlsbTablePart> encode_sheet_tables(const Workbook& wb, std::size_t sheet_index,
                                               const FeatureFormulaWriteContext& ctx,
                                               const std::unordered_set<std::string>& taken_rel_ids,
                                               XlsbTableAllocator& allocator, std::uint32_t* lost) {
  std::vector<XlsbTablePart> parts;
  std::size_t next_rel = 1;
  for (std::size_t index = 0; index < wb.tables().size(); ++index) {
    const TableMetadata& table = wb.tables()[index];
    if (table.sheet_index != sheet_index) {
      continue;
    }
    const std::uint32_t id = xlsb_table_id(wb.tables(), index, ctx.sheet_ranges.reserved_table_ids);
    std::vector<std::uint8_t> bytes = write_table_bin(table, id, ctx, lost);
    if (bytes.empty()) {
      continue;
    }
    XlsbTablePart part;
    part.path = allocator.path_for(id);
    do {
      part.rel_id = "rId" + std::to_string(next_rel++);
    } while (taken_rel_ids.count(part.rel_id) != 0U);
    part.bytes = std::move(bytes);
    parts.push_back(std::move(part));
  }
  return parts;
}

void splice_list_parts(std::vector<std::uint8_t>& after_hyperlinks, const std::vector<std::string>& rel_ids) {
  if (rel_ids.empty()) {
    return;
  }
  std::vector<std::uint8_t> parts;
  for (const std::string& rel_id : rel_ids) {
    const std::vector<std::uint8_t> rec = ListPartRecord(rel_id);
    parts.insert(parts.end(), rec.begin(), rec.end());
  }
  const std::vector<FramedRecord> recs = split_records(after_hyperlinks);
  std::vector<std::uint8_t> out;
  out.reserve(after_hyperlinks.size() + parts.size() + 16U);
  if (const std::optional<ListPartsBlock> block = FindListParts(recs)) {
    // A retained block keeps its own parts; ours join it before its end.
    for (std::size_t i = 0; i < recs.size(); ++i) {
      if (i == block->begin) {
        std::vector<std::uint8_t> payload;
        emit_u32(payload, static_cast<std::uint32_t>(block->rel_ids.size() + rel_ids.size()));
        emit_record(out, kBrtBeginListParts, payload);
        continue;
      }
      if (i == block->end) {
        out.insert(out.end(), parts.begin(), parts.end());
      }
      append_record(out, after_hyperlinks, recs[i]);
    }
    after_hyperlinks = std::move(out);
    return;
  }
  std::size_t insert_at = 0;
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (PrecedesListParts(recs[i].type)) {
      insert_at = i + 1U;
    }
  }
  std::vector<std::uint8_t> block;
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(rel_ids.size()));
  emit_record(block, kBrtBeginListParts, payload);
  block.insert(block.end(), parts.begin(), parts.end());
  emit_record(block, kBrtEndListParts, ByteSpan{});
  if (recs.empty()) {
    // Empty, or not whole records: append rather than split it.
    after_hyperlinks.insert(after_hyperlinks.end(), block.begin(), block.end());
    return;
  }
  for (std::size_t i = 0; i < recs.size(); ++i) {
    if (i == insert_at) {
      out.insert(out.end(), block.begin(), block.end());
    }
    append_record(out, after_hyperlinks, recs[i]);
  }
  if (insert_at == recs.size()) {
    out.insert(out.end(), block.begin(), block.end());
  }
  after_hyperlinks = std::move(out);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
