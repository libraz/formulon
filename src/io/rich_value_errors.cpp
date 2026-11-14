// @size-budget: 18 KB
//
// Rich-value error table, its XML parts, the strict reader and the load-time
// restore. See rich_value_errors.h.

#include "io/rich_value_errors.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "io/ooxml_defs.h"
#include "io/rich_values.h"
#include "io/xlsb/func_id_table.h"
#include "io/xlsb/ptg.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "io/xsd_int.h"
#include "parser/ast.h"
#include "pivot/pivot_table.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/index_sort.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

struct RichValueIftabRow {
  std::uint16_t id;
  std::uint16_t iftab_mask;
  std::uint16_t plain_mask;
  std::uint8_t tail_start;
  std::uint8_t group;
};

#include "io/generated/rich_value_iftab.inc"

constexpr std::uint32_t kTypeValue = 2U;
constexpr std::uint32_t kTypeSpill = 8U;
constexpr std::uint32_t kTypeCalc = 13U;
constexpr std::uint32_t kSubOccupied = 1U;
constexpr std::uint32_t kSubPastEdge = 3U;
constexpr std::uint32_t kSubMerged = 6U;
constexpr std::uint32_t kSubPivot = 7U;
constexpr std::uint32_t kSubCap = 9U;
constexpr std::uint16_t kSyntheticIdBase = 0x8000U;

constexpr std::string_view kXmlDeclCrLf = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";

// Keys of rdRichValueTypes.xml's key flags, space-separated in Excel's order.
constexpr std::string_view kTypeFlagKeys =
    "_Self _DisplayString _Flags _Format _SubLabel _Attribution _Icon _Display _CanonicalPropertyNames "
    "_ClassificationId";

constexpr std::string_view kTypeAttrs =
    " minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" "
    "rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\"";

// Keys of each kind, alphabetical as Excel writes them; `b` marks the boolean key.
constexpr const char* kKeys[][4] = {
    {"colOffset", "errorType", "rwOffset", "subType"}, {"errorType", "subType", nullptr, nullptr},
    {"errorType", "iftab", "subType", nullptr},        {"errorType", "propagated", nullptr, nullptr},
    {"errorType", "propagated", "ptg", nullptr},       {"errorType", "iftab", "propagated", nullptr},
};

std::size_t KindIndex(RichError::Kind kind) {
  return static_cast<std::size_t>(kind);
}

// The field a key names; null for `propagated`, which is the constant 1.
std::uint32_t* KeyField(RichError& v, std::string_view key) {
  return key == "colOffset"   ? &v.col_offset
         : key == "errorType" ? &v.error_type
         : key == "rwOffset"  ? &v.rw_offset
         : key == "subType"   ? &v.sub_type
         : key == "iftab"     ? &v.iftab
         : key == "ptg"       ? &v.ptg
                              : nullptr;
}

std::uint32_t KeyValue(RichError v, std::string_view key) {
  const std::uint32_t* field = KeyField(v, key);
  return field != nullptr ? *field : 1U;
}

std::optional<std::uint32_t> TypeOfCode(ErrorCode code) {
  switch (code) {
    case ErrorCode::Value:
      return kTypeValue;
    case ErrorCode::Spill:
      return kTypeSpill;
    case ErrorCode::Calc:
      return kTypeCalc;
    default:
      return std::nullopt;
  }
}

ErrorCode CodeOfType(std::uint32_t type) {
  return type == kTypeSpill ? ErrorCode::Spill : type == kTypeCalc ? ErrorCode::Calc : ErrorCode::Value;
}

RichError SpillValue(std::uint32_t col_offset, std::uint32_t rw_offset, std::uint32_t sub_type) {
  RichError v;
  v.kind = RichError::Kind::kSpill;
  v.error_type = kTypeSpill;
  v.col_offset = col_offset;
  v.rw_offset = rw_offset;
  v.sub_type = sub_type;
  return v;
}

// The function id the measured table is keyed by, or 0 when the function is in neither table.
std::uint16_t FunctionId(std::string_view name) {
  if (const xlsb::XlsbFuncEntry* entry = xlsb::lookup_func_by_name(name)) {
    return entry->id;
  }
  const std::uint16_t* begin = std::begin(kRichValueIftabNameOffsets);
  const std::uint16_t* end = std::end(kRichValueIftabNameOffsets);
  const std::uint16_t* it = std::lower_bound(begin, end, name, [](std::uint16_t offset, std::string_view key) {
    return std::string_view(kRichValueIftabNameText + offset) < key;
  });
  if (it != end && name == std::string_view(kRichValueIftabNameText + *it)) {
    return kRichValueIftabNameIds[it - begin];
  }
  return 0;
}

// The measured row of function `id`; false when it has none.
bool IftabRow(std::uint16_t id, RichValueIftabRow* out) {
  const std::uint16_t* ids_end = std::end(kRichValueIftabFirstOnlyIds);
  const std::uint16_t* first = std::lower_bound(std::begin(kRichValueIftabFirstOnlyIds), ids_end, id);
  if (first != ids_end && *first == id) {
    *out = RichValueIftabRow{id, 1U, 0U, 0U, 0U};
    return true;
  }
  const auto* end = std::end(kRichValueIftabRows);
  const auto* row = std::lower_bound(std::begin(kRichValueIftabRows), end, id,
                                     [](const RichValueIftabRow& r, std::uint16_t key) { return r.id < key; });
  if (row == end || row->id != id) {
    return false;
  }
  *out = *row;
  return true;
}

// The propagated `#VALUE!` a root function call at `argpos` is saved as; nullopt for a plain error.
std::optional<RichError> PropagatedThroughFunction(std::string_view name, std::uint32_t argpos) {
  RichError v;
  v.error_type = kTypeValue;
  const std::uint16_t id = FunctionId(name);
  RichValueIftabRow row{};
  if (id == 0U || argpos == 0U || !IftabRow(id, &row)) {
    return v;
  }
  const std::uint32_t tail = row.tail_start;
  const std::uint32_t group = row.group;
  if (group != 0U && argpos >= tail + group) {
    argpos = tail + (argpos - tail) % group;
  }
  const std::uint32_t bit = argpos <= 16U ? 1U << (argpos - 1U) : 0U;
  if ((row.plain_mask & bit) != 0U) {
    return std::nullopt;
  }
  if ((row.iftab_mask & bit) != 0U && id < kSyntheticIdBase) {
    v.kind = RichError::Kind::kPropagatedIftab;
    v.iftab = id;
  }
  return v;
}

RichError PivotValue(const pivot::PivotTable& table, const pivot::BlockedReport& state) {
  const std::uint32_t body = state.rows > state.rows_above ? state.rows - state.rows_above : 1U;
  const std::uint32_t cols = state.cols != 0U ? state.cols : 1U;
  if (!state.past_edge) {
    return SpillValue(cols - 1U, body - 1U, kSubPivot);
  }
  const std::uint64_t row = table.anchor_row();
  const std::uint64_t col = table.anchor_col();
  const std::uint32_t col_offset =
      col + cols > Sheet::kMaxCols ? static_cast<std::uint32_t>(Sheet::kMaxCols - 1U - col) : cols - 1U;
  const std::uint32_t rw_offset =
      row + body > Sheet::kMaxRows ? static_cast<std::uint32_t>(Sheet::kMaxRows - 1U - row) : body - 1U;
  return SpillValue(col_offset, rw_offset, kSubPastEdge);
}

bool CachesError(const Cell* cell, ErrorCode code) {
  return cell != nullptr && cell->cached_value.is_error() && cell->cached_value.as_error() == code;
}

struct Candidate {
  std::uint32_t row;
  std::uint32_t col;
  RichError value;
  const OpaqueRichValue* opaque = nullptr;  // set for an opaque value; `value` is then unused
};

// Every rich error of `sheet`, unordered.
std::vector<Candidate> SheetCandidates(const Workbook& wb, std::size_t sheet_index, PivotAnchorFilter pivot_filter) {
  const Sheet& sheet = wb.sheet(sheet_index);
  std::vector<Candidate> out;
  for (std::size_t p = 0; p < sheet.pivot_tables().size(); ++p) {
    const std::unique_ptr<pivot::PivotTable>& table = sheet.pivot_tables()[p];
    if (table != nullptr && table->blocked().has_value() &&
        (pivot_filter == nullptr || pivot_filter(wb, sheet_index, p)) &&
        CachesError(sheet.cell_at(table->anchor_row(), table->anchor_col()), ErrorCode::Spill)) {
      out.push_back({table->anchor_row(), table->anchor_col(), PivotValue(*table, *table->blocked())});
    }
  }
  for (const BlockedSpillFootprint& fp : sheet.blocked_spill_footprints()) {
    const Cell* cell = sheet.cell_at(fp.anchor_row, fp.anchor_col);
    if (cell == nullptr || cell->formula_text.empty() || !CachesError(cell, ErrorCode::Spill)) {
      continue;
    }
    const RichError v =
        fp.cause == SpillCause::kBeyondEdge
            ? SpillValue(0U, 0U, kSubPastEdge)
            : SpillValue(fp.cols - 1U, fp.rows - 1U, fp.cause == SpillCause::kMerged ? kSubMerged : kSubOccupied);
    out.push_back({fp.anchor_row, fp.anchor_col, v});
  }
  if (!sheet.has_error_details()) {
    return out;
  }
  for (const auto& [address, detail] : sheet.error_details()) {
    const std::optional<RichError> v = rich_error_of_detail(detail);
    if (!v.has_value()) {
      continue;
    }
    if (CachesError(sheet.cell_at(address.row, address.col), detail.code)) {
      out.push_back({address.row, address.col, *v});
    }
    const SpillRegion* region = detail.elements ? sheet.spill_region_at_anchor(address.row, address.col) : nullptr;
    for (std::size_t i = 1; region != nullptr && i < region->cells.size(); ++i) {
      const Value& cell = region->cells[i];
      if (cell.is_error() && cell.as_error() == detail.code) {
        out.push_back({static_cast<std::uint32_t>(address.row + i / region->cols),
                       static_cast<std::uint32_t>(address.col + i % region->cols), *v});
      }
    }
  }
  return out;
}

bool AnyRichErrorSource(const Workbook& wb, bool with_opaque) {
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const Sheet& sheet = wb.sheet(s);
    if (sheet.has_error_details() || !sheet.blocked_spill_anchors().empty() ||
        (with_opaque && sheet.has_rich_values())) {
      return true;
    }
    for (const std::unique_ptr<pivot::PivotTable>& table : sheet.pivot_tables()) {
      if (table != nullptr && table->blocked().has_value()) {
        return true;
      }
    }
  }
  return false;
}

void AppendCount(std::string& out, std::string_view open, std::size_t count) {
  out.append(open);
  append_xml_attr_uint(out, "count", count);
  out.push_back('>');
}

// Loads part `part` into `doc` and finds its `root` element by local name; false when it does not parse.
bool LoadPart(pugi::xml_document& doc, const std::vector<std::uint8_t>& bytes, const char* part, const char* root,
              pugi::xml_node* out) {
  if (!load_xml_buffer(doc, bytes, "rich_value_errors", part)) {
    return false;
  }
  *out = child_local(doc, root);
  return true;
}

// ---------------------------------------------------------------------------
// Reading
// ---------------------------------------------------------------------------

bool UintAttr(const pugi::xml_node& node, const char* name, std::uint32_t* out) {
  const pugi::xml_attribute attr = node.attribute(name);
  return attr && parse_xsd_nonneg_int(attr.value(), out);
}

// Accepts `v` only when its key set and values are measured.
bool Measured(const RichError& v) {
  switch (v.kind) {
    case RichError::Kind::kSpill:
      return v.error_type == kTypeSpill && (v.sub_type == kSubOccupied || v.sub_type == kSubPastEdge ||
                                            v.sub_type == kSubMerged || v.sub_type == kSubPivot);
    case RichError::Kind::kSubType:
      return v.error_type == kTypeValue || (v.error_type == kTypeCalc && (v.sub_type == 1U || v.sub_type == 2U ||
                                                                          v.sub_type == 3U || v.sub_type == 24U));
    case RichError::Kind::kIftab:
      return v.error_type == kTypeCalc && v.sub_type == kSubCap;
    case RichError::Kind::kPropagated:
      return v.error_type == kTypeValue || v.error_type == kTypeSpill || v.error_type == kTypeCalc;
    case RichError::Kind::kPropagatedPtg:
      return v.error_type == kTypeValue && v.ptg >= 3U && v.ptg <= 14U;
    case RichError::Kind::kPropagatedIftab:
      return v.error_type == kTypeValue;
  }
  return false;
}

// The kind whose key set is `keys` (any order), or nullopt.
std::optional<RichError::Kind> KindOfKeys(std::vector<std::string_view> keys) {
  sort_by_index(keys, [](std::string_view a, std::string_view b) { return a < b; });
  for (std::size_t k = 0; k < std::size(kKeys); ++k) {
    std::size_t n = 0;
    while (n < 4U && kKeys[k][n] != nullptr) {
      ++n;
    }
    bool same = n == keys.size();
    for (std::size_t i = 0; same && i < n; ++i) {
      same = keys[i] == kKeys[k][i];
    }
    if (same) {
      return static_cast<RichError::Kind>(k);
    }
  }
  return std::nullopt;
}

// The XLDAPR block Excel writes for every dynamic-array formula.
bool StandardDynamicArrayBlock(const pugi::xml_node& bk) {
  pugi::xml_node props;
  std::size_t count = 0;
  for (pugi::xml_node ext : child_local(bk, "extLst").children()) {
    for (pugi::xml_node p : ext.children()) {
      props = p;
      ++count;
    }
  }
  return count == 1U && local_name(props) == "dynamicArrayProperties" &&
         std::string_view(props.attribute("fDynamic").value()) == "1" &&
         std::string_view(props.attribute("fCollapsed").as_string("0")) == "0";
}

CellErrorDetail DetailOf(const RichError& v) {
  CellErrorDetail d;
  d.code = CodeOfType(v.error_type);
  d.sub_type = v.sub_type;
  d.iftab = v.iftab;
  d.ptg = v.ptg;
  switch (v.kind) {
    case RichError::Kind::kSubType:
      d.shape = CellErrorDetail::Shape::kSubType;
      break;
    case RichError::Kind::kIftab:
      d.shape = CellErrorDetail::Shape::kIftab;
      break;
    case RichError::Kind::kPropagatedPtg:
      d.shape = CellErrorDetail::Shape::kPropagatedPtg;
      break;
    case RichError::Kind::kPropagatedIftab:
      d.shape = CellErrorDetail::Shape::kPropagatedIftab;
      break;
    default:
      d.shape = CellErrorDetail::Shape::kPropagated;
      break;
  }
  return d;
}

SpillCause CauseOfSubType(std::uint32_t sub_type) {
  return sub_type == kSubPastEdge ? SpillCause::kBeyondEdge
         : sub_type == kSubMerged ? SpillCause::kMerged
                                  : SpillCause::kOccupied;
}

// The anchor whose footprint holds `(row, col)` as a phantom, or null.
const ArrayAnchor* PhantomOf(const std::vector<ArrayAnchor>& anchors, std::uint32_t row, std::uint32_t col) {
  for (const ArrayAnchor& a : anchors) {
    if (row >= a.row && row <= a.last_row && col >= a.col && col <= a.last_col && (row != a.row || col != a.col)) {
      return &a;
    }
  }
  return nullptr;
}

// The index of `rel` in `rels`, appending it on first use.
std::uint32_t RelIndex(std::vector<std::pair<std::string, std::string>>& rels, const OpaqueRichValueKey& key) {
  std::size_t index = 0;
  while (index < rels.size() && !(rels[index].first == key.rel_target && rels[index].second == key.rel_type)) {
    ++index;
  }
  if (index == rels.size()) {
    rels.emplace_back(key.rel_target, key.rel_type);
  }
  return static_cast<std::uint32_t>(index);
}

// Same key declarations: name, type and style, in order.
bool SameStructure(const OpaqueRichValue& a, const OpaqueRichValue& b) {
  if (a.structure_type != b.structure_type || a.keys.size() != b.keys.size()) {
    return false;
  }
  for (std::size_t k = 0; k < a.keys.size(); ++k) {
    if (a.keys[k].name != b.keys[k].name || a.keys[k].type != b.keys[k].type || a.keys[k].style != b.keys[k].style) {
      return false;
    }
  }
  return true;
}

// The table of `build_rich_error_table` and, `with_opaque`, of `build_rich_value_table`.
RichErrorTable BuildTable(const Workbook& wb, PivotAnchorFilter pivot_filter, bool with_opaque) {
  RichErrorTable table;
  table.vm_by_sheet.resize(wb.sheet_count());
  const bool kept_verbatim = with_opaque ? rich_value_parts_kept_verbatim(wb.passthrough_parts())
                                         : keeps_foreign_rich_data(wb.passthrough_parts());
  if (kept_verbatim || !AnyRichErrorSource(wb, with_opaque)) {
    return table;
  }
  // Opaque values as read, parallel to `table.opaque`, for deduplication before renumbering.
  std::vector<OpaqueRichValue> opaque_source;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    std::vector<Candidate> cells = SheetCandidates(wb, s, pivot_filter);
    // Appended after the errors, so the stable sort keeps a cell's rich error first.
    const std::vector<CellRichValue> opaque_cells =
        with_opaque ? wb.sheet(s).rich_values() : std::vector<CellRichValue>();
    for (const CellRichValue& c : opaque_cells) {
      if (CachesError(wb.sheet(s).cell_at(c.row, c.col), ErrorCode::Value)) {
        cells.push_back({c.row, c.col, RichError{}, &c.value});
      }
    }
    sort_by_index(
        cells, [](const Candidate& a, const Candidate& b) { return a.row != b.row ? a.row < b.row : a.col < b.col; });
    RichErrorVmMap& vm = table.vm_by_sheet[s];
    for (const Candidate& c : cells) {
      const std::uint64_t key = rich_error_cell_key(c.row, c.col);
      if (!vm.empty() && vm.back().first == key) {
        continue;
      }
      std::size_t index = 0;
      while (index < table.values.size()) {
        const std::uint32_t opaque = index < table.opaque_of.size() ? table.opaque_of[index] : 0U;
        if (c.opaque == nullptr ? opaque == 0U && table.values[index] == c.value
                                : opaque != 0U && opaque_source[opaque - 1U] == *c.opaque) {
          break;
        }
        ++index;
      }
      if (index == table.values.size()) {
        std::size_t structure = 0;
        while (structure < table.structures.size()) {
          const std::uint32_t opaque =
              structure < table.structure_opaque.size() ? table.structure_opaque[structure] : 0U;
          if (c.opaque == nullptr ? opaque == 0U && table.structures[structure] == c.value.kind
                                  : opaque != 0U && SameStructure(opaque_source[opaque - 1U], *c.opaque)) {
            break;
          }
          ++structure;
        }
        std::uint32_t opaque = 0U;
        if (c.opaque != nullptr) {
          OpaqueRichValue saved = *c.opaque;
          for (OpaqueRichValueKey& k : saved.keys) {
            if (is_rich_value_rel_key(k.name)) {
              k.value = std::to_string(RelIndex(table.rels, k));
            }
          }
          table.opaque.push_back(std::move(saved));
          opaque_source.push_back(*c.opaque);
          opaque = static_cast<std::uint32_t>(table.opaque.size());
        }
        if (structure == table.structures.size()) {
          table.structures.push_back(c.value.kind);
          table.structure_opaque.push_back(opaque);
        }
        table.values.push_back(c.value);
        table.structure_of.push_back(static_cast<std::uint32_t>(structure));
        table.opaque_of.push_back(opaque);
      }
      vm.emplace_back(key, static_cast<std::uint32_t>(index + 1U));
    }
  }
  return table;
}

}  // namespace

bool operator==(const RichError& a, const RichError& b) noexcept {
  return a.kind == b.kind && a.error_type == b.error_type && a.col_offset == b.col_offset &&
         a.rw_offset == b.rw_offset && a.sub_type == b.sub_type && a.iftab == b.iftab && a.ptg == b.ptg;
}

bool keeps_foreign_rich_data(const std::vector<PassthroughPart>& parts) {
  for (const PassthroughPart& part : parts) {
    if (part.path == "xl/metadata.xml" || part.path.compare(0, kRichDataDirPrefix.size(), kRichDataDirPrefix) == 0) {
      return true;
    }
  }
  return false;
}

std::optional<RichError> rich_error_of_detail(const CellErrorDetail& detail) {
  const std::optional<std::uint32_t> type = TypeOfCode(detail.code);
  if (!type.has_value()) {
    return std::nullopt;
  }
  RichError v;
  v.error_type = *type;
  switch (detail.shape) {
    case CellErrorDetail::Shape::kSubType:
      v.kind = RichError::Kind::kSubType;
      v.sub_type = detail.sub_type;
      return v;
    case CellErrorDetail::Shape::kIftab:
      v.kind = RichError::Kind::kIftab;
      v.sub_type = detail.sub_type;
      v.iftab = detail.iftab;
      return v;
    case CellErrorDetail::Shape::kPropagated:
      return v;
    case CellErrorDetail::Shape::kPropagatedPtg:
      v.kind = RichError::Kind::kPropagatedPtg;
      v.ptg =
          detail.root_binop != 0U ? xlsb::ptg_of_binop(static_cast<parser::BinOp>(detail.root_binop - 1U)) : detail.ptg;
      return v;
    case CellErrorDetail::Shape::kPropagatedIftab:
      if (!detail.root_function.empty()) {
        return PropagatedThroughFunction(detail.root_function, detail.root_argpos);
      }
      v.kind = RichError::Kind::kPropagatedIftab;
      v.iftab = detail.iftab;
      return v;
  }
  return std::nullopt;
}

std::uint32_t rich_error_vm_at(const RichErrorVmMap& map, std::uint32_t row, std::uint32_t col) noexcept {
  const std::uint64_t key = rich_error_cell_key(row, col);
  const auto it = std::lower_bound(
      map.begin(), map.end(), key,
      [](const std::pair<std::uint64_t, std::uint32_t>& entry, std::uint64_t k) { return entry.first < k; });
  return it != map.end() && it->first == key ? it->second : 0U;
}

RichErrorTable build_rich_error_table(const Workbook& wb, PivotAnchorFilter pivot_filter) {
  return BuildTable(wb, pivot_filter, /*with_opaque=*/false);
}

RichErrorTable build_rich_value_table(const Workbook& wb) {
  return BuildTable(wb, nullptr, /*with_opaque=*/true);
}

MetadataModel metadata_model(bool dynamic_arrays, std::size_t rich_values) {
  MetadataModel model;
  if (dynamic_arrays) {
    model.dynamic_type = ++model.type_count;
    model.dynamic_blocks = 1U;
    model.cell_entries.push_back({model.dynamic_type, 0U});
  }
  if (rich_values != 0U) {
    model.rich_type = ++model.type_count;
    for (std::size_t i = 0; i < rich_values; ++i) {
      model.rich_blocks.push_back(static_cast<std::uint32_t>(i));
      model.value_entries.push_back({model.rich_type, static_cast<std::uint32_t>(i)});
    }
  }
  return model;
}

bool add_metadata_type(MetadataModel& model, std::string_view name) {
  const bool dynamic = name == kDynamicArrayMetadataType;
  std::uint32_t& slot = dynamic ? model.dynamic_type : model.rich_type;
  if ((!dynamic && name != kRichValueMetadataType) || slot != 0U) {
    return false;
  }
  slot = ++model.type_count;
  return true;
}

RichValueMetadata resolve_metadata(const MetadataModel& model) {
  RichValueMetadata out;
  for (const MetadataEntry& e : model.cell_entries) {
    if (model.dynamic_type == 0U || e.type != model.dynamic_type || e.id != 0U) {
      return RichValueMetadata{};
    }
  }
  for (const MetadataEntry& e : model.value_entries) {
    if (model.rich_type == 0U || e.type != model.rich_type || e.id >= model.rich_blocks.size()) {
      return RichValueMetadata{};
    }
    out.vm_rv.push_back(model.rich_blocks[e.id]);
  }
  out.consumable = true;
  return out;
}

std::string build_metadata_xml(bool dynamic_arrays, std::size_t rich_values) {
  const MetadataModel model = metadata_model(dynamic_arrays, rich_values);
  std::string out(kXmlDeclCrLf);
  out.append("<metadata");
  append_xml_ns_decl(out, "", kNsSpreadsheetMl);
  if (model.rich_type != 0U) {
    append_xml_ns_decl(out, "xlrd", kNsRichData);
  }
  if (model.dynamic_type != 0U) {
    append_xml_ns_decl(out, "xda", kNsDynamicArray);
  }
  out.push_back('>');
  AppendCount(out, "<metadataTypes", model.type_count);
  for (std::uint32_t t = 1; t <= model.type_count; ++t) {
    const bool dynamic = t == model.dynamic_type;
    out.append("<metadataType name=\"")
        .append(dynamic ? kDynamicArrayMetadataType : kRichValueMetadataType)
        .append("\"")
        .append(kTypeAttrs)
        .append(dynamic ? " cellMeta=\"1\"/>" : "/>");
  }
  out.append("</metadataTypes>");
  if (model.dynamic_blocks != 0U) {
    AppendCount(out, "<futureMetadata name=\"XLDAPR\"", model.dynamic_blocks);
    for (std::size_t i = 0; i < model.dynamic_blocks; ++i) {
      out.append(
          "<bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\">"
          "<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk>");
    }
    out.append("</futureMetadata>");
  }
  if (!model.rich_blocks.empty()) {
    AppendCount(out, "<futureMetadata name=\"XLRICHVALUE\"", model.rich_blocks.size());
    for (const std::uint32_t rv : model.rich_blocks) {
      out.append("<bk><extLst><ext uri=\"{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}\"><xlrd:rvb");
      append_xml_attr_uint(out, "i", rv);
      out.append("/></ext></extLst></bk>");
    }
    out.append("</futureMetadata>");
  }
  const auto append_entries = [&out](const char* open, const char* close, const std::vector<MetadataEntry>& entries) {
    if (entries.empty()) {
      return;
    }
    AppendCount(out, open, entries.size());
    for (const MetadataEntry& e : entries) {
      out.append("<bk><rc");
      append_xml_attr_uint(out, "t", e.type);
      append_xml_attr_uint(out, "v", e.id);
      out.append("/></bk>");
    }
    out.append(close);
  };
  append_entries("<cellMetadata", "</cellMetadata>", model.cell_entries);
  append_entries("<valueMetadata", "</valueMetadata>", model.value_entries);
  out.append("</metadata>");
  return out;
}

std::string serialize_rdrichvalue(const RichErrorTable& table) {
  std::string out(kXmlDeclCrLf);
  out.append("<rvData");
  append_xml_ns_decl(out, "", kNsRichData);
  AppendCount(out, "", table.values.size());
  for (std::size_t i = 0; i < table.values.size(); ++i) {
    const RichError& v = table.values[i];
    out.append("<rv");
    append_xml_attr_uint(out, "s", table.structure_of[i]);
    out.push_back('>');
    if (i < table.opaque_of.size() && table.opaque_of[i] != 0U) {
      for (const OpaqueRichValueKey& key : table.opaque[table.opaque_of[i] - 1U].keys) {
        out.append("<v>");
        AppendXmlRawTextEscaped(out, key.value);
        out.append("</v>");
      }
      out.append("</rv>");
      continue;
    }
    for (const char* key : kKeys[KindIndex(v.kind)]) {
      if (key != nullptr) {
        out.append("<v>").append(std::to_string(KeyValue(v, key))).append("</v>");
      }
    }
    out.append("</rv>");
  }
  out.append("</rvData>");
  return out;
}

std::string serialize_rdrichvaluestructure(const RichErrorTable& table) {
  std::string out(kXmlDeclCrLf);
  out.append("<rvStructures");
  append_xml_ns_decl(out, "", kNsRichData);
  AppendCount(out, "", table.structures.size());
  for (std::size_t i = 0; i < table.structures.size(); ++i) {
    const RichError::Kind kind = table.structures[i];
    if (i < table.structure_opaque.size() && table.structure_opaque[i] != 0U) {
      const OpaqueRichValue& value = table.opaque[table.structure_opaque[i] - 1U];
      out.append("<s");
      append_xml_attr(out, "t", value.structure_type);
      out.push_back('>');
      for (const OpaqueRichValueKey& key : value.keys) {
        out.append("<k");
        append_xml_attr(out, "n", key.name);
        if (!key.type.empty()) {
          append_xml_attr(out, "t", key.type);
        }
        if (!key.style.empty()) {
          append_xml_attr(out, "s", key.style);
        }
        out.append("/>");
      }
      out.append("</s>");
      continue;
    }
    out.append("<s t=\"_error\">");
    for (const char* key : kKeys[KindIndex(kind)]) {
      if (key != nullptr) {
        out.append("<k n=\"").append(key).append(std::string_view(key) == "propagated" ? "\" t=\"b\"/>"
                                                                                       : "\" t=\"i\"/>");
      }
    }
    out.append("</s>");
  }
  out.append("</rvStructures>");
  return out;
}

std::string rich_value_types_xml() {
  std::string out(kXmlDeclCrLf);
  out.append(
      "<rvTypesInfo xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\" "
      "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" "
      "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><global><keyFlags>");
  for (std::string_view keys = kTypeFlagKeys; !keys.empty();) {
    const std::size_t space = keys.find(' ');
    const std::string_view key = keys.substr(0, space);
    keys = space == std::string_view::npos ? std::string_view() : keys.substr(space + 1U);
    out.append("<key name=\"").append(key).append("\">");
    if (key == "_Self") {
      out.append("<flag name=\"ExcludeFromFile\" value=\"1\"/>");
    }
    out.append("<flag name=\"ExcludeFromCalcComparison\" value=\"1\"/></key>");
  }
  out.append("</keyFlags></global></rvTypesInfo>");
  return out;
}

std::vector<std::optional<RichError>> parse_rich_errors(const std::vector<std::uint8_t>& rdrichvalue,
                                                        const std::vector<std::uint8_t>& rdrichvaluestructure) {
  pugi::xml_document values_doc;
  pugi::xml_document structures_doc;
  pugi::xml_node values_root;
  pugi::xml_node structures_root;
  if (!LoadPart(values_doc, rdrichvalue, "rdrichvalue.xml", "rvData", &values_root) ||
      !LoadPart(structures_doc, rdrichvaluestructure, "rdrichvaluestructure.xml", "rvStructures", &structures_root)) {
    return {};
  }
  // Per structure: its kind and its keys in document order, or no kind when foreign.
  struct Structure {
    std::optional<RichError::Kind> kind;
    std::vector<std::string_view> keys;
  };
  std::vector<Structure> structures;
  for (pugi::xml_node s : structures_root.children()) {
    Structure entry;
    bool typed = std::string_view(s.attribute("t").value()) == "_error";
    for (pugi::xml_node k : s.children()) {
      const std::string_view name = k.attribute("n").value();
      typed = typed && std::string_view(k.attribute("t").value()) == (name == "propagated" ? "b" : "i");
      entry.keys.push_back(name);
    }
    entry.kind = typed ? KindOfKeys(entry.keys) : std::nullopt;
    structures.push_back(std::move(entry));
  }
  std::vector<std::optional<RichError>> out;
  for (pugi::xml_node rv : values_root.children()) {
    std::uint32_t s = 0;
    std::optional<RichError> parsed;
    if (UintAttr(rv, "s", &s) && s < structures.size() && structures[s].kind.has_value()) {
      RichError v;
      v.kind = *structures[s].kind;
      std::size_t i = 0;
      bool ok = true;
      for (pugi::xml_node cell : rv.children()) {
        std::uint32_t n = 0;
        ok = ok && i < structures[s].keys.size() && parse_xsd_nonneg_int(cell.text().get(), &n);
        if (!ok) {
          break;
        }
        std::uint32_t* field = KeyField(v, structures[s].keys[i++]);
        if (field != nullptr) {
          *field = n;
        } else {
          ok = n == 1U;
        }
      }
      if (ok && i == structures[s].keys.size() && Measured(v)) {
        parsed = v;
      }
    }
    out.push_back(parsed);
  }
  return out;
}

RichValueMetadata parse_metadata_xml(const std::vector<std::uint8_t>& metadata_xml) {
  const auto refuse = [](std::string_view reason) {
    RichValueMetadata out;
    out.fallback_reason = reason;
    return out;
  };
  pugi::xml_document doc;
  pugi::xml_node root;
  if (!LoadPart(doc, metadata_xml, "metadata.xml", "metadata", &root)) {
    return refuse(kRichFallbackMetadataBlock);
  }
  MetadataModel model;
  for (pugi::xml_node type : child_local(root, "metadataTypes").children()) {
    if (!add_metadata_type(model, type.attribute("name").value())) {
      return refuse(kRichFallbackMetadataType);
    }
  }
  for (pugi::xml_node child : root.children()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    const std::string_view element = local_name(child);
    const std::string_view name = child.attribute("name").value();
    if (element == "metadataTypes") {
      continue;
    }
    if (element == "futureMetadata" && name == kDynamicArrayMetadataType) {
      std::size_t blocks = 0;
      for (pugi::xml_node bk : child.children()) {
        if (!StandardDynamicArrayBlock(bk)) {
          return refuse(kRichFallbackMetadataBlock);
        }
        ++blocks;
      }
      if (blocks != 1U) {
        return refuse(kRichFallbackMetadataBlock);
      }
      ++model.dynamic_blocks;
    } else if (element == "futureMetadata" && name == kRichValueMetadataType) {
      for (pugi::xml_node bk : child.children()) {
        pugi::xml_node ext = child_local(child_local(bk, "extLst"), "ext");
        const pugi::xml_node rvb = child_local(ext, "rvb");
        std::uint32_t i = 0;
        if (!UintAttr(rvb, "i", &i)) {
          return refuse(kRichFallbackMetadataBlock);
        }
        for (pugi::xml_attribute attr = rvb.first_attribute(); attr; attr = attr.next_attribute()) {
          const std::string_view attr_name = attr.name();
          if (attr_name != "i" && attr_name.compare(0, 5, "xmlns") != 0) {
            return refuse(kRichFallbackRvbAttribute);
          }
        }
        model.rich_blocks.push_back(i);
      }
    } else if (element == "cellMetadata" || element == "valueMetadata") {
      const bool cell = element == "cellMetadata";
      for (pugi::xml_node bk : child.children()) {
        const pugi::xml_node rc = child_local(bk, "rc");
        MetadataEntry entry;
        if (!UintAttr(rc, "t", &entry.type) || !UintAttr(rc, "v", &entry.id) || rc.next_sibling()) {
          return refuse(kRichFallbackMetadataBlock);
        }
        // A value entry names a block declared before it.
        if (!cell && entry.id >= model.rich_blocks.size()) {
          return refuse(kRichFallbackMetadataBlock);
        }
        (cell ? model.cell_entries : model.value_entries).push_back(entry);
      }
    } else {
      return refuse(kRichFallbackMetadataBlock);
    }
  }
  RichValueMetadata out = resolve_metadata(model);
  if (!out.consumable) {
    out.fallback_reason = kRichFallbackMetadataBlock;
  }
  return out;
}

void restore_rich_error_cells(Workbook& wb, std::size_t sheet_index, const std::vector<RichErrorCell>& cells,
                              const std::vector<ArrayAnchor>& anchors, const RichErrorSet& set,
                              std::vector<DeferredSpillAnchor>* deferred) {
  Sheet& sheet = wb.sheet(sheet_index);
  // Phantom values move to their anchor, which carries them for the elements.
  std::vector<std::pair<const ArrayAnchor*, RichError>> element_values;
  for (const RichErrorCell& c : cells) {
    if (c.vm == 0U || c.vm > set.vm_rv.size() || set.vm_rv[c.vm - 1U] >= set.values.size()) {
      continue;
    }
    const RichError& v = set.values[set.vm_rv[c.vm - 1U]];
    const Cell* cell = sheet.cell_at(c.row, c.col);
    if (cell == nullptr || !cell->cached_value.is_error()) {
      continue;
    }
    if (v.kind == RichError::Kind::kSpill) {
      if (cell->formula_text.empty()) {
        deferred->push_back({sheet_index, c.row, c.col, v});
      } else {
        (void)wb.restore_blocked_spill(sheet_index, c.row, c.col, v.rw_offset + 1U, v.col_offset + 1U,
                                       CauseOfSubType(v.sub_type));
      }
      continue;
    }
    if (const ArrayAnchor* anchor = PhantomOf(anchors, c.row, c.col)) {
      sheet.set_cell_cached_value(c.row, c.col, Value::error(CodeOfType(v.error_type)));
      element_values.emplace_back(anchor, v);
      continue;
    }
    (void)wb.restore_error_detail(sheet_index, c.row, c.col, CodeOfType(v.error_type), DetailOf(v));
  }
  for (const auto& [anchor, v] : element_values) {
    const std::optional<CellErrorDetail> own = sheet.error_detail_at(anchor->row, anchor->col);
    if (own.has_value() && !(std::optional<RichError>(v) == rich_error_of_detail(*own))) {
      continue;
    }
    CellErrorDetail detail = DetailOf(v);
    detail.elements = true;
    sheet.set_error_detail(anchor->row, anchor->col, detail);
  }
}

void restore_deferred_spill_anchors(Workbook& wb, const std::vector<DeferredSpillAnchor>& deferred) {
  const auto zero = [](std::optional<std::uint32_t> v) { return v.has_value() && *v == 0U; };
  for (const DeferredSpillAnchor& d : deferred) {
    const auto& tables = wb.sheet(d.sheet).pivot_tables();
    bool restored = false;
    for (std::size_t p = 0; !restored && p < tables.size(); ++p) {
      const pivot::PivotTable* table = tables[p].get();
      if (table == nullptr || table->anchor_row() != d.row || table->anchor_col() != d.col || !table->span_read() ||
          table->span_rows() != 1U || table->span_cols() != 1U || !zero(table->location_first_header_row()) ||
          !zero(table->location_first_data_row()) || !zero(table->location_first_data_col()) ||
          (d.value.sub_type != kSubPivot && d.value.sub_type != kSubPastEdge)) {
        continue;
      }
      restored = static_cast<bool>(
          wb.restore_blocked_pivot(d.sheet, p, d.value.rw_offset + 1U, d.value.col_offset + 1U, d.value.sub_type));
    }
    if (!restored) {
      wb.sheet(d.sheet).set_cell_cached_value(d.row, d.col, Value::error(ErrorCode::Spill));
    }
  }
}

}  // namespace io
}  // namespace formulon
