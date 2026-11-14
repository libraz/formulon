//
// Rich-value errors shared by the OOXML and XLSB packages. Excel stores
// #SPILL!, #CALC! and some #VALUE! as a cell caching #VALUE! whose value
// metadata points at an `_error` rich value carrying the real kind, the
// reason and, for a spill, the needed size. The rich values live in the
// `xl/richData/` XML parts in both formats; only the metadata part differs.
//
// One workbook-wide table numbers the values by first occurrence, sheets in
// tab order and cells row-major; values are deduplicated by content and
// structures by key set. Reading accepts only measured key sets and values,
// so a book carrying anything else keeps its parts verbatim. A save also lists
// the cells' opaque rich values (`rich_values.h`) in that table.

#ifndef FORMULON_IO_RICH_VALUE_ERRORS_H_
#define FORMULON_IO_RICH_VALUE_ERRORS_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/array_anchor_budget.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "sheet_metadata.h"

namespace formulon {
class Workbook;
namespace io {

/// One `_error` rich value. `error_type` is 2 (#VALUE!), 8 (#SPILL!) or 13
/// (#CALC!); the other fields belong to the kinds noted beside them.
struct RichError {
  /// The structure (key set) of the value.
  enum class Kind : std::uint8_t {
    kSpill,            ///< colOffset, errorType, rwOffset, subType
    kSubType,          ///< errorType, subType
    kIftab,            ///< errorType, iftab, subType
    kPropagated,       ///< errorType, propagated
    kPropagatedPtg,    ///< errorType, propagated, ptg
    kPropagatedIftab,  ///< errorType, iftab, propagated
  };
  Kind kind = Kind::kPropagated;
  std::uint32_t error_type = 0;
  std::uint32_t col_offset = 0;  // kSpill
  std::uint32_t rw_offset = 0;   // kSpill
  std::uint32_t sub_type = 0;    // kSpill, kSubType, kIftab
  std::uint32_t iftab = 0;       // kIftab, kPropagatedIftab
  std::uint32_t ptg = 0;         // kPropagatedPtg
};

bool operator==(const RichError& a, const RichError& b) noexcept;

/// Cell key of `RichErrorVmMap`.
inline std::uint64_t rich_error_cell_key(std::uint32_t row, std::uint32_t col) noexcept {
  return (static_cast<std::uint64_t>(row) << 32) | col;
}

/// A sheet's 1-based value-metadata index (`vm`) per cell key, sorted by key.
using RichErrorVmMap = std::vector<std::pair<std::uint64_t, std::uint32_t>>;

/// The `vm` of `(row, col)` in `map`, or 0 for none.
std::uint32_t rich_error_vm_at(const RichErrorVmMap& map, std::uint32_t row, std::uint32_t col) noexcept;

/// The rich values a save writes. `vm` n names value n-1. A value or
/// structure whose `opaque_of` / `structure_opaque` entry is non-zero is an
/// opaque one, and its `values` / `structures` slot is unused.
struct RichErrorTable {
  std::vector<RichError> values;
  std::vector<std::uint32_t> structure_of;      // per value, its structure index
  std::vector<RichError::Kind> structures;      // in first-use order
  std::vector<RichErrorVmMap> vm_by_sheet;      // parallel to the workbook's sheets
  std::vector<std::uint32_t> opaque_of;         // per value: 0, or 1 + its index in `opaque`
  std::vector<std::uint32_t> structure_opaque;  // per structure: 0, or 1 + the `opaque` value declaring it
  /// Opaque values, each `_rvRel:` key's `value` set to its index in `rels`.
  std::vector<OpaqueRichValue> opaque;
  /// `xl/richData/richValueRel.xml` in index order: (target part, relationship type).
  std::vector<std::pair<std::string, std::string>> rels;
  bool empty() const noexcept { return values.empty(); }
};

/// True when `parts` keeps a metadata or richData part the reader did not
/// consume; `build_rich_error_table` then lists nothing for the book.
bool keeps_foreign_rich_data(const std::vector<PassthroughPart>& parts);

/// The rich value a recorded detail is saved as, or nullopt when it is saved
/// as a plain error. An engine-made propagated `#VALUE!` takes its Ptg id from
/// the root operator and its iftab from the measured (function, argument
/// position) table.
std::optional<RichError> rich_error_of_detail(const CellErrorDetail& detail);

/// Decides whether the blocked pivot `pivot` of sheet `sheet` is saved with its
/// anchor's rich value; null keeps every one.
using PivotAnchorFilter = bool (*)(const Workbook& wb, std::size_t sheet, std::size_t pivot);

/// Builds the table for `wb`: every blocked pivot anchor and blocked formula
/// spill anchor, every cell whose detail matches its cached error, and the
/// phantoms such a detail covers. Empty when the book keeps foreign rich data.
RichErrorTable build_rich_error_table(const Workbook& wb, PivotAnchorFilter pivot_filter = nullptr);

/// The table a save writes, in either format: `build_rich_error_table`'s
/// values with each error cell's opaque rich value in its row-major place (a
/// cell's rich error wins), and the `_rvRel:` targets numbered in first-use
/// order. Empty when the book keeps its rich-value parts verbatim
/// (`rich_value_parts_kept_verbatim`).
RichErrorTable build_rich_value_table(const Workbook& wb, PivotAnchorFilter pivot_filter = nullptr);

/// `xl/metadata.xml` listing XLDAPR when `dynamic_arrays` and XLRICHVALUE
/// with `rich_values` blocks when non-zero, as Excel writes it.
std::string build_metadata_xml(bool dynamic_arrays, std::size_t rich_values);

/// `xl/richData/rdrichvalue.xml` for `table`.
std::string serialize_rdrichvalue(const RichErrorTable& table);

/// `xl/richData/rdrichvaluestructure.xml` for `table`.
std::string serialize_rdrichvaluestructure(const RichErrorTable& table);

/// `xl/richData/rdRichValueTypes.xml`: one constant document.
std::string rich_value_types_xml();

/// One entry per `<rv>` of `rdrichvalue`: the value, or nullopt when its key
/// set or a value is not one Excel was measured writing. Empty when either
/// part does not parse.
std::vector<std::optional<RichError>> parse_rich_errors(const std::vector<std::uint8_t>& rdrichvalue,
                                                        const std::vector<std::uint8_t>& rdrichvaluestructure);

/// What `xl/metadata.xml` says about rich values.
struct RichValueMetadata {
  /// Its types are XLDAPR (the standard block only) and XLRICHVALUE, and
  /// every block resolves.
  bool consumable = false;
  /// Per `vm` - 1, the rich value index.
  std::vector<std::uint32_t> vm_rv;
  /// Why the part is not consumable, when `parse_metadata_xml` can name it.
  std::string_view fallback_reason;
};

RichValueMetadata parse_metadata_xml(const std::vector<std::uint8_t>& metadata_xml);

/// Names of the dynamic-array and rich-value metadata types.
constexpr std::string_view kDynamicArrayMetadataType = "XLDAPR";
constexpr std::string_view kRichValueMetadataType = "XLRICHVALUE";

/// One cell- or value-metadata entry: a 1-based type ordinal and an id.
struct MetadataEntry {
  std::uint32_t type = 0;
  std::uint32_t id = 0;
};

/// The metadata part as rich errors use it, in either format: both
/// `xl/metadata.xml` and `xl/metadata.bin` are built from and parsed into it.
struct MetadataModel {
  std::uint32_t type_count = 0;
  std::uint32_t dynamic_type = 0;          ///< 1-based ordinal of XLDAPR, 0 when absent
  std::uint32_t rich_type = 0;             ///< 1-based ordinal of XLRICHVALUE, 0 when absent
  std::size_t dynamic_blocks = 0;          ///< standard XLDAPR future-metadata blocks
  std::vector<std::uint32_t> rich_blocks;  ///< per XLRICHVALUE block, its rich value index
  std::vector<MetadataEntry> cell_entries;
  std::vector<MetadataEntry> value_entries;
};

/// The part Excel writes: XLDAPR with its block and one cell entry when
/// `dynamic_arrays`, then XLRICHVALUE with `rich_values` blocks and as many
/// value entries, entry n naming block n.
MetadataModel metadata_model(bool dynamic_arrays, std::size_t rich_values);

/// Declares type `name` at the next ordinal; false for an unknown or repeated name.
bool add_metadata_type(MetadataModel& model, std::string_view name);

/// Consumable when every cell entry names XLDAPR id 0 and every value entry
/// an XLRICHVALUE block; `vm_rv` follows the value entries.
RichValueMetadata resolve_metadata(const MetadataModel& model);

/// A cell read with a `vm`.
struct RichErrorCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t vm = 0;
};

/// A consumed rich-value set: `values` indexed through `vm_rv`.
struct RichErrorSet {
  std::vector<std::uint32_t> vm_rv;
  std::vector<RichError> values;
};

/// A constant cell saved with a spill rich value, left for the pivot hook.
struct DeferredSpillAnchor {
  std::size_t sheet = 0;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  RichError value;
};

/// Restores sheet `sheet_index`'s rich errors while its cells are decoded,
/// before `register_array_spills`: a formula with a spill value becomes a
/// blocked anchor, any other value restores the real error and its detail,
/// and a phantom of `anchors` gets only the real error, its detail moving to
/// the anchor. A constant with a spill value is appended to `deferred`.
void restore_rich_error_cells(Workbook& wb, std::size_t sheet_index, const std::vector<RichErrorCell>& cells,
                              const std::vector<ArrayAnchor>& anchors, const RichErrorSet& set,
                              std::vector<DeferredSpillAnchor>* deferred);

/// After the pivots attach: a deferred cell anchoring a pivot saved blocked
/// (one-cell location, zero offsets, subType 7 or 3) restores that pivot as
/// blocked; any other is a plain constant #SPILL!.
void restore_deferred_spill_anchors(Workbook& wb, const std::vector<DeferredSpillAnchor>& deferred);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_RICH_VALUE_ERRORS_H_
