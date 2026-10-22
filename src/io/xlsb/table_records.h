//
// XLSB table parts (`xl/tables/tableN.bin`, [MS-XLSB] 2.1.7.51) and the
// worksheet LISTPARTS block that names them, to and from `TableMetadata`.
//
// TABLE = [ACUID] BrtBeginList(343) [AUTOFILTER] [SORTSTATE] LISTCOLS
//         BrtTableStyleClient(513) FRTTABLE BrtEndList(344)
// LISTCOLS = BrtBeginListCols(345) 1*LISTCOL BrtEndListCols(346)
// LISTCOL = [ACUID] BrtBeginListCol(347) [BrtListCCFmla(351)]
//           [BrtListTrFmla(352)] [LISTXMLCPR] *FRT BrtEndListCol(348)
// LISTPARTS = BrtBeginListParts(660) *BrtListPart(661) BrtEndListParts(662)
//
// Tables are model-owned: the reader decodes each part into
// `Workbook::tables()` and drops its source part, relationship and
// BrtListPart record; the writer regenerates them from the model whatever
// container the workbook came from. A part carrying a record the model does
// not hold (an XML-map column, an unknown future record, an array calculated
// column, an undecodable formula) stays verbatim, beside the generated ones
// in the same LISTPARTS block.
//
// The XML-only table attributes the model keeps as raw attribute text map
// onto BrtBeginList / BrtBeginListCol fields: comment, tableType, insertRow,
// insertRowShift, published, connectionId, the six table and three column
// dxf ids and cell-style names, and a column's uniqueName and
// queryTableFieldId.

#ifndef FORMULON_IO_XLSB_TABLE_RECORDS_H_
#define FORMULON_IO_XLSB_TABLE_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <unordered_set>
#include <vector>

#include "io/xlsb/feature_formula.h"
#include "io/zip_reader.h"
#include "table.h"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Content type of an XLSB table part.
inline constexpr std::string_view kCtTableXlsb = "application/vnd.ms-excel.table";

/// The formulas of one table part, still encoded: a column's calculated
/// column formula and totals-row formula as their ListParsedFormula bytes
/// (`cce` / `rgce` / `cb` / `rgcb`), empty when the record is absent.
struct TableFormulaBytes {
  std::vector<std::vector<std::uint8_t>> calculated;
  std::vector<std::vector<std::uint8_t>> totals;
};

/// Decodes one table part. Formulas are returned encoded in `formulas`
/// because a structured reference inside them can name any table of the
/// workbook; `decode_table_formulas` renders them once every part is read.
/// `kIoXlsbUnknownRecord` names content the model does not hold.
Expected<TableMetadata, Error> read_table_bin(ByteSpan part, std::size_t sheet_index, TableFormulaBytes* formulas);

/// Renders the formulas `read_table_bin` returned into `table`; `tables`
/// resolves the structured references inside them.
Expected<void, Error> decode_table_formulas(TableMetadata& table, const TableFormulaBytes& formulas,
                                            const FeatureFormulaReadContext& ctx,
                                            const std::vector<TableMetadata>* tables);

/// Serialises `table` as a table part with BrtBeginList `idList` = `id`.
/// `lost` counts the settings the records cannot hold and the formulas that
/// failed to encode.
std::vector<std::uint8_t> write_table_bin(const TableMetadata& table, std::uint32_t id,
                                          const FeatureFormulaWriteContext& ctx, std::uint32_t* lost);

/// The table parts a workbook's sheets relate to, read before any formula
/// so a structured reference (PtgList) in a cell resolves.
struct XlsbTableParts {
  struct Part {
    std::size_t sheet_index = 0;
    std::string rel_id;
    std::string path;
    TableMetadata table;
    TableFormulaBytes formulas;
  };
  std::vector<Part> parts;
};

/// Appends the readable table parts sheet `sheet_index`'s relationships
/// (`rels`) name to `out`. A part the model cannot hold is left out.
void read_xlsb_table_parts(const ZipReader& zip, std::size_t sheet_index, const std::vector<UnknownRelationship>& rels,
                           XlsbTableParts& out);

/// The tables of `parts`, formulas not yet rendered: what cell formulas
/// resolve structured references against while the sheets decode.
std::vector<TableMetadata> xlsb_table_directory(const XlsbTableParts& parts);

/// Moves the tables of `parts` whose formulas decode into `wb.tables()`, in
/// each sheet's LISTPARTS order, removing their BrtListPart records and
/// relationships from the sheet and adding their paths to `consumed_parts`.
/// Runs after every sheet body is decoded. Any other table part stays
/// passthrough, its record and relationship retained.
void load_xlsb_tables(Workbook& wb, XlsbTableParts parts, const std::vector<std::string>& sheet_names,
                      const std::vector<XlsbName>& name_table, const std::vector<XlsbSheetRange>& sheet_ranges,
                      const XlsbExternalBooks& external_books, std::unordered_set<std::string>& consumed_parts);

/// Identity, sheet and column captions of every table part kept verbatim in
/// `wb`'s passthrough parts, read from its BrtBeginList and BrtBeginListCol
/// records. A part whose list record cannot be read is left out.
std::vector<XlsbRetainedTable> xlsb_retained_tables(const Workbook& wb);

/// One generated table part and the sheet relationship that names it.
struct XlsbTablePart {
  std::string path;
  std::string rel_id;
  std::vector<std::uint8_t> bytes;
};

/// Hands out table part paths unique across a package.
class XlsbTableAllocator {
 public:
  /// `taken_paths` are part paths already in the package (passthrough).
  explicit XlsbTableAllocator(std::unordered_set<std::string> taken_paths);

  /// `xl/tables/table<id>.bin`, or a free variant of it.
  std::string path_for(std::uint32_t id);

 private:
  std::unordered_set<std::string> taken_paths_;
};

/// Encodes the model tables of sheet `sheet_index`, each with the id
/// `xlsb_table_id` gives it (the one PtgList tokens name). Relationship ids avoid
/// `taken_rel_ids` (the sheet's retained and hyperlink ids). Adds the
/// settings the parts could not hold to `*lost`.
std::vector<XlsbTablePart> encode_sheet_tables(const Workbook& wb, std::size_t sheet_index,
                                               const FeatureFormulaWriteContext& ctx,
                                               const std::unordered_set<std::string>& taken_rel_ids,
                                               XlsbTableAllocator& allocator, std::uint32_t* lost);

/// Adds one BrtListPart per id in `rel_ids` to the worksheet-tail slot that
/// follows the hyperlinks: into a retained LISTPARTS block when the slot has
/// one, otherwise as a new block after the last record the worksheet grammar
/// orders before LISTPARTS.
void splice_list_parts(std::vector<std::uint8_t>& after_hyperlinks, const std::vector<std::string>& rel_ids);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_TABLE_RECORDS_H_
