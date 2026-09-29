//
// Formulas owned by conditional-format rules and data validations.
//
// Unlike a cell formula, these are stored relative to a base cell -- the
// top-left of the owning sqref's bounding box, which is also what the
// .xlsx formula text is relative to (measured) -- so references with a
// relative axis are PtgRefN / PtgAreaN (see `PtgBaseCell`). This file
// frames them (`cce` / `rgce` / `cb` / `rgcb`) and supplies the base.

#ifndef FORMULON_IO_XLSB_FEATURE_FORMULA_H_
#define FORMULON_IO_XLSB_FEATURE_FORMULA_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/ptg_writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Workbook tables a feature formula resolves names and sheets through.
struct FeatureFormulaReadContext {
  const std::vector<std::string>& sheet_names;
  const std::vector<XlsbName>& name_table;
  const std::vector<XlsbSheetRange>& sheet_ranges;
  const XlsbExternalBooks& external_books;
  std::size_t sheet_index = 0;
};

struct FeatureFormulaWriteContext {
  const std::vector<std::string>& sheet_names;
  const SheetRangeTable& sheet_ranges;
  const NameTable& name_table;
};

/// A feature formula as its records carry it: `cce` + `rgce` + `cb` +
/// `rgcb` (`bytes`), plus the size field BrtBeginCFRule / BrtCFVO keep
/// alongside it (`cb_fmla`). Excel writes `cce` plus 2 per reference token
/// and 4 per PtgName there but reads it only as present / absent: a file
/// storing `cce` opens and re-saves identically (measured), so `cce` is
/// what is written.
struct EncodedFeatureFormula {
  std::vector<std::uint8_t> bytes;
  std::uint32_t cb_fmla = 0;
};

/// Reads an Sqrfx: u32 count, then per range rwFirst, rwLast, colFirst,
/// colLast (u32 each). False for an empty list or a range outside the grid.
bool read_sqref(ByteSpan& cursor, std::vector<MergeRange>& out);

void emit_sqref(std::vector<std::uint8_t>& dst, const std::vector<MergeRange>& ranges);

/// Top-left of the ranges' bounding box: the cell feature formulas are
/// relative to.
MergeRange sqref_base(const std::vector<MergeRange>& ranges);

/// Reads `cce` / `rgce` / `cb` / `rgcb` from `cursor` and decodes the
/// formula anchored at (`base_row`, `base_col`) to its text, without a
/// leading `=`.
Expected<std::string, Error> read_feature_formula(ByteSpan& cursor, std::uint32_t base_row, std::uint32_t base_col,
                                                  const FeatureFormulaReadContext& ctx);

/// Parses `text` (no leading `=`) and encodes it anchored at
/// (`base_row`, `base_col`); a conditional format's formulas pass
/// `PtgEvaluation::kConditionalFormat`.
Expected<EncodedFeatureFormula, Error> encode_feature_formula(std::string_view text, std::uint32_t base_row,
                                                              std::uint32_t base_col, PtgRootClass root_class,
                                                              const FeatureFormulaWriteContext& ctx,
                                                              PtgEvaluation evaluation = PtgEvaluation::kDynamicArray);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_FEATURE_FORMULA_H_
