//
// Coordinate rewriting inside a raw worksheet `<extLst>` for row / column
// structural edits.
//
// The extension payload is retained verbatim, but its `<xm:sqref>` and
// `<xm:f>` elements name cells: the x14 conditional-format ranges, each
// sparkline's location and source, the x14 data-validation ranges. Excel
// moves them with the edit (measured on Windows Excel 365: an x14 CF
// `xm:sqref` mirrors the legacy union sqref and moves with it, a sparkline's
// `xm:f` and `xm:sqref` move like ordinary references), so these helpers
// let the structural-edit pass apply the same mapping the modelled
// structures use. Only those two element kinds are touched; everything
// else keeps its bytes, and an extension nothing changed is not
// re-serialised.

#ifndef FORMULON_IO_EXT_LST_REFS_H_
#define FORMULON_IO_EXT_LST_REFS_H_

#include <cstdint>
#include <functional>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "parser/ast_shift.h"
#include "sheet.h"

namespace formulon::io {

/// Edits one sqref's rectangles in place; leaving it empty removes the
/// element that owns the sqref.
using SqrefRemap = std::function<void(std::vector<MergeRange>&)>;

/// Returns the rewritten text of one formula (no leading `=`), or nothing
/// when it is unchanged.
using FormulaRemap = std::function<std::optional<std::string>(std::string_view)>;

/// Maps every `<xm:sqref>` in `xml` through `remap`. An emptied sqref
/// removes its owner -- the `<x14:sparkline>` (and its group once no
/// sparkline is left), the `<x14:conditionalFormatting>`, or whatever
/// other element holds it -- and any container that leaves empty; a
/// container's `count` attribute is kept in step. An sqref that does not
/// parse is left alone. Returns true when `xml` changed.
bool remap_ext_lst_sqrefs(std::string& xml, const SqrefRemap& remap);

/// Maps the text of every `<xm:f>` in `xml` through `remap`. Returns true
/// when `xml` changed.
bool remap_ext_lst_formulas(std::string& xml, const FormulaRemap& remap);

/// The sparklines an insert copies: those hosted in row (or, with `row_axis`
/// false, column) `source`, the cell next to the inserted band. The k-th of
/// `count` copies is hosted k rows (columns) further on, with its source
/// translated by the same distance.
struct SparklineClone {
  bool row_axis = true;
  std::uint32_t source = 0;
  std::uint32_t count = 0;

  /// Whether a sparkline hosted at `host` is one to copy.
  bool hosted(const MergeRange& host) const {
    return row_axis ? host.first_row == source && host.last_row == source
                    : host.first_col == source && host.last_col == source;
  }

  /// `host` moved `step` rows (columns) on, or nothing past the sheet edge.
  std::optional<MergeRange> host_at(MergeRange host, std::uint32_t step) const {
    (row_axis ? host.first_row : host.first_col) += step;
    (row_axis ? host.last_row : host.last_col) += step;
    if ((row_axis ? host.last_row : host.last_col) >= (row_axis ? Sheet::kMaxRows : Sheet::kMaxCols)) {
      return std::nullopt;
    }
    return host;
  }
};

/// Copy of a formula `steps` rows or columns further on. `$`-absolute
/// coordinates move too; only a whole-column (whole-row) reference stays put
/// on the axis it spans.
class TranslateTransform final : public parser::RefTransform {
 public:
  TranslateTransform(bool row_axis, std::uint32_t steps) noexcept : row_axis_(row_axis), steps_(steps) {}

  std::optional<parser::Reference> apply(const parser::Reference& ref) const override {
    parser::Reference out = ref;
    std::uint32_t& coord = row_axis_ ? out.row : out.col;
    const bool spans_axis = row_axis_ ? ref.is_full_col : ref.is_full_row;
    if (!spans_axis) {
      coord += steps_;
      if (coord >= (row_axis_ ? Sheet::kMaxRows : Sheet::kMaxCols)) {
        return std::nullopt;
      }
    }
    return out;
  }

 private:
  bool row_axis_;
  std::uint32_t steps_;
};

/// Returns the text of one formula translated through `transform`, or
/// nothing when it is unchanged or cannot be translated.
using FormulaTranslate = std::function<std::optional<std::string>(std::string_view, const parser::RefTransform&)>;

/// Appends each copy `clone` describes after its sparkline, translating the
/// source with `translate`. A container's `count` attribute is kept in step.
/// Returns true when `xml` changed.
bool clone_ext_lst_sparklines(std::string& xml, const SparklineClone& clone, const FormulaTranslate& translate);

}  // namespace formulon::io

#endif  // FORMULON_IO_EXT_LST_REFS_H_
