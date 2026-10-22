//
// Table model: per-column and per-table metadata for one
// `xl/tables/tableN.xml` part. The tables reader in `io/tables_reader.h`
// fills these; `Workbook` owns them and the structured-reference resolver
// reads them.

#ifndef FORMULON_TABLE_H_
#define FORMULON_TABLE_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "auto_filter.h"

namespace formulon {

/// Per-column metadata for a table.
///
/// `id` is the 1-based column id Excel assigns; `name` is the displayed
/// column header. `totals_label` and `totals_function` are populated
/// when the table has a totals row that selects either a literal label
/// (`totalsRowLabel="..."`) or a built-in function
/// (`totalsRowFunction="sum"|...`); both default to empty.
/// `calculated_column_formula` carries the raw text of the column's
/// `<calculatedColumnFormula>` child, when present. An empty string
/// means the element was absent in the source XML; the writer therefore
/// omits the element entirely on round-trip.
struct TableColumn {
  TableColumn() = default;
  // Keeps the pre-extra-attributes construction shape source-compatible for
  // embedders and tests that initialise the five modelled fields directly.
  TableColumn(std::uint32_t id_in, std::string name_in, std::string totals_label_in, std::string totals_function_in,
              std::string calculated_column_formula_in)
      : id(id_in),
        name(std::move(name_in)),
        totals_label(std::move(totals_label_in)),
        totals_function(std::move(totals_function_in)),
        calculated_column_formula(std::move(calculated_column_formula_in)) {}

  std::uint32_t id = 0;
  std::string name;
  std::string totals_label;
  std::string totals_function;
  /// Raw text of `<calculatedColumnFormula>` (no leading `=`). Empty
  /// when the element was absent in the source XML; structured-
  /// references inside the text are NOT parsed or evaluated here.
  std::string calculated_column_formula;
  /// Text of `<totalsRowFormula>` (no leading `=`), the formula a `custom`
  /// totals-row function stands for. Empty when the element was absent.
  std::string totals_row_formula;
  /// Whether `<totalsRowFormula>` carried `array="1"`.
  bool totals_row_formula_array = false;
  /// `<xmlColumnPr>` and `<extLst>` children, retained verbatim in their
  /// schema positions after the two formula children.
  std::string trailing_children_xml;
  /// Attributes not represented by the table model (notably data/header/
  /// totals-row dxf ids) retained as escaped ` name="value"` pairs.
  std::string extra_attrs;
};

/// A table's `<tableStyleInfo>`: its style name and the four toggles that
/// pick which of the style's regions apply. Absent toggles read as false.
struct TableStyleInfo {
  /// Whether the table carries the element at all.
  bool present = false;
  /// Built-in or custom (`<tableStyles>`) style name; empty for no style.
  std::string name;
  bool show_first_column = false;
  bool show_last_column = false;
  bool show_row_stripes = false;
  bool show_column_stripes = false;
  /// Unmodelled attributes as escaped ` name="value"` pairs.
  std::string extra_attrs;
};

/// In-memory representation of one `xl/tables/tableN.xml` part.
///
/// `id`, `name`, `display_name` and `ref` come straight from the root
/// `<table>` attributes. `sheet_index` is the workbook-relative index
/// of the sheet that owns the table (resolved by the caller via the
/// sheet rels file); the reader does not infer this from the part name.
/// `header_row` and `totals_row` reflect the `headerRowCount` /
/// `totalsRowCount` attributes (any value >= 1 enables the row;
/// `headerRowCount="0"` explicitly disables the header).
struct TableMetadata {
  std::uint32_t id = 0;
  std::string name;
  std::string display_name;
  std::string ref;
  std::size_t sheet_index = 0;
  bool header_row = true;
  bool totals_row = false;
  /// `totalsRowShown`: whether a totals row has ever been shown, which keeps
  /// the columns' totals settings while the row is off. `nullopt` when the
  /// attribute is absent (its default is true).
  std::optional<bool> totals_row_shown;
  std::vector<TableColumn> columns;
  TableStyleInfo style_info;
  /// Table AutoFilter. Keeps the name of the raw element it replaced;
  /// `io/auto_filter_xml.h` reads and writes the element.
  /// `Workbook::set_table_auto_filter` also dirties row-visibility dependents.
  AutoFilterSlot auto_filter_xml;
  /// Unmodelled table-level sort and extension payloads retained verbatim in
  /// their schema positions on write.
  std::string sort_state_xml;
  std::string ext_lst_xml;
  /// Table-root attributes not represented by the model, including extra
  /// namespace declarations and compatibility attributes needed by raw
  /// column or extension payloads.
  std::string root_extra_attrs;
};

}  // namespace formulon

#endif  // FORMULON_TABLE_H_
