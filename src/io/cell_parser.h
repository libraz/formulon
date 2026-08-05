//
// OOXML `<c>` (cell) element parser. Decodes a single cell node from a
// `sheet*.xml` document into a (row, col, value, formula) tuple suitable
// for handoff to `Workbook::set_cell_value` / `set_cell_formula`. The
// surrounding `<sheetData>` walk and shared-formula bookkeeping live in
// `sheet_reader.{h,cpp}`; this header is the leaf that touches one `<c>`
// at a time.
//
// Shared-strings handling: when a cell carries `t="s"`, this slice does
// not yet resolve the index against the workbook's shared-string table
// (Bundle 2.3 will). Instead the parser surfaces the index through
// `ParsedCell::sst_index` and sets `is_sst_index = true`; the consumer
// records the (row, col, sst_index) triple in a side table for later
// resolution and writes a `Text("")` placeholder into the cell.

#ifndef FORMULON_IO_CELL_PARSER_H_
#define FORMULON_IO_CELL_PARSER_H_

#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <utility>

#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace io {

/// Outcome of parsing a single OOXML `<c>` element.
///
/// `value` carries the literal payload as decoded from `<v>` / `<is>`. For
/// shared-string cells (`t="s"`) the parser stores `Text("")` and surfaces
/// the SST index via `sst_index`; the caller resolves it in a follow-up
/// bundle.
///
/// The text payload of `value` (when applicable) is a non-owning
/// `string_view` into the `text_storage` argument supplied by the caller
/// to `parse_cell_element`. The caller therefore controls text lifetime
/// directly: typically by appending into a workbook-lifetime
/// `std::deque<std::string>` so subsequent appends do not invalidate
/// earlier views.
struct ParsedCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  /// Empty when the cell has no formula. Otherwise the formula text WITHOUT
  /// the leading `=` (OOXML convention; `Workbook::set_cell_formula` accepts
  /// both spellings, but we strip the `=` here to keep the contract narrow).
  std::string formula;
  /// The literal / cached value:
  ///   * `Bool`  — parsed from `<v>0</v>` / `<v>1</v>` when `t="b"`
  ///   * `Error` — parsed from `<v>#DIV/0!</v>` etc. when `t="e"`
  ///   * `Number` — parsed from `<v>3.14</v>` (default or `t="n"`)
  ///   * `Text`  — concatenation of all `<t>` text nodes inside `<is>` when
  ///               `t="inlineStr"` (rich-text `<r>` children are walked in
  ///               document order; formatting is dropped). Backed by an
  ///               entry appended into the caller's `text_storage`.
  ///   * `Text("")` — placeholder when `t="s"`; the SST index is in
  ///                  `sst_index`, with `is_sst_index = true`
  ///   * `Blank` — when `<v>` and `<is>` are both absent
  Value value = Value::blank();
  /// Bundle 2.3 hand-off: when `t="s"`, the parser surfaces the SST index
  /// here instead of resolving it. The reader records (row, col,
  /// sst_index) in a side table and replaces the placeholder text once the
  /// SST is loaded.
  bool is_sst_index = false;
  std::uint32_t sst_index = 0;
  /// Concatenated kana text from any `<rPh>` annotations attached to the
  /// cell's `<is>` (inline-string) block, in document order. Empty when
  /// the cell has no `<is>` block or the block carries no `<rPh>`. SST-
  /// referenced cells (`t="s"`) carry their phonetic on the matching SST
  /// `<si>` instead, surfaced via `SharedStringTable::phonetic_for_entries`.
  /// Backed by a fresh entry in the caller-supplied `text_storage`, with
  /// the same lifetime contract as `value`'s text payload.
  std::string_view phonetic_text;
  /// Index into the workbook's `StylesTable::cell_xfs`, sourced from the
  /// `s=` attribute on the `<c>` element. Defaults to `0` (the default
  /// xf) when the attribute is absent.
  std::uint32_t xf_index = 0;
};

/// Parses one `<c>` element. `node` must be a valid `<c>` element node;
/// the caller is responsible for checking `node.name() == "c"`.
///
/// `text_storage` is the durable backing-store for any inline-string
/// payload the cell carries. The parser appends the decoded string to
/// the back of `text_storage` and points `ParsedCell::value`'s
/// `string_view` at that entry. The caller MUST use a container with
/// pointer / iterator stability across appends (e.g.
/// `std::deque<std::string>`); using a `std::vector` would invalidate
/// previously stored views.
///
/// Error paths (all surfaced as `kIoSheetCorrupt` with `<c>`-scoped
/// context):
///   * missing `r=` attribute
///   * malformed cell reference (whole-row, whole-col, range, sheet-qualified)
///   * `t=` value outside the supported set (b / e / n / s / str /
///     inlineStr / d are accepted; anything else fails)
///   * `<v>` content unparseable for the declared `t=`
Expected<ParsedCell, Error> parse_cell_element(const pugi::xml_node& node, std::deque<std::string>& text_storage);

/// Helper: parses an A1-style reference into 0-based `(row, col)`.
///
/// Examples: `"A1"` -> `(0, 0)`, `"Z9"` -> `(8, 25)`, `"AA1"` -> `(0, 26)`,
/// `"XFD1048576"` -> `(1048575, 16383)`.
///
/// Whole-row (`"1"`), whole-column (`"A"`), range (`"A1:B2"`), and
/// sheet-qualified (`"Sheet1!A1"`) shapes are rejected: this helper is
/// strictly for cell references. Out-of-range coordinates (column past
/// `XFD`, row past `1048576`, or row `0`) are also rejected.
Expected<std::pair<std::uint32_t, std::uint32_t>, Error> parse_a1(std::string_view ref);

/// String-view-only cell-payload decoder, shared between the DOM and
/// SAX read paths. Encodes the type-specific behaviour ("n", "b",
/// "e", "s", "str", "inlineStr", "d") of the OOXML `<c>` payload
/// without requiring a pugixml node.
///
/// Inputs:
///   * `t`                  - the `t=` attribute value, or empty (=> "n")
///   * `v_text`             - text content of `<v>` (entity-decoded)
///   * `value_present`      - true if a `<v>` or `<is>` child existed
///   * `is_inline_string`   - true if the value came from `<is>` rather
///                            than `<v>` (i.e. `t="inlineStr"`)
///   * `text_storage`       - durable backing-store for owned strings
///                            (the same `std::deque<std::string>` the
///                            DOM path uses).
/// Outputs (via `result`):
///   * `value`              - decoded `Value` per the OOXML rules
///   * `is_sst_index`       - true iff `t="s"`; the index is in
///                            `sst_index`
///
/// Returns `kIoSheetCorrupt` on any per-type contract violation
/// (unsupported `t=`, malformed `<v>` payload, ...).
Expected<ParsedCell, Error> decode_cell_payload(std::string_view t, std::string_view v_text, bool value_present,
                                                bool is_inline_string, std::deque<std::string>& text_storage);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_CELL_PARSER_H_
