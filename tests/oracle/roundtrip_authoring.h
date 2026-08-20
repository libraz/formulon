//
// Workbook-oracle test harness: the native half of a `roundtrip` case.
//
// Every other workbook suite builds its workbook *inside* Excel and asks
// how Excel lays it out; Formulon's own bytes never reach Excel on that
// path. A `roundtrip` case closes that gap. The capture side authors the
// fixture through Formulon's Python binding, hands the saved xlsx to
// Excel, and records what Excel resolved the settings to
// (`tools/oracle/print_roundtrip.py` plus the Windows driver). This
// header is the verifier's mirror image: it authors the same fixture
// through the same C ABI the binding wraps, saves it, and reports what the
// resulting package literally says.
//
// "Literally" is load-bearing. The observation is parsed out of the saved
// xlsx directly, not by loading it back through Formulon's own reader: a
// reader that undoes a writer mistake hands back the model the case
// started from, and that agreement is evidence about nothing. The
// manual-page-break `id` off-by-one was exactly this shape -- the writer
// added one, the reader subtracted one, every Formulon-only round trip
// looked clean, and only Excel could see that the file said 21 where the
// author meant 20. So the reader stays out of the loop here; it has its
// own coverage in the OOXML integration tests.
//
// Authoring through the C ABI rather than the engine types is equally
// deliberate. It is the surface the capture drove, so both halves make the
// same calls in the same order; going through `Workbook` directly would
// leave the print-setter layer -- where the raw XML fragments are
// synthesised -- untested.
//
// Declarative `roundtrip` block shape
// -----------------------------------
// A workbook case `spec` may carry an optional `roundtrip` object. Every
// member is optional except `sheet`, and each maps to one authoring call;
// an omitted member is not authored at all, so the file stays silent
// about it (which matters: see `stated` below).
//
//   {
//     "sheet": "Sheet1",                 // REQUIRED: sheet name or index
//     "page_setup": {                    // ints, as the C ABI takes them
//       "orientation": 2,                // 1 = portrait, 2 = landscape
//       "paper_size": 9,                 // OOXML paperSize code
//       "scale": 75,                     // percent
//       "fit_to_width": 1,               // pages
//       "fit_to_height": 1,              // pages
//       "fit_to_page": true              // <sheetPr><pageSetUpPr>
//     },
//     "page_margins": {"left": 0.5, ...},          // inches
//     "print_options": {"grid_lines": true, ...},  // booleans
//     "header_footer": {                 // section text carries Excel's
//       "odd_header": "&L left&C mid",   // &L / &C / &R codes verbatim
//       "different_odd_even": true, ...
//     },
//     "print_area": "A1:H60",            // an A1 range
//     "print_titles": {"repeat_rows": "1:2", "repeat_cols": "A:A"},
//     "row_breaks": [21, 41],            // 1-based Excel row numbers
//     "col_breaks": ["D"]                // column letters
//   }
//
// `row_breaks` / `col_breaks` place a break *before* the named track, so
// row 21 and column D become the 0-based ids 20 and 3 -- which is what
// OOXML's `<brk id>` stores and what Excel reports back.

#ifndef FORMULON_TESTS_ORACLE_ROUNDTRIP_AUTHORING_H_
#define FORMULON_TESTS_ORACLE_ROUNDTRIP_AUTHORING_H_

#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <vector>

#include "tests/oracle/json_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace tests {
namespace oracle {

/// The six header/footer sections plus the four `<headerFooter>` flags, as
/// the saved file states them. Section text is decoded (`&amp;` is already
/// an `&`) and keeps Excel's `&L` / `&C` / `&R` position codes, so
/// it can be compared against the golden's per-section capture by
/// reassembling that capture rather than by re-implementing Excel's split.
struct RoundtripHeaderFooter {
  bool different_odd_even = false;
  bool different_first = false;
  /// ECMA-376 §18.3.1.46 defaults both of these to true.
  bool scale_with_doc = true;
  bool align_with_margins = true;
  std::string odd_header;
  std::string odd_footer;
  std::string even_header;
  std::string even_footer;
  std::string first_header;
  std::string first_footer;
};

/// What the file Formulon just wrote says about itself.
///
/// The `*_stated` flags mean "this attribute is actually present in the
/// XML". They are not decoration. A `<pageSetup>` attribute the case never
/// authored is absent from the file, and Excel then answers from the
/// *printer's* defaults -- Letter paper on the capture host, whatever the
/// OOXML default may be -- so an unstated attribute is not evidence about
/// our writer and must not be compared.
struct RoundtripObservation {
  bool orientation_stated = false;
  /// `FM_ORIENTATION_*`, which coincides with Excel's `XlPageOrientation`
  /// (1 = portrait, 2 = landscape) so the golden needs no translation.
  std::uint32_t orientation = 0;
  bool paper_size_stated = false;
  std::uint32_t paper_size = 0;
  bool scale_stated = false;
  std::uint32_t scale = 0;
  bool fit_to_width_stated = false;
  std::uint32_t fit_to_width = 0;
  bool fit_to_height_stated = false;
  std::uint32_t fit_to_height = 0;
  bool fit_to_page = false;

  /// Inches, matching both the golden (which converts COM's points) and
  /// the OOXML attribute.
  double margin_left = 0.0;
  double margin_right = 0.0;
  double margin_top = 0.0;
  double margin_bottom = 0.0;
  double margin_header = 0.0;
  double margin_footer = 0.0;

  bool grid_lines = false;
  bool headings = false;
  bool horizontal_centered = false;
  bool vertical_centered = false;

  RoundtripHeaderFooter header_footer;

  /// A1 strings, empty when the sheet declares none -- the same shape the
  /// capture records from `PageSetup.PrintArea` / `PrintTitleRows`.
  std::string print_area;
  std::string print_title_rows;
  std::string print_title_cols;

  /// `<brk id>` values carrying `man="1"`, ascending. The id is already
  /// the 0-based index the break precedes, so it is compared against
  /// Excel's report unchanged. An automatic break Excel re-derived at the
  /// same position is not the break we wrote, so both sides keep only the
  /// manual ones.
  std::vector<std::uint32_t> manual_row_breaks;
  std::vector<std::uint32_t> manual_col_breaks;

  /// Length of the authored xlsx. The golden records the length and the
  /// SHA-256 of the bytes Excel opened; the digest cannot be reproduced
  /// here because the zip stamps each entry with the wall clock, but the
  /// length is content-determined and so still detects a golden that
  /// predates a change in what the writer emits.
  std::size_t xlsx_bytes = 0;
};

/// Authors a `roundtrip` case through the C ABI, saves it as xlsx, and
/// reports the print settings the saved package states.
///
/// `spec` is a workbook case `spec` object carrying a `sheets` block and a
/// `roundtrip` block (shape documented at the top of this header).
/// Case-level `column_widths` / `row_heights` are applied to the first
/// sheet, matching the capture-side authoring.
///
/// Errors (all `FormulonErrorCode::kInvalidArgument`):
///   * `spec` is not an object, or has no `roundtrip` block.
///   * the `roundtrip` block is missing `sheet`, or names an unknown one.
///   * a cell record, A1 address, or break token fails to parse.
///   * any C ABI call reports a non-OK status (the message names it).
///   * the saved package cannot be opened, or is missing the workbook
///     part, the relationship, or the worksheet the sheet index names.
Expected<RoundtripObservation, Error> observe_roundtrip_from_spec(const JsonValue& spec);

/// Materialises `spec`'s `sheets` block on a fresh workbook and returns the
/// resulting sheet name -> index map -- the same map
/// `observe_roundtrip_from_spec` resolves `roundtrip.sheet` through.
///
/// Indices follow the block's declaration order, and a case that declares
/// no sheets maps the workbook's default sheet to index 0. Both halves of
/// the round trip must agree on this map or they author and inspect
/// different sheets; the capture half's `_apply_sheets`
/// (tools/oracle/print_roundtrip.py) is its mirror.
Expected<std::map<std::string, std::uint32_t>, Error> roundtrip_sheet_indices(const JsonValue& spec);

}  // namespace oracle
}  // namespace tests
}  // namespace formulon

#endif  // FORMULON_TESTS_ORACLE_ROUNDTRIP_AUTHORING_H_
