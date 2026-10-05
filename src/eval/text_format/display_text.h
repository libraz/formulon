//
// Cell display text: the string Excel shows for a value under a number format.
//
// The result is independent of column width: a `*x` repeat fill emits no
// characters and `_x` emits one space. A value that cannot be rendered in the
// format (for example a date serial outside the calendar range) is reported
// through `DisplayStatus` rather than by a width-dependent `#` run.

#ifndef FORMULON_EVAL_TEXT_FORMAT_DISPLAY_TEXT_H_
#define FORMULON_EVAL_TEXT_FORMAT_DISPLAY_TEXT_H_

#include <cstdint>
#include <string>
#include <string_view>

#include "value.h"

namespace formulon {

class Workbook;
class Sheet;
struct CellXf;
struct StylesTable;

namespace text_format {

/// Outcome of a display-text rendering.
enum class DisplayStatus : int {
  kOk = 0,
  /// The value does not fit the format (e.g. date serial out of range);
  /// `DisplayText::text` is `"########"`.
  kOverflow = 1,
  /// The format code is malformed; `DisplayText::text` is the General
  /// rendering of the value.
  kInvalidFormat = 2,
};

/// Display string plus the status it was produced under.
struct DisplayText {
  std::string text;
  DisplayStatus status = DisplayStatus::kOk;
};

/// Renders `value` as Excel displays it under number-format `code`.
/// Blank is empty, errors show their name, booleans show `TRUE`/`FALSE`,
/// text uses the format's text section when it has one, and numbers go
/// through `apply_format`. An empty `code` means General. `date1904`
/// selects the workbook date epoch.
DisplayText format_value_for_display(const Value& value, std::string_view code, bool date1904);

/// Returns the number-format code of `xf` as resolved against `styles`
/// (custom formats first, then the built-in table); empty for a null `xf`.
std::string_view number_format_code_for_xf(const StylesTable& styles, const CellXf* xf);

/// Display text of the cell at (`row`, `col`) of `sheet`, using the cell's xf
/// number format and the workbook's date system. An absent cell is empty.
DisplayText format_cell_for_display(const Workbook& workbook, const Sheet& sheet, std::uint32_t row, std::uint32_t col);

}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_DISPLAY_TEXT_H_
