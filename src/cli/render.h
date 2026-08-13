//
// Shared value-rendering helpers used by `eval` and `dump` to produce
// stable, diff-friendly text for `fm_value_t`. Lives next to the CLI
// command sources because the formatting rules are CLI-policy, not
// engine policy.
//
// The format mirrors Excel's General format only: rendering is driven by
// the value kind alone, because `fm_value_t` carries no number format. A
// date-formatted cell therefore prints its raw serial, and currency,
// percentage and similar formats are likewise not applied.
//
// Per-kind rules:
//
//   * Number  — `format_double` (locale-independent shortest form).
//   * Bool    — `TRUE` / `FALSE`.
//   * Text    — verbatim UTF-8, no quoting (callers that need a
//               diff-friendly format with embedded text should layer
//               their own quoting on top).
//   * Error   — Excel display name (`#NAME?`, `#DIV/0!`, …).
//   * Blank   — empty string.
//   * Array / Ref / Lambda — placeholder string (the C ABI does not
//                            yet expose array payloads to the CLI).
//
// A1 column rendering follows Excel's bijective base-26 (`A`, `B`, …,
// `Z`, `AA`, `AB`, …, `XFD`).

#ifndef FORMULON_CLI_RENDER_H_
#define FORMULON_CLI_RENDER_H_

#include <cstdint>
#include <string>
#include <string_view>

#include "c_api/formulon_c.h"

namespace formulon {
namespace cli {

/// Appends the bijective base-26 column letters for `col` (0-based) to
/// `out`. `col == 0` yields `"A"`, `col == 25` yields `"Z"`,
/// `col == 26` yields `"AA"`, and so on through Excel's `XFD` cap at
/// `col == 16383`.
void append_column_letters(std::string& out, std::uint32_t col);

/// Returns the A1-style cell address (e.g. `"A1"`) for `(row, col)`,
/// both 0-based.
std::string format_a1(std::uint32_t row, std::uint32_t col);

/// Renders `v` as a single-line plain-text string. See the header
/// comment for the per-kind rules.
std::string render_value(const fm_value_t& v);

/// Returns `text` with JSON-style escapes for backslash, quote, and ASCII
/// control characters, but without surrounding quotes. Use for line-oriented
/// output such as `dump`, where embedded newlines must not create records.
std::string escape_single_line(std::string_view text);

/// Renders `v` as a JSON object: `{"kind": "...", "value": ...}`.
/// Strings are JSON-escaped; numbers use `format_double`.
std::string render_value_json(const fm_value_t& v);

}  // namespace cli
}  // namespace formulon

#endif  // FORMULON_CLI_RENDER_H_
