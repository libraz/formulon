// Copyright 2026 libraz. Licensed under the MIT License.
//
// `Reference` is the parser's structural representation of an A1-style cell
// reference. It is shared by every node kind that needs to talk about a cell
// or a range endpoint: `NodeKind::Ref`, `NodeKind::ExternalRef`, and the two
// endpoints of `NodeKind::RangeOp`.
//
// The structure is layout-flat (no heap, no destructors) so it can live
// directly inside an `AstNode` payload. String views point into arena-interned
// storage owned by the producing parser; callers must keep that arena alive
// for as long as the reference is used.

#ifndef FORMULON_PARSER_REFERENCE_H_
#define FORMULON_PARSER_REFERENCE_H_

#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
namespace parser {

/// Structural A1-style cell reference.
///
/// All fields are populated by the parser; the AST factories merely copy the
/// value verbatim. `col` and `row` are 0-based to match Excel's internal
/// model (column A == 0, row 1 == 0). The `*_abs` flags carry the dollar-sign
/// markers needed for round-trip serialisation.
///
/// `sheet` is empty when the reference is local to the formula's owning
/// sheet. `sheet_quoted` is a hint for round-trip formatting only: it is set
/// to `true` when the original source spelled the sheet name in single quotes
/// (typically because the name contained spaces or other punctuation).
struct Reference {
  /// Sheet qualifier without surrounding quotes; empty when absent.
  std::string_view sheet;
  /// Round-trip hint: true if the source spelled the sheet name in quotes.
  bool sheet_quoted = false;
  /// 0-based column index. Excel's last column XFD == 16383.
  std::uint32_t col = 0;
  /// 0-based row index. Excel's last row 1048576 == 1048575.
  std::uint32_t row = 0;
  /// True if the column was prefixed with `$`.
  bool col_abs = false;
  /// True if the row was prefixed with `$`.
  bool row_abs = false;
};

/// Formats `r` as canonical A1 syntax.
///
/// Examples: `A1`, `$A$1`, `Sheet1!A1`, `'Sheet 1'!$A$1`. The sheet name is
/// wrapped in single quotes iff `r.sheet_quoted` is true; the parser is
/// responsible for setting that flag, and any embedded single quotes are
/// doubled per Excel's escaping convention.
std::string format_a1(const Reference& r);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_REFERENCE_H_
