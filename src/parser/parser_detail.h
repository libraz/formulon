//
// Internal helpers shared between parser.cpp, parser_atoms.cpp,
// parser_reference.cpp, and parser_let_lambda.cpp. Not a public header; do not
// include from outside src/parser/.

#ifndef FORMULON_PARSER_PARSER_DETAIL_H_
#define FORMULON_PARSER_PARSER_DETAIL_H_

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "parser/token.h"
#include "utils/a1_ref.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {
namespace detail {

// Binding-power constants. See parser.cpp header comment for the precedence
// table.
//
// Postfix `(` (immediately-invoked LAMBDA / curried call) sits above the
// spilled-range operator so `LAMBDA(x, x+1)(5)` and `LAMBDA(x, LAMBDA(y, x+y))
// (3)(4)` left-associate naturally without having to special-case the LHS
// shape. The Pratt loop already consumes a normal `Ident(args)` call site
// inside the atom dispatcher, so this rule only fires when the most recent
// LHS is an already-parsed expression (a Lambda atom, a parenthesised
// expression, or a previous LambdaCall).
inline constexpr int kBpPostfixCall = 95;
// Postfix `#` (spilled-range operator) sits above `:` so that `=A1:B2#`
// parses as `RangeOp(A1, SpillRef(B2))`; the `:` RHS shape check then
// rejects the SpillRef since spill anchors are single cells, never range
// endpoints.
inline constexpr int kBpPostfixHash = 90;
inline constexpr int kBpRange = 80;
// Space-as-intersection sits below `:` (range) and above prefix unary, matching
// Excel's precedence table. The token only retains binding power when it sits
// between two reference-shaped operands; see the whitespace-retention pass in
// `Parser::parse()`.
inline constexpr int kBpIntersect = 75;
inline constexpr int kBpUnaryPrefix = 70;
inline constexpr int kBpPostfixPercent = 60;
inline constexpr int kBpPow = 50;
inline constexpr int kBpMulDiv = 40;
inline constexpr int kBpAddSub = 30;
inline constexpr int kBpConcat = 20;
inline constexpr int kBpComparison = 10;
// `@` (implicit-intersection prefix) binds tighter than every arithmetic /
// comparison operator but looser than the reference operators (`:` range,
// space intersect, `#` spill). So `=@D3:D5*2` parses as `(@D3:D5)*2` — the
// `@` first binds the whole reference `D3:D5`, then the intersected scalar
// is multiplied — matching Mac Excel 365. Sitting above `kBpPostfixPercent`
// keeps `@A1%` / `@A1^2` as `(@A1)%` / `(@A1)^2`, and below `kBpIntersect`
// lets the operand absorb `:` / space so `@A1:B2` is `@(A1:B2)`.
inline constexpr int kBpAtPrefix = 65;

// Keep the historical parser-detail aliases for callers that use them, while
// sourcing the workbook grid limits from the shared A1 reference constants.
inline constexpr std::uint32_t kMaxColumn = a1::kMaxCols;
inline constexpr std::uint32_t kMaxRow = a1::kMaxRows;

// ASCII helpers. Re-implemented locally to avoid depending on the tokenizer's
// privates and to keep the parser self-contained.
inline bool IsAsciiLetter(char c) noexcept {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}

inline bool IsAsciiDigit(char c) noexcept {
  return c >= '0' && c <= '9';
}

// Byte length of the bare qualifier run starting at `pos` of `text`: ASCII
// letters, digits, `.`, `_` and non-ASCII characters other than U+3000 and
// U+FEFF. It is the shape of an unquoted local sheet qualifier and of the
// book and sheet parts of an unquoted cross-workbook qualifier.
inline std::size_t BareQualifierRunLength(std::string_view text, std::size_t pos) noexcept {
  std::size_t i = pos;
  while (i < text.size()) {
    const unsigned char c = static_cast<unsigned char>(text[i]);
    if (c < 0x80) {
      if (!IsAsciiLetter(static_cast<char>(c)) && !IsAsciiDigit(static_cast<char>(c)) && c != '.' && c != '_') {
        break;
      }
      ++i;
      continue;
    }
    const std::string_view rest = text.substr(i);
    if (rest.substr(0, 3) == "\xE3\x80\x80" || rest.substr(0, 3) == "\xEF\xBB\xBF") {
      break;
    }
    ++i;
  }
  return i - pos;
}

// True when `name` ends with a workbook file extension (ASCII
// case-insensitive): what marks `Book.xlsx!Name` as a book-scope name of
// another workbook rather than a name local to a sheet.
inline bool HasWorkbookExtension(std::string_view name) noexcept {
  constexpr std::string_view kExtensions[] = {".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx",
                                              ".xltm", ".xlt",  ".xlam", ".xla"};
  for (const std::string_view ext : kExtensions) {
    if (name.size() > ext.size() && strings::case_insensitive_eq(name.substr(name.size() - ext.size()), ext)) {
      return true;
    }
  }
  return false;
}

// Decodes a run of ASCII decimal digits into an unsigned magnitude,
// stopping accumulation the moment the running value exceeds `limit`.
// This guards `std::uint64_t` accumulation against wraparound on
// pathological inputs (e.g. a 20-digit row literal): once the value is
// already past any valid row/column limit, further digits cannot change
// that outcome, so there is no need to keep multiplying toward overflow.
// The caller only needs to know whether the result exceeds `limit`, which
// stays true for the returned value even though it is not the exact
// decoded magnitude of arbitrarily long digit runs. `digits` must contain
// only ASCII '0'-'9' characters; the caller validates that beforehand.
inline std::uint64_t DecodeDigitRunClamped(std::string_view digits, std::uint64_t limit) noexcept {
  std::uint64_t value = 0;
  for (char c : digits) {
    if (value > limit) {
      break;
    }
    value = value * 10u + static_cast<std::uint32_t>(c - '0');
  }
  return value;
}

// Builds a TextRange that spans from `a.start` (using a's line/column) to
// `b.end`. Used to attach a source span to a node assembled from children.
inline TextRange SpanRange(TextRange a, TextRange b) noexcept {
  TextRange r;
  r.start = a.start;
  r.end = b.end;
  r.line = a.line;
  r.column = a.column;
  return r;
}

}  // namespace detail
}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_PARSER_DETAIL_H_
