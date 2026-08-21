//
// Internal helpers shared between parser.cpp, parser_atoms.cpp, and
// parser_reference.cpp. Not a public header; do not include from outside
// src/parser/.

#ifndef FORMULON_PARSER_PARSER_DETAIL_H_
#define FORMULON_PARSER_PARSER_DETAIL_H_

#include <cstdint>
#include <string_view>

#include "parser/token.h"

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

inline constexpr std::uint32_t kMaxColumn = 16384;  // XFD
inline constexpr std::uint32_t kMaxRow = 1048576;   // 2^20

// Ceiling on the `[N]` supporting-workbook index of a cross-workbook
// reference. Excel sets no documented limit; this one exists so a long
// digit run inside the brackets stops accumulating instead of
// overflowing, and sits far above any plausible number of external links
// so it never rejects a real file. An index past the workbook's actual
// link count resolves to `#REF!` at evaluation, not here.
inline constexpr std::uint32_t kMaxExternalBookIndex = 65535;

// ASCII helpers. Re-implemented locally to avoid depending on the tokenizer's
// privates and to keep the parser self-contained.
inline bool IsAsciiLetter(char c) noexcept {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}

inline bool IsAsciiDigit(char c) noexcept {
  return c >= '0' && c <= '9';
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
