// Copyright 2026 libraz. Licensed under the MIT License.
//
// Internal helpers shared between parser.cpp, parser_atoms.cpp, and
// parser_reference.cpp. Not a public header; do not include from outside
// src/parser/.

#ifndef FORMULON_PARSER_PARSER_DETAIL_H_
#define FORMULON_PARSER_PARSER_DETAIL_H_

#include <cstdint>

#include "parser/token.h"

namespace formulon {
namespace parser {
namespace detail {

// Binding-power constants. See parser.cpp header comment for the precedence
// table.
inline constexpr int kBpRange = 80;
inline constexpr int kBpUnaryPrefix = 70;
inline constexpr int kBpPostfixPercent = 60;
inline constexpr int kBpPow = 50;
inline constexpr int kBpMulDiv = 40;
inline constexpr int kBpAddSub = 30;
inline constexpr int kBpConcat = 20;
inline constexpr int kBpComparison = 10;
inline constexpr int kBpAtPrefix = 1;

inline constexpr std::uint32_t kMaxColumn = 16384;  // XFD
inline constexpr std::uint32_t kMaxRow = 1048576;   // 2^20

// ASCII helpers. Re-implemented locally to avoid depending on the tokenizer's
// privates and to keep the parser self-contained.
inline bool IsAsciiLetter(char c) noexcept {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}

inline bool IsAsciiDigit(char c) noexcept {
  return c >= '0' && c <= '9';
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
