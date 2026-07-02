// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Scanning primitives consumed by the format-string tokenizer
// (`number_format_tokenizer.cpp`) and the section classifier
// (`number_format_section.cpp`). These helpers are pure / `noexcept` and
// do not allocate beyond the bounded `std::string` scratch buffer used by
// `parse_cond_directive` to terminate the numeric tail for `std::strtod`.

#include "eval/text_format/number_format_scanner.h"

#include <cstddef>
#include <cstdlib>
#include <string>
#include <string_view>

#include "eval/number_parse.h"
#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

bool is_date_letter(char c) noexcept {
  return c == 'y' || c == 'Y' || c == 'm' || c == 'M' || c == 'd' || c == 'D' || c == 'h' || c == 'H' || c == 's' ||
         c == 'S';
}

std::size_t scan_run(std::string_view fmt, std::size_t& i, char letter) noexcept {
  const char upper = letter >= 'a' && letter <= 'z' ? static_cast<char>(letter - 32) : letter;
  const char lower = letter >= 'A' && letter <= 'Z' ? static_cast<char>(letter + 32) : letter;
  std::size_t start = i;
  while (i < fmt.size() && (fmt[i] == upper || fmt[i] == lower)) {
    ++i;
  }
  return i - start;
}

bool is_color_specifier(std::string_view body) noexcept {
  if (body.empty()) {
    return false;
  }
  auto eq_ci = [](std::string_view a, const char* b) {
    std::size_t n = 0;
    while (b[n] != '\0') {
      ++n;
    }
    if (a.size() != n) {
      return false;
    }
    for (std::size_t k = 0; k < n; ++k) {
      char ac = a[k];
      char bc = b[k];
      if (ac >= 'A' && ac <= 'Z') {
        ac = static_cast<char>(ac + 32);
      }
      if (bc >= 'A' && bc <= 'Z') {
        bc = static_cast<char>(bc + 32);
      }
      if (ac != bc) {
        return false;
      }
    }
    return true;
  };
  static const char* kNames[] = {"red", "blue", "green", "black", "white", "yellow", "cyan", "magenta"};
  for (const char* name : kNames) {
    if (eq_ci(body, name)) {
      return true;
    }
  }
  // `Color` followed by an integer in 1..56.
  if (body.size() < 6) {
    return false;
  }
  const std::string_view prefix = body.substr(0, 5);
  if (!eq_ci(prefix, "color")) {
    return false;
  }
  const std::string_view num = body.substr(5);
  if (num.empty()) {
    return false;
  }
  int value = 0;
  for (char ch : num) {
    if (ch < '0' || ch > '9') {
      return false;
    }
    value = value * 10 + (ch - '0');
    if (value > 56) {
      return false;
    }
  }
  return value >= 1 && value <= 56;
}

int parse_cond_directive(std::string_view body, CondOp* out_op, double* out_value) noexcept {
  std::size_t k = 0;
  while (k < body.size() && (body[k] == ' ' || body[k] == '\t')) {
    ++k;
  }
  if (k >= body.size()) {
    return 0;
  }
  CondOp op = CondOp::kNone;
  // Match longest operator first (so `>=`/`<=`/`<>` win over `>`/`<`/`=`).
  if (k + 1 < body.size() && body[k] == '>' && body[k + 1] == '=') {
    op = CondOp::kGe;
    k += 2;
  } else if (k + 1 < body.size() && body[k] == '<' && body[k + 1] == '=') {
    op = CondOp::kLe;
    k += 2;
  } else if (k + 1 < body.size() && body[k] == '<' && body[k + 1] == '>') {
    op = CondOp::kNe;
    k += 2;
  } else if (body[k] == '>') {
    op = CondOp::kGt;
    k += 1;
  } else if (body[k] == '<') {
    op = CondOp::kLt;
    k += 1;
  } else if (body[k] == '=') {
    op = CondOp::kEq;
    k += 1;
  } else {
    return 0;
  }
  // Parse the remaining bytes as a finite double. We copy into a stack buffer
  // so `std::strtod` sees a NUL terminator without depending on `body`'s
  // backing being NUL-terminated.
  if (k >= body.size()) {
    return -1;
  }
  std::string buf(body.substr(k));
  // Trim trailing whitespace.
  while (!buf.empty() && (buf.back() == ' ' || buf.back() == '\t')) {
    buf.pop_back();
  }
  if (buf.empty()) {
    return -1;
  }
  char* endp = nullptr;
  const double v = eval::parse_double_c_locale(buf.c_str(), &endp);
  if (endp == nullptr || endp == buf.c_str() || static_cast<std::size_t>(endp - buf.c_str()) != buf.size()) {
    return -1;
  }
  *out_op = op;
  *out_value = v;
  return 1;
}

int parse_dbnum_directive(std::string_view body) noexcept {
  // Expected form: `dbnumN` where N is `1`, `2`, or `3`.
  if (body.size() != 6) {
    return 0;
  }
  static const char kPrefix[5] = {'d', 'b', 'n', 'u', 'm'};
  for (std::size_t k = 0; k < 5; ++k) {
    char ch = body[k];
    if (ch >= 'A' && ch <= 'Z') {
      ch = static_cast<char>(ch + 32);
    }
    if (ch != kPrefix[k]) {
      return 0;
    }
  }
  const char d = body[5];
  if (d == '1') {
    return 1;
  }
  if (d == '2') {
    return 2;
  }
  if (d == '3') {
    return 3;
  }
  return 0;
}

bool is_date_tok(Tok t) noexcept {
  switch (t) {
    case Tok::DateY2:
    case Tok::DateY4:
    case Tok::DateMOrMin:
    case Tok::DateMMM:
    case Tok::DateMMMM:
    case Tok::DateMMMMM:
    case Tok::DateD:
    case Tok::DateDD:
    case Tok::DateDDD:
    case Tok::DateDDDD:
    case Tok::DateH:
    case Tok::DateHH:
    case Tok::DateS:
    case Tok::DateSS:
    case Tok::DateElapsedH:
    case Tok::DateElapsedM:
    case Tok::DateElapsedS:
    case Tok::AmPm:
    case Tok::AP:
    case Tok::EraG:
    case Tok::EraGG:
    case Tok::EraGGG:
    case Tok::EraE:
    case Tok::EraEE:
    case Tok::DateAaa:
    case Tok::DateAaaa:
    case Tok::DateM:
    case Tok::DateMM:
    case Tok::DateMin:
    case Tok::DateMMMin:
      return true;
    default:
      return false;
  }
}

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon
