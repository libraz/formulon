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

namespace {

// The eight color names a ja-JP format string may use, listed in the order
// the English UI names them (black, blue, cyan, green, magenta, red, white,
// yellow). Each is a single UTF-8 CJK character.
constexpr std::string_view kColorNames[] = {"黒", "青", "水", "緑", "紫", "赤", "白", "黄"};

// Localized spelling of the indexed `ColorN` form.
constexpr std::string_view kColorIndexPrefix = "色";

// Highest index the indexed form accepts; `色57` is rejected.
constexpr int kMaxColorIndex = 56;

bool starts_with(std::string_view s, std::string_view prefix) noexcept {
  return s.size() >= prefix.size() && s.substr(0, prefix.size()) == prefix;
}

// Reads one decimal digit at `body[i]`, accepting the ASCII digits and their
// full-width forms U+FF10..U+FF19 alike (Excel folds full-width ASCII across
// the whole format string, and the color index inherits that). Advances `i`
// past the digit and returns its value, or -1 when `i` is not on a digit.
int take_digit(std::string_view body, std::size_t& i) noexcept {
  const char c = body[i];
  if (c >= '0' && c <= '9') {
    ++i;
    return c - '0';
  }
  if (i + 2 < body.size() && static_cast<unsigned char>(body[i]) == 0xEFU &&
      static_cast<unsigned char>(body[i + 1]) == 0xBCU) {
    const unsigned char lo = static_cast<unsigned char>(body[i + 2]);
    if (lo >= 0x90U && lo <= 0x99U) {
      i += 3;
      return static_cast<int>(lo - 0x90U);
    }
  }
  return -1;
}

}  // namespace

bool is_color_specifier(std::string_view body) noexcept {
  for (const std::string_view name : kColorNames) {
    if (starts_with(body, name)) {
      return true;
    }
  }
  if (!starts_with(body, kColorIndexPrefix)) {
    return false;
  }
  // `色` then optional blanks then the index. Trailing bytes after the digit
  // run are ignored the same way they are after a name.
  std::size_t i = kColorIndexPrefix.size();
  while (i < body.size() && (body[i] == ' ' || body[i] == '\t')) {
    ++i;
  }
  int value = 0;
  int digits = 0;
  while (i < body.size()) {
    const int d = take_digit(body, i);
    if (d < 0) {
      break;
    }
    ++digits;
    // The run is greedy, so a value that has already overshot can only grow:
    // `色12345` is out of range, not index 1 with `2345` trailing.
    value = value * 10 + d;
    if (value > kMaxColorIndex) {
      return false;
    }
  }
  // Leading zeros are fine (`色001`), an index of zero is not.
  return digits > 0 && value >= 1;
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
