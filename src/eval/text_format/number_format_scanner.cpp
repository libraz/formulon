//
// Scanning primitives consumed by the format-string tokenizer
// (`number_format_tokenizer.cpp`) and the section classifier
// (`number_format_section.cpp`). The scanning helpers are pure / `noexcept`;
// the public normalizer is the one deliberate allocation point that owns the
// syntax-normalized format view for a single TEXT() call.

#include "eval/text_format/number_format_scanner.h"

#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <string>
#include <string_view>

#include "eval/number_parse.h"
#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

namespace {

struct Utf8Scalar {
  std::uint32_t codepoint = 0;
  std::size_t width = 1;
  bool valid = false;
};

Utf8Scalar decode_utf8(std::string_view fmt, std::size_t i) noexcept {
  if (i >= fmt.size()) {
    return {};
  }
  const auto byte = [&](std::size_t offset) -> std::uint8_t {
    return static_cast<std::uint8_t>(static_cast<unsigned char>(fmt[offset]));
  };
  const std::uint8_t first = byte(i);
  if (first < 0x80U) {
    return {first, 1, true};
  }

  // Reject overlong encodings, UTF-16 surrogate code points, and values above
  // U+10FFFF. On failure the caller copies just the original lead byte and
  // continues, which is the safest recovery for a user-supplied format.
  if (first >= 0xC2U && first <= 0xDFU && i + 1 < fmt.size()) {
    const std::uint8_t second = byte(i + 1);
    if ((second & 0xC0U) == 0x80U) {
      return {static_cast<std::uint32_t>(first & 0x1FU) << 6U | (second & 0x3FU), 2, true};
    }
  } else if (first >= 0xE0U && first <= 0xEFU && i + 2 < fmt.size()) {
    const std::uint8_t second = byte(i + 1);
    const std::uint8_t third = byte(i + 2);
    if ((second & 0xC0U) == 0x80U && (third & 0xC0U) == 0x80U && (first != 0xE0U || second >= 0xA0U) &&
        (first != 0xEDU || second <= 0x9FU)) {
      return {static_cast<std::uint32_t>(first & 0x0FU) << 12U | static_cast<std::uint32_t>(second & 0x3FU) << 6U |
                  (third & 0x3FU),
              3, true};
    }
  } else if (first >= 0xF0U && first <= 0xF4U && i + 3 < fmt.size()) {
    const std::uint8_t second = byte(i + 1);
    const std::uint8_t third = byte(i + 2);
    const std::uint8_t fourth = byte(i + 3);
    if ((second & 0xC0U) == 0x80U && (third & 0xC0U) == 0x80U && (fourth & 0xC0U) == 0x80U &&
        (first != 0xF0U || second >= 0x90U) && (first != 0xF4U || second <= 0x8FU)) {
      return {static_cast<std::uint32_t>(first & 0x07U) << 18U | static_cast<std::uint32_t>(second & 0x3FU) << 12U |
                  static_cast<std::uint32_t>(third & 0x3FU) << 6U | (fourth & 0x3FU),
              4, true};
    }
  }
  return {first, 1, false};
}

bool is_fullwidth_ascii(std::uint32_t cp) noexcept {
  return cp >= 0xFF01U && cp <= 0xFF5EU;
}

char fullwidth_ascii(std::uint32_t cp) noexcept {
  return static_cast<char>(cp - 0xFEE0U);
}

bool is_date_codepoint(std::uint32_t cp) noexcept {
  if (!is_fullwidth_ascii(cp)) {
    return false;
  }
  const char c = fullwidth_ascii(cp);
  return c == 'y' || c == 'Y' || c == 'm' || c == 'M' || c == 'd' || c == 'D' || c == 'h' || c == 'H' || c == 's' ||
         c == 'S' || c == 'g' || c == 'G' || c == 'e';
}

char date_codepoint_letter(std::uint32_t cp) noexcept {
  if (cp >= 'A' && cp <= 'Z') {
    return static_cast<char>(cp + ('a' - 'A'));
  }
  if (cp >= 'a' && cp <= 'z') {
    return static_cast<char>(cp);
  }
  if (!is_fullwidth_ascii(cp)) {
    return '\0';
  }
  const char c = fullwidth_ascii(cp);
  switch (c) {
    case 'y':
    case 'Y':
    case 'm':
    case 'M':
    case 'd':
    case 'D':
    case 'h':
    case 'H':
    case 's':
    case 'S':
    case 'g':
    case 'G':
    case 'e':
    case 'E':
      return c >= 'A' && c <= 'Z' ? static_cast<char>(c + ('a' - 'A')) : c;
    default:
      return '\0';
  }
}

bool is_latin_codepoint(std::uint32_t cp) noexcept {
  if ((cp >= 'A' && cp <= 'Z') || (cp >= 'a' && cp <= 'z')) {
    return true;
  }
  if (!is_fullwidth_ascii(cp)) {
    return false;
  }
  const char c = fullwidth_ascii(cp);
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}

std::size_t previous_scalar_begin(std::string_view fmt, std::size_t i) noexcept {
  if (i == 0) {
    return 0;
  }
  std::size_t begin = i - 1;
  while (begin > 0) {
    const auto byte = static_cast<unsigned char>(fmt[begin]);
    if ((byte & 0xC0U) != 0x80U) {
      break;
    }
    --begin;
  }
  return begin;
}

bool should_fold_date_codepoint(std::string_view fmt, std::size_t i, std::size_t width, std::uint32_t cp) noexcept {
  const char letter = date_codepoint_letter(cp);
  if (letter == '\0') {
    return false;
  }
  if (i > 0) {
    const Utf8Scalar previous = decode_utf8(fmt, previous_scalar_begin(fmt, i));
    if (previous.valid && is_latin_codepoint(previous.codepoint) &&
        date_codepoint_letter(previous.codepoint) != letter) {
      return false;
    }
  }
  if (i + width < fmt.size()) {
    const Utf8Scalar next = decode_utf8(fmt, i + width);
    if (next.valid && is_latin_codepoint(next.codepoint) && date_codepoint_letter(next.codepoint) != letter) {
      return false;
    }
  }
  return true;
}

bool is_a_codepoint(std::uint32_t cp) noexcept {
  return cp == 'a' || cp == 'A' || cp == 0xFF41U || cp == 0xFF21U;
}

char ascii_lower(char c) noexcept {
  return c >= 'A' && c <= 'Z' ? static_cast<char>(c + ('a' - 'A')) : c;
}

bool scalar_matches_ascii(std::string_view fmt, std::size_t i, char expected, std::size_t* width) noexcept {
  const Utf8Scalar scalar = decode_utf8(fmt, i);
  if (!scalar.valid) {
    return false;
  }
  if (expected == '/') {
    if (scalar.codepoint == '/') {
      *width = scalar.width;
      return true;
    }
    if (scalar.codepoint == 0xFF0FU) {
      *width = scalar.width;
      return true;
    }
    return false;
  }
  if (scalar.codepoint < 0x80U) {
    if (ascii_lower(static_cast<char>(scalar.codepoint)) == ascii_lower(expected)) {
      *width = scalar.width;
      return true;
    }
    return false;
  }
  if (is_fullwidth_ascii(scalar.codepoint) && ascii_lower(fullwidth_ascii(scalar.codepoint)) == ascii_lower(expected)) {
    *width = scalar.width;
    return true;
  }
  return false;
}

// Returns the byte end of an ASCII/full-width spelling of `word` at `start`.
// The caller can then append the canonical ASCII word. This is used only for
// markers whose complete spelling is unambiguous; arbitrary full-width Latin
// text is intentionally left untouched.
bool match_marker(std::string_view fmt, std::size_t start, std::string_view word, std::size_t* end) noexcept {
  std::size_t i = start;
  for (const char expected : word) {
    std::size_t width = 0;
    if (!scalar_matches_ascii(fmt, i, expected, &width)) {
      return false;
    }
    i += width;
  }
  *end = i;
  return true;
}

bool match_dbnum(std::string_view fmt, std::size_t start, std::size_t* end, std::uint32_t* digit_codepoint) noexcept {
  std::size_t i = start;
  if (!match_marker(fmt, i, "dbnum", &i)) {
    return false;
  }
  const Utf8Scalar digit = decode_utf8(fmt, i);
  if (!digit.valid || digit.codepoint < '1' || digit.codepoint > '3') {
    if (!is_fullwidth_ascii(digit.codepoint)) {
      return false;
    }
    const char c = fullwidth_ascii(digit.codepoint);
    if (c < '1' || c > '3') {
      return false;
    }
  }
  i += digit.width;
  *end = i;
  *digit_codepoint = digit.codepoint;
  return true;
}

char syntax_ascii(std::uint32_t cp) noexcept {
  if (cp >= 0xFF10U && cp <= 0xFF19U) {
    return static_cast<char>('0' + (cp - 0xFF10U));
  }
  if (!is_fullwidth_ascii(cp)) {
    return '\0';
  }
  const char c = fullwidth_ascii(cp);
  switch (c) {
    case '#':
    case '?':
    case '.':
    case ',':
    case '%':
    case ';':
    case '[':
    case ']':
    case '<':
    case '>':
    case '=':
    case '+':
    case '-':
    case ':':
    case '/':
    case '\\':
    case '!':
    case '_':
    case '*':
    case '@':
    case '$':
      return c;
    default:
      return '\0';
  }
}

void append_raw(std::string& out, std::string_view fmt, std::size_t begin, std::size_t end) {
  out.append(fmt.data() + begin, end - begin);
}

}  // namespace

std::size_t utf8_scalar_width(std::string_view fmt, std::size_t i) noexcept {
  return decode_utf8(fmt, i).width;
}

std::string normalize_ja_jp_format_syntax(std::string_view fmt) {
  std::string out;
  out.reserve(fmt.size());
  std::size_t i = 0;
  bool in_bracket = false;
  while (i < fmt.size()) {
    const Utf8Scalar scalar = decode_utf8(fmt, i);
    if (!scalar.valid) {
      out.push_back(fmt[i]);
      ++i;
      continue;
    }
    const std::size_t width = scalar.width;
    const std::uint32_t cp = scalar.codepoint;

    // A quoted string is opaque. In particular, full-width digits and
    // punctuation in a quoted label must remain exactly as supplied.
    if (cp == '"') {
      out.push_back('"');
      i += width;
      while (i < fmt.size()) {
        const Utf8Scalar quoted = decode_utf8(fmt, i);
        append_raw(out, fmt, i, i + quoted.width);
        i += quoted.width;
        if (quoted.valid && quoted.codepoint == '"') {
          break;
        }
      }
      continue;
    }

    // Escapes, underscore skips, and asterisk fills reserve/consume one
    // complete UTF-8 scalar. Their payload is always literal, even when it
    // happens to be a full-width syntax character.
    const char mapped_syntax = syntax_ascii(cp);
    if (cp == '\\' || cp == '!' || cp == '_' || cp == '*' || mapped_syntax == '\\' || mapped_syntax == '!' ||
        mapped_syntax == '_' || mapped_syntax == '*') {
      out.push_back(mapped_syntax != '\0' ? mapped_syntax : static_cast<char>(cp));
      i += width;
      if (i < fmt.size()) {
        const Utf8Scalar payload = decode_utf8(fmt, i);
        append_raw(out, fmt, i, i + payload.width);
        i += payload.width;
      }
      continue;
    }

    if (in_bracket) {
      // `[DBNum1]` is the one bracket directive whose full-width Latin
      // spelling is syntax. Keep other full-width Latin words opaque.
      std::size_t marker_end = i;
      std::uint32_t digit_codepoint = 0;
      if (match_dbnum(fmt, i, &marker_end, &digit_codepoint)) {
        out.append("DBNum");
        const char digit_ascii = digit_codepoint >= '1' && digit_codepoint <= '3' ? static_cast<char>(digit_codepoint)
                                                                                  : fullwidth_ascii(digit_codepoint);
        out.push_back(digit_ascii);
        i = marker_end;
        continue;
      }
    } else {
      // AM/PM and A/P are multi-scalar markers. Match before considering a
      // lone full-width A as syntax; Excel treats that lone A as literal text.
      std::size_t marker_end = i;
      if (match_marker(fmt, i, "AM/PM", &marker_end)) {
        out.append("AM/PM");
        i = marker_end;
        continue;
      }
      if (match_marker(fmt, i, "A/P", &marker_end)) {
        out.append("A/P");
        i = marker_end;
        continue;
      }

      // Weekday `aaa`/`aaaa` is the only full-width A run accepted as a
      // date token. A or AA remain literal full-width text.
      if (is_a_codepoint(cp)) {
        std::size_t run_end = i;
        std::size_t run = 0;
        while (run_end < fmt.size()) {
          const Utf8Scalar a = decode_utf8(fmt, run_end);
          if (!a.valid || !is_a_codepoint(a.codepoint)) {
            break;
          }
          run_end += a.width;
          ++run;
        }
        if (run >= 3) {
          out.append(run, 'a');
          i = run_end;
          continue;
        }
      }
    }

    // A full-width scientific E is syntax only when it follows a digit
    // placeholder. This prevents a literal full-width E from being silently
    // converted while still accepting `0.00Ｅ＋00`.
    if (!in_bracket && (cp == 0xFF25U || cp == 'E') && !out.empty() &&
        (out.back() == '0' || out.back() == '#' || out.back() == '?')) {
      std::size_t sign_end = i + width;
      const Utf8Scalar sign = decode_utf8(fmt, sign_end);
      if (sign.valid &&
          (sign.codepoint == '+' || sign.codepoint == '-' || sign.codepoint == 0xFF0BU || sign.codepoint == 0xFF0DU)) {
        out.push_back('E');
        const char sign_ascii = sign.codepoint == 0xFF0BU   ? '+'
                                : sign.codepoint == 0xFF0DU ? '-'
                                                            : static_cast<char>(sign.codepoint);
        out.push_back(sign_ascii);
        i = sign_end + sign.width;
        continue;
      }
    }

    // Outside opaque payloads, the ASCII punctuation and digits in the
    // full-width block are format syntax. Bracket state is tracked after the
    // fold so full-width brackets can delimit a normalised bracket body.
    const char syntax = syntax_ascii(cp);
    if (syntax != '\0') {
      out.push_back(syntax);
      if (syntax == '[') {
        in_bracket = true;
      } else if (syntax == ']') {
        in_bracket = false;
      }
      i += width;
      continue;
    }

    if (cp == '[') {
      out.push_back('[');
      in_bracket = true;
      i += width;
      continue;
    }
    if (cp == ']') {
      out.push_back(']');
      in_bracket = false;
      i += width;
      continue;
    }

    // Full-width date/time runs (`ｙ`, `ｍ`, `ｄ`, `ｈ`, `ｓ`, `ｇ`, `ｅ`) are
    // canonicalised one scalar at a time. Other full-width Latin letters,
    // including the lone `Ａ` literal observed in Excel, are preserved.
    if (is_date_codepoint(cp) && should_fold_date_codepoint(fmt, i, width, cp)) {
      const char c = fullwidth_ascii(cp);
      out.push_back(ascii_lower(c));
      i += width;
      continue;
    }

    append_raw(out, fmt, i, i + width);
    i += width;
  }
  return out;
}

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
