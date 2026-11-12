//
// Parser for REGEXREPLACE replacement templates (`$` and backslash layers).
// See `regex_replacement.h`.

#include "eval/regex_replacement.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/coerce.h"
#include "utils/case_map.h"
#include "utils/strings.h"
#include "utils/text_ops.h"

namespace formulon::eval::detail {

// --- REGEXREPLACE substitution -------------------------------------------
//
// The replacement string is parsed once per call into a token list and
// expanded per match. Two layers:
//
//   `$` layer       `$n` (one or two digits, greedy), `${n}`, `${name}`,
//                   `$name`, `$<name>`, `$&` (whole match), `` $` `` / `$'` (text before /
//                   after the match), `$_` (whole subject), `$$`.
//   backslash layer `\1`..`\99` and `\g<n>` / `\g<name>` (group references),
//                   `\n \t \r \e \a \f \v \b`, `\xHH`, `\x{H..}`, octal `\0HH`
//                   / `\ooo`, `\\ \$ \&`, `\u` (upper-case the next
//                   character), `\l` (lower-case the next), `\U`..`\E` /
//                   `\L`..`\E` (upper- / lower-case a run), `\Q`..`\E`
//                   (literal span).
//
// Anything else (a bare `$`, `$+`, `\0`, `\g1`, `\k<name>`, an
// unknown escape, a trailing `\`) is `#VALUE!`, as is a reference to a group
// the pattern lacks. Errors surface when a match is expanded, so a pattern
// that matches nothing leaves the subject unchanged. A group that did not
// participate expands to nothing.

namespace {

bool is_name_byte(char c) noexcept {
  return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_';
}

bool is_octal_byte(char c) noexcept {
  return c >= '0' && c <= '7';
}

// Appends the UTF-8 encoding of `cp`; false for a surrogate or a value past
// U+10FFFF.
bool append_code_point(std::string& out, std::uint32_t cp) {
  const std::string encoded = encode_utf8_codepoint(cp);
  out += encoded;
  return !encoded.empty();
}

// Reads `[0-9]{1,max_digits}` at `i`, advancing it.
std::uint32_t read_digits(std::string_view s, std::size_t& i, std::size_t max_digits) {
  std::uint32_t n = 0;
  for (std::size_t k = 0; k < max_digits && i < s.size() && strings::is_ascii_digit(s[i]); ++k, ++i) {
    n = n * 10U + static_cast<std::uint32_t>(s[i] - '0');
  }
  return n;
}

// A group reference spelled with a name or digits (`${..}`, `\g<..>`): all
// digits is a number, otherwise a name. Empty or non-word text is invalid.
ReplToken group_reference_token(std::string_view body) {
  ReplToken tok;
  tok.kind = ReplTokenKind::kInvalid;
  if (body.empty() || body.size() > 32U) {
    return tok;
  }
  bool all_digits = true;
  for (const char c : body) {
    if (!is_name_byte(c)) {
      return tok;
    }
    all_digits = all_digits && strings::is_ascii_digit(c);
  }
  if (all_digits) {
    std::size_t i = 0;
    tok.kind = ReplTokenKind::kGroup;
    tok.number = read_digits(body, i, body.size());
    return tok;
  }
  tok.kind = ReplTokenKind::kNamed;
  tok.text = std::string(body);
  return tok;
}

// Parses the `$` token starting at `i` (pointing at the `$`), advancing past it.
ReplToken parse_dollar(std::string_view r, std::size_t& i) {
  ReplToken tok;
  ++i;
  if (i >= r.size()) {
    tok.kind = ReplTokenKind::kInvalid;
    return tok;
  }
  const char c = r[i];
  if (strings::is_ascii_digit(c)) {
    tok.kind = ReplTokenKind::kGroup;
    tok.number = read_digits(r, i, 2U);
    return tok;
  }
  if (c == '{') {
    const std::size_t close = r.find('}', i);
    if (close == std::string_view::npos) {
      tok.kind = ReplTokenKind::kInvalid;
      return tok;
    }
    tok = group_reference_token(r.substr(i + 1U, close - i - 1U));
    i = close + 1U;
    return tok;
  }
  if (c == '<') {
    const std::size_t close = r.find('>', i);
    if (close == std::string_view::npos) {
      tok.kind = ReplTokenKind::kInvalid;
      return tok;
    }
    tok = group_reference_token(r.substr(i + 1U, close - i - 1U));
    if (tok.kind == ReplTokenKind::kGroup) {
      tok.kind = ReplTokenKind::kInvalid;  // `$<n>` names groups only.
    }
    i = close + 1U;
    return tok;
  }
  if (c == '_') {
    ++i;
    tok.kind = ReplTokenKind::kSubject;
    return tok;
  }
  if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_') {
    std::size_t end = i;
    while (end < r.size() && is_name_byte(r[end])) {
      ++end;
    }
    tok = group_reference_token(r.substr(i, end - i));
    i = end;
    return tok;
  }
  tok.kind = ReplTokenKind::kLiteral;
  ++i;
  switch (c) {
    case '$':
      tok.text = "$";
      return tok;
    case '&':
      tok.kind = ReplTokenKind::kGroup;
      return tok;
    case '`':
      tok.kind = ReplTokenKind::kBefore;
      return tok;
    case '\'':
      tok.kind = ReplTokenKind::kAfter;
      return tok;
    default:
      tok.kind = ReplTokenKind::kInvalid;
      return tok;
  }
}

// Parses the backslash token starting at `i` (pointing at the `\`), advancing
// past it.
ReplToken parse_backslash(std::string_view r, std::size_t& i) {
  ReplToken tok;
  ++i;
  if (i >= r.size()) {
    tok.kind = ReplTokenKind::kInvalid;
    return tok;
  }
  const char c = r[i++];
  const auto literal = [&tok](char ch) {
    tok.kind = ReplTokenKind::kLiteral;
    tok.text.assign(1U, ch);
  };
  switch (c) {
    case '\\':
    case '$':
    case '&':
      literal(c);
      return tok;
    case 'n':
      literal('\n');
      return tok;
    case 't':
      literal('\t');
      return tok;
    case 'r':
      literal('\r');
      return tok;
    case 'e':
      literal('\x1B');
      return tok;
    case 'a':
      literal('\a');
      return tok;
    case 'f':
      literal('\f');
      return tok;
    case 'v':
      literal('\v');
      return tok;
    case 'b':
      literal('\b');
      return tok;
    case 'u':
      tok.kind = ReplTokenKind::kUpperNext;
      return tok;
    case 'l':
      tok.kind = ReplTokenKind::kLowerNext;
      return tok;
    case 'U':
      tok.kind = ReplTokenKind::kUpperRun;
      return tok;
    case 'L':
      tok.kind = ReplTokenKind::kLowerRun;
      return tok;
    case 'E':
      tok.kind = ReplTokenKind::kCaseEnd;
      return tok;
    case 'Q': {
      const std::size_t end = r.find("\\E", i);
      tok.kind = ReplTokenKind::kLiteral;
      tok.text = std::string(r.substr(i, end == std::string_view::npos ? std::string_view::npos : end - i));
      i = end == std::string_view::npos ? r.size() : end + 2U;
      return tok;
    }
    case 'x': {
      std::uint32_t cp = 0;
      bool ok = false;
      if (i < r.size() && r[i] == '{') {
        std::size_t k = i + 1U;
        std::size_t digits = 0;
        while (k < r.size() && strings::hex_digit_value(r[k]) >= 0 && digits < 8U) {
          cp = cp * 16U + static_cast<std::uint32_t>(strings::hex_digit_value(r[k]));
          ++k;
          ++digits;
        }
        ok = digits > 0U && k < r.size() && r[k] == '}';
        i = ok ? k + 1U : i;
      } else {
        std::size_t digits = 0;
        while (digits < 2U && i < r.size() && strings::hex_digit_value(r[i]) >= 0) {
          cp = cp * 16U + static_cast<std::uint32_t>(strings::hex_digit_value(r[i]));
          ++i;
          ++digits;
        }
        ok = digits > 0U;
      }
      tok.kind = ReplTokenKind::kLiteral;
      if (!ok || !append_code_point(tok.text, cp)) {
        tok.kind = ReplTokenKind::kInvalid;
      }
      return tok;
    }
    case 'g': {
      if (i < r.size() && r[i] == '<') {
        const std::size_t close = r.find('>', i);
        if (close != std::string_view::npos) {
          tok = group_reference_token(r.substr(i + 1U, close - i - 1U));
          i = close + 1U;
          return tok;
        }
      }
      tok.kind = ReplTokenKind::kInvalid;
      return tok;
    }
    default:
      break;
  }
  if (strings::is_ascii_digit(c)) {
    --i;  // Back onto the first digit.
    const bool octal3 = is_octal_byte(c) && i + 2U < r.size() && is_octal_byte(r[i + 1U]) && is_octal_byte(r[i + 2U]);
    if (c == '0' || octal3) {
      std::uint32_t cp = 0;
      std::size_t digits = 0;
      while (digits < 3U && i < r.size() && is_octal_byte(r[i])) {
        cp = cp * 8U + static_cast<std::uint32_t>(r[i] - '0');
        ++i;
        ++digits;
      }
      tok.kind = ReplTokenKind::kLiteral;
      if (digits < 2U || !append_code_point(tok.text, cp)) {
        tok.kind = ReplTokenKind::kInvalid;  // `\0` alone is not a reference.
      }
      return tok;
    }
    tok.kind = ReplTokenKind::kGroup;
    tok.number = read_digits(r, i, 2U);
    return tok;
  }
  tok.kind = ReplTokenKind::kInvalid;
  return tok;
}

}  // namespace

std::vector<ReplToken> parse_replacement(std::string_view r) {
  std::vector<ReplToken> tokens;
  std::string pending;
  const auto flush = [&] {
    if (!pending.empty()) {
      ReplToken lit;
      lit.text = std::move(pending);
      pending.clear();
      tokens.push_back(std::move(lit));
    }
  };
  for (std::size_t i = 0; i < r.size();) {
    if (r[i] != '$' && r[i] != '\\') {
      pending.push_back(r[i++]);
      continue;
    }
    ReplToken tok = r[i] == '$' ? parse_dollar(r, i) : parse_backslash(r, i);
    if (tok.kind == ReplTokenKind::kLiteral) {
      pending += tok.text;
      continue;
    }
    flush();
    tokens.push_back(std::move(tok));
  }
  flush();
  return tokens;
}

namespace {

// Conservative byte bound for the cap: a UTF-16 unit never needs more than
// 4 UTF-8 bytes, so a required substitution output beyond this many bytes
// always exceeds the cap and can be rejected before any buffer growth.
// @limit: internal
constexpr std::size_t kMaxSubstituteOutputBytes = static_cast<std::size_t>(kExcelTextCapUnits) * 4U;

// Output of one expansion. A pending `\u` / `\l` maps the next character on its
// own; text appended under `\U` / `\L` is gathered and mapped as one span when
// the run ends, so a sigma inside it sees its neighbours.
class ReplacementWriter {
 public:
  explicit ReplacementWriter(std::string& out) : out_(out) {}

  void append(std::string_view text) {
    if (next_ != CaseMode::kNone && !text.empty()) {
      std::size_t bytes = 0;
      const std::uint32_t cp = decode_utf8_step(text, 0, &bytes);
      const bool valid = cp != 0xFFFDU || bytes != 1U;
      const std::uint32_t mapped = !valid                      ? cp
                                   : next_ == CaseMode::kUpper ? case_map_upper(cp)
                                                               : case_map_lower_isolated(cp);
      emit(mapped == cp ? std::string(text.substr(0, bytes)) : encode_utf8_codepoint(mapped));
      next_ = CaseMode::kNone;
      text.remove_prefix(bytes);
    }
    if (run_ == CaseMode::kNone) {
      out_.append(text);
    } else {
      span_.append(text);
    }
  }
  void next(bool upper) { next_ = upper ? CaseMode::kUpper : CaseMode::kLower; }
  void run(bool upper) {
    flush_span();
    run_ = upper ? CaseMode::kUpper : CaseMode::kLower;
  }
  void end_run() {
    flush_span();
    run_ = CaseMode::kNone;
  }
  // Ends an unterminated run at the end of the replacement.
  void finish() { end_run(); }
  // Bytes written so far, including a case span not yet flushed.
  std::size_t size() const noexcept { return out_.size() + span_.size(); }

 private:
  enum class CaseMode : std::uint8_t { kNone, kUpper, kLower };

  // A one-shot character sits outside the run, after the span gathered so far.
  void emit(const std::string& piece) {
    flush_span();
    out_.append(piece);
  }
  void flush_span() {
    if (run_ != CaseMode::kNone) {
      out_ += run_ == CaseMode::kUpper ? case_map_to_upper(span_) : case_map_to_lower(span_);
    }
    span_.clear();
  }

  std::string& out_;
  std::string span_;
  CaseMode run_ = CaseMode::kNone;
  CaseMode next_ = CaseMode::kNone;
};

// Resolves a group number from a name; -1 when absent or ambiguous.
long long group_number_for_name(const CompiledPattern& program, const std::string& name) {
  const int n = pcre2_substring_number_from_name(program.code(), reinterpret_cast<PCRE2_SPTR>(name.c_str()));
  return n < 0 ? -1 : n;
}

}  // namespace

// Appends the expansion of `tokens` for one match to `out`. `ovec` is the
// PCRE2 ovector of the match. False when a token is invalid or names a group
// the pattern lacks.
bool expand_replacement(const std::vector<ReplToken>& tokens, const CompiledPattern& program, std::string_view subject,
                        const PCRE2_SIZE* ovec, std::string& out) {
  ReplacementWriter writer(out);
  const auto group_text = [&](std::size_t group) -> std::string_view {
    if (ovec[2 * group] == PCRE2_UNSET) {
      return {};
    }
    return subject.substr(static_cast<std::size_t>(ovec[2 * group]),
                          static_cast<std::size_t>(ovec[2 * group + 1] - ovec[2 * group]));
  };
  for (const ReplToken& tok : tokens) {
    switch (tok.kind) {
      case ReplTokenKind::kLiteral:
        writer.append(tok.text);
        break;
      case ReplTokenKind::kGroup:
        if (tok.number > program.capture_count()) {
          return false;
        }
        writer.append(group_text(tok.number));
        break;
      case ReplTokenKind::kNamed: {
        const long long group = group_number_for_name(program, tok.text);
        if (group < 0) {
          return false;
        }
        writer.append(group_text(static_cast<std::size_t>(group)));
        break;
      }
      case ReplTokenKind::kBefore:
        writer.append(subject.substr(0, static_cast<std::size_t>(ovec[0])));
        break;
      case ReplTokenKind::kAfter:
        writer.append(subject.substr(static_cast<std::size_t>(ovec[1])));
        break;
      case ReplTokenKind::kSubject:
        writer.append(subject);
        break;
      case ReplTokenKind::kUpperNext:
        writer.next(/*upper=*/true);
        break;
      case ReplTokenKind::kLowerNext:
        writer.next(/*upper=*/false);
        break;
      case ReplTokenKind::kUpperRun:
        writer.run(/*upper=*/true);
        break;
      case ReplTokenKind::kLowerRun:
        writer.run(/*upper=*/false);
        break;
      case ReplTokenKind::kCaseEnd:
        writer.end_run();
        break;
      case ReplTokenKind::kInvalid:
        return false;
    }
    if (writer.size() > kMaxSubstituteOutputBytes) {
      return false;
    }
  }
  writer.finish();
  return out.size() <= kMaxSubstituteOutputBytes;
}

}  // namespace formulon::eval::detail
