// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the M2.2 Excel formula tokenizer. See the header for
// the high-level contract. Key implementation choices:
//
//   * UTF-8 decoding is hand-rolled (no `<codecvt>` / ICU): we only need
//     byte-length and UTF-16 code-unit count per codepoint.
//   * Numbers go through `std::strtod` for M2; M4 swaps in `fast_float`
//     once 1-bit Excel parity becomes mandatory.
//   * String / quoted-sheet-name escapes expand into the tokenizer's arena
//     so token views remain stable for the lifetime of the tokenizer.
//   * The spilled-range `#` operator is disambiguated from `#error!` by
//     lookahead: `#` immediately after a CellRef with no whitespace gap
//     becomes `Hash`; otherwise we try to match an error-literal catalog.
//
// The file intentionally keeps each scanner small and self-contained so
// future harderning (e.g. dot-notation ranges, structured-ref keywords) can
// slot in without touching the main dispatch loop.

#include "parser/tokenizer.h"

#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <string>
#include <string_view>

namespace formulon {
namespace parser {

namespace {

// UTF-8 BOM and UTF-16 LE BOM (as 2 bytes) for the source-prefix check.
constexpr std::string_view kUtf8Bom = "\xEF\xBB\xBF";
constexpr std::string_view kUtf16LeBom = "\xFF\xFE";

// Excel 365 workbook limits.
constexpr std::uint32_t kMaxColumn = 16384;  // XFD
constexpr std::uint32_t kMaxRow = 1048576;   // 2^20

// Lookup for the 17 canonical Excel error literals. Kept sorted by the
// longest-prefix-first ordering so scan_error_literal can commit on first
// match without ambiguity.
struct ErrorLiteralEntry {
  std::string_view text;
  ErrorCode code;
};

constexpr ErrorLiteralEntry kErrorLiterals[] = {
    {"#GETTING_DATA", ErrorCode::GettingData},  // 13
    {"#EXTERNAL!", ErrorCode::External},        // 10
    {"#BLOCKED!", ErrorCode::Blocked},          //  9
    {"#CONNECT!", ErrorCode::Connect},          //  9
    {"#UNKNOWN!", ErrorCode::Unknown},          //  9
    {"#PYTHON!", ErrorCode::Python},            //  8
    {"#DIV/0!", ErrorCode::Div0},               //  7
    {"#VALUE!", ErrorCode::Value},              //  7
    {"#SPILL!", ErrorCode::Spill},              //  7
    {"#FIELD!", ErrorCode::Field},              //  7
    {"#BUSY!", ErrorCode::Busy},                //  6
    {"#CALC!", ErrorCode::Calc},                //  6
    {"#NAME?", ErrorCode::Name},                //  6
    {"#NULL!", ErrorCode::Null},                //  6
    {"#NUM!", ErrorCode::Num},                  //  5
    {"#REF!", ErrorCode::Ref},                  //  5
    {"#N/A", ErrorCode::NA},                    //  4
};

// ASCII case-insensitive prefix match over a known-ASCII catalog entry.
bool ieq_prefix(std::string_view haystack, std::string_view needle) noexcept {
  if (haystack.size() < needle.size()) {
    return false;
  }
  for (std::size_t i = 0; i < needle.size(); ++i) {
    char a = haystack[i];
    char b = needle[i];
    if (a >= 'a' && a <= 'z') {
      a = static_cast<char>(a - ('a' - 'A'));
    }
    if (b >= 'a' && b <= 'z') {
      b = static_cast<char>(b - ('a' - 'A'));
    }
    if (a != b) {
      return false;
    }
  }
  return true;
}

// Converts a column-letter run (1..3 ASCII letters) to a 1-based column index.
// Returns 0 on overflow past Excel's 16384 column cap.
std::uint32_t column_letters_to_index(std::string_view letters) noexcept {
  std::uint32_t v = 0;
  for (char ch : letters) {
    if (ch >= 'a' && ch <= 'z') {
      ch = static_cast<char>(ch - ('a' - 'A'));
    }
    if (ch < 'A' || ch > 'Z') {
      return 0;
    }
    v = v * 26u + static_cast<std::uint32_t>(ch - 'A' + 1);
    if (v > kMaxColumn) {
      return 0;
    }
  }
  return v;
}

// Converts a row-digit run to an integer. Returns 0 on overflow past the
// 1048576 cap.
std::uint32_t row_digits_to_index(std::string_view digits) noexcept {
  std::uint64_t v = 0;
  for (char ch : digits) {
    if (ch < '0' || ch > '9') {
      return 0;
    }
    v = v * 10u + static_cast<std::uint32_t>(ch - '0');
    if (v > kMaxRow) {
      return 0;
    }
  }
  return static_cast<std::uint32_t>(v);
}

}  // namespace

// ---------------------------------------------------------------------------
// Static helpers
// ---------------------------------------------------------------------------

bool Tokenizer::is_ascii_letter(char c) noexcept {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
}

bool Tokenizer::is_ascii_digit(char c) noexcept {
  return c >= '0' && c <= '9';
}

bool Tokenizer::is_ident_start_byte(unsigned char c) noexcept {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_' || c >= 0x80;
}

bool Tokenizer::is_ident_cont_byte(unsigned char c) noexcept {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_' || c == '.' ||
         c >= 0x80;
}

bool Tokenizer::is_formula_whitespace(char c) noexcept {
  // Excel accepts ASCII space / tab / CR / LF as formula whitespace. The
  // full-width space U+3000 is intentionally *not* accepted here; it is
  // flagged as InvalidCharacter by the main loop.
  return c == ' ' || c == '\t' || c == '\r' || c == '\n';
}

bool Tokenizer::is_bool_word(std::string_view word, bool* out) noexcept {
  if (word.size() == 4) {
    if (ieq_prefix(word, "TRUE")) {
      *out = true;
      return true;
    }
  } else if (word.size() == 5) {
    if (ieq_prefix(word, "FALSE")) {
      *out = false;
      return true;
    }
  }
  return false;
}

bool Tokenizer::match_error_literal(std::string_view run, ErrorCode* out) noexcept {
  for (const auto& e : kErrorLiterals) {
    if (run.size() == e.text.size() && ieq_prefix(run, e.text)) {
      *out = e.code;
      return true;
    }
  }
  return false;
}

bool Tokenizer::looks_like_cellref(std::string_view run, bool* letters_only) noexcept {
  *letters_only = false;
  std::size_t i = 0;
  if (i < run.size() && run[i] == '$') {
    ++i;
  }
  const std::size_t letters_begin = i;
  while (i < run.size() && is_ascii_letter(run[i])) {
    ++i;
  }
  const std::size_t letters_len = i - letters_begin;
  if (letters_len == 0 || letters_len > 3) {
    return false;
  }
  if (i < run.size() && run[i] == '$') {
    ++i;
  }
  const std::size_t digits_begin = i;
  while (i < run.size() && is_ascii_digit(run[i])) {
    ++i;
  }
  const std::size_t digits_len = i - digits_begin;
  if (i != run.size()) {
    return false;
  }
  // Check column limit.
  if (column_letters_to_index(run.substr(letters_begin, letters_len)) == 0) {
    return false;
  }
  if (digits_len == 0) {
    // Letter-only with an optional leading `$`: treated as identifier-like
    // in M2.2. The caller re-emits as Ident.
    if (run[0] == '$' || run.back() == '$') {
      // `$A` or `A$` without digits is still malformed.
      return false;
    }
    *letters_only = true;
    return false;
  }
  if (digits_len > 7) {
    return false;
  }
  if (row_digits_to_index(run.substr(digits_begin, digits_len)) == 0) {
    return false;
  }
  return true;
}

// ---------------------------------------------------------------------------
// Construction and public API
// ---------------------------------------------------------------------------

Tokenizer::Tokenizer(std::string_view source, TokenizerOptions opts) noexcept
    : source_(source), opts_(opts), arena_(4096) {}

const std::vector<Token>& Tokenizer::tokens() {
  if (done_) {
    return tokens_;
  }
  done_ = true;

  // BOM handling: strip a leading UTF-8 or UTF-16-LE BOM silently.
  if (source_.size() >= kUtf8Bom.size() && std::memcmp(source_.data(), kUtf8Bom.data(), kUtf8Bom.size()) == 0) {
    byte_pos_ = kUtf8Bom.size();
  } else if (source_.size() >= kUtf16LeBom.size() &&
             std::memcmp(source_.data(), kUtf16LeBom.data(), kUtf16LeBom.size()) == 0) {
    byte_pos_ = kUtf16LeBom.size();
  }
  // Note: BOM bytes are *not* counted in UTF-16 offsets, so utf16_pos_ stays
  // at 0 here. This matches Excel's "BOM is invisible" contract.

  while (byte_pos_ < source_.size()) {
    // Enforce the UTF-16 length cap. The comparison is against the offset
    // *before* consuming the next codepoint, which means the last accepted
    // token ends at exactly max_formula_length_utf16 code units.
    if (utf16_pos_ >= opts_.max_formula_length_utf16) {
      if (!truncated_) {
        const std::size_t err_start = byte_pos_;
        // Advance to end of source so the error lexeme covers the overflow.
        byte_pos_ = source_.size();
        record_error(LexerErrorCode::ExcessiveLength, err_start);
        truncated_ = true;
      }
      break;
    }

    const unsigned char c = static_cast<unsigned char>(source_[byte_pos_]);

    // Mid-input BOM: consume the 3 bytes (or 2 for UTF-16 LE) and record an
    // error. We handle UTF-8 BOM here because the sentinel is unambiguous.
    if (c == 0xEF && byte_pos_ + 2 < source_.size() && static_cast<unsigned char>(source_[byte_pos_ + 1]) == 0xBB &&
        static_cast<unsigned char>(source_[byte_pos_ + 2]) == 0xBF) {
      const std::size_t err_start = byte_pos_;
      byte_pos_ += 3;
      column_ += 1;  // zero-width BOM counts as one column for diagnostics.
      record_error(LexerErrorCode::InvalidCharacter, err_start);
      continue;
    }

    // Whitespace.
    if (is_formula_whitespace(static_cast<char>(c))) {
      scan_whitespace();
      continue;
    }

    // Non-whitespace "space-like" bytes that Excel doesn't accept. Fullwidth
    // space U+3000 (0xE3 0x80 0x80) falls into this bucket because we must
    // not confuse it with the intersection operator.
    if (c == 0xE3 && byte_pos_ + 2 < source_.size() && static_cast<unsigned char>(source_[byte_pos_ + 1]) == 0x80 &&
        static_cast<unsigned char>(source_[byte_pos_ + 2]) == 0x80) {
      const std::size_t err_start = byte_pos_;
      // Advance by one codepoint (3 bytes) and record.
      advance_one();
      record_error(LexerErrorCode::InvalidCharacter, err_start);
      continue;
    }

    // Dispatch on the leading byte.
    switch (c) {
      case '"':
        scan_string();
        continue;
      case '\'':
        scan_quoted_sheet_name();
        continue;
      case '#':
        scan_error_literal();
        continue;
      case '(': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::LParen, start);
        continue;
      }
      case ')': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::RParen, start);
        continue;
      }
      case '{': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::LBrace, start);
        continue;
      }
      case '}': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::RBrace, start);
        continue;
      }
      case '[': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::LBracket, start);
        continue;
      }
      case ']': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::RBracket, start);
        continue;
      }
      case ',': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Comma, start);
        continue;
      }
      case ';': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Semicolon, start);
        continue;
      }
      case ':': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Colon, start);
        continue;
      }
      case '!': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Bang, start);
        continue;
      }
      case '+': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Plus, start);
        continue;
      }
      case '-': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Minus, start);
        continue;
      }
      case '*': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Star, start);
        continue;
      }
      case '/': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Slash, start);
        continue;
      }
      case '^': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Caret, start);
        continue;
      }
      case '%': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Percent, start);
        continue;
      }
      case '&': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Ampersand, start);
        continue;
      }
      case '=': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::Eq, start);
        continue;
      }
      case '<':
        scan_lt();
        continue;
      case '>':
        scan_gt();
        continue;
      case '@': {
        const std::size_t start = byte_pos_;
        mark_start();
        advance_one();
        emit(TokenKind::At, start);
        continue;
      }
      default:
        break;
    }

    // Digits: number literal.
    if (is_ascii_digit(static_cast<char>(c)) ||
        (c == '.' && byte_pos_ + 1 < source_.size() && is_ascii_digit(source_[byte_pos_ + 1]))) {
      scan_number();
      continue;
    }

    // `$` only makes sense as the leading anchor of a CellRef. Route it
    // into the ident/cellref scanner if it's followed by a letter; otherwise
    // flag it as InvalidCharacter.
    if (c == '$') {
      if (byte_pos_ + 1 < source_.size() && is_ascii_letter(static_cast<char>(source_[byte_pos_ + 1]))) {
        scan_ident_or_cellref_or_bool();
        continue;
      }
      const std::size_t err_start = byte_pos_;
      mark_start();
      advance_one();
      emit(TokenKind::Invalid, err_start);
      record_error(LexerErrorCode::InvalidReference, err_start);
      continue;
    }

    // Identifier / cell-ref / bool-literal start.
    if (is_ident_start_byte(c)) {
      scan_ident_or_cellref_or_bool();
      continue;
    }

    // Unclassifiable byte: emit an error and consume one codepoint.
    {
      const std::size_t err_start = byte_pos_;
      advance_one();
      record_error(LexerErrorCode::InvalidCharacter, err_start);
    }
  }

  // Always terminate with Eof. The range is a zero-width slice at the
  // current UTF-16 offset.
  Token eof;
  eof.kind = TokenKind::Eof;
  eof.range.start = utf16_pos_;
  eof.range.end = utf16_pos_;
  eof.range.line = line_;
  eof.range.column = column_;
  eof.lexeme = std::string_view(source_.data() + byte_pos_, 0);
  tokens_.push_back(eof);
  return tokens_;
}

// ---------------------------------------------------------------------------
// Position bookkeeping and per-codepoint advance
// ---------------------------------------------------------------------------

Tokenizer::CodepointInfo Tokenizer::peek_codepoint(std::size_t i) const noexcept {
  CodepointInfo info;
  if (i >= source_.size()) {
    return info;
  }
  const unsigned char c0 = static_cast<unsigned char>(source_[i]);
  if (c0 < 0x80) {
    info.codepoint = c0;
    info.byte_len = 1;
    info.utf16_units = 1;
    info.valid = true;
    return info;
  }
  std::uint32_t need = 0;
  std::uint32_t value = 0;
  if ((c0 & 0xE0) == 0xC0) {
    need = 1;
    value = c0 & 0x1F;
  } else if ((c0 & 0xF0) == 0xE0) {
    need = 2;
    value = c0 & 0x0F;
  } else if ((c0 & 0xF8) == 0xF0) {
    need = 3;
    value = c0 & 0x07;
  } else {
    info.byte_len = 1;  // skip one malformed byte
    return info;
  }
  if (i + need >= source_.size()) {
    info.byte_len = 1;
    return info;
  }
  for (std::uint32_t k = 0; k < need; ++k) {
    const unsigned char ck = static_cast<unsigned char>(source_[i + 1 + k]);
    if ((ck & 0xC0) != 0x80) {
      info.byte_len = 1;
      return info;
    }
    value = (value << 6) | (ck & 0x3F);
  }
  info.codepoint = value;
  info.byte_len = need + 1;
  info.utf16_units = value > 0xFFFF ? 2 : 1;
  info.valid = true;
  return info;
}

void Tokenizer::advance_one() {
  if (byte_pos_ >= source_.size()) {
    return;
  }
  const CodepointInfo info = peek_codepoint(byte_pos_);
  const char ch = source_[byte_pos_];
  if (ch == '\n') {
    ++line_;
    column_ = 1;
  } else if (ch == '\r') {
    ++line_;
    column_ = 1;
    // Swallow a paired LF if present so "\r\n" counts as one line break.
    if (byte_pos_ + 1 < source_.size() && source_[byte_pos_ + 1] == '\n') {
      // consume the CR here, the LF will advance on the next iteration but
      // will see column_ == 1 already; suppress the second line bump.
      byte_pos_ += 1;
      utf16_pos_ += info.utf16_units;
      // Advance across the LF without bumping line.
      const CodepointInfo lf = peek_codepoint(byte_pos_);
      byte_pos_ += lf.byte_len;
      utf16_pos_ += lf.utf16_units;
      return;
    }
  } else {
    column_ += info.utf16_units;
  }
  byte_pos_ += info.byte_len == 0 ? 1 : info.byte_len;
  utf16_pos_ += info.utf16_units;
}

void Tokenizer::mark_start() noexcept {
  start_utf16_ = utf16_pos_;
  start_line_ = line_;
  start_column_ = column_;
}

TextRange Tokenizer::make_range() const noexcept {
  TextRange r;
  r.start = start_utf16_;
  r.end = utf16_pos_;
  r.line = start_line_;
  r.column = start_column_;
  return r;
}

void Tokenizer::emit(TokenKind kind, std::size_t lex_start) {
  Token t;
  t.kind = kind;
  t.range = make_range();
  t.lexeme = std::string_view(source_.data() + lex_start, byte_pos_ - lex_start);
  tokens_.push_back(t);
}

void Tokenizer::record_error(LexerErrorCode code, std::size_t err_start) {
  LexerError e;
  e.code = code;
  e.range = make_range();
  e.lexeme = std::string_view(source_.data() + err_start, byte_pos_ - err_start);
  errors_.push_back(e);
}

// ---------------------------------------------------------------------------
// Per-kind scanners
// ---------------------------------------------------------------------------

void Tokenizer::scan_whitespace() {
  const std::size_t start = byte_pos_;
  mark_start();
  while (byte_pos_ < source_.size() && is_formula_whitespace(source_[byte_pos_])) {
    advance_one();
  }
  emit(TokenKind::Whitespace, start);
}

void Tokenizer::scan_string() {
  const std::size_t start = byte_pos_;
  mark_start();
  advance_one();  // consume leading '"'.

  std::string buf;
  bool terminated = false;
  while (byte_pos_ < source_.size()) {
    const char ch = source_[byte_pos_];
    if (ch == '"') {
      // Doubled quote escape.
      if (byte_pos_ + 1 < source_.size() && source_[byte_pos_ + 1] == '"') {
        buf.push_back('"');
        advance_one();
        advance_one();
        continue;
      }
      // Closing quote.
      advance_one();
      terminated = true;
      break;
    }
    const CodepointInfo info = peek_codepoint(byte_pos_);
    if (!info.valid) {
      // Invalid UTF-8 inside a string: append the raw byte and advance.
      buf.push_back(ch);
      advance_one();
      continue;
    }
    buf.append(source_.data() + byte_pos_, info.byte_len);
    advance_one();
  }

  Token t;
  t.kind = TokenKind::String;
  t.range = make_range();
  t.lexeme = std::string_view(source_.data() + start, byte_pos_ - start);
  // Intern the resolved payload into the arena so the view outlives `buf`.
  t.text = arena_.intern(std::string_view(buf.data(), buf.size()));
  tokens_.push_back(t);

  if (!terminated) {
    record_error(LexerErrorCode::UnterminatedString, start);
  }
}

void Tokenizer::scan_quoted_sheet_name() {
  const std::size_t start = byte_pos_;
  mark_start();
  advance_one();  // consume leading '.

  std::string buf;
  bool terminated = false;
  while (byte_pos_ < source_.size()) {
    const char ch = source_[byte_pos_];
    if (ch == '\'') {
      if (byte_pos_ + 1 < source_.size() && source_[byte_pos_ + 1] == '\'') {
        buf.push_back('\'');
        advance_one();
        advance_one();
        continue;
      }
      advance_one();
      terminated = true;
      break;
    }
    const CodepointInfo info = peek_codepoint(byte_pos_);
    if (!info.valid) {
      buf.push_back(ch);
      advance_one();
      continue;
    }
    buf.append(source_.data() + byte_pos_, info.byte_len);
    advance_one();
  }

  Token t;
  t.kind = TokenKind::SheetName;
  t.range = make_range();
  t.lexeme = std::string_view(source_.data() + start, byte_pos_ - start);
  t.text = arena_.intern(std::string_view(buf.data(), buf.size()));
  tokens_.push_back(t);

  if (!terminated) {
    record_error(LexerErrorCode::UnterminatedSheetQuote, start);
  }
}

void Tokenizer::scan_number() {
  const std::size_t start = byte_pos_;
  mark_start();

  bool saw_dot = false;
  bool saw_exp = false;
  bool is_integer = true;

  // Integer / fractional parts.
  while (byte_pos_ < source_.size()) {
    const char ch = source_[byte_pos_];
    if (is_ascii_digit(ch)) {
      advance_one();
      continue;
    }
    if (ch == '.' && !saw_dot && !saw_exp) {
      saw_dot = true;
      is_integer = false;
      advance_one();
      continue;
    }
    if ((ch == 'e' || ch == 'E') && !saw_exp) {
      saw_exp = true;
      is_integer = false;
      advance_one();
      if (byte_pos_ < source_.size() && (source_[byte_pos_] == '+' || source_[byte_pos_] == '-')) {
        advance_one();
      }
      // Require at least one exponent digit.
      bool have_exp_digit = false;
      while (byte_pos_ < source_.size() && is_ascii_digit(source_[byte_pos_])) {
        have_exp_digit = true;
        advance_one();
      }
      if (!have_exp_digit) {
        emit(TokenKind::Invalid, start);
        record_error(LexerErrorCode::InvalidNumberLiteral, start);
        return;
      }
      break;
    }
    break;
  }

  std::string_view lex(source_.data() + start, byte_pos_ - start);

  // Reject degenerate runs: a bare '.' (no digits around it), or an empty
  // fractional part like `1.`. We allow `.5` because the main loop only
  // entered scan_number when `.` is followed by a digit. For `1.`, the
  // fractional branch sets saw_dot but consumes no digit, which is invalid.
  if (lex == ".") {
    emit(TokenKind::Invalid, start);
    record_error(LexerErrorCode::InvalidNumberLiteral, start);
    return;
  }
  if (saw_dot && !saw_exp) {
    // Ensure there's at least one digit either before or after the dot.
    bool has_digit_before = false;
    bool has_digit_after = false;
    bool seen_dot = false;
    for (char ch : lex) {
      if (ch == '.') {
        seen_dot = true;
        continue;
      }
      if (is_ascii_digit(ch)) {
        if (seen_dot) {
          has_digit_after = true;
        } else {
          has_digit_before = true;
        }
      }
    }
    if (!has_digit_before && !has_digit_after) {
      emit(TokenKind::Invalid, start);
      record_error(LexerErrorCode::InvalidNumberLiteral, start);
      return;
    }
  }

  // Reject an immediately-following second '.' as in `1.2.3`. Without this
  // check the main loop would produce Number("1.2"), Invalid or similar;
  // flagging it here yields the more actionable diagnostic.
  if (byte_pos_ < source_.size() && source_[byte_pos_] == '.') {
    // Absorb the offending run up to the next whitespace / operator so the
    // diagnostic lexeme covers the whole malformed literal.
    while (byte_pos_ < source_.size() && (is_ascii_digit(source_[byte_pos_]) || source_[byte_pos_] == '.')) {
      advance_one();
    }
    emit(TokenKind::Invalid, start);
    record_error(LexerErrorCode::InvalidNumberLiteral, start);
    return;
  }

  // Parse via strtod over a NUL-terminated copy (strtod requires C-strings).
  char buf[64];
  const std::size_t n = (lex.size() < sizeof(buf) - 1) ? lex.size() : sizeof(buf) - 1;
  std::memcpy(buf, lex.data(), n);
  buf[n] = '\0';
  char* end_ptr = nullptr;
  const double value = std::strtod(buf, &end_ptr);
  if (end_ptr != buf + n) {
    emit(TokenKind::Invalid, start);
    record_error(LexerErrorCode::InvalidNumberLiteral, start);
    return;
  }

  Token t;
  t.kind = TokenKind::Number;
  t.range = make_range();
  t.lexeme = lex;
  t.number = value;
  t.is_integer = is_integer;
  tokens_.push_back(t);
}

void Tokenizer::scan_error_literal() {
  const std::size_t start = byte_pos_;
  mark_start();

  // Spilled-range `#`: emitted if the previous non-whitespace token was a
  // CellRef and there was no whitespace between. Whitespace between a
  // CellRef and `#` becomes Whitespace then Hash (the parser decides).
  if (last_cellref_end_byte_ == byte_pos_) {
    advance_one();
    emit(TokenKind::Hash, start);
    last_cellref_end_byte_ = static_cast<std::size_t>(-1);
    return;
  }

  // Otherwise try to match an error literal. Scan a reasonable run first so
  // we can compare it against the catalog verbatim.
  std::size_t probe = byte_pos_ + 1;  // skip '#'
  while (probe < source_.size()) {
    const unsigned char c = static_cast<unsigned char>(source_[probe]);
    // Accept letters, digits, `/`, `_`, `?`, `!`: enough to cover every
    // error-literal spelling. Stop before whitespace / operators.
    if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '/' || c == '_' ||
        c == '?' || c == '!') {
      ++probe;
      continue;
    }
    break;
  }
  std::string_view run(source_.data() + byte_pos_, probe - byte_pos_);
  ErrorCode code;
  if (match_error_literal(run, &code)) {
    while (byte_pos_ < probe) {
      advance_one();
    }
    Token t;
    t.kind = TokenKind::ErrorLiteral;
    t.range = make_range();
    t.lexeme = std::string_view(source_.data() + start, byte_pos_ - start);
    t.error_code = code;
    tokens_.push_back(t);
    return;
  }

  // Fallback: consume the run as Invalid so the parser can proceed past it.
  while (byte_pos_ < probe) {
    advance_one();
  }
  emit(TokenKind::Invalid, start);
  record_error(LexerErrorCode::InvalidErrorLiteral, start);
}

void Tokenizer::scan_ident_or_cellref_or_bool() {
  const std::size_t start = byte_pos_;
  mark_start();

  // Optional leading `$` (only relevant if we end up being a CellRef).
  if (byte_pos_ < source_.size() && source_[byte_pos_] == '$') {
    advance_one();
  }

  while (byte_pos_ < source_.size()) {
    const unsigned char c = static_cast<unsigned char>(source_[byte_pos_]);
    if (c == '$') {
      // Accept at most one internal `$` between column and row.
      advance_one();
      continue;
    }
    if (c < 0x80) {
      if (is_ident_cont_byte(c)) {
        advance_one();
        continue;
      }
      break;
    }
    // Multi-byte: decode the whole codepoint and reject sentinels that are
    // not legitimately identifier-eligible (U+FEFF BOM, U+3000 full-width
    // space). Every other codepoint >= 0x80 is accepted as ident-eligible
    // for M2.2.
    const CodepointInfo info = peek_codepoint(byte_pos_);
    if (!info.valid) {
      break;
    }
    if (info.codepoint == 0xFEFF || info.codepoint == 0x3000) {
      break;
    }
    advance_one();
  }

  std::string_view run(source_.data() + start, byte_pos_ - start);

  // Classify: CellRef first, then Bool, otherwise Ident.
  bool letters_only = false;
  if (looks_like_cellref(run, &letters_only)) {
    Token t;
    t.kind = TokenKind::CellRef;
    t.range = make_range();
    t.lexeme = run;
    tokens_.push_back(t);
    last_cellref_end_byte_ = byte_pos_;
    return;
  }

  bool b = false;
  if (is_bool_word(run, &b)) {
    Token t;
    t.kind = TokenKind::Bool;
    t.range = make_range();
    t.lexeme = run;
    t.boolean = b;
    tokens_.push_back(t);
    return;
  }

  // Degenerate forms: `$A` or `A$` without a row are not refs and not
  // identifiers; flag as InvalidReference and emit Invalid so the parser
  // doesn't try to treat them as names.
  if (run.size() >= 2 && (run.front() == '$' || run.back() == '$')) {
    // Allow `_xlfn.` and similar which don't contain `$`; we only reach
    // this path if `$` is present.
    Token t;
    t.kind = TokenKind::Invalid;
    t.range = make_range();
    t.lexeme = run;
    tokens_.push_back(t);
    record_error(LexerErrorCode::InvalidReference, start);
    return;
  }

  Token t;
  t.kind = TokenKind::Ident;
  t.range = make_range();
  t.lexeme = run;
  tokens_.push_back(t);
  (void)letters_only;
}

void Tokenizer::scan_lt() {
  const std::size_t start = byte_pos_;
  mark_start();
  advance_one();  // '<'
  if (byte_pos_ < source_.size()) {
    const char next = source_[byte_pos_];
    if (next == '=') {
      advance_one();
      emit(TokenKind::LtEq, start);
      return;
    }
    if (next == '>') {
      advance_one();
      emit(TokenKind::NotEq, start);
      return;
    }
  }
  emit(TokenKind::Lt, start);
}

void Tokenizer::scan_gt() {
  const std::size_t start = byte_pos_;
  mark_start();
  advance_one();  // '>'
  if (byte_pos_ < source_.size() && source_[byte_pos_] == '=') {
    advance_one();
    emit(TokenKind::GtEq, start);
    return;
  }
  emit(TokenKind::Gt, start);
}

}  // namespace parser
}  // namespace formulon
