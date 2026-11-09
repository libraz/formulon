//
// Implementation of the Excel formula tokenizer. See the header for
// the high-level contract. Key implementation choices:
//
//   * UTF-8 decoding is hand-rolled (no `<codecvt>` / ICU): we only need
//     byte-length and UTF-16 code-unit count per codepoint.
//   * Numbers keep 15 significant digits and are scaled the way Excel does
//     (`parse_excel_decimal` in `utils/double_parse.h`).
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
#include <cstring>
#include <iterator>
#include <string>
#include <string_view>

#include "excel_locale.h"
#include "parser/formula_prefix.h"
#include "parser/parser_detail.h"
#include "utils/double_parse.h"
#include "utils/strings.h"
#include "utils/text_ops.h"

namespace formulon {
namespace parser {

namespace {

// UTF-8 BOM and UTF-16 LE BOM (as 2 bytes) for the source-prefix check.
constexpr std::string_view kUtf16LeBom = "\xFF\xFE";

constexpr std::uint32_t kBomCodepoint = 0xFEFF;
constexpr std::uint32_t kIdeographicSpaceCodepoint = 0x3000;

// Byte length of a `[Book]Name` written with no `!` at the `[` at `pos`, a
// name Excel stores verbatim; 0 when the text there is not that shape. A
// decimal book, and a name followed by `!`, `(` or `:`, are excluded.
std::size_t BracketNameLength(std::string_view source, std::size_t pos) noexcept {
  const std::size_t book_len = detail::BareQualifierRunLength(source, pos + 1);
  std::size_t end = pos + 1 + book_len;
  if (book_len == 0 || end >= source.size() || source[end] != ']') {
    return 0;
  }
  const bool decimal = detail::IsAllAsciiDigits(source.substr(pos + 1, end - pos - 1));
  const std::size_t name_len = detail::BareQualifierRunLength(source, end + 1);
  end += 1 + name_len;
  if (decimal || name_len == 0) {
    return 0;
  }
  if (end < source.size() && (source[end] == '!' || source[end] == '(' || source[end] == ':')) {
    return 0;
  }
  return end - pos;
}

}  // namespace

// ---------------------------------------------------------------------------
// Static helpers
// ---------------------------------------------------------------------------

bool Tokenizer::is_ident_start_byte(unsigned char c) noexcept {
  // Only UTF-8 leading bytes (>= 0xC0) are valid identifier starts.
  // U+0080..U+00BF are continuation bytes; treating them as start bytes
  // lets a stray continuation slip into the identifier slot and pulls
  // subsequent token boundaries off by one byte.
  // `\` is Excel's third legal name-manager start character (alongside a
  // letter and `_`) for a defined name -- `scan_ident_or_cellref_or_bool`
  // consumes it explicitly before the continuation loop, since it is not
  // itself a `formula_prefix_is_ident_byte`.
  return strings::is_ascii_alpha(static_cast<char>(c)) || c == '_' || c == '\\' || c >= 0xC0;
}

bool Tokenizer::is_formula_whitespace(char c) noexcept {
  // Excel accepts ASCII space / tab / CR / LF as formula whitespace. The
  // full-width space U+3000 is intentionally *not* accepted here; it is
  // flagged as InvalidCharacter by the main loop.
  return detail::IsFormulaWhitespace(c);
}

bool Tokenizer::is_bool_word(std::string_view word, bool* out) noexcept {
  if (strings::case_insensitive_eq(word, "TRUE")) {
    *out = true;
    return true;
  }
  if (strings::case_insensitive_eq(word, "FALSE")) {
    *out = false;
    return true;
  }
  return false;
}

bool Tokenizer::is_bool_name(std::string_view word, bool* out) const noexcept {
  if (is_bool_word(word, out)) {
    return true;
  }
  if (opts_.locale == nullptr) {
    return false;
  }
  const LocaleFacts& f = *opts_.locale;
  if (strings::case_insensitive_eq(word, f.true_name)) {
    *out = true;
    return true;
  }
  if (strings::case_insensitive_eq(word, f.false_name)) {
    *out = false;
    return true;
  }
  return false;
}

bool Tokenizer::match_locale_error(std::string_view run, ErrorCode* out, std::size_t* match_len) const noexcept {
  if (opts_.locale == nullptr) {
    return false;
  }
  for (std::size_t i = 0; i < opts_.locale->error_names.size(); ++i) {
    const std::string_view name = opts_.locale->error_names[i];
    if (name != kErrorTable[i].display_name && strings::case_insensitive_starts_with(run, name)) {
      *out = static_cast<ErrorCode>(i);
      *match_len = name.size();
      return true;
    }
  }
  return false;
}

bool Tokenizer::scan_locale_char(unsigned char c) {
  const LocaleFacts& f = *opts_.locale;
  const char ch = static_cast<char>(c);
  if (brace_depth_ > 0) {
    if (ch == f.array_column_separator) {
      emit_single_char(TokenKind::Comma);
      return true;
    }
    if (ch == f.array_row_separator) {
      emit_single_char(TokenKind::Semicolon);
      return true;
    }
  } else if (ch == f.list_separator) {
    emit_single_char(TokenKind::Comma);
    return true;
  }
  if (ch == decimal_ && ch != '.' && byte_pos_ + 1 < source_.size() &&
      strings::is_ascii_digit(source_[byte_pos_ + 1])) {
    scan_number();
    return true;
  }
  return false;
}

bool Tokenizer::match_error_literal(std::string_view run, ErrorCode* out, std::size_t* match_len) noexcept {
  // Longest prefix wins so an error literal followed by an operator or
  // reference (`#REF!/2`, `#N/A/B1`, `#DIV/0!/A1`) commits on the literal alone.
  std::size_t best_len = 0;
  for (std::size_t i = 0; i < std::size(kErrorTable); ++i) {
    const std::string_view text = kErrorTable[i].display_name;
    if (text.size() > best_len && strings::case_insensitive_starts_with(run, text)) {
      *out = static_cast<ErrorCode>(i);
      best_len = text.size();
    }
  }
  if (best_len == 0) {
    return false;
  }
  *match_len = best_len;
  return true;
}

bool Tokenizer::looks_like_cellref(std::string_view run, bool* letters_only) noexcept {
  *letters_only = false;
  std::size_t i = 0;
  if (i < run.size() && run[i] == '$') {
    ++i;
  }
  const std::size_t letters_begin = i;
  while (i < run.size() && strings::is_ascii_alpha(run[i])) {
    ++i;
  }
  const std::size_t letters_len = i - letters_begin;
  if (letters_len == 0 || letters_len > detail::kMaxColumnLetters) {
    return false;
  }
  if (i < run.size() && run[i] == '$') {
    ++i;
  }
  const std::size_t digits_begin = i;
  while (i < run.size() && strings::is_ascii_digit(run[i])) {
    ++i;
  }
  const std::size_t digits_len = i - digits_begin;
  if (i != run.size()) {
    return false;
  }
  // Check column limit.
  if (detail::DecodeColumnLetters(run.substr(letters_begin, letters_len)) == 0) {
    return false;
  }
  if (digits_len == 0) {
    // Letter-only with an optional leading `$`: treated as identifier-like.
    // The caller re-emits as Ident.
    if (run[0] == '$' || run.back() == '$') {
      // `$A` or `A$` without digits is still malformed.
      return false;
    }
    *letters_only = true;
    return false;
  }
  if (digits_len > detail::kMaxRowDigits) {
    return false;
  }
  const std::uint64_t row = detail::DecodeDigitRunClamped(run.substr(digits_begin, digits_len), detail::kMaxRow);
  return row != 0 && row <= detail::kMaxRow;
}

// ---------------------------------------------------------------------------
// Construction and public API
// ---------------------------------------------------------------------------

Tokenizer::Tokenizer(std::string_view source, TokenizerOptions opts) noexcept
    : source_(source),
      opts_(opts),
      arena_(4096),
      decimal_(opts.locale != nullptr ? opts.locale->decimal_separator : '.') {}

const std::vector<Token>& Tokenizer::tokens() {
  if (done_) {
    return tokens_;
  }
  done_ = true;

  // BOM handling: strip a leading UTF-8 or UTF-16-LE BOM silently.
  if (source_.size() >= detail::kUtf8Bom.size() &&
      std::memcmp(source_.data(), detail::kUtf8Bom.data(), detail::kUtf8Bom.size()) == 0) {
    byte_pos_ = detail::kUtf8Bom.size();
  } else if (source_.size() >= kUtf16LeBom.size() &&
             std::memcmp(source_.data(), kUtf16LeBom.data(), kUtf16LeBom.size()) == 0) {
    byte_pos_ = kUtf16LeBom.size();
  }
  // Note: BOM bytes are *not* counted in UTF-16 offsets, so utf16_pos_ stays
  // at 0 here. This matches Excel's "BOM is invisible" contract.

  // Bound the scan to the length cap before any scanner runs. The loop
  // below only re-checks the cap between tokens, so an input that is one
  // enormous token (a string literal, an identifier, a quoted sheet name)
  // would otherwise be consumed whole no matter how long it is. Trimming
  // the view keeps every scanner inside the budget; `source_` still points
  // into the caller's buffer, so token lexemes stay valid.
  const std::size_t capped_end = length_capped_end(byte_pos_);
  const bool over_length = capped_end < source_.size();
  if (over_length) {
    source_ = source_.substr(0, capped_end);
  }

  while (byte_pos_ < source_.size()) {
    const unsigned char c = static_cast<unsigned char>(source_[byte_pos_]);

    // Mid-input BOM: consume the 3 bytes (or 2 for UTF-16 LE) and record an
    // error. We handle UTF-8 BOM here because the sentinel is unambiguous.
    if (c == 0xEFU && source_.substr(byte_pos_, detail::kUtf8Bom.size()) == detail::kUtf8Bom) {
      const std::size_t err_start = byte_pos_;
      mark_start();
      // `advance_one()` (rather than a manual `byte_pos_ += 3`) is required
      // here: it is the only thing that advances `utf16_pos_` for this
      // codepoint. The previous manual-jump form left `utf16_pos_` unmoved,
      // so a BOM appearing mid-formula desynchronised the UTF-16 offset of
      // every token scanned after it. U+FEFF is a single BMP code unit, so
      // this also reproduces the prior "zero-width BOM counts as one column"
      // column bump as a side effect.
      advance_one();
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
    if (c == 0xE3U && source_.substr(byte_pos_, detail::kIdeographicSpace.size()) == detail::kIdeographicSpace) {
      const std::size_t err_start = byte_pos_;
      // Mark the range start before advancing past the codepoint, so
      // `start_utf16_` reflects `err_start` rather than the previous token.
      mark_start();
      // Advance by one codepoint (3 bytes) and record.
      advance_one();
      record_error(LexerErrorCode::InvalidCharacter, err_start);
      continue;
    }

    if (opts_.locale != nullptr && bracket_depth_ == 0 && scan_locale_char(c)) {
      continue;
    }

    // Dispatch on the leading byte.
    switch (c) {
      case '"':
        // Inside a bracket a double quote is a column-name character.
        if (bracket_depth_ > 0) {
          emit_single_char(TokenKind::Ident);
        } else {
          scan_quoted('"', TokenKind::String, LexerErrorCode::UnterminatedString);
        }
        continue;
      case '\'':
        // Inside a bracket an apostrophe is the ECMA-376 18.5.1.10
        // structured-reference escape prefix, never a sheet-name quote.
        if (bracket_depth_ > 0) {
          scan_structured_ref_escape();
        } else {
          scan_quoted('\'', TokenKind::SheetName, LexerErrorCode::UnterminatedSheetQuote);
        }
        continue;
      case '#':
        scan_error_literal();
        continue;
      case '(':
        emit_single_char(TokenKind::LParen);
        continue;
      case ')':
        emit_single_char(TokenKind::RParen);
        // `OFFSET(A1,1,0)#` and `(A4)#` both anchor on what the closing
        // parenthesis ends.
        last_anchor_tail_end_byte_ = byte_pos_;
        continue;
      case '{':
        emit_single_char(TokenKind::LBrace);
        ++brace_depth_;
        continue;
      case '}':
        emit_single_char(TokenKind::RBrace);
        if (brace_depth_ > 0) {
          --brace_depth_;
        }
        continue;
      case '[':
        if (bracket_depth_ == 0 && at_operand_start()) {
          if (try_scan_external_qualifier()) {
            continue;
          }
          if (const std::size_t len = BracketNameLength(source_, byte_pos_); len != 0) {
            const std::size_t start = byte_pos_;
            mark_start();
            while (byte_pos_ < start + len) {
              advance_one();
            }
            emit(TokenKind::Ident, start);
            last_anchor_tail_end_byte_ = byte_pos_;
            continue;
          }
        }
        emit_single_char(TokenKind::LBracket);
        ++bracket_depth_;
        continue;
      case ']':
        emit_single_char(TokenKind::RBracket);
        if (bracket_depth_ > 0) {
          --bracket_depth_;
        }
        continue;
      case ',':
        emit_single_char(TokenKind::Comma);
        continue;
      case ';':
        emit_single_char(TokenKind::Semicolon);
        continue;
      case ':':
        scan_colon();
        continue;
      case '!':
        emit_single_char(TokenKind::Bang);
        continue;
      case '+':
        emit_single_char(TokenKind::Plus);
        continue;
      case '-':
        emit_single_char(TokenKind::Minus);
        continue;
      case '*':
        emit_single_char(TokenKind::Star);
        continue;
      case '/':
        emit_single_char(TokenKind::Slash);
        continue;
      case '^':
        emit_single_char(TokenKind::Caret);
        continue;
      case '%':
        emit_single_char(TokenKind::Percent);
        continue;
      case '&':
        emit_single_char(TokenKind::Ampersand);
        continue;
      case '=':
        emit_single_char(TokenKind::Eq);
        continue;
      case '<':
        scan_lt();
        continue;
      case '>':
        scan_gt();
        continue;
      case '@':
        emit_single_char(TokenKind::At);
        continue;
      default:
        break;
    }

    // A sheet qualifier, whatever the run would otherwise tokenize as.
    if (bracket_depth_ == 0 &&
        (strings::is_ascii_digit(static_cast<char>(c)) || (c != '\\' && is_ident_start_byte(c))) &&
        try_scan_local_sheet_qualifier()) {
      continue;
    }

    // `.:` / `.:.` after a range endpoint the scanners stopped short of.
    if (c == '.' && byte_pos_ + 1 < source_.size() && source_[byte_pos_ + 1] == ':') {
      scan_colon();
      continue;
    }

    // Digits: number literal.
    if (strings::is_ascii_digit(static_cast<char>(c)) ||
        (c == static_cast<unsigned char>(decimal_) && byte_pos_ + 1 < source_.size() &&
         strings::is_ascii_digit(source_[byte_pos_ + 1]))) {
      scan_number();
      continue;
    }

    // `$` only makes sense as the leading anchor of a reference. Route `$`
    // followed by a letter into the ident/cellref scanner (`$A1`, `$A:$A`);
    // `$` followed by a digit is an absolute row anchor (`$1:$1`), scanned
    // here as a Number token whose lexeme retains the `$` so the parser's
    // whole-row path can carry the row_abs flag. Anything else is invalid.
    if (c == '$') {
      if (byte_pos_ + 1 < source_.size() && strings::is_ascii_alpha(static_cast<char>(source_[byte_pos_ + 1]))) {
        scan_ident_or_cellref_or_bool();
        continue;
      }
      if (byte_pos_ + 1 < source_.size() && strings::is_ascii_digit(source_[byte_pos_ + 1])) {
        const std::size_t num_start = byte_pos_;
        mark_start();
        advance_one();  // consume '$'
        double value = 0.0;
        while (byte_pos_ < source_.size() && strings::is_ascii_digit(source_[byte_pos_])) {
          value = value * 10.0 + static_cast<double>(source_[byte_pos_] - '0');
          advance_one();
        }
        Token& t = emit(TokenKind::Number, num_start);
        t.number = value;
        t.is_integer = true;
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
      // Mark the range start before advancing, so `start_utf16_` reflects
      // `err_start` rather than the previous token.
      mark_start();
      advance_one();
      record_error(LexerErrorCode::InvalidCharacter, err_start);
    }
  }

  // The scan was bounded above, so the loop ends at the cap without noticing;
  // report the truncation here.
  if (over_length) {
    mark_start();
    record_error(LexerErrorCode::ExcessiveLength, byte_pos_);
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
  const Utf8DecodeResult decoded = decode_first_utf8_codepoint(source_.substr(i));
  if (!decoded.valid) {
    info.byte_len = 1;  // skip one malformed byte
    info.utf16_units = 1;
    return info;
  }
  info.codepoint = decoded.codepoint;
  info.byte_len = static_cast<std::uint32_t>(decoded.byte_len);
  info.utf16_units = decoded.codepoint > 0xFFFFu ? 2U : 1U;
  info.valid = true;
  return info;
}

std::size_t Tokenizer::length_capped_end(std::size_t start) const noexcept {
  // UTF-16 units never exceed the byte count, so a short enough tail cannot hit the cap.
  if (start <= source_.size() && source_.size() - start <= opts_.max_formula_length_utf16) {
    return source_.size();
  }
  std::uint32_t units = 0;
  std::size_t pos = start;
  while (pos < source_.size()) {
    if (units >= opts_.max_formula_length_utf16) {
      return pos;
    }
    const CodepointInfo info = peek_codepoint(pos);
    pos += info.byte_len == 0 ? 1 : info.byte_len;
    units += info.utf16_units;
  }
  return source_.size();
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

Token& Tokenizer::emit(TokenKind kind, std::size_t lex_start) {
  Token t;
  t.kind = kind;
  t.range = make_range();
  t.lexeme = std::string_view(source_.data() + lex_start, byte_pos_ - lex_start);
  tokens_.push_back(t);
  return tokens_.back();
}

void Tokenizer::emit_single_char(TokenKind kind) {
  const std::size_t start = byte_pos_;
  mark_start();
  advance_one();
  emit(kind, start);
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

void Tokenizer::scan_quoted(char quote, TokenKind kind, LexerErrorCode unterminated) {
  const std::size_t start = byte_pos_;
  mark_start();
  advance_one();  // consume the opening quote.

  std::string buf;
  bool terminated = false;
  while (byte_pos_ < source_.size()) {
    const char ch = source_[byte_pos_];
    if (ch == quote) {
      // Doubled quote escape.
      if (byte_pos_ + 1 < source_.size() && source_[byte_pos_ + 1] == quote) {
        buf.push_back(quote);
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
      // Invalid UTF-8 inside the quotes: append the raw byte and advance.
      buf.push_back(ch);
      advance_one();
      continue;
    }
    buf.append(source_.data() + byte_pos_, info.byte_len);
    advance_one();
  }

  Token& t = emit(kind, start);
  // Intern the resolved payload into the arena so the view outlives `buf`.
  t.text = arena_.intern(std::string_view(buf.data(), buf.size()));

  if (!terminated) {
    record_error(unterminated, start);
  }
}

void Tokenizer::scan_structured_ref_escape() {
  const std::size_t start = byte_pos_;
  mark_start();
  advance_one();  // consume the escape apostrophe.
  if (byte_pos_ < source_.size()) {
    advance_one();  // the escaped codepoint is literal, whatever it is.
  }
  emit(TokenKind::Ident, start);
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
    if (strings::is_ascii_digit(ch)) {
      advance_one();
      continue;
    }
    if (ch == decimal_ && !saw_dot && !saw_exp) {
      // `1.:3` is a leading-trim row range, not the literal `1.`.
      if (decimal_ == '.' && byte_pos_ + 1 < source_.size() && source_[byte_pos_ + 1] == ':' && byte_pos_ > start) {
        break;
      }
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
      while (byte_pos_ < source_.size() && strings::is_ascii_digit(source_[byte_pos_])) {
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

  // Reject an immediately-following second '.' as in `1.2.3`. Without this
  // check the main loop would produce Number("1.2"), Invalid or similar;
  // flagging it here yields the more actionable diagnostic.
  if (byte_pos_ < source_.size() && source_[byte_pos_] == decimal_ &&
      !(decimal_ == '.' && byte_pos_ + 1 < source_.size() && source_[byte_pos_ + 1] == ':')) {
    // Absorb the offending run up to the next whitespace / operator so the
    // diagnostic lexeme covers the whole malformed literal.
    while (byte_pos_ < source_.size() &&
           (strings::is_ascii_digit(source_[byte_pos_]) || source_[byte_pos_] == decimal_)) {
      advance_one();
    }
    emit(TokenKind::Invalid, start);
    record_error(LexerErrorCode::InvalidNumberLiteral, start);
    return;
  }

  // Apply Excel's 15-significant-digit rule before parsing. The original
  // lexeme is still recorded on the token for diagnostics.
  std::string invariant;
  std::string_view numeric_lex = lex;
  if (decimal_ != '.') {
    invariant.assign(lex);
    for (char& ch : invariant) {
      if (ch == decimal_) {
        ch = '.';
      }
    }
    numeric_lex = invariant;
  }
  double value = 0.0;
  // A nonzero literal whose leading digit lies outside Excel's range is a
  // syntax error, so a literal never overflows.
  if (!parse_excel_literal(numeric_lex, &value)) {
    emit(TokenKind::Invalid, start);
    record_error(LexerErrorCode::InvalidNumberLiteral, start);
    return;
  }

  Token& t = emit(TokenKind::Number, start);
  t.number = value;
  t.is_integer = is_integer;
}

void Tokenizer::scan_error_literal() {
  const std::size_t start = byte_pos_;
  mark_start();

  // Spilled-range `#`: emitted if the previous token can end a spill
  // anchor and there was no whitespace between. Whitespace before the `#`
  // becomes Whitespace then Hash (the parser decides). The parser is what
  // rules on whether the anchor shape is legal; the tokenizer only has to
  // stop reading `#` as the opening byte of an error literal.
  if (last_anchor_tail_end_byte_ == byte_pos_) {
    advance_one();
    emit(TokenKind::Hash, start);
    last_anchor_tail_end_byte_ = static_cast<std::size_t>(-1);
    return;
  }

  // Otherwise try to match an error literal. Scan a reasonable run first so
  // we can compare it against the catalog verbatim.
  std::size_t probe = byte_pos_ + 1;  // skip '#'
  while (probe < source_.size()) {
    const unsigned char c = static_cast<unsigned char>(source_[probe]);
    // Accept letters, digits, `/`, `_`, `?`, `!`, `.` and non-ASCII bytes:
    // enough to cover every error-literal spelling, localized ones
    // (`#Н/Д`, `#DELING.DOOR.0!`) included. Stop before whitespace / operators.
    if (strings::is_ascii_alpha(static_cast<char>(c)) || strings::is_ascii_digit(static_cast<char>(c)) || c == '/' ||
        c == '_' || c == '?' || c == '!' || c == '.' || c >= 0x80) {
      ++probe;
      continue;
    }
    break;
  }
  std::string_view run(source_.data() + byte_pos_, probe - byte_pos_);
  ErrorCode code;
  std::size_t match_len = 0;
  if (match_error_literal(run, &code, &match_len) || match_locale_error(run, &code, &match_len)) {
    // Consume only the matched literal. Any trailing run bytes (an operator
    // or reference glued to the literal) stay for the main loop.
    const std::size_t match_end = byte_pos_ + match_len;
    while (byte_pos_ < match_end) {
      advance_one();
    }
    emit(TokenKind::ErrorLiteral, start).error_code = code;
    return;
  }

  // `#REF!#`: a `#` glued to a Ref error literal that opens no known error
  // literal is the spill operator on a collapsed anchor.
  if (!tokens_.empty()) {
    const Token& prev = tokens_.back();
    if (prev.kind == TokenKind::ErrorLiteral && prev.error_code == ErrorCode::Ref &&
        prev.lexeme.data() + prev.lexeme.size() == source_.data() + byte_pos_) {
      advance_one();
      emit(TokenKind::Hash, start);
      return;
    }
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

  // Optional leading `\` (defined-name start character; see
  // `is_ident_start_byte`). It is not a `formula_prefix_is_ident_byte`, so it has to
  // be consumed here rather than falling into the loop below; a CellRef
  // never carries one, so `looks_like_cellref` rejects the run unchanged.
  if (byte_pos_ < source_.size() && source_[byte_pos_] == '\\') {
    advance_one();
  }

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
      // A cell or column endpoint stops before the `.:` of a trim operator; a
      // `.` inside a name (`Tax.Rate`) does not.
      if (c == '.' && byte_pos_ + 1 < source_.size() && source_[byte_pos_ + 1] == ':') {
        const std::string_view head(source_.data() + start, byte_pos_ - start);
        bool letters_only = false;
        bool anchored_letters = false;
        if (head.size() >= 2 && head.front() == '$') {
          (void)looks_like_cellref(head.substr(1), &anchored_letters);
        }
        if (looks_like_cellref(head, &letters_only) || letters_only || anchored_letters) {
          break;
        }
      }
      // `.` separates the columns of an array constant in a locale that says so.
      if (c == '.' && brace_depth_ > 0 && opts_.locale != nullptr && opts_.locale->array_column_separator == '.') {
        break;
      }
      if (formula_prefix_is_ident_byte(static_cast<char>(c))) {
        advance_one();
        continue;
      }
      break;
    }
    // Multi-byte: decode the whole codepoint and reject sentinels that are
    // not legitimately identifier-eligible (U+FEFF BOM, U+3000 full-width
    // space). Every other codepoint >= 0x80 is accepted as ident-eligible.
    const CodepointInfo info = peek_codepoint(byte_pos_);
    if (!info.valid) {
      break;
    }
    if (info.codepoint == kBomCodepoint || info.codepoint == kIdeographicSpaceCodepoint) {
      break;
    }
    advance_one();
  }

  // A lead byte that starts no decodable codepoint — a truncated or
  // otherwise malformed UTF-8 sequence — leaves the loop above having
  // consumed nothing. Emitting a zero-width token here would hand control
  // back to the dispatcher with `byte_pos_` unmoved, and since the same
  // byte still classifies as an identifier start the pair would spin
  // forever, appending empty tokens until the process runs out of memory.
  // Consume the byte as unclassifiable instead, which is how the dispatcher
  // treats every other byte it cannot begin a token with.
  if (byte_pos_ == start) {
    advance_one();
    record_error(LexerErrorCode::InvalidCharacter, start);
    return;
  }

  std::string_view run(source_.data() + start, byte_pos_ - start);

  // Classify: CellRef first, then Bool, otherwise Ident.
  bool letters_only = false;
  if (looks_like_cellref(run, &letters_only)) {
    emit(TokenKind::CellRef, start);
    last_anchor_tail_end_byte_ = byte_pos_;
    return;
  }

  bool b = false;
  if (is_bool_name(run, &b)) {
    emit(TokenKind::Bool, start).boolean = b;
    return;
  }

  // Absolute whole-column anchor: `$A` (a single leading `$` followed by a
  // valid column-letters run and nothing else) is the endpoint of an
  // absolute whole-column reference such as `$A:$A` / `$A:$C`. Emit it as an
  // Ident so the parser's whole-column path (which decodes the `$` via
  // `decode_column_letters`) can pair it across the `:`. A standalone `$A`
  // then parses as a NameRef and resolves to #NAME?, which is acceptable for
  // that malformed lone form. `A$`, `$A$`, and multi-`$` runs stay Invalid.
  if (run.size() >= 2 && run.front() == '$' && run.back() != '$') {
    if (detail::DecodeColumnLetters(run.substr(1)) != 0) {
      emit(TokenKind::Ident, start);
      return;
    }
  }

  // A run that contains `$` but did not classify as a CellRef (or the one
  // absolute whole-column form above) is neither a legal name nor a valid
  // reference. In particular this rejects repeated anchors such as `A$$1`;
  // treating those as identifiers silently defers a syntax error to #NAME?.
  if (run.find('$') != std::string_view::npos) {
    emit(TokenKind::Invalid, start);
    record_error(LexerErrorCode::InvalidReference, start);
    return;
  }

  emit(TokenKind::Ident, start);
  // A defined name or a LET binding can anchor a spill (`Anchor#`,
  // `LET(x, A4, SUM(x#))`), so an identifier arms the operator too.
  last_anchor_tail_end_byte_ = byte_pos_;
}

bool Tokenizer::try_scan_local_sheet_qualifier() {
  const std::size_t start = byte_pos_;
  const std::size_t len = detail::BareQualifierRunLength(source_, start);
  const std::size_t end = start + len;
  if (len == 0 || end >= source_.size()) {
    return false;
  }
  const std::string_view run = source_.substr(start, len);
  bool b = false;
  if (is_bool_name(run, &b)) {
    return false;
  }
  if (source_[end] != '!') {
    // `X:Y!`: the run is a 3-D span's first sheet unless it is cell-shaped
    // or the second sheet opens with a digit.
    if (source_[end] != ':') {
      return false;
    }
    const std::size_t second_len = detail::BareQualifierRunLength(source_, end + 1);
    const std::size_t second_end = end + 1 + second_len;
    bool letters_only = false;
    if (second_len == 0 || second_end >= source_.size() || source_[second_end] != '!' ||
        strings::is_ascii_digit(source_[end + 1]) || looks_like_cellref(run, &letters_only)) {
      return false;
    }
  }
  mark_start();
  while (byte_pos_ < end) {
    advance_one();
  }
  emit(TokenKind::Ident, start);
  last_anchor_tail_end_byte_ = byte_pos_;
  return true;
}

bool Tokenizer::try_scan_external_qualifier() {
  const std::size_t start = byte_pos_;
  std::size_t pos = start + 1;
  const std::size_t book_len = detail::BareQualifierRunLength(source_, pos);
  pos += book_len;
  if (book_len == 0 || pos >= source_.size() || source_[pos] != ']') {
    return false;
  }
  ++pos;
  const std::size_t sheet_len = detail::BareQualifierRunLength(source_, pos);
  pos += sheet_len;
  if (sheet_len == 0) {
    return false;
  }
  if (pos < source_.size() && source_[pos] == ':') {
    const std::size_t end_len = detail::BareQualifierRunLength(source_, pos + 1);
    if (end_len == 0) {
      return false;
    }
    pos += 1 + end_len;
  }
  if (pos >= source_.size() || source_[pos] != '!') {
    return false;
  }
  mark_start();
  while (byte_pos_ < pos) {
    advance_one();
  }
  emit(TokenKind::ExternalQualifier, start);
  return true;
}

bool Tokenizer::at_operand_start() const noexcept {
  if (tokens_.empty()) {
    return true;
  }
  switch (tokens_.back().kind) {
    case TokenKind::Whitespace:
    case TokenKind::LParen:
    case TokenKind::LBrace:
    case TokenKind::Comma:
    case TokenKind::Semicolon:
    case TokenKind::Colon:
    case TokenKind::Plus:
    case TokenKind::Minus:
    case TokenKind::Star:
    case TokenKind::Slash:
    case TokenKind::Caret:
    case TokenKind::Ampersand:
    case TokenKind::Eq:
    case TokenKind::NotEq:
    case TokenKind::Lt:
    case TokenKind::LtEq:
    case TokenKind::Gt:
    case TokenKind::GtEq:
    case TokenKind::At:
      return true;
    default:
      return false;
  }
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

void Tokenizer::scan_colon() {
  const std::size_t start = byte_pos_;
  mark_start();
  const bool leading = source_[byte_pos_] == '.';
  if (leading) {
    advance_one();
  }
  advance_one();  // ':'
  // `:.` only before a reference endpoint (`A1:.A10`, `1:.1`, `A:.$A`).
  bool trailing = false;
  if (byte_pos_ + 1 < source_.size() && source_[byte_pos_] == '.') {
    const char next = source_[byte_pos_ + 1];
    if (strings::is_ascii_alpha(next) || strings::is_ascii_digit(next) || next == '$') {
      trailing = true;
      advance_one();
    }
  }
  emit(TokenKind::Colon, start);
  tokens_.back().trim = leading && trailing ? TrimRefMode::Both
                        : leading           ? TrimRefMode::Leading
                        : trailing          ? TrimRefMode::Trailing
                                            : TrimRefMode::None;
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
