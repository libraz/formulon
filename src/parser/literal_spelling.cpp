//
// Number-literal respelling of formula text. See `literal_spelling.h` for
// the contract.

#include "parser/literal_spelling.h"

#include <cstddef>
#include <vector>

#include "parser/token.h"
#include "parser/tokenizer.h"
#include "utils/double_format.h"

namespace formulon {
namespace parser {
namespace {

// True when a token of `kind` can end an operand, making a following sign
// binary. A rejected run counts as one, so no sign after it is taken.
bool ends_operand(TokenKind kind) noexcept {
  switch (kind) {
    case TokenKind::Invalid:
    case TokenKind::Number:
    case TokenKind::String:
    case TokenKind::Bool:
    case TokenKind::ErrorLiteral:
    case TokenKind::Ident:
    case TokenKind::SheetName:
    case TokenKind::ExternalQualifier:
    case TokenKind::CellRef:
    case TokenKind::RParen:
    case TokenKind::RBrace:
    case TokenKind::RBracket:
    case TokenKind::Percent:
    case TokenKind::Hash:
      return true;
    default:
      return false;
  }
}

// Index of the nearest non-whitespace token before `i`, or `i` when none.
std::size_t previous_significant(const std::vector<Token>& tokens, std::size_t i) noexcept {
  for (std::size_t j = i; j > 0; --j) {
    if (tokens[j - 1].kind != TokenKind::Whitespace) {
      return j - 1;
    }
  }
  return i;
}

// Index of the nearest non-whitespace token after `i` (the `Eof` at worst).
std::size_t next_significant(const std::vector<Token>& tokens, std::size_t i) noexcept {
  std::size_t j = i + 1;
  while (j + 1 < tokens.size() && tokens[j].kind == TokenKind::Whitespace) {
    ++j;
  }
  return j;
}

// A number beside a `:` is a row of a row range (`1:3`), not a literal.
bool is_row_number(const std::vector<Token>& tokens, std::size_t i) noexcept {
  const std::size_t prev = previous_significant(tokens, i);
  if (prev != i && tokens[prev].kind == TokenKind::Colon) {
    return true;
  }
  const std::size_t next = next_significant(tokens, i);
  return next < tokens.size() && tokens[next].kind == TokenKind::Colon;
}

// True when the token before the number at `i` is a unary sign written
// directly against it.
bool has_touching_unary_sign(const std::vector<Token>& tokens, std::size_t i) noexcept {
  if (i == 0) {
    return false;
  }
  const Token& sign = tokens[i - 1];
  if (sign.kind != TokenKind::Plus && sign.kind != TokenKind::Minus) {
    return false;
  }
  if (sign.lexeme.data() + sign.lexeme.size() != tokens[i].lexeme.data()) {
    return false;
  }
  const std::size_t prev = previous_significant(tokens, i - 1);
  return prev == i - 1 || !ends_operand(tokens[prev].kind);
}

}  // namespace

std::string spell_number_literals(std::string_view formula) {
  std::string_view body = formula;
  if (!body.empty() && body.front() == '=') {
    body.remove_prefix(1);
  }
  Tokenizer tokenizer(body);
  const std::vector<Token>& tokens = tokenizer.tokens();

  std::string out;
  out.reserve(formula.size());
  std::size_t copied = 0;
  int brackets = 0;
  for (std::size_t i = 0; i < tokens.size(); ++i) {
    const Token& token = tokens[i];
    if (token.kind == TokenKind::Eof) {
      break;
    }
    if (token.kind == TokenKind::LBracket) {
      ++brackets;
    } else if (token.kind == TokenKind::RBracket && brackets > 0) {
      --brackets;
    }
    if (token.kind != TokenKind::Number || brackets > 0 || is_row_number(tokens, i)) {
      continue;
    }
    std::size_t from = static_cast<std::size_t>(token.lexeme.data() - formula.data());
    double value = token.number;
    if (has_touching_unary_sign(tokens, i)) {
      from -= tokens[i - 1].lexeme.size();
      if (tokens[i - 1].kind == TokenKind::Minus) {
        value = -value;
      }
    }
    out.append(formula.substr(copied, from - copied));
    format_formula_number(out, value);
    copied = static_cast<std::size_t>(token.lexeme.data() - formula.data()) + token.lexeme.size();
  }
  out.append(formula.substr(copied));
  return out;
}

}  // namespace parser
}  // namespace formulon
