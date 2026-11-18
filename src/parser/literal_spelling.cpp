//
// Number-literal respelling and whitespace dropping of formula text. See `literal_spelling.h` for
// the contract.

#include "parser/literal_spelling.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "parser/ast_format.h"
#include "parser/parser_detail.h"
#include "parser/token.h"
#include "parser/tokenizer.h"
#include "utils/double_format.h"
#include "utils/formula_text.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {
namespace {

// Walks the tokens of a formula (leading `=` dropped) up to the `Eof`, tracking
// the depth of `[...]` and of `{...}` outside brackets.
struct TokenWalk {
  explicit TokenWalk(std::string_view formula);

  // Moves to the next token and updates the depths; false at the `Eof`.
  bool advance() noexcept;

  Tokenizer tokenizer;
  const std::vector<Token>& tokens;
  std::size_t i = static_cast<std::size_t>(-1);
  int brackets = 0;
  int braces = 0;
};

[[gnu::noinline]] TokenWalk::TokenWalk(std::string_view formula)
    : tokenizer(strip_formula_prefix(formula)), tokens(tokenizer.tokens()) {}

[[gnu::noinline]] bool TokenWalk::advance() noexcept {
  if (++i >= tokens.size() || tokens[i].kind == TokenKind::Eof) {
    return false;
  }
  const TokenKind kind = tokens[i].kind;
  if (kind == TokenKind::LBracket) {
    ++brackets;
  } else if (kind == TokenKind::RBracket && brackets > 0) {
    --brackets;
  } else if (brackets == 0 && kind == TokenKind::LBrace) {
    ++braces;
  } else if (brackets == 0 && kind == TokenKind::RBrace && braces > 0) {
    --braces;
  }
  return true;
}

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

std::string ascii_upper(std::string_view text) {
  std::string out(text);
  for (char& c : out) {
    c = strings::ascii_to_upper(c);
  }
  return out;
}

// True when `token` is a column endpoint (`c`, `$C`) of a whole-column range.
bool is_column_token(const Token& token) noexcept {
  if (token.kind != TokenKind::Ident) {
    return false;
  }
  std::string_view letters = token.lexeme;
  if (!letters.empty() && letters.front() == '$') {
    letters.remove_prefix(1);
  }
  return detail::DecodeColumnLetters(letters) != 0U;
}

// True when the column-shaped name at `i` is one end of a `:` whose other end
// is column-shaped too (`c:c`, `$a:B`).
bool is_column_endpoint(const std::vector<Token>& tokens, std::size_t i) noexcept {
  if (!is_column_token(tokens[i])) {
    return false;
  }
  const std::size_t prev = previous_significant(tokens, i);
  if (prev != i && tokens[prev].kind == TokenKind::Colon) {
    const std::size_t other = previous_significant(tokens, prev);
    if (other != prev && is_column_token(tokens[other])) {
      return true;
    }
  }
  const std::size_t next = next_significant(tokens, i);
  if (next < tokens.size() && tokens[next].kind == TokenKind::Colon) {
    const std::size_t other = next_significant(tokens, next);
    return other < tokens.size() && is_column_token(tokens[other]);
  }
  return false;
}

// True when `kind` can spell a sheet before a `!`.
bool may_name_sheet(TokenKind kind) noexcept {
  return kind == TokenKind::Ident || kind == TokenKind::CellRef || kind == TokenKind::SheetName ||
         kind == TokenKind::Bool;
}

// True when token `j` follows token `i` with nothing in between.
bool touches(const std::vector<Token>& tokens, std::size_t i, std::size_t j) noexcept {
  return j < tokens.size() && tokens[i].lexeme.data() + tokens[i].lexeme.size() == tokens[j].lexeme.data();
}

// The unquoted sheet qualifying the name at `i` (after its `!`): empty for a
// 3-D span or an external book, which keep the unqualified lookup.
std::string_view name_qualifier(const std::vector<Token>& tokens, std::size_t i) {
  if (i < 2U || !may_name_sheet(tokens[i - 2].kind) || !touches(tokens, i - 2, i - 1)) {
    return {};
  }
  if (i >= 3U && touches(tokens, i - 3, i - 2) &&
      (tokens[i - 3].kind == TokenKind::Colon || tokens[i - 3].kind == TokenKind::RBracket)) {
    return {};
  }
  const Token& sheet = tokens[i - 2];
  const std::string_view text = sheet.kind == TokenKind::SheetName ? sheet.text : sheet.lexeme;
  return text.find(':') == std::string_view::npos && text.find('[') == std::string_view::npos ? text
                                                                                              : std::string_view();
}

// `name` without a typed `_xlfn.` / `_xlfn._xlws.` / `_xlws.` storage prefix.
std::string_view without_storage_prefix(std::string_view name) noexcept {
  for (const std::string_view prefix : {std::string_view("_xlfn."), std::string_view("_xlws.")}) {
    if (name.size() > prefix.size() && strings::case_insensitive_eq(name.substr(0, prefix.size()), prefix)) {
      name.remove_prefix(prefix.size());
    }
  }
  return name;
}

// A name and the spelling every later occurrence takes.
struct Spelling {
  std::string key;
  std::string_view text;
};

const Spelling* find_spelling(const std::vector<Spelling>& spellings, std::string_view name) {
  const std::string key = strings::to_ascii_lower(name);
  for (std::size_t k = spellings.size(); k > 0; --k) {
    if (spellings[k - 1].key == key) {
      return &spellings[k - 1];
    }
  }
  return nullptr;
}

// A `(` open while the walk is inside it: a call (LET / LAMBDA bind their
// parameters) or a grouping.
struct OpenParen {
  bool let = false;
  bool lambda = false;
  std::uint32_t argument = 0;
  std::size_t scope_size = 0;
};

}  // namespace

std::string spell_number_literals(std::string_view formula) {
  TokenWalk walk(formula);
  const std::vector<Token>& tokens = walk.tokens;

  std::string out;
  out.reserve(formula.size());
  std::size_t copied = 0;
  while (walk.advance()) {
    const std::size_t i = walk.i;
    const Token& token = tokens[i];
    // A `$` lexeme is a row anchor (`$1`), not a literal, even where no `:` follows.
    if (token.kind != TokenKind::Number || walk.brackets > 0 || is_row_number(tokens, i) ||
        token.lexeme.front() == '$') {
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

std::string drop_unkept_whitespace(std::string_view formula) {
  TokenWalk walk(formula);
  const std::vector<Token>& tokens = walk.tokens;
  // Adjacency is checked by byte offset, since a run the tokenizer only
  // records as an error has no token of its own.
  const auto begin_of = [&formula](const Token& token) {
    return static_cast<std::size_t>(token.lexeme.data() - formula.data());
  };
  const auto end_of = [&begin_of](const Token& token) { return begin_of(token) + token.lexeme.size(); };

  std::string out;
  out.reserve(formula.size());
  std::size_t copied = 0;
  while (walk.advance()) {
    const std::size_t i = walk.i;
    const Token& token = tokens[i];
    if (token.kind != TokenKind::Whitespace || walk.brackets > 0) {
      continue;
    }
    const std::size_t from = begin_of(token);
    const std::size_t to = end_of(token);
    const Token& next = tokens[i + 1];
    const bool before_separator =
        (next.kind == TokenKind::Comma || next.kind == TokenKind::Colon) && begin_of(next) == to;
    const bool after_colon = i > 0 && tokens[i - 1].kind == TokenKind::Colon && end_of(tokens[i - 1]) == from;
    out.append(formula.substr(copied, from - copied));
    copied = to;
    if (to != formula.size() && !before_separator && !after_colon && walk.braces == 0) {
      // Spaces directly before a line feed go.
      const std::string_view run = token.lexeme;
      for (std::size_t k = 0; k < run.size();) {
        const std::size_t end = run[k] == ' ' ? std::min(run.find_first_not_of(' ', k), run.size()) : k + 1U;
        if (run[k] != ' ' || end == run.size() || run[end] != '\n') {
          out.append(run.substr(k, end - k));
        }
        k = end;
      }
    }
  }
  out.append(formula.substr(copied));
  return out;
}

bool has_refused_whitespace(std::string_view formula, bool (*is_function_name)(std::string_view)) {
  TokenWalk walk(formula);
  const std::vector<Token>& tokens = walk.tokens;
  while (walk.advance()) {
    const std::size_t i = walk.i;
    const Token& token = tokens[i];
    if (token.kind != TokenKind::Whitespace || walk.brackets > 0) {
      continue;
    }
    if (token.lexeme.find_first_of("\t\r") != std::string_view::npos) {
      return true;
    }
    // An unqualified name before the run, a `(` after it.
    if (i > 0 && tokens[i + 1].kind == TokenKind::LParen &&
        (tokens[i - 1].kind == TokenKind::Ident || tokens[i - 1].kind == TokenKind::CellRef) &&
        (i < 2 || tokens[i - 2].kind != TokenKind::Bang) && is_function_name(tokens[i - 1].lexeme)) {
      return true;
    }
  }
  return false;
}

std::string spell_entered_formula(std::string_view formula) {
  return spell_number_literals(drop_unkept_whitespace(formula));
}

std::string spell_entered_names(std::string_view formula, const EnteredNames& names) {
  TokenWalk walk(formula);
  const std::vector<Token>& tokens = walk.tokens;
  const auto begin_of = [&formula](const Token& token) {
    return static_cast<std::size_t>(token.lexeme.data() - formula.data());
  };

  std::string out;
  out.reserve(formula.size());
  std::size_t copied = 0;
  const auto replace = [&](const Token& token, std::string_view text) {
    out.append(formula.substr(copied, begin_of(token) - copied));
    out.append(text);
    copied = begin_of(token) + token.lexeme.size();
  };

  std::vector<OpenParen> parens;
  std::vector<Spelling> scope;
  OpenParen pending;
  std::size_t skip_to = 0;
  while (walk.advance()) {
    const std::size_t i = walk.i;
    const Token& token = tokens[i];
    if (walk.brackets > 0 || token.kind == TokenKind::RBracket || i < skip_to) {
      continue;
    }
    const TokenKind kind = token.kind;
    if (kind == TokenKind::LParen) {
      pending.scope_size = scope.size();
      parens.push_back(pending);
      pending = OpenParen();
      continue;
    }
    if (kind == TokenKind::RParen) {
      if (!parens.empty()) {
        scope.resize(parens.back().scope_size);
        parens.pop_back();
      }
      continue;
    }
    if (kind == TokenKind::Comma && walk.braces == 0) {
      if (!parens.empty()) {
        ++parens.back().argument;
      }
      continue;
    }
    if (kind == TokenKind::ErrorLiteral) {
      replace(token, ascii_upper(token.lexeme));
      continue;
    }

    // A sheet qualifier: `name!` or the 3-D `first:last!`.
    if (may_name_sheet(kind) && touches(tokens, i, i + 1)) {
      const TokenKind next = tokens[i + 1].kind;
      if (next == TokenKind::Bang) {
        const std::string_view typed = kind == TokenKind::SheetName ? token.text : token.lexeme;
        const std::string_view own = typed.find(':') == std::string_view::npos
                                         ? names.own(names.ctx, EnteredNameKind::kSheet, typed, {})
                                         : std::string_view();
        if (!own.empty()) {
          std::string spelled;
          append_sheet_name(own, false, spelled);
          replace(token, spelled);
        }
        continue;
      }
      if (next == TokenKind::Colon && kind != TokenKind::SheetName && touches(tokens, i + 1, i + 2) &&
          may_name_sheet(tokens[i + 2].kind) && tokens[i + 2].kind != TokenKind::SheetName &&
          touches(tokens, i + 2, i + 3) && tokens[i + 3].kind == TokenKind::Bang) {
        const std::string_view first = names.own(names.ctx, EnteredNameKind::kSheet, token.lexeme, {});
        const std::string_view last = names.own(names.ctx, EnteredNameKind::kSheet, tokens[i + 2].lexeme, {});
        std::string first_spelled;
        std::string last_spelled;
        append_sheet_name(first, false, first_spelled);
        append_sheet_name(last, false, last_spelled);
        if (!first.empty() && !last.empty() && first_spelled == first && last_spelled == last) {
          replace(token, first);
          replace(tokens[i + 2], last);
        }
        // Any other 3-D qualifier stays as typed.
        skip_to = i + 3;
        continue;
      }
    }

    const bool qualified = i > 0 && tokens[i - 1].kind == TokenKind::Bang;
    if ((kind == TokenKind::Ident || kind == TokenKind::CellRef || kind == TokenKind::Bool) &&
        touches(tokens, i, i + 1) && tokens[i + 1].kind == TokenKind::LParen) {
      if (qualified) {
        continue;
      }
      if (kind == TokenKind::Bool) {
        replace(token, ascii_upper(token.lexeme));
      } else if (const Spelling* param = find_spelling(scope, token.lexeme)) {
        replace(token, param->text);
      } else if (const char* builtin = names.builtin(without_storage_prefix(token.lexeme))) {
        const std::string_view canonical(builtin);
        pending.let = canonical == "LET";
        pending.lambda = canonical == "LAMBDA";
        replace(token, canonical);
      } else {
        replace(token, names.unknown_call(names.ctx, token.lexeme));
      }
      continue;
    }
    if (kind == TokenKind::Bool) {
      replace(token, ascii_upper(token.lexeme));
      continue;
    }
    if (kind != TokenKind::Ident && kind != TokenKind::CellRef) {
      continue;
    }
    // A parameter, cell-shaped (`B1`) or not, is a whole argument of LET
    // (every other one but the last) or LAMBDA (every one but the last).
    if (!qualified && !parens.empty() && (parens.back().let || parens.back().lambda)) {
      const OpenParen& open = parens.back();
      const std::size_t prev = previous_significant(tokens, i);
      const std::size_t next = next_significant(tokens, i);
      const bool whole = prev != i &&
                         (tokens[prev].kind == TokenKind::LParen || tokens[prev].kind == TokenKind::Comma) &&
                         next < tokens.size() && tokens[next].kind == TokenKind::Comma;
      if (whole && (open.lambda || open.argument % 2U == 0U)) {
        scope.push_back(Spelling{strings::to_ascii_lower(token.lexeme), token.lexeme});
        continue;
      }
    }
    if (const Spelling* param = qualified ? nullptr : find_spelling(scope, token.lexeme)) {
      replace(token, param->text);
    } else if (kind == TokenKind::CellRef || is_column_endpoint(tokens, i)) {
      replace(token, ascii_upper(token.lexeme));
    } else {
      // A table before its `[...]`, else a defined name.
      const bool table = touches(tokens, i, i + 1) && tokens[i + 1].kind == TokenKind::LBracket;
      const std::string_view own = table ? names.own(names.ctx, EnteredNameKind::kTable, token.lexeme, {})
                                         : names.own(names.ctx, EnteredNameKind::kDefinedName, token.lexeme,
                                                     qualified ? name_qualifier(tokens, i) : std::string_view());
      if (!own.empty()) {
        replace(token, own);
      }
    }
  }
  out.append(formula.substr(copied));
  return out;
}

}  // namespace parser
}  // namespace formulon
