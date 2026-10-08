// Locale rewriting of stored formula text.
//
// The stored text is tokenized and re-emitted token by token. Token lexemes
// are views into the source, so the bytes between tokens (and any run the
// tokenizer skipped) are copied verbatim; only the token kinds Excel
// localizes are replaced. Structured-reference brackets are copied as-is
// because no capture covers their localized form.

#include "eval/formula_localize.h"

#include <cstddef>
#include <cstring>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "eval/locale_function_names.h"
#include "eval/locale_text.h"
#include "excel_locale.h"
#include "parser/token.h"
#include "parser/tokenizer.h"
#include "utils/strings.h"

namespace formulon {
namespace eval {
namespace {

// Scans the generated records for the first one `match(canonical, field)`
// accepts in `locale`'s column, skipping empty (English) fields.
template <typename Match>
const char* find_function_name(ExcelLocale locale, Match match) noexcept {
  const int column = locale_function_name_column(locale);
  if (column < 0) {
    return nullptr;
  }
  const char* p = kLocaleFunctionNames;
  const char* const end = kLocaleFunctionNames + kLocaleFunctionNamesSize;
  while (p < end) {
    const char* canonical = p;
    p += std::strlen(p) + 1;
    const char* field = nullptr;
    for (int c = 0; c < kLocaleFunctionNameColumns; ++c) {
      if (c == column) {
        field = p;
      }
      p += std::strlen(p) + 1;
    }
    if (*field != '\0') {
      if (const char* hit = match(canonical, field); hit != nullptr) {
        return hit;
      }
    }
  }
  return nullptr;
}

bool is_function_name_token(parser::TokenKind kind) noexcept {
  // `LOG10(` lexes as a cell reference and `TRUE(` as a boolean.
  return kind == parser::TokenKind::Ident || kind == parser::TokenKind::CellRef || kind == parser::TokenKind::Bool;
}

// An exponent literal reads back as its regenerated value (`1.5E-3` ->
// `0.0015`, locale_tokens.formulatext_sci_number); other literals keep
// their digits and only take the locale decimal separator.
void append_number(const parser::Token& token, char decimal_separator, std::string& out) {
  if (token.lexeme.find_first_of("Ee") != std::string_view::npos) {
    out.append(locale_number_text(token.number));
    return;
  }
  for (const char c : token.lexeme) {
    out.push_back(c == '.' ? decimal_separator : c);
  }
}

}  // namespace

const char* localized_function_name(ExcelLocale locale, std::string_view canonical) noexcept {
  return find_function_name(locale, [canonical](const char* name, const char* field) -> const char* {
    return strings::case_insensitive_eq(name, canonical) ? field : nullptr;
  });
}

const char* canonical_function_name(ExcelLocale locale, std::string_view localized) noexcept {
  return find_function_name(locale, [localized](const char* name, const char* field) -> const char* {
    return strings::case_insensitive_eq(field, localized) ? name : nullptr;
  });
}

std::string localize_formula_text(std::string_view formula) {
  const ExcelProfile profile = current_eval_profile();
  const LocaleFacts& facts = locale_facts(profile);
  parser::Tokenizer tokenizer(formula);
  const std::vector<parser::Token>& tokens = tokenizer.tokens();

  std::string out;
  out.reserve(formula.size() + formula.size() / 4);
  std::size_t copied = 0;
  int braces = 0;
  int brackets = 0;
  for (std::size_t i = 0; i < tokens.size(); ++i) {
    const parser::Token& token = tokens[i];
    if (token.kind == parser::TokenKind::Eof) {
      break;
    }
    const auto start = static_cast<std::size_t>(token.lexeme.data() - formula.data());
    out.append(formula.substr(copied, start - copied));
    copied = start + token.lexeme.size();

    if (token.kind == parser::TokenKind::LBracket) {
      ++brackets;
    } else if (token.kind == parser::TokenKind::RBracket && brackets > 0) {
      --brackets;
      out.append(token.lexeme);
      continue;
    }
    if (brackets > 0) {
      out.append(token.lexeme);
      continue;
    }

    const bool is_call = i + 1 < tokens.size() && tokens[i + 1].kind == parser::TokenKind::LParen;
    if (is_call && is_function_name_token(token.kind)) {
      const char* local = localized_function_name(profile.locale, token.lexeme);
      out.append(local != nullptr ? std::string_view(local) : token.lexeme);
      continue;
    }
    switch (token.kind) {
      case parser::TokenKind::LBrace:
        ++braces;
        out.append(token.lexeme);
        break;
      case parser::TokenKind::RBrace:
        braces = braces > 0 ? braces - 1 : 0;
        out.append(token.lexeme);
        break;
      case parser::TokenKind::Comma:
        out.push_back(braces > 0 ? facts.array_column_separator : facts.list_separator);
        break;
      case parser::TokenKind::Semicolon:
        if (braces > 0) {
          out.push_back(facts.array_row_separator);
        } else {
          out.append(token.lexeme);
        }
        break;
      case parser::TokenKind::Number:
        append_number(token, facts.decimal_separator, out);
        break;
      case parser::TokenKind::Bool:
        out.append(locale_bool_text(token.boolean));
        break;
      case parser::TokenKind::ErrorLiteral:
        out.append(locale_error_text(token.error_code));
        break;
      default:
        out.append(token.lexeme);
        break;
    }
  }
  out.append(formula.substr(copied));
  return out;
}

}  // namespace eval
}  // namespace formulon
