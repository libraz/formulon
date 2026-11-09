// Locale rewriting of formula text, in both directions.
//
// The text is tokenized and re-emitted token by token. Token lexemes
// are views into the source, so the bytes between tokens (and any run the
// tokenizer skipped) are copied verbatim; only the token kinds Excel
// localizes are replaced. A structured-reference bracket group is rewritten
// separately: only its `#Spec` words and the separator between bracketed
// items change; column names are copied as-is. Each written piece can be
// recorded with the input span it came from, so a position in the result
// maps back to the input.

#include "eval/formula_localize.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "eval/locale_function_names.h"
#include "excel_locale.h"
#include "parser/token.h"
#include "parser/tokenizer.h"
#include "utils/double_format.h"
#include "utils/strings.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"

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
    const std::size_t from = out.size();
    format_formula_number(out, token.number);
    std::replace(out.begin() + static_cast<std::ptrdiff_t>(from), out.end(), '.', decimal_separator);
    return;
  }
  for (const char c : token.lexeme) {
    out.push_back(c == '.' ? decimal_separator : c);
  }
}

// Item specifiers in ExcelLocale order: All, Data, Headers, Totals.
constexpr const char* kCanonicalSpecifiers[4] = {"All", "Data", "Headers", "Totals"};
constexpr const char* kSpecifierWords[14][4] = {
    {"すべて", "データ", "見出し", "集計"},
    {"All", "Data", "Headers", "Totals"},
    {"Alle", "Daten", "Kopfzeilen", "Ergebnisse"},
    {"Tout", "Données", "En-têtes", "Totaux"},
    {"全部", "数据", "标题", "汇总"},
    {"모두", "데이터", "머리글", "요약"},
    {"All", "Data", "Headers", "Totals"},
    {"Todo", "Datos", "Encabezados", "Totales"},
    {"Todo", "Datos", "Encabezados", "Totales"},
    {"Tudo", "Dados", "Cabeçalhos", "Totais"},
    {"Все", "Данные", "Заголовки", "Итоги"},
    {"全部", "資料", "標題", "總計"},
    {"Tutti", "Dati", "Intestazioni", "Totali"},
    {"Alles", "Gegevens", "Kopteksten", "Totalen"},
};

// Spelling of the specifier `word` (no `#`) in the other direction, or nullptr.
const char* translate_specifier(std::string_view word, ExcelLocale locale, bool to_locale) noexcept {
  const char* const* local = kSpecifierWords[static_cast<std::size_t>(locale)];
  for (std::size_t k = 0; k < 4; ++k) {
    if (strings::case_insensitive_eq(word, to_locale ? kCanonicalSpecifiers[k] : local[k])) {
      return to_locale ? local[k] : kCanonicalSpecifiers[k];
    }
  }
  return nullptr;
}

// Records that `out[out_begin, out.size())` was written for `f[in_begin,
// in_end)`, merging a verbatim piece into the verbatim piece it continues.
void record_span(std::vector<FormulaSourceSpan>* spans, std::string_view f, std::size_t in_begin, std::size_t in_end,
                 const std::string& out, std::size_t out_begin) {
  if (spans == nullptr || (in_begin == in_end && out_begin == out.size())) {
    return;
  }
  const bool verbatim = std::string_view(out).substr(out_begin) == f.substr(in_begin, in_end - in_begin);
  if (verbatim && !spans->empty() && spans->back().verbatim && spans->back().in_end == in_begin &&
      spans->back().out_end == out_begin) {
    spans->back().in_end = in_end;
    spans->back().out_end = out.size();
    return;
  }
  spans->push_back(FormulaSourceSpan{out_begin, out.size(), in_begin, in_end, verbatim});
}

// Rewrites the bracket group opening at `begin` (a `[`) into `out` and returns
// the offset past its closing `]`. A `'` escapes the next byte; the separator
// is only the one that directly follows a closing `]`, so `t[a,b]` is kept.
std::size_t rewrite_structured_group(std::string_view f, std::size_t begin, ExcelProfile profile, bool to_locale,
                                     const LocaleFacts& facts, std::string& out,
                                     std::vector<FormulaSourceSpan>* spans) {
  const char from = to_locale ? ',' : facts.list_separator;
  int depth = 0;
  std::size_t i = begin;
  bool closed = false;
  while (i < f.size() && !closed) {
    const std::size_t in_begin = i;
    const std::size_t out_begin = out.size();
    const char c = f[i];
    const char* word = nullptr;
    std::size_t close = std::string_view::npos;
    if (c == '#' && depth > 0) {
      close = f.find_first_of("[]", i + 1);
      if (close != std::string_view::npos && f[close] == ']') {
        word = translate_specifier(f.substr(i + 1, close - i - 1), profile.locale, to_locale);
      }
    }
    if (c == '\'' && i + 1 < f.size()) {
      out.append(f.substr(i, 2));
      i += 2;
    } else if (word != nullptr) {
      out.push_back('#');
      out.append(word);
      i = close;
    } else if (c == from && depth == 1 && f[i - 1] == ']') {
      out.push_back(to_locale ? facts.list_separator : ',');
      ++i;
    } else {
      out.push_back(c);
      ++i;
      if (c == '[') {
        ++depth;
      } else if (c == ']' && --depth == 0) {
        closed = true;
      }
    }
    record_span(spans, f, in_begin, i, out, out_begin);
  }
  return i;
}

enum class Direction { kToLocale, kToCanonical };

// Writes the rewritten spelling of `tokens[i]` (a non-bracket token) to `out`.
void rewrite_token(const std::vector<parser::Token>& tokens, std::size_t i, ExcelProfile profile, bool to_locale,
                   const LocaleFacts& facts, int& braces, std::string& out) {
  const parser::Token& token = tokens[i];
  const bool is_qualified_name =
      i > 0 && tokens[i - 1].kind == parser::TokenKind::Bang && is_function_name_token(token.kind);
  if (is_qualified_name) {
    // A qualified name (`Sheet1!TRUE`, `[0]!SUM`) is scoped, never a builtin to localize.
    out.append(token.lexeme);
    return;
  }
  const bool is_call = i + 1 < tokens.size() && tokens[i + 1].kind == parser::TokenKind::LParen;
  if (is_call && is_function_name_token(token.kind)) {
    const char* renamed = to_locale ? localized_function_name(profile.locale, token.lexeme)
                                    : canonical_function_name(profile.locale, token.lexeme);
    out.append(renamed != nullptr ? std::string_view(renamed) : token.lexeme);
    return;
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
      if (!to_locale) {
        out.push_back(',');
      } else {
        out.push_back(braces > 0 ? facts.array_column_separator : facts.list_separator);
      }
      break;
    case parser::TokenKind::Semicolon:
      if (braces > 0) {
        out.push_back(to_locale ? facts.array_row_separator : ';');
      } else {
        out.append(token.lexeme);
      }
      break;
    case parser::TokenKind::Number:
      if (to_locale) {
        append_number(token, facts.decimal_separator, out);
      } else {
        for (const char c : token.lexeme) {
          out.push_back(c == facts.decimal_separator ? '.' : c);
        }
      }
      break;
    case parser::TokenKind::Bool:
      if (to_locale) {
        out.append(token.boolean ? facts.true_name : facts.false_name);
      } else {
        out.append(token.boolean ? "TRUE" : "FALSE");
      }
      break;
    case parser::TokenKind::ErrorLiteral: {
      const auto ordinal = static_cast<std::size_t>(token.error_code);
      if (to_locale) {
        out.append(facts.error_names[ordinal]);
      } else {
        out.append(kErrorTable[ordinal].display_name);
      }
      break;
    }
    default:
      out.append(token.lexeme);
      break;
  }
}

// Re-emits `formula` token by token. Inter-token bytes are copied verbatim
// and only the token kinds Excel localizes change; nothing is validated.
// `spans`, when non-null, receives the position map of the result.
std::string rewrite_formula_text(std::string_view formula, ExcelProfile profile, Direction direction,
                                 std::vector<FormulaSourceSpan>* spans) {
  const bool to_locale = direction == Direction::kToLocale;
  const LocaleFacts& facts = locale_facts(profile);
  parser::TokenizerOptions options;
  if (!to_locale) {
    options.locale = &facts;
  }
  parser::Tokenizer tokenizer(formula, options);
  const std::vector<parser::Token>& tokens = tokenizer.tokens();

  std::string out;
  out.reserve(formula.size() + formula.size() / 4);
  std::size_t copied = 0;
  int braces = 0;
  for (std::size_t i = 0; i < tokens.size(); ++i) {
    const parser::Token& token = tokens[i];
    if (token.kind == parser::TokenKind::Eof) {
      break;
    }
    const auto start = static_cast<std::size_t>(token.lexeme.data() - formula.data());
    if (start < copied) {
      continue;
    }
    std::size_t out_begin = out.size();
    out.append(formula.substr(copied, start - copied));
    record_span(spans, formula, copied, start, out, out_begin);

    if (token.kind == parser::TokenKind::LBracket) {
      copied = rewrite_structured_group(formula, start, profile, to_locale, facts, out, spans);
      continue;
    }
    copied = start + token.lexeme.size();
    out_begin = out.size();
    rewrite_token(tokens, i, profile, to_locale, facts, braces, out);
    record_span(spans, formula, start, copied, out, out_begin);
  }
  const std::size_t out_begin = out.size();
  out.append(formula.substr(copied));
  record_span(spans, formula, copied, formula.size(), out, out_begin);
  return out;
}

// Line and column of byte `offset` in `text`, counted as the tokenizer counts
// them: CR, LF and CRLF each end a line, and a column is one UTF-16 unit.
void line_and_column(std::string_view text, std::size_t offset, parser::TextRange& range) {
  std::uint32_t line = 1;
  std::size_t line_start = 0;
  for (std::size_t i = 0; i < offset && i < text.size(); ++i) {
    if (text[i] == '\r' && i + 1 < text.size() && text[i + 1] == '\n' && i + 1 < offset) {
      continue;
    }
    if (text[i] == '\n' || text[i] == '\r') {
      ++line;
      line_start = i + 1;
    }
  }
  range.line = line;
  range.column = 1U + utf16_units_in(text.substr(line_start, offset - line_start));
}

}  // namespace

parser::TextRange MappedFormulaText::source_range(std::string_view source, parser::TextRange range) const {
  // Offsets of `range` in bytes of `text`; the start rounds down, the end up.
  const std::size_t out_start = utf16_to_byte_offset_floor(text, range.start);
  const std::size_t out_end = std::max(out_start, utf16_to_byte_offset(text, range.end));
  // A start at the end of `text` and an end at its beginning match no span.
  std::size_t in_start = source.size();
  std::size_t in_end = out_end == 0 ? 0 : source.size();
  for (const FormulaSourceSpan& span : spans) {
    if (out_start >= span.out_begin && out_start < span.out_end) {
      in_start = span.verbatim ? span.in_begin + (out_start - span.out_begin) : span.in_begin;
    }
    if (out_end > span.out_begin && out_end <= span.out_end) {
      in_end = span.verbatim ? span.in_begin + (out_end - span.out_begin) : span.in_end;
    }
  }
  in_end = std::max(in_start, std::min(in_end, source.size()));
  parser::TextRange mapped;
  mapped.start = utf16_units_in(source.substr(0, in_start));
  mapped.end = utf16_units_in(source.substr(0, in_end));
  line_and_column(source, in_start, mapped);
  return mapped;
}

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

std::string localize_formula_text(std::string_view formula, ExcelProfile profile) {
  return rewrite_formula_text(formula, profile, Direction::kToLocale, nullptr);
}

std::string canonicalize_formula_text(std::string_view formula, ExcelProfile profile) {
  return rewrite_formula_text(formula, profile, Direction::kToCanonical, nullptr);
}

MappedFormulaText localize_formula_text_mapped(std::string_view formula, ExcelProfile profile) {
  MappedFormulaText mapped;
  mapped.text = rewrite_formula_text(formula, profile, Direction::kToLocale, &mapped.spans);
  return mapped;
}

MappedFormulaText canonicalize_formula_text_mapped(std::string_view formula, ExcelProfile profile) {
  MappedFormulaText mapped;
  mapped.text = rewrite_formula_text(formula, profile, Direction::kToCanonical, &mapped.spans);
  return mapped;
}

}  // namespace eval
}  // namespace formulon
