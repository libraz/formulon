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
//
// R1C1 localization takes the reference nodes of the parsed formula and
// replaces their cell parts inside the same token pass; R1C1 canonicalization
// reads the endpoints back into A1 first (`rewrite_r1c1_text`) and composes
// the two position maps.
//
// @size-budget: 8 KB (R1C1 notation and qualifier / name quoting)

#include "eval/formula_localize.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/a1_parse.h"
#include "eval/builtin_names.h"
#include "eval/formula_entry_r1c1.h"
#include "eval/locale_function_names.h"
#include "excel_locale.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_format_r1c1.h"
#include "parser/formula_prefix.h"
#include "parser/parse_error.h"
#include "parser/reference.h"
#include "parser/token.h"
#include "parser/tokenizer.h"
#include "utils/arena.h"
#include "utils/double_format.h"
#include "utils/formula_text.h"
#include "utils/index_sort.h"
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

// Which rule a local sheet qualifier is requoted by.
enum class QualifierRule : std::uint8_t { kStored, kA1Display, kR1C1Display };

struct RewriteRules {
  Direction direction = Direction::kToLocale;
  QualifierRule qualifiers = QualifierRule::kStored;
  /// R1C1 output prints a cell-shaped name bare (`_xlpm.B1` -> `B1`).
  bool bare_cell_names = false;
};

// The cell part of one reference in the formula and its R1C1 spelling.
struct ReferenceRewrite {
  std::size_t begin = 0;
  std::size_t end = 0;
  std::string text;
};

// Writes the rewritten spelling of `tokens[i]` (a non-bracket token) to `out`.
void rewrite_token(const std::vector<parser::Token>& tokens, std::size_t i, ExcelProfile profile,
                   const RewriteRules& rules, const LocaleFacts& facts, int& braces, std::string& out) {
  const bool to_locale = rules.direction == Direction::kToLocale;
  const parser::Token& token = tokens[i];
  const std::string_view lexeme = rules.bare_cell_names && token.kind == parser::TokenKind::Ident
                                      ? parser::cell_shaped_name_bare(token.lexeme)
                                      : token.lexeme;
  const bool is_qualified_name =
      i > 0 && tokens[i - 1].kind == parser::TokenKind::Bang && is_function_name_token(token.kind);
  if (is_qualified_name) {
    // A qualified name (`Sheet1!TRUE`, `[0]!SUM`) is scoped, never a builtin to localize.
    out.append(lexeme);
    return;
  }
  const bool is_call = i + 1 < tokens.size() && tokens[i + 1].kind == parser::TokenKind::LParen;
  if (is_call && is_function_name_token(token.kind)) {
    const char* renamed = to_locale ? localized_function_name(profile.locale, token.lexeme)
                                    : canonical_function_name(profile.locale, token.lexeme);
    out.append(renamed != nullptr ? std::string_view(renamed) : lexeme);
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
      out.append(lexeme);
      break;
  }
}

bool needs_quoting(std::string_view sheet, QualifierRule rule, const LocaleFacts& facts) {
  switch (rule) {
    case QualifierRule::kStored:
      return parser::local_sheet_needs_quoting_a1(sheet);
    case QualifierRule::kA1Display:
      return parser::local_sheet_needs_quoting_display(sheet, facts);
    case QualifierRule::kR1C1Display:
      return parser::local_sheet_needs_quoting_r1c1_display(sheet, facts);
  }
  return true;
}

std::size_t token_start(std::string_view formula, const parser::Token& token) {
  return static_cast<std::size_t>(token.lexeme.data() - formula.data());
}

// True when `tokens[i]` exists, is one of the sheet-name kinds, and starts
// where `tokens[i - 1]` ends.
bool sheet_token_at(std::string_view formula, const std::vector<parser::Token>& tokens, std::size_t i,
                    bool cell_shaped) {
  if (i >= tokens.size()) {
    return false;
  }
  const parser::TokenKind kind = tokens[i].kind;
  const bool sheet_kind = kind == parser::TokenKind::Ident || kind == parser::TokenKind::SheetName ||
                          (cell_shaped && kind == parser::TokenKind::CellRef);
  return sheet_kind && (i == 0U || token_start(formula, tokens[i - 1U]) + tokens[i - 1U].lexeme.size() ==
                                       token_start(formula, tokens[i]));
}

bool bang_at(std::string_view formula, const std::vector<parser::Token>& tokens, std::size_t i) {
  return i < tokens.size() && tokens[i].kind == parser::TokenKind::Bang &&
         token_start(formula, tokens[i - 1U]) + tokens[i - 1U].lexeme.size() == token_start(formula, tokens[i]);
}

std::string_view sheet_text(const parser::Token& token) {
  return token.kind == parser::TokenKind::SheetName ? token.text : token.lexeme;
}

// A local sheet qualifier starting at `tokens[i]`: a name and `!`, or two
// names joined by `:` and `!`. Writes it under `rule` (a partly quoted 3-D
// span as written) and returns the index of its `!`; 0 when none starts here.
std::size_t append_qualifier_at(std::string_view formula, const std::vector<parser::Token>& tokens, std::size_t i,
                                QualifierRule rule, const LocaleFacts& facts, std::string& out) {
  std::string_view first;
  std::string_view last;
  std::size_t bang = 0;
  bool partial = false;
  if (sheet_token_at(formula, tokens, i, true) && bang_at(formula, tokens, i + 1U)) {
    first = sheet_text(tokens[i]);
    bang = i + 1U;
    if (tokens[i].kind == parser::TokenKind::SheetName) {
      if (first.find('[') != std::string_view::npos) {
        return 0;  // a cross-workbook qualifier
      }
      if (const std::size_t colon = first.find(':'); colon != std::string_view::npos) {
        last = first.substr(colon + 1U);
        first = first.substr(0, colon);
      }
    }
  } else if (sheet_token_at(formula, tokens, i, false) && i + 1U < tokens.size() &&
             tokens[i + 1U].kind == parser::TokenKind::Colon && sheet_token_at(formula, tokens, i + 2U, true) &&
             bang_at(formula, tokens, i + 3U)) {
    first = sheet_text(tokens[i]);
    last = sheet_text(tokens[i + 2U]);
    bang = i + 3U;
    partial = tokens[i].kind == parser::TokenKind::SheetName || tokens[i + 2U].kind == parser::TokenKind::SheetName;
  } else {
    return 0;
  }
  if (partial) {
    out.append(formula.substr(token_start(formula, tokens[i]),
                              token_start(formula, tokens[bang]) - token_start(formula, tokens[i])));
    return bang;
  }
  const bool quote = needs_quoting(first, rule, facts) || (!last.empty() && needs_quoting(last, rule, facts));
  append_sheet_qualifier(out, first, last, quote);
  return bang;
}

// Re-emits `formula` token by token. Inter-token bytes are copied verbatim
// and only the token kinds Excel localizes change, with local sheet
// qualifiers requoted under `rules`; nothing is validated. `spans`, when
// non-null, receives the position map of the result. Each of `rewrites`
// (ascending, each starting at a token) replaces its bytes.
std::string rewrite_formula_text(std::string_view formula, ExcelProfile profile, const RewriteRules& rules,
                                 std::vector<FormulaSourceSpan>* spans, const std::vector<ReferenceRewrite>& rewrites) {
  const bool to_locale = rules.direction == Direction::kToLocale;
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
  std::size_t next_rewrite = 0;
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
    if (next_rewrite < rewrites.size() && rewrites[next_rewrite].begin <= start) {
      const ReferenceRewrite& rewrite = rewrites[next_rewrite++];
      std::size_t out_begin = out.size();
      out.append(formula.substr(copied, rewrite.begin - copied));
      record_span(spans, formula, copied, rewrite.begin, out, out_begin);
      out_begin = out.size();
      out.append(rewrite.text);
      record_span(spans, formula, rewrite.begin, rewrite.end, out, out_begin);
      copied = rewrite.end;
      continue;
    }
    std::size_t out_begin = out.size();
    out.append(formula.substr(copied, start - copied));
    record_span(spans, formula, copied, start, out, out_begin);

    if (token.kind == parser::TokenKind::LBracket) {
      copied = rewrite_structured_group(formula, start, profile, to_locale, facts, out, spans);
      continue;
    }
    out_begin = out.size();
    if (const std::size_t bang = append_qualifier_at(formula, tokens, i, rules.qualifiers, facts, out); bang != 0U) {
      copied = token_start(formula, tokens[bang]);
      record_span(spans, formula, start, copied, out, out_begin);
      continue;
    }
    copied = start + token.lexeme.size();
    rewrite_token(tokens, i, profile, rules, facts, braces, out);
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

bool is_endpoint_byte(char c) noexcept {
  return c == '$' || strings::is_ascii_alpha(c) || strings::is_ascii_digit(c);
}

// Collects the R1C1 spelling of every reference node of a parsed formula.
class ReferenceCollector {
 public:
  ReferenceCollector(std::string_view formula, std::size_t body_at, const LocaleFacts& facts,
                     const FormulaNotation& notation)
      : formula_(formula), body_(formula.substr(body_at)), body_at_(body_at), facts_(facts), notation_(notation) {}

  void visit(const parser::AstNode& node) {
    switch (node.kind()) {
      case parser::NodeKind::Ref:
        add(node, node.as_ref(), nullptr, false);
        return;
      case parser::NodeKind::SpillRef:
        if (node.as_spill_ref_anchor_expr() == nullptr) {
          add(node, node.as_spill_ref(), nullptr, true);
          return;
        }
        break;
      case parser::NodeKind::Ref3D:
        add(node, node.as_ref3d_cell(), node.as_ref3d_is_range() ? &node.as_ref3d_cell_end() : nullptr, false);
        return;
      case parser::NodeKind::ExternalRef:
        if (node.as_external_ref_name().empty()) {
          add(node, node.as_external_ref_cell(),
              node.as_external_ref_is_range() ? &node.as_external_ref_cell_end() : nullptr, false);
        }
        return;
      case parser::NodeKind::ErrorLiteral: {
        // A dead sheet's address (`#REF!A1`) is relative to the formula's cell too.
        const parser::RefErrorSite* site = node.as_ref_error_site();
        if (site != nullptr && site->form == parser::RefErrorSite::Form::kDeadSheet && site->has_address) {
          add(node, site->first, site->is_range ? &site->last : nullptr, false);
        }
        return;
      }
      default:
        break;
    }
    parser::any_child_node(node, [this](const parser::AstNode& child) {
      visit(child);
      return false;
    });
  }

  /// Quotes every unbound name that reads as a reference in the locale's
  /// R1C1 letters (de `z` -> `'z'`); LET / LAMBDA names stay as written.
  void quote_names(const parser::AstNode& root) {
    for_each_unbound_name(
        root,
        [](void* ctx, const parser::AstNode& node) {
          auto& self = *static_cast<ReferenceCollector*>(ctx);
          const std::string_view name = node.as_name();
          const refs_internal::R1C1Letters letters{self.facts_.r1c1_row, self.facts_.r1c1_col, self.facts_.r1c1_open,
                                                   self.facts_.r1c1_close};
          const refs_internal::R1C1Endpoint ep = refs_internal::parse_r1c1_endpoint_axes(name, 0, letters);
          if (ep.kind != refs_internal::R1C1EndpointKind::kReference || ep.end != name.size()) {
            return;
          }
          const parser::TextRange range = node.range();
          const std::size_t begin = self.body_at_ + utf16_to_byte_offset(self.body_, range.start);
          const std::size_t end = self.body_at_ + utf16_to_byte_offset(self.body_, range.end);
          std::string text;
          append_sheet_qualifier(text, name, {}, true);
          self.rewrites_.push_back(ReferenceRewrite{begin, end, std::move(text)});
        },
        this);
  }

  /// True when a reference holds a relative axis and there is no base.
  bool failed() const noexcept { return failed_begin_ != std::string_view::npos; }
  std::size_t failed_begin() const noexcept { return failed_begin_; }
  std::size_t failed_end() const noexcept { return failed_end_; }

  /// The rewrites in source order.
  std::vector<ReferenceRewrite> take_sorted() {
    const auto earlier = [this](std::uint32_t lhs, std::uint32_t rhs) {
      return rewrites_[lhs].begin < rewrites_[rhs].begin;
    };
    std::vector<std::uint32_t> order;
    sorted_index_order(order, static_cast<std::uint32_t>(rewrites_.size()), make_index_less(earlier));
    apply_index_order(rewrites_, order);
    return std::move(rewrites_);
  }

 private:
  void append_axis(char letter, std::uint32_t index, std::uint32_t base, bool absolute, std::string& out) {
    relative_ = relative_ || !absolute;
    parser::append_r1c1_axis(letter, facts_.r1c1_open, facts_.r1c1_close, index, base, absolute, out);
  }

  // A whole column or row prints its own axis only.
  void append_cell(const parser::Reference& r, std::string& out) {
    if (!r.is_full_col) {
      append_axis(facts_.r1c1_row, r.row, notation_.base_row, r.row_abs, out);
    }
    if (!r.is_full_row) {
      append_axis(facts_.r1c1_col, r.col, notation_.base_col, r.col_abs, out);
    }
  }

  // Rewrites the cell part of `node`'s text, after its last `!` and before a
  // spill `#`: `first`, or `first`, the operator as written and `last`.
  void add(const parser::AstNode& node, const parser::Reference& first, const parser::Reference* last, bool spill) {
    const parser::TextRange range = node.range();
    const std::size_t begin = body_at_ + utf16_to_byte_offset(body_, range.start);
    const std::size_t end = body_at_ + utf16_to_byte_offset(body_, range.end);
    if (end <= begin) {
      return;
    }
    std::string_view span = formula_.substr(begin, end - begin);
    if (const std::size_t hash = spill ? span.rfind('#') : std::string_view::npos; hash != std::string_view::npos) {
      span = span.substr(0, hash);
    }
    const std::size_t bang = span.rfind('!');
    const std::size_t tail_at = bang == std::string_view::npos ? 0U : bang + 1U;
    const std::string_view tail = span.substr(tail_at);
    std::size_t first_end = 0;
    while (first_end < tail.size() && is_endpoint_byte(tail[first_end])) {
      ++first_end;
    }
    std::size_t last_begin = tail.size();
    while (last_begin > first_end && is_endpoint_byte(tail[last_begin - 1U])) {
      --last_begin;
    }
    relative_ = false;
    ReferenceRewrite rewrite{begin + tail_at, begin + span.size(), {}};
    append_cell(first, rewrite.text);
    if (first_end < tail.size()) {
      // One whole column or row prints as one axis unless a trim operator joins its halves.
      const std::string_view op = tail.substr(first_end, last_begin - first_end);
      if (last != nullptr || op != ":") {
        rewrite.text.append(op);
        append_cell(last != nullptr ? *last : first, rewrite.text);
      }
    }
    if (relative_ && !notation_.has_base && begin < failed_begin_) {
      failed_begin_ = begin;
      failed_end_ = end;
    }
    rewrites_.push_back(std::move(rewrite));
  }

  std::string_view formula_;
  std::string_view body_;
  std::size_t body_at_;
  const LocaleFacts& facts_;
  FormulaNotation notation_;
  std::vector<ReferenceRewrite> rewrites_;
  bool relative_ = false;
  std::size_t failed_begin_ = std::string_view::npos;
  std::size_t failed_end_ = 0;
};

// `outer` maps a text onto an intermediate text, which `inner` maps onto the
// input; the result maps the text onto the input. A verbatim piece splits at
// the inner pieces it crosses; a rewritten one widens to every input byte
// its intermediate bytes came from.
std::vector<FormulaSourceSpan> compose_spans(const std::vector<FormulaSourceSpan>& outer,
                                             const std::vector<FormulaSourceSpan>& inner) {
  std::vector<FormulaSourceSpan> composed;
  composed.reserve(outer.size() + inner.size());
  std::size_t first = 0;
  for (const FormulaSourceSpan& o : outer) {
    while (first < inner.size() && inner[first].out_end < o.in_begin) {
      ++first;
    }
    if (o.verbatim) {
      for (std::size_t k = first; k < inner.size() && inner[k].out_begin < o.in_end; ++k) {
        const FormulaSourceSpan& in = inner[k];
        const std::size_t a = std::max(o.in_begin, in.out_begin);
        const std::size_t b = std::min(o.in_end, in.out_end);
        if (a >= b) {
          continue;
        }
        FormulaSourceSpan piece;
        piece.out_begin = o.out_begin + (a - o.in_begin);
        piece.out_end = o.out_begin + (b - o.in_begin);
        piece.in_begin = in.verbatim ? in.in_begin + (a - in.out_begin) : in.in_begin;
        piece.in_end = in.verbatim ? in.in_begin + (b - in.out_begin) : in.in_end;
        piece.verbatim = in.verbatim;
        composed.push_back(piece);
      }
      continue;
    }
    std::size_t lo = std::string_view::npos;
    std::size_t hi = 0;
    for (std::size_t k = first; k < inner.size() && inner[k].out_begin <= o.in_end; ++k) {
      const FormulaSourceSpan& in = inner[k];
      const bool overlaps = in.out_begin < o.in_end && o.in_begin < in.out_end;
      // An empty intermediate range maps to its point in the input.
      const bool holds_point = o.in_begin == o.in_end && in.out_begin <= o.in_begin && o.in_begin <= in.out_end;
      if (overlaps) {
        lo = std::min(lo, in.in_begin);
        hi = std::max(hi, in.in_end);
      } else if (holds_point && lo == std::string_view::npos) {
        lo = in.verbatim ? in.in_begin + (o.in_begin - in.out_begin) : in.in_begin;
        hi = lo;
      }
    }
    if (lo == std::string_view::npos) {
      lo = 0;
      hi = 0;
    }
    composed.push_back(FormulaSourceSpan{o.out_begin, o.out_end, lo, hi, false});
  }
  return composed;
}

RewriteRules a1_display_rules() {
  RewriteRules rules;
  rules.qualifiers = QualifierRule::kA1Display;
  return rules;
}

RewriteRules stored_rules() {
  RewriteRules rules;
  rules.direction = Direction::kToCanonical;
  return rules;
}

// A parse failure over a body, moved past the `prefix` (the `=`) before it.
parser::ParseFailure shift_past_prefix(parser::ParseFailure failure, std::string_view prefix) {
  const std::uint32_t units = utf16_units_in(prefix);
  failure.range.start += units;
  failure.range.end += units;
  if (failure.range.line == 1U) {
    failure.range.column += units;
  }
  return failure;
}

// Parses `formula` once and collects its rewrites in source order: the
// quoted spelling of every name that reads as a reference in `profile`'s R1C1
// letters and, with `notation.r1c1`, the R1C1 spelling of every reference.
// Fails with the parse failure, or over a relative reference with no base.
Expected<std::vector<ReferenceRewrite>, parser::ParseFailure> collect_rewrites(std::string_view formula,
                                                                               ExcelProfile profile,
                                                                               const FormulaNotation& notation) {
  if (formula.empty()) {
    return std::vector<ReferenceRewrite>{};
  }
  const std::string_view body = strip_formula_prefix(formula);
  const std::size_t body_at = formula.size() - body.size();
  Arena arena;
  parser::ParseFailure failure;
  const parser::AstNode* root = parse_formula_entry_with_diagnostics(body, arena, &failure);
  if (root == nullptr) {
    return shift_past_prefix(failure, formula.substr(0, body_at));
  }
  ReferenceCollector collector(formula, body_at, locale_facts(profile), notation);
  if (notation.r1c1) {
    collector.visit(*root);
    if (collector.failed()) {
      return parser::ParseFailure{parser::ParseErrorCode::InvalidReference,
                                  formula_text_range(formula, collector.failed_begin(), collector.failed_end()),
                                  parser::default_message(parser::ParseErrorCode::InvalidReference)};
    }
  }
  collector.quote_names(*root);
  return collector.take_sorted();
}

// The A1 display's rewrites: the quoted names, none when `formula` does not parse.
std::vector<ReferenceRewrite> quoted_names(std::string_view formula, ExcelProfile profile) {
  Expected<std::vector<ReferenceRewrite>, parser::ParseFailure> rewrites =
      collect_rewrites(formula, profile, FormulaNotation{});
  return rewrites ? std::move(rewrites.value()) : std::vector<ReferenceRewrite>{};
}

}  // namespace

parser::TextRange formula_text_range(std::string_view text, std::size_t begin, std::size_t end) {
  parser::TextRange range;
  range.start = utf16_units_in(text.substr(0, begin));
  range.end = utf16_units_in(text.substr(0, end));
  line_and_column(text, begin, range);
  return range;
}

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

void append_sheet_qualifier(std::string& out, std::string_view first, std::string_view last, bool quote) {
  if (quote) {
    out.push_back('\'');
  }
  const auto append = [&](std::string_view sheet) {
    if (quote) {
      parser::append_quote_escaped(sheet, out);
    } else {
      out.append(sheet);
    }
  };
  append(first);
  if (!last.empty()) {
    out.push_back(':');
    append(last);
  }
  if (quote) {
    out.push_back('\'');
  }
}

std::string localize_formula_text(std::string_view formula, ExcelProfile profile) {
  return rewrite_formula_text(formula, profile, a1_display_rules(), nullptr, quoted_names(formula, profile));
}

std::string canonicalize_formula_text(std::string_view formula, ExcelProfile profile) {
  return rewrite_formula_text(formula, profile, stored_rules(), nullptr, {});
}

MappedFormulaText localize_formula_text_mapped(std::string_view formula, ExcelProfile profile) {
  MappedFormulaText mapped;
  mapped.text =
      rewrite_formula_text(formula, profile, a1_display_rules(), &mapped.spans, quoted_names(formula, profile));
  return mapped;
}

MappedFormulaText canonicalize_formula_text_mapped(std::string_view formula, ExcelProfile profile) {
  MappedFormulaText mapped;
  mapped.text = rewrite_formula_text(formula, profile, stored_rules(), &mapped.spans, {});
  return mapped;
}

Expected<MappedFormulaText, parser::ParseFailure> localize_formula_text_r1c1(std::string_view formula,
                                                                             ExcelProfile profile,
                                                                             const FormulaNotation& notation) {
  const Expected<std::vector<ReferenceRewrite>, parser::ParseFailure> rewrites =
      collect_rewrites(formula, profile, notation);
  if (!rewrites) {
    return rewrites.error();
  }
  RewriteRules rules = a1_display_rules();
  if (notation.r1c1) {
    rules.qualifiers = QualifierRule::kR1C1Display;
    rules.bare_cell_names = true;
  }
  MappedFormulaText mapped;
  mapped.text = rewrite_formula_text(formula, profile, rules, &mapped.spans, rewrites.value());
  return mapped;
}

Expected<MappedFormulaText, parser::ParseFailure> canonicalize_formula_text_r1c1(std::string_view formula,
                                                                                 ExcelProfile profile,
                                                                                 const FormulaNotation& notation) {
  if (!notation.r1c1) {
    return canonicalize_formula_text_mapped(formula, profile);
  }
  const LocaleFacts& facts = locale_facts(profile);
  R1C1RewriteOptions options;
  options.letters = refs_internal::R1C1Letters{facts.r1c1_row, facts.r1c1_col, facts.r1c1_open, facts.r1c1_close};
  options.array_row_separator = facts.array_row_separator;
  options.has_base = notation.has_base;
  options.base_row = notation.base_row;
  options.base_col = notation.base_col;
  Expected<MappedFormulaText, parser::ParseFailure> a1 = rewrite_r1c1_text(formula, options);
  if (!a1) {
    return a1.error();
  }
  MappedFormulaText canonical = canonicalize_formula_text_mapped(a1.value().text, profile);
  canonical.spans = compose_spans(canonical.spans, a1.value().spans);
  if (canonical.text.empty()) {
    return canonical;
  }
  const std::string_view body = strip_formula_prefix(canonical.text);
  Arena arena;
  parser::ParseFailure failure;
  const parser::AstNode* root = parse_formula_entry_with_diagnostics(body, arena, &failure);
  if (root == nullptr) {
    const std::string_view prefix = std::string_view(canonical.text).substr(0, canonical.text.size() - body.size());
    failure = shift_past_prefix(failure, prefix);
    failure.range = canonical.source_range(formula, failure.range);
    return failure;
  }
  // A cell-shaped name reads back with `_xlpm.` when a LET / LAMBDA binds it, else `_xlnm.`.
  respell_unbound_cell_names(*root, canonical.text);
  return canonical;
}

}  // namespace eval
}  // namespace formulon
