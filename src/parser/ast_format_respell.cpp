//
// Respelling of a formula's storage text over its typed source, and the
// sheet-quoting check that decides whether storage text needs rewriting.

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_format_internal.h"
#include "parser/formula_prefix.h"
#include "parser/reference.h"
#include "parser/token.h"
#include "parser/tokenizer.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {

namespace {

// True when `stored` is `written` behind one or more storage prefixes; the
// name itself may differ in case (`_xleta.` upper-cases it).
bool IsPrefixedSpelling(std::string_view written, std::string_view stored) {
  static constexpr std::array<std::string_view, 5> kPrefixes = {kFnPrefix, kWsPrefix, kParamPrefix, kNamePrefix,
                                                                kEtaPrefix};
  if (stored.size() <= written.size() ||
      !strings::case_insensitive_eq(stored.substr(stored.size() - written.size()), written)) {
    return false;
  }
  std::string_view head = stored.substr(0, stored.size() - written.size());
  while (!head.empty()) {
    const auto it = std::find_if(kPrefixes.begin(), kPrefixes.end(),
                                 [head](std::string_view prefix) { return strings::starts_with(head, prefix); });
    if (it == kPrefixes.end()) {
      return false;
    }
    head.remove_prefix(it->size());
  }
  return true;
}

}  // namespace

std::string respell_storage_over_source(std::string_view source, std::string_view storage) {
  TokenizerOptions options;
  options.max_formula_length_utf16 = std::numeric_limits<std::uint32_t>::max();
  Tokenizer source_lexer(source, options);
  Tokenizer storage_lexer(storage, options);
  const std::vector<Token>& written = source_lexer.tokens();
  const std::vector<Token>& stored = storage_lexer.tokens();
  if (!source_lexer.errors().empty() || !storage_lexer.errors().empty()) {
    return std::string(storage);
  }
  std::string out;
  out.reserve(source.size() + storage.size());
  std::size_t j = 0;
  bool stored_space = false;
  const auto skip_stored_space = [&] {
    stored_space = false;
    while (stored[j].kind == TokenKind::Whitespace) {
      stored_space = true;
      ++j;
    }
  };
  skip_stored_space();
  bool written_space = false;
  // A written `[b]` against a stored `_xlop.b`: 1 after `[`, 2 after the name (already written), 0 outside.
  int in_optional = 0;
  for (const Token& token : written) {
    if (token.kind == TokenKind::Whitespace) {
      out.append(token.lexeme);
      written_space = true;
      continue;
    }
    if (token.kind == TokenKind::Eof) {
      break;
    }
    const Token& counterpart = stored[j];
    // A stored space is the intersection operator; the source must have one there.
    if (stored_space && !written_space) {
      return std::string(storage);
    }
    if (token.kind == TokenKind::At && counterpart.kind != TokenKind::At) {
      continue;
    }
    if (token.kind == TokenKind::LBracket && in_optional == 0 && counterpart.kind == TokenKind::Ident &&
        strings::case_insensitive_starts_with(counterpart.lexeme, kOptionalParamPrefix)) {
      in_optional = 1;
      continue;
    }
    if (in_optional == 1) {
      if (token.kind != TokenKind::Ident ||
          !strings::case_insensitive_eq(cell_shaped_name_bare(token.lexeme),
                                        counterpart.lexeme.substr(kOptionalParamPrefix.size()))) {
        return std::string(storage);
      }
      out.append(counterpart.lexeme);
      in_optional = 2;
      continue;
    }
    if (in_optional == 2) {
      if (token.kind != TokenKind::RBracket) {
        return std::string(storage);
      }
      in_optional = 0;
      written_space = false;
      ++j;
      skip_stored_space();
      continue;
    }
    if (token.kind == counterpart.kind && token.lexeme == counterpart.lexeme) {
      out.append(token.lexeme);
    } else if (counterpart.kind == TokenKind::Ident && IsPrefixedSpelling(token.lexeme, counterpart.lexeme)) {
      out.append(counterpart.lexeme);
    } else {
      return std::string(storage);
    }
    written_space = false;
    ++j;
    skip_stored_space();
  }
  return stored[j].kind == TokenKind::Eof ? out : std::string(storage);
}

bool formula_needs_storage_requote(const AstNode& root) {
  const auto bare_but_needs_quotes = [](std::string_view sheet, bool quoted) {
    return !sheet.empty() && !quoted && local_sheet_needs_quoting_a1(sheet);
  };
  std::vector<const AstNode*> pending{&root};
  while (!pending.empty()) {
    const AstNode& node = *pending.back();
    pending.pop_back();
    switch (node.kind()) {
      case NodeKind::Ref:
        if (bare_but_needs_quotes(node.as_ref().sheet, node.as_ref().sheet_quoted)) {
          return true;
        }
        break;
      case NodeKind::SpillRef:
        if (node.as_spill_ref_anchor_expr() == nullptr &&
            bare_but_needs_quotes(node.as_spill_ref().sheet, node.as_spill_ref().sheet_quoted)) {
          return true;
        }
        break;
      case NodeKind::NameRef:
        if (bare_but_needs_quotes(node.as_name_sheet(), node.as_name_sheet_quoted())) {
          return true;
        }
        break;
      case NodeKind::ErrorLiteral:
        if (const RefErrorSite* site = node.as_ref_error_site();
            site != nullptr && site->form == RefErrorSite::Form::kQualified &&
            bare_but_needs_quotes(site->sheet, site->sheet_quoted)) {
          return true;
        }
        break;
      case NodeKind::Ref3D:
        // The span keeps no quoting hint; requoting text that was already
        // quoted rewrites it to the same spelling.
        if (local_sheet_needs_quoting_a1(node.as_ref3d_sheet_begin()) ||
            local_sheet_needs_quoting_a1(node.as_ref3d_sheet_end())) {
          return true;
        }
        break;
      default:
        break;
    }
    if (node.kind() != NodeKind::ArrayLiteral) {
      any_child_node(node, [&](const AstNode& child) {
        pending.push_back(&child);
        return false;
      });
    }
  }
  return false;
}

}  // namespace parser
}  // namespace formulon
