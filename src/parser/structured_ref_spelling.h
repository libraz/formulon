//
// Converts the structured references of formula text between the stored
// spelling and the formula-bar spelling of a given cell position.
//
// Both directions parse the text and replace node spans in place, so every
// byte outside a rewritten reference is kept. Text that does not parse is
// returned unchanged.

#ifndef FORMULON_PARSER_STRUCTURED_REF_SPELLING_H_
#define FORMULON_PARSER_STRUCTURED_REF_SPELLING_H_

#include <optional>
#include <string>
#include <string_view>

#include "parser/ast.h"
#include "sheet_name.h"

namespace formulon {
namespace parser {

/// Whether an unqualified name entered in a formula names a table (and no
/// defined name visible where the formula is entered hides it). A null
/// `is_bare_table` matches nothing.
struct BareTableNamePredicate {
  bool (*is_bare_table)(const void* ctx, std::string_view name);
  const void* ctx;
};

/// `text` (a leading `=` optional) with every structured reference in the
/// display form for a formula in `own_table`'s range; `own_table` is empty
/// outside every table. A table spelling resolves among `table_list` (the
/// case fold against `own_table` without one).
/// Display form of one `StructuredRef` node, or `std::nullopt` when its payload does not parse.
/// `own_table` / `table_list` decide whether the reference sits inside its own table.
std::optional<std::string> display_structured_ref_node(const AstNode& node, std::string_view own_table,
                                                       const sheet_names::NameList* table_list);

std::string display_structured_refs(std::string_view text, std::string_view own_table,
                                    const sheet_names::NameList* table_list = nullptr);

/// The stored form of entered `text` at a cell in `own_table`'s range (empty
/// outside every table): an unqualified `[...]` gains `own_table` and the
/// stored payload, `T[@...]` takes the stored payload, and an unqualified
/// name the predicate accepts, unless bound by LET or LAMBDA, gains `[]`.
/// Other references are kept as entered.
std::string bind_structured_refs(std::string_view text, std::string_view own_table,
                                 BareTableNamePredicate is_table_name);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_STRUCTURED_REF_SPELLING_H_
