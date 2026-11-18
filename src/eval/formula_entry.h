// Formula entry: the text Excel stores for a formula typed into a cell, or
// into a conditional-format, data-validation or defined-name field. Entry
// refuses what Excel refuses (a second leading '=', refused whitespace, a
// reference-only parameter given a value, a wrong argument count, a typed
// table name read as an R1C1 reference), binds structured references at the
// cell, falls back to R1C1 at the cell, normalises number literals and
// whitespace, and respells names and references as Excel stores them. Every
// binding enters through the C API, and the oracle harness through the same
// functions, so both store what Excel stores.

#ifndef FORMULON_EVAL_FORMULA_ENTRY_H_
#define FORMULON_EVAL_FORMULA_ENTRY_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <utility>

#include "excel_profile.h"
#include "parser/parse_status.h"

namespace formulon {

class Arena;
class Workbook;

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

/// The first spelling of each unknown function name a workbook has seen, by
/// ASCII lower-case key: an entered call takes it, as Excel decides that
/// spelling workbook-wide. Filled on first use from the workbook's stored
/// formulas in Excel's load order (defined names, then each sheet in tab
/// order, row by row), then from each entry.
class EnteredCallSpellings {
 public:
  /// The spelling an entered call to unknown `typed` takes in `book`.
  std::string_view spelling(const Workbook& book, std::string_view typed);

 private:
  static std::string_view record(const void* self, std::string_view typed);

  // Written through the `const void*` the spelling callbacks take.
  mutable std::deque<std::pair<std::string, std::string>> spellings_;
  bool seeded_ = false;
};

enum class EntryResult : std::uint8_t { kStored, kRefused, kInvalidUtf8 };

/// Entry of `formula` at (`sheet`, `row`, `col`) of `book`. On `kStored`,
/// `*stored` is the text to store, with one leading '='; on `kRefused`,
/// `*failure` locates the refusal in `formula`.
EntryResult enter_cell_formula(const Workbook& book, EnteredCallSpellings& calls, std::size_t sheet, std::uint32_t row,
                               std::uint32_t col, std::string_view formula, std::string* stored,
                               parser::ParseFailure* failure);

/// Entry of non-empty `formula` stored outside a cell (no structured-reference
/// binding or R1C1 fallback; defined names looked up workbook-wide). On
/// `kStored`, `*stored` is the text to store, with one leading '='.
EntryResult enter_stored_formula(const Workbook& book, EnteredCallSpellings& calls, std::string_view formula,
                                 std::string* stored, parser::ParseFailure* failure);

/// `enter_stored_formula` for a defined-name body, which may also be a bare
/// comma list of references (`Sheet1!$A$1,Sheet1!$B$1`).
EntryResult enter_name_body(const Workbook& book, EnteredCallSpellings& calls, std::string_view formula,
                            std::string* stored, parser::ParseFailure* failure);

/// `formula`, a cell formula read from a file onto `sheet` of `book` opened
/// under `locale`, with the name and reference spelling entry stores; an
/// unknown function keeps its spelling. Nothing is validated.
std::string spell_loaded_cell_formula(const Workbook& book, std::size_t sheet, ExcelLocale locale,
                                      std::string_view formula);

/// AST of ad-hoc `formula`, bound to `bound`, owned by `arena`; nullptr with
/// `*failure` located in `formula` when it does not parse.
const parser::AstNode* parse_evaluated_formula(std::string_view formula, std::string_view bound, Arena& arena,
                                               parser::ParseFailure* failure);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_FORMULA_ENTRY_H_
