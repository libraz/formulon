#include "eval/formula_entry.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/builtin_names.h"
#include "eval/formula_entry_r1c1.h"
#include "eval/formula_localize.h"
#include "excel_locale.h"
#include "parser/ast.h"
#include "parser/formula_prefix.h"
#include "parser/literal_spelling.h"
#include "parser/parse_error.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/arena.h"
#include "utils/formula_text.h"
#include "utils/index_sort.h"
#include "utils/strings.h"
#include "utils/utf8_length.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using parser::AstNode;
using parser::ParseErrorCode;
using parser::ParseFailure;

bool is_builtin_function_name(std::string_view name) {
  return resolve_builtin_function_name(name) != nullptr;
}

std::string with_formula_prefix(std::string_view text) {
  if (!text.empty() && text.front() == '=') {
    return std::string(text);
  }
  std::string out;
  out.reserve(text.size() + 1U);
  out += '=';
  out.append(text);
  return out;
}

// Where entered text is spelled: the workbook, its unknown-function
// spellings, the sheet a defined name is looked up from, or `kNoSheet`, and
// for a formula read from a file the locale it is opened under.
struct EntrySpellingContext {
  const Workbook* book;
  EnteredCallSpellings* calls;
  std::size_t sheet;
  std::optional<ExcelLocale> loaded_in = std::nullopt;
};

// The defined name `typed` reaches in scope `local` (a sheet index, or negative
// for workbook scope): its exact spelling, or with `builtin` the `_xlnm.` name.
const DefinedName* scoped_name(const Workbook& book, std::string_view typed, std::int32_t local, bool builtin) {
  for (const DefinedName& name : book.defined_names()) {
    if (local < 0 ? name.local_sheet_id >= 0 : name.local_sheet_id != local) {
      continue;
    }
    std::string_view stored = name.name;
    if (builtin) {
      if (!strings::case_insensitive_starts_with(stored, parser::kNamePrefix)) {
        continue;
      }
      stored.remove_prefix(parser::kNamePrefix.size());
    }
    if (strings::case_insensitive_eq(stored, typed)) {
      return &name;
    }
  }
  return nullptr;
}

// The workbook's own spelling of the sheet, defined name or table `typed` names.
parser::OwnName own_entry_name(const void* ctx, parser::EnteredNameKind kind, std::string_view typed,
                               std::string_view sheet) {
  const auto& entry = *static_cast<const EntrySpellingContext*>(ctx);
  const Workbook& book = *entry.book;
  switch (kind) {
    case parser::EnteredNameKind::kSheet: {
      const Sheet* found = book.sheet_by_name(typed);
      return {found != nullptr ? std::string_view(found->name()) : std::string_view()};
    }
    case parser::EnteredNameKind::kTable: {
      const std::size_t index = sheet_names::resolve_table(book.tables(), typed);
      return {index != sheet_names::kNotFound ? std::string_view(book.tables()[index].display_name)
                                              : std::string_view()};
    }
    case parser::EnteredNameKind::kDefinedName:
      break;
  }
  // A typed built-in name without its `_xlnm.` prefix reaches the visible
  // built-in, and a stored exact spelling wins; in a file, only where the
  // opening locale spells that built-in in English. Unqualified, a sheet-local
  // name hides a workbook one; a qualified name its sheet does not define is
  // read at workbook scope. A built-in that resolves to nothing still takes
  // its canonical spelling.
  const bool reaches_builtin =
      !entry.loaded_in.has_value() || !bare_builtin_defined_name(*entry.loaded_in, typed).empty();
  if (!sheet.empty()) {
    const std::size_t qualifier = book.sheet_index_by_name(sheet);
    if (qualifier != Workbook::kNoSheet) {
      for (const bool builtin : {false, reaches_builtin}) {
        if (const DefinedName* name = scoped_name(book, typed, static_cast<std::int32_t>(qualifier), builtin)) {
          return {name->name};
        }
      }
      for (const bool builtin : {false, reaches_builtin}) {
        if (const DefinedName* name = scoped_name(book, typed, -1, builtin)) {
          return {name->name, true};
        }
      }
    }
    return {builtin_defined_name(typed)};
  }
  const auto local = entry.sheet != Workbook::kNoSheet ? static_cast<std::int32_t>(entry.sheet) : -1;
  for (const bool builtin : {false, reaches_builtin}) {
    for (const std::int32_t scope : {local, -1}) {
      if (const DefinedName* name = scoped_name(book, typed, scope, builtin)) {
        return {name->name};
      }
    }
  }
  return {builtin_defined_name(typed)};
}

parser::OwnName no_entry_name(const void*, parser::EnteredNameKind, std::string_view, std::string_view) {
  return {};
}

std::string_view entry_unknown_call(const void* ctx, std::string_view typed) {
  const auto& entry = *static_cast<const EntrySpellingContext*>(ctx);
  return entry.calls->spelling(*entry.book, typed);
}

std::string_view kept_unknown_call(const void*, std::string_view typed) {
  return typed;
}

// Tab position of `sheet` in this workbook (empty `book`) or in the cached
// sheet list of the external link `path` / `book` names.
std::size_t entry_sheet_position(const void* ctx, std::string_view path, std::string_view book,
                                 std::string_view sheet) {
  const Workbook& wb = *static_cast<const Workbook*>(ctx);
  if (book.empty()) {
    const std::size_t index = wb.sheet_index_by_name(sheet);
    return index != Workbook::kNoSheet ? index : parser::kUnknownSheet;
  }
  const ExternalLinkRecord* link = wb.find_external_link(path, book);
  const std::uint32_t index = link != nullptr ? link->book.sheet_index(sheet) : ExternalBook::kNoSheet;
  return index != ExternalBook::kNoSheet ? index : parser::kUnknownSheet;
}

// Entered formula `text` with the name and reference spelling Excel stores,
// defined names looked up from `sheet` (`Workbook::kNoSheet` for none), and
// each 3-D span in tab order.
std::string spell_entered_names_in(const Workbook& book, EnteredCallSpellings& calls, std::size_t sheet,
                                   std::string_view text) {
  const EntrySpellingContext ctx{&book, &calls, sheet};
  const parser::EnteredNames names{&resolve_builtin_function_name, &own_entry_name, &entry_unknown_call, &ctx};
  return parser::order_sheet_spans(parser::spell_entered_names(text, names),
                                   parser::SheetOrder{&entry_sheet_position, &book});
}

// `code` over the whole of `text`.
ParseFailure whole_text_failure(ParseErrorCode code, std::string_view text) {
  ParseFailure failure;
  failure.code = code;
  failure.range.end = utf16_units_in(text);
  failure.message = parser::default_message(code);
  return failure;
}

// `failure` moved right by `units` UTF-16 code units of first-line text.
ParseFailure shifted(ParseFailure failure, std::uint32_t units) {
  failure.range.start += units;
  failure.range.end += units;
  if (failure.range.line == 1U) {
    failure.range.column += units;
  }
  return failure;
}

// UTF-16 code units of the text ahead of `body`, a suffix of `text`.
std::uint32_t prefix_units(std::string_view text, std::string_view body) {
  return utf16_units_in(text.substr(0, text.size() - body.size()));
}

// Bytes spanned by the first `units` UTF-16 code units of UTF-8 `text`.
std::size_t utf8_prefix_bytes(std::string_view text, std::uint32_t units) {
  std::size_t i = 0;
  std::uint32_t seen = 0;
  while (i < text.size() && seen < units) {
    const auto lead = static_cast<unsigned char>(text[i]);
    const std::size_t len = lead < 0x80U ? 1U : lead >= 0xF0U ? 4U : lead >= 0xE0U ? 3U : lead >= 0xC0U ? 2U : 1U;
    seen += len == 4U ? 2U : 1U;
    i += len;
  }
  return std::min(i, text.size());
}

// The call in `node` that takes `arg` as one of its arguments.
const AstNode* call_taking(const AstNode& node, const AstNode* arg) {
  if (node.kind() == parser::NodeKind::Call) {
    for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
      if (&node.as_call_arg(i) == arg) {
        return &node;
      }
    }
  }
  const AstNode* hit = nullptr;
  parser::any_child_node(node, [&](const AstNode& child) {
    hit = call_taking(child, arg);
    return hit != nullptr;
  });
  return hit;
}

// AST of formula body `src` (no leading '=') as entry accepts it, or nullptr
// with `*failure` set in `src`'s coordinates: the parse diagnostic, a
// qualified built-in, and with `refusals` a non-reference argument to a
// reference-only parameter (`SUBTOTAL(9,)`), an argument count the callee
// does not take (`ROWS()`), a newer error literal (`#CALC!`), or a typed
// `ANCHORARRAY` call.
// The grammar entry parses in: a cell formula's, or a defined-name body's.
using EntryParse = AstNode* (*)(std::string_view, Arena&, ParseFailure*);

const AstNode* diagnose_entry(std::string_view src, Arena& arena, bool refusals, ParseFailure* failure,
                              EntryParse parse = &parse_formula_entry_with_diagnostics) {
  const AstNode* root = parse(src, arena, failure);
  if (root == nullptr || !refusals) {
    return root;
  }
  ParseErrorCode code = ParseErrorCode::EntryReferenceRequired;
  const AstNode* call = nullptr;
  if (const AstNode* arg = find_non_reference_argument(*root); arg != nullptr) {
    call = call_taking(*root, arg);
  } else if ((call = find_wrong_arity_call(*root)) != nullptr) {
    code = ParseErrorCode::EntryWrongArity;
  } else if ((call = find_refused_error_literal(*root)) != nullptr) {
    code = ParseErrorCode::LexerInvalidErrorLiteral;
  } else if ((call = find_anchor_array_call(*root)) != nullptr) {
    code = ParseErrorCode::EntryAnchorArrayCall;
  } else if ((call = find_overlong_external_sheet(*root)) != nullptr) {
    code = ParseErrorCode::InvalidReference;
  } else {
    return root;
  }
  failure->code = code;
  failure->range = call != nullptr ? call->range() : whole_text_failure(code, src).range;
  failure->message = parser::default_message(code);
  return nullptr;
}

// `failure`, from `diagnose_entry` on the body of `parsed` (the caller's
// `caller` after normalization), placed in `caller`: as is when `parsed` and
// `caller` agree up to its end, else as a re-diagnosis of `caller` finds it
// when that yields the same code, else over all of `caller`.
ParseFailure place_failure(std::string_view caller, std::string_view parsed, bool refusals, const ParseFailure& failure,
                           EntryParse parse = &parse_formula_entry_with_diagnostics) {
  const ParseFailure at = shifted(failure, prefix_units(parsed, strip_formula_prefix(parsed)));
  const std::size_t bytes = utf8_prefix_bytes(parsed, at.range.end);
  if (caller.substr(0, bytes) == parsed.substr(0, bytes)) {
    return at;
  }
  const std::string_view body = strip_formula_prefix(caller);
  Arena arena;
  ParseFailure again;
  if (diagnose_entry(body, arena, refusals, &again, parse) == nullptr && again.code == failure.code) {
    return shifted(again, prefix_units(caller, body));
  }
  return whole_text_failure(failure.code, caller);
}

// The refusal of a leading byte-order mark or a second leading '=', which
// precede any parse.
std::optional<ParseFailure> refused_entry_prefix(std::string_view formula) {
  ParseFailure failure;
  if (formula.substr(0, 3) == "\xEF\xBB\xBF") {
    failure.code = ParseErrorCode::EntryBomNotSupported;
    failure.range.end = 1;
  } else if (formula.substr(0, 2) == "==") {
    failure.code = ParseErrorCode::EntryDoubleEquals;
    failure.range = parser::TextRange{1U, 2U, 1U, 2U};
  } else {
    return std::nullopt;
  }
  failure.message = parser::default_message(failure.code);
  return failure;
}

// The refusal of typed `formula` text, a table reference whose table name opens with an R1C1
// reference spelled in the profile locale of `book` and that names a table of `book`: Excel reads
// the typed `F[` as a row reference. Run on the text as typed, before structured references are
// bound, so a bare `[@Col]` has no such token and text naming no table falls through to R1C1.
std::optional<ParseFailure> refused_typed_table_reference(std::string_view formula, const Workbook& book) {
  const LocaleFacts& locale = locale_facts(book.excel_profile());
  const std::string_view body = strip_formula_prefix(formula);
  Arena arena;
  ParseFailure ignored;
  const AstNode* root = parse_formula_entry_with_diagnostics(body, arena, &ignored);
  const AstNode* ref = root != nullptr ? find_refused_table_reference(
                                             *root, locale.r1c1_row, locale.r1c1_col,
                                             [](const void* ctx, std::string_view name) {
                                               return static_cast<const Workbook*>(ctx)->is_table_name(name);
                                             },
                                             &book)
                                       : nullptr;
  if (ref == nullptr) {
    return std::nullopt;
  }
  ParseFailure failure;
  failure.code = ParseErrorCode::InvalidReference;
  failure.range = ref->range();
  failure.range.end = failure.range.start + utf16_units_in(ref->as_structured_ref_table());
  failure.message = parser::default_message(failure.code);
  return shifted(failure, prefix_units(formula, body));
}

// The refusal of typed `formula` text naming a chartsheet of `book` as the
// sheet of a cell reference: Excel stores no formula that reads one.
std::optional<ParseFailure> refused_chartsheet_reference(std::string_view formula, const Workbook& book) {
  const std::string_view body = strip_formula_prefix(formula);
  Arena arena;
  ParseFailure ignored;
  const AstNode* root = parse_formula_entry_with_diagnostics(body, arena, &ignored);
  const AstNode* ref = root != nullptr
                           ? find_reference_to_sheet(
                                 *root,
                                 [](const void* ctx, std::string_view sheet) {
                                   const Sheet* found = static_cast<const Workbook*>(ctx)->sheet_by_name(sheet);
                                   return found != nullptr && found->kind() == SheetKind::kChartsheet;
                                 },
                                 &book)
                           : nullptr;
  if (ref == nullptr) {
    return std::nullopt;
  }
  ParseFailure failure;
  failure.code = ParseErrorCode::InvalidReference;
  failure.range = ref->range();
  failure.message = parser::default_message(failure.code);
  return shifted(failure, prefix_units(formula, body));
}

// The refusal of an `ANCHORARRAY` call in typed `formula`, in the coordinates
// of its body. Run on the text as typed: normalization respells the call as
// the `#` operator Excel accepts.
std::optional<ParseFailure> refused_anchor_array_call(std::string_view formula) {
  if (!strings::case_insensitive_contains(formula, parser::kAnchorArrayFn)) {
    return std::nullopt;
  }
  Arena arena;
  ParseFailure ignored;
  const AstNode* root = parse_formula_entry_with_diagnostics(strip_formula_prefix(formula), arena, &ignored);
  const AstNode* call = root != nullptr ? find_anchor_array_call(*root) : nullptr;
  if (call == nullptr) {
    return std::nullopt;
  }
  ParseFailure failure;
  failure.code = ParseErrorCode::EntryAnchorArrayCall;
  failure.range = call->range();
  failure.message = parser::default_message(failure.code);
  return failure;
}

// AST of entered `formula`'s stored spelling, written to `*stored`, or
// nullptr with `*failure` placed in `caller` (the text the caller passed,
// which `formula` binds): whitespace Excel refuses, or a stored spelling
// `diagnose_entry` refuses, or an `ANCHORARRAY` call as typed.
const AstNode* parse_entered_formula(std::string_view caller, std::string_view formula, Arena& arena,
                                     std::string* stored, ParseFailure* failure,
                                     EntryParse parse = &parse_formula_entry_with_diagnostics) {
  *stored = Workbook::normalize_formula_text(std::string(formula), Workbook::FormulaSource::kEntered);
  if (parser::has_refused_whitespace(formula, &is_builtin_function_name)) {
    *failure = whole_text_failure(ParseErrorCode::EntryRefusedWhitespace, caller);
    return nullptr;
  }
  if (const std::optional<ParseFailure> typed = refused_anchor_array_call(formula); typed.has_value()) {
    *failure = place_failure(caller, formula, /*refusals=*/true, *typed, parse);
    return nullptr;
  }
  ParseFailure raw;
  const AstNode* root = diagnose_entry(strip_formula_prefix(*stored), arena, /*refusals=*/true, &raw, parse);
  if (root == nullptr) {
    *failure = place_failure(caller, *stored, /*refusals=*/true, raw, parse);
  }
  return root;
}

// Entry of `formula` stored outside a cell, parsed by `parse`.
EntryResult enter_outside_cell(const Workbook& book, EnteredCallSpellings& calls, std::string_view formula,
                               std::string* stored, ParseFailure* failure, EntryParse parse) {
  // The parser skips one '=' of its own, so a second one is refused here.
  if (const std::optional<ParseFailure> refused = refused_entry_prefix(formula)) {
    *failure = *refused;
    return EntryResult::kRefused;
  }
  if (const std::optional<ParseFailure> refused = refused_typed_table_reference(formula, book)) {
    *failure = *refused;
    return EntryResult::kRefused;
  }
  if (const std::optional<ParseFailure> refused = refused_chartsheet_reference(formula, book)) {
    *failure = *refused;
    return EntryResult::kRefused;
  }
  if (!sheet_names::valid_utf8(formula)) {
    return EntryResult::kInvalidUtf8;
  }
  if (has_entered_cell_shaped_param(formula)) {
    *failure = whole_text_failure(ParseErrorCode::EntryParamShapedName, formula);
    return EntryResult::kRefused;
  }
  std::string normalized;
  Arena arena;
  if (parse_entered_formula(formula, formula, arena, &normalized, failure, parse) == nullptr) {
    return EntryResult::kRefused;
  }
  *stored = spell_entered_names_in(book, calls, Workbook::kNoSheet, normalized);
  return EntryResult::kStored;
}

}  // namespace

std::string_view EnteredCallSpellings::record(const void* self, std::string_view typed) {
  auto& spellings = static_cast<const EnteredCallSpellings*>(self)->spellings_;
  std::string key = strings::to_ascii_lower(typed);
  for (const auto& [seen_key, spelling] : spellings) {
    if (seen_key == key) {
      return spelling;
    }
  }
  spellings.emplace_back(std::move(key), std::string(typed));
  return spellings.back().second;
}

std::string_view EnteredCallSpellings::spelling(const Workbook& book, std::string_view typed) {
  if (!seeded_) {
    seeded_ = true;
    const parser::EnteredNames names{&resolve_builtin_function_name, &no_entry_name, &record, this};
    for (const DefinedName& name : book.defined_names()) {
      (void)parser::spell_entered_names(name.formula, names);
    }
    for (std::size_t s = 0; s < book.sheet_count(); ++s) {
      const Sheet& sheet = book.sheet(s);
      std::vector<CellAddress> cells = sheet.all_formula_cells();
      sort_by_index(cells, [](const CellAddress& a, const CellAddress& b) {
        return a.row != b.row ? a.row < b.row : a.col < b.col;
      });
      for (const CellAddress& address : cells) {
        Sheet::CellRead read;
        sheet.read_formula_cell(address.row, address.col, read);
        (void)parser::spell_entered_names(read.formula_text(), names);
      }
    }
  }
  return record(this, typed);
}

EntryResult enter_cell_formula(const Workbook& book, EnteredCallSpellings& calls, std::size_t sheet, std::uint32_t row,
                               std::uint32_t col, std::string_view formula, std::string* stored,
                               ParseFailure* failure) {
  if (!sheet_names::valid_utf8(formula)) {
    return EntryResult::kInvalidUtf8;
  }
  if (const std::optional<ParseFailure> refused = refused_entry_prefix(formula)) {
    *failure = *refused;
    return EntryResult::kRefused;
  }
  if (const std::optional<ParseFailure> refused = refused_typed_table_reference(formula, book)) {
    *failure = *refused;
    return EntryResult::kRefused;
  }
  if (const std::optional<ParseFailure> refused = refused_chartsheet_reference(formula, book)) {
    *failure = *refused;
    return EntryResult::kRefused;
  }
  const std::string bound = book.bind_entered_structured_refs(formula, sheet, row, col);
  if (has_entered_cell_shaped_param(bound)) {
    *failure = whole_text_failure(ParseErrorCode::EntryParamShapedName, formula);
    return EntryResult::kRefused;
  }
  std::string normalized;
  ParseFailure a1_failure;
  bool a1_failed = false;
  {
    Arena arena;
    const AstNode* root = parse_entered_formula(formula, bound, arena, &normalized, &a1_failure);
    if (root != nullptr && !needs_r1c1_entry(*root, bound, book)) {
      *stored = with_formula_prefix(spell_entered_names_in(book, calls, sheet, bound));
      return EntryResult::kStored;
    }
    a1_failed = root == nullptr;
  }
  // Excel reads text that fails as A1 again as R1C1 relative to the cell.
  std::optional<std::string> translated = translate_r1c1_entry(bound, row, col);
  if (translated) {
    Arena arena;
    ParseFailure r1c1_failure;
    if (parse_entered_formula(*translated, *translated, arena, &normalized, &r1c1_failure) != nullptr) {
      *stored = with_formula_prefix(spell_entered_names_in(book, calls, sheet, *translated));
      return EntryResult::kStored;
    }
  }
  // A failed A1 reading is the one the caller's text locates; text that read
  // as A1 failed only as the R1C1 reference it names.
  *failure = a1_failed ? a1_failure : whole_text_failure(ParseErrorCode::InvalidReference, formula);
  return EntryResult::kRefused;
}

EntryResult enter_stored_formula(const Workbook& book, EnteredCallSpellings& calls, std::string_view formula,
                                 std::string* stored, ParseFailure* failure) {
  return enter_outside_cell(book, calls, formula, stored, failure, &parse_formula_entry_with_diagnostics);
}

EntryResult enter_name_body(const Workbook& book, EnteredCallSpellings& calls, std::string_view formula,
                            std::string* stored, ParseFailure* failure) {
  return enter_outside_cell(book, calls, formula, stored, failure, &parse_name_body_with_diagnostics);
}

std::string spell_loaded_cell_formula(const Workbook& book, std::size_t sheet, ExcelLocale locale,
                                      std::string_view formula) {
  const EntrySpellingContext ctx{&book, nullptr, sheet, locale};
  const parser::EnteredNames names{&resolve_builtin_function_name, &own_entry_name, &kept_unknown_call, &ctx};
  return parser::spell_entered_names(formula, names);
}

const AstNode* parse_evaluated_formula(std::string_view formula, std::string_view bound, Arena& arena,
                                       ParseFailure* failure) {
  ParseFailure raw;
  const AstNode* root = diagnose_entry(strip_formula_prefix(bound), arena, /*refusals=*/false, &raw);
  if (root == nullptr) {
    *failure = place_failure(formula, bound, /*refusals=*/false, raw);
  }
  return root;
}

}  // namespace eval
}  // namespace formulon
