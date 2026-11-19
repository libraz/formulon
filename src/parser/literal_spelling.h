//
// Respells the number literals of formula text the way Excel stores a formula
// it has parsed from text.
//
// Each literal is printed from its value (`format_formula_number`), so
// `1.50`, `1E3` and `0.30000000000000004` read back as `1.5`, `1000` and
// `0.3`. A unary `+` or `-` written directly against the literal, with no
// whitespace in between, belongs to the literal: `+` disappears and `-`
// makes it negative, so `-0` prints as `0`. Only that one sign is taken;
// any sign before it, a binary operator, and a sign over a reference, a
// parenthesis, a call, an array or any non-number literal stay as written.
//
// Entered text also loses the whitespace its parsed form has no place for
// (`drop_unkept_whitespace`); `spell_entered_formula` applies both steps.
// `spell_entered_names` then gives names and references the case Excel
// stores, which needs the workbook's built-in and sheet names.
//
// The rewrite is not idempotent where Excel's is not either: a surviving
// sign touching a literal (`=+5` stored for `=++5`) is absorbed when the
// stored text is parsed again.

#ifndef FORMULON_PARSER_LITERAL_SPELLING_H_
#define FORMULON_PARSER_LITERAL_SPELLING_H_

#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {
namespace parser {

/// Returns invariant-spelled `formula` (a leading `=` optional) with every
/// number literal respelled and its touching unary sign absorbed as the
/// header describes. Every other byte, whitespace and name case included,
/// is copied verbatim. Literals inside structured-reference brackets, row
/// numbers of a `1:3` row range and runs the tokenizer rejects (`1E-310`)
/// are left alone; nothing is validated.
std::string spell_number_literals(std::string_view formula);

/// Returns `formula` (a leading `=` optional) without the whitespace Excel
/// drops when it parses entered text: the trailing run, a run before a `,`,
/// a run on either side of a `:`, every run inside an array constant and,
/// in any other run, the spaces directly before a line feed. Every other
/// run is kept, the intersection operator, a leading run and line feeds
/// among them. Text inside structured-reference brackets and string
/// literals is left alone; nothing is validated.
std::string drop_unkept_whitespace(std::string_view formula);

/// True when entered `formula` (a leading `=` optional) holds whitespace
/// Excel refuses to store: a tab or a carriage return between tokens, or a
/// run between a function name (`is_function_name`) and its `(`. Text inside
/// structured-reference brackets and string literals is not looked at.
bool has_refused_whitespace(std::string_view formula, bool (*is_function_name)(std::string_view));

/// Stored spelling of entered formula text: `drop_unkept_whitespace`, then
/// `spell_number_literals`.
std::string spell_entered_formula(std::string_view formula);

/// What a name in entered text may denote, for `EnteredNames::own`.
enum class EnteredNameKind : std::uint8_t { kSheet, kDefinedName, kTable };

/// How `EnteredNames::own` reads a typed name.
struct OwnName {
  /// The spelling to store, or empty to keep the typed text.
  std::string_view spelling;
  /// A qualified defined name its sheet does not define, reached at workbook
  /// scope and stored as `[0]!Name`.
  bool workbook_scope = false;
};

/// The spelling of the external link an unknown sheet qualifier binds, for
/// `EnteredNames::link`.
struct LinkSpelling {
  /// The link's book, or empty when no existing link matches.
  std::string book;
  /// The link's sheet the qualifier names.
  std::string sheet;
};

/// What `spell_entered_names` reads from outside the formula.
struct EnteredNames {
  /// Canonical spelling of the built-in `name`, or nullptr when it is none.
  const char* (*builtin)(std::string_view name);
  /// The workbook's own spelling of the sheet, defined name or table `typed`
  /// resolves to. `sheet` is the unquoted sheet a defined name is qualified
  /// with (`Sheet1!Name`), empty otherwise.
  OwnName (*own)(const void* ctx, EnteredNameKind kind, std::string_view typed, std::string_view sheet);
  /// The spelling of a call of the unknown function `typed`: the first one
  /// the workbook has seen, `typed` itself when it is the first (and is
  /// remembered as such).
  std::string_view (*unknown_call)(const void* ctx, std::string_view typed);
  const void* ctx;
  /// The book and sheet spelling of the existing extension-less external
  /// link the sheet qualifier `typed`, which names no local sheet, folds to;
  /// an empty book binds a new link spelled as typed. Null leaves every
  /// unknown qualifier as typed.
  LinkSpelling (*link)(const void* ctx, std::string_view typed) = nullptr;
};

/// Where `bind_unknown_qualifiers` runs, which decides the `:` operands it binds.
enum class QualifierBindingRoute : std::uint8_t {
  /// Entered text: an unknown qualifier that is an operand of `:` binds only in `X:Y!ref` with `X` an
  /// identifier.
  kEntry,
  /// Text read from a file: an unknown qualifier binds as an operand of `:` whatever the other operand is.
  kLoad,
};

/// What `bind_unknown_qualifiers` reads from outside the formula.
struct UnknownQualifiers {
  /// True when `typed` names a sheet of this workbook.
  bool (*is_local_sheet)(const void* ctx, std::string_view typed);
  /// The existing extension-less link `typed` folds to, as `EnteredNames::link`.
  LinkSpelling (*link)(const void* ctx, std::string_view typed);
  const void* ctx;
};

/// Returns `formula` (a leading `=` optional) with every bare or quoted sheet qualifier that names no sheet of the
/// workbook rewritten to the external link of that name: `[book]sheet!` before a reference, the book-scope
/// `[book]!` before a name. As an operand of `:` it binds per `route` and is written quoted on the right of the
/// `:` (`X:'[book]sheet'!ref`), where `X:Y!ref` with only `Y` local also writes `Y` quoted. A qualifier that is a
/// span, a book, a path or a workbook file name, a Boolean word, or one before a call, a table or a spill anchor is
/// left as typed, as are string literals, the inside of brackets and external qualifiers.
std::string bind_unknown_qualifiers(std::string_view formula, const UnknownQualifiers& unknown,
                                    QualifierBindingRoute route);

/// Returns entered `formula` (a leading `=` optional) with the spelling Excel
/// stores for its names: a built-in function name in its canonical spelling,
/// TRUE / FALSE (also as `FALSE()`), error literals and A1 cell, column and
/// range references upper-cased; a sheet qualifier, a defined name and the
/// table of a structured reference respelled to what they name, a qualifier
/// quoted only where needed, a qualified name its sheet does not define
/// requalified as `[0]!` when it is a workbook-scope name; every use of a LET / LAMBDA parameter in the
/// spelling of its definition; an unknown function as `unknown_call` spells
/// it. With `link`, unknown qualifiers are first bound by
/// `bind_unknown_qualifiers` on the entry route. A name that
/// resolves to nothing, string
/// literals, the inside of structured-reference brackets, external
/// qualifiers and whitespace are copied verbatim; a 3-D qualifier is
/// respelled only when neither sheet needs quotes. Nothing is validated.
std::string spell_entered_names(std::string_view formula, const EnteredNames& names);

/// Where a sheet sits in the tab order of the book a qualifier names, for
/// `order_sheet_spans`.
struct SheetOrder {
  /// Tab position of `sheet` in the book written with directory `path`
  /// (empty for none) and file `book`; an empty `book` is this workbook.
  /// `kUnknownSheet` when the book or the sheet is not known.
  std::size_t (*position)(const void* ctx, std::string_view path, std::string_view book, std::string_view sheet);
  const void* ctx;
};

inline constexpr std::size_t kUnknownSheet = static_cast<std::size_t>(-1);

/// Returns entered `formula` (a leading `=` optional) with every 3-D span
/// whose first sheet comes after its last in tab order written first-to-last
/// (`Sheet3:Sheet1!A1` -> `Sheet1:Sheet3!A1`), local or external, quoted or
/// bare. A span with an end `order` does not know is copied verbatim, as is
/// everything else.
std::string order_sheet_spans(std::string_view formula, const SheetOrder& order);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_LITERAL_SPELLING_H_
