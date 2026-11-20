//
// Defined-name model: the round-trip carrier for one `<definedName>`
// element of `xl/workbook.xml`. The defined formulas are NOT validated or
// evaluated here — that is the parser/evaluator's job at use-site. The
// pugi-typed reader lives in `io/defined_names_internal.h`.

#ifndef FORMULON_DEFINED_NAME_H_
#define FORMULON_DEFINED_NAME_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "excel_profile.h"

namespace formulon {

/// The macro a defined name names. Only an .xlsb carries one (`fProc`, with
/// `fFunc` for a function macro); an .xlsx without a macro sheet cannot, so
/// Excel's .xlsx save leaves such a name out.
enum class NameMacroKind : std::uint8_t {
  kNone = 0,
  kCommand = 1,
  kFunction = 2,
};

/// Where `Workbook::set_defined_name_scoped` stored a name. A name that
/// spells a built-in can be stored under another spelling and scope than
/// the ones asked for.
struct DefinedNameKey {
  std::string name;
  std::int32_t local_sheet_id = -1;
};

/// In-memory representation of one `<definedName>` element from
/// `xl/workbook.xml`. `name` is preserved with its authored case
/// (although Excel matches names case-insensitively at use-site);
/// `formula` is the raw formula text (Excel sometimes wraps the body in
/// surrounding whitespace, which the reader strips). `local_sheet_id`
/// follows the OOXML attribute semantics: `-1` means workbook scope,
/// `>= 0` is a 0-based sheet index for sheet-scoped names. `comment`
/// is the optional human-readable note Excel exposes in the Name
/// Manager dialog.
struct DefinedName {
  DefinedName() = default;
  // Keeps the five-field construction shape for call sites that initialise
  // the modelled fields directly.
  DefinedName(std::string name_in, std::string formula_in, std::int32_t local_sheet_id_in, bool hidden_in,
              std::string comment_in)
      : name(std::move(name_in)),
        formula(std::move(formula_in)),
        local_sheet_id(local_sheet_id_in),
        hidden(hidden_in),
        comment(std::move(comment_in)) {}

  std::string name;
  std::string formula;
  std::int32_t local_sheet_id = -1;
  bool hidden = false;
  std::string comment;
  /// Set by the .xlsb reader for a name flagged as a macro; the name is live
  /// and evaluates as any other.
  NameMacroKind macro_kind = NameMacroKind::kNone;
  /// The macro name's `fFutureFunction` flag, kept for the .xlsb writer.
  bool macro_future_function = false;
  /// The .xlsb function category (`fGrp`, 0-511), kept for the .xlsb writer.
  /// Excel sets 14 when a macro name's formula is assigned.
  std::uint16_t function_group = 0;
  /// Attributes not represented by the model (`function`, `vbProcedure`,
  /// `shortcutKey`, `description`, ...) retained as escaped ` name="value"`
  /// pairs.
  std::string extra_attrs;
};

/// A defined name a load renamed to the built-in its stored spelling names in
/// the load locale: `typed` was its stored name, `stored` is its `_xlnm.` one.
struct LoadedNameFold {
  std::string typed;
  std::string stored;
  std::int32_t local_sheet_id = -1;
};

/// A name a reader holds before installing it: its stored spelling and scope.
struct LoadedNameScope {
  std::string_view name;
  std::int32_t local_sheet_id = -1;
};

/// Per entry of `names`, in order: the `_xlnm.` name it folds to on a load
/// under `locale`, or empty. An entry folds when it has no `_xlnm.` prefix,
/// equals the locale's spelling of a built-in X, and no `_xlnm.X` of `names`
/// is visible from its scope (that scope, or workbook scope for a sheet-local
/// entry); the folded spelling is X's canonical one.
std::vector<std::string> loaded_builtin_name_folds(const std::vector<LoadedNameScope>& names, ExcelLocale locale);

}  // namespace formulon

#endif  // FORMULON_DEFINED_NAME_H_
