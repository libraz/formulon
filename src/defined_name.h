//
// Defined-name model: the round-trip carrier for one `<definedName>`
// element of `xl/workbook.xml`. The defined formulas are NOT validated or
// evaluated here — that is the parser/evaluator's job at use-site. The
// pugi-typed reader lives in `io/defined_names_internal.h`.

#ifndef FORMULON_DEFINED_NAME_H_
#define FORMULON_DEFINED_NAME_H_

#include <cstdint>
#include <string>
#include <utility>

namespace formulon {

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
  /// Attributes not represented by the model (`function`, `vbProcedure`,
  /// `shortcutKey`, `description`, ...) retained as escaped ` name="value"`
  /// pairs.
  std::string extra_attrs;
};

}  // namespace formulon

#endif  // FORMULON_DEFINED_NAME_H_
