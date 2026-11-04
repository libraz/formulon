//
// Pivot formula host backed by the standard parser and tree walker.
//
// Pivot names (quoted names, `Field[...]` references and identifiers equal to
// a field or item name) are rewritten to synthetic identifiers before the
// standard parser sees the text, then bound through `NameEnv` at evaluation
// time, so defined names are never reached.

#ifndef FORMULON_EVAL_PIVOT_FORMULA_H_
#define FORMULON_EVAL_PIVOT_FORMULA_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "parser/ast.h"
#include "pivot/pivot_formula_host.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::eval {

/// Compiles and evaluates calculated-field / calculated-item formulas.
///
/// Ids are indices into this instance; each compiled AST lives as long as the
/// host. A text result stays valid until the next `evaluate` call.
class PivotFormulaEvaluator final : public pivot::PivotFormulaHost {
 public:
  Expected<std::uint32_t, Error> compile(std::string_view formula, const pivot::FormulaScope& scope,
                                         std::vector<pivot::FormulaRef>* refs, std::string* canonical) override;
  Value evaluate(std::uint32_t formula_id, const Value* ref_values, std::size_t ref_count) override;

  /// Capacity held by the per-call scratch arena; for tests and diagnostics.
  std::size_t scratch_bytes_allocated() const noexcept { return scratch_.bytes_allocated(); }

 private:
  struct Compiled {
    Arena arena;
    const parser::AstNode* root = nullptr;
    std::size_t ref_count = 0;
  };

  std::vector<Compiled> formulas_;
  /// `names_[i]` is the synthetic identifier bound to `ref_values[i]`.
  std::vector<std::string> names_;
  Arena scratch_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_PIVOT_FORMULA_H_
