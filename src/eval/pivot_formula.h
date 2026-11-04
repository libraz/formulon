//
// Pivot formula host backed by the standard parser and tree walker.

#ifndef FORMULON_EVAL_PIVOT_FORMULA_H_
#define FORMULON_EVAL_PIVOT_FORMULA_H_

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "pivot/pivot_formula_host.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::eval {

/// Compiles and evaluates calculated-field / calculated-item formulas.
class PivotFormulaEvaluator final : public pivot::PivotFormulaHost {
 public:
  Expected<std::uint32_t, Error> compile(std::string_view formula, const pivot::FormulaScope& scope,
                                         std::vector<pivot::FormulaRef>* refs) override;
  Value evaluate(std::uint32_t formula_id, const Value* ref_values, std::size_t ref_count) override;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_PIVOT_FORMULA_H_
