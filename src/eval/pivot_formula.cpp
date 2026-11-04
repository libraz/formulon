//
// Pivot formula host backed by the standard parser and tree walker.

#include "eval/pivot_formula.h"

namespace formulon::eval {

Expected<std::uint32_t, Error> PivotFormulaEvaluator::compile(std::string_view /*formula*/,
                                                              const pivot::FormulaScope& /*scope*/,
                                                              std::vector<pivot::FormulaRef>* /*refs*/) {
  return make_error(FormulonErrorCode::kInvalidArgument, "pivot formulas are not supported yet");
}

Value PivotFormulaEvaluator::evaluate(std::uint32_t /*formula_id*/, const Value* /*ref_values*/,
                                      std::size_t /*ref_count*/) {
  return Value::error(ErrorCode::Value);
}

}  // namespace formulon::eval
