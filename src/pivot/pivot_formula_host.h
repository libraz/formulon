//
// Host interface through which the pivot layer evaluates calculated-field and
// calculated-item formulas without depending on the parser or evaluator.

#ifndef FORMULON_PIVOT_PIVOT_FORMULA_HOST_H_
#define FORMULON_PIVOT_PIVOT_FORMULA_HOST_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "pivot/pivot_cache.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot {

/// One pivot reference found in a formula.
struct FormulaRef {
  enum class Kind : std::uint8_t { Field, Item, ItemAt, ItemOffset };
  Kind kind;
  std::size_t field;      ///< Cache field index.
  std::size_t item;       ///< Shared-item index (`Item`).
  std::int32_t position;  ///< 1-based (`ItemAt`) or signed offset (`ItemOffset`).
};

/// Names a formula may use.
struct FormulaScope {
  const PivotCache* cache;                ///< Names come from cache fields / shared items.
  std::optional<std::size_t> item_field;  ///< Set for a calculated-item formula.
};

/// Compiles and evaluates pivot formulas on behalf of the pivot layer.
class PivotFormulaHost {
 public:
  /// Parses once and fills `refs` in first-appearance order. A non-null
  /// `canonical` receives the stored form: the text without a leading `=`, and
  /// in a calculated-item formula every bare item name qualified as `Field[Item]`.
  ///
  /// Errors: `kInvalidArgument` with a reason for syntax errors, names outside
  /// the scope, cell/range/structured/external references, defined names,
  /// array constants, the disallowed functions, and All/Blank.
  virtual Expected<std::uint32_t, Error> compile(std::string_view formula, const FormulaScope& scope,
                                                 std::vector<FormulaRef>* refs, std::string* canonical) = 0;

  /// Evaluates a compiled formula; `ref_values[i]` is the value of `refs[i]`
  /// for the cell being evaluated.
  virtual Value evaluate(std::uint32_t formula_id, const Value* ref_values, std::size_t ref_count) = 0;

 protected:
  ~PivotFormulaHost() = default;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_FORMULA_HOST_H_
