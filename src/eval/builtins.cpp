// Copyright 2026 libraz. Licensed under the MIT License.
//
// Thin registrar that stitches together the per-family builtin catalogs.
// Each family lives in its own translation unit under `src/eval/builtins/`
// and publishes a single `register_<family>_builtins(FunctionRegistry&)`
// entry point. `register_builtins` simply invokes each in turn, keeping
// the public surface in `eval/builtins.h` stable.
//
// `IF`, `IFERROR`, and `IFNA` are intentionally absent: they short-circuit
// and are special-cased in the tree walker before the registry is consulted.

#include "eval/builtins.h"

#include "eval/builtins/aggregate.h"
#include "eval/builtins/datetime.h"
#include "eval/builtins/financial.h"
#include "eval/builtins/info.h"
#include "eval/builtins/logical.h"
#include "eval/builtins/math.h"
#include "eval/builtins/math_trig.h"
#include "eval/builtins/stats.h"
#include "eval/builtins/text.h"
#include "eval/function_registry.h"

namespace formulon {
namespace eval {

void register_builtins(FunctionRegistry& registry) {
  register_aggregate_builtins(registry);
  register_logical_builtins(registry);
  register_math_builtins(registry);
  register_math_trig_builtins(registry);
  register_stats_builtins(registry);
  register_text_builtins(registry);
  register_info_builtins(registry);
  register_datetime_builtins(registry);
  register_financial_builtins(registry);
}

}  // namespace eval
}  // namespace formulon
