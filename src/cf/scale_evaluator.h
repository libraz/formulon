//
// Color-scale / databar / iconSet resolution for the CF evaluator.
//
// This header is internal to the `cf/` subsystem. The public entry
// points are in `cf/cf_evaluator.h`; this file exposes the per-visual
// resolvers so `cf_evaluator.cpp` can build the `CFMatch` render
// payload and `rule_match.cpp` can ask "does this rule produce a
// usable visual?" without itself depending on the rendering details.

#ifndef FORMULON_CF_SCALE_EVALUATOR_H_
#define FORMULON_CF_SCALE_EVALUATOR_H_

#include <optional>

#include "cf/cf_evaluator.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "value.h"

namespace formulon::cf::scales {

/// Computes the resolved fill colour for a `colorScale` rule applied to
/// `cell_value`. Returns `nullopt` for empty populations, malformed
/// thresholds, or non-numeric cells.
std::optional<Color> resolve_color_scale(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx);

/// Boolean "rule applies" check: identical to "the colour resolves".
bool match_color_scale(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx);

/// Computes the bar length, axis position, and fill / border for a
/// `dataBar` rule applied to `cell_value`. Returns `nullopt` when the
/// thresholds collapse or are unresolvable.
std::optional<DataBarRender> resolve_data_bar(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx);

bool match_data_bar(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx);

/// Computes the icon index for an `iconSet` rule applied to
/// `cell_value`, honouring `gte` boundaries and the spec's `reverse`
/// flag.
std::optional<IconRender> resolve_icon_set(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx);

bool match_icon_set(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx);

}  // namespace formulon::cf::scales

#endif  // FORMULON_CF_SCALE_EVALUATOR_H_
