// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Per-rule-kind matching for the CF evaluator: cellIs / expression /
// timePeriod / text-family / duplicate-unique / above-average / top10.
//
// The public-facing `match_rule(...)` overloads are declared in
// `cf/cf_evaluator.h`; this header exposes a few internal entry points
// used by `cf_evaluator.cpp` (for example `match_above_average` is
// also a CF helper consumed from the cross-block walker indirectly via
// `match_rule`, so the file lives here for cohesion even though the
// caller goes through the public overload).
//
// Visual rule kinds (ColorScale / DataBar / IconSet) live in
// `scale_evaluator.h`. The context-aware `match_rule(rule, cell, ctx)`
// dispatch in `rule_match.cpp` calls into both this and that header.

#ifndef FORMULON_CF_RULE_MATCH_H_
#define FORMULON_CF_RULE_MATCH_H_

// The public overloads `match_rule(rule, cell_value)` and
// `match_rule(rule, cell_value, ctx)` are declared in
// `cf/cf_evaluator.h`; including this header here keeps consumers from
// having to know which TU defines them.
#include "cf/cf_evaluator.h"  // IWYU pragma: export

#endif  // FORMULON_CF_RULE_MATCH_H_
