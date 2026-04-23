// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's extended probability-distribution catalog: the beta,
// gamma, Weibull, lognormal, and hypergeometric families. The normal,
// binomial, Poisson, exponential, chi-squared, Student-t, and Snedecor-F
// distributions live in `eval/builtins/stats.cpp` because they share the
// MEDIAN / STDEV argument-coercion conventions; this translation unit
// specialises in distributions whose inverses lean on bracket-then-Newton
// iteration and whose PDFs / PMFs benefit from log-space evaluation.
//
// Functions registered here:
//   BETA.DIST / BETA.INV
//   GAMMA / GAMMALN / GAMMALN.PRECISE / GAMMA.DIST / GAMMA.INV
//   WEIBULL.DIST
//   LOGNORM.DIST / LOGNORM.INV
//   HYPGEOM.DIST
//
// Kept separate from `stats.cpp` so the extended distribution catalog can
// evolve (and share its root-finding helpers) without growing the already
// sizeable stats TU.

#ifndef FORMULON_EVAL_BUILTINS_DISTRIBUTIONS_H_
#define FORMULON_EVAL_BUILTINS_DISTRIBUTIONS_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the extended probability-distribution built-ins (beta, gamma,
/// Weibull, lognormal, hypergeometric) into `registry`. Called from
/// `register_builtins` alongside the other per-family registrars.
void register_distribution_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_DISTRIBUTIONS_H_
