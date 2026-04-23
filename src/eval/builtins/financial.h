// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's core time-value-of-money built-ins (PV / FV / PMT /
// NPER / NPV / RATE / IPMT / PPMT / CUMIPMT / CUMPRINC) into a
// FunctionRegistry. IRR is not registered here — it lives on the lazy-
// dispatch seam in `eval/financial_lazy.h` because its first argument
// must reach the impl as an un-flattened AST shape (so a bare Ref, a
// RangeOp, or an ArrayLiteral can all be walked cell-by-cell for
// Newton-Raphson).
//
// Kept in its own translation unit so the financial family can evolve
// independently of the rest of the builtin catalog.

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the eager financial built-ins (PV, FV, PMT, NPER, NPV,
/// RATE, IPMT, PPMT, CUMIPMT, CUMPRINC, SLN, SYD, DDB, DB, DOLLARDE,
/// DOLLARFR, EFFECT, NOMINAL, FVSCHEDULE, PDURATION, RRI, ISPMT, DISC,
/// INTRATE, RECEIVED, TBILLPRICE, TBILLYIELD, TBILLEQ) into `registry`.
/// IRR and MIRR are wired separately via the lazy dispatch table in
/// `tree_walker.cpp`; see `eval/financial_lazy.h`.
void register_financial_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_H_
