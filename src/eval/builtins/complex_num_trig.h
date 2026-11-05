//
// Registers Excel's complex-number trigonometric and hyperbolic built-ins
// (IMSIN/IMCOS/IMTAN/IMSEC/IMCSC/IMCOT/IMSINH/IMCOSH/IMSECH/IMCSCH) into a
// FunctionRegistry. Kept in its own translation unit apart from the rest of
// the complex-number family.

#ifndef FORMULON_EVAL_BUILTINS_COMPLEX_NUM_TRIG_H_
#define FORMULON_EVAL_BUILTINS_COMPLEX_NUM_TRIG_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the complex-number trigonometric and hyperbolic built-in
/// functions (IMSIN, IMCOS, IMTAN, IMSEC, IMCSC, IMCOT, IMSINH, IMCOSH,
/// IMSECH, IMCSCH) into `registry`. Intended to be invoked from
/// `register_builtins`.
void register_complex_num_trig_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_COMPLEX_NUM_TRIG_H_
