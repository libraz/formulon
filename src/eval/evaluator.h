// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Skeleton header kept for project-layout symmetry. The two real evaluator
// entry points live elsewhere:
//
//   * `tree_walker.h` -- the production tree-walking evaluator. Holds the
//     public `evaluate(...)` surface, operator dispatch, and the lazy
//     special-form glue that drives the engine today.
//   * `vm.h` / `vm.cpp` -- the stack-machine VM that executes bytecode
//     produced by `compiler.h`. Currently exercised in parallel under
//     `FORMULON_VM_PARITY_CHECK`; will become the default entry point in
//     a later bundle.
//
// External readers grep-ing for `evaluator.h` land here rather than
// concluding the file is missing. **Do not include this header from
// production code** — it declares no symbols.
//
// Design references:
//   * `src/eval/tree_walker.h`   - tree-walking evaluator (production)
//   * `src/eval/vm.h`            - bytecode VM (parity-mode today)
//   * `src/eval/compiler.h`      - AST -> bytecode lowering
//   * `src/eval/bytecode.h`      - bytecode IR definition

#ifndef FORMULON_EVAL_EVALUATOR_H_
#define FORMULON_EVAL_EVALUATOR_H_

#endif  // FORMULON_EVAL_EVALUATOR_H_
