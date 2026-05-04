// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Reserved for the future bytecode VM. The current evaluator entry points
// are in `tree_walker.h`; that file holds the public `evaluate(...)`
// surface, the operator dispatch, and the lazy-special-form glue that the
// rest of the engine drives today.
//
// This header is an intentional placeholder so the project layout matches
// what the design corpus calls out (a dedicated `evaluator` translation
// unit alongside the tree-walker), and so external readers grep-ing the
// source tree for "evaluator.h" land here rather than concluding the file
// is missing. **Do not include this header from production code** — it
// declares no symbols. The bytecode VM will land in this slot when the
// `bytecode.h` IR has stabilised; until then any inclusion is a build
// error by virtue of producing no public surface.
//
// Design references:
//   * `src/eval/tree_walker.h`   - current evaluator entry points
//   * `src/eval/bytecode.h`      - IR sketch the future VM will consume

#ifndef FORMULON_EVAL_EVALUATOR_H_
#define FORMULON_EVAL_EVALUATOR_H_

#endif  // FORMULON_EVAL_EVALUATOR_H_
