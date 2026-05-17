// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Internal header for the rectangle-resolution / intersection routines
// that sit alongside INDIRECT and OFFSET in this subdirectory. The
// public declarations live in `eval/range_resolvers.h`; this header is
// reserved for forward declarations consumed only by sibling TUs in
// `reference/`. It is currently empty because the public surface needs
// no internal complement, but it stays around so the file split is
// symmetric with `indirect.h` / `offset.h` and so future helpers can
// land here without churning the build.

#ifndef FORMULON_EVAL_REFERENCE_INTERSECTION_H_
#define FORMULON_EVAL_REFERENCE_INTERSECTION_H_

#include "eval/range_resolvers.h"

#endif  // FORMULON_EVAL_REFERENCE_INTERSECTION_H_
