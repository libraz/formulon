// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Per-thread re-entrancy flag shared between the serial and parallel recalc
// entry points. A nested invocation on the same thread (typically a future
// host-callable function or a progress callback that calls back into
// `Workbook::recalc()` / `recalc_parallel()`) trips the flag and surfaces
// `kGraphRecalcReentrant` rather than deadlocking on the engine mutex.
//
// The flag is `inline thread_local`: one logical variable shared across all
// translation units, with a per-thread instance. Both the parallel scheduler
// and `RecalcEngine::recalc` consult it so a callback issued from inside a
// parallel pass cannot re-enter the serial engine, and vice versa.

#ifndef FORMULON_EVAL_RECALC_REENTRY_H_
#define FORMULON_EVAL_RECALC_REENTRY_H_

namespace formulon {
namespace eval {
namespace detail {

inline thread_local bool g_in_recalc = false;

/// RAII guard that scopes `g_in_recalc` to a recalc invocation. Captures
/// the prior value so it restores correctly even when nested invocations
/// are blocked further up the stack.
struct RecalcReentryGuard {
  bool prev;
  RecalcReentryGuard() noexcept : prev(g_in_recalc) { g_in_recalc = true; }
  ~RecalcReentryGuard() { g_in_recalc = prev; }
  RecalcReentryGuard(const RecalcReentryGuard&) = delete;
  RecalcReentryGuard& operator=(const RecalcReentryGuard&) = delete;
  RecalcReentryGuard(RecalcReentryGuard&&) = delete;
  RecalcReentryGuard& operator=(RecalcReentryGuard&&) = delete;
};

}  // namespace detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_RECALC_REENTRY_H_
