// Copyright 2026 libraz. Licensed under the MIT License.
//
// `DirtySet` is implemented entirely as inline members in the header
// (a thin wrapper around `std::unordered_set<CellNodeId>`). This
// translation unit exists so that the CMake `FORMULON_CORE_SOURCES` list
// has a stable anchor for the helper — and so that future additions
// (e.g. shard-aware merge, atomic-friendly variant for the MT scheduler)
// have an obvious home that does not force every consumer to recompile.

#include "eval/dirty_set.h"

namespace formulon::eval {

// Intentionally empty — see dirty_set.h.

}  // namespace formulon::eval
