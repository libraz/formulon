// Copyright 2026 libraz. Licensed under the MIT License.
//
// Anchor projection for oracle comparison.
//
// Excel 365 spills multi-cell results into the cells adjacent to the
// formula cell, and xlwings reads back only the top-left "anchor" cell
// of that spill region. Formulon, evaluating in isolation, surfaces the
// full Array Value. To compare against the anchor scalar Excel reports,
// the oracle comparator projects the Array down to cell [0][0].

#ifndef FORMULON_TESTS_ORACLE_ORACLE_ANCHOR_H_
#define FORMULON_TESTS_ORACLE_ORACLE_ANCHOR_H_

#include "value.h"

namespace formulon {
namespace tests {
namespace oracle {

/// Returns the anchor cell (top-left) of `v` when it is a non-empty
/// Array Value, or `v` itself for any other kind.
///
/// Empty arrays (rows == 0 || cols == 0) are returned unchanged so the
/// caller can surface a kind mismatch through its existing error path.
const Value& anchor_or_self(const Value& v);

}  // namespace oracle
}  // namespace tests
}  // namespace formulon

#endif  // FORMULON_TESTS_ORACLE_ORACLE_ANCHOR_H_
