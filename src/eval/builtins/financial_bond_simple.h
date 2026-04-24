// Copyright 2026 libraz. Licensed under the MIT License.
//
// Internal header — do not include outside `src/eval/builtins/financial*`.
//
// Forward declarations for the closed-form bond-pricing builtins and the
// STOCKHISTORY stub. Implementations live in `financial_bond_simple.cpp`
// and are registered from `financial.cpp` via
// `register_financial_builtins`.
//
// Functions declared here:
//   * PRICEDISC   -- price per 100 face for a discounted security.
//   * PRICEMAT    -- price per 100 face for a security paying interest
//                    at maturity.
//   * YIELDDISC   -- annual yield for a discounted security.
//   * YIELDMAT    -- annual yield for a security paying interest at
//                    maturity.
//   * STOCKHISTORY -- intentional stub returning `#VALUE!`; Formulon does
//                    not perform network I/O (see the WEBSERVICE/PY
//                    precedent in `web.cpp`).

#ifndef FORMULON_EVAL_BUILTINS_FINANCIAL_BOND_SIMPLE_H_
#define FORMULON_EVAL_BUILTINS_FINANCIAL_BOND_SIMPLE_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace financial_detail {

Value PriceDisc(const Value* args, std::uint32_t arity, Arena& arena);
Value PriceMat(const Value* args, std::uint32_t arity, Arena& arena);
Value YieldDisc(const Value* args, std::uint32_t arity, Arena& arena);
Value YieldMat(const Value* args, std::uint32_t arity, Arena& arena);
Value StockHistory(const Value* args, std::uint32_t arity, Arena& arena);

}  // namespace financial_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_FINANCIAL_BOND_SIMPLE_H_
