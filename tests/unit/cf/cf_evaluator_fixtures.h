//
// Shared fixtures for the conditional-format evaluator tests: rule and
// range builders, a sheet-backed evaluation harness, and colour-scale specs.

#ifndef FORMULON_TESTS_UNIT_CF_CF_EVALUATOR_FIXTURES_H_
#define FORMULON_TESTS_UNIT_CF_CF_EVALUATOR_FIXTURES_H_

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_evaluator.h"
#include "cf/cf_types.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon::cf::test {

inline CFRule MakeRule(RuleType t) {
  CFRule r;
  r.type = t;
  r.priority = 5;
  r.id = "rule-x";
  r.dxf_id = 7u;
  return r;
}

inline CellAddress At(std::uint32_t row, std::uint32_t col) {
  CellAddress addr{};
  addr.row = row;
  addr.col = col;
  return addr;
}

struct CFEvalHarness {
  Sheet sheet{"Sheet1"};
  Arena arena;
  eval::EvalContext eval_ctx{sheet};

  CFEvalContext context(CellAddress anchor, CellAddress target) {
    CFEvalContext ctx;
    ctx.anchor = anchor;
    ctx.target = target;
    ctx.arena = &arena;
    ctx.registry = &eval::default_registry();
    ctx.eval_ctx = &eval_ctx;
    return ctx;
  }
};

inline CFCellRange MakeRange(std::uint32_t r1, std::uint32_t c1, std::uint32_t r2, std::uint32_t c2) {
  CFCellRange range{};
  range.first.row = r1;
  range.first.col = c1;
  range.last.row = r2;
  range.last.col = c2;
  return range;
}

inline CFEvalContext SqrefContext(CFEvalHarness& harness, const std::vector<CFCellRange>& sqref) {
  CFEvalContext ctx = harness.context(At(0, 0), At(0, 0));
  ctx.sqref = &sqref;
  return ctx;
}

inline void PopulateLinearPopulation(CFEvalHarness& harness) {
  harness.sheet.set_cell_value(0, 0, Value::number(10.0));
  harness.sheet.set_cell_value(1, 0, Value::number(20.0));
  harness.sheet.set_cell_value(2, 0, Value::number(30.0));
  harness.sheet.set_cell_value(3, 0, Value::number(40.0));
  harness.sheet.set_cell_value(4, 0, Value::number(50.0));
}

inline Color RGB(std::uint8_t red, std::uint8_t green, std::uint8_t blue) {
  Color color{};
  color.r = red;
  color.g = green;
  color.b = blue;
  color.a = 255;
  return color;
}

inline CfValueObject Cfvo(CfvoType type, std::string value = "") {
  CfValueObject cfvo;
  cfvo.type = type;
  cfvo.value = std::move(value);
  return cfvo;
}

inline ColorScaleSpec TwoStopMinMax(Color lo, Color hi) {
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Min), Cfvo(CfvoType::Max)};
  spec.colors = {lo, hi};
  return spec;
}

}  // namespace formulon::cf::test

#endif  // FORMULON_TESTS_UNIT_CF_CF_EVALUATOR_FIXTURES_H_
