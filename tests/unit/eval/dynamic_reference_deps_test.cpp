//
// Recalculation through references OFFSET and INDIRECT resolve while the
// formula runs. The circularity rules are measured on Mac Excel 365
// (16.113.2): ROW / ROWS / ISREF / AREAS on their own cell are not
// circular, OFFSET / INDIRECT resolving onto their own cell are, and the
// verdict follows the resolved target; a circular cell keeps its last value.

#include <cstdint>
#include <string>
#include <vector>

#include "calc_settings.h"
#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::uint32_t kA = 0U;
constexpr std::uint32_t kB = 1U;
constexpr std::uint32_t kC = 2U;
constexpr std::uint32_t kD = 3U;
constexpr std::uint32_t kE = 4U;
constexpr std::uint32_t kF = 5U;

// Row / column of an A1 cell as 0-based coordinates.
struct At {
  std::uint32_t row;
  std::uint32_t col;
};

void Formula(Workbook& wb, At at, const char* text) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, at.row, at.col, text))) << text;
}

void Number(Workbook& wb, At at, double v) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, at.row, at.col, Value::number(v))));
}

eval::RecalcStats Recalc(Workbook& wb) {
  auto stats = wb.recalc(eval::default_registry());
  EXPECT_TRUE(static_cast<bool>(stats)) << stats.error().message;
  return stats ? stats.value() : eval::RecalcStats{};
}

Value At_(const Workbook& wb, At at) {
  return wb.sheet(0).resolve_cell_value(at.row, at.col);
}

void ExpectNumberAt(const Workbook& wb, At at, double want, const char* what) {
  const Value v = At_(wb, at);
  ASSERT_TRUE(v.is_number()) << what << " -> " << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), want) << what;
}

TEST(DynamicReferenceRecalc, ReferenceOnlyFunctionsOnTheirOwnCellAreNotCircular) {
  const char* formulas[] = {"=ROW(A1)", "=ROWS(A1)", "=AREAS(A1)", "=ROWS(OFFSET(A1,0,0))"};
  for (const char* f : formulas) {
    Workbook wb = Workbook::create();
    Formula(wb, {0U, kA}, f);
    const eval::RecalcStats stats = Recalc(wb);
    EXPECT_EQ(stats.cycle_cells, 0U) << f;
    ExpectNumberAt(wb, {0U, kA}, 1.0, f);
  }
  Workbook wb = Workbook::create();
  Formula(wb, {0U, kA}, "=ISREF(A1)");
  EXPECT_EQ(Recalc(wb).cycle_cells, 0U);
  ASSERT_TRUE(At_(wb, {0U, kA}).is_boolean());
  EXPECT_TRUE(At_(wb, {0U, kA}).as_boolean());
}

// A reference bound to a LET name, a LAMBDA parameter or a helper callback
// parameter reads its cell only where the name is used as a value, so a
// position-only use of the formula's own cell is not circular. Measured on
// Mac Excel 365 with each formula in A1 and A2 = 5.
Workbook OwnCellBook(const char* formula) {
  Workbook wb = Workbook::create();
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("RowOf", "=LAMBDA(r,ROW(r))")));
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("ValOf", "=LAMBDA(r,r+1)")));
  Number(wb, {1U, kA}, 5.0);
  Formula(wb, {0U, kA}, formula);
  return wb;
}

TEST(DynamicReferenceRecalc, PositionOnlyUseOfBoundOwnCellIsNotCircular) {
  struct Case {
    const char* formula;
    double want;
  };
  const Case cases[] = {
      {"=LET(r,A1,ROW(r))", 1.0},
      {"=LET(r,A1,ROW(OFFSET(r,1,0)))", 2.0},
      {"=LET(r,A1,OFFSET(r,1,0))", 5.0},
      {"=LET(r,A1,ROWS(r)+COLUMNS(r)+AREAS(r))", 3.0},
      {"=LET(r,A1,s,r,ROW(s))", 1.0},
      {"=LET(r,A1,SUM(ROW(r),1))", 2.0},
      {"=LET(r,A1,LAMBDA(q,ROW(q))(r))", 1.0},
      {"=LET(f,LAMBDA(r,ROW(r)),f(A1))", 1.0},
      {"=LAMBDA(r,ROW(r))(A1)", 1.0},
      {"=RowOf(A1)", 1.0},
      {"=MAP(A1,LAMBDA(x,ROW(x)))", 1.0},
      {"=BYROW(A1,LAMBDA(x,ROW(x)))", 1.0},
      {"=REDUCE(0,A1,LAMBDA(a,v,ROW(v)))", 1.0},
      {"=SCAN(0,A1,LAMBDA(a,v,ROW(v)))", 1.0},
  };
  for (const Case& c : cases) {
    Workbook wb = OwnCellBook(c.formula);
    EXPECT_EQ(Recalc(wb).cycle_cells, 0U) << c.formula;
    ExpectNumberAt(wb, {0U, kA}, c.want, c.formula);
  }
  Workbook wb = OwnCellBook("=LET(r,A1,ISREF(r))");
  EXPECT_EQ(Recalc(wb).cycle_cells, 0U);
  ASSERT_TRUE(At_(wb, {0U, kA}).is_boolean());
  EXPECT_TRUE(At_(wb, {0U, kA}).as_boolean());
}

TEST(DynamicReferenceRecalc, ValueUseOfBoundOwnCellIsCircular) {
  const char* formulas[] = {
      "=LET(r,A1,r+1)",           "=LET(r,A1,ROW(r)+r)",
      "=LET(r,A1,OFFSET(r,0,0))", "=LET(r,A1,CELL(\"contents\",r))",
      "=LET(r,A1,INDEX(r,1,1))",  "=LET(r,A1,s,r,s+1)",
      "=LAMBDA(r,ROW(r)+r)(A1)",  "=ValOf(A1)",
      "=MAP(A1,LAMBDA(x,x+1))",   "=REDUCE(0,A1,LAMBDA(a,v,a+v))",
  };
  for (const char* f : formulas) {
    Workbook wb = OwnCellBook(f);
    EXPECT_GT(Recalc(wb).cycle_cells, 0U) << f;
  }
}

// A reference bound unwalked still reaches the graph through a value use.
TEST(DynamicReferenceRecalc, ValueUseOfBoundReferenceRecalcs) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("ValOf", "=LAMBDA(r,r+1)")));
  Number(wb, {0U, kB}, 3.0);
  const char* formulas[] = {"=LET(r,B1,s,r,s+1)", "=LAMBDA(r,r+1)(B1)", "=ValOf(B1)", "=MAP(B1,LAMBDA(x,x+1))",
                            "=LET(f,LAMBDA(r,r+1),LET(q,B1,f(q)))"};
  for (std::uint32_t row = 0U; row < 5U; ++row) {
    Formula(wb, {row, kD}, formulas[row]);
  }
  Recalc(wb);
  Number(wb, {0U, kB}, 9.0);
  Recalc(wb);
  for (std::uint32_t row = 0U; row < 5U; ++row) {
    ExpectNumberAt(wb, {row, kD}, 10.0, formulas[row]);
  }
}

// CELL reads its reference's value only for "contents" and "type", and a
// reference-returning call in a position-only slot only positions the
// arguments it can return; its other arguments are value reads. Measured on
// Mac Excel 365 with each formula in A1 and B1 = 1.
TEST(DynamicReferenceRecalc, PositionOnlyCellInfoAndReferenceCallsOnOwnCell) {
  const char* not_circular[] = {
      "=CELL(\"address\",A1)",     "=CELL(\"col\",A1)",          "=CELL(\"Row\",A1)",
      "=CELL(\"filename\",A1)",    "=CELL(\"format\",A1)",       "=CELL(\"color\",A1)",
      "=CELL(\"prefix\",A1)",      "=CELL(\"protect\",A1)",      "=CELL(\"width\",A1)",
      "=CELL(\"parentheses\",A1)", "=LET(r,A1,CELL(\"row\",r))", "=ROW(INDEX(A1,1,1))",
      "=ROW(IF(TRUE,A1,B1))",      "=ROW(CHOOSE(1,A1,B1))",      "=CELL(\"row\",INDEX(A1:A2,1))",
      "=ISREF(INDEX(A1:A2,1))",    "=ROW(INDEX((A1,B1),1,1,2))", "=LET(r,A1,ROW(INDEX(r,1,1)))",
  };
  for (const char* f : not_circular) {
    Workbook wb = Workbook::create();
    Number(wb, {0U, kB}, 1.0);
    Formula(wb, {0U, kA}, f);
    EXPECT_EQ(Recalc(wb).cycle_cells, 0U) << f;
  }
  const char* circular[] = {
      "=CELL(\"contents\",A1)",        "=CELL(\"CONTENTS\",A1)", "=CELL(\"Type\",A1)",
      "=ROW(INDEX(A1:A3,A1+1))",       "=ROW(OFFSET(A1,A1,0))",  "=ROW(IF(A1=0,A1,B1))",
      "=ROWS(XLOOKUP(1,A1:A3,B1:B3))",
  };
  for (const char* f : circular) {
    Workbook wb = Workbook::create();
    Number(wb, {0U, kB}, 1.0);
    Formula(wb, {0U, kA}, f);
    EXPECT_GT(Recalc(wb).cycle_cells, 0U) << f;
  }
}

// A computed CELL info_type leaves the reference unread by the formula text:
// a position key on the formula's own cell is not circular (measured on Mac
// Excel 365), and a value key's read is learned as the formula runs.
TEST(DynamicReferenceRecalc, ComputedCellInfoTypeReadsAtRunTime) {
  for (const char* info : {"row", "address", "prefix"}) {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, kB, Value::text(info))));
    Formula(wb, {0U, kA}, "=CELL(B2,A1)");
    EXPECT_EQ(Recalc(wb).cycle_cells, 0U) << info;
  }
  for (const char* info : {"contents", "type"}) {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, kB, Value::text(info))));
    Formula(wb, {0U, kA}, "=LEN(CELL(B2,A1)&\"\")+100");
    EXPECT_GT(Recalc(wb).cycle_cells, 0U) << info;
    ExpectNumberAt(wb, {0U, kA}, 0.0, info);
  }
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, kB, Value::text("contents"))));
  Number(wb, {4U, kD}, 3.0);
  Formula(wb, {0U, kA}, "=CELL(B2,C5)");
  Formula(wb, {4U, kC}, "=D5*2");
  Recalc(wb);
  ExpectNumberAt(wb, {0U, kA}, 6.0, "CELL(B2,C5)");
  Number(wb, {4U, kD}, 4.0);
  Recalc(wb);
  ExpectNumberAt(wb, {0U, kA}, 8.0, "CELL(B2,C5)");
}

TEST(DynamicReferenceRecalc, SumOverItsOwnCellIsCircular) {
  Workbook wb = Workbook::create();
  Formula(wb, {0U, kA}, "=SUM(A1:A2)");
  EXPECT_GT(Recalc(wb).cycle_cells, 0U);
}

TEST(DynamicReferenceRecalc, OffsetOntoAnotherCellIsNotCircular) {
  // D1 = 7, D2 = OFFSET(D2,-1,0): 7, no circular reference.
  Workbook wb = Workbook::create();
  Number(wb, {0U, kD}, 7.0);
  Formula(wb, {1U, kD}, "=OFFSET(D2,-1,0)");
  EXPECT_EQ(Recalc(wb).cycle_cells, 0U);
  ExpectNumberAt(wb, {1U, kD}, 7.0, "D2");
  // A nested base only positions the outer OFFSET.
  Formula(wb, {2U, kD}, "=OFFSET(OFFSET(D3,0,0),-2,0)");
  EXPECT_EQ(Recalc(wb).cycle_cells, 0U);
  ExpectNumberAt(wb, {2U, kD}, 7.0, "D3");
}

// The circular verdict follows the resolved target, and a circular cell
// keeps the value it showed before it became circular.
void ExpectCircularityFollowsTarget(const char* formula, At switch_cell, const Value& self, const Value& other) {
  Workbook wb = Workbook::create();
  Number(wb, {0U, kD}, 7.0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, switch_cell.row, switch_cell.col, other)));
  Formula(wb, {1U, kD}, formula);
  EXPECT_EQ(Recalc(wb).cycle_cells, 0U) << formula;
  ExpectNumberAt(wb, {1U, kD}, 7.0, formula);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, switch_cell.row, switch_cell.col, self)));
  EXPECT_GT(Recalc(wb).cycle_cells, 0U) << formula;
  ExpectNumberAt(wb, {1U, kD}, 7.0, formula);
  // Still circular on the next pass, still at its last value.
  EXPECT_GT(Recalc(wb).cycle_cells, 0U) << formula;
  ExpectNumberAt(wb, {1U, kD}, 7.0, formula);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, switch_cell.row, switch_cell.col, other)));
  Number(wb, {0U, kD}, 5.0);
  EXPECT_EQ(Recalc(wb).cycle_cells, 0U) << formula;
  ExpectNumberAt(wb, {1U, kD}, 5.0, formula);
}

TEST(DynamicReferenceRecalc, OffsetCircularityFollowsItsTarget) {
  ExpectCircularityFollowsTarget("=OFFSET(D2,E1,0)", {0U, kE}, Value::number(0.0), Value::number(-1.0));
}

TEST(DynamicReferenceRecalc, IndirectCircularityFollowsItsTarget) {
  ExpectCircularityFollowsTarget("=INDIRECT(E1)", {0U, kE}, Value::text("D2"), Value::text("D1"));
}

TEST(DynamicReferenceRecalc, CircularCellKeepsItsLastValueEvenWhenItWouldChangeIt) {
  // OFFSET(D2,E1,0)+1 onto itself must not step once per recalc.
  Workbook wb = Workbook::create();
  Number(wb, {0U, kD}, 7.0);
  Number(wb, {0U, kE}, -1.0);
  Formula(wb, {1U, kD}, "=OFFSET(D2,E1,0)+1");
  Recalc(wb);
  ExpectNumberAt(wb, {1U, kD}, 8.0, "D2");
  Number(wb, {0U, kE}, 0.0);
  for (int pass = 0; pass < 3; ++pass) {
    EXPECT_GT(Recalc(wb).cycle_cells, 0U);
    ExpectNumberAt(wb, {1U, kD}, 8.0, "D2 circular");
  }
}

TEST(DynamicReferenceRecalc, IterativeOffsetOntoItselfMatchesADirectReference) {
  IterativeOptions opts;
  opts.enabled = true;
  Workbook direct = Workbook::create();
  direct.set_iterative_options(opts);
  Formula(direct, {0U, kA}, "=A1+1");
  Workbook through_offset = Workbook::create();
  through_offset.set_iterative_options(opts);
  Formula(through_offset, {0U, kA}, "=OFFSET(A1,0,0)+1");
  // Only the entry pass is compared: OFFSET is volatile, so later passes
  // iterate it again where the non-volatile A1+1 is left alone.
  Recalc(direct);
  Recalc(through_offset);
  const Value want = At_(direct, {0U, kA});
  ASSERT_TRUE(want.is_number()) << want.debug_to_string();
  ExpectNumberAt(through_offset, {0U, kA}, want.as_number(), "OFFSET(A1,0,0)+1");
}

// Formulas that read, through OFFSET / INDIRECT, a formula whose value
// changes see the new value after one recalc, including when the OFFSET
// arguments move the target.
Workbook FreshReadBook() {
  Workbook wb = Workbook::create();
  Number(wb, {0U, kB}, 3.0);
  Number(wb, {0U, kF}, 5.0);
  // The readers sit above and left of their targets so evaluation order
  // alone would reach them first.
  Formula(wb, {0U, kC}, "=OFFSET(A1,5,0)");
  Formula(wb, {1U, kC}, "=INDIRECT(\"A6\")");
  Formula(wb, {2U, kC}, "=OFFSET(A1,F1,0)");
  Formula(wb, {3U, kC}, "=LET(r,A1,OFFSET(r,5,0))");
  Formula(wb, {4U, kC}, "=SUM(A1:OFFSET(A1,5,0))");
  Formula(wb, {5U, kA}, "=B1*2");
  Formula(wb, {6U, kA}, "=B1*3");
  return wb;
}

void ExpectFreshReads(const Workbook& wb, double b1) {
  ExpectNumberAt(wb, {0U, kC}, b1 * 2.0, "OFFSET(A1,5,0)");
  ExpectNumberAt(wb, {1U, kC}, b1 * 2.0, "INDIRECT(\"A6\")");
  ExpectNumberAt(wb, {3U, kC}, b1 * 2.0, "LET(r,A1,OFFSET(r,5,0))");
  ExpectNumberAt(wb, {4U, kC}, b1 * 2.0, "SUM(A1:OFFSET(A1,5,0))");
}

TEST(DynamicReferenceRecalc, ReadersSeeTheNewValueAfterOneRecalc) {
  Workbook wb = FreshReadBook();
  Recalc(wb);
  ExpectFreshReads(wb, 3.0);
  ExpectNumberAt(wb, {2U, kC}, 6.0, "OFFSET(A1,F1,0) -> A6");
  Number(wb, {0U, kB}, 4.0);
  Recalc(wb);
  ExpectFreshReads(wb, 4.0);
  ExpectNumberAt(wb, {2U, kC}, 8.0, "OFFSET(A1,F1,0) -> A6");
  Number(wb, {0U, kF}, 6.0);
  Number(wb, {0U, kB}, 5.0);
  Recalc(wb);
  ExpectFreshReads(wb, 5.0);
  ExpectNumberAt(wb, {2U, kC}, 15.0, "OFFSET(A1,F1,0) -> A7");
}

TEST(DynamicReferenceRecalc, ParallelRecalcMatchesSerial) {
  for (int run = 0; run < 8; ++run) {
    Workbook wb = FreshReadBook();
    eval::SchedulerConfig cfg;
    cfg.num_threads = 2U;
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(eval::default_registry(), cfg, nullptr)));
    ExpectFreshReads(wb, 3.0);
    Number(wb, {0U, kB}, 4.0);
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(eval::default_registry(), cfg, nullptr)));
    ExpectFreshReads(wb, 4.0);
  }
}

TEST(DynamicReferenceRecalc, PartialRecalcReadsATargetOutsideTheStaticClosure) {
  Workbook wb = FreshReadBook();
  Recalc(wb);
  Number(wb, {0U, kB}, 9.0);
  auto stats = wb.partial_recalc(eval::default_registry(), eval::SheetCellRange{0U, 0U, 0U, kC, kC});
  ASSERT_TRUE(static_cast<bool>(stats)) << stats.error().message;
  ExpectNumberAt(wb, {0U, kC}, 18.0, "C1 in viewport");
}

TEST(DynamicReferenceRecalc, OffsetChainSettlesInOneRecalc) {
  Workbook wb = Workbook::create();
  Number(wb, {0U, kA}, 1.0);
  for (std::uint32_t row = 1U; row < 100U; ++row) {
    const std::string f = "=OFFSET(A" + std::to_string(row + 1U) + ",-1,0)+1";
    Formula(wb, {row, kA}, f.c_str());
  }
  Recalc(wb);
  ExpectNumberAt(wb, {99U, kA}, 100.0, "A100");
  Number(wb, {0U, kA}, 11.0);
  Recalc(wb);
  ExpectNumberAt(wb, {99U, kA}, 110.0, "A100");
}

TEST(DynamicReferenceRecalc, ReplacingAReaderDropsItsLearnedEdges) {
  Workbook wb = FreshReadBook();
  Recalc(wb);
  const eval::CellNodeId c1{0U, 0U, kC};
  const eval::CellNodeId a6{0U, 5U, kA};
  auto engine_graph = [&]() -> const eval::DepGraph& { return wb.recalc_engine().dep_graph(); };
  EXPECT_TRUE(engine_graph().has_dependency_source(c1, a6, eval::DepGraph::DependencySource::kDynamicReference));
  Number(wb, {0U, kC}, 1.0);
  EXPECT_FALSE(engine_graph().has_dynamic_dependencies(c1));
}

// A reference bound to a LAMBDA parameter, a helper callback argument or a
// named LAMBDA's parameter stays a reference, so OFFSET inside the body reads
// through it and recalc learns the read; a plain reference argument remains
// a static precedent.
TEST(DynamicReferenceRecalc, ReferenceBoundLambdaParametersRecalc) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Below5", "=LAMBDA(r,OFFSET(r,5,0))")));
  Number(wb, {0U, kB}, 3.0);
  Formula(wb, {0U, kD}, "=LAMBDA(r,OFFSET(r,5,0))(A1)");
  Formula(wb, {1U, kD}, "=MAP(A1,LAMBDA(r,OFFSET(r,5,0)))");
  Formula(wb, {2U, kD}, "=Below5(A1)");
  Formula(wb, {3U, kD}, "=REDUCE(0,A1,LAMBDA(a,v,OFFSET(v,5,0)))");
  Formula(wb, {4U, kD}, "=LAMBDA(r,LAMBDA(s,OFFSET(s,5,0))(r))(A1)");
  Formula(wb, {5U, kD}, "=LAMBDA(r,r*10)(B1)");
  Formula(wb, {5U, kA}, "=B1*2");
  const char* readers[] = {"LAMBDA(r,OFFSET(r,5,0))(A1)", "MAP(A1,...)", "Below5(A1)", "REDUCE(0,A1,...)",
                           "nested LAMBDA"};
  Recalc(wb);
  for (std::uint32_t row = 0U; row < 5U; ++row) {
    ExpectNumberAt(wb, {row, kD}, 6.0, readers[row]);
  }
  ExpectNumberAt(wb, {5U, kD}, 30.0, "LAMBDA(r,r*10)(B1)");
  const eval::DepGraph& graph = wb.recalc_engine().dep_graph();
  EXPECT_TRUE(graph.has_dependency_source(eval::CellNodeId{0U, 0U, kD}, eval::CellNodeId{0U, 5U, kA},
                                          eval::DepGraph::DependencySource::kDynamicReference));
  EXPECT_TRUE(graph.has_dependency_source(eval::CellNodeId{0U, 5U, kD}, eval::CellNodeId{0U, 0U, kB},
                                          eval::DepGraph::DependencySource::kAuthored));
  Number(wb, {0U, kB}, 4.0);
  Recalc(wb);
  for (std::uint32_t row = 0U; row < 5U; ++row) {
    ExpectNumberAt(wb, {row, kD}, 8.0, readers[row]);
  }
  ExpectNumberAt(wb, {5U, kD}, 40.0, "LAMBDA(r,r*10)(B1)");
}

// A helper's callback is invoked, so the cells its body reads are
// precedents whether the lambda is written inline, bound by LET or named.
TEST(DynamicReferenceRecalc, HelperCallbackBodyReadsRecalc) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("TimesB1", "=LAMBDA(x,x*B1)")));
  Number(wb, {0U, kA}, 1.0);
  Number(wb, {0U, kB}, 3.0);
  Formula(wb, {0U, kD}, "=MAP(A1,LAMBDA(x,x*B1))");
  Formula(wb, {1U, kD}, "=LET(f,LAMBDA(x,x*B1),MAP(A1,f))");
  Formula(wb, {2U, kD}, "=MAP(A1,TimesB1)");
  Formula(wb, {3U, kD}, "=REDUCE(0,A1,LAMBDA(a,v,v*B1))");
  Recalc(wb);
  for (std::uint32_t row = 0U; row < 4U; ++row) {
    ExpectNumberAt(wb, {row, kD}, 3.0, "callback reads B1");
  }
  Number(wb, {0U, kB}, 5.0);
  Recalc(wb);
  for (std::uint32_t row = 0U; row < 4U; ++row) {
    ExpectNumberAt(wb, {row, kD}, 5.0, "callback reads B1");
  }
}

// With iterative calculation off, a cycle closing through an OFFSET read
// leaves every member -- the static one included -- at the value it showed
// before the recalc, an uncomputed member at 0, and recalculating changes
// nothing. Measured on Mac Excel 365 (A1 `=B1+1`, B1 `=OFFSET(A1,0,0)`).
enum class CycleDriver { kSerial, kParallel };

void RecalcWith(Workbook& wb, CycleDriver driver) {
  if (driver == CycleDriver::kSerial) {
    Recalc(wb);
    return;
  }
  eval::SchedulerConfig cfg;
  cfg.num_threads = 2U;
  const auto done = wb.recalc_parallel(eval::default_registry(), cfg, nullptr);
  EXPECT_TRUE(static_cast<bool>(done)) << done.error().message;
}

void ExpectCycleValues(const Workbook& wb, double a1, double b1, const char* what) {
  ExpectNumberAt(wb, {0U, kA}, a1, what);
  ExpectNumberAt(wb, {0U, kB}, b1, what);
}

class DynamicCycleRecalc : public ::testing::TestWithParam<CycleDriver> {};

TEST_P(DynamicCycleRecalc, ClosingMemberShowsZeroOthersKeepTheirValue) {
  Workbook wb = Workbook::create();
  Formula(wb, {0U, kA}, "=B1+1");
  RecalcWith(wb, GetParam());
  Formula(wb, {0U, kB}, "=OFFSET(A1,0,0)");
  RecalcWith(wb, GetParam());
  ExpectCycleValues(wb, 1.0, 0.0, "S1");
  RecalcWith(wb, GetParam());
  RecalcWith(wb, GetParam());
  ExpectCycleValues(wb, 1.0, 0.0, "S1 recalculated");
}

TEST_P(DynamicCycleRecalc, StaticMemberEnteredLastShowsZero) {
  Workbook wb = Workbook::create();
  Formula(wb, {0U, kB}, "=OFFSET(A1,0,0)");
  RecalcWith(wb, GetParam());
  Formula(wb, {0U, kA}, "=B1+1");
  RecalcWith(wb, GetParam());
  ExpectCycleValues(wb, 0.0, 0.0, "S2");
}

TEST_P(DynamicCycleRecalc, ValueReplacedByFormulaShowsZero) {
  Workbook wb = Workbook::create();
  Number(wb, {0U, kA}, 10.0);
  Formula(wb, {0U, kB}, "=OFFSET(A1,0,0)");
  RecalcWith(wb, GetParam());
  ExpectNumberAt(wb, {0U, kB}, 10.0, "S4 before");
  Formula(wb, {0U, kA}, "=B1+1");
  RecalcWith(wb, GetParam());
  ExpectCycleValues(wb, 0.0, 10.0, "S4");
}

TEST_P(DynamicCycleRecalc, MemberIsNotRecomputedWhenItsOtherInputChanges) {
  Workbook wb = Workbook::create();
  Number(wb, {0U, kC}, 3.0);
  Formula(wb, {0U, kA}, "=B1+C1");
  RecalcWith(wb, GetParam());
  Formula(wb, {0U, kB}, "=OFFSET(A1,0,0)");
  RecalcWith(wb, GetParam());
  ExpectCycleValues(wb, 3.0, 0.0, "S6");
  Number(wb, {0U, kC}, 4.0);
  RecalcWith(wb, GetParam());
  ExpectCycleValues(wb, 3.0, 0.0, "S6 after C1 = 4");
}

INSTANTIATE_TEST_SUITE_P(Drivers, DynamicCycleRecalc, ::testing::Values(CycleDriver::kSerial, CycleDriver::kParallel));

// Excel judges circularity by what evaluation reads: a formula whose only
// path back to its own cell sits in a branch never taken is not circular.
// Measured on Mac Excel 365 with each formula in A1.
TEST(ReadCircularity, BranchNotTakenIsNotCircular) {
  struct Case {
    const char* formula;
    double want;
  };
  const Case cases[] = {
      {"=LET(r,A1,IF(FALSE,r,ROW(r)))", 1.0},
      {"=IF(FALSE,A1,5)", 5.0},
      {"=IF(FALSE,A1+1,3)", 3.0},
      {"=CHOOSE(2,A1,7)", 7.0},
      {"=IFS(FALSE,A1,TRUE,8)", 8.0},
      {"=SWITCH(2,1,A1,2,9)", 9.0},
      {"=IFERROR(5,A1)", 5.0},
      {"=IFNA(5,A1)", 5.0},
      {"=IF(B1>0,A1,5)", 5.0},
      {"=XLOOKUP(1,{1},{6},A1)", 6.0},
      // Reads Excel does not count: formula text, formula-ness and
      // phonetic text.
      {"=LEN(FORMULATEXT(A1))", 21.0},
      {"=IF(ISFORMULA(A1),7,8)", 7.0},
      {"=LEN(PHONETIC(A1))+100", 100.0},
  };
  for (const Case& c : cases) {
    Workbook wb = Workbook::create();
    Formula(wb, {0U, kA}, c.formula);
    EXPECT_EQ(Recalc(wb).cycle_cells, 0U) << c.formula;
    ExpectNumberAt(wb, {0U, kA}, c.want, c.formula);
  }
}

TEST(ReadCircularity, ReadingTheOwnCellStaysCircular) {
  const char* formulas[] = {
      "=A1+1",
      "=IF(TRUE,A1,5)",
      "=CHOOSE(1,A1,7)",
      "=AND(FALSE,A1)",
      "=OR(TRUE,A1)",
      "=LET(x,A1+1,5)",
      "=LAMBDA(x,5)(A1+1)",
      "=SUM(A1:A3)",
      "=COUNTIF(A1:A3,1)",
      "=VLOOKUP(1,A1:B3,2,0)",
      "=SUMPRODUCT(A1:A3)",
      "=MATCH(1,A1:A3,0)",
      "=AVERAGE(A1:A3)",
      "=XLOOKUP(1,A1:A3,B1:B3)",
      "=SUMIF(A1:A3,\">0\")",
      "=MAX(A:A)",
      "=IF(ISBLANK(A1),7,8)",
      "=IF(ISNUMBER(A1),7,8)",
      "=N(A1)+100",
  };
  for (const char* f : formulas) {
    Workbook wb = Workbook::create();
    Formula(wb, {0U, kA}, f);
    EXPECT_GT(Recalc(wb).cycle_cells, 0U) << f;
  }
}

// Reading a spill reads its anchor's value: B1 `=SEQUENCE(2)+D1` with D1
// `=SUM(B1#)+100` is circular (measured on Mac Excel 365).
TEST(ReadCircularity, ReadingASpillOfTheComponentIsCircular) {
  Workbook wb = Workbook::create();
  Formula(wb, {0U, kB}, "=SEQUENCE(2)+D1");
  Formula(wb, {0U, kD}, "=SUM(B1#)+100");
  EXPECT_GT(Recalc(wb).cycle_cells, 0U);
}

// A1 `=IF(FALSE,B1,1)`, B1 `=A1+1`: a static cycle whose reads run one way,
// so both cells compute (Excel: 1 and 2), under either recalc driver and
// whichever cell was entered last.
TEST_P(DynamicCycleRecalc, StaticCycleWithOneWayReadsComputes) {
  for (const bool a_first : {true, false}) {
    Workbook wb = Workbook::create();
    if (a_first) {
      Formula(wb, {0U, kA}, "=IF(FALSE,B1,1)");
      Formula(wb, {0U, kB}, "=A1+1");
    } else {
      Formula(wb, {0U, kB}, "=A1+1");
      Formula(wb, {0U, kA}, "=IF(FALSE,B1,1)");
    }
    RecalcWith(wb, GetParam());
    ExpectCycleValues(wb, 1.0, 2.0, a_first ? "A1 first" : "B1 first");
  }
}

// The verdict follows the values: once the branch back to A1 is taken, the
// cycle is real and reported as every static cycle is (#REF!, the accepted
// divergence from Excel, which keeps the last value).
TEST(ReadCircularity, VerdictFollowsTheBranchTaken) {
  Workbook wb = Workbook::create();
  Number(wb, {0U, kB}, 0.0);
  Formula(wb, {0U, kA}, "=IF(B1>0,A1+1,5)");
  EXPECT_EQ(Recalc(wb).cycle_cells, 0U);
  ExpectNumberAt(wb, {0U, kA}, 5.0, "B1 = 0");
  Number(wb, {0U, kB}, 1.0);
  EXPECT_GT(Recalc(wb).cycle_cells, 0U);
  Number(wb, {0U, kB}, 0.0);
  EXPECT_EQ(Recalc(wb).cycle_cells, 0U);
  ExpectNumberAt(wb, {0U, kA}, 5.0, "B1 back to 0");
}

TEST(ReadCircularity, IterativeCalcStillSolvesTheComponent) {
  IterativeOptions opts;
  opts.enabled = true;
  Workbook wb = Workbook::create();
  wb.set_iterative_options(opts);
  Formula(wb, {0U, kA}, "=IF(FALSE,A1,5)");
  Recalc(wb);
  ExpectNumberAt(wb, {0U, kA}, 5.0, "IF(FALSE,A1,5)");
}

}  // namespace
}  // namespace formulon
