//
// Stable C ABI trace (precedents / dependents) regression tests.

#include <cstdint>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
};

struct CellNodesGuard {
  fm_cell_nodes_t* handle = nullptr;
  ~CellNodesGuard() { fm_cell_nodes_destroy(handle); }
};

}  // namespace

TEST(FormulonCApiTrace, EmptyWorkbookReturnsEmpty) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  CellNodesGuard nodes;
  ASSERT_EQ(fm_workbook_precedents(wb.handle, 0, 0, 0, 1, &nodes.handle), 0);
  EXPECT_EQ(fm_cell_nodes_count(nodes.handle), 0U);
}

TEST(FormulonCApiTrace, DirectPrecedentsOfFormula) {
  // A1=10, B1=20, C1=A1+B1 -> precedents(C1) = {A1, B1}
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=10"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 1, "=20"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 2, "=A1+B1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  CellNodesGuard nodes;
  ASSERT_EQ(fm_workbook_precedents(wb.handle, 0, 0, 2, 1, &nodes.handle), 0);
  ASSERT_EQ(fm_cell_nodes_count(nodes.handle), 2U);
  bool saw_a1 = false;
  bool saw_b1 = false;
  for (std::size_t i = 0; i < 2; ++i) {
    fm_cell_node_t n{};
    ASSERT_EQ(fm_cell_nodes_at(nodes.handle, i, &n), 0);
    EXPECT_EQ(n.sheet, 0U);
    EXPECT_EQ(n.row, 0U);
    if (n.col == 0U)
      saw_a1 = true;
    if (n.col == 1U)
      saw_b1 = true;
  }
  EXPECT_TRUE(saw_a1);
  EXPECT_TRUE(saw_b1);
}

TEST(FormulonCApiTrace, DirectDependentsOfRoot) {
  // A1=10, B1=A1*2, C1=A1+B1 -> dependents(A1) = {B1, C1}
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=10"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 1, "=A1*2"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 2, "=A1+B1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  CellNodesGuard nodes;
  ASSERT_EQ(fm_workbook_dependents(wb.handle, 0, 0, 0, 1, &nodes.handle), 0);
  ASSERT_EQ(fm_cell_nodes_count(nodes.handle), 2U);
}

TEST(FormulonCApiTrace, DepthExpandsTransitively) {
  // A1=1, B1=A1, C1=B1 -> precedents(C1, depth=2) includes A1
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=1"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 1, "=A1"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 2, "=B1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  // depth=1 → only B1
  CellNodesGuard d1;
  ASSERT_EQ(fm_workbook_precedents(wb.handle, 0, 0, 2, 1, &d1.handle), 0);
  EXPECT_EQ(fm_cell_nodes_count(d1.handle), 1U);

  // depth=2 → B1 + A1
  CellNodesGuard d2;
  ASSERT_EQ(fm_workbook_precedents(wb.handle, 0, 0, 2, 2, &d2.handle), 0);
  EXPECT_EQ(fm_cell_nodes_count(d2.handle), 2U);
}

TEST(FormulonCApiTrace, DepthCapPreventsRunaway) {
  // Single-cell self-loop wouldn't add itself, but 100 should still cap at 32.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  CellNodesGuard n;
  ASSERT_EQ(fm_workbook_precedents(wb.handle, 0, 0, 0, 1000, &n.handle), 0);
  // No precedents (literal) → result empty regardless of depth.
  EXPECT_EQ(fm_cell_nodes_count(n.handle), 0U);
}

TEST(FormulonCApiTrace, OutOfRangeSheetReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_nodes_t* out = nullptr;
  fm_status_t rc = fm_workbook_precedents(wb.handle, 99, 0, 0, 1, &out);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(out, nullptr);
}

TEST(FormulonCApiTrace, NullArgsReturnBindingNullPointer) {
  fm_cell_nodes_t* out = nullptr;
  EXPECT_EQ(fm_workbook_precedents(nullptr, 0, 0, 0, 1, &out),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_dependents(nullptr, 0, 0, 0, 1, &out),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  fm_cell_node_t node{};
  EXPECT_EQ(fm_cell_nodes_at(nullptr, 0, &node),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiTrace, DestroyHandlesNullSafely) {
  fm_cell_nodes_destroy(nullptr);
  EXPECT_EQ(fm_cell_nodes_count(nullptr), 0U);
  SUCCEED();
}
