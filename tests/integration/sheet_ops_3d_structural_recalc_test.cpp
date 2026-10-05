// Structural workbook mutation tests grouped by public surface.
#include "sheet_ops_test_support.h"

namespace formulon {
namespace {
using namespace sheet_ops_test;

TEST(WorkbookRowColEdits, RowDeletesPreserveThreeDReferenceCoordinates) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "First")));
  wb.add_sheet("Second");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A1:A3)")));

  // Excel 365 16.111.3 keeps the shared 3-D coordinate tail unchanged when
  // a sheet inside the span is edited.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
  ASSERT_NE(wb.sheet(2).cell_at(9, 0), nullptr);
  EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->formula_text, "=SUM(First:Second!A1:A3)");

  // Editing the formula owner moves the formula cell, but still does not
  // rewrite the 3-D tail.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(2, /*row=*/0, /*count=*/1)));
  ASSERT_NE(wb.sheet(2).cell_at(8, 0), nullptr);
  EXPECT_EQ(wb.sheet(2).cell_at(8, 0)->formula_text, "=SUM(First:Second!A1:A3)");
}

TEST(WorkbookRowColEdits, ThreeDRowEditsDirtySpanMembersOnly) {
  auto make_workbook = [] {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    wb.add_sheet("Outside");
    for (std::uint32_t row = 0; row < 4U; ++row) {
      EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, row, 0, Value::number(static_cast<double>(row + 1U)))));
    }
    for (std::uint32_t row = 0; row < 3U; ++row) {
      EXPECT_TRUE(
          static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
    }
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A1:A3)")));
    EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    return wb;
  };

  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->formula_text, "=SUM(First:Second!A1:A3)");
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
    EXPECT_EQ(stats.value().cells_evaluated, 1U);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(1, /*row=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 56.0);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 63.0);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_rows(1, /*row=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 36.0);
    EXPECT_EQ(stats.value().cells_evaluated, 1U);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_rows(3, /*row=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    EXPECT_EQ(stats.value().cells_evaluated, 0U);
  }
}

TEST(WorkbookRowColEdits, ThreeDDefinedNamesAndNamedLambdasRecalculateAfterRowEdit) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("First");
  wb.add_sheet("Second");
  wb.add_sheet("Summary");
  wb.set_defined_names({
      DefinedName{"Inner", "SUM(First:Second!A1:A3)", -1, false, ""},
      DefinedName{"Outer", "Inner+0", -1, false, ""},
      DefinedName{"F", "LAMBDA(x,Outer+x)", -1, false, ""},
  });
  for (std::uint32_t row = 0; row < 4U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, row, 0, Value::number(static_cast<double>(row + 1U)))));
  }
  for (std::uint32_t row = 0; row < 3U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=Outer")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 10, 0, "=F(2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 11, 0, "=A10+A11")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(wb.sheet(2).cell_at(9, 0)->cached_value.is_number())
      << wb.sheet(2).cell_at(9, 0)->cached_value.debug_to_string();
  ASSERT_TRUE(wb.sheet(2).cell_at(10, 0)->cached_value.is_number())
      << wb.sheet(2).cell_at(10, 0)->cached_value.debug_to_string();
  ASSERT_TRUE(wb.sheet(2).cell_at(11, 0)->cached_value.is_number())
      << wb.sheet(2).cell_at(11, 0)->cached_value.debug_to_string();
  EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
  EXPECT_EQ(wb.sheet(2).cell_at(10, 0)->cached_value.as_number(), 68.0);
  EXPECT_EQ(wb.sheet(2).cell_at(11, 0)->cached_value.as_number(), 134.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
  EXPECT_EQ(wb.sheet(2).cell_at(10, 0)->cached_value.as_number(), 71.0);
  EXPECT_EQ(wb.sheet(2).cell_at(11, 0)->cached_value.as_number(), 140.0);
  EXPECT_EQ(stats.value().cells_evaluated, 3U);
}

TEST(WorkbookRowColEdits, ThreeDReverseSpanWakesMiddleMemberOnly) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("First");
  wb.add_sheet("Middle");
  wb.add_sheet("Last");
  wb.add_sheet("Summary");
  wb.add_sheet("Outside");
  for (std::uint32_t row = 0; row < 4U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, row, 0, Value::number(static_cast<double>(row + 1U)))));
  }
  for (std::uint32_t row = 0; row < 3U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 100U)))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(2, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(3, 9, 0, "=SUM(Last:First!A1:A3)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(3).cell_at(9, 0)->cached_value.as_number(), 666.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(1, /*row=*/0, /*count=*/1)));
  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(wb.sheet(3).cell_at(9, 0)->formula_text, "=SUM(Last:First!A1:A3)");
  EXPECT_EQ(wb.sheet(3).cell_at(9, 0)->cached_value.as_number(), 566.0);
  EXPECT_EQ(stats.value().cells_evaluated, 1U);

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(4, /*row=*/0, /*count=*/1)));
  stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(wb.sheet(3).cell_at(9, 0)->cached_value.as_number(), 566.0);
  EXPECT_EQ(stats.value().cells_evaluated, 0U);
}

TEST(WorkbookRowColEdits, ThreeDOwnerMoveReordersMovedSourceAndDownstream) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("First");
  wb.add_sheet("Second");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 0, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 3, 0, "=1+3")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 5, 0, "=SUM(First:Second!A1:A3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 0, 0, "=First!A6*2")));
  for (std::uint32_t row = 0; row < 3U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(5, 0)->cached_value.as_number(), 66.0);
  ASSERT_EQ(wb.sheet(2).cell_at(0, 0)->cached_value.as_number(), 132.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(4, 0)->formula_text, "=SUM(First:Second!A1:A3)");
  EXPECT_EQ(wb.sheet(0).cell_at(4, 0)->cached_value.as_number(), 69.0);
  EXPECT_EQ(wb.sheet(2).cell_at(0, 0)->formula_text, "=First!A5*2");
  EXPECT_EQ(wb.sheet(2).cell_at(0, 0)->cached_value.as_number(), 138.0);
}

TEST(WorkbookRowColEdits, ThreeDOwnerMoveWithDefinedNameReindexKeepsRegistryCoordinates) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("First");
  wb.add_sheet("Second");
  wb.set_defined_names({DefinedName{"Region", "First!A2:A3", -1, false, ""}});
  for (std::uint32_t row = 0; row < 4U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, row, 0, Value::number(static_cast<double>(row + 1U)))));
  }
  for (std::uint32_t row = 0; row < 3U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
  }
  // The ordinary formula moves immediately below the 3D owner. If a name
  // rewrite rebuilds the registry before the physical move, both entries are
  // left at their old coordinates and the later span edit can dirty this
  // unrelated formula as though it were the owner.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 5, 0, "=SUM(First:Second!A1:A3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 6, 0, "=1+3")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(5, 0)->cached_value.as_number(), 66.0);
  ASSERT_EQ(wb.sheet(0).cell_at(6, 0)->cached_value.as_number(), 4.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
  ASSERT_EQ(wb.sheet(0).cell_at(4, 0)->formula_text, "=SUM(First:Second!A1:A3)");
  ASSERT_EQ(wb.sheet(0).cell_at(5, 0)->formula_text, "=1+3");
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(4, 0)->cached_value.as_number(), 69.0);

  // The owner is on First, while this edit is on the other span member.
  // Exactly the owner should be woken; a stale pre-move registry entry would
  // also wake the unrelated formula now occupying row 5.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(1, /*row=*/0, /*count=*/1)));
  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(wb.sheet(0).cell_at(4, 0)->cached_value.as_number(), 59.0);
  EXPECT_EQ(wb.sheet(0).cell_at(5, 0)->cached_value.as_number(), 4.0);
  EXPECT_EQ(stats.value().cells_evaluated, 1U);
}

TEST(WorkbookRowColEdits, ThreeDOwnerDeletedByEditIsDroppedFromRegistry) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("First");
  wb.add_sheet("Second");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 0, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 3, 0, "=SUM(First:Second!A1:A3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 1, 0, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 2, 0, Value::number(30.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(3, 0)->cached_value.as_number(), 66.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/3, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* deleted = wb.sheet(0).cell_at(3, 0);
  EXPECT_TRUE(deleted == nullptr || deleted->formula_text.empty());
}

TEST(WorkbookRowColEdits, ThreeDColumnEditsDirtySpanMembersOnly) {
  auto make_workbook = [] {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    wb.add_sheet("Outside");
    for (std::uint32_t col = 0; col < 4U; ++col) {
      EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, col, Value::number(static_cast<double>(col + 1U)))));
    }
    for (std::uint32_t col = 0; col < 3U; ++col) {
      EXPECT_TRUE(
          static_cast<bool>(wb.set_cell_value(1, 0, col, Value::number(static_cast<double>((col + 1U) * 10U)))));
    }
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A1:C1)")));
    EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    return wb;
  };

  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->formula_text, "=SUM(First:Second!A1:C1)");
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
    EXPECT_EQ(stats.value().cells_evaluated, 1U);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(1, /*col=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 56.0);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, /*col=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 63.0);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_cols(1, /*col=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 36.0);
    EXPECT_EQ(stats.value().cells_evaluated, 1U);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_cols(3, /*col=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    EXPECT_EQ(stats.value().cells_evaluated, 0U);
  }
}

TEST(WorkbookRowColEdits, ThreeDWholeAxisEditsRecalculate) {
  {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    for (std::uint32_t row = 0; row < 4U; ++row) {
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, row, 0, Value::number(static_cast<double>(row + 1U)))));
    }
    for (std::uint32_t row = 0; row < 3U; ++row) {
      ASSERT_TRUE(
          static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
    }
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A:A)")));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 70.0);

    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
  }
  {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    for (std::uint32_t col = 0; col < 4U; ++col) {
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, col, Value::number(static_cast<double>(col + 1U)))));
    }
    for (std::uint32_t col = 0; col < 3U; ++col) {
      ASSERT_TRUE(
          static_cast<bool>(wb.set_cell_value(1, 0, col, Value::number(static_cast<double>((col + 1U) * 10U)))));
    }
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!1:1)")));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 70.0);

    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
  }
}

TEST(WorkbookRowColEdits, ThreeDFormulaSourcesRecalculateBeforeFiniteAndWholeAxisOwners) {
  {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(2.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 2, 0, "=ROW()")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(10.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 1, 0, Value::number(20.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 2, 0, Value::number(30.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A1:A3)")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 10, 0, "=SUM(First:Second!A:A)")));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    ASSERT_EQ(wb.sheet(2).cell_at(10, 0)->cached_value.as_number(), 66.0);

    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_EQ(wb.sheet(0).cell_at(1, 0)->formula_text, "=ROW()");
    EXPECT_EQ(wb.sheet(0).cell_at(1, 0)->cached_value.as_number(), 2.0);
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 64.0);
    EXPECT_EQ(wb.sheet(2).cell_at(10, 0)->cached_value.as_number(), 64.0);
  }
  {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 1, Value::number(2.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 2, "=COLUMN()")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(10.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 1, Value::number(20.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 2, Value::number(30.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A1:C1)")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 10, 0, "=SUM(First:Second!1:1)")));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    ASSERT_EQ(wb.sheet(2).cell_at(10, 0)->cached_value.as_number(), 66.0);

    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_EQ(wb.sheet(0).cell_at(0, 1)->formula_text, "=COLUMN()");
    EXPECT_EQ(wb.sheet(0).cell_at(0, 1)->cached_value.as_number(), 2.0);
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 64.0);
    EXPECT_EQ(wb.sheet(2).cell_at(10, 0)->cached_value.as_number(), 64.0);
  }
}

TEST(WorkbookSheetOps, RenameUsesAstForThreeDFormulaAndPreservesStringLiteral) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Alpha");
  wb.add_sheet("Beta");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 0, 0, "=SUM(Alpha:Beta!A1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 1, 0, "=\"Alpha!x\"")));
  wb.set_defined_names({DefinedName{"Span", "SUM(Alpha:Beta!A1)", -1, false, ""}});
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(2).cell_at(0, 0)->cached_value.as_number(), 5.0);

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "Alpha New")));
  EXPECT_EQ(wb.sheet(2).cell_at(0, 0)->formula_text, "=SUM('Alpha New:Beta'!A1)");
  EXPECT_EQ(wb.sheet(2).cell_at(1, 0)->formula_text, "=\"Alpha!x\"");
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].formula, "SUM('Alpha New:Beta'!A1)");
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(2).cell_at(0, 0)->cached_value.as_number(), 5.0);
}

TEST(WorkbookSheetOps, RenameQuotesApostropheInThreeDEndpoint) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("2026 Q1");
  wb.add_sheet("Feb");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 0, 0, "=SUM('2026 Q1:Feb'!A1)")));
  wb.set_defined_names({DefinedName{"Span", "SUM('2026 Q1:Feb'!A1)", -1, false, ""}});
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "2026 Q1's")));
  EXPECT_EQ(wb.sheet(2).cell_at(0, 0)->formula_text, "=SUM('2026 Q1''s:Feb'!A1)");
  EXPECT_EQ(wb.defined_names()[0].formula, "SUM('2026 Q1''s:Feb'!A1)");
}

TEST(WorkbookSheetOps, RemoveRewritesOnlyAffectedAstSubtreesAndDoesNotRebind) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Drop");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(42.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(1, 2, 0, "Drop!x")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 0, 0, "=IFERROR(Drop!A1,99)+5")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 3, 0, "=IF(A3=\"Drop!x\",1,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 4, 0, "=Drop!A1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));
  ASSERT_EQ(wb.sheet_count(), 1U);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->formula_text, "=IFERROR(#REF!,99)+5");
  EXPECT_EQ(wb.sheet(0).cell_at(3, 0)->formula_text, "=IF(A3=\"Drop!x\",1,2)");
  EXPECT_EQ(wb.sheet(0).cell_at(4, 0)->formula_text, "=#REF!+1");
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 104.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(3, 0)->cached_value.as_number(), 1.0);

  // A later sheet with the deleted display name must not rebind the frozen
  // subtree or the direct reference after the final graph rebuild.
  wb.add_sheet("Drop");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(99.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_TRUE(wb.sheet(0).cell_at(4, 0)->cached_value.is_error());
  EXPECT_EQ(wb.sheet(0).cell_at(4, 0)->cached_value.as_error(), ErrorCode::Ref);
}

TEST(WorkbookSheetOps, RemoveAdjustsThreeDBeginMiddleEndAndReverseSpans) {
  const auto run = [](std::uint32_t removed_index, std::string expected_formula, double expected_value,
                      std::string source_formula) {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("Jan");
    wb.add_sheet("Feb");
    wb.add_sheet("Mar");
    wb.add_sheet("Summary");
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(2.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(2, 0, 0, Value::number(3.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(3, 0, 0, std::move(source_formula))));
    ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(removed_index)));
    const std::size_t summary = wb.sheet_index_by_name("Summary");
    ASSERT_NE(summary, static_cast<std::size_t>(-1));
    EXPECT_EQ(wb.sheet(summary).cell_at(0, 0)->formula_text, expected_formula);
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_DOUBLE_EQ(wb.sheet(summary).cell_at(0, 0)->cached_value.as_number(), expected_value);
  };

  run(0, "=SUM(Feb:Mar!A1)", 5.0, "=SUM(Jan:Mar!A1)");
  run(1, "=SUM(Jan:Mar!A1)", 4.0, "=SUM(Jan:Mar!A1)");
  run(2, "=SUM(Jan:Feb!A1)", 3.0, "=SUM(Jan:Mar!A1)");
  run(2, "=SUM(Feb:Jan!A1)", 3.0, "=SUM(Mar:Jan!A1)");
}

TEST(WorkbookSheetOps, RemoveTransformsAllFormulaMetadataAndDropsDanglingPivotState) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Drop");
  wb.add_sheet("Survivor");

  cf::ConditionalFormat conditional_format;
  cf::CFRule rule;
  rule.formula1 = "Drop!A1";
  rule.formula2 = "Drop!B1";
  rule.color_scale = cf::ColorScaleSpec{};
  rule.color_scale->thresholds.push_back(cf::CfValueObject{cf::CfvoType::Formula, "Drop!C1", true});
  rule.icon_set = cf::IconSetSpec{};
  rule.icon_set->thresholds.push_back(cf::CfValueObject{cf::CfvoType::Formula, "Drop!D1", true});
  rule.data_bar = cf::DataBarSpec{};
  rule.data_bar->min = cf::CfValueObject{cf::CfvoType::Formula, "Drop!E1", true};
  rule.data_bar->max = cf::CfValueObject{cf::CfvoType::Formula, "Drop!F1", true};
  conditional_format.rules.push_back(std::move(rule));
  wb.sheet(1).mutable_conditional_formats().push_back(std::move(conditional_format));

  Hyperlink hyperlink;
  hyperlink.location = "#Drop!A1";
  wb.sheet(1).mutable_hyperlinks().push_back(hyperlink);
  DataValidation validation;
  validation.formula1 = "Drop!A1";
  validation.formula2 = "Drop!B1";
  wb.sheet(1).mutable_validations().push_back(validation);

  TableMetadata dropped_table;
  dropped_table.id = 1;
  dropped_table.name = "DropTable";
  dropped_table.display_name = "DropTable";
  dropped_table.ref = "A1:B2";
  dropped_table.sheet_index = 0;
  TableMetadata survivor_table;
  survivor_table.id = 2;
  survivor_table.name = "SurvivorTable";
  survivor_table.display_name = "SurvivorTable";
  survivor_table.ref = "A1:B2";
  survivor_table.sheet_index = 1;
  survivor_table.columns.push_back(TableColumn{1, "Value", {}, {}, "Drop!A1"});
  wb.set_tables({std::move(dropped_table), std::move(survivor_table)});

  wb.set_defined_names({DefinedName{"DropName", "Drop!A1", -1, false, ""}});
  auto dropped_cache = std::make_unique<pivot::PivotCache>();
  dropped_cache->set_cache_id(11);
  dropped_cache->mutable_worksheet_source() = {true, "$A$1:$B$2", "Drop", ""};
  wb.add_pivot_cache(std::move(dropped_cache));
  auto survivor_cache = std::make_unique<pivot::PivotCache>();
  survivor_cache->set_cache_id(12);
  survivor_cache->mutable_worksheet_source() = {true, "$A$1:$B$2", "Survivor", ""};
  wb.add_pivot_cache(std::move(survivor_cache));
  auto dropped_pivot = std::make_unique<pivot::PivotTable>();
  dropped_pivot->set_pivot_cache_id(11);
  wb.sheet(1).add_pivot_table(std::move(dropped_pivot));
  auto survivor_pivot = std::make_unique<pivot::PivotTable>();
  survivor_pivot->set_pivot_cache_id(12);
  wb.sheet(1).add_pivot_table(std::move(survivor_pivot));

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));
  ASSERT_EQ(wb.sheet_count(), 1U);
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].formula1.value(), "#REF!");
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].formula2.value(), "#REF!");
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].color_scale->thresholds[0].value, "#REF!");
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].icon_set->thresholds[0].value, "#REF!");
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].data_bar->min.value, "#REF!");
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].data_bar->max.value, "#REF!");
  EXPECT_EQ(wb.sheet(0).hyperlinks()[0].location, "#REF!");
  EXPECT_EQ(wb.sheet(0).validations()[0].formula1, "#REF!");
  EXPECT_EQ(wb.sheet(0).validations()[0].formula2, "#REF!");
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables()[0].sheet_index, 0U);
  EXPECT_EQ(wb.tables()[0].columns[0].calculated_column_formula, "#REF!");
  EXPECT_EQ(wb.defined_names()[0].formula, "#REF!");
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  EXPECT_EQ(wb.pivot_caches()[0]->cache_id(), 12U);
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);
  EXPECT_EQ(wb.sheet(0).pivot_tables()[0]->pivot_cache_id(), 12U);

  // Re-adding the removed display name must not rebind any surviving holder
  // or resurrect the dropped cache/pivot relationship.
  wb.add_sheet("Drop");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(99.0))));
  EXPECT_EQ(wb.tables()[0].columns[0].calculated_column_formula, "#REF!");
  EXPECT_EQ(wb.defined_names()[0].formula, "#REF!");
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  EXPECT_EQ(wb.pivot_caches()[0]->cache_id(), 12U);
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);
  EXPECT_EQ(wb.sheet(0).pivot_tables()[0]->pivot_cache_id(), 12U);
}

TEST(WorkbookSheetOps, RenameDoesNotClearUnrelatedCommittedSpill) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Alpha");
  wb.add_sheet("Beta");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 0, 0, "=SEQUENCE(2,1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 1, "=Beta!A1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(1).resolve_cell_value(1, 0).as_number(), 2.0);

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "Alpha New")));
  EXPECT_DOUBLE_EQ(wb.sheet(1).resolve_cell_value(1, 0).as_number(), 2.0);
}

TEST(WorkbookSheetOps, RemoveReindexesSurvivingScopedNameBeforeDependentRecalc) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Drop");
  wb.add_sheet("Survivor");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::text("Label"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 1, Value::text("Value"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 1, 0, Value::text("row-1"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 1, 1, Value::number(2.0))));
  TableMetadata table;
  table.id = 7U;
  table.name = "SurvivorTable";
  table.display_name = "SurvivorTable";
  table.ref = "A1:B2";
  table.sheet_index = 1U;
  table.header_row = true;
  table.columns = {TableColumn{1U, "Label", {}, {}, {}}, TableColumn{2U, "Value", {}, {}, {}}};
  wb.set_tables({table});
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Local", "$B$2", 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 0, 2, "=Local")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 0, 0, "=SUM(SurvivorTable[Value])")));

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].local_sheet_id, 0);
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables()[0].sheet_index, 0U);
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(0, 2)->cached_value.as_number(), 2.0);
  EXPECT_DOUBLE_EQ(wb.sheet(1).cell_at(0, 0)->cached_value.as_number(), 2.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 1, Value::number(7.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(0, 2)->cached_value.as_number(), 7.0);
  EXPECT_DOUBLE_EQ(wb.sheet(1).cell_at(0, 0)->cached_value.as_number(), 7.0);
}

}  // namespace
}  // namespace formulon
