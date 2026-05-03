// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// End-to-end regression test for the recalc engine's Text-payload lifetime
// across `Arena::reset()` cycles. The recalc engine resets its per-pass
// `Arena` between cells; without `Sheet::set_cell_cached_value` deep-copying
// Text payloads into per-cell `Cell::cached_text_owned` storage, every Text
// scalar produced by an earlier cell would dangle by the time a subsequent
// cell (or downstream consumer) reads it back.
//
// The scenario chains three formula cells whose results are all Text:
//
//   A1: =CONCAT("foo","bar")  -> "foobar"
//   A2: =UPPER("hello")        -> "HELLO"
//   A3: =A1 & A2               -> "foobarHELLO"
//
// A3 reads A1's and A2's cached Text values during its own evaluation, so
// the bug shows up immediately on the first recalc: A3 either reads
// corrupted bytes or trips a sanitizer use-after-free. The test then
// triggers a second recalc with A1 re-marked dirty to exercise the worst
// case (a fresh arena cycle reuses the bytes that the previous pass's Text
// scalars referenced) and re-asserts every cell.

#include <cstdint>
#include <string>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

TEST(RecalcTextLifetime, RecalcTextResultSurvivesArenaReset) {
  Workbook wb = Workbook::create();

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=CONCAT(\"foo\",\"bar\")")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=UPPER(\"hello\")")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 0U, "=A1 & A2")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  // First pass: every cell must hold its expected Text result. Reading via
  // `cell_at` exercises the post-recalc view; under the bug the bytes
  // would be reused arena memory and either mismatch or trigger a
  // sanitizer report.
  const Sheet& sheet = wb.sheet(0U);

  const Cell* a1 = sheet.cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text())
      << "A1 expected Text; got error="
      << (a1->cached_value.is_error() ? static_cast<int>(a1->cached_value.as_error()) : -1);
  EXPECT_EQ(a1->cached_value.as_text(), "foobar");

  const Cell* a2 = sheet.cell_at(1U, 0U);
  ASSERT_NE(a2, nullptr);
  ASSERT_TRUE(a2->cached_value.is_text())
      << "A2 expected Text; got error="
      << (a2->cached_value.is_error() ? static_cast<int>(a2->cached_value.as_error()) : -1);
  EXPECT_EQ(a2->cached_value.as_text(), "HELLO");

  const Cell* a3 = sheet.cell_at(2U, 0U);
  ASSERT_NE(a3, nullptr);
  ASSERT_TRUE(a3->cached_value.is_text())
      << "A3 expected Text; got error="
      << (a3->cached_value.is_error() ? static_cast<int>(a3->cached_value.as_error()) : -1);
  EXPECT_EQ(a3->cached_value.as_text(), "foobarHELLO");

  // Second pass: re-mark A1 by re-installing its formula (this routes
  // through `set_cell_formula` and marks A1 dirty in the engine). The dep
  // graph propagates dirtiness to A3 so all three cells re-evaluate. A
  // second arena cycle reuses the bytes the previous pass's Text scalars
  // referenced, so any lingering pointer into arena memory would surface
  // as a corrupt read on this assertion.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=CONCAT(\"foo\",\"bar\")")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Cell* a1b = sheet.cell_at(0U, 0U);
  ASSERT_NE(a1b, nullptr);
  ASSERT_TRUE(a1b->cached_value.is_text());
  EXPECT_EQ(a1b->cached_value.as_text(), "foobar");

  const Cell* a2b = sheet.cell_at(1U, 0U);
  ASSERT_NE(a2b, nullptr);
  ASSERT_TRUE(a2b->cached_value.is_text());
  EXPECT_EQ(a2b->cached_value.as_text(), "HELLO");

  const Cell* a3b = sheet.cell_at(2U, 0U);
  ASSERT_NE(a3b, nullptr);
  ASSERT_TRUE(a3b->cached_value.is_text());
  EXPECT_EQ(a3b->cached_value.as_text(), "foobarHELLO");
}

}  // namespace
}  // namespace formulon
