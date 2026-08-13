//
// ThreadSanitizer-targeted tests for `RecalcEngine`'s public-API thread
// safety. The engine's internal mutex is supposed to serialise every
// mutating entry (`register_formula`, `unregister_formula`,
// `clear_cell_dependencies`, `mark_dirty`, the `recalc*` family) against
// every other mutating entry, even when callers race them from disjoint
// threads.
//
// These tests do NOT exercise the documented "no concurrent recalc on
// the same workbook" invariant — that contract is the caller's
// responsibility — but they DO cover the lower-level guarantee that the
// engine's bookkeeping survives an adversarial caller that races
// `register_formula` / `mark_dirty` from multiple threads. The lock
// keeps the dep graph and dirty set internally consistent so a future
// API surface that exposes parallel mutation paths cannot regress
// silently.
//
// Detached threads are forbidden by project policy; every thread is
// joined.

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <thread>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/tables_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Helper: a workbook with N rows in column A laid out as
//   A1 = 1, A2 = A1+1, ..., AN = A(N-1)+1.
// Sized small enough to keep the TSan event budget low while still
// producing a non-trivial dep graph with O(N) edges.
constexpr std::uint32_t kChainRows = 32U;

void SeedChain(Workbook& wb) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  for (std::uint32_t r = 1U; r < kChainRows; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, r, 0U, "=A" + std::to_string(r) + "+1")));
  }
}

Value OwnedPivotText(pivot::PivotCache& cache, std::string text) {
  cache.mutable_text_storage().push_back(std::move(text));
  return Value::text(cache.text_storage().back());
}

std::unique_ptr<pivot::PivotCache> BuildRacePivotCache(std::uint32_t cache_id, std::string sheet_name) {
  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(cache_id);

  pivot::PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(OwnedPivotText(*cache, "North"));
  region.shared_items.push_back(OwnedPivotText(*cache, "South"));
  cache->mutable_fields().push_back(std::move(region));

  pivot::PivotCacheField amount;
  amount.name = "Amount";
  cache->mutable_fields().push_back(std::move(amount));
  for (const auto& record_values : {std::pair{0.0, 100.0}, std::pair{0.0, 200.0}, std::pair{1.0, 300.0}}) {
    pivot::PivotCacheRecord record;
    record.cells.push_back(Value::number(record_values.first));
    record.cells.push_back(Value::number(record_values.second));
    cache->mutable_records().push_back(std::move(record));
  }
  cache->mutable_worksheet_source() = {true, "$A$1:$B$2", std::move(sheet_name), ""};
  return cache;
}

std::unique_ptr<pivot::PivotTable> BuildRacePivotTable(std::uint32_t cache_id) {
  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("RacePivot");
  table->set_pivot_cache_id(cache_id);

  pivot::PivotField region;
  region.source_name = "Region";
  region.axis = pivot::PivotAxis::Row;
  table->mutable_fields().push_back(std::move(region));
  pivot::PivotField amount;
  amount.source_name = "Amount";
  amount.axis = pivot::PivotAxis::Value;
  table->mutable_fields().push_back(std::move(amount));

  pivot::PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1U;
  sum.aggregation = pivot::Aggregation::Sum;
  table->mutable_data_fields().push_back(std::move(sum));
  table->mutable_row_field_order().push_back(0U);
  table->set_anchor(2U, 0U, 5U, 2U);  // A3:B7
  return table;
}

// ---------------------------------------------------------------------------
// Concurrent register_formula / mark_dirty must not crash or corrupt state.
// ---------------------------------------------------------------------------

TEST(RecalcEngineThreadSafety, ConcurrentMutatorsRaceFreeUnderLock) {
  // Two writer threads hammer the workbook's mutating recalc-engine
  // entry points in parallel: one keeps re-issuing `set_cell_formula`
  // (which routes through `register_formula` + `mark_dirty`), the other
  // keeps re-issuing `set_cell_value` on a sibling cell (which routes
  // through `mark_dirty` for every dependent). With the engine mutex in
  // place these calls serialise cleanly; without it the dep graph and
  // dirty set would race and TSan would flag the unsynchronised access.
  Workbook wb = Workbook::create();
  SeedChain(wb);

  std::atomic<bool> stop{false};
  std::atomic<int> failures{0};

  // Writer 1: rewrites A2's formula in a tight loop. Each rewrite drops
  // the cell's previous outgoing edges and adds the new one, exercising
  // both `unregister_formula_locked` and `register_formula_locked`.
  std::thread writer_a([&] {
    for (int i = 0; !stop.load(std::memory_order_acquire) && i < 200; ++i) {
      const std::string formula = (i % 2 == 0) ? "=A1+1" : "=A1+2";
      if (!static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, formula))) {
        ++failures;
        return;
      }
    }
  });

  // Writer 2: re-stamps the seed cell so every dependent gets dirtied
  // through the BFS in `set_cell_value`. The volatile / dirty bookkeeping
  // is the most TSan-visible mutable state during the race.
  std::thread writer_b([&] {
    for (int i = 0; !stop.load(std::memory_order_acquire) && i < 200; ++i) {
      if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(i + 1))))) {
        ++failures;
        return;
      }
    }
  });

  writer_a.join();
  writer_b.join();
  stop.store(true, std::memory_order_release);
  EXPECT_EQ(failures.load(), 0);

  // Drive a final recalc and confirm the engine state is still healthy.
  // We do not assert specific values because the writers raced — only
  // that the post-race recalc completes successfully.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
}

// ---------------------------------------------------------------------------
// Concurrent recalc + read-only dep_graph() observer.
// ---------------------------------------------------------------------------

TEST(RecalcEngineThreadSafety, ParallelRecalcSerialisesAgainstReaderObserver) {
  // While `recalc_parallel` is in flight the engine mutex is held for
  // the entire pass. A separate observer thread that calls into the
  // mutating API (`mark_dirty` on any of the chain's cells) should
  // block until the recalc returns; once it does, the dep graph state
  // remains coherent and a subsequent recalc still succeeds.
  Workbook wb = Workbook::create();
  SeedChain(wb);

  // Initial recalc to commit values.
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));

  // Re-dirty the chain by mutating the seed.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(2.0))));

  std::atomic<int> failures{0};
  std::atomic<bool> observer_done{false};

  // Observer thread: races a `set_cell_value` against the parallel
  // recalc on the main thread. The lock makes this a happens-before
  // edge instead of a race; either the observer's mutation lands
  // before the recalc's lock acquisition (and the recalc picks it up)
  // or after (and a follow-up recalc would). Either ordering must be
  // race-free under TSan.
  std::thread observer([&] {
    if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::number(99.0)))) {
      ++failures;
    }
    observer_done.store(true, std::memory_order_release);
  });

  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  observer.join();
  EXPECT_TRUE(observer_done.load());
  EXPECT_EQ(failures.load(), 0);

  // Final recalc must succeed regardless of the observer's interleaving.
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
}

// A stable pair of SEQUENCE spills gives the whole-axis readers committed
// phantom coordinates to inspect, while the independent producers below are
// re-registered before every pass so their workers commit fresh spill regions
// on the same Sheet. The aggregate SCCs do not depend on those producers, so
// the scheduler can place both kinds of work in one parallel layer. This is
// deliberately a real scheduler workload rather than a concurrent caller
// racing two recalc_parallel invocations (which the public contract forbids).
TEST(RecalcEngineThreadSafety, ParallelWholeAxisReadsWhileSpillsCommit) {
  constexpr std::uint32_t kProducerCount = 16U;
  constexpr std::uint32_t kProducerRow = 9U;  // Excel row 10.
  constexpr std::uint32_t kAggregateCount = 16U;
  constexpr int kPasses = 12;

  Workbook wb = Workbook::create();
  // A1:D1 supplies row-1 phantom values 1..4, and AA1:AA4 supplies a
  // vertical phantom-only tail for the whole-column aggregate.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(1,4)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 26U, "=SEQUENCE(4,1)")));  // AA1

  // Keep the aggregate formulas away from row 1: SUM(1:1) must observe only
  // the stable initial spill, not include its own formula cell. They also do
  // not cover the producer row, so they are independent scheduler tasks.
  for (std::uint32_t i = 0U; i < kAggregateCount; ++i) {
    const std::uint32_t row = 1U + i;
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 25U, "=SUM(AA:AA)")));  // Z2:Z17
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 24U, "=SUM(1:1)")));    // Y2:Y17
  }

  const auto producer_col = [](std::uint32_t index) {
    // Leave Y:AA free for the stable spills and aggregate formulas.
    const std::uint32_t raw = 4U + index * 3U;
    return raw >= 24U ? raw + 4U : raw;
  };
  const auto register_producers = [&] {
    for (std::uint32_t i = 0U; i < kProducerCount; ++i) {
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, kProducerRow, producer_col(i), "=SEQUENCE(16,1)")));
    }
  };
  register_producers();

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  SchedulerStats initial_stats;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &initial_stats)));
  ASSERT_GT(initial_stats.parallel_steps, 0U);

  for (int pass = 0; pass < kPasses; ++pass) {
    // Re-registering a producer clears its previous committed spill and marks
    // it dirty. Re-registering the readers makes each pass exercise the same
    // parallel layer, without relying on volatile timing or sleeps.
    register_producers();
    for (std::uint32_t i = 0U; i < kAggregateCount; ++i) {
      const std::uint32_t row = 1U + i;
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 25U, "=SUM(AA:AA)"))) << "pass " << pass;
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 24U, "=SUM(1:1)"))) << "pass " << pass;
    }

    SchedulerStats stats;
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats))) << "pass " << pass;
    ASSERT_GT(stats.parallel_steps, 0U) << "pass " << pass;

    // The stable vertical spill is visible through the shared populated
    // extent, and the stable horizontal spill is visible through the whole
    // row. The assertions are exact across all passes, while the producer
    // spills concurrently mutate the same sheet's spill table.
    EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(1U, 25U).as_number(), 10.0) << "pass " << pass;
    EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(1U, 24U).as_number(), 11.0) << "pass " << pass;
    for (std::uint32_t i = 0U; i < kProducerCount; ++i) {
      const Value tail = wb.sheet(0).resolve_cell_value(kProducerRow + 15U, producer_col(i));
      ASSERT_TRUE(tail.is_number()) << "pass " << pass << " producer " << i;
      EXPECT_DOUBLE_EQ(tail.as_number(), 16.0) << "pass " << pass << " producer " << i;
    }
  }
}

void AssertParallelPhantomOnlyRangeSequence(bool producer_first) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  const auto install_watcher = [&] { ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)"))); };
  const auto install_producer = [&] {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(3,2)")));
  };
  if (producer_first) {
    install_producer();
    install_watcher();
  } else {
    install_watcher();
    install_producer();
  }

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_number(), 12.0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(3,2,10)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_number(), 39.0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=C1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(1U, 0U).as_number(), 1.0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_number(), 0.0);
}

TEST(RecalcEngineThreadSafety, ParallelPhantomOnlyRangeWatcherWatcherFirst) {
  AssertParallelPhantomOnlyRangeSequence(false);
}

TEST(RecalcEngineThreadSafety, ParallelPhantomOnlyRangeWatcherProducerFirst) {
  AssertParallelPhantomOnlyRangeSequence(true);
}

TEST(RecalcEngineThreadSafety, ParallelDefinedNameReindexClearsStaleSpill) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("MAKE", "SEQUENCE(3,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=MAKE")));      // A2
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)")));  // C1

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  SchedulerStats initial;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &initial)));
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_number(), 12.0);

  // Reindexing the defined name must clear the old committed footprint before
  // the parallel SCC pass sees the new scalar dependency C1 -> A2.
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("MAKE", "C1+1")));
  SchedulerStats rewrite;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &rewrite)));
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(1U, 0U).as_number(), 1.0);
  EXPECT_TRUE(wb.sheet(0).resolve_cell_value(1U, 1U).is_blank());

  // A retry with no mutation must preserve the recovered scalar values and
  // must not resurrect the stale phantom-derived cycle.
  SchedulerStats retry;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &retry)));
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(1U, 0U).as_number(), 1.0);
  EXPECT_TRUE(wb.sheet(0).resolve_cell_value(1U, 1U).is_blank());
}

TEST(RecalcEngineThreadSafety, ParallelRecalcSerialisesAgainstRemoveSheetTransaction) {
  constexpr std::size_t kMetadataCopies = 96U;
  constexpr int kRacePasses = 16;
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Drop");
  wb.add_sheet("Survivor");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(42.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::text("Label"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 1U, Value::text("Value"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 1U, 0U, Value::text("row-1"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 1U, 1U, Value::number(1.0))));

  std::vector<io::DefinedName> names;
  names.reserve(kMetadataCopies);
  for (std::size_t i = 0; i < kMetadataCopies; ++i) {
    names.push_back(io::DefinedName{"DropName" + std::to_string(i), "Drop!A1", -1, false, ""});
  }
  wb.set_defined_names(std::move(names));

  io::TableMetadata table;
  table.id = 1U;
  table.name = "DropTable";
  table.display_name = "DropTable";
  table.ref = "A1:B2";
  table.sheet_index = 0U;
  table.columns = {io::TableColumn{1U, "Label", {}, {}, {}}, io::TableColumn{2U, "Value", {}, {}, {}}};
  std::vector<io::TableMetadata> tables;
  tables.reserve(kMetadataCopies + 1U);
  tables.push_back(std::move(table));
  for (std::size_t i = 0; i < kMetadataCopies; ++i) {
    io::TableMetadata survivor_table;
    survivor_table.id = static_cast<std::uint32_t>(100U + i);
    survivor_table.name = "SurvivorTable" + std::to_string(i);
    survivor_table.display_name = survivor_table.name;
    survivor_table.ref = "A1:B2";
    survivor_table.sheet_index = 1U;
    survivor_table.columns = {io::TableColumn{1U, "Label", {}, {}, {}}, io::TableColumn{2U, "Value", {}, {}, {}}};
    tables.push_back(std::move(survivor_table));
  }
  wb.set_tables(std::move(tables));
  for (std::size_t i = 0; i < kMetadataCopies; ++i) {
    ASSERT_TRUE(static_cast<bool>(
        wb.set_cell_formula(1U, static_cast<std::uint32_t>(2U + i), 0U, "=DropName" + std::to_string(i))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, static_cast<std::uint32_t>(2U + i), 1U,
                                                      "=SUM(SurvivorTable" + std::to_string(i) + "[Value])")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, static_cast<std::uint32_t>(2U + i), 2U,
                                                      "=GETPIVOTDATA(\"Sum of Amount\",A3,\"Region\",\"North\")")));
  }

  wb.add_pivot_cache(BuildRacePivotCache(21U, "Drop"));
  for (std::size_t i = 0; i < kMetadataCopies; ++i) {
    wb.add_pivot_cache(BuildRacePivotCache(static_cast<std::uint32_t>(1000U + i), "Survivor"));
  }
  wb.sheet(1).add_pivot_table(BuildRacePivotTable(21U));
  wb.sheet(1).add_pivot_table(BuildRacePivotTable(1000U));

  // Seed the dependency graph after all name/table definitions are present,
  // then dirty both evaluator paths immediately before the race.
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(43.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 1U, 1U, Value::number(2.0))));
  for (std::size_t i = 0; i < kMetadataCopies; ++i) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, static_cast<std::uint32_t>(2U + i), 2U,
                                                      "=GETPIVOTDATA(\"Sum of Amount\",A3,\"Region\",\"North\")")));
  }

  std::atomic<int> ready{0};
  std::atomic<bool> go{false};
  std::atomic<int> failures{0};
  const auto barrier = [&] {
    ready.fetch_add(1, std::memory_order_acq_rel);
    while (!go.load(std::memory_order_acquire)) {
      std::this_thread::yield();
    }
  };

  std::thread recalc_thread([&] {
    barrier();
    for (int pass = 0; pass < kRacePasses; ++pass) {
      if (!static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr))) {
        failures.fetch_add(1, std::memory_order_relaxed);
      }
    }
  });
  std::thread remove_thread([&] {
    barrier();
    if (!static_cast<bool>(wb.remove_sheet(0U))) {
      failures.fetch_add(1, std::memory_order_relaxed);
    }
  });
  while (ready.load(std::memory_order_acquire) != 2) {
    std::this_thread::yield();
  }
  go.store(true, std::memory_order_release);
  recalc_thread.join();
  remove_thread.join();

  EXPECT_EQ(failures.load(std::memory_order_relaxed), 0);
  ASSERT_EQ(wb.sheet_count(), 1U);
  EXPECT_EQ(wb.sheet(0U).name(), "Survivor");
  ASSERT_EQ(wb.tables().size(), kMetadataCopies);
  EXPECT_EQ(wb.tables()[0].sheet_index, 0U);
  ASSERT_EQ(wb.pivot_caches().size(), kMetadataCopies);
  ASSERT_EQ(wb.sheet(0U).pivot_tables().size(), 1U);
  EXPECT_EQ(wb.sheet(0U).pivot_tables()[0]->pivot_cache_id(), 1000U);
  ASSERT_EQ(wb.defined_names().size(), kMetadataCopies);
  EXPECT_EQ(wb.defined_names()[0].formula, "#REF!");
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  for (std::size_t i = 0; i < kMetadataCopies; ++i) {
    EXPECT_TRUE(wb.sheet(0U).cell_at(static_cast<std::uint32_t>(2U + i), 0U)->cached_value.is_error());
    EXPECT_DOUBLE_EQ(wb.sheet(0U).cell_at(static_cast<std::uint32_t>(2U + i), 1U)->cached_value.as_number(), 2.0);
    EXPECT_DOUBLE_EQ(wb.sheet(0U).cell_at(static_cast<std::uint32_t>(2U + i), 2U)->cached_value.as_number(), 300.0);
  }
}

TEST(RecalcEngineThreadSafety, ParallelRecalcSerialisesAgainstDefinedNameUpdate) {
  constexpr int kRacePasses = 128;
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("N", "1", -1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=N+1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(0U).cell_at(0U, 0U)->cached_value.as_number(), 2.0);

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  std::atomic<int> ready{0};
  std::atomic<bool> go{false};
  std::atomic<int> failures{0};
  const auto barrier = [&] {
    ready.fetch_add(1, std::memory_order_acq_rel);
    while (!go.load(std::memory_order_acquire)) {
      std::this_thread::yield();
    }
  };

  std::thread recalc_thread([&] {
    barrier();
    for (int pass = 0; pass < kRacePasses; ++pass) {
      if (!static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr))) {
        failures.fetch_add(1, std::memory_order_relaxed);
      }
    }
  });
  std::thread update_thread([&] {
    barrier();
    for (int pass = 0; pass < kRacePasses; ++pass) {
      const std::string formula = (pass % 2 == 0) ? "2" : "4";
      if (!static_cast<bool>(wb.set_defined_name_scoped("N", formula, -1))) {
        failures.fetch_add(1, std::memory_order_relaxed);
      }
    }
  });
  while (ready.load(std::memory_order_acquire) != 2) {
    std::this_thread::yield();
  }
  go.store(true, std::memory_order_release);
  recalc_thread.join();
  update_thread.join();

  EXPECT_EQ(failures.load(std::memory_order_relaxed), 0);
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("N", "2", -1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  EXPECT_DOUBLE_EQ(wb.sheet(0U).cell_at(0U, 0U)->cached_value.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// Concurrent independent workbooks: per-workbook mutex must not contend.
// ---------------------------------------------------------------------------

TEST(RecalcEngineThreadSafety, IndependentWorkbookEnginesIndependentLocks) {
  // Eight threads each own a workbook and drive a tight
  // mutate-recalc-mutate loop. Per-workbook mutexes mean these never
  // contend; TSan must observe no race even though the total mutation
  // rate is high.
  constexpr int kThreads = 8;
  std::vector<std::thread> workers;
  workers.reserve(kThreads);
  std::atomic<int> failures{0};

  for (int t = 0; t < kThreads; ++t) {
    workers.emplace_back([&failures, t] {
      Workbook wb = Workbook::create();
      if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(t + 1))))) {
        ++failures;
        return;
      }
      if (!static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+A1"))) {
        ++failures;
        return;
      }
      for (int i = 0; i < 25; ++i) {
        if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(i))))) {
          ++failures;
          return;
        }
        if (!static_cast<bool>(wb.recalc(default_registry()))) {
          ++failures;
          return;
        }
      }
    });
  }
  for (std::thread& th : workers) {
    th.join();
  }
  EXPECT_EQ(failures.load(), 0);
}

}  // namespace
}  // namespace formulon::eval
