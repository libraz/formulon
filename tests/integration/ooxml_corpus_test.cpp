// OOXML round-trip parity over a 20-book minimal corpus.
//
// The parameterized fixture and its formatter stay together so every corpus
// id remains a stable gtest name. Builders, invariants, and package helpers
// live in the adjacent topic files.

#include <string>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "ooxml_corpus_support.h"

namespace formulon {
namespace {

using namespace ooxml_corpus;

class RoundTripParity : public ::testing::TestWithParam<CorpusBook> {};

TEST_P(RoundTripParity, TwoCyclePipeline) {
  const CorpusBook& book = GetParam();

  // (1) build source bytes.
  auto src_or = book.build();
  ASSERT_TRUE(static_cast<bool>(src_or)) << "build failed for '" << book.id << "': " << src_or.error().message;

  // (2) read into workbook A.
  auto first_or = io::read_ooxml(SpanOf(src_or.value()));
  ASSERT_TRUE(static_cast<bool>(first_or))
      << "first read_ooxml failed for '" << book.id << "': " << first_or.error().message;
  io::OoxmlReadResult& first = first_or.value();

  // (3) recalc workbook A so cached values are populated. Iterative
  // options round-trip through `<calcPr>` (parser + writer), so the
  // cyclic SCC books no longer need their options re-applied here.
  auto stats_or = first.workbook.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats_or))
      << "first recalc failed for '" << book.id << "': " << stats_or.error().message;

  // Per-book invariants on A.
  book.assert_invariants(first.workbook);

  // (4) write A back to bytes.
  auto bytes_or = first.workbook.save();
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "first save failed for '" << book.id << "': " << bytes_or.error().message;

  // (5) read again into workbook C.
  auto second_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(second_or))
      << "second read_ooxml failed for '" << book.id << "': " << second_or.error().message;
  io::OoxmlReadResult& second = second_or.value();

  // (6) recalc workbook C. Iterative options round-trip via `<calcPr>`
  // and so are already in place from the second read.
  auto stats2_or = second.workbook.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats2_or))
      << "second recalc failed for '" << book.id << "': " << stats2_or.error().message;

  // Re-assert per-book invariants on C — the second read must produce
  // the same shape.
  book.assert_invariants(second.workbook);

  // (7) cross-cycle equality contract.
  EXPECT_TRUE(sheets_have_same_shape(first.workbook, second.workbook)) << "sheet shape drifted for '" << book.id << "'";
  EXPECT_TRUE(defined_names_equal(first.workbook, second.workbook)) << "defined names drifted for '" << book.id << "'";
  EXPECT_TRUE(tables_equal(first.workbook, second.workbook)) << "tables drifted for '" << book.id << "'";
  EXPECT_TRUE(passthrough_parts_equal(first.workbook, second.workbook))
      << "passthrough parts drifted for '" << book.id << "'";
  EXPECT_TRUE(formula_texts_equal_for_each_cell(first.workbook, second.workbook))
      << "formula texts drifted for '" << book.id << "'";
}

/// Custom GTest parameter name printer: each test is named after the
/// `id` field of the corpus book so failures read as
/// `OoxmlCorpus/RoundTripParity.TwoCyclePipeline/<id>`.
struct CorpusNameFormatter {
  std::string operator()(const ::testing::TestParamInfo<CorpusBook>& info) const { return info.param.id; }
};

INSTANTIATE_TEST_SUITE_P(OoxmlCorpus, RoundTripParity, ::testing::ValuesIn(make_corpus()), CorpusNameFormatter());

}  // namespace
}  // namespace formulon
