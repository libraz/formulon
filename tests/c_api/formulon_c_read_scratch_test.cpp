//
// Regression test for the per-handle read-path text storage. A long-lived
// handle that loops over text reads must not accumulate one scratch entry
// per call: read-path strings live in `read_scratch`, which is cleared after
// argument/model validation by each successful scratch-backed producer so it
// only ever holds the most recent successful output. Validation-rejected
// calls leave it untouched. This test reaches into the TU-public handle struct
// (`src/c_api/parts/common.h`) to assert the bound directly.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/lambda_value.h"
#include "gtest/gtest.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

}  // namespace

// Looping `fm_workbook_get_value` over a text cell must keep the per-handle
// read scratch bounded (one entry, not one-per-iteration).
TEST(FormulonCApiReadScratch, GetValueDoesNotGrowScratch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "hello"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  constexpr int kIterations = 10000;
  for (int i = 0; i < kIterations; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
    ASSERT_EQ(v.kind, FM_VAL_TEXT);
    ASSERT_NE(v.u.text, nullptr);
    EXPECT_STREQ(v.u.text, "hello");
  }

  // The scratch must hold at most this call's output, regardless of how
  // many reads ran. Without the per-call reset it would hold `kIterations`.
  EXPECT_LE(wb.handle->read_scratch.size(), 1U);
}

// `fm_workbook_lambda_text_at` shares the same read scratch. A validation-
// rejected call on a non-lambda cell must leave the previous successful text
// pointer, value, and scratch size intact.
TEST(FormulonCApiReadScratch, MixedReadsStayBounded) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "alpha"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 0, 42.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  for (int i = 0; i < 5000; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
    fm_value_t n{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 1, 0, &n), 0);
    EXPECT_EQ(n.kind, FM_VAL_NUMBER);
  }
  EXPECT_LE(wb.handle->read_scratch.size(), 1U);
}

TEST(FormulonCApiReadScratch, RejectedLambdaReadPreservesPreviousScratch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "alpha"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_TEXT);
  ASSERT_NE(value.u.text, nullptr);
  const char* previous_pointer = value.u.text;
  const std::string previous_value = previous_pointer;
  const std::size_t previous_size = wb.handle->read_scratch.size();

  const char* rejected_output = nullptr;
  EXPECT_EQ(fm_workbook_lambda_text_at(wb.handle, 0, 0, 0, &rejected_output),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(rejected_output, nullptr);
  EXPECT_EQ(wb.handle->read_scratch.size(), previous_size);
  EXPECT_EQ(value.u.text, previous_pointer);
  EXPECT_STREQ(previous_pointer, previous_value.c_str());
}

// Alternating between two text cells must reuse the same bounded scratch
// store. A caller-owned copy remains valid after the subsequent read replaces
// the handle's transient view.
TEST(FormulonCApiReadScratch, AlternatingReadsStayBoundedAndCopiedTextSurvives) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "first"), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 1, "second"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  std::string retained;
  for (int i = 0; i < 5000; ++i) {
    fm_value_t first{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &first), 0);
    ASSERT_EQ(first.kind, FM_VAL_TEXT);
    ASSERT_NE(first.u.text, nullptr);
    if (i == 0) {
      retained = first.u.text;  // Explicit caller copy before the next read.
    }

    fm_value_t second{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 1, &second), 0);
    ASSERT_EQ(second.kind, FM_VAL_TEXT);
    ASSERT_NE(second.u.text, nullptr);
    EXPECT_STREQ(second.u.text, "second");
  }
  EXPECT_EQ(retained, "first");
  EXPECT_LE(wb.handle->read_scratch.size(), 1U);
}

// Model-backed views have their own lifetime contract: an unrelated
// read-scratch refresh must not alter a previously returned sheet name.
TEST(FormulonCApiReadScratch, ModelBackedSheetNameSurvivesScratchRead) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "payload"), 0);

  const char* sheet_name = nullptr;
  ASSERT_EQ(fm_workbook_sheet_name(wb.handle, 0, &sheet_name), 0);
  ASSERT_NE(sheet_name, nullptr);
  const std::string expected_name = sheet_name;

  fm_value_t value{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &value), 0);
  ASSERT_EQ(value.kind, FM_VAL_TEXT);
  ASSERT_NE(value.u.text, nullptr);
  EXPECT_STREQ(value.u.text, "payload");
  EXPECT_EQ(std::string(sheet_name), expected_name);
}

// Every producer that returns a pointer into per-handle scratch must refresh
// its own bounded store after validation. Exercise the producer families in
// an alternating sequence, but deliberately do not dereference any pointer
// after another producer runs: callers must make their own copy when they
// need retention.
TEST(FormulonCApiReadScratch, AllScratchProducersStayBoundedWhenAlternated) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "cell"), 0);
  ASSERT_EQ(fm_workbook_set_cell_phonetic(wb.handle, 0, 0, 0, "kana"), 0);
  ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, "<autoFilter ref=\"A1:C10\"/>"), 0);

  // Build a lambda value directly in an arena so the lambda text producer
  // succeeds without relying on the top-level `=LAMBDA(...) -> #CALC!`
  // surface rule. The arena outlives every read in this test.
  formulon::Arena lambda_arena;
  formulon::parser::Parser lambda_parser("LAMBDA(x, x*2)", lambda_arena);
  formulon::parser::AstNode* lambda_node = lambda_parser.parse();
  ASSERT_NE(lambda_node, nullptr);
  ASSERT_EQ(lambda_node->kind(), formulon::parser::NodeKind::Lambda);
  auto* lambda = lambda_arena.create<formulon::eval::LambdaValue>();
  ASSERT_NE(lambda, nullptr);
  const std::uint32_t lambda_param_count = lambda_node->as_lambda_param_count();
  auto* lambda_params = lambda_arena.create_array<std::string_view>(lambda_param_count);
  ASSERT_NE(lambda_params, nullptr);
  for (std::uint32_t i = 0; i < lambda_param_count; ++i) {
    lambda_params[i] = lambda_node->as_lambda_param(i);
  }
  lambda->params = lambda_params;
  lambda->param_count = lambda_param_count;
  lambda->optional_count = lambda_node->as_lambda_optional_count();
  lambda->body = &lambda_node->as_lambda_body();
  lambda->captured_env = nullptr;
  wb.handle->workbook().sheet(0).set_cell_cached_value(0, 1, formulon::Value::lambda(lambda));

  fm_cf_cell_range_t cf_range{0, 0, 0, 0};
  fm_cfvo_t thresholds[3]{};
  thresholds[0].type = 3;  // Min.
  thresholds[1].type = 1;  // Percent.
  thresholds[1].value = "50";
  thresholds[2].type = 4;  // Max.
  fm_cf_color_t colors[3]{{255, 0, 0, 255}, {255, 255, 0, 255}, {0, 255, 0, 255}};
  fm_cf_rule_t cf_input{};
  cf_input.type = 2;  // ColorScale.
  cf_input.sqref = &cf_range;
  cf_input.sqref_count = 1;
  cf_input.color_scale_thresholds = thresholds;
  cf_input.color_scale_colors = colors;
  cf_input.color_scale_count = 3;
  std::size_t cf_index = 0;
  ASSERT_EQ(fm_sheet_cf_add_rule(wb.handle, 0, cf_input, &cf_index), 0);
  ASSERT_EQ(cf_index, 0U);

  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  constexpr int kIterations = 500;
  for (int i = 0; i < kIterations; ++i) {
    fm_value_t cell_value{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &cell_value), 0);

    const char* phonetic = nullptr;
    ASSERT_EQ(fm_workbook_get_cell_phonetic(wb.handle, 0, 0, 0, &phonetic), 0);

    const char* lambda_text = nullptr;
    ASSERT_EQ(fm_workbook_lambda_text_at(wb.handle, 0, 0, 1, &lambda_text), 0);

    fm_value_t adhoc_value{};
    ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=\"adhoc\"", &adhoc_value), 0);

    uint32_t rows = 0;
    uint32_t cols = 0;
    ASSERT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 0, "=\"array\"", &rows, &cols), 0);
    ASSERT_EQ(rows, 1U);
    ASSERT_EQ(cols, 1U);
    fm_value_t array_value{};
    ASSERT_EQ(fm_workbook_evaluate_formula_array_cell(wb.handle, 0, &array_value), 0);

    const char* auto_filter = nullptr;
    ASSERT_EQ(fm_sheet_get_auto_filter_xml(wb.handle, 0, &auto_filter), 0);

    fm_cf_rule_t cf_value{};
    ASSERT_EQ(fm_sheet_cf_get_at(wb.handle, 0, 0, &cf_value), 0);

    EXPECT_LE(wb.handle->read_scratch.size(), 3U);
    EXPECT_LE(wb.handle->cfvo_scratch.size(), 3U);
    EXPECT_LE(wb.handle->cf_color_scratch.size(), 3U);
  }
}

// Text writes own their current bytes in the cell rather than retaining every
// historical buffer in a handle-global store. Repeated overwrites must leave
// the workbook-level shared-string storage untouched and preserve the final
// text through later reads.
TEST(FormulonCApiReadScratch, RepeatedTextOverwritesDoNotRetainHistoricalBuffers) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  constexpr int kWrites = 1000;
  for (int i = 0; i < kWrites; ++i) {
    const std::string text = "value-" + std::to_string(i);
    ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, text.c_str()), 0);
  }
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  for (int i = 0; i < 1000; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  }

  EXPECT_TRUE(wb.handle->workbook().text_storage().empty());
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_TEXT);
  EXPECT_STREQ(v.u.text, "value-999");
}
