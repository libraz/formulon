// Consumers outside the evaluator that read a picture cell: PivotTable refresh, charts, table totals labels,
// what-if data tables, conditional formats, data validation and AutoFilter. Each pins the reading the consumer
// had before a cell could hold a picture.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "c_api/formulon_c.h"
#include "cell.h"
#include "cf/auto_filter_eval.h"
#include "cf/cf_evaluator.h"
#include "cf/cf_types.h"
#include "cf/validation_eval.h"
#include "chart/chart_data.h"
#include "chart/chart_model.h"
#include "data_table.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_writer.h"
#include "io/read_options.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "merge_range.h"
#include "miniz.h"
#include "passthrough_part.h"
#include "picture_table.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "table.h"
#include "table_totals_sync.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// Registers one picture and its media part in `wb`, so a save can write it.
Value MakePicture(Workbook& wb) {
  const std::vector<std::uint8_t> png = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0x01, 0x02};
  EXPECT_TRUE(wb.add_passthrough_part(PassthroughPart("xl/media/image1.png", "image/png", png)).has_value());
  PictureTable& table = wb.picture_table();
  const std::uint32_t media =
      table.add_media_at(table.media_count(), MediaKey{png.size(), 0x1234U}, "xl/media/image1.png");
  const std::uint32_t id = table.add_entry(media, std::nullopt);
  return Value::picture(PicturePayload{id, media}, 5U);
}

// ---------------------------------------------------------------------------
// PivotTable refresh
// ---------------------------------------------------------------------------

constexpr std::size_t kDataSheet = 0U;
constexpr std::size_t kPivotSheet = 1U;
constexpr std::uint32_t kCacheId = 1U;
constexpr std::uint32_t kProductField = 1U;

void SetDataText(Workbook& wb, std::uint32_t row, std::uint32_t col, const char* text) {
  ASSERT_TRUE(wb.set_cell_text(kDataSheet, row, col, text).has_value());
}

// Region / Product / Amt over five rows, with a picture in the Product column of row 3 (A3 is `East`), and one
// pivot of Region on a second sheet.
Workbook PivotBook() {
  Workbook wb = Workbook::create_for_profile(mac_365_en_us_profile());
  const char* regions[] = {"North", "South", "East", "North", "South"};
  const char* products[] = {"A", "B", "A", "B", "A"};
  SetDataText(wb, 0U, 0U, "Region");
  SetDataText(wb, 0U, 1U, "Product");
  SetDataText(wb, 0U, 2U, "Amt");
  for (std::uint32_t r = 0; r < 5U; ++r) {
    SetDataText(wb, r + 1U, 0U, regions[r]);
    SetDataText(wb, r + 1U, 1U, products[r]);
    EXPECT_TRUE(wb.set_cell_value(kDataSheet, r + 1U, 2U, Value::number(100.0 * (r + 1U))).has_value());
  }
  EXPECT_TRUE(wb.set_cell_value(kDataSheet, 3U, 1U, MakePicture(wb)).has_value());

  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(kCacheId);
  cache->mutable_fields().push_back(pivot::PivotCacheField{"Region", {}});
  cache->mutable_fields().push_back(pivot::PivotCacheField{"Product", {}});
  cache->mutable_fields().push_back(pivot::PivotCacheField{"Amt", {}});
  pivot::WorksheetSource& source = cache->mutable_worksheet_source();
  source.present = true;
  source.ref = "A1:C6";
  source.sheet = wb.sheet(kDataSheet).name();
  wb.add_pivot_cache(std::move(cache));

  EXPECT_TRUE(static_cast<bool>(wb.add_sheet_entered("Pivot")));
  auto table = std::make_unique<pivot::PivotTable>();
  table->set_name("PT1");
  table->set_pivot_cache_id(kCacheId);
  const char* names[] = {"Region", "Product", "Amt"};
  const pivot::PivotAxis axes[] = {pivot::PivotAxis::Row, pivot::PivotAxis::Row, pivot::PivotAxis::Value};
  for (std::size_t i = 0; i < 3U; ++i) {
    pivot::PivotField field;
    field.source_name = names[i];
    field.axis = axes[i];
    table->mutable_fields().push_back(std::move(field));
  }
  table->mutable_row_field_order() = {0U, 1U};
  pivot::PivotDataField sum;
  sum.name = "Sum of Amt";
  sum.field_index = 2U;
  table->mutable_data_fields().push_back(std::move(sum));
  table->set_anchor(2U, 0U, 1U, 1U);
  // A new pivot's style and defaults, as the C API sets them; the xlsb writer needs both.
  table->set_data_caption("Values");
  pivot::PivotStyleInfo style;
  style.name = "PivotStyleLight16";
  style.show_row_headers = true;
  style.show_col_headers = true;
  style.show_last_column = true;
  table->set_style_info(style);
  table->apply_new_pivot_defaults();
  EXPECT_TRUE(static_cast<bool>(wb.add_pivot(kPivotSheet, std::move(table))));
  EXPECT_TRUE(static_cast<bool>(wb.project_pivot(kPivotSheet, 0U, ReportPlacement::kBlock)));
  return wb;
}

bool HoldsValueError(const std::vector<Value>& values) {
  for (const Value& v : values) {
    if (v.is_error() && v.as_error() == ErrorCode::Value) {
      return true;
    }
  }
  return false;
}

TEST(PictureConsumers, PivotRefreshReadsAPictureSourceCellAsAValueErrorItem) {
  Workbook wb = PivotBook();
  ASSERT_TRUE(wb.refresh_pivot_cache(kCacheId, eval::default_registry()).has_value());
  const pivot::PivotCache* cache = wb.find_pivot_cache(kCacheId);
  ASSERT_NE(cache, nullptr);
  ASSERT_GT(cache->fields().size(), kProductField);
  EXPECT_TRUE(HoldsValueError(cache->fields()[kProductField].shared_items));
}

// The text of the pivotCacheDefinition part of an .xlsx package.
std::string CacheDefinitionXml(const std::vector<std::uint8_t>& package) {
  std::string out;
  mz_zip_archive zip{};
  if (mz_zip_reader_init_mem(&zip, package.data(), package.size(), 0) == MZ_FALSE) {
    return out;
  }
  for (mz_uint i = 0; i < mz_zip_reader_get_num_files(&zip); ++i) {
    char name[256];
    mz_zip_reader_get_filename(&zip, i, name, sizeof name);
    if (std::string_view(name).rfind("xl/pivotCache/pivotCacheDefinition", 0) != 0) {
      continue;
    }
    std::size_t size = 0;
    void* data = mz_zip_reader_extract_to_heap(&zip, i, &size, 0);
    out.assign(static_cast<const char*>(data), size);
    mz_free(data);
  }
  mz_zip_reader_end(&zip);
  return out;
}

// Excel stores the locale text of a picture (`画像` in ja-JP) as the cache item; the engine stores `#VALUE!` until
// the PivotTable follow-up models that caption, which changes this expectation.
TEST(PictureConsumers, PivotCacheOfAPictureSourceSavesAValueErrorItemInXlsx) {
  Workbook wb = PivotBook();
  ASSERT_TRUE(wb.refresh_pivot_cache(kCacheId, eval::default_registry()).has_value());
  const Expected<std::vector<std::uint8_t>, Error> saved = io::write_ooxml(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  const std::string definition = CacheDefinitionXml(saved.value());
  ASSERT_FALSE(definition.empty());
  EXPECT_NE(definition.find("<e v=\"#VALUE!\"/>"), std::string::npos) << definition;
  EXPECT_EQ(definition.find("<m/>"), std::string::npos) << definition;
}

// As above for .xlsb, read back through the reader: the error item survives the container. The pivot is built
// through the C API, which gives it the style and defaults the xlsb writer needs.
TEST(PictureConsumers, PivotCacheOfAPictureSourceSavesAValueErrorItemInXlsb) {
  const std::uint8_t png[] = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I',  'H',  'D',  'R', 0,
                              0,    0,   2,   0,   0,    0,    2,    8,    6, 0, 0, 0,  0x1F, 0x15, 0xC4, 0x89};
  fm_workbook_t* handle = nullptr;
  ASSERT_EQ(fm_workbook_create_default(&handle), 0);
  const char* const products[] = {"Pen", nullptr, "Ink", "Pen"};
  ASSERT_EQ(fm_sheet_set_text(handle, 0U, 0U, 0U, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(handle, 0U, 0U, 1U, "Amount"), 0);
  for (std::uint32_t r = 0; r < 4U; ++r) {
    if (products[r] != nullptr) {
      ASSERT_EQ(fm_sheet_set_text(handle, 0U, r + 1U, 0U, products[r]), 0);
    } else {
      fm_cell_picture_t picture = FM_STRUCT_INIT(fm_cell_picture_t);
      picture.bytes = png;
      picture.len = sizeof(png);
      ASSERT_EQ(fm_sheet_set_cell_picture(handle, 0U, r + 1U, 0U, &picture), 0) << fm_get_last_error_message();
    }
    ASSERT_EQ(fm_sheet_set_number(handle, 0U, r + 1U, 1U, 10.0 * (r + 1U)), 0);
  }
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(handle, 0U, &cache_id), 0) << fm_get_last_error_message();
  fm_pivot_worksheet_source_t source{sizeof(fm_pivot_worksheet_source_t), "A1:B5", "Sheet1", nullptr};
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(handle, cache_id, &source), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_refresh_pivot_cache(handle, cache_id), 0) << fm_get_last_error_message();
  fm_pivot_spec_t spec{sizeof(fm_pivot_spec_t), "PT", cache_id, 0U, 3U};
  std::size_t pivot = 0;
  ASSERT_EQ(fm_sheet_add_pivot(handle, 0U, &spec, &pivot), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(handle, 0U, pivot, 0U, FM_PIVOT_AXIS_ROW), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t data{};
  data.struct_size = sizeof(fm_pivot_data_field_t);
  data.field_index = 1U;
  data.aggregation = FM_PIVOT_AGGREGATION_SUM;
  std::size_t data_index = 0;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(handle, 0U, pivot, &data, &data_index), 0) << fm_get_last_error_message();
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  const fm_status_t status = fm_workbook_save(handle, FM_WORKBOOK_FORMAT_XLSB, &bytes, &len);
  fm_workbook_destroy(handle);
  ASSERT_EQ(status, 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> saved(bytes, bytes + len);
  fm_buffer_destroy(bytes);

  const auto read = io::xlsb::read_xlsb(io::ByteSpan{saved.data(), saved.size()});
  ASSERT_TRUE(read.has_value()) << read.error().message;
  ASSERT_EQ(read.value().workbook.pivot_caches().size(), 1U);
  const pivot::PivotCache& cache = *read.value().workbook.pivot_caches()[0];
  ASSERT_FALSE(cache.fields().empty());
  EXPECT_TRUE(HoldsValueError(cache.fields()[0].shared_items));
  for (const Value& v : cache.fields()[0].shared_items) {
    EXPECT_FALSE(v.is_blank());
  }
}

// ---------------------------------------------------------------------------
// Charts
// ---------------------------------------------------------------------------

TEST(PictureConsumers, ChartPointOverAPictureMatchesAValueErrorPoint) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(1, 1, Value::number(1.0));
  sheet.set_cell_value(2, 1, MakePicture(wb));
  sheet.set_cell_value(3, 1, Value::error(ErrorCode::Value));

  ChartDataSource source;
  source.kind = ChartDataSource::Kind::kRef;
  source.ref_element = ChartDataSource::RefElement::kNumRef;
  source.formula = "Sheet1!$B$2:$B$4";
  ChartSeries series;
  series.val = source;
  ChartPlot plot;
  plot.type = ChartPlotType::kBar;
  plot.series.push_back(std::move(series));
  ChartModel model;
  model.plots.push_back(std::move(plot));

  const auto data = resolve_chart_data(wb, 0, model);
  ASSERT_TRUE(data.has_value()) << data.error().message;
  ASSERT_EQ(data.value().series.size(), 1U);
  const std::vector<ChartDataPoint>& points = data.value().series[0].points;
  ASSERT_EQ(points.size(), 3U);
  EXPECT_EQ(points[1].value, points[2].value);
  EXPECT_EQ(points[1].state, points[2].state);
}

// ---------------------------------------------------------------------------
// Table totals label
// ---------------------------------------------------------------------------

TEST(PictureConsumers, TotalsLabelOverAPictureIsTheValueErrorLabel) {
  Workbook wb = Workbook::create();
  TableMetadata table;
  table.sheet_index = 0;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:B3";
  table.totals_row = true;
  for (const char* name : {"Region", "Amount"}) {
    TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  ASSERT_TRUE(wb.set_cell_text(0, 1U, 0U, "East").has_value());
  ASSERT_TRUE(wb.set_cell_value(0, 1U, 1U, Value::number(10.0)).has_value());
  ASSERT_TRUE(wb.add_table(std::move(table)).has_value());
  ASSERT_TRUE(wb.set_cell_value(0, 2U, 0U, MakePicture(wb)).has_value());

  const Cell* cell = wb.sheet(0).cell_at(2U, 0U);
  ASSERT_NE(cell, nullptr);
  const TotalsCellClass cls = classify_totals_cell(wb, wb.tables()[0], 0U, *cell);
  EXPECT_EQ(cls.kind, TotalsCellClass::Kind::kLabel);
  EXPECT_EQ(cls.text, "#VALUE!");
}

// ---------------------------------------------------------------------------
// What-if data table
// ---------------------------------------------------------------------------

DataTableInput InputAt(std::uint32_t row, std::uint32_t col) {
  DataTableInput input;
  input.present = true;
  input.row = row;
  input.col = col;
  return input;
}

// D2:D4 over input A1, with the header values in C2:C4; the formula tells an error input from any other.
TEST(PictureConsumers, DataTableSubstitutesAPictureHeaderAsAValueError) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::number(1.0)).has_value());
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 3, "=ISERROR(A1)").has_value());
  ASSERT_TRUE(wb.set_cell_value(0, 1, 2, Value::number(10.0)).has_value());
  ASSERT_TRUE(wb.set_cell_value(0, 2, 2, MakePicture(wb)).has_value());
  ASSERT_TRUE(wb.set_cell_value(0, 3, 2, Value::number(30.0)).has_value());
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{1, 3, 3, 3}, DataTableInput{}, InputAt(0, 0)).has_value());
  ASSERT_TRUE(wb.recalc(eval::default_registry()).has_value());
  const bool expected[] = {false, true, false};
  for (std::uint32_t k = 0; k < 3U; ++k) {
    const Cell* cell = wb.sheet(0).cell_at(1U + k, 3U);
    ASSERT_NE(cell, nullptr) << k;
    ASSERT_TRUE(cell->cached_value.is_boolean()) << k;
    EXPECT_EQ(cell->cached_value.as_boolean(), expected[k]) << k;
  }
}

// ---------------------------------------------------------------------------
// Conditional formats
// ---------------------------------------------------------------------------

// Whether a rule of `type` (a cellIs rule taking `op` and `operand`) matches A1 holding a picture; B1 holds the
// same picture when `picture_in_b1`.
bool CfMatchesPictureAtA1(cf::RuleType type, cf::CellIsOperator op = cf::CellIsOperator::GreaterThan,
                          const char* operand = "0", bool picture_in_b1 = false) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  const Value picture = MakePicture(wb);
  EXPECT_TRUE(wb.set_cell_value(0U, 0U, 0U, picture).has_value());
  if (picture_in_b1) {
    EXPECT_TRUE(wb.set_cell_value(0U, 0U, 1U, picture).has_value());
  }

  cf::ConditionalFormat block;
  cf::CFCellRange range{};
  block.sqref.push_back(range);
  cf::CFRule rule;
  rule.type = type;
  rule.priority = 1;
  rule.dxf_id = 7U;
  if (type == cf::RuleType::CellIs) {
    rule.op = op;
    rule.formula1 = operand;
  }
  block.rules.push_back(std::move(rule));
  sheet.mutable_conditional_formats().push_back(std::move(block));

  Arena arena;
  eval::EvalState state;
  eval::EvalContext eval_ctx(wb, sheet, state);
  cf::CFHost host;
  host.arena = &arena;
  host.registry = &eval::default_registry();
  host.eval_ctx = &eval_ctx;
  return !cf::evaluate_cf_at(sheet, CellAddress{0U, 0U}, host).empty();
}

TEST(PictureConsumers, ConditionalFormatContainsErrorsDoesNotMatchAPicture) {
  EXPECT_FALSE(CfMatchesPictureAtA1(cf::RuleType::ContainsErrors));
}

TEST(PictureConsumers, ConditionalFormatNotContainsErrorsMatchesAPicture) {
  EXPECT_TRUE(CfMatchesPictureAtA1(cf::RuleType::NotContainsErrors));
}

TEST(PictureConsumers, ConditionalFormatCellIsGreaterThanZeroDoesNotMatchAPicture) {
  EXPECT_FALSE(CfMatchesPictureAtA1(cf::RuleType::CellIs));
}

// Records today's result: the cellIs operand reads a picture as no operand, so `=` against the same picture does
// not match. Excel's result is not measured (divergence cell_picture_formula_value).
TEST(PictureConsumers, ConditionalFormatCellIsEqualToTheSamePictureDoesNotMatch) {
  EXPECT_FALSE(CfMatchesPictureAtA1(cf::RuleType::CellIs, cf::CellIsOperator::Equal, "$B$1", true));
}

// ---------------------------------------------------------------------------
// Data validation
// ---------------------------------------------------------------------------

TEST(PictureConsumers, ValidationRejectsAPictureForAWholeNumberRule) {
  Workbook wb = Workbook::create_empty();
  const std::size_t sheet_index = wb.add_sheet("S");
  Sheet& sheet = wb.sheet(sheet_index);
  DataValidation dv;
  dv.ranges.push_back(MergeRange{0, 0, 0, 0});
  dv.type = 1U;  // whole
  dv.op = 0U;    // between
  dv.allow_blank = false;
  dv.formula1 = "0";
  dv.formula2 = "10";
  sheet.mutable_validations().push_back(std::move(dv));

  const Value picture = MakePicture(wb);
  const auto outcome = validate_value(wb, sheet, 0U, 0U, picture);
  ASSERT_TRUE(outcome.has_value()) << outcome.error().message;
  EXPECT_TRUE(outcome.value().has_rule);
  EXPECT_FALSE(outcome.value().valid);
}

// ---------------------------------------------------------------------------
// AutoFilter
// ---------------------------------------------------------------------------

// Body-row visibility of A2:A4 (number, picture, #VALUE! error) under one custom condition on column A.
std::vector<bool> FilterRows(FilterOperator op) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EXPECT_TRUE(wb.set_cell_text(0U, 0U, 0U, "H").has_value());
  EXPECT_TRUE(wb.set_cell_value(0U, 1U, 0U, Value::number(1.0)).has_value());
  EXPECT_TRUE(wb.set_cell_value(0U, 2U, 0U, MakePicture(wb)).has_value());
  EXPECT_TRUE(wb.set_cell_value(0U, 3U, 0U, Value::error(ErrorCode::Value)).has_value());

  AutoFilter filter;
  filter.range = MergeRange{0, 0, 3, 0};
  FilterColumn column;
  column.col_id = 0U;
  column.kind = FilterKind::kCustom;
  CustomFilter condition;
  condition.op = op;
  condition.val = "#VALUE!";
  column.custom.filters.push_back(std::move(condition));
  filter.columns.push_back(std::move(column));

  const auto flags = evaluate_auto_filter(wb, sheet, filter);
  EXPECT_TRUE(flags.has_value());
  return flags.has_value() ? flags.value() : std::vector<bool>{};
}

TEST(PictureConsumers, AutoFilterEqualValueErrorDropsAPicture) {
  EXPECT_EQ(FilterRows(FilterOperator::kEqual), (std::vector<bool>{false, false, true}));
}

TEST(PictureConsumers, AutoFilterNotEqualValueErrorKeepsAPicture) {
  EXPECT_EQ(FilterRows(FilterOperator::kNotEqual), (std::vector<bool>{true, true, false}));
}

}  // namespace
}  // namespace formulon
