#include "io/xlsb/pivot_table_records.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot_twin_corpus.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using pivot_twins::CorpusPaths;
using pivot_twins::Parts;

constexpr const char* kTablePrefix = "xl/pivotTables/pivotTable";

/// Bytes of pivotTable1.bin in the corpus book whose file name starts with `stem`.
std::vector<std::uint8_t> TablePart(const std::string& stem) {
  for (const std::string& path : CorpusPaths()) {
    const std::string file = path.substr(path.find_last_of('/') + 1U);
    if (file.compare(0, stem.size(), stem) != 0) {
      continue;
    }
    for (auto& part : Parts(path, kTablePrefix)) {
      if (part.first == "xl/pivotTables/pivotTable1.bin") {
        return std::move(part.second);
      }
    }
  }
  ADD_FAILURE() << "no pivot table part for " << stem;
  return {};
}

PivotTableDoc Decode(const std::string& stem) {
  const std::vector<std::uint8_t> bytes = TablePart(stem);
  auto doc = decode_pivot_table(SpanOf(bytes));
  EXPECT_TRUE(doc.has_value()) << stem << ": " << (doc ? "" : doc.error().message);
  return doc ? std::move(doc.value()) : PivotTableDoc{};
}

/// The M-pivot table-side variants under tests/fixtures/excel/xlsb_parity/pivot.
constexpr const char* kParityVariants[] = {"formats_areas", "formats_dxf_order", "chart_formats", "x14_fill_down",
                                           "x14_show_as",   "hide_values_row",   "unknown_ext"};

std::string ParityPath(const std::string& file) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/pivot/" + file;
}

/// Bytes of pivotTable1.bin in the parity fixture `file`.
std::vector<std::uint8_t> ParityTablePart(const std::string& file) {
  for (auto& part : Parts(ParityPath(file), kTablePrefix)) {
    if (part.first == "xl/pivotTables/pivotTable1.bin") {
      return std::move(part.second);
    }
  }
  ADD_FAILURE() << "no pivot table part in " << file;
  return {};
}

PivotTableDoc DecodeParity(const std::string& variant) {
  const std::vector<std::uint8_t> bytes = ParityTablePart(variant + ".xlsb");
  auto doc = decode_pivot_table(SpanOf(bytes));
  EXPECT_TRUE(doc.has_value()) << variant << ": " << (doc ? "" : doc.error().message);
  return doc ? std::move(doc.value()) : PivotTableDoc{};
}

/// First item of field 0 whose flags carry `flag`, or nullptr.
const SxItem* FlaggedItem(const PivotTableDoc& doc, std::uint16_t flag) {
  if (doc.fields.empty() || !doc.fields[0].items) {
    return nullptr;
  }
  for (const SxItem& item : *doc.fields[0].items) {
    if ((item.flags & flag) != 0U) {
      return &item;
    }
  }
  return nullptr;
}

TEST(XlsbPivotTableRecords, EveryCorpusPartRoundTripsByteForByte) {
  std::size_t checked = 0;
  for (const std::string& path : CorpusPaths()) {
    for (const auto& part : Parts(path, kTablePrefix)) {
      SCOPED_TRACE(path + " " + part.first);
      auto doc = decode_pivot_table(SpanOf(part.second));
      ASSERT_TRUE(doc.has_value()) << doc.error().message;
      auto bytes = encode_pivot_table(doc.value());
      ASSERT_TRUE(bytes.has_value()) << bytes.error().message;
      EXPECT_EQ(bytes.value(), part.second);
      ++checked;
    }
  }
  EXPECT_EQ(checked, 89U);
}

TEST(XlsbPivotTableRecords, ViewHeaderBytes) {
  const PivotTableDoc base = Decode("B01_base");
  EXPECT_EQ(base.view.created_version, 8U);
  EXPECT_EQ(base.view.updated_version, 8U);
  EXPECT_EQ(base.view.min_refreshable_version, 3U);
  EXPECT_EQ(base.view.cache_id, 221U);
  EXPECT_EQ(base.view.data_position, 0xFFFFFFFFU);
  EXPECT_EQ(base.view.data_axis, 2U);
  EXPECT_EQ(base.view.name, "PivotTable1");
  EXPECT_EQ(base.view.data_caption, "Values");
  EXPECT_NE(base.view.flags5 & kSxViewAutoFormat, 0);
  EXPECT_NE(base.view.flags9 & kSxViewSingleFieldFilter, 0);
  EXPECT_EQ(base.view.flags3, 0U);
  EXPECT_FALSE(base.view.grand_total_caption || base.view.error_caption || base.view.missing_caption);

  EXPECT_EQ(Decode("B06_values_on_rows").view.data_axis, 1U);
  EXPECT_EQ(Decode("E_ft_m02_two_filters_mff1").view.flags9 & kSxViewSingleFieldFilter, 0);
  EXPECT_EQ(Decode("B15_indent").view.flags3 & kSxViewIndentMask, 2);
  EXPECT_NE(Decode("B08_headers_off").view.flags3 & kSxViewNoHeaders, 0);
  EXPECT_EQ(Decode("B04_rowgrand_off").view.flags5 & kSxViewColGrand, 0);
  EXPECT_EQ(Decode("B05_colgrand_off").view.flags5 & kSxViewRowGrand, 0);
  EXPECT_NE(Decode("B14_merge_items").view.flags6 & kSxViewMergeItem, 0);
  EXPECT_EQ(Decode("B09_grid_drop_zones").view.flags8 & kSxViewNoDropZones, 0);
  EXPECT_EQ(Decode("B18_name").view.name, "SalesPivot");
  EXPECT_EQ(Decode("B10_data_caption").view.data_caption, "Measures");
  EXPECT_EQ(Decode("B11_grand_total_caption").view.grand_total_caption, std::optional<std::string>("All Regions"));
  const PivotTableDoc error = Decode("B12_error_caption");
  EXPECT_EQ(error.view.error_caption, std::optional<std::string>("ERR"));
  EXPECT_NE(error.view.flags5 & kSxViewShowError, 0);
  EXPECT_EQ(Decode("B13_missing_caption").view.missing_caption, std::optional<std::string>("n/a"));
}

TEST(XlsbPivotTableRecords, UidIsTheMixedEndianGuid) {
  // xr:uid="{55D3C1B4-BB27-F64C-B779-ECFCC1DF7353}"
  const std::array<std::uint8_t, 16> expected = {0xb4, 0xc1, 0xd3, 0x55, 0x27, 0xbb, 0x4c, 0xf6,
                                                 0xb7, 0x79, 0xec, 0xfc, 0xc1, 0xdf, 0x73, 0x53};
  const PivotTableDoc doc = Decode("B01_base");
  ASSERT_TRUE(doc.uid.has_value());
  EXPECT_EQ(*doc.uid, expected);
}

TEST(XlsbPivotTableRecords, LocationAndPageField) {
  const PivotTableDoc doc = Decode("B17_page_field");
  // location ref="A3:C7" firstHeaderRow="0" firstDataRow="1" firstDataCol="1" rowPageCount="1" colPageCount="1"
  EXPECT_EQ(doc.location.first_row, 2U);
  EXPECT_EQ(doc.location.last_row, 6U);
  EXPECT_EQ(doc.location.first_col, 0U);
  EXPECT_EQ(doc.location.last_col, 2U);
  EXPECT_EQ(doc.location.first_header_row, 2U);
  EXPECT_EQ(doc.location.first_data_row, 3U);
  EXPECT_EQ(doc.location.first_data_col, 1U);
  EXPECT_EQ(doc.location.row_page_count, 1U);
  EXPECT_EQ(doc.location.col_page_count, 1U);
  ASSERT_TRUE(doc.page_fields.has_value());
  ASSERT_EQ(doc.page_fields->size(), 1U);
  EXPECT_EQ((*doc.page_fields)[0].field, 1U);
  EXPECT_EQ((*doc.page_fields)[0].item, 0x001000FEU);
  EXPECT_EQ((*doc.page_fields)[0].hierarchy, 0xFFFFFFFFU);
  ASSERT_GT(doc.fields.size(), 1U);
  EXPECT_EQ(doc.fields[1].axis, kSxFieldAxisPage);

  const PivotTableDoc base = Decode("B01_base");
  EXPECT_FALSE(base.page_fields.has_value());
  EXPECT_EQ(base.location.row_page_count, 0U);
}

TEST(XlsbPivotTableRecords, FieldHeaderBytes) {
  const PivotTableDoc base = Decode("B01_base");
  ASSERT_EQ(base.fields.size(), 5U);
  EXPECT_EQ(base.fields[0].axis, kSxFieldAxisRow);
  EXPECT_EQ(base.fields[0].subtotals, kSxFieldSubtotalDefault);
  EXPECT_NE(base.fields[0].name_flags & kSxFieldCompact, 0);
  EXPECT_EQ(base.fields[2].num_fmt, 14U);
  EXPECT_EQ(base.fields[3].axis, kSxFieldAxisData);
  EXPECT_EQ(base.fields[0].item_page_count, 10U);
  EXPECT_NE(base.fields[0].flags & kSxFieldOutline, 0U);
  EXPECT_NE(base.fields[0].flags & kSxFieldSubtotalTop, 0U);
  EXPECT_EQ(base.fields[0].flags & kSxFieldShowAll, 0U);
  ASSERT_TRUE(base.fields[0].items.has_value());
  EXPECT_EQ(base.fields[0].items->size(), 4U);
  EXPECT_FALSE(base.fields[1].items.has_value());

  EXPECT_EQ(Decode("C08_field_caption").fields[0].display_name, std::optional<std::string>("Area"));
  const PivotTableDoc numfmt = Decode("C09_field_numfmt");
  ASSERT_EQ(numfmt.fields.size(), 5U);
  EXPECT_EQ(numfmt.fields[4].num_fmt, 2U);
  EXPECT_EQ(numfmt.fields[4].axis, kSxFieldAxisRow | kSxFieldAxisData);
  EXPECT_EQ(Decode("C03_default_subtotal_off").fields[0].subtotals, 0U);
  EXPECT_NE(Decode("C04_show_all").fields[0].flags & kSxFieldShowAll, 0U);
  EXPECT_NE(Decode("C02_insert_blank_row").fields[0].flags & kSxFieldInsertBlankRow, 0U);
  EXPECT_EQ(Decode("C01_subtotal_bottom").fields[0].flags & kSxFieldSubtotalTop, 0U);
  const std::uint32_t asc = Decode("C05_sort_asc").fields[0].flags;
  EXPECT_NE(asc & kSxFieldAutoSort, 0U);
  EXPECT_NE(asc & kSxFieldAscendSort, 0U);
  const std::uint32_t desc = Decode("C06_sort_desc").fields[0].flags;
  EXPECT_NE(desc & kSxFieldAutoSort, 0U);
  EXPECT_EQ(desc & kSxFieldAscendSort, 0U);
  const PivotTableDoc measure = Decode("E_ft_t01_top2");
  bool any_measure_filter = false;
  for (const SxField& field : measure.fields) {
    any_measure_filter = any_measure_filter || (field.flags & kSxFieldMeasureFilter) != 0U;
  }
  EXPECT_TRUE(any_measure_filter);
}

TEST(XlsbPivotTableRecords, ItemFlags) {
  const PivotTableDoc hidden_doc = Decode("C10_item_hidden");
  const SxItem* hidden = FlaggedItem(hidden_doc, kSxItemHidden);
  ASSERT_NE(hidden, nullptr);
  EXPECT_EQ(hidden->kind, 0U);
  EXPECT_EQ(hidden->index, 2U);

  const PivotTableDoc collapsed_doc = Decode("C11_item_collapsed");
  const SxItem* collapsed = FlaggedItem(collapsed_doc, kSxItemCollapsed);
  ASSERT_NE(collapsed, nullptr);
  EXPECT_EQ(collapsed->index, 0U);

  const PivotTableDoc caption_doc = Decode("C12_item_caption");
  const SxItem* caption = FlaggedItem(caption_doc, kSxItemHasCaption);
  ASSERT_NE(caption, nullptr);
  EXPECT_EQ(caption->index, 0U);
  EXPECT_EQ(caption->caption, std::optional<std::string>("Eastern"));

  const PivotTableDoc calculated_doc = Decode("C13_calc_item");
  const SxItem* calculated = FlaggedItem(calculated_doc, kSxItemCalculated);
  ASSERT_NE(calculated, nullptr);
  EXPECT_EQ(calculated->index, 3U);

  const PivotTableDoc missing_doc = Decode("A5_C9__sum");
  const SxItem* missing = FlaggedItem(missing_doc, kSxItemMissing);
  ASSERT_NE(missing, nullptr);
  EXPECT_EQ(missing->kind, 0U);
}

TEST(XlsbPivotTableRecords, CustomSubtotalItems) {
  struct Case {
    const char* stem;
    std::uint8_t kind;
    std::uint16_t subtotal_bit;
  };
  const Case cases[] = {
      {"A8_custom_subtotal_countA", 3, 0x0004},   {"A8_custom_subtotal_max.", 5, 0x0010},
      {"A8_custom_subtotal_min.", 6, 0x0020},     {"A8_custom_subtotal_stdDev.", 9, 0x0100},
      {"A8_custom_subtotal_stdDevP", 10, 0x0200}, {"A8_custom_subtotal_varP", 12, 0x0800},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.stem);
    const PivotTableDoc doc = Decode(c.stem);
    ASSERT_FALSE(doc.fields.empty());
    EXPECT_EQ(doc.fields[0].subtotals, c.subtotal_bit);
    ASSERT_TRUE(doc.fields[0].items.has_value());
    std::size_t subtotal_items = 0;
    for (const SxItem& item : *doc.fields[0].items) {
      if (item.kind != 0U) {
        EXPECT_EQ(item.kind, c.kind);
        EXPECT_EQ(item.index, 0xFFFFFFFFU);
        ++subtotal_items;
      }
    }
    EXPECT_EQ(subtotal_items, 1U);
  }
}

TEST(XlsbPivotTableRecords, AutoSortScopeIsCarriedVerbatim) {
  const PivotTableDoc doc = Decode("C07_sort_by_data");
  ASSERT_FALSE(doc.fields.empty());
  const auto& scope = doc.fields[0].auto_sort_scope;
  ASSERT_EQ(scope.size(), 10U);
  EXPECT_EQ(scope.front().type, pivot_records::kBeginAutoSortScope);
  EXPECT_EQ(scope.back().type, pivot_records::kEndAutoSortScope);
  EXPECT_EQ(scope[1].type, pivot_records::kBeginPRule);
  EXPECT_EQ(scope[3].type, pivot_records::kBeginPRFilter);
  const std::vector<std::uint8_t> reference = {0xfe, 0xff, 0xff, 0xff, 0x01, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00};
  EXPECT_EQ(scope[3].payload, reference);
  for (std::size_t k = 1; k < doc.fields.size(); ++k) {
    EXPECT_TRUE(doc.fields[k].auto_sort_scope.empty());
  }
}

TEST(XlsbPivotTableRecords, DataFields) {
  const PivotTableDoc base = Decode("B01_base");
  ASSERT_EQ(base.data_fields.size(), 2U);
  EXPECT_EQ(base.data_fields[0].field, 3U);
  EXPECT_EQ(base.data_fields[0].name, "Sum of Amt");
  EXPECT_EQ(base.data_fields[1].field, 4U);
  EXPECT_EQ(base.data_fields[1].name, "Sum of Qty");

  // dataField fld="3" showDataAs="difference" baseField="0" baseItem="2"
  const PivotTableDoc diff = Decode("D02_showas_difference_from");
  ASSERT_FALSE(diff.data_fields.empty());
  EXPECT_EQ(diff.data_fields[0].show_data_as, 1U);
  EXPECT_EQ(diff.data_fields[0].base_field, 0U);
  EXPECT_EQ(diff.data_fields[0].base_item, 2U);

  // dataField fld="3" numFmtId="176"
  const PivotTableDoc numfmt = Decode("D03_data_numfmt");
  ASSERT_FALSE(numfmt.data_fields.empty());
  EXPECT_EQ(numfmt.data_fields[0].num_fmt, 176U);
  EXPECT_EQ(numfmt.data_fields[0].base_item, 0U);
}

TEST(XlsbPivotTableRecords, AxisFieldsStyleFiltersAndFrt) {
  const PivotTableDoc base = Decode("B01_base");
  ASSERT_TRUE(base.row_fields.has_value());
  EXPECT_EQ(*base.row_fields, std::vector<std::uint32_t>({0U}));
  ASSERT_TRUE(base.col_fields.has_value());
  EXPECT_EQ(*base.col_fields, std::vector<std::uint32_t>({0xFFFFFFFEU}));
  EXPECT_EQ(base.style.name, std::optional<std::string>("PivotStyleLight16"));
  EXPECT_EQ(base.style.flags, 0x0032U);
  EXPECT_TRUE(base.filters.empty());
  ASSERT_TRUE(base.view14.has_value());
  EXPECT_EQ(base.view14->frt_payload, std::vector<std::uint8_t>({0x02, 0x0e, 0x00, 0x00}));
  EXPECT_EQ(base.view14->payload.size(), 21U);
  EXPECT_EQ(base.view16.frt_payload, std::vector<std::uint8_t>({0x03, 0x11, 0x00, 0x00}));
  EXPECT_EQ(base.view16.payload, std::vector<std::uint8_t>({0x01, 0x01, 0x00}));

  const PivotTableDoc filtered = Decode("E_ft_m02_two_filters_mff1");
  ASSERT_FALSE(filtered.filters.empty());
  EXPECT_EQ(filtered.filters.front().type, pivot_records::kBeginSXFilters);
  EXPECT_EQ(filtered.filters.back().type, pivot_records::kEndSXFilters);
}

TEST(XlsbPivotTableRecords, EncodeRefusesSeveralViewCaptions) {
  PivotTableDoc doc = Decode("B01_base");
  doc.view.error_caption = "ERR";
  doc.view.missing_caption = "n/a";
  auto bytes = encode_pivot_table(doc);
  ASSERT_FALSE(bytes.has_value());
  EXPECT_EQ(bytes.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(XlsbPivotTableRecords, DecodeRefusesTrailingRecord) {
  std::vector<std::uint8_t> bytes = TablePart("B01_base");
  bytes.push_back(0x01);  // a BrtCellBlank with an empty payload after BrtEndSXView
  bytes.push_back(0x00);
  auto doc = decode_pivot_table(SpanOf(bytes));
  ASSERT_FALSE(doc.has_value());
  EXPECT_EQ(doc.error().code, FormulonErrorCode::kIoXlsbUnknownRecord);
}

TEST(XlsbPivotTableRecords, ParityVariantsRoundTripByteForByte) {
  for (const char* variant : kParityVariants) {
    for (const std::string& file : {std::string(variant) + ".xlsb", std::string(variant) + ".2.xlsb"}) {
      SCOPED_TRACE(file);
      const std::vector<std::uint8_t> bytes = ParityTablePart(file);
      ASSERT_FALSE(bytes.empty());
      auto doc = decode_pivot_table(SpanOf(bytes));
      ASSERT_TRUE(doc.has_value()) << doc.error().message;
      auto encoded = encode_pivot_table(doc.value());
      ASSERT_TRUE(encoded.has_value()) << encoded.error().message;
      EXPECT_EQ(encoded.value(), bytes);
    }
  }
}

TEST(XlsbPivotTableRecords, FormatsAreas) {
  const PivotTableDoc doc = DecodeParity("formats_areas");
  EXPECT_TRUE(doc.chart_formats.empty());
  ASSERT_EQ(doc.formats.size(), 5U);
  // dxfId 4: outline="0" fieldPosition="0", reference field 0 x 1 selected="0".
  EXPECT_EQ(doc.formats[0].dxf_id, 4U);
  const pivot_area::PivotArea& item = doc.formats[0].area;
  EXPECT_EQ(item.field, pivot_area::kNoField);
  EXPECT_EQ(item.type, pivot_area::kAreaTypeNormal);
  EXPECT_EQ(item.flags, pivot_area::kAreaDataOnly);
  ASSERT_TRUE(item.filters.has_value());
  ASSERT_EQ(item.filters->size(), 1U);
  EXPECT_EQ((*item.filters)[0].field, 0U);
  EXPECT_EQ((*item.filters)[0].flags, pivot_area::kFilterAlways);
  EXPECT_EQ((*item.filters)[0].items, std::vector<std::uint32_t>({1U}));
  // dxfId 3: dataOnly="0" labelOnly="1", reference selected by default.
  EXPECT_EQ(doc.formats[1].area.flags, pivot_area::kAreaLabelOnly);
  ASSERT_TRUE(doc.formats[1].area.filters.has_value());
  EXPECT_EQ((*doc.formats[1].area.filters)[0].flags, pivot_area::kFilterAlways | pivot_area::kFilterSelected);
  // dxfId 2: field="0" type="button" labelOnly="1" axis="axisRow".
  const pivot_area::PivotArea& button = doc.formats[2].area;
  EXPECT_EQ(button.field, 0U);
  EXPECT_EQ(button.type, pivot_area::kAreaTypeButton);
  EXPECT_EQ(button.flags, pivot_area::kAreaLabelOnly | pivot_area::kAreaAxisRow);
  EXPECT_FALSE(button.filters.has_value());
  // dxfId 1 and 0: grandRow and grandCol, label only.
  EXPECT_EQ(doc.formats[3].area.flags, pivot_area::kAreaLabelOnly | pivot_area::kAreaGrandRow);
  EXPECT_EQ(doc.formats[4].area.flags, pivot_area::kAreaLabelOnly | pivot_area::kAreaGrandCol);
  EXPECT_EQ(doc.formats[4].dxf_id, 0U);
}

TEST(XlsbPivotTableRecords, FormatsDxfOrder) {
  const PivotTableDoc doc = DecodeParity("formats_dxf_order");
  ASSERT_EQ(doc.formats.size(), 2U);
  EXPECT_EQ(doc.formats[0].dxf_id, 17U);
  EXPECT_EQ(doc.formats[0].area.type, pivot_area::kAreaTypeAll);
  EXPECT_EQ(doc.formats[0].area.flags, 0U);
  EXPECT_EQ(doc.formats[1].dxf_id, 16U);
  EXPECT_EQ(doc.formats[1].area.flags, pivot_area::kAreaDataOnly | pivot_area::kAreaCollapsedLevelsAreSubtotals);
}

TEST(XlsbPivotTableRecords, ChartFormats) {
  const PivotTableDoc doc = DecodeParity("chart_formats");
  EXPECT_TRUE(doc.formats.empty());
  ASSERT_EQ(doc.chart_formats.size(), 2U);
  for (std::uint32_t i = 0; i < 2U; ++i) {
    const SxChartFormat& format = doc.chart_formats[i];
    EXPECT_EQ(format.format, i);
    EXPECT_EQ(format.area.type, pivot_area::kAreaTypeData);
    EXPECT_EQ(format.area.flags, pivot_area::kAreaDataOnly);
    ASSERT_TRUE(format.area.filters.has_value());
    EXPECT_EQ((*format.area.filters)[0].field, pivot_area::kDataField);
    EXPECT_EQ((*format.area.filters)[0].items, std::vector<std::uint32_t>({i}));
  }
  // The root `chartFormat="1"`.
  EXPECT_EQ(doc.view.reserved24, 1U);
}

TEST(XlsbPivotTableRecords, X14FieldAndDataField) {
  const PivotTableDoc fill = DecodeParity("x14_fill_down");
  ASSERT_EQ(fill.fields.size(), 5U);
  EXPECT_TRUE(fill.fields[0].fill_down_labels);
  for (std::size_t i = 1; i < fill.fields.size(); ++i) {
    EXPECT_FALSE(fill.fields[i].fill_down_labels) << i;
  }
  const PivotTableDoc show = DecodeParity("x14_show_as");
  ASSERT_EQ(show.data_fields.size(), 2U);
  EXPECT_EQ(show.data_fields[0].pivot_show_as, std::optional<std::uint32_t>(kSxDataShowAsPercentOfParentRow));
  EXPECT_EQ(show.data_fields[0].show_data_as, 0U);
  EXPECT_FALSE(show.data_fields[1].pivot_show_as.has_value());
}

TEST(XlsbPivotTableRecords, HideValuesRowOffHasNoView14) {
  EXPECT_FALSE(DecodeParity("hide_values_row").view14.has_value());
  const PivotTableDoc base = DecodeParity("unknown_ext");
  EXPECT_TRUE(base.view14.has_value());
  EXPECT_TRUE(base.formats.empty());
  EXPECT_TRUE(base.chart_formats.empty());
}

TEST(XlsbPivotTableRecords, EncodeRefusesUnmeasuredFormatStates) {
  const PivotTableDoc formats = DecodeParity("formats_areas");
  const PivotTableDoc charts = DecodeParity("chart_formats");
  const PivotTableDoc show = DecodeParity("x14_show_as");
  const auto refused = [](const PivotTableDoc& doc) {
    auto bytes = encode_pivot_table(doc);
    return !bytes.has_value() && bytes.error().code == FormulonErrorCode::kInvalidArgument;
  };

  PivotTableDoc both = formats;
  both.chart_formats = charts.chart_formats;
  EXPECT_TRUE(refused(both));

  PivotTableDoc origin = formats;
  origin.formats[0].area.type = 4U;  // `origin`
  EXPECT_TRUE(refused(origin));

  PivotTableDoc col_axis = formats;
  col_axis.formats[2].area.flags = pivot_area::kAreaLabelOnly | 0x000200U;  // `axisCol`
  EXPECT_TRUE(refused(col_axis));

  PivotTableDoc empty_ref = formats;
  (*empty_ref.formats[0].area.filters)[0].items.clear();
  EXPECT_TRUE(refused(empty_ref));

  PivotTableDoc show_col = show;
  show_col.data_fields[0].pivot_show_as = 10U;  // `percentOfParentCol`
  EXPECT_TRUE(refused(show_col));

  PivotTableDoc fill_no_items = DecodeParity("x14_fill_down");
  fill_no_items.fields[1].fill_down_labels = true;  // a field without items
  EXPECT_TRUE(refused(fill_no_items));
}

TEST(XlsbPivotTableRecords, DecodeRefusesUnmeasuredFormatAction) {
  std::vector<std::uint8_t> bytes = ParityTablePart("formats_areas.xlsb");
  // BrtBeginSXFormat (303 = AF 02), 6-byte payload opening with action bytes 01 00.
  const std::vector<std::uint8_t> head = {0xAF, 0x02, 0x06, 0x01, 0x00};
  const auto at = std::search(bytes.begin(), bytes.end(), head.begin(), head.end());
  ASSERT_NE(at, bytes.end());
  *(at + 3) = 0x02;
  auto doc = decode_pivot_table(SpanOf(bytes));
  ASSERT_FALSE(doc.has_value());
  EXPECT_EQ(doc.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
