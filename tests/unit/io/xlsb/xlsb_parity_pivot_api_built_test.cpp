// PivotTables built through the C API and saved as `.xlsb`: the cases of tests/fixtures/excel/xlsb_parity/pivot/
// api_built, each built as the workbook Excel 16.115 opened from the engine's `.xlsx` save and re-saved as `.xlsb`
// (`<case>.xlsb`). The save keeps every pivot part, passes the package validator, reads back to the same tables, and
// its pivot records match Excel's but for the revision uid Excel adds. A calculated field has no record layout yet,
// so that case is left out and counted.

#include <cstddef>
#include <cstdint>
#include <set>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/pivot_records_common.h"
#include "io/xlsb/record_ids.h"
#include "support/xlsb_package_validator.h"
#include "writer_test_helpers.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace tc = test;

constexpr const char* kRegions[] = {"East", "West", "North"};
constexpr const char* kProducts[] = {"Apple", "Pear", "Plum", "Fig"};
constexpr std::uint32_t kRows = 12;
/// Mac Excel saved the twins from a workbook built under English captions.
constexpr const char* kProfile = "mac-365-en_US";

/// The Data sheet A1:E13 (Region, Product, Date, Amt, Qty) and a refreshed cache on it.
std::uint32_t BuildSource(fm_workbook_t* wb) {
  EXPECT_EQ(fm_workbook_rename_sheet(wb, 0U, "Data"), 0) << fm_get_last_error_message();
  const char* const headers[] = {"Region", "Product", "Date", "Amt", "Qty"};
  for (std::uint32_t c = 0; c < 5U; ++c) {
    EXPECT_EQ(fm_sheet_set_text(wb, 0U, 0U, c, headers[c]), 0);
  }
  std::uint32_t num_fmt = 0;
  EXPECT_EQ(fm_workbook_add_number_format(wb, "yyyy-mm-dd", &num_fmt), 0) << fm_get_last_error_message();
  fm_cell_xf_t xf{};
  xf.num_fmt_id = num_fmt;
  std::uint32_t date_xf = 0;
  EXPECT_EQ(fm_workbook_add_cell_xf(wb, &xf, &date_xf), 0) << fm_get_last_error_message();
  for (std::uint32_t i = 0; i < kRows; ++i) {
    EXPECT_EQ(fm_sheet_set_text(wb, 0U, i + 1U, 0U, kRegions[i % 3U]), 0);
    EXPECT_EQ(fm_sheet_set_text(wb, 0U, i + 1U, 1U, kProducts[i % 4U]), 0);
    EXPECT_EQ(fm_sheet_set_number(wb, 0U, i + 1U, 2U, 45292.0 + 31.0 * i), 0);
    EXPECT_EQ(fm_sheet_set_cell_xf_index(wb, 0U, i + 1U, 2U, date_xf), 0);
    EXPECT_EQ(fm_sheet_set_number(wb, 0U, i + 1U, 3U, 100.0 + 10.0 * i), 0);
    EXPECT_EQ(fm_sheet_set_number(wb, 0U, i + 1U, 4U, 1.0 + i % 5U), 0);
  }
  std::uint32_t cache_id = 0;
  EXPECT_EQ(fm_workbook_add_pivot_cache(wb, 0U, &cache_id), 0) << fm_get_last_error_message();
  fm_pivot_worksheet_source_t source{sizeof(fm_pivot_worksheet_source_t), "A1:E13", "Data", nullptr};
  EXPECT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb, cache_id, &source), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_workbook_refresh_pivot_cache(wb, cache_id), 0) << fm_get_last_error_message();
  return cache_id;
}

/// A new sheet `sheet` holding pivot `name` on `cache_id` at row `anchor_row`, column A.
struct Placed {
  std::size_t sheet = 0;
  std::size_t pivot = 0;
};
Placed AddPivot(fm_workbook_t* wb, const char* sheet, std::uint32_t cache_id, const char* name,
                std::uint32_t anchor_row = 2U) {
  Placed out;
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_get_sheet_count(wb, &count), 0);
  EXPECT_EQ(fm_workbook_add_sheet(wb, sheet, count, &out.sheet), 0) << fm_get_last_error_message();
  fm_pivot_spec_t spec{sizeof(fm_pivot_spec_t), name, cache_id, anchor_row, 0U};
  EXPECT_EQ(fm_sheet_add_pivot(wb, out.sheet, &spec, &out.pivot), 0) << fm_get_last_error_message();
  return out;
}

void Axis(fm_workbook_t* wb, const Placed& p, std::size_t field, int32_t axis) {
  EXPECT_EQ(fm_sheet_set_pivot_field_axis(wb, p.sheet, p.pivot, field, axis), 0) << fm_get_last_error_message();
}

void SumOf(fm_workbook_t* wb, const Placed& p, std::size_t field) {
  fm_pivot_data_field_t data{};
  data.struct_size = sizeof(fm_pivot_data_field_t);
  data.field_index = static_cast<std::uint32_t>(field);
  data.aggregation = FM_PIVOT_AGGREGATION_SUM;
  std::size_t index = 0;
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb, p.sheet, p.pivot, &data, &index), 0) << fm_get_last_error_message();
}

void Project(fm_workbook_t* wb, const Placed& p) {
  EXPECT_EQ(fm_sheet_project_pivot(wb, p.sheet, p.pivot), 0) << fm_get_last_error_message();
}

void BuildCase(const std::string& name, fm_workbook_t* wb) {
  const std::uint32_t cache = BuildSource(wb);
  if (name == "a_plain") {
    const Placed p = AddPivot(wb, "PA", cache, "PivotA");
    Axis(wb, p, 0U, FM_PIVOT_AXIS_ROW);
    Axis(wb, p, 1U, FM_PIVOT_AXIS_COLUMN);
    SumOf(wb, p, 3U);
    Project(wb, p);
  } else if (name == "b_page_filter") {
    const Placed p = AddPivot(wb, "PB", cache, "PivotB", 3U);
    Axis(wb, p, 0U, FM_PIVOT_AXIS_PAGE);
    Axis(wb, p, 1U, FM_PIVOT_AXIS_ROW);
    SumOf(wb, p, 3U);
    std::size_t items = 0;
    EXPECT_EQ(fm_sheet_get_pivot_field_item_count(wb, p.sheet, p.pivot, 0U, &items), 0);
    for (std::size_t i = 0; i < items; ++i) {
      fm_pivot_item_info_t info{};
      info.struct_size = sizeof(info);
      EXPECT_EQ(fm_sheet_get_pivot_field_item_at(wb, p.sheet, p.pivot, 0U, i, &info), 0);
      if (info.has_cache_index != 0 && std::string(info.name != nullptr ? info.name : "") != "East") {
        fm_pivot_item_patch_t patch{};
        patch.struct_size = sizeof(patch);
        patch.visible_op = FM_PATCH_OP_SET;
        patch.visible = 0;
        EXPECT_EQ(fm_sheet_update_pivot_field_item(wb, p.sheet, p.pivot, 0U, i, &patch), 0)
            << fm_get_last_error_message();
      }
    }
    fm_pivot_filter_t filter{};
    filter.struct_size = sizeof(filter);
    filter.field_index = 1U;
    filter.kind = FM_PIVOT_FILTER_KIND_LABEL_BEGINS_WITH;
    filter.value_kind = FM_PIVOT_FILTER_OPERAND_KIND_TEXT;
    filter.value_text = "P";
    std::size_t filter_index = 0;
    EXPECT_EQ(fm_sheet_add_pivot_filter(wb, p.sheet, p.pivot, &filter, &filter_index), 0)
        << fm_get_last_error_message();
    Project(wb, p);
  } else if (name == "c_date_group") {
    const Placed p = AddPivot(wb, "PC", cache, "PivotC");
    Axis(wb, p, 2U, FM_PIVOT_AXIS_ROW);
    fm_pivot_date_group_t group{};
    group.struct_size = sizeof(group);
    group.periods = FM_PIVOT_PERIOD_YEARS | FM_PIVOT_PERIOD_MONTHS;
    group.calendar = FM_PIVOT_CALENDAR_GREGORIAN;
    group.interval_days = 1U;
    EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb, p.sheet, p.pivot, 2U, &group), 0) << fm_get_last_error_message();
    SumOf(wb, p, 3U);
    Project(wb, p);
  } else if (name == "d_calc_field") {
    std::size_t margin = 0;
    EXPECT_EQ(fm_workbook_add_pivot_cache_calculated_field(wb, cache, "Margin", "=Amt*0.1", &margin), 0)
        << fm_get_last_error_message();
    const Placed p = AddPivot(wb, "PD", cache, "PivotD");
    Axis(wb, p, 0U, FM_PIVOT_AXIS_ROW);
    SumOf(wb, p, 3U);
    SumOf(wb, p, margin);
    Project(wb, p);
  } else if (name == "e_shared_cache") {
    const Placed p1 = AddPivot(wb, "PE1", cache, "PivotE1");
    Axis(wb, p1, 0U, FM_PIVOT_AXIS_ROW);
    SumOf(wb, p1, 3U);
    Project(wb, p1);
    const Placed p2 = AddPivot(wb, "PE2", cache, "PivotE2");
    Axis(wb, p2, 1U, FM_PIVOT_AXIS_ROW);
    SumOf(wb, p2, 4U);
    Project(wb, p2);
  } else {
    ADD_FAILURE() << "unknown case " << name;
  }
}

/// Each pivot's name and the display text of its sheet's A1:H20, per sheet.
std::vector<std::string> Report(const fm_workbook_t* wb) {
  std::vector<std::string> out;
  std::size_t sheets = 0;
  EXPECT_EQ(fm_workbook_get_sheet_count(wb, &sheets), 0);
  for (std::size_t s = 0; s < sheets; ++s) {
    std::size_t pivots = 0;
    EXPECT_EQ(fm_sheet_get_pivot_count(wb, s, &pivots), 0);
    for (std::size_t p = 0; p < pivots; ++p) {
      const char* name = nullptr;
      EXPECT_EQ(fm_sheet_get_pivot_name(wb, s, p, &name), 0);
      out.push_back("sheet " + std::to_string(s) + " pivot " + (name != nullptr ? name : ""));
    }
    if (pivots == 0U) {
      continue;
    }
    for (std::uint32_t r = 0; r < 20U; ++r) {
      std::string line = "  row " + std::to_string(r) + ":";
      for (std::uint32_t c = 0; c < 8U; ++c) {
        const char* text = nullptr;
        std::int32_t status = 0;
        EXPECT_EQ(fm_sheet_get_display_text(wb, s, r, c, &text, &status), 0);
        line += std::string(" |") + (text != nullptr ? text : "");
      }
      out.push_back(line);
    }
  }
  return out;
}

/// The pivot lines of a `Report`.
std::vector<std::string> PivotNames(const std::vector<std::string>& report) {
  std::vector<std::string> out;
  for (const std::string& line : report) {
    if (line.rfind("sheet ", 0) == 0) {
      out.push_back(line);
    }
  }
  return out;
}

struct Saved {
  std::vector<std::uint8_t> bytes;
  std::size_t deferred = 0;
  std::vector<std::string> report;  ///< `Report` of the built workbook.
};

Saved SaveCase(const std::string& name) {
  Saved out;
  fm_workbook_t* wb = nullptr;
  EXPECT_EQ(fm_workbook_create_for_profile(kProfile, &wb), 0) << fm_get_last_error_message();
  BuildCase(name, wb);
  out.report = Report(wb);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  fm_save_diagnostics_info_t info{};
  info.struct_size = sizeof(fm_save_diagnostics_info_t);
  const fm_status_t status = fm_workbook_save_with_diagnostics(wb, FM_WORKBOOK_FORMAT_XLSB, &bytes, &len, &info);
  fm_workbook_destroy(wb);
  EXPECT_EQ(status, 0) << fm_get_last_error_message();
  if (status == 0) {
    out.bytes.assign(bytes, bytes + len);
    fm_buffer_destroy(bytes);
  }
  out.deferred = info.deferred_feature_count;
  return out;
}

std::vector<std::string> ReloadedReport(const std::vector<std::uint8_t>& bytes) {
  fm_workbook_t* wb = nullptr;
  fm_load_options_spec_t options{sizeof(fm_load_options_spec_t), kProfile};
  EXPECT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &options, &wb), 0) << fm_get_last_error_message();
  if (wb == nullptr) {
    return {};
  }
  std::vector<std::string> out = Report(wb);
  fm_workbook_destroy(wb);
  return out;
}

void ExpectValid(const std::vector<std::uint8_t>& bytes) {
  const std::vector<std::string> findings = formulon::test::ValidateXlsbPackage(SpanOf(bytes));
  EXPECT_TRUE(findings.empty()) << findings.front();
}

/// Every record type of Excel's pivot parts but the revision-uid wrapper Excel adds on save, which the engine's
/// saves never write.
std::set<std::uint16_t> PivotTypesButUid(const tc::Bytes& excel) {
  std::set<std::uint16_t> types;
  for (const auto& [part, records] : tc::ExtractRecords(excel, {})) {
    if (part.rfind("xl/pivot", 0) == 0) {
      for (const tc::IndexedRecord& rec : records) {
        types.insert(rec.type);
      }
    }
  }
  for (const std::uint16_t uid_record : {kAcBegin, kUid, kAcEnd}) {
    types.erase(uid_record);
  }
  return types;
}

/// What a comparison ignores: the union Excel leaves unset before a string filter operand, and the ids Excel
/// renumbers on load, the table's idCache (BrtBeginSXView bytes 28-31) and a custom number format's id
/// (BrtBeginPCDField bytes 2-3).
tc::ByteMask Unmatched(const tc::Bytes& excel) {
  constexpr std::uint16_t kFirstCustomNumFmt = 164;
  constexpr std::uint8_t kVtsString = 0x06;
  tc::ByteMask mask;
  // The union before a string operand is unset memory in Excel's save.
  for (const auto& [part, records] : tc::ExtractRecords(excel, {kBrtCustomFilter})) {
    for (const tc::IndexedRecord& rec : records) {
      if (!rec.payload.empty() && rec.payload[0] == kVtsString) {
        for (std::size_t offset = 2; offset < 10U; ++offset) {
          mask.bytes[part][rec.index].insert(offset);
        }
      }
    }
  }
  for (const auto& [part, records] :
       tc::ExtractRecords(excel, {pivot_records::kBeginSXView, pivot_records::kBeginPCDField})) {
    for (const tc::IndexedRecord& rec : records) {
      const bool view = rec.type == pivot_records::kBeginSXView;
      if (!view && (rec.payload.size() < 4U || (rec.payload[2] | rec.payload[3] << 8U) < kFirstCustomNumFmt)) {
        continue;
      }
      const std::size_t first = view ? 28U : 2U;
      const std::size_t width = view ? 4U : 2U;
      for (std::size_t offset = first; offset < first + width; ++offset) {
        mask.bytes[part][rec.index].insert(offset);
      }
    }
  }
  return mask;
}

class XlsbParityPivotApiBuilt : public ::testing::TestWithParam<const char*> {};

TEST_P(XlsbParityPivotApiBuilt, WritesExcelsPivotRecords) {
  const std::string name = GetParam();
  const Saved saved = SaveCase(name);
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(saved.deferred, 0U);
  ExpectValid(saved.bytes);
  EXPECT_EQ(ReloadedReport(saved.bytes), saved.report);

  const tc::Bytes excel =
      tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/pivot/api_built/" + name + ".xlsb");
  ASSERT_FALSE(excel.empty());
  // Excel's save, which names no refreshedBy, reads back with the same pivots; its cells are an older projection.
  EXPECT_EQ(PivotNames(ReloadedReport(excel)), PivotNames(saved.report));
  std::string diff;
  for (const std::string& line : tc::CompareRecordSets(saved.bytes, excel, PivotTypesButUid(excel), Unmatched(excel))) {
    if (line.find("xl/pivot") != std::string::npos) {
      diff += "  " + line + "\n";
    }
  }
  EXPECT_TRUE(diff.empty()) << diff;
}

INSTANTIATE_TEST_SUITE_P(Cases, XlsbParityPivotApiBuilt,
                         ::testing::Values("a_plain", "b_page_filter", "e_shared_cache"));

// The date group's cache refreshes on load, so Excel's save is a refreshed state with its own group labels; the
// group and range records still hold the base field, the parent field and the months and years periods as Excel's.
TEST(XlsbParityPivotApiBuiltGroups, WritesExcelsDateGroupRecords) {
  const Saved saved = SaveCase("c_date_group");
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(saved.deferred, 0U);
  ExpectValid(saved.bytes);
  EXPECT_EQ(ReloadedReport(saved.bytes), saved.report);

  const tc::Bytes excel =
      tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/pivot/api_built/c_date_group.xlsb");
  ASSERT_FALSE(excel.empty());
  const std::set<std::uint16_t> types = {pivot_records::kBeginPCDFGroup, pivot_records::kBeginPCDFGRange};
  std::string diff;
  for (const std::string& line : tc::CompareRecordSets(saved.bytes, excel, types, tc::ByteMask{})) {
    if (line.find("xl/pivot") != std::string::npos) {
      diff += "  " + line + "\n";
    }
  }
  EXPECT_TRUE(diff.empty()) << diff;
  const tc::RecordSet ours = tc::ExtractRecords(saved.bytes, types);
  ASSERT_EQ(ours.count("xl/pivotCache/pivotCacheDefinition1.bin"), 1U);
  EXPECT_EQ(ours.at("xl/pivotCache/pivotCacheDefinition1.bin").size(), 5U);
}

// A calculated cache field has no record layout here, so its cache group is left out whole and counted.
TEST(XlsbParityPivotApiBuiltDeferred, CalculatedFieldLeavesItsGroupOut) {
  const Saved saved = SaveCase("d_calc_field");
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(saved.deferred, 1U);
  ExpectValid(saved.bytes);
  for (const std::string& part : tc::ListParts(saved.bytes)) {
    EXPECT_EQ(part.find("xl/pivot"), std::string::npos) << part;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
