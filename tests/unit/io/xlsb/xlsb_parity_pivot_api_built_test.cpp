// PivotTables built through the C API and saved as `.xlsb`: the cases of tests/fixtures/excel/xlsb_parity/pivot/
// api_built, each built as the workbook Excel 16.115 opened from the engine's `.xlsx` save and re-saved as `.xlsb`
// (`<case>.xlsb`). The save keeps every pivot part, passes the package validator, reads back to the same tables, and
// its pivot records match Excel's but for the revision uid Excel adds. A pivot whose state has no measured layout
// (a calculated field, a quarter or numeric range group, a rank or percent-of-running-total display) is left out
// and counted.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <set>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/read_options.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/pivot_records_common.h"
#include "io/xlsb/record_ids.h"
#include "support/xlsb_package_validator.h"
#include "utils/structured_log.h"
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
  } else if (name == "f_quarter_group") {
    const Placed p = AddPivot(wb, "PF", cache, "PivotF");
    Axis(wb, p, 2U, FM_PIVOT_AXIS_ROW);
    fm_pivot_date_group_t group{};
    group.struct_size = sizeof(group);
    group.periods = FM_PIVOT_PERIOD_YEARS | FM_PIVOT_PERIOD_QUARTERS;
    group.calendar = FM_PIVOT_CALENDAR_GREGORIAN;
    group.interval_days = 1U;
    EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb, p.sheet, p.pivot, 2U, &group), 0) << fm_get_last_error_message();
    SumOf(wb, p, 3U);
    Project(wb, p);
  } else if (name == "g_numeric_group") {
    const Placed p = AddPivot(wb, "PG", cache, "PivotG");
    Axis(wb, p, 3U, FM_PIVOT_AXIS_ROW);
    fm_pivot_numeric_group_t group{};
    group.struct_size = sizeof(group);
    group.interval = 50.0;
    EXPECT_EQ(fm_sheet_set_pivot_field_numeric_group(wb, p.sheet, p.pivot, 3U, &group), 0)
        << fm_get_last_error_message();
    SumOf(wb, p, 4U);
    Project(wb, p);
  } else if (name == "h_rank" || name == "i_running_percent") {
    const Placed p = AddPivot(wb, name == "h_rank" ? "PH" : "PI", cache, name == "h_rank" ? "PivotH" : "PivotI");
    Axis(wb, p, 0U, FM_PIVOT_AXIS_ROW);
    fm_pivot_data_field_t data{};
    data.struct_size = sizeof(data);
    data.field_index = 3U;
    data.aggregation = FM_PIVOT_AGGREGATION_SUM;
    data.show_as =
        name == "h_rank" ? FM_PIVOT_SHOW_VALUES_AS_RANK_DESCENDING : FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_RUNNING_TOTAL;
    data.has_show_as_base_field = 1;
    data.show_as_base_field = 0U;
    std::size_t index = 0;
    EXPECT_EQ(fm_sheet_add_pivot_data_field(wb, p.sheet, p.pivot, &data, &index), 0) << fm_get_last_error_message();
    Project(wb, p);
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

/// What a comparison ignores: the union bytes Excel leaves unset before a string filter operand, the four unset bytes
/// ending a cache definition with no refreshedBy, which ours leaves out, and the ids Excel renumbers on load, the
/// table's idCache (BrtBeginSXView bytes 28-31) and a custom number format's id (BrtBeginPCDField bytes 2-3).
tc::ByteMask Unmatched(const tc::Bytes& excel) {
  constexpr std::uint16_t kFirstCustomNumFmt = 164;
  constexpr std::uint8_t kVtsString = 0x06;
  constexpr std::size_t kRefreshedByFlagByte = 16;
  constexpr std::size_t kUnsetTail = 4;
  tc::ByteMask mask;
  for (const auto& [part, records] : tc::ExtractRecords(excel, {pivot_records::kBeginPivotCacheDef})) {
    for (const tc::IndexedRecord& rec : records) {
      if (rec.payload.size() > kRefreshedByFlagByte + kUnsetTail && (rec.payload[kRefreshedByFlagByte] & 0x01U) == 0U) {
        mask.unset_tails[{part, rec.index}] = kUnsetTail;
      }
    }
  }
  // Past its first two bytes, the union before a string operand is unset memory in Excel's save.
  for (const auto& [part, records] : tc::ExtractRecords(excel, {kBrtCustomFilter})) {
    for (const tc::IndexedRecord& rec : records) {
      if (!rec.payload.empty() && rec.payload[0] == kVtsString) {
        for (std::size_t offset = 4; offset < 10U; ++offset) {
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

// A source row left wholly or partly blank (xlsb_parity/pivot/blank, Excel 16.115's saves after a refresh): the
// .xlsx read and written as .xlsb holds Excel's own pivot records, the blank cells and missing items included.
class XlsbParityPivotBlank : public ::testing::TestWithParam<const char*> {};

TEST_P(XlsbParityPivotBlank, WritesExcelsPivotRecords) {
  const std::string stem = GetParam();
  ReadOptions mac;
  mac.profile = mac_365_en_us_profile();
  const tc::Converted ours =
      tc::ConvertToXlsb(std::string(FORMULON_FIXTURES_DIR), "excel/xlsb_parity/pivot/blank/" + stem + ".xlsx", mac);
  ASSERT_TRUE(ours.error.empty()) << ours.error;
  const tc::Bytes excel =
      tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/pivot/blank/" + stem + ".xlsb");
  ASSERT_FALSE(excel.empty());
  std::string diff;
  for (const std::string& line : tc::CompareRecordSets(ours.bytes, excel, PivotTypesButUid(excel), Unmatched(excel))) {
    if (line.find("xl/pivot") != std::string::npos) {
      diff += "  " + line + "\n";
    }
  }
  EXPECT_TRUE(diff.empty()) << diff;
}

INSTANTIATE_TEST_SUITE_P(Blanks, XlsbParityPivotBlank,
                         ::testing::Values("blank_control", "blank_whole_row", "blank_row_field_only",
                                           "blank_value_only", "blank_date_not_on_axis", "blank_control_date_rows",
                                           "blank_date_row_field"));

// The date group's cache refreshes on load, so Excel's save is a refreshed state with its own group labels; the
// group and range records still hold the base field, the parent field and the months and years periods as Excel's.
/// The record types of every field group block in the package's first cache definition, in stream order.
std::vector<std::uint16_t> GroupBlockTypes(const tc::Bytes& package) {
  std::vector<std::uint16_t> types;
  const auto part = tc::ReadPart(package, "xl/pivotCache/pivotCacheDefinition1.bin");
  if (!part) {
    return types;
  }
  bool in_group = false;
  for (const tc::IndexedRecord& rec : tc::SplitPart(*part)) {
    in_group = in_group || rec.type == pivot_records::kBeginPCDFGroup;
    if (in_group) {
      types.push_back(rec.type);
    }
    in_group = in_group && rec.type != pivot_records::kEndPCDFGroup;
  }
  return types;
}

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
  // The item texts follow the locale Excel refreshed under; the item lists' records do not: the 14 months as one
  // run, and the one-year years group's 3 items as single string records.
  const std::vector<std::uint16_t> excel_types = GroupBlockTypes(excel);
  ASSERT_FALSE(excel_types.empty());
  EXPECT_EQ(GroupBlockTypes(saved.bytes), excel_types);
}

void AppendRecord(std::string_view record, void* log) {
  static_cast<std::string*>(log)->append(record);
}

/// A pivot on PV over Sheet1!K1:L4 whose Amt column holds `second` in row 2 (blank when empty) and 1 and 2 below it.
/// Returns the xlsb save's deferred count and the warn log written while saving.
std::pair<std::size_t, std::string> SaveRefreshedOver(const char* second) {
  fm_workbook_t* wb = nullptr;
  EXPECT_EQ(fm_workbook_create_for_profile(kProfile, &wb), 0);
  EXPECT_EQ(fm_sheet_set_text(wb, 0U, 0U, 10U, "Item"), 0);
  EXPECT_EQ(fm_sheet_set_text(wb, 0U, 0U, 11U, "Amt"), 0);
  EXPECT_EQ(fm_sheet_set_text(wb, 0U, 1U, 10U, "C"), 0);
  if (*second != '\0') {
    EXPECT_EQ(fm_sheet_set_text(wb, 0U, 1U, 11U, second), 0);
  }
  EXPECT_EQ(fm_sheet_set_text(wb, 0U, 2U, 10U, "A"), 0);
  EXPECT_EQ(fm_sheet_set_number(wb, 0U, 2U, 11U, 1.0), 0);
  EXPECT_EQ(fm_sheet_set_text(wb, 0U, 3U, 10U, "B"), 0);
  EXPECT_EQ(fm_sheet_set_number(wb, 0U, 3U, 11U, 2.0), 0);
  std::uint32_t cache_id = 0;
  EXPECT_EQ(fm_workbook_add_pivot_cache(wb, 0U, &cache_id), 0) << fm_get_last_error_message();
  fm_pivot_worksheet_source_t source{sizeof(fm_pivot_worksheet_source_t), "K1:L4", "Sheet1", nullptr};
  EXPECT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb, cache_id, &source), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_workbook_refresh_pivot_cache(wb, cache_id), 0) << fm_get_last_error_message();
  const Placed p = AddPivot(wb, "PV", cache_id, "PivotS");
  Axis(wb, p, 0U, FM_PIVOT_AXIS_ROW);
  SumOf(wb, p, 1U);
  Project(wb, p);
  std::string log;
  set_structured_log_sink(&AppendRecord, &log);
  set_structured_log_min_level(StructuredLogLevel::kWarn);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  fm_save_diagnostics_info_t info{};
  info.struct_size = sizeof(info);
  EXPECT_EQ(fm_workbook_save_with_diagnostics(wb, FM_WORKBOOK_FORMAT_XLSB, &bytes, &len, &info), 0);
  set_structured_log_min_level(StructuredLogLevel::kOff);
  set_structured_log_sink(nullptr);
  fm_buffer_destroy(bytes);
  fm_workbook_destroy(wb);
  return {info.deferred_feature_count, log};
}

// A blank source cell in a field with no shared items is written as Excel writes it, so the pivot is not left out.
TEST(XlsbParityPivotApiBuiltBlank, ABlankInAnInlineFieldIsWritten) {
  const auto [deferred, log] = SaveRefreshedOver("");
  EXPECT_EQ(deferred, 0U) << log;
}

// Text and numbers in one field with no shared items have no measured record cells yet: the group is left out,
// counted, and logged with that reason.
TEST(XlsbParityPivotApiBuiltDeferred, MixedTypesInAnInlineFieldAreLeftOutWithTheirReason) {
  const auto [deferred, log] = SaveRefreshedOver("x");
  EXPECT_EQ(deferred, 1U);
  EXPECT_NE(log.find("unmeasured xlsb pivot cache state: record cells of a field without shared items holding blanks "
                     "or more than one value type"),
            std::string::npos)
      << log;
}

/// The xlsx cache definition of `wb`'s first cache.
std::string CacheDefinitionXml(fm_workbook_t* wb) {
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  EXPECT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len), 0) << fm_get_last_error_message();
  const tc::Bytes package(bytes, bytes + len);
  fm_buffer_destroy(bytes);
  const std::optional<tc::Bytes> part = tc::ReadPart(package, "xl/pivotCache/pivotCacheDefinition1.xml");
  return part.has_value() ? std::string(part->begin(), part->end()) : std::string();
}

// Adding, editing or removing a calculated field leaves the records what the refresh read, so the cache built through
// the API saves without refreshOnLoad.
TEST(XlsbParityPivotApiBuiltCalculatedField, FieldEditsKeepTheCacheRefreshed) {
  fm_workbook_t* wb = nullptr;
  ASSERT_EQ(fm_workbook_create_for_profile(kProfile, &wb), 0);
  BuildCase("d_calc_field", wb);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_id_at(wb, 0U, &cache_id), 0);
  constexpr std::size_t kMargin = 5;
  std::string xml = CacheDefinitionXml(wb);
  EXPECT_NE(xml.find("name=\"Margin\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("refreshOnLoad"), std::string::npos) << xml;
  ASSERT_EQ(fm_workbook_set_pivot_cache_calculated_field_formula(wb, cache_id, kMargin, "=Amt*0.2"), 0)
      << fm_get_last_error_message();
  xml = CacheDefinitionXml(wb);
  EXPECT_EQ(xml.find("refreshOnLoad"), std::string::npos) << xml;
  ASSERT_EQ(fm_workbook_remove_pivot_cache_calculated_field(wb, cache_id, kMargin), 0) << fm_get_last_error_message();
  xml = CacheDefinitionXml(wb);
  EXPECT_EQ(xml.find("name=\"Margin\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("refreshOnLoad"), std::string::npos) << xml;
  fm_workbook_destroy(wb);
}

/// The pivot table parts of a saved package.
std::size_t PivotTableParts(const std::vector<std::uint8_t>& bytes) {
  std::size_t out = 0;
  for (const std::string& part : tc::ListParts(bytes)) {
    out += part.rfind("xl/pivotTables/pivotTable", 0) == 0U ? 1U : 0U;
  }
  return out;
}

/// The xlsb save of case `name` after the engine's own xlsx save of it is read back.
Saved SaveThroughXlsx(const std::string& name) {
  Saved out;
  fm_workbook_t* built = nullptr;
  EXPECT_EQ(fm_workbook_create_for_profile(kProfile, &built), 0) << fm_get_last_error_message();
  BuildCase(name, built);
  out.report = Report(built);
  std::uint8_t* xlsx = nullptr;
  std::size_t xlsx_len = 0;
  EXPECT_EQ(fm_workbook_save(built, FM_WORKBOOK_FORMAT_XLSX, &xlsx, &xlsx_len), 0) << fm_get_last_error_message();
  fm_workbook_destroy(built);
  fm_load_options_spec_t options{sizeof(fm_load_options_spec_t), kProfile};
  fm_workbook_t* wb = nullptr;
  EXPECT_EQ(fm_workbook_load(xlsx, xlsx_len, &options, &wb), 0) << fm_get_last_error_message();
  fm_buffer_destroy(xlsx);
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

/// The BrtBeginPivotCacheDef payloads of a package.
std::vector<tc::Bytes> CacheDefinitions(const tc::Bytes& package) {
  std::vector<tc::Bytes> out;
  for (const auto& [part, records] : tc::ExtractRecords(package, {pivot_records::kBeginPivotCacheDef})) {
    for (const tc::IndexedRecord& rec : records) {
      out.push_back(rec.payload);
    }
  }
  return out;
}

// A pivot reopened from the engine's xlsx save, whose cache no refresh dated, is written to xlsb as Excel saves that
// xlsx without a refresh: the version triple, refreshedDate 0 and no refreshedBy.
TEST(XlsbParityPivotApiBuiltReopened, WritesTheCacheAsExcelSavesItWithoutARefresh) {
  for (const char* name : {"a_plain", "b_page_filter"}) {
    const Saved saved = SaveThroughXlsx(name);
    ASSERT_FALSE(saved.bytes.empty()) << name;
    EXPECT_EQ(saved.deferred, 0U) << name;
    ExpectValid(saved.bytes);
    EXPECT_EQ(PivotNames(ReloadedReport(saved.bytes)), PivotNames(saved.report)) << name;
    const tc::Bytes excel =
        tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/pivot/api_built/" + name + ".xlsb");
    const std::vector<tc::Bytes> ours = CacheDefinitions(saved.bytes);
    const std::vector<tc::Bytes> theirs = CacheDefinitions(excel);
    ASSERT_EQ(ours.size(), 1U) << name;
    ASSERT_EQ(theirs.size(), 1U) << name;
    ASSERT_GE(ours[0].size(), 17U);
    ASSERT_GE(theirs[0].size(), 17U);
    // Bytes 0-2 the version triple, 8-15 refreshedDate, 16 the flag naming a refreshedBy.
    for (const std::size_t offset : {0U, 1U, 2U, 8U, 9U, 10U, 11U, 12U, 13U, 14U, 15U, 16U}) {
      EXPECT_EQ(ours[0][offset], theirs[0][offset]) << name << " byte " << offset;
    }
  }
}

// A pivot whose state has no measured record layout is left out with its cache group, and counted.
class XlsbParityPivotApiBuiltDeferred : public ::testing::TestWithParam<const char*> {};

TEST_P(XlsbParityPivotApiBuiltDeferred, LeavesItsGroupOutAndCountsIt) {
  const Saved saved = SaveCase(GetParam());
  ASSERT_FALSE(saved.bytes.empty());
  EXPECT_EQ(saved.deferred, 1U);
  ExpectValid(saved.bytes);
  for (const std::string& part : tc::ListParts(saved.bytes)) {
    EXPECT_EQ(part.find("xl/pivot"), std::string::npos) << part;
  }
}

INSTANTIATE_TEST_SUITE_P(Causes, XlsbParityPivotApiBuiltDeferred,
                         ::testing::Values("d_calc_field", "f_quarter_group", "g_numeric_group", "h_rank",
                                           "i_running_percent"));

// Every pivot the model holds is either written or counted as left out.
TEST(XlsbParityPivotApiBuiltCount, EveryPivotIsWrittenOrCounted) {
  for (const char* name : {"a_plain", "b_page_filter", "c_date_group", "d_calc_field", "e_shared_cache",
                           "f_quarter_group", "g_numeric_group", "h_rank", "i_running_percent"}) {
    const Saved saved = SaveCase(name);
    ASSERT_FALSE(saved.bytes.empty()) << name;
    EXPECT_EQ(PivotTableParts(saved.bytes) + saved.deferred, PivotNames(saved.report).size()) << name;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
