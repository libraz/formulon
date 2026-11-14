//
// Edge values of the sparkline group add and update: the location is one
// cell, and a source or date-axis reference is one reference or name on any
// sheet, a 3-D span or another workbook, or `#REF!`.

#include <cstddef>
#include <iterator>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

// Sources Excel refuses to open a workbook with.
const char* const kRefusedSources[] = {
    "(Sheet1!E1:E3,Sheet1!E5:E6)", "tblData[Val]", "Sheet1!E1:E6*2", "SUM(Sheet1!E1:E6)", "Sheet1!#REF!",
};
// Shapes no Excel save holds; refused with them.
const char* const kOtherRefusedSources[] = {"Sheet1!E1:E6 Sheet1!E2:E3", "(Sheet1!E1:E6)", "Sheet1!E1:", "1"};
// Sources Excel stores.
const char* const kAcceptedSources[] = {
    "Sheet1!E1:E6",        "=Sheet1!E1:E6",   "E1:E6", "Sheet1!E1", "Data!A1:A6",
    "Sheet1:Sheet2!E1:E6", "[1]Sheet1!E1:E6", "Rates", "#REF!",
};

fm_sparkline_group_t LineGroup() {
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  g.type = FM_SPARKLINE_TYPE_LINE;
  g.line_width_pt = 0.75;
  g.empty_cells_as = FM_BLANKS_AS_GAP;
  return g;
}

size_t GroupCount(const fm_workbook_t* wb) {
  size_t count = 99;
  EXPECT_EQ(fm_sheet_get_sparkline_group_count(wb, 0, &count), 0);
  return count;
}

// Adds a one-sparkline group over Sheet1!E1:E6 at G1 and returns its id.
std::string AddGroup(fm_workbook_t* wb) {
  const fm_sparkline_t items[] = {{"Sheet1!E1:E6", "G1"}};
  const fm_sparkline_group_t g = LineGroup();
  const char* id = nullptr;
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb, 0, &g, items, 1, &id), 0) << fm_get_last_error_message();
  return id != nullptr ? std::string(id) : std::string();
}

std::string FormulaAt(const fm_workbook_t* wb, size_t group, size_t index) {
  fm_sparkline_t sp{};
  EXPECT_EQ(fm_sheet_get_sparkline_at(wb, 0, group, index, &sp), 0);
  return sp.formula != nullptr ? std::string(sp.formula) : std::string();
}

TEST(ApiEdge, SparklineLocationMultiCell) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_group_t g = LineGroup();
  const fm_sparkline_t one[] = {{"Sheet1!E1:E6", "G1"}};
  for (const char* location : {"G1:G2", "G1,G3"}) {
    SCOPED_TRACE(location);
    const fm_sparkline_t items[] = {{"Sheet1!E1:E6", location}};
    EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_id()), kInvalidArgument);
    EXPECT_EQ(GroupCount(wb.handle), 0U);
  }
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, one, 1, ignored_id()), 0) << fm_get_last_error_message();

  const std::string id = AddGroup(wb.handle);
  const fm_sparkline_t wide[] = {{"Sheet1!E1:E6", "H1:H2"}};
  fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  patch.items_op = FM_PATCH_OP_SET;
  patch.items = wide;
  patch.items_count = 1;
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, id.c_str(), &patch), kInvalidArgument);
  fm_sparkline_t sp{};
  ASSERT_EQ(fm_sheet_get_sparkline_at(wb.handle, 0, 1, 0, &sp), 0);
  EXPECT_STREQ(sp.location, "G1");
}

TEST(ApiEdge, SparklineSourceRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string id = AddGroup(wb.handle);
  const fm_sparkline_group_t g = LineGroup();
  std::vector<const char*> sources(std::begin(kRefusedSources), std::end(kRefusedSources));
  sources.insert(sources.end(), std::begin(kOtherRefusedSources), std::end(kOtherRefusedSources));
  for (const char* source : sources) {
    SCOPED_TRACE(source);
    const fm_sparkline_t items[] = {{source, "G2"}};
    EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_id()), kInvalidArgument);
    EXPECT_EQ(std::string(fm_get_last_error_context()).rfind("formula=", 0), 0U) << fm_get_last_error_context();
    EXPECT_EQ(GroupCount(wb.handle), 1U);

    fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
    patch.items_op = FM_PATCH_OP_SET;
    patch.items = items;
    patch.items_count = 1;
    EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, id.c_str(), &patch), kInvalidArgument);
    EXPECT_EQ(FormulaAt(wb.handle, 0, 0), "=Sheet1!E1:E6");
  }
}

TEST(ApiEdge, SparklineSourceAccepted) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_group_t g = LineGroup();
  for (const char* source : kAcceptedSources) {
    SCOPED_TRACE(source);
    const fm_sparkline_t items[] = {{source, "G1"}};
    EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_id()), 0) << fm_get_last_error_message();
  }
  EXPECT_EQ(GroupCount(wb.handle), std::size(kAcceptedSources));
}

TEST(ApiEdge, SparklineDeletedSourceIsStoredAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"#REF!", "G1"}};
  const fm_sparkline_group_t g = LineGroup();
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_id()), 0) << fm_get_last_error_message();
  const std::string before = FormulaAt(wb.handle, 0, 0);
  for (const fm_workbook_format_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    SCOPED_TRACE(format);
    BufferGuard buf;
    ASSERT_EQ(fm_workbook_save(wb.handle, format, &buf.data, &buf.len), 0) << fm_get_last_error_message();
    WorkbookGuard again;
    ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &again.handle), 0) << fm_get_last_error_message();
    ASSERT_EQ(GroupCount(again.handle), 1U);
    EXPECT_EQ(FormulaAt(again.handle, 0, 0), before);
  }
}

TEST(ApiEdge, SparklineDateAxisShape) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_sparkline_t items[] = {{"Sheet1!E1:E6", "G1"}};
  fm_sparkline_group_t g = LineGroup();
  g.date_axis_ref = "(Sheet1!A1:D1,Sheet1!A3:D3)";
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_id()), kInvalidArgument);
  EXPECT_EQ(std::string(fm_get_last_error_context()).rfind("date_axis_ref=", 0), 0U) << fm_get_last_error_context();
  EXPECT_EQ(GroupCount(wb.handle), 0U);
  g.date_axis_ref = "Sheet1:Sheet2!A1:D1";
  EXPECT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, ignored_id()), 0) << fm_get_last_error_message();

  // A deleted date range is stored absent.
  g.date_axis_ref = "#REF!";
  const char* added = nullptr;
  ASSERT_EQ(fm_sheet_add_sparkline_group(wb.handle, 0, &g, items, 1, &added), 0) << fm_get_last_error_message();
  const std::string id = added;
  fm_sparkline_group_t out = FM_STRUCT_INIT(fm_sparkline_group_t);
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 1, &out), 0);
  EXPECT_EQ(out.date_axis_ref, nullptr);

  fm_sparkline_group_patch_t patch = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  patch.date_axis_ref_op = FM_PATCH_OP_SET;
  patch.date_axis_ref = "SUM(Sheet1!A1:D1)";
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, id.c_str(), &patch), kInvalidArgument);
  patch.date_axis_ref = "Sheet1!A1:D1";
  EXPECT_EQ(fm_sheet_update_sparkline_group(wb.handle, 0, id.c_str(), &patch), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_sparkline_group_at(wb.handle, 0, 1, &out), 0);
  ASSERT_NE(out.date_axis_ref, nullptr);
  EXPECT_STREQ(out.date_axis_ref, "=Sheet1!A1:D1");
}

}  // namespace
