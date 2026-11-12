// Engine ceilings that are not Excel limits fail with
// `FM_STATUS_SECURITY_RESOURCE_LIMIT` and a context that leads with
// `part=<name> requested=<n> limit=<n>`.

#include "formulon_c_resource_limit.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

TEST(FormulonCApiResourceLimit, RangeCountPastTheCallCapNamesTheRangesPart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t rect{0, 0, 0, 0};
  EXPECT_EQ(ExpectResourceLimit(fm_sheet_set_print_area(wb.handle, 0, &rect, 16385U), "ranges"),
            "part=ranges requested=16385 limit=16384");
}

TEST(FormulonCApiResourceLimit, OversizedPrintFragmentNamesItsMember) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string header_footer = "<headerFooter>" + std::string(70000, 'x') + "</headerFooter>";
  EXPECT_EQ(
      ExpectResourceLimit(fm_sheet_set_header_footer_xml(wb.handle, 0, header_footer.c_str()), "header_footer_xml"),
      "part=header_footer_xml requested=" + std::to_string(header_footer.size()) + " limit=65536");
  const std::string sheet_pr = "<sheetPr>" + std::string(9000, 'x') + "</sheetPr>";
  EXPECT_EQ(ExpectResourceLimit(fm_sheet_set_sheet_pr_xml(wb.handle, 0, sheet_pr.c_str()), "sheet_pr_xml"),
            "part=sheet_pr_xml requested=" + std::to_string(sheet_pr.size()) + " limit=8192");
}

TEST(FormulonCApiResourceLimit, PictureOver32MiBNamesTheImageBytesPart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::size_t len = 32U * 1024U * 1024U + 1U;
  const std::vector<std::uint8_t> bytes(len, 0U);
  fm_drawing_placement_spec_t opts = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  opts.anchor.kind = FM_ANCHOR_KIND_ONE_CELL;
  opts.anchor.edit_as = FM_ANCHOR_EDIT_AS_TWO_CELL;
  uint32_t id = 0;
  EXPECT_EQ(
      ExpectResourceLimit(fm_sheet_add_picture(wb.handle, 0, bytes.data(), bytes.size(), &opts, &id), "image_bytes"),
      "part=image_bytes requested=33554433 limit=33554432");
}

TEST(FormulonCApiResourceLimit, LocationPathPast1024BytesNamesTheLocationPathPart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string directory = "/" + std::string(1017, 'd');
  fm_workbook_location_t record = FM_STRUCT_INIT(fm_workbook_location_t);
  record.directory = directory.c_str();
  record.file_name = "b.xlsx";
  EXPECT_EQ(ExpectResourceLimit(fm_workbook_set_location(wb.handle, &record), "location_path"),
            "part=location_path requested=1025 limit=1024");
}

TEST(FormulonCApiResourceLimit, GradientStopsPast1024NameTheGradientStopsPart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::vector<fm_gradient_stop_t> stops(1025U);
  fm_fill_t fill{};
  fill.has_gradient = 1;
  fill.gradient_stops = stops.data();
  fill.gradient_stop_count = 1024U;
  uint32_t index = 0;
  ASSERT_EQ(fm_workbook_add_fill(wb.handle, &fill, &index), 0) << fm_get_last_error_message();
  fill.gradient_stop_count = 1025U;
  EXPECT_EQ(ExpectResourceLimit(fm_workbook_add_fill(wb.handle, &fill, &index), "gradient_stops"),
            "part=gradient_stops requested=1025 limit=1024");
  EXPECT_EQ(ExpectResourceLimit(fm_workbook_set_fill(wb.handle, index, &fill), "gradient_stops"),
            "part=gradient_stops requested=1025 limit=1024");
}

TEST(FormulonCApiResourceLimit, ThreadCountPastTheSchedulerCapNamesTheThreadsPart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(ExpectResourceLimit(fm_workbook_recalc_parallel(wb.handle, 9U, nullptr), "threads"),
            "part=threads requested=9 limit=8");
  EXPECT_EQ(fm_workbook_recalc_parallel(wb.handle, 8U, nullptr), 0) << fm_get_last_error_message();
}

}  // namespace
