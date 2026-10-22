// Stable C ABI: pivot filters and subtotal functions survive save + reload.

#include <cmath>
#include <cstdint>
#include <limits>
#include <set>
#include <string>
#include <string_view>
#include <tuple>
#include <vector>

#include "formulon_c_pivot_test_helpers.h"

namespace {

// Region / Date / Amount over four records:
//   North 45300 100, South 45400 200, East 45500 50, North 45310 300.
// Region sums: North 400, South 200, East 50.
std::vector<std::uint8_t> BuildFilterPackage(std::string_view region_field_attrs, std::string_view filters_xml) {
  const std::string content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      "<Override PartName=\"/xl/pivotCache/pivotCacheDefinition1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"/>"
      "<Override PartName=\"/xl/pivotCache/pivotCacheRecords1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml\"/>"
      "<Override PartName=\"/xl/pivotTables/pivotTable1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\"/>"
      "</Types>";
  const std::string package_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/></Relationships>";
  const std::string workbook_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
      "<pivotCaches><pivotCache cacheId=\"0\" r:id=\"rId2\"/></pivotCaches></workbook>";
  const std::string workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>"
      "<Relationship Id=\"rId2\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" "
      "Target=\"pivotCache/pivotCacheDefinition1.xml\"/></Relationships>";
  const std::string sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>";
  const std::string sheet_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" "
      "Target=\"../pivotTables/pivotTable1.xml\"/></Relationships>";
  const std::string cache_def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" "
      "recordCount=\"4\"><cacheSource type=\"worksheet\"><worksheetSource ref=\"A1:C5\" sheet=\"Sheet1\"/>"
      "</cacheSource><cacheFields count=\"3\">"
      "<cacheField name=\"Region\"><sharedItems count=\"3\"><s v=\"North\"/><s v=\"South\"/><s v=\"East\"/>"
      "</sharedItems></cacheField>"
      "<cacheField name=\"Date\"><sharedItems containsNumber=\"1\"/></cacheField>"
      "<cacheField name=\"Amount\"><sharedItems containsNumber=\"1\"/></cacheField>"
      "</cacheFields></pivotCacheDefinition>";
  const std::string cache_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords\" "
      "Target=\"pivotCacheRecords1.xml\"/></Relationships>";
  const std::string cache_records =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"4\">"
      "<r><x v=\"0\"/><n v=\"45300\"/><n v=\"100\"/></r>"
      "<r><x v=\"1\"/><n v=\"45400\"/><n v=\"200\"/></r>"
      "<r><x v=\"2\"/><n v=\"45500\"/><n v=\"50\"/></r>"
      "<r><x v=\"0\"/><n v=\"45310\"/><n v=\"300\"/></r>"
      "</pivotCacheRecords>";
  std::string table_def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "name=\"PT\" cacheId=\"0\" dataCaption=\"Values\" updatedVersion=\"8\">"
      "<location ref=\"E1:F5\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>"
      "<pivotFields count=\"3\"><pivotField axis=\"axisRow\" showAll=\"0\"";
  table_def.append(region_field_attrs);
  table_def.append(
      "><items count=\"4\"><item x=\"0\"/><item x=\"1\"/><item x=\"2\"/><item t=\"default\"/></items>"
      "</pivotField><pivotField showAll=\"0\"/><pivotField dataField=\"1\" showAll=\"0\"/></pivotFields>"
      "<rowFields count=\"1\"><field x=\"0\"/></rowFields>"
      "<dataFields count=\"1\"><dataField name=\"Sum of Amount\" fld=\"2\" baseField=\"0\" baseItem=\"0\"/>"
      "</dataFields>"
      "<pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\" showColHeaders=\"1\" "
      "showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\"/>");
  table_def.append(filters_xml);
  table_def.append("</pivotTableDefinition>");
  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/worksheets/_rels/sheet1.xml.rels", sheet_rels},
      {"xl/pivotCache/pivotCacheDefinition1.xml", cache_def},
      {"xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", cache_rels},
      {"xl/pivotCache/pivotCacheRecords1.xml", cache_records},
      {"xl/pivotTables/pivotTable1.xml", table_def},
  });
}

std::set<std::string> RowLabels(fm_workbook_t* wb) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_workbook_pivot_layout(wb, 0, 0, &projected.handle), 0) << fm_last_error_message();
  std::set<std::string> out;
  for (const fm_pivot_cell_t& cell : CollectCells(projected.handle)) {
    if (cell.kind == FM_PIVOT_CELL_ROW_LABEL && cell.value.kind == FM_VAL_TEXT && cell.value.u.text != nullptr) {
      out.insert(cell.value.u.text);
    }
  }
  return out;
}

std::vector<std::uint8_t> Save(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0) << fm_last_error_message();
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

std::string FiltersBlock(const std::string& xml) {
  const std::size_t begin = xml.find("<filters");
  const std::size_t end = xml.rfind("</filters>");
  if (begin == std::string::npos || end == std::string::npos) {
    return {};
  }
  return xml.substr(begin, end + std::string_view("</filters>").size() - begin);
}

fm_pivot_filter_spec_t TextSpec(fm_pivot_filter_type_t type, const char* field, const char* text) {
  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = field;
  spec.type = type;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = text;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  return spec;
}

fm_pivot_filter_spec_t NumberSpec(fm_pivot_filter_type_t type, const char* field, double low, double high) {
  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = field;
  spec.type = type;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  spec.value_double = low;
  spec.value_high_kind = std::isnan(high) ? FM_PIVOT_FILTER_VALUE_NONE : FM_PIVOT_FILTER_VALUE_DOUBLE;
  spec.value_high_double = std::isnan(high) ? 0.0 : high;
  return spec;
}

struct FilterCase {
  const char* label;
  fm_pivot_filter_spec_t spec;
  std::set<std::string> kept;
};

}  // namespace

// Each public filter type, added to a loaded pivot, is saved, reads back with
// the same spec, and filters the reloaded report the same way.
TEST(FormulonCApiPivotFilterPersistence, EachFilterTypeSurvivesSaveAndReload) {
  const double kNone = std::numeric_limits<double>::quiet_NaN();
  fm_pivot_filter_spec_t top = NumberSpec(FM_PIVOT_FILTER_VALUE_TOP_10, "Region", 0.0, kNone);
  top.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  top.value_int = 1;
  const FilterCase cases[] = {
      {"contains", TextSpec(FM_PIVOT_FILTER_LABEL_CONTAINS, "Region", "orth"), {"North"}},
      {"begins", TextSpec(FM_PIVOT_FILTER_LABEL_BEGINS_WITH, "Region", "S"), {"South"}},
      {"top", top, {"North"}},
      {"greater", NumberSpec(FM_PIVOT_FILTER_VALUE_GREATER_THAN, "Region", 150.0, kNone), {"North", "South"}},
      {"between", NumberSpec(FM_PIVOT_FILTER_VALUE_BETWEEN, "Region", 100.0, 250.0), {"South"}},
      {"date", NumberSpec(FM_PIVOT_FILTER_LABEL_DATE, "Date", 45390.0, 45410.0), {"South"}},
  };
  const std::vector<std::uint8_t> package = BuildFilterPackage("", "");
  for (const FilterCase& c : cases) {
    SCOPED_TRACE(c.label);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_load(package.data(), package.size(), &wb.handle), 0) << fm_last_error_message();
    ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, 0, &c.spec), 0) << fm_last_error_message();
    EXPECT_EQ(RowLabels(wb.handle), c.kept);

    const std::vector<std::uint8_t> saved = Save(wb.handle);
    const std::string saved_xml = ExtractZipEntry(saved, "xl/pivotTables/pivotTable1.xml");
    EXPECT_NE(saved_xml.find("<filters count=\"1\">"), std::string::npos) << saved_xml;

    WorkbookGuard back;
    ASSERT_EQ(fm_workbook_load(saved.data(), saved.size(), &back.handle), 0) << fm_last_error_message();
    std::size_t count = 0;
    ASSERT_EQ(fm_workbook_pivot_filter_count(back.handle, 0, 0, &count), 0);
    ASSERT_EQ(count, 1U);
    fm_pivot_filter_spec_t read{};
    ASSERT_EQ(fm_workbook_pivot_filter_at(back.handle, 0, 0, 0, &read), 0) << fm_last_error_message();
    EXPECT_EQ(read.type, c.spec.type);
    EXPECT_STREQ(read.field_name, c.spec.field_name);
    EXPECT_EQ(read.axis, FM_PIVOT_AXIS_ROW);
    EXPECT_EQ(read.value_kind, c.spec.value_kind);
    EXPECT_EQ(read.value_int, c.spec.value_int);
    EXPECT_DOUBLE_EQ(read.value_double, c.spec.value_double);
    if (c.spec.value_kind == FM_PIVOT_FILTER_VALUE_TEXT) {
      EXPECT_STREQ(read.value_text, c.spec.value_text);
    }
    EXPECT_EQ(read.value_high_kind, c.spec.value_high_kind);
    EXPECT_DOUBLE_EQ(read.value_high_double, c.spec.value_high_double);
    EXPECT_EQ(read.data_field_index, c.spec.data_field_index);
    EXPECT_EQ(RowLabels(back.handle), c.kept);
    // A second save re-emits the decoded entry byte for byte.
    EXPECT_EQ(FiltersBlock(ExtractZipEntry(Save(back.handle), "xl/pivotTables/pivotTable1.xml")),
              FiltersBlock(saved_xml));
  }
}

// A pivot assembled entirely through the C ABI persists its filter too.
TEST(FormulonCApiPivotFilterPersistence, ScratchPivotFilterSurvivesSaveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  const fm_pivot_filter_spec_t spec = TextSpec(FM_PIVOT_FILTER_LABEL_BEGINS_WITH, "Region", "N");
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> saved = Save(wb.handle);

  WorkbookGuard back;
  ASSERT_EQ(fm_workbook_load(saved.data(), saved.size(), &back.handle), 0) << fm_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_count(back.handle, 0, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  EXPECT_TRUE(LayoutHasText(back.handle, 0, "North"));
  EXPECT_FALSE(LayoutHasText(back.handle, 0, "South"));
}

namespace {

constexpr std::string_view kMixedFilters =
    "<filters count=\"4\">"
    "<filter fld=\"0\" type=\"captionNotEqual\" evalOrder=\"-1\" id=\"1\" stringValue1=\"East\"><autoFilter "
    "ref=\"A1\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"notEqual\" val=\"East\"/>"
    "</customFilters></filterColumn></autoFilter></filter>"
    "<filter fld=\"0\" type=\"count\" evalOrder=\"-1\" id=\"2\" iMeasureFld=\"0\"><autoFilter ref=\"A1\">"
    "<filterColumn colId=\"0\"><top10 val=\"2\" filterVal=\"2\"/></filterColumn></autoFilter></filter>"
    "<filter fld=\"0\" type=\"percent\" evalOrder=\"-1\" id=\"3\" iMeasureFld=\"0\"><autoFilter ref=\"A1\">"
    "<filterColumn colId=\"0\"><top10 percent=\"1\" val=\"50\" filterVal=\"50\"/></filterColumn></autoFilter>"
    "</filter>"
    "<filter fld=\"1\" type=\"someFutureFamily\" evalOrder=\"-1\" id=\"4\"><autoFilter ref=\"A1\"/></filter>"
    "</filters>";

}  // namespace

// Entries the C ABI cannot list are kept and evaluated, and the whole block
// is saved byte for byte; editing the listed entry leaves them in place.
TEST(FormulonCApiPivotFilterPersistence, LoadedUnsupportedFiltersAreKeptVerbatim) {
  const std::vector<std::uint8_t> package = BuildFilterPackage("", kMixedFilters);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(package.data(), package.size(), &wb.handle), 0) << fm_last_error_message();
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_pivot_filter_spec_t read{};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, 0, 0, &read), 0) << fm_last_error_message();
  EXPECT_EQ(read.type, FM_PIVOT_FILTER_VALUE_TOP_10);
  EXPECT_STREQ(read.field_name, "Region");
  EXPECT_EQ(read.value_kind, FM_PIVOT_FILTER_VALUE_INT);
  EXPECT_EQ(read.value_int, 2);
  // not-East, then top 2 of North 400 / South 200, then the top 50% of 600.
  const std::set<std::string> kept = {"North"};
  EXPECT_EQ(RowLabels(wb.handle), kept);

  const std::vector<std::uint8_t> saved = Save(wb.handle);
  EXPECT_EQ(FiltersBlock(ExtractZipEntry(saved, "xl/pivotTables/pivotTable1.xml")), std::string(kMixedFilters));
  WorkbookGuard back;
  ASSERT_EQ(fm_workbook_load(saved.data(), saved.size(), &back.handle), 0) << fm_last_error_message();
  EXPECT_EQ(RowLabels(back.handle), kept);

  ASSERT_EQ(fm_workbook_pivot_filter_remove_at(back.handle, 0, 0, 0), 0) << fm_last_error_message();
  const std::string edited = FiltersBlock(SavedPivotXml(back.handle));
  EXPECT_EQ(edited.rfind("<filters count=\"3\">", 0), 0U) << edited;
  EXPECT_EQ(edited.find("type=\"count\""), std::string::npos) << edited;
  EXPECT_LT(edited.find("captionNotEqual"), edited.find("percent"));
  EXPECT_LT(edited.find("percent"), edited.find("someFutureFamily"));
  ASSERT_EQ(fm_workbook_pivot_filter_clear(back.handle, 0, 0), 0);
  EXPECT_EQ(FiltersBlock(SavedPivotXml(back.handle)), edited);
}

// The subtotal selection reads back before a save, after one, and from a
// file that carries it.
TEST(FormulonCApiPivotFilterPersistence, SubtotalFunctionsReadBack) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  fm_pivot_field_info_t info{};
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot_idx, 0, &info), 0);
  EXPECT_EQ(info.default_subtotal, 1);
  EXPECT_EQ(info.subtotal_fns, 0U);

  ASSERT_EQ(fm_workbook_pivot_field_add_subtotal_fn(wb.handle, 0, pivot_idx, 0, FM_PIVOT_AGG_COUNT), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_subtotal_fn(wb.handle, 0, pivot_idx, 0, FM_PIVOT_AGG_SUM), 0);
  ASSERT_EQ(fm_workbook_pivot_field_at(wb.handle, 0, pivot_idx, 0, &info), 0);
  const std::uint32_t sum_count = (1U << FM_PIVOT_AGG_SUM) | (1U << FM_PIVOT_AGG_COUNT);
  EXPECT_EQ(info.default_subtotal, 0);
  EXPECT_EQ(info.subtotal_fns, sum_count);

  const std::vector<std::uint8_t> saved = Save(wb.handle);
  WorkbookGuard back;
  ASSERT_EQ(fm_workbook_load(saved.data(), saved.size(), &back.handle), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_field_at(back.handle, 0, 0, 0, &info), 0);
  EXPECT_EQ(info.default_subtotal, 0);
  EXPECT_EQ(info.subtotal_fns, sum_count);
  // The value field carries the default.
  ASSERT_EQ(fm_workbook_pivot_field_at(back.handle, 0, 0, 1, &info), 0);
  EXPECT_EQ(info.default_subtotal, 1);
  EXPECT_EQ(info.subtotal_fns, 0U);

  for (const auto& [attrs, want_default, want_fns] :
       {std::tuple<std::string_view, std::int32_t, std::uint32_t>{" defaultSubtotal=\"0\"", 0, 0U},
        std::tuple<std::string_view, std::int32_t, std::uint32_t>{
            " avgSubtotal=\"1\" maxSubtotal=\"1\"", 0, (1U << FM_PIVOT_AGG_AVERAGE) | (1U << FM_PIVOT_AGG_MAX)}}) {
    SCOPED_TRACE(std::string(attrs));
    const std::vector<std::uint8_t> package = BuildFilterPackage(attrs, "");
    WorkbookGuard loaded;
    ASSERT_EQ(fm_workbook_load(package.data(), package.size(), &loaded.handle), 0) << fm_last_error_message();
    ASSERT_EQ(fm_workbook_pivot_field_at(loaded.handle, 0, 0, 0, &info), 0);
    EXPECT_EQ(info.default_subtotal, want_default);
    EXPECT_EQ(info.subtotal_fns, want_fns);
  }
}
