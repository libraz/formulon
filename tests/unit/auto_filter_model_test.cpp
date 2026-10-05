#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

constexpr std::string_view kFilterDatabase = "_xlnm._FilterDatabase";

// The structure-edit setup measured in Excel: header row 2 of B2:F12, filters on c2
// (colId 1, value `1`) and c4 (colId 3, value `2`).
constexpr const char* kProbeFilter =
    "<autoFilter ref=\"B2:F12\"><filterColumn colId=\"1\"><filters><filter val=\"1\"/></filters></filterColumn>"
    "<filterColumn colId=\"3\"><filters><filter val=\"2\"/></filters></filterColumn></autoFilter>";

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(file);
  return out;
}

std::vector<std::uint8_t> FixtureBytes() {
  return ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/auto_filter_kinds.xlsx");
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto result = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>();
}

/// Every `<autoFilter>` element in the package's worksheet parts, as written.
std::vector<std::string> WorksheetAutoFilters(const std::vector<std::uint8_t>& package) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{package.data(), package.size()})));
  std::vector<std::string> out;
  for (const std::string& name : zip.list_entries()) {
    if (name.rfind("xl/worksheets/sheet", 0) != 0) {
      continue;
    }
    auto bytes = zip.read_entry(name);
    EXPECT_TRUE(static_cast<bool>(bytes));
    const std::string xml(bytes.value().begin(), bytes.value().end());
    const std::size_t begin = xml.find("<autoFilter");
    if (begin == std::string::npos) {
      continue;
    }
    const std::size_t tag_end = xml.find('>', begin);
    const std::size_t end = xml[tag_end - 1U] == '/'
                                ? tag_end + 1U
                                : xml.find("</autoFilter>", begin) + std::string("</autoFilter>").size();
    out.push_back(xml.substr(begin, end - begin));
  }
  std::sort(out.begin(), out.end());
  return out;
}

std::vector<std::string> ModelAutoFilters(const Workbook& wb) {
  std::vector<std::string> out;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (wb.sheet(i).has_auto_filter()) {
      out.push_back(wb.sheet(i).auto_filter_xml());
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

const AutoFilter& SheetFilter(const Workbook& wb, std::string_view sheet_name) {
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (wb.sheet(i).name() == sheet_name && wb.sheet(i).auto_filter() != nullptr) {
      return *wb.sheet(i).auto_filter();
    }
  }
  ADD_FAILURE() << "no AutoFilter on sheet " << sheet_name;
  static const AutoFilter kEmpty;
  return kEmpty;
}

/// `B2:E12; colId 2: 2` — the shape of the measured expectation table.
std::string Describe(const AutoFilter* filter) {
  if (filter == nullptr) {
    return "none";
  }
  std::string out;
  append_a1_rectangle(out, filter->range);
  for (const FilterColumn& col : filter->columns) {
    out += "; colId " + std::to_string(col.col_id) + ":";
    for (const std::string& value : col.values.values) {
      out += " " + value;
    }
  }
  return out;
}

Workbook NewWorkbook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  return wb;
}

Workbook ProbeWorkbook() {
  Workbook wb = NewWorkbook();
  wb.sheet(0).set_auto_filter_xml(kProbeFilter);
  return wb;
}

Workbook ProbeTableWorkbook() {
  Workbook wb = NewWorkbook();
  TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.display_name = "Table1";
  table.ref = "B2:F12";
  table.sheet_index = 0;
  for (std::uint32_t i = 1; i <= 5; ++i) {
    table.columns.emplace_back(i, "c" + std::to_string(i), "", "", "");
  }
  table.auto_filter_xml = kProbeFilter;
  wb.mutable_tables().push_back(std::move(table));
  return wb;
}

const DefinedName* FilterDatabaseName(const Workbook& wb, std::int32_t sheet) {
  for (const DefinedName& name : wb.defined_names()) {
    if (name.name == kFilterDatabase && name.local_sheet_id == sheet) {
      return &name;
    }
  }
  return nullptr;
}

TEST(AutoFilterModel, EveryFixtureKindRoundTrips) {
  const std::vector<std::uint8_t> fixture = FixtureBytes();
  const std::vector<std::string> written = WorksheetAutoFilters(fixture);
  ASSERT_EQ(written.size(), 60U);
  const Workbook wb = Load(fixture);
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (const AutoFilter* filter = wb.sheet(i).auto_filter()) {
      EXPECT_FALSE(filter->is_opaque()) << wb.sheet(i).name();
    }
  }
  EXPECT_EQ(ModelAutoFilters(wb), written);

  const std::vector<std::uint8_t> saved = Save(wb);
  EXPECT_EQ(WorksheetAutoFilters(saved), written);
  EXPECT_EQ(ModelAutoFilters(Load(saved)), written);
}

TEST(AutoFilterModel, FixtureKindsParseIntoTypedCriteria) {
  const Workbook wb = Load(FixtureBytes());

  const FilterColumn& values = SheetFilter(wb, "val_multi_a_c").columns.at(0);
  EXPECT_EQ(values.kind, FilterKind::kValues);
  EXPECT_EQ(values.values.values, (std::vector<std::string>{"a", "ab"}));
  EXPECT_TRUE(SheetFilter(wb, "val_blank_eq").columns.at(0).values.blank);

  const FilterColumn& custom = SheetFilter(wb, "cus_and_gt2_lt5").columns.at(0);
  ASSERT_EQ(custom.kind, FilterKind::kCustom);
  EXPECT_TRUE(custom.custom.and_join);
  ASSERT_EQ(custom.custom.filters.size(), 2U);
  EXPECT_EQ(custom.custom.filters[0].op, FilterOperator::kGreaterThan);
  EXPECT_EQ(custom.custom.filters[0].val, "2");
  EXPECT_EQ(custom.custom.filters[1].op, FilterOperator::kLessThan);
  EXPECT_EQ(SheetFilter(wb, "cus_ne_x").columns.at(0).custom.filters.at(0).op, FilterOperator::kNotEqual);

  const FilterColumn& bottom = SheetFilter(wb, "bottom_pct_10_n15").columns.at(0);
  ASSERT_EQ(bottom.kind, FilterKind::kTop10);
  EXPECT_FALSE(bottom.top10.top);
  EXPECT_TRUE(bottom.top10.percent);
  EXPECT_EQ(bottom.top10.val, 10.0);
  EXPECT_EQ(bottom.top10.filter_val, -1.0);

  const FilterColumn& week = SheetFilter(wb, "dyn_thisWeek").columns.at(0);
  ASSERT_EQ(week.kind, FilterKind::kDynamic);
  EXPECT_EQ(week.dynamic.type, DynamicFilterType::kThisWeek);
  EXPECT_EQ(week.dynamic.val, 46299.0);
  EXPECT_EQ(week.dynamic.max_val, 46306.0);
  EXPECT_EQ(SheetFilter(wb, "dyn_Oct").columns.at(0).dynamic.type, DynamicFilterType::kM10);
  EXPECT_FALSE(SheetFilter(wb, "dyn_Q4").columns.at(0).dynamic.val.has_value());

  const FilterColumn& hour = SheetFilter(wb, "dgi_hour_2025_03_03_12").columns.at(0);
  ASSERT_EQ(hour.kind, FilterKind::kValues);
  ASSERT_EQ(hour.values.date_groups.size(), 1U);
  EXPECT_EQ(hour.values.date_groups[0].year, 2025U);
  EXPECT_EQ(hour.values.date_groups[0].month, 3U);
  EXPECT_EQ(hour.values.date_groups[0].day, 3U);
  EXPECT_EQ(hour.values.date_groups[0].hour, 12U);
  EXPECT_EQ(hour.values.date_groups[0].grouping, DateTimeGrouping::kHour);
  EXPECT_TRUE(hour.ext_xml.empty());

  const FilterColumn& font = SheetFilter(wb, "col_font_red").columns.at(0);
  ASSERT_EQ(font.kind, FilterKind::kColor);
  EXPECT_EQ(font.color.dxf_id, 1U);
  EXPECT_FALSE(font.color.cell_color);
  EXPECT_TRUE(SheetFilter(wb, "col_cell_yellow").columns.at(0).color.cell_color);

  // Excel-saved filters carry `xr:uid`, retained as an unmodelled attribute.
  ASSERT_EQ(SheetFilter(wb, "top_items_1").extra_attrs.size(), 1U);
  EXPECT_EQ(SheetFilter(wb, "top_items_1").extra_attrs[0].first, "xr:uid");
}

TEST(AutoFilterModel, ColIdShiftsOnColumnInsert) {
  struct Case {
    std::uint32_t col;
    const char* expected;
  };
  // O5a sheet_insert_col_* rows.
  const Case cases[] = {
      {0U, "C2:G12; colId 1: 1; colId 3: 2"},  // A, outside left
      {1U, "C2:G12; colId 1: 1; colId 3: 2"},  // B, left edge
      {2U, "B2:G12; colId 2: 1; colId 4: 2"},  // C, inside
      {4U, "B2:G12; colId 1: 1; colId 4: 2"},  // E, between filters
      {6U, "B2:F12; colId 1: 1; colId 3: 2"},  // G, outside right
  };
  for (const Case& c : cases) {
    Workbook wb = ProbeWorkbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, c.col, 1U)));
    EXPECT_EQ(Describe(wb.sheet(0).auto_filter()), c.expected) << "insert at column " << c.col;
  }
}

TEST(AutoFilterModel, ColumnDeleteRemovesCriterionAndRenumbers) {
  struct Case {
    std::uint32_t col;
    std::uint32_t count;
    const char* expected;
  };
  // O5a sheet_delete_col_* rows.
  const Case cases[] = {
      {1U, 1U, "B2:E12; colId 0: 1; colId 2: 2"},  // B, first
      {2U, 1U, "B2:E12; colId 2: 2"},              // C, filtered c2
      {3U, 1U, "B2:E12; colId 1: 1; colId 2: 2"},  // D, unfiltered
      {4U, 1U, "B2:E12; colId 1: 1"},              // E, filtered c4
      {5U, 1U, "B2:E12; colId 1: 1; colId 3: 2"},  // F, last
      {2U, 2U, "B2:D12; colId 1: 2"},              // C:D
  };
  for (const Case& c : cases) {
    Workbook wb = ProbeWorkbook();
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, c.col, c.count)));
    EXPECT_EQ(Describe(wb.sheet(0).auto_filter()), c.expected) << "delete at column " << c.col;
  }
}

TEST(AutoFilterModel, RowEditsMoveTheSheetFilter) {
  struct Case {
    bool is_delete;
    std::uint32_t row;
    const char* expected;
  };
  // O5a sheet_*_row_* rows (0-based row indices).
  const Case cases[] = {
      {true, 1U, "none"},                             // header row
      {true, 11U, "B2:F11; colId 1: 1; colId 3: 2"},  // last body row
      {true, 0U, "B1:F11; colId 1: 1; colId 3: 2"},   // above the range
      {true, 4U, "B2:F11; colId 1: 1; colId 3: 2"},   // in the body
      {false, 1U, "B3:F13; colId 1: 1; colId 3: 2"},  // above the header
      {false, 4U, "B2:F13; colId 1: 1; colId 3: 2"},  // in the body
  };
  for (const Case& c : cases) {
    Workbook wb = ProbeWorkbook();
    ASSERT_TRUE(static_cast<bool>(c.is_delete ? wb.delete_rows(0, c.row, 1U) : wb.insert_rows(0, c.row, 1U)));
    EXPECT_EQ(Describe(wb.sheet(0).auto_filter()), c.expected) << (c.is_delete ? "delete" : "insert") << c.row;
  }
}

TEST(AutoFilterModel, HeaderRowDeleteRemovesFilterDatabaseName) {
  Workbook wb = NewWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter_xml(0, kProbeFilter)));
  ASSERT_NE(FilterDatabaseName(wb, 0), nullptr);
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 1U, 1U)));
  EXPECT_FALSE(wb.sheet(0).has_auto_filter());
  EXPECT_EQ(FilterDatabaseName(wb, 0), nullptr);
}

TEST(AutoFilterModel, TableFilterFollowsRowDelete) {
  // O5a tbl_delete_row_in_body: table and filter shrink, colIds stay.
  Workbook wb = ProbeTableWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 4U, 1U)));
  EXPECT_EQ(wb.tables()[0].ref, "B2:F11");
  EXPECT_EQ(Describe(wb.tables()[0].auto_filter_xml.get()), "B2:F11; colId 1: 1; colId 3: 2");

  // O5a tbl_delete_col_C_filtered(c2).
  Workbook cols = ProbeTableWorkbook();
  ASSERT_TRUE(static_cast<bool>(cols.delete_cols(0, 2U, 1U)));
  EXPECT_EQ(Describe(cols.tables()[0].auto_filter_xml.get()), "B2:E12; colId 2: 2");

  // O5a tbl_insert_col_at_C_inside.
  Workbook inserted = ProbeTableWorkbook();
  ASSERT_TRUE(static_cast<bool>(inserted.insert_cols(0, 2U, 1U)));
  EXPECT_EQ(Describe(inserted.tables()[0].auto_filter_xml.get()), "B2:G12; colId 2: 1; colId 4: 2");

  // A table filter is never dropped for its header row.
  Workbook header = ProbeTableWorkbook();
  ASSERT_TRUE(static_cast<bool>(header.delete_rows(0, 1U, 1U)));
  EXPECT_NE(header.tables()[0].auto_filter_xml.get(), nullptr);
  EXPECT_TRUE(header.defined_names().empty());
}

TEST(AutoFilterModel, OpaqueFragmentSurvives) {
  const std::string unknown_type =
      "<autoFilter ref=\"A1:B5\"><filterColumn colId=\"0\"><dynamicFilter type=\"fortnight\"/></filterColumn>"
      "</autoFilter>";
  const std::string qualified_ref = "<autoFilter ref=\"Sheet1!D6:E7\"/>";
  for (const std::string& xml : {unknown_type, qualified_ref}) {
    EXPECT_EQ(parse_auto_filter_xml(xml).error().code, FormulonErrorCode::kAutoFilterInvalid) << xml;
    Workbook wb = NewWorkbook();
    ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter_xml(0, xml)));
    ASSERT_NE(wb.sheet(0).auto_filter(), nullptr);
    EXPECT_TRUE(wb.sheet(0).auto_filter()->is_opaque());
    EXPECT_EQ(wb.sheet(0).auto_filter_xml(), xml);
    EXPECT_EQ(FilterDatabaseName(wb, 0), nullptr);

    ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 0U, 1U)));
    EXPECT_EQ(wb.sheet(0).auto_filter_xml(), xml);
    EXPECT_EQ(Load(Save(wb)).sheet(0).auto_filter_xml(), xml);

    const auto typed = wb.set_sheet_auto_filter(0, *wb.sheet(0).auto_filter());
    ASSERT_FALSE(static_cast<bool>(typed));
    EXPECT_EQ(typed.error().code, FormulonErrorCode::kAutoFilterInvalid);
  }
}

TEST(AutoFilterModel, UnknownColumnContentIsRetainedTyped) {
  const std::string xml =
      "<autoFilter ref=\"A1:C9\" xr:uid=\"{00000000-0000-0000-0000-000000000001}\">"
      "<filterColumn colId=\"2\" hiddenButton=\"1\" x:future=\"7\"><filters><filter val=\"x\"/></filters>"
      "<futureChild a=\"1\"/></filterColumn></autoFilter>";
  auto parsed = parse_auto_filter_xml(xml);
  ASSERT_TRUE(static_cast<bool>(parsed)) << parsed.error().message;
  const FilterColumn& col = parsed.value().columns.at(0);
  EXPECT_TRUE(col.hidden_button);
  ASSERT_EQ(col.extra_attrs.size(), 1U);
  EXPECT_EQ(col.extra_attrs[0].first, "x:future");
  EXPECT_EQ(col.extra_xml, "<futureChild a=\"1\"/>");
  EXPECT_EQ(serialize_auto_filter(parsed.value()), xml);
}

TEST(AutoFilterModel, ExtLstRoundTrips) {
  // Date groups lift out of the richdata2 extension; another extension in
  // the same column and the AutoFilter's own extLst stay verbatim.
  const std::string xml =
      "<autoFilter ref=\"A1:A10\"><filterColumn colId=\"0\"><extLst>"
      "<ext uri=\"{1AD28BCE-077C-4C59-8B6E-1921CE8616D4}\" "
      "xmlns:xlrd2=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\"><xlrd2:filterColumn>"
      "<xlrd2:filters><xlrd2:dateGroupItem year=\"2025\" month=\"3\" dateTimeGrouping=\"month\"/></xlrd2:filters>"
      "</xlrd2:filterColumn></ext><ext uri=\"{00000000-0000-0000-0000-00000000FFFF}\"><y:other/></ext></extLst>"
      "</filterColumn><extLst><ext uri=\"{00000000-0000-0000-0000-00000000EEEE}\"/></extLst></autoFilter>";
  auto parsed = parse_auto_filter_xml(xml);
  ASSERT_TRUE(static_cast<bool>(parsed)) << parsed.error().message;
  const FilterColumn& col = parsed.value().columns.at(0);
  ASSERT_EQ(col.kind, FilterKind::kValues);
  ASSERT_EQ(col.values.date_groups.size(), 1U);
  EXPECT_EQ(col.values.date_groups[0].year, 2025U);
  EXPECT_EQ(col.values.date_groups[0].month, 3U);
  EXPECT_EQ(col.values.date_groups[0].grouping, DateTimeGrouping::kMonth);
  EXPECT_EQ(col.ext_xml, "<ext uri=\"{00000000-0000-0000-0000-00000000FFFF}\"><y:other/></ext>");
  EXPECT_EQ(parsed.value().ext_lst_xml, "<extLst><ext uri=\"{00000000-0000-0000-0000-00000000EEEE}\"/></extLst>");
  EXPECT_EQ(serialize_auto_filter(parsed.value()), xml);

  // A legacy `<filters><dateGroupItem>` is written in the extension form
  // Excel uses (O5b dgi_year_2024).
  auto legacy = parse_auto_filter_xml(
      "<autoFilter ref=\"A1:A10\"><filterColumn colId=\"0\"><filters>"
      "<dateGroupItem year=\"2024\" dateTimeGrouping=\"year\"/></filters></filterColumn></autoFilter>");
  ASSERT_TRUE(static_cast<bool>(legacy));
  EXPECT_EQ(serialize_auto_filter(legacy.value()),
            "<autoFilter ref=\"A1:A10\"><filterColumn colId=\"0\"><extLst>"
            "<ext uri=\"{1AD28BCE-077C-4C59-8B6E-1921CE8616D4}\" "
            "xmlns:xlrd2=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\"><xlrd2:filterColumn>"
            "<xlrd2:filters><xlrd2:dateGroupItem year=\"2024\" dateTimeGrouping=\"year\"/></xlrd2:filters>"
            "</xlrd2:filterColumn></ext></extLst></filterColumn></autoFilter>");
}

TEST(AutoFilterModel, SortStateRoundTripsAndMoves) {
  const std::string xml =
      "<autoFilter ref=\"A1:C9\"><sortState caseSensitive=\"1\" ref=\"A2:C9\">"
      "<sortCondition descending=\"1\" ref=\"B2:B9\"/><sortCondition sortBy=\"icon\" ref=\"C2:C9\" "
      "iconSet=\"3Flags\" iconId=\"2\"/></sortState></autoFilter>";
  auto parsed = parse_auto_filter_xml(xml);
  ASSERT_TRUE(static_cast<bool>(parsed)) << parsed.error().message;
  ASSERT_TRUE(parsed.value().sort.has_value());
  EXPECT_EQ(parsed.value().sort->conditions.at(1).icon_set, cf::IconSetName::Three_Flags);
  EXPECT_EQ(serialize_auto_filter(parsed.value()), xml);

  AutoFilter shifted = parsed.value();
  ASSERT_TRUE(shift_auto_filter(shifted, 1U, 1U, /*is_delete=*/true, /*row_axis=*/false,
                                /*header_delete_removes=*/true));
  EXPECT_EQ(serialize_auto_filter(shifted),
            "<autoFilter ref=\"A1:B9\"><sortState caseSensitive=\"1\" ref=\"A2:B9\">"
            "<sortCondition sortBy=\"icon\" ref=\"B2:B9\" iconSet=\"3Flags\" iconId=\"2\"/></sortState></autoFilter>");
}

TEST(AutoFilterModel, FilterDatabaseNameCreatedAndRemoved) {
  Workbook wb = NewWorkbook();
  const std::size_t quoted = wb.add_sheet("My Data");
  ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter_xml(0, kProbeFilter)));
  const DefinedName* name = FilterDatabaseName(wb, 0);
  ASSERT_NE(name, nullptr);
  EXPECT_TRUE(name->hidden);
  EXPECT_EQ(name->formula, "Sheet1!$B$2:$F$12");

  auto narrowed = parse_auto_filter_xml("<autoFilter ref=\"A1:A16\"/>");
  ASSERT_TRUE(static_cast<bool>(narrowed));
  ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter(quoted, narrowed.value())));
  ASSERT_NE(FilterDatabaseName(wb, static_cast<std::int32_t>(quoted)), nullptr);
  EXPECT_EQ(FilterDatabaseName(wb, static_cast<std::int32_t>(quoted))->formula, "'My Data'!$A$1:$A$16");
  ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter(0, narrowed.value())));
  EXPECT_EQ(FilterDatabaseName(wb, 0)->formula, "Sheet1!$A$1:$A$16");
  EXPECT_EQ(wb.defined_names().size(), 2U);

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet_auto_filter(0)));
  EXPECT_FALSE(wb.sheet(0).has_auto_filter());
  EXPECT_EQ(FilterDatabaseName(wb, 0), nullptr);
  ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter_xml(quoted, "")));
  EXPECT_TRUE(wb.defined_names().empty());

  // Tables carry no `_FilterDatabase` name.
  Workbook tables = ProbeTableWorkbook();
  auto probe = parse_auto_filter_xml(kProbeFilter);
  ASSERT_TRUE(static_cast<bool>(tables.set_table_auto_filter(0, probe.value())));
  EXPECT_TRUE(tables.defined_names().empty());
}

TEST(AutoFilterModel, SetDefinedNameHiddenTogglesTheFlag) {
  Workbook wb = NewWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Data", "Sheet1!$A$1", 0)));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_hidden("data", 0, true)));
  EXPECT_TRUE(wb.defined_names()[0].hidden);
  const auto missing = wb.set_defined_name_hidden("Data", -1, true);
  ASSERT_FALSE(static_cast<bool>(missing));
  EXPECT_EQ(missing.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(AutoFilterModel, ValidationRejectsMalformedModels) {
  auto base = parse_auto_filter_xml(kProbeFilter);
  ASSERT_TRUE(static_cast<bool>(base));
  EXPECT_TRUE(static_cast<bool>(validate_auto_filter(base.value())));

  AutoFilter outside = base.value();
  outside.columns[1].col_id = 5U;  // B..F is five columns wide.
  AutoFilter unordered = base.value();
  std::swap(unordered.columns[0], unordered.columns[1]);
  AutoFilter empty_custom = base.value();
  empty_custom.columns[0].kind = FilterKind::kCustom;
  for (const AutoFilter& bad : {outside, unordered, empty_custom}) {
    Workbook wb = NewWorkbook();
    const auto result = wb.set_sheet_auto_filter(0, bad);
    ASSERT_FALSE(static_cast<bool>(result));
    EXPECT_EQ(result.error().code, FormulonErrorCode::kAutoFilterInvalid);
    EXPECT_FALSE(wb.sheet(0).has_auto_filter());
  }
}

TEST(AutoFilterModel, ChangesDirtySubtotalDependents) {
  Workbook wb = NewWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0U, 2U, "=SUBTOTAL(9,A1:A3)")));
  const auto evaluated = [&wb]() {
    auto stats = wb.recalc(eval::default_registry());
    EXPECT_TRUE(static_cast<bool>(stats));
    return stats ? stats.value().cells_evaluated : 0U;
  };
  evaluated();
  ASSERT_EQ(evaluated(), 0U);

  auto probe = parse_auto_filter_xml(kProbeFilter);
  ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter(0, probe.value())));
  EXPECT_EQ(evaluated(), 1U) << "set";
  ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter_xml(0, kProbeFilter)));
  EXPECT_EQ(evaluated(), 1U) << "xml view";
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, 3U, 1U)));
  EXPECT_GE(evaluated(), 1U) << "structural edit";
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet_auto_filter(0)));
  EXPECT_EQ(evaluated(), 1U) << "remove";

  TableMetadata table;
  table.name = "T";
  table.display_name = "T";
  table.ref = "E1:F4";
  wb.mutable_tables().push_back(std::move(table));
  ASSERT_TRUE(static_cast<bool>(wb.set_table_auto_filter(0, probe.value())));
  EXPECT_EQ(evaluated(), 1U) << "table set";
  ASSERT_TRUE(static_cast<bool>(wb.remove_table_auto_filter(0)));
  EXPECT_EQ(evaluated(), 1U) << "table remove";
}

TEST(AutoFilterModel, XlsbSaveReportsTheDeferredFilter) {
  Workbook wb = NewWorkbook();
  auto probe = parse_auto_filter_xml(kProbeFilter);
  ASSERT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter(0, probe.value())));
  auto written = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  EXPECT_EQ(written.value().diagnostics.deferred_feature_count, 1U);
}

}  // namespace
}  // namespace formulon
