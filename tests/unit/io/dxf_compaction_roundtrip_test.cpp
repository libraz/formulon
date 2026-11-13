//
// An .xlsx save writes only the dxfs some written reference names, one record
// per user, in Excel's order, renumbers every
// reference (modelled and raw XML), and leaves the model's dxf table
// untouched. A styles part whose slicer styles may index the main table, or a
// raw fragment the scanner cannot read, saves the table as held.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <initializer_list>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/dxf_xml_ids.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/styles_writer.h"
#include "miniz.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_passthrough.h"
#include "styles.h"
#include "table.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io {
namespace {

/// A dxf whose saved body is unique to `tag`.
DifferentialFormat TaggedDxf(std::uint32_t tag) {
  DifferentialFormat dxf;
  dxf.has_num_fmt = true;
  dxf.num_fmt_id = static_cast<std::uint16_t>(200U + tag);
  dxf.num_fmt_code = "\"dxf" + std::to_string(tag) + "\"0";
  return dxf;
}

std::string Body(const DifferentialFormat& dxf) {
  return dxf_body_fragment(dxf);
}

std::vector<std::string> Bodies(const std::vector<DifferentialFormat>& dxfs) {
  std::vector<std::string> out;
  for (const DifferentialFormat& dxf : dxfs) {
    out.push_back(Body(dxf));
  }
  return out;
}

bool HasBody(const Workbook& wb, const std::string& body) {
  for (const DifferentialFormat& dxf : wb.styles().dxfs) {
    if (Body(dxf) == body) {
      return true;
    }
  }
  return false;
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto bytes_or = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << (bytes_or ? "" : bytes_or.error().message);
  return bytes_or ? std::move(bytes_or.value()) : std::vector<std::uint8_t>{};
}

Workbook Reload(const std::vector<std::uint8_t>& bytes) {
  auto read_or = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << (read_or ? "" : read_or.error().message);
  return read_or ? std::move(read_or.value().workbook) : Workbook::create_empty();
}

/// Every zip entry of `bytes` by path.
std::map<std::string, std::string> Parts(const std::vector<std::uint8_t>& bytes) {
  std::map<std::string, std::string> out;
  mz_zip_archive zip{};
  if (mz_zip_reader_init_mem(&zip, bytes.data(), bytes.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return out;
  }
  for (mz_uint i = 0; i < mz_zip_reader_get_num_files(&zip); ++i) {
    char name[512] = {};
    if (mz_zip_reader_get_filename(&zip, i, name, sizeof(name)) == 0) {
      continue;
    }
    std::size_t size = 0;
    void* body = mz_zip_reader_extract_to_heap(&zip, i, &size, 0);
    if (body != nullptr) {
      out.emplace(name, std::string(static_cast<const char*>(body), size));
      mz_free(body);
    }
  }
  mz_zip_reader_end(&zip);
  return out;
}

/// Number of `<dxf>` records in the saved `<dxfs>`.
std::size_t SavedDxfCount(const std::map<std::string, std::string>& parts) {
  const auto it = parts.find("xl/styles.xml");
  if (it == parts.end()) {
    ADD_FAILURE() << "no xl/styles.xml";
    return 0;
  }
  const std::string& styles = it->second;
  const std::size_t begin = styles.find("<dxfs");
  const std::size_t end = begin == std::string::npos ? begin : styles.find("</dxfs>", begin);
  if (end == std::string::npos) {
    return 0;
  }
  std::size_t count = 0;
  for (std::size_t at = styles.find("<dxf>", begin); at < end; at = styles.find("<dxf>", at + 1U)) {
    ++count;
  }
  return count;
}

/// Every dxf attribute in every saved XML part names a saved dxf.
void ExpectEveryIdResolves(const std::map<std::string, std::string>& parts) {
  const std::size_t saved = SavedDxfCount(parts);
  for (const auto& [path, xml] : parts) {
    if (path.size() < 4U || path.compare(path.size() - 4U, 4U, ".xml") != 0) {
      continue;
    }
    const auto ids = visit_dxf_xml_ids(xml);
    ASSERT_TRUE(static_cast<bool>(ids)) << path;
    for (const DxfXmlId& id : ids.value()) {
      EXPECT_LT(id.value, saved) << path << " " << id.attribute;
    }
  }
}

/// The value of dxf attribute `name` in `text`, a fragment or an attribute list.
std::optional<std::uint32_t> DxfAttr(std::string_view text, std::string_view name, bool attr_list) {
  const auto ids = attr_list ? visit_dxf_attr_ids(text) : visit_dxf_xml_ids(text);
  EXPECT_TRUE(static_cast<bool>(ids)) << text;
  if (!ids) {
    return std::nullopt;
  }
  for (const DxfXmlId& id : ids.value()) {
    if (id.attribute == name) {
      return id.value;
    }
  }
  return std::nullopt;
}

cf::ConditionalFormat CellIsBlock(std::uint32_t col, std::optional<std::uint32_t> dxf_id, std::string formula,
                                  int priority) {
  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, col}, {9, col}});
  cf::CFRule rule;
  rule.type = cf::RuleType::CellIs;
  rule.op = cf::CellIsOperator::GreaterThan;
  rule.priority = priority;
  rule.dxf_id = dxf_id;
  rule.formula1 = std::move(formula);
  block.rules.push_back(std::move(rule));
  return block;
}

TableMetadata TwoColumnTable(std::string root_attrs, std::string column_attrs) {
  TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.display_name = "Table1";
  table.ref = "A1:B3";
  table.sheet_index = 0;
  table.columns.emplace_back(1U, "H1", "", "", "");
  table.columns.emplace_back(2U, "H2", "", "", "");
  table.root_extra_attrs = std::move(root_attrs);
  table.columns[0].extra_attrs = std::move(column_attrs);
  return table;
}

void SetTableHeaders(Workbook& wb) {
  wb.set_cell_value(0, 0, 0, Value::text("H1"));
  wb.set_cell_value(0, 0, 1, Value::text("H2"));
}

/// A sheet AutoFilter on D1:E3 with a colour filter and a colour sort.
AutoFilter ColorFilterAndSort(std::optional<std::uint32_t> filter_dxf, std::optional<std::uint32_t> sort_dxf) {
  AutoFilter filter;
  filter.range = MergeRange{0, 3, 2, 4};
  FilterColumn col;
  col.col_id = 0;
  col.kind = FilterKind::kColor;
  col.color.dxf_id = filter_dxf;
  filter.columns.push_back(col);
  SortState sort;
  sort.ref = MergeRange{1, 3, 2, 4};
  SortCondition cond;
  cond.ref = MergeRange{1, 3, 2, 3};
  cond.sort_by = SortBy::kCellColor;
  cond.dxf_id = sort_dxf;
  sort.conditions.push_back(cond);
  filter.sort = sort;
  return filter;
}

TEST(DxfCompactionXlsx, OrphansDropAndEveryReferenceResolvesToItsRecord) {
  Workbook wb = Workbook::create();
  SetTableHeaders(wb);
  for (std::uint32_t i = 0; i < 11U; ++i) {
    wb.mutable_styles().dxfs.push_back(TaggedDxf(i));
  }
  // 0, 2 and 8 are orphans; every other dxf has one user of its own kind.
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(0, 1U, "50", 1));
  TableMetadata table = TwoColumnTable(" headerRowDxfId=\"3\"", " dataDxfId=\"4\"");
  table.sort_state_xml =
      "<sortState ref=\"A2:B3\"><sortCondition ref=\"A2:A3\" sortBy=\"cellColor\" dxfId=\"9\"/></sortState>";
  wb.mutable_tables().push_back(std::move(table));
  CustomTableStyle style;
  style.name = "Custom1";
  style.elements.push_back(TableStyleElement{"wholeTable", 1U, 5U, ""});
  wb.mutable_styles().table_styles.present = true;
  wb.mutable_styles().table_styles.styles.push_back(style);
  wb.sheet(0).set_auto_filter(ColorFilterAndSort(6U, 7U));
  wb.sheet(0).mutable_raw_extensions().push_back(WorksheetRawChild{
      static_cast<std::uint32_t>(worksheet_child::slot_of("sortState")),
      "<sortState ref=\"G1:G3\"><sortCondition ref=\"G1:G3\" sortBy=\"fontColor\" dxfId=\"10\"/></sortState>"});
  const std::vector<std::string> model_before = Bodies(wb.styles().dxfs);

  const std::vector<std::uint8_t> bytes = Save(wb);
  // The save leaves the model's table as it was.
  EXPECT_EQ(Bodies(wb.styles().dxfs), model_before);

  const std::map<std::string, std::string> parts = Parts(bytes);
  EXPECT_EQ(SavedDxfCount(parts), 8U);
  ExpectEveryIdResolves(parts);

  const Workbook back = Reload(bytes);
  ASSERT_EQ(back.styles().dxfs.size(), 8U);
  for (const std::uint32_t orphan : {0U, 2U, 8U}) {
    EXPECT_FALSE(HasBody(back, model_before[orphan])) << orphan;
  }
  const auto expect_resolves = [&](std::optional<std::uint32_t> saved, std::uint32_t model, const char* holder) {
    ASSERT_TRUE(saved.has_value()) << holder;
    ASSERT_LT(*saved, back.styles().dxfs.size()) << holder;
    EXPECT_EQ(Body(back.styles().dxfs[*saved]), model_before[model]) << holder;
  };
  expect_resolves(back.sheet(0).conditional_formats().at(0).rules.at(0).dxf_id, 1U, "cfRule");
  ASSERT_EQ(back.tables().size(), 1U);
  expect_resolves(DxfAttr(back.tables()[0].root_extra_attrs, "headerRowDxfId", true), 3U, "table");
  expect_resolves(DxfAttr(back.tables()[0].columns.at(0).extra_attrs, "dataDxfId", true), 4U, "tableColumn");
  expect_resolves(back.styles().table_styles.styles.at(0).elements.at(0).dxf_id, 5U, "tableStyleElement");
  ASSERT_TRUE(back.sheet(0).has_auto_filter());
  expect_resolves(back.sheet(0).auto_filter()->columns.at(0).color.dxf_id, 6U, "colorFilter");
  expect_resolves(back.sheet(0).auto_filter()->sort->conditions.at(0).dxf_id, 7U, "sortCondition");
  expect_resolves(DxfAttr(back.tables()[0].sort_state_xml, "dxfId", false), 9U, "table sortState");
  std::optional<std::uint32_t> raw_sort;
  for (const WorksheetRawChild& child : back.sheet(0).raw_extensions()) {
    if (child.xml.find("<sortState") != std::string::npos) {
      raw_sort = DxfAttr(child.xml, "dxfId", false);
    }
  }
  expect_resolves(raw_sort, 10U, "worksheet sortState");
}

TEST(DxfCompactionXlsx, DxfSharedByTwoKindsSavesOneRecordPerUser) {
  Workbook wb = Workbook::create();
  wb.mutable_styles().dxfs.push_back(TaggedDxf(0U));
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(0, 0U, "50", 1));
  CustomTableStyle style;
  style.name = "Custom1";
  style.elements.push_back(TableStyleElement{"wholeTable", 1U, 0U, ""});
  wb.mutable_styles().table_styles.present = true;
  wb.mutable_styles().table_styles.styles.push_back(style);
  const std::string body = Body(wb.styles().dxfs[0]);

  const std::vector<std::uint8_t> bytes = Save(wb);
  EXPECT_EQ(wb.styles().dxfs.size(), 1U);
  const std::map<std::string, std::string> parts = Parts(bytes);
  EXPECT_EQ(SavedDxfCount(parts), 2U);
  ExpectEveryIdResolves(parts);

  const Workbook back = Reload(bytes);
  ASSERT_EQ(back.styles().dxfs.size(), 2U);
  EXPECT_EQ(Bodies(back.styles().dxfs), (std::vector<std::string>{body, body}));
  const std::optional<std::uint32_t> cf_id = back.sheet(0).conditional_formats().at(0).rules.at(0).dxf_id;
  const std::optional<std::uint32_t> element_id = back.styles().table_styles.styles.at(0).elements.at(0).dxf_id;
  // The styles-part user registers first, so it is written last.
  EXPECT_EQ(cf_id, std::optional<std::uint32_t>(0U));
  EXPECT_EQ(element_id, std::optional<std::uint32_t>(1U));
}

TEST(DxfCompactionXlsx, CfRulesSharingADxfTakeARecordEach) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.mutable_styles().dxfs.push_back(TaggedDxf(0U));
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(0, 0U, "50", 1));
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(1, 0U, "60", 2));
  const std::string body = Body(wb.styles().dxfs[0]);

  const std::vector<std::uint8_t> one_sheet = Save(wb);
  EXPECT_EQ(SavedDxfCount(Parts(one_sheet)), 2U);
  const Workbook back = Reload(one_sheet);
  ASSERT_EQ(back.styles().dxfs.size(), 2U);
  // Blocks register in document order; the later block's record is written first.
  EXPECT_EQ(back.sheet(0).conditional_formats().at(0).rules.at(0).dxf_id, std::optional<std::uint32_t>(1U));
  EXPECT_EQ(back.sheet(0).conditional_formats().at(1).rules.at(0).dxf_id, std::optional<std::uint32_t>(0U));

  wb.sheet(1).mutable_conditional_formats().push_back(CellIsBlock(0, 0U, "70", 3));
  const std::vector<std::uint8_t> two_sheets = Save(wb);
  const std::map<std::string, std::string> parts = Parts(two_sheets);
  EXPECT_EQ(SavedDxfCount(parts), 3U);
  ExpectEveryIdResolves(parts);
  const Workbook again = Reload(two_sheets);
  ASSERT_EQ(again.styles().dxfs.size(), 3U);
  EXPECT_EQ(Bodies(again.styles().dxfs), (std::vector<std::string>{body, body, body}));
  // CF registers by sheet in tab order; the later sheet's record is written first.
  EXPECT_EQ(again.sheet(0).conditional_formats().at(0).rules.at(0).dxf_id, std::optional<std::uint32_t>(2U));
  EXPECT_EQ(again.sheet(0).conditional_formats().at(1).rules.at(0).dxf_id, std::optional<std::uint32_t>(1U));
  EXPECT_EQ(again.sheet(1).conditional_formats().at(0).rules.at(0).dxf_id, std::optional<std::uint32_t>(0U));
}

TEST(DxfCompactionXlsx, RecordsFollowExcelOrderNotModelOrder) {
  Workbook wb = Workbook::create();
  SetTableHeaders(wb);
  for (std::uint32_t i = 0; i < 3U; ++i) {
    wb.mutable_styles().dxfs.push_back(TaggedDxf(i));
  }
  // Model order: table 0, style element 1, CF 2.
  wb.mutable_tables().push_back(TwoColumnTable(" headerRowDxfId=\"0\"", ""));
  CustomTableStyle style;
  style.name = "Custom1";
  style.elements.push_back(TableStyleElement{"wholeTable", 1U, 1U, ""});
  wb.mutable_styles().table_styles.present = true;
  wb.mutable_styles().table_styles.styles.push_back(style);
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(0, 2U, "50", 1));
  const std::vector<std::string> model = Bodies(wb.styles().dxfs);

  const std::vector<std::uint8_t> bytes = Save(wb);
  EXPECT_EQ(Bodies(wb.styles().dxfs), model);
  ExpectEveryIdResolves(Parts(bytes));
  const Workbook back = Reload(bytes);
  // Registered style, table, CF; written reversed.
  EXPECT_EQ(Bodies(back.styles().dxfs), (std::vector<std::string>{model[2], model[0], model[1]}));
  EXPECT_EQ(back.sheet(0).conditional_formats().at(0).rules.at(0).dxf_id, std::optional<std::uint32_t>(0U));
  ASSERT_EQ(back.tables().size(), 1U);
  EXPECT_EQ(DxfAttr(back.tables()[0].root_extra_attrs, "headerRowDxfId", true), std::optional<std::uint32_t>(1U));
  EXPECT_EQ(back.styles().table_styles.styles.at(0).elements.at(0).dxf_id, std::optional<std::uint32_t>(2U));
}

TEST(DxfCompactionXlsx, PivotFormatInRawXmlIsRenumbered) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_layout/en_c0_rows.xlsx";
  std::vector<std::uint8_t> source;
  if (FILE* file = std::fopen(path.c_str(), "rb"); file != nullptr) {
    for (int c = std::fgetc(file); c != EOF; c = std::fgetc(file)) {
      source.push_back(static_cast<std::uint8_t>(c));
    }
    std::fclose(file);
  }
  ASSERT_FALSE(source.empty()) << path;
  Workbook wb = Reload(source);
  pivot::PivotTable* pivot = nullptr;
  for (std::size_t s = 0; s < wb.sheet_count() && pivot == nullptr; ++s) {
    if (!wb.sheet(s).pivot_tables().empty()) {
      pivot = wb.sheet(s).mutable_pivot_tables().front().get();
    }
  }
  ASSERT_NE(pivot, nullptr);
  const auto base = static_cast<std::uint32_t>(wb.styles().dxfs.size());
  wb.mutable_styles().dxfs.push_back(TaggedDxf(900U));  // orphan
  wb.mutable_styles().dxfs.push_back(TaggedDxf(901U));
  pivot->mutable_raw_passthrough_before_style_info() =
      "<formats count=\"1\"><format dxfId=\"" + std::to_string(base + 1U) +
      "\"><pivotArea type=\"all\" dataOnly=\"0\" outline=\"0\" fieldPosition=\"0\"/></format></formats>";
  const std::string used = Body(wb.styles().dxfs[base + 1U]);
  const std::string orphan = Body(wb.styles().dxfs[base]);

  const std::vector<std::uint8_t> bytes = Save(wb);
  ExpectEveryIdResolves(Parts(bytes));
  const Workbook back = Reload(bytes);
  EXPECT_FALSE(HasBody(back, orphan));
  const pivot::PivotTable* saved = nullptr;
  for (std::size_t s = 0; s < back.sheet_count() && saved == nullptr; ++s) {
    if (!back.sheet(s).pivot_tables().empty()) {
      saved = back.sheet(s).pivot_tables().front().get();
    }
  }
  ASSERT_NE(saved, nullptr);
  const std::optional<std::uint32_t> id = DxfAttr(saved->raw_passthrough_before_style_info(), "dxfId", false);
  ASSERT_TRUE(id.has_value());
  ASSERT_LT(*id, back.styles().dxfs.size());
  EXPECT_EQ(Body(back.styles().dxfs[*id]), used);
}

TEST(DxfCompactionXlsx, IdNamingNoModelDxfIsOmittedInEveryHolder) {
  Workbook wb = Workbook::create();
  SetTableHeaders(wb);
  wb.mutable_styles().dxfs.push_back(TaggedDxf(0U));
  wb.mutable_styles().dxfs.push_back(TaggedDxf(1U));
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(0, 1U, "50", 1));
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(1, 2U, "60", 2));
  TableMetadata table = TwoColumnTable(" dataDxfId=\"99\"", " totalsRowDxfId=\"4294967295\"");
  table.sort_state_xml =
      "<sortState ref=\"A2:B3\"><sortCondition ref=\"A2:A3\" sortBy=\"cellColor\" dxfId=\"7\"/></sortState>";
  wb.mutable_tables().push_back(std::move(table));
  CustomTableStyle style;
  style.name = "Custom1";
  style.elements.push_back(TableStyleElement{"wholeTable", 1U, 77U, ""});
  wb.mutable_styles().table_styles.present = true;
  wb.mutable_styles().table_styles.styles.push_back(style);
  wb.sheet(0).set_auto_filter(ColorFilterAndSort(2U, 0xFFFFFFFFU));
  const std::string used = Body(wb.styles().dxfs[1]);

  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::map<std::string, std::string> parts = Parts(bytes);
  EXPECT_EQ(SavedDxfCount(parts), 1U);
  ExpectEveryIdResolves(parts);
  const std::string& table_xml = parts.at("xl/tables/table1.xml");
  EXPECT_EQ(table_xml.find("dataDxfId"), std::string::npos) << table_xml;
  EXPECT_EQ(table_xml.find("totalsRowDxfId"), std::string::npos) << table_xml;
  EXPECT_EQ(table_xml.find("dxfId"), std::string::npos) << table_xml;
  EXPECT_EQ(parts.at("xl/styles.xml").find("dxfId"), std::string::npos);

  const Workbook back = Reload(bytes);
  ASSERT_EQ(back.styles().dxfs.size(), 1U);
  EXPECT_EQ(Body(back.styles().dxfs[0]), used);
  const auto& blocks = back.sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 2U);
  EXPECT_EQ(blocks[0].rules.at(0).dxf_id, std::optional<std::uint32_t>(0U));
  EXPECT_FALSE(blocks[1].rules.at(0).dxf_id.has_value());
  ASSERT_TRUE(back.sheet(0).has_auto_filter());
  EXPECT_FALSE(back.sheet(0).auto_filter()->columns.at(0).color.dxf_id.has_value());
  EXPECT_FALSE(back.sheet(0).auto_filter()->sort->conditions.at(0).dxf_id.has_value());
  EXPECT_FALSE(back.styles().table_styles.styles.at(0).elements.at(0).dxf_id.has_value());
}

// Excel keeps an x14-only rule's inline format in `<dxfs>` too (structural_ref_m6b.xlsx).
TEST(DxfCompactionXlsx, X14OnlyRuleWritesItsDxfInlineAndKeepsItInTheTable) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.mutable_styles().dxfs.push_back(TaggedDxf(0U));  // orphan
  wb.mutable_styles().dxfs.push_back(TaggedDxf(1U));
  wb.mutable_styles().dxfs.push_back(TaggedDxf(2U));
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(0, 1U, "Sheet2!$A$1", 1));
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(1, 2U, "50", 2));
  const std::string orphan_body = Body(wb.styles().dxfs[0]);
  const std::string inline_body = Body(wb.styles().dxfs[1]);
  const std::string legacy_body = Body(wb.styles().dxfs[2]);

  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::map<std::string, std::string> parts = Parts(bytes);
  EXPECT_EQ(SavedDxfCount(parts), 2U);
  ExpectEveryIdResolves(parts);
  EXPECT_NE(parts.at("xl/worksheets/sheet1.xml").find("<x14:dxf>" + inline_body + "</x14:dxf>"), std::string::npos);

  const Workbook back = Reload(bytes);
  ASSERT_EQ(back.styles().dxfs.size(), 2U);
  EXPECT_FALSE(HasBody(back, orphan_body));
  bool saw_inline = false;
  bool saw_legacy = false;
  for (const cf::ConditionalFormat& block : back.sheet(0).conditional_formats()) {
    for (const cf::CFRule& rule : block.rules) {
      ASSERT_TRUE(rule.dxf_id.has_value());
      ASSERT_LT(*rule.dxf_id, back.styles().dxfs.size());
      const std::string body = Body(back.styles().dxfs[*rule.dxf_id]);
      const bool x14 = rule.formula1.has_value() && rule.formula1->find("Sheet2") != std::string::npos;
      EXPECT_EQ(body, x14 ? inline_body : legacy_body);
      (x14 ? saw_inline : saw_legacy) = true;
    }
  }
  EXPECT_TRUE(saw_inline);
  EXPECT_TRUE(saw_legacy);
}

TEST(DxfCompactionXlsx, NoReferenceSavesNoDxfs) {
  Workbook wb = Workbook::create();
  wb.mutable_styles().dxfs.push_back(TaggedDxf(0U));
  wb.mutable_styles().dxfs.push_back(TaggedDxf(1U));

  const std::vector<std::uint8_t> bytes = Save(wb);
  EXPECT_EQ(wb.styles().dxfs.size(), 2U);
  EXPECT_EQ(SavedDxfCount(Parts(bytes)), 0U);
  EXPECT_TRUE(Reload(bytes).styles().dxfs.empty());
}

/// A styles `<extLst>` whose slicer style names `dxfId="0"`; `extension_dxfs`
/// adds the `x14:dxfs` collection that element indexes.
std::string SlicerStylesExt(bool extension_dxfs) {
  std::string ext =
      "<extLst><ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">";
  if (extension_dxfs) {
    ext.append("<x14:dxfs count=\"1\"><dxf><font><b/></font></dxf></x14:dxfs>");
  }
  ext.append(
      "<x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"><x14:slicerStyle name=\"S1\">"
      "<x14:slicerStyleElements><x14:slicerStyleElement type=\"selectedItemWithData\" dxfId=\"0\"/>"
      "</x14:slicerStyleElements></x14:slicerStyle></x14:slicerStyles></ext></extLst>");
  return ext;
}

TEST(DxfCompactionXlsx, SlicerStyleWithoutExtensionDxfsSavesTheTableAsHeld) {
  for (const bool extension_dxfs : {false, true}) {
    Workbook wb = Workbook::create();
    wb.mutable_styles().dxfs.push_back(TaggedDxf(0U));  // orphan
    wb.mutable_styles().dxfs.push_back(TaggedDxf(1U));
    wb.mutable_styles().ext_lst_xml = SlicerStylesExt(extension_dxfs);
    wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(0, 1U, "50", 1));

    const std::map<std::string, std::string> parts = Parts(Save(wb));
    const std::string& sheet = parts.at("xl/worksheets/sheet1.xml");
    if (extension_dxfs) {
      EXPECT_EQ(SavedDxfCount(parts), 1U);
      EXPECT_NE(sheet.find("dxfId=\"0\""), std::string::npos) << sheet;
    } else {
      EXPECT_EQ(SavedDxfCount(parts), 2U);
      EXPECT_NE(sheet.find("dxfId=\"1\""), std::string::npos) << sheet;
    }
    // The slicer element is never renumbered.
    EXPECT_NE(parts.at("xl/styles.xml").find("type=\"selectedItemWithData\" dxfId=\"0\""), std::string::npos);
  }
}

TEST(DxfCompactionXlsx, UnreadableRawFragmentSavesTheTableAsHeld) {
  Workbook wb = Workbook::create();
  wb.mutable_styles().dxfs.push_back(TaggedDxf(0U));  // orphan
  wb.mutable_styles().dxfs.push_back(TaggedDxf(1U));
  wb.sheet(0).mutable_conditional_formats().push_back(CellIsBlock(0, 1U, "50", 1));
  const std::string unreadable = "<customSheetViews><customSheetView guid=\"{00000000-0000-0000-0000-000000000000}\"";
  wb.sheet(0).mutable_raw_extensions().push_back(
      WorksheetRawChild{static_cast<std::uint32_t>(worksheet_child::slot_of("customSheetViews")), unreadable});

  const std::map<std::string, std::string> parts = Parts(Save(wb));
  EXPECT_EQ(SavedDxfCount(parts), 2U);
  const std::string& sheet = parts.at("xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet.find("dxfId=\"1\""), std::string::npos) << sheet;
  EXPECT_NE(sheet.find(unreadable), std::string::npos);
}

}  // namespace
}  // namespace formulon::io
