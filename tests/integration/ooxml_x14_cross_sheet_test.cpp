//
// Conditional-format rules that name another sheet, through the x14
// extension: the Excel-saved fixtures load them into the model, and a save
// writes each one back in the extension alone, in Excel's own XML.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <initializer_list>
#include <map>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

std::vector<std::uint8_t> ReadFixture(const std::string& file) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + file;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    out.clear();
  }
  std::fclose(f);
  return out;
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto r = io::read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto r = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value()) : std::vector<std::uint8_t>{};
}

std::string ReadPart(const std::vector<std::uint8_t>& bytes, const char* path) {
  io::ZipReader zip;
  if (!zip.open(SpanOf(bytes))) {
    ADD_FAILURE() << "not a zip";
    return {};
  }
  auto entry = zip.read_entry(path);
  if (!entry) {
    ADD_FAILURE() << "no part " << path;
    return {};
  }
  return std::string(entry.value().begin(), entry.value().end());
}

std::size_t Count(const std::string& haystack, std::string_view needle) {
  std::size_t n = 0;
  for (std::size_t at = haystack.find(needle); at != std::string::npos; at = haystack.find(needle, at + 1U)) {
    ++n;
  }
  return n;
}

/// Every `<x14:conditionalFormatting>` of the worksheet part `xml`, keyed by
/// the id of its first rule and re-serialised by pugixml, so two documents
/// compare by element, attribute order and namespace placement alone.
std::map<std::string, std::string> X14Blocks(const std::string& xml) {
  std::map<std::string, std::string> out;
  pugi::xml_document doc;
  if (!doc.load_string(xml.c_str())) {
    ADD_FAILURE() << "worksheet does not parse";
    return out;
  }
  for (pugi::xml_node ext : doc.child("worksheet").child("extLst").children("ext")) {
    for (pugi::xml_node block : ext.child("x14:conditionalFormattings").children("x14:conditionalFormatting")) {
      out[block.child("x14:cfRule").attribute("id").value()] = io::raw_xml(block);
    }
  }
  return out;
}

/// The `<conditionalFormatting>` elements of the worksheet part `xml`.
std::string MainBlocks(const std::string& xml) {
  std::string out;
  for (std::size_t at = xml.find("<conditionalFormatting "); at != std::string::npos;
       at = xml.find("<conditionalFormatting ", at + 1U)) {
    const std::size_t end = xml.find("</conditionalFormatting>", at);
    out.append(xml, at, end == std::string::npos ? std::string::npos : end - at);
  }
  return out;
}

const cf::CFRule* RuleWithPriority(const Sheet& sheet, std::int32_t priority) {
  for (const cf::ConditionalFormat& format : sheet.conditional_formats()) {
    for (const cf::CFRule& rule : format.rules) {
      if (rule.priority == priority) {
        return &rule;
      }
    }
  }
  return nullptr;
}

/// Asserts that re-saving `fixture` writes Sheet2's x14 blocks exactly as
/// Excel did, each rule once and none of them in the main element.
void ExpectResaveMatchesExcel(const char* fixture, std::initializer_list<const char*> ids) {
  const std::vector<std::uint8_t> excel = ReadFixture(fixture);
  const std::string excel_xml = ReadPart(excel, "xl/worksheets/sheet2.xml");
  const std::string saved_xml = ReadPart(Save(Load(excel)), "xl/worksheets/sheet2.xml");
  const std::map<std::string, std::string> expected = X14Blocks(excel_xml);
  const std::map<std::string, std::string> actual = X14Blocks(saved_xml);
  for (const char* id : ids) {
    ASSERT_EQ(expected.count(id), 1U) << id;
    ASSERT_EQ(actual.count(id), 1U) << id << "\n" << saved_xml;
    EXPECT_EQ(actual.at(id), expected.at(id)) << id;
    EXPECT_EQ(Count(saved_xml, id), 1U) << id << "\n" << saved_xml;
  }
  // INDIRECT("Sheet1!A1") is no reference and stays in the main element.
  EXPECT_EQ(MainBlocks(saved_xml).find(">Sheet1!"), std::string::npos) << saved_xml;
  EXPECT_EQ(MainBlocks(saved_xml).find("(Sheet1!"), std::string::npos) << saved_xml;
  EXPECT_EQ(MainBlocks(saved_xml).find(">Sheet3!"), std::string::npos) << saved_xml;
}

constexpr const char* kM6Id1 = "{00000000-000E-0000-0100-000001000000}";
constexpr const char* kM6Id4 = "{00000000-000E-0000-0100-000004000000}";
constexpr const char* kM6Id6 = "{00000000-000E-0000-0100-000006000000}";
constexpr const char* kM10Id2 = "{00000000-000E-0000-0100-000002000000}";
constexpr const char* kM10Id3 = "{00000000-000E-0000-0100-000003000000}";

TEST(OoxmlX14CrossSheet, CfExcelRulesLoadIntoTheModel) {
  const Workbook wb = Load(ReadFixture("structural_ref_m6.xlsx"));
  const Sheet& sheet = wb.sheet(1);
  const cf::CFRule* p1 = RuleWithPriority(sheet, 1);
  const cf::CFRule* p4 = RuleWithPriority(sheet, 4);
  const cf::CFRule* p6 = RuleWithPriority(sheet, 6);
  ASSERT_NE(p1, nullptr);
  ASSERT_NE(p4, nullptr);
  ASSERT_NE(p6, nullptr);
  EXPECT_EQ(p1->type, cf::RuleType::Expression);
  EXPECT_EQ(p1->formula1.value_or(""), "Sheet1!$A$1>0");
  EXPECT_EQ(p1->id, kM6Id1);
  EXPECT_FALSE(p1->dxf_id.has_value());
  EXPECT_EQ(p4->type, cf::RuleType::CellIs);
  ASSERT_TRUE(p4->op.has_value());
  EXPECT_EQ(*p4->op, cf::CellIsOperator::GreaterThan);
  EXPECT_EQ(p4->formula1.value_or(""), "Sheet1!$A$1");
  EXPECT_EQ(p6->formula1.value_or(""), "SUM(Sheet1!A:A)>0");
  // The same-sheet rules of the main element are still there.
  EXPECT_NE(RuleWithPriority(sheet, 2), nullptr);
  EXPECT_NE(RuleWithPriority(sheet, 8), nullptr);
  // Adopted rules leave the captured extension.
  EXPECT_EQ(sheet.ext_lst_xml().find(kM6Id1), std::string::npos) << sheet.ext_lst_xml();
}

TEST(OoxmlX14CrossSheet, CfExcelRuleWithFormattingLoadsItsDxf) {
  const Workbook wb = Load(ReadFixture("structural_ref_m6b.xlsx"));
  const cf::CFRule* rule = RuleWithPriority(wb.sheet(1), 1);
  ASSERT_NE(rule, nullptr);
  EXPECT_TRUE(rule->stop_if_true);
  EXPECT_EQ(rule->formula1.value_or(""), "Sheet1!$A$1>0");
  // Excel also lists the inline format in `<dxfs>`; the rule reuses that entry.
  ASSERT_EQ(wb.styles().dxfs.size(), 2U);
  ASSERT_TRUE(rule->dxf_id.has_value());
  EXPECT_EQ(*rule->dxf_id, 1U);
  EXPECT_TRUE(wb.styles().dxfs[1].has_font);
  EXPECT_TRUE(wb.styles().dxfs[1].font.bold);
}

TEST(OoxmlX14CrossSheet, CfColourScaleWithCrossSheetCfvoLoads) {
  const Workbook wb = Load(ReadFixture("structural_ref_m10b.xlsx"));
  const cf::CFRule* scale = RuleWithPriority(wb.sheet(1), 3);
  ASSERT_NE(scale, nullptr);
  EXPECT_EQ(scale->type, cf::RuleType::ColorScale);
  EXPECT_EQ(scale->id, kM10Id3);
  ASSERT_TRUE(scale->color_scale.has_value());
  ASSERT_EQ(scale->color_scale->thresholds.size(), 2U);
  EXPECT_EQ(scale->color_scale->thresholds[0].type, cf::CfvoType::Formula);
  EXPECT_EQ(scale->color_scale->thresholds[0].value, "Sheet3!$A$1");
  EXPECT_EQ(scale->color_scale->thresholds[1].type, cf::CfvoType::Max);
  ASSERT_EQ(scale->color_scale->colors.size(), 2U);
  EXPECT_EQ(scale->color_scale->colors[0], (cf::Color{0xFF, 0x00, 0x00, 0xFF}));
  const cf::CFRule* p2 = RuleWithPriority(wb.sheet(1), 2);
  ASSERT_NE(p2, nullptr);
  EXPECT_EQ(p2->formula1.value_or(""), "Sheet3!$A$1>0");
  // The data validations of the same extension list are in the model too.
  EXPECT_EQ(wb.sheet(1).ext_lst_xml().find("x14:dataValidations"), std::string::npos);
  EXPECT_EQ(wb.sheet(1).validations().size(), 2U);
}

TEST(OoxmlX14CrossSheet, CfResaveWritesExcelXml) {
  ExpectResaveMatchesExcel("structural_ref_m6.xlsx", {kM6Id1, kM6Id4, kM6Id6});
}

TEST(OoxmlX14CrossSheet, CfResaveWritesExcelXmlWithInlineDxf) {
  ExpectResaveMatchesExcel("structural_ref_m6b.xlsx", {kM6Id1});
}

TEST(OoxmlX14CrossSheet, CfResaveWritesExcelColourScaleXml) {
  ExpectResaveMatchesExcel("structural_ref_m10b.xlsx", {kM6Id1, kM10Id2, kM10Id3});
}

TEST(OoxmlX14CrossSheet, CfSecondRoundTripIsStable) {
  const std::vector<std::uint8_t> once = Save(Load(ReadFixture("structural_ref_m6b.xlsx")));
  const std::vector<std::uint8_t> twice = Save(Load(once));
  const std::string styles_once = ReadPart(once, "xl/styles.xml");
  const std::string styles_twice = ReadPart(twice, "xl/styles.xml");
  EXPECT_NE(styles_once.find("<dxfs count=\"2\">"), std::string::npos) << styles_once;
  EXPECT_NE(styles_twice.find("<dxfs count=\"2\">"), std::string::npos) << styles_twice;
  const std::string sheet_twice = ReadPart(twice, "xl/worksheets/sheet2.xml");
  EXPECT_EQ(Count(sheet_twice, kM6Id1), 1U) << sheet_twice;
  EXPECT_EQ(X14Blocks(sheet_twice), X14Blocks(ReadPart(once, "xl/worksheets/sheet2.xml")));
}

TEST(OoxmlX14CrossSheet, CfApiRuleGoesToX14WithAnExcelFormId) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  cf::ConditionalFormat block;
  block.sqref = {cf::CFCellRange{{0, 2}, {2, 2}}};
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = 1;
  rule.formula1 = "Sheet1!$A$1>0";
  block.rules.push_back(rule);
  wb.sheet(1).mutable_conditional_formats().push_back(block);

  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::string xml = ReadPart(bytes, "xl/worksheets/sheet2.xml");
  EXPECT_EQ(xml.find("<conditionalFormatting "), std::string::npos) << xml;
  EXPECT_NE(xml.find("<x14:conditionalFormatting xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
                     "<x14:cfRule type=\"expression\" priority=\"1\" id=\"{00000000-000E-0000-0100-000001000000}\">"
                     "<xm:f>Sheet1!$A$1&gt;0</xm:f><x14:dxf/></x14:cfRule><xm:sqref>C1:C3</xm:sqref>"
                     "</x14:conditionalFormatting>"),
            std::string::npos)
      << xml;
  // The id is minted for the write; the model keeps none.
  EXPECT_TRUE(wb.sheet(1).conditional_formats()[0].rules[0].id.empty());
  const Workbook again = Load(bytes);
  const cf::CFRule* loaded = RuleWithPriority(again.sheet(1), 1);
  ASSERT_NE(loaded, nullptr);
  EXPECT_EQ(loaded->formula1.value_or(""), "Sheet1!$A$1>0");
}

// Excel writes an own-sheet reference inside x14 `<xm:f>` bare
// (x14_formula_edit_0_base.xlsx: `Sheet2!$A$1&lt;A1`).
TEST(OoxmlX14CrossSheet, CfOwnSheetReferenceStaysBareInX14) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  cf::ConditionalFormat block;
  block.sqref = {cf::CFCellRange{{0, 0}, {0, 0}}};
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  rule.priority = 1;
  rule.formula1 = "Sheet1!$A$1>SUM($B$1:$B$3)+C1";
  block.rules.push_back(rule);
  wb.sheet(1).mutable_conditional_formats().push_back(block);

  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::string xml = ReadPart(bytes, "xl/worksheets/sheet2.xml");
  EXPECT_NE(xml.find("<xm:f>Sheet1!$A$1&gt;SUM($B$1:$B$3)+C1</xm:f>"), std::string::npos) << xml;
  const Workbook again = Load(bytes);
  const cf::CFRule* loaded = RuleWithPriority(again.sheet(1), 1);
  ASSERT_NE(loaded, nullptr);
  EXPECT_EQ(loaded->formula1.value_or(""), "Sheet1!$A$1>SUM($B$1:$B$3)+C1");
}

TEST(OoxmlX14CrossSheet, CfRuleWhoseTargetSheetIsDeletedStaysInX14) {
  Workbook wb = Load(ReadFixture("structural_ref_m10b.xlsx"));
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));
  const std::string xml = ReadPart(Save(wb), "xl/worksheets/sheet1.xml");
  const std::map<std::string, std::string> blocks = X14Blocks(xml);
  ASSERT_EQ(blocks.count(kM6Id1), 1U) << xml;
  EXPECT_NE(blocks.at(kM6Id1).find("<xm:f>#REF!&gt;0</xm:f>"), std::string::npos) << blocks.at(kM6Id1);
  EXPECT_EQ(MainBlocks(xml).find("#REF!"), std::string::npos) << xml;
}

// Excel moves a bare `#REF!` rule it read from the extension to the main element.
TEST(OoxmlX14CrossSheet, CfBareRefErrorFromX14MovesToBaseElement) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.sheet(1).set_ext_lst_xml(
      "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
      "<x14:conditionalFormattings><x14:conditionalFormatting "
      "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
      "<x14:cfRule type=\"expression\" priority=\"1\" id=\"{00000000-000E-0000-0100-000001000000}\">"
      "<xm:f>#REF!&gt;0</xm:f><x14:dxf/></x14:cfRule><xm:sqref>C1:C3</xm:sqref>"
      "</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>");
  const Workbook loaded = Load(Save(wb));
  const cf::CFRule* rule = RuleWithPriority(loaded.sheet(1), 1);
  ASSERT_NE(rule, nullptr);
  EXPECT_EQ(rule->formula1.value_or(""), "#REF!>0");
  const std::string xml = ReadPart(Save(loaded), "xl/worksheets/sheet2.xml");
  EXPECT_NE(MainBlocks(xml).find("#REF!"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<x14:conditionalFormatting"), std::string::npos) << xml;
}

/// Every `<x14:dataValidation>` of the worksheet part `xml`, keyed by sqref,
/// without its `xr:uid` (the only attribute a save leaves out).
std::map<std::string, std::string> X14Dvs(const std::string& xml) {
  std::map<std::string, std::string> out;
  pugi::xml_document doc;
  if (!doc.load_string(xml.c_str())) {
    ADD_FAILURE() << "worksheet does not parse";
    return out;
  }
  for (pugi::xml_node ext : doc.child("worksheet").child("extLst").children("ext")) {
    for (pugi::xml_node dv : ext.child("x14:dataValidations").children("x14:dataValidation")) {
      dv.remove_attribute("xr:uid");
      out[dv.child("xm:sqref").text().get()] = io::raw_xml(dv);
    }
  }
  return out;
}

const DataValidation* DvAt(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  for (const DataValidation& v : sheet.validations()) {
    for (const MergeRange& r : v.ranges) {
      if (r.first_row == row && r.first_col == col) {
        return &v;
      }
    }
  }
  return nullptr;
}

TEST(OoxmlX14CrossSheet, DvExcelValidationsLoadIntoTheModel) {
  const Workbook wb = Load(ReadFixture("structural_ref_m7.xlsx"));
  const Sheet& sheet = wb.sheet(1);
  ASSERT_EQ(sheet.validations().size(), 5U);
  const DataValidation* own = DvAt(sheet, 0, 6);
  const DataValidation* list = DvAt(sheet, 3, 6);
  const DataValidation* custom = DvAt(sheet, 4, 6);
  ASSERT_NE(own, nullptr);
  ASSERT_NE(list, nullptr);
  ASSERT_NE(custom, nullptr);
  EXPECT_EQ(own->formula1, "$A$1:$A$3");
  EXPECT_EQ(list->formula1, "Sheet1!$A$1:$A$3");
  EXPECT_EQ(custom->formula1, "Sheet1!$A$1>0");
  EXPECT_TRUE(list->extra_attrs.empty()) << list->extra_attrs;
  EXPECT_EQ(sheet.ext_lst_xml().find("x14:dataValidations"), std::string::npos) << sheet.ext_lst_xml();
}

TEST(OoxmlX14CrossSheet, DvResaveWritesExcelXml) {
  const std::vector<std::uint8_t> excel = ReadFixture("structural_ref_m7.xlsx");
  const std::string excel_xml = ReadPart(excel, "xl/worksheets/sheet2.xml");
  const std::string saved_xml = ReadPart(Save(Load(excel)), "xl/worksheets/sheet2.xml");
  const std::map<std::string, std::string> expected = X14Dvs(excel_xml);
  const std::map<std::string, std::string> actual = X14Dvs(saved_xml);
  ASSERT_EQ(expected.size(), 2U);
  EXPECT_EQ(actual, expected) << saved_xml;
  EXPECT_EQ(Count(saved_xml, "Sheet1!$A$1:$A$3"), 1U) << saved_xml;
  EXPECT_EQ(Count(saved_xml, "Sheet1!$A$1&gt;0"), 1U) << saved_xml;
  EXPECT_NE(saved_xml.find("<dataValidations count=\"3\">"), std::string::npos) << saved_xml;
  EXPECT_NE(saved_xml.find("<x14:dataValidations count=\"2\" xmlns:xm="), std::string::npos) << saved_xml;
}

TEST(OoxmlX14CrossSheet, DvSecondRoundTripIsStable) {
  const std::vector<std::uint8_t> once = Save(Load(ReadFixture("structural_ref_m7.xlsx")));
  const std::vector<std::uint8_t> twice = Save(Load(once));
  const std::string sheet_twice = ReadPart(twice, "xl/worksheets/sheet2.xml");
  EXPECT_EQ(X14Dvs(sheet_twice), X14Dvs(ReadPart(once, "xl/worksheets/sheet2.xml")));
  EXPECT_EQ(Count(sheet_twice, "<x14:dataValidation "), 2U) << sheet_twice;
}

TEST(OoxmlX14CrossSheet, DvSheetOfCrossSheetValidationsOmitsTheMainElement) {
  const std::vector<std::uint8_t> excel = ReadFixture("structural_ref_m10b.xlsx");
  const std::string saved_xml = ReadPart(Save(Load(excel)), "xl/worksheets/sheet2.xml");
  EXPECT_EQ(saved_xml.find("<dataValidations"), std::string::npos) << saved_xml;
  EXPECT_EQ(X14Dvs(saved_xml), X14Dvs(ReadPart(excel, "xl/worksheets/sheet2.xml"))) << saved_xml;
  EXPECT_EQ(Count(saved_xml, "<x14:dataValidation "), 2U) << saved_xml;
}

TEST(OoxmlX14CrossSheet, DvContainerAttributesMoveToTheX14Container) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  DataValidation v;
  v.ranges.push_back(MergeRange{0, 0, 0, 0});
  v.type = 3;
  v.formula1 = "Sheet1!$A$1:$A$3";
  wb.sheet(1).mutable_validations().push_back(v);
  wb.sheet(1).set_validations_extra_attrs(" disablePrompts=\"1\"");
  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::string xml = ReadPart(bytes, "xl/worksheets/sheet2.xml");
  EXPECT_EQ(xml.find("<dataValidations"), std::string::npos) << xml;
  EXPECT_NE(xml.find(" disablePrompts=\"1\""), std::string::npos) << xml;
  const Workbook again = Load(bytes);
  ASSERT_EQ(again.sheet(1).validations().size(), 1U);
  EXPECT_EQ(again.sheet(1).validations_extra_attrs(), " disablePrompts=\"1\"");
}

TEST(OoxmlX14CrossSheet, DvValidationWhoseTargetSheetIsDeletedStaysInX14) {
  Workbook wb = Load(ReadFixture("structural_ref_m10b.xlsx"));
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));
  const std::string xml = ReadPart(Save(wb), "xl/worksheets/sheet1.xml");
  const std::map<std::string, std::string> dvs = X14Dvs(xml);
  ASSERT_EQ(dvs.count("G1"), 1U) << xml;
  EXPECT_NE(dvs.at("G1").find("<xm:f>#REF!</xm:f>"), std::string::npos) << dvs.at("G1");
  EXPECT_EQ(xml.find("<dataValidations"), std::string::npos) << xml;
}

TEST(OoxmlX14CrossSheet, DvBareRefErrorFromX14MovesToBaseElement) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.sheet(1).set_ext_lst_xml(
      "<extLst><ext uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
      "<x14:dataValidations count=\"1\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
      "<x14:dataValidation type=\"list\"><x14:formula1><xm:f>#REF!</xm:f></x14:formula1>"
      "<xm:sqref>G1</xm:sqref></x14:dataValidation>"
      "</x14:dataValidations></ext></extLst>");
  const Workbook loaded = Load(Save(wb));
  ASSERT_EQ(loaded.sheet(1).validations().size(), 1U);
  const std::string xml = ReadPart(Save(loaded), "xl/worksheets/sheet2.xml");
  EXPECT_NE(xml.find("<dataValidations"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<x14:dataValidation"), std::string::npos) << xml;
}

TEST(OoxmlX14CrossSheet, DvUnreadableEntrySurvivesAndAdoptedOnesAreNotDuplicated) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.sheet(1).set_ext_lst_xml(
      "<extLst><ext uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
      "<x14:dataValidations count=\"2\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
      "<x14:dataValidation type=\"list\"><x14:formula1><xm:f>Sheet1!$A$1:$A$3</xm:f></x14:formula1>"
      "<xm:sqref>G1</xm:sqref></x14:dataValidation>"
      "<x14:dataValidation type=\"list\"><x14:formula1><xm:f>Sheet1!$A$1:$A$3</xm:f></x14:formula1>"
      "<xm:sqref>not a range</xm:sqref></x14:dataValidation>"
      "</x14:dataValidations></ext></extLst>");
  const std::vector<std::uint8_t> once = Save(wb);
  const Workbook loaded = Load(once);
  ASSERT_EQ(loaded.sheet(1).validations().size(), 1U);
  EXPECT_NE(loaded.sheet(1).ext_lst_xml().find("not a range"), std::string::npos);
  EXPECT_EQ(loaded.sheet(1).ext_lst_xml().find("<xm:sqref>G1<"), std::string::npos);
  const std::string xml = ReadPart(Save(loaded), "xl/worksheets/sheet2.xml");
  EXPECT_EQ(Count(xml, "<xm:sqref>not a range</xm:sqref>"), 1U) << xml;
  EXPECT_EQ(Count(xml, "<xm:sqref>G1</xm:sqref>"), 1U) << xml;
  EXPECT_EQ(Count(xml, "<x14:dataValidation "), 2U) << xml;
  EXPECT_EQ(Count(xml, "<x14:dataValidations "), 1U) << xml;
}

TEST(OoxmlX14CrossSheet, CfCrossSheetDataBarIsOneX14Rule) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  cf::ConditionalFormat block;
  block.sqref = {cf::CFCellRange{{0, 0}, {4, 0}}};
  cf::CFRule rule;
  rule.type = cf::RuleType::DataBar;
  rule.priority = 2;
  cf::DataBarSpec bar;
  bar.min = {cf::CfvoType::Formula, "Sheet1!$A$1", true};
  bar.max = {cf::CfvoType::Max, "", true};
  bar.fill = {0x63, 0x8E, 0xC6, 255};
  bar.negative_fill = cf::Color{0xFF, 0x00, 0x00, 255};
  rule.data_bar = bar;
  block.rules.push_back(rule);
  wb.sheet(1).mutable_conditional_formats().push_back(block);

  const std::vector<std::uint8_t> bytes = Save(wb);
  const std::string xml = ReadPart(bytes, "xl/worksheets/sheet2.xml");
  EXPECT_EQ(xml.find("<conditionalFormatting "), std::string::npos) << xml;
  EXPECT_EQ(Count(xml, "<x14:cfRule "), 1U) << xml;
  EXPECT_NE(xml.find("<x14:cfRule type=\"dataBar\" priority=\"2\" id=\"{00000000-000E-0000-0100-000002000000}\">"),
            std::string::npos)
      << xml;
  const Workbook again = Load(bytes);
  const cf::CFRule* loaded = RuleWithPriority(again.sheet(1), 2);
  ASSERT_NE(loaded, nullptr);
  ASSERT_TRUE(loaded->data_bar.has_value());
  EXPECT_EQ(loaded->data_bar->min.value, "Sheet1!$A$1");
  EXPECT_EQ(loaded->data_bar->fill, bar.fill);
  EXPECT_EQ(loaded->data_bar->negative_fill, bar.negative_fill);
}

}  // namespace
}  // namespace formulon
