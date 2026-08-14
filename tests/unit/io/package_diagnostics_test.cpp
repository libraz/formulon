//
// Reachability tests for the save/load loss counters that reach callers
// through `fm_read_diagnostics_t` / `fm_save_diagnostics_t`.
//
// The events these count are also emitted as structured-log records, but
// the shipping default for that log is off, so the counters are the only
// caller-visible signal. Each test therefore drives the workbook into the
// state that produces the event and asserts the counter, rather than
// asserting the counter's presence.

#include "io/package_diagnostics.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/cf_reader.h"
#include "io/default_content_type.h"
#include "io/ooxml/emission_plan.h"
#include "io/ooxml/workbook_xml_builder.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/passthrough_part.h"
#include "io/sheet_reader.h"
#include "io/tables_reader.h"
#include "io/unknown_relationship.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "support/ooxml_package_fixture.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

Workbook OneSheetWorkbook() {
  Workbook wb = Workbook::create();
  wb.set_cell_value(0, 0, 0, Value::number(1.0));
  return wb;
}

pugi::xml_document ParseWorksheet(const std::string& xml) {
  pugi::xml_document doc;
  const pugi::xml_parse_result parsed = doc.load_string(xml.c_str());
  EXPECT_TRUE(parsed) << parsed.description();
  return doc;
}

// ---------------------------------------------------------------------------
// Save counters
// ---------------------------------------------------------------------------

TEST(OoxmlWriteDiagnostics, CleanWorkbookReportsNothingLost) {
  const Workbook wb = OneSheetWorkbook();
  auto result = write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const WriteDiagnostics& d = result.value().diagnostics;
  EXPECT_EQ(d.downgraded_formula_count, 0U);
  EXPECT_EQ(d.deferred_feature_count, 0U);
  EXPECT_EQ(d.dropped_part_count, 0U);
  EXPECT_EQ(d.dropped_relationship_count, 0U);
  EXPECT_EQ(d.renumbered_part_count, 0U);
}

TEST(OoxmlWriteDiagnostics, PassthroughPartCollidingWithAGeneratedPathIsCounted) {
  Workbook wb = OneSheetWorkbook();
  // `xl/styles.xml` is always generated, so a preserved copy loses the
  // collision and is dropped from the package.
  PassthroughPart stale;
  stale.path = "xl/styles.xml";
  stale.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
  stale.bytes = {'<', '/', '>'};
  wb.set_passthrough_parts({std::move(stale)});

  auto result = write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.dropped_part_count, 1U);
  EXPECT_EQ(result.value().diagnostics.dropped_relationship_count, 0U);
}

TEST(OoxmlWriteDiagnostics, ASinglePackageScopeRelationshipDropCountsExactlyOne) {
  // `dropped_relationship_count` is fed from several sites; each has to
  // count one per dropped relationship, not one per pass. Drive exactly one
  // drop and assert equality rather than a threshold.
  Workbook wb = OneSheetWorkbook();
  UnknownRelationship package_rel;
  package_rel.id = "rId9";
  package_rel.type = "http://schemas.example.com/orphan";
  package_rel.target = "docProps/missing.xml";
  wb.set_unknown_package_rels({std::move(package_rel)});

  auto result = write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.dropped_relationship_count, 1U);
  EXPECT_EQ(result.value().diagnostics.dropped_part_count, 0U);
}

TEST(OoxmlWriteDiagnostics, ASingleWorkbookScopeRelationshipDropCountsExactlyOne) {
  Workbook wb = OneSheetWorkbook();
  UnknownRelationship workbook_rel;
  workbook_rel.id = "rId9";
  workbook_rel.type = "http://schemas.example.com/orphan";
  workbook_rel.target = "xl/missing.xml";
  wb.set_unknown_workbook_rels({std::move(workbook_rel)});

  auto result = write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.dropped_relationship_count, 1U);
  EXPECT_EQ(result.value().diagnostics.dropped_part_count, 0U);
}

TEST(OoxmlWriteDiagnostics, TwoScopesDroppingOneRelationshipEachSumToTwo) {
  Workbook wb = OneSheetWorkbook();
  UnknownRelationship package_rel;
  package_rel.id = "rId9";
  package_rel.type = "http://schemas.example.com/orphan";
  package_rel.target = "docProps/missing.xml";
  wb.set_unknown_package_rels({std::move(package_rel)});

  UnknownRelationship workbook_rel;
  workbook_rel.id = "rId9";
  workbook_rel.type = "http://schemas.example.com/orphan";
  workbook_rel.target = "xl/missing.xml";
  wb.set_unknown_workbook_rels({std::move(workbook_rel)});

  auto result = write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.dropped_relationship_count, 2U);
  EXPECT_EQ(result.value().diagnostics.dropped_part_count, 0U);
}

TEST(OoxmlWriteDiagnostics, OneLostPartRaisesBothTheDroppedPartAndRelationshipCounters) {
  // Pins the documented co-firing: the passthrough copy loses the path
  // collision, and the relationship that pointed at it can no longer be
  // emitted. That is ONE lost part observed twice, so the two counters read
  // 1 and 1 -- not 2 and 0, and not 1 and 0. A later change that
  // "de-duplicates" the apparent double count has to break this test.
  Workbook wb = OneSheetWorkbook();
  PassthroughPart stale;
  stale.path = "xl/styles.xml";
  stale.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
  stale.bytes = {'<', '/', '>'};
  wb.set_passthrough_parts({std::move(stale)});

  UnknownRelationship rel_to_dropped_part;
  rel_to_dropped_part.id = "rId9";
  rel_to_dropped_part.type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
  rel_to_dropped_part.target = "xl/styles.xml";
  wb.set_unknown_workbook_rels({std::move(rel_to_dropped_part)});

  auto result = write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const WriteDiagnostics& d = result.value().diagnostics;
  EXPECT_EQ(d.dropped_part_count, 1U);
  EXPECT_EQ(d.dropped_relationship_count, 1U);
  // Nothing else moved: one part, one relationship, no feature loss.
  EXPECT_EQ(d.downgraded_formula_count, 0U);
  EXPECT_EQ(d.deferred_feature_count, 0U);
  EXPECT_EQ(d.renumbered_part_count, 0U);
}

TEST(OoxmlWriteDiagnostics, TableWithoutAnIdIsCountedAsRenumbered) {
  Workbook wb = OneSheetWorkbook();
  TableMetadata table;
  table.id = 0;  // no usable id: the writer mints one
  table.name = "Table1";
  table.display_name = "Table1";
  table.ref = "A1:B2";
  table.sheet_index = 0;
  wb.mutable_tables().push_back(std::move(table));

  auto result = write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.renumbered_part_count, 1U);
  // The table part itself still ships; only its id changed.
  EXPECT_EQ(result.value().diagnostics.dropped_part_count, 0U);
}

TEST(OoxmlWriteDiagnostics, TableNamingARemovedSheetFailsTheSaveRatherThanCountingALoss) {
  // The writer refuses this workbook outright, which is why
  // `deferred_feature_count` has no OOXML source: the caller loses the save,
  // not the table.
  Workbook wb = OneSheetWorkbook();
  TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.ref = "A1:B2";
  table.sheet_index = 7;
  wb.mutable_tables().push_back(std::move(table));

  auto result = write_ooxml_with_result(wb);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoWriteFailed);
}

TEST(OoxmlWriteDiagnostics, ANullSinkDiscardsTheCountsAndChangesNothingElse) {
  // Both halves of this feature accept a NULL sink. The write side is
  // header-exposed, so a caller that declines to collect must degrade
  // rather than crash -- and must still produce identical bytes.
  Workbook wb = OneSheetWorkbook();
  PassthroughPart stale;
  stale.path = "xl/styles.xml";
  stale.bytes = {'<', '/', '>'};
  UnknownRelationship orphan;
  orphan.id = "rId9";
  orphan.type = "http://schemas.example.com/orphan";
  orphan.target = "xl/missing.xml";
  TableMetadata table;
  table.id = 0;
  table.name = "Table1";
  table.ref = "A1:B2";
  table.sheet_index = 0;
  wb.set_passthrough_parts({std::move(stale)});
  wb.set_unknown_workbook_rels({std::move(orphan)});
  wb.mutable_tables().push_back(std::move(table));

  auto collected = write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(collected)) << collected.error().message;
  ASSERT_GT(collected.value().diagnostics.dropped_part_count, 0U);

  WriteDiagnostics unused;
  const EmissionPlan plan = BuildEmissionPlan(wb, /*generated_shared_strings=*/false, nullptr);
  const std::string package_rels = BuildPackageRels(wb, plan, nullptr);
  const std::string workbook_rels = BuildWorkbookRels(wb.sheet_count(), plan, wb, nullptr);
  EXPECT_FALSE(package_rels.empty());
  EXPECT_FALSE(workbook_rels.empty());
  EXPECT_EQ(unused.dropped_part_count, 0U);

  // The same plan built with a sink must agree on what it emitted.
  WriteDiagnostics collected_again;
  const EmissionPlan planned = BuildEmissionPlan(wb, /*generated_shared_strings=*/false, &collected_again);
  EXPECT_EQ(BuildPackageRels(wb, planned, &collected_again), package_rels);
  EXPECT_EQ(BuildWorkbookRels(wb.sheet_count(), planned, wb, &collected_again), workbook_rels);
  EXPECT_EQ(collected_again.dropped_part_count, 1U);
  EXPECT_EQ(collected_again.dropped_relationship_count, 1U);
  EXPECT_EQ(collected_again.renumbered_part_count, 1U);
}

// ---------------------------------------------------------------------------
// Save counters: the XLSB half of the fields both writers produce
// ---------------------------------------------------------------------------

TEST(XlsbWriteDiagnostics, CleanWorkbookReportsNothingLost) {
  const Workbook wb = OneSheetWorkbook();
  auto result = xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const WriteDiagnostics& d = result.value().diagnostics;
  EXPECT_EQ(d.downgraded_formula_count, 0U);
  EXPECT_EQ(d.deferred_feature_count, 0U);
  EXPECT_EQ(d.dropped_part_count, 0U);
  EXPECT_EQ(d.dropped_relationship_count, 0U);
  EXPECT_EQ(d.renumbered_part_count, 0U);
}

TEST(XlsbWriteDiagnostics, PassthroughPartCollidingWithAGeneratedPathIsCounted) {
  Workbook wb = OneSheetWorkbook();
  PassthroughPart stale;
  stale.path = "xl/workbook.bin";
  stale.content_type = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
  stale.bytes = {0x00};
  wb.set_passthrough_parts({std::move(stale)});

  auto result = xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.dropped_part_count, 1U);
}

TEST(XlsbWriteDiagnostics, TwoPassthroughEntriesClaimingOnePathCountOneDrop) {
  // The second site inside the XLSB emission plan: a duplicate passthrough
  // path. The first entry wins, the second is dropped exactly once.
  Workbook wb = OneSheetWorkbook();
  PassthroughPart first;
  first.path = "xl/custom/keeper.bin";
  first.bytes = {0x01};
  PassthroughPart duplicate = first;
  duplicate.bytes = {0x02};
  wb.set_default_content_types({DefaultContentType{"bin", "application/octet-stream"}});
  wb.set_passthrough_parts({std::move(first), std::move(duplicate)});

  auto result = xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.dropped_part_count, 1U);
}

TEST(XlsbWriteDiagnostics, EachRelationshipScopeDropsExactlyOne) {
  // Package, workbook and sheet scope each feed `dropped_relationship_count`.
  // Sheet scope has no OOXML counterpart, so it is the one that disappears
  // if the field is unified from the OOXML side alone.
  UnknownRelationship orphan;
  orphan.id = "rId9";
  orphan.type = "http://schemas.example.com/orphan";

  {
    Workbook wb = OneSheetWorkbook();
    UnknownRelationship rel = orphan;
    rel.target = "docProps/missing.xml";
    wb.set_unknown_package_rels({std::move(rel)});
    auto result = xlsb::write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
    EXPECT_EQ(result.value().diagnostics.dropped_relationship_count, 1U) << "package scope";
  }
  {
    Workbook wb = OneSheetWorkbook();
    UnknownRelationship rel = orphan;
    rel.target = "xl/missing.bin";
    wb.set_unknown_workbook_rels({std::move(rel)});
    auto result = xlsb::write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
    EXPECT_EQ(result.value().diagnostics.dropped_relationship_count, 1U) << "workbook scope";
  }
  {
    Workbook wb = OneSheetWorkbook();
    UnknownRelationship rel = orphan;
    rel.target = "xl/drawings/missing.bin";
    wb.sheet(0).set_unknown_relationships({std::move(rel)});
    auto result = xlsb::write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
    EXPECT_EQ(result.value().diagnostics.dropped_relationship_count, 1U) << "sheet scope";
  }
}

// ---------------------------------------------------------------------------
// Load counters
// ---------------------------------------------------------------------------

TEST(OoxmlReadDiagnostics, ConditionalFormatBlockWithAnUnusableSqrefIsCounted) {
  const pugi::xml_document doc = ParseWorksheet(
      "<worksheet>"
      "<conditionalFormatting><cfRule type=\"expression\" priority=\"1\"/></conditionalFormatting>"
      "<conditionalFormatting sqref=\"!!!\"><cfRule type=\"expression\" priority=\"2\"/></conditionalFormatting>"
      "<conditionalFormatting sqref=\"A1:A2\"><cfRule type=\"expression\" priority=\"3\"/></conditionalFormatting>"
      "</worksheet>");
  ReadDiagnostics d;
  auto cfs = read_conditional_formats(doc.child("worksheet"), &d);
  ASSERT_TRUE(static_cast<bool>(cfs));
  EXPECT_EQ(cfs.value().size(), 1U);
  EXPECT_EQ(d.skipped_feature_count, 2U);
  EXPECT_EQ(d.unknown_content_type_count, 0U);
}

// `skipped_feature_count` is fed from eight distinct call sites (two in the
// CF reader, six in the sheet-overlay readers). Each has to contribute
// exactly one per dropped entry: a site that counts twice, or a site that
// forgets, is invisible to any test that only checks "greater than zero".
TEST(OoxmlReadDiagnostics, EachOverlaySkipSiteCountsExactlyOne) {
  struct Case {
    const char* what;
    const char* sheet_xml;
  };
  const Case merge_cases[] = {
      {"merge ref missing", "<worksheet><mergeCells><mergeCell ref=\"\"/></mergeCells></worksheet>"},
      {"merge ref unparseable", "<worksheet><mergeCells><mergeCell ref=\"nope\"/></mergeCells></worksheet>"},
  };
  for (const Case& c : merge_cases) {
    ReadDiagnostics d;
    const pugi::xml_document doc = ParseWorksheet(c.sheet_xml);
    auto merges = read_merges(doc.child("worksheet"), &d);
    ASSERT_TRUE(static_cast<bool>(merges)) << c.what;
    EXPECT_TRUE(merges.value().empty()) << c.what;
    EXPECT_EQ(d.skipped_feature_count, 1U) << c.what;
  }

  const Case hyperlink_cases[] = {
      {"hyperlink ref missing", "<worksheet><hyperlinks><hyperlink ref=\"\"/></hyperlinks></worksheet>"},
      {"hyperlink ref unparseable", "<worksheet><hyperlinks><hyperlink ref=\"nope\"/></hyperlinks></worksheet>"},
  };
  for (const Case& c : hyperlink_cases) {
    ReadDiagnostics d;
    const pugi::xml_document doc = ParseWorksheet(c.sheet_xml);
    auto links = read_hyperlinks(doc.child("worksheet"), &d);
    ASSERT_TRUE(static_cast<bool>(links)) << c.what;
    EXPECT_TRUE(links.value().empty()) << c.what;
    EXPECT_EQ(d.skipped_feature_count, 1U) << c.what;
  }

  const Case validation_cases[] = {
      {"validation sqref missing",
       "<worksheet><dataValidations><dataValidation sqref=\"\"/></dataValidations></worksheet>"},
      {"validation sqref unparseable",
       "<worksheet><dataValidations><dataValidation sqref=\"nope\"/></dataValidations></worksheet>"},
  };
  for (const Case& c : validation_cases) {
    ReadDiagnostics d;
    const pugi::xml_document doc = ParseWorksheet(c.sheet_xml);
    auto validations = read_data_validations(doc.child("worksheet"), &d);
    ASSERT_TRUE(static_cast<bool>(validations)) << c.what;
    EXPECT_TRUE(validations.value().empty()) << c.what;
    EXPECT_EQ(d.skipped_feature_count, 1U) << c.what;
  }

  const Case cf_cases[] = {
      {"cf sqref missing",
       "<worksheet><conditionalFormatting><cfRule type=\"expression\" priority=\"1\"/>"
       "</conditionalFormatting></worksheet>"},
      {"cf sqref unparseable",
       "<worksheet><conditionalFormatting sqref=\"!!!\"><cfRule type=\"expression\" priority=\"1\"/>"
       "</conditionalFormatting></worksheet>"},
  };
  for (const Case& c : cf_cases) {
    ReadDiagnostics d;
    const pugi::xml_document doc = ParseWorksheet(c.sheet_xml);
    auto cfs = read_conditional_formats(doc.child("worksheet"), &d);
    ASSERT_TRUE(static_cast<bool>(cfs)) << c.what;
    EXPECT_TRUE(cfs.value().empty()) << c.what;
    EXPECT_EQ(d.skipped_feature_count, 1U) << c.what;
  }
}

TEST(OoxmlReadDiagnostics, OneSinkAccumulatesAcrossEveryOverlayReader) {
  // The per-site counts above only add up if the reader threads one sink
  // through all of them.
  ReadDiagnostics d;
  {
    const pugi::xml_document doc = ParseWorksheet(
        "<worksheet><mergeCells><mergeCell ref=\"\"/><mergeCell ref=\"nope\"/>"
        "<mergeCell ref=\"A1:B2\"/></mergeCells></worksheet>");
    auto merges = read_merges(doc.child("worksheet"), &d);
    ASSERT_TRUE(static_cast<bool>(merges));
    EXPECT_EQ(merges.value().size(), 1U);
  }
  {
    const pugi::xml_document doc =
        ParseWorksheet("<worksheet><hyperlinks><hyperlink ref=\"\"/></hyperlinks></worksheet>");
    auto links = read_hyperlinks(doc.child("worksheet"), &d);
    ASSERT_TRUE(static_cast<bool>(links));
    EXPECT_TRUE(links.value().empty());
  }
  {
    const pugi::xml_document doc =
        ParseWorksheet("<worksheet><dataValidations><dataValidation sqref=\"\"/></dataValidations></worksheet>");
    auto validations = read_data_validations(doc.child("worksheet"), &d);
    ASSERT_TRUE(static_cast<bool>(validations));
    EXPECT_TRUE(validations.value().empty());
  }
  EXPECT_EQ(d.skipped_feature_count, 4U);
}

TEST(OoxmlReadDiagnostics, CleanPackageReportsNothingSkipped) {
  const std::string content_types = test::OoxmlContentTypes(test::kXlsxWorkbookContentType);
  const std::vector<std::uint8_t> bytes = test::BuildOoxmlPackage(
      content_types,
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/>"
      "<mergeCells><mergeCell ref=\"A1:B2\"/></mergeCells></worksheet>");
  auto result = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.skipped_feature_count, 0U);
  EXPECT_EQ(result.value().diagnostics.unknown_content_type_count, 0U);
  EXPECT_EQ(result.value().workbook.sheet(0).merges().size(), 1U);
}

TEST(OoxmlReadDiagnostics, SkippedOverlayEntriesReachTheReadResult) {
  // Per-reader counters only matter if `read_ooxml` threads one sink
  // through every sheet reader, so drive one bad entry per overlay kind.
  const std::string content_types = test::OoxmlContentTypes(test::kXlsxWorkbookContentType);
  const std::vector<std::uint8_t> bytes = test::BuildOoxmlPackage(
      content_types,
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/>"
      "<mergeCells><mergeCell ref=\"nope\"/></mergeCells>"
      "<hyperlinks><hyperlink ref=\"\"/></hyperlinks>"
      "<conditionalFormatting><cfRule type=\"expression\" priority=\"1\"/></conditionalFormatting>"
      "<dataValidations><dataValidation sqref=\"\"/></dataValidations>"
      "</worksheet>");
  auto result = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.skipped_feature_count, 4U);
  EXPECT_TRUE(result.value().workbook.sheet(0).merges().empty());
}

TEST(OoxmlReadDiagnostics, UnrecognisedWorkbookContentTypeReachesTheReadResult) {
  const std::string content_types = test::OoxmlContentTypes("application/vnd.bogus.foo+xml");
  const std::vector<std::uint8_t> bytes = test::BuildOoxmlPackage(content_types, test::kEmptySheetXml);
  auto result = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.unknown_content_type_count, 1U);
  EXPECT_EQ(result.value().diagnostics.skipped_feature_count, 0U);
}

}  // namespace
}  // namespace io
}  // namespace formulon
