//
// Contracts the OOXML writer owes a consumer that reads only the bytes.
//
// Each group pins a place where two parts of one saved package have to
// agree with each other, which no single-element test can express:
//
//   * `<dimension>` must cover everything the same save writes into
//     `<sheetData>`, including a spill footprint the anchor declares.
//   * `<pageMargins>` is required by ECMA-376 to carry all six margins,
//     whichever subset of them a caller patched.
//   * A relationship and the element naming its rId must appear or
//     disappear together with the part they point at.

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "cell.h"
#include "gtest/gtest.h"
#include "io/external_links.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/package_diagnostics.h"
#include "io/passthrough_part.h"
#include "io/unknown_relationship.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

/// Saves `wb` and returns the text of `part`, failing the test when either
/// step does not complete.
std::string SavedPart(const Workbook& wb, const char* part) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  if (!saved) {
    return {};
  }
  std::string body;
  EXPECT_TRUE(test::extract_part(test::span_of(saved.value()), part, &body));
  return body;
}

/// Returns the `ref` attribute value of the first sheet's `<dimension>`.
std::string SavedDimensionRef(const Workbook& wb) {
  const std::string sheet_xml = SavedPart(wb, "xl/worksheets/sheet1.xml");
  const std::size_t open = sheet_xml.find("<dimension ref=\"");
  if (open == std::string::npos) {
    ADD_FAILURE() << "no <dimension> in: " << sheet_xml;
    return {};
  }
  const std::size_t start = open + std::string("<dimension ref=\"").size();
  const std::size_t end = sheet_xml.find('"', start);
  if (end == std::string::npos) {
    ADD_FAILURE() << "unterminated <dimension ref>: " << sheet_xml;
    return {};
  }
  return sheet_xml.substr(start, end - start);
}

// ---------------------------------------------------------------------------
// <dimension> versus <sheetData>.
// ---------------------------------------------------------------------------

TEST(OoxmlDimension, CoversStyleOnlyCells) {
  // A blank cell carrying a style index ships as `<c r="C3" s="N"/>`, so
  // the declared used range has to reach it. A consumer that trusts
  // `<dimension>` would otherwise stop short of a cell the same file
  // contains.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_xf_index(0, 2, 2, 7U)));

  const std::string sheet_xml = SavedPart(wb, "xl/worksheets/sheet1.xml");
  ASSERT_NE(sheet_xml.find("<c r=\"C3\""), std::string::npos) << sheet_xml;
  EXPECT_EQ(SavedDimensionRef(wb), "A1:C3") << sheet_xml;
}

TEST(OoxmlDimension, CoversSpillFootprint) {
  // `SEQUENCE(10)` down one column writes `<f t="array" ref="A1:A10">` on
  // the anchor while its phantoms live only in the spill table. The box
  // must follow the footprint the anchor itself declares.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_formula(0, 0, "=SEQUENCE(10)");
  std::vector<Value> cells;
  cells.reserve(10U);
  for (std::uint32_t i = 0; i < 10U; ++i) {
    cells.push_back(Value::number(static_cast<double>(i + 1U)));
  }
  ASSERT_TRUE(wb.sheet(0).commit_spill(0, 0, /*rows=*/10U, /*cols=*/1U, std::move(cells)));

  const std::string sheet_xml = SavedPart(wb, "xl/worksheets/sheet1.xml");
  ASSERT_NE(sheet_xml.find("ref=\"A1:A10\""), std::string::npos) << sheet_xml;
  EXPECT_EQ(SavedDimensionRef(wb), "A1:A10") << sheet_xml;
}

TEST(OoxmlDimension, EmptySheetStillDeclaresA1) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  EXPECT_EQ(SavedDimensionRef(wb), "A1");
}

// ---------------------------------------------------------------------------
// <pageMargins> required attributes.
// ---------------------------------------------------------------------------

struct WorkbookHandle {
  fm_workbook_t* handle = nullptr;
  ~WorkbookHandle() { fm_workbook_destroy(handle); }
  WorkbookHandle() = default;
  WorkbookHandle(const WorkbookHandle&) = delete;
  WorkbookHandle& operator=(const WorkbookHandle&) = delete;
};

TEST(OoxmlPageMargins, PartialPatchStillWritesEveryRequiredAttribute) {
  WorkbookHandle wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Engage two of the six. ECMA-376 makes all six required, so the other
  // four have to be filled from the schema defaults rather than omitted.
  fm_page_margins_t margins{};
  margins.left_engaged = 1;
  margins.left = 1.25;
  margins.footer_engaged = 1;
  margins.footer = 0.5;
  ASSERT_EQ(fm_sheet_set_page_margins(wb.handle, 0, &margins), 0);

  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb.handle, &bytes, &len), 0);
  ASSERT_NE(bytes, nullptr);
  const std::vector<std::uint8_t> package(bytes, bytes + len);
  fm_buffer_free(bytes);

  std::string sheet_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(package), "xl/worksheets/sheet1.xml", &sheet_xml));
  const std::size_t open = sheet_xml.find("<pageMargins");
  ASSERT_NE(open, std::string::npos) << sheet_xml;
  const std::string element = sheet_xml.substr(open, sheet_xml.find('>', open) - open);
  for (const char* attr : {"left=", "right=", "top=", "bottom=", "header=", "footer="}) {
    EXPECT_NE(element.find(attr), std::string::npos) << attr << " missing from " << element;
  }
  // The patched values survive; only the unengaged ones take a default.
  EXPECT_NE(element.find("left=\"1.25\""), std::string::npos) << element;
  EXPECT_NE(element.find("footer=\"0.5\""), std::string::npos) << element;
  EXPECT_NE(element.find("right=\"0.7\""), std::string::npos) << element;

  // Completing the element is a write-time step. Storing the defaults in
  // the model instead would report every margin as engaged here, and the
  // caller could no longer tell a stated margin from a defaulted one.
  fm_page_margins_t read{};
  ASSERT_EQ(fm_sheet_get_page_margins(wb.handle, 0, &read), 0);
  EXPECT_EQ(read.left_engaged, 1);
  EXPECT_EQ(read.footer_engaged, 1);
  EXPECT_EQ(read.right_engaged, 0);
  EXPECT_EQ(read.top_engaged, 0);
}

TEST(OoxmlPageMargins, IncompleteSourceElementIsCompletedOnSave) {
  // A third-party file can arrive with an element the schema would reject.
  // Completing only at authoring time would round-trip the defect; the
  // write-time step covers this path too.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).mutable_print_settings().page_margins_xml = "<pageMargins left=\"0.25\"/>";

  const std::string sheet_xml = SavedPart(wb, "xl/worksheets/sheet1.xml");
  const std::size_t open = sheet_xml.find("<pageMargins");
  ASSERT_NE(open, std::string::npos) << sheet_xml;
  const std::string element = sheet_xml.substr(open, sheet_xml.find('>', open) - open);
  for (const char* attr : {"left=", "right=", "top=", "bottom=", "header=", "footer="}) {
    EXPECT_NE(element.find(attr), std::string::npos) << attr << " missing from " << element;
  }
  EXPECT_NE(element.find("left=\"0.25\""), std::string::npos) << element;
}

// ---------------------------------------------------------------------------
// Relationships may not outlive the part they name.
// ---------------------------------------------------------------------------

/// Builds a workbook declaring one external link whose body part is or is
/// not present in `passthrough_parts()`.
Workbook WorkbookWithExternalLink(bool with_body_part) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));

  io::ExternalLinkRecord record;
  record.index = 1;
  record.rel_id = "rId9";
  record.part_path = "xl/externalLinks/externalLink1.xml";
  record.target = "file:///tmp/remote.xlsx";
  record.target_external = true;
  wb.set_external_links({record});

  if (with_body_part) {
    io::PassthroughPart body;
    body.path = "xl/externalLinks/externalLink1.xml";
    body.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
    const std::string xml = "<externalLink/>";
    body.bytes.assign(xml.begin(), xml.end());
    wb.set_passthrough_parts({std::move(body)});
  }
  return wb;
}

/// Builds a workbook whose sheet names a printer-settings part that is or
/// is not present in `passthrough_parts()`. The reader records the path
/// from the source sheet rels whether or not the archive carried the
/// part, so the absent case is reachable from a real (truncated) file.
Workbook WorkbookWithPrinterSettings(bool with_body_part) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));

  SheetPrintSettings& print = wb.sheet(0).mutable_print_settings();
  print.page_setup_xml = "<pageSetup orientation=\"portrait\" r:id=\"rId4\"/>";
  print.printer_settings_rid = "rId4";
  print.printer_settings_path = "xl/printerSettings/printerSettings1.bin";

  if (with_body_part) {
    io::PassthroughPart body;
    body.path = "xl/printerSettings/printerSettings1.bin";
    body.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
    body.bytes = {0x00, 0x01};
    wb.set_passthrough_parts({std::move(body)});
  }
  return wb;
}

TEST(OoxmlPrinterSettings, PresentBodyPartKeepsTheRelationship) {
  const Workbook wb = WorkbookWithPrinterSettings(/*with_body_part=*/true);
  const std::string sheet_rels = SavedPart(wb, "xl/worksheets/_rels/sheet1.xml.rels");
  const std::string sheet_xml = SavedPart(wb, "xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet_rels.find("printerSettings1.bin"), std::string::npos) << sheet_rels;
  EXPECT_NE(sheet_xml.find("<pageSetup"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("r:id=\"rId4\""), std::string::npos) << sheet_xml;
}

TEST(OoxmlPrinterSettings, MissingBodyPartDropsTheRelationshipAndTheReference) {
  const Workbook wb = WorkbookWithPrinterSettings(/*with_body_part=*/false);
  auto saved = io::write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;

  std::string sheet_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value().bytes), "xl/worksheets/sheet1.xml", &sheet_xml));
  // The relationship is gone, and so is the `<pageSetup r:id>` naming it:
  // an rId with no `<Relationship>` dangles just as badly as the reverse.
  EXPECT_EQ(sheet_xml.find("r:id="), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("<pageSetup orientation=\"portrait\"/>"), std::string::npos) << sheet_xml;
  std::string sheet_rels;
  if (test::extract_part(test::span_of(saved.value().bytes), "xl/worksheets/_rels/sheet1.xml.rels", &sheet_rels)) {
    EXPECT_EQ(sheet_rels.find("printerSettings"), std::string::npos) << sheet_rels;
  }
  EXPECT_GE(saved.value().diagnostics.dropped_relationship_count, 1U);
}

/// Resolves a `<Relationship Target>` against the part directory that owns
/// `rels_part`, giving the package path the target names.
///
/// A rels file at `<dir>/_rels/<name>.rels` describes a part in `<dir>`, so
/// targets are relative to `<dir>` and commonly step out of it (`../`).
std::string ResolveRelsTarget(std::string_view rels_part, std::string_view target) {
  const std::size_t rels_dir = rels_part.rfind("_rels/");
  std::string base(rels_part.substr(0, rels_dir == std::string_view::npos ? 0 : rels_dir));
  std::vector<std::string> segments;
  const std::string joined = base + std::string(target);
  std::size_t start = 0;
  while (start <= joined.size()) {
    const std::size_t slash = joined.find('/', start);
    const std::string segment = joined.substr(start, slash == std::string::npos ? std::string::npos : slash - start);
    if (segment == ".." && !segments.empty()) {
      segments.pop_back();
    } else if (!segment.empty() && segment != ".") {
      segments.push_back(segment);
    }
    if (slash == std::string::npos) {
      break;
    }
    start = slash + 1;
  }
  std::string out;
  for (const std::string& segment : segments) {
    if (!out.empty()) {
      out.push_back('/');
    }
    out.append(segment);
  }
  return out;
}

/// The file name a rels part describes: `<dir>/_rels/<name>.rels` belongs
/// to `<name>`, which `ResolveRelsTarget` then places in `<dir>`.
std::string OwnedPartName(std::string_view rels_part) {
  const std::size_t slash = rels_part.rfind('/');
  const std::string_view file = slash == std::string_view::npos ? rels_part : rels_part.substr(slash + 1);
  return std::string(file.substr(0, file.size() - std::string_view(".rels").size()));
}

/// Builds a workbook whose sheet names a DrawingML part that is or is not
/// present in `passthrough_parts()`. Like the printer settings above, the
/// reader records the target from the source sheet rels without checking
/// the archive carried it.
Workbook WorkbookWithDrawing(bool with_body_part) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_drawing_rel_target("xl/drawings/drawing1.xml");

  if (with_body_part) {
    io::PassthroughPart body;
    body.path = "xl/drawings/drawing1.xml";
    body.content_type = "application/vnd.openxmlformats-officedocument.drawing+xml";
    const std::string xml = "<xdr:wsDr/>";
    body.bytes.assign(xml.begin(), xml.end());
    wb.set_passthrough_parts({std::move(body)});
  }
  return wb;
}

TEST(OoxmlDrawing, PresentBodyPartKeepsTheRelationship) {
  const Workbook wb = WorkbookWithDrawing(/*with_body_part=*/true);
  const std::string sheet_rels = SavedPart(wb, "xl/worksheets/_rels/sheet1.xml.rels");
  const std::string sheet_xml = SavedPart(wb, "xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet_rels.find("drawing1.xml"), std::string::npos) << sheet_rels;
  EXPECT_NE(sheet_xml.find("<drawing r:id="), std::string::npos) << sheet_xml;
}

TEST(OoxmlDrawing, MissingBodyPartDropsTheRelationshipAndTheReference) {
  const Workbook wb = WorkbookWithDrawing(/*with_body_part=*/false);
  auto saved = io::write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;

  std::string sheet_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value().bytes), "xl/worksheets/sheet1.xml", &sheet_xml));
  EXPECT_EQ(sheet_xml.find("<drawing"), std::string::npos) << sheet_xml;
  std::string sheet_rels;
  if (test::extract_part(test::span_of(saved.value().bytes), "xl/worksheets/_rels/sheet1.xml.rels", &sheet_rels)) {
    EXPECT_EQ(sheet_rels.find("drawing1.xml"), std::string::npos) << sheet_rels;
  }
  EXPECT_GE(saved.value().diagnostics.dropped_relationship_count, 1U);
}

TEST(OoxmlSheetRels, EveryInternalRelationshipNamesAWrittenPart) {
  // The closure check the individual cases above cannot state: whatever a
  // sheet or the workbook references, no `<Relationship>` with an internal
  // target may name a part the same save did not write. Driven by a
  // workbook carrying one of each guarded edge with its payload absent,
  // plus the generated edges (tables, comments) whose parts this save
  // mints itself.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_drawing_rel_target("xl/drawings/drawing1.xml");
  SheetPrintSettings& print = wb.sheet(0).mutable_print_settings();
  print.printer_settings_rid = "rId4";
  print.printer_settings_path = "xl/printerSettings/printerSettings1.bin";
  io::UnknownRelationship rel;
  rel.id = "rId7";
  rel.type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
  rel.target = "xl/embeddings/oleObject1.bin";
  wb.sheet(0).set_unknown_relationships({rel});
  io::ExternalLinkRecord link;
  link.index = 1;
  link.rel_id = "rId9";
  link.part_path = "xl/externalLinks/externalLink1.xml";
  link.target = "file:///tmp/remote.xlsx";
  link.target_external = true;
  wb.set_external_links({link});

  auto saved = io::write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const std::vector<std::uint8_t>& package = saved.value().bytes;

  // Enumerated from the archive rather than from a fixed list, so a rels
  // part nobody thought to name is checked too -- which is the only way
  // this test can outlive the specific edges it was written for.
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(test::span_of(package))));
  const std::vector<std::string> entries = zip.list_entries();

  std::size_t internal_checked = 0;
  for (const std::string& rels_part : entries) {
    if (rels_part.size() < 5U || rels_part.compare(rels_part.size() - 5U, 5U, ".rels") != 0) {
      continue;
    }
    // A rels part describes one part, and `_rels/.rels` describes the
    // package root. An orphan rels file is the same defect seen from the
    // other side: the edges inside it point somewhere, but nothing points
    // at the part they belong to.
    if (rels_part != "_rels/.rels") {
      const std::string owner = ResolveRelsTarget(rels_part, OwnedPartName(rels_part));
      std::string owner_body;
      EXPECT_TRUE(test::extract_part(test::span_of(package), owner.c_str(), &owner_body))
          << rels_part << " describes " << owner << ", which the package does not contain";
    }
    std::string rels_xml;
    ASSERT_TRUE(test::extract_part(test::span_of(package), rels_part.c_str(), &rels_xml));
    pugi::xml_document doc;
    const pugi::xml_parse_result parsed = doc.load_buffer(rels_xml.data(), rels_xml.size());
    ASSERT_TRUE(parsed) << rels_part << ": " << parsed.description();
    for (pugi::xml_node r = doc.child("Relationships").child("Relationship"); r; r = r.next_sibling("Relationship")) {
      if (std::string_view(r.attribute("TargetMode").value()) == "External") {
        continue;
      }
      const std::string resolved = ResolveRelsTarget(rels_part, r.attribute("Target").value());
      std::string body;
      EXPECT_TRUE(test::extract_part(test::span_of(package), resolved.c_str(), &body))
          << rels_part << " -> " << resolved << " (from Target=\"" << r.attribute("Target").value() << "\")";
      ++internal_checked;
    }
  }
  // Without this the check would still pass on a package that emitted no
  // relationships at all, which is not the property being asserted.
  EXPECT_GT(internal_checked, 2U);
}

TEST(OoxmlSheetRels, DroppedUnknownRelationshipIsCounted) {
  // The guard that skips an unknown relationship whose payload is absent
  // predates the counter, so a sheet could lose an edge with every
  // diagnostic reading zero.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  io::UnknownRelationship rel;
  rel.id = "rId7";
  rel.type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
  rel.target = "xl/embeddings/oleObject1.bin";
  rel.target_external = false;
  wb.sheet(0).set_unknown_relationships({rel});

  auto saved = io::write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_GE(saved.value().diagnostics.dropped_relationship_count, 1U);
}

TEST(OoxmlExternalLinks, PresentBodyPartKeepsBothHalvesOfTheReference) {
  const Workbook wb = WorkbookWithExternalLink(/*with_body_part=*/true);
  const std::string workbook_xml = SavedPart(wb, "xl/workbook.xml");
  const std::string workbook_rels = SavedPart(wb, "xl/_rels/workbook.xml.rels");
  EXPECT_NE(workbook_xml.find("<externalReference"), std::string::npos) << workbook_xml;
  EXPECT_NE(workbook_rels.find("externalLinks/externalLink1.xml"), std::string::npos) << workbook_rels;
}

TEST(OoxmlExternalLinks, MissingBodyPartDropsBothHalvesAndCounts) {
  const Workbook wb = WorkbookWithExternalLink(/*with_body_part=*/false);
  auto saved = io::write_ooxml_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;

  std::string workbook_xml;
  std::string workbook_rels;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value().bytes), "xl/workbook.xml", &workbook_xml));
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value().bytes), "xl/_rels/workbook.xml.rels", &workbook_rels));

  // Neither half may survive: a `<Relationship>` with no part and an
  // `<externalReference r:id>` with no relationship are both repair-mode
  // triggers, so the two are gated on the same answer.
  EXPECT_EQ(workbook_xml.find("<externalReference"), std::string::npos) << workbook_xml;
  EXPECT_EQ(workbook_rels.find("externalLink"), std::string::npos) << workbook_rels;
  EXPECT_GE(saved.value().diagnostics.dropped_relationship_count, 1U);
}

// ---------------------------------------------------------------------------
// Unmodelled `<worksheet>` children survive in their schema position.
// ---------------------------------------------------------------------------

/// Names of `worksheet`'s element children, in document order.
std::vector<std::string> ChildElementNames(const pugi::xml_node& worksheet) {
  std::vector<std::string> names;
  for (pugi::xml_node child = worksheet.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element) {
      names.emplace_back(child.name());
    }
  }
  return names;
}

/// Parses `xml` and returns its `<worksheet>` child element names.
std::vector<std::string> WorksheetChildNames(const std::string& xml, pugi::xml_document* doc) {
  const pugi::xml_parse_result parsed = doc->load_buffer(xml.data(), xml.size());
  EXPECT_TRUE(parsed) << parsed.description();
  return ChildElementNames(doc->child("worksheet"));
}

TEST(WorksheetChildren, SchemaOrderTableHasNoDuplicates) {
  // The table is the single answer to "where does this element go", so a
  // duplicate name would make that answer ambiguous and silently move
  // whichever element resolved to the later slot.
  for (std::size_t i = 0; i < worksheet_child::kCount; ++i) {
    EXPECT_EQ(worksheet_child::slot_of(worksheet_child::kOrder[i]), i) << worksheet_child::kOrder[i];
  }
  EXPECT_EQ(worksheet_child::slot_of("noSuchElement"), worksheet_child::kCount);
}

TEST(WorksheetChildren, RealExcelSheetKeepsAnUnmodelledChildAndItsRelationship) {
  // An Excel-authored sheet carrying `<customProperties>` -- one of the
  // worksheet children no part of the model interprets, and one that also
  // names a relationship, so keeping the element only helps if the rel and
  // its part survive with it.
  // The only Excel-authored workbook in the tree whose worksheet carries a
  // child outside the set the model interprets. It lives in the vendored
  // corpus rather than under `fixtures/` because no hand-made fixture can
  // stand in here: the point is to check the placement against Excel's own
  // emission order, not against the table the writer uses.
  const std::string kCustomPropertiesWorkbook =
      std::string(FORMULON_FIXTURES_DIR) + "/../oracle/external/ironcalc/fixtures/calc_tests/UNICODE.xlsx";
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(kCustomPropertiesWorkbook);
  ASSERT_FALSE(bytes.empty());
  auto loaded = io::read_ooxml(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;

  const std::string saved_sheet = SavedPart(loaded.value().workbook, "xl/worksheets/sheet1.xml");
  ASSERT_NE(saved_sheet.find("<customProperties>"), std::string::npos) << saved_sheet;
  EXPECT_NE(saved_sheet.find("name=\"OrphanNamesChecked\""), std::string::npos) << saved_sheet;

  auto saved = io::write_ooxml(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  std::string sheet_rels;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/worksheets/_rels/sheet1.xml.rels", &sheet_rels));
  EXPECT_NE(sheet_rels.find("customProperty1.bin"), std::string::npos) << sheet_rels;

  // Excel puts `<customProperties>` after `<pageMargins>`; the saved sheet
  // has to agree, or the file is schema-invalid however well it preserves
  // the bytes.
  pugi::xml_document doc;
  const std::vector<std::string> names = WorksheetChildNames(saved_sheet, &doc);
  const auto margins = std::find(names.begin(), names.end(), "pageMargins");
  const auto props = std::find(names.begin(), names.end(), "customProperties");
  ASSERT_NE(margins, names.end());
  ASSERT_NE(props, names.end());
  EXPECT_LT(margins - names.begin(), props - names.begin());
}

TEST(WorksheetChildren, SavedSheetChildrenFollowTheSchemaSequence) {
  // Whatever the source contained, the children a save emits must be in
  // non-decreasing schema order. Driven by an Excel-authored workbook so
  // the ordering is checked against Excel's own emission rather than
  // against the same table the writer uses.
  for (const char* fixture : {"/excel/xlsb_fidelity_base.xlsx", "/excel/formula_corpus.xlsx"}) {
    const std::vector<std::uint8_t> bytes = test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + fixture);
    ASSERT_FALSE(bytes.empty()) << fixture;
    auto loaded = io::read_ooxml(test::span_of(bytes));
    ASSERT_TRUE(static_cast<bool>(loaded)) << fixture << ": " << loaded.error().message;

    const std::string saved_sheet = SavedPart(loaded.value().workbook, "xl/worksheets/sheet1.xml");
    pugi::xml_document doc;
    const std::vector<std::string> names = WorksheetChildNames(saved_sheet, &doc);
    std::size_t previous = 0;
    for (const std::string& name : names) {
      const std::size_t slot = worksheet_child::slot_of(name);
      ASSERT_NE(slot, worksheet_child::kCount) << fixture << ": unplaceable child " << name;
      EXPECT_GE(slot, previous) << fixture << ": " << name << " is out of schema order in " << saved_sheet;
      previous = slot;
    }
  }
}

TEST(WorksheetChildren, UnmodelledChildIsPreservedWithoutBeingNamedInAdvance) {
  // The sweep is by exclusion, so an element no release has modelled --
  // here `<sortState>`, which Excel writes for any plain sorted range --
  // round-trips without anyone having listed it.
  const std::vector<std::uint8_t> bytes =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_fidelity_base.xlsx");
  ASSERT_FALSE(bytes.empty());
  auto loaded = io::read_ooxml(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;

  Workbook wb = std::move(loaded.value().workbook);
  WorksheetRawExtensions& raw = wb.sheet(0).mutable_raw_extensions();
  raw.push_back(WorksheetRawChild{static_cast<std::uint32_t>(worksheet_child::slot_of("sortState")),
                                  "<sortState ref=\"A1:B3\"><sortCondition ref=\"A1:A3\"/></sortState>"});
  raw.push_back(WorksheetRawChild{static_cast<std::uint32_t>(worksheet_child::slot_of("cellWatches")),
                                  "<cellWatches><cellWatch r=\"B2\"/></cellWatches>"});
  std::stable_sort(raw.begin(), raw.end(),
                   [](const WorksheetRawChild& a, const WorksheetRawChild& b) { return a.slot < b.slot; });

  const std::string saved_sheet = SavedPart(wb, "xl/worksheets/sheet1.xml");
  EXPECT_NE(saved_sheet.find("<sortState ref=\"A1:B3\">"), std::string::npos) << saved_sheet;
  EXPECT_NE(saved_sheet.find("<cellWatch r=\"B2\"/>"), std::string::npos) << saved_sheet;

  pugi::xml_document doc;
  const std::vector<std::string> names = WorksheetChildNames(saved_sheet, &doc);
  std::size_t previous = 0;
  for (const std::string& name : names) {
    const std::size_t slot = worksheet_child::slot_of(name);
    ASSERT_NE(slot, worksheet_child::kCount) << "unplaceable child " << name;
    EXPECT_GE(slot, previous) << name << " is out of schema order in " << saved_sheet;
    previous = slot;
  }
}

}  // namespace
}  // namespace formulon
