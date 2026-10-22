//
// Workbook globals through `xl/workbook.bin`: BrtCalcProp (`<calcPr>`),
// BrtWbProp (`<workbookPr>`), BrtBookView (`<workbookView>`),
// BrtBookProtection(Iso) and the BrtName hidden flag and comment. Each survives an xlsx-sourced model saved as .xlsb
// and an .xlsb re-save; the Excel twin fixtures pin the bytes.

#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <vector>

#include "calc_settings.h"
#include "defined_name.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/workbook_props_records.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

ByteSpan Span(const std::vector<std::uint8_t>& v) {
  return ByteSpan{v.data(), v.size()};
}

std::vector<std::uint8_t> ReadFixture(const std::string& relative) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/" + relative;
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
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(f);
  return out;
}

Workbook FromXlsx(const std::vector<std::uint8_t>& bytes) {
  auto r = read_ooxml(Span(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

Workbook FromXlsb(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(Span(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

/// The model as an xlsx save and reload leaves it.
Workbook ThroughXlsx(const Workbook& wb) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? FromXlsx(saved.value()) : Workbook::create_empty();
}

std::vector<std::uint8_t> ToXlsb(const Workbook& wb, std::uint32_t* deferred = nullptr) {
  auto saved = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  if (!saved) {
    return {};
  }
  if (deferred != nullptr) {
    *deferred = saved.value().diagnostics.deferred_feature_count;
  }
  return std::move(saved.value().bytes);
}

/// Payloads of the workbook.bin records of `type`, in stream order.
std::vector<std::vector<std::uint8_t>> WorkbookRecords(const std::vector<std::uint8_t>& package, std::uint16_t type) {
  std::vector<std::vector<std::uint8_t>> out;
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(Span(package))));
  auto book = zip.read_entry("xl/workbook.bin");
  if (!book) {
    ADD_FAILURE() << "no workbook.bin";
    return out;
  }
  ByteSpan cursor = Span(book.value());
  while (cursor.size > 0U) {
    auto record = read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    if (record.value().type == type) {
      const ByteSpan p = record.value().payload;
      out.emplace_back(p.data, p.data + p.size);
    }
  }
  return out;
}

void ExpectSameCalc(const Workbook& actual, const Workbook& expected, const char* where) {
  EXPECT_EQ(actual.calc_mode(), expected.calc_mode()) << where;
  const IterativeOptions& ai = actual.iterative_options();
  const IterativeOptions& ei = expected.iterative_options();
  EXPECT_EQ(ai.enabled, ei.enabled) << where;
  EXPECT_EQ(ai.max_iterations, ei.max_iterations) << where;
  EXPECT_EQ(ai.max_change, ei.max_change) << where;
  const CalcOptions& a = actual.calc_options();
  const CalcOptions& e = expected.calc_options();
  EXPECT_EQ(a.full_precision, e.full_precision) << where;
  EXPECT_EQ(a.calc_on_save, e.calc_on_save) << where;
  EXPECT_EQ(a.ref_mode, e.ref_mode) << where;
  EXPECT_EQ(a.has_calc_id, e.has_calc_id) << where;
  EXPECT_EQ(a.calc_id, e.calc_id) << where;
  EXPECT_EQ(a.concurrent_calc, e.concurrent_calc) << where;
  EXPECT_EQ(a.has_concurrent_manual_count, e.has_concurrent_manual_count) << where;
  EXPECT_EQ(a.concurrent_manual_count, e.concurrent_manual_count) << where;
  EXPECT_EQ(a.force_full_calc, e.force_full_calc) << where;
  EXPECT_EQ(a.full_calc_on_load, e.full_calc_on_load) << where;
  EXPECT_EQ(a.calc_completed, e.calc_completed) << where;
}

/// Every `<calcPr>` attribute off its default.
Workbook NonDefaultCalcWorkbook(CalcMode mode) {
  Workbook wb = Workbook::create();
  wb.set_calc_mode(mode);
  IterativeOptions iter;
  iter.enabled = true;
  iter.max_iterations = 250U;
  iter.max_change = 0.0005;
  wb.set_iterative_options(iter);
  CalcOptions& calc = wb.mutable_calc_options();
  calc.full_precision = false;
  calc.calc_on_save = false;
  calc.ref_mode = RefMode::kR1C1;
  calc.has_calc_id = true;
  calc.calc_id = 191029U;
  calc.concurrent_calc = false;
  calc.has_concurrent_manual_count = true;
  calc.concurrent_manual_count = 3U;
  calc.force_full_calc = true;
  calc.full_calc_on_load = true;
  calc.calc_completed = false;
  return wb;
}

TEST(XlsbWorkbookFidelity, CalcPropertiesSurviveXlsxToXlsbAndXlsbResave) {
  for (const CalcMode mode : {CalcMode::kManual, CalcMode::kAutoNoTable}) {
    const Workbook source = ThroughXlsx(NonDefaultCalcWorkbook(mode));
    ASSERT_EQ(source.calc_mode(), mode);
    std::uint32_t deferred = 99U;
    const std::vector<std::uint8_t> first = ToXlsb(source, &deferred);
    EXPECT_EQ(deferred, 0U);
    const Workbook loaded = FromXlsb(first);
    ExpectSameCalc(loaded, source, "xlsx -> xlsb");
    ExpectSameCalc(FromXlsb(ToXlsb(loaded)), source, "xlsb -> xlsb");
  }
}

// Layout check against [MS-XLSB] 3.7.12's example record.
TEST(XlsbWorkbookFidelity, CalcPropMatchesTheSpecExample) {
  XlsbCalcProps props;
  props.options.has_calc_id = true;
  props.options.calc_id = 0x0001EB1DU;
  std::vector<std::uint8_t> out;
  emit_calc_prop(out, props);
  ByteSpan cursor = Span(out);
  auto record = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(record));
  EXPECT_EQ(record.value().type, kBrtCalcProp);
  const std::vector<std::uint8_t> expected = {0x1D, 0xEB, 0x01, 0x00, 0x01, 0x00, 0x00, 0x00, 0x64,
                                              0x00, 0x00, 0x00, 0xFC, 0xA9, 0xF1, 0xD2, 0x4D, 0x62,
                                              0x50, 0x3F, 0x01, 0x00, 0x00, 0x00, 0x6A, 0x00};
  const ByteSpan p = record.value().payload;
  EXPECT_EQ(std::vector<std::uint8_t>(p.data, p.data + p.size), expected);
}

TEST(XlsbWorkbookFidelity, DefaultCalcSettingsWriteNoCalcProp) {
  EXPECT_TRUE(WorkbookRecords(ToXlsb(Workbook::create()), kBrtCalcProp).empty());
}

// Excel's .xlsb twin of an .xlsx pins both records: the xlsx-sourced save
// writes Excel's bytes (cUserThreadCount aside, which Excel fills with the
// machine's core count while fUserSetThreadCount is clear) and both packages
// read to the same model.
TEST(XlsbWorkbookFidelity, WbPropAndCalcPropMatchExcelsTwin) {
  const std::string base = "excel/win/book_protection/com_s1_struct_win_nopw";
  const std::vector<std::uint8_t> excel_xlsb = ReadFixture(base + ".xlsb");
  const Workbook from_xlsx = FromXlsx(ReadFixture(base + ".xlsx"));
  const Workbook from_xlsb = FromXlsb(excel_xlsb);
  EXPECT_EQ(from_xlsb.workbook_pr_xml(), from_xlsx.workbook_pr_xml());
  EXPECT_EQ(from_xlsb.workbook_pr_xml(), "<workbookPr defaultThemeVersion=\"202300\"/>");
  ExpectSameCalc(from_xlsb, from_xlsx, "excel xlsb vs xlsx");

  const std::vector<std::uint8_t> ours = ToXlsb(from_xlsx);
  EXPECT_EQ(WorkbookRecords(ours, static_cast<std::uint16_t>(XlsbRecordType::BrtWbProp)),
            WorkbookRecords(excel_xlsb, static_cast<std::uint16_t>(XlsbRecordType::BrtWbProp)));
  auto our_calc = WorkbookRecords(ours, kBrtCalcProp);
  auto excel_calc = WorkbookRecords(excel_xlsb, kBrtCalcProp);
  ASSERT_EQ(our_calc.size(), 1U);
  ASSERT_EQ(excel_calc.size(), 1U);
  ASSERT_EQ(our_calc[0].size(), 26U);
  ASSERT_EQ(excel_calc[0].size(), 26U);
  for (std::size_t i = 20; i < 24; ++i) {
    our_calc[0][i] = excel_calc[0][i] = 0U;
  }
  EXPECT_EQ(our_calc[0], excel_calc[0]);
}

TEST(XlsbWorkbookFidelity, WorkbookPrSurvivesXlsxToXlsbAndXlsbResave) {
  // Every BrtWbProp field off its default, attributes in schema order.
  const std::string xml =
      "<workbookPr date1904=\"1\" showObjects=\"placeholders\" showBorderUnselectedTables=\"0\" "
      "filterPrivacy=\"1\" promptedSolutions=\"1\" showInkAnnotation=\"0\" backupFile=\"1\" "
      "saveExternalLinkValues=\"0\" updateLinks=\"always\" codeName=\"ThisWorkbook\" hidePivotFieldList=\"1\" "
      "showPivotChartFilter=\"1\" publishItems=\"1\" checkCompatibility=\"1\" autoCompressPictures=\"0\" "
      "refreshAllConnections=\"1\" defaultThemeVersion=\"166925\"/>";
  Workbook wb = Workbook::create();
  wb.set_workbook_pr_xml(xml);
  wb.set_date1904(true);
  const Workbook source = ThroughXlsx(wb);
  ASSERT_EQ(source.workbook_pr_xml(), xml);
  std::uint32_t deferred = 99U;
  const Workbook loaded = FromXlsb(ToXlsb(source, &deferred));
  EXPECT_EQ(deferred, 0U);
  EXPECT_TRUE(loaded.date1904());
  EXPECT_EQ(loaded.workbook_pr_xml(), xml);
  const Workbook resaved = FromXlsb(ToXlsb(loaded));
  EXPECT_TRUE(resaved.date1904());
  EXPECT_EQ(resaved.workbook_pr_xml(), xml);
}

TEST(XlsbWorkbookFidelity, WorkbookPrWithoutAFieldIsReportedDeferred) {
  std::uint32_t baseline = 99U;
  ToXlsb(Workbook::create(), &baseline);
  for (const char* xml : {"<workbookPr dateCompatibility=\"0\"/>", "<workbookPr allowRefreshQuery=\"1\"/>"}) {
    Workbook wb = Workbook::create();
    wb.set_workbook_pr_xml(xml);
    std::uint32_t deferred = 0U;
    ToXlsb(wb, &deferred);
    EXPECT_EQ(deferred, baseline + 1U) << xml;
  }
  // Their default values lose nothing.
  Workbook wb = Workbook::create();
  wb.set_workbook_pr_xml("<workbookPr dateCompatibility=\"1\" allowRefreshQuery=\"0\"/>");
  std::uint32_t deferred = 0U;
  ToXlsb(wb, &deferred);
  EXPECT_EQ(deferred, baseline);
}

TEST(XlsbWorkbookFidelity, WorkbookProtectionHashesSurviveBothPaths) {
  // Modern hashes for both passwords plus every lock flag.
  const std::string xml =
      "<workbookProtection workbookAlgorithmName=\"SHA-512\" workbookHashValue=\"Zm9vYmFy\" "
      "workbookSaltValue=\"Zm9v\" workbookSpinCount=\"100000\" revisionsAlgorithmName=\"SHA-256\" "
      "revisionsHashValue=\"Zm9vYg==\" revisionsSaltValue=\"Zg==\" revisionsSpinCount=\"5000\" "
      "lockStructure=\"1\" lockWindows=\"1\" lockRevision=\"1\"/>";
  Workbook wb = Workbook::create();
  wb.set_workbook_protection_xml(xml);
  const Workbook source = ThroughXlsx(wb);
  ASSERT_EQ(source.workbook_protection_xml(), xml);
  std::uint32_t deferred = 99U;
  const std::vector<std::uint8_t> first = ToXlsb(source, &deferred);
  EXPECT_EQ(deferred, 0U);
  const auto book = WorkbookRecords(first, 534U);
  ASSERT_EQ(book.size(), 1U);
  EXPECT_EQ(book[0], (std::vector<std::uint8_t>{0x00, 0x00, 0x00, 0x00, 0x07, 0x00}));
  EXPECT_EQ(WorkbookRecords(first, 677U).size(), 1U);
  const Workbook loaded = FromXlsb(first);
  EXPECT_EQ(loaded.workbook_protection_xml(), xml);
  EXPECT_EQ(FromXlsb(ToXlsb(loaded)).workbook_protection_xml(), xml);

  // Legacy verifiers for both passwords.
  const std::string legacy =
      "<workbookProtection workbookPassword=\"CC1A\" revisionsPassword=\"CC4B\" lockStructure=\"1\"/>";
  wb.set_workbook_protection_xml(legacy);
  const Workbook legacy_back = FromXlsb(ToXlsb(ThroughXlsx(wb), &deferred));
  EXPECT_EQ(deferred, 0U);
  EXPECT_EQ(legacy_back.workbook_protection_xml(), legacy);
}

// After an Iso record the 534 verifiers must be zero, so a legacy verifier
// beside a modern hash is left out and reported.
TEST(XlsbWorkbookFidelity, LegacyVerifierBesideModernHashIsReported) {
  Workbook wb = Workbook::create();
  std::uint32_t baseline = 99U;
  ToXlsb(wb, &baseline);
  wb.set_workbook_protection_xml(
      "<workbookProtection workbookPassword=\"CC1A\" workbookAlgorithmName=\"SHA-512\" workbookHashValue=\"Zg==\" "
      "workbookSaltValue=\"Zg==\" workbookSpinCount=\"1\" lockStructure=\"1\"/>");
  std::uint32_t deferred = 0U;
  const std::vector<std::uint8_t> saved = ToXlsb(wb, &deferred);
  EXPECT_EQ(deferred, baseline + 1U);
  const auto book = WorkbookRecords(saved, 534U);
  ASSERT_EQ(book.size(), 1U);
  EXPECT_EQ(book[0], (std::vector<std::uint8_t>{0x00, 0x00, 0x00, 0x00, 0x01, 0x00}));
}

TEST(XlsbWorkbookFidelity, DefinedNameHiddenFlagAndCommentSurviveBothPaths) {
  Workbook wb = Workbook::create();
  std::vector<DefinedName> names(3);
  names[0].name = "Rate";
  names[0].formula = "Sheet1!$A$1";
  names[0].comment = "Interest rate";
  names[1].name = "Secret";
  names[1].formula = "Sheet1!$B$2:$C$3";
  names[1].hidden = true;
  names[2].name = "Local";
  names[2].formula = "42";
  names[2].local_sheet_id = 0;
  names[2].hidden = true;
  names[2].comment = "Sheet-scoped";
  wb.set_defined_names(names);
  const Workbook source = ThroughXlsx(wb);
  ASSERT_EQ(source.defined_names().size(), 3U);
  const Workbook loaded = FromXlsb(ToXlsb(source));
  const Workbook resaved = FromXlsb(ToXlsb(loaded));
  for (const Workbook* back : {&loaded, &resaved}) {
    ASSERT_EQ(back->defined_names().size(), 3U);
    for (std::size_t i = 0; i < 3U; ++i) {
      const DefinedName& a = back->defined_names()[i];
      const DefinedName& e = source.defined_names()[i];
      EXPECT_EQ(a.name, e.name);
      EXPECT_EQ(a.formula, e.formula) << a.name;
      EXPECT_EQ(a.local_sheet_id, e.local_sheet_id) << a.name;
      EXPECT_EQ(a.hidden, e.hidden) << a.name;
      EXPECT_EQ(a.comment, e.comment) << a.name;
    }
  }
}

std::vector<std::uint8_t> SheetPart(const std::vector<std::uint8_t>& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(Span(package))));
  auto part = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(part)) << path;
  return part ? std::move(part.value()) : std::vector<std::uint8_t>();
}

std::size_t RecordCount(const std::vector<std::uint8_t>& part, std::uint16_t type) {
  std::size_t count = 0;
  ByteSpan cursor = Span(part);
  while (cursor.size > 0U) {
    auto record = read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    count += record.value().type == type ? 1U : 0U;
  }
  return count;
}

/// Two workbook windows, the second showing sheet 1 hidden and minimized,
/// and an extra sheet view in that window.
Workbook TwoWindowWorkbook() {
  Workbook wb = Workbook::create();
  wb.add_sheet("Second");
  wb.set_book_views_xml(
      "<bookViews><workbookView xWindow=\"-103\" yWindow=\"-103\" windowWidth=\"27634\" "
      "windowHeight=\"16629\" activeTab=\"1\"/><workbookView visibility=\"hidden\" minimized=\"1\" "
      "showHorizontalScroll=\"0\" showVerticalScroll=\"0\" showSheetTabs=\"0\" xWindow=\"120\" "
      "yWindow=\"240\" windowWidth=\"9000\" windowHeight=\"7000\" tabRatio=\"350\" firstSheet=\"1\" "
      "activeTab=\"1\" autoFilterDateGrouping=\"0\"/></bookViews>");
  wb.set_active_sheet(1U);
  wb.sheet(0).mutable_view().extra_views_xml = "<sheetView workbookViewId=\"1\" zoomScale=\"80\"/>";
  return wb;
}

constexpr std::uint16_t kBrtBeginWsView = 137;

TEST(XlsbWorkbookFidelity, WorkbookWindowsSurviveXlsxToXlsbAndXlsbResave) {
  const Workbook source = ThroughXlsx(TwoWindowWorkbook());
  ASSERT_EQ(source.book_views_xml(), TwoWindowWorkbook().book_views_xml());
  std::uint32_t deferred = 99U;
  const std::vector<std::uint8_t> first = ToXlsb(source, &deferred);
  EXPECT_EQ(deferred, 0U);
  EXPECT_EQ(WorkbookRecords(first, kBrtBookView).size(), 2U);
  EXPECT_EQ(RecordCount(SheetPart(first, "xl/worksheets/sheet1.bin"), kBrtBeginWsView), 2U);
  EXPECT_EQ(RecordCount(SheetPart(first, "xl/worksheets/sheet2.bin"), kBrtBeginWsView), 1U);

  const Workbook loaded = FromXlsb(first);
  EXPECT_EQ(loaded.book_views_xml(), source.book_views_xml());
  EXPECT_EQ(loaded.active_sheet(), 1U);
  const std::string& extra = loaded.sheet(0).view().extra_views_xml;
  EXPECT_NE(extra.find("workbookViewId=\"1\""), std::string::npos) << extra;
  EXPECT_NE(extra.find("zoomScale=\"80\""), std::string::npos) << extra;

  const std::vector<std::uint8_t> second = ToXlsb(loaded, &deferred);
  EXPECT_EQ(deferred, 0U);
  EXPECT_EQ(WorkbookRecords(second, kBrtBookView), WorkbookRecords(first, kBrtBookView));
  const Workbook resaved = FromXlsb(second);
  EXPECT_EQ(resaved.book_views_xml(), source.book_views_xml());
  EXPECT_EQ(resaved.sheet(0).view().extra_views_xml, extra);
}

// A BrtBookView too short to decode loses that window, not the workbook.
TEST(XlsbWorkbookFidelity, TruncatedBookViewIsSkipped) {
  const std::vector<std::uint8_t> package = ToXlsb(TwoWindowWorkbook());
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(Span(package))));
  auto book = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(book));
  std::vector<std::uint8_t> patched;
  for (const FramedRecord& rec : split_records(book.value())) {
    if (rec.type == kBrtBookView) {
      emit_record(patched, kBrtBookView, std::vector<std::uint8_t>{0x01, 0x00});
    } else {
      append_record(patched, book.value(), rec);
    }
  }
  auto loaded = read_xlsb(Span(RepackWithEntries(package, {{"xl/workbook.bin", patched}})));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  EXPECT_EQ(loaded.value().workbook.sheet_count(), 2U);
  EXPECT_TRUE(loaded.value().workbook.book_views_xml().empty());
}

// Custom table / PivotTable styles have no binary form here and are reported.
TEST(XlsbWorkbookFidelity, CustomTableStylesAreReportedDeferred) {
  Workbook wb = Workbook::create();
  std::uint32_t baseline = 99U;
  ToXlsb(wb, &baseline);

  CustomTableStyle style;
  style.name = "Mine";
  wb.mutable_styles().table_styles.present = true;
  wb.mutable_styles().table_styles.styles.push_back(style);
  std::uint32_t deferred = 0U;
  ToXlsb(wb, &deferred);
  EXPECT_EQ(deferred, baseline + 1U);

  // Only default style names that differ from Excel's are a setting to lose.
  Workbook defaults = Workbook::create();
  defaults.mutable_styles().table_styles.present = true;
  defaults.mutable_styles().table_styles.default_table_style = "TableStyleMedium2";
  defaults.mutable_styles().table_styles.default_pivot_style = "PivotStyleLight16";
  ToXlsb(defaults, &deferred);
  EXPECT_EQ(deferred, baseline);
  defaults.mutable_styles().table_styles.default_table_style = "TableStyleLight1";
  ToXlsb(defaults, &deferred);
  EXPECT_EQ(deferred, baseline + 1U);
}

// A sheet view naming a window the workbook does not have is still reported.
TEST(XlsbWorkbookFidelity, SheetViewForAMissingWindowIsReported) {
  Workbook wb = TwoWindowWorkbook();
  std::uint32_t baseline = 99U;
  ToXlsb(wb, &baseline);
  wb.sheet(1).mutable_view().extra_views_xml = "<sheetView workbookViewId=\"2\" zoomScale=\"90\"/>";
  std::uint32_t deferred = 0U;
  ToXlsb(wb, &deferred);
  EXPECT_EQ(deferred, baseline + 1U);
}

// Excel's .xlsb twins pin BrtBookView: the xlsx-sourced save writes Excel's
// bytes, and the .xlsb reads to the xlsx window minus its revision uid.
TEST(XlsbWorkbookFidelity, BookViewMatchesExcelsTwin) {
  for (const char* base : {"excel/win/pivot_week_1900/pivot_week_1900", "excel/xlsb_fidelity_base"}) {
    const std::vector<std::uint8_t> excel_xlsb = ReadFixture(std::string(base) + ".xlsb");
    const Workbook from_xlsx = FromXlsx(ReadFixture(std::string(base) + ".xlsx"));
    EXPECT_EQ(WorkbookRecords(ToXlsb(from_xlsx), kBrtBookView), WorkbookRecords(excel_xlsb, kBrtBookView)) << base;
    std::string expected = from_xlsx.book_views_xml();
    const std::size_t uid = expected.find(" xr2:uid=\"");
    ASSERT_NE(uid, std::string::npos) << base;
    expected.erase(uid, expected.find('"', uid + 10U) + 1U - uid);
    EXPECT_EQ(FromXlsb(excel_xlsb).book_views_xml(), expected) << base;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
