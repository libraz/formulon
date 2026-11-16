// The worksheet print records the xlsb writer derives from the model (print options, margins, page setup, header and
// footer, manual breaks), held against Excel 16.115's own xlsb saves of the same sources
// (tests/fixtures/excel/xlsb_parity/print), and a fresh book of every profile saving with nothing deferred.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/print_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;

Bytes ReadFile(const std::string& path) {
  Bytes out;
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
  }
  std::fclose(f);
  return out;
}

std::string Fixture(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/print/" + name;
}

bool IsPrintRecord(std::uint16_t type) {
  switch (type) {
    case kBrtPrintOptions:
    case kBrtMargins:
    case kBrtPageSetup:
    case kBrtBeginHeaderFooter:
    case kBrtEndHeaderFooter:
    case kBrtBeginRwBrk:
    case kBrtEndRwBrk:
    case kBrtBeginColBrk:
    case kBrtEndColBrk:
    case kBrtBrk:
      return true;
    default:
      return false;
  }
}

/// The type and payload of each print record in the first worksheet, in stream order.
std::vector<std::pair<std::uint16_t, Bytes>> PrintRecords(const Bytes& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(part));
  std::vector<std::pair<std::uint16_t, Bytes>> out;
  if (!part) {
    return out;
  }
  for (const FramedRecord& rec : split_records(part.value())) {
    if (IsPrintRecord(rec.type)) {
      out.emplace_back(rec.type, Bytes(rec.payload.data, rec.payload.data + rec.payload.size));
    }
  }
  return out;
}

TEST(XlsbParityPrint, PrintOptionsAndMarginsMatchExcel) {
  for (const char* name : {"base",
                           "defaults_only",
                           "margins_fine",
                           "margins_zero",
                           "opt_vcenter",
                           "opt_gridlinesset0",
                           "paper_custom",
                           "ps_bw",
                           "ps_comments_atend",
                           "ps_comments_displayed",
                           "ps_copies",
                           "ps_dpi",
                           "ps_draft",
                           "ps_errors_blank",
                           "ps_errors_dash",
                           "ps_errors_na",
                           "ps_firstpage_unused",
                           "ps_orient_portrait",
                           "ps_pageorder",
                           "ps_paper_b5",
                           "ps_paper_letter",
                           "ps_printerdefaults0",
                           "scale_edges",
                           "hf_codes",
                           "hf_empty_header",
                           "hf_odd_even_first",
                           "hf_scale_align0",
                           "header_255",
                           "breaks_many",
                           "breaks_row_1024",
                           "breaks_col_1024",
                           "brk_auto_pt",
                           "brk_minmax",
                           "wsprop_bits",
                           "outline_apply_styles",
                           "outline_symbols0"}) {
    SCOPED_TRACE(name);
    auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(std::string(name) + "_src.xlsx"))));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    auto saved = write_xlsb_with_result(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(PrintRecords(saved.value().bytes), PrintRecords(ReadFile(Fixture(std::string(name) + ".xlsb"))));
  }
}

/// The first worksheet's BrtWsProp payload.
Bytes WsProp(const Bytes& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(part));
  if (part) {
    for (const FramedRecord& rec : split_records(part.value())) {
      if (rec.type == static_cast<std::uint16_t>(XlsbRecordType::BrtWsProp)) {
        return Bytes(rec.payload.data, rec.payload.data + rec.payload.size);
      }
    }
  }
  ADD_FAILURE() << "no BrtWsProp";
  return Bytes{};
}

TEST(XlsbParityPrint, SheetPropertiesMatchExcelAndDeferNoPrintSettings) {
  for (const char* name : {"defaults_only", "wsprop_bits", "outline_apply_styles", "outline_symbols0"}) {
    SCOPED_TRACE(name);
    auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(std::string(name) + "_src.xlsx"))));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    auto saved = write_xlsb_with_result(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(WsProp(saved.value().bytes), WsProp(ReadFile(Fixture(std::string(name) + ".xlsb"))));
    EXPECT_FALSE(has_deferred_print_settings(loaded.value().workbook.sheet(0).print_settings()));
  }
}

TEST(XlsbParityPrint, UnwrittenSheetPropertiesStayDeferred) {
  // Sheets two to four set autoPageBreaks, summaryBelow and summaryRight off, none of which the record carries yet.
  auto loaded = read_ooxml(SpanOf(ReadFile(Fixture("wsprop_bits_src.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 4U);
  for (std::size_t i = 1; i < wb.sheet_count(); ++i) {
    EXPECT_TRUE(has_deferred_print_settings(wb.sheet(i).print_settings())) << i;
  }
}

TEST(XlsbParityPrint, ExcelXlsbLoadsTheTwinsPageSetupAndMarginsAndResavesExcelsRecords) {
  for (const char* name : {"base",
                           "defaults_only",
                           "margins_fine",
                           "margins_zero",
                           "paper_custom",
                           "ps_bw",
                           "ps_comments_atend",
                           "ps_comments_displayed",
                           "ps_copies",
                           "ps_dpi",
                           "ps_draft",
                           "ps_errors_blank",
                           "ps_errors_dash",
                           "ps_errors_na",
                           "ps_firstpage_unused",
                           "ps_orient_portrait",
                           "ps_pageorder",
                           "ps_paper_b5",
                           "ps_paper_letter",
                           "ps_printerdefaults0",
                           "scale_edges",
                           "hf_codes",
                           "breaks_many",
                           "wsprop_bits"}) {
    SCOPED_TRACE(name);
    const Bytes excel = ReadFile(Fixture(std::string(name) + ".xlsb"));
    auto twin = read_ooxml(SpanOf(ReadFile(Fixture(std::string(name) + ".xlsx"))));
    ASSERT_TRUE(static_cast<bool>(twin)) << twin.error().message;
    auto loaded = read_xlsb(SpanOf(excel));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const SheetPrintSettings& want = twin.value().workbook.sheet(0).print_settings();
    const SheetPrintSettings& got = loaded.value().workbook.sheet(0).print_settings();
    EXPECT_EQ(got.page_setup_xml, want.page_setup_xml);
    EXPECT_EQ(got.page_setup.orientation, want.page_setup.orientation);
    EXPECT_EQ(got.page_setup.paper_size, want.page_setup.paper_size);
    EXPECT_EQ(got.page_setup.scale, want.page_setup.scale);
    EXPECT_EQ(got.page_setup.fit_to_width, want.page_setup.fit_to_width);
    EXPECT_EQ(got.page_setup.fit_to_height, want.page_setup.fit_to_height);
    EXPECT_EQ(got.page_setup.fit_to_page, want.page_setup.fit_to_page);
    EXPECT_EQ(got.page_setup.page_order, want.page_setup.page_order);
    EXPECT_EQ(got.page_margins.left, want.page_margins.left);
    EXPECT_EQ(got.page_margins.right, want.page_margins.right);
    EXPECT_EQ(got.page_margins.top, want.page_margins.top);
    EXPECT_EQ(got.page_margins.bottom, want.page_margins.bottom);
    EXPECT_EQ(got.page_margins.header, want.page_margins.header);
    EXPECT_EQ(got.page_margins.footer, want.page_margins.footer);

    auto saved = write_xlsb_with_result(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(PrintRecords(saved.value().bytes), PrintRecords(excel));
  }
}

/// The first worksheet's XML part of an `.xlsx` package.
std::string SheetXml(const Bytes& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto part = zip.read_entry("xl/worksheets/sheet1.xml");
  EXPECT_TRUE(static_cast<bool>(part));
  return part ? std::string(part.value().begin(), part.value().end()) : std::string();
}

/// The `<element ...>...</element>` of `xml`, empty when it has none.
std::string ElementOf(const std::string& xml, const std::string& element) {
  const std::size_t begin = xml.find("<" + element);
  const std::string close = "</" + element + ">";
  const std::size_t end = xml.find(close, begin);
  return begin == std::string::npos || end == std::string::npos ? std::string()
                                                                : xml.substr(begin, end + close.size() - begin);
}

void ExpectSameBreaks(const std::vector<ManualBreak>& got, const std::vector<ManualBreak>& want) {
  ASSERT_EQ(got.size(), want.size());
  for (std::size_t i = 0; i < got.size(); ++i) {
    EXPECT_EQ(got[i].id, want[i].id) << i;
    EXPECT_EQ(got[i].min, want[i].min) << i;
    EXPECT_EQ(got[i].max, want[i].max) << i;
    EXPECT_EQ(got[i].manual, want[i].manual) << i;
  }
}

TEST(XlsbParityPrint, ExcelXlsbLoadsTheTwinsHeaderFooterAndBreaksAndKeepsThemOnBothSaves) {
  for (const char* name : {"hf_codes", "hf_empty_header", "hf_odd_even_first", "hf_scale_align0", "brk_auto_pt",
                           "brk_minmax", "breaks_many", "breaks_row_1024", "breaks_col_1024"}) {
    SCOPED_TRACE(name);
    const Bytes excel = ReadFile(Fixture(std::string(name) + ".xlsb"));
    const Bytes twin_package = ReadFile(Fixture(std::string(name) + ".xlsx"));
    auto twin = read_ooxml(SpanOf(twin_package));
    ASSERT_TRUE(static_cast<bool>(twin)) << twin.error().message;
    auto loaded = read_xlsb(SpanOf(excel));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const SheetPrintSettings& want = twin.value().workbook.sheet(0).print_settings();
    const SheetPrintSettings& got = loaded.value().workbook.sheet(0).print_settings();
    EXPECT_EQ(got.header_footer_xml, want.header_footer_xml);
    ExpectSameBreaks(got.manual_row_breaks, want.manual_row_breaks);
    ExpectSameBreaks(got.manual_col_breaks, want.manual_col_breaks);

    auto xlsx = write_ooxml(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
    const std::string wrote = SheetXml(xlsx.value());
    const std::string twin_xml = SheetXml(twin_package);
    for (const char* element : {"headerFooter", "rowBreaks", "colBreaks"}) {
      EXPECT_EQ(ElementOf(wrote, element), ElementOf(twin_xml, element)) << element;
    }

    auto saved = write_xlsb_with_result(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(PrintRecords(saved.value().bytes), PrintRecords(excel));
  }
}

/// A header/footer family with the default flag word and an odd header only.
Bytes HeaderFooterFamily() {
  Bytes out;
  Bytes head;
  emit_u16(head, 0x0CU);
  emit_xlnullablewidestring(head, std::string_view("&CPage"));
  for (int i = 0; i < 5; ++i) {
    emit_xlnullablewidestring(head, std::nullopt);
  }
  emit_record(out, kBrtBeginHeaderFooter, head);
  emit_record(out, kBrtEndHeaderFooter, Bytes());
  return out;
}

/// A break family on one axis holding a single break before index 10.
Bytes BreakFamily(bool row_axis, bool manual) {
  Bytes out;
  Bytes head;
  emit_u32(head, 1U);
  emit_u32(head, manual ? 1U : 0U);
  emit_record(out, row_axis ? kBrtBeginRwBrk : kBrtBeginColBrk, head);
  Bytes brk;
  emit_u32(brk, 10U);
  emit_u32(brk, 0U);
  emit_u32(brk, row_axis ? 16383U : 1048575U);
  emit_u32(brk, manual ? 1U : 0U);
  emit_u32(brk, 0U);
  emit_record(out, kBrtBrk, brk);
  emit_record(out, row_axis ? kBrtEndRwBrk : kBrtEndColBrk, Bytes());
  return out;
}

Bytes PrintOptionsRecord(std::uint16_t flags) {
  Bytes out;
  Bytes payload;
  emit_u16(payload, flags);
  emit_record(out, kBrtPrintOptions, payload);
  return out;
}

SheetPrintSettings Decoded(const Bytes& stream) {
  SheetPrintSettings print;
  decode_print_records(stream, print);
  return print;
}

TEST(XlsbPrintFamilies, HeaderFooterTheModelStatesIsWrittenAndOneItLacksIsCounted) {
  const Bytes stream = HeaderFooterFamily();
  SheetPrintSettings print = Decoded(stream);
  ASSERT_FALSE(print.header_footer_xml.empty());
  EXPECT_EQ(unwritten_print_families(stream, print), 0U);
  // The same stream against a model without it.
  print.header_footer_xml.clear();
  EXPECT_EQ(unwritten_print_families(stream, print), 1U);
}

TEST(XlsbPrintFamilies, ManualRowBreakIsReadAndAutomaticOneStaysRetainedAndCounted) {
  const Bytes manual = BreakFamily(true, true);
  const SheetPrintSettings read = Decoded(manual);
  ASSERT_EQ(read.manual_row_breaks.size(), 1U);
  EXPECT_EQ(read.manual_row_breaks[0].id, 10U);
  EXPECT_TRUE(read.manual_row_breaks[0].manual);
  EXPECT_EQ(unwritten_print_families(manual, read), 0U);

  const Bytes automatic = BreakFamily(true, false);
  const SheetPrintSettings kept = Decoded(automatic);
  EXPECT_TRUE(kept.manual_row_breaks.empty());
  EXPECT_EQ(unwritten_print_families(automatic, kept), 1U);
}

TEST(XlsbPrintFamilies, RowAndColumnBreaksCountAsSeparateFamilies) {
  Bytes stream = BreakFamily(true, false);
  const Bytes cols = BreakFamily(false, false);
  stream.insert(stream.end(), cols.begin(), cols.end());
  EXPECT_EQ(unwritten_print_families(stream, Decoded(stream)), 2U);
}

TEST(XlsbPrintFamilies, DefaultPrintOptionsStateNothingAndAreNotCounted) {
  const Bytes stream = PrintOptionsRecord(0x10U);
  const SheetPrintSettings print = Decoded(stream);
  EXPECT_TRUE(print.print_options_xml.empty());
  EXPECT_EQ(unwritten_print_families(stream, print), 0U);
}

TEST(XlsbPrintFamilies, GridLinesAndHeadingsPrintOptionsAreReadAsExcelStatesThem) {
  const Bytes stream = PrintOptionsRecord(0x1CU);
  const SheetPrintSettings print = Decoded(stream);
  EXPECT_EQ(print.print_options_xml, "<printOptions headings=\"1\" gridLines=\"1\"/>");
  EXPECT_EQ(unwritten_print_families(stream, print), 0U);
}

TEST(XlsbPrintFamilies, PrintOptionsWithAnUnmeasuredBitStayRetainedAndCounted) {
  const Bytes stream = PrintOptionsRecord(0x30U);
  const SheetPrintSettings print = Decoded(stream);
  EXPECT_TRUE(print.print_options_xml.empty());
  EXPECT_EQ(unwritten_print_families(stream, print), 1U);
}

/// The deferred-feature count of an `.xlsx` save of the Excel xlsb `hf_codes` with `extra` appended to its sheet tail.
std::uint32_t XlsxDeferredCountWithTailExtra(const Bytes& extra) {
  auto loaded = read_xlsb(SpanOf(ReadFile(Fixture("hf_codes.xlsb"))));
  EXPECT_TRUE(static_cast<bool>(loaded));
  if (!loaded) {
    return 0U;
  }
  Workbook& wb = loaded.value().workbook;
  XlsbSheetTail tail = wb.sheet(0).xlsb_tail();
  tail.after_hyperlinks.insert(tail.after_hyperlinks.end(), extra.begin(), extra.end());
  wb.sheet(0).set_xlsb_tail(std::move(tail));
  auto saved = write_ooxml_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(saved));
  return saved ? saved.value().diagnostics.deferred_feature_count : 0U;
}

TEST(XlsbPrintFamilies, XlsxSaveCountsEachPrintFamilyItCannotWrite) {
  const std::uint32_t base = XlsxDeferredCountWithTailExtra(Bytes());
  // Print options at their defaults state nothing.
  EXPECT_EQ(XlsxDeferredCountWithTailExtra(PrintOptionsRecord(0x10U)), base);
  EXPECT_EQ(XlsxDeferredCountWithTailExtra(BreakFamily(true, false)), base + 1U);
  Bytes both = BreakFamily(true, false);
  const Bytes cols = BreakFamily(false, false);
  both.insert(both.end(), cols.begin(), cols.end());
  EXPECT_EQ(XlsxDeferredCountWithTailExtra(both), base + 2U);
  EXPECT_EQ(XlsxDeferredCountWithTailExtra(PrintOptionsRecord(0x30U)), base + 1U);
}

TEST(XlsbPrintFamilies, XlsxSaveOfExcelsXlsbCountsNoPrintFamily) {
  for (const char* name : {"hf_codes", "breaks_many", "base", "brk_minmax"}) {
    SCOPED_TRACE(name);
    auto loaded = read_xlsb(SpanOf(ReadFile(Fixture(std::string(name) + ".xlsb"))));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const Sheet& sheet = loaded.value().workbook.sheet(0);
    EXPECT_EQ(unwritten_print_families(sheet.xlsb_tail().after_hyperlinks, sheet.print_settings()), 0U);
  }
}

TEST(XlsbParityPrint, FreshBookOfEveryProfileDefersNothing) {
  for (const ::formulon::detail::ExcelProfileIdEntry& entry : ::formulon::detail::kExcelProfileIds) {
    Workbook wb = Workbook::create_for_profile(entry.profile);
    wb.add_sheet("Second");
    auto saved = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << entry.id;
    EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U) << entry.id;
  }
  auto plain = write_xlsb_with_result(Workbook::create());
  ASSERT_TRUE(static_cast<bool>(plain));
  EXPECT_EQ(plain.value().diagnostics.deferred_feature_count, 0U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
