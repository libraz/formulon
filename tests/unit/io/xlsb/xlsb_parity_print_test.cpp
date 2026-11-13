// The worksheet print records the xlsb writer derives from the model (print options, margins, page setup, header and
// footer, manual breaks), held against Excel 16.115's own xlsb saves of the same sources
// (tests/fixtures/excel/xlsb_parity/print), and a fresh book of every profile saving with nothing deferred.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
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
