// The custom table styles block of styles.bin the xlsb writer writes from the model, and the dxfs its elements name,
// held against Excel 16.115's own xlsb saves of the same sources (tests/fixtures/excel/xlsb_parity/tablestyles).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "styles.h"
#include "utils/structured_log.h"
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
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/tablestyles/" + name;
}

/// A part's records as `type:hex` lines, which keep a failure readable.
std::vector<std::string> RecordLines(const Bytes& part) {
  static const char kHex[] = "0123456789abcdef";
  std::vector<std::string> out;
  for (const FramedRecord& rec : split_records(part)) {
    std::string line = std::to_string(rec.type) + ":";
    for (std::size_t i = 0; i < rec.payload.size; ++i) {
      line.push_back(kHex[rec.payload.data[i] >> 4U]);
      line.push_back(kHex[rec.payload.data[i] & 0x0FU]);
    }
    out.push_back(std::move(line));
  }
  return out;
}

/// The styles part's dxf list and table styles block (records 505 through 509) as `type:hex` lines.
std::vector<std::string> TableStyleBlock(const Bytes& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto styles = zip.read_entry("xl/styles.bin");
  std::vector<std::string> out;
  if (!styles) {
    ADD_FAILURE() << "no styles part";
    return out;
  }
  bool inside = false;
  for (const std::string& line : RecordLines(styles.value())) {
    inside = inside || line.rfind("505:", 0) == 0 || line.rfind("508:", 0) == 0;
    if (inside) {
      out.push_back(line);
    }
    if (line.rfind("509:", 0) == 0) {
      break;
    }
  }
  return out;
}

/// `lines` without the dxf records and with each element's dxf slot blanked.
std::vector<std::string> WithoutDxfSlots(const std::vector<std::string>& lines) {
  std::vector<std::string> out;
  for (const std::string& line : lines) {
    if (line.rfind("507:", 0) == 0) {
      continue;
    }
    out.push_back(line.rfind("512:", 0) == 0 ? line.substr(0, line.size() - 8U) : line);
  }
  return out;
}

Bytes StylesPart(const Bytes& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto styles = zip.read_entry("xl/styles.bin");
  return styles ? std::move(styles.value()) : Bytes{};
}

/// The deferred-feature log records of the saves made while it lives, which name what a save left out.
class DeferredRecords {
 public:
  DeferredRecords() {
    set_structured_log_sink(&Collect, &records_);
    set_structured_log_min_level(StructuredLogLevel::kWarn);
  }
  ~DeferredRecords() {
    set_structured_log_min_level(StructuredLogLevel::kOff);
    set_structured_log_sink(nullptr);
  }
  DeferredRecords(const DeferredRecords&) = delete;
  DeferredRecords& operator=(const DeferredRecords&) = delete;
  const std::string& text() const noexcept { return records_; }

 private:
  static void Collect(std::string_view record, void* user) {
    if (record.find("xlsb.writer.deferred") != std::string_view::npos) {
      static_cast<std::string*>(user)->append(record);
    }
  }
  std::string records_;
};

const char* const kCases[] = {"custom_style", "medium9_default", "element_types"};

TEST(XlsbParityTableStyleTwins, BlockFromTheSourceMatchesExcel) {
  // Custom styles with their uid wrappers, elements in type order, the dxfs they name in Excel's order, and the
  // default names, from each source's .xlsx.
  for (const char* name : kCases) {
    SCOPED_TRACE(name);
    auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(std::string(name) + "_src.xlsx"))));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const DeferredRecords deferred;
    auto saved = write_xlsb_with_result(loaded.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(TableStyleBlock(saved.value().bytes), TableStyleBlock(ReadFile(Fixture(std::string(name) + ".xlsb"))));
    EXPECT_EQ(deferred.text().find("table_styles"), std::string::npos) << deferred.text();
  }
}

TEST(XlsbParityTableStyleTwins, NewWorkbookDefaultsMatchExcel) {
  auto loaded = read_ooxml(SpanOf(ReadFile(Fixture("new_workbook.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto saved = write_xlsb_with_result(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const auto ours = TableStyleBlock(saved.value().bytes);
  const auto theirs = TableStyleBlock(ReadFile(Fixture("new_workbook.xlsb")));
  ASSERT_FALSE(theirs.empty());
  EXPECT_EQ(ours, theirs);
}

TEST(XlsbParityTableStyleEdit, ExcelSaveReadsAndWritesBack) {
  for (const char* name : kCases) {
    SCOPED_TRACE(name);
    const Bytes excel = ReadFile(Fixture(std::string(name) + ".xlsb"));
    auto loaded = read_xlsb(SpanOf(excel));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    Workbook& wb = loaded.value().workbook;
    EXPECT_TRUE(wb.styles().table_styles.present);
    // Untouched, the styles part passes through byte for byte.
    auto kept = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(kept)) << kept.error().message;
    EXPECT_EQ(StylesPart(kept.value().bytes), StylesPart(excel));
    // Regenerated, it keeps every style Excel wrote, each element in place; which saved index a dxf takes after an
    // xlsb load is not measured, so the dxf slots are compared by content in XlsbParityDxf.
    wb.mutable_styles().fonts.push_back(FontRecord{});
    auto regenerated = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(regenerated)) << regenerated.error().message;
    EXPECT_NE(StylesPart(regenerated.value().bytes), StylesPart(excel));
    EXPECT_EQ(WithoutDxfSlots(TableStyleBlock(regenerated.value().bytes)), WithoutDxfSlots(TableStyleBlock(excel)));
  }
}

TEST(XlsbParityTableStyleEdit, StylesReadFromExcelsSave) {
  auto loaded = read_xlsb(SpanOf(ReadFile(Fixture("element_types.xlsb"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const TableStylesSection& section = loaded.value().workbook.styles().table_styles;
  EXPECT_EQ(section.default_table_style, "TableStyleMedium2");
  EXPECT_EQ(section.default_pivot_style, "PivotStyleLight16");
  ASSERT_EQ(section.styles.size(), 3U);
  EXPECT_EQ(section.styles[0].name, "AllTable");
  EXPECT_TRUE(section.styles[0].table);
  EXPECT_FALSE(section.styles[0].pivot);
  EXPECT_EQ(section.styles[0].elements.size(), 13U);
  EXPECT_EQ(section.styles[1].elements.size(), 28U);
  EXPECT_TRUE(section.styles[1].pivot);
  EXPECT_FALSE(section.styles[1].table);
  EXPECT_EQ(section.styles[1].elements[27].type, "pageFieldValues");
  EXPECT_EQ(section.styles[2].elements[1].type, "headerRow");
}

TEST(XlsbParityTableStyleEdit, AStyleOfAnUnmeasuredFormIsLeftOutAndReported) {
  // Usable by both tables and PivotTables: not a form Excel was measured writing.
  auto loaded = read_ooxml(SpanOf(ReadFile(Fixture("custom_style_src.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  wb.mutable_styles().table_styles.styles[0].pivot = true;
  const DeferredRecords deferred;
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_NE(deferred.text().find("\"kind\":\"table_styles\",\"count\":1"), std::string::npos) << deferred.text();
  for (const std::string& line : TableStyleBlock(saved.value().bytes)) {
    EXPECT_NE(line.rfind("510:", 0), 0U) << line;
  }
}

TEST(XlsbParityDxf, TableStyleElementDxfsSurviveCompaction) {
  // Every dxf a written element names is kept, and the element names the saved index of that dxf.
  auto loaded = read_ooxml(SpanOf(ReadFile(Fixture("element_types_src.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto saved = write_xlsb_with_result(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto back = read_xlsb(SpanOf(saved.value().bytes));
  ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
  const StylesTable& before = loaded.value().workbook.styles();
  const StylesTable& after = back.value().workbook.styles();
  ASSERT_EQ(after.dxfs.size(), before.dxfs.size());
  for (std::size_t s = 0; s < before.table_styles.styles.size(); ++s) {
    for (const TableStyleElement& element : before.table_styles.styles[s].elements) {
      bool found = false;
      for (const TableStyleElement& written : after.table_styles.styles[s].elements) {
        if (written.type == element.type) {
          found = true;
          EXPECT_EQ(after.dxfs[*written.dxf_id].font.color_argb, before.dxfs[*element.dxf_id].font.color_argb)
              << element.type;
        }
      }
      EXPECT_TRUE(found) << element.type;
    }
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
