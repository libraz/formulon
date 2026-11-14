// An .xlsb loaded and saved with no style edit keeps its styles.bin byte for byte, whatever the workbook profile, and
// the parts written beside it follow the stored part: cell format indices as loaded, no sheet phonetic setting naming
// a font the stored part lacks, table cell-style names the stored part holds, and dxf indices as stored. The part is
// regenerated once the style tables or the cell formats in use change since load; formats in use count cells, style
// runs, row and column styles and merged-range references (`collect_used_cell_xfs`), so clearing the last cell, row or
// column using a format also regenerates it (the cs1 freshness goldens pin the Excel-measured cell case).

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/xlsb/emission_plan.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "styles.h"
#include "table.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr std::uint16_t kBrtPhoneticInfo = 537;
constexpr std::uint16_t kBrtBeginDxfs = 505;
constexpr std::uint16_t kBrtDxf = 507;
constexpr std::uint16_t kBrtBeginTableStyles = 508;

Bytes ReadFile(const std::string& path) {
  std::ifstream in(path, std::ios::binary);
  return Bytes(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
}

ByteSpan SpanOf(const Bytes& bytes) {
  return ByteSpan{bytes.data(), bytes.size()};
}

Bytes Entry(const Bytes& package, const char* name) {
  ZipReader zip;
  if (!zip.open(SpanOf(package)) || !zip.has_entry(name)) {
    return {};
  }
  auto bytes = zip.read_entry(name);
  return bytes ? std::move(bytes.value()) : Bytes{};
}

// Excel-saved .xlsb files, relative to tests/fixtures.
std::vector<std::string> ExcelXlsbFixtures() {
  std::vector<std::string> out;
  std::ifstream in(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_flag_corpus_manifest.txt");
  std::string line;
  while (std::getline(in, line)) {
    std::istringstream words(line.substr(0, line.find('#')));
    std::string path;
    if (words >> path) {
      out.push_back(path);
    }
  }
  return out;
}

struct Record {
  std::uint16_t id = 0;
  Bytes bytes;  // header and payload
};

// Every record of an xlsb part, in order.
std::vector<Record> Records(const Bytes& part) {
  std::vector<Record> out;
  std::size_t i = 0;
  while (i < part.size()) {
    const std::size_t begin = i;
    std::uint32_t id = 0;
    for (int k = 0; k < 2 && i < part.size(); ++k) {
      const std::uint8_t b = part[i++];
      id |= static_cast<std::uint32_t>(b & 0x7FU) << (7 * k);
      if ((b & 0x80U) == 0U) {
        break;
      }
    }
    std::uint32_t len = 0;
    for (int k = 0; k < 4 && i < part.size(); ++k) {
      const std::uint8_t b = part[i++];
      len |= static_cast<std::uint32_t>(b & 0x7FU) << (7 * k);
      if ((b & 0x80U) == 0U) {
        break;
      }
    }
    i = std::min(part.size(), i + len);
    out.push_back(Record{static_cast<std::uint16_t>(id), Bytes(part.begin() + static_cast<std::ptrdiff_t>(begin),
                                                               part.begin() + static_cast<std::ptrdiff_t>(i))});
  }
  return out;
}

std::vector<std::uint16_t> RecordIds(const Bytes& part) {
  std::vector<std::uint16_t> ids;
  for (const Record& r : Records(part)) {
    ids.push_back(r.id);
  }
  return ids;
}

TEST(XlsbStylesPassthrough, EveryExcelXlsbKeepsItsStylesPartOnAnUntouchedSave) {
  const std::vector<std::string> fixtures = ExcelXlsbFixtures();
  ASSERT_GT(fixtures.size(), 100U);
  std::vector<std::string> failed_to_load;
  std::vector<std::string> stale_on_load;
  std::vector<std::string> changed;
  for (const std::string& rel : fixtures) {
    const Bytes original = ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/" + rel);
    const Bytes stored = Entry(original, "xl/styles.bin");
    auto loaded = read_xlsb(SpanOf(original));
    if (!loaded || stored.empty()) {
      failed_to_load.push_back(rel);
      continue;
    }
    if (xlsb_styles_regenerated(loaded.value().workbook)) {
      stale_on_load.push_back(rel);
      continue;
    }
    auto saved = write_xlsb(loaded.value().workbook);
    if (!saved || Entry(saved.value(), "xl/styles.bin") != stored) {
      changed.push_back(rel);
    }
  }
  const auto joined = [](const std::vector<std::string>& v) {
    std::string out;
    for (const std::string& s : v) {
      out += "\n  " + s;
    }
    return out;
  };
  EXPECT_TRUE(changed.empty()) << "styles.bin changed on an untouched save:" << joined(changed);
  // Files the corpus check cannot cover, each listed with its reason.
  EXPECT_EQ(failed_to_load, std::vector<std::string>{}) << "failed to load:" << joined(failed_to_load);
  EXPECT_EQ(stale_on_load, std::vector<std::string>{}) << "stale on load:" << joined(stale_on_load);
}

// A ja-written Excel book with duplicate cell formats, loaded under the default en-US profile: the save passes
// styles.bin through, so every cell keeps the format index it was loaded with instead of a compacted one.
TEST(XlsbStylesPassthrough, CellFormatIndicesStayAsLoadedBesideAStoredPart) {
  const Bytes original =
      ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/cellxfs_compaction/mg2_B14_merge_items_b.xlsb");
  auto loaded = read_xlsb(SpanOf(original));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;
  ASSERT_FALSE(xlsb_styles_regenerated(wb));
  auto saved = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_EQ(Entry(saved.value(), "xl/styles.bin"), Entry(original, "xl/styles.bin"));
  auto back = read_xlsb(SpanOf(saved.value()));
  ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
  std::size_t compared = 0;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const auto& [row, cells] : wb.sheet(s).rows()) {
      for (const auto& [col, cell] : cells) {
        const Cell* again = back.value().workbook.sheet(s).cell_at(row, col);
        ASSERT_NE(again, nullptr) << "sheet " << s << " row " << row << " col " << col;
        EXPECT_EQ(again->xf_index, cell.xf_index) << "sheet " << s << " row " << row << " col " << col;
        ++compared;
      }
    }
  }
  EXPECT_GT(compared, 0U);
}

// A book saved under en-US and reloaded under ja-JP: a regenerated part would gain the ja phonetic font and every
// sheet would name it; beside the stored part no sheet names a phonetic font.
TEST(XlsbStylesPassthrough, NoSheetNamesAPhoneticFontTheStoredPartLacks) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  auto first = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(first)) << first.error().message;
  const std::vector<std::uint16_t> before = RecordIds(Entry(first.value(), "xl/worksheets/sheet1.bin"));
  ASSERT_EQ(std::count(before.begin(), before.end(), kBrtPhoneticInfo), 0);
  auto loaded = read_xlsb(SpanOf(first.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& book = loaded.value().workbook;
  book.set_excel_profile(mac_365_ja_jp_profile());
  auto second = write_xlsb(book);
  ASSERT_TRUE(static_cast<bool>(second)) << second.error().message;
  EXPECT_EQ(Entry(second.value(), "xl/styles.bin"), Entry(first.value(), "xl/styles.bin"));
  const std::vector<std::uint16_t> ids = RecordIds(Entry(second.value(), "xl/worksheets/sheet1.bin"));
  EXPECT_EQ(std::count(ids.begin(), ids.end(), kBrtPhoneticInfo), 0);
}

// A CF-only edit on an Excel xlsb renumbers the stored part's dxf list in place: the style tables and the formats in
// use are unchanged, so the part is not regenerated, and every record but the dxf list, its table and slicer style
// blocks included, stays as stored.
TEST(XlsbStylesPassthrough, ACfOnlyEditRenumbersTheStoredDxfListInPlace) {
  const Bytes original = ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_feature_cellis_ops.xlsb");
  auto loaded = read_xlsb(SpanOf(original));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  ASSERT_FALSE(wb.sheet(0).conditional_formats().empty());
  wb.sheet(0).mutable_conditional_formats().pop_back();
  EXPECT_FALSE(xlsb_styles_regenerated(wb));
  EXPECT_TRUE(xlsb_dxf_references_changed(wb));
  auto saved = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;

  std::vector<Bytes> kept_before;
  std::size_t dxfs_before = 0;
  bool table_styles = false;
  for (const Record& r : Records(Entry(original, "xl/styles.bin"))) {
    table_styles = table_styles || r.id == kBrtBeginTableStyles;
    if (r.id == kBrtDxf) {
      ++dxfs_before;
    } else if (r.id != kBrtBeginDxfs) {
      kept_before.push_back(r.bytes);
    }
  }
  ASSERT_TRUE(table_styles);
  std::vector<Bytes> kept_after;
  std::size_t dxfs_after = 0;
  for (const Record& r : Records(Entry(saved.value(), "xl/styles.bin"))) {
    if (r.id == kBrtDxf) {
      ++dxfs_after;
    } else if (r.id != kBrtBeginDxfs) {
      kept_after.push_back(r.bytes);
    }
  }
  EXPECT_EQ(kept_after, kept_before);
  EXPECT_LT(dxfs_after, dxfs_before);
}

// A table naming the English built-in spelling "Normal", saved under en-US and reloaded under ja-JP: a regenerated
// part would name the style 標準 and drop the attribute; beside the stored part, which holds "Normal", it is kept.
TEST(XlsbStylesPassthrough, TableCellStyleNamesFollowTheStoredPart) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 0U, 0U, "Region")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 1U, 0U, "North")));
  TableMetadata table;
  table.id = 1;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:A2";
  table.sheet_index = 0;
  table.header_row = true;
  table.root_extra_attrs = " dataCellStyle=\"Normal\"";
  table.columns = {TableColumn(1, "Region", "", "", "")};
  wb.mutable_tables().push_back(table);
  auto first = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(first)) << first.error().message;
  auto loaded = read_xlsb(SpanOf(first.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ASSERT_EQ(loaded.value().workbook.tables().size(), 1U);
  ASSERT_NE(loaded.value().workbook.tables()[0].root_extra_attrs.find("dataCellStyle=\"Normal\""), std::string::npos);

  Workbook& book = loaded.value().workbook;
  book.set_excel_profile(mac_365_ja_jp_profile());
  auto second = write_xlsb(book);
  ASSERT_TRUE(static_cast<bool>(second)) << second.error().message;
  EXPECT_EQ(Entry(second.value(), "xl/styles.bin"), Entry(first.value(), "xl/styles.bin"));
  auto back = read_xlsb(SpanOf(second.value()));
  ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
  ASSERT_EQ(back.value().workbook.tables().size(), 1U);
  EXPECT_NE(back.value().workbook.tables()[0].root_extra_attrs.find("dataCellStyle=\"Normal\""), std::string::npos)
      << back.value().workbook.tables()[0].root_extra_attrs;
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
