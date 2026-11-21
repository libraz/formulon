//
// A cell delete on the far line of an x14-only conditional format (a stock icon set Excel stores only in the
// extension) keeps the range whole unless the rule ranks the range or has a min, max or percent threshold, as for
// a legacy rule. `x14_far_edge_ctl` is Excel's save of the authored book, one rule per sheet; `x14_far_edge_sit` is
// Excel's save after the same delete on every sheet, in both formats. Pieces compare as a set: the order Excel
// writes a split range's bands in is not the engine's, with no visible effect.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <iterator>
#include <sstream>
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
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

constexpr std::uint16_t kBeginConditionalFormatting14 = 1046U;

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_x14_far_edge/" + name;
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
    out.clear();
  }
  std::fclose(file);
  return out;
}

Workbook Load(const std::vector<std::uint8_t>& bytes, bool xlsb) {
  const io::ByteSpan span{bytes.data(), bytes.size()};
  if (xlsb) {
    auto read = io::xlsb::read_xlsb(span);
    EXPECT_TRUE(static_cast<bool>(read));
    return read ? std::move(read.value().workbook) : Workbook::create_empty();
  }
  auto read = io::read_ooxml(span);
  EXPECT_TRUE(static_cast<bool>(read));
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Save(const Workbook& wb, bool xlsb) {
  if (xlsb) {
    auto saved = io::xlsb::write_xlsb(wb);
    EXPECT_TRUE(static_cast<bool>(saved));
    return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
  }
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved));
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

std::string ReadEntry(const std::vector<std::uint8_t>& package, const std::string& name) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{package.data(), package.size()})));
  auto entry = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(entry)) << name;
  return entry ? std::string(entry.value().begin(), entry.value().end()) : std::string();
}

std::string CellText(std::uint32_t row, std::uint32_t col) {
  std::string letters;
  for (std::uint32_t n = col + 1U; n > 0U; n = (n - 1U) / 26U) {
    letters.insert(letters.begin(), static_cast<char>('A' + (n - 1U) % 26U));
  }
  return letters + std::to_string(row + 1U);
}

std::uint32_t LoadU32(const std::uint8_t* at) {
  return static_cast<std::uint32_t>(at[0]) | (static_cast<std::uint32_t>(at[1]) << 8U) |
         (static_cast<std::uint32_t>(at[2]) << 16U) | (static_cast<std::uint32_t>(at[3]) << 24U);
}

// The pieces of one x14 block's sqref, A1 text sorted, from a BrtBeginConditionalFormatting14 payload: three
// four-byte fields, the piece count, then rwFirst, rwLast, colFirst, colLast per piece.
std::vector<std::string> XlsbPieces(const io::xlsb::FramedRecord& rec) {
  constexpr std::size_t kCountAt = 12U;
  constexpr std::size_t kRfxBytes = 16U;
  std::vector<std::string> out;
  if (rec.payload.size < kCountAt + 4U) {
    return out;
  }
  const std::uint32_t count = LoadU32(rec.payload.data + kCountAt);
  for (std::uint32_t k = 0; k < count && kCountAt + 4U + (k + 1U) * kRfxBytes <= rec.payload.size; ++k) {
    const std::uint8_t* rfx = rec.payload.data + kCountAt + 4U + k * kRfxBytes;
    const std::string first = CellText(LoadU32(rfx), LoadU32(rfx + 8U));
    const std::string last = CellText(LoadU32(rfx + 4U), LoadU32(rfx + 12U));
    out.push_back(first == last ? first : first + ":" + last);
  }
  std::sort(out.begin(), out.end());
  return out;
}

// The x14 conditional-format blocks of the `index`-th sheet part, each as its sorted sqref pieces.
std::vector<std::vector<std::string>> X14Ranges(const std::vector<std::uint8_t>& package, std::size_t index,
                                                bool xlsb) {
  const std::string part =
      "xl/worksheets/sheet" + std::to_string(index + 1U) + (xlsb ? std::string(".bin") : std::string(".xml"));
  const std::string body = ReadEntry(package, part);
  std::vector<std::vector<std::string>> out;
  if (xlsb) {
    const std::vector<std::uint8_t> bytes(body.begin(), body.end());
    for (const io::xlsb::FramedRecord& rec : io::xlsb::split_records(bytes)) {
      if (rec.type == kBeginConditionalFormatting14) {
        out.push_back(XlsbPieces(rec));
      }
    }
    return out;
  }
  constexpr std::string_view kOpen = "<xm:sqref>";
  for (std::size_t at = body.find(kOpen); at != std::string::npos; at = body.find(kOpen, at)) {
    at += kOpen.size();
    std::istringstream text(body.substr(at, body.find("</xm:sqref>", at) - at));
    std::vector<std::string> pieces{std::istream_iterator<std::string>(text), std::istream_iterator<std::string>()};
    std::sort(pieces.begin(), pieces.end());
    out.push_back(std::move(pieces));
  }
  return out;
}

struct SheetEdit {
  const char* sheet;
  const char* first;  // top-left of the deleted cells
  const char* last;
  CellShift shift;
};

// One delete per sheet: a middle cell of the right line, two cells of the bottom line, and a corner of a
// two-row range. B23 (a data bar) and B24 (an icon set with formula thresholds) are the controls.
const SheetEdit kEdits[] = {
    {"B01", "E4", "E4", CellShift::kLeft},   {"B02", "E4", "E4", CellShift::kLeft},
    {"B03", "E4", "E4", CellShift::kLeft},   {"B04", "E4", "E4", CellShift::kLeft},
    {"B05", "E4", "E4", CellShift::kLeft},   {"B06", "E4", "E4", CellShift::kLeft},
    {"B07", "E4", "E4", CellShift::kLeft},   {"B08", "E4", "E4", CellShift::kLeft},
    {"B09", "E4", "E4", CellShift::kLeft},   {"B10", "E4", "E4", CellShift::kLeft},
    {"B11", "E4", "E4", CellShift::kLeft},   {"B12", "E4", "E4", CellShift::kLeft},
    {"B13", "E4", "E4", CellShift::kLeft},   {"B14", "E4", "E4", CellShift::kLeft},
    {"B15", "E4", "E4", CellShift::kLeft},   {"B16", "D6", "E6", CellShift::kUp},
    {"B17", "D6", "E6", CellShift::kUp},     {"B18", "D6", "E6", CellShift::kUp},
    {"B19", "D6", "E6", CellShift::kUp},     {"B20", "D6", "E6", CellShift::kUp},
    {"B21", "E11", "E11", CellShift::kLeft}, {"B22", "E11", "E11", CellShift::kLeft},
    {"B23", "E4", "E4", CellShift::kLeft},   {"B24", "E4", "E4", CellShift::kLeft},
};

std::pair<std::uint32_t, std::uint32_t> Cell(std::string_view a1) {
  std::uint32_t col = 0;
  std::size_t i = 0;
  for (; i < a1.size() && a1[i] >= 'A' && a1[i] <= 'Z'; ++i) {
    col = col * 26U + static_cast<std::uint32_t>(a1[i] - 'A' + 1);
  }
  std::uint32_t row = 0;
  for (; i < a1.size(); ++i) {
    row = row * 10U + static_cast<std::uint32_t>(a1[i] - '0');
  }
  return {row - 1U, col - 1U};
}

void CheckFormat(bool xlsb) {
  Workbook wb = Load(ReadFixture(xlsb ? "x14_far_edge_ctl.xlsb" : "x14_far_edge_ctl.xlsx"), xlsb);
  for (const SheetEdit& edit : kEdits) {
    std::size_t index = wb.sheet_count();
    for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
      if (wb.sheet(i).name() == edit.sheet) {
        index = i;
      }
    }
    ASSERT_LT(index, wb.sheet_count()) << edit.sheet;
    const auto first = Cell(edit.first);
    const auto last = Cell(edit.last);
    ASSERT_TRUE(static_cast<bool>(
        wb.delete_cells(index, MergeRange{first.first, first.second, last.first, last.second}, edit.shift)))
        << edit.sheet;
  }
  const std::vector<std::uint8_t> engine = Save(wb, xlsb);
  const std::vector<std::uint8_t> excel = ReadFixture(xlsb ? "x14_far_edge_sit.xlsb" : "x14_far_edge_sit.xlsx");
  for (std::size_t i = 0; i < std::size(kEdits); ++i) {
    const std::vector<std::vector<std::string>> expected = X14Ranges(excel, i, xlsb);
    EXPECT_EQ(expected.size(), 1U) << kEdits[i].sheet;
    EXPECT_EQ(X14Ranges(engine, i, xlsb), expected) << kEdits[i].sheet;
  }
}

TEST(CfX14FarEdgeFixture, XlsxRangeSplitsOnlyForPercentThresholds) {
  CheckFormat(false);
}

TEST(CfX14FarEdgeFixture, XlsbRangeSplitsOnlyForPercentThresholds) {
  CheckFormat(true);
}

}  // namespace
}  // namespace formulon
