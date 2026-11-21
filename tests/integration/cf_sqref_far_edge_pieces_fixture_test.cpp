//
// The pieces a far-line cell delete leaves of a legacy conditional format. `far_edge_a` holds blocks of two or
// three rules mixing kinds that split with kinds that stay whole (the block splits when any rule does);
// `far_edge_d` holds multi-area ranges with one rule. `*_ctl` is Excel's save of the authored book and `*_sit`
// Excel's save after the same delete on every sheet. Pieces compare as a set per block: Excel writes a range split
// around a cut line as top, bottom, then the cut line where the engine writes row order, and the engine's load
// swaps the first two areas of an Excel-saved multi-area data bar or colour scale that Excel keeps; neither
// changes which cells a rule covers.

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
#include "io/zip_reader.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_sqref_far_edge_pieces/" + name;
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

// The `sqref` pieces of every legacy `<conditionalFormatting>` in the `index`-th sheet part, sorted per block.
std::vector<std::vector<std::string>> LegacySqrefs(const std::vector<std::uint8_t>& package, std::size_t index) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{package.data(), package.size()})));
  const std::string part = "xl/worksheets/sheet" + std::to_string(index + 1U) + ".xml";
  auto entry = zip.read_entry(part);
  EXPECT_TRUE(static_cast<bool>(entry)) << part;
  const std::string body = entry ? std::string(entry.value().begin(), entry.value().end()) : std::string();
  std::vector<std::vector<std::string>> out;
  constexpr std::string_view kOpen = "<conditionalFormatting sqref=\"";
  for (std::size_t at = body.find(kOpen); at != std::string::npos; at = body.find(kOpen, at)) {
    at += kOpen.size();
    std::istringstream text(body.substr(at, body.find('"', at) - at));
    std::vector<std::string> pieces{std::istream_iterator<std::string>(text), std::istream_iterator<std::string>()};
    std::sort(pieces.begin(), pieces.end());
    out.push_back(std::move(pieces));
  }
  return out;
}

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

struct SheetEdit {
  std::string sheet;
  const char* first;  // top-left of the deleted cells
  const char* last;
  CellShift shift;
};

// Applies `edits` to Excel's save of the authored `book` and compares each edited sheet's saved sqref text with
// Excel's save after the same edits.
void CheckBook(const std::string& book, const std::vector<SheetEdit>& edits) {
  const std::vector<std::uint8_t> authored = ReadFixture(book + "_ctl.xlsx");
  auto loaded = io::read_ooxml(io::ByteSpan{authored.data(), authored.size()});
  ASSERT_TRUE(static_cast<bool>(loaded));
  Workbook wb = std::move(loaded.value().workbook);
  std::vector<std::size_t> indices;
  for (const SheetEdit& edit : edits) {
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
    indices.push_back(index);
  }
  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  const std::vector<std::uint8_t> excel = ReadFixture(book + "_sit.xlsx");
  for (std::size_t k = 0; k < edits.size(); ++k) {
    EXPECT_EQ(LegacySqrefs(saved.value(), indices[k]), LegacySqrefs(excel, indices[k])) << edits[k].sheet;
  }
}

// The same delete on sheets `prefix`NN for NN in [`from`, `to`].
void AddRun(std::vector<SheetEdit>& edits, char prefix, int from, int to, const char* first, const char* last,
            CellShift shift) {
  for (int n = from; n <= to; ++n) {
    std::string name(1, prefix);
    name += static_cast<char>('0' + n / 10);
    name += static_cast<char>('0' + n % 10);
    edits.push_back(SheetEdit{name, first, last, shift});
  }
}

TEST(CfSqrefFarEdgePiecesFixture, BlocksMixingSplittingAndWholeKinds) {
  std::vector<SheetEdit> edits;
  AddRun(edits, 'A', 1, 8, "E4", "E4", CellShift::kLeft);
  AddRun(edits, 'A', 9, 16, "D6", "E6", CellShift::kUp);
  AddRun(edits, 'A', 17, 22, "E11", "E11", CellShift::kLeft);
  AddRun(edits, 'A', 23, 26, "E4", "E4", CellShift::kLeft);
  CheckBook("far_edge_a", edits);
}

TEST(CfSqrefFarEdgePiecesFixture, MultiAreaRanges) {
  std::vector<SheetEdit> edits;
  const char* const kLeftAtF5[] = {"D01", "D03", "D05", "D09", "D10", "D11", "D13", "D16", "D17"};
  const char* const kUpAtE6[] = {"D02", "D04", "D06", "D07", "D08", "D12", "D14", "D15", "D18"};
  for (const char* sheet : kLeftAtF5) {
    edits.push_back(SheetEdit{sheet, "F5", "F5", CellShift::kLeft});
  }
  for (const char* sheet : kUpAtE6) {
    edits.push_back(SheetEdit{sheet, "E6", "E6", CellShift::kUp});
  }
  CheckBook("far_edge_d", edits);
}

}  // namespace
}  // namespace formulon
