#include "support/xlsb_package_validator.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/zip_reader.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::test {
namespace {

using io::xlsb::ReadFixtureFile;
using io::xlsb::RepackWithEntries;
using io::xlsb::store_u32;

const std::string kExcel = std::string(FORMULON_FIXTURES_DIR) + "/excel/";
const std::string kChildrenBase = kExcel + "retained_children_edit/children_0_base.xlsb";
const std::string kPivot = kExcel + "xlsb_parity/pivot/item_b_e.xlsb";
const std::string kWorkbook = "xl/workbook.bin";
const std::string kSheet1 = "xl/worksheets/sheet1.bin";

struct Rec {
  std::uint16_t type;
  std::vector<std::uint8_t> payload;
};

std::vector<std::uint8_t> Entry(const std::vector<std::uint8_t>& package, const std::string& name) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{package.data(), package.size()})));
  auto part = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(part)) << name;
  return part ? part.value() : std::vector<std::uint8_t>{};
}

/// Re-zips `package` with `part` rewritten by `edit` over its decoded records.
template <typename EditFn>
std::vector<std::uint8_t> Twin(const std::vector<std::uint8_t>& package, const std::string& part, EditFn edit) {
  const std::vector<std::uint8_t> body = Entry(package, part);
  std::vector<Rec> recs;
  for (const io::xlsb::FramedRecord& r : io::xlsb::split_records(body)) {
    recs.push_back(Rec{r.type, std::vector<std::uint8_t>(r.payload.data, r.payload.data + r.payload.size)});
  }
  EXPECT_FALSE(recs.empty()) << part;
  edit(recs);
  std::vector<std::uint8_t> out;
  for (const Rec& r : recs) {
    io::xlsb::emit_record(out, r.type, r.payload);
  }
  return RepackWithEntries(package, {{part, out}});
}

std::vector<Rec>::iterator FindRec(std::vector<Rec>& recs, std::uint16_t type) {
  const auto it = std::find_if(recs.begin(), recs.end(), [type](const Rec& r) { return r.type == type; });
  EXPECT_NE(it, recs.end()) << "no record " << type;
  return it;
}

std::vector<std::string> Validate(const std::vector<std::uint8_t>& package) {
  return ValidateXlsbPackage(io::ByteSpan{package.data(), package.size()});
}

/// Runs one rule over `package`.
template <typename RuleFn>
std::vector<std::string> RunRule(const std::vector<std::uint8_t>& package, RuleFn rule) {
  std::vector<std::string> out;
  XlsbParts parts;
  if (UnzipXlsbParts(io::ByteSpan{package.data(), package.size()}, parts, out)) {
    rule(parts, out);
  }
  return out;
}

bool Contains(const std::vector<std::string>& findings, const std::string& entry) {
  return std::find(findings.begin(), findings.end(), entry) != findings.end();
}

std::string Joined(const std::vector<std::string>& findings) {
  std::string out;
  for (const std::string& f : findings) {
    out += f + "\n";
  }
  return out;
}

/// Every committed Excel-saved .xlsb under the two swept fixture trees.
std::vector<std::string> SweptFixtures() {
  std::vector<std::string> out;
  for (const char* dir : {"xlsb_parity", "retained_children_edit"}) {
    for (const auto& entry : std::filesystem::recursive_directory_iterator(kExcel + dir)) {
      if (entry.is_regular_file() && entry.path().extension() == ".xlsb") {
        out.push_back(entry.path().string());
      }
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

std::vector<std::string> PivotFixtures() {
  std::vector<std::string> out;
  for (const auto& entry : std::filesystem::directory_iterator(kExcel + "xlsb_parity/pivot")) {
    if (entry.path().extension() == ".xlsb") {
      out.push_back(entry.path().string());
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

TEST(XlsbPackageValidator, SweepIsCleanExceptRecordMaxWidth) {
  const std::vector<std::string> files = SweptFixtures();
  ASSERT_GT(files.size(), 100U);
  for (const std::string& path : files) {
    const std::vector<std::string> findings = Validate(ReadFixtureFile(path));
    if (std::filesystem::path(path).filename() == "width_record_max.xlsb") {
      ASSERT_EQ(findings.size(), 1U) << path << "\n" << Joined(findings);
      EXPECT_EQ(findings[0].rfind("ws_fmt_info_range: ", 0), 0U) << findings[0];
      continue;
    }
    EXPECT_TRUE(findings.empty()) << path << "\n" << Joined(findings);
  }
}

TEST(XlsbPackageValidator, UnreadablePackage) {
  const std::vector<std::uint8_t> junk = {'n', 'o', 't', ' ', 'a', ' ', 'z', 'i', 'p'};
  const std::vector<std::string> findings = Validate(junk);
  ASSERT_EQ(findings.size(), 1U);
  EXPECT_EQ(findings[0].rfind("package_unreadable: ", 0), 0U) << findings[0];
}

TEST(XlsbPackageValidator, UserBookViewPairing) {
  const std::vector<std::uint8_t> base = ReadFixtureFile(kChildrenBase);
  EXPECT_TRUE(RunRule(base, CheckUserBookViewPairing).empty());

  const auto no_book_view = Twin(base, kWorkbook, [](std::vector<Rec>& recs) { recs.erase(FindRec(recs, 397)); });
  const std::vector<std::string> sheet_only = RunRule(no_book_view, CheckUserBookViewPairing);
  ASSERT_EQ(sheet_only.size(), 1U) << Joined(sheet_only);
  EXPECT_NE(sheet_only[0].find("in a sheet has no workbook view"), std::string::npos) << sheet_only[0];

  const auto no_sheet_view = Twin(base, kSheet1, [](std::vector<Rec>& recs) {
    const auto first = FindRec(recs, 422);
    const auto last = FindRec(recs, 425);
    recs.erase(first, last + 1);
  });
  const std::vector<std::string> book_only = RunRule(no_sheet_view, CheckUserBookViewPairing);
  ASSERT_EQ(book_only.size(), 1U) << Joined(book_only);
  EXPECT_NE(book_only[0].find("in workbook.bin has no sheet view"), std::string::npos) << book_only[0];

  const auto flipped = Twin(base, kWorkbook, [](std::vector<Rec>& recs) { FindRec(recs, 397)->payload[24] ^= 0x01U; });
  EXPECT_EQ(RunRule(flipped, CheckUserBookViewPairing).size(), 2U);
}

TEST(XlsbPackageValidator, PivotCacheIdResolves) {
  const std::vector<std::string> pivots = PivotFixtures();
  ASSERT_FALSE(pivots.empty());
  for (const std::string& path : pivots) {
    EXPECT_TRUE(RunRule(ReadFixtureFile(path), CheckPivotCacheIdResolves).empty()) << path;
  }
  const std::vector<std::uint8_t> base = ReadFixtureFile(kPivot);
  const std::vector<std::uint8_t> base_wb = Entry(base, kWorkbook);
  std::uint32_t id_sx = 0;
  for (const io::xlsb::FramedRecord& r : io::xlsb::split_records(base_wb)) {
    if (r.type == 386) {
      id_sx = io::xlsb::load_u32(r.payload.data);
    }
  }
  const std::string other = std::to_string(id_sx + 1U);

  const auto dangling = Twin(base, kWorkbook, [id_sx](std::vector<Rec>& recs) {
    std::vector<std::uint8_t> payload(4);
    store_u32(payload.data(), id_sx + 1U);
    io::xlsb::emit_xlwidestring(payload, "rId99");
    const auto end = FindRec(recs, 385);
    recs.insert(end, {Rec{386, payload}, Rec{387, {}}});
  });
  EXPECT_EQ(RunRule(dangling, CheckPivotCacheIdResolves),
            std::vector<std::string>{"pivot_cache_id_resolves: idSx " + other +
                                     " r:id rId99 names no pivotCacheDefinition relationship"});

  const auto renumbered = Twin(
      base, kWorkbook, [id_sx](std::vector<Rec>& recs) { store_u32(FindRec(recs, 386)->payload.data(), id_sx + 1U); });
  EXPECT_EQ(RunRule(renumbered, CheckPivotCacheIdResolves),
            std::vector<std::string>{"pivot_cache_id_resolves: xl/pivotTables/pivotTable1.bin cacheId " +
                                     std::to_string(id_sx) + " matches no idSx"});

  const auto unlisted = Twin(base, kWorkbook, [](std::vector<Rec>& recs) {
    recs.erase(FindRec(recs, 386));
    recs.erase(FindRec(recs, 387));
  });
  const std::vector<std::string> unlisted_findings = RunRule(unlisted, CheckPivotCacheIdResolves);
  EXPECT_TRUE(
      Contains(unlisted_findings,
               "pivot_cache_id_resolves: xl/pivotCache/pivotCacheDefinition1.bin is named by no BrtPivotCacheID"))
      << Joined(unlisted_findings);

  const auto repeated = Twin(base, kWorkbook, [](std::vector<Rec>& recs) {
    const Rec copy = *FindRec(recs, 386);
    recs.insert(FindRec(recs, 385), {copy, Rec{387, {}}});
  });
  EXPECT_EQ(RunRule(repeated, CheckPivotCacheIdResolves),
            std::vector<std::string>{"pivot_cache_id_resolves: idSx " + std::to_string(id_sx) + " repeats"});
}

TEST(XlsbPackageValidator, ActiveSheetIsId) {
  const std::vector<std::uint8_t> base = ReadFixtureFile(kChildrenBase);
  EXPECT_TRUE(RunRule(base, CheckActiveSheetIsId).empty());
  EXPECT_TRUE(RunRule(ReadFixtureFile(kPivot), CheckActiveSheetIsId).empty());

  const auto unknown_id =
      Twin(base, kWorkbook, [](std::vector<Rec>& recs) { store_u32(FindRec(recs, 397)->payload.data() + 16, 3U); });
  EXPECT_EQ(RunRule(unknown_id, CheckActiveSheetIsId),
            std::vector<std::string>{"active_sheet_is_id: 397 names sheetId 3, book has 1"});

  const auto past_end =
      Twin(base, kWorkbook, [](std::vector<Rec>& recs) { store_u32(FindRec(recs, 158)->payload.data() + 24, 1U); });
  EXPECT_EQ(RunRule(past_end, CheckActiveSheetIsId),
            std::vector<std::string>{"active_sheet_is_id: 158 itabCur 1 with 1 sheets"});
}

TEST(XlsbPackageValidator, OutlineLevel) {
  const std::vector<std::uint8_t> base = ReadFixtureFile(kExcel + "xlsb_parity/format/outline_row_7.xlsb");
  EXPECT_TRUE(RunRule(base, CheckOutlineLevel).empty());

  const auto row_8 = Twin(base, kSheet1, [](std::vector<Rec>& recs) { FindRec(recs, 485)->payload[10] = 8U; });
  EXPECT_EQ(RunRule(row_8, CheckOutlineLevel), std::vector<std::string>{"outline_level: " + kSheet1 + " row 8 col 0"});

  const auto col_8 = Twin(base, kSheet1, [](std::vector<Rec>& recs) { FindRec(recs, 485)->payload[11] = 8U; });
  EXPECT_EQ(RunRule(col_8, CheckOutlineLevel), std::vector<std::string>{"outline_level: " + kSheet1 + " row 7 col 8"});
}

TEST(XlsbPackageValidator, WsFmtInfoRange) {
  for (const char* name : {"width_255", "base_255", "height_409_5"}) {
    EXPECT_TRUE(RunRule(ReadFixtureFile(kExcel + "xlsb_parity/format/" + name + ".xlsb"), CheckWsFmtInfoRange).empty())
        << name;
  }
  EXPECT_EQ(RunRule(ReadFixtureFile(kExcel + "xlsb_parity/format/width_record_max.xlsb"), CheckWsFmtInfoRange),
            std::vector<std::string>{"ws_fmt_info_range: " + kSheet1 + " dxGCol 65536"});

  const std::vector<std::uint8_t> base = ReadFixtureFile(kExcel + "xlsb_parity/format/width_255.xlsb");
  const auto over = Twin(
      base, kSheet1, [](std::vector<Rec>& recs) { store_u32(FindRec(recs, 485)->payload.data(), 255U * 256U + 1U); });
  EXPECT_EQ(RunRule(over, CheckWsFmtInfoRange),
            std::vector<std::string>{"ws_fmt_info_range: " + kSheet1 + " dxGCol 65281"});
}

TEST(XlsbPackageValidator, PageBreakCount) {
  for (const char* axis : {"row", "col"}) {
    const std::vector<std::uint8_t> base =
        ReadFixtureFile(kExcel + "xlsb_parity/print/breaks_" + std::string(axis) + "_1024.xlsb");
    EXPECT_TRUE(RunRule(base, CheckPageBreakCount).empty()) << axis;

    const auto extra = Twin(base, kSheet1, [](std::vector<Rec>& recs) {
      const auto brk = FindRec(recs, 396);
      const Rec copy = *brk;
      recs.insert(brk, copy);
    });
    EXPECT_EQ(RunRule(extra, CheckPageBreakCount),
              std::vector<std::string>{"page_break_count: " + kSheet1 +
                                       (std::string(axis) == "row" ? " row" : " column") + " block has 1025 breaks"});
  }
}

TEST(XlsbPackageValidator, SheetXfInRange) {
  const std::vector<std::uint8_t> base = ReadFixtureFile(kChildrenBase);
  EXPECT_TRUE(RunRule(base, CheckSheetXfInRange).empty());

  const auto cell = Twin(base, kSheet1, [](std::vector<Rec>& recs) {
    Rec& rk = *FindRec(recs, 2);
    rk.payload[4] = rk.payload[5] = rk.payload[6] = 0xFFU;
  });
  const std::vector<std::string> cell_findings = RunRule(cell, CheckSheetXfInRange);
  ASSERT_EQ(cell_findings.size(), 1U) << Joined(cell_findings);
  EXPECT_NE(cell_findings[0].find(" xf 16777215 of "), std::string::npos) << cell_findings[0];

  const auto row = Twin(base, kSheet1, [](std::vector<Rec>& recs) {
    Rec& hdr = *FindRec(recs, 0);
    hdr.payload[11] |= io::xlsb::kRowGhostDirty;
    store_u32(hdr.payload.data() + 4, 0xFFFFU);
  });
  const std::vector<std::string> row_findings = RunRule(row, CheckSheetXfInRange);
  ASSERT_EQ(row_findings.size(), 1U) << Joined(row_findings);
  EXPECT_NE(row_findings[0].find(" xf 65535 of "), std::string::npos) << row_findings[0];
}

TEST(XlsbPackageValidator, WorkbookRecordOrder) {
  const std::vector<std::uint8_t> base = ReadFixtureFile(kChildrenBase);
  EXPECT_TRUE(RunRule(base, CheckWorkbookRecordOrder).empty());
  EXPECT_TRUE(RunRule(ReadFixtureFile(kPivot), CheckWorkbookRecordOrder).empty());

  const auto swapped =
      Twin(base, kWorkbook, [](std::vector<Rec>& recs) { std::iter_swap(FindRec(recs, 157), FindRec(recs, 397)); });
  EXPECT_EQ(RunRule(swapped, CheckWorkbookRecordOrder),
            std::vector<std::string>{"workbook_record_order: 157 after 397"});
}

}  // namespace
}  // namespace formulon::test
