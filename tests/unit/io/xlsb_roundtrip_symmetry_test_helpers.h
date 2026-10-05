#pragma once

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "styles.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace xlsb_roundtrip_test_support {

inline std::string FixturePath(const char* name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
}

// Reads the shared fixture from both formats. Returns false (with a gtest
// failure) if either read fails; otherwise fills the two out-workbooks.
inline ::testing::AssertionResult LoadBothFormats(Workbook* xlsb_out, Workbook* xlsx_out) {
  const std::vector<std::uint8_t> xlsb_bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  const std::vector<std::uint8_t> xlsx_bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsx"));
  if (xlsb_bytes.empty() || xlsx_bytes.empty()) {
    return ::testing::AssertionFailure() << "fixture bytes empty";
  }
  auto xb = io::xlsb::read_xlsb(test::span_of(xlsb_bytes));
  if (!xb) {
    return ::testing::AssertionFailure() << "read_xlsb failed: " << xb.error().message;
  }
  auto xx = io::read_ooxml(test::span_of(xlsx_bytes));
  if (!xx) {
    return ::testing::AssertionFailure() << "read_ooxml failed: " << xx.error().message;
  }
  *xlsb_out = std::move(xb.value().workbook);
  *xlsx_out = std::move(xx.value().workbook);
  return ::testing::AssertionSuccess();
}

inline void ExpectColorSpecEqual(const ColorSpec& xlsb, const ColorSpec& xlsx) {
  ASSERT_EQ(static_cast<int>(xlsb.kind), static_cast<int>(xlsx.kind));
  switch (xlsb.kind) {
    case ColorSpec::Kind::kRgb:
      EXPECT_EQ(xlsb.rgb, xlsx.rgb);
      break;
    case ColorSpec::Kind::kTheme:
      EXPECT_EQ(xlsb.theme, xlsx.theme);
      EXPECT_NEAR(xlsb.tint, xlsx.tint, 1e-9);
      break;
    case ColorSpec::Kind::kIndexed:
      EXPECT_EQ(xlsb.indexed, xlsx.indexed);
      break;
    case ColorSpec::Kind::kNone:
    case ColorSpec::Kind::kAuto:
      break;
  }
}

inline std::string CellXfsBlockOfSavedPackage(const std::vector<std::uint8_t>& package) {
  std::string styles;
  if (!test::extract_part(test::span_of(package), "xl/styles.xml", &styles)) {
    return std::string();
  }
  const std::size_t begin = styles.find("<cellXfs");
  const std::size_t end = styles.find("</cellXfs>");
  if (begin == std::string::npos || end == std::string::npos || end < begin) {
    return std::string();
  }
  return styles.substr(begin, end - begin + std::strlen("</cellXfs>"));
}

/// Counts non-overlapping occurrences of `needle` in `haystack`.
inline std::size_t CountOccurrences(const std::string& haystack, const std::string& needle) {
  std::size_t count = 0;
  for (std::size_t pos = haystack.find(needle); pos != std::string::npos; pos = haystack.find(needle, pos + 1)) {
    ++count;
  }
  return count;
}

// The in-memory table is only half the claim: an `.xlsx` save serialises the
// model, never the retained `xl/styles.bin` bytes, so what a host or another
// spreadsheet application sees after a conversion is whatever reached
// `xl/styles.xml`. Asserting on the saved part -- and against the same part
// produced from the workbook's own `.xlsx` export, which is the identical
// writer on an identical model -- is what states that a `.xlsb` source loses
// nothing on the way out.
inline std::uint32_t ResolvedNumFmtId(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Cell* c = wb.sheet(0).cell_at(row, col);
  if (c == nullptr) {
    return 0xFFFFFFFFU;
  }
  const StylesTable& st = wb.styles();
  if (c->xf_index >= st.cell_xfs.size()) {
    return 0xFFFFFFFFU;  // dangling index -> style table did not round-trip
  }
  return st.cell_xfs[c->xf_index].num_fmt_id;
}
struct WireName {
  std::string name;
  std::int32_t itab = -1;
};

// Decodes `xl/workbook.bin`'s `BrtName` records in emission order, so
// entry `i` is what a `PtgName` with `ilbl == i + 1` reaches.
inline std::vector<WireName> ReadWireNames(const std::vector<std::uint8_t>& workbook_bin) {
  std::vector<WireName> out;
  io::ByteSpan cursor = test::span_of(workbook_bin);
  while (cursor.size > 0U) {
    auto rec_or = io::xlsb::read_record(cursor);
    if (!rec_or) {
      return out;
    }
    if (rec_or.value().type != static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtName)) {
      continue;
    }
    // BrtName: grbit (u32) + chKey (u8) + itab (i32) + the name string.
    io::ByteSpan p = rec_or.value().payload;
    auto grbit = io::xlsb::read_u32(p);
    auto ch_key = io::xlsb::read_u8(p);
    auto itab = io::xlsb::read_u32(p);
    if (!grbit || !ch_key || !itab) {
      return out;
    }
    auto name = io::xlsb::read_xlwidestring(p);
    if (!name) {
      return out;
    }
    out.push_back(WireName{name.value(), static_cast<std::int32_t>(itab.value())});
  }
  return out;
}

// Returns the `ilbl` encoded by the lone `PtgName` token of the formula
// stored at row 0 / `col` of `sheet_bin`, or 0 when the cell is absent
// or its token stream is not a single bare name reference.
inline std::uint32_t IlblOfNameOnlyFormula(const std::vector<std::uint8_t>& sheet_bin, std::uint32_t col) {
  io::ByteSpan cursor = test::span_of(sheet_bin);
  while (cursor.size > 0U) {
    auto rec_or = io::xlsb::read_record(cursor);
    if (!rec_or) {
      return 0U;
    }
    // A name formula is entered as a dynamic-array formula, whose tokens a
    // `BrtArrFmla` carries: rwFirst, rwLast, colFirst, colLast (u32 each),
    // a flag byte, then cce + rgce.
    if (rec_or.value().type == static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtArrFmla)) {
      io::ByteSpan p = rec_or.value().payload;
      if (p.size < 26U || p.data[8] != col || p.data[17] != 5U || p.data[21] != 0x43U) {
        continue;
      }
      p.data += 22U;
      p.size -= 22U;
      auto ilbl = io::xlsb::read_u32(p);
      return ilbl ? ilbl.value() : 0U;
    }
    if (rec_or.value().type != static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtFmlaNum)) {
      continue;
    }
    // BrtFmlaNum: cell header (col u32 + iStyleRef 3B + fPhShow u8),
    // the cached double, grbitFlags (u16), then the CellParsedFormula
    // (cce + rgce + cb + rgcb).
    io::ByteSpan p = rec_or.value().payload;
    auto cell_col = io::xlsb::read_u32(p);
    if (!cell_col || cell_col.value() != col) {
      continue;
    }
    if (p.size < 14U) {
      return 0U;
    }
    p.data += 4U + 8U + 2U;  // iStyleRef + fPhShow, cached value, grbitFlags
    p.size -= 4U + 8U + 2U;
    auto cce = io::xlsb::read_u32(p);
    if (cce && p.size > 0U && p.data[0] == 0x01U) {
      continue;  // PtgExp: the tokens follow in the cell's BrtArrFmla.
    }
    // `=Foo` lowers to exactly one value-class PtgName: opcode 0x43 + a u32 ilbl.
    if (!cce || cce.value() != 5U || p.size < 5U || p.data[0] != 0x43U) {
      return 0U;
    }
    p.data += 1U;
    p.size -= 1U;
    auto ilbl = io::xlsb::read_u32(p);
    return ilbl ? ilbl.value() : 0U;
  }
  return 0U;
}

// Builds the workbook both scope-resolution cases share: a
// workbook-scoped `Foo` (Sheet1!A1 = 10) and a Sheet1-local `Foo`
// (Sheet1!B1 = 20), with `=Foo` on both sheets. `local_first` flips the
// declaration order of the two names, which is the tie-break a
// text-keyed name table falls back on.
inline void RunScopeResolutionCase(bool local_first) {
  SCOPED_TRACE(local_first ? "sheet-local Foo declared first" : "workbook-scoped Foo declared first");
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  if (local_first) {
    ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$B$1", 0)));
    ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$A$1", -1)));
  } else {
    ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$A$1", -1)));
    ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Foo", "Sheet1!$B$1", 0)));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));  // Sheet1!A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::number(20.0))));  // Sheet1!B1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=Foo")));             // Sheet1!D1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 3U, "=Foo")));             // Sheet2!D1
  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << recalc_or.error().message;
  // The engine's own resolution is the reference the encoding has to
  // agree with: Sheet1 sees the local `Foo` (B1), Sheet2 the global one.
  ASSERT_EQ(wb.sheet(0).resolve_cell_value(0U, 3U).as_number(), 20.0);
  ASSERT_EQ(wb.sheet(1).resolve_cell_value(0U, 3U).as_number(), 10.0);

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << "write_xlsb failed: " << saved.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(test::span_of(saved.value()))));
  auto workbook_bin = zip.read_entry("xl/workbook.bin");
  auto sheet1_bin = zip.read_entry("xl/worksheets/sheet1.bin");
  auto sheet2_bin = zip.read_entry("xl/worksheets/sheet2.bin");
  ASSERT_TRUE(static_cast<bool>(workbook_bin));
  ASSERT_TRUE(static_cast<bool>(sheet1_bin));
  ASSERT_TRUE(static_cast<bool>(sheet2_bin));

  const std::vector<WireName> wire_names = ReadWireNames(workbook_bin.value());
  ASSERT_EQ(wire_names.size(), 2U);

  const std::uint32_t sheet1_ilbl = IlblOfNameOnlyFormula(sheet1_bin.value(), 3U);
  const std::uint32_t sheet2_ilbl = IlblOfNameOnlyFormula(sheet2_bin.value(), 3U);
  ASSERT_GE(sheet1_ilbl, 1U);
  ASSERT_LE(sheet1_ilbl, wire_names.size());
  ASSERT_GE(sheet2_ilbl, 1U);
  ASSERT_LE(sheet2_ilbl, wire_names.size());

  // Sheet1's reference must land on the record scoped to Sheet1.
  EXPECT_EQ(wire_names[sheet1_ilbl - 1U].name, "Foo");
  EXPECT_EQ(wire_names[sheet1_ilbl - 1U].itab, 0) << "Sheet1 must resolve the sheet-local Foo";
  // Sheet2 has no local override, so it must land on the workbook one.
  EXPECT_EQ(wire_names[sheet2_ilbl - 1U].name, "Foo");
  EXPECT_EQ(wire_names[sheet2_ilbl - 1U].itab, -1) << "Sheet2 must resolve the workbook-scoped Foo";

  // The values the round trip produces are unchanged either way (the
  // reader re-resolves by name text), so they are a guard against the
  // scope fix disturbing them, not the detector for it.
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << "read_xlsb failed: " << reloaded.error().message;
  Workbook after = std::move(reloaded.value().workbook);
  auto after_recalc = after.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(after_recalc)) << after_recalc.error().message;
  EXPECT_EQ(after.sheet(0).resolve_cell_value(0U, 3U).as_number(), 20.0);
  EXPECT_EQ(after.sheet(1).resolve_cell_value(0U, 3U).as_number(), 10.0);
}
inline void ExpectSelfBookFixture(Workbook& wb, const char* label) {
  const Sheet& s1 = wb.sheet(0);
  const Sheet& s2 = wb.sheet(1);
  EXPECT_EQ(s1.cell_at(0U, 2U)->formula_text, "=[0]!Fn(3)") << label;
  EXPECT_EQ(s1.cell_at(0U, 3U)->formula_text, "=SUM([0]!Rng)") << label;
  EXPECT_EQ(s2.cell_at(0U, 0U)->formula_text, "=[0]!G") << label;
  EXPECT_EQ(s2.cell_at(0U, 2U)->formula_text, "=ROWS([0]!Rng)") << label;
  EXPECT_EQ(s2.cell_at(0U, 3U)->formula_text, "=Sheet1!A10:[0]!Rng") << label;
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry()))) << label;
  const double sheet1[] = {100.0, 100.0, 6.0, 7.0, 100.0};
  for (std::uint32_t col = 0; col < 5U; ++col) {
    const Value v = wb.sheet(0).resolve_cell_value(0U, col);
    ASSERT_TRUE(v.is_number()) << label << " Sheet1 col=" << col;
    EXPECT_EQ(v.as_number(), sheet1[col]) << label << " Sheet1 col=" << col;
  }
  const double sheet2[] = {5.0, 5.0, 2.0, 3.0};
  for (std::uint32_t col = 0; col < 4U; ++col) {
    const Value v = wb.sheet(1).resolve_cell_value(0U, col);
    ASSERT_TRUE(v.is_number()) << label << " Sheet2 col=" << col;
    EXPECT_EQ(v.as_number(), sheet2[col]) << label << " Sheet2 col=" << col;
  }
  EXPECT_EQ(wb.sheet(1).resolve_cell_value(1U, 3U).as_number(), 4.0) << label << " Sheet2!D2 spill";
}
inline std::uint64_t BitsOf(double v) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof(v));
  return bits;
}

// The whole point of the predicate: a `true` answer is a promise that
// `BrtCellRk` is lossless for that value. Sweep a deterministic spread
// of finite doubles -- including the currency-shaped band where the
inline ::testing::AssertionResult ThroughXlsb(const Workbook& wb, Workbook* out) {
  auto saved = io::xlsb::write_xlsb(wb);
  if (!saved) {
    return ::testing::AssertionFailure() << "write_xlsb failed: " << saved.error().message;
  }
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  if (!reloaded) {
    return ::testing::AssertionFailure() << "read_xlsb failed: " << reloaded.error().message;
  }
  *out = std::move(reloaded.value().workbook);
  return ::testing::AssertionSuccess();
}

/// Saves `wb` as `.xlsx` and reads it back, or fails the test.
inline ::testing::AssertionResult ThroughXlsx(const Workbook& wb, Workbook* out) {
  auto saved = io::write_ooxml(wb);
  if (!saved) {
    return ::testing::AssertionFailure() << "write_ooxml failed: " << saved.error().message;
  }
  auto reloaded = io::read_ooxml(test::span_of(saved.value()));
  if (!reloaded) {
    return ::testing::AssertionFailure() << "read_ooxml failed: " << reloaded.error().message;
  }
  *out = std::move(reloaded.value().workbook);
  return ::testing::AssertionSuccess();
}

}  // namespace xlsb_roundtrip_test_support
}  // namespace formulon
