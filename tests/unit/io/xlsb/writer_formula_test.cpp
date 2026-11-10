#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"
namespace formulon {
namespace io {
namespace xlsb {
namespace {
// Ptg class byte at the formula's own root position, as in real
// Excel 365 `.xlsb` output (spill anchors and a workbook-qualified name): a
// bare reference
// that is the *entire* formula body -- nothing else consumes it -- is
// written in value class, not reference class. A reference-class token
// cannot stand alone as a formula's result; without this a cell whose
// whole formula is a bare range reference reads back `#VALUE!` in Excel.
// A reference used as a function argument, or as an operand the `:` fast
// path collapses into an Area/Area3d from within an otherwise-nested
// position, is unaffected -- see the "stays reference class" tests below.
namespace {

// Returns the `rgce` bytes of the first `BrtArrFmla` record whose `colFirst`
// equals `want_col`, or an empty vector if none matches.
std::vector<std::uint8_t> FindArrFmlaRgce(const std::vector<std::uint8_t>& sheet_bytes, std::uint32_t want_col) {
  ByteSpan cursor = SpanOf(sheet_bytes);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla)) {
      continue;
    }
    ByteSpan p = record_or.value().payload;
    auto rw_first_or = read_u32(p);
    auto rw_last_or = read_u32(p);
    auto col_first_or = read_u32(p);
    auto col_last_or = read_u32(p);
    (void)rw_first_or;
    (void)rw_last_or;
    (void)col_last_or;
    if (!col_first_or || col_first_or.value() != want_col) {
      continue;
    }
    if (p.size < 1) {
      return {};
    }
    p.data += 1;  // flag byte
    p.size -= 1;
    auto cce_or = read_u32(p);
    if (!cce_or || cce_or.value() > p.size) {
      return {};
    }
    return std::vector<std::uint8_t>(p.data, p.data + cce_or.value());
  }
  return {};
}

// Returns the `rgce` bytes of the `BrtFmlaNum` record at `want_col` (a
// plain, non-array formula cell), or an empty vector if none matches.
std::vector<std::uint8_t> FindFmlaNumRgce(const std::vector<std::uint8_t>& sheet_bytes, std::uint32_t want_col) {
  ByteSpan cursor = SpanOf(sheet_bytes);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaNum)) {
      continue;
    }
    ByteSpan p = record_or.value().payload;
    auto col_or = read_u32(p);
    if (!col_or || col_or.value() != want_col) {
      continue;
    }
    // Skip iStyleRef(3) + fPhShow(1) + the double value(8) + grbitFlags(2).
    if (p.size < 14) {
      return {};
    }
    p.data += 14;
    p.size -= 14;
    auto cce_or = read_u32(p);
    if (!cce_or || cce_or.value() > p.size) {
      return {};
    }
    return std::vector<std::uint8_t>(p.data, p.data + cce_or.value());
  }
  return {};
}
}  // namespace

TEST(XlsbWriter, RootLevelSameSheetRangeUsesValueClassArea) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Sheet1"));
  sheet.set_cell_formula(1U, 2U, "=A10:A11");  // C2
  ASSERT_TRUE(sheet.commit_spill(1U, 2U, 2U, 1U, {Value::number(5.0), Value::number(7.0)}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindArrFmlaRgce(sheet_or.value(), 2U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x45U);  // PtgArea, value class (0x25 | value-class bit)
}

TEST(XlsbWriter, RootLevelCrossSheetRangeUsesValueClassArea3d) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  Sheet& sheet2 = wb.sheet(wb.add_sheet("Sheet2"));
  sheet2.set_cell_formula(1U, 1U, "=Sheet1!A10:A11");  // B2
  ASSERT_TRUE(sheet2.commit_spill(1U, 1U, 2U, 1U, {Value::number(5.0), Value::number(7.0)}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet2.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindArrFmlaRgce(sheet_or.value(), 1U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x5BU);  // PtgArea3d, value class (0x3B | value-class bit)
}

TEST(XlsbWriter, RootLevelSingleCellReferenceUsesValueClassRef) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Sheet1"));
  sheet.set_cell_formula(1U, 2U, "=A10");  // C2, plain (non-array) formula cell

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindFmlaNumRgce(sheet_or.value(), 2U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x44U);  // PtgRef, value class (0x24 | value-class bit)
}

TEST(XlsbWriter, RootLevelCrossSheetSingleCellReferenceUsesValueClassRef3d) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  Sheet& sheet2 = wb.sheet(wb.add_sheet("Sheet2"));
  sheet2.set_cell_formula(1U, 2U, "=Sheet1!A10");  // C2

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet2.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindFmlaNumRgce(sheet_or.value(), 2U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x5AU);  // PtgRef3d, value class (0x3A | value-class bit)
}

TEST(XlsbWriter, NestedRangeArgumentStaysReferenceClass) {
  // A reference consumed as a function argument is unaffected by the
  // root-position promotion above: it must stay reference class, exactly
  // as before. Real Excel 365 reads a value/array-class 3-D reference used
  // this way as `#REF!`.
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Sheet1"));
  sheet.set_cell_formula(1U, 2U, "=SUM(A10:A11)");  // C2, plain (non-array) formula cell

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindFmlaNumRgce(sheet_or.value(), 2U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x25U);  // PtgArea, reference class -- SUM's argument
}

TEST(XlsbWriter, RootLevelChainedRangeWithDefinedNameWrapsInValueClassMemFunc) {
  // The general (non-collapsible) form of a `:` range at root position:
  // one endpoint is a defined name, so the fast-path Area/Area3d collapse
  // in `emit_range` does not apply. For an Excel 365 cell
  // `=Sheet1!A10:Book!Rng`: real Excel wraps `operand + operand +
  // PtgRange` in a value-class `PtgMemFunc` ("function returns a range",
  // per ptg.h) carrying the wrapped run's byte length ahead of it, because
  // `PtgRange` itself carries no class bits to promote directly.
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Sheet1"));
  DefinedName dn;
  dn.name = "Rng";
  dn.formula = "$A$10:$A$11";
  dn.local_sheet_id = -1;
  wb.set_defined_names({dn});
  sheet.set_cell_formula(1U, 2U, "=A10:Rng");  // C2
  ASSERT_TRUE(sheet.commit_spill(1U, 2U, 2U, 1U, {Value::number(5.0), Value::number(7.0)}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindArrFmlaRgce(sheet_or.value(), 2U);
  ASSERT_GE(rgce.size(), 3U);
  EXPECT_EQ(rgce[0], 0x49U);  // PtgMemFunc, value class (0x29 | value-class bit)
  const std::uint16_t wrapped_len = static_cast<std::uint16_t>(rgce[1] | (rgce[2] << 8U));
  ASSERT_EQ(static_cast<std::size_t>(wrapped_len), rgce.size() - 3U);
  EXPECT_EQ(rgce.back(), 0x11U);  // the wrapped run ends with PtgRange (`:`)
}

namespace {

// Returns the `rgce` bytes of the `BrtName` record spelling `want_name`, or
// an empty vector if none matches.
std::vector<std::uint8_t> FindNameRgce(const std::vector<std::uint8_t>& workbook_bin, std::string_view want_name) {
  ByteSpan cursor = SpanOf(workbook_bin);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtName)) {
      continue;
    }
    ByteSpan p = record_or.value().payload;
    if (p.size < 9U) {
      return {};
    }
    p.data += 9;  // grbit(4) + chKey(1) + itab(4)
    p.size -= 9;
    auto cch_or = read_u32(p);
    if (!cch_or || cch_or.value() * 2U > p.size) {
      return {};
    }
    std::string decoded;
    decoded.reserve(cch_or.value());
    for (std::uint32_t i = 0; i < cch_or.value(); ++i) {
      decoded.push_back(static_cast<char>(p.data[i * 2U]));
    }
    p.data += cch_or.value() * 2U;
    p.size -= cch_or.value() * 2U;
    auto cce_or = read_u32(p);
    if (!cce_or || cce_or.value() > p.size) {
      return {};
    }
    if (decoded != want_name) {
      continue;
    }
    return std::vector<std::uint8_t>(p.data, p.data + cce_or.value());
  }
  return {};
}

// Returns the first nine payload bytes (grbit, chKey, itab) of the `BrtName`
// record spelling `want_name`, or an empty vector if none matches.
std::vector<std::uint8_t> FindNameRecordPrefix(const std::vector<std::uint8_t>& workbook_bin,
                                               std::string_view want_name) {
  ByteSpan cursor = SpanOf(workbook_bin);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    const ByteSpan p = record_or.value().payload;
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtName) || p.size < 13U) {
      continue;
    }
    const std::uint32_t cch = static_cast<std::uint32_t>(p.data[9]) | (static_cast<std::uint32_t>(p.data[10]) << 8);
    if (cch != want_name.size() || p.size < 13U + cch * 2U) {
      continue;
    }
    std::string decoded;
    for (std::uint32_t i = 0; i < cch; ++i) {
      decoded.push_back(static_cast<char>(p.data[13U + i * 2U]));
    }
    if (decoded == want_name) {
      return std::vector<std::uint8_t>(p.data, p.data + 9);
    }
  }
  return {};
}

// Returns `xl/workbook.bin`'s `BrtWbView` `itabCur` field (u32 at payload
// offset 24), or `0xFFFFFFFF` if the record is missing or truncated.
std::uint32_t FindItabCur(const std::vector<std::uint8_t>& workbook_bin) {
  ByteSpan cursor = SpanOf(workbook_bin);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return 0xFFFFFFFFU;
    }
    if (record_or.value().type != 158U) {  // BrtWbView
      continue;
    }
    ByteSpan p = record_or.value().payload;
    if (p.size < 28U) {
      return 0xFFFFFFFFU;
    }
    p.data += 24;
    auto itab_or = read_u32(p);
    return itab_or ? itab_or.value() : 0xFFFFFFFFU;
  }
  return 0xFFFFFFFFU;
}

}  // namespace

TEST(XlsbWriter, DefinedNameBodyRootReferenceStaysReferenceClass) {
  // For an Excel 365 BrtName "Rng" (=Sheet1!$A$10:$A$11):
  // unlike a cell formula, a defined
  // name's own body keeps reference class at its root.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  DefinedName dn;
  dn.name = "Rng";
  dn.formula = "Sheet1!$A$10:$A$11";
  dn.local_sheet_id = -1;
  wb.set_defined_names({dn});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));

  const std::vector<std::uint8_t> rgce = FindNameRgce(wb_bin_or.value(), "Rng");
  const std::vector<std::uint8_t> want = {0x3B, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x0A,
                                          0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
  EXPECT_EQ(rgce, want);
}

TEST(XlsbWriter, ItabCurMatchesTheTabSelectedSheet) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  Sheet& sheet2 = wb.sheet(wb.add_sheet("Sheet2"));
  sheet2.mutable_view().tab_selected = true;

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));

  EXPECT_EQ(FindItabCur(wb_bin_or.value()), 1U);
}

TEST(XlsbWriter, ItabCurIsZeroWhenNoSheetIsTabSelected) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));

  EXPECT_EQ(FindItabCur(wb_bin_or.value()), 0U);
}

TEST(XlsbWriter, RealFormulaRoundTripsAsFormulaCell) {
  // An engine-authored formula encodes to a Ptg stream, survives the
  // write, and decodes back to the same formula text on read.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(2U, 3U, "=A1+B2*3");

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;

  const Cell* c = read_or.value().workbook.sheet(0).cell_at(2U, 3U);
  ASSERT_NE(c, nullptr);
  EXPECT_EQ(c->formula_text, "=A1+B2*3");
}

TEST(XlsbWriter, RespelledLiteralsRoundTripWithIdenticalFormulaText) {
  // Entered text stores its literals respelled; the Ptg stream reads back
  // as that same text, including a sign the entry kept (`=++5` -> `=+5`,
  // `=--0` -> `=-0`).
  struct Case {
    const char* typed;
    const char* stored;
  };
  const Case kCases[] = {
      {"=1.50+1E21*2.5E-18", "=1.5+1E+21*0.0000000000000000025"},
      {"=1E20-1E-20", "=100000000000000000000-1E-20"},
      {"=++5", "=+5"},
      {"=--0", "=-0"},
      {"=--5", "=--5"},
      {"=+-5", "=+-5"},
      {"=1-+5", "=1-5"},
      {"=-5^2", "=-5^2"},
      {"=SUM(-0,+1)", "=SUM(0,1)"},
      {"={+5,-0,1.50}", "={5,0,1.5}"},
  };
  Workbook wb = Workbook::create_empty();
  const std::size_t sheet = wb.add_sheet("F");
  std::uint32_t row = 0U;
  for (const Case& c : kCases) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(sheet, row, 0U, c.typed)));
    EXPECT_EQ(wb.sheet(sheet).cell_at(row, 0U)->formula_text, c.stored);
    ++row;
  }

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& rt = read_or.value().workbook.sheet(0);
  row = 0U;
  for (const Case& c : kCases) {
    const Cell* cell = rt.cell_at(row, 0U);
    ASSERT_NE(cell, nullptr) << c.typed;
    EXPECT_EQ(cell->formula_text, c.stored) << c.typed;
    ++row;
  }
}

// On Excel 365: it saves
// `S!Name` that neither `S` nor the workbook defines against an empty stub
// scoped to `S`, never against another sheet's definition, so `Sheet2!SVal`
// stays #NAME? while only Sheet1 defines SVal.
TEST(XlsbWriter, SheetQualifiedNamesKeepTheirQualifierAndScope) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  DefinedName dn;
  dn.name = "SVal";
  dn.formula = "9";
  dn.local_sheet_id = 0;
  wb.set_defined_names({dn});
  struct Case {
    const char* formula;
    Value want;
  };
  const Case cases[] = {
      {"=Sheet2!SVal", Value::error(ErrorCode::Name)},      {"=Sheet1!SVal", Value::number(9.0)},
      {"=Sheet1!NOSUCH(1)", Value::error(ErrorCode::Name)}, {"=Sheet2!NOSUCH(1)", Value::error(ErrorCode::Name)},
      {"=NOSUCH(1)", Value::error(ErrorCode::Name)},        {"=Sheet1!A1(1)", Value::error(ErrorCode::Ref)},
  };
  for (std::uint32_t row = 0; row < std::size(cases); ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, row, 1U, cases[row].formula)));
  }

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  Workbook& back = read_or.value().workbook;
  ASSERT_TRUE(static_cast<bool>(back.recalc(eval::default_registry())));

  for (std::uint32_t row = 0; row < std::size(cases); ++row) {
    const Cell* c = back.sheet(0).cell_at(row, 1U);
    ASSERT_NE(c, nullptr) << cases[row].formula;
    EXPECT_EQ(c->formula_text, cases[row].formula);
    EXPECT_EQ(c->cached_value, cases[row].want) << cases[row].formula;
  }
}

TEST(XlsbWriter, NameSpelledInAnotherCaseReusesTheDefinedNameRecord) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Sheet1"));
  sheet.set_cell_value(0, 0, Value::number(5.0));
  wb.set_defined_names({DefinedName{"Äpfel", "Sheet1!$A$1", -1, false, ""}});
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 1U, "=äPFEL*2")));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  Workbook& back = read_or.value().workbook;
  ASSERT_EQ(back.defined_names().size(), 1U);
  EXPECT_EQ(back.defined_names()[0].name, "Äpfel");
  ASSERT_TRUE(static_cast<bool>(back.recalc(eval::default_registry())));
  const Cell* c = back.sheet(0).cell_at(0, 1U);
  ASSERT_NE(c, nullptr);
  EXPECT_EQ(c->formula_text, "=Äpfel*2");
  EXPECT_EQ(c->cached_value, Value::number(10.0));
}

TEST(XlsbWriter, DefinedNameWithFutureFunctionRoundTrips) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Data");
  wb.set_defined_names({DefinedName{"Joined", "TEXTJOIN(\",\",TRUE,A1:A2)", -1, false, ""}});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  ASSERT_EQ(read_or.value().workbook.defined_names().size(), 1U);
  EXPECT_EQ(read_or.value().workbook.defined_names()[0].name, "Joined");
  // Read back in formula-bar spelling, as the OOXML reader does.
  EXPECT_EQ(read_or.value().workbook.defined_names()[0].formula, "TEXTJOIN(\",\",TRUE,A1:A2)");
}

// True when `haystack` contains `needle` encoded the way `BrtName`
// stores a name: UTF-16LE, no BOM. Excel names are ASCII, so each byte
// is followed by a zero byte.
bool ContainsUtf16Le(const std::vector<std::uint8_t>& haystack, std::string_view needle) {
  std::vector<std::uint8_t> wide;
  wide.reserve(needle.size() * 2U);
  for (const char c : needle) {
    wide.push_back(static_cast<std::uint8_t>(c));
    wide.push_back(0U);
  }
  return std::search(haystack.begin(), haystack.end(), wide.begin(), wide.end()) != haystack.end();
}

TEST(XlsbWriter, FutureFunctionCallRegistersHiddenName) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=XLOOKUP(\"k\",A1:A3,B1:B3)");

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value().bytes))));
  auto workbook_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(workbook_or));
  EXPECT_TRUE(ContainsUtf16Le(workbook_or.value(), "_xlfn.XLOOKUP"));
}

// On Excel 365: it stores a cached #SPILL! or #CALC! as #VALUE! (the real
// error lives in the rich value the cell's `BrtValueMeta` names; a cell with
// no recorded detail is the plain #VALUE!) and #GETTING_DATA as #N/A. A newer
// error byte in a cell made Excel refuse the whole file.
TEST(XlsbWriter, NewerErrorsAreStoredAsTheirLegacyFallback) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("E"));
  const ErrorCode errors[] = {ErrorCode::Spill, ErrorCode::Calc, ErrorCode::GettingData, ErrorCode::Div0};
  const ErrorCode stored[] = {ErrorCode::Value, ErrorCode::Value, ErrorCode::NA, ErrorCode::Div0};
  for (std::uint32_t row = 0; row < std::size(errors); ++row) {
    s.set_cell_formula(row, 0U, "=1/0");
    s.set_cell_cached_value(row, 0U, Value::error(errors[row]));
    s.set_cell_value(row, 1U, Value::error(errors[row]));
  }

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& back = read_or.value().workbook.sheet(0);
  for (std::uint32_t row = 0; row < std::size(errors); ++row) {
    for (std::uint32_t col = 0; col < 2U; ++col) {
      const Cell* c = back.cell_at(row, col);
      ASSERT_NE(c, nullptr);
      const Value& v = c->cached_value;
      EXPECT_EQ(v, Value::error(stored[row])) << "row " << row << " col " << col;
    }
  }
}

// On Excel 365: it keeps `_xlfn.` on a name it
// does not know, in the formula bar and in the file.
TEST(XlsbWriter, UnknownPrefixedNameKeepsItsPrefix) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0U, 0U, "=_xlfn.FOOBAR(1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1U, 0U, "=_xlfn.XLOOKUP(1,B1:B2,C1:C2)")));
  EXPECT_EQ(wb.sheet(0).cell_at(0U, 0U)->formula_text, "=_xlfn.FOOBAR(1)");
  EXPECT_EQ(wb.sheet(0).cell_at(1U, 0U)->formula_text, "=XLOOKUP(1,B1:B2,C1:C2)");

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  // `_xlfn.FOOBAR` is a plain undefined-name stub (flags 0, empty body), and
  // only a function Excel knows gets the hidden future-function record.
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));
  EXPECT_EQ(FindNameRecordPrefix(wb_bin_or.value(), "_xlfn.FOOBAR"),
            (std::vector<std::uint8_t>{0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF}));
  EXPECT_TRUE(FindNameRgce(wb_bin_or.value(), "_xlfn.FOOBAR").empty());
  EXPECT_EQ(FindNameRecordPrefix(wb_bin_or.value(), "_xlfn.XLOOKUP"),
            (std::vector<std::uint8_t>{0x0B, 0x00, 0x02, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF}));

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& back = read_or.value().workbook.sheet(0);
  EXPECT_EQ(back.cell_at(0U, 0U)->formula_text, "=_xlfn.FOOBAR(1)");
  EXPECT_EQ(back.cell_at(1U, 0U)->formula_text, "=XLOOKUP(1,B1:B2,C1:C2)");
}

// A loaded formula keeps the meaning its file gives it. With no dynamic-array
// mark (`cm` / `BrtCellMeta`), a fixed-size CSE block and an
// implicit-intersection =SUM(A1:A2*2) are not dynamic-array formulas, and a
// save through either container keeps them so: the block its array range,
// neither of them a mark. A formula entered through `Workbook` is marked and
// keeps the mark across both containers.
TEST(XlsbWriter, LegacyFormulasKeepTheirMeaningAcrossContainers) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("Sheet1"));
  s.set_cell_value(0U, 0U, Value::number(1.0));
  s.set_cell_value(1U, 0U, Value::number(2.0));
  s.set_cell_formula(0U, 2U, "=A1:A2*2");
  ASSERT_TRUE(s.commit_spill(0U, 2U, 2U, 1U, {Value::number(2.0), Value::number(4.0)}));
  s.set_cell_formula(0U, 4U, "=SUM(A1:A2*2)");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0U, 5U, "=SUM(A1:A2*2)")));

  auto check = [](const Workbook& back, const char* label) {
    const Sheet& sheet = back.sheet(0);
    EXPECT_FALSE(sheet.cell_at(0U, 2U)->dynamic_array) << label;
    const SpillRegion* block = sheet.spill_region_at_anchor(0U, 2U);
    ASSERT_NE(block, nullptr) << label;
    EXPECT_EQ(block->rows, 2U) << label;
    EXPECT_FALSE(sheet.cell_at(0U, 4U)->dynamic_array) << label;
    EXPECT_TRUE(sheet.cell_at(0U, 5U)->dynamic_array) << label;
  };

  auto xlsx_or = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx_or)) << xlsx_or.error().message;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(xlsx_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  const std::string sheet_xml(sheet_or.value().begin(), sheet_or.value().end());
  EXPECT_NE(sheet_xml.find("<c r=\"C1\"><f t=\"array\" ref=\"C1:C2\">"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("<c r=\"E1\"><f>SUM(A1:A2*2)</f>"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("<c r=\"F1\" cm=\"1\"><f t=\"array\" ref=\"F1\">"), std::string::npos) << sheet_xml;
  auto from_xlsx = read_ooxml(SpanOf(xlsx_or.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  check(from_xlsx.value().workbook, "xlsx");

  auto xlsb_or = write_xlsb(from_xlsx.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message;
  auto from_xlsb = read_xlsb(SpanOf(xlsb_or.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  check(from_xlsb.value().workbook, "xlsb");
}

TEST(XlsbWriter, CubeFunctionsEncodeWithTheirFunctionIds) {
  // Excel 365 saves `CUBEVALUE("c","m")` as two strings and
  // `PtgFuncVar(2, 380)`, not through a hidden `_xlfn.` name.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=CUBEVALUE(\"c\",\"m\")");

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value().bytes))));
  auto workbook_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(workbook_or));
  EXPECT_FALSE(ContainsUtf16Le(workbook_or.value(), "CUBEVALUE"));

  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Cell* cell = read_or.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=CUBEVALUE(\"c\",\"m\")");
}

TEST(XlsbWriter, UndefinedNameCallIsStoredAgainstAWorkbookStub) {
  // On Excel 365: a lone `NOSUCH(1)` is
  // `PtgName` + the argument + `PtgFuncVar(255)`, the name an empty,
  // visible, workbook-scoped BrtName.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0U, 0U, "=NOSUCH(1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1U, 0U, "=NOSUCHREF")));

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value().bytes))));
  auto workbook_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(workbook_or));
  EXPECT_EQ(FindNameRecordPrefix(workbook_or.value(), "NOSUCH"),
            (std::vector<std::uint8_t>{0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF}));

  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  Workbook& back = read_or.value().workbook;
  ASSERT_TRUE(static_cast<bool>(back.recalc(eval::default_registry())));
  for (std::uint32_t row = 0; row < 2U; ++row) {
    const Cell* cell = back.sheet(0).cell_at(row, 0U);
    ASSERT_NE(cell, nullptr);
    EXPECT_EQ(cell->formula_text, row == 0U ? "=NOSUCH(1)" : "=NOSUCHREF");
    EXPECT_EQ(cell->cached_value, Value::error(ErrorCode::Name));
  }
}

TEST(XlsbWriter, LocalisedJisSpellingSavesAsTheStoredDbcsSpelling) {
  // Excel localises the formula bar but not the file: `JIS` is the ja-JP
  // spelling of `DBCS`, and Excel stores the call as `DBCS` with
  // function id 215 in both containers. A workbook carrying either
  // spelling must save without degrading, and both come back as the
  // stored spelling.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=JIS(\"ABC\")");
  s.set_cell_formula(1U, 0U, "=DBCS(\"ABC\")");

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);

  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& rt = read_or.value().workbook.sheet(0);
  const Cell* from_jis = rt.cell_at(0U, 0U);
  ASSERT_NE(from_jis, nullptr);
  EXPECT_EQ(from_jis->formula_text, "=DBCS(\"ABC\")");
  const Cell* from_dbcs = rt.cell_at(1U, 0U);
  ASSERT_NE(from_dbcs, nullptr);
  EXPECT_EQ(from_dbcs->formula_text, "=DBCS(\"ABC\")");
}

TEST(XlsbWriter, HarvestedFuncIdCallsRoundTripWithIdenticalFormulaText) {
  // The ids for these callees were decoded from an Excel-produced
  // workbook (`tests/fixtures/excel/xlsb_func_ids.xlsb`). Each spells a
  // different encoding decision the table drives: a fixed-arity
  // `PtgFunc` (MROUND), a `PtgFuncVar` at its minimum and its maximum
  // arity (WEEKNUM), and an open-ended variadic (GCD). None may degrade
  // to a cached literal, and every formula must come back byte-identical.
  constexpr std::uint32_t kFormulaCount = 5U;
  const char* kFormulas[kFormulaCount] = {
      "=MROUND(17,5)", "=WEEKNUM(43922)", "=WEEKNUM(43922,1)", "=GCD(24,36,60)", "=YEARFRAC(43831,44197,0)",
  };
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  for (std::uint32_t i = 0; i < kFormulaCount; ++i) {
    s.set_cell_formula(i, 0U, kFormulas[i]);
  }

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);

  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& rt = read_or.value().workbook.sheet(0);
  for (std::uint32_t i = 0; i < kFormulaCount; ++i) {
    const Cell* cell = rt.cell_at(i, 0U);
    ASSERT_NE(cell, nullptr) << kFormulas[i];
    EXPECT_EQ(cell->formula_text, kFormulas[i]);
  }
}

TEST(XlsbWriter, UnencodableFormulaDowngradesToCachedLiteralAndReportsIt) {
  // A structured reference to a table the workbook lacks (`T[C]`) has no
  // Ptg lowering. A single unsupported formula must not make the
  // whole workbook unsaveable: it degrades to its cached literal and the
  // explicit result count records the loss.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=T[C]");
  s.set_cell_cached_value(0U, 0U, Value::number(42.0));

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 1U);
  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Cell* cell = read_or.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->formula_text.empty());
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 42.0);
}

TEST(XlsbWriter, UnencodableSpillFormulaDowngradesAnchorWithoutPhantoms) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=T[C]");
  ASSERT_TRUE(s.commit_spill(0U, 0U, 1U, 2U, {Value::number(11.0), Value::number(12.0)}));

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 1U);
  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& rt = read_or.value().workbook.sheet(0);
  const Cell* anchor = rt.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_TRUE(anchor->formula_text.empty());
  ASSERT_TRUE(anchor->cached_value.is_number());
  EXPECT_DOUBLE_EQ(anchor->cached_value.as_number(), 11.0);
  EXPECT_EQ(rt.cell_at(0U, 1U), nullptr);
}

TEST(XlsbWriter, GeneratedPartsBeatPassthroughOnCollision) {
  // Passthrough trying to ride on the writer's reserved
  // `xl/workbook.bin` path: must be dropped (writer logs a warning,
  // not asserted here), and the round-trip must still succeed.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("S1"));
  s.set_cell_value(0U, 0U, Value::number(1.0));

  std::vector<PassthroughPart> parts;
  PassthroughPart bogus_workbook;
  bogus_workbook.path = "xl/workbook.bin";
  bogus_workbook.content_type = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
  bogus_workbook.bytes = {0xDE, 0xAD, 0xBE, 0xEF};  // garbage — would break the package if it survived
  parts.push_back(std::move(bogus_workbook));
  wb.set_passthrough_parts(std::move(parts));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  // Workbook is intact: we got our sheet back.
  ASSERT_EQ(read_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(read_or.value().workbook.sheet(0).name(), "S1");
}
TEST(XlsbWriter, WholeColumnFormulaSavesWithoutDowngrade) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Ranges"));
  sheet.set_cell_formula(0U, 0U, "=SUM(A:A)");

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Cell* cell = read_or.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=SUM(A:A)");
}

namespace {

// One `BrtExternSheet` entry as written: iSupBook, itabFirst, itabLast.
struct WrittenXti {
  std::uint32_t book = 0;
  std::uint32_t first = 0;
  std::uint32_t last = 0;
};

// Returns the entries of `xl/workbook.bin`'s `BrtExternSheet`, or an empty
// vector when the record is missing or truncated.
std::vector<WrittenXti> FindExternSheet(const std::vector<std::uint8_t>& workbook_bin) {
  ByteSpan cursor = SpanOf(workbook_bin);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtExternSheet)) {
      continue;
    }
    ByteSpan p = record_or.value().payload;
    auto count_or = read_u32(p);
    if (!count_or) {
      return {};
    }
    std::vector<WrittenXti> out;
    for (std::uint32_t i = 0; i < count_or.value(); ++i) {
      auto book_or = read_u32(p);
      auto first_or = read_u32(p);
      auto last_or = read_u32(p);
      if (!book_or || !first_or || !last_or) {
        return {};
      }
      out.push_back(WrittenXti{book_or.value(), first_or.value(), last_or.value()});
    }
    return out;
  }
  return {};
}

std::uint16_t U16At(const std::vector<std::uint8_t>& bytes, std::size_t at) {
  return static_cast<std::uint16_t>(bytes[at] | (bytes[at + 1U] << 8U));
}

std::uint32_t U32At(const std::vector<std::uint8_t>& bytes, std::size_t at) {
  return static_cast<std::uint32_t>(bytes[at]) | (static_cast<std::uint32_t>(bytes[at + 1U]) << 8U) |
         (static_cast<std::uint32_t>(bytes[at + 2U]) << 16U) | (static_cast<std::uint32_t>(bytes[at + 3U]) << 24U);
}

// The XTI the ixti at `rgce[at]` names, or {~0, ~0, ~0} when out of range.
WrittenXti XtiAt(const std::vector<WrittenXti>& xti, const std::vector<std::uint8_t>& rgce, std::size_t at) {
  const std::uint16_t ixti = U16At(rgce, at);
  return ixti < xti.size() ? xti[ixti] : WrittenXti{0xFFFFFFFFU, 0xFFFFFFFFU, 0xFFFFFFFFU};
}

constexpr std::uint8_t kClassStrip = 0x9FU;  // clears the two class bits of a Ptg byte

}  // namespace

TEST(XlsbWriter, SitedRefErrorsUseTheirSheetsXti) {
  Workbook wb = Workbook::create_empty();
  const std::size_t first = wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  wb.add_sheet("Sheet3");
  Sheet& sheet1 = wb.sheet(first);
  sheet1.set_cell_formula(0U, 0U, "=Sheet3!#REF!*2");   // A1
  sheet1.set_cell_formula(0U, 1U, "=#REF!A1");          // B1
  sheet1.set_cell_formula(0U, 2U, "=SUM(#REF!A1:B2)");  // C1
  sheet1.set_cell_formula(0U, 3U, "=SUM(#REF!A:A)");    // D1
  sheet1.set_cell_formula(0U, 4U, "=#REF!#REF!");       // E1
  sheet1.set_cell_formula(0U, 5U, "=Nope!#REF!");       // F1
  sheet1.set_cell_formula(0U, 6U, "=#REF!");            // G1

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value().bytes))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));
  const std::vector<WrittenXti> xti = FindExternSheet(wb_bin_or.value());

  // Sheet3!#REF!: PtgRefErr3d through Sheet3's XTI, zero location.
  const std::vector<std::uint8_t> qualified = FindFmlaNumRgce(sheet_or.value(), 0U);
  ASSERT_GE(qualified.size(), 9U);
  EXPECT_EQ(qualified[0] & kClassStrip, 0x3CU & kClassStrip);
  const WrittenXti sheet3 = XtiAt(xti, qualified, 1U);
  EXPECT_EQ(sheet3.book, 0U);
  EXPECT_EQ(sheet3.first, 2U);
  EXPECT_EQ(sheet3.last, 2U);
  EXPECT_EQ(U32At(qualified, 3U), 0U);
  EXPECT_EQ(U16At(qualified, 7U), 0U);

  // #REF!A1: PtgRef3d through the deleted-sheet XTI {0, -1, -1}.
  const std::vector<std::uint8_t> cell = FindFmlaNumRgce(sheet_or.value(), 1U);
  ASSERT_GE(cell.size(), 9U);
  EXPECT_EQ(cell[0] & kClassStrip, 0x3AU & kClassStrip);
  const WrittenXti dead = XtiAt(xti, cell, 1U);
  EXPECT_EQ(dead.book, 0U);
  EXPECT_EQ(dead.first, 0xFFFFFFFFU);
  EXPECT_EQ(dead.last, 0xFFFFFFFFU);
  EXPECT_EQ(U32At(cell, 3U), 0U);            // row 1
  EXPECT_EQ(U16At(cell, 7U) & 0x3FFFU, 0U);  // column A

  // SUM(#REF!A1:B2): PtgArea3d through the same XTI.
  const std::vector<std::uint8_t> area = FindFmlaNumRgce(sheet_or.value(), 2U);
  ASSERT_GE(area.size(), 15U);
  EXPECT_EQ(area[0] & kClassStrip, 0x3BU & kClassStrip);
  EXPECT_EQ(U16At(area, 1U), U16At(cell, 1U));
  EXPECT_EQ(U32At(area, 3U), 0U);
  EXPECT_EQ(U32At(area, 7U), 1U);
  EXPECT_EQ(U16At(area, 11U) & 0x3FFFU, 0U);
  EXPECT_EQ(U16At(area, 13U) & 0x3FFFU, 1U);

  // SUM(#REF!A:A): PtgArea3d over every row.
  const std::vector<std::uint8_t> column = FindFmlaNumRgce(sheet_or.value(), 3U);
  ASSERT_GE(column.size(), 15U);
  EXPECT_EQ(column[0] & kClassStrip, 0x3BU & kClassStrip);
  EXPECT_EQ(U16At(column, 1U), U16At(cell, 1U));
  EXPECT_EQ(U32At(column, 3U), 0U);
  EXPECT_EQ(U32At(column, 7U), 1048575U);
  EXPECT_EQ(U16At(column, 11U) & 0x3FFFU, 0U);
  EXPECT_EQ(U16At(column, 13U) & 0x3FFFU, 0U);

  // #REF!#REF!: PtgRefErr3d through the deleted-sheet XTI.
  const std::vector<std::uint8_t> no_address = FindFmlaNumRgce(sheet_or.value(), 4U);
  ASSERT_GE(no_address.size(), 9U);
  EXPECT_EQ(no_address[0] & kClassStrip, 0x3CU & kClassStrip);
  EXPECT_EQ(U16At(no_address, 1U), U16At(cell, 1U));
  EXPECT_EQ(U32At(no_address, 3U), 0U);
  EXPECT_EQ(U16At(no_address, 7U), 0U);

  // A sheet the workbook lacks, and an unsited #REF!: PtgRefErr.
  const std::vector<std::uint8_t> missing = FindFmlaNumRgce(sheet_or.value(), 5U);
  ASSERT_FALSE(missing.empty());
  EXPECT_EQ(missing[0] & kClassStrip, 0x2AU & kClassStrip);
  const std::vector<std::uint8_t> bare = FindFmlaNumRgce(sheet_or.value(), 6U);
  ASSERT_FALSE(bare.empty());
  EXPECT_EQ(bare[0] & kClassStrip, 0x2AU & kClassStrip);
}

TEST(XlsbWriter, DeletedSheetXtiStaysDistinctFromTheBookScopeXti) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet1 = wb.sheet(wb.add_sheet("Sheet1"));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("G", "5")));
  sheet1.set_cell_formula(0U, 0U, "=[0]!G");    // A1
  sheet1.set_cell_formula(0U, 1U, "=#REF!A1");  // B1

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));
  const std::vector<WrittenXti> xti = FindExternSheet(wb_bin_or.value());
  const auto has = [&xti](std::uint32_t itab) {
    return std::any_of(xti.begin(), xti.end(),
                       [itab](const WrittenXti& e) { return e.book == 0U && e.first == itab && e.last == itab; });
  };
  EXPECT_TRUE(has(0xFFFFFFFEU));
  EXPECT_TRUE(has(0xFFFFFFFFU));
}

TEST(XlsbWriter, SpillOfAnotherSheetsAnchorRegistersItsXti) {
  Workbook wb = Workbook::create_empty();
  const std::size_t first = wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  wb.add_sheet("Sheet3");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(first, 0U, 0U, "=Sheet3!D5#")));

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Cell* cell = read_or.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=Sheet3!D5#");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
