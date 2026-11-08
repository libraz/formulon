// Stable C ABI tests for the R1C1 reading of entered cell formulas. Each
// test rebuilds one Excel probe sheet cell for cell: A1=100, A2=20 (B1=1000
// on the second and third), and on each probe row i: D=7, E=the entered
// formula, F=IFERROR(ERROR.TYPE(Ei),-1), G=FORMULATEXT(Ei). One recalc runs
// after every entry. Expected spellings and values are Excel 365's.

#include <cstddef>
#include <cstdint>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_out_index.h"
#include "formulon_c_setters.h"
#include "gtest/gtest.h"
#include "parser/parse_status.h"
#include "table.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kParserUnexpectedToken =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kParserUnexpectedToken);

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

// One probe row: the 1-based row, the formula typed into column E, Excel's
// FORMULATEXT (nullptr when Excel refused the entry), its value and its
// ERROR.TYPE (-1 for none; `value` is meaningful only then).
struct ProbeRow {
  std::uint32_t row;
  const char* formula;
  const char* stored;
  double value;
  int error_type;
};

constexpr std::uint32_t kColD = 3;
constexpr std::uint32_t kColE = 4;
constexpr std::uint32_t kColF = 5;
constexpr std::uint32_t kColG = 6;

const ProbeRow kOut6[] = {
    {3, "=RC[-1]", "=D3", 7, -1},
    {4, "=R1C1", "=$A$1", 100, -1},
    {5, "=R[0]C[-1]", "=D5", 7, -1},
    {6, "=LET(B1,5,B1+R1C1)", "=LET(B1,5,B1+$A$1)", 105, -1},
    {7, "=LET(B1,5,B1+RC[-1])", "=LET(B1,5,B1+D7)", 12, -1},
    {8, "=LET(B1,5,B1+R[0]C[-1])", "=LET(B1,5,B1+D8)", 12, -1},
    {9, "=LET(B1,5,B1+R1C1:R2C1)", "=LET(B1,5,B1+$A$1:$A$2)", 0, 9},
    {10, "=LET(B1,5,SUM(B1,R1C1:R2C1))", "=LET(B1,5,SUM(B1,$A$1:$A$2))", 125, -1},
    {11, "=LET(B1,5,B1+Sheet1!R1C1)", "=LET(B1,5,B1+Sheet1!$A$1)", 105, -1},
    {12, "=LET(B1,5,SUM(B1,C1))", "=LET(B1,5,SUM(B1,$A:$A))", 125, -1},
    {13, "=LET(B1,5,COUNT(R1))", "=LET(B1,5,COUNT($1:$1))", 1, -1},
    {14, "=B2+LET(B1,5,B1)", "='B2'+LET(B1,5,B1)", 0, 5},
    {15, "=LET(B1,5,B1)+B2", "=LET(B1,5,B1)+'B2'", 0, 5},
    {16, "=LET(x,1,x)+LET(B1,5,B1)", "=LET(x,1,x)+LET(B1,5,B1)", 6, -1},
    {17, "=LET(x,5,x+B2)", "=LET(x,5,x+B2)", 5, -1},
    {18, "=LET(B1,100,LOG10(B1))", "=LET(B1,100,LOG10(B1))", 2, -1},
    {19, "=LET(B1,5,B1+XFE1)", "=LET(B1,5,B1+XFE1)", 0, 5},
    {20, "=LET(B1,5,R1,6,B1+R1)", nullptr, 0, -1},
    {21, "=LAMBDA(B1,B1+R1C1)(4)", "=LAMBDA(B1,B1+$A$1)(4)", 104, -1},
    {22, "=LET(B1,5,B1+'B2')", nullptr, 0, -1},
    {23, "='B2'", nullptr, 0, -1},
    {24, "=LET(x,5,x+'B2')", nullptr, 0, -1},
    {25, "=LET(B1,5,B1+A:A)", "=LET(B1,5,B1+A:(A))", 0, 5},
    {26, "=LET(B1,5,B1+$B1)", nullptr, 0, -1},
    {27, "=LET(B1,5,B1+R[1]C)", "=LET(B1,5,B1+E28)", 8, -1},
    {28, "=LET(B1,5,ROWS(R1C1:R3C1))", "=LET(B1,5,ROWS($A$1:$A$3))", 3, -1},
    {29, "=LET(B1,5,ROW(R2C1))", "=LET(B1,5,ROW($A$2))", 2, -1},
    {30, "=LET(B1,5,INDIRECT(\"A1\")+B1)", "=LET(B1,5,INDIRECT(\"A1\")+B1)", 105, -1},
    {31, "=LET(B1,5,INDIRECT(\"R1C1\",FALSE)+B1)", "=LET(B1,5,INDIRECT(\"R1C1\",FALSE)+B1)", 105, -1},
    {32, "=LET(B1,5,B1+A1)", "=LET(B1,5,B1+'A1')", 0, 5},
    {33, "=LET(A1,5,A1+R1C1)", "=LET(A1,5,A1+$A$1)", 105, -1},
    {34, "=LET(B1,2,B1)*R1C1", "=LET(B1,2,B1)*$A$1", 200, -1},
};

const ProbeRow kOut7[] = {
    {3, "=LET(x,5,x+R1C1)", "=LET(x,5,x+$A$1)", 105, -1},
    {4, "=R1C1+1", "=$A$1+1", 101, -1},
    {5, "=LET(RC,1,RC)+R1C1", nullptr, 0, -1},
    {6, "=LET(R1C1,3,R1C1)+R1C1", nullptr, 0, -1},
    {7, "=LET(R1C1,3,R1C1)", "=LET(R1C1,3,R1C1)", 3, -1},
    {8, "=SUM(R1:R2)", "=SUM(R1:R2)", 0, -1},
    {9, "=R1C1:R2C1", "=$A$1:$A$2", 0, 9},
    {10, "=SUM(R[1]:R[2])", "=SUM(11:12)", 1064972, -1},
    {11, "=ROWS(C[-1])", "=ROWS(D:D)", 1048576, -1},
    {12, "=COLUMNS(R[1])", "=COLUMNS(13:13)", 16384, -1},
    {13, "=RC[-10]", "=XEY13", 0, -1},
    {14, "=R[-5000000]C[-1]", nullptr, 0, -1},
    {15, "=Sheet1!R1C1", "=Sheet1!$A$1", 100, -1},
    {16, "=r1c1", "=$A$1", 100, -1},
    {17, "=rc[-1]", "=D17", 7, -1},
    {18, "=LET(B1,5,Sheet1!B1)", "=LET(B1,5,Sheet1!'B1')", 0, 5},
    {19, "=LET(B1,5,B1+A1:B2)", "=LET(B1,5,B1+'A1':'B2')", 0, 5},
    {20, "=LET(B1,5,B1+A)", "=LET(B1,5,B1+A)", 0, 5},
    {21, "=LET(B1,5,B1+AB)", "=LET(B1,5,B1+AB)", 0, 5},
    {22, "=LET(B1,5,B1+XFD1048577)", "=LET(B1,5,B1+XFD1048577)", 0, 5},
    {23, "=LET(R1C,2,R1C)", "=LET(R1C,2,R1C)", 2, -1},
    {24, "=LET(R1C,2,R1C)+R1C1", nullptr, 0, -1},
    {25, "=R1C1C", nullptr, 0, -1},
    {26, "=LET(B1,5,SUM(A:B))", "=LET(B1,5,SUM(A:(B)))", 0, 5},
    {27, "=LET(B1,5,SUM(1:2))", nullptr, 0, -1},
    {28, "=LET(B1,5,SUM(R1:R2))", "=LET(B1,5,SUM($1:$2))", 1120, -1},
    {29, "=LET(B1,5,B1+R1C1:R1C2)", "=LET(B1,5,B1+$A$1:$B$1)", 0, 9},
    {30, "=LET(_xlpm.B1,5,_xlpm.B1)", nullptr, 0, -1},
    {31, "=_xlnm.B2", "=_xlnm.B2", 0, 5},
    {32, "=LET(B1,5,B1+_xlnm.B2)", "=LET(B1,5,B1+_xlnm.B2)", 0, 5},
    {33, "=LET(B1,5,B1)+R[1]C[-3]", "=LET(B1,5,B1)+B34", 5, -1},
    {34, "=LAMBDA(R1C,R1C)(3)", "=LAMBDA(R1C,R1C)(3)", 3, -1},
    {35, "=LET(B1,5,B1+$A$1)", nullptr, 0, -1},
    {36, "=LET(B1,5,B1+R[-35]C[-3])", "=LET(B1,5,B1+B1)", 1005, -1},
    {37, "=LET(B1,5,B1*0+R[-36]C[-3])", "=LET(B1,5,B1*0+B1)", 1000, -1},
};

const ProbeRow kOut8New[] = {
    {3, "=LET(B1,5,B1+R1048577)", "=LET(B1,5,B1+R1048577)", 0, 5},
    {4, "=LET(B1,5,B1+C16385)", "=LET(B1,5,B1+'C16385')", 0, 5},
    {5, "=LET(B1,5,B1+R0)", "=LET(B1,5,B1+R0)", 0, 5},
    {6, "=LET(B1,5,B1+R4294967296)", "=LET(B1,5,B1+R4294967296)", 0, 5},
    {7, "=LET(B1,LAMBDA(x,x+1),B1(2))", "=LET(B1,LAMBDA(x,x+1),B1(2))", 3, -1},
    {8, "=LET(B1,5,B1 + R1C1)", "=LET(B1,5,B1 + $A$1)", 105, -1},
    {9, "= LET( B1 , 5 , B1+B2 )", "= LET( B1, 5, B1+'B2' )", 0, 5},
    {10, "=LET(B1,5,SUM(Sheet1!A1:B2))", "=LET(B1,5,SUM(Sheet1!'A1':'B2'))", 0, 5},
    {11, "=LET(B1,5,SUM(R1C1 :R2C1))", "=LET(B1,5,SUM($A$1:$A$2))", 120, -1},
    {12, "=RC[-1] + 1", "=D12 + 1", 8, -1},
    {13, "=SUM( RC[-1] , 1 )", "=SUM( D13, 1 )", 8, -1},
    {14, "=LAMBDA(B1,B1(1))(LAMBDA(x,x*2))", "=LAMBDA(B1,B1(1))(LAMBDA(x,x*2))", 2, -1},
    {15, "=LET(B1,5,R1C1:B1)", "=LET(B1,5,$A$1:B1)", 0, 3},
    {16, "=LET(B1,5,B1+A1:A)", "=LET(B1,5,B1+'A1':(A))", 0, 5},
    {17, "=LET(LE,5,LE:NA())", nullptr, 0, -1},
    {18, "=LET(B1,5,B1+Sheet1!R1C1:R2C1)", "=LET(B1,5,B1+Sheet1!$A$1:$A$2)", 105, -1},
};

std::string ProbeCell(std::uint32_t row, char column) {
  return std::string(1, column) + std::to_string(row);
}

// Builds the probe sheet for `rows`; `*status` receives each E entry's status.
template <std::size_t N>
void BuildProbeSheet(WorkbookGuard* wb, const ProbeRow (&rows)[N], bool with_b1, std::vector<fm_status_t>* status) {
  ASSERT_EQ(fm_workbook_create_default(&wb->handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb->handle, 0, 0, 0, 100.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb->handle, 0, 1, 0, 20.0), 0);
  if (with_b1) {
    ASSERT_EQ(fm_sheet_set_number(wb->handle, 0, 0, 1, 1000.0), 0);
  }
  for (const ProbeRow& r : rows) {
    const std::uint32_t row0 = r.row - 1U;
    ASSERT_EQ(fm_sheet_set_number(wb->handle, 0, row0, kColD, 7.0), 0);
    status->push_back(fm_sheet_set_formula(wb->handle, 0, row0, kColE, r.formula));
    if (status->back() != 0) {
      // A refusal carries the parser status of the parse error it records.
      fm_parse_error_t err = FM_STRUCT_INIT(fm_parse_error_t);
      ASSERT_EQ(fm_get_last_parse_error(&err), 0) << r.formula;
      EXPECT_EQ(status->back(), static_cast<fm_status_t>(formulon::parser::parse_error_status(
                                    static_cast<formulon::parser::ParseErrorCode>(err.code))))
          << r.formula;
    }
    const std::string e = ProbeCell(r.row, 'E');
    ASSERT_EQ(fm_sheet_set_formula(wb->handle, 0, row0, kColF, ("=IFERROR(ERROR.TYPE(" + e + "),-1)").c_str()), 0)
        << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_set_formula(wb->handle, 0, row0, kColG, ("=FORMULATEXT(" + e + ")").c_str()), 0)
        << fm_get_last_error_message();
  }
  ASSERT_EQ(fm_workbook_recalc(wb->handle), 0);
}

std::string FormulaAt(const fm_workbook_t* wb, std::uint32_t row0, std::uint32_t col) {
  const char* text = nullptr;
  EXPECT_EQ(fm_sheet_get_formula(wb, 0, row0, col, &text), 0) << fm_get_last_error_message();
  return text != nullptr ? text : "<null>";
}

fm_value_t ValueAt(const fm_workbook_t* wb, std::uint32_t row0, std::uint32_t col) {
  fm_value_t value{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row0, col, &value), 0) << fm_get_last_error_message();
  return value;
}

// Checks one row against Excel: refusal, or the formula-bar spelling, the
// FORMULATEXT result, the value and its ERROR.TYPE.
void ExpectRowMatches(const fm_workbook_t* wb, const ProbeRow& r, fm_status_t status) {
  SCOPED_TRACE(std::string("row ") + std::to_string(r.row) + " " + r.formula);
  const std::uint32_t row0 = r.row - 1U;
  const fm_value_t f = ValueAt(wb, row0, kColF);
  ASSERT_EQ(f.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(f.u.number, static_cast<double>(r.error_type));
  if (r.stored == nullptr) {
    EXPECT_NE(status, 0);
    EXPECT_EQ(FormulaAt(wb, row0, kColE), "");
    EXPECT_EQ(ValueAt(wb, row0, kColE).kind, FM_VALUE_KIND_BLANK);
    return;
  }
  ASSERT_EQ(status, 0) << fm_get_last_error_message();
  EXPECT_EQ(FormulaAt(wb, row0, kColE), r.stored);
  const fm_value_t g = ValueAt(wb, row0, kColG);
  ASSERT_EQ(g.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(g.u.text, r.stored);
  const fm_value_t e = ValueAt(wb, row0, kColE);
  if (r.error_type == -1) {
    ASSERT_EQ(e.kind, FM_VALUE_KIND_NUMBER);
    EXPECT_EQ(e.u.number, r.value);
  } else {
    EXPECT_EQ(e.kind, FM_VALUE_KIND_ERROR);
  }
}

template <std::size_t N>
void ExpectSheetMatches(const ProbeRow (&rows)[N], bool with_b1, const std::set<std::uint32_t>& skipped) {
  WorkbookGuard wb;
  std::vector<fm_status_t> status;
  ASSERT_NO_FATAL_FAILURE(BuildProbeSheet(&wb, rows, with_b1, &status));
  ASSERT_EQ(status.size(), N);
  for (std::size_t i = 0; i < N; ++i) {
    if (skipped.count(rows[i].row) == 0U) {
      ExpectRowMatches(wb.handle, rows[i], status[i]);
    }
  }
}

// Formulon's documented spelling for a row where it parts from Excel: the
// formula-bar text, and #NAME? for the value.
void ExpectOut7Row(std::uint32_t row, const char* shown) {
  WorkbookGuard wb;
  std::vector<fm_status_t> status;
  ASSERT_NO_FATAL_FAILURE(BuildProbeSheet(&wb, kOut7, true, &status));
  std::size_t i = 0;
  while (kOut7[i].row != row) {
    ++i;
  }
  ASSERT_EQ(status[i], 0) << fm_get_last_error_message();
  EXPECT_EQ(FormulaAt(wb.handle, row - 1U, kColE), shown);
  const fm_value_t g = ValueAt(wb.handle, row - 1U, kColG);
  ASSERT_EQ(g.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(g.u.text, shown);
  const fm_value_t f = ValueAt(wb.handle, row - 1U, kColF);
  ASSERT_EQ(f.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(f.u.number, 5.0);  // #NAME?
}

}  // namespace

TEST(FormulonCApiEntryR1C1, Out6SheetMatchesExcel) {
  ExpectSheetMatches(kOut6, false, {});
}

TEST(FormulonCApiEntryR1C1, Out7SheetMatchesExcel) {
  ExpectSheetMatches(kOut7, true, {25U, 31U, 32U});
}

TEST(FormulonCApiEntryR1C1, Out8NewSheetMatchesExcel) {
  ExpectSheetMatches(kOut8New, true, {});
}

// Excel refuses `=R1C1C`; it is neither an R1C1 reference nor a cell shape,
// so Formulon keeps it as an undefined name.
TEST(FormulonCApiEntryR1C1, Out7R1C1CStaysAName) {
  ExpectOut7Row(25U, "=R1C1C");
}

// Typed `_xlnm.B2` shares the stored form of the name `B2`, so it shows as
// Excel shows both after a save and reload.
TEST(FormulonCApiEntryR1C1, Out7TypedXlnmNameShowsQuoted) {
  ExpectOut7Row(31U, "='B2'");
}

TEST(FormulonCApiEntryR1C1, Out7LetWithTypedXlnmNameShowsQuoted) {
  ExpectOut7Row(32U, "=LET(B1,5,B1+'B2')");
}

// A table named `RC` keeps `RC[Qty]` a structured reference. Excel refuses
// that display name, so only a loaded file carries one.
TEST(FormulonCApiEntryR1C1, BracketedEndpointNamingATableStaysStructured) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 0, "Qty"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0, 1, "X"), 0);
  const fm_cell_rect_t rect{0, 0, 2, 1};
  const char* const columns[] = {"Qty", "X"};
  ASSERT_EQ(AddTable(wb.handle, 0, &rect, "Tbl", columns, 2, nullptr, 1, 0, ignored_index()), 0)
      << fm_get_last_error_message();
  std::vector<formulon::TableMetadata> tables = wb.handle->workbook().tables();
  tables[0].name = "RC";
  tables[0].display_name = "RC";
  wb.handle->workbook().set_tables(std::move(tables));
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 4, 6, "=SUM(RC[Qty])"), 0) << fm_get_last_error_message();
  EXPECT_EQ(FormulaAt(wb.handle, 4, 6), "=SUM(RC[Qty])");
}

// An array formula reads R1C1 at its top-left cell.
TEST(FormulonCApiEntryR1C1, ArrayFormulaReadsR1C1AtItsTopLeft) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t rect{2, 4, 3, 4};
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, &rect, "=RC[-1]"), 0) << fm_get_last_error_message();
  fm_array_formula_t info = FM_STRUCT_INIT(fm_array_formula_t);
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 3, 4, &info), 0);
  ASSERT_EQ(info.engaged, 1);
  EXPECT_STREQ(info.formula, "=D3");
}

// Typed `_xlpm.` before a cell shape is refused, in a cell and a defined name.
TEST(FormulonCApiEntryR1C1, TypedParamPrefixOnACellShapeIsRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=LET(_xlpm.B1,5,_xlpm.B1)"), kParserUnexpectedToken);
  EXPECT_EQ(SetDefinedName(wb.handle, "Rate", "LET(_xlpm.B1,5,_xlpm.B1)"), kParserUnexpectedToken);
  EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=LET(_xlpm.x,5,_xlpm.x)"), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 0, "=\"_xlpm.B1\""), 0) << fm_get_last_error_message();
}

// The fallback output passes the whole entry check: whitespace and arity
// rules Excel applies to A1 entry still refuse it.
TEST(FormulonCApiEntryR1C1, FallbackKeepsTheEntryRefusals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const char* formula : {"=SUM (RC[-1])", "=ROWS()+RC[-1]", "=LET(B1,5,SUBTOTAL(9,))", "=LET(B1,5,B1\t+1)"}) {
    SCOPED_TRACE(formula);
    EXPECT_EQ(fm_sheet_set_formula(wb.handle, 0, 4, 4, formula), kParserUnexpectedToken);
  }
}
