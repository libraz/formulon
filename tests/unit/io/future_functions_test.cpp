//
// Storage-prefix classification (`io/future_functions.h`).
//
// The expectations below are drawn from the `<f>` text of the
// Excel-authored workbooks shipped under `tests/fixtures/excel/` and
// `tests/oracle/external/ironcalc/fixtures/`: whatever spelling those
// files carry for a function is the spelling this classifier must
// reproduce.

#include "io/future_functions.h"

#include <cstddef>
#include <fstream>
#include <sstream>
#include <string>

#include "gtest/gtest.h"
#include "parser/ast_format.h"

namespace formulon {
namespace io {
namespace {

using parser::StoragePrefixKind;

// ---------------------------------------------------------------------------
// Functions native since Excel 2007 or earlier keep their bare spelling.
// The Analysis ToolPak group Excel 2007 absorbed is the interesting part:
// none of it has an entry in `io/xlsb/func_id_table.h`, so a classifier
// keyed on that table would prefix all of it and real Excel would render
// every one as #NAME?.
// ---------------------------------------------------------------------------

TEST(ClassifyStoragePrefix, AnalysisToolPakFunctionsAreNotPrefixed) {
  for (const char* name :
       {"MROUND",     "YEARFRAC",   "WEEKNUM",     "ISEVEN",    "ISODD",       "GCD",      "LCM",       "QUOTIENT",
        "SQRTPI",     "FACTDOUBLE", "MULTINOMIAL", "SERIESSUM", "RANDBETWEEN", "EDATE",    "EOMONTH",   "NETWORKDAYS",
        "WORKDAY",    "DELTA",      "GESTEP",      "CONVERT",   "COMPLEX",     "BESSELI",  "BIN2HEX",   "DEC2BIN",
        "IMPOWER",    "IMSUM",      "ERF",         "ERFC",      "CUMIPMT",     "CUMPRINC", "XIRR",      "XNPV",
        "EFFECT",     "NOMINAL",    "DOLLARDE",    "DOLLARFR",  "ACCRINT",     "COUPNUM",  "ODDFPRICE", "PRICEMAT",
        "TBILLYIELD", "YIELDDISC",  "AMORLINC",    "DURATION"}) {
    SCOPED_TRACE(name);
    EXPECT_EQ(classify_storage_prefix(name), StoragePrefixKind::None);
    EXPECT_EQ(storage_function_name(name), std::string(name));
  }
}

TEST(ClassifyStoragePrefix, ClassicFunctionsAreNotPrefixed) {
  for (const char* name : {"SUM", "VLOOKUP", "IF", "INDEX", "MATCH", "IFERROR", "SUMIFS", "COUNTIFS", "AVERAGEIF",
                           "TRANSPOSE", "ASC", "JIS", "LENB", "PHONETIC", "GETPIVOTDATA", "CUBEVALUE", "RTD"}) {
    SCOPED_TRACE(name);
    EXPECT_EQ(classify_storage_prefix(name), StoragePrefixKind::None);
  }
}

// ISO.CEILING, NETWORKDAYS.INTL and WORKDAY.INTL arrived with Excel 2010 yet
// Excel stores them bare, so the classification cannot be derived from the
// introduction date alone. The `.xlsb` encodes all three by hidden name.
TEST(ClassifyStoragePrefix, BareExcel2010FunctionsAreNotPrefixed) {
  for (const char* name : {"ISO.CEILING", "NETWORKDAYS.INTL", "WORKDAY.INTL"}) {
    SCOPED_TRACE(name);
    EXPECT_EQ(classify_storage_prefix(name), StoragePrefixKind::None);
    EXPECT_EQ(storage_function_name(name), name);
    EXPECT_TRUE(xlsb_uses_hidden_name(name));
  }
}

// ---------------------------------------------------------------------------
// Post-2007 functions carry `_xlfn.`; the two worksheet-only dynamic-array
// forms carry `_xlfn._xlws.`.
// ---------------------------------------------------------------------------

TEST(ClassifyStoragePrefix, FutureFunctionsCarryXlfn) {
  for (const char* name : {"XLOOKUP",      "XMATCH",      "TEXTJOIN",    "CONCAT",    "IFS",        "SEQUENCE",
                           "LET",          "LAMBDA",      "UNIQUE",      "SORTBY",    "STDEV.P",    "PERCENTILE.INC",
                           "CEILING.MATH", "NORM.S.DIST", "FORMULATEXT", "ISFORMULA", "ISOWEEKNUM", "UNICODE",
                           "XOR",          "RRI",         "PDURATION",   "AGGREGATE", "STDEV.S",    "GROUPBY",
                           "REGEXTEST",    "TEXTSPLIT",   "HSTACK",      "TAKE"}) {
    SCOPED_TRACE(name);
    EXPECT_EQ(classify_storage_prefix(name), StoragePrefixKind::Xlfn);
    EXPECT_EQ(storage_function_name(name), std::string("_xlfn.") + name);
  }
}

TEST(ClassifyStoragePrefix, WorksheetOnlyDynamicArrayFunctionsCarryXlfnXlws) {
  EXPECT_EQ(classify_storage_prefix("FILTER"), StoragePrefixKind::XlfnXlws);
  EXPECT_EQ(storage_function_name("FILTER"), "_xlfn._xlws.FILTER");
  EXPECT_EQ(classify_storage_prefix("SORT"), StoragePrefixKind::XlfnXlws);
  EXPECT_EQ(storage_function_name("SORT"), "_xlfn._xlws.SORT");
}

TEST(ClassifyStoragePrefix, PythonCallCarriesXlfnXlws) {
  EXPECT_EQ(classify_storage_prefix("PY"), StoragePrefixKind::XlfnXlws);
  EXPECT_EQ(storage_function_name("PY"), "_xlfn._xlws.PY");
}

// SORTBY / UNIQUE sit next to FILTER / SORT in the same dynamic-array
// family but Excel writes them with the single prefix.
TEST(ClassifyStoragePrefix, DynamicArraySiblingsDoNotCarryXlws) {
  EXPECT_EQ(storage_function_name("SORTBY"), "_xlfn.SORTBY");
  EXPECT_EQ(storage_function_name("UNIQUE"), "_xlfn.UNIQUE");
  EXPECT_EQ(storage_function_name("SEQUENCE"), "_xlfn.SEQUENCE");
}

// ---------------------------------------------------------------------------
// Lookup shape: case-insensitive, whole-name, and closed under the
// unknown-name default.
// ---------------------------------------------------------------------------

TEST(ClassifyStoragePrefix, LookupIsCaseInsensitiveAndUpperCasesTheStoredName) {
  EXPECT_EQ(classify_storage_prefix("xlookup"), StoragePrefixKind::Xlfn);
  EXPECT_EQ(storage_function_name("xLoOkUp"), "_xlfn.XLOOKUP");
  EXPECT_EQ(storage_function_name("mround"), "MROUND");
}

// A dotted family name must match as a whole entry: `NORM.DIST` is listed
// but must not be found as a substring of `NORM.S.DIST`, and a truncation
// of a listed name must not match at all.
TEST(ClassifyStoragePrefix, PartialNamesDoNotMatch) {
  EXPECT_EQ(classify_storage_prefix("NORM"), StoragePrefixKind::None);
  EXPECT_EQ(classify_storage_prefix("DIST"), StoragePrefixKind::None);
  EXPECT_EQ(classify_storage_prefix("XLOOKU"), StoragePrefixKind::None);
  EXPECT_EQ(classify_storage_prefix("LOOKUP"), StoragePrefixKind::None);
  EXPECT_EQ(classify_storage_prefix(""), StoragePrefixKind::None);
}

TEST(ClassifyStoragePrefix, UnknownNamesAreStoredAsSpelled) {
  EXPECT_EQ(classify_storage_prefix("MY.UDF"), StoragePrefixKind::None);
  EXPECT_EQ(storage_function_name("MY.UDF"), "MY.UDF");
}

// Every name in the committed map of stored spellings Excel 365 writes to `<f>` classifies to that spelling.
TEST(ClassifyStoragePrefix, MatchesTheSpellingExcelStores) {
  std::ifstream in(FORMULON_FIXTURES_DIR "/excel/function_storage_spelling.txt");
  ASSERT_TRUE(in.good());
  std::size_t checked = 0U;
  std::string line;
  while (std::getline(in, line)) {
    if (line.empty() || line[0] == '#') {
      continue;
    }
    std::istringstream fields(line);
    std::string name;
    std::string spelling;
    ASSERT_TRUE(static_cast<bool>(fields >> name >> spelling)) << line;
    StoragePrefixKind want = StoragePrefixKind::None;
    if (spelling == "_xlfn.") {
      want = StoragePrefixKind::Xlfn;
    } else if (spelling == "_xlfn._xlws.") {
      want = StoragePrefixKind::XlfnXlws;
    } else {
      ASSERT_EQ(spelling, "bare") << line;
    }
    EXPECT_EQ(classify_storage_prefix(name), want) << name;
    ++checked;
  }
  EXPECT_GT(checked, 500U);
}

}  // namespace
}  // namespace io
}  // namespace formulon
