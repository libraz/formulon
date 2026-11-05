//
// Tests for built-in name resolution and the qualified-built-in-call entry
// rule. Every formula below was entered into Mac Excel 365 (ja-JP); the
// rejected set is what Excel refused at entry, the accepted set is what it
// stored.

#include "eval/builtin_names.h"

#include <cstdint>
#include <set>
#include <string>
#include <string_view>

#include "eval/datetime_lazy.h"
#include "eval/financial_lazy.h"
#include "eval/function_registry.h"
#include "eval/special_forms_catalog.h"
#include "eval/tree_walker.h"
#include "eval/tree_walker_lazy_table.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/formula_prefix.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace eval {
namespace {

// Mirrors formula entry: storage prefixes are stripped before the parse.
const parser::AstNode* ParseEntry(std::string_view formula, Arena& arena, std::string& storage) {
  storage = parser::strip_storage_prefixes(formula);
  return parser::parse_strict(storage, arena);
}

TEST(BuiltinNames, ResolvesEverySource) {
  EXPECT_STREQ(resolve_builtin_function_name("sum"), "SUM");
  EXPECT_STREQ(resolve_builtin_function_name("XLOOKUP"), "XLOOKUP");
  EXPECT_STREQ(resolve_builtin_function_name("Let"), "LET");
  EXPECT_STREQ(resolve_builtin_function_name("LAMBDA"), "LAMBDA");
  EXPECT_EQ(resolve_builtin_function_name("NOSUCH"), nullptr);
  EXPECT_EQ(resolve_builtin_function_name("EUROCONVERT"), nullptr);
}

TEST(BuiltinNames, QualifiedBuiltinCallIsFound) {
  for (const char* src :
       {"Sheet1!SUM(1)", "'My Sheet'!SUM(1)", "[0]!SUM(1)", "Sheet1!sum(1)", "Sheet1!XLOOKUP(1,A1:A5,A1:A5)",
        "Sheet1!_xlfn.XLOOKUP(1,A1:A5,A1:A5)", "Sheet1!LET(a,1,a)", "Sheet1!LAMBDA(x,x)(1)", "Sheet1!IF(1,2)",
        "Sheet1!NOW()", "Sheet1!ISOMITTED(1)", "SUM(Sheet1!SUM(1))", "Sheet1!SUM(1)+1"}) {
    Arena a;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, a, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_NE(find_qualified_builtin_call(*root), nullptr) << src;
  }
}

TEST(BuiltinNames, RejectedSpellingsTheParserAlreadyRefuses) {
  for (const char* src : {"Sheet1!TRUE(1)", "Sheet1!SUM (1)", "Sheet1:Sheet2!SUM(1)"}) {
    Arena a;
    std::string storage;
    EXPECT_EQ(ParseEntry(src, a, storage), nullptr) << src;
  }
}

TEST(BuiltinNames, AcceptedQualifiedFormsAreNotFlagged) {
  for (const char* src :
       {"Sheet1!MyFn(1)", "MyFn(1)", "[0]!WFn(1)", "WFn(1)", "'My Sheet'!QFn(1)", "Sheet1!NOSUCH(1)", "Sheet1!SUMX(1)",
        "Sheet1!EUROCONVERT(1,\"FRF\",\"EUR\")", "Sheet1!SUM", "Sheet1!SVal", "SUM(1)", "LAMBDA(x,x)(1)"}) {
    Arena a;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, a, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_EQ(find_qualified_builtin_call(*root), nullptr) << src;
  }
}

TEST(BuiltinNames, FormulaEntryGateRejectsQualifiedBuiltinsOnly) {
  for (const char* src : {"Sheet1!SUM(1)", "[0]!LET(a,1,a)", "SUM(Sheet1!SUM(1))"}) {
    Arena arena;
    std::string storage = parser::strip_storage_prefixes(src);
    EXPECT_EQ(parse_formula_entry(storage, arena), nullptr) << src;
  }

  for (const char* src : {"Sheet1!MyFn(1)", "Sheet1!NOSUCH(1)", "SUM(1)", "A1+1"}) {
    Arena arena;
    std::string storage = parser::strip_storage_prefixes(src);
    EXPECT_NE(parse_formula_entry(storage, arena), nullptr) << src;
  }

  Arena malformed_arena;
  EXPECT_EQ(parse_formula_entry("SUM(", malformed_arena), nullptr);
}

// Mac Excel 365 at entry: what a reference-only parameter (SUMIF's range)
// accepts and refuses.
TEST(BuiltinNames, ReferenceOnlyParameterTakesReferenceExpressionsOnly) {
  for (const char* arg : {"A1",
                          "(A1)",
                          "(A1,A2)",
                          "A1:A2 A1",
                          "A:A",
                          "1:1",
                          "Sheet2!A1",
                          "Sheet2:Sheet3!A1",
                          "[Book9.xlsx]Sheet1!A1",
                          "A1#",
                          "@A1",
                          "#REF!",
                          "Nm",
                          "t[c]",
                          "INDEX(A1:A2,1):A2",
                          "A1:INDEX(A1:A2,2)",
                          "LAMBDA(x,x)(A1)",
                          "FOO(A1)",
                          "INDEX({1,2},1)",
                          "INDIRECT(1)",
                          "OFFSET(A1,0,0)",
                          "IF(1,1,1)",
                          "IF(1,A1)",
                          "CHOOSE(1,1)",
                          "XLOOKUP(1,{1},{1})",
                          "IFS(1,1)",
                          "SWITCH(1,1,1)",
                          "LET(x,1,x)",
                          "TRIMRANGE(1)",
                          "SINGLE(1)",
                          "TAKE(A1:A2,1)",
                          "DROP(A1:A2,1)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(std::string("SUMIF(") + arg + ",1)", arena, storage);
    ASSERT_NE(root, nullptr) << arg;
    EXPECT_EQ(find_non_reference_argument(*root), nullptr) << arg;
  }
  for (const char* arg : {"",
                          "1",
                          "\"A1\"",
                          "TRUE",
                          "{1}",
                          "#N/A",
                          "1%",
                          "(1)",
                          "(A1,1)",
                          "A1+0",
                          "-A1",
                          "+A1",
                          "A1&\"\"",
                          "SUM(A1)",
                          "IFERROR(A1,A1)",
                          "IFNA(A1,A1)",
                          "HYPERLINK(A1)",
                          "HSTACK(A1)",
                          "VSTACK(A1)",
                          "EXPAND(A1,1)",
                          "CHOOSEROWS(A1:A2,1)",
                          "FILTER(A1:A2,1)",
                          "SORT(A1:A2)",
                          "UNIQUE(A1:A2)",
                          "XMATCH(1,A1)",
                          "ROWS(A1)",
                          "CELL(\"row\",A1)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(std::string("SUMIF(") + arg + ",1)", arena, storage);
    ASSERT_NE(root, nullptr) << arg;
    EXPECT_NE(find_non_reference_argument(*root), nullptr) << arg;
  }
}

// Each reference-only parameter Excel has, given an empty argument, and the
// parameters beside it, which take any value.
TEST(BuiltinNames, ReferenceOnlyParametersPerFunction) {
  for (const char* src : {"AREAS(1)",           "COUNTBLANK(1)",
                          "FORMULATEXT(1)",     "ISFORMULA(1)",
                          "PHONETIC(1)",        "ROW(1)",
                          "COLUMN(1)",          "SUMIF(,1)",
                          "SUMIF(A1,1,)",       "AVERAGEIF(,1)",
                          "AVERAGEIF(A1,1,1)",  "COUNTIF(,1)",
                          "COUNTIFS(A1,1,,1)",  "SUMIFS(,A1,1)",
                          "SUMIFS(A1,,1)",      "SUMIFS(A1,A1,1,1,1)",
                          "AVERAGEIFS(A1,1,1)", "MAXIFS(,A1,1)",
                          "MINIFS(A1,A1,1,,1)", "DSUM(,1,1)",
                          "DGET(1,1,1)",        "DCOUNTA(,1,1)",
                          "CELL(\"row\",)",     "OFFSET(,1,1)",
                          "RANK(1,)",           "RANK.EQ(1,1)",
                          "RANK.AVG(1,,1)",     "SUBTOTAL(9,)",
                          "SUBTOTAL(9,A1,)",    "SUBTOTAL(9,A1,A1,1)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, arena, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_NE(find_non_reference_argument(*root), nullptr) << src;
  }
  for (const char* src : {"SUM(,)",
                          "IF(A1,,1)",
                          "COUNT(1,,1)",
                          "INDEX(A1,,)",
                          "SUMIF(A1,)",
                          "SUMIF(A1,1)",
                          "COUNTIFS(A1,,A1,)",
                          "SUMIFS(A1,A1,)",
                          "DSUM(A1,,)",
                          "CELL(,A1)",
                          "OFFSET(A1,,)",
                          "RANK(,A1,)",
                          "SUBTOTAL(,A1)",
                          "SUBTOTAL(1,A1)",
                          "AGGREGATE(9,4,)",
                          "AGGREGATE(9,4,1)",
                          "AGGREGATE(14,4,,1)",
                          "ROW()",
                          "COLUMN()",
                          "SUM(SUMIF(A1,1),1)",
                          "ISREF(1)",
                          "ROWS(1)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, arena, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_EQ(find_non_reference_argument(*root), nullptr) << src;
  }
  // Found at any depth, including inside LET and LAMBDA bodies.
  for (const char* src : {"1+SUBTOTAL(9,)", "LET(x,1,SUMIF(1,x))", "IF(A1,COUNTIF(1,1),0)"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(src, arena, storage);
    ASSERT_NE(root, nullptr) << src;
    EXPECT_NE(find_non_reference_argument(*root), nullptr) << src;
  }
}

struct ExpectedArity {
  const char* name;
  std::uint32_t min_arity;
  std::uint32_t max_arity;
};

// Written argument counts Excel 365 accepts at entry, measured per built-in
// with every count up to 21; open lists are kVariadic. Names Mac Excel does
// not know (COPILOT ENCODEURL FILTERXML JIS WEBSERVICE) and PY keep their
// registered bounds; LET and LAMBDA carry Excel's documented limits.
constexpr ExpectedArity kExpectedArity[] = {
    {"ABS", 1U, 1U},
    {"ACCRINT", 6U, 8U},
    {"ACCRINTM", 4U, 5U},
    {"ACOS", 1U, 1U},
    {"ACOSH", 1U, 1U},
    {"ACOT", 1U, 1U},
    {"ACOTH", 1U, 1U},
    {"ADDRESS", 2U, 5U},
    {"AGGREGATE", 3U, kVariadic},
    {"AMORDEGRC", 6U, 7U},
    {"AMORLINC", 6U, 7U},
    {"ANCHORARRAY", 1U, 1U},
    {"AND", 1U, kVariadic},
    {"ARABIC", 1U, 1U},
    {"AREAS", 1U, 1U},
    {"ARRAYTOTEXT", 1U, 2U},
    {"ASC", 1U, 1U},
    {"ASIN", 1U, 1U},
    {"ASINH", 1U, 1U},
    {"ATAN", 1U, 1U},
    {"ATAN2", 2U, 2U},
    {"ATANH", 1U, 1U},
    {"AVEDEV", 1U, kVariadic},
    {"AVERAGE", 1U, kVariadic},
    {"AVERAGEA", 1U, kVariadic},
    {"AVERAGEIF", 2U, 3U},
    {"AVERAGEIFS", 3U, kVariadic},
    {"BAHTTEXT", 1U, 1U},
    {"BASE", 2U, 3U},
    {"BESSELI", 2U, 2U},
    {"BESSELJ", 2U, 2U},
    {"BESSELK", 2U, 2U},
    {"BESSELY", 2U, 2U},
    {"BETA.DIST", 4U, 6U},
    {"BETA.INV", 3U, 5U},
    {"BETADIST", 3U, 5U},
    {"BETAINV", 3U, 5U},
    {"BIN2DEC", 1U, 1U},
    {"BIN2HEX", 1U, 2U},
    {"BIN2OCT", 1U, 2U},
    {"BINOM.DIST", 4U, 4U},
    {"BINOM.DIST.RANGE", 3U, 4U},
    {"BINOM.INV", 3U, 3U},
    {"BINOMDIST", 4U, 4U},
    {"BITAND", 2U, 2U},
    {"BITLSHIFT", 2U, 2U},
    {"BITOR", 2U, 2U},
    {"BITRSHIFT", 2U, 2U},
    {"BITXOR", 2U, 2U},
    {"BYCOL", 1U, 2U},
    {"BYROW", 1U, 2U},
    {"CEILING", 2U, 2U},
    {"CEILING.MATH", 1U, 3U},
    {"CEILING.PRECISE", 1U, 2U},
    {"CELL", 1U, 2U},
    {"CHAR", 1U, 1U},
    {"CHIDIST", 2U, 2U},
    {"CHIINV", 2U, 2U},
    {"CHISQ.DIST", 3U, 3U},
    {"CHISQ.DIST.RT", 2U, 2U},
    {"CHISQ.INV", 2U, 2U},
    {"CHISQ.INV.RT", 2U, 2U},
    {"CHISQ.TEST", 2U, 2U},
    {"CHITEST", 2U, 2U},
    {"CHOOSE", 2U, kVariadic},
    {"CHOOSECOLS", 2U, kVariadic},
    {"CHOOSEROWS", 2U, kVariadic},
    {"CLEAN", 1U, 1U},
    {"CODE", 1U, 1U},
    {"COLUMN", 0U, 1U},
    {"COLUMNS", 1U, 1U},
    {"COMBIN", 2U, 2U},
    {"COMBINA", 2U, 2U},
    {"COMPLEX", 2U, 3U},
    {"CONCAT", 1U, kVariadic},
    {"CONCATENATE", 1U, kVariadic},
    {"CONFIDENCE", 3U, 3U},
    {"CONFIDENCE.NORM", 3U, 3U},
    {"CONFIDENCE.T", 3U, 3U},
    {"CONVERT", 3U, 3U},
    {"COPILOT", 1U, kVariadic},
    {"CORREL", 2U, 2U},
    {"COS", 1U, 1U},
    {"COSH", 1U, 1U},
    {"COT", 1U, 1U},
    {"COTH", 1U, 1U},
    {"COUNT", 1U, kVariadic},
    {"COUNTA", 1U, kVariadic},
    {"COUNTBLANK", 1U, 1U},
    {"COUNTIF", 2U, 2U},
    {"COUNTIFS", 2U, kVariadic},
    {"COUPDAYBS", 3U, 4U},
    {"COUPDAYS", 3U, 4U},
    {"COUPDAYSNC", 3U, 4U},
    {"COUPNCD", 3U, 4U},
    {"COUPNUM", 3U, 4U},
    {"COUPPCD", 3U, 4U},
    {"COVAR", 2U, 2U},
    {"COVARIANCE.P", 2U, 2U},
    {"COVARIANCE.S", 2U, 2U},
    {"CRITBINOM", 3U, 3U},
    {"CSC", 1U, 1U},
    {"CSCH", 1U, 1U},
    {"CUBEKPIMEMBER", 3U, 4U},
    {"CUBEMEMBER", 2U, 3U},
    {"CUBEMEMBERPROPERTY", 3U, 3U},
    {"CUBERANKEDMEMBER", 3U, 4U},
    {"CUBESET", 2U, 5U},
    {"CUBESETCOUNT", 1U, 1U},
    {"CUBEVALUE", 1U, kVariadic},
    {"CUMIPMT", 6U, 6U},
    {"CUMPRINC", 6U, 6U},
    {"DATE", 3U, 3U},
    {"DATEDIF", 3U, 3U},
    {"DATEVALUE", 1U, 1U},
    {"DAVERAGE", 3U, 3U},
    {"DAY", 1U, 1U},
    {"DAYS", 2U, 2U},
    {"DAYS360", 2U, 3U},
    {"DB", 4U, 5U},
    {"DBCS", 1U, 1U},
    {"DCOUNT", 3U, 3U},
    {"DCOUNTA", 3U, 3U},
    {"DDB", 4U, 5U},
    {"DEC2BIN", 1U, 2U},
    {"DEC2HEX", 1U, 2U},
    {"DEC2OCT", 1U, 2U},
    {"DECIMAL", 2U, 2U},
    {"DEGREES", 1U, 1U},
    {"DELTA", 1U, 2U},
    {"DETECTLANGUAGE", 1U, 1U},
    {"DEVSQ", 1U, kVariadic},
    {"DGET", 3U, 3U},
    {"DISC", 4U, 5U},
    {"DMAX", 3U, 3U},
    {"DMIN", 3U, 3U},
    {"DOLLAR", 1U, 2U},
    {"DOLLARDE", 2U, 2U},
    {"DOLLARFR", 2U, 2U},
    {"DPRODUCT", 3U, 3U},
    {"DROP", 2U, 3U},
    {"DSTDEV", 3U, 3U},
    {"DSTDEVP", 3U, 3U},
    {"DSUM", 3U, 3U},
    {"DURATION", 5U, 6U},
    {"DVAR", 3U, 3U},
    {"DVARP", 3U, 3U},
    {"EDATE", 2U, 2U},
    {"EFFECT", 2U, 2U},
    {"ENCODEURL", 1U, 1U},
    {"EOMONTH", 2U, 2U},
    {"ERF", 1U, 2U},
    {"ERF.PRECISE", 1U, 1U},
    {"ERFC", 1U, 1U},
    {"ERFC.PRECISE", 1U, 1U},
    {"ERROR.TYPE", 1U, 1U},
    {"EVEN", 1U, 1U},
    {"EXACT", 2U, 2U},
    {"EXP", 1U, 1U},
    {"EXPAND", 2U, 4U},
    {"EXPON.DIST", 3U, 3U},
    {"EXPONDIST", 3U, 3U},
    {"F.DIST", 4U, 4U},
    {"F.DIST.RT", 3U, 3U},
    {"F.INV", 3U, 3U},
    {"F.INV.RT", 3U, 3U},
    {"F.TEST", 2U, 2U},
    {"FACT", 1U, 1U},
    {"FACTDOUBLE", 1U, 1U},
    {"FALSE", 0U, 0U},
    {"FDIST", 3U, 3U},
    {"FILTER", 2U, 3U},
    {"FILTERXML", 2U, 2U},
    {"FIND", 2U, 3U},
    {"FINDB", 2U, 3U},
    {"FINV", 3U, 3U},
    {"FISHER", 1U, 1U},
    {"FISHERINV", 1U, 1U},
    {"FIXED", 1U, 3U},
    {"FLOOR", 2U, 2U},
    {"FLOOR.MATH", 1U, 3U},
    {"FLOOR.PRECISE", 1U, 2U},
    {"FORECAST", 3U, 3U},
    {"FORECAST.ETS", 3U, 6U},
    {"FORECAST.ETS.CONFINT", 3U, 7U},
    {"FORECAST.ETS.SEASONALITY", 2U, 4U},
    {"FORECAST.ETS.STAT", 3U, 6U},
    {"FORECAST.LINEAR", 3U, 3U},
    {"FORMULATEXT", 1U, 1U},
    {"FREQUENCY", 2U, 2U},
    {"FTEST", 2U, 2U},
    {"FV", 3U, 5U},
    {"FVSCHEDULE", 2U, 2U},
    {"GAMMA", 1U, 1U},
    {"GAMMA.DIST", 4U, 4U},
    {"GAMMA.INV", 3U, 3U},
    {"GAMMADIST", 4U, 4U},
    {"GAMMAINV", 3U, 3U},
    {"GAMMALN", 1U, 1U},
    {"GAMMALN.PRECISE", 1U, 1U},
    {"GAUSS", 1U, 1U},
    {"GCD", 1U, kVariadic},
    {"GEOMEAN", 1U, kVariadic},
    {"GESTEP", 1U, 2U},
    {"GETPIVOTDATA", 2U, kVariadic},
    {"GROUPBY", 3U, 8U},
    {"GROWTH", 1U, 4U},
    {"HARMEAN", 1U, kVariadic},
    {"HEX2BIN", 1U, 2U},
    {"HEX2DEC", 1U, 1U},
    {"HEX2OCT", 1U, 2U},
    {"HLOOKUP", 3U, 4U},
    {"HOUR", 1U, 1U},
    {"HSTACK", 1U, kVariadic},
    {"HYPERLINK", 1U, 2U},
    {"HYPGEOM.DIST", 5U, 5U},
    {"HYPGEOMDIST", 4U, 4U},
    {"IF", 2U, 3U},
    {"IFERROR", 2U, 2U},
    {"IFNA", 2U, 2U},
    {"IFS", 2U, kVariadic},
    {"IMABS", 1U, 1U},
    {"IMAGE", 1U, 5U},
    {"IMAGINARY", 1U, 1U},
    {"IMARGUMENT", 1U, 1U},
    {"IMCONJUGATE", 1U, 1U},
    {"IMCOS", 1U, 1U},
    {"IMCOSH", 1U, 1U},
    {"IMCOT", 1U, 1U},
    {"IMCSC", 1U, 1U},
    {"IMCSCH", 1U, 1U},
    {"IMDIV", 2U, 2U},
    {"IMEXP", 1U, 1U},
    {"IMLN", 1U, 1U},
    {"IMLOG10", 1U, 1U},
    {"IMLOG2", 1U, 1U},
    {"IMPOWER", 2U, 2U},
    {"IMPRODUCT", 1U, kVariadic},
    {"IMREAL", 1U, 1U},
    {"IMSEC", 1U, 1U},
    {"IMSECH", 1U, 1U},
    {"IMSIN", 1U, 1U},
    {"IMSINH", 1U, 1U},
    {"IMSQRT", 1U, 1U},
    {"IMSUB", 2U, 2U},
    {"IMSUM", 1U, kVariadic},
    {"IMTAN", 1U, 1U},
    {"INDEX", 2U, 4U},
    {"INDIRECT", 1U, 2U},
    {"INFO", 1U, 1U},
    {"INT", 1U, 1U},
    {"INTERCEPT", 2U, 2U},
    {"INTRATE", 4U, 5U},
    {"IPMT", 4U, 6U},
    {"IRR", 1U, 2U},
    {"ISBLANK", 1U, 1U},
    {"ISERR", 1U, 1U},
    {"ISERROR", 1U, 1U},
    {"ISEVEN", 1U, 1U},
    {"ISFORMULA", 1U, 1U},
    {"ISLOGICAL", 1U, 1U},
    {"ISNA", 1U, 1U},
    {"ISNONTEXT", 1U, 1U},
    {"ISNUMBER", 1U, 1U},
    {"ISO.CEILING", 1U, 2U},
    {"ISODD", 1U, 1U},
    {"ISOMITTED", 1U, 1U},
    {"ISOWEEKNUM", 1U, 1U},
    {"ISPMT", 4U, 4U},
    {"ISREF", 1U, 1U},
    {"ISTEXT", 1U, 1U},
    {"JIS", 1U, 1U},
    {"KURT", 1U, kVariadic},
    {"LAMBDA", 1U, 254U},
    {"LARGE", 2U, 2U},
    {"LCM", 1U, kVariadic},
    {"LEFT", 1U, 2U},
    {"LEFTB", 1U, 2U},
    {"LEN", 1U, 1U},
    {"LENB", 1U, 1U},
    {"LET", 3U, 253U},
    {"LINEST", 1U, 4U},
    {"LN", 1U, 1U},
    {"LOG", 1U, 2U},
    {"LOG10", 1U, 1U},
    {"LOGEST", 1U, 4U},
    {"LOGINV", 3U, 3U},
    {"LOGNORM.DIST", 4U, 4U},
    {"LOGNORM.INV", 3U, 3U},
    {"LOGNORMDIST", 3U, 3U},
    {"LOOKUP", 2U, 3U},
    {"LOWER", 1U, 1U},
    {"MAKEARRAY", 3U, 3U},
    {"MAP", 2U, kVariadic},
    {"MATCH", 2U, 3U},
    {"MAX", 1U, kVariadic},
    {"MAXA", 1U, kVariadic},
    {"MAXIFS", 3U, kVariadic},
    {"MDETERM", 1U, 1U},
    {"MDURATION", 5U, 6U},
    {"MEDIAN", 1U, kVariadic},
    {"MID", 3U, 3U},
    {"MIDB", 3U, 3U},
    {"MIN", 1U, kVariadic},
    {"MINA", 1U, kVariadic},
    {"MINIFS", 3U, kVariadic},
    {"MINUTE", 1U, 1U},
    {"MINVERSE", 1U, 1U},
    {"MIRR", 3U, 3U},
    {"MMULT", 2U, 2U},
    {"MOD", 2U, 2U},
    {"MODE", 1U, kVariadic},
    {"MODE.MULT", 1U, kVariadic},
    {"MODE.SNGL", 1U, kVariadic},
    {"MONTH", 1U, 1U},
    {"MROUND", 2U, 2U},
    {"MULTINOMIAL", 1U, kVariadic},
    {"MUNIT", 1U, 1U},
    {"N", 1U, 1U},
    {"NA", 0U, 0U},
    {"NEGBINOM.DIST", 4U, 4U},
    {"NEGBINOMDIST", 3U, 3U},
    {"NETWORKDAYS", 2U, 3U},
    {"NETWORKDAYS.INTL", 2U, 4U},
    {"NOMINAL", 2U, 2U},
    {"NORM.DIST", 4U, 4U},
    {"NORM.INV", 3U, 3U},
    {"NORM.S.DIST", 2U, 2U},
    {"NORM.S.INV", 1U, 1U},
    {"NORMDIST", 4U, 4U},
    {"NORMINV", 3U, 3U},
    {"NORMSDIST", 1U, 1U},
    {"NORMSINV", 1U, 1U},
    {"NOT", 1U, 1U},
    {"NOW", 0U, 0U},
    {"NPER", 3U, 5U},
    {"NPV", 2U, kVariadic},
    {"NUMBERVALUE", 1U, 3U},
    {"OCT2BIN", 1U, 2U},
    {"OCT2DEC", 1U, 1U},
    {"OCT2HEX", 1U, 2U},
    {"ODD", 1U, 1U},
    {"ODDFPRICE", 8U, 9U},
    {"ODDFYIELD", 8U, 9U},
    {"ODDLPRICE", 7U, 8U},
    {"ODDLYIELD", 7U, 8U},
    {"OFFSET", 3U, 5U},
    {"OR", 1U, kVariadic},
    {"PDURATION", 3U, 3U},
    {"PEARSON", 2U, 2U},
    {"PERCENTILE", 2U, 2U},
    {"PERCENTILE.EXC", 2U, 2U},
    {"PERCENTILE.INC", 2U, 2U},
    {"PERCENTOF", 2U, 2U},
    {"PERCENTRANK", 2U, 3U},
    {"PERCENTRANK.EXC", 2U, 3U},
    {"PERCENTRANK.INC", 2U, 3U},
    {"PERMUT", 2U, 2U},
    {"PERMUTATIONA", 2U, 2U},
    {"PHI", 1U, 1U},
    {"PHONETIC", 1U, 1U},
    {"PI", 0U, 0U},
    {"PIVOTBY", 4U, 11U},
    {"PMT", 3U, 5U},
    {"POISSON", 3U, 3U},
    {"POISSON.DIST", 3U, 3U},
    {"POWER", 2U, 2U},
    {"PPMT", 4U, 6U},
    {"PRICE", 6U, 7U},
    {"PRICEDISC", 4U, 5U},
    {"PRICEMAT", 5U, 6U},
    {"PROB", 3U, 4U},
    {"PRODUCT", 1U, kVariadic},
    {"PROPER", 1U, 1U},
    {"PV", 3U, 5U},
    {"PY", 1U, 1U},
    {"QUARTILE", 2U, 2U},
    {"QUARTILE.EXC", 2U, 2U},
    {"QUARTILE.INC", 2U, 2U},
    {"QUOTIENT", 2U, 2U},
    {"RADIANS", 1U, 1U},
    {"RAND", 0U, 0U},
    {"RANDARRAY", 0U, 5U},
    {"RANDBETWEEN", 2U, 2U},
    {"RANK", 2U, 3U},
    {"RANK.AVG", 2U, 3U},
    {"RANK.EQ", 2U, 3U},
    {"RATE", 3U, 6U},
    {"RECEIVED", 4U, 5U},
    {"REDUCE", 3U, 3U},
    {"REGEXEXTRACT", 2U, 4U},
    {"REGEXREPLACE", 3U, 5U},
    {"REGEXTEST", 2U, 3U},
    {"REPLACE", 4U, 4U},
    {"REPLACEB", 4U, 4U},
    {"REPT", 2U, 2U},
    {"RIGHT", 1U, 2U},
    {"RIGHTB", 1U, 2U},
    {"ROMAN", 1U, 2U},
    {"ROUND", 2U, 2U},
    {"ROUNDDOWN", 2U, 2U},
    {"ROUNDUP", 2U, 2U},
    {"ROW", 0U, 1U},
    {"ROWS", 1U, 1U},
    {"RRI", 3U, 3U},
    {"RSQ", 2U, 2U},
    {"RTD", 3U, kVariadic},
    {"SCAN", 3U, 3U},
    {"SEARCH", 2U, 3U},
    {"SEARCHB", 2U, 3U},
    {"SEC", 1U, 1U},
    {"SECH", 1U, 1U},
    {"SECOND", 1U, 1U},
    {"SEQUENCE", 1U, 4U},
    {"SERIESSUM", 4U, 4U},
    {"SHEET", 0U, 1U},
    {"SHEETS", 0U, 1U},
    {"SIGN", 1U, 1U},
    {"SIN", 1U, 1U},
    {"SINGLE", 1U, 1U},
    {"SINH", 1U, 1U},
    {"SKEW", 1U, kVariadic},
    {"SKEW.P", 1U, kVariadic},
    {"SLN", 3U, 3U},
    {"SLOPE", 2U, 2U},
    {"SMALL", 2U, 2U},
    {"SORT", 1U, 4U},
    {"SORTBY", 2U, kVariadic},
    {"SQRT", 1U, 1U},
    {"SQRTPI", 1U, 1U},
    {"STANDARDIZE", 3U, 3U},
    {"STDEV", 1U, kVariadic},
    {"STDEV.P", 1U, kVariadic},
    {"STDEV.S", 1U, kVariadic},
    {"STDEVA", 1U, kVariadic},
    {"STDEVP", 1U, kVariadic},
    {"STDEVPA", 1U, kVariadic},
    {"STEYX", 2U, 2U},
    {"STOCKHISTORY", 2U, kVariadic},
    {"SUBSTITUTE", 3U, 4U},
    {"SUBTOTAL", 2U, kVariadic},
    {"SUM", 1U, kVariadic},
    {"SUMIF", 2U, 3U},
    {"SUMIFS", 3U, kVariadic},
    {"SUMPRODUCT", 1U, kVariadic},
    {"SUMSQ", 1U, kVariadic},
    {"SUMX2MY2", 2U, 2U},
    {"SUMX2PY2", 2U, 2U},
    {"SUMXMY2", 2U, 2U},
    {"SWITCH", 3U, kVariadic},
    {"SYD", 4U, 4U},
    {"T", 1U, 1U},
    {"T.DIST", 3U, 3U},
    {"T.DIST.2T", 2U, 2U},
    {"T.DIST.RT", 2U, 2U},
    {"T.INV", 2U, 2U},
    {"T.INV.2T", 2U, 2U},
    {"T.TEST", 4U, 4U},
    {"TAKE", 2U, 3U},
    {"TAN", 1U, 1U},
    {"TANH", 1U, 1U},
    {"TBILLEQ", 3U, 3U},
    {"TBILLPRICE", 3U, 3U},
    {"TBILLYIELD", 3U, 3U},
    {"TDIST", 3U, 3U},
    {"TEXT", 2U, 2U},
    {"TEXTAFTER", 2U, 6U},
    {"TEXTBEFORE", 2U, 6U},
    {"TEXTJOIN", 3U, kVariadic},
    {"TEXTSPLIT", 2U, 6U},
    {"TIME", 3U, 3U},
    {"TIMEVALUE", 1U, 1U},
    {"TINV", 2U, 2U},
    {"TOCOL", 1U, 3U},
    {"TODAY", 0U, 0U},
    {"TOROW", 1U, 3U},
    {"TRANSLATE", 1U, 3U},
    {"TRANSPOSE", 1U, 1U},
    {"TREND", 1U, 4U},
    {"TRIM", 1U, 1U},
    {"TRIMMEAN", 2U, 2U},
    {"TRIMRANGE", 1U, 3U},
    {"TRUE", 0U, 0U},
    {"TRUNC", 1U, 2U},
    {"TTEST", 4U, 4U},
    {"TYPE", 1U, 1U},
    {"UNICHAR", 1U, 1U},
    {"UNICODE", 1U, 1U},
    {"UNIQUE", 1U, 3U},
    {"UPPER", 1U, 1U},
    {"USDOLLAR", 1U, 2U},
    {"VALUE", 1U, 1U},
    {"VALUETOTEXT", 1U, 2U},
    {"VAR", 1U, kVariadic},
    {"VAR.P", 1U, kVariadic},
    {"VAR.S", 1U, kVariadic},
    {"VARA", 1U, kVariadic},
    {"VARP", 1U, kVariadic},
    {"VARPA", 1U, kVariadic},
    {"VDB", 5U, 7U},
    {"VLOOKUP", 3U, 4U},
    {"VSTACK", 1U, kVariadic},
    {"WEBSERVICE", 1U, 1U},
    {"WEEKDAY", 1U, 2U},
    {"WEEKNUM", 1U, 2U},
    {"WEIBULL", 4U, 4U},
    {"WEIBULL.DIST", 4U, 4U},
    {"WORKDAY", 2U, 3U},
    {"WORKDAY.INTL", 2U, 4U},
    {"WRAPCOLS", 2U, 3U},
    {"WRAPROWS", 2U, 3U},
    {"XIRR", 2U, 3U},
    {"XLOOKUP", 3U, 6U},
    {"XMATCH", 2U, 4U},
    {"XNPV", 3U, 3U},
    {"XOR", 1U, kVariadic},
    {"YEAR", 1U, 1U},
    {"YEARFRAC", 2U, 3U},
    {"YIELD", 6U, 7U},
    {"YIELDDISC", 4U, 5U},
    {"YIELDMAT", 5U, 6U},
    {"Z.TEST", 2U, 3U},
    {"ZTEST", 2U, 3U},
    {"_TRO_ALL", 1U, 1U},
    {"_TRO_LEADING", 1U, 1U},
    {"_TRO_TRAILING", 1U, 1U},
};

TEST(BuiltinNames, BuiltinArityMatchesExcel) {
  std::set<std::string> expected;
  for (const ExpectedArity& row : kExpectedArity) {
    expected.insert(row.name);
    std::uint32_t min_arity = 0U;
    std::uint32_t max_arity = 0U;
    ASSERT_TRUE(builtin_arity(row.name, &min_arity, &max_arity)) << row.name;
    EXPECT_EQ(min_arity, row.min_arity) << row.name;
    EXPECT_EQ(max_arity, row.max_arity) << row.name;
  }
  std::set<std::string> known;
  default_registry().for_each_name(
      [](std::string_view name, void* out) { static_cast<std::set<std::string>*>(out)->emplace(name); }, &known);
  for (const char* const* p = lazy_form_names(); *p != nullptr; ++p) {
    known.emplace(*p);
  }
  for (const char* const* p = parser_special_form_names(); *p != nullptr; ++p) {
    known.emplace(*p);
  }
  EXPECT_EQ(known, expected);
}

TEST(BuiltinNames, EagerAndLazyRegistrationsAgreeOnArity) {
  for (const char* const* p = lazy_form_names(); *p != nullptr; ++p) {
    std::uint32_t min_arity = 0U;
    std::uint32_t max_arity = 0U;
    ASSERT_TRUE(find_lazy_arity(*p, &min_arity, &max_arity)) << *p;
    if (const FunctionDef* def = default_registry().lookup(*p); def != nullptr) {
      EXPECT_EQ(def->min_arity, min_arity) << *p;
      EXPECT_EQ(def->max_arity, max_arity) << *p;
    }
    if (const DateEntry* entry = find_date_entry(*p); entry != nullptr) {
      EXPECT_EQ(entry->min_arity, min_arity) << *p;
      EXPECT_EQ(entry->max_arity, max_arity) << *p;
    }
    if (const FinancialDateEntry* entry = find_financial_date_entry(*p); entry != nullptr) {
      EXPECT_EQ(entry->min_arity, min_arity) << *p;
      EXPECT_EQ(entry->max_arity, max_arity) << *p;
    }
  }
}

std::string Arguments(std::uint32_t count) {
  std::string out;
  for (std::uint32_t i = 0; i < count; ++i) {
    out += i == 0U ? "A1" : ",A1";
  }
  return out;
}

TEST(BuiltinNames, WrongArgumentCountIsFound) {
  for (const std::string& formula :
       {std::string("ROWS()"), std::string("ABS(1,2)"), std::string("ABS(1,)"), std::string("SUM()"),
        std::string("IF(1)"), std::string("1+ROWS()"), std::string("COUNTIFS(A1,1,A1)"),
        std::string("SUMIFS(A1,A1,1,A1)"), std::string("SORTBY(A1,A1,1,A1)"),
        std::string("GETPIVOTDATA(A1,A1,A1,A1,A1)"), std::string("LARGE(A1,A1,1)"),
        std::string("REDUCE(A1,LAMBDA(a,b,a))"), std::string("_xlfn.IFS(1)"), "SUM(" + Arguments(256) + ")",
        "FOO(" + Arguments(256) + ")"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(formula, arena, storage);
    ASSERT_NE(root, nullptr) << formula;
    EXPECT_NE(find_wrong_arity_call(*root), nullptr) << formula;
  }
  for (const std::string& formula :
       {std::string("ROW()"), std::string("NOW()"), std::string("SUM(A1)"), std::string("SUM(A1,)"),
        std::string("COUNTIFS(A1,1,A1,1)"), std::string("SORTBY(A1,A1)"), std::string("SORTBY(A1,A1,1)"),
        std::string("GETPIVOTDATA(A1,A1,A1)"), std::string("GETPIVOTDATA(A1,A1,A1,A1)"), std::string("IFS(1,1)"),
        std::string("SWITCH(1,1,1)"), std::string("TRANSLATE(\"a\")"), std::string("FOO(1,2,3)"),
        std::string("AGGREGATE(9,4,A1,1,)"), std::string("LET(x,1,x)"), std::string("LAMBDA(x,x)(1,2)"),
        "SUM(" + Arguments(255) + ")"}) {
    Arena arena;
    std::string storage;
    const parser::AstNode* root = ParseEntry(formula, arena, storage);
    ASSERT_NE(root, nullptr) << formula;
    EXPECT_EQ(find_wrong_arity_call(*root), nullptr) << formula;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
