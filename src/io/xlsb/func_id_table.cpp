//
// Implementation of the MS-XLSB function-id mapping. The table is a
// flat `constexpr` array sorted by id, so a binary search lands the row
// in O(log N) without `<map>` / `<unordered_map>`.
//
// Coverage policy: every row is either an id documented in [MS-XLS]
// §2.5.198.17 (Cetab) / [MS-XLSB] §2.5.97.74, or an id decoded from the
// Ptg stream of an Excel-produced workbook. The Reader surfaces `#NAME?`
// for any id missing from this table, so the table can grow
// incrementally as additional names are needed; we deliberately do not
// pad it out with guessed ids, because a wrong id substitutes a
// different function in both directions.

#include "io/xlsb/func_id_table.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cstddef>
#include <cstdint>

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Sentinel for "no documented upper bound" used by variadic functions.
constexpr std::uint8_t kVariadicMax = 255;

constexpr std::size_t kEntriesCountConst = 357;
using FuncEntryArray = std::array<XlsbFuncEntry, kEntriesCountConst>;

constexpr FuncEntryArray kEntries = {{
    // ---- Classic Excel functions ([MS-XLS] §2.5.198.17 Cetab) -------------
    // The id range 0..0x16C is documented; we include the entries we have
    // a credible source for. Functions absent from this list fall back to
    // `#NAME?` at read time, which is the documented behaviour for
    // unknown ids.
    {0, "COUNT", 0, kVariadicMax, true},
    {1, "IF", 2, 3, false},
    {2, "ISNA", 1, 1, false},
    {3, "ISERROR", 1, 1, false},
    {4, "SUM", 0, kVariadicMax, true},
    {5, "AVERAGE", 0, kVariadicMax, true},
    {6, "MIN", 0, kVariadicMax, true},
    {7, "MAX", 0, kVariadicMax, true},
    {8, "ROW", 0, 1, true},
    {9, "COLUMN", 0, 1, true},
    {10, "NA", 0, 0, false},
    {11, "NPV", 2, kVariadicMax, true},
    {12, "STDEV", 0, kVariadicMax, true},
    {13, "DOLLAR", 1, 2, true},
    {14, "FIXED", 1, 3, true},
    {15, "SIN", 1, 1, false},
    {16, "COS", 1, 1, false},
    {17, "TAN", 1, 1, false},
    {18, "ATAN", 1, 1, false},
    {19, "PI", 0, 0, false},
    {20, "SQRT", 1, 1, false},
    {21, "EXP", 1, 1, false},
    {22, "LN", 1, 1, false},
    {23, "LOG10", 1, 1, false},
    {24, "ABS", 1, 1, false},
    {25, "INT", 1, 1, false},
    {26, "SIGN", 1, 1, false},
    {27, "ROUND", 2, 2, false},
    {28, "LOOKUP", 2, 3, true},
    {29, "INDEX", 2, 4, true},
    {30, "REPT", 2, 2, false},
    {31, "MID", 3, 3, false},
    {32, "LEN", 1, 1, false},
    {33, "VALUE", 1, 1, false},
    {34, "TRUE", 0, 0, false},
    {35, "FALSE", 0, 0, false},
    {36, "AND", 1, kVariadicMax, true},
    {37, "OR", 1, kVariadicMax, true},
    {38, "NOT", 1, 1, false},
    {39, "MOD", 2, 2, false},
    {40, "DCOUNT", 3, 3, false},
    {41, "DSUM", 3, 3, false},
    {42, "DAVERAGE", 3, 3, false},
    {43, "DMIN", 3, 3, false},
    {44, "DMAX", 3, 3, false},
    {45, "DSTDEV", 3, 3, false},
    {46, "VAR", 0, kVariadicMax, true},
    {47, "DVAR", 3, 3, false},
    {48, "TEXT", 2, 2, false},
    {49, "LINEST", 1, 4, true},
    {50, "TREND", 1, 4, true},
    {51, "LOGEST", 1, 4, true},
    {52, "GROWTH", 1, 4, true},
    {56, "PV", 3, 5, true},
    {57, "FV", 3, 5, true},
    {58, "NPER", 3, 5, true},
    {59, "PMT", 3, 5, true},
    {60, "RATE", 3, 6, true},
    {61, "MIRR", 3, 3, false},
    {62, "IRR", 1, 2, true},
    {63, "RAND", 0, 0, false},
    {64, "MATCH", 2, 3, true},
    {65, "DATE", 3, 3, false},
    {66, "TIME", 3, 3, false},
    {67, "DAY", 1, 1, false},
    {68, "MONTH", 1, 1, false},
    {69, "YEAR", 1, 1, false},
    {70, "WEEKDAY", 1, 2, true},
    {71, "HOUR", 1, 1, false},
    {72, "MINUTE", 1, 1, false},
    {73, "SECOND", 1, 1, false},
    {74, "NOW", 0, 0, false},
    {75, "AREAS", 1, 1, false},
    {76, "ROWS", 1, 1, false},
    {77, "COLUMNS", 1, 1, false},
    {78, "OFFSET", 3, 5, true},
    {82, "SEARCH", 2, 3, true},
    {83, "TRANSPOSE", 1, 1, false},
    {86, "TYPE", 1, 1, false},
    {97, "ATAN2", 2, 2, false},
    {98, "ASIN", 1, 1, false},
    {99, "ACOS", 1, 1, false},
    {100, "CHOOSE", 2, kVariadicMax, true},
    {101, "HLOOKUP", 3, 4, true},
    {102, "VLOOKUP", 3, 4, true},
    {105, "ISREF", 1, 1, false},
    {109, "LOG", 1, 2, true},
    {111, "CHAR", 1, 1, false},
    {112, "LOWER", 1, 1, false},
    {113, "UPPER", 1, 1, false},
    {114, "PROPER", 1, 1, false},
    {115, "LEFT", 1, 2, true},
    {116, "RIGHT", 1, 2, true},
    {117, "EXACT", 2, 2, false},
    {118, "TRIM", 1, 1, false},
    {119, "REPLACE", 4, 4, false},
    {120, "SUBSTITUTE", 3, 4, true},
    {121, "CODE", 1, 1, false},
    {124, "FIND", 2, 3, true},
    {125, "CELL", 1, 2, true},
    {126, "ISERR", 1, 1, false},
    {127, "ISTEXT", 1, 1, false},
    {128, "ISNUMBER", 1, 1, false},
    {129, "ISBLANK", 1, 1, false},
    {130, "T", 1, 1, false},
    {131, "N", 1, 1, false},
    {140, "DATEVALUE", 1, 1, false},
    {141, "TIMEVALUE", 1, 1, false},
    {142, "SLN", 3, 3, false},
    {143, "SYD", 4, 4, false},
    {144, "DDB", 4, 5, true},
    {148, "INDIRECT", 1, 2, true},
    {162, "CLEAN", 1, 1, false},
    {163, "MDETERM", 1, 1, false},
    {164, "MINVERSE", 1, 1, false},
    {165, "MMULT", 2, 2, false},
    {167, "IPMT", 4, 6, true},
    {168, "PPMT", 4, 6, true},
    {169, "COUNTA", 0, kVariadicMax, true},
    {183, "PRODUCT", 0, kVariadicMax, true},
    {184, "FACT", 1, 1, false},
    {189, "DPRODUCT", 3, 3, false},
    {190, "ISNONTEXT", 1, 1, false},
    {193, "STDEVP", 0, kVariadicMax, true},
    {194, "VARP", 0, kVariadicMax, true},
    {195, "DSTDEVP", 3, 3, false},
    {196, "DVARP", 3, 3, false},
    {197, "TRUNC", 1, 2, true},
    {198, "ISLOGICAL", 1, 1, false},
    {199, "DCOUNTA", 3, 3, false},
    {204, "USDOLLAR", 1, 2, true},
    {205, "FINDB", 2, 3, true},
    {206, "SEARCHB", 2, 3, true},
    {207, "REPLACEB", 4, 4, false},
    {208, "LEFTB", 1, 2, true},
    {209, "RIGHTB", 1, 2, true},
    {210, "MIDB", 3, 3, false},
    {211, "LENB", 1, 1, false},
    {212, "ROUNDUP", 2, 2, false},
    {213, "ROUNDDOWN", 2, 2, false},
    {214, "ASC", 1, 1, false},
    {215, "DBCS", 1, 1, false},
    {216, "RANK", 2, 3, true},
    {219, "ADDRESS", 2, 5, true},
    {220, "DAYS360", 2, 3, true},
    {221, "TODAY", 0, 0, false},
    {222, "VDB", 5, 7, true},
    {227, "MEDIAN", 0, kVariadicMax, true},
    {228, "SUMPRODUCT", 1, kVariadicMax, true},
    {229, "SINH", 1, 1, false},
    {230, "COSH", 1, 1, false},
    {231, "TANH", 1, 1, false},
    {232, "ASINH", 1, 1, false},
    {233, "ACOSH", 1, 1, false},
    {234, "ATANH", 1, 1, false},
    {235, "DGET", 3, 3, false},
    {244, "INFO", 1, 1, false},
    {247, "DB", 4, 5, true},
    {252, "FREQUENCY", 2, 2, false},
    {261, "ERROR.TYPE", 1, 1, false},
    {269, "AVEDEV", 1, kVariadicMax, true},
    {270, "BETADIST", 3, 5, true},
    {271, "GAMMALN", 1, 1, false},
    {272, "BETAINV", 3, 5, true},
    {273, "BINOMDIST", 4, 4, false},
    {274, "CHIDIST", 2, 2, false},
    {275, "CHIINV", 2, 2, false},
    {276, "COMBIN", 2, 2, false},
    {277, "CONFIDENCE", 3, 3, false},
    {278, "CRITBINOM", 3, 3, false},
    {279, "EVEN", 1, 1, false},
    {280, "EXPONDIST", 3, 3, false},
    {281, "FDIST", 3, 3, false},
    {282, "FINV", 3, 3, false},
    {283, "FISHER", 1, 1, false},
    {284, "FISHERINV", 1, 1, false},
    {285, "FLOOR", 2, 2, false},
    {286, "GAMMADIST", 4, 4, false},
    {287, "GAMMAINV", 3, 3, false},
    {288, "CEILING", 2, 2, false},
    {289, "HYPGEOMDIST", 4, 4, false},
    {290, "LOGNORMDIST", 3, 3, false},
    {291, "LOGINV", 3, 3, false},
    {292, "NEGBINOMDIST", 3, 3, false},
    {293, "NORMDIST", 4, 4, false},
    {294, "NORMSDIST", 1, 1, false},
    {295, "NORMINV", 3, 3, false},
    {296, "NORMSINV", 1, 1, false},
    {297, "STANDARDIZE", 3, 3, false},
    {298, "ODD", 1, 1, false},
    {299, "PERMUT", 2, 2, false},
    {300, "POISSON", 3, 3, false},
    {301, "TDIST", 3, 3, false},
    {302, "WEIBULL", 4, 4, false},
    {303, "SUMXMY2", 2, 2, false},
    {304, "SUMX2MY2", 2, 2, false},
    {305, "SUMX2PY2", 2, 2, false},
    {306, "CHITEST", 2, 2, false},
    {307, "CORREL", 2, 2, false},
    {308, "COVAR", 2, 2, false},
    {309, "FORECAST", 3, 3, false},
    {310, "FTEST", 2, 2, false},
    {311, "INTERCEPT", 2, 2, false},
    {312, "PEARSON", 2, 2, false},
    {313, "RSQ", 2, 2, false},
    {314, "STEYX", 2, 2, false},
    {315, "SLOPE", 2, 2, false},
    {316, "TTEST", 4, 4, false},
    {317, "PROB", 3, 4, true},
    {318, "DEVSQ", 1, kVariadicMax, true},
    {319, "GEOMEAN", 1, kVariadicMax, true},
    {320, "HARMEAN", 1, kVariadicMax, true},
    {321, "SUMSQ", 0, kVariadicMax, true},
    {322, "KURT", 1, kVariadicMax, true},
    {323, "SKEW", 1, kVariadicMax, true},
    {324, "ZTEST", 2, 3, true},
    {325, "LARGE", 2, 2, false},
    {326, "SMALL", 2, 2, false},
    {327, "QUARTILE", 2, 2, false},
    {328, "PERCENTILE", 2, 2, false},
    {329, "PERCENTRANK", 2, 3, true},
    {330, "MODE", 1, kVariadicMax, true},
    {331, "TRIMMEAN", 2, 2, false},
    {332, "TINV", 2, 2, false},
    {336, "CONCATENATE", 0, kVariadicMax, true},
    {337, "POWER", 2, 2, false},
    {342, "RADIANS", 1, 1, false},
    {343, "DEGREES", 1, 1, false},
    {344, "SUBTOTAL", 2, kVariadicMax, true},
    {345, "SUMIF", 2, 3, true},
    {346, "COUNTIF", 2, 2, false},
    {347, "COUNTBLANK", 1, 1, false},
    {350, "ISPMT", 4, 4, false},
    {351, "DATEDIF", 3, 3, false},
    {352, "DATESTRING", 1, 1, false},
    {353, "NUMBERSTRING", 2, 2, false},
    {354, "ROMAN", 1, 2, true},
    {358, "GETPIVOTDATA", 2, kVariadicMax, true},
    {359, "HYPERLINK", 1, 2, true},
    {360, "PHONETIC", 1, 1, false},
    {361, "AVERAGEA", 0, kVariadicMax, true},
    {362, "MAXA", 0, kVariadicMax, true},
    {363, "MINA", 0, kVariadicMax, true},
    {364, "STDEVPA", 0, kVariadicMax, true},
    {365, "VARPA", 0, kVariadicMax, true},
    {366, "STDEVA", 0, kVariadicMax, true},
    {367, "VARA", 0, kVariadicMax, true},
    {368, "BAHTTEXT", 1, 1, false},
    {369, "THAIDAYOFWEEK", 1, 1, false},
    {370, "THAIDIGIT", 1, 1, false},
    {371, "THAIMONTHOFYEAR", 1, 1, false},
    {372, "THAINUMSOUND", 1, 1, false},
    {373, "THAINUMSTRING", 1, 1, false},
    {374, "THAISTRINGLENGTH", 1, 1, false},
    {375, "ISTHAIDIGIT", 1, 1, false},
    {376, "ROUNDBAHTDOWN", 1, 1, false},
    {377, "ROUNDBAHTUP", 1, 1, false},
    {378, "THAIYEAR", 1, 1, false},
    {379, "RTD", 2, kVariadicMax, true},
    // ---- Analysis ToolPak functions Excel 2007 made native ---------------
    // These ids were decoded from the `PtgFunc` / `PtgFuncVar` tokens of an
    // Excel-365-produced workbook (`tests/fixtures/excel/xlsb_func_ids.xlsb`,
    // one probe formula per function per arity); the harvester that produced
    // it is `tools/dev/xlsb_func_id_harvest.py`. A row is fixed-arity
    // (`variadic = false`) exactly when Excel encoded the call as `PtgFunc`,
    // in which case `arg_min` is the arity Excel accepted -- the Reader pops
    // that many operands for such a call.
    {384, "HEX2BIN", 1, 2, true},
    {385, "HEX2DEC", 1, 1, false},
    {386, "HEX2OCT", 1, 2, true},
    {387, "DEC2BIN", 1, 2, true},
    {388, "DEC2HEX", 1, 2, true},
    {389, "DEC2OCT", 1, 2, true},
    {390, "OCT2BIN", 1, 2, true},
    {391, "OCT2HEX", 1, 2, true},
    {392, "OCT2DEC", 1, 1, false},
    {393, "BIN2DEC", 1, 1, false},
    {394, "BIN2OCT", 1, 2, true},
    {395, "BIN2HEX", 1, 2, true},
    {396, "IMSUB", 2, 2, false},
    {397, "IMDIV", 2, 2, false},
    {398, "IMPOWER", 2, 2, false},
    {399, "IMABS", 1, 1, false},
    {400, "IMSQRT", 1, 1, false},
    {401, "IMLN", 1, 1, false},
    {402, "IMLOG2", 1, 1, false},
    {403, "IMLOG10", 1, 1, false},
    {404, "IMSIN", 1, 1, false},
    {405, "IMCOS", 1, 1, false},
    {406, "IMEXP", 1, 1, false},
    {407, "IMARGUMENT", 1, 1, false},
    {408, "IMCONJUGATE", 1, 1, false},
    {409, "IMAGINARY", 1, 1, false},
    {410, "IMREAL", 1, 1, false},
    {411, "COMPLEX", 2, 3, true},
    {412, "IMSUM", 1, kVariadicMax, true},
    {413, "IMPRODUCT", 1, kVariadicMax, true},
    {414, "SERIESSUM", 4, 4, false},
    {415, "FACTDOUBLE", 1, 1, false},
    {416, "SQRTPI", 1, 1, false},
    {417, "QUOTIENT", 2, 2, false},
    {418, "DELTA", 1, 2, true},
    {419, "GESTEP", 1, 2, true},
    {420, "ISEVEN", 1, 1, false},
    {421, "ISODD", 1, 1, false},
    {422, "MROUND", 2, 2, false},
    {423, "ERF", 1, 2, true},
    {424, "ERFC", 1, 1, false},
    {425, "BESSELJ", 2, 2, false},
    {426, "BESSELK", 2, 2, false},
    {427, "BESSELY", 2, 2, false},
    {428, "BESSELI", 2, 2, false},
    {429, "XIRR", 2, 3, true},
    {430, "XNPV", 3, 3, false},
    {431, "PRICEMAT", 5, 6, true},
    {432, "YIELDMAT", 5, 6, true},
    {433, "INTRATE", 4, 5, true},
    {434, "RECEIVED", 4, 5, true},
    {435, "DISC", 4, 5, true},
    {436, "PRICEDISC", 4, 5, true},
    {437, "YIELDDISC", 4, 5, true},
    {438, "TBILLEQ", 3, 3, false},
    {439, "TBILLPRICE", 3, 3, false},
    {440, "TBILLYIELD", 3, 3, false},
    {441, "PRICE", 6, 7, true},
    {442, "YIELD", 6, 7, true},
    {443, "DOLLARDE", 2, 2, false},
    {444, "DOLLARFR", 2, 2, false},
    {445, "NOMINAL", 2, 2, false},
    {446, "EFFECT", 2, 2, false},
    {447, "CUMPRINC", 6, 6, false},
    {448, "CUMIPMT", 6, 6, false},
    {449, "EDATE", 2, 2, false},
    {450, "EOMONTH", 2, 2, false},
    {451, "YEARFRAC", 2, 3, true},
    {452, "COUPDAYBS", 3, 4, true},
    {453, "COUPDAYS", 3, 4, true},
    {454, "COUPDAYSNC", 3, 4, true},
    {455, "COUPNCD", 3, 4, true},
    {456, "COUPNUM", 3, 4, true},
    {457, "COUPPCD", 3, 4, true},
    {458, "DURATION", 5, 6, true},
    {459, "MDURATION", 5, 6, true},
    {460, "ODDLPRICE", 7, 8, true},
    {461, "ODDLYIELD", 7, 8, true},
    {462, "ODDFPRICE", 8, 9, true},
    {463, "ODDFYIELD", 8, 9, true},
    {464, "RANDBETWEEN", 2, 2, false},
    {465, "WEEKNUM", 1, 2, true},
    {466, "AMORDEGRC", 6, 7, true},
    {467, "AMORLINC", 6, 7, true},
    {468, "CONVERT", 3, 3, false},
    {469, "ACCRINT", 6, 8, true},
    {470, "ACCRINTM", 4, 5, true},
    {471, "WORKDAY", 2, 3, true},
    {472, "NETWORKDAYS", 2, 3, true},
    {473, "GCD", 1, kVariadicMax, true},
    {474, "MULTINOMIAL", 1, kVariadicMax, true},
    {475, "LCM", 1, kVariadicMax, true},
    {476, "FVSCHEDULE", 2, 2, false},
    {480, "IFERROR", 2, 2, false},
    {481, "COUNTIFS", 2, kVariadicMax, true},
    {482, "SUMIFS", 3, kVariadicMax, true},
    {483, "AVERAGEIF", 2, 3, true},
    {484, "AVERAGEIFS", 3, kVariadicMax, true},
}};

constexpr std::size_t kEntriesCount = kEntries.size();

constexpr bool TableIsSortedById() {
  for (std::size_t i = 1; i < kEntriesCount; ++i) {
    if (kEntries[i - 1].id >= kEntries[i].id) {
      return false;
    }
  }
  return true;
}

static_assert(TableIsSortedById(), "kEntries must be sorted by id");

// A name Excel accepts but never stores (the ja-JP formula bar's `JIS`
// for `DBCS`) is resolved by `io::canonical_function_name` before it
// reaches this table, so both persistence paths agree on the callee and
// there is exactly one place aliases are enumerated. This table stays a
// pure id <-> stored-name mapping.

/// ASCII upper-case fold for case-insensitive name lookup. Excel names
/// are ASCII so a `std::toupper`-equivalent is sufficient.
char AsciiUpper(char c) {
  if (c >= 'a' && c <= 'z') {
    return static_cast<char>(c - ('a' - 'A'));
  }
  return c;
}

bool NameEqualsIgnoreCase(std::string_view a, std::string_view b) {
  if (a.size() != b.size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.size(); ++i) {
    if (AsciiUpper(a[i]) != AsciiUpper(b[i])) {
      return false;
    }
  }
  return true;
}

}  // namespace

const std::size_t kXlsbFuncEntryCount = kEntriesCount;
const XlsbFuncEntry* const kXlsbFuncEntries = kEntries.data();

const XlsbFuncEntry* lookup_func_by_id(std::uint16_t id) {
  // NOLINTNEXTLINE(readability-qualified-auto): the iterator type may be
  // a raw pointer (libc++ macOS) or a wrapped iterator (libc++ wasm); a
  // bare `auto` works for both.
  auto it = std::lower_bound(kEntries.begin(), kEntries.end(), id,
                             [](const XlsbFuncEntry& entry, std::uint16_t v) { return entry.id < v; });
  if (it == kEntries.end() || it->id != id) {
    return nullptr;
  }
  return &(*it);
}

const XlsbFuncEntry* lookup_func_by_name(std::string_view name) {
  for (std::size_t i = 0; i < kEntriesCount; ++i) {
    if (NameEqualsIgnoreCase(name, kEntries[i].name)) {
      return &kEntries[i];
    }
  }
  return nullptr;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
