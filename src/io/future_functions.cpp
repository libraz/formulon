//
// Implementation of the storage-prefix classifier. See
// `io/future_functions.h`.
//
// The enumeration is a single `|`-delimited blob rather than an array of
// string views: one contiguous literal costs no per-entry pointer pair
// and no load-time relocation, which matters at ~170 entries in the WASM
// data segment. Lookup is a substring search for `|NAME|`, so the
// delimiters on both sides prevent a partial hit (`NORM.DIST` inside
// `NORM.S.DIST`). Grouping is by the Excel release that introduced the
// function, which is what decides membership.
//
// Membership was checked against the `<f>` text of the Excel-authored
// workbooks under `tests/fixtures/excel/` and
// `tests/oracle/external/ironcalc/fixtures/`; names those files do not
// exercise are classified by introduction date.

#include "io/future_functions.h"

#include <cstddef>
#include <string>
#include <string_view>

#include "parser/ast_format.h"

namespace formulon {
namespace io {
namespace {

/// Functions Excel stores with the `_xlfn.` prefix.
constexpr std::string_view kXlfnFunctions =
    "|"
    // Excel 2010.
    "AGGREGATE|BETA.DIST|BETA.INV|BINOM.DIST|BINOM.INV|CEILING.PRECISE|CHISQ.DIST|CHISQ.DIST.RT|CHISQ.INV|"
    "CHISQ.INV.RT|CHISQ.TEST|CONFIDENCE.NORM|CONFIDENCE.T|COVARIANCE.P|COVARIANCE.S|ERF.PRECISE|ERFC.PRECISE|"
    "EXPON.DIST|F.DIST|F.DIST.RT|F.INV|F.INV.RT|F.TEST|FLOOR.PRECISE|GAMMA.DIST|GAMMA.INV|GAMMALN.PRECISE|"
    "HYPGEOM.DIST|LOGNORM.DIST|LOGNORM.INV|MODE.MULT|MODE.SNGL|NEGBINOM.DIST|NETWORKDAYS.INTL|NORM.DIST|"
    "NORM.INV|NORM.S.DIST|NORM.S.INV|PERCENTILE.EXC|PERCENTILE.INC|PERCENTRANK.EXC|PERCENTRANK.INC|"
    "POISSON.DIST|QUARTILE.EXC|QUARTILE.INC|RANK.AVG|RANK.EQ|STDEV.P|STDEV.S|T.DIST|T.DIST.2T|T.DIST.RT|"
    "T.INV|T.INV.2T|T.TEST|VAR.P|VAR.S|WEIBULL.DIST|WORKDAY.INTL|Z.TEST|"
    // Excel 2013.
    "ACOT|ACOTH|ARABIC|BASE|BINOM.DIST.RANGE|BITAND|BITLSHIFT|BITOR|BITRSHIFT|BITXOR|CEILING.MATH|COMBINA|"
    "COT|COTH|CSC|CSCH|DAYS|DECIMAL|ENCODEURL|FILTERXML|FLOOR.MATH|FORMULATEXT|GAMMA|GAUSS|IFNA|IMCOSH|IMCOT|"
    "IMCSC|IMCSCH|IMSEC|IMSECH|IMSINH|IMTAN|ISFORMULA|ISOWEEKNUM|MUNIT|NUMBERVALUE|PDURATION|PERMUTATIONA|"
    "PHI|RRI|SEC|SECH|SHEET|SHEETS|SKEW.P|UNICHAR|UNICODE|WEBSERVICE|XOR|"
    // Excel 2016.
    "CONCAT|FORECAST.ETS|FORECAST.ETS.CONFINT|FORECAST.ETS.SEASONALITY|FORECAST.ETS.STAT|FORECAST.LINEAR|IFS|"
    "MAXIFS|MINIFS|SWITCH|TEXTJOIN|"
    // Excel 2019 / Microsoft 365.
    "ANCHORARRAY|ARRAYTOTEXT|BYCOL|BYROW|CHOOSECOLS|CHOOSEROWS|COPILOT|DETECTLANGUAGE|DROP|EXPAND|GROUPBY|"
    "HSTACK|IMAGE|ISOMITTED|LAMBDA|LET|MAKEARRAY|MAP|PERCENTOF|PIVOTBY|PY|RANDARRAY|REDUCE|REGEXEXTRACT|"
    "REGEXREPLACE|REGEXTEST|SCAN|SEQUENCE|SINGLE|SORTBY|STOCKHISTORY|TAKE|TEXTAFTER|TEXTBEFORE|TEXTSPLIT|"
    "TOCOL|TOROW|TRANSLATE|TRIMRANGE|UNIQUE|VALUETOTEXT|VSTACK|WRAPCOLS|WRAPROWS|XLOOKUP|XMATCH|";

/// Functions Excel spells bare in the OOXML `<f>` text but has no
/// classic function id for, so an XLSB call to one has to go through the
/// hidden-name route. Kept apart from `kXlfnFunctions` because the two
/// answer different questions (see `xlsb_uses_hidden_name`); a name here
/// is NOT `_xlfn.`-prefixed in `<f>`.
///
/// Membership requires bytes decoded out of an Excel-produced workbook.
/// `ISO.CEILING` was observed in both containers from one workbook saved
/// twice by Excel 365: `xl/worksheets/sheet1.xml` holds
/// `<f>ISO.CEILING(4.3)</f>` with no `<definedNames>` element at all,
/// while the `.xlsb` sibling holds a `BrtName` spelling
/// `_xlfn.ISO.CEILING` and a cell stream of `PtgName`, `PtgNum`,
/// `PtgFuncVar(cparams=2, id=255)`.
constexpr std::string_view kXlsbHiddenNameFunctions = "|ISO.CEILING|";

/// One localised formula-bar spelling and the name Excel stores for it.
///
/// Verified against Excel 365 (ja-JP): `Range.formula` and the `<f>` text
/// read `DBCS("ABC")` where `Range.formula_local` reads `JIS("ABC")`,
/// assigning `JIS(...)` through the localised property produces a cell
/// whose stored formula is `DBCS(...)`, the call encodes as `PtgFunc`
/// id 215 in `.xlsb`, and `JIS` sent through the invariant API is not a
/// function at all (`#NAME?` plus a bare defined name).
struct FunctionNameAlias {
  std::string_view localised;  ///< Spelling Excel accepts but never stores.
  std::string_view stored;     ///< Spelling Excel writes to the file.
};

constexpr FunctionNameAlias kFunctionNameAliases[] = {
    {"JIS", "DBCS"},
};

/// The subset Excel additionally marks worksheet-only, stored as
/// `_xlfn._xlws.<NAME>`. The sibling dynamic-array functions (SORTBY,
/// UNIQUE, SEQUENCE, ...) take the plain `_xlfn.` form; only these two
/// carry the doubled prefix in Excel-authored files.
constexpr std::string_view kXlwsFunctions = "|FILTER|SORT|";

/// ASCII upper-case fold. Excel function names are ASCII.
std::string AsciiUpper(std::string_view name) {
  std::string upper;
  upper.reserve(name.size());
  for (char c : name) {
    upper.push_back((c >= 'a' && c <= 'z') ? static_cast<char>(c - 'a' + 'A') : c);
  }
  return upper;
}

/// True when `blob` lists `upper` as a whole `|`-delimited entry.
bool BlobContains(std::string_view blob, std::string_view upper) {
  std::string needle;
  needle.reserve(upper.size() + 2U);
  needle.push_back('|');
  needle.append(upper);
  needle.push_back('|');
  return blob.find(needle) != std::string_view::npos;
}

}  // namespace

parser::StoragePrefixKind classify_storage_prefix(std::string_view canonical_name) {
  const std::string upper = AsciiUpper(canonical_name);
  if (BlobContains(kXlwsFunctions, upper)) {
    return parser::StoragePrefixKind::XlfnXlws;
  }
  if (BlobContains(kXlfnFunctions, upper)) {
    return parser::StoragePrefixKind::Xlfn;
  }
  return parser::StoragePrefixKind::None;
}

std::string storage_function_name(std::string_view canonical_name) {
  const std::string upper = AsciiUpper(canonical_function_name(canonical_name));
  switch (classify_storage_prefix(upper)) {
    case parser::StoragePrefixKind::XlfnXlws:
      return std::string("_xlfn._xlws.") + upper;
    case parser::StoragePrefixKind::Xlfn:
      return std::string("_xlfn.") + upper;
    case parser::StoragePrefixKind::None:
      break;
  }
  return upper;
}

std::string_view canonical_function_name(std::string_view name) {
  const std::string upper = AsciiUpper(name);
  for (const FunctionNameAlias& alias : kFunctionNameAliases) {
    if (upper == alias.localised) {
      return alias.stored;
    }
  }
  return name;
}

std::string storage_call_name(std::string_view name) {
  const std::string_view stored = canonical_function_name(name);
  switch (classify_storage_prefix(stored)) {
    case parser::StoragePrefixKind::XlfnXlws:
      return std::string("_xlfn._xlws.") + std::string(stored);
    case parser::StoragePrefixKind::Xlfn:
      return std::string("_xlfn.") + std::string(stored);
    case parser::StoragePrefixKind::None:
      break;
  }
  return std::string(stored);
}

bool xlsb_uses_hidden_name(std::string_view canonical_name) {
  return BlobContains(kXlsbHiddenNameFunctions, AsciiUpper(canonical_name));
}

std::string xlsb_hidden_function_name(std::string_view canonical_name) {
  const std::string upper = AsciiUpper(canonical_name);
  if (classify_storage_prefix(upper) != parser::StoragePrefixKind::None) {
    return storage_function_name(upper);
  }
  return std::string("_xlfn.") + upper;
}

}  // namespace io
}  // namespace formulon
