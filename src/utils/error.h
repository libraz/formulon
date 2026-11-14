//
// Formulon internal error representation.
//
// `FormulonErrorCode` partitions the 0-9999 space across 10 modules.
// `Error` is the payload type used by `Expected<T, Error>` throughout the
// engine. It is distinct from the Excel-visible `ErrorCode`, which
// represents formula-level business errors like `#DIV/0!`.

#ifndef FORMULON_UTILS_ERROR_H_
#define FORMULON_UTILS_ERROR_H_

#include <cstdint>
#include <string>
#include <utility>

namespace formulon {

/// Enumeration of every internal error that the Formulon engine may surface.
///
/// Codes are partitioned by module in 1000-wide bands so a decimal digit of
/// the code identifies its origin. The exact catalog is the single source of
/// truth for the C API (`fm_status_t`) and the WASM / Python bindings. A number
/// missing from a band is retired and is never reassigned.
enum class FormulonErrorCode : int32_t {
  // ===== 0-999: General =====
  kOk = 0,
  kInvalidArgument = 2,
  kNotImplemented = 3,
  kOutOfMemory = 4,
  kNotFound = 6,
  kAlreadyExists = 7,
  kInternalError = 8,
  kPreconditionFailed = 9,
  /// Sheet index passed to a sheet-level operation is `>= sheet_count()`.
  kSheetIndexOutOfRange = 10,
  /// Sheet name fails structural validation: malformed UTF-8, empty, longer
  /// than 31 UTF-16 units, or contains a forbidden character (`: \ / ? * [ ]`).
  /// A well-formed name already taken is `kAlreadyExists`.
  kInvalidSheetName = 11,
  /// A non-sheet index addresses past the end of a counted collection.
  kIndexOutOfRange = 13,

  // ===== 1000-1999: Parser / Tokenizer =====
  kParserUnexpectedToken = 1000,
  kParserUnterminatedString = 1001,
  kParserInvalidNumber = 1003,
  kParserInvalidReference = 1004,
  kParserInvalidSheetName = 1005,
  kParserMismatchedBrackets = 1006,
  kParserExcessiveNesting = 1007,
  kParserInvalidArrayLiteral = 1009,
  kParserInvalidStructRef = 1010,
  kParserBomNotSupported = 1013,
  kParserTooManyErrors = 1014,

  // ===== 2000-2999: Evaluator =====
  /// Pivot evaluator was handed a `PivotTable` whose `pivot_cache_id` does
  /// not match the supplied `PivotCache`. Indicates an unresolved cache
  /// binding at evaluation time.
  kEvalPivotMissing = 2009,
  /// Pivot evaluator detected a structurally invalid table — most commonly
  /// a `PivotDataField::field_index` that points outside
  /// `PivotCache::fields()`.
  kEvalPivotInvalid = 2010,

  // ===== 3000-3999: Functions =====
  kFunctionOverflow = 3004,

  // ===== 4000-4999: Dependency graph / Recalc =====
  kGraphScheduleFailed = 4003,
  /// `recalc_parallel` was invoked while another `recalc_parallel` call on
  /// the same thread was still in flight. The scheduler does not support
  /// nested recalc; callers (typically a UDF that re-enters the engine)
  /// must avoid triggering recalc from inside an evaluator callback.
  kGraphRecalcReentrant = 4005,

  // ===== 5000-5999: I/O (OOXML / XLSB / CSV) =====
  kIoFileNotFound = 5000,
  kIoZipCorrupt = 5003,
  kIoZipEncrypted = 5004,
  kIoZipBomb = 5005,
  kIoZipSlip = 5006,
  kIoXmlParse = 5007,
  kIoRelationshipBroken = 5010,
  kIoContentTypeInvalid = 5011,
  kIoSheetCorrupt = 5012,
  kIoXlsbRecordCorrupt = 5014,
  kIoWriteFailed = 5016,
  kIoXlsbRecordTruncated = 5018,
  kIoXlsbUnknownRecord = 5019,
  kIoXlsbUnsupportedPtg = 5020,
  kIoXlsbCorrupt = 5021,
  /// An XLSB save found a retained pivot or styles part whose model state
  /// changed since load; the writer re-emits those bytes verbatim, so the
  /// save is refused rather than dropping the change.
  kIoXlsbRetainedPartStale = 5022,
  /// Image bytes passed for insertion are not a format the drawing writer
  /// can embed (PNG, JPEG, GIF, BMP), or their header is truncated.
  kIoImageUnsupported = 5023,
  /// A sheet's drawing part could not be parsed, so typed drawing edits on
  /// it are refused rather than rewriting bytes the model cannot see.
  kIoDrawingUnparseable = 5024,
  /// A chart part (or the drawing part reaching it) could not be parsed, so
  /// typed chart reads and edits are refused rather than rewriting bytes the
  /// model cannot see.
  kIoChartUnparseable = 5025,
  /// A chart series or title reference does not resolve to cell ranges of
  /// this workbook (bad syntax, `#REF!`, external, or not a reference).
  kIoChartInvalidReference = 5026,
  /// The chart's kind or shape does not support the requested edit (chartex,
  /// shared axes, pivot-chart series edits).
  kIoChartUnsupportedKind = 5027,
  /// The workbook holds a feature this writer cannot write to xlsb; nothing
  /// is written. The error context's `kind=` names the feature and, for a
  /// sheet-scoped one, `sheet_index=` the sheet.
  kIoFeatureNotWritable = 5028,

  // ===== 6000-6999: Crypto / Security =====
  kSecurityResourceLimit = 6005,

  // ===== 7000-7999: Bindings / C API =====
  kBindingInvalidHandle = 7000,
  kBindingNullPointer = 7001,
  kBindingUtf8EncodingError = 7002,
  kBindingCallbackException = 7003,

  // ===== 8000-8999: Pivot / Advanced =====
  kDataTableInvalid = 8005,
  kAutoFilterInvalid = 8006,

  // ===== 9000-9999: UI integration / Printing =====
  kPrintInvalidArea = 9002,
};

/// Structured error payload returned by every fallible internal API.
///
/// The engine never throws; instead the usual return type is
/// `Expected<T, Error>` (see `expected.h`). `context` carries an optional
/// free-form `key=value` string intended for logging and diagnostics; the
/// top-level `message` must stay concise and in English.
struct Error {
  FormulonErrorCode code = FormulonErrorCode::kInternalError;
  std::string message;
  std::string context;
  const char* file = nullptr;
  int line = 0;
};

/// Builds an `Error` with the given code and message.
inline Error make_error(FormulonErrorCode code, std::string message) {
  Error err;
  err.code = code;
  err.message = std::move(message);
  return err;
}

/// Builds an `Error` with the given code, message and free-form context.
inline Error make_error(FormulonErrorCode code, std::string message, std::string context) {
  Error err;
  err.code = code;
  err.message = std::move(message);
  err.context = std::move(context);
  return err;
}

/// Literal-text overloads, defined out of line so a call site passes two
/// pointers instead of materialising both strings inline.
Error make_error(FormulonErrorCode code, const char* message);
Error make_error(FormulonErrorCode code, const char* message, const char* context);
Error make_error(FormulonErrorCode code, const char* message, std::string context);

/// Returns the textual name of an error code (e.g. `"kParserUnexpectedToken"`).
///
/// Useful for structured logs and diagnostic assertions. The returned pointer
/// references a static string literal with program lifetime.
inline const char* to_cstring(FormulonErrorCode code) {
  switch (code) {
    // General
    case FormulonErrorCode::kOk:
      return "kOk";
    case FormulonErrorCode::kInvalidArgument:
      return "kInvalidArgument";
    case FormulonErrorCode::kNotImplemented:
      return "kNotImplemented";
    case FormulonErrorCode::kOutOfMemory:
      return "kOutOfMemory";
    case FormulonErrorCode::kNotFound:
      return "kNotFound";
    case FormulonErrorCode::kAlreadyExists:
      return "kAlreadyExists";
    case FormulonErrorCode::kInternalError:
      return "kInternalError";
    case FormulonErrorCode::kPreconditionFailed:
      return "kPreconditionFailed";
    case FormulonErrorCode::kSheetIndexOutOfRange:
      return "kSheetIndexOutOfRange";
    case FormulonErrorCode::kInvalidSheetName:
      return "kInvalidSheetName";
    case FormulonErrorCode::kIndexOutOfRange:
      return "kIndexOutOfRange";

    // Parser
    case FormulonErrorCode::kParserUnexpectedToken:
      return "kParserUnexpectedToken";
    case FormulonErrorCode::kParserUnterminatedString:
      return "kParserUnterminatedString";
    case FormulonErrorCode::kParserInvalidNumber:
      return "kParserInvalidNumber";
    case FormulonErrorCode::kParserInvalidReference:
      return "kParserInvalidReference";
    case FormulonErrorCode::kParserInvalidSheetName:
      return "kParserInvalidSheetName";
    case FormulonErrorCode::kParserMismatchedBrackets:
      return "kParserMismatchedBrackets";
    case FormulonErrorCode::kParserExcessiveNesting:
      return "kParserExcessiveNesting";
    case FormulonErrorCode::kParserInvalidArrayLiteral:
      return "kParserInvalidArrayLiteral";
    case FormulonErrorCode::kParserInvalidStructRef:
      return "kParserInvalidStructRef";
    case FormulonErrorCode::kParserBomNotSupported:
      return "kParserBomNotSupported";
    case FormulonErrorCode::kParserTooManyErrors:
      return "kParserTooManyErrors";

    // Evaluator
    case FormulonErrorCode::kEvalPivotMissing:
      return "kEvalPivotMissing";
    case FormulonErrorCode::kEvalPivotInvalid:
      return "kEvalPivotInvalid";

    // Functions
    case FormulonErrorCode::kFunctionOverflow:
      return "kFunctionOverflow";

    // Graph / Recalc
    case FormulonErrorCode::kGraphScheduleFailed:
      return "kGraphScheduleFailed";
    case FormulonErrorCode::kGraphRecalcReentrant:
      return "kGraphRecalcReentrant";

    // I/O
    case FormulonErrorCode::kIoFileNotFound:
      return "kIoFileNotFound";
    case FormulonErrorCode::kIoZipCorrupt:
      return "kIoZipCorrupt";
    case FormulonErrorCode::kIoZipEncrypted:
      return "kIoZipEncrypted";
    case FormulonErrorCode::kIoZipBomb:
      return "kIoZipBomb";
    case FormulonErrorCode::kIoZipSlip:
      return "kIoZipSlip";
    case FormulonErrorCode::kIoXmlParse:
      return "kIoXmlParse";
    case FormulonErrorCode::kIoRelationshipBroken:
      return "kIoRelationshipBroken";
    case FormulonErrorCode::kIoContentTypeInvalid:
      return "kIoContentTypeInvalid";
    case FormulonErrorCode::kIoSheetCorrupt:
      return "kIoSheetCorrupt";
    case FormulonErrorCode::kIoXlsbRecordCorrupt:
      return "kIoXlsbRecordCorrupt";
    case FormulonErrorCode::kIoWriteFailed:
      return "kIoWriteFailed";
    case FormulonErrorCode::kIoXlsbRecordTruncated:
      return "kIoXlsbRecordTruncated";
    case FormulonErrorCode::kIoXlsbUnknownRecord:
      return "kIoXlsbUnknownRecord";
    case FormulonErrorCode::kIoXlsbUnsupportedPtg:
      return "kIoXlsbUnsupportedPtg";
    case FormulonErrorCode::kIoXlsbCorrupt:
      return "kIoXlsbCorrupt";
    case FormulonErrorCode::kIoXlsbRetainedPartStale:
      return "kIoXlsbRetainedPartStale";
    case FormulonErrorCode::kIoImageUnsupported:
      return "kIoImageUnsupported";
    case FormulonErrorCode::kIoDrawingUnparseable:
      return "kIoDrawingUnparseable";
    case FormulonErrorCode::kIoChartUnparseable:
      return "kIoChartUnparseable";
    case FormulonErrorCode::kIoChartInvalidReference:
      return "kIoChartInvalidReference";
    case FormulonErrorCode::kIoChartUnsupportedKind:
      return "kIoChartUnsupportedKind";
    case FormulonErrorCode::kIoFeatureNotWritable:
      return "kIoFeatureNotWritable";

    // Crypto / Security
    case FormulonErrorCode::kSecurityResourceLimit:
      return "kSecurityResourceLimit";

    // Bindings / C API
    case FormulonErrorCode::kBindingInvalidHandle:
      return "kBindingInvalidHandle";
    case FormulonErrorCode::kBindingNullPointer:
      return "kBindingNullPointer";
    case FormulonErrorCode::kBindingUtf8EncodingError:
      return "kBindingUtf8EncodingError";
    case FormulonErrorCode::kBindingCallbackException:
      return "kBindingCallbackException";

    // Pivot / Advanced
    case FormulonErrorCode::kDataTableInvalid:
      return "kDataTableInvalid";
    case FormulonErrorCode::kAutoFilterInvalid:
      return "kAutoFilterInvalid";

    // UI / Printing
    case FormulonErrorCode::kPrintInvalidArea:
      return "kPrintInvalidArea";
  }
  return "kUnknownError";
}

}  // namespace formulon

#endif  // FORMULON_UTILS_ERROR_H_
