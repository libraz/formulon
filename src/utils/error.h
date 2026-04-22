// Copyright 2026 libraz. Licensed under the MIT License.
//
// Formulon internal error representation.
//
// `FormulonErrorCode` partitions the 0-9999 space across 10 modules; see
// backup/plans/23-error-codes.md for the authoritative specification.
// `Error` is the payload type used by `Expected<T, Error>` throughout the
// engine. It is distinct from the Excel-visible `ErrorCode` (see 02 §2.1),
// which represents formula-level business errors like `#DIV/0!`.

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
/// truth for the C API (`fm_error_t::code`) and the WASM / Python bindings.
enum class FormulonErrorCode : int32_t {
  // ===== 0-999: General =====
  kOk = 0,
  kUnknownError = 1,
  kInvalidArgument = 2,
  kNotImplemented = 3,
  kOutOfMemory = 4,
  kCancelRequested = 5,
  kNotFound = 6,
  kAlreadyExists = 7,
  kInternalError = 8,
  kPreconditionFailed = 9,

  // ===== 1000-1999: Parser / Tokenizer =====
  kParserUnexpectedToken = 1000,
  kParserUnterminatedString = 1001,
  kParserUnterminatedComment = 1002,
  kParserInvalidNumber = 1003,
  kParserInvalidReference = 1004,
  kParserInvalidSheetName = 1005,
  kParserMismatchedBrackets = 1006,
  kParserExcessiveNesting = 1007,
  kParserUnknownFunction = 1008,
  kParserInvalidArrayLiteral = 1009,
  kParserInvalidStructRef = 1010,
  kParserInvalidR1C1 = 1011,
  kParserInvalidDotNotation = 1012,
  kParserBomNotSupported = 1013,
  kParserTooManyErrors = 1014,

  // ===== 2000-2999: Evaluator / VM =====
  kEvalStackOverflow = 2000,
  kEvalCircularReference = 2001,
  kEvalLambdaArityMismatch = 2002,
  kEvalInvalidByteCode = 2003,
  kEvalTimeLimit = 2004,
  kEvalMemoryLimit = 2005,
  kEvalInvalidReference = 2006,
  kEvalLambdaNotFound = 2007,
  kEvalSpillCollision = 2008,

  // ===== 3000-3999: Functions =====
  kFnNotRegistered = 3000,
  kFnWrongArgCount = 3001,
  kFnArgTypeMismatch = 3002,
  kFnDomainError = 3003,
  kFnOverflow = 3004,
  kFnUnderflow = 3005,
  kFnRegexCompile = 3006,
  kFnRegexMatchLimit = 3007,
  kFnHostNotAvailable = 3008,

  // ===== 4000-4999: Dependency graph / Recalc =====
  kGraphCycleDetected = 4000,
  kGraphIterativeDiverged = 4001,
  kGraphInvalidNodeRef = 4002,
  kGraphScheduleFailed = 4003,
  kGraphThreadPoolError = 4004,

  // ===== 5000-5999: I/O (OOXML / XLSB / CSV) =====
  kIoFileNotFound = 5000,
  kIoFilePermission = 5001,
  kIoFileTooLarge = 5002,
  kIoZipCorrupt = 5003,
  kIoZipEncrypted = 5004,
  kIoZipBomb = 5005,
  kIoZipSlip = 5006,
  kIoXmlParse = 5007,
  kIoXmlDoctype = 5008,
  kIoXmlEntityExplosion = 5009,
  kIoRelationshipBroken = 5010,
  kIoContentTypeInvalid = 5011,
  kIoSheetCorrupt = 5012,
  kIoUnsupportedVariant = 5013,
  kIoXlsbRecordCorrupt = 5014,
  kIoCsvEncodingDetect = 5015,
  kIoWriteFailed = 5016,
  kIoCalcChainMismatch = 5017,

  // ===== 6000-6999: Crypto / Security =====
  kCryptoAgileNotSupported = 6000,
  kCryptoStandardNotSupported = 6001,
  kCryptoBadPassword = 6002,
  kCryptoHashMismatch = 6003,
  kCryptoKeyDerivationFailed = 6004,
  kSecResourceLimit = 6005,
  kSecExternalNotAllowed = 6006,
  kSecPolicyBlocked = 6007,

  // ===== 7000-7999: Bindings / C API =====
  kBindingInvalidHandle = 7000,
  kBindingNullPointer = 7001,
  kBindingUtf8EncodingError = 7002,
  kBindingCallbackException = 7003,
  kBindingWasmInitFailed = 7004,

  // ===== 8000-8999: Pivot / Advanced =====
  kPivotSourceInvalid = 8000,
  kPivotFieldNotFound = 8001,
  kPivotAggregationFailed = 8002,
  kSlicerNotConnected = 8003,
  kSparklineInvalid = 8004,
  kDataTableInvalid = 8005,
  kAutoFilterInvalid = 8006,

  // ===== 9000-9999: UI integration / Printing =====
  kPrintFontMetricsMissing = 9000,
  kPrintLayoutConvergence = 9001,
  kPrintInvalidArea = 9002,
  kUiViewStateInvalid = 9003,
  kUiSnapshotFailed = 9004,
};

/// Structured error payload returned by every fallible internal API.
///
/// The engine never throws; instead the usual return type is
/// `Expected<T, Error>` (see `expected.h`). `context` carries an optional
/// free-form `key=value` string intended for logging and diagnostics; the
/// top-level `message` must stay concise and in English.
struct Error {
  FormulonErrorCode code = FormulonErrorCode::kUnknownError;
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

/// Returns the textual name of an error code (e.g. `"kParserUnexpectedToken"`).
///
/// Useful for structured logs and diagnostic assertions. The returned pointer
/// references a static string literal with program lifetime.
inline const char* to_cstring(FormulonErrorCode code) {
  switch (code) {
    // General
    case FormulonErrorCode::kOk:
      return "kOk";
    case FormulonErrorCode::kUnknownError:
      return "kUnknownError";
    case FormulonErrorCode::kInvalidArgument:
      return "kInvalidArgument";
    case FormulonErrorCode::kNotImplemented:
      return "kNotImplemented";
    case FormulonErrorCode::kOutOfMemory:
      return "kOutOfMemory";
    case FormulonErrorCode::kCancelRequested:
      return "kCancelRequested";
    case FormulonErrorCode::kNotFound:
      return "kNotFound";
    case FormulonErrorCode::kAlreadyExists:
      return "kAlreadyExists";
    case FormulonErrorCode::kInternalError:
      return "kInternalError";
    case FormulonErrorCode::kPreconditionFailed:
      return "kPreconditionFailed";

    // Parser
    case FormulonErrorCode::kParserUnexpectedToken:
      return "kParserUnexpectedToken";
    case FormulonErrorCode::kParserUnterminatedString:
      return "kParserUnterminatedString";
    case FormulonErrorCode::kParserUnterminatedComment:
      return "kParserUnterminatedComment";
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
    case FormulonErrorCode::kParserUnknownFunction:
      return "kParserUnknownFunction";
    case FormulonErrorCode::kParserInvalidArrayLiteral:
      return "kParserInvalidArrayLiteral";
    case FormulonErrorCode::kParserInvalidStructRef:
      return "kParserInvalidStructRef";
    case FormulonErrorCode::kParserInvalidR1C1:
      return "kParserInvalidR1C1";
    case FormulonErrorCode::kParserInvalidDotNotation:
      return "kParserInvalidDotNotation";
    case FormulonErrorCode::kParserBomNotSupported:
      return "kParserBomNotSupported";
    case FormulonErrorCode::kParserTooManyErrors:
      return "kParserTooManyErrors";

    // Evaluator
    case FormulonErrorCode::kEvalStackOverflow:
      return "kEvalStackOverflow";
    case FormulonErrorCode::kEvalCircularReference:
      return "kEvalCircularReference";
    case FormulonErrorCode::kEvalLambdaArityMismatch:
      return "kEvalLambdaArityMismatch";
    case FormulonErrorCode::kEvalInvalidByteCode:
      return "kEvalInvalidByteCode";
    case FormulonErrorCode::kEvalTimeLimit:
      return "kEvalTimeLimit";
    case FormulonErrorCode::kEvalMemoryLimit:
      return "kEvalMemoryLimit";
    case FormulonErrorCode::kEvalInvalidReference:
      return "kEvalInvalidReference";
    case FormulonErrorCode::kEvalLambdaNotFound:
      return "kEvalLambdaNotFound";
    case FormulonErrorCode::kEvalSpillCollision:
      return "kEvalSpillCollision";

    // Functions
    case FormulonErrorCode::kFnNotRegistered:
      return "kFnNotRegistered";
    case FormulonErrorCode::kFnWrongArgCount:
      return "kFnWrongArgCount";
    case FormulonErrorCode::kFnArgTypeMismatch:
      return "kFnArgTypeMismatch";
    case FormulonErrorCode::kFnDomainError:
      return "kFnDomainError";
    case FormulonErrorCode::kFnOverflow:
      return "kFnOverflow";
    case FormulonErrorCode::kFnUnderflow:
      return "kFnUnderflow";
    case FormulonErrorCode::kFnRegexCompile:
      return "kFnRegexCompile";
    case FormulonErrorCode::kFnRegexMatchLimit:
      return "kFnRegexMatchLimit";
    case FormulonErrorCode::kFnHostNotAvailable:
      return "kFnHostNotAvailable";

    // Graph / Recalc
    case FormulonErrorCode::kGraphCycleDetected:
      return "kGraphCycleDetected";
    case FormulonErrorCode::kGraphIterativeDiverged:
      return "kGraphIterativeDiverged";
    case FormulonErrorCode::kGraphInvalidNodeRef:
      return "kGraphInvalidNodeRef";
    case FormulonErrorCode::kGraphScheduleFailed:
      return "kGraphScheduleFailed";
    case FormulonErrorCode::kGraphThreadPoolError:
      return "kGraphThreadPoolError";

    // I/O
    case FormulonErrorCode::kIoFileNotFound:
      return "kIoFileNotFound";
    case FormulonErrorCode::kIoFilePermission:
      return "kIoFilePermission";
    case FormulonErrorCode::kIoFileTooLarge:
      return "kIoFileTooLarge";
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
    case FormulonErrorCode::kIoXmlDoctype:
      return "kIoXmlDoctype";
    case FormulonErrorCode::kIoXmlEntityExplosion:
      return "kIoXmlEntityExplosion";
    case FormulonErrorCode::kIoRelationshipBroken:
      return "kIoRelationshipBroken";
    case FormulonErrorCode::kIoContentTypeInvalid:
      return "kIoContentTypeInvalid";
    case FormulonErrorCode::kIoSheetCorrupt:
      return "kIoSheetCorrupt";
    case FormulonErrorCode::kIoUnsupportedVariant:
      return "kIoUnsupportedVariant";
    case FormulonErrorCode::kIoXlsbRecordCorrupt:
      return "kIoXlsbRecordCorrupt";
    case FormulonErrorCode::kIoCsvEncodingDetect:
      return "kIoCsvEncodingDetect";
    case FormulonErrorCode::kIoWriteFailed:
      return "kIoWriteFailed";
    case FormulonErrorCode::kIoCalcChainMismatch:
      return "kIoCalcChainMismatch";

    // Crypto / Security
    case FormulonErrorCode::kCryptoAgileNotSupported:
      return "kCryptoAgileNotSupported";
    case FormulonErrorCode::kCryptoStandardNotSupported:
      return "kCryptoStandardNotSupported";
    case FormulonErrorCode::kCryptoBadPassword:
      return "kCryptoBadPassword";
    case FormulonErrorCode::kCryptoHashMismatch:
      return "kCryptoHashMismatch";
    case FormulonErrorCode::kCryptoKeyDerivationFailed:
      return "kCryptoKeyDerivationFailed";
    case FormulonErrorCode::kSecResourceLimit:
      return "kSecResourceLimit";
    case FormulonErrorCode::kSecExternalNotAllowed:
      return "kSecExternalNotAllowed";
    case FormulonErrorCode::kSecPolicyBlocked:
      return "kSecPolicyBlocked";

    // Bindings / C API
    case FormulonErrorCode::kBindingInvalidHandle:
      return "kBindingInvalidHandle";
    case FormulonErrorCode::kBindingNullPointer:
      return "kBindingNullPointer";
    case FormulonErrorCode::kBindingUtf8EncodingError:
      return "kBindingUtf8EncodingError";
    case FormulonErrorCode::kBindingCallbackException:
      return "kBindingCallbackException";
    case FormulonErrorCode::kBindingWasmInitFailed:
      return "kBindingWasmInitFailed";

    // Pivot / Advanced
    case FormulonErrorCode::kPivotSourceInvalid:
      return "kPivotSourceInvalid";
    case FormulonErrorCode::kPivotFieldNotFound:
      return "kPivotFieldNotFound";
    case FormulonErrorCode::kPivotAggregationFailed:
      return "kPivotAggregationFailed";
    case FormulonErrorCode::kSlicerNotConnected:
      return "kSlicerNotConnected";
    case FormulonErrorCode::kSparklineInvalid:
      return "kSparklineInvalid";
    case FormulonErrorCode::kDataTableInvalid:
      return "kDataTableInvalid";
    case FormulonErrorCode::kAutoFilterInvalid:
      return "kAutoFilterInvalid";

    // UI / Printing
    case FormulonErrorCode::kPrintFontMetricsMissing:
      return "kPrintFontMetricsMissing";
    case FormulonErrorCode::kPrintLayoutConvergence:
      return "kPrintLayoutConvergence";
    case FormulonErrorCode::kPrintInvalidArea:
      return "kPrintInvalidArea";
    case FormulonErrorCode::kUiViewStateInvalid:
      return "kUiViewStateInvalid";
    case FormulonErrorCode::kUiSnapshotFailed:
      return "kUiSnapshotFailed";
  }
  return "kUnknownError";
}

}  // namespace formulon

#endif  // FORMULON_UTILS_ERROR_H_
