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
  /// Pivot evaluator was handed a `PivotTable` whose `pivot_cache_id` does
  /// not match the supplied `PivotCache`. Indicates an unresolved cache
  /// binding at evaluation time.
  kEvalPivotMissing = 2009,
  /// Pivot evaluator detected a structurally invalid table — most commonly
  /// a `PivotDataField::field_index` that points outside
  /// `PivotCache::fields()`.
  kEvalPivotInvalid = 2010,
  // --- AST -> ByteCode compiler (2050-2069) -------------------------------
  /// Generic compile failure that does not match a more specific code below.
  kVmCompileFailed = 2050,
  /// The compiler encountered an AST node it cannot lower (e.g. an
  /// `ErrorPlaceholder` from panic-mode parser recovery, or a node kind
  /// not yet supported by the backend).
  kVmUnsupportedNode = 2051,
  /// The constants pool has overflowed its 24-bit operand budget (more than
  /// 2^24 distinct constants in a single bytecode body).
  kVmConstPoolOverflow = 2052,
  /// The names pool has overflowed its 24-bit operand budget.
  kVmNamePoolOverflow = 2053,
  /// The instruction stream has overflowed its 32-bit operand budget (jump
  /// targets > 2^32 instructions).
  kVmInstructionLimit = 2054,
  /// A `LetBinding` references more local slots than the compiler can
  /// encode in 24 bits (more than 2^24 LET-bound names in a single body).
  kVmLetSlotOverflow = 2055,
  /// A `Lambda` declares more parameters than the operand encoding allows
  /// (more than 2^16 params).
  kVmLambdaParamOverflow = 2056,
  /// The bytecode body has no instructions; the VM cannot decide what to
  /// return. Indicates a compiler bug, since `compile()` always emits at
  /// least a `Return`.
  kVmEmptyBytecode = 2057,
  /// An opcode that pops `N` operands found fewer than `N` values on the
  /// operand stack.
  kVmStackUnderflow = 2058,
  /// The operand stack grew past the VM's hard cap. Defends against runaway
  /// `Union` / pathological array literals in handcrafted bytecode.
  kVmStackOverflow = 2059,
  /// An instruction word carries an opcode value outside the declared
  /// `OpCode` enum. Indicates corrupted or hand-rolled bytecode.
  kVmInvalidOpcode = 2060,
  /// A `Jump` / `JumpIfFalse` target points outside the current code stream.
  kVmInvalidJumpTarget = 2061,
  /// `LoadLet` referenced a slot that no `StoreLet` has populated yet.
  kVmLetSlotMissing = 2062,
  /// `CallLambda` was invoked with an argument count that does not satisfy
  /// the closure's `[required..param_count]` arity range.
  kVmLambdaArityMismatch = 2063,
  /// The bytecode optimiser failed to lower a pass (constant fold / name
  /// inline / range canonicalise / branch hoist) without producing
  /// behaviour-preserving output. The raw input bytecode is returned
  /// unchanged whenever this is observed; the error is reserved for hard
  /// failures (e.g. constants-pool overflow when re-pooling a folded
  /// result). The optimiser never invents Excel-visible faults.
  kVmOptimizerFailed = 2064,

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
  /// `recalc_parallel` was invoked while another `recalc_parallel` call on
  /// the same thread was still in flight. The scheduler does not support
  /// nested recalc; callers (typically a UDF that re-enters the engine)
  /// must avoid triggering recalc from inside an evaluator callback.
  kGraphRecalcReentrant = 4005,
  /// The iterative solver returned early because the user-supplied
  /// progress callback requested cancellation. The cell store is left in
  /// its current partially-converged state and the caller should treat
  /// the result as "not converged".
  kGraphIterationAborted = 4006,

  // ===== 5000-5999: I/O (OOXML / XLSB / CSV) =====
  kIoFileNotFound = 5000,
  kIoFilePermission = 5001,
  kIoFileTooLarge = 5002,
  kIoZipCorrupt = 5003,
  kIoZipEncrypted = 5004,
  kIoZipBomb = 5005,
  kIoZipSlip = 5006,
  kIoXmlParse = 5007,
  /// Allocated but not currently produced. The XML readers never expand a
  /// `<!DOCTYPE>` internal subset or a custom entity: pugixml is configured
  /// without `parse_doctype`, and the SAX reader decodes only the five
  /// predefined entity references and numeric character references. A
  /// document carrying either construct is therefore parsed without an
  /// entity-expansion step rather than rejected with a dedicated code, so
  /// nothing reaches a site that would build these. The slots stay allocated
  /// (not removed, not renumbered) because the C API promises numeric
  /// identity with this enum; a reader that later refuses these constructs
  /// outright should claim them and drop these notes.
  kIoXmlDoctype = 5008,
  /// Allocated but not currently produced. See `kIoXmlDoctype`.
  kIoXmlEntityExplosion = 5009,
  kIoRelationshipBroken = 5010,
  kIoContentTypeInvalid = 5011,
  kIoSheetCorrupt = 5012,
  kIoUnsupportedVariant = 5013,
  kIoXlsbRecordCorrupt = 5014,
  /// Reserved for a future CSV reader/writer (not yet implemented: no
  /// `src/io/csv_*`, CLI option, or test currently exercises this code).
  /// Kept allocated rather than removed or renumbered so any code already
  /// persisting or comparing this numeric value is not silently broken;
  /// the 5000-5999 I/O range documented in `CLAUDE.md` reserves this slot
  /// for CSV specifically.
  kIoCsvEncodingDetect = 5015,
  kIoWriteFailed = 5016,
  kIoCalcChainMismatch = 5017,
  kIoXlsbRecordTruncated = 5018,
  kIoXlsbUnknownRecord = 5019,
  kIoXlsbUnsupportedPtg = 5020,
  kIoXlsbCorrupt = 5021,

  // ===== 6000-6999: Crypto / Security =====
  /// Allocated but not currently produced: Formulon does not decrypt
  /// password-protected packages, so there is no key-derivation or
  /// verifier step that could fail. An encrypted package is refused at the
  /// archive boundary with `kIoZipEncrypted` instead, which is the code a
  /// binding must branch on to offer a password prompt. These five slots
  /// stay allocated (not removed, not renumbered) because the C API
  /// promises numeric identity with this enum; a decryption path added
  /// later should claim them and drop these notes.
  kCryptoAgileNotSupported = 6000,
  /// Allocated but not currently produced. See `kCryptoAgileNotSupported`.
  kCryptoStandardNotSupported = 6001,
  /// Allocated but not currently produced. See `kCryptoAgileNotSupported`.
  kCryptoBadPassword = 6002,
  /// Allocated but not currently produced. See `kCryptoAgileNotSupported`.
  kCryptoHashMismatch = 6003,
  /// Allocated but not currently produced. See `kCryptoAgileNotSupported`.
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
  kCliInvalidArgs = 7005,
  kCliFileNotFound = 7006,
  kCliOutputFailed = 7007,

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
  /// The physical page count of a pagination request exceeds
  /// `kMaxPaginationPages`. A page grid may be split by one manual break per
  /// grid row and per grid column, so a crafted break configuration reaches a
  /// count no 32-bit total can hold; pagination reports this instead of
  /// returning a truncated count.
  kPrintPageCountOverflow = 9005,

  // ----- Workbook structural mutation (5050-5069) -----
  // Reuse the I/O band: sheet name validation, sheet rearrangement, and
  // defined-name editing all sit at the workbook-level mutation surface,
  // which is logically adjacent to the OOXML (de)serialisation that owns
  // the same metadata.
  /// Sheet index passed to a structural mutation (`rename_sheet`,
  /// `remove_sheet`, `move_sheet`) is `>= sheet_count()`.
  kSheetIndexOutOfRange = 5050,
  /// Sheet name fails structural validation: malformed UTF-8, empty, longer
  /// than 31 UTF-16 units, contains a forbidden character (`: \ / ? * [ ]`),
  /// or collides under locale-independent Unicode simple case folding with
  /// an existing sheet's name. This is not a claim of full Excel-equivalence.
  kInvalidSheetName = 5051,
  /// `remove_sheet` was invoked on a workbook that has only one sheet
  /// remaining. Excel's UI rejects the same operation; we mirror it so
  /// `save()` cannot land in the empty-sheet-list state Excel itself
  /// refuses to open.
  kCannotRemoveLastSheet = 5052,
  /// Appending a sheet would push `sheet_count()` past `Workbook::kMaxSheets`.
  /// The dependency graph identifies a cell by a 16-bit sheet id, so a
  /// workbook holding more sheets than that could not address the excess
  /// ones without aliasing an existing sheet.
  kSheetCountLimitExceeded = 5053,
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
    case FormulonErrorCode::kEvalPivotMissing:
      return "kEvalPivotMissing";
    case FormulonErrorCode::kEvalPivotInvalid:
      return "kEvalPivotInvalid";
    case FormulonErrorCode::kVmCompileFailed:
      return "kVmCompileFailed";
    case FormulonErrorCode::kVmUnsupportedNode:
      return "kVmUnsupportedNode";
    case FormulonErrorCode::kVmConstPoolOverflow:
      return "kVmConstPoolOverflow";
    case FormulonErrorCode::kVmNamePoolOverflow:
      return "kVmNamePoolOverflow";
    case FormulonErrorCode::kVmInstructionLimit:
      return "kVmInstructionLimit";
    case FormulonErrorCode::kVmLetSlotOverflow:
      return "kVmLetSlotOverflow";
    case FormulonErrorCode::kVmLambdaParamOverflow:
      return "kVmLambdaParamOverflow";
    case FormulonErrorCode::kVmEmptyBytecode:
      return "kVmEmptyBytecode";
    case FormulonErrorCode::kVmStackUnderflow:
      return "kVmStackUnderflow";
    case FormulonErrorCode::kVmStackOverflow:
      return "kVmStackOverflow";
    case FormulonErrorCode::kVmInvalidOpcode:
      return "kVmInvalidOpcode";
    case FormulonErrorCode::kVmInvalidJumpTarget:
      return "kVmInvalidJumpTarget";
    case FormulonErrorCode::kVmLetSlotMissing:
      return "kVmLetSlotMissing";
    case FormulonErrorCode::kVmLambdaArityMismatch:
      return "kVmLambdaArityMismatch";
    case FormulonErrorCode::kVmOptimizerFailed:
      return "kVmOptimizerFailed";

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
    case FormulonErrorCode::kGraphRecalcReentrant:
      return "kGraphRecalcReentrant";
    case FormulonErrorCode::kGraphIterationAborted:
      return "kGraphIterationAborted";

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
    case FormulonErrorCode::kIoXlsbRecordTruncated:
      return "kIoXlsbRecordTruncated";
    case FormulonErrorCode::kIoXlsbUnknownRecord:
      return "kIoXlsbUnknownRecord";
    case FormulonErrorCode::kIoXlsbUnsupportedPtg:
      return "kIoXlsbUnsupportedPtg";
    case FormulonErrorCode::kIoXlsbCorrupt:
      return "kIoXlsbCorrupt";

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
    case FormulonErrorCode::kCliInvalidArgs:
      return "kCliInvalidArgs";
    case FormulonErrorCode::kCliFileNotFound:
      return "kCliFileNotFound";
    case FormulonErrorCode::kCliOutputFailed:
      return "kCliOutputFailed";

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
    case FormulonErrorCode::kPrintPageCountOverflow:
      return "kPrintPageCountOverflow";
    case FormulonErrorCode::kUiViewStateInvalid:
      return "kUiViewStateInvalid";
    case FormulonErrorCode::kUiSnapshotFailed:
      return "kUiSnapshotFailed";

    // Workbook structural mutation
    case FormulonErrorCode::kSheetIndexOutOfRange:
      return "kSheetIndexOutOfRange";
    case FormulonErrorCode::kInvalidSheetName:
      return "kInvalidSheetName";
    case FormulonErrorCode::kCannotRemoveLastSheet:
      return "kCannotRemoveLastSheet";
    case FormulonErrorCode::kSheetCountLimitExceeded:
      return "kSheetCountLimitExceeded";
  }
  return "kUnknownError";
}

}  // namespace formulon

#endif  // FORMULON_UTILS_ERROR_H_
