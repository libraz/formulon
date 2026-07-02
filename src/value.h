// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `Value` is Formulon's scalar atom: the tagged union that the parser,
// evaluator, and Excel error model all sit on top of.
//
// The current scope of this header covers the scalar variants `Blank`,
// `Number`, `Bool`, `Error`, and `Text`, plus `Array` (non-owning pointer
// to an arena-backed `ArrayValue`) and `Lambda` (non-owning pointer to an
// arena-backed `eval::LambdaValue`). All pointer-shaped variants share the
// `Text` lifetime contract: callers must keep the arena alive for as long
// as any `Value` references it. The `Ref` variant still reserves a slot in
// `ValueKind` but has no factory yet — that follows once the cell reference
// representation lands.
//
// `Value` is intentionally trivially copyable so it can be passed freely
// through the evaluator's value stack without heap allocation. The `Array`
// variant preserves this property because its payload is a single 8-byte
// pointer.

#ifndef FORMULON_VALUE_H_
#define FORMULON_VALUE_H_

#include <cstddef>
#include <cstdint>
#include <iterator>
#include <string>
#include <string_view>
#include <type_traits>

namespace formulon {

namespace eval {
struct LambdaValue;
}  // namespace eval

struct ArrayValue;

/// Discriminator tag for every variant a cell may hold.
///
/// The numeric assignments are load-bearing: downstream code (evaluator
/// dispatch tables, bindings) relies on the dense 0..7 range to index
/// jump tables. Do not reorder.
enum class ValueKind : std::uint8_t {
  Blank = 0,
  Number = 1,
  Bool = 2,
  Text = 3,
  Error = 4,
  Array = 5,
  Ref = 6,
  Lambda = 7,
};

/// Excel-visible error codes that surface in cell values.
///
/// This enum is distinct from Formulon's internal `FormulonErrorCode`
/// (see `utils/error.h`): `ErrorCode` represents business-level formula
/// errors that appear inside a cell (`#DIV/0!`, `#VALUE!`, ...), whereas
/// `FormulonErrorCode` represents engine-level failures that propagate
/// through `Expected<T, Error>`.
///
/// The sequential numbering here is an implementation detail of the
/// in-memory tag. It is *not* the OOXML wire code; use `ooxml_code()` for
/// that mapping. The 17 variants below cover:
///
///   * Classic 7 errors (pre-2007): `Null`, `Div0`, `Value`, `Ref`, `Name`,
///     `Num`, `NA`.
///   * Excel 2007+: `GettingData`.
///   * Dynamic array era (2018+): `Spill`, `Calc`.
///   * Linked data types (2019+): `Field`.
///   * 2021+: `Blocked`, `Connect`, `External`.
///   * Excel 365 (2023+): `Busy`, `Python`, plus the catch-all `Unknown`.
enum class ErrorCode : std::uint16_t {
  // Classic 7 (pre-2007).
  Null = 0,
  Div0,
  Value,
  Ref,
  Name,
  Num,
  NA,
  // Excel 2007+.
  GettingData,
  // Dynamic array (2018+).
  Spill,
  Calc,
  // Linked data types (2019+).
  Field,
  // 2021+.
  Blocked,
  Connect,
  External,
  // Excel 365 (2023+).
  Busy,
  Python,
  Unknown,
};

/// Per-`ErrorCode` lookup row consolidating the Excel-visible display
/// name and the OOXML wire code into a single source of truth. The table
/// `kErrorTable` below is indexed directly by the `ErrorCode` ordinal,
/// which is why `ErrorCode` must remain contiguous and start at 0.
struct ErrorCodeInfo {
  /// Tokenised Excel display name (e.g. `"#DIV/0!"`). Static string
  /// literal with program lifetime.
  const char* display_name;
  /// ECMA-376 / [MS-XLSB] wire code (>= 0 for every known code).
  std::int32_t ooxml_code;
};

/// One entry per `ErrorCode`, in enum order. Adding a new `ErrorCode`
/// requires appending a row here — the `static_assert` below catches a
/// missed update at compile time.
inline constexpr ErrorCodeInfo kErrorTable[] = {
    {"#NULL!", 0},          // ErrorCode::Null
    {"#DIV/0!", 7},         // ErrorCode::Div0
    {"#VALUE!", 15},        // ErrorCode::Value
    {"#REF!", 23},          // ErrorCode::Ref
    {"#NAME?", 29},         // ErrorCode::Name
    {"#NUM!", 36},          // ErrorCode::Num
    {"#N/A", 42},           // ErrorCode::NA
    {"#GETTING_DATA", 43},  // ErrorCode::GettingData
    {"#SPILL!", 14},        // ErrorCode::Spill
    {"#CALC!", 13},         // ErrorCode::Calc
    {"#FIELD!", 10},        // ErrorCode::Field
    {"#BLOCKED!", 12},      // ErrorCode::Blocked
    {"#CONNECT!", 11},      // ErrorCode::Connect
    {"#EXTERNAL!", 19},     // ErrorCode::External
    {"#BUSY!", 16},         // ErrorCode::Busy
    {"#PYTHON!", 17},       // ErrorCode::Python
    {"#UNKNOWN!", 9},       // ErrorCode::Unknown
};

// `ErrorCode::Unknown` is the last enumerator; the table length must match
// the enum size exactly so the `kErrorTable[ordinal]` lookup is total.
static_assert(std::size(kErrorTable) == static_cast<std::size_t>(ErrorCode::Unknown) + 1,
              "kErrorTable must have one entry per ErrorCode, in enum order");

/// Returns the lookup row for `e`. Precondition: `e` is a valid enumerator
/// (the in-memory tag is `std::uint16_t` but only the 17 declared codes
/// produce a defined result).
constexpr const ErrorCodeInfo& error_info(ErrorCode e) noexcept {
  return kErrorTable[static_cast<std::size_t>(e)];
}

/// Returns the OOXML wire code for `e`.
///
/// The mapping follows the ECMA-376 / [MS-XLSB] convention.
constexpr std::int32_t ooxml_code(ErrorCode e) noexcept {
  return error_info(e).ooxml_code;
}

/// Inverse of `ooxml_code`: maps a wire code back to its `ErrorCode`.
///
/// Every `kErrorTable` row round-trips through this pair of functions, so
/// a reader and a writer that both go through `ooxml_code` /
/// `error_from_ooxml_code` can never drift apart the way two
/// hand-duplicated switch statements can. Wire codes with no matching row
/// (including genuinely unknown codes) resolve to `ErrorCode::Unknown`.
constexpr ErrorCode error_from_ooxml_code(std::int32_t code) noexcept {
  for (std::size_t i = 0; i < std::size(kErrorTable); ++i) {
    if (kErrorTable[i].ooxml_code == code) {
      return static_cast<ErrorCode>(i);
    }
  }
  return ErrorCode::Unknown;
}

/// Returns the tokenised Excel display name for `e` (e.g. `"#DIV/0!"`).
///
/// The pointer references a static string literal with program lifetime.
constexpr const char* display_name(ErrorCode e) noexcept {
  return error_info(e).display_name;
}

/// Scalar `Value` atom of the Formulon calc engine.
///
/// The scalar variants (`Blank`, `Number`, `Bool`, `Error`, `Text`) carry
/// factories, as do `Array` (payload: non-owning pointer to an arena-backed
/// `ArrayValue`) and `Lambda` (payload: non-owning pointer to an
/// arena-backed `eval::LambdaValue`). The kind query for `Ref` exists but
/// always returns false until that variant is implemented. All factories
/// are `noexcept` and never allocate. The `Text`, `Array`, and `Lambda`
/// payloads are non-owning views into arena storage: the caller is
/// responsible for keeping the underlying storage alive for at least the
/// lifetime of the `Value`.
///
/// Accessors (`as_number()`, `as_boolean()`, `as_error()`, `as_text()`,
/// `as_array()` / `as_array_rows()` / `as_array_cols()` /
/// `as_array_cells()`, `as_lambda()`) are precondition-checked: invoking
/// one on a mismatched kind aborts the process via `FM_CHECK`. Callers
/// must gate access with the corresponding `is_*()` query, or branch on
/// `kind()`.
class Value {
 public:
  /// Builds a `Blank` value. This is the zero-state used by empty cells.
  static Value blank() noexcept {
    Value v;
    v.kind_ = ValueKind::Blank;
    v.data_.number = 0.0;
    return v;
  }

  /// Builds a `Number` value. `v` may be any finite, infinite, or NaN double.
  static Value number(double v) noexcept {
    Value out;
    out.kind_ = ValueKind::Number;
    out.data_.number = v;
    return out;
  }

  /// Builds a `Bool` value.
  static Value boolean(bool v) noexcept {
    Value out;
    out.kind_ = ValueKind::Bool;
    out.data_.boolean = v;
    return out;
  }

  /// Builds an `Error` value carrying the Excel-visible code `c`.
  static Value error(ErrorCode c) noexcept {
    Value out;
    out.kind_ = ValueKind::Error;
    out.data_.error = c;
    return out;
  }

  /// Builds a `Text` value referencing `s`. The caller owns the backing
  /// storage and must keep it alive for the lifetime of the returned value.
  static Value text(std::string_view s) noexcept {
    Value out;
    out.kind_ = ValueKind::Text;
    out.data_.text = s;
    return out;
  }

  /// Builds an `Array` value referencing `arr`. The caller owns the
  /// arena-backed storage (both the `ArrayValue` itself and its `cells`
  /// buffer) and must keep it alive for the lifetime of the returned value,
  /// matching the `Text` lifetime contract.
  static Value array(const ArrayValue* arr) noexcept {
    Value out;
    out.kind_ = ValueKind::Array;
    out.data_.array = arr;
    return out;
  }

  /// Builds a `Lambda` value referencing the closure `lv`. The caller owns
  /// the arena-backed `LambdaValue` (which itself transitively borrows the
  /// parameter array, the body AST, and the captured environment) and must
  /// keep it alive for the lifetime of the returned value, matching the
  /// `Text` / `Array` lifetime contract.
  static Value lambda(const eval::LambdaValue* lv) noexcept {
    Value out;
    out.kind_ = ValueKind::Lambda;
    out.data_.lambda = lv;
    return out;
  }

  /// Returns the discriminator tag for this value.
  ValueKind kind() const noexcept { return kind_; }

  // Kind queries: one per variant. `is_ref()` and `is_lambda()` always
  // return false until those variants are implemented.
  bool is_blank() const noexcept { return kind_ == ValueKind::Blank; }
  bool is_number() const noexcept { return kind_ == ValueKind::Number; }
  bool is_boolean() const noexcept { return kind_ == ValueKind::Bool; }
  bool is_error() const noexcept { return kind_ == ValueKind::Error; }
  bool is_text() const noexcept { return kind_ == ValueKind::Text; }
  bool is_array() const noexcept { return kind_ == ValueKind::Array; }
  bool is_ref() const noexcept { return kind_ == ValueKind::Ref; }
  bool is_lambda() const noexcept { return kind_ == ValueKind::Lambda; }

  /// Returns the numeric payload. Aborts if `kind() != Number`.
  double as_number() const;

  /// Returns the boolean payload. Aborts if `kind() != Bool`.
  bool as_boolean() const;

  /// Returns the error-code payload. Aborts if `kind() != Error`.
  ErrorCode as_error() const;

  /// Returns the text payload as a non-owning view. Aborts if
  /// `kind() != Text`.
  std::string_view as_text() const;

  /// Returns the (non-owning) array payload pointer. Aborts if
  /// `kind() != Array`.
  const ArrayValue* as_array() const;

  /// Returns the array's row count. Aborts if `kind() != Array`.
  std::uint32_t as_array_rows() const;

  /// Returns the array's column count. Aborts if `kind() != Array`.
  std::uint32_t as_array_cols() const;

  /// Returns the array's row-major cell pointer. Aborts if
  /// `kind() != Array`. Indexing: cell `(r, c)` is `cells[r * cols + c]`.
  const Value* as_array_cells() const;

  /// Returns the (non-owning) lambda payload pointer. Aborts if
  /// `kind() != Lambda`.
  const eval::LambdaValue* as_lambda() const;

  /// Debug-formatting helper, not an Excel display string.
  ///
  /// Examples: `"Blank"`, `"Number(42)"`, `"Bool(true)"`,
  /// `"Error(#DIV/0!)"`. The exact number format is not load-bearing and
  /// uses `std::to_string`; Excel-exact formatting is implemented separately.
  std::string debug_to_string() const;

  /// Value-level equality.
  ///
  /// `Blank == Blank`. `Number` values compare via `operator==` on
  /// `double`, which means `NaN != NaN` (aligned with Excel's `#NUM!`
  /// propagation model: the evaluator handles NaN coercion separately).
  /// `Bool` compares bools; `Error` compares `ErrorCode`. Values with
  /// different kinds are never equal.
  friend bool operator==(const Value& a, const Value& b) noexcept;
  friend bool operator!=(const Value& a, const Value& b) noexcept { return !(a == b); }

 private:
  Value() noexcept : kind_(ValueKind::Blank), data_{} {}

  union Payload {
    double number;
    bool boolean;
    ErrorCode error;
    /// Transitional text payload: a non-owning view borrowed from
    /// arena-interned storage. The long-term plan is to replace this with a
    /// `uint32_t text_id` indexing a workbook-scoped SharedStringPool,
    /// which will shrink the union back toward 8 bytes.
    std::string_view text;
    /// Non-owning pointer into arena-allocated `ArrayValue` storage; same
    /// lifetime contract as `text`.
    const ArrayValue* array;
    /// Non-owning pointer into arena-allocated `eval::LambdaValue` storage;
    /// same lifetime contract as `text` / `array`. Holds parameter list,
    /// body AST, and captured `NameEnv` chain for an Excel `LAMBDA` form.
    const eval::LambdaValue* lambda;
    Payload() noexcept : number(0.0) {}
  };

  ValueKind kind_;
  Payload data_;
};

/// Backing storage for `Value::Array`. Allocated in the per-evaluation
/// `Arena` alongside its `cells` array (also arena-allocated). Both pointers
/// must remain valid for as long as any `Value::Array` references them, the
/// same contract as the non-owning `string_view` of `Value::Text`. Cells are
/// stored in row-major order; index `(r, c)` is `cells[r * cols + c]`.
struct ArrayValue {
  std::uint32_t rows;
  std::uint32_t cols;
  const Value* cells;  // arena-owned; row-major; size = rows * cols
};

// `ArrayValue` is created via `Arena::create<>` which forbids destructors.
static_assert(std::is_trivially_destructible_v<ArrayValue>,
              "ArrayValue must be trivially destructible to live in Arena::create");

// Keeping `Value` trivially copyable means the evaluator can pass values by
// value through the VM stack and arg packs without moves or allocations.
// Every variant payload (including the `Array` and `Lambda` 8-byte pointers)
// preserves this property. This invariant must be revisited when `Ref`
// lands.
static_assert(std::is_trivially_copyable_v<Value>, "Value must be trivially copyable");

// The payload union is driven by the 16-byte `string_view` text member
// (and will later be driven by a 16-byte `Reference` payload). The
// `Array` and `Lambda` payloads are 8-byte pointers that fit inside the
// existing budget. With a 1-byte tag and alignment padding the struct
// lands at 24 bytes on every platform Formulon targets.
static_assert(sizeof(Value) <= 24, "Value must fit within 24 bytes");

}  // namespace formulon

#endif  // FORMULON_VALUE_H_
