// Copyright 2026 libraz. Licensed under the MIT License.
//
// `Value` is Formulon's scalar atom: the tagged union that the parser,
// evaluator, and Excel error model all sit on top of. See
// backup/plans/02-calc-engine.md §2.1 for the authoritative specification.
//
// The current scope of this header covers the scalar variants `Blank`,
// `Number`, `Bool`, `Error`, and `Text`, plus `Array`, which is implemented
// as a non-owning pointer to an arena-backed `ArrayValue` (same lifetime
// contract as the `Text` `string_view`: callers must keep the arena alive
// for as long as any `Value::Array` references it). The `Ref` and `Lambda`
// variants still reserve slots in `ValueKind` but do not yet have factories
// or accessors: those will follow once their backing types exist (cell
// reference representation, LAMBDA closures).
//
// `Value` is intentionally trivially copyable so it can be passed freely
// through the evaluator's value stack without heap allocation. The `Array`
// variant preserves this property because its payload is a single 8-byte
// pointer.

#ifndef FORMULON_VALUE_H_
#define FORMULON_VALUE_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <type_traits>

namespace formulon {

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

/// Returns the OOXML wire code for `e`, or `-1` when `e` is out of range.
///
/// The mapping follows the ECMA-376 / [MS-XLSB] convention. See
/// backup/plans/02-calc-engine.md §2.1 for the authoritative table.
constexpr std::int32_t ooxml_code(ErrorCode e) noexcept {
  switch (e) {
    case ErrorCode::Null:
      return 0;
    case ErrorCode::Div0:
      return 7;
    case ErrorCode::Value:
      return 15;
    case ErrorCode::Ref:
      return 23;
    case ErrorCode::Name:
      return 29;
    case ErrorCode::Num:
      return 36;
    case ErrorCode::NA:
      return 42;
    case ErrorCode::GettingData:
      return 43;
    case ErrorCode::Spill:
      return 14;
    case ErrorCode::Calc:
      return 13;
    case ErrorCode::Unknown:
      return 9;
    case ErrorCode::Field:
      return 10;
    case ErrorCode::Connect:
      return 11;
    case ErrorCode::Blocked:
      return 12;
    case ErrorCode::External:
      return 19;
    case ErrorCode::Busy:
      return 16;
    case ErrorCode::Python:
      return 17;
  }
  return -1;
}

/// Returns the tokenised Excel display name for `e` (e.g. `"#DIV/0!"`).
///
/// The pointer references a static string literal with program lifetime.
constexpr const char* display_name(ErrorCode e) noexcept {
  switch (e) {
    case ErrorCode::Null:
      return "#NULL!";
    case ErrorCode::Div0:
      return "#DIV/0!";
    case ErrorCode::Value:
      return "#VALUE!";
    case ErrorCode::Ref:
      return "#REF!";
    case ErrorCode::Name:
      return "#NAME?";
    case ErrorCode::Num:
      return "#NUM!";
    case ErrorCode::NA:
      return "#N/A";
    case ErrorCode::GettingData:
      return "#GETTING_DATA";
    case ErrorCode::Spill:
      return "#SPILL!";
    case ErrorCode::Calc:
      return "#CALC!";
    case ErrorCode::Field:
      return "#FIELD!";
    case ErrorCode::Blocked:
      return "#BLOCKED!";
    case ErrorCode::Connect:
      return "#CONNECT!";
    case ErrorCode::External:
      return "#EXTERNAL!";
    case ErrorCode::Busy:
      return "#BUSY!";
    case ErrorCode::Python:
      return "#PYTHON!";
    case ErrorCode::Unknown:
      return "#UNKNOWN!";
  }
  return "#UNKNOWN!";
}

/// Scalar `Value` atom of the Formulon calc engine.
///
/// The scalar variants (`Blank`, `Number`, `Bool`, `Error`, `Text`) carry
/// factories, as does `Array` (whose payload is a non-owning pointer to an
/// arena-backed `ArrayValue`). The kind queries for `Ref`/`Lambda` exist but
/// always return false until those variants are implemented. All factories
/// are `noexcept` and never allocate. The `Text` and `Array` payloads are
/// non-owning views into arena storage: the caller is responsible for
/// keeping the underlying storage alive for at least the lifetime of the
/// `Value`.
///
/// Accessors (`as_number()`, `as_boolean()`, `as_error()`, `as_text()`,
/// `as_array()` / `as_array_rows()` / `as_array_cols()` /
/// `as_array_cells()`) are precondition-checked: invoking one on a
/// mismatched kind aborts the process via `FM_CHECK`. Callers must gate
/// access with the corresponding `is_*()` query, or branch on `kind()`.
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
    /// arena-interned storage. The long-term plan (per
    /// backup/plans/02-calc-engine.md §2.1) is to replace this with a
    /// `uint32_t text_id` indexing a workbook-scoped SharedStringPool,
    /// which will shrink the union back toward 8 bytes.
    std::string_view text;
    /// Non-owning pointer into arena-allocated `ArrayValue` storage; same
    /// lifetime contract as `text`.
    const ArrayValue* array;
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
// Every variant payload (including the `Array` 8-byte pointer) preserves
// this property. This invariant must be revisited when `Ref`/`Lambda` land.
static_assert(std::is_trivially_copyable_v<Value>, "Value must be trivially copyable");

// The payload union is driven by the 16-byte `string_view` text member
// (and will later be driven by a 16-byte `Reference` payload, see
// backup/plans/02-calc-engine.md §2.1). The `Array` payload is an 8-byte
// pointer that fits inside the existing budget. With a 1-byte tag and
// alignment padding the struct lands at 24 bytes on every platform
// Formulon targets.
static_assert(sizeof(Value) <= 24, "Value must fit within 24 bytes");

}  // namespace formulon

#endif  // FORMULON_VALUE_H_
