// Copyright 2026 libraz. Licensed under the MIT License.
//
// `Value` is Formulon's scalar atom: the tagged union that the parser,
// evaluator, and Excel error model all sit on top of. See
// backup/plans/02-calc-engine.md §2.1 for the authoritative specification.
//
// The current scope of this header covers the scalar variants `Blank`,
// `Number`, `Bool`, `Error`, and `Text`. The `Array`, `Ref`, and `Lambda`
// variants reserve slots in `ValueKind` but do not yet have factories or
// accessors: those will follow once their backing types exist
// (`ArrayValue`, cell reference representation, LAMBDA closures).
//
// `Value` is intentionally trivially copyable so it can be passed freely
// through the evaluator's value stack without heap allocation. This
// invariant must be revisited once the non-scalar variants land.

#ifndef FORMULON_VALUE_H_
#define FORMULON_VALUE_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <type_traits>

namespace formulon {

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
/// factories; the kind queries for `Array`/`Ref`/`Lambda` exist but always
/// return false until those variants are implemented. All factories are
/// `noexcept` and never allocate. The `Text` payload is a non-owning
/// `string_view`: the caller is responsible for keeping the underlying
/// storage alive for at least the lifetime of the `Value`.
///
/// Accessors (`as_number()`, `as_boolean()`, `as_error()`, `as_text()`) are
/// precondition-checked: invoking one on a mismatched kind aborts the
/// process via `FM_CHECK`. Callers must gate access with the corresponding
/// `is_*()` query, or branch on `kind()`.
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

  /// Returns the discriminator tag for this value.
  ValueKind kind() const noexcept { return kind_; }

  // Kind queries: one per variant. Queries for not-yet-implemented variants
  // (`Text`, `Array`, `Ref`, `Lambda`) always return false for now.
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
    Payload() noexcept : number(0.0) {}
  };

  ValueKind kind_;
  Payload data_;
};

// `Value` is currently scalar-only. Keeping it trivially copyable means the
// evaluator can pass values by value through the VM stack and arg packs
// without moves or allocations. This invariant must be revisited when
// `Text`/`Array`/`Ref`/`Lambda` land.
static_assert(std::is_trivially_copyable_v<Value>, "Value is scalar-only and must be trivially copyable");

// The payload union is now driven by the 16-byte `string_view` text
// member (and will later be driven by a 16-byte `Reference` payload, see
// backup/plans/02-calc-engine.md §2.1). With a 1-byte tag and 7 bytes of
// alignment padding the struct lands at 24 bytes on every platform
// Formulon targets.
static_assert(sizeof(Value) <= 24, "Value must fit within 24 bytes");

}  // namespace formulon

#endif  // FORMULON_VALUE_H_
