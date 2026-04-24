// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the simple integer-only engineering built-ins:
//   * Base conversion (BIN/OCT/HEX <-> DEC).
//   * Bit manipulation (BITAND / BITOR / BITXOR / BITLSHIFT / BITRSHIFT).
//   * Comparators (DELTA / GESTEP).
//
// Excel-observable semantics are the authority here. The base conversion
// family uses fixed-width two's complement on the *input* side (10 digits
// in the source base) and always emits a 10-digit two's complement string
// for negative results; for non-negative results the emission is
// minimum-width unless `places` is supplied.
//
// Signed range per destination base (see backup/plans/02-calc-engine.md
// §2.7.3 and Excel's own documentation):
//   * 2 / "BIN":  -2^9  .. 2^9  - 1   = -512          .. 511
//   * 8 / "OCT":  -2^29 .. 2^29 - 1   = -536870912    .. 536870911
//   * 16 / "HEX": -2^39 .. 2^39 - 1   = -549755813888 .. 549755813887
//
// Bit operations accept 48-bit non-negative integers (0..2^48-1) and use a
// shift magnitude cap of 53 (matching Excel's double-precision mantissa).

#include "eval/builtins/engineering.h"

#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Base conversion helpers
// ---------------------------------------------------------------------------

// Destination base descriptor. Centralises the per-base signed range so the
// DEC2* / *2DEC / *2* functions stay symmetric.
struct BaseSpec {
  int radix;           // 2, 8, or 16
  int max_digits;      // 10 for the three source bases we support
  std::int64_t two_c;  // 2^(bits-per-10-digit-word); 1024 / 2^30 / 2^40
  std::int64_t min_signed;
  std::int64_t max_signed;
};

constexpr BaseSpec kSpecBin{2, 10, 1024LL, -512LL, 511LL};
constexpr BaseSpec kSpecOct{8, 10, 1LL << 30, -(1LL << 29), (1LL << 29) - 1};
constexpr BaseSpec kSpecHex{16, 10, 1LL << 40, -(1LL << 39), (1LL << 39) - 1};

// Returns the numeric value of a single digit in `radix`, or -1 on failure.
// Accepts 0-9 for any radix; uppercase and lowercase A-F for hex.
int digit_value(char c, int radix) {
  int v = -1;
  if (c >= '0' && c <= '9') {
    v = c - '0';
  } else if (radix == 16 && c >= 'A' && c <= 'F') {
    v = 10 + (c - 'A');
  } else if (radix == 16 && c >= 'a' && c <= 'f') {
    v = 10 + (c - 'a');
  }
  if (v < 0 || v >= radix) {
    return -1;
  }
  return v;
}

// Parses a BIN/OCT/HEX digit string, applying two's complement decoding
// when the input has the full 10 digits and the leading digit's sign-bit
// region is set. Empty input yields #NUM!. Invalid digits yield #NUM!.
// Length > 10 yields #NUM!.
Expected<std::int64_t, ErrorCode> decode_digit_string(std::string_view s, const BaseSpec& spec) {
  if (s.empty()) {
    return ErrorCode::Num;
  }
  if (s.size() > static_cast<std::size_t>(spec.max_digits)) {
    return ErrorCode::Num;
  }
  std::uint64_t acc = 0;
  for (char c : s) {
    const int d = digit_value(c, spec.radix);
    if (d < 0) {
      return ErrorCode::Num;
    }
    acc = acc * static_cast<std::uint64_t>(spec.radix) + static_cast<std::uint64_t>(d);
  }
  auto value = static_cast<std::int64_t>(acc);
  // Two's complement decoding only applies at the full 10-digit width. The
  // sign bit is the MSB of that 10-digit word: BIN -> bit 9, OCT -> bit 29,
  // HEX -> bit 39. In each case `>= two_c / 2` iff the sign bit is set.
  if (s.size() == static_cast<std::size_t>(spec.max_digits)) {
    const std::int64_t half = spec.two_c / 2;
    if (value >= half) {
      value -= spec.two_c;
    }
  }
  return value;
}

// Converts a non-negative int64 to a string in `radix`, most-significant
// first. Never produces an empty string: zero renders as "0".
std::string to_digits(std::int64_t value, int radix) {
  if (value == 0) {
    return std::string("0");
  }
  std::string out;
  auto u = static_cast<std::uint64_t>(value);
  while (u > 0) {
    const int d = static_cast<int>(u % static_cast<std::uint64_t>(radix));
    u /= static_cast<std::uint64_t>(radix);
    out.push_back(static_cast<char>(d < 10 ? ('0' + d) : ('A' + (d - 10))));
  }
  for (std::size_t i = 0, j = out.size() - 1; i < j; ++i, --j) {
    std::swap(out[i], out[j]);
  }
  return out;
}

// Renders a signed value as a string in the destination base. Negative
// values always produce a full 10-digit two's complement string; `places`
// is ignored for negatives (Excel does not pad the negative form).
// Positive values are rendered at minimum width and then left-padded with
// '0' to `places`. If the minimum-width representation exceeds `places`,
// yields #NUM!.
Expected<std::string, ErrorCode> encode_base_string(std::int64_t value, const BaseSpec& spec, const int* places_opt) {
  if (value < 0) {
    const std::int64_t two_c_form = value + spec.two_c;
    std::string digits = to_digits(two_c_form, spec.radix);
    if (digits.size() < static_cast<std::size_t>(spec.max_digits)) {
      digits.insert(digits.begin(), static_cast<std::size_t>(spec.max_digits) - digits.size(), '0');
    }
    return digits;
  }
  std::string digits = to_digits(value, spec.radix);
  if (places_opt != nullptr) {
    const int places = *places_opt;
    if (static_cast<int>(digits.size()) > places) {
      return ErrorCode::Num;
    }
    if (static_cast<int>(digits.size()) < places) {
      digits.insert(digits.begin(), static_cast<std::size_t>(places) - digits.size(), '0');
    }
  }
  return digits;
}

// Coerces a `places` argument: must be a finite number, truncated toward
// zero, and in [1, 10]. Out-of-range / non-finite values yield #NUM!.
// Excel rejects Bool directly for `places` with `#VALUE!` — strict-Bool
// rejection mirroring the input-digit side of BIN2*/OCT2*/HEX2*.
Expected<int, ErrorCode> coerce_places(const Value& v) {
  if (v.kind() == ValueKind::Bool) {
    return ErrorCode::Value;
  }
  auto p = coerce_to_number(v);
  if (!p) {
    return p.error();
  }
  const double d = p.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  const double t = std::trunc(d);
  if (t < 1.0 || t > 10.0) {
    return ErrorCode::Num;
  }
  return static_cast<int>(t);
}

// Extracts the input digit string for *2* source-side parsing. For Text,
// returns the view as-is (caller trims nothing -- Excel rejects leading /
// trailing whitespace with #NUM!). For Number and Bool, renders the
// truncated integer form. Empty text yields #NUM!.
Expected<std::string, ErrorCode> input_digit_string(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Text: {
      const std::string_view s = v.as_text();
      if (s.empty()) {
        return ErrorCode::Num;
      }
      return std::string(s);
    }
    case ValueKind::Number: {
      const double d = v.as_number();
      if (std::isnan(d) || std::isinf(d)) {
        return ErrorCode::Num;
      }
      // Truncate toward zero. Negative numeric inputs are not a valid
      // BIN/OCT/HEX "digit string" and are rejected at parse time because
      // decoded strings never carry a sign -- decode_digit_string will
      // reject the '-' character as an invalid digit.
      const double t = std::trunc(d);
      // Use a fixed-width scratch buffer: the magnitude fits in a 64-bit
      // integer (signed range is well below 2^63 for our bases).
      char buf[32];
      std::snprintf(buf, sizeof(buf), "%lld", static_cast<long long>(t));
      return std::string(buf);
    }
    case ValueKind::Bool:
      // Excel rejects a direct Bool argument to BIN2*/OCT2*/HEX2* with
      // `#VALUE!` — mirrors the strict-Bool rejection in DEC2*.
      return ErrorCode::Value;
    case ValueKind::Blank:
      // A blank cell reference coerces to 0 in numeric context, so Excel
      // treats BIN2OCT(blank) as BIN2OCT(0) -> "0" (likewise for the other
      // *2* conversions). Reproduced by IronCalc on calc_tests/Decimal I32.
      return std::string("0");
    default:
      // Array / Ref / Lambda fall through. The eager dispatcher surfaces
      // scalar mismatches as #VALUE! -- mirror that here defensively.
      return ErrorCode::Value;
  }
}

// Dispatches a *2* conversion. Decodes the input from `src` and encodes the
// result to `dst`, honouring an optional `places` argument. The decoded
// signed value must fit in the destination's signed range; otherwise
// surface #NUM!, matching Excel's behaviour for e.g. HEX2BIN("200").
Value convert_bases(const Value* args, std::uint32_t arity, Arena& arena, const BaseSpec& src, const BaseSpec& dst,
                    bool allow_places) {
  auto s = input_digit_string(args[0]);
  if (!s) {
    return Value::error(s.error());
  }
  auto decoded = decode_digit_string(s.value(), src);
  if (!decoded) {
    return Value::error(decoded.error());
  }
  const std::int64_t value = decoded.value();
  if (value < dst.min_signed || value > dst.max_signed) {
    return Value::error(ErrorCode::Num);
  }
  int places = 0;
  const int* places_ptr = nullptr;
  if (allow_places && arity >= 2) {
    auto p = coerce_places(args[1]);
    if (!p) {
      return Value::error(p.error());
    }
    places = p.value();
    places_ptr = &places;
  }
  auto out = encode_base_string(value, dst, places_ptr);
  if (!out) {
    return Value::error(out.error());
  }
  return Value::text(arena.intern(out.value()));
}

// Dispatches a DEC -> * conversion. Takes a numeric input (truncated toward
// zero), range-checks against `dst`, and encodes.
Value convert_from_dec(const Value* args, std::uint32_t arity, Arena& arena, const BaseSpec& dst) {
  // Excel-quirk: DEC2BIN / DEC2OCT / DEC2HEX reject a direct Bool argument
  // with `#VALUE!` rather than coercing TRUE/FALSE to 1/0. Matches EFFECT /
  // NOMINAL's strict-Bool rejection (see `financial_misc.cpp`). Bool inside
  // a range cell would have been flattened earlier; direct scalars only.
  if (args[0].kind() == ValueKind::Bool) {
    return Value::error(ErrorCode::Value);
  }
  auto n = coerce_to_number(args[0]);
  if (!n) {
    return Value::error(n.error());
  }
  const double d = n.value();
  if (std::isnan(d) || std::isinf(d)) {
    return Value::error(ErrorCode::Num);
  }
  const double t = std::trunc(d);
  if (t < static_cast<double>(dst.min_signed) || t > static_cast<double>(dst.max_signed)) {
    return Value::error(ErrorCode::Num);
  }
  const auto value = static_cast<std::int64_t>(t);
  int places = 0;
  const int* places_ptr = nullptr;
  if (arity >= 2) {
    auto p = coerce_places(args[1]);
    if (!p) {
      return Value::error(p.error());
    }
    places = p.value();
    places_ptr = &places;
  }
  auto out = encode_base_string(value, dst, places_ptr);
  if (!out) {
    return Value::error(out.error());
  }
  return Value::text(arena.intern(out.value()));
}

// Dispatches a * -> DEC conversion. Returns the signed integer as a Number.
Value convert_to_dec(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/, const BaseSpec& src) {
  auto s = input_digit_string(args[0]);
  if (!s) {
    return Value::error(s.error());
  }
  auto decoded = decode_digit_string(s.value(), src);
  if (!decoded) {
    return Value::error(decoded.error());
  }
  return Value::number(static_cast<double>(decoded.value()));
}

// ---------------------------------------------------------------------------
// Concrete base-conversion impls
// ---------------------------------------------------------------------------

Value Bin2Dec(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_to_dec(args, arity, arena, kSpecBin);
}
Value Oct2Dec(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_to_dec(args, arity, arena, kSpecOct);
}
Value Hex2Dec(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_to_dec(args, arity, arena, kSpecHex);
}

Value Bin2Oct(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_bases(args, arity, arena, kSpecBin, kSpecOct, /*allow_places=*/true);
}
Value Bin2Hex(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_bases(args, arity, arena, kSpecBin, kSpecHex, /*allow_places=*/true);
}
Value Oct2Bin(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_bases(args, arity, arena, kSpecOct, kSpecBin, /*allow_places=*/true);
}
Value Oct2Hex(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_bases(args, arity, arena, kSpecOct, kSpecHex, /*allow_places=*/true);
}
Value Hex2Bin(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_bases(args, arity, arena, kSpecHex, kSpecBin, /*allow_places=*/true);
}
Value Hex2Oct(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_bases(args, arity, arena, kSpecHex, kSpecOct, /*allow_places=*/true);
}

Value Dec2Bin(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_from_dec(args, arity, arena, kSpecBin);
}
Value Dec2Oct(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_from_dec(args, arity, arena, kSpecOct);
}
Value Dec2Hex(const Value* args, std::uint32_t arity, Arena& arena) {
  return convert_from_dec(args, arity, arena, kSpecHex);
}

// ---------------------------------------------------------------------------
// Bit operations
// ---------------------------------------------------------------------------

// Maximum representable unsigned 48-bit value. Shared by BITAND / BITOR /
// BITXOR / BITLSHIFT (result check) / BITRSHIFT (input check).
constexpr std::uint64_t kBit48Max = (static_cast<std::uint64_t>(1) << 48u) - 1u;
constexpr int kBitShiftMagnitudeMax = 53;

// Coerces a BIT* integer argument: must be a finite integer-valued number
// in [0, 2^48 - 1]. Non-integer (e.g. 12.5) -> #NUM!; non-numeric -> #VALUE!;
// out-of-range -> #NUM!. Excel rejects fractional inputs here rather than
// silently truncating.
Expected<std::uint64_t, ErrorCode> coerce_48_bit_unsigned(const Value& v) {
  auto n = coerce_to_number(v);
  if (!n) {
    return n.error();
  }
  const double d = n.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  if (d != std::trunc(d)) {
    return ErrorCode::Num;
  }
  if (d < 0.0 || d > static_cast<double>(kBit48Max)) {
    return ErrorCode::Num;
  }
  return static_cast<std::uint64_t>(d);
}

// Coerces a shift magnitude: finite number truncated toward zero, in
// [-53, 53]. Non-numeric -> #VALUE!; out-of-range -> #NUM!. Unlike the
// value argument, the shift is Excel-truncated (34.67 -> 34) rather than
// rejected.
Expected<int, ErrorCode> coerce_shift(const Value& v) {
  auto n = coerce_to_number(v);
  if (!n) {
    return n.error();
  }
  const double d = n.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  const double t = std::trunc(d);
  if (t < -static_cast<double>(kBitShiftMagnitudeMax) || t > static_cast<double>(kBitShiftMagnitudeMax)) {
    return ErrorCode::Num;
  }
  return static_cast<int>(t);
}

Value BitAnd(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto a = coerce_48_bit_unsigned(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  auto b = coerce_48_bit_unsigned(args[1]);
  if (!b) {
    return Value::error(b.error());
  }
  return Value::number(static_cast<double>(a.value() & b.value()));
}

Value BitOr(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto a = coerce_48_bit_unsigned(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  auto b = coerce_48_bit_unsigned(args[1]);
  if (!b) {
    return Value::error(b.error());
  }
  return Value::number(static_cast<double>(a.value() | b.value()));
}

Value BitXor(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto a = coerce_48_bit_unsigned(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  auto b = coerce_48_bit_unsigned(args[1]);
  if (!b) {
    return Value::error(b.error());
  }
  return Value::number(static_cast<double>(a.value() ^ b.value()));
}

// Applies a left-then-right or right-then-left shift according to sign.
// Shared between BITLSHIFT and BITRSHIFT so the overflow-check policy stays
// in a single spot: *left* shifts must fit in 48 bits, right shifts simply
// collapse to 0 once the shift consumes every bit.
Expected<std::uint64_t, ErrorCode> apply_bit_shift(std::uint64_t n, int shift_left) {
  if (shift_left == 0) {
    return n;
  }
  if (shift_left > 0) {
    // Left shift. A shift >= 64 in C++ is UB, so short-circuit any shift
    // >= 48 that would be guaranteed to overflow anyway (n <= 2^48-1).
    if (shift_left >= 48 && n != 0) {
      return ErrorCode::Num;
    }
    if (shift_left >= 64) {
      // n must be 0 here (caught above), but keep the guard explicit.
      return static_cast<std::uint64_t>(0);
    }
    const std::uint64_t shifted = n << static_cast<unsigned>(shift_left);
    if (shifted > kBit48Max) {
      return ErrorCode::Num;
    }
    return shifted;
  }
  // Right shift. A shift >= 64 in C++ is UB; any such shift fully drains n.
  const int amount = -shift_left;
  if (amount >= 64) {
    return static_cast<std::uint64_t>(0);
  }
  return n >> static_cast<unsigned>(amount);
}

Value BitLShift(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto n = coerce_48_bit_unsigned(args[0]);
  if (!n) {
    return Value::error(n.error());
  }
  auto shift = coerce_shift(args[1]);
  if (!shift) {
    return Value::error(shift.error());
  }
  // BITLSHIFT interprets positive shift as "to the left".
  auto result = apply_bit_shift(n.value(), shift.value());
  if (!result) {
    return Value::error(result.error());
  }
  return Value::number(static_cast<double>(result.value()));
}

Value BitRShift(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto n = coerce_48_bit_unsigned(args[0]);
  if (!n) {
    return Value::error(n.error());
  }
  auto shift = coerce_shift(args[1]);
  if (!shift) {
    return Value::error(shift.error());
  }
  // BITRSHIFT interprets positive shift as "to the right" -- invert for
  // the shared helper whose convention is "positive = left".
  auto result = apply_bit_shift(n.value(), -shift.value());
  if (!result) {
    return Value::error(result.error());
  }
  return Value::number(static_cast<double>(result.value()));
}

// ---------------------------------------------------------------------------
// DELTA / GESTEP
// ---------------------------------------------------------------------------

Value Delta(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  // Excel-quirk: DELTA rejects a direct Bool argument with `#VALUE!`
  // rather than coercing it to 0/1. Matches Excel 365 / IronCalc oracle.
  if (args[0].kind() == ValueKind::Bool) {
    return Value::error(ErrorCode::Value);
  }
  if (arity >= 2 && args[1].kind() == ValueKind::Bool) {
    return Value::error(ErrorCode::Value);
  }
  auto a = coerce_to_number(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  double b = 0.0;
  if (arity >= 2) {
    auto bv = coerce_to_number(args[1]);
    if (!bv) {
      return Value::error(bv.error());
    }
    b = bv.value();
  }
  return Value::number(a.value() == b ? 1.0 : 0.0);
}

Value Gestep(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  // Excel-quirk: GESTEP rejects a direct Bool argument with `#VALUE!`
  // rather than coercing it to 0/1. Matches Excel 365 / IronCalc oracle.
  if (args[0].kind() == ValueKind::Bool) {
    return Value::error(ErrorCode::Value);
  }
  if (arity >= 2 && args[1].kind() == ValueKind::Bool) {
    return Value::error(ErrorCode::Value);
  }
  auto a = coerce_to_number(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  double step = 0.0;
  if (arity >= 2) {
    auto sv = coerce_to_number(args[1]);
    if (!sv) {
      return Value::error(sv.error());
    }
    step = sv.value();
  }
  return Value::number(a.value() >= step ? 1.0 : 0.0);
}

}  // namespace

void register_engineering_builtins(FunctionRegistry& registry) {
  // Base conversion. The spec pins BIN2DEC / OCT2DEC / HEX2DEC to exactly
  // one argument (no `places`); all others accept an optional `places`.
  registry.register_function(FunctionDef{"BIN2DEC", 1u, 1u, &Bin2Dec});
  registry.register_function(FunctionDef{"BIN2OCT", 1u, 2u, &Bin2Oct});
  registry.register_function(FunctionDef{"BIN2HEX", 1u, 2u, &Bin2Hex});
  registry.register_function(FunctionDef{"OCT2DEC", 1u, 1u, &Oct2Dec});
  registry.register_function(FunctionDef{"OCT2BIN", 1u, 2u, &Oct2Bin});
  registry.register_function(FunctionDef{"OCT2HEX", 1u, 2u, &Oct2Hex});
  registry.register_function(FunctionDef{"HEX2DEC", 1u, 1u, &Hex2Dec});
  registry.register_function(FunctionDef{"HEX2BIN", 1u, 2u, &Hex2Bin});
  registry.register_function(FunctionDef{"HEX2OCT", 1u, 2u, &Hex2Oct});
  registry.register_function(FunctionDef{"DEC2BIN", 1u, 2u, &Dec2Bin});
  registry.register_function(FunctionDef{"DEC2OCT", 1u, 2u, &Dec2Oct});
  registry.register_function(FunctionDef{"DEC2HEX", 1u, 2u, &Dec2Hex});

  // Bit operations.
  registry.register_function(FunctionDef{"BITAND", 2u, 2u, &BitAnd});
  registry.register_function(FunctionDef{"BITOR", 2u, 2u, &BitOr});
  registry.register_function(FunctionDef{"BITXOR", 2u, 2u, &BitXor});
  registry.register_function(FunctionDef{"BITLSHIFT", 2u, 2u, &BitLShift});
  registry.register_function(FunctionDef{"BITRSHIFT", 2u, 2u, &BitRShift});

  // Comparators.
  registry.register_function(FunctionDef{"DELTA", 1u, 2u, &Delta});
  registry.register_function(FunctionDef{"GESTEP", 1u, 2u, &Gestep});
}

}  // namespace eval
}  // namespace formulon
