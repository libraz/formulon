//
// MS-XLSB record-emission primitives — symmetric inverse of
// `io/xlsb/record.{h,cpp}`. Every helper here writes the exact byte
// sequence that the matching reader helper consumes, so a record
// emitted by `emit_record` round-trips through `read_record` to the
// same `(type, payload)` pair.
//
// `io/xlsb/writer.cpp` composes these primitives to build:
//
//   * variable-length record headers (1- or 2-byte type, 1..4-byte
//     payload size with MSB-set continuation, per [MS-XLSB] §2.1.4);
//   * little-endian fixed-width unsigned integers;
//   * `XLWideString` (length-prefixed UTF-16LE);
//   * `XLNullableWideString` (XLWideString + the `0xFFFFFFFF` sentinel
//     for null);
//   * `RkNumber` (integer or scaled IEEE 754 double, [MS-XLSB] §2.5.121).
//
// Design references:
//   * [MS-XLSB] §2.1 (Record framing) and §2.5.121 (RkNumber)

#ifndef FORMULON_IO_XLSB_RECORD_WRITER_H_
#define FORMULON_IO_XLSB_RECORD_WRITER_H_

#include <cstdint>
#include <optional>
#include <string_view>
#include <vector>

#include "io/zip_reader.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Appends a single 8-bit unsigned integer to `dst`.
void emit_u8(std::vector<std::uint8_t>& dst, std::uint8_t value);

/// Appends a little-endian 16-bit unsigned integer to `dst`.
void emit_u16(std::vector<std::uint8_t>& dst, std::uint16_t value);

/// Appends a little-endian 32-bit unsigned integer to `dst`.
void emit_u32(std::vector<std::uint8_t>& dst, std::uint32_t value);

/// Appends a little-endian IEEE 754 64-bit double to `dst`.
void emit_double(std::vector<std::uint8_t>& dst, double value);

/// Appends an `XLWideString` ([MS-XLSB] §2.5.166) to `dst`:
///   * `u32` count of UTF-16 code units (cch);
///   * `cch` UTF-16LE code units of the input UTF-8 `text`.
///
/// Surrogate pairs are emitted for code points outside the BMP. The
/// input must be valid UTF-8; ill-formed sequences are surfaced as
/// U+FFFD, matching what we do for malformed bytes elsewhere in the
/// engine. Empty input emits a 4-byte length prefix of zero followed
/// by no UTF-16 code units.
void emit_xlwidestring(std::vector<std::uint8_t>& dst, std::string_view text);

/// Appends an `XLNullableWideString` ([MS-XLSB] §2.5.167) to `dst`.
///   * `std::nullopt`         -> 4-byte sentinel `0xFFFFFFFF`.
///   * present empty string  -> length-zero XLWideString (4 bytes).
///   * present non-empty     -> regular XLWideString.
void emit_xlnullablewidestring(std::vector<std::uint8_t>& dst, std::optional<std::string_view> text);

/// Encodes `value` as an MS-XLSB `RkNumber` ([MS-XLSB] §2.5.121) and
/// appends the 4-byte little-endian encoding to `dst`.
///
/// Encoding choices:
///   * Integer form (`fInt=1`) when `value` is an exact integer in the
///     30-bit signed range `[-(1<<29), (1<<29) - 1]` and not negative
///     zero. The 30-bit two's-complement payload is shifted into the
///     upper 30 bits.
///   * Scaled-x100 form (`fX100=1`) when `value * 100` is an exact
///     integer in the 30-bit signed range AND has more precision than
///     the bare integer encoding could carry (used for currency-style
///     values like `123.45`). We only opt into x100 when `value` is
///     not already an exact integer — otherwise the bare integer form
///     is shorter and cleaner.
///   * IEEE 754 form otherwise: the upper 32 bits of the double's bit
///     pattern are masked with the low two bits cleared. This is
///     lossy unless the lower 34 bits of the double's bit pattern are
///     zero, so callers that need to preserve full precision (e.g.
///     `1.0/3.0`) must use `BrtCellReal` instead. The writer's cell
///     dispatcher (`cell_writer.{h,cpp}`) makes that choice.
///
/// The companion predicate `rk_round_trips_value` reports whether
/// `decode_rk_number(emit_rk_number(...))` will reproduce `value`
/// exactly. Callers that demand exact round-trip should consult it
/// before choosing between `BrtCellRk` and `BrtCellReal`.
void emit_rk_number(std::vector<std::uint8_t>& dst, double value);

/// Returns `true` when `value` can be encoded as an `RkNumber`
/// without precision loss (integer form or x100 form, or an IEEE 754
/// representation whose lower 34 bits are zero).
bool rk_round_trips_value(double value);

/// Appends one framed record to `dst`:
///   * variable-length record-type (1 or 2 MSB-continuation bytes);
///   * variable-length payload size (1..4 MSB-continuation bytes);
///   * `payload` bytes verbatim.
///
/// Round-tripping: bytes produced by this helper are accepted by
/// `read_record` and yield the same `(type, payload)` tuple. The
/// type and size encodings are minimal-length (we emit only as many
/// continuation bytes as required to represent the value).
void emit_record(std::vector<std::uint8_t>& dst, std::uint16_t type, ByteSpan payload);

/// Convenience overload: same as the `ByteSpan` form, but takes the
/// payload as a `std::vector` (which is what the cell/sheet writers
/// build their bodies into). Avoids forcing every caller to construct
/// a `ByteSpan` at the call site.
void emit_record(std::vector<std::uint8_t>& dst, std::uint16_t type, const std::vector<std::uint8_t>& payload);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_RECORD_WRITER_H_
