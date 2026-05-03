// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the MS-XLSB record-emission primitives. The byte
// layout matches `io/xlsb/record.cpp` exactly so every helper is the
// inverse of its `read_*` counterpart.

#include "io/xlsb/record_writer.h"

#include <cmath>
#include <cstdint>
#include <cstring>
#include <optional>
#include <string_view>
#include <vector>

#include "io/zip_reader.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Appends `value` as an MSB-continuation varint with at most
/// `max_bytes` 7-bit groups. Mirrors `record.cpp::ReadVarInt`. Each
/// emitted byte's MSB is set when more groups follow; the last group
/// is emitted with the MSB clear. Always emits at least one byte (the
/// minimum representation for `value == 0`).
void EmitVarInt(std::vector<std::uint8_t>& dst, std::uint32_t value, std::size_t max_bytes) {
  for (std::size_t i = 0; i < max_bytes; ++i) {
    const std::uint8_t low7 = static_cast<std::uint8_t>(value & 0x7FU);
    value >>= 7;
    if (value == 0U) {
      dst.push_back(low7);
      return;
    }
    dst.push_back(static_cast<std::uint8_t>(low7 | 0x80U));
  }
  // Caller bug: value didn't fit in `max_bytes` 7-bit groups. Defensive
  // fallback: drop the MSB on the last byte we wrote so the stream
  // doesn't claim a continuation that isn't there. The reader will
  // surface a `kIoXlsbRecordTruncated` if the value was actually used.
  if (!dst.empty()) {
    dst.back() &= 0x7FU;
  }
}

/// Appends one UTF-16LE code unit to `dst`.
inline void AppendCodeUnit(std::vector<std::uint8_t>& dst, std::uint16_t cu) {
  dst.push_back(static_cast<std::uint8_t>(cu & 0xFFU));
  dst.push_back(static_cast<std::uint8_t>((cu >> 8) & 0xFFU));
}

/// Decodes one UTF-8 codepoint from `text[i..]` and advances `i` past
/// the consumed bytes. Returns `0xFFFD` for ill-formed sequences,
/// consuming exactly one byte so we make forward progress.
std::uint32_t DecodeUtf8(std::string_view text, std::size_t& i) {
  const auto remaining = text.size() - i;
  const std::uint8_t b0 = static_cast<std::uint8_t>(text[i]);
  if (b0 < 0x80U) {
    ++i;
    return b0;
  }
  if ((b0 & 0xE0U) == 0xC0U && remaining >= 2U) {
    const std::uint8_t b1 = static_cast<std::uint8_t>(text[i + 1]);
    if ((b1 & 0xC0U) == 0x80U) {
      i += 2U;
      return (static_cast<std::uint32_t>(b0 & 0x1FU) << 6) | static_cast<std::uint32_t>(b1 & 0x3FU);
    }
  } else if ((b0 & 0xF0U) == 0xE0U && remaining >= 3U) {
    const std::uint8_t b1 = static_cast<std::uint8_t>(text[i + 1]);
    const std::uint8_t b2 = static_cast<std::uint8_t>(text[i + 2]);
    if ((b1 & 0xC0U) == 0x80U && (b2 & 0xC0U) == 0x80U) {
      i += 3U;
      return (static_cast<std::uint32_t>(b0 & 0x0FU) << 12) | (static_cast<std::uint32_t>(b1 & 0x3FU) << 6) |
             static_cast<std::uint32_t>(b2 & 0x3FU);
    }
  } else if ((b0 & 0xF8U) == 0xF0U && remaining >= 4U) {
    const std::uint8_t b1 = static_cast<std::uint8_t>(text[i + 1]);
    const std::uint8_t b2 = static_cast<std::uint8_t>(text[i + 2]);
    const std::uint8_t b3 = static_cast<std::uint8_t>(text[i + 3]);
    if ((b1 & 0xC0U) == 0x80U && (b2 & 0xC0U) == 0x80U && (b3 & 0xC0U) == 0x80U) {
      i += 4U;
      return (static_cast<std::uint32_t>(b0 & 0x07U) << 18) | (static_cast<std::uint32_t>(b1 & 0x3FU) << 12) |
             (static_cast<std::uint32_t>(b2 & 0x3FU) << 6) | static_cast<std::uint32_t>(b3 & 0x3FU);
    }
  }
  // Ill-formed: skip exactly one byte and surface U+FFFD so we always
  // make progress.
  ++i;
  return 0xFFFDU;
}

/// Returns the count of UTF-16 code units a UTF-8 string would expand
/// to. Used to size the cch length prefix without buffering the
/// encoded body in a separate vector first.
std::uint32_t Utf16Length(std::string_view text) {
  std::uint32_t cch = 0;
  std::size_t i = 0;
  while (i < text.size()) {
    const std::uint32_t cp = DecodeUtf8(text, i);
    cch += (cp >= 0x10000U) ? 2U : 1U;
  }
  return cch;
}

}  // namespace

void emit_u8(std::vector<std::uint8_t>& dst, std::uint8_t value) {
  dst.push_back(value);
}

void emit_u16(std::vector<std::uint8_t>& dst, std::uint16_t value) {
  dst.push_back(static_cast<std::uint8_t>(value & 0xFFU));
  dst.push_back(static_cast<std::uint8_t>((value >> 8) & 0xFFU));
}

void emit_u32(std::vector<std::uint8_t>& dst, std::uint32_t value) {
  dst.push_back(static_cast<std::uint8_t>(value & 0xFFU));
  dst.push_back(static_cast<std::uint8_t>((value >> 8) & 0xFFU));
  dst.push_back(static_cast<std::uint8_t>((value >> 16) & 0xFFU));
  dst.push_back(static_cast<std::uint8_t>((value >> 24) & 0xFFU));
}

void emit_double(std::vector<std::uint8_t>& dst, double value) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &value, sizeof(value));
  emit_u32(dst, static_cast<std::uint32_t>(bits & 0xFFFFFFFFU));
  emit_u32(dst, static_cast<std::uint32_t>((bits >> 32) & 0xFFFFFFFFU));
}

void emit_xlwidestring(std::vector<std::uint8_t>& dst, std::string_view text) {
  const std::uint32_t cch = Utf16Length(text);
  emit_u32(dst, cch);
  std::size_t i = 0;
  while (i < text.size()) {
    const std::uint32_t cp = DecodeUtf8(text, i);
    if (cp < 0x10000U) {
      AppendCodeUnit(dst, static_cast<std::uint16_t>(cp));
    } else {
      // Surrogate pair.
      const std::uint32_t adjusted = cp - 0x10000U;
      const std::uint16_t high = static_cast<std::uint16_t>(0xD800U + (adjusted >> 10));
      const std::uint16_t low = static_cast<std::uint16_t>(0xDC00U + (adjusted & 0x3FFU));
      AppendCodeUnit(dst, high);
      AppendCodeUnit(dst, low);
    }
  }
}

void emit_xlnullablewidestring(std::vector<std::uint8_t>& dst, std::optional<std::string_view> text) {
  if (!text.has_value()) {
    emit_u32(dst, 0xFFFFFFFFU);
    return;
  }
  emit_xlwidestring(dst, text.value());
}

bool rk_round_trips_value(double value) {
  if (std::isnan(value) || std::isinf(value)) {
    return false;
  }
  // Integer form covers any exact integer in the 30-bit signed range,
  // excluding negative zero (which loses its sign in the integer
  // representation).
  if (value == 0.0 && std::signbit(value)) {
    return false;
  }
  if (value == std::trunc(value)) {
    constexpr double kMin = -static_cast<double>(1 << 29);
    constexpr double kMax = static_cast<double>((1 << 29) - 1);
    if (value >= kMin && value <= kMax) {
      return true;
    }
  }
  // Scaled-x100 form: `value * 100` is an exact integer in the
  // 30-bit range. This catches typical currency-style values.
  const double scaled = value * 100.0;
  if (scaled == std::trunc(scaled)) {
    constexpr double kMin = -static_cast<double>(1 << 29);
    constexpr double kMax = static_cast<double>((1 << 29) - 1);
    if (scaled >= kMin && scaled <= kMax) {
      return true;
    }
  }
  // IEEE 754 form is exact only when the lower 34 bits of the bit
  // pattern are zero; otherwise we'd lose precision on the round-trip.
  std::uint64_t bits = 0;
  std::memcpy(&bits, &value, sizeof(value));
  return (bits & 0x3FFFFFFFFULL) == 0ULL;
}

void emit_rk_number(std::vector<std::uint8_t>& dst, double value) {
  // 1) Integer form: shortest representation when applicable.
  if (value != 0.0 || !std::signbit(value)) {
    if (value == std::trunc(value)) {
      constexpr double kMin = -static_cast<double>(1 << 29);
      constexpr double kMax = static_cast<double>((1 << 29) - 1);
      if (value >= kMin && value <= kMax) {
        const std::int32_t signed_payload = static_cast<std::int32_t>(value);
        // Pack the 30-bit signed integer into the upper 30 bits with
        // the low two bits as flags. The reader decodes by arithmetic-
        // shifting `static_cast<int32_t>(payload) >> 2`, which we mirror
        // here by shifting the signed value left by 2 (UB-free for
        // values in `[-(1<<29), (1<<29)-1]`, which is what we just
        // checked).
        const std::uint32_t shifted = static_cast<std::uint32_t>(signed_payload)
                                      << 2;                        // wraps for negatives, decoded via arithmetic shift
        const std::uint32_t rk = (shifted & 0xFFFFFFFCU) | 0x02U;  // fInt=1, fX100=0
        emit_u32(dst, rk);
        return;
      }
    }
  }

  // 2) x100 form: integer * 0.01 in the 30-bit range. Note we only
  // reach here when the integer form did not apply (which means
  // `value` is not already an exact integer in the 30-bit range);
  // this avoids preferring the longer `123 * 100` form over the
  // bare integer `123`. Negative zero is also rejected here because
  // the integer-form path already declined it (sign bit is lost).
  if (value != 0.0 || !std::signbit(value)) {
    const double scaled = value * 100.0;
    if (scaled == std::trunc(scaled)) {
      constexpr double kMin = -static_cast<double>(1 << 29);
      constexpr double kMax = static_cast<double>((1 << 29) - 1);
      if (scaled >= kMin && scaled <= kMax) {
        const std::int32_t signed_payload = static_cast<std::int32_t>(scaled);
        const std::uint32_t shifted = static_cast<std::uint32_t>(signed_payload) << 2;
        const std::uint32_t rk = (shifted & 0xFFFFFFFCU) | 0x03U;  // fInt=1, fX100=1
        emit_u32(dst, rk);
        return;
      }
    }
  }

  // 3) IEEE 754 form: store the upper 32 bits with the low two bits
  // cleared. Lossy unless the lower 34 bits of the bit pattern were
  // already zero — callers that need lossless output must use
  // BrtCellReal instead (see `rk_round_trips_value`).
  std::uint64_t bits = 0;
  std::memcpy(&bits, &value, sizeof(value));
  const std::uint32_t high = static_cast<std::uint32_t>(bits >> 32) & 0xFFFFFFFCU;
  emit_u32(dst, high);  // fInt=0, fX100=0
}

void emit_record(std::vector<std::uint8_t>& dst, std::uint16_t type, ByteSpan payload) {
  EmitVarInt(dst, static_cast<std::uint32_t>(type), /*max_bytes=*/2);
  EmitVarInt(dst, static_cast<std::uint32_t>(payload.size), /*max_bytes=*/4);
  if (payload.size > 0 && payload.data != nullptr) {
    dst.insert(dst.end(), payload.data, payload.data + payload.size);
  }
}

void emit_record(std::vector<std::uint8_t>& dst, std::uint16_t type, const std::vector<std::uint8_t>& payload) {
  emit_record(dst, type, ByteSpan{payload.data(), payload.size()});
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
