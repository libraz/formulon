//
// Implementation of the MS-XLSB record framing primitives. See
// `io/xlsb/record.h` for the wire-format reference.

#include "io/xlsb/record.h"

#include <cstdint>
#include <cstring>
#include <string>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Reads up to `max_bytes` MSB-continuation bytes into `out`, where each
/// byte contributes 7 bits in little-endian order. Returns the actual
/// number of bytes consumed in `bytes_consumed`. Used for both the
/// record-type and payload-size variable-length integers.
Expected<std::uint32_t, Error> ReadVarInt(ByteSpan& cursor, std::size_t max_bytes, const char* what) {
  std::uint32_t value = 0;
  for (std::size_t i = 0; i < max_bytes; ++i) {
    if (cursor.size == 0) {
      std::string ctx("context=xlsb.record what=");
      ctx.append(what);
      ctx.append(" byte_index=");
      ctx.append(std::to_string(i));
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb varint truncated", std::move(ctx));
    }
    const std::uint8_t b = cursor.data[0];
    cursor.data += 1;
    cursor.size -= 1;
    value |= static_cast<std::uint32_t>(b & 0x7F) << (7 * i);
    if ((b & 0x80) == 0) {
      return value;
    }
  }
  // All `max_bytes` had the continuation bit set, which is malformed.
  std::string ctx("context=xlsb.record what=");
  ctx.append(what);
  return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb varint exceeded max length", std::move(ctx));
}

}  // namespace

Expected<std::uint8_t, Error> read_u8(ByteSpan& cursor) {
  if (cursor.size < 1) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb u8 read out of range", "context=xlsb.record");
  }
  const std::uint8_t v = cursor.data[0];
  cursor.data += 1;
  cursor.size -= 1;
  return v;
}

Expected<std::uint16_t, Error> read_u16(ByteSpan& cursor) {
  if (cursor.size < 2) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb u16 read out of range", "context=xlsb.record");
  }
  const std::uint32_t lo = cursor.data[0];
  const std::uint32_t hi = cursor.data[1];
  const std::uint16_t v = static_cast<std::uint16_t>(lo | (hi << 8));
  cursor.data += 2;
  cursor.size -= 2;
  return v;
}

Expected<std::uint32_t, Error> read_u32(ByteSpan& cursor) {
  if (cursor.size < 4) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb u32 read out of range", "context=xlsb.record");
  }
  const std::uint32_t v =
      static_cast<std::uint32_t>(cursor.data[0]) | (static_cast<std::uint32_t>(cursor.data[1]) << 8) |
      (static_cast<std::uint32_t>(cursor.data[2]) << 16) | (static_cast<std::uint32_t>(cursor.data[3]) << 24);
  cursor.data += 4;
  cursor.size -= 4;
  return v;
}

Expected<std::string, Error> read_xlwidestring(ByteSpan& cursor) {
  auto len_or = read_u32(cursor);
  if (!len_or) {
    return len_or.error();
  }
  const std::uint32_t cch = len_or.value();
  // 16-bit code units per char; cap at the buffer to avoid large
  // synthetic allocations on hostile input. The bound on `cch` is also
  // what makes the `cch * 3` reservation below safe to compute as
  // `std::size_t`: at this point we know `cch <= cursor.size / 2`, so
  // `cch * 3 <= cursor.size * 1.5`, which cannot overflow even on the
  // 32-bit `size_t` of a WASM build (`cursor.size` is always at most the
  // enclosing record's payload size).
  if (cch > cursor.size / 2) {
    std::string ctx("context=xlsb.record what=XLWideString cch=");
    ctx.append(std::to_string(cch));
    ctx.append(" cursor_size=");
    ctx.append(std::to_string(cursor.size));
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb wide-string body truncated", std::move(ctx));
  }
  const std::size_t needed = static_cast<std::size_t>(cch) * 2;
  std::string out;
  // Reserve a generous upper bound: every code unit takes at most 3
  // UTF-8 bytes (BMP) and surrogate pairs take 4 over two units. Safe
  // from overflow per the bound established above.
  out.reserve(static_cast<std::size_t>(cch) * 3);
  for (std::uint32_t i = 0; i < cch; ++i) {
    const std::uint32_t cu_lo = cursor.data[2 * i];
    const std::uint32_t cu_hi = cursor.data[2 * i + 1];
    const std::uint16_t cu = static_cast<std::uint16_t>(cu_lo | (cu_hi << 8));
    std::uint32_t cp = cu;
    // Surrogate pair: high surrogate followed by low surrogate.
    if (cu >= 0xD800 && cu <= 0xDBFF && i + 1 < cch) {
      const std::uint32_t low_lo = cursor.data[2 * (i + 1)];
      const std::uint32_t low_hi = cursor.data[2 * (i + 1) + 1];
      const std::uint16_t low = static_cast<std::uint16_t>(low_lo | (low_hi << 8));
      if (low >= 0xDC00 && low <= 0xDFFF) {
        cp =
            0x10000U + ((static_cast<std::uint32_t>(cu) - 0xD800U) << 10) + (static_cast<std::uint32_t>(low) - 0xDC00U);
        ++i;  // consumed the low surrogate as well
      }
    }
    if (cp < 0x80) {
      out.push_back(static_cast<char>(cp));
    } else if (cp < 0x800) {
      out.push_back(static_cast<char>(0xC0 | (cp >> 6)));
      out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else if (cp < 0x10000) {
      out.push_back(static_cast<char>(0xE0 | (cp >> 12)));
      out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
      out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    } else {
      out.push_back(static_cast<char>(0xF0 | (cp >> 18)));
      out.push_back(static_cast<char>(0x80 | ((cp >> 12) & 0x3F)));
      out.push_back(static_cast<char>(0x80 | ((cp >> 6) & 0x3F)));
      out.push_back(static_cast<char>(0x80 | (cp & 0x3F)));
    }
  }
  cursor.data += needed;
  cursor.size -= needed;
  return out;
}

Expected<std::string, Error> read_xlnullablewidestring(ByteSpan& cursor) {
  if (cursor.size < 4) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb nullable-wide-string len truncated",
                      "context=xlsb.record");
  }
  // Peek at the length prefix without consuming, so the null sentinel
  // (`0xFFFFFFFF`) advances exactly four bytes and produces an empty
  // string without invoking the wide-string body decoder.
  const std::uint32_t len =
      static_cast<std::uint32_t>(cursor.data[0]) | (static_cast<std::uint32_t>(cursor.data[1]) << 8) |
      (static_cast<std::uint32_t>(cursor.data[2]) << 16) | (static_cast<std::uint32_t>(cursor.data[3]) << 24);
  if (len == 0xFFFFFFFFU) {
    cursor.data += 4;
    cursor.size -= 4;
    return std::string{};
  }
  return read_xlwidestring(cursor);
}

double decode_rk_number(std::uint32_t rk) {
  // Bit 0: fX100 — divide result by 100.
  // Bit 1: fInt  — upper 30 bits are a signed integer (vs. truncated
  //                IEEE-754 double).
  const bool f_x100 = (rk & 0x1U) != 0;
  const bool f_int = (rk & 0x2U) != 0;
  const std::uint32_t payload = rk & 0xFFFFFFFCU;  // upper 30 bits in place
  double out = 0.0;
  if (f_int) {
    // Sign-extend the 30-bit signed integer payload.
    const std::int32_t signed_payload = static_cast<std::int32_t>(payload) >> 2;  // arithmetic shift
    out = static_cast<double>(signed_payload);
  } else {
    // Reconstruct the IEEE 754 double: the 30 bits go into the high 30
    // of the 64-bit double; the low 34 bits are zero.
    const std::uint64_t bits = static_cast<std::uint64_t>(payload) << 32;
    double d;
    std::memcpy(&d, &bits, sizeof(d));
    out = d;
  }
  if (f_x100) {
    out /= 100.0;
  }
  return out;
}

Expected<XlsbRecord, Error> read_record(ByteSpan& cursor) {
  // Record-type: up to 2 MSB-continuation bytes encoding 14 bits.
  auto type_or = ReadVarInt(cursor, /*max_bytes=*/2, "record_type");
  if (!type_or) {
    return type_or.error();
  }
  // Payload-size: up to 4 MSB-continuation bytes encoding 28 bits.
  auto size_or = ReadVarInt(cursor, /*max_bytes=*/4, "record_size");
  if (!size_or) {
    return size_or.error();
  }
  const std::uint32_t payload_size = size_or.value();
  if (cursor.size < payload_size) {
    std::string ctx("context=xlsb.record need=");
    ctx.append(std::to_string(payload_size));
    ctx.append(" have=");
    ctx.append(std::to_string(cursor.size));
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb record payload truncated", std::move(ctx));
  }
  XlsbRecord rec;
  rec.type = static_cast<std::uint16_t>(type_or.value());
  rec.payload = ByteSpan{cursor.data, payload_size};
  cursor.data += payload_size;
  cursor.size -= payload_size;
  return rec;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
