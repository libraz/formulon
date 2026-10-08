// Double-byte code page lookups and the decoder for the generated tables.

#include "eval/dbcs_table.h"

namespace formulon {
namespace eval {
namespace dbcs_detail {
namespace {

// Must match ACC_SHIFT / K_BIAS / ACC_INIT in tools/dbcs/generate_tables.py.
constexpr unsigned kAccShift = 1;
constexpr unsigned kKBias = 2;
constexpr std::uint32_t kAccInit = 32;

class BitReader {
 public:
  BitReader(const std::uint8_t* data, std::size_t size) noexcept : data_(data), bits_(size * 8) {}

  bool exhausted() const noexcept { return pos_ >= bits_; }

  std::uint32_t bit() noexcept {
    if (pos_ >= bits_) {
      return 0;
    }
    const std::uint32_t b = (data_[pos_ >> 3] >> (7 - (pos_ & 7))) & 1u;
    ++pos_;
    return b;
  }

  std::uint32_t bits(unsigned width) noexcept {
    std::uint32_t v = 0;
    for (unsigned i = 0; i < width; ++i) {
      v = (v << 1) | bit();
    }
    return v;
  }

  std::uint32_t exp_golomb(unsigned k) noexcept {
    unsigned length = 0;
    while (bit() == 0) {
      if (exhausted() || ++length > 24) {
        return 0;
      }
    }
    const std::uint32_t w = (1u << length) | bits(length);
    return ((w - 1u) << k) | bits(k);
  }

 private:
  const std::uint8_t* data_;
  std::size_t bits_;
  std::size_t pos_ = 0;
};

unsigned bit_length(std::uint32_t v) noexcept {
  unsigned n = 0;
  while (v != 0) {
    v >>= 1;
    ++n;
  }
  return n;
}

}  // namespace

DbcsCells::DbcsCells(const std::uint8_t* encoded, std::size_t size) noexcept : unicode{} {
  constexpr std::size_t kSlots = kDbcsGridSize * kDbcsGridSize;
  BitReader reader(encoded, size);
  std::size_t slot = 0;
  std::uint32_t prev = 0;
  std::uint32_t acc = kAccInit;
  while (slot < kSlots && !reader.exhausted()) {
    const unsigned len = bit_length(acc >> kAccShift);
    const std::uint32_t v = reader.exp_golomb(len > kKBias ? len - kKBias : 0);
    if (v == 0) {
      // A run of unassigned slots; they are already zero.
      slot += static_cast<std::size_t>(reader.exp_golomb(0)) + 1;
      continue;
    }
    // Zigzag: even values are positive deltas, odd ones negative.
    prev = (v & 1u) != 0 ? prev - ((v + 1u) >> 1) : prev + (v >> 1);
    unicode[slot++] = static_cast<std::uint16_t>(prev);
    acc += v - (acc >> kAccShift);
  }
}

}  // namespace dbcs_detail

namespace {

const std::uint16_t* cells_for(DbcsCodepage codepage) noexcept {
  switch (codepage) {
    case DbcsCodepage::kJis0208:
      return dbcs_detail::jis0208_cells();
    case DbcsCodepage::kGb2312:
      return dbcs_detail::gb2312_cells();
    case DbcsCodepage::kKsX1001:
      return dbcs_detail::ksx1001_cells();
    case DbcsCodepage::kNone:
      break;
  }
  return nullptr;
}

}  // namespace

std::uint16_t lookup_unicode_to_dbcs(DbcsCodepage codepage, std::uint32_t codepoint) noexcept {
  // Slot value 0 marks an unassigned cell, so U+0000 never matches.
  if (codepoint == 0u || codepoint > 0xFFFFu) {
    return 0u;
  }
  const std::uint16_t* cells = cells_for(codepage);
  if (cells == nullptr) {
    return 0u;
  }
  constexpr std::size_t kSlots = kDbcsGridSize * kDbcsGridSize;
  for (std::size_t i = 0; i < kSlots; ++i) {
    if (cells[i] == codepoint) {
      return static_cast<std::uint16_t>(((i / kDbcsGridSize + 1) << 8) | (i % kDbcsGridSize + 1));
    }
  }
  return 0u;
}

std::uint16_t lookup_dbcs_to_unicode(DbcsCodepage codepage, std::uint8_t row, std::uint8_t cell) noexcept {
  if (row < 1u || row > kDbcsGridSize || cell < 1u || cell > kDbcsGridSize) {
    return 0u;
  }
  const std::uint16_t* cells = cells_for(codepage);
  if (cells == nullptr) {
    return 0u;
  }
  return cells[(row - 1u) * kDbcsGridSize + (cell - 1u)];
}

std::uint16_t dbcs_code_bias(DbcsCodepage codepage) noexcept {
  switch (codepage) {
    case DbcsCodepage::kJis0208:
      return 0x2020u;
    case DbcsCodepage::kGb2312:
    case DbcsCodepage::kKsX1001:
      return 0xA0A0u;
    case DbcsCodepage::kNone:
      break;
  }
  return 0u;
}

}  // namespace eval
}  // namespace formulon
