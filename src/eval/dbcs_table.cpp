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

void decode_dbcs_cells(const std::uint8_t* encoded, std::size_t size, std::uint16_t* out, std::size_t slots) noexcept {
  BitReader reader(encoded, size);
  std::size_t slot = 0;
  std::uint32_t prev = 0;
  std::uint32_t acc = kAccInit;
  while (slot < slots && !reader.exhausted()) {
    const unsigned len = bit_length(acc >> kAccShift);
    const std::uint32_t v = reader.exp_golomb(len > kKBias ? len - kKBias : 0);
    if (v == 0) {
      // A run of unassigned slots; they are already zero.
      slot += static_cast<std::size_t>(reader.exp_golomb(0)) + 1;
      continue;
    }
    // Zigzag: even values are positive deltas, odd ones negative.
    prev = (v & 1u) != 0 ? prev - ((v + 1u) >> 1) : prev + (v >> 1);
    out[slot++] = static_cast<std::uint16_t>(prev);
    acc += v - (acc >> kAccShift);
  }
}

}  // namespace dbcs_detail

namespace {

const dbcs_detail::DbcsGrid* grid_for(DbcsCodepage codepage) noexcept {
  switch (codepage) {
    case DbcsCodepage::kJis0208:
      return &dbcs_detail::jis0208_grid();
    case DbcsCodepage::kGb2312:
      return &dbcs_detail::gb2312_grid();
    case DbcsCodepage::kKsX1001:
      return &dbcs_detail::ksx1001_grid();
    case DbcsCodepage::kBig5:
      return &dbcs_detail::big5_grid();
    case DbcsCodepage::kNone:
      break;
  }
  return nullptr;
}

std::size_t trail_total(const dbcs_detail::DbcsShape& shape) noexcept {
  return static_cast<std::size_t>(shape.trails[0].count) + shape.trails[1].count;
}

// Offset of `trail` within the shape's trail ranges, or trail_total when outside.
std::size_t trail_index(const dbcs_detail::DbcsShape& shape, std::uint32_t trail) noexcept {
  std::size_t before = 0;
  for (const auto& range : shape.trails) {
    if (trail >= range.first && trail < static_cast<std::uint32_t>(range.first) + range.count) {
      return before + (trail - range.first);
    }
    before += range.count;
  }
  return before;
}

}  // namespace

std::uint16_t lookup_unicode_to_dbcs(DbcsCodepage codepage, std::uint32_t codepoint) noexcept {
  // Slot value 0 marks an unassigned cell, so U+0000 never matches.
  if (codepoint == 0u || codepoint > 0xFFFFu) {
    return 0u;
  }
  const dbcs_detail::DbcsGrid* grid = grid_for(codepage);
  if (grid == nullptr) {
    return 0u;
  }
  const dbcs_detail::DbcsShape& shape = grid->shape;
  const std::size_t trails = trail_total(shape);
  const std::size_t slots = static_cast<std::size_t>(shape.lead_count) * trails;
  for (std::size_t i = 0; i < slots; ++i) {
    if (grid->unicode[i] != codepoint) {
      continue;
    }
    std::size_t offset = i % trails;
    std::uint32_t trail = 0;
    for (const auto& range : shape.trails) {
      if (offset < range.count) {
        trail = range.first + static_cast<std::uint32_t>(offset);
        break;
      }
      offset -= range.count;
    }
    const std::uint32_t lead = shape.lead_first + static_cast<std::uint32_t>(i / trails);
    return static_cast<std::uint16_t>((lead << 8) | trail);
  }
  return 0u;
}

std::uint16_t lookup_dbcs_to_unicode(DbcsCodepage codepage, std::uint8_t lead, std::uint8_t trail) noexcept {
  const dbcs_detail::DbcsGrid* grid = grid_for(codepage);
  if (grid == nullptr) {
    return 0u;
  }
  const dbcs_detail::DbcsShape& shape = grid->shape;
  const std::size_t trails = trail_total(shape);
  const std::size_t t = trail_index(shape, trail);
  if (lead < shape.lead_first || lead >= static_cast<unsigned>(shape.lead_first) + shape.lead_count || t >= trails) {
    return 0u;
  }
  return grid->unicode[static_cast<std::size_t>(lead - shape.lead_first) * trails + t];
}

std::uint16_t dbcs_code_bias(DbcsCodepage codepage) noexcept {
  switch (codepage) {
    case DbcsCodepage::kJis0208:
      return 0x2020u;
    case DbcsCodepage::kGb2312:
    case DbcsCodepage::kKsX1001:
      return 0xA0A0u;
    case DbcsCodepage::kBig5:
    case DbcsCodepage::kNone:
      break;
  }
  return 0u;
}

}  // namespace eval
}  // namespace formulon
