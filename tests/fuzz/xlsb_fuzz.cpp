//
// libFuzzer harness for the MS-XLSB (.xlsb) reader.
//
// XLSB shares OOXML's ZIP envelope but feeds several length- and offset-driven
// binary record readers after package extraction. Keep this harness separate
// from xlsx_fuzz so mutations always reach `read_xlsb` and its Ptg decoder.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <vector>

#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/reader.h"
#include "utils/arena.h"

namespace {

constexpr char kPtgHexPrefix[] = "FORMULON_XLSB_PTG_HEX:";

int hex_digit(std::uint8_t c) {
  if (c >= '0' && c <= '9') {
    return c - '0';
  }
  if (c >= 'a' && c <= 'f') {
    return c - 'a' + 10;
  }
  if (c >= 'A' && c <= 'F') {
    return c - 'A' + 10;
  }
  return -1;
}

bool decode_ptg_hex_corpus(const std::uint8_t* data, std::size_t size, std::vector<std::uint8_t>* ptgs) {
  constexpr std::size_t prefix_size = sizeof(kPtgHexPrefix) - 1;
  if (size < prefix_size || std::memcmp(data, kPtgHexPrefix, prefix_size) != 0) {
    return false;
  }

  int high_nibble = -1;
  for (std::size_t i = prefix_size; i < size; ++i) {
    const int digit = hex_digit(data[i]);
    if (digit < 0) {
      // Whitespace keeps corpus seeds reviewable; every other character
      // rejects the special format and falls through to the XLSB reader.
      if (data[i] == ' ' || data[i] == '\n' || data[i] == '\r' || data[i] == '\t') {
        continue;
      }
      return false;
    }
    if (high_nibble < 0) {
      high_nibble = digit;
    } else {
      ptgs->push_back(static_cast<std::uint8_t>((high_nibble << 4) | digit));
      high_nibble = -1;
    }
  }
  return high_nibble < 0;
}

}  // namespace

extern "C" int LLVMFuzzerTestOneInput(const uint8_t* data, size_t size) {
  if (size > 4 * 1024 * 1024) {
    return 0;
  }

  // A portable, textual Ptg seed exercises decoder depths that are not
  // practical to embed in a complete XLSB package. Normal inputs still take
  // the whole-package reader path below.
  std::vector<std::uint8_t> ptgs;
  if (decode_ptg_hex_corpus(data, size, &ptgs)) {
    formulon::Arena arena;
    formulon::io::ByteSpan bytes{ptgs.data(), ptgs.size()};
    (void)formulon::io::xlsb::decode_ptgs(bytes, {}, arena, {}, {}, {});
    return 0;
  }

  formulon::io::ByteSpan bytes{data, size};
  (void)formulon::io::xlsb::read_xlsb(bytes);
  return 0;
}
