#pragma once

#include <cstdint>
#include <vector>

#include "io/xlsb/record.h"

namespace formulon::io::xlsb {

inline ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return ByteSpan{v.data(), v.size()};
}

}  // namespace formulon::io::xlsb
