//
// libFuzzer harness for the MS-XLSB (.xlsb) reader.
//
// XLSB shares OOXML's ZIP envelope but feeds several length- and offset-driven
// binary record readers after package extraction. Keep this harness separate
// from xlsx_fuzz so mutations always reach `read_xlsb` and its Ptg decoder.

#include <cstddef>
#include <cstdint>

#include "io/xlsb/reader.h"

extern "C" int LLVMFuzzerTestOneInput(const uint8_t* data, size_t size) {
  if (size > 4 * 1024 * 1024) {
    return 0;
  }
  formulon::io::ByteSpan bytes{data, size};
  (void)formulon::io::xlsb::read_xlsb(bytes);
  return 0;
}
