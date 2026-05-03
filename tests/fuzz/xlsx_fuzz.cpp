// Copyright 2026 libraz. Licensed under the MIT License.
//
// libFuzzer harness for the OOXML (.xlsx) reader.
//
// Fuzz goal: feed arbitrary bytes through the entire OOXML pipeline —
// zip extraction, XML parsing, cell decoding, shared-strings resolution,
// styles parsing — and detect crashes, leaks, or undefined behaviour.
// Excel files are zip archives, so the fuzzer effectively also fuzzes
// miniz and pugixml at their consumed surface.

#include <cstddef>
#include <cstdint>

#include "io/ooxml_reader.h"
#include "io/zip_reader.h"

extern "C" int LLVMFuzzerTestOneInput(const uint8_t* data, size_t size) {
  if (size > 4 * 1024 * 1024) {
    return 0;
  }
  formulon::io::ByteSpan bytes{data, size};
  (void)formulon::io::read_ooxml(bytes);
  return 0;
}
