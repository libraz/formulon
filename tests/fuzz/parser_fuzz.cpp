// Copyright 2026 libraz. Licensed under the MIT License.
//
// libFuzzer harness for the Formulon Pratt parser.
//
// Fuzz goal: feed arbitrary bytes as Excel formula text into the parser
// and assert no crash, ASan violation, UBSan violation, or infinite loop.
// The parser must reject malformed input via the diagnostics list rather
// than aborting.
//
// Build: clang only, with `-DFM_BUILD_FUZZ=ON`.
// Run: `./build/bin/parser_fuzz -runs=1000` (smoke) or `-max_total_time=21600`
// (6h nightly).

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "parser/parser.h"
#include "utils/arena.h"

extern "C" int LLVMFuzzerTestOneInput(const uint8_t* data, size_t size) {
  if (size > 65536) {
    return 0;
  }
  formulon::Arena arena;
  std::string_view src(reinterpret_cast<const char*>(data), size);
  formulon::parser::Parser parser(src, arena);
  (void)parser.parse();
  (void)parser.errors();
  return 0;
}
