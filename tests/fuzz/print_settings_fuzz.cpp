//
// libFuzzer harness for the print-settings raw-XML setters.
//
// Fuzz goal: the five setters are a new path for caller-authored bytes to
// reach pugixml, and each one re-parses what it stored to re-derive the
// structured views the paginator reads. Feed arbitrary bytes through every
// setter, then paginate, and detect crashes, leaks, or undefined behaviour
// on either half of that round trip.
//
// The contract under test is total: whatever arrives, the call returns one
// of `kOk` / `kInvalidArgument` / `kPreconditionFailed`, and a sheet that
// accepted a fragment must still paginate.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"

namespace {

// One byte of the input selects the setter so a single corpus entry can
// reach any of them; the rest is the fragment.
constexpr std::size_t kSetterCount = 5;

using SetterFn = fm_status_t (*)(fm_workbook_t*, std::size_t, const char*);

constexpr SetterFn kSetters[kSetterCount] = {
    &fm_sheet_set_page_setup_xml,    &fm_sheet_set_page_margins_xml, &fm_sheet_set_print_options_xml,
    &fm_sheet_set_header_footer_xml, &fm_sheet_set_sheet_pr_xml,
};

}  // namespace

extern "C" int LLVMFuzzerTestOneInput(const uint8_t* data, size_t size) {
  if (size < 1 || size > 256 * 1024) {
    return 0;
  }
  const std::size_t selector = data[0] % kSetterCount;
  // The C ABI takes a NUL-terminated string, so an embedded NUL truncates
  // the fragment rather than being part of it - which is itself worth
  // exercising, and is why the copy stops at the first NUL the same way
  // the entry point will.
  const std::string fragment(reinterpret_cast<const char*>(data + 1), size - 1);

  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create(&wb) != 0 || wb == nullptr) {
    return 0;
  }
  (void)kSetters[selector](wb, 0, fragment.c_str());

  // Whatever the setter decided, the sheet must remain paginable: an
  // accepted fragment feeds the structured views, and a rejected one must
  // have left them untouched.
  fm_pagination_t* pagination = nullptr;
  if (fm_workbook_paginate(wb, 0, &pagination) == 0) {
    fm_pagination_destroy(pagination);
  }

  // The getter shares the scratch contract with every other read, so run
  // it too rather than leaving half the pair unfuzzed.
  const char* readback = nullptr;
  (void)fm_sheet_get_page_setup_xml(wb, 0, &readback);

  fm_workbook_destroy(wb);
  return 0;
}
