#pragma once

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"

// Small ownership guards shared by the C ABI topic tests.  Keeping the
// lifetime rules in one header makes each topic file focus on its contract.
struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct BufferGuard {
  std::uint8_t* data = nullptr;
  std::size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};
