// Small RAII wrappers for C ABI resources used by CLI commands.

#ifndef FORMULON_CLI_C_API_RAII_H_
#define FORMULON_CLI_C_API_RAII_H_

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"

namespace formulon {
namespace cli {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
};

struct ValueArrayGuard {
  fm_value_array_t* handle = nullptr;
  ValueArrayGuard() = default;
  ValueArrayGuard(const ValueArrayGuard&) = delete;
  ValueArrayGuard& operator=(const ValueArrayGuard&) = delete;
  ~ValueArrayGuard() { fm_value_array_destroy(handle); }
};

struct PaginationGuard {
  fm_pagination_t* handle = nullptr;
  PaginationGuard() = default;
  PaginationGuard(const PaginationGuard&) = delete;
  PaginationGuard& operator=(const PaginationGuard&) = delete;
  ~PaginationGuard() { fm_pagination_destroy(handle); }
};

struct SaveBuffer {
  std::uint8_t* data = nullptr;
  std::size_t len = 0;
  SaveBuffer() = default;
  SaveBuffer(const SaveBuffer&) = delete;
  SaveBuffer& operator=(const SaveBuffer&) = delete;
  ~SaveBuffer() { fm_buffer_destroy(data); }
};

}  // namespace cli
}  // namespace formulon

#endif  // FORMULON_CLI_C_API_RAII_H_
