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

struct SaveBuffer {
  std::uint8_t* data = nullptr;
  std::size_t len = 0;
  SaveBuffer() = default;
  SaveBuffer(const SaveBuffer&) = delete;
  SaveBuffer& operator=(const SaveBuffer&) = delete;
  ~SaveBuffer() { fm_buffer_free(data); }
};

}  // namespace cli
}  // namespace formulon

#endif  // FORMULON_CLI_C_API_RAII_H_
