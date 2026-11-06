#include "cli/workbook_load.h"

#include <cerrno>
#include <cstdint>
#include <cstring>
#include <vector>

#include "cli/diagnostics.h"
#include "cli/file_io.h"

namespace formulon::cli {

fm_status_t load_workbook_file(const std::string& path, const SharedOptions& shared, WorkbookGuard& wb,
                               std::ostream& err, std::string_view subcommand) {
  std::vector<std::uint8_t> bytes;
  const bool from_stdin = path == kStdioPath;
  if (auto rc = from_stdin ? read_stdin(bytes) : read_file(path, bytes); rc != 0) {
    if (rc == kExitUsage) {
      err << "formulon: " << subcommand << ": stdin is a terminal; pipe a workbook into '-' or name a file\n";
    } else if (from_stdin) {
      err << "formulon: " << subcommand << ": cannot read stdin: " << std::strerror(errno) << '\n';
    } else {
      err << "formulon: " << subcommand << ": cannot read '" << path << "': " << std::strerror(errno) << '\n';
    }
    return rc;
  }
  if (auto rc = fm_workbook_load(bytes.data(), bytes.size(), &wb.handle); rc != 0) {
    emit_last_error(err, subcommand);
    return rc;
  }
  if (!from_stdin) {
    set_file_name_from_path(wb.handle, path);
  }
  if (auto rc = apply_shared_options(wb.handle, shared); rc != 0) {
    emit_last_error(err, subcommand);
    return rc;
  }
  // Load losses are reported before any output is produced from the workbook.
  if (auto rc = emit_read_diagnostics(wb.handle, err, subcommand); rc != 0) {
    emit_last_error(err, subcommand);
    return rc;
  }
  return 0;
}

}  // namespace formulon::cli
