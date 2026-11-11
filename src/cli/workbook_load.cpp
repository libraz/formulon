#include "cli/workbook_load.h"

#include <cerrno>
#include <cstdint>
#include <cstring>
#include <vector>

#include "cli/diagnostics.h"
#include "cli/file_io.h"

namespace formulon::cli {

fm_status_t load_workbook_file(const std::string& path, const SharedOptions& shared, WorkbookGuard& wb,
                               std::ostream& err, std::string_view subcommand, bool report_loss) {
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
    return fail_with_last_error(err, subcommand, rc);
  }
  if (!from_stdin) {
    set_location_from_path(wb.handle, path);
  }
  if (auto rc = apply_shared_options(wb.handle, shared); rc != 0) {
    return fail_with_last_error(err, subcommand, rc);
  }
  // Load losses are reported before any output is produced from the workbook.
  if (!report_loss) {
    return 0;
  }
  if (auto rc = emit_read_diagnostics(wb.handle, err, subcommand); rc != 0) {
    return fail_with_last_error(err, subcommand, rc);
  }
  return 0;
}

fm_status_t use_manual_calculation(fm_workbook_t* wb, std::ostream& err, std::string_view subcommand) {
  if (auto rc = fm_workbook_set_calc_mode(wb, FM_CALC_MODE_MANUAL); rc != 0) {
    return fail_with_last_error(err, subcommand, rc);
  }
  return 0;
}

fm_status_t resolve_sheet(const fm_workbook_t* wb, std::string_view spec, std::size_t* out, std::ostream& err,
                          std::string_view subcommand) {
  const std::string name(spec);
  std::size_t index = 0;
  std::int32_t present = 0;
  if (auto rc = fm_workbook_get_sheet_index(wb, name.c_str(), &present, &index); rc != 0) {
    return fail_with_last_error(err, subcommand, rc);
  }
  if (present == 0) {
    std::size_t count = 0;
    if (auto rc = fm_workbook_get_sheet_count(wb, &count); rc != 0) {
      return fail_with_last_error(err, subcommand, rc);
    }
    if (!parse_unsigned(spec, &index) || index >= count) {
      err << "formulon: " << subcommand << ": no sheet named or numbered '" << spec << "'\n";
      return kExitUsage;
    }
  }
  *out = index;
  return 0;
}

}  // namespace formulon::cli
