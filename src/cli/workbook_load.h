// Opening a workbook file for a CLI subcommand.

#ifndef FORMULON_CLI_WORKBOOK_LOAD_H_
#define FORMULON_CLI_WORKBOOK_LOAD_H_

#include <ostream>
#include <string>
#include <string_view>

#include "c_api/formulon_c.h"
#include "cli/c_api_raii.h"
#include "cli/cli.h"

namespace formulon::cli {

/// Reads and loads `path` into `wb`, names it after the file, applies the
/// shared options and reports load diagnostics. Every failure is written to
/// `err` under `subcommand` and its status returned.
fm_status_t load_workbook_file(const std::string& path, const SharedOptions& shared, WorkbookGuard& wb,
                               std::ostream& err, std::string_view subcommand);

}  // namespace formulon::cli

#endif  // FORMULON_CLI_WORKBOOK_LOAD_H_
