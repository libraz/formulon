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

/// Reads and loads `path` into `wb` (`-` reads stdin to the end and leaves the
/// workbook unnamed), names it after the file, applies the
/// shared options and reports load diagnostics. Every failure is written to
/// `err` under `subcommand` and its status returned.
fm_status_t load_workbook_file(const std::string& path, const SharedOptions& shared, WorkbookGuard& wb,
                               std::ostream& err, std::string_view subcommand);

/// Switches a loaded workbook to manual calculation so that reading it shows
/// the file's stored values. Failures are written to `err` under `subcommand`.
fm_status_t use_manual_calculation(fm_workbook_t* wb, std::ostream& err, std::string_view subcommand);

/// Resolves `spec`, a sheet name or a 0-based sheet index (a name wins over an
/// index), to `*out`. An unknown sheet is written to `err` and returns
/// `kExitUsage`; an engine failure returns its status after writing it.
fm_status_t resolve_sheet(const fm_workbook_t* wb, std::string_view spec, std::size_t* out, std::ostream& err,
                          std::string_view subcommand);

}  // namespace formulon::cli

#endif  // FORMULON_CLI_WORKBOOK_LOAD_H_
