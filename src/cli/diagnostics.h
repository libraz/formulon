//
// Shared load / save diagnostic reporting for the CLI subcommands.
//
// The engine reports data it could not represent through counter structs
// rather than a status code, so a lossy load or save still succeeds. Any
// subcommand whose output a user reads as a faithful representation of the
// input therefore has to surface those counters itself, and has to phrase
// them the same way every other subcommand does. The emitters live here so
// there is one place that decides what a lossy round-trip looks like on
// stderr, instead of one copy per command TU.

#ifndef FORMULON_CLI_DIAGNOSTICS_H_
#define FORMULON_CLI_DIAGNOSTICS_H_

#include <ostream>
#include <string_view>

#include "c_api/formulon_c.h"

namespace formulon {
namespace cli {

/// Queries `wb`'s load counters and writes one warning line per non-empty
/// counter group to `err`, prefixed with `subcommand`.
///
/// Call this immediately after `fm_workbook_load` and before any primary
/// output: exit status 0 must never imply a lossless load, and the warning
/// has to reach the user whether or not the command goes on to succeed.
/// Verbosity flags do not gate it — a dropped formula is not chatter.
///
/// Returns the status of the underlying `fm_workbook_read_diagnostics`
/// query so a handle failure propagates to the caller unchanged.
fm_status_t emit_read_diagnostics(const fm_workbook_t* wb, std::ostream& err, std::string_view subcommand);

/// Writes one warning line to `err` for the non-zero counters in `d`,
/// labelled with the container `format` actually written and prefixed with
/// `subcommand`. Emits nothing when the save lost nothing.
void emit_write_diagnostics(std::ostream& err, std::string_view subcommand, fm_workbook_format_t format,
                            const fm_save_diagnostics_t& d);

}  // namespace cli
}  // namespace formulon

#endif  // FORMULON_CLI_DIAGNOSTICS_H_
