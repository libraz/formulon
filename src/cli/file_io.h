//
// Shared file I/O primitives for native CLI commands. The C ABI remains
// bytes-in / bytes-out; only the CLI maps paths to byte buffers.

#ifndef FORMULON_CLI_FILE_IO_H_
#define FORMULON_CLI_FILE_IO_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "cli/exit_codes.h"

namespace formulon::cli {

/// The path operand that stands for stdin (input) or stdout (`-o`).
inline constexpr std::string_view kStdioPath = "-";

/// Reads an entire binary file into `out`. Returns 0, `kExitNoInput` when the
/// file cannot be opened or is a directory, or `kExitIoError` when the read
/// fails; `errno` describes the failure.
fm_status_t read_file(const std::string& path, std::vector<std::uint8_t>& out);

/// Reads stdin to the end into `out`. Returns 0, `kExitUsage` when stdin is a
/// terminal, or `kExitIoError` when the read fails; `errno` describes it.
fm_status_t read_stdin(std::vector<std::uint8_t>& out);

/// True when stdout is a terminal.
bool stdout_is_terminal();

/// Writes `len` bytes to stdout (descriptor 1), bypassing the C++ streams so
/// binary data is not altered. Returns 0 or `kExitIoError`; `errno` describes it.
fm_status_t write_stdout(const std::uint8_t* bytes, std::size_t len);

/// Atomically replaces `path` with `bytes`. The temporary file is created
/// exclusively with a random suffix in the target directory, preserving the
/// atomic-rename guarantee and preventing predictable-name collisions.
///
/// A `path` that is a symlink is resolved first, so the replace lands on
/// the file the link names and the link itself survives. A dangling link
/// is replaced as-is: there is no target to keep.
///
/// Returns 0, `kExitCantCreate` when the temporary cannot be created or the
/// final rename fails, or `kExitIoError` when writing it fails; `errno`
/// describes the failure.
///
/// No exit path leaves an open descriptor or a leftover temporary behind:
/// on failure the temporary is unlinked and `path` keeps whatever content
/// it already had.
///
/// Permissions on success: an existing `path` keeps its own mode, and a
/// path created by this call gets `0666` masked by the process umask, the
/// same mode a plain `open(path, O_CREAT|O_WRONLY, 0666)` would produce.
/// The temporary's own restrictive creation mode is never observable.
fm_status_t write_file_atomically(const std::string& path, const std::uint8_t* bytes, std::size_t len);

/// Names `wb` after the last component of `path` (`fm_workbook_set_file_name`),
/// as Excel names a workbook it opened; a name the setter refuses, such as one
/// without a workbook extension, leaves the workbook unnamed.
void set_file_name_from_path(fm_workbook_t* wb, const std::string& path);

}  // namespace formulon::cli

#endif  // FORMULON_CLI_FILE_IO_H_
