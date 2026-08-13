//
// Shared file I/O primitives for native CLI commands. The C ABI remains
// bytes-in / bytes-out; only the CLI maps paths to byte buffers.

#ifndef FORMULON_CLI_FILE_IO_H_
#define FORMULON_CLI_FILE_IO_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"

namespace formulon::cli {

/// Reads an entire binary file into `out`.
fm_status_t read_file(const std::string& path, std::vector<std::uint8_t>& out);

/// Atomically replaces `path` with `bytes`. The temporary file is created
/// exclusively with a random suffix in the target directory, preserving the
/// atomic-rename guarantee and preventing predictable-name collisions.
///
/// A `path` that is a symlink is resolved first, so the replace lands on
/// the file the link names and the link itself survives. A dangling link
/// is replaced as-is: there is no target to keep.
///
/// Permissions on success: an existing `path` keeps its own mode, and a
/// path created by this call gets `0666` masked by the process umask, the
/// same mode a plain `open(path, O_CREAT|O_WRONLY, 0666)` would produce.
/// The temporary's own restrictive creation mode is never observable.
fm_status_t write_file_atomically(const std::string& path, const std::uint8_t* bytes, std::size_t len);

}  // namespace formulon::cli

#endif  // FORMULON_CLI_FILE_IO_H_
