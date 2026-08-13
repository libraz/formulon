
#include "cli/file_io.h"

#include <sys/stat.h>
#include <unistd.h>

#include <algorithm>
#include <cerrno>
#include <climits>
#include <cstdio>
#include <cstdlib>
#include <fstream>
#include <ios>
#include <limits>
#include <string>
#include <vector>

#include "utils/error.h"

namespace formulon::cli {

fm_status_t read_file(const std::string& path, std::vector<std::uint8_t>& out) {
  std::ifstream in(path, std::ios::binary);
  if (!in) {
    return static_cast<fm_status_t>(FormulonErrorCode::kCliFileNotFound);
  }
  in.seekg(0, std::ios::end);
  const std::streamsize size = in.tellg();
  if (size < 0) {
    return static_cast<fm_status_t>(FormulonErrorCode::kCliFileNotFound);
  }
  in.seekg(0, std::ios::beg);
  out.resize(static_cast<std::size_t>(size));
  if (size > 0) {
    in.read(reinterpret_cast<char*>(out.data()), size);
    if (!in) {
      return static_cast<fm_status_t>(FormulonErrorCode::kCliFileNotFound);
    }
  }
  return 0;
}

namespace {

/// Resolves `path` to the file a symlink chain ultimately names, so an
/// atomic replace updates that file instead of clobbering the link with
/// a regular file. Someone who saves through a symlink expects the link
/// to survive; replacing it also leaves the real workbook stale, which
/// is the silent-data-loss shape this write path exists to avoid.
///
/// Returns `path` unchanged when it is not a symlink, when it dangles
/// (there is no target to preserve), or when resolution fails for any
/// other reason — in each of those cases replacing the entry itself is
/// the correct outcome.
///
/// Resolving also settles the aliasing question for `recalc -o`: the
/// input is fully read and serialized before the rename, so an output
/// that resolves to the same inode as the input is the in-place case,
/// which the temp-then-rename sequence already handles.
std::string resolve_link_target(const std::string& path) {
  struct stat link_stat {};
  if (::lstat(path.c_str(), &link_stat) != 0 || !S_ISLNK(link_stat.st_mode)) {
    return path;
  }
  std::string resolved(PATH_MAX, '\0');
  if (::realpath(path.c_str(), resolved.data()) == nullptr) {
    return path;
  }
  resolved.resize(std::char_traits<char>::length(resolved.c_str()));
  return resolved;
}

/// The process umask, sampled once before `main` runs.
///
/// `umask(2)` has no read-only form: the only way to observe the mask is
/// to install a new one and put the old one back. Doing that during a
/// write would leave a window in which any concurrent file creation sees
/// mask 0, so the set-and-restore pair runs during static initialization
/// instead — before `main`, hence before any thread exists and before the
/// engine creates anything. The umask is inherited at exec and nothing in
/// the CLI changes it, so one sample stays accurate for the whole run and
/// the process mask is left exactly as it was found.
const mode_t kProcessUmask = [] {
  const mode_t previous = ::umask(0);
  ::umask(previous);
  return previous;
}();

}  // namespace

fm_status_t write_file_atomically(const std::string& link_or_path, const std::uint8_t* bytes, std::size_t len) {
  const std::string path = resolve_link_target(link_or_path);
  struct stat target_stat {};
  const bool preserve_existing_mode = ::stat(path.c_str(), &target_stat) == 0;
  std::string tmp_path = path + ".formulon-tmp.XXXXXX";
  const int fd = ::mkstemp(tmp_path.data());
  if (fd < 0) {
    return static_cast<fm_status_t>(FormulonErrorCode::kCliOutputFailed);
  }

  const auto fail = [&](bool close_fd) {
    if (close_fd) {
      ::close(fd);
    }
    std::remove(tmp_path.c_str());
    return static_cast<fm_status_t>(FormulonErrorCode::kCliOutputFailed);
  };
  // Replacing a file keeps its identity, including its permissions; creating
  // one gives what an ordinary `open(path, O_CREAT|O_WRONLY, 0666)` would
  // give under the current umask. Either way the 0600 that `mkstemp` forces
  // on its temporary must not be observable in the final artifact.
  const mode_t result_mode = preserve_existing_mode
                                 ? static_cast<mode_t>(target_stat.st_mode & 0777U)
                                 : static_cast<mode_t>(0666U & ~static_cast<unsigned int>(kProcessUmask));
  if (::fchmod(fd, result_mode) != 0) {
    return fail(/*close_fd=*/true);
  }

  std::size_t written = 0;
  while (written < len) {
    const std::size_t remaining = len - written;
    const std::size_t chunk = std::min(remaining, static_cast<std::size_t>(std::numeric_limits<ssize_t>::max()));
    const ssize_t count = ::write(fd, bytes + written, chunk);
    if (count < 0) {
      if (errno == EINTR) {
        continue;
      }
      return fail(/*close_fd=*/true);
    }
    if (count == 0) {
      return fail(/*close_fd=*/true);
    }
    written += static_cast<std::size_t>(count);
  }
  if (::fsync(fd) != 0 || ::close(fd) != 0) {
    return fail(/*close_fd=*/false);
  }

  if (std::rename(tmp_path.c_str(), path.c_str()) != 0) {
    std::remove(tmp_path.c_str());
    return static_cast<fm_status_t>(FormulonErrorCode::kCliOutputFailed);
  }
  return 0;
}

}  // namespace formulon::cli
