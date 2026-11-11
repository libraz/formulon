
#include "cli/file_io.h"

#include <fcntl.h>
#include <sys/stat.h>
#include <unistd.h>

#include <algorithm>
#include <cerrno>
#include <climits>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <ios>
#include <limits>
#include <string>
#include <vector>

namespace formulon::cli {

namespace {

/// Reads descriptor `fd` to the end into `out`.
fm_status_t read_descriptor(int fd, std::vector<std::uint8_t>& out) {
  out.clear();
  constexpr std::size_t kChunk = 1U << 16;
  for (;;) {
    const std::size_t used = out.size();
    out.resize(used + kChunk);
    const ssize_t count = ::read(fd, out.data() + used, kChunk);
    if (count < 0) {
      if (errno == EINTR) {
        out.resize(used);
        continue;
      }
      return kExitIoError;
    }
    out.resize(used + static_cast<std::size_t>(count));
    if (count == 0) {
      return 0;
    }
  }
}

}  // namespace

fm_status_t read_file(const std::string& path, std::vector<std::uint8_t>& out) {
  const int fd = ::open(path.c_str(), O_RDONLY);
  if (fd < 0) {
    return kExitNoInput;
  }
  struct stat info {};
  if (::fstat(fd, &info) == 0 && S_ISDIR(info.st_mode)) {
    ::close(fd);
    errno = EISDIR;
    return kExitNoInput;
  }
  const fm_status_t rc = read_descriptor(fd, out);
  const int saved_errno = errno;
  ::close(fd);
  errno = saved_errno;
  return rc;
}

fm_status_t read_stdin(std::vector<std::uint8_t>& out) {
  if (::isatty(STDIN_FILENO) != 0) {
    return kExitUsage;
  }
  return read_descriptor(STDIN_FILENO, out);
}

bool stdout_is_terminal() {
  return ::isatty(STDOUT_FILENO) != 0;
}

namespace {

/// Writes all `len` bytes to `fd`, retrying on EINTR. A write that makes no
/// progress fails rather than looping.
bool write_all(int fd, const std::uint8_t* bytes, std::size_t len) {
  std::size_t written = 0;
  while (written < len) {
    const std::size_t chunk = std::min(len - written, static_cast<std::size_t>(std::numeric_limits<ssize_t>::max()));
    const ssize_t count = ::write(fd, bytes + written, chunk);
    if (count < 0) {
      if (errno == EINTR) {
        continue;
      }
      return false;
    }
    if (count == 0) {
      return false;
    }
    written += static_cast<std::size_t>(count);
  }
  return true;
}

}  // namespace

fm_status_t write_stdout(const std::uint8_t* bytes, std::size_t len) {
  return write_all(STDOUT_FILENO, bytes, len) ? 0 : kExitIoError;
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

/// Owns an open descriptor for the rest of its scope.
///
/// The write sequence below has several failure points between `mkstemp`
/// and the final `close`, and every one of them returns early. Tying the
/// close to the scope rather than to each `return` is what keeps a new
/// failure branch from leaking the descriptor the way a hand-written
/// close on each path eventually does.
class FdGuard {
 public:
  explicit FdGuard(int fd) : fd_(fd) {}
  FdGuard(const FdGuard&) = delete;
  FdGuard& operator=(const FdGuard&) = delete;
  ~FdGuard() {
    if (fd_ >= 0) {
      ::close(fd_);
    }
  }

  int get() const { return fd_; }

  /// Hands ownership back so the caller can close explicitly and inspect
  /// the result: a failing `close` is sometimes the first report of a
  /// write error, and the destructor cannot report anything.
  int release() {
    const int owned = fd_;
    fd_ = -1;
    return owned;
  }

 private:
  int fd_;
};

}  // namespace

fm_status_t write_file_atomically(const std::string& link_or_path, const std::uint8_t* bytes, std::size_t len) {
  const std::string path = resolve_link_target(link_or_path);
  struct stat target_stat {};
  const bool preserve_existing_mode = ::stat(path.c_str(), &target_stat) == 0;
  std::string tmp_path = path + ".formulon-tmp.XXXXXX";
  FdGuard tmp_fd(::mkstemp(tmp_path.data()));
  if (tmp_fd.get() < 0) {
    return kExitCantCreate;
  }

  // Removing the temp is all a failure path has to do by hand; `tmp_fd`
  // closes the descriptor when the function returns.
  const auto fail = [&] {
    const int saved_errno = errno;
    std::remove(tmp_path.c_str());
    errno = saved_errno;
    return static_cast<fm_status_t>(kExitIoError);
  };
  // Replacing a file keeps its identity, including its permissions; creating
  // one gives what an ordinary `open(path, O_CREAT|O_WRONLY, 0666)` would
  // give under the current umask. Either way the 0600 that `mkstemp` forces
  // on its temporary must not be observable in the final artifact.
  const mode_t result_mode = preserve_existing_mode
                                 ? static_cast<mode_t>(target_stat.st_mode & 0777U)
                                 : static_cast<mode_t>(0666U & ~static_cast<unsigned int>(kProcessUmask));
  if (::fchmod(tmp_fd.get(), result_mode) != 0) {
    return fail();
  }

  if (!write_all(tmp_fd.get(), bytes, len)) {
    return fail();
  }
  if (::fsync(tmp_fd.get()) != 0) {
    return fail();
  }
  // POSIX leaves the descriptor's state unspecified after a failed
  // `close`, so ownership is given up before the call either way.
  if (::close(tmp_fd.release()) != 0) {
    return fail();
  }

  if (std::rename(tmp_path.c_str(), path.c_str()) != 0) {
    const int saved_errno = errno;
    std::remove(tmp_path.c_str());
    errno = saved_errno;
    return kExitCantCreate;
  }
  return 0;
}

void set_location_from_path(fm_workbook_t* wb, const std::string& path) {
  std::string absolute = path;
  if (absolute.empty() || absolute.front() != '/') {
    std::vector<char> cwd(PATH_MAX + 1, '\0');
    if (::getcwd(cwd.data(), cwd.size()) == nullptr) {
      (void)fm_workbook_set_location(wb, nullptr);
      return;
    }
    absolute = std::string(cwd.data()) + "/" + absolute;
  }
  // Lexical: `.` and empty segments drop, `..` climbs and stops at the root.
  std::vector<std::string> segments;
  std::size_t start = 0;
  while (start <= absolute.size()) {
    const std::size_t end = std::min(absolute.find('/', start), absolute.size());
    const std::string segment = absolute.substr(start, end - start);
    if (segment == "..") {
      if (!segments.empty()) {
        segments.pop_back();
      }
    } else if (!segment.empty() && segment != ".") {
      segments.push_back(segment);
    }
    start = end + 1;
  }
  if (segments.empty()) {
    (void)fm_workbook_set_location(wb, nullptr);
    return;
  }
  std::string directory;
  for (std::size_t i = 0; i + 1 < segments.size(); ++i) {
    directory += "/" + segments[i];
  }
  if (directory.empty()) {
    directory = "/";
  }
  fm_workbook_location_t location = FM_STRUCT_INIT(fm_workbook_location_t);
  location.directory = directory.c_str();
  location.file_name = segments.back().c_str();
  // A refused location leaves the workbook without one, which is not an error here.
  if (fm_workbook_set_location(wb, &location) != 0) {
    (void)fm_workbook_set_location(wb, nullptr);
  }
}

}  // namespace formulon::cli
