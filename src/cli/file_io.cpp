
#include "cli/file_io.h"

#include <sys/stat.h>
#include <unistd.h>

#include <algorithm>
#include <cerrno>
#include <cstdio>
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

fm_status_t write_file_atomically(const std::string& path, const std::uint8_t* bytes, std::size_t len) {
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
  if (preserve_existing_mode && ::fchmod(fd, target_stat.st_mode & 0777U) != 0) {
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
