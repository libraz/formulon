//
// miniz wrappers used by the OOXML and XLSB writers to add parts to the
// in-memory zip archive. The RAII guard ensures the writer state is torn down
// even on the error path. Internal to `src/io/`; not part of the public API.

#ifndef FORMULON_IO_OOXML_ZIP_PART_WRITER_H_
#define FORMULON_IO_OOXML_ZIP_PART_WRITER_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "miniz.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// RAII guard around an initialised `mz_zip_archive` writer. The destructor
/// releases any heap buffer retained by miniz when the writer is abandoned
/// mid-flight (e.g. an `mz_zip_writer_add_mem` call failed and we early-
/// returned an error).
class ZipWriterGuard {
 public:
  ZipWriterGuard() = default;
  ZipWriterGuard(const ZipWriterGuard&) = delete;
  ZipWriterGuard& operator=(const ZipWriterGuard&) = delete;
  ZipWriterGuard(ZipWriterGuard&&) = delete;
  ZipWriterGuard& operator=(ZipWriterGuard&&) = delete;

  ~ZipWriterGuard() {
    if (active_) {
      // Best-effort cleanup; we're already on an error path.
      mz_zip_writer_end(&archive_);
    }
  }

  bool init() {
    if (mz_zip_writer_init_heap(&archive_, /*size_to_reserve_at_beginning=*/0,
                                /*initial_allocation_size=*/8 * 1024) == MZ_FALSE) {
      return false;
    }
    active_ = true;
    return true;
  }

  mz_zip_archive* get() noexcept { return &archive_; }

  /// Releases ownership of the underlying archive to the caller. Subsequent
  /// destruction no longer touches miniz state.
  void release() noexcept { active_ = false; }

 private:
  mz_zip_archive archive_{};
  bool active_ = false;
};

/// Adds a single text part to the archive. Returns an `Error` tagged
/// with the part path when miniz refuses the write.
///
/// `seen_paths`, when non-null, is both consulted and updated: a path
/// already present is refused with an `Error` instead of silently
/// producing a duplicate zip entry (which reader implementations
/// resolve inconsistently). Callers that want this guard own the set's
/// lifetime across a single archive's worth of `AddPart` /
/// `AddPartBytes` calls; passing `nullptr` (the default) skips the
/// check, which existing callers that have not opted in still rely on.
Expected<void, Error> AddPart(mz_zip_archive* archive, std::string_view path, const std::string& body,
                              std::unordered_set<std::string>* seen_paths = nullptr);

/// Adds a binary part — a passthrough blob, or an XLSB record stream.
/// Same error contract and `seen_paths` behaviour as `AddPart`.
Expected<void, Error> AddPartBytes(mz_zip_archive* archive, std::string_view path,
                                   const std::vector<std::uint8_t>& body,
                                   std::unordered_set<std::string>* seen_paths = nullptr);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_ZIP_PART_WRITER_H_
