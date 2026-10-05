//
// Implementation of the miniz part-add wrappers.

#include "io/ooxml/zip_part_writer.h"

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "miniz.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

namespace {

/// Returns an `Error` when `path` is already in `*seen_paths`, else
/// records it and returns success. A no-op (always succeeds) when
/// `seen_paths` is null, so callers that have not opted into the guard
/// see no behaviour change.
Expected<void, Error> CheckAndRecordPath(std::string_view path, std::unordered_set<std::string>* seen_paths) {
  if (seen_paths == nullptr) {
    return Expected<void, Error>::Ok();
  }
  if (!seen_paths->emplace(path).second) {
    std::string context("part=");
    context.append(path);
    return make_error(FormulonErrorCode::kIoWriteFailed, "duplicate zip entry path", std::move(context));
  }
  return Expected<void, Error>::Ok();
}

/// Adds `size` bytes at `data` as entry `path`; `failure` is the error
/// message when miniz rejects the entry.
Expected<void, Error> AddEntry(mz_zip_archive* archive, std::string_view path, const void* data, std::size_t size,
                               std::unordered_set<std::string>* seen_paths, const char* failure) {
  if (auto dup_check = CheckAndRecordPath(path, seen_paths); !dup_check) {
    return dup_check.error();
  }
  const mz_bool ok = mz_zip_writer_add_mem(archive, std::string(path).c_str(), data, size,
                                           static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
  if (ok == MZ_FALSE) {
    std::string context("part=");
    context.append(path);
    return make_error(FormulonErrorCode::kIoWriteFailed, failure, std::move(context));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<void, Error> AddPart(mz_zip_archive* archive, std::string_view path, const std::string& body,
                              std::unordered_set<std::string>* seen_paths) {
  return AddEntry(archive, path, body.data(), body.size(), seen_paths, "miniz mz_zip_writer_add_mem failed");
}

Expected<void, Error> AddPartBytes(mz_zip_archive* archive, std::string_view path,
                                   const std::vector<std::uint8_t>& body, std::unordered_set<std::string>* seen_paths) {
  return AddEntry(archive, path, body.data(), body.size(), seen_paths,
                  "miniz mz_zip_writer_add_mem failed (binary part)");
}

Expected<std::vector<std::uint8_t>, Error> FinalizeArchive(ZipWriterGuard& writer, const char* context) {
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  if (mz_zip_writer_finalize_heap_archive(writer.get(), &archive_ptr, &archive_size) == MZ_FALSE) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_finalize_heap_archive failed",
                      std::string(context));
  }
  if (mz_zip_writer_end(writer.get()) == MZ_FALSE) {
    // Finalize succeeded but end failed: still free the buffer miniz handed
    // us before surfacing the error.
    if (archive_ptr != nullptr) {
      mz_free(archive_ptr);
    }
    writer.release();
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_end failed", std::string(context));
  }
  writer.release();

  // Copy into a std::vector so the caller owns the bytes through normal RAII.
  std::vector<std::uint8_t> bytes;
  bytes.resize(archive_size);
  if (archive_size > 0 && archive_ptr != nullptr) {
    std::memcpy(bytes.data(), archive_ptr, archive_size);
  }
  if (archive_ptr != nullptr) {
    mz_free(archive_ptr);
  }
  return bytes;
}

}  // namespace io
}  // namespace formulon
