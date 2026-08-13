//
// Implementation of the miniz part-add wrappers.

#include "io/ooxml/zip_part_writer.h"

#include <cstdint>
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

}  // namespace

Expected<void, Error> AddPart(mz_zip_archive* archive, std::string_view path, const std::string& body,
                              std::unordered_set<std::string>* seen_paths) {
  if (auto dup_check = CheckAndRecordPath(path, seen_paths); !dup_check) {
    return dup_check.error();
  }
  const mz_bool ok = mz_zip_writer_add_mem(archive, std::string(path).c_str(), body.data(), body.size(),
                                           static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
  if (ok == MZ_FALSE) {
    std::string context("part=");
    context.append(path);
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_add_mem failed", std::move(context));
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> AddPartBytes(mz_zip_archive* archive, std::string_view path,
                                   const std::vector<std::uint8_t>& body, std::unordered_set<std::string>* seen_paths) {
  if (auto dup_check = CheckAndRecordPath(path, seen_paths); !dup_check) {
    return dup_check.error();
  }
  const mz_bool ok = mz_zip_writer_add_mem(archive, std::string(path).c_str(), body.data(), body.size(),
                                           static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
  if (ok == MZ_FALSE) {
    std::string context("part=");
    context.append(path);
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_add_mem failed (binary part)",
                      std::move(context));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace io
}  // namespace formulon
