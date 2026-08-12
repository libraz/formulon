//
// Implementation of the miniz part-add wrappers.

#include "io/ooxml/zip_part_writer.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "miniz.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

Expected<void, Error> AddPart(mz_zip_archive* archive, std::string_view path, const std::string& body) {
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
                                   const std::vector<std::uint8_t>& body) {
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
