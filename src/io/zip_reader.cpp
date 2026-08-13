//
// `ZipReader` implementation. Wraps the miniz `mz_zip_reader_*` API in a
// PIMPL so the header is miniz-free. All methods are non-throwing and
// surface errors via `Expected<T, Error>` per Formulon convention.

#include "io/zip_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <memory>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "miniz.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"

namespace formulon {
namespace io {

namespace {

FormulonErrorCode ErrorCodeForMiniz(mz_zip_error error) {
  return error == MZ_ZIP_UNSUPPORTED_ENCRYPTION ? FormulonErrorCode::kIoZipEncrypted : FormulonErrorCode::kIoZipCorrupt;
}

}  // namespace

struct ZipReader::Impl {
  mz_zip_archive archive{};
  bool open = false;
  // Reusable scratch buffer for filename queries. miniz's
  // `mz_zip_reader_get_filename` is the cheapest way to probe a name; it
  // requires a caller-supplied buffer. We keep one resident on the impl
  // so `entry_name` does not allocate per call.
  mutable std::string name_buf;
  // Running charge for every extraction attempt in the current open
  // session. `ResourceBudget` deliberately charges before allocation, so a
  // failed extraction remains charged and cannot be retried indefinitely.
  mutable ResourceBudget total_extracted_bytes;

  Impl() : total_extracted_bytes(kMaxTotalExtractedBytes, FormulonErrorCode::kIoFileTooLarge) {
    name_buf.resize(kInitialNameCapacity);
  }

  // The miniz archive holds internal pointers into the user's input
  // buffer (after `init_mem`); copying or moving the struct would
  // silently corrupt those references. Make the breakage compile-time
  // instead by disabling all four; ownership is conveyed through the
  // outer `ZipReader`'s `unique_ptr<Impl>` instead.
  Impl(const Impl&) = delete;
  Impl& operator=(const Impl&) = delete;
  Impl(Impl&&) = delete;
  Impl& operator=(Impl&&) = delete;

  ~Impl() { close(); }

  void close() noexcept {
    if (open) {
      // best-effort: the destructor never returns a failure
      mz_zip_reader_end(&archive);
      open = false;
    }
  }

  // OOXML part names are typically <100 bytes; pre-size accordingly so the
  // common case avoids any growth. The buffer auto-grows in `entry_name`
  // when miniz reports a longer name.
  static constexpr std::size_t kInitialNameCapacity = 256;
};

ZipReader::ZipReader() : impl_(std::make_unique<Impl>()) {}
ZipReader::~ZipReader() = default;
ZipReader::ZipReader(ZipReader&&) noexcept = default;
ZipReader& ZipReader::operator=(ZipReader&&) noexcept = default;

Expected<void, Error> ZipReader::open(ByteSpan bytes, std::size_t max_total_extracted_bytes) {
  // Reset any previously opened archive first so callers can reuse the
  // instance (idempotent open).
  impl_->close();
  // Clamped, so the parameter can only tighten the bound.
  const std::size_t ceiling = std::min(max_total_extracted_bytes, kMaxTotalExtractedBytes);
  impl_->total_extracted_bytes = ResourceBudget(ceiling, FormulonErrorCode::kIoFileTooLarge);

  if (bytes.data == nullptr || bytes.size == 0) {
    return make_error(FormulonErrorCode::kIoZipCorrupt, "ZipReader::open: empty buffer",
                      "context=zip_reader bytes_size=0");
  }

  // miniz's mz_zip_archive must be zero-initialised before init.
  std::memset(&impl_->archive, 0, sizeof(impl_->archive));

  if (mz_zip_reader_init_mem(&impl_->archive, bytes.data, bytes.size, /*flags=*/0) == MZ_FALSE) {
    const mz_zip_error err = mz_zip_get_last_error(&impl_->archive);
    std::string ctx("context=zip_reader miniz_error=");
    ctx.append(std::to_string(static_cast<int>(err)));
    return make_error(ErrorCodeForMiniz(err), "ZipReader::open: mz_zip_reader_init_mem failed", std::move(ctx));
  }

  // Reject archives whose central directory advertises an unreasonable
  // number of parts before we hand a usable reader back. Doing this at
  // open-time means callers never start iterating through entries that
  // the policy would refuse anyway, and the cap is enforced even if the
  // caller never exercises `read_entry`.
  const mz_uint num_files = mz_zip_reader_get_num_files(&impl_->archive);
  if (static_cast<std::size_t>(num_files) > kMaxParts) {
    mz_zip_reader_end(&impl_->archive);
    std::string ctx("context=zip_reader limit=parts num_files=");
    ctx.append(std::to_string(static_cast<std::uint64_t>(num_files)));
    ctx.append(" cap=");
    ctx.append(std::to_string(static_cast<std::uint64_t>(kMaxParts)));
    return make_error(FormulonErrorCode::kIoZipBomb, "ZipReader::open: archive entry count exceeds cap",
                      std::move(ctx));
  }

  // Preflight the complete central directory before exposing the archive to
  // callers. Duplicate non-directory names are rejected: miniz's lookup
  // policy for duplicate catalogue records is not a safe content-addressing
  // contract, and accepting first/last-wins ambiguity would let an attacker
  // swap the bytes of a critical or passthrough part between inventory and
  // extraction. Per-entry size and compression-ratio policies remain in
  // `read_entry`, where the returned buffer is actually allocated.
  std::unordered_set<std::string> seen_names;
  seen_names.reserve(static_cast<std::size_t>(num_files));
  for (mz_uint i = 0; i < num_files; ++i) {
    mz_zip_archive_file_stat stat{};
    if (mz_zip_reader_file_stat(&impl_->archive, i, &stat) == MZ_FALSE) {
      const mz_zip_error err = mz_zip_get_last_error(&impl_->archive);
      mz_zip_reader_end(&impl_->archive);
      std::string ctx("context=zip_reader preflight miniz_error=");
      ctx.append(std::to_string(static_cast<int>(err)));
      return make_error(ErrorCodeForMiniz(err), "ZipReader::open: central-directory preflight failed", std::move(ctx));
    }
    if (stat.m_is_directory) {
      continue;
    }
    const mz_uint needed = mz_zip_reader_get_filename(&impl_->archive, i, nullptr, 0);
    if (needed == 0U) {
      mz_zip_reader_end(&impl_->archive);
      return make_error(FormulonErrorCode::kIoZipCorrupt, "ZipReader::open: central-directory entry has no filename",
                        "context=zip_reader preflight");
    }
    if (impl_->name_buf.size() < needed) {
      impl_->name_buf.resize(needed);
    }
    const mz_uint written = mz_zip_reader_get_filename(&impl_->archive, i, impl_->name_buf.data(),
                                                       static_cast<mz_uint>(impl_->name_buf.size()));
    if (written == 0U) {
      mz_zip_reader_end(&impl_->archive);
      return make_error(FormulonErrorCode::kIoZipCorrupt, "ZipReader::open: central-directory filename lookup failed",
                        "context=zip_reader preflight");
    }
    std::size_t name_len = written;
    if (name_len > 0U && impl_->name_buf[name_len - 1U] == '\0') {
      --name_len;
    }
    if (!seen_names.emplace(impl_->name_buf.data(), name_len).second) {
      std::string ctx("context=zip_reader duplicate_entry=");
      ctx.append(impl_->name_buf.data(), name_len);
      mz_zip_reader_end(&impl_->archive);
      return make_error(FormulonErrorCode::kIoZipCorrupt, "ZipReader::open: duplicate non-directory entry name",
                        std::move(ctx));
    }
  }

  impl_->open = true;
  return Expected<void, Error>::Ok();
}

std::size_t ZipReader::entry_count() const noexcept {
  if (!impl_->open) {
    return 0;
  }
  return static_cast<std::size_t>(mz_zip_reader_get_num_files(&impl_->archive));
}

std::string_view ZipReader::entry_name(std::size_t index) const noexcept {
  if (!impl_->open) {
    return {};
  }
  const mz_uint count = mz_zip_reader_get_num_files(&impl_->archive);
  if (index >= static_cast<std::size_t>(count)) {
    return {};
  }
  const auto idx = static_cast<mz_uint>(index);

  // Probe the filename length first. miniz returns the byte length
  // including the NUL terminator; passing 0 as the buffer size is the
  // documented way to query length without writing anything.
  const mz_uint needed = mz_zip_reader_get_filename(&impl_->archive, idx, nullptr, 0);
  if (needed == 0) {
    return {};
  }
  if (impl_->name_buf.size() < needed) {
    impl_->name_buf.resize(needed);
  }
  const mz_uint written = mz_zip_reader_get_filename(&impl_->archive, idx, impl_->name_buf.data(),
                                                     static_cast<mz_uint>(impl_->name_buf.size()));
  if (written == 0) {
    return {};
  }
  // miniz includes the terminating NUL in the returned length; trim it for
  // the string_view we hand back.
  std::size_t len = written;
  if (len > 0 && impl_->name_buf[len - 1] == '\0') {
    --len;
  }
  return std::string_view{impl_->name_buf.data(), len};
}

bool ZipReader::has_entry(std::string_view name) const noexcept {
  if (!impl_->open || name.empty()) {
    return false;
  }
  // mz_zip_reader_locate_file requires a NUL-terminated C string. We pass
  // MZ_ZIP_FLAG_CASE_SENSITIVE so OOXML part names compare exactly --
  // miniz's default is case-insensitive (a Windows-ism hold-over) which
  // would let `XL/WORKBOOK.XML` match a real `xl/workbook.xml` entry.
  // `unique_ptr::operator->` returns a non-const pointer regardless of
  // outer constness, so `&impl_->archive` is already `mz_zip_archive*`
  // and no cast is needed.
  std::string c_name(name);
  const int idx = mz_zip_reader_locate_file(&impl_->archive, c_name.c_str(), nullptr,
                                            /*flags=*/MZ_ZIP_FLAG_CASE_SENSITIVE);
  return idx >= 0;
}

Expected<std::vector<std::uint8_t>, Error> ZipReader::read_entry(std::string_view name) const {
  if (!impl_->open) {
    return make_error(FormulonErrorCode::kIoZipCorrupt, "ZipReader::read_entry: archive not open",
                      "context=zip_reader entry=" + std::string(name));
  }

  std::string c_name(name);
  const int located = mz_zip_reader_locate_file(&impl_->archive, c_name.c_str(), nullptr,
                                                /*flags=*/MZ_ZIP_FLAG_CASE_SENSITIVE);
  if (located < 0) {
    return make_error(FormulonErrorCode::kIoFileNotFound, "ZipReader::read_entry: entry not found",
                      "context=zip_reader entry=" + c_name);
  }

  // Probe the central directory entry first so we can reject ZIP bombs
  // before miniz allocates the full uncompressed size on the heap. A
  // crafted entry can claim 99:1 compression: a few KB of archive bytes
  // expanding into hundreds of MB of memory. Refuse anything past
  // `kMaxExtractedBytesPerEntry`.
  mz_zip_archive_file_stat stat{};
  if (mz_zip_reader_file_stat(&impl_->archive, static_cast<mz_uint>(located), &stat) == MZ_FALSE) {
    const mz_zip_error err = mz_zip_get_last_error(&impl_->archive);
    std::string ctx("context=zip_reader entry=");
    ctx.append(c_name);
    ctx.append(" miniz_error=");
    ctx.append(std::to_string(static_cast<int>(err)));
    return make_error(ErrorCodeForMiniz(err), "ZipReader::read_entry: stat failed", std::move(ctx));
  }
  if (stat.m_is_encrypted) {
    std::string ctx("context=zip_reader entry=");
    ctx.append(c_name);
    return make_error(FormulonErrorCode::kIoZipEncrypted, "ZipReader::read_entry: encrypted entry is unsupported",
                      std::move(ctx));
  }
  if (stat.m_uncomp_size > static_cast<mz_uint64>(kMaxExtractedBytesPerEntry)) {
    std::string ctx("context=zip_reader limit=entry entry=");
    ctx.append(c_name);
    ctx.append(" uncompressed_size=");
    ctx.append(std::to_string(static_cast<std::uint64_t>(stat.m_uncomp_size)));
    ctx.append(" cap=");
    ctx.append(std::to_string(static_cast<std::uint64_t>(kMaxExtractedBytesPerEntry)));
    return make_error(FormulonErrorCode::kIoZipBomb, "ZipReader::read_entry: uncompressed size exceeds per-entry cap",
                      std::move(ctx));
  }

  // Reject pathological compression ratios. The compressed size comes from
  // the central directory; it can be zero for an empty/malformed entry, so
  // keep the guard to avoid division by zero. A stored non-empty entry has a
  // non-zero compressed size equal to its uncompressed size.
  if (stat.m_comp_size > 0) {
    const std::uint64_t ratio =
        static_cast<std::uint64_t>(stat.m_uncomp_size) / static_cast<std::uint64_t>(stat.m_comp_size);
    if (ratio > static_cast<std::uint64_t>(kMaxRatio)) {
      std::string ctx("context=zip_reader limit=ratio entry=");
      ctx.append(c_name);
      ctx.append(" uncompressed_size=");
      ctx.append(std::to_string(static_cast<std::uint64_t>(stat.m_uncomp_size)));
      ctx.append(" compressed_size=");
      ctx.append(std::to_string(static_cast<std::uint64_t>(stat.m_comp_size)));
      ctx.append(" ratio=");
      ctx.append(std::to_string(ratio));
      ctx.append(" cap=");
      ctx.append(std::to_string(static_cast<std::uint64_t>(kMaxRatio)));
      return make_error(FormulonErrorCode::kIoZipBomb, "ZipReader::read_entry: compression ratio exceeds cap",
                        std::move(ctx));
    }
  }

  // Charge the advertised size before touching the destination allocation.
  // `ResourceBudget::consume()` is uint64/overflow-safe and leaves the used
  // count unchanged on refusal. A successful charge is intentionally not
  // rolled back if miniz later rejects the bytes: failed extraction attempts
  // still consume the session's bounded work budget.
  auto charged = impl_->total_extracted_bytes.consume(static_cast<std::uint64_t>(stat.m_uncomp_size));
  if (!charged) {
    const std::string_view budget_context = charged.error().context;
    std::string context("limit=total entry=");
    context.append(c_name);
    if (!budget_context.empty()) {
      context.push_back(' ');
      context.append(budget_context);
    }
    Error error = charged.error();
    error.message = "ZipReader::read_entry: cumulative extracted bytes exceed cap";
    error.context = std::move(context);
    return error;
  }

  const std::size_t advertised = static_cast<std::size_t>(stat.m_uncomp_size);

  // Extract straight into the caller-owned vector. `advertised` is the
  // stat's `m_uncomp_size`, already validated against the per-entry, ratio,
  // and cumulative caps above. Sizing the destination once and
  // decompressing into it avoids the second full-size buffer that
  // `extract_to_heap` + `memcpy` + `mz_free` would hold simultaneously,
  // halving the extraction's peak footprint.
  std::vector<std::uint8_t> bytes;
  bytes.resize(advertised);
  if (advertised > 0) {
    if (mz_zip_reader_extract_to_mem(&impl_->archive, static_cast<mz_uint>(located), bytes.data(), bytes.size(),
                                     /*flags=*/0) == MZ_FALSE) {
      const mz_zip_error err = mz_zip_get_last_error(&impl_->archive);
      std::string ctx("context=zip_reader entry=");
      ctx.append(c_name);
      ctx.append(" miniz_error=");
      ctx.append(std::to_string(static_cast<int>(err)));
      return make_error(ErrorCodeForMiniz(err), "ZipReader::read_entry: extraction failed", std::move(ctx));
    }
  }
  return bytes;
}

std::vector<std::string> ZipReader::list_entries() const {
  std::vector<std::string> out;
  const std::size_t n = entry_count();
  out.reserve(n);
  for (std::size_t i = 0; i < n; ++i) {
    out.emplace_back(entry_name(i));
  }
  return out;
}

}  // namespace io
}  // namespace formulon
