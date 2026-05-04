// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `ZipReader` implementation. Wraps the miniz `mz_zip_reader_*` API in a
// PIMPL so the header is miniz-free. All methods are non-throwing and
// surface errors via `Expected<T, Error>` per Formulon convention.

#include "io/zip_reader.h"

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "miniz.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

namespace {

// Hard cap on the uncompressed size of any single ZIP entry. Excel parts
// rarely exceed a few MB; 100 MiB is well above any realistic per-part
// payload (sharedStrings.xml and large sheets are the worst offenders and
// stay under ~30 MB even on million-cell workbooks). The cap blunts ZIP
// bombs whose 99:1 compression ratio would otherwise allow a few KB of
// archive bytes to expand into multi-GiB allocations and trigger OOM /
// DoS during `read_entry`.
constexpr std::size_t kMaxExtractedBytesPerEntry = 100ULL * 1024ULL * 1024ULL;

}  // namespace

struct ZipReader::Impl {
  mz_zip_archive archive{};
  bool open = false;
  // Reusable scratch buffer for filename queries. miniz's
  // `mz_zip_reader_get_filename` is the cheapest way to probe a name; it
  // requires a caller-supplied buffer. We keep one resident on the impl
  // so `entry_name` does not allocate per call.
  mutable std::string name_buf;

  Impl() { name_buf.resize(kInitialNameCapacity); }

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

Expected<void, Error> ZipReader::open(ByteSpan bytes) {
  // Reset any previously opened archive first so callers can reuse the
  // instance (idempotent open).
  impl_->close();

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
    return make_error(FormulonErrorCode::kIoZipCorrupt, "ZipReader::open: mz_zip_reader_init_mem failed",
                      std::move(ctx));
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
    return make_error(FormulonErrorCode::kIoZipCorrupt, "ZipReader::read_entry: stat failed", std::move(ctx));
  }
  if (stat.m_uncomp_size > static_cast<mz_uint64>(kMaxExtractedBytesPerEntry)) {
    std::string ctx("context=zip_reader entry=");
    ctx.append(c_name);
    ctx.append(" uncompressed_size=");
    ctx.append(std::to_string(static_cast<std::uint64_t>(stat.m_uncomp_size)));
    ctx.append(" cap=");
    ctx.append(std::to_string(static_cast<std::uint64_t>(kMaxExtractedBytesPerEntry)));
    return make_error(FormulonErrorCode::kIoZipBomb, "ZipReader::read_entry: uncompressed size exceeds per-entry cap",
                      std::move(ctx));
  }

  std::size_t extracted_size = 0;
  void* extracted = mz_zip_reader_extract_to_heap(&impl_->archive, static_cast<mz_uint>(located), &extracted_size,
                                                  /*flags=*/0);
  if (extracted == nullptr) {
    const mz_zip_error err = mz_zip_get_last_error(&impl_->archive);
    std::string ctx("context=zip_reader entry=");
    ctx.append(c_name);
    ctx.append(" miniz_error=");
    ctx.append(std::to_string(static_cast<int>(err)));
    return make_error(FormulonErrorCode::kIoZipCorrupt, "ZipReader::read_entry: extraction failed", std::move(ctx));
  }

  std::vector<std::uint8_t> bytes;
  bytes.resize(extracted_size);
  if (extracted_size > 0) {
    std::memcpy(bytes.data(), extracted, extracted_size);
  }
  mz_free(extracted);
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
