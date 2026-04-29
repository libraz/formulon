// Copyright 2026 libraz. Licensed under the MIT License.
//
// Read-only ZIP archive accessor. The OOXML reader pipeline opens an
// `.xlsx` package by handing the raw bytes to a `ZipReader`, then
// requesting individual parts (`xl/workbook.xml`, sheet rels, etc.)
// by name. Higher-level OOXML semantics live in `ooxml_reader.{h,cpp}`;
// this module only knows about ZIP central directories and entry
// names. miniz is hidden behind a PIMPL so the header stays free of
// third-party includes.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.4 (Reader pipeline)
//   * backup/plans/04-xlsx-io.md §4.4.4 (zip64, BOM, malformed input)

#ifndef FORMULON_IO_ZIP_READER_H_
#define FORMULON_IO_ZIP_READER_H_

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// Non-owning view over a contiguous byte buffer. C++17 lacks `std::span`,
/// so the I/O layer carries this minimal stand-in. The pointer/size pair
/// must remain valid for the lifetime of any consumer that holds it; the
/// view does not copy or own the underlying memory.
struct ByteSpan {
  const std::uint8_t* data = nullptr;
  std::size_t size = 0;
};

/// Read-only accessor for a ZIP archive backed by an in-memory buffer.
///
/// The archive bytes are *not* copied: the buffer passed to `open()` must
/// outlive the `ZipReader`. This matches the OOXML reader's typical
/// flow (load file once, parse parts directly off the loaded buffer)
/// and avoids a second large allocation.
///
/// The reader is move-only (it holds a miniz handle). Constructing a
/// `ZipReader` does not open anything; call `open()` and inspect the
/// returned `Expected<void, Error>` before invoking any other method.
/// Calling lookup / read methods on a never-opened reader returns
/// "no such entry" results without a crash, but is not the supported
/// way to use the API.
class ZipReader {
 public:
  ZipReader();
  ~ZipReader();
  ZipReader(const ZipReader&) = delete;
  ZipReader& operator=(const ZipReader&) = delete;
  ZipReader(ZipReader&&) noexcept;
  ZipReader& operator=(ZipReader&&) noexcept;

  /// Opens the archive backed by `bytes`. The buffer must outlive this
  /// reader. Returns `FormulonErrorCode::kIoZipCorrupt` if miniz cannot
  /// initialise from the buffer (truncated central directory, wrong
  /// magic, etc.). Calling `open()` a second time on the same instance
  /// closes the previous archive first.
  Expected<void, Error> open(ByteSpan bytes);

  /// Number of entries in the open archive. Returns 0 if the reader has
  /// not been opened.
  std::size_t entry_count() const noexcept;

  /// Returns the part name (UTF-8) at `index`. Names use forward slashes
  /// (OOXML convention). Returns an empty `string_view` when `index >=
  /// entry_count()`. The returned view points into a buffer owned by the
  /// reader and is invalidated by the next `entry_name(...)` call on the
  /// same instance — copy into a `std::string` if you need to retain it.
  std::string_view entry_name(std::size_t index) const noexcept;

  /// Whether `name` exists in the open archive. Case-sensitive, exact
  /// match. Returns `false` if the reader has not been opened.
  bool has_entry(std::string_view name) const noexcept;

  /// Reads the entire decompressed contents of `name` into a freshly
  /// allocated buffer. Returns `kIoFileNotFound` when the entry is
  /// absent and `kIoZipCorrupt` on miniz extraction failure (e.g.
  /// stored size disagrees with central directory). The returned buffer
  /// is independent of the underlying ZIP bytes.
  Expected<std::vector<std::uint8_t>, Error> read_entry(std::string_view name) const;

  /// Returns every entry name in archive order. Convenience wrapper
  /// around `entry_count()` + `entry_name(i)` that materialises the
  /// names into owning strings so callers can hold them past further
  /// reader calls.
  std::vector<std::string> list_entries() const;

 private:
  // PIMPL: the miniz handle and a small scratch buffer for filename
  // queries live behind this pointer so the public header does not
  // include `miniz.h`.
  struct Impl;
  std::unique_ptr<Impl> impl_;
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_ZIP_READER_H_
