//
// `PartHandler`: future-facing abstraction for one OOXML part (the
// reader-writer pair for a `[Content_Types].xml` Override entry).
//
// Today, supporting a new OOXML part touches 4-5 files: a per-part
// `<part>_reader.{h,cpp}`, a `<part>_writer.{h,cpp}`, the central
// `ooxml_reader.cpp` + `ooxml_writer.cpp` dispatch, and at least one
// `Workbook` field. Each part also reimplements its own bytes-in /
// model-out / model-in / bytes-out boundary by hand.
//
// The eventual end-state collapses that into a single `PartHandler`
// implementation per part, registered in a process-wide list that the
// reader and writer iterate. Adding a new part becomes a single TU
// touching only this header + the new handler.
//
// This file currently introduces the interface only. The first wave of
// part migrations (likely `comments` and `tables`, both small enough to
// validate the contract) is scheduled as a follow-up; until then,
// existing reader / writer pairs continue to work unchanged.

#ifndef FORMULON_IO_PART_HANDLER_H_
#define FORMULON_IO_PART_HANDLER_H_

#include <string>
#include <string_view>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace io {

/// Abstract handler for one OOXML part.
///
/// A handler binds together:
///   * `content_type()` -- the `[Content_Types].xml` Override
///                         `ContentType=` value the package uses to
///                         identify this part. The dispatcher matches
///                         incoming Override entries by exact-string
///                         compare against this value.
///   * `read(bytes, wb)` -- decodes the raw decompressed part bytes
///                         and folds the resulting model into `wb`.
///                         Returns `Error` on parse failure; the
///                         dispatcher propagates that error to the
///                         top-level reader.
///   * `write(wb)`      -- serialises the relevant slice of `wb` back
///                         to a UTF-8 string. An empty string instructs
///                         the writer to skip emitting this part (the
///                         workbook does not currently carry any data
///                         for it). Returns `Error` on serialisation
///                         failure.
///
/// All three methods are stateless beyond the pure function contract:
/// the handler instance itself stores no per-call state, so a single
/// instance per content-type is shared across all reads / writes. A
/// process-wide registry hands them to the dispatcher in declaration
/// order.
///
/// Lifetime: handler instances have program lifetime (registered at
/// startup, never freed). The dispatcher holds non-owning pointers.
///
/// Thread-safety: all three methods must be safe to invoke concurrently
/// from multiple threads on the same instance, since the workbook
/// serialiser parallelises across parts. Implementations that need
/// per-call scratch state must allocate it on the stack or per-call.
class PartHandler {
 public:
  virtual ~PartHandler() = default;

  /// The OOXML `ContentType=` string this handler claims. Returned as a
  /// `string_view` into static storage; the dispatcher does not copy.
  virtual std::string_view content_type() const = 0;

  /// Decodes `bytes` (the raw decompressed part body, already stripped
  /// of zip framing) into the workbook model on `wb`. The handler may
  /// touch any field of `wb`; it must not assume single-threaded
  /// access to the dispatcher's queue but may assume the caller has
  /// serialised access to `wb` itself for the duration of the call.
  virtual Expected<void, Error> read(std::string_view bytes, Workbook& wb) const = 0;

  /// Serialises the relevant model slice of `wb` to a UTF-8 string
  /// suitable for writing into the package as the part's body. An
  /// empty string instructs the writer to skip the part entirely (the
  /// workbook carries no content for it).
  virtual Expected<std::string, Error> write(const Workbook& wb) const = 0;

 protected:
  PartHandler() = default;
  PartHandler(const PartHandler&) = delete;
  PartHandler& operator=(const PartHandler&) = delete;
  PartHandler(PartHandler&&) = delete;
  PartHandler& operator=(PartHandler&&) = delete;
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_PART_HANDLER_H_
