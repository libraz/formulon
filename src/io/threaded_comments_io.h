//
// Read and write the threadedComments and persons package parts.
//
// `xl/threadedComments/threadedComment<N>.xml` holds one sheet's threaded
// comments; `xl/persons/person.xml` the workbook's person list. Both are
// fully model-owned: the reader takes them out of passthrough and the writer
// regenerates them. Each thread also has a legacy note stub in the sheet's
// comments part, identified by the author `tc={thread id}`; the helpers here
// separate stubs from real notes on read and rebuild them on write.

#ifndef FORMULON_IO_THREADED_COMMENTS_IO_H_
#define FORMULON_IO_THREADED_COMMENTS_IO_H_

#include <cstdint>
#include <string>
#include <vector>

#include "sheet.h"
#include "threaded_comment.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {

/// Parses a threadedComments part. Fails with `kIoSheetCorrupt` on a missing
/// root or an absent / unparseable `ref`.
Expected<std::vector<ThreadedComment>, Error> read_threaded_comments(const std::vector<std::uint8_t>& bytes);

/// Parses a persons part. Fails with `kIoSheetCorrupt` on a missing root.
Expected<std::vector<Person>, Error> read_persons(const std::vector<std::uint8_t>& bytes);

/// Serialises one sheet's threaded comments; empty input returns "".
std::string write_threaded_comments(const std::vector<ThreadedComment>& comments);

/// Serialises the person list; empty input returns "".
std::string write_persons(const std::vector<Person>& persons);

/// Removes from `notes` every legacy stub of a thread in `threads`: a note
/// whose author is `tc=` + the id of a thread anchored at the same cell.
/// A `tc=` note without its thread stays a note.
void drop_thread_stubs(std::vector<CellComment>& notes, const std::vector<ThreadedComment>& threads);

/// The comments-part list for `sheet`: its notes in order, then one stub per
/// thread (author `tc={id}`, `xr:uid` = thread id). A note sharing a cell with
/// a thread is left out, since a cell carries one comment.
std::vector<CellComment> legacy_comments_with_stubs(const Sheet& sheet);

}  // namespace formulon::io

#endif  // FORMULON_IO_THREADED_COMMENTS_IO_H_
