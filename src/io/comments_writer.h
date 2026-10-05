//
// Writer for `xl/comments<N>.xml` and its legacy VML drawing. Emits the
// modern shape with an inline author table plus a flat `<commentList>` of
// `<comment>` elements (one per `CellComment`). Rich-text fidelity is
// intentionally dropped: each comment is emitted as a single
// `<r><t>...</t></r>` run.
//
// The VML drawing carries one hidden note shape per commented cell, which
// is what Excel needs to show the comment indicator. The OOXML writer
// re-emits a sheet's source VML verbatim while its comment anchors are
// unchanged and regenerates it with `write_vml_drawing` otherwise.

#ifndef FORMULON_IO_COMMENTS_WRITER_H_
#define FORMULON_IO_COMMENTS_WRITER_H_

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "sheet.h"

namespace formulon::io {

/// Serialises a `CellComment` list into a complete `xl/comments<N>.xml`
/// document. The author table is built from the unique `author` strings
/// present in `comments` (in first-occurrence order); each `<comment>`
/// references its author by index. Empty `comments` returns the empty
/// string — the writer must NOT emit an empty comments part. A non-empty
/// `uid` is written as the comment's `xr:uid` attribute.
std::string write_comments(const std::vector<CellComment>& comments);

/// Serialises a VML drawing with one hidden note shape per `(row, col)` in
/// `anchors`. `drawing_id` (1-based, unique per package) sets the shape-id
/// block: `<o:idmap data>` is `drawing_id` and shapes are numbered from
/// `1024 * drawing_id + 1`.
std::string write_vml_drawing(const std::vector<std::pair<std::uint32_t, std::uint32_t>>& anchors,
                              std::uint32_t drawing_id);

/// A VML drawing with no shapes (`write_vml_drawing({}, 1)`).
std::string write_vml_drawing_stub();

}  // namespace formulon::io

#endif  // FORMULON_IO_COMMENTS_WRITER_H_
