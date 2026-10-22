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
/// document. The author table starts with `known_authors` (in order) and is
/// extended with the unique `author` strings of `comments` in
/// first-occurrence order; each `<comment>` references its author by index. Empty `comments` returns the empty
/// string — the writer must NOT emit an empty comments part. A non-empty
/// `uid` is written as the comment's `xr:uid` attribute.
std::string write_comments(const std::vector<CellComment>& comments,
                           const std::vector<std::string>& known_authors = {});

/// Serialises a VML drawing with one note shape per entry of `shapes`. A shape
/// with source `xml` is re-emitted as it was with only its anchor, visibility
/// and cell patched; any other is laid out in Excel's default box unless it
/// carries its own, and hidden unless it is marked visible. `drawing_id`
/// (1-based, unique per package) sets the shape-id block: `<o:idmap data>` is
/// `drawing_id` and new shapes are numbered from `1024 * drawing_id + 1`,
/// skipping the ids source shapes carry.
std::string write_vml_drawing(const std::vector<CommentShape>& shapes, std::uint32_t drawing_id);

/// A VML drawing with no shapes (`write_vml_drawing({}, 1)`).
std::string write_vml_drawing_stub();

}  // namespace formulon::io

#endif  // FORMULON_IO_COMMENTS_WRITER_H_
