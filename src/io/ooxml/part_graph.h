//
// Relationship graph over a workbook's passthrough parts.
//
// Parts the engine edits without modelling them (drawings, media, the theme)
// are linked by rels parts and by the sheet and workbook relationships. These
// functions read one rels part and drop parts that nothing references any
// more, so every edit that removes a link cleans up the same way.

#ifndef FORMULON_IO_OOXML_PART_GRAPH_H_
#define FORMULON_IO_OOXML_PART_GRAPH_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;
struct PassthroughPart;

namespace io {
namespace ooxml {

/// One relationship of a part's rels part. `target` is the resolved package
/// path for an internal target and the raw value for an external one.
struct DrawingRel {
  std::string id;
  std::string type;
  std::string target;
  bool external = false;
};

/// Parses the rels part of a part in directory `owner_dir` (no trailing slash). Fails with
/// `kIoDrawingUnparseable` on malformed XML or a target that cannot resolve.
Expected<std::vector<DrawingRel>, Error> parse_part_rels(const std::vector<std::uint8_t>& bytes,
                                                         std::string_view owner_dir);

/// True when a sheet relationship, a workbook relationship or any rels part
/// in `parts` targets `path`. An unreadable rels part counts as a reference.
bool referenced(const Workbook& wb, const std::vector<PassthroughPart>& parts, const std::string& path);

/// Drops each candidate part nothing references any more, with its rels part,
/// and then the targets only that rels part referenced.
void remove_orphans(Workbook& wb, std::vector<std::string> candidates);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_PART_GRAPH_H_
