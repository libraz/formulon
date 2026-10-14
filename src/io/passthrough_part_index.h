// Transient path lookup for an immutable passthrough-part collection.
//
// Writers and chart refresh build this index for one operation, then keep
// iterating their original vectors for output order. The map stores string
// views into those parts and therefore never copies part paths; callers must
// keep the collection unchanged until the index is discarded.

#ifndef FORMULON_IO_PASSTHROUGH_PART_INDEX_H_
#define FORMULON_IO_PASSTHROUGH_PART_INDEX_H_

#include <cstddef>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "passthrough_part.h"

namespace formulon {
namespace io {

/// Finds the first passthrough part for each package path in source order;
/// of duplicate paths, the first one wins.
class PassthroughPartIndex {
 public:
  explicit PassthroughPartIndex(const std::vector<PassthroughPart>& parts) {
    first_by_path_.reserve(parts.size());
    for (const PassthroughPart& part : parts) {
      first_by_path_.emplace(std::string_view(part.path), &part);
    }
  }

  /// Returns the first source-order part whose path equals `path`, or null.
  const PassthroughPart* find(std::string_view path) const noexcept {
    const auto it = first_by_path_.find(path);
    return it == first_by_path_.end() ? nullptr : it->second;
  }

 private:
  std::unordered_map<std::string_view, const PassthroughPart*> first_by_path_;
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_PASSTHROUGH_PART_INDEX_H_
