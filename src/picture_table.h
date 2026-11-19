//
// The workbook's picture table: one entry per distinct (media, alt) pair a
// `Value::picture` names, and the byte order of the media. Append-only for
// the life of the workbook, so picture and media ids are never recycled and
// `entry()` pointers stay valid across additions. The table holds no bytes;
// the caller orders a new media and passes its position.

#ifndef FORMULON_PICTURE_TABLE_H_
#define FORMULON_PICTURE_TABLE_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <optional>
#include <string>
#include <vector>

namespace formulon {

/// Identity of a media part's bytes: their length and 64-bit FNV-1a digest.
struct MediaKey {
  std::uint64_t size;
  std::uint64_t digest;
};

/// One picture: the media it shows and its alt text (absent when none).
struct PictureEntry {
  std::string media_path;
  std::uint32_t media_id;
  std::optional<std::string> alt;
};

/// Returns -1, 0 or 1 as media bytes `a` sort before, equal to or after `b`
/// in ascending picture order: the sequence greater at the first differing
/// byte, compared unsigned, sorts first, and of a sequence and its proper
/// prefix the longer sorts first.
int compare_media_bytes(const std::uint8_t* a, std::size_t a_len, const std::uint8_t* b, std::size_t b_len) noexcept;

class PictureTable {
 public:
  /// Adds a media taking rank `order_position` in ascending picture order;
  /// the media at that rank and above move up by one. Returns its media id.
  /// Aborts if `order_position > media_count()`.
  std::uint32_t add_media_at(std::size_t order_position, MediaKey key, std::string media_path = {});

  /// Adds a picture of `media_id` with `alt` (already normalised) and
  /// returns its picture id. Aborts if `media_id` is unknown.
  std::uint32_t add_entry(std::uint32_t media_id, std::optional<std::string> alt);

  /// The entry of `picture_id`, or nullptr when the table does not hold it.
  const PictureEntry* entry(std::uint32_t picture_id) const noexcept;

  /// Rank of `media_id` in ascending picture order; 0 sorts first. Aborts if
  /// `media_id` is unknown.
  std::uint32_t media_rank(std::uint32_t media_id) const;

  /// Number of entries.
  std::size_t size() const noexcept { return entries_.size(); }

  /// Number of media.
  std::size_t media_count() const noexcept { return media_keys_.size(); }

 private:
  // @limit: internal — no bound below UINT32_MAX entries; the table grows by one per distinct (media, alt) placed or
  // loaded during the handle's life
  std::deque<PictureEntry> entries_;
  // Indexed by media id.
  std::vector<MediaKey> media_keys_;
  std::vector<std::string> media_paths_;
  std::vector<std::uint32_t> media_ranks_;
};

}  // namespace formulon

#endif  // FORMULON_PICTURE_TABLE_H_
