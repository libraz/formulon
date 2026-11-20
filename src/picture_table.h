//
// The workbook's picture table: one entry per distinct (media, alt) pair a
// `Value::picture` names, and the byte order of the media. Append-only for
// the life of the workbook, so picture and media ids are never recycled and
// `entry()` pointers stay valid across additions. The table holds no bytes;
// the caller orders a new media and passes its position. Two maps find an
// existing entry by (media, alt) and the media sharing a key.

#ifndef FORMULON_PICTURE_TABLE_H_
#define FORMULON_PICTURE_TABLE_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
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

/// The key of media bytes `data[0..len)`.
MediaKey media_key_of(const std::uint8_t* data, std::size_t len) noexcept;

/// Orders two media whose bytes are not available by their keys, as
/// `compare_media_bytes` would order bytes: the larger size first, then the
/// larger digest.
int compare_media_keys(MediaKey a, MediaKey b) noexcept;

/// `alt` with trailing U+0020 removed; nullopt when that leaves it empty.
std::optional<std::string_view> normalise_picture_alt(std::optional<std::string_view> alt) noexcept;

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

  /// The first picture id added for `media_id` with exactly `alt`, or nullopt.
  std::optional<std::uint32_t> find_entry(std::uint32_t media_id, std::optional<std::string_view> alt) const;

  /// The media ids whose key equals `key`.
  std::vector<std::uint32_t> media_with_key(MediaKey key) const;

  /// The media id at rank `rank`. Aborts if `rank >= media_count()`.
  std::uint32_t media_at_rank(std::size_t rank) const;

  /// Key and part path of `media_id`. Abort if `media_id` is unknown.
  MediaKey media_key(std::uint32_t media_id) const;
  const std::string& media_path(std::uint32_t media_id) const;

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
  // Indexed by rank.
  std::vector<std::uint32_t> rank_order_;
  // (media id, alt) as `entry_key` spells it -> first picture id.
  std::unordered_map<std::string, std::uint32_t> entry_ids_;
  // Digest -> media ids holding it; the size is checked on lookup.
  std::unordered_multimap<std::uint64_t, std::uint32_t> media_by_digest_;
};

}  // namespace formulon

#endif  // FORMULON_PICTURE_TABLE_H_
