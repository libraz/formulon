#include "picture_table.h"

#include <limits>
#include <utility>

#include "utils/expected.h"

namespace formulon {

int compare_media_bytes(const std::uint8_t* a, std::size_t a_len, const std::uint8_t* b, std::size_t b_len) noexcept {
  const std::size_t n = a_len < b_len ? a_len : b_len;
  for (std::size_t i = 0; i < n; ++i) {
    if (a[i] != b[i]) {
      return a[i] > b[i] ? -1 : 1;
    }
  }
  if (a_len == b_len) {
    return 0;
  }
  return a_len > b_len ? -1 : 1;
}

std::uint32_t PictureTable::add_media_at(std::size_t order_position, MediaKey key, std::string media_path) {
  FM_CHECK(order_position <= media_keys_.size(), "PictureTable::add_media_at: position past the ranked media");
  FM_CHECK(media_keys_.size() < std::numeric_limits<std::uint32_t>::max(), "PictureTable: media id space exhausted");
  const auto rank = static_cast<std::uint32_t>(order_position);
  for (std::uint32_t& r : media_ranks_) {
    if (r >= rank) {
      ++r;
    }
  }
  const auto id = static_cast<std::uint32_t>(media_keys_.size());
  media_keys_.push_back(key);
  media_paths_.push_back(std::move(media_path));
  media_ranks_.push_back(rank);
  return id;
}

std::uint32_t PictureTable::add_entry(std::uint32_t media_id, std::optional<std::string> alt) {
  FM_CHECK(media_id < media_keys_.size(), "PictureTable::add_entry: unknown media id");
  FM_CHECK(entries_.size() < std::numeric_limits<std::uint32_t>::max(), "PictureTable: picture id space exhausted");
  const auto id = static_cast<std::uint32_t>(entries_.size());
  entries_.push_back(PictureEntry{media_paths_[media_id], media_id, std::move(alt)});
  return id;
}

const PictureEntry* PictureTable::entry(std::uint32_t picture_id) const noexcept {
  return picture_id < entries_.size() ? &entries_[picture_id] : nullptr;
}

std::uint32_t PictureTable::media_rank(std::uint32_t media_id) const {
  FM_CHECK(media_id < media_ranks_.size(), "PictureTable::media_rank: unknown media id");
  return media_ranks_[media_id];
}

}  // namespace formulon
