#include "picture_table.h"

#include <limits>
#include <string_view>
#include <utility>

#include "utils/expected.h"
#include "utils/fnv1a.h"
#include "utils/strings.h"

namespace formulon {
namespace {

// The `entry_ids_` key of (media id, alt): the id's four bytes, then 0 for no alt or 1 and the alt.
std::string entry_key(std::uint32_t media_id, std::optional<std::string_view> alt) {
  std::string key(4U, '\0');
  for (std::size_t i = 0; i < 4U; ++i) {
    key[i] = static_cast<char>((media_id >> (8U * i)) & 0xFFU);
  }
  key.push_back(alt.has_value() ? '\1' : '\0');
  if (alt.has_value()) {
    key.append(alt->data(), alt->size());
  }
  return key;
}

}  // namespace

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

MediaKey media_key_of(const std::uint8_t* data, std::size_t len) noexcept {
  return MediaKey{static_cast<std::uint64_t>(len), fnv1a64(std::string_view(reinterpret_cast<const char*>(data), len))};
}

int compare_media_keys(MediaKey a, MediaKey b) noexcept {
  if (a.size != b.size) {
    return a.size > b.size ? -1 : 1;
  }
  if (a.digest != b.digest) {
    return a.digest > b.digest ? -1 : 1;
  }
  return 0;
}

std::optional<std::string_view> normalise_picture_alt(std::optional<std::string_view> alt) noexcept {
  if (!alt.has_value()) {
    return std::nullopt;
  }
  const std::string_view text = strings::rtrim_spaces(*alt);
  return text.empty() ? std::nullopt : std::optional<std::string_view>(text);
}

std::uint32_t PictureTable::add_media_at(std::size_t order_position, MediaKey key, std::string media_path) {
  FM_CHECK(order_position <= media_keys_.size(), "PictureTable::add_media_at: position past the ranked media");
  FM_CHECK(media_keys_.size() < std::numeric_limits<std::uint32_t>::max(), "PictureTable: media id space exhausted");
  const auto id = static_cast<std::uint32_t>(media_keys_.size());
  media_keys_.push_back(key);
  media_paths_.push_back(std::move(media_path));
  media_ranks_.push_back(0U);
  rank_order_.insert(rank_order_.begin() + static_cast<std::ptrdiff_t>(order_position), id);
  for (std::size_t r = order_position; r < rank_order_.size(); ++r) {
    media_ranks_[rank_order_[r]] = static_cast<std::uint32_t>(r);
  }
  media_by_digest_.emplace(key.digest, id);
  return id;
}

std::uint32_t PictureTable::add_entry(std::uint32_t media_id, std::optional<std::string> alt) {
  FM_CHECK(media_id < media_keys_.size(), "PictureTable::add_entry: unknown media id");
  FM_CHECK(entries_.size() < std::numeric_limits<std::uint32_t>::max(), "PictureTable: picture id space exhausted");
  const auto id = static_cast<std::uint32_t>(entries_.size());
  entry_ids_.emplace(entry_key(media_id, alt.has_value() ? std::optional<std::string_view>(*alt) : std::nullopt), id);
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

std::optional<std::uint32_t> PictureTable::find_entry(std::uint32_t media_id,
                                                      std::optional<std::string_view> alt) const {
  const auto found = entry_ids_.find(entry_key(media_id, alt));
  return found == entry_ids_.end() ? std::nullopt : std::optional<std::uint32_t>(found->second);
}

std::vector<std::uint32_t> PictureTable::media_with_key(MediaKey key) const {
  std::vector<std::uint32_t> out;
  const auto range = media_by_digest_.equal_range(key.digest);
  for (auto it = range.first; it != range.second; ++it) {
    if (media_keys_[it->second].size == key.size) {
      out.push_back(it->second);
    }
  }
  return out;
}

std::uint32_t PictureTable::media_at_rank(std::size_t rank) const {
  FM_CHECK(rank < rank_order_.size(), "PictureTable::media_at_rank: rank past the ranked media");
  return rank_order_[rank];
}

MediaKey PictureTable::media_key(std::uint32_t media_id) const {
  FM_CHECK(media_id < media_keys_.size(), "PictureTable::media_key: unknown media id");
  return media_keys_[media_id];
}

const std::string& PictureTable::media_path(std::uint32_t media_id) const {
  FM_CHECK(media_id < media_paths_.size(), "PictureTable::media_path: unknown media id");
  return media_paths_[media_id];
}

}  // namespace formulon
