//
// Address-stable owners for the payloads handed across the C ABI as
// borrowed views.
//
// `fm_cf_rule_t` and its siblings carry `const char*` and `const T*`
// fields the callee only borrows: every one of them must still point at
// live, unmoved storage when the call that reads them runs. Owning that
// storage in a `std::vector` is the recurring defect - a later
// `push_back` reallocates and silently invalidates a pointer already
// published into the record - so the storage lives here instead, where
// address stability is a property of the type rather than of the order
// in which the surrounding statements happen to be written.
//
// Neither type exposes its container and neither offers a capacity
// hint: there is no API through which a caller can reintroduce the
// reallocation.

#ifndef FORMULON_C_API_BORROWED_ARENA_H_
#define FORMULON_C_API_BORROWED_ARENA_H_

#include <cstddef>
#include <cstring>
#include <deque>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace formulon {
namespace c_api {

/// Byte written over released transient storage in builds with assertions.
inline constexpr unsigned char kReleasedByte = 0xDD;

/// Empties `live`, the storage behind borrowed views a caller may still hold. With assertions (no `NDEBUG`) the
/// released bytes are overwritten with `kReleasedByte` and kept, unreused, until the next release: a view read past
/// its documented lifetime then sees the marker every time instead of whatever the allocator reused the memory for.
/// `poison` overwrites one element of `live`.
template <typename Container, typename Poison>
void release_transient(Container& live, [[maybe_unused]] Container& retired, [[maybe_unused]] Poison poison) {
#ifndef NDEBUG
  for (auto& element : live) {
    poison(element);
  }
  retired = std::move(live);
#endif
  live.clear();
}

/// `release_transient` for a store of `std::string`s; the terminating NUL stays, so a stale view reads a marker
/// string of the original length.
inline void release_text(std::deque<std::string>& live, std::deque<std::string>& retired) {
  release_transient(live, retired, [](std::string& text) { std::memset(&text[0], kReleasedByte, text.size()); });
}

/// Owns the UTF-8 bytes behind borrowed `const char*` fields.
///
/// `std::deque` never relocates the elements it already holds, so every
/// pointer handed out by `emplace` stays valid until `clear` runs or the
/// arena dies - no matter how many strings are added afterwards.
class BorrowedStringArena {
 public:
  /// Copies `text` into the arena and returns a NUL-terminated view of
  /// the stored copy. Never returns null; an empty input yields a
  /// pointer to an empty string.
  const char* emplace(std::string_view text) {
    storage_.emplace_back(text);
    return storage_.back().c_str();
  }

  /// Drops every stored string, invalidating all previously returned
  /// pointers (see `release_text`).
  void clear() { release_text(storage_, retired_); }

  /// Number of strings currently retained.
  std::size_t size() const noexcept { return storage_.size(); }

  bool empty() const noexcept { return storage_.empty(); }

 private:
  std::deque<std::string> storage_;
  std::deque<std::string> retired_;
};

/// Owns the contiguous arrays behind borrowed `const T*` / count field
/// pairs.
///
/// Each adopted block keeps the heap allocation it arrived with, and
/// moving a `std::vector` transfers that allocation rather than copying
/// it. Adopting further blocks therefore never moves the elements of an
/// earlier one, even when the outer container grows.
template <typename T>
class BorrowedArrayArena {
 public:
  /// Takes ownership of `block` and returns a pointer to its elements,
  /// or null when the block is empty (the C ABI spells an absent array
  /// as a null pointer with a zero count).
  const T* adopt(std::vector<T>&& block) {
    if (block.empty()) {
      return nullptr;
    }
    storage_.push_back(std::move(block));
    return storage_.back().data();
  }

  /// Drops every stored block, invalidating all previously returned
  /// pointers (see `release_transient`).
  void clear() {
    release_transient(storage_, retired_, [](std::vector<T>& block) {
      std::memset(static_cast<void*>(block.data()), kReleasedByte, block.size() * sizeof(T));
    });
  }

  /// Total number of elements retained across all blocks.
  std::size_t size() const noexcept {
    std::size_t total = 0;
    for (const auto& block : storage_) {
      total += block.size();
    }
    return total;
  }

  bool empty() const noexcept { return storage_.empty(); }

 private:
  std::vector<std::vector<T>> storage_;
  std::vector<std::vector<T>> retired_;
};

}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_BORROWED_ARENA_H_
