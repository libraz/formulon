// Copyright 2026 libraz. Licensed under the MIT License.
//
// A simple bump-allocating arena for trivially-destructible payloads.
//
// The arena owns a singly-linked list of heap chunks. `allocate()` hands out
// aligned slices of the head chunk; when the head runs out of room we allocate
// a fresh chunk whose size doubles up to a 1 MiB cap (or grows to fit the
// current request, whichever is larger). No destructors are ever invoked:
// callers must only store trivially-destructible types, which is enforced by
// `static_assert` on the `create<T>()` entry points.
//
// Rationale:
//   - `-fno-exceptions` means we cannot rely on `operator new` to throw on
//     allocation failure. We use `std::malloc` so failure is a clean
//     `nullptr` which we propagate back to the caller.
//   - Trivially-destructible objects dominate AST / bytecode payloads; a bump
//     allocator removes per-node `delete` traffic and keeps parse/compile
//     working-set memory contiguous.
//
// The implementation is header-only and inline so the arena can be used from
// any translation unit without a separate object file.

#ifndef FORMULON_UTILS_ARENA_H_
#define FORMULON_UTILS_ARENA_H_

#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <new>
#include <string_view>
#include <type_traits>
#include <utility>

namespace formulon {

/// Bump allocator for trivially-destructible payloads.
///
/// `Arena` is not thread-safe: concurrent writers must synchronise externally.
/// Allocation failures surface as `nullptr`; the arena never terminates the
/// process on its own.
class Arena {
 public:
  /// Builds an empty arena. The first chunk is allocated lazily on the first
  /// `allocate()` request, sized to `initial_chunk_bytes`.
  explicit Arena(std::size_t initial_chunk_bytes = 4096) noexcept
      : initial_chunk_bytes_(initial_chunk_bytes == 0 ? 4096 : initial_chunk_bytes),
        next_chunk_bytes_(initial_chunk_bytes == 0 ? 4096 : initial_chunk_bytes) {}

  /// Releases every owned chunk.
  ~Arena() { release_all(); }

  Arena(const Arena&) = delete;
  Arena& operator=(const Arena&) = delete;

  /// Transfers ownership of every chunk from `other`, leaving it empty.
  Arena(Arena&& other) noexcept
      : head_(other.head_),
        initial_chunk_bytes_(other.initial_chunk_bytes_),
        next_chunk_bytes_(other.next_chunk_bytes_),
        chunk_count_(other.chunk_count_),
        bytes_allocated_(other.bytes_allocated_) {
    other.head_ = nullptr;
    other.chunk_count_ = 0;
    other.bytes_allocated_ = 0;
    other.next_chunk_bytes_ = other.initial_chunk_bytes_;
  }

  /// Transfers ownership of every chunk from `other`, releasing whatever this
  /// arena previously owned.
  Arena& operator=(Arena&& other) noexcept {
    if (this != &other) {
      release_all();
      head_ = other.head_;
      initial_chunk_bytes_ = other.initial_chunk_bytes_;
      next_chunk_bytes_ = other.next_chunk_bytes_;
      chunk_count_ = other.chunk_count_;
      bytes_allocated_ = other.bytes_allocated_;
      other.head_ = nullptr;
      other.chunk_count_ = 0;
      other.bytes_allocated_ = 0;
      other.next_chunk_bytes_ = other.initial_chunk_bytes_;
    }
    return *this;
  }

  /// Returns an aligned pointer to `bytes` of raw storage, or `nullptr` if the
  /// underlying chunk allocation failed. `alignment` must be a power of two.
  void* allocate(std::size_t bytes, std::size_t alignment) noexcept {
    if (bytes == 0) {
      bytes = 1;
    }
    if (alignment == 0 || (alignment & (alignment - 1)) != 0) {
      return nullptr;
    }

    if (head_ != nullptr) {
      void* p = try_allocate_in(head_, bytes, alignment);
      if (p != nullptr) {
        return p;
      }
    }

    // Head is either missing or too full. Size the new chunk so it can satisfy
    // the current request even after worst-case alignment padding.
    const std::size_t needed = bytes + alignment;
    std::size_t chunk_bytes = next_chunk_bytes_;
    if (chunk_bytes < needed) {
      chunk_bytes = needed;
    }
    Chunk* c = new_chunk(chunk_bytes);
    if (c == nullptr) {
      return nullptr;
    }
    c->next = head_;
    head_ = c;
    ++chunk_count_;
    bytes_allocated_ += c->capacity;

    // Advance the default growth policy: double up to kMaxChunkBytes.
    if (next_chunk_bytes_ < kMaxChunkBytes) {
      std::size_t doubled = next_chunk_bytes_ * 2;
      if (doubled < next_chunk_bytes_) {
        doubled = kMaxChunkBytes;  // overflow guard
      }
      next_chunk_bytes_ = doubled > kMaxChunkBytes ? kMaxChunkBytes : doubled;
    }

    return try_allocate_in(head_, bytes, alignment);
  }

  /// Allocates and constructs a single `T`, forwarding `args...`. `T` must be
  /// trivially destructible because the arena never calls destructors.
  template <class T, class... Args>
  T* create(Args&&... args) {
    static_assert(std::is_trivially_destructible_v<T>,
                  "Arena only stores trivially-destructible types (no destructors are called)");
    static_assert((alignof(T) & (alignof(T) - 1)) == 0, "alignof(T) must be a power of two");
    void* raw = allocate(sizeof(T), alignof(T));
    if (raw == nullptr) {
      return nullptr;
    }
    return ::new (raw) T(std::forward<Args>(args)...);
  }

  /// Allocates an uninitialised contiguous array of `n` `T`s. `T` must be
  /// trivially destructible. The returned pointer is nullptr if allocation
  /// fails or if `n * sizeof(T)` would overflow.
  template <class T>
  T* create_array(std::size_t n) {
    static_assert(std::is_trivially_destructible_v<T>,
                  "Arena only stores trivially-destructible types (no destructors are called)");
    static_assert((alignof(T) & (alignof(T) - 1)) == 0, "alignof(T) must be a power of two");
    if (n == 0) {
      return nullptr;
    }
    if (n > SIZE_MAX / sizeof(T)) {
      return nullptr;
    }
    void* raw = allocate(sizeof(T) * n, alignof(T));
    return static_cast<T*>(raw);
  }

  /// Copies `src` into arena memory and returns a view of the copy. The
  /// stored bytes are NUL-terminated so callers can treat the storage as a
  /// C string; the returned view's length excludes the terminator. Returns an
  /// empty view on allocation failure or empty input.
  std::string_view intern(std::string_view src) noexcept {
    if (src.empty()) {
      return {};
    }
    void* raw = allocate(src.size() + 1, alignof(char));
    if (raw == nullptr) {
      return {};
    }
    auto* dst = static_cast<char*>(raw);
    std::memcpy(dst, src.data(), src.size());
    dst[src.size()] = '\0';
    return std::string_view(dst, src.size());
  }

  /// Returns the total bytes handed out to callers across every chunk.
  std::size_t bytes_used() const noexcept {
    std::size_t sum = 0;
    for (const Chunk* c = head_; c != nullptr; c = c->next) {
      sum += c->used;
    }
    return sum;
  }

  /// Returns the total capacity of every owned chunk.
  std::size_t bytes_allocated() const noexcept { return bytes_allocated_; }

  /// Returns the number of chunks currently owned.
  std::size_t chunk_count() const noexcept { return chunk_count_; }

  /// Releases every chunk except the most-recently-added (which is reset to
  /// empty). If the arena currently owns no chunks this is a no-op. After
  /// `reset()` the arena may be reused; the retained chunk is typically the
  /// largest one, so subsequent allocations often avoid re-growing.
  void reset() noexcept {
    if (head_ == nullptr) {
      return;
    }
    Chunk* keep = head_;
    Chunk* older = keep->next;
    while (older != nullptr) {
      Chunk* next = older->next;
      bytes_allocated_ -= older->capacity;
      --chunk_count_;
      free_chunk(older);
      older = next;
    }
    keep->next = nullptr;
    keep->used = 0;
    head_ = keep;
    next_chunk_bytes_ = initial_chunk_bytes_;
  }

 private:
  struct Chunk {
    Chunk* next;
    std::size_t capacity;
    std::size_t used;
    // `data` is the payload; sized by the allocator. We over-allocate the
    // surrounding buffer (see `new_chunk`) so the usable bytes start here.
    alignas(std::max_align_t) unsigned char data[1];
  };

  static constexpr std::size_t kMaxChunkBytes = 1u << 20;  // 1 MiB

  static std::size_t align_up(std::size_t value, std::size_t alignment) noexcept {
    return (value + (alignment - 1)) & ~(alignment - 1);
  }

  static void* try_allocate_in(Chunk* c, std::size_t bytes, std::size_t alignment) noexcept {
    auto base = reinterpret_cast<std::uintptr_t>(c->data);
    const std::uintptr_t cursor = base + c->used;
    const std::uintptr_t aligned = (cursor + (alignment - 1)) & ~(static_cast<std::uintptr_t>(alignment) - 1);
    const std::size_t pad = static_cast<std::size_t>(aligned - cursor);
    if (c->used + pad + bytes > c->capacity) {
      return nullptr;
    }
    c->used += pad + bytes;
    return reinterpret_cast<void*>(aligned);
  }

  static Chunk* new_chunk(std::size_t capacity) noexcept {
    // Place the Chunk header in the same allocation as its data; the `data`
    // array is just a placeholder that marks where the payload begins.
    const std::size_t header_bytes = offsetof(Chunk, data);
    const std::size_t total = header_bytes + capacity;
    void* raw = std::malloc(total);
    if (raw == nullptr) {
      return nullptr;
    }
    auto* c = static_cast<Chunk*>(raw);
    c->next = nullptr;
    c->capacity = capacity;
    c->used = 0;
    return c;
  }

  static void free_chunk(Chunk* c) noexcept { std::free(c); }

  void release_all() noexcept {
    Chunk* c = head_;
    while (c != nullptr) {
      Chunk* next = c->next;
      free_chunk(c);
      c = next;
    }
    head_ = nullptr;
    chunk_count_ = 0;
    bytes_allocated_ = 0;
    next_chunk_bytes_ = initial_chunk_bytes_;
  }

  Chunk* head_ = nullptr;
  std::size_t initial_chunk_bytes_ = 4096;
  std::size_t next_chunk_bytes_ = 4096;
  std::size_t chunk_count_ = 0;
  std::size_t bytes_allocated_ = 0;
};

}  // namespace formulon

#endif  // FORMULON_UTILS_ARENA_H_
