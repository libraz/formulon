// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for utils/arena.h.
//
// Note: a compile-time guard ensures `Arena::create<T>` rejects non-trivially
// destructible `T` (for example `std::string`). That is verified via
// `static_assert` inside the header; we do not exercise compile-failure cases
// from the test binary.

#include "utils/arena.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>

#include "gtest/gtest.h"

namespace formulon {
namespace {

struct Pod {
  int a;
  double b;
  char c;
};

TEST(ArenaTest, AllocateAlignedRaw) {
  Arena arena(128);
  void* p8 = arena.allocate(16, 8);
  ASSERT_NE(nullptr, p8);
  EXPECT_EQ(0u, reinterpret_cast<std::uintptr_t>(p8) % 8);

  void* p16 = arena.allocate(32, 16);
  ASSERT_NE(nullptr, p16);
  EXPECT_EQ(0u, reinterpret_cast<std::uintptr_t>(p16) % 16);

  void* p64 = arena.allocate(8, 64);
  ASSERT_NE(nullptr, p64);
  EXPECT_EQ(0u, reinterpret_cast<std::uintptr_t>(p64) % 64);
}

TEST(ArenaTest, AllocateRejectsNonPowerOfTwoAlignment) {
  Arena arena(128);
  EXPECT_EQ(nullptr, arena.allocate(8, 3));
  EXPECT_EQ(nullptr, arena.allocate(8, 0));
}

TEST(ArenaTest, CreateDefaultConstructsPod) {
  Arena arena(256);
  Pod* p = arena.create<Pod>();
  ASSERT_NE(nullptr, p);
  p->a = 42;
  p->b = 3.14;
  p->c = 'z';
  EXPECT_EQ(42, p->a);
  EXPECT_DOUBLE_EQ(3.14, p->b);
  EXPECT_EQ('z', p->c);
  EXPECT_EQ(0u, reinterpret_cast<std::uintptr_t>(p) % alignof(Pod));
}

TEST(ArenaTest, CreateForwardsConstructorArgs) {
  struct Point {
    int x;
    int y;
  };
  Arena arena(64);
  Point* q = arena.create<Point>(Point{7, 11});
  ASSERT_NE(nullptr, q);
  EXPECT_EQ(7, q->x);
  EXPECT_EQ(11, q->y);
}

TEST(ArenaTest, CreateArrayAllocatesContiguous) {
  Arena arena(1024);
  constexpr std::size_t kN = 32;
  int* xs = arena.create_array<int>(kN);
  ASSERT_NE(nullptr, xs);
  for (std::size_t i = 0; i < kN; ++i) {
    xs[i] = static_cast<int>(i);
  }
  for (std::size_t i = 0; i < kN; ++i) {
    EXPECT_EQ(static_cast<int>(i), xs[i]) << "index " << i;
  }
  // Contiguity: &xs[i] == xs + i by definition; verify the strides match.
  EXPECT_EQ(xs + 1, &xs[1]);
  EXPECT_EQ(xs + kN - 1, &xs[kN - 1]);
}

TEST(ArenaTest, CreateArrayZeroReturnsNull) {
  Arena arena(64);
  EXPECT_EQ(nullptr, arena.create_array<int>(0));
}

TEST(ArenaTest, InternStoresCopyWithDistinctAddress) {
  Arena arena(128);
  const std::string source = "formulon";
  std::string_view view = arena.intern(source);
  ASSERT_EQ(source.size(), view.size());
  EXPECT_EQ(source, std::string(view));
  // The arena copy must live in arena-owned memory, not the caller's string.
  EXPECT_NE(source.data(), view.data());
  // NUL terminator is present so callers can treat the storage as a C string.
  EXPECT_EQ('\0', view.data()[view.size()]);
}

TEST(ArenaTest, InternEmptyIsEmpty) {
  Arena arena(64);
  EXPECT_TRUE(arena.intern("").empty());
}

TEST(ArenaTest, GrowsToNewChunkWhenExhausted) {
  Arena arena(64);
  // First allocation opens the initial chunk.
  ASSERT_NE(nullptr, arena.allocate(32, 1));
  EXPECT_EQ(1u, arena.chunk_count());

  // Request more than the initial chunk can hold. This must allocate a new
  // chunk whose capacity covers the request plus alignment padding.
  void* big = arena.allocate(256, 8);
  ASSERT_NE(nullptr, big);
  EXPECT_GE(arena.chunk_count(), 2u);
  EXPECT_GE(arena.bytes_allocated(), 256u + 64u);
}

TEST(ArenaTest, BytesUsedAccumulatesAcrossAllocations) {
  Arena arena(256);
  arena.allocate(10, 1);
  arena.allocate(20, 1);
  EXPECT_GE(arena.bytes_used(), 30u);
  EXPECT_LE(arena.bytes_used(), arena.bytes_allocated());
}

TEST(ArenaTest, MoveTransfersOwnership) {
  Arena src(128);
  int* p = src.create<int>(99);
  ASSERT_NE(nullptr, p);
  EXPECT_EQ(1u, src.chunk_count());

  Arena dst(std::move(src));
  EXPECT_EQ(0u, src.chunk_count());      // NOLINT(bugprone-use-after-move) -- intentional
  EXPECT_EQ(0u, src.bytes_allocated());  // NOLINT(bugprone-use-after-move)
  EXPECT_EQ(1u, dst.chunk_count());
  EXPECT_EQ(99, *p);  // Pointer still valid; storage moved with the arena.
}

TEST(ArenaTest, MoveAssignReleasesPriorChunks) {
  Arena a(64);
  a.allocate(16, 1);
  Arena b(64);
  b.allocate(16, 1);
  const std::size_t b_before = b.bytes_allocated();
  a = std::move(b);
  EXPECT_EQ(b_before, a.bytes_allocated());
  EXPECT_EQ(0u, b.chunk_count());  // NOLINT(bugprone-use-after-move)
}

TEST(ArenaTest, ResetKeepsFirstChunkReusable) {
  Arena arena(64);
  // Force multiple chunks.
  ASSERT_NE(nullptr, arena.allocate(32, 1));
  ASSERT_NE(nullptr, arena.allocate(512, 1));
  ASSERT_GE(arena.chunk_count(), 2u);
  const std::size_t before = arena.bytes_allocated();
  ASSERT_GT(before, 0u);

  arena.reset();

  EXPECT_EQ(1u, arena.chunk_count());
  EXPECT_EQ(0u, arena.bytes_used());
  EXPECT_GT(arena.bytes_allocated(), 0u);
  EXPECT_LE(arena.bytes_allocated(), before);

  // The retained chunk is still usable for further allocations.
  void* p = arena.allocate(8, 1);
  EXPECT_NE(nullptr, p);
  EXPECT_GT(arena.bytes_used(), 0u);
}

TEST(ArenaTest, ResetOnEmptyArenaIsNoOp) {
  Arena arena(64);
  arena.reset();
  EXPECT_EQ(0u, arena.chunk_count());
  EXPECT_EQ(0u, arena.bytes_allocated());
  // Still usable afterwards.
  int* p = arena.create<int>(1);
  ASSERT_NE(nullptr, p);
  EXPECT_EQ(1, *p);
}

}  // namespace
}  // namespace formulon
