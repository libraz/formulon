// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the function dispatch table.

#include "eval/function_registry.h"

#include <cstdint>
#include <string_view>

#include "gtest/gtest.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Trivial impl that returns a fixed sentinel so we can prove the dispatcher
// found the right entry.
Value StubImpl(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::number(7.0);
}

Value SecondImpl(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::number(99.0);
}

TEST(FunctionRegistry, RegisterAndExactCaseLookup) {
  FunctionRegistry r;
  ASSERT_TRUE(r.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  const FunctionDef* def = r.lookup("FOO");
  ASSERT_NE(def, nullptr);
  EXPECT_EQ(def->canonical_name, "FOO");
}

TEST(FunctionRegistry, LookupIsCaseInsensitiveLowercase) {
  FunctionRegistry r;
  ASSERT_TRUE(r.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  EXPECT_NE(r.lookup("foo"), nullptr);
}

TEST(FunctionRegistry, LookupIsCaseInsensitiveMixedCase) {
  FunctionRegistry r;
  ASSERT_TRUE(r.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  EXPECT_NE(r.lookup("FoO"), nullptr);
}

TEST(FunctionRegistry, LookupNonExistentReturnsNullptr) {
  FunctionRegistry r;
  ASSERT_TRUE(r.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  EXPECT_EQ(r.lookup("BAR"), nullptr);
}

TEST(FunctionRegistry, DuplicateRegistrationReturnsFalseAndPreservesFirst) {
  FunctionRegistry r;
  ASSERT_TRUE(r.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  EXPECT_FALSE(r.register_function(FunctionDef{"FOO", 1u, 1u, &SecondImpl}));
  const FunctionDef* def = r.lookup("FOO");
  ASSERT_NE(def, nullptr);
  Arena a;
  // Confirm the original implementation survives.
  const Value v = def->impl(nullptr, 0u, a);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
  EXPECT_EQ(def->min_arity, 0u);
}

TEST(FunctionRegistry, SizeReflectsCount) {
  FunctionRegistry r;
  EXPECT_EQ(r.size(), 0u);
  ASSERT_TRUE(r.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  EXPECT_EQ(r.size(), 1u);
  ASSERT_TRUE(r.register_function(FunctionDef{"BAR", 0u, kVariadic, &StubImpl}));
  EXPECT_EQ(r.size(), 2u);
  // Duplicate does not bump the count.
  EXPECT_FALSE(r.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  EXPECT_EQ(r.size(), 2u);
}

TEST(FunctionRegistry, MoveConstructionPreservesContents) {
  FunctionRegistry src;
  ASSERT_TRUE(src.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  FunctionRegistry dst(std::move(src));
  EXPECT_NE(dst.lookup("FOO"), nullptr);
  EXPECT_EQ(dst.size(), 1u);
}

TEST(FunctionRegistry, DefaultRegistryContainsExpectedBuiltins) {
  const FunctionRegistry& r = default_registry();
  EXPECT_NE(r.lookup("SUM"), nullptr);
  EXPECT_NE(r.lookup("CONCAT"), nullptr);
  EXPECT_NE(r.lookup("CONCATENATE"), nullptr);
  EXPECT_NE(r.lookup("LEN"), nullptr);
}

TEST(FunctionRegistry, DefaultRegistryDoesNotContainIf) {
  // IF is special-cased in the evaluator's call dispatcher; it must not
  // appear in the table.
  const FunctionRegistry& r = default_registry();
  EXPECT_EQ(r.lookup("IF"), nullptr);
}

TEST(FunctionRegistry, DefaultRegistryIsCaseInsensitive) {
  const FunctionRegistry& r = default_registry();
  EXPECT_NE(r.lookup("sum"), nullptr);
  EXPECT_NE(r.lookup("Concat"), nullptr);
}

TEST(FunctionRegistry, DefaultRegistryIsSingleton) {
  const FunctionRegistry& a = default_registry();
  const FunctionRegistry& b = default_registry();
  EXPECT_EQ(&a, &b);
}

TEST(FunctionRegistry, DefaultRegistryIsNonEmpty) {
  const FunctionRegistry& r = default_registry();
  EXPECT_GT(r.size(), 0u);
}

TEST(FunctionRegistry, DefaultAcceptsRangesFalse) {
  // A bare FunctionDef defaults to non-range-aware; only explicitly
  // opted-in entries should see range expansion in the dispatcher.
  FunctionRegistry r;
  ASSERT_TRUE(r.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  const FunctionDef* def = r.lookup("FOO");
  ASSERT_NE(def, nullptr);
  EXPECT_FALSE(def->accepts_ranges);
}

TEST(FunctionRegistry, AggregatorsOptIntoRangeExpansion) {
  // The five aggregators wired in `builtins.cpp` are the only built-ins
  // that accept RangeOp arguments.
  const FunctionRegistry& r = default_registry();
  for (const char* name : {"SUM", "AVERAGE", "MIN", "MAX", "PRODUCT"}) {
    const FunctionDef* def = r.lookup(name);
    ASSERT_NE(def, nullptr) << name;
    EXPECT_TRUE(def->accepts_ranges) << name;
  }
  // A representative non-aggregator stays scalar-only.
  const FunctionDef* len = r.lookup("LEN");
  ASSERT_NE(len, nullptr);
  EXPECT_FALSE(len->accepts_ranges);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
