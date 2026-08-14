//
// Unit tests for the function dispatch table.

#include "eval/function_registry.h"

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/spill_potential.h"
#include "gtest/gtest.h"
#include "parser/parser.h"
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

bool FormulaMaySpill(std::string_view formula) {
  Arena arena;
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  parser::AstNode* root = parser::parse_strict(formula, arena);
  return root != nullptr && may_produce_spill(*root);
}

SpillPotential FormulaPotential(std::string_view formula, const FunctionRegistry* registry = nullptr) {
  Arena arena;
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  parser::AstNode* root = parser::parse_strict(formula, arena);
  if (root == nullptr) {
    return SpillPotential::kMaySpill;
  }
  return registry == nullptr ? spill_potential(*root) : spill_potential(*root, *registry);
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

TEST(FunctionRegistry, LookupRejectsUnknownNameLongerThanEveryRegisteredName) {
  FunctionRegistry r;
  ASSERT_TRUE(r.register_function(FunctionDef{"FOO", 0u, kVariadic, &StubImpl}));
  const std::string too_long(4096, 'x');
  EXPECT_EQ(r.lookup(too_long), nullptr);
}

TEST(FunctionRegistry, LookupStillSupportsARegisteredLongName) {
  FunctionRegistry r;
  const std::string name(256, 'x');
  ASSERT_TRUE(r.register_function(FunctionDef{name, 0u, kVariadic, &StubImpl}));
  EXPECT_NE(r.lookup(name), nullptr);
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
  ASSERT_NE(r.lookup("SEQUENCE"), nullptr);
  EXPECT_EQ(r.lookup("SEQUENCE")->result_shape, FunctionDef::ResultShape::kArray);
  ASSERT_NE(r.lookup("ABS"), nullptr);
  EXPECT_EQ(r.lookup("ABS")->result_shape, FunctionDef::ResultShape::kBroadcast);
  ASSERT_NE(r.lookup("FILTERXML"), nullptr);
  EXPECT_EQ(r.lookup("FILTERXML")->result_shape, FunctionDef::ResultShape::kArray);
}

TEST(FunctionRegistry, CustomFunctionDefaultsToArrayCapableShape) {
  FunctionRegistry r;
  ASSERT_TRUE(r.register_function(FunctionDef{"CUSTOMSPILL", 0u, kVariadic, &StubImpl}));
  ASSERT_NE(r.lookup("CUSTOMSPILL"), nullptr);
  EXPECT_EQ(r.lookup("CUSTOMSPILL")->result_shape, FunctionDef::ResultShape::kArray);
}

TEST(FunctionRegistry, SpillPotentialHonoursScalarizationAndBroadcast) {
  EXPECT_FALSE(FormulaMaySpill("=SUM(B:B)"));
  EXPECT_FALSE(FormulaMaySpill("=NOW()"));
  EXPECT_FALSE(FormulaMaySpill("=@SEQUENCE(1,2)"));
  EXPECT_TRUE(FormulaMaySpill("=ABS(SEQUENCE(1,2))"));
  EXPECT_TRUE(FormulaMaySpill("=SIN(SEQUENCE(1,2))"));
  EXPECT_TRUE(FormulaMaySpill("=IF(TRUE,SEQUENCE(1,2),0)"));
  EXPECT_TRUE(FormulaMaySpill("=LET(x,SEQUENCE(1,2),x)"));
  EXPECT_TRUE(FormulaMaySpill("=MY_DEFINED_NAME"));
  EXPECT_TRUE(FormulaMaySpill("=CUSTOMSPILL(1)"));
}

TEST(FunctionRegistry, SpillPotentialCoversLazyReturnShapes) {
  EXPECT_FALSE(FormulaMaySpill("=DATE(2020,1,1)"));
  EXPECT_TRUE(FormulaMaySpill("=DATE(SEQUENCE(1,2),1,1)"));
  EXPECT_FALSE(FormulaMaySpill("=REGEXTEST(\"a\",\"a\")"));
  EXPECT_FALSE(FormulaMaySpill("=REGEXREPLACE(\"a\",\"a\",\"b\")"));
  EXPECT_TRUE(FormulaMaySpill("=REGEXREPLACE(SEQUENCE(1,2),\"1\",\"x\")"));
  EXPECT_TRUE(FormulaMaySpill("=REGEXEXTRACT(\"a\",\"a\")"));
  EXPECT_TRUE(FormulaMaySpill("=LINEST(A:A,B:B)"));
  EXPECT_TRUE(FormulaMaySpill("=TRIMRANGE(A1:B2)"));
  EXPECT_TRUE(FormulaMaySpill("=XLOOKUP(1,A:A,B:C)"));
  EXPECT_TRUE(FormulaMaySpill("=VLOOKUP(1,A:A,1)"));
  EXPECT_TRUE(FormulaMaySpill("=HLOOKUP(1,1:1,1)"));
  EXPECT_TRUE(FormulaMaySpill("=MATCH(1,A:A)"));
  EXPECT_FALSE(FormulaMaySpill("=CHOOSE(1,10,20)"));
  EXPECT_TRUE(FormulaMaySpill("=CHOOSE(SEQUENCE(1),10,20)"));
  EXPECT_TRUE(FormulaMaySpill("=INDEX(A1:A2,1)"));
  EXPECT_FALSE(FormulaMaySpill("=LOOKUP(1,A:A,B:B)"));
  EXPECT_FALSE(FormulaMaySpill("=ROW(A1)"));
  EXPECT_TRUE(FormulaMaySpill("=ROW(A1:A2)"));
  EXPECT_TRUE(FormulaMaySpill("=CELL(\"width\",A1)"));
  EXPECT_TRUE(FormulaMaySpill("=TEXT(SEQUENCE(1,2),\"0\")"));
  EXPECT_TRUE(FormulaMaySpill("=FILTERXML(\"<a><b>1</b><b>2</b></a>\",\"//b\")"));
  EXPECT_TRUE(FormulaMaySpill("=REDUCE(0,A1:B2,LAMBDA(a,b,a+b))"));
  EXPECT_TRUE(FormulaMaySpill("=MODE.MULT(1,2,1,2)"));
}

TEST(FunctionRegistry, SpillPotentialLetUsesOnlyTheBodyShape) {
  EXPECT_FALSE(FormulaMaySpill("=LET(x,SEQUENCE(1,2),42)"));
  EXPECT_TRUE(FormulaMaySpill("=LET(x,SEQUENCE(1,2),x)"));
}

TEST(FunctionRegistry, SpillPotentialRechecksKnownNamesAgainstRuntimeRegistry) {
  EXPECT_EQ(FormulaPotential("=ABS(1)"), SpillPotential::kNeedsRegistry);
  EXPECT_EQ(FormulaPotential("=SEQUENCE(1,2)"), SpillPotential::kNeedsRegistry);

  FunctionRegistry custom;
  ASSERT_TRUE(custom.register_function(FunctionDef{"ABS", 1U, 1U, &StubImpl, true, false, false, false, false,
                                                   FunctionDef::BlankScalarPolicy::Allow, ErrorCode::Value,
                                                   FunctionDef::ResultShape::kArray}));
  FunctionDef sum_def{"SUM", 1U, kVariadic, &StubImpl};
  sum_def.accepts_ranges = true;
  sum_def.result_shape = FunctionDef::ResultShape::kReduce;
  ASSERT_TRUE(custom.register_function(sum_def));
  for (const std::string_view name : {"SIN", "LEN"}) {
    FunctionDef array_def{name, 1U, 1U, &StubImpl};
    array_def.result_shape = FunctionDef::ResultShape::kArray;
    ASSERT_TRUE(custom.register_function(array_def));
  }
  EXPECT_EQ(FormulaPotential("=ABS(1)", &custom), SpillPotential::kMaySpill);
  EXPECT_EQ(FormulaPotential("=SUM(A:A)", &custom), SpillPotential::kNever);
  EXPECT_EQ(FormulaPotential("=SIN(1)", &custom), SpillPotential::kMaySpill);
  EXPECT_EQ(FormulaPotential("=LEN(1)", &custom), SpillPotential::kMaySpill);

  FunctionRegistry scalar_override;
  FunctionDef sequence_def{"SEQUENCE", 1U, kVariadic, &StubImpl};
  sequence_def.result_shape = FunctionDef::ResultShape::kScalar;
  ASSERT_TRUE(scalar_override.register_function(sequence_def));
  EXPECT_EQ(FormulaPotential("=SEQUENCE(1,2)", &scalar_override), SpillPotential::kNever);
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
