//
// Unit tests for `eval::format_lambda_value`. Builds `LambdaValue`
// instances directly (without driving the parser / evaluator) so the
// rendering contract can be pinned independently of the surrounding
// pipeline. End-to-end coverage that exercises this through the C ABI
// boundary lives in `tests/c_api/formulon_c_lambda_text_test.cpp`.

#include "eval/lambda_format.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/lambda_value.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using parser::AstNode;
using parser::BinOp;
using parser::make_binary_op;
using parser::make_literal;
using parser::make_name_ref;

struct LambdaBuilder {
  Arena arena;
  std::vector<std::vector<std::string_view>> param_holders;

  LambdaValue* make(const std::vector<std::string_view>& params, std::uint32_t optional_count, AstNode* body) {
    param_holders.push_back(params);
    auto* lv = arena.create<LambdaValue>();
    lv->params = param_holders.back().data();
    lv->param_count = static_cast<std::uint32_t>(params.size());
    lv->optional_count = optional_count;
    lv->body = body;
    lv->captured_env = nullptr;
    return lv;
  }

  AstNode* num(double v) { return make_literal(arena, Value::number(v)); }
  AstNode* nameref(std::string_view name) { return make_name_ref(arena, name); }
  AstNode* mul(AstNode* a, AstNode* b) { return make_binary_op(arena, BinOp::Mul, a, b); }
  AstNode* add(AstNode* a, AstNode* b) { return make_binary_op(arena, BinOp::Add, a, b); }
};

TEST(LambdaFormat, SingleParamScaledByLiteral) {
  LambdaBuilder b;
  AstNode* body = b.mul(b.nameref("x"), b.num(2.0));
  const LambdaValue* lv = b.make({"x"}, /*optional_count=*/0, body);
  EXPECT_EQ(format_lambda_value(*lv), "LAMBDA(x,x*2)");
}

TEST(LambdaFormat, MultipleParamsCommaSeparated) {
  LambdaBuilder b;
  // `format_formula` may emit redundant parens around right-children to
  // preserve round-trip equivalence (the contract is parse-equivalent
  // text, not byte-stable). Asserting on the substring shape rather
  // than exact text keeps the test robust to defensive parenthesisation.
  AstNode* body = b.add(b.nameref("a"), b.add(b.nameref("b"), b.nameref("c")));
  const LambdaValue* lv = b.make({"a", "b", "c"}, /*optional_count=*/0, body);
  const std::string out = format_lambda_value(*lv);
  EXPECT_TRUE(out.find("LAMBDA(a,b,c,") == 0) << out;
  EXPECT_NE(out.find("a+"), std::string::npos) << out;
  EXPECT_NE(out.find("b+c"), std::string::npos) << out;
  EXPECT_TRUE(out.back() == ')') << out;
}

TEST(LambdaFormat, OptionalTrailingParamsRenderWithBrackets) {
  LambdaBuilder b;
  // Only `extra` is optional; `base` is required.
  AstNode* body = b.add(b.nameref("base"), b.nameref("extra"));
  const LambdaValue* lv = b.make({"base", "extra"}, /*optional_count=*/1, body);
  EXPECT_EQ(format_lambda_value(*lv), "LAMBDA(base,[extra],base+extra)");
}

TEST(LambdaFormat, AllParamsOptional) {
  LambdaBuilder b;
  AstNode* body = b.num(0.0);
  const LambdaValue* lv = b.make({"x", "y"}, /*optional_count=*/2, body);
  EXPECT_EQ(format_lambda_value(*lv), "LAMBDA([x],[y],0)");
}

TEST(LambdaFormat, ZeroParamsOmitsTrailingComma) {
  LambdaBuilder b;
  AstNode* body = b.num(42.0);
  const LambdaValue* lv = b.make({}, /*optional_count=*/0, body);
  EXPECT_EQ(format_lambda_value(*lv), "LAMBDA(42)");
}

TEST(LambdaFormat, NullBodyEmitsEmptyBody) {
  // Defensive: a malformed LambdaValue with `body == nullptr` (only
  // ever produced via the test fixture in `lambda_test.cpp`) still
  // renders without crashing.
  LambdaBuilder b;
  const LambdaValue* lv = b.make({"x"}, /*optional_count=*/0, /*body=*/nullptr);
  EXPECT_EQ(format_lambda_value(*lv), "LAMBDA(x,)");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
