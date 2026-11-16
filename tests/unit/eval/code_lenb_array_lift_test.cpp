//
// CODE and LENB lift element by element over an array argument, keeping each element's own result.

#include <string_view>

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace {

using formulon::test::EvalSource;

double Number(std::string_view src) {
  const Value v = EvalSource(src);
  EXPECT_TRUE(v.is_number()) << src;
  return v.is_number() ? v.as_number() : 0.0;
}

TEST(CodeLenbArrayLift, CodeLiftsToItsShape) {
  EXPECT_EQ(Number("=COLUMNS(CODE({\"A\",\"B\"}))"), 2.0);
  EXPECT_EQ(Number("=INDEX(CODE({\"A\",\"B\"}),2)"), 66.0);
  const Value v = EvalSource("=INDEX(CODE({\"A\",\"\"}),2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(CodeLenbArrayLift, LenbLiftsToItsShape) {
  EXPECT_EQ(Number("=ROWS(LENB({\"a\";\"bc\";\"def\"}))"), 3.0);
  EXPECT_EQ(Number("=INDEX(LENB({\"a\";\"bc\";\"def\"}),3,1)"), 3.0);
}

TEST(CodeLenbArrayLift, ScalarCallsStayScalar) {
  EXPECT_EQ(Number("=CODE(\"A\")"), 65.0);
  EXPECT_EQ(Number("=LENB(\"abc\")"), 3.0);
}

}  // namespace
}  // namespace formulon
