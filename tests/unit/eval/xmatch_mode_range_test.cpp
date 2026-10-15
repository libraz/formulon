// XMATCH / XLOOKUP reject match_mode and search_mode values far outside the
// int range the same way they reject any other unsupported mode.

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {
using formulon::test::EvalSource;

void ExpectValueError(const char* formula) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_error()) << formula;
  EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
}

TEST(XmatchModeRange, HugeMatchModeIsValueError) {
  ExpectValueError("=XMATCH(1,{1,2,3},1E300)");
  ExpectValueError("=XMATCH(1,{1,2,3},-1E300)");
}

TEST(XmatchModeRange, HugeSearchModeIsValueError) {
  ExpectValueError("=XMATCH(1,{1,2,3},0,1E300)");
  ExpectValueError("=XMATCH(1,{1,2,3},0,-1E300)");
}

TEST(XlookupModeRange, HugeModesAreValueError) {
  ExpectValueError("=XLOOKUP(1,{1,2,3},{4,5,6},0,1E300)");
  ExpectValueError("=XLOOKUP(1,{1,2,3},{4,5,6},0,-1E300)");
  ExpectValueError("=XLOOKUP(1,{1,2,3},{4,5,6},0,0,1E300)");
  ExpectValueError("=XLOOKUP(1,{1,2,3},{4,5,6},0,0,-1E300)");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
