// A malformed UTF-8 byte counts as one UTF-16 unit in token positions.

#include "gtest/gtest.h"
#include "parser/lexer_error.h"
#include "parser/token.h"
#include "parser/tokenizer.h"

namespace formulon {
namespace parser {
namespace {

TEST(TokenizerMalformedUtf8, MalformedByteAdvancesOneUtf16Unit) {
  Tokenizer tz(
      "A1+\xFF"
      "+B2");
  const auto& v = tz.tokens();
  ASSERT_EQ(v.size(), 5u);
  EXPECT_EQ(v[2].kind, TokenKind::Plus);
  EXPECT_EQ(v[2].range.start, 4u);
  EXPECT_EQ(v[3].kind, TokenKind::CellRef);
  EXPECT_EQ(v[3].range.start, 5u);
  EXPECT_EQ(v[3].range.end, 7u);
}

}  // namespace
}  // namespace parser
}  // namespace formulon
