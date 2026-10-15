// PtgStr and the shared UTF-16LE decoder join surrogate pairs into 4-byte UTF-8.

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "parser/ast_format.h"
#include "ptg_codec_test_helpers.h"
#include "utils/arena.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr const char* kEmoji = "\xF0\x9F\x98\x80";  // U+1F600

TEST(XlsbPtgStringSurrogate, PairInRgceDecodesToFourByteUtf8) {
  // PtgStr, cch=2, D83D DE00.
  const std::vector<std::uint8_t> rgce = {0x17, 0x02, 0x00, 0x3D, 0xD8, 0x00, 0xDE};
  Arena arena;
  auto decoded = decode_ptgs(ByteSpan{rgce.data(), rgce.size()}, ByteSpan{}, arena, {"Sheet1"}, {}, {}, {});
  ASSERT_TRUE(static_cast<bool>(decoded));
  EXPECT_EQ(parser::format_formula(*decoded.value()), std::string("\"") + kEmoji + "\"");
}

TEST(XlsbPtgStringSurrogate, EmojiLiteralRoundTrips) {
  const std::string formula = std::string("\"") + kEmoji + "\"";
  EXPECT_EQ(ptg_codec_test_support::RoundTrip(formula), formula);
}

TEST(XlsbPtgStringSurrogate, SharedDecoderJoinsPairsAndKeepsLoneSurrogates) {
  const std::uint8_t pair[] = {0x3D, 0xD8, 0x00, 0xDE};
  std::string out;
  append_utf16le_as_utf8(pair, 2U, &out);
  EXPECT_EQ(out, kEmoji);

  const std::uint8_t lone[] = {0x3D, 0xD8, 0x41, 0x00};  // high surrogate then 'A'
  out.clear();
  append_utf16le_as_utf8(lone, 2U, &out);
  EXPECT_EQ(out,
            "\xED\xA0\xBD"
            "A");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
