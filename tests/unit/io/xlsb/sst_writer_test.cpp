//
// Unit tests for the XLSB shared-string-table builder + emitter.
// Symmetry with `read_xlsb`'s SST decoder is checked by re-walking the
// emitted bytes through `io/xlsb/record.h` (the writer's reader-side
// counterpart) — Bundle 4.1's SST decoder lives inside
// `read_xlsb` and isn't directly callable, so we hand-decode the
// records the same way the reader does.

#include "io/xlsb/sst_writer.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/text_ops.h"
#include "eval/utf8_length.h"
#include "gtest/gtest.h"
#include "io/xlsb/record.h"
#include "io/zip_reader.h"
#include "phonetic.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

const std::vector<PhoneticRun> kNoPhonetic;
constexpr PhoneticProperties kDefaultProps{};

ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return ByteSpan{v.data(), v.size()};
}

// Walks a serialised SST body the way `read_xlsb` does: skip records
// other than `BrtSSTItem`, decode each item as `(u8 flags, XLWideString)`.
// Returns the decoded items in stream order.
std::vector<std::string> DecodeSstStream(const std::vector<std::uint8_t>& body) {
  std::vector<std::string> out;
  ByteSpan cursor = SpanOf(body);
  while (cursor.size > 0) {
    auto rec = read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << "read_record failed: " << rec.error().message;
      return out;
    }
    if (rec.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem)) {
      continue;
    }
    ByteSpan p = rec.value().payload;
    auto flags = read_u8(p);
    if (!flags) {
      ADD_FAILURE() << "BrtSSTItem flags truncated";
      return out;
    }
    (void)flags.value();
    auto s = read_xlwidestring(p);
    if (!s) {
      ADD_FAILURE() << "BrtSSTItem string truncated: " << s.error().message;
      return out;
    }
    out.push_back(s.value());
  }
  return out;
}

/// One `BrtSSTItem` decoded far enough to see the phonetic tail.
struct DecodedSstItem {
  std::string text;
  std::vector<PhoneticRun> phonetic;
};

// Walks the emitted stream the way `read_xlsb` does, but decodes the
// phonetic tail as well: the kana concatenation, the run count, and each
// `(ichFirst, ichMom, cchMom)` triple, closed by `(ifnt, flags)`.
std::vector<DecodedSstItem> DecodeSstItems(const std::vector<std::uint8_t>& body) {
  std::vector<DecodedSstItem> out;
  ByteSpan cursor = SpanOf(body);
  while (cursor.size > 0) {
    auto rec = read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << "read_record failed: " << rec.error().message;
      return out;
    }
    if (rec.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem)) {
      continue;
    }
    ByteSpan p = rec.value().payload;
    auto flags = read_u8(p);
    auto text = read_xlwidestring(p);
    if (!flags || !text) {
      ADD_FAILURE() << "BrtSSTItem truncated";
      return out;
    }
    DecodedSstItem item;
    item.text = text.value();
    if ((flags.value() & 0x02U) != 0U) {
      auto kana = read_xlwidestring(p);
      auto count = read_u32(p);
      if (!kana || !count) {
        ADD_FAILURE() << "phonetic tail truncated";
        return out;
      }
      std::vector<std::array<std::uint16_t, 3>> runs;
      for (std::uint32_t i = 0; i < count.value(); ++i) {
        std::array<std::uint16_t, 3> fields{};
        for (std::uint16_t& field : fields) {
          auto value = read_u16(p);
          if (!value) {
            ADD_FAILURE() << "phonetic run truncated";
            return out;
          }
          field = value.value();
        }
        runs.push_back(fields);
      }
      const std::uint32_t kana_units = eval::utf16_units_in(kana.value());
      for (std::size_t i = 0; i < runs.size(); ++i) {
        const std::uint32_t end = (i + 1U < runs.size()) ? runs[i + 1U][0] : kana_units;
        item.phonetic.push_back(PhoneticRun{runs[i][1], static_cast<std::uint32_t>(runs[i][1] + runs[i][2]),
                                            eval::utf16_substring(kana.value(), runs[i][0], end - runs[i][0])});
      }
      auto ifnt = read_u16(p);
      auto tail_flags = read_u16(p);
      if (!ifnt || !tail_flags) {
        ADD_FAILURE() << "phonetic properties truncated";
        return out;
      }
      // Excel's own defaults for a guide that arrived without a
      // `<phoneticPr>`: the workbook font, halfwidthKatakana, noControl.
      EXPECT_EQ(ifnt.value(), 0U);
      EXPECT_EQ(tail_flags.value(), 0x0030U);
    }
    out.push_back(std::move(item));
  }
  return out;
}

TEST(XlsbSstBuilder, EmptyBuilderEmitsBeginEndFraming) {
  SstBuilder sst;
  EXPECT_TRUE(sst.empty());
  EXPECT_EQ(sst.size(), 0U);
  auto body_or = emit_sst(sst);
  ASSERT_TRUE(static_cast<bool>(body_or));

  // Just begin + end records, no items.
  const std::vector<std::string> items = DecodeSstStream(body_or.value());
  EXPECT_TRUE(items.empty());
}

TEST(XlsbSstBuilder, InternsIdenticalStringsToSameIndex) {
  SstBuilder sst;
  EXPECT_EQ(sst.intern("apple", kNoPhonetic, kDefaultProps), 0U);
  EXPECT_EQ(sst.intern("banana", kNoPhonetic, kDefaultProps), 1U);
  EXPECT_EQ(sst.intern("apple", kNoPhonetic, kDefaultProps), 0U);
  EXPECT_EQ(sst.intern("cherry", kNoPhonetic, kDefaultProps), 2U);
  EXPECT_EQ(sst.intern("banana", kNoPhonetic, kDefaultProps), 1U);
  EXPECT_EQ(sst.size(), 3U);

  ASSERT_EQ(sst.entries().size(), 3U);
  EXPECT_EQ(sst.entries()[0].text, "apple");
  EXPECT_EQ(sst.entries()[1].text, "banana");
  EXPECT_EQ(sst.entries()[2].text, "cherry");
}

TEST(XlsbSstBuilder, EmittedStreamRoundTripsThroughReader) {
  SstBuilder sst;
  sst.intern("alpha", kNoPhonetic, kDefaultProps);
  sst.intern("beta", kNoPhonetic, kDefaultProps);
  sst.intern("alpha", kNoPhonetic, kDefaultProps);
  sst.intern("gamma", kNoPhonetic, kDefaultProps);

  auto body_or = emit_sst(sst);
  ASSERT_TRUE(static_cast<bool>(body_or));
  const std::vector<std::string> items = DecodeSstStream(body_or.value());
  ASSERT_EQ(items.size(), 3U);
  EXPECT_EQ(items[0], "alpha");
  EXPECT_EQ(items[1], "beta");
  EXPECT_EQ(items[2], "gamma");
}

TEST(XlsbSstBuilder, InternHandlesBmpAndSurrogatePairStrings) {
  SstBuilder sst;
  // BMP only ("日本") and a string that triggers surrogate pairs ("🌟ok")
  // to exercise the writer's UTF-16 expansion.
  EXPECT_EQ(sst.intern("\xE6\x97\xA5\xE6\x9C\xAC", kNoPhonetic, kDefaultProps), 0U);
  EXPECT_EQ(sst.intern("\xF0\x9F\x8C\x9F"
                       "ok",
                       kNoPhonetic, kDefaultProps),
            1U);

  auto body_or = emit_sst(sst);
  ASSERT_TRUE(static_cast<bool>(body_or));
  const std::vector<std::string> items = DecodeSstStream(body_or.value());
  ASSERT_EQ(items.size(), 2U);
  EXPECT_EQ(items[0], "\xE6\x97\xA5\xE6\x9C\xAC");
  EXPECT_EQ(items[1],
            "\xF0\x9F\x8C\x9F"
            "ok");
}

TEST(XlsbSstBuilder, PhoneticGuideKeepsItsSpansAndSplitsTheInternKey) {
  SstBuilder sst;
  const std::vector<PhoneticRun> tokyo{{0U, 2U, "トウキョウ"}, {2U, 3U, "ト"}};
  const std::vector<PhoneticRun> other{{0U, 3U, "ヒガシキョウト"}};
  // Same surface text, different readings: the guide is part of the entry,
  // so merging them would move one cell's furigana onto the other.
  EXPECT_EQ(sst.intern("東京都", tokyo, kDefaultProps), 0U);
  EXPECT_EQ(sst.intern("東京都", other, kDefaultProps), 1U);
  EXPECT_EQ(sst.intern("東京都", tokyo, kDefaultProps), 0U);
  EXPECT_EQ(sst.intern("東京都", kNoPhonetic, kDefaultProps), 2U);
  EXPECT_EQ(sst.size(), 3U);

  auto body_or = emit_sst(sst);
  ASSERT_TRUE(static_cast<bool>(body_or));
  const std::vector<DecodedSstItem> items = DecodeSstItems(body_or.value());
  ASSERT_EQ(items.size(), 3U);
  EXPECT_EQ(items[0].text, "東京都");
  ASSERT_EQ(items[0].phonetic.size(), 2U);
  EXPECT_EQ(items[0].phonetic[0].sb, 0U);
  EXPECT_EQ(items[0].phonetic[0].eb, 2U);
  EXPECT_EQ(items[0].phonetic[0].text, "トウキョウ");
  EXPECT_EQ(items[0].phonetic[1].sb, 2U);
  EXPECT_EQ(items[0].phonetic[1].eb, 3U);
  EXPECT_EQ(items[0].phonetic[1].text, "ト");
  ASSERT_EQ(items[1].phonetic.size(), 1U);
  EXPECT_EQ(items[1].phonetic[0].text, "ヒガシキョウト");
  EXPECT_TRUE(items[2].phonetic.empty());
}

TEST(XlsbSstBuilder, BeginRecordCarriesCountFields) {
  SstBuilder sst;
  sst.intern("a", kNoPhonetic, kDefaultProps);
  sst.intern("b", kNoPhonetic, kDefaultProps);
  auto body_or = emit_sst(sst);
  ASSERT_TRUE(static_cast<bool>(body_or));
  const std::vector<std::uint8_t>& body = body_or.value();

  ByteSpan cursor = SpanOf(body);
  auto rec = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(rec));
  EXPECT_EQ(rec.value().type, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSst));
  ByteSpan p = rec.value().payload;
  auto total = read_u32(p);
  auto unique = read_u32(p);
  ASSERT_TRUE(static_cast<bool>(total));
  ASSERT_TRUE(static_cast<bool>(unique));
  EXPECT_EQ(total.value(), 2U);
  EXPECT_EQ(unique.value(), 2U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
