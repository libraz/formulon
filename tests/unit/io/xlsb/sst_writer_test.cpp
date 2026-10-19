//
// Unit tests for the XLSB shared-string-table builder, emitter, and decoder.
// Emitted records are checked through the production `DecodeSharedStringsBin`
// path as well as the low-level record framing helpers.

#include "io/xlsb/sst_writer.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sst_reader.h"
#include "io/zip_reader.h"
#include "phonetic.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/text_ops.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

const std::vector<PhoneticRun> kNoPhonetic;
constexpr PhoneticProperties kDefaultProps{};

ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return ByteSpan{v.data(), v.size()};
}

/// One `BrtSSTItem` decoded far enough to see the phonetic tail.
struct DecodedSstItem {
  std::string text;
  std::vector<PhoneticRun> phonetic;
  PhoneticProperties phonetic_props;
};

// Uses the same decoder as `read_xlsb`, retaining owned text for the
// non-owning views returned by `DecodeSharedStringsBin`.
std::vector<DecodedSstItem> DecodeSstItems(const std::vector<std::uint8_t>& body) {
  std::deque<std::string> text_storage;
  std::vector<std::vector<PhoneticRun>> phonetic;
  std::vector<PhoneticProperties> phonetic_props;
  auto entries = DecodeSharedStringsBin(body, text_storage, phonetic, phonetic_props);
  if (!entries) {
    ADD_FAILURE() << "DecodeSharedStringsBin failed: " << entries.error().message;
    return {};
  }

  std::vector<DecodedSstItem> out;
  out.reserve(entries.value().size());
  for (std::size_t i = 0; i < entries.value().size(); ++i) {
    out.push_back(DecodedSstItem{std::string(entries.value()[i]), std::move(phonetic[i]), phonetic_props[i]});
  }
  return out;
}

std::vector<std::string> DecodeSstStream(const std::vector<std::uint8_t>& body) {
  const std::vector<DecodedSstItem> items = DecodeSstItems(body);
  std::vector<std::string> out;
  out.reserve(items.size());
  for (const DecodedSstItem& item : items) {
    out.push_back(item.text);
  }
  return out;
}

void AppendPhoneticSliceCase(std::vector<std::uint8_t>& body, std::string_view kana, std::uint16_t start,
                             std::uint16_t next_start) {
  std::vector<std::uint8_t> payload;
  emit_u8(payload, 0x02U);
  emit_xlwidestring(payload, "surface");
  emit_xlwidestring(payload, kana);
  emit_u32(payload, 2U);
  emit_u16(payload, start);
  emit_u16(payload, 0U);
  emit_u16(payload, 1U);
  emit_u16(payload, next_start);
  emit_u16(payload, 1U);
  emit_u16(payload, 1U);
  emit_u16(payload, 0U);
  emit_u16(payload, 0x0030U);
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem), payload);
}

std::vector<std::uint8_t> BuildSstItemWithRichRuns(std::uint8_t flags, std::uint32_t run_count, std::size_t run_bytes,
                                                   bool include_phonetic_tail = false) {
  std::vector<std::uint8_t> payload;
  emit_u8(payload, flags);
  emit_xlwidestring(payload, "A");
  if ((flags & 0x01U) != 0U) {
    emit_u32(payload, run_count);
    payload.insert(payload.end(), run_bytes, 0U);
  }
  if (include_phonetic_tail && (flags & 0x02U) != 0U) {
    emit_xlwidestring(payload, "");
    emit_u32(payload, 0U);
  }

  std::vector<std::uint8_t> body;
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem), payload);
  return body;
}

Expected<std::vector<std::string>, Error> DecodeSharedStringRecords(const std::vector<std::uint8_t>& body) {
  std::deque<std::string> text_storage;
  std::vector<std::vector<PhoneticRun>> phonetic;
  std::vector<PhoneticProperties> phonetic_props;
  auto entries = DecodeSharedStringsBin(body, text_storage, phonetic, phonetic_props);
  if (!entries) {
    return std::move(entries.error());
  }
  std::vector<std::string> decoded;
  decoded.reserve(entries.value().size());
  for (const std::string_view entry : entries.value()) {
    decoded.emplace_back(entry);
  }
  return decoded;
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
  EXPECT_EQ(items[0].phonetic_props.font_id, 0U);
  EXPECT_EQ(items[0].phonetic_props.type, 0U);
  EXPECT_EQ(items[0].phonetic_props.alignment, 0U);
  ASSERT_EQ(items[1].phonetic.size(), 1U);
  EXPECT_EQ(items[1].phonetic[0].text, "ヒガシキョウト");
  EXPECT_TRUE(items[2].phonetic.empty());
}

TEST(XlsbSstBuilder, PhoneticSlicesMatchUtf16SubstringForEveryBoundaryPair) {
  const std::string kana = "A🌟BC";
  const std::vector<std::uint16_t> boundaries{0U, 1U, 2U, 3U, 4U, 5U, 6U, UINT16_MAX};
  std::vector<std::uint8_t> body;
  std::vector<std::pair<std::uint16_t, std::uint16_t>> cases;
  cases.reserve(boundaries.size() * boundaries.size());
  for (const std::uint16_t start : boundaries) {
    for (const std::uint16_t next_start : boundaries) {
      AppendPhoneticSliceCase(body, kana, start, next_start);
      cases.emplace_back(start, next_start);
    }
  }

  const std::vector<DecodedSstItem> decoded = DecodeSstItems(body);
  ASSERT_EQ(decoded.size(), cases.size());
  for (std::size_t i = 0; i < cases.size(); ++i) {
    const auto [start, next_start] = cases[i];
    const std::uint32_t length = next_start > start ? next_start - start : 0U;
    ASSERT_EQ(decoded[i].phonetic.size(), 2U);
    EXPECT_EQ(decoded[i].phonetic[0].text, utf16_substring(kana, start, length))
        << "case " << start << " -> " << next_start;
    if (next_start <= start) {
      // Equal and backwards ranges are empty, including at the surrogate midpoint.
      EXPECT_EQ(decoded[i].phonetic[0].text, "") << "case " << start << " -> " << next_start;
    }
    if (start == 1U && next_start == 2U) {
      // A positive range ending inside 🌟 rounds up and includes the codepoint.
      EXPECT_EQ(decoded[i].phonetic[0].text, "🌟");
    }
  }
}

TEST(XlsbSstBuilder, RichRunValidationAcceptsAllFlagsWithExactRunBytes) {
  for (std::uint8_t flags = 0U; flags < 4U; ++flags) {
    const std::size_t run_bytes = (flags & 0x01U) != 0U ? 4U : 0U;
    const std::vector<std::uint8_t> body = BuildSstItemWithRichRuns(flags, 1U, run_bytes, true);
    const auto decoded = DecodeSharedStringRecords(body);
    ASSERT_TRUE(static_cast<bool>(decoded)) << "flags=" << static_cast<unsigned int>(flags);
    ASSERT_EQ(decoded.value().size(), 1U);
    EXPECT_EQ(decoded.value()[0], "A");
  }
}

TEST(XlsbSstBuilder, RichRunTruncationIsRejectedWithOrWithoutPhoneticTail) {
  const std::array<std::uint8_t, 2> flag_cases{{0x01U, 0x03U}};
  for (const std::uint8_t flags : flag_cases) {
    for (std::size_t remaining = 0U; remaining < 4U; ++remaining) {
      const std::vector<std::uint8_t> body = BuildSstItemWithRichRuns(flags, 1U, remaining);
      const auto decoded = DecodeSharedStringRecords(body);
      ASSERT_FALSE(static_cast<bool>(decoded))
          << "flags=" << static_cast<unsigned int>(flags) << " remaining=" << remaining;
      EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
      EXPECT_NE(decoded.error().message.find("rich-text run array truncated"), std::string::npos);
    }
  }

  const std::vector<std::uint8_t> huge_count = BuildSstItemWithRichRuns(0x01U, UINT32_MAX, 0U);
  const auto decoded = DecodeSharedStringRecords(huge_count);
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}

TEST(XlsbSstBuilder, RichRunCannotReadBytesFromFollowingRecord) {
  std::vector<std::uint8_t> body = BuildSstItemWithRichRuns(0x01U, 1U, 0U);
  const std::vector<std::uint8_t> following = BuildSstItemWithRichRuns(0x00U, 0U, 0U);
  body.insert(body.end(), following.begin(), following.end());

  const auto decoded = DecodeSharedStringRecords(body);
  ASSERT_FALSE(static_cast<bool>(decoded));
  EXPECT_EQ(decoded.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
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

TEST(XlsbSstBuilder, BeginRecordSeparatesRepeatedReferencesFromUniqueEntries) {
  SstBuilder sst;
  EXPECT_EQ(sst.intern("same", kNoPhonetic, kDefaultProps), 0U);
  EXPECT_EQ(sst.intern("other", kNoPhonetic, kDefaultProps), 1U);
  EXPECT_EQ(sst.intern("same", kNoPhonetic, kDefaultProps), 0U);
  EXPECT_EQ(sst.total_count(), 3U);
  EXPECT_EQ(sst.size(), 2U);

  auto body_or = emit_sst(sst);
  ASSERT_TRUE(static_cast<bool>(body_or));
  ByteSpan cursor = SpanOf(body_or.value());
  auto rec = read_record(cursor);
  ASSERT_TRUE(static_cast<bool>(rec));
  ASSERT_EQ(rec.value().type, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSst));
  ByteSpan payload = rec.value().payload;
  auto total = read_u32(payload);
  auto unique = read_u32(payload);
  ASSERT_TRUE(static_cast<bool>(total));
  ASSERT_TRUE(static_cast<bool>(unique));
  EXPECT_EQ(total.value(), 3U);
  EXPECT_EQ(unique.value(), 2U);
  EXPECT_EQ(DecodeSstStream(body_or.value()), (std::vector<std::string>{"same", "other"}));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
