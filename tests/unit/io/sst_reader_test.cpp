// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `formulon::io::read_shared_strings`. The reader works
// against an in-memory `<sst>` byte stream, so each test builds the
// payload from a string literal and feeds it to the reader directly. A
// `std::deque<std::string>` owns the resulting payloads; the
// `string_view`s in `SharedStringTable::entries` alias entries in that
// deque, so the deque must outlive the table.

#include "io/sst_reader.h"

#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "utils/error.h"

namespace formulon {
namespace io {
namespace {

/// Wraps a string literal into the byte vector the reader consumes.
std::vector<std::uint8_t> Bytes(std::string_view xml) {
  return std::vector<std::uint8_t>(xml.begin(), xml.end());
}

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

// ---------------------------------------------------------------------------
// Happy-path coverage
// ---------------------------------------------------------------------------

TEST(SstReader, EmptySstYieldsZeroEntries) {
  std::string xml(kXmlDecl);
  xml.append(
      "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read failed: " << result_or.error().message;
  EXPECT_EQ(result_or.value().entries.size(), 0U);
  EXPECT_TRUE(storage.empty());
}

TEST(SstReader, SinglePlainEntry) {
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si><t>hello</t></si>");
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().entries.size(), 1U);
  EXPECT_EQ(result_or.value().entries[0], "hello");
}

TEST(SstReader, MultipleEntriesPreserveOrder) {
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si><t>alpha</t></si>");
  xml.append("<si><t>beta</t></si>");
  xml.append("<si><t>gamma</t></si>");
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().entries.size(), 3U);
  EXPECT_EQ(result_or.value().entries[0], "alpha");
  EXPECT_EQ(result_or.value().entries[1], "beta");
  EXPECT_EQ(result_or.value().entries[2], "gamma");
}

TEST(SstReader, XmlSpacePreserveKeepsLeadingAndTrailingWhitespace) {
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si><t xml:space=\"preserve\">  spaced  </t></si>");
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().entries.size(), 1U);
  EXPECT_EQ(result_or.value().entries[0], "  spaced  ");
}

TEST(SstReader, RichTextRunsConcatenateInOrder) {
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si>");
  xml.append("<r><rPr><b/></rPr><t>foo</t></r>");
  xml.append("<r><t>bar</t></r>");
  xml.append("<r><rPr><i/></rPr><t>baz</t></r>");
  xml.append("</si>");
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().entries.size(), 1U);
  EXPECT_EQ(result_or.value().entries[0], "foobarbaz");
}

TEST(SstReader, PhoneticGuidesAreIgnored) {
  // <si><t>kanji</t><rPh sb="0" eb="2"><t>furigana</t></rPh></si>
  // The phonetic guide must be skipped: we keep only the base "kanji"
  // payload. (The <t> inside <rPh> is intentionally NOT concatenated.)
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si>");
  xml.append("<t>kanji</t>");
  xml.append("<rPh sb=\"0\" eb=\"2\"><t>furigana</t></rPh>");
  xml.append("</si>");
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().entries.size(), 1U);
  EXPECT_EQ(result_or.value().entries[0], "kanji");
}

TEST(SstReader, JapaneseUnicode) {
  // "日本語" (E6 97 A5 E6 9C AC E8 AA 9E)
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si><t>\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E</t></si>");
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().entries.size(), 1U);
  EXPECT_EQ(result_or.value().entries[0], "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E");
}

TEST(SstReader, EmojiUnicode) {
  // "😀" (F0 9F 98 80) — outside the BMP. Should round-trip byte-for-byte.
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si><t>\xF0\x9F\x98\x80</t></si>");
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().entries.size(), 1U);
  EXPECT_EQ(result_or.value().entries[0], "\xF0\x9F\x98\x80");
}

// ---------------------------------------------------------------------------
// Lifetime
// ---------------------------------------------------------------------------

TEST(SstReader, EntriesViewsAliasTextStorage) {
  // Append many strings to force the deque to allocate multiple chunks,
  // then verify every view still resolves to the right content. This
  // exercises pugixml-driven appends where the underlying storage grows
  // mid-walk.
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  for (int i = 0; i < 64; ++i) {
    xml.append("<si><t>entry-");
    xml.append(std::to_string(i));
    xml.append("</t></si>");
  }
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().entries.size(), 64U);
  for (int i = 0; i < 64; ++i) {
    std::string expected = "entry-";
    expected.append(std::to_string(i));
    EXPECT_EQ(result_or.value().entries[static_cast<std::size_t>(i)], expected);
  }
  // text_storage must hold every payload still — its pointer stability
  // is the lifetime guarantee these views rely on.
  EXPECT_EQ(storage.size(), 64U);
}

// ---------------------------------------------------------------------------
// Error paths
// ---------------------------------------------------------------------------

TEST(SstReader, RejectsMalformedXml) {
  std::string xml = "<sst><si><t>oops";  // unterminated
  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoXmlParse);
}

TEST(SstReader, RejectsMissingSstRoot) {
  std::string xml(kXmlDecl);
  xml.append("<somethingElse/>");
  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoXmlParse);
}

TEST(SstReader, RejectsSiWithNoText) {
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si></si>");  // no <t>, no <r>, no <rPh>
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
  // The error context should pinpoint which entry failed.
  EXPECT_NE(result_or.error().context.find("index=0"), std::string::npos);
}

TEST(SstReader, RejectsSiWithOnlyPhoneticGuides) {
  // No base <t> and no <r><t> — only an rPh subtree — is a degenerate
  // case Excel never emits; we treat it as data loss rather than letting
  // it round-trip to an empty string.
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si><rPh sb=\"0\" eb=\"1\"><t>furigana</t></rPh></si>");
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(SstReader, EmptyTextElementIsAccepted) {
  // A <si><t></t></si> is legal: the writer might emit it for an
  // explicit empty-string literal. We must NOT error on it.
  std::string xml(kXmlDecl);
  xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
  xml.append("<si><t></t></si>");
  xml.append("</sst>");

  std::deque<std::string> storage;
  auto result_or = read_shared_strings(Bytes(xml), storage);
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().entries.size(), 1U);
  EXPECT_EQ(result_or.value().entries[0], "");
}

}  // namespace
}  // namespace io
}  // namespace formulon
