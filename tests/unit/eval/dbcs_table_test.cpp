// Unit tests for the DBCS code page tables behind CHAR / CODE. Expected values
// come from code_char_jp_probes on the Mac ja-JP, zh-CN and ko-KR targets.

#include "eval/dbcs_table.h"

#include <cstdint>

#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace {

std::uint32_t code_of(DbcsCodepage codepage, std::uint32_t codepoint) {
  const std::uint16_t packed = lookup_unicode_to_dbcs(codepage, codepoint);
  return packed == 0u ? 95u : static_cast<std::uint32_t>(packed + dbcs_code_bias(codepage));
}

TEST(DbcsTable, Jis0208UnicodeToCode) {
  EXPECT_EQ(code_of(DbcsCodepage::kJis0208, 0x3042u), 9250u);   // あ
  EXPECT_EQ(code_of(DbcsCodepage::kJis0208, 0x4E00u), 12396u);  // 一
  EXPECT_EQ(code_of(DbcsCodepage::kJis0208, 0x6F22u), 13377u);  // 漢
  EXPECT_EQ(code_of(DbcsCodepage::kJis0208, 0xFFE5u), 8559u);   // ￥
}

TEST(DbcsTable, Jis0208Unmapped) {
  // NEC extension, supplementary plane, and U+0000 (the empty-slot sentinel).
  EXPECT_EQ(lookup_unicode_to_dbcs(DbcsCodepage::kJis0208, 0x9AD9u), 0u);
  EXPECT_EQ(lookup_unicode_to_dbcs(DbcsCodepage::kJis0208, 0x1F600u), 0u);
  EXPECT_EQ(lookup_unicode_to_dbcs(DbcsCodepage::kJis0208, 0u), 0u);
}

TEST(DbcsTable, Jis0208RowCellToUnicode) {
  EXPECT_EQ(lookup_dbcs_to_unicode(DbcsCodepage::kJis0208, 1, 1), 0x3000u);
  EXPECT_EQ(lookup_dbcs_to_unicode(DbcsCodepage::kJis0208, 4, 2), 0x3042u);
  EXPECT_EQ(lookup_dbcs_to_unicode(DbcsCodepage::kJis0208, 0, 2), 0u);
  EXPECT_EQ(lookup_dbcs_to_unicode(DbcsCodepage::kJis0208, 95, 2), 0u);
  EXPECT_EQ(lookup_dbcs_to_unicode(DbcsCodepage::kJis0208, 4, 0), 0u);
  EXPECT_EQ(lookup_dbcs_to_unicode(DbcsCodepage::kJis0208, 4, 95), 0u);
}

TEST(DbcsTable, Gb2312UnicodeToCode) {
  EXPECT_EQ(code_of(DbcsCodepage::kGb2312, 0x3042u), 42146u);  // あ
  EXPECT_EQ(code_of(DbcsCodepage::kGb2312, 0x30F4u), 42484u);  // ヴ
  EXPECT_EQ(code_of(DbcsCodepage::kGb2312, 0x4E00u), 53947u);  // 一
  EXPECT_EQ(code_of(DbcsCodepage::kGb2312, 0xFFE5u), 41892u);  // ￥
  EXPECT_EQ(code_of(DbcsCodepage::kGb2312, 0x6F22u), 95u);     // 漢, GBK only
  EXPECT_EQ(code_of(DbcsCodepage::kGb2312, 0xFF71u), 95u);     // ｱ
}

TEST(DbcsTable, KsX1001UnicodeToCode) {
  EXPECT_EQ(code_of(DbcsCodepage::kKsX1001, 0x3042u), 43682u);  // あ
  EXPECT_EQ(code_of(DbcsCodepage::kKsX1001, 0x4E00u), 60649u);  // 一
  EXPECT_EQ(code_of(DbcsCodepage::kKsX1001, 0x6F22u), 63955u);  // 漢
  EXPECT_EQ(code_of(DbcsCodepage::kKsX1001, 0x93ACu), 64480u);  // 鎬
  EXPECT_EQ(code_of(DbcsCodepage::kKsX1001, 0xD55Cu), 51153u);  // 한
  EXPECT_EQ(code_of(DbcsCodepage::kKsX1001, 0xFFE5u), 41421u);  // ￥
}

TEST(DbcsTable, CodeBias) {
  EXPECT_EQ(dbcs_code_bias(DbcsCodepage::kJis0208), 0x2020u);
  EXPECT_EQ(dbcs_code_bias(DbcsCodepage::kGb2312), 0xA0A0u);
  EXPECT_EQ(dbcs_code_bias(DbcsCodepage::kKsX1001), 0xA0A0u);
  EXPECT_EQ(lookup_unicode_to_dbcs(DbcsCodepage::kNone, 0x3042u), 0u);
}

TEST(DbcsTable, EveryMappedCellRoundTrips) {
  for (DbcsCodepage codepage : {DbcsCodepage::kJis0208, DbcsCodepage::kGb2312, DbcsCodepage::kKsX1001}) {
    int mapped = 0;
    for (std::uint8_t row = 1; row <= kDbcsGridSize; ++row) {
      for (std::uint8_t cell = 1; cell <= kDbcsGridSize; ++cell) {
        const std::uint16_t cp = lookup_dbcs_to_unicode(codepage, row, cell);
        if (cp == 0u) {
          continue;
        }
        ++mapped;
        ASSERT_EQ(lookup_unicode_to_dbcs(codepage, cp), static_cast<std::uint16_t>((row << 8) | cell))
            << "codepage " << static_cast<int>(codepage) << " U+" << std::hex << cp;
      }
    }
    // Mapped-cell counts of the generated tables.
    const int expected = codepage == DbcsCodepage::kJis0208 ? 6879 : codepage == DbcsCodepage::kGb2312 ? 7445 : 8225;
    EXPECT_EQ(mapped, expected);
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
