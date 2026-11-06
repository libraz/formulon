// Name syntax against the names Mac Excel 365 accepts as a table display name
// and as a defined name.

#include "utils/name_syntax.h"

#include <string>

#include "gtest/gtest.h"

namespace formulon {
namespace {

TEST(NameSyntax, TableNamesExcelKeeps) {
  for (const char* name : {"Tbl",
                           "Table1",
                           "_a",
                           "_",
                           "_1",
                           "\\",
                           "\\a",
                           "a\\b",
                           "a.b",
                           "a.",
                           "a?b",
                           "x",
                           "SUM",
                           "Sheet1",
                           "A1A",
                           "XFE1",
                           "AAAA1",
                           "A0",
                           "XFD0",
                           "A1048577",
                           "R0",
                           "R0C1",
                           "RC0",
                           "C0",
                           "R1048577",
                           "Rx",
                           "RC_",
                           "\xEF\xBC\xB2\xEF\xBC\xA3_",
                           "\xE3\x81\x82",
                           "\xE8\xA1\xA8"
                           "1",
                           "\xCE\xA3x",
                           "\xEF\xBD\x81\xEF\xBD\x82"}) {
    EXPECT_TRUE(is_valid_table_name(name)) << name;
  }
}

TEST(NameSyntax, TableNamesExcelRewrites) {
  for (const char* name : {"",
                           "T1",
                           "A1",
                           "a1",
                           "XFD1",
                           "XFD1048576",
                           "A01",
                           "ab1",
                           "Abc123",
                           "Ra1",
                           "C16385",
                           "R",
                           "C",
                           "r",
                           "c",
                           "RC",
                           "R1",
                           "C1",
                           "RC1",
                           "R1C",
                           "R1C1",
                           "R1C1A",
                           "R2D2",
                           "C1a",
                           "r1x",
                           "R1C0",
                           "R1048576",
                           "C16384",
                           "R1C16385",
                           "TRUE",
                           "FALSE",
                           "true",
                           "1a",
                           ".a",
                           "?a",
                           "a b",
                           "a-b",
                           "a[b",
                           "a]b",
                           "a!b",
                           "a$b",
                           "a,b",
                           "a:b",
                           "R[1]C",
                           "a\xE3\x80\x80"
                           "b",
                           "\xEF\xBC\x91"
                           "a",
                           "\xEF\xBD\x81\xEF\xBC\x8D\xEF\xBD\x82",
                           "\xEF\xBC\xB4\xEF\xBC\xB2\xEF\xBC\xB5\xEF\xBC\xA5",
                           "\xEF\xBC\xA1"
                           "1",
                           "\xEF\xBC\xB2"
                           "1",
                           "\xEF\xBD\x92\xEF\xBD\x83"}) {
    EXPECT_FALSE(is_valid_table_name(name)) << name;
  }
}

TEST(NameSyntax, LengthCountsUtf16Units) {
  EXPECT_TRUE(is_valid_table_name(std::string(255, 'a')));
  EXPECT_FALSE(is_valid_table_name(std::string(256, 'a')));
  std::string kana;
  for (int i = 0; i < 255; ++i) {
    kana += "\xE3\x81\x82";
  }
  EXPECT_TRUE(is_valid_table_name(kana));
  EXPECT_FALSE(is_valid_table_name(kana + "\xE3\x81\x82"));
  std::string emoji;
  for (int i = 0; i < 127; ++i) {
    emoji += "\xF0\x9F\x98\x80";
  }
  EXPECT_TRUE(is_valid_table_name(emoji + "a"));
  EXPECT_FALSE(is_valid_table_name(emoji + "\xF0\x9F\x98\x80"));
}

TEST(NameSyntax, DefinedNamesRefuseBackslashAndOneCharacter) {
  for (const char* name : {"\\a", "\\b", "\\A", "\\_", "\\1", "\\.", "\\\\", "\\z", "\\\xE3\x81\x82"}) {
    EXPECT_FALSE(is_valid_defined_name(name)) << name;
  }
  for (const char* name : {"\\", "\\ab", "\\x1", "\\abc", "x\\a", "a\\b", "_a", "ab", "_1", "a_", "_", "x",
                           "Print_Area", "_xlnm.Print_Area", "XFE1", "A1048577"}) {
    EXPECT_TRUE(is_valid_defined_name(name)) << name;
  }
  for (const char* name : {"T1", "A1", "R1C1", "RC", "R", "C", "1a", ".a", "a b", "a-b", "TRUE"}) {
    EXPECT_FALSE(is_valid_defined_name(name)) << name;
  }
}

}  // namespace
}  // namespace formulon
