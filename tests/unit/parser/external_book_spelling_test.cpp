// Rewriting `[N]` in stored formula text to the formula-bar spelling, and
// mapping a link target to the directory a qualifier displays.

#include "parser/external_book_spelling.h"

#include <cstdint>
#include <string>
#include <string_view>

#include "gtest/gtest.h"

namespace formulon {
namespace parser {
namespace {

bool Resolve(const void* /*ctx*/, std::uint32_t index, ExternalBookDisplay* out) {
  switch (index) {
    case 1:
      out->path = "/Users/libraz/Documents/ext_link_probe/";
      out->book = "ExtSource.xlsx";
      return true;
    case 2:
      out->path.clear();
      out->book = "ExtSource.xlsx";
      return true;
    case 3:
      out->path.clear();
      out->book = "My Book.xlsx";
      return true;
    case 4:
      out->path = "C:\\a b\\";
      out->book = "It's.xlsx";
      return true;
    default:
      return false;
  }
}

std::string Spell(std::string_view text) {
  const ExternalBookResolver resolver{&Resolve, nullptr};
  return spell_external_books(text, resolver);
}

TEST(ExternalBookSpellingRewrite, PathQualifiedSheetIsQuotedWhole) {
  EXPECT_EQ(Spell("[1]Data!A1"), "'/Users/libraz/Documents/ext_link_probe/[ExtSource.xlsx]Data'!A1");
}

TEST(ExternalBookSpellingRewrite, PathlessSheetStaysBare) {
  EXPECT_EQ(Spell("[2]Data!A1"), "[ExtSource.xlsx]Data!A1");
}

TEST(ExternalBookSpellingRewrite, BookScopeName) {
  EXPECT_EQ(Spell("[1]!SrcTotal"), "'/Users/libraz/Documents/ext_link_probe/ExtSource.xlsx'!SrcTotal");
  EXPECT_EQ(Spell("[2]!SrcTotal"), "ExtSource.xlsx!SrcTotal");
  EXPECT_EQ(Spell("[3]!SrcTotal"), "'My Book.xlsx'!SrcTotal");
}

TEST(ExternalBookSpellingRewrite, QuotedInputIsRequotedByRule) {
  EXPECT_EQ(Spell("'[2]Sheet 1'!A1"), "'[ExtSource.xlsx]Sheet 1'!A1");
  EXPECT_EQ(Spell("'[2]Data'!A1"), "[ExtSource.xlsx]Data!A1");
  EXPECT_EQ(Spell("'[2]2024'!A1"), "'[ExtSource.xlsx]2024'!A1");
}

TEST(ExternalBookSpellingRewrite, ThreeDStaysQuotedInA1) {
  EXPECT_EQ(Spell("'[2]S1:S2'!A1"), "'[ExtSource.xlsx]S1:S2'!A1");
  EXPECT_EQ(Spell("[2]S1:S2!A1"), "'[ExtSource.xlsx]S1:S2'!A1");
}

TEST(ExternalBookSpellingRewrite, BookNeedingQuotesQuotesTheQualifier) {
  EXPECT_EQ(Spell("[3]Data!A1"), "'[My Book.xlsx]Data'!A1");
}

TEST(ExternalBookSpellingRewrite, EmbeddedQuotesAreDoubled) {
  EXPECT_EQ(Spell("[4]Data!A1"), "'C:\\a b\\[It''s.xlsx]Data'!A1");
  EXPECT_EQ(Spell("'[2]It''s'!A1"), "'[ExtSource.xlsx]It''s'!A1");
}

TEST(ExternalBookSpellingRewrite, UnmatchedIndexIsSpelledDecimal) {
  EXPECT_EQ(Spell("[7]Data!A1"), "[7]Data!A1");
  EXPECT_EQ(Spell("[7]!Name"), "7!Name");
}

TEST(ExternalBookSpellingRewrite, SpacesAndCaseOutsideQualifiersSurvive) {
  EXPECT_EQ(Spell("=SUM( [2]Data!A1 , 2 )"), "=SUM( [ExtSource.xlsx]Data!A1 , 2 )");
  EXPECT_EQ(Spell("sum([2]Data!A1:B2)+[2]Data!C1"), "sum([ExtSource.xlsx]Data!A1:B2)+[ExtSource.xlsx]Data!C1");
}

TEST(ExternalBookSpellingRewrite, MultipleQualifiersInOneFormula) {
  EXPECT_EQ(Spell("[2]A!A1+[3]B!A1+[2]!N"), "[ExtSource.xlsx]A!A1+'[My Book.xlsx]B'!A1+ExtSource.xlsx!N");
}

TEST(ExternalBookSpellingUntouched, TextWithoutBracketIsReturnedAsIs) {
  EXPECT_EQ(Spell("SUM(A1:B2)+'My Sheet'!C3"), "SUM(A1:B2)+'My Sheet'!C3");
}

TEST(ExternalBookSpellingUntouched, SelfBookName) {
  EXPECT_EQ(Spell("[0]!Name+1"), "[0]!Name+1");
}

TEST(ExternalBookSpellingUntouched, StructuredReferences) {
  EXPECT_EQ(Spell("Table1[Col]"), "Table1[Col]");
  EXPECT_EQ(Spell("Table1[[#This Row],[Col]]"), "Table1[[#This Row],[Col]]");
  EXPECT_EQ(Spell("[@Col]+[1]"), "[@Col]+[1]");
  EXPECT_EQ(Spell("Table1[1]!A1"), "Table1[1]!A1");
}

TEST(ExternalBookSpellingUntouched, BracketsInsideStringLiterals) {
  EXPECT_EQ(Spell("\"[1]Data!A1\"&[2]Data!A1"), "\"[1]Data!A1\"&[ExtSource.xlsx]Data!A1");
  EXPECT_EQ(Spell("\"a\"\"[1]x!\""), "\"a\"\"[1]x!\"");
}

TEST(ExternalBookSpellingUntouched, QuotedLocalSheetWithBracketText) {
  EXPECT_EQ(Spell("'a[b]c'!A1"), "'a[b]c'!A1");
  EXPECT_EQ(Spell("'[x]Sheet'!A1"), "'[x]Sheet'!A1");
}

TEST(ExternalBookSpellingUntouched, IncompleteQualifier) {
  EXPECT_EQ(Spell("[2]Data"), "[2]Data");
}

TEST(ExternalBookSpellingDisplayPath, FileUrlWithDrive) {
  EXPECT_EQ(display_path_for_link_target("file:///C:/a/b/Book.xlsx"), "C:\\a\\b\\");
  EXPECT_EQ(display_path_for_link_target("file:///C:\\a\\b\\Book.xlsx"), "C:\\a\\b\\");
}

TEST(ExternalBookSpellingDisplayPath, PosixPaths) {
  EXPECT_EQ(display_path_for_link_target("file:///Users/x/Book.xlsx"), "/Users/x/");
  EXPECT_EQ(display_path_for_link_target("/Users/x/Book.xlsx"), "/Users/x/");
}

TEST(ExternalBookSpellingDisplayPath, UncShare) {
  EXPECT_EQ(display_path_for_link_target("file://server/share/dir/Book.xlsx"), "\\\\server\\share\\dir\\");
  EXPECT_EQ(display_path_for_link_target("file://server/share/Book.xlsx"), "\\\\server\\share\\");
}

TEST(ExternalBookSpellingDisplayPath, HttpDirectory) {
  EXPECT_EQ(display_path_for_link_target("https://host/a/b/Book.xlsx"), "https://host/a/b/");
  EXPECT_EQ(display_path_for_link_target("http://host/Book.xlsx"), "http://host/");
}

TEST(ExternalBookSpellingDisplayPath, RelativeAndFileNameOnlyHaveNoPath) {
  EXPECT_EQ(display_path_for_link_target("../x/Book.xlsx"), "");
  EXPECT_EQ(display_path_for_link_target("Book.xlsx"), "");
  EXPECT_EQ(display_path_for_link_target(""), "");
}

TEST(ExternalBookSpellingDisplayPath, FileUrlPercentEscapesAreDecoded) {
  EXPECT_EQ(display_path_for_link_target("file:///Users/my%20dir/Book.xlsx"), "/Users/my dir/");
  EXPECT_EQ(display_path_for_link_target("file:///C:/My%20Docs/Book.xlsx"), "C:\\My Docs\\");
}

TEST(ExternalBookSpellingDisplayPath, PercentInPlainPathIsLiteral) {
  EXPECT_EQ(display_path_for_link_target("/Users/a%20b/Book.xlsx"), "/Users/a%20b/");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
