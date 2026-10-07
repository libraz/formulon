//
// Cross-workbook references as the formula bar spells them: which spellings
// parse, the AST they build, and their A1 / R1C1 / storage display, checked
// against Mac Excel 365 ja-JP entry and display.
//
// Pairwise model for `PairwiseEntryAndDisplay` (every pair of levels across
// the three factors appears):
//   book  : the measured bare-safe book names (letters, leading digit, dots,
//           underscore, Japanese, fullwidth, no extension, cell- and
//           bool-shaped) and one name per measured forbidden character
//           (space - ( & + , $ # % @ ~ = ; ^ { !)                    (26)
//   form  : cell, area, whole column, 3-D, book-scope name, path       (6)
//   quote : bare, quoted                                              (2)
// Rows are book x form, with quote level `(book + form) % 2`; each book and
// each form then meets both quote levels, so the 156 rows cover every pair.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/ast_format.h"
#include "parser/ast_format_r1c1.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

constexpr std::string_view kSafeBooks[] = {"Book.xlsx",
                                           "2024.xlsx",
                                           "1.xlsx",
                                           "a.b.c.xlsx",
                                           "my_book.xlsx",
                                           "\xE5\xA3\xB2\xE4\xB8\x8A.xlsx",  // 売上.xlsx
                                           "\xEF\xBD\x81\xEF\xBD\x82.xlsx",  // ａｂ.xlsx
                                           "Book",
                                           "A1.xlsx",
                                           "TRUE.xlsx"};
constexpr std::string_view kForbiddenBooks[] = {"a b.xlsx", "a-b.xlsx", "a(b.xlsx", "a&b.xlsx", "a+b.xlsx", "a,b.xlsx",
                                                "a$b.xlsx", "a#b.xlsx", "a%b.xlsx", "a@b.xlsx", "a~b.xlsx", "a=b.xlsx",
                                                "a;b.xlsx", "a^b.xlsx", "a{b.xlsx", "a!b.xlsx"};

enum class Form { kCell, kArea, kColumn, kSpan, kBookName, kPath };
constexpr Form kForms[] = {Form::kCell, Form::kArea, Form::kColumn, Form::kSpan, Form::kBookName, Form::kPath};

constexpr std::string_view kDir = "/tmp/x/";

bool HasExtension(std::string_view book) {
  return book.size() > 5 && book.substr(book.size() - 5) == ".xlsx";
}

// The qualifier through `!` for the cell forms.
std::string Qualifier(std::string_view book, std::string_view sheets, bool quoted, std::string_view dir = {}) {
  std::string q = std::string(dir) + "[" + std::string(book) + "]" + std::string(sheets);
  return quoted ? "'" + q + "'!" : q + "!";
}

std::string Entered(std::string_view book, Form form, bool quoted) {
  switch (form) {
    case Form::kCell:
      return "=" + Qualifier(book, "Sheet", quoted) + "A1";
    case Form::kArea:
      return "=SUM(" + Qualifier(book, "Sheet", quoted) + "$A$1:B2)";
    case Form::kColumn:
      return "=SUM(" + Qualifier(book, "Sheet", quoted) + "A:A)";
    case Form::kSpan:
      return "=SUM(" + Qualifier(book, "Sheet:S2", quoted) + "A1)";
    case Form::kBookName:
      return quoted ? "='" + std::string(book) + "'!Name" : "=" + std::string(book) + "!Name";
    case Form::kPath:
      return "=" + Qualifier(book, "Sheet", quoted, kDir) + "A1";
  }
  return {};
}

enum class Outcome {
  kExternal,     // an external reference to `book`, displayed as `a1`
  kRejected,     // a parse error
  kNotThisBook,  // parses, but not as a reference into `book`
};

struct Expectation {
  Outcome outcome = Outcome::kRejected;
  std::string a1;
};

Expectation Expected(std::string_view book, bool forbidden, Form form, bool quoted) {
  // Quoting is kept on display only where some part needs it, and an A1
  // 3-D span is always one quoted unit.
  const bool needs_quotes = forbidden;
  if (form == Form::kBookName) {
    // Only a file name with a workbook extension marks a book-scope name;
    // `Book!Name` is a sheet's local name, and a bare forbidden character
    // splits the run into other tokens.
    if (!HasExtension(book) || (!quoted && forbidden)) {
      return {Outcome::kNotThisBook, {}};
    }
    return {Outcome::kExternal, (needs_quotes ? "'" + std::string(book) + "'" : std::string(book)) + "!Name"};
  }
  if (!quoted && (forbidden || form == Form::kPath)) {
    return {Outcome::kRejected, {}};
  }
  switch (form) {
    case Form::kCell:
      return {Outcome::kExternal, Qualifier(book, "Sheet", needs_quotes) + "A1"};
    case Form::kArea:
      return {Outcome::kExternal, "SUM(" + Qualifier(book, "Sheet", needs_quotes) + "$A$1:B2)"};
    case Form::kColumn:
      return {Outcome::kExternal, "SUM(" + Qualifier(book, "Sheet", needs_quotes) + "A:A)"};
    case Form::kSpan:
      return {Outcome::kExternal, "SUM(" + Qualifier(book, "Sheet:S2", true) + "A1)"};
    case Form::kPath:
      return {Outcome::kExternal, Qualifier(book, "Sheet", true, kDir) + "A1"};
    case Form::kBookName:
      break;
  }
  return {};
}

// The ExternalRef the formula is or wraps (`SUM(<ref>)`), or null.
const AstNode* FindExternalRef(const AstNode& root) {
  if (root.kind() == NodeKind::ExternalRef) {
    return &root;
  }
  if (root.kind() == NodeKind::Call && root.as_call_arity() == 1U &&
      root.as_call_arg(0).kind() == NodeKind::ExternalRef) {
    return &root.as_call_arg(0);
  }
  return nullptr;
}

TEST(ExternalRefParse, PairwiseEntryAndDisplay) {
  std::size_t rows = 0;
  const std::size_t safe_count = std::size(kSafeBooks);
  const std::size_t book_count = safe_count + std::size(kForbiddenBooks);
  for (std::size_t b = 0; b < book_count; ++b) {
    const bool forbidden = b >= safe_count;
    const std::string_view book = forbidden ? kForbiddenBooks[b - safe_count] : kSafeBooks[b];
    for (std::size_t f = 0; f < std::size(kForms); ++f) {
      const bool quoted = (b + f) % 2U == 1U;
      const std::string src = Entered(book, kForms[f], quoted);
      const Expectation want = Expected(book, forbidden, kForms[f], quoted);
      ++rows;
      Arena arena;
      Parser p(src, arena);
      const AstNode* root = p.parse();
      const bool parsed = root != nullptr && p.errors().empty();
      const AstNode* ext = parsed ? FindExternalRef(*root) : nullptr;
      switch (want.outcome) {
        case Outcome::kRejected:
          EXPECT_FALSE(parsed) << src;
          break;
        case Outcome::kNotThisBook:
          EXPECT_TRUE(ext == nullptr || ext->as_external_ref_book() != book) << src;
          break;
        case Outcome::kExternal:
          ASSERT_TRUE(parsed) << src;
          ASSERT_NE(ext, nullptr) << src;
          EXPECT_EQ(ext->as_external_ref_book(), book) << src;
          EXPECT_EQ(format_formula(*root), want.a1) << src;
          break;
      }
    }
  }
  EXPECT_EQ(rows, 156U);
}

std::string Dump(std::string_view src) {
  Arena arena;
  Parser p(src, arena);
  const AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << src;
  EXPECT_TRUE(p.errors().empty()) << src;
  return root == nullptr ? std::string() : dump_sexpr(*root);
}

TEST(ExternalRefParse, QuotedAndBareSpellingsBuildTheSameTree) {
  for (const std::string_view book : kSafeBooks) {
    const std::string b(book);
    EXPECT_EQ(Dump("=[" + b + "]Sheet!A1"), Dump("='[" + b + "]Sheet'!A1")) << book;
    EXPECT_EQ(Dump("=[" + b + "]S1:S2!A1:B2"), Dump("='[" + b + "]S1:S2'!A1:B2")) << book;
  }
  EXPECT_EQ(Dump("=Book.xlsx!Name"), Dump("='Book.xlsx'!Name"));
}

TEST(ExternalRefParse, TreeCarriesPathBookAndSheets) {
  Arena arena;
  Parser p("='C:\\x\\[Book.xlsx]S1:S2'!A1:B2", arena);
  const AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  ASSERT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::ExternalRef);
  EXPECT_EQ(root->as_external_ref_path(), "C:\\x\\");
  EXPECT_EQ(root->as_external_ref_book(), "Book.xlsx");
  EXPECT_EQ(root->as_external_ref_sheet(), "S1");
  EXPECT_EQ(root->as_external_ref_sheet_end(), "S2");
  EXPECT_TRUE(root->as_external_ref_name().empty());
  EXPECT_TRUE(root->as_external_ref_is_range());
  EXPECT_FALSE(is_self_book_name_ref(*root));
}

TEST(ExternalRefParse, SheetLocalAndBookScopeNames) {
  {
    Arena arena;
    Parser p("=[Book.xlsx]Sheet!LName", arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr);
    ASSERT_TRUE(p.errors().empty());
    ASSERT_EQ(root->kind(), NodeKind::ExternalRef);
    EXPECT_EQ(root->as_external_ref_sheet(), "Sheet");
    EXPECT_EQ(root->as_external_ref_name(), "LName");
    EXPECT_EQ(format_formula(*root), "[Book.xlsx]Sheet!LName");
  }
  {
    Arena arena;
    Parser p("='/Users/x/Book.xlsx'!Name", arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr);
    ASSERT_TRUE(p.errors().empty());
    ASSERT_EQ(root->kind(), NodeKind::ExternalRef);
    EXPECT_EQ(root->as_external_ref_path(), "/Users/x/");
    EXPECT_EQ(root->as_external_ref_book(), "Book.xlsx");
    EXPECT_TRUE(root->as_external_ref_sheet().empty());
    EXPECT_EQ(format_formula(*root), "'/Users/x/Book.xlsx'!Name");
  }
  {
    // With cells after the `!`, a workbook-shaped qualifier is a sheet.
    Arena arena;
    Parser p("=Book.xlsx!A1", arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr);
    ASSERT_TRUE(p.errors().empty());
    ASSERT_EQ(root->kind(), NodeKind::Ref);
    EXPECT_EQ(root->as_ref().sheet, "Book.xlsx");
  }
}

TEST(ExternalRefParse, RejectedSpellings) {
  for (const char* src : {
           "=[Book.xlsx]!Name",             // a book-scope name is spelled without brackets
           "=/tmp/x/[Book.xlsx]Sheet!A1",   // a path must be quoted
           "=[Book.xlsx]Sheet",             // no `!`
           "=[Book 1.xlsx]Sheet!A1",        // the book needs quoting
           "='/tmp/x/Sheet'!A1",            // a directory with no bracketed book
           "='[Book.xlsx]'!A1",             // no sheet
           "=SUM([Book.xlsx]S1:S2!LName)",  // no name across a span
           "=[0]Sheet!A1",                  // `[0]` only names the own workbook's names
       }) {
    Arena arena;
    Parser p(src, arena);
    (void)p.parse();
    EXPECT_FALSE(p.errors().empty()) << src;
  }
}

TEST(ExternalRefParse, SelfBookNameStaysSelf) {
  Arena arena;
  Parser p("=[0]!Rate", arena);
  const AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  ASSERT_TRUE(p.errors().empty());
  EXPECT_TRUE(is_self_book_name_ref(*root));
  EXPECT_EQ(root->as_external_ref_book(), "0");
  EXPECT_EQ(format_formula(*root), "[0]!Rate");
}

TEST(ExternalRefParse, StructuredReferencesAreUntouched) {
  for (const char* src : {"=Table[a.b]", "=Table[@Col]", "=Table[[#This Row],[Col]]", "=SUM(Table[Col])"}) {
    Arena arena;
    Parser p(src, arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << src;
    EXPECT_TRUE(p.errors().empty()) << src;
    const AstNode& ref = root->kind() == NodeKind::Call ? root->as_call_arg(0) : *root;
    EXPECT_EQ(ref.kind(), NodeKind::StructuredRef) << src;
  }
  // A bare structured reference has no table and stays unsupported.
  Arena arena;
  Parser p("=[@Col]", arena);
  (void)p.parse();
  EXPECT_TRUE(!p.errors().empty());
}

struct DisplayRow {
  const char* entered;
  std::uint32_t row;  // 0-based host cell for R1C1
  std::uint32_t col;
  const char* a1;
  const char* r1c1;
};

constexpr DisplayRow kDisplayRows[] = {
    {"=[Book.xlsx]Sheet!A1", 0, 0, "[Book.xlsx]Sheet!A1", "[Book.xlsx]Sheet!RC"},
    {"=SUM([Book.xlsx]Sheet!A1:B2)", 1, 0, "SUM([Book.xlsx]Sheet!A1:B2)", "SUM([Book.xlsx]Sheet!R[-1]C:RC[1])"},
    {"=[Book.xlsx]Sheet!$A$1:$B$2", 0, 0, "[Book.xlsx]Sheet!$A$1:$B$2", "[Book.xlsx]Sheet!R1C1:R2C2"},
    {"=SUM([Book.xlsx]Sheet!A:A)", 0, 0, "SUM([Book.xlsx]Sheet!A:A)", "SUM([Book.xlsx]Sheet!C)"},
    {"=SUM([Book.xlsx]Sheet:S2!A1)", 31, 0, "SUM('[Book.xlsx]Sheet:S2'!A1)", "SUM([Book.xlsx]Sheet:S2!R[-31]C)"},
    {"=[Book.xlsx]2024!A1", 0, 0, "'[Book.xlsx]2024'!A1", "'[Book.xlsx]2024'!RC"},
    {"='[Book.xlsx]Sheet'!A1", 0, 0, "[Book.xlsx]Sheet!A1", "[Book.xlsx]Sheet!RC"},
    {"=[Book.xlsx]S2!A1", 0, 0, "[Book.xlsx]S2!A1", "[Book.xlsx]S2!RC"},
    {"=[Book.xlsx]TRUE!A1", 0, 0, "[Book.xlsx]TRUE!A1", "[Book.xlsx]TRUE!RC"},
    {"=[Book.xlsx]R1C1!A1", 0, 0, "[Book.xlsx]R1C1!A1", "[Book.xlsx]R1C1!RC"},
    {"='/private/tmp/extref-probe/src/[Book.xlsx]Sheet'!A1", 0, 0,
     "'/private/tmp/extref-probe/src/[Book.xlsx]Sheet'!A1", "'/private/tmp/extref-probe/src/[Book.xlsx]Sheet'!RC"},
    {"=SUM('/private/tmp/extref-probe/src/[Book.xlsx]Sheet:S2'!A1)", 6, 0,
     "SUM('/private/tmp/extref-probe/src/[Book.xlsx]Sheet:S2'!A1)",
     "SUM('/private/tmp/extref-probe/src/[Book.xlsx]Sheet:S2'!R[-6]C)"},
    {"='/private/tmp/extref-probe/src/Book.xlsx'!Name", 0, 0, "'/private/tmp/extref-probe/src/Book.xlsx'!Name",
     "'/private/tmp/extref-probe/src/Book.xlsx'!Name"},
    {"='C:\\x\\[Book.xlsx]Sheet'!A1", 0, 0, "'C:\\x\\[Book.xlsx]Sheet'!A1", "'C:\\x\\[Book.xlsx]Sheet'!RC"},
    {"=Book.xlsx!Name", 0, 0, "Book.xlsx!Name", "Book.xlsx!Name"},
};

TEST(ExternalRefParse, DisplayMatchesTheFormulaBar) {
  for (const DisplayRow& row : kDisplayRows) {
    Arena arena;
    Parser p(row.entered, arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << row.entered;
    ASSERT_TRUE(p.errors().empty()) << row.entered;
    EXPECT_EQ(format_formula(*root), row.a1) << row.entered;
    EXPECT_EQ(format_formula_r1c1(*root, row.row, row.col), row.r1c1) << row.entered;
  }
}

std::string SpellAsIs(std::string_view name) {
  return std::string(name);
}

// Binds `Book.xlsx` (with or without a directory) to link 1 and `Other.xlsx`
// to link 2.
std::uint32_t IndexBook(const void* /*ctx*/, std::string_view /*path*/, std::string_view book) {
  if (book == "Book.xlsx") {
    return 1U;
  }
  return book == "Other.xlsx" ? 2U : 0U;
}

TEST(ExternalRefParse, StorageSpellingNamesTheLinkIndex) {
  const ExternalBookIndexer indexer{&IndexBook, nullptr};
  struct Row {
    const char* entered;
    const char* stored;
  };
  constexpr Row kRows[] = {
      {"=[Book.xlsx]Sheet!A1", "[1]Sheet!A1"},
      {"='[Other.xlsx]Sheet 1'!A1", "'[2]Sheet 1'!A1"},
      {"=Book.xlsx!Name", "[1]!Name"},
      {"=SUM([Book.xlsx]S1:S2!A1)", "SUM('[1]S1:S2'!A1)"},
      {"='/Users/x/[Book.xlsx]Sheet'!A1", "[1]Sheet!A1"},
      {"='/Users/x/Book.xlsx'!Name", "[1]!Name"},
      {"=[Book.xlsx]Sheet!LName", "[1]Sheet!LName"},
      {"=[0]!Rate", "[0]!Rate"},
  };
  for (const Row& row : kRows) {
    Arena arena;
    Parser p(row.entered, arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << row.entered;
    ASSERT_TRUE(p.errors().empty()) << row.entered;
    EXPECT_EQ(format_formula_storage(*root, &SpellAsIs, nullptr, &indexer), row.stored) << row.entered;
    // Without an indexer the book keeps its spelling.
    EXPECT_EQ(format_formula_storage(*root, &SpellAsIs), format_formula(*root)) << row.entered;
  }
}

}  // namespace
}  // namespace parser
}  // namespace formulon
