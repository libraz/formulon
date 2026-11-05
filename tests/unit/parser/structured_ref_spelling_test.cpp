#include "parser/structured_ref_spelling.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_format_r1c1.h"
#include "parser/parse_error.h"
#include "parser/parser.h"
#include "parser/structured_ref_payload.h"
#include "utils/arena.h"
#include "utils/strings.h"

// Expectations are Mac Excel 16.114 (ja-JP) readings: `Range.Formula` for the
// display form, the saved file for the stored form.

namespace formulon {
namespace parser {
namespace {

using Specs = StructuredRefSpecifiers;

bool IsTestTable(const void* /*ctx*/, std::string_view name) {
  return strings::case_insensitive_eq(name, "tblSales") || strings::case_insensitive_eq(name, "tblOther");
}

constexpr BareTableNamePredicate kTables{&IsTestTable, nullptr};

std::optional<StructuredRefPayload> Parse(std::string_view text) {
  return parse_structured_ref_payload_text(text);
}

// --- Payload grammar --------------------------------------------------------

TEST(StructuredRefPayloadParse, Shapes) {
  auto p = Parse("");
  ASSERT_TRUE(p);
  EXPECT_EQ(p->specifiers, 0U);
  EXPECT_TRUE(p->columns.empty());

  p = Parse("@");
  ASSERT_TRUE(p);
  EXPECT_EQ(p->specifiers, Specs::kThisRow);
  EXPECT_TRUE(p->columns.empty());

  p = Parse("@Qty");
  ASSERT_TRUE(p);
  EXPECT_EQ(p->specifiers, Specs::kThisRow);
  EXPECT_EQ(p->columns, std::vector<std::string>{"Qty"});

  p = Parse("@[Unit Cost]");
  ASSERT_TRUE(p);
  EXPECT_EQ(p->columns, std::vector<std::string>{"Unit Cost"});

  p = Parse("@[Qty]:[Price]");
  ASSERT_TRUE(p);
  EXPECT_EQ(p->specifiers, Specs::kThisRow);
  EXPECT_EQ(p->columns, (std::vector<std::string>{"Qty", "Price"}));

  p = Parse("#totals");
  ASSERT_TRUE(p);
  EXPECT_EQ(p->specifiers, Specs::kTotals);

  p = Parse("[#Headers],[#Data],[Qty]");
  ASSERT_TRUE(p);
  EXPECT_EQ(p->specifiers, Specs::kHeaders | Specs::kData);
  EXPECT_EQ(p->columns, std::vector<std::string>{"Qty"});

  p = Parse("[#This Row], [Qty]:[Price]");
  ASSERT_TRUE(p);
  EXPECT_EQ(p->specifiers, Specs::kThisRow);
  EXPECT_EQ(p->columns, (std::vector<std::string>{"Qty", "Price"}));

  p = Parse("[ ab]");
  ASSERT_TRUE(p);
  EXPECT_EQ(p->columns, std::vector<std::string>{" ab"});
}

// An unbracketed payload is one column name, `,` and `:` included.
TEST(StructuredRefPayloadParse, UnbracketedIsOneColumn) {
  for (const char* name : {"a,b", "a:b", "Unit Cost", "a@b"}) {
    SCOPED_TRACE(name);
    const auto p = Parse(name);
    ASSERT_TRUE(p);
    EXPECT_EQ(p->specifiers, 0U);
    EXPECT_EQ(p->columns, std::vector<std::string>{name});
  }
}

// The parsed columns, or `<rejected>`.
std::vector<std::string> Columns(std::string_view text) {
  const auto p = Parse(text);
  return p ? p->columns : std::vector<std::string>{"<rejected>"};
}

TEST(StructuredRefPayloadParse, Unescape) {
  EXPECT_EQ(Columns("a''b"), std::vector<std::string>{"a'b"});
  EXPECT_EQ(Columns("a'[b"), std::vector<std::string>{"a[b"});
  EXPECT_EQ(Columns("a']b"), std::vector<std::string>{"a]b"});
  EXPECT_EQ(Columns("a'#b"), std::vector<std::string>{"a#b"});
  EXPECT_EQ(Columns("@[a'@b]"), std::vector<std::string>{"a@b"});
  EXPECT_EQ(Columns("'@x"), std::vector<std::string>{"@x"});
}

TEST(StructuredRefPayloadParse, Rejects) {
  for (const char* bad : {"#Bogus", "[#Bogus],[Qty]", "[Qty", "[A],[B]", "@[#Headers]", "[A]:", "a[b", "[A]x"}) {
    SCOPED_TRACE(bad);
    EXPECT_FALSE(Parse(bad));
  }
}

// --- Stored form (S1) --------------------------------------------------------

TEST(StructuredRefPayloadStored, ExcelSavedSpelling) {
  struct Row {
    const char* written;
    const char* stored;
  };
  const Row rows[] = {
      {"[#This Row],[a''b]", "[#This Row],[a''b]"},
      {"a'[b", "a'[b"},
      {"a']b", "a']b"},
      {"a'#b", "a'#b"},
      {"@[a'@b]", "[#This Row],[a@b]"},
      {"a'@b", "a@b"},
      {"a,b", "a,b"},
      {"a:b", "a:b"},
      {"@[a:b]", "[#This Row],[a:b]"},
      {"[ ab]", "[ ab]"},
      {"@[ ab]", "[#This Row],[ ab]"},
      {"@Qty", "[#This Row],[Qty]"},
      {"@", "#This Row"},
      {"", ""},
      {"#All", "#All"},
      {"[#Data],[#Headers],[Qty]", "[#Headers],[#Data],[Qty]"},
      {"[Qty]:[Price]", "[Qty]:[Price]"},
      {"@[Qty]:[Price]", "[#This Row],[Qty]:[Price]"},
  };
  for (const Row& row : rows) {
    SCOPED_TRACE(row.written);
    const auto p = Parse(row.written);
    ASSERT_TRUE(p);
    EXPECT_EQ(stored_payload(*p), row.stored);
  }
}

// --- Display form (D1, D2, D3) ----------------------------------------------

// D1: tblSales with a header row, 20 data rows and a totals row.
struct D1Row {
  const char* stored;
  const char* inside;
  const char* outside;
};

const D1Row kD1[] = {
    {"=tblSales[[#This Row],[Qty]]", "=[@Qty]", "=tblSales[@Qty]"},
    {"=tblSales[[#This Row],[Unit Cost]]", "=[@[Unit Cost]]", "=tblSales[@[Unit Cost]]"},
    {"=SUM(tblSales[Qty])", "=SUM([Qty])", "=SUM(tblSales[Qty])"},
    {"=SUM(tblSales[Unit Cost])", "=SUM([Unit Cost])", "=SUM(tblSales[Unit Cost])"},
    {"=tblSales[[#This Row],[a'#b]]", "=[@[a'#b]]", "=tblSales[@[a'#b]]"},
    {"=SUM(tblSales[[#This Row],[Qty]:[Price]])", "=SUM(tblSales[@[Qty]:[Price]])", "=SUM(tblSales[@[Qty]:[Price]])"},
    {"=SUM(tblSales[#This Row])", "=SUM(tblSales[@])", "=SUM(tblSales[@])"},
    {"=tblSales[[#Headers],[Qty]]", "=tblSales[[#Headers],[Qty]]", "=tblSales[[#Headers],[Qty]]"},
    {"=tblSales[[#Totals],[Qty]]", "=tblSales[[#Totals],[Qty]]", "=tblSales[[#Totals],[Qty]]"},
    {"=ROWS(tblSales[#All])", "=ROWS(tblSales[#All])", "=ROWS(tblSales[#All])"},
    {"=ROWS(tblSales[#Data])", "=ROWS(tblSales[#Data])", "=ROWS(tblSales[#Data])"},
    {"=COUNTA(tblSales[#Headers])", "=COUNTA(tblSales[#Headers])", "=COUNTA(tblSales[#Headers])"},
    {"=ROWS(tblSales[])", "=ROWS(tblSales)", "=ROWS(tblSales)"},
    {"=ROWS(tblSales[[#Headers],[#Data],[Qty]])", "=ROWS(tblSales[[#Headers],[#Data],[Qty]])",
     "=ROWS(tblSales[[#Headers],[#Data],[Qty]])"},
    {"=SUM(tblSales[[Qty]:[Price]])", "=SUM(tblSales[[Qty]:[Price]])", "=SUM(tblSales[[Qty]:[Price]])"},
    {"=SUM(tblOther[Val])", "=SUM(tblOther[Val])", "=SUM(tblOther[Val])"},
    // Totals-row cells lie inside the table's range.
    {"=SUBTOTAL(109,tblSales[Qty])", "=SUBTOTAL(109,[Qty])", "=SUBTOTAL(109,tblSales[Qty])"},
    {"=tblSales[[#Totals],[Qty]]+1", "=tblSales[[#Totals],[Qty]]+1", "=tblSales[[#Totals],[Qty]]+1"},
};

TEST(StructuredRefSpellingDisplay, D1InsideOutsideOtherSheet) {
  for (const D1Row& row : kD1) {
    SCOPED_TRACE(row.stored);
    EXPECT_EQ(display_structured_refs(row.stored, "tblSales"), row.inside);
    // Outside the table and on another sheet no table holds the cell.
    EXPECT_EQ(display_structured_refs(row.stored, ""), row.outside);
  }
}

// D2: one name per row, read at the four positions.
struct D2Row {
  const char* name;
  const char* in_row;
  const char* in_col;
  const char* out_row;
  const char* out_col;
};

const D2Row kD2[] = {
    {"ab", "[@ab]", "[ab]", "t[@ab]", "t[ab]"},
    {"1a", "[@1a]", "[1a]", "t[@1a]", "t[1a]"},
    {"a1", "[@a1]", "[a1]", "t[@a1]", "t[a1]"},
    {"日本", "[@日本]", "[日本]", "t[@日本]", "t[日本]"},
    {"a|b", "[@a|b]", "[a|b]", "t[@a|b]", "t[a|b]"},
    {"a b", "[@[a b]]", "[a b]", "t[@[a b]]", "t[a b]"},
    {"a-b", "[@[a-b]]", "[a-b]", "t[@[a-b]]", "t[a-b]"},
    {"a.b", "[@[a.b]]", "[a.b]", "t[@[a.b]]", "t[a.b]"},
    {"a_b", "[@[a_b]]", "[a_b]", "t[@[a_b]]", "t[a_b]"},
    {"a'b", "[@[a''b]]", "[a''b]", "t[@[a''b]]", "t[a''b]"},
    {"a[b", "[@[a'[b]]", "[a'[b]", "t[@[a'[b]]", "t[a'[b]"},
    {"a]b", "[@[a']b]]", "[a']b]", "t[@[a']b]]", "t[a']b]"},
    {"a#b", "[@[a'#b]]", "[a'#b]", "t[@[a'#b]]", "t[a'#b]"},
    {"a@b", "[@[a'@b]]", "[a'@b]", "t[@[a'@b]]", "t[a'@b]"},
    {"a,b", "[@[a,b]]", "[a,b]", "t[@[a,b]]", "t[a,b]"},
    {"a:b", "[@[a:b]]", "[a:b]", "t[@[a:b]]", "t[a:b]"},
    {"a+b", "[@[a+b]]", "[a+b]", "t[@[a+b]]", "t[a+b]"},
    {"a/b", "[@[a/b]]", "[a/b]", "t[@[a/b]]", "t[a/b]"},
    {"a$b", "[@[a$b]]", "[a$b]", "t[@[a$b]]", "t[a$b]"},
    {"a&b", "[@[a&b]]", "[a&b]", "t[@[a&b]]", "t[a&b]"},
    {"a!b", "[@[a!b]]", "[a!b]", "t[@[a!b]]", "t[a!b]"},
    {"a\"b", "[@[a\"b]]", "[a\"b]", "t[@[a\"b]]", "t[a\"b]"},
    {"a(b", "[@[a(b]]", "[a(b]", "t[@[a(b]]", "t[a(b]"},
    {"a;b", "[@[a;b]]", "[a;b]", "t[@[a;b]]", "t[a;b]"},
    {"a=b", "[@[a=b]]", "[a=b]", "t[@[a=b]]", "t[a=b]"},
    {"a?b", "[@[a?b]]", "[a?b]", "t[@[a?b]]", "t[a?b]"},
    {"a{b", "[@[a{b]]", "[a{b]", "t[@[a{b]]", "t[a{b]"},
    {"a~b", "[@[a~b]]", "[a~b]", "t[@[a~b]]", "t[a~b]"},
    {"a%b", "[@[a%b]]", "[a%b]", "t[@[a%b]]", "t[a%b]"},
    {"a\\b", "[@[a\\b]]", "[a\\b]", "t[@[a\\b]]", "t[a\\b]"},
    {"a*b", "[@[a*b]]", "[a*b]", "t[@[a*b]]", "t[a*b]"},
    {"a^b", "[@[a^b]]", "[a^b]", "t[@[a^b]]", "t[a^b]"},
    {"a<b", "[@[a<b]]", "[a<b]", "t[@[a<b]]", "t[a<b]"},
    {"a`b", "[@[a`b]]", "[a`b]", "t[@[a`b]]", "t[a`b]"},
    {" ab", "[@[ ab]]", "[[ ab]]", "t[@[ ab]]", "t[[ ab]]"},
    {"ab ", "[@[ab ]]", "[[ab ]]", "t[@[ab ]]", "t[[ab ]]"},
};

StructuredRefPayload ThisRow(const char* name) {
  return StructuredRefPayload{Specs::kThisRow, {name}};
}

StructuredRefPayload Column(const char* name) {
  return StructuredRefPayload{0U, {name}};
}

TEST(StructuredRefPayloadDisplay, D2ColumnNames) {
  for (const D2Row& row : kD2) {
    SCOPED_TRACE(row.name);
    EXPECT_EQ(display_structured_ref("t", ThisRow(row.name), true), row.in_row);
    EXPECT_EQ(display_structured_ref("t", Column(row.name), true), row.in_col);
    EXPECT_EQ(display_structured_ref("t", ThisRow(row.name), false), row.out_row);
    EXPECT_EQ(display_structured_ref("t", Column(row.name), false), row.out_col);
  }
}

// The same names spliced from stored formula text.
TEST(StructuredRefSpellingDisplay, D2FromStoredText) {
  for (const D2Row& row : kD2) {
    SCOPED_TRACE(row.name);
    const std::string col = stored_payload(Column(row.name));
    const std::string this_row = stored_payload(ThisRow(row.name));
    const std::string col_text = "=t[" + col + "]";
    const std::string row_text = "=t[" + this_row + "]";
    EXPECT_EQ(display_structured_refs(row_text, "t"), std::string("=") + row.in_row);
    EXPECT_EQ(display_structured_refs(col_text, "t"), std::string("=") + row.in_col);
    EXPECT_EQ(display_structured_refs(row_text, ""), std::string("=") + row.out_row);
    EXPECT_EQ(display_structured_refs(col_text, ""), std::string("=") + row.out_col);
  }
}

TEST(StructuredRefSpellingDisplay, D3OtherContexts) {
  // A table with one data row.
  EXPECT_EQ(display_structured_refs("=tblOne[[#This Row],[q]]", "tblOne"), "=[@q]");
  EXPECT_EQ(display_structured_refs("=tblOne[[#This Row],[q]]", ""), "=tblOne[@q]");
  EXPECT_EQ(display_structured_refs("=SUM(tblOne[q])", ""), "=SUM(tblOne[q])");
  // Another table's this-row reference from inside tblSales.
  EXPECT_EQ(display_structured_refs("=tblOther[[#This Row],[Val]]", "tblSales"), "=tblOther[@Val]");
  // Defined names read with the outside rule.
  EXPECT_EQ(display_structured_refs("=tblSales[[#This Row],[Qty]]", ""), "=tblSales[@Qty]");
  EXPECT_EQ(display_structured_refs("=tblSales[[#Headers],[Qty]]", ""), "=tblSales[[#Headers],[Qty]]");
  EXPECT_EQ(display_structured_refs("=tblSales[Qty]", ""), "=tblSales[Qty]");
  EXPECT_EQ(display_structured_refs("=tblSales[[#Totals],[Price]]", ""), "=tblSales[[#Totals],[Price]]");
}

TEST(StructuredRefSpellingDisplay, NonAsciiSplice) {
  EXPECT_EQ(display_structured_refs("=\"😀日本\"&tblSales[[#This Row],[日本]]&\"é\"&SUM(tblSales[Qty])", "tblSales"),
            "=\"😀日本\"&[@日本]&\"é\"&SUM([Qty])");
  EXPECT_EQ(display_structured_refs("=(tblSales[[#This Row],[Qty]])*2", "tblSales"), "=([@Qty])*2");
}

TEST(StructuredRefSpellingDisplay, UnchangedText) {
  EXPECT_EQ(display_structured_refs("=SUM(A1:B2)", "tblSales"), "=SUM(A1:B2)");
  EXPECT_EQ(display_structured_refs("=SUM(", "tblSales"), "=SUM(");
  EXPECT_EQ(display_structured_refs("=tblSales[#Bogus]", "tblSales"), "=tblSales[#Bogus]");
}

// FormulaR1C1 matches Formula for every structured reference (D3).
TEST(StructuredRefSpellingDisplay, R1C1MatchesA1) {
  for (const D1Row& row : kD1) {
    for (const char* own : {"tblSales", ""}) {
      SCOPED_TRACE(std::string(row.stored) + " in " + own);
      Arena arena;
      const AstNode* root = parse_strict(std::string_view(row.stored).substr(1), arena);
      ASSERT_NE(root, nullptr);
      const StructuredRefDisplay display{own};
      EXPECT_EQ("=" + format_formula_r1c1(*root, 1, 5, &display), display_structured_refs(row.stored, own));
    }
  }
}

// --- Binding entered text (S3) ----------------------------------------------

TEST(StructuredRefSpellingBind, BareTableNameGainsBrackets) {
  EXPECT_EQ(bind_structured_refs("=ROWS(tblSales)", "", kTables), "=ROWS(tblSales[])");
  EXPECT_EQ(bind_structured_refs("=ROWS(TBLSALES)+tblOther", "", kTables), "=ROWS(TBLSALES[])+tblOther[]");
  EXPECT_EQ(bind_structured_refs("=ROWS(Rate)", "", kTables), "=ROWS(Rate)");
  EXPECT_EQ(bind_structured_refs("=Sheet1!tblSales", "", kTables), "=Sheet1!tblSales");
  EXPECT_EQ(bind_structured_refs("=LET(tblSales,1,tblSales+1)", "", kTables), "=LET(tblSales,1,tblSales+1)");
  EXPECT_EQ(bind_structured_refs("=LAMBDA(tblSales,tblSales)(1)", "", kTables), "=LAMBDA(tblSales,tblSales)(1)");
  EXPECT_EQ(bind_structured_refs("=ROWS(tblSales)", "", BareTableNamePredicate{nullptr, nullptr}), "=ROWS(tblSales)");
}

TEST(StructuredRefSpellingBind, ThisRowAndBareForms) {
  EXPECT_EQ(bind_structured_refs("=[@Qty]", "tblSales", kTables), "=tblSales[[#This Row],[Qty]]");
  EXPECT_EQ(bind_structured_refs("=SUM([Qty])", "tblSales", kTables), "=SUM(tblSales[Qty])");
  EXPECT_EQ(bind_structured_refs("=tblSales[@Qty]", "", kTables), "=tblSales[[#This Row],[Qty]]");
  EXPECT_EQ(bind_structured_refs("=tblSales[@]", "", kTables), "=tblSales[#This Row]");
  // Payloads without `@` stay as entered.
  EXPECT_EQ(bind_structured_refs("=SUM(tblSales[ Qty ])", "", kTables), "=SUM(tblSales[ Qty ])");
}

TEST(StructuredRefSpellingBind, BareBracketOutsideTableFails) {
  Arena arena;
  EXPECT_EQ(parse_strict("[Qty]", arena), nullptr);
  Parser parser("[Qty]", arena);
  parser.parse();
  ASSERT_FALSE(parser.errors().empty());
  EXPECT_EQ(parser.errors().front().code, ParseErrorCode::UnsupportedConstruct);
  EXPECT_EQ(bind_structured_refs("=[Qty]", "", kTables), "=[Qty]");
}

TEST(StructuredRefSpellingBind, BareArmBuildsStructuredRef) {
  Arena arena;
  ParserOptions opts;
  opts.bare_structured_ref_table = "tblSales";
  const AstNode* root = parse_strict("SUM([[#Headers],[Unit Cost]])", arena, opts);
  ASSERT_NE(root, nullptr);
  ASSERT_EQ(root->kind(), NodeKind::Call);
  const AstNode& arg = root->as_call_arg(0);
  ASSERT_EQ(arg.kind(), NodeKind::StructuredRef);
  EXPECT_EQ(arg.as_structured_ref_table(), "tblSales");
  EXPECT_EQ(arg.as_structured_ref_column(), "[#Headers],[Unit Cost]");
  EXPECT_EQ(arg.range().start, 4U);
  // `[0]!Name` keeps its meaning.
  const AstNode* self = parse_strict("[0]!Rate", arena, opts);
  ASSERT_NE(self, nullptr);
  EXPECT_EQ(self->kind(), NodeKind::ExternalRef);
}

// Display then bind gives back what Excel stores.
TEST(StructuredRefSpellingBind, RoundTrip) {
  for (const D1Row& row : kD1) {
    SCOPED_TRACE(row.stored);
    EXPECT_EQ(bind_structured_refs(display_structured_refs(row.stored, "tblSales"), "tblSales", kTables), row.stored);
    EXPECT_EQ(bind_structured_refs(display_structured_refs(row.stored, ""), "", kTables), row.stored);
  }
  for (const D2Row& row : kD2) {
    SCOPED_TRACE(row.name);
    for (const StructuredRefPayload& payload : {ThisRow(row.name), Column(row.name)}) {
      const std::string stored = "=tblSales[" + stored_payload(payload) + "]";
      EXPECT_EQ(bind_structured_refs(display_structured_refs(stored, "tblSales"), "tblSales", kTables), stored);
    }
  }
}

}  // namespace
}  // namespace parser
}  // namespace formulon
