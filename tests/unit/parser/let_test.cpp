//
// Parser tests for the `LET` special form. The Pratt parser recognises
// `LET(name, expr, [name, expr, ...], body)` and emits a `LetBinding` AST
// node; binding-name slots accept bare identifiers that are NOT resolved as
// cell references, and odd-arity plus name-shape validation runs at parse
// time.

#include <algorithm>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/parse_error.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

bool HasErrorCode(const std::vector<ParseError>& errors, ParseErrorCode code) {
  return std::any_of(errors.begin(), errors.end(), [code](const ParseError& e) { return e.code == code; });
}

std::string ParseToSexpr(std::string_view src) {
  Arena a;
  Parser p(src, a);
  AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  EXPECT_TRUE(p.errors().empty()) << "unexpected errors for: " << src;
  if (root == nullptr) {
    return "<null>";
  }
  return dump_sexpr(*root);
}

// ---------------------------------------------------------------------------
// Well-formed LET shapes
// ---------------------------------------------------------------------------

TEST(ParserLet, SingleBindingScalarBody) {
  EXPECT_EQ(ParseToSexpr("=LET(x, 1, x+2)"), "(let ((x (num 1))) (binary + (name x) (num 2)))");
}

TEST(ParserLet, TwoBindingsSequential) {
  EXPECT_EQ(ParseToSexpr("=LET(x, 1, y, x+1, x*y)"),
            "(let ((x (num 1)) (y (binary + (name x) (num 1)))) (binary * (name x) (name y)))");
}

TEST(ParserLet, LowerCaseKeyword) {
  // LET is case-insensitive at the call-name level, matching all other
  // Excel function dispatch.
  EXPECT_EQ(ParseToSexpr("=let(x, 5, x)"), "(let ((x (num 5))) (name x))");
}

TEST(ParserLet, XlfnStoragePrefixIsRecognizedAsSpecialForm) {
  // Excel's name manager stores a reusable LET as `_xlfn.LET(...)` so older
  // Excel versions do not misinterpret it. `parser::strip_storage_prefixes`
  // canonicalises `formula_text` at every known ingestion point, but the
  // parser must also recognise the prefixed spelling directly: without
  // this, a caller handing the parser un-normalised text would get an
  // ordinary (and wrong) function-call parse instead of LET's dedicated
  // binding-name grammar.
  EXPECT_EQ(ParseToSexpr("=_xlfn.LET(x, 1, x+2)"), "(let ((x (num 1))) (binary + (name x) (num 2)))");
  EXPECT_EQ(ParseToSexpr("=_XLFN.let(x, 1, x+2)"), "(let ((x (num 1))) (binary + (name x) (num 2)))");
}

TEST(ParserLet, XlfnStoragePrefixIsRecognizedForLambda) {
  // Same mechanism as the LET case above, for `_xlfn.LAMBDA(...)`.
  EXPECT_EQ(ParseToSexpr("=_xlfn.LAMBDA(x, x+1)"), "(lambda (x) (binary + (name x) (num 1)))");
}

TEST(ParserLet, NestedLetInBindingValue) {
  EXPECT_EQ(ParseToSexpr("=LET(x, LET(y, 2, y*y), x+1)"),
            "(let ((x (let ((y (num 2))) (binary * (name y) (name y))))) (binary + (name x) (num 1)))");
}

TEST(ParserLet, ShadowingRedefinesName) {
  // Inner `x` binding is the second (name, expr) pair; the trailing `x`
  // in the body reads the shadowed definition.
  EXPECT_EQ(ParseToSexpr("=LET(x, 1, x, x+1, x)"), "(let ((x (num 1)) (x (binary + (name x) (num 1)))) (name x))");
}

TEST(ParserLet, BindingNameWithUnderscore) {
  EXPECT_EQ(ParseToSexpr("=LET(_foo, 3, _foo)"), "(let ((_foo (num 3))) (name _foo))");
}

TEST(ParserLet, BindingNameWithBackslashStart) {
  // `\` is Excel's third legal name-manager start character, alongside a
  // letter and `_`; `IsLetNameShape` mirrors the tokenizer's identifier
  // rule so a LET binding admits it too.
  EXPECT_EQ(ParseToSexpr("=LET(\\foo, 3, \\foo)"), "(let ((\\foo (num 3))) (name \\foo))");
}

TEST(ParserLet, BindingNameWithDigitsAndDot) {
  // Allowed shape: leading letter, subsequent letters/digits/underscore/
  // period. `a1_test.v2` has trailing non-digit characters so is NOT an
  // A1 cell ref and parses as a LET binding name.
  EXPECT_EQ(ParseToSexpr("=LET(a1_test.v2, 3, a1_test.v2)"), "(let ((a1_test.v2 (num 3))) (name a1_test.v2))");
}

TEST(ParserLet, FunctionNameAsBindingAllowed) {
  // Excel allows shadowing a builtin inside the LET scope; the parser
  // treats it as an ordinary binding name.
  EXPECT_EQ(ParseToSexpr("=LET(SUM, 1, SUM)"), "(let ((SUM (num 1))) (name SUM))");
}

TEST(ParserLet, BodyReferencesCellAndName) {
  // Mixed body: cell ref stays a Ref, bound identifier stays a NameRef.
  EXPECT_EQ(ParseToSexpr("=LET(n, A1, n*2)"), "(let ((n (ref A1))) (binary * (name n) (num 2)))");
}

// ---------------------------------------------------------------------------
// Name-validation errors
// ---------------------------------------------------------------------------

TEST(ParserLet, InvalidNameA1IsRejected) {
  Arena a;
  Parser p("=LET(A1, 1, A1)", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LetInvalidName))
      << "LET with cell-ref-shaped binding name should emit LetInvalidName";
}

TEST(ParserLet, InvalidNameStartingWithDigitIsRejected) {
  Arena a;
  Parser p("=LET(1x, 1, 1x)", a);
  (void)p.parse();
  // The tokenizer splits `1x` into Number + Ident, so the parser sees a
  // Number in the binding-name slot; the resulting diagnostic is a
  // name-shape / shape error but need not specifically be LetInvalidName.
  // We assert only that the parse reports at least one error.
  EXPECT_FALSE(p.errors().empty());
}

TEST(ParserLet, InvalidNameMultiLetterCellRefIsRejected) {
  Arena a;
  Parser p("=LET(AA10, 1, AA10)", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LetInvalidName));
}

TEST(ParserLet, ZeroPaddedRowCellRefIsRejected) {
  // The tokenizer reads `A01` as a cell reference, so it is not a legal name.
  Arena a;
  Parser p("=LET(A01, 1, A01)", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LetInvalidName));
}

TEST(ParserLet, OutOfGridColumnIsAcceptedAsName) {
  // XFE and ZZZ lie past the last column (XFD), so the tokenizer reads these as names.
  EXPECT_EQ(ParseToSexpr("=LET(XFE1, 1, XFE1)"), "(let ((XFE1 (num 1))) (name XFE1))");
  EXPECT_EQ(ParseToSexpr("=LET(ZZZ1, 1, ZZZ1)"), "(let ((ZZZ1 (num 1))) (name ZZZ1))");
}

TEST(ParserLet, OutOfGridRowIsAcceptedAsName) {
  EXPECT_EQ(ParseToSexpr("=LET(A1048577, 1, A1048577)"), "(let ((A1048577 (num 1))) (name A1048577))");
}

TEST(ParserLet, LastGridCellIsRejected) {
  Arena a;
  Parser p("=LET(XFD1048576, 1, XFD1048576)", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LetInvalidName));
}

// ---------------------------------------------------------------------------
// LAMBDA parameter-name validation (same grid bound as LET binding names)
// ---------------------------------------------------------------------------

TEST(ParserLet, LambdaOutOfGridParamIsAcceptedAsName) {
  EXPECT_EQ(ParseToSexpr("=LAMBDA(XFE1, XFE1)"), "(lambda (XFE1) (name XFE1))");
  EXPECT_EQ(ParseToSexpr("=LAMBDA(ZZZ1, ZZZ1)"), "(lambda (ZZZ1) (name ZZZ1))");
  EXPECT_EQ(ParseToSexpr("=LAMBDA(A1048577, A1048577)"), "(lambda (A1048577) (name A1048577))");
}

TEST(ParserLet, LambdaZeroPaddedRowParamIsRejected) {
  Arena a;
  Parser p("=LAMBDA(A01, A01)", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LambdaInvalidParam));
}

TEST(ParserLet, LambdaLastGridCellParamIsRejected) {
  Arena a;
  Parser p("=LAMBDA(XFD1048576, 1)", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LambdaInvalidParam));
}

// ---------------------------------------------------------------------------
// Arity errors
// ---------------------------------------------------------------------------

TEST(ParserLet, MissingBodyIsArityError) {
  // LET(x, 1) — 2 args, even. No body.
  Arena a;
  Parser p("=LET(x, 1)", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LetWrongArity));
}

TEST(ParserLet, SingleArgIsArityError) {
  Arena a;
  Parser p("=LET(x)", a);
  (void)p.parse();
  // Either an arity or a name-shape diagnostic; main contract is "not OK".
  EXPECT_FALSE(p.errors().empty());
}

TEST(ParserLet, EmptyArgListIsArityError) {
  Arena a;
  Parser p("=LET()", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LetWrongArity));
}

TEST(ParserLet, EvenArityWithTrailingCommaIsArityError) {
  // Four args: (name, expr, name, expr) with no body.
  Arena a;
  Parser p("=LET(x, 1, y, 2)", a);
  (void)p.parse();
  EXPECT_TRUE(HasErrorCode(p.errors(), ParseErrorCode::LetWrongArity));
}

TEST(ParserLet, BindingNameCapCountsUtf16Units) {
  // 86 kanji are 258 UTF-8 bytes but 86 UTF-16 units; 255 ASCII letters fit, 256 do not.
  std::string kanji;
  for (int k = 0; k < 86; ++k) {
    kanji += "\xE6\xBC\xA2";
  }
  Arena a;
  const std::string kanji_src = "=LET(" + kanji + ",1," + kanji + ")";
  Parser p(kanji_src, a);
  EXPECT_NE(p.parse(), nullptr);
  EXPECT_TRUE(p.errors().empty());

  const std::string fits(255, 'a');
  const std::string too_long(256, 'a');
  Arena a2;
  const std::string fits_src = "=LET(" + fits + ",1," + fits + ")";
  Parser fits_parser(fits_src, a2);
  (void)fits_parser.parse();
  EXPECT_TRUE(fits_parser.errors().empty());
  Arena a3;
  const std::string long_src = "=LET(" + too_long + ",1," + too_long + ")";
  Parser long_parser(long_src, a3);
  (void)long_parser.parse();
  EXPECT_FALSE(long_parser.errors().empty());
}

}  // namespace
}  // namespace parser
}  // namespace formulon
