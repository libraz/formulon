//
// Names whose rest is an A1 cell (`B1`, `A1`) keep a prefix in stored text so
// they cannot read back as cells: `_xlpm.` for a LET / LAMBDA binding and its
// in-scope uses, `_xlnm.` for every other one. The stored forms below are the
// `<f>` text of Excel-saved workbooks; the displayed forms are Excel's
// FORMULATEXT.

#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_format_r1c1.h"
#include "parser/formula_prefix.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace parser {
namespace {

const AstNode* ParseOk(std::string_view src, Arena& arena) {
  Parser p(src, arena);
  const AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << src;
  EXPECT_TRUE(p.errors().empty()) << src;
  return root;
}

std::string Format(std::string_view src) {
  Arena arena;
  const AstNode* root = ParseOk(src, arena);
  return root == nullptr ? std::string("<null>") : format_formula(*root);
}

// The writer's speller for the functions these formulas use: only LET and
// LAMBDA are future functions.
std::string SpellFuture(std::string_view name) {
  if (name == "LET" || name == "LAMBDA") {
    return "_xlfn." + std::string(name);
  }
  return std::string(name);
}

std::string Store(std::string_view src) {
  Arena arena;
  const AstNode* root = ParseOk(src, arena);
  return root == nullptr ? std::string("<null>") : format_formula_storage(*root, &SpellFuture);
}

bool KnowsEverything(std::string_view /*name*/) {
  return true;
}

TEST(CellShapedNames, ColumnShape) {
  EXPECT_TRUE(is_a1_column_shaped("A"));
  EXPECT_TRUE(is_a1_column_shaped("xfd"));
  EXPECT_TRUE(is_a1_column_shaped("LE"));
  EXPECT_FALSE(is_a1_column_shaped("XFE"));
  EXPECT_FALSE(is_a1_column_shaped("ABCD"));
  EXPECT_FALSE(is_a1_column_shaped("A1"));
  EXPECT_FALSE(is_a1_column_shaped(""));
}

TEST(CellShapedNames, BareAndPrefix) {
  EXPECT_EQ(cell_shaped_name_bare("_xlpm.B1"), "B1");
  EXPECT_EQ(cell_shaped_name_prefix("_xlpm.B1"), "_xlpm.");
  EXPECT_EQ(cell_shaped_name_bare("_XLNM.b2"), "b2");
  EXPECT_EQ(cell_shaped_name_prefix("_XLNM.b2"), "_xlnm.");
  EXPECT_EQ(cell_shaped_name_bare("B1"), "B1");
  EXPECT_EQ(cell_shaped_name_prefix("B1"), "");
  // The rest is no cell: nothing is dropped.
  EXPECT_EQ(cell_shaped_name_bare("_xlpm.x"), "_xlpm.x");
  EXPECT_EQ(cell_shaped_name_prefix("_xlpm.x"), "");
  EXPECT_EQ(cell_shaped_name_bare("_xlnm.XFE1"), "_xlnm.XFE1");
  EXPECT_EQ(cell_shaped_name_bare("_xlnm._FilterDatabase"), "_xlnm._FilterDatabase");
  EXPECT_EQ(cell_shaped_name_bare("_xlfn.B1"), "_xlfn.B1");
}

TEST(CellShapedNames, StripKeepsParamPrefixBeforeCellShapedName) {
  EXPECT_EQ(strip_storage_prefixes("=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)"), "=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)");
  EXPECT_EQ(strip_storage_prefixes("=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)", &KnowsEverything),
            "=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)");
  EXPECT_EQ(strip_storage_prefixes("_xlfn.LET(_xlpm.B1,LAMBDA(x,x+1),_xlpm.B1(2))"),
            "LET(_xlpm.B1,LAMBDA(x,x+1),_xlpm.B1(2))");
  EXPECT_EQ(strip_storage_prefixes("_xlfn.LET(_xlpm.b1,5,_xlpm.b1)"), "LET(_xlpm.b1,5,_xlpm.b1)");
  EXPECT_EQ(strip_storage_prefixes("_xlfn.LET(_xlpm.B1,5,Sheet1!_xlnm.B1)"), "LET(_xlpm.B1,5,Sheet1!_xlnm.B1)");
}

TEST(CellShapedNames, StripStillDropsParamPrefixBeforeOtherNames) {
  EXPECT_EQ(strip_storage_prefixes("_xlfn.LET(_xlpm.x,1,_xlpm.x)+_xlfn.LET(_xlpm.B1,5,_xlpm.B1)"),
            "LET(x,1,x)+LET(_xlpm.B1,5,_xlpm.B1)");
  EXPECT_EQ(strip_storage_prefixes("_xlfn.LET(_xlpm.R1C1,3,_xlpm.R1C1)"), "LET(R1C1,3,R1C1)");
  EXPECT_EQ(strip_storage_prefixes("_xlfn.LET(_xlpm.R1C,2,_xlpm.R1C)"), "LET(R1C,2,R1C)");
  EXPECT_EQ(strip_storage_prefixes("LET(_xlpm.XFE1,1,_xlpm.XFE1)"), "LET(XFE1,1,XFE1)");
  EXPECT_EQ(strip_storage_prefixes("=\"_xlpm.B1\"&_xlpm.x"), "=\"_xlpm.B1\"&x");
}

// Each pair is the canonical text (the stored `<f>` without `_xlfn.`) and the
// `<f>` Excel saved for it.
struct StoredForm {
  const char* canonical;
  const char* stored;
};

constexpr StoredForm kExcelStoredForms[] = {
    {"LET(_xlpm.B1,5,_xlpm.B1+$A$1)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+$A$1)"},
    {"LET(_xlpm.B1,5,SUM(_xlpm.B1,$A:$A))", "_xlfn.LET(_xlpm.B1,5,SUM(_xlpm.B1,$A:$A))"},
    {"_xlnm.B2+LET(_xlpm.B1,5,_xlpm.B1)", "_xlnm.B2+_xlfn.LET(_xlpm.B1,5,_xlpm.B1)"},
    {"LET(_xlpm.B1,5,_xlpm.B1)+_xlnm.B2", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1)+_xlnm.B2"},
    {"LET(x,1,x)+LET(_xlpm.B1,5,_xlpm.B1)", "_xlfn.LET(_xlpm.x,1,_xlpm.x)+_xlfn.LET(_xlpm.B1,5,_xlpm.B1)"},
    {"LET(x,5,x+B2)", "_xlfn.LET(_xlpm.x,5,_xlpm.x+B2)"},
    {"LET(_xlpm.B1,100,LOG10(_xlpm.B1))", "_xlfn.LET(_xlpm.B1,100,LOG10(_xlpm.B1))"},
    {"LET(_xlpm.B1,5,_xlpm.B1+XFE1)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+XFE1)"},
    {"LAMBDA(_xlpm.B1,_xlpm.B1+$A$1)(4)", "_xlfn.LAMBDA(_xlpm.B1,_xlpm.B1+$A$1)(4)"},
    {"LET(_xlpm.B1,5,_xlpm.B1+A:(A))", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+A:(A))"},
    {"LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1)"},
    {"LET(_xlpm.A1,5,_xlpm.A1+$A$1)", "_xlfn.LET(_xlpm.A1,5,_xlpm.A1+$A$1)"},
    {"LET(_xlpm.B1,2,_xlpm.B1)*$A$1", "_xlfn.LET(_xlpm.B1,2,_xlpm.B1)*$A$1"},
    {"LET(R1C1,3,R1C1)", "_xlfn.LET(_xlpm.R1C1,3,_xlpm.R1C1)"},
    {"LET(_xlpm.B1,5,Sheet1!_xlnm.B1)", "_xlfn.LET(_xlpm.B1,5,Sheet1!_xlnm.B1)"},
    {"LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:_xlnm.B2)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:_xlnm.B2)"},
    {"LET(_xlpm.B1,5,_xlpm.B1+A)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+A)"},
    {"LET(_xlpm.B1,5,_xlpm.B1+AB)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+AB)"},
    {"LET(_xlpm.B1,5,_xlpm.B1+XFD1048577)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+XFD1048577)"},
    {"LET(R1C,2,R1C)", "_xlfn.LET(_xlpm.R1C,2,_xlpm.R1C)"},
    {"LET(_xlpm.B1,5,SUM(A:(B)))", "_xlfn.LET(_xlpm.B1,5,SUM(A:(B)))"},
    {"_xlnm.B2", "_xlnm.B2"},
    {"LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)"},
    {"LAMBDA(R1C,R1C)(3)", "_xlfn.LAMBDA(_xlpm.R1C,_xlpm.R1C)(3)"},
    {"LET(_xlpm.B1,5,_xlpm.B1+B1)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+B1)"},
    {"LET(_xlpm.B1,5,_xlpm.B1*0+B1)", "_xlfn.LET(_xlpm.B1,5,_xlpm.B1*0+B1)"},
};

TEST(CellShapedNames, ExcelStoredFormsReadToCanonicalText) {
  for (const StoredForm& form : kExcelStoredForms) {
    EXPECT_EQ(strip_storage_prefixes(form.stored), form.canonical) << form.stored;
  }
}

TEST(CellShapedNames, FormatIsIdempotentOnCanonicalText) {
  for (const StoredForm& form : kExcelStoredForms) {
    EXPECT_EQ(Format(form.canonical), form.canonical) << form.canonical;
  }
  // A binding called as a function, and a cell-shaped LAMBDA parameter called.
  EXPECT_EQ(Format("LET(_xlpm.B1,LAMBDA(x,x+1),_xlpm.B1(2))"), "LET(_xlpm.B1,LAMBDA(x,x+1),_xlpm.B1(2))");
  EXPECT_EQ(Format("LAMBDA(_xlpm.B1,_xlpm.B1(1))(LAMBDA(x,x*2))"), "LAMBDA(_xlpm.B1,_xlpm.B1(1))(LAMBDA(x,x*2))");
  EXPECT_EQ(Format("LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:(A))"), "LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:(A))");
}

TEST(CellShapedNames, StorageMatchesExcelWithoutDoublePrefix) {
  for (const StoredForm& form : kExcelStoredForms) {
    EXPECT_EQ(Store(form.canonical), form.stored) << form.canonical;
  }
  EXPECT_EQ(Store("LET(_xlpm.B1,LAMBDA(x,x+1),_xlpm.B1(2))"),
            "_xlfn.LET(_xlpm.B1,_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1),_xlpm.B1(2))");
  // A non-cell-shaped binding already spelled with the prefix gets no second one.
  EXPECT_EQ(Store("LET(_xlpm.x,1,_xlpm.x+1)"), "_xlfn.LET(_xlpm.x,1,_xlpm.x+1)");
  EXPECT_EQ(Store("LAMBDA([_xlpm.B1],_xlpm.B1)"), "_xlfn.LAMBDA([_xlpm.B1],_xlpm.B1)");
}

TEST(CellShapedNames, BindingScopeDecidesThePrefix) {
  // A cell-shaped name with no binding in scope is a workbook name.
  EXPECT_EQ(Format("LET(_xlpm.B1,5,_xlpm.B1+_xlpm.B2)"), "LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)");
  EXPECT_EQ(Format("LET(_xlpm.B1,5,_xlpm.B1)+_xlpm.B1"), "LET(_xlpm.B1,5,_xlpm.B1)+_xlnm.B1");
  // Scope matches ASCII case-insensitively; each use keeps its own spelling.
  EXPECT_EQ(Format("LET(_xlpm.b1,5,_xlpm.B1)"), "LET(_xlpm.b1,5,_xlpm.B1)");
  // `_xlnm.` stays a workbook name even where a binding of that name is in scope.
  EXPECT_EQ(Format("LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B1)"), "LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B1)");
  // A sheet-qualified name is never a binding.
  EXPECT_EQ(Format("LET(_xlpm.B1,5,Sheet1!_xlpm.B1)"), "LET(_xlpm.B1,5,Sheet1!_xlnm.B1)");
  // A call outside any binding of its name.
  EXPECT_EQ(Format("_xlpm.B1(2)"), "_xlnm.B1(2)");
  // Names that are no cell print as before.
  EXPECT_EQ(Format("LET(x,1,x)"), "LET(x,1,x)");
  EXPECT_EQ(Format("LET(_xlpm.x,1,_xlpm.x)"), "LET(_xlpm.x,1,_xlpm.x)");
}

TEST(CellShapedNames, BareNamesFromBinaryRecordsTakeTheirPrefix) {
  Arena a;
  // LET(B1,5,B1+B2) as decoded from binary tokens, prefixes already dropped.
  const std::string_view names[] = {"B1"};
  const AstNode* exprs[] = {make_literal(a, Value::number(5.0))};
  AstNode* body = make_binary_op(a, BinOp::Add, make_name_ref(a, "B1"), make_name_ref(a, "B2"));
  AstNode* let = make_let_binding(a, names, exprs, 1, body);
  ASSERT_NE(let, nullptr);
  EXPECT_EQ(format_formula(*let), "LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)");
  EXPECT_EQ(format_formula_storage(*let, &SpellFuture), "_xlfn.LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)");

  AstNode* qualified = make_let_binding(a, names, exprs, 1, make_sheet_name_ref(a, "Sheet1", "B1", false));
  ASSERT_NE(qualified, nullptr);
  EXPECT_EQ(format_formula(*qualified), "LET(_xlpm.B1,5,Sheet1!_xlnm.B1)");
  EXPECT_EQ(format_formula_storage(*qualified, &SpellFuture), "_xlfn.LET(_xlpm.B1,5,Sheet1!_xlnm.B1)");

  const AstNode* args[] = {make_literal(a, Value::number(2.0))};
  AstNode* called = make_let_binding(a, names, exprs, 1, make_call(a, "B1", args, 1));
  ASSERT_NE(called, nullptr);
  EXPECT_EQ(format_formula(*called), "LET(_xlpm.B1,5,_xlpm.B1(2))");
  EXPECT_EQ(format_formula_storage(*called, &SpellFuture), "_xlfn.LET(_xlpm.B1,5,_xlpm.B1(2))");

  // The built-in LOG10 is a function, not a name.
  const AstNode* log_args[] = {make_name_ref(a, "B1")};
  AstNode* log = make_let_binding(a, names, exprs, 1, make_call(a, "LOG10", log_args, 1));
  ASSERT_NE(log, nullptr);
  EXPECT_EQ(format_formula(*log), "LET(_xlpm.B1,5,LOG10(_xlpm.B1))");
}

TEST(CellShapedNames, ColumnShapedNameRightOfANameTakesParens) {
  Arena a;
  AstNode* cols = make_range_op(a, make_name_ref(a, "A"), make_name_ref(a, "B"), /*spliced=*/false);
  ASSERT_NE(cols, nullptr);
  EXPECT_EQ(format_formula(*cols), "A:(B)");
  AstNode* cell_col = make_range_op(a, make_name_ref(a, "A1"), make_name_ref(a, "A"), /*spliced=*/false);
  ASSERT_NE(cell_col, nullptr);
  EXPECT_EQ(format_formula(*cell_col), "_xlnm.A1:(A)");
  AstNode* cells = make_range_op(a, make_name_ref(a, "A1"), make_name_ref(a, "B2"), /*spliced=*/false);
  ASSERT_NE(cells, nullptr);
  EXPECT_EQ(format_formula(*cells), "_xlnm.A1:_xlnm.B2");
  // A call on the right keeps the left endpoint's parentheses.
  EXPECT_EQ(Format("=LE:(NA())"), "(LE):(NA())");
}

TEST(CellShapedNames, FormatInScopeSeesOuterBindings) {
  Arena arena;
  const AstNode* body = ParseOk("_xlpm.B1+_xlpm.B2", arena);
  ASSERT_NE(body, nullptr);
  EXPECT_EQ(format_formula(*body), "_xlnm.B1+_xlnm.B2");
  const std::vector<std::string_view> prefixed = {"_xlpm.B1"};
  EXPECT_EQ(format_formula_in_scope(*body, prefixed), "_xlpm.B1+_xlnm.B2");
  const std::vector<std::string_view> bare = {"b1"};
  EXPECT_EQ(format_formula_in_scope(*body, bare), "_xlpm.B1+_xlnm.B2");
}

TEST(CellShapedNames, DisplayMatchesFormulaText) {
  struct Case {
    const char* stored;
    const char* shown;
  };
  constexpr Case kCases[] = {
      {"=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)", "=LET(B1,5,B1+'B2')"},
      {"=_xlnm.B2", "='B2'"},
      {"=_xlnm.B2+LET(_xlpm.B1,5,_xlpm.B1)", "='B2'+LET(B1,5,B1)"},
      {"=LET(_xlpm.B1,5,Sheet1!_xlnm.B1)", "=LET(B1,5,Sheet1!'B1')"},
      {"=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:_xlnm.B2)", "=LET(B1,5,B1+'A1':'B2')"},
      {"=LET(_xlpm.B1,5,SUM(Sheet1!_xlnm.A1:_xlnm.B2))", "=LET(B1,5,SUM(Sheet1!'A1':'B2'))"},
      {"=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:(A))", "=LET(B1,5,B1+'A1':(A))"},
      {"=LET(_xlpm.B1,LAMBDA(x,x+1),_xlpm.B1(2))", "=LET(B1,LAMBDA(x,x+1),B1(2))"},
      {"=LAMBDA(_xlpm.B1,_xlpm.B1(1))(LAMBDA(x,x*2))", "=LAMBDA(B1,B1(1))(LAMBDA(x,x*2))"},
      {"=LET(_xlpm.B1,5,_xlpm.B1+B1)", "=LET(B1,5,B1+B1)"},
      {"= LET( _xlpm.B1, 5, _xlpm.B1+_xlnm.B2 )", "= LET( B1, 5, B1+'B2' )"},
      {"=LET(x,5,x+B2)", "=LET(x,5,x+B2)"},
  };
  for (const Case& c : kCases) {
    EXPECT_EQ(display_cell_shaped_names(c.stored), c.shown) << c.stored;
  }
}

TEST(CellShapedNames, DisplayLeavesOtherTextAlone) {
  EXPECT_EQ(display_cell_shaped_names("=\"_xlnm.B2\"&_xlnm.B2"), "=\"_xlnm.B2\"&'B2'");
  EXPECT_EQ(display_cell_shaped_names("='_xlnm.B1'!A1+'It''s'!_xlnm.B1"), "='_xlnm.B1'!A1+'It''s'!'B1'");
  EXPECT_EQ(display_cell_shaped_names("=foo_xlnm.B1+_xlnm._FilterDatabase+_xlnm.XFE1+_xlpm.x"),
            "=foo_xlnm.B1+_xlnm._FilterDatabase+_xlnm.XFE1+_xlpm.x");
  EXPECT_EQ(display_cell_shaped_names("=Tbl[a'#b]+_xlnm.B2"), "=Tbl[a'#b]+'B2'");
  EXPECT_EQ(display_cell_shaped_names("=LAMBDA([_xlpm.B1],1)"), "=LAMBDA([B1],1)");
}

TEST(CellShapedNames, R1C1PrintsBareNames) {
  Arena a1;
  const AstNode* let = ParseOk("LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2)", a1);
  ASSERT_NE(let, nullptr);
  EXPECT_EQ(format_formula_r1c1(*let, 0, 0), "LET(B1,5,B1+B2)");
  Arena a2;
  const AstNode* qualified = ParseOk("LET(_xlpm.B1,5,Sheet1!_xlnm.B1)", a2);
  ASSERT_NE(qualified, nullptr);
  EXPECT_EQ(format_formula_r1c1(*qualified, 0, 0), "LET(B1,5,Sheet1!B1)");
  Arena a3;
  const AstNode* called = ParseOk("LET(_xlpm.B1,LAMBDA(x,x+1),_xlpm.B1(2))", a3);
  ASSERT_NE(called, nullptr);
  EXPECT_EQ(format_formula_r1c1(*called, 0, 0), "LET(B1,LAMBDA(x,x+1),B1(2))");
  Arena a4;
  const AstNode* lambda = ParseOk("LAMBDA(_xlpm.B1,_xlpm.B1+$A$1)(4)", a4);
  ASSERT_NE(lambda, nullptr);
  EXPECT_EQ(format_formula_r1c1(*lambda, 20, 4), "LAMBDA(B1,B1+R1C1)(4)");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
