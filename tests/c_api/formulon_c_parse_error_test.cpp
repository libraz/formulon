// fm_get_last_parse_error over tools/api/parse_error_cases.json: every row of
// an entry point this suite drives fails with the row's status, records one of
// its codes, names its field in the context, and, when the row pins one, the
// position in the caller's text.

#include <cstddef>
#include <cstdint>
#include <map>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "support/json_cases.h"

namespace {

using formulon::test_support::Json;

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

bool LoadCases(Json* out) {
  return formulon::test_support::LoadRepoJson(__FILE__, "tests/c_api/", "tools/api/parse_error_cases.json", out);
}

const std::map<std::string, int32_t>& CodeValues() {
  static const std::map<std::string, int32_t> kCodes = {
      {"LexerInvalidCharacter", FM_PARSE_ERROR_LEXER_INVALID_CHARACTER},
      {"LexerUnterminatedString", FM_PARSE_ERROR_LEXER_UNTERMINATED_STRING},
      {"LexerUnterminatedSheetQuote", FM_PARSE_ERROR_LEXER_UNTERMINATED_SHEET_QUOTE},
      {"LexerInvalidNumberLiteral", FM_PARSE_ERROR_LEXER_INVALID_NUMBER_LITERAL},
      {"LexerInvalidErrorLiteral", FM_PARSE_ERROR_LEXER_INVALID_ERROR_LITERAL},
      {"LexerInvalidEscape", FM_PARSE_ERROR_LEXER_INVALID_ESCAPE},
      {"LexerInvalidReference", FM_PARSE_ERROR_LEXER_INVALID_REFERENCE},
      {"LexerExcessiveLength", FM_PARSE_ERROR_LEXER_EXCESSIVE_LENGTH},
      {"UnexpectedToken", FM_PARSE_ERROR_UNEXPECTED_TOKEN},
      {"UnexpectedEof", FM_PARSE_ERROR_UNEXPECTED_EOF},
      {"ExpectedExpression", FM_PARSE_ERROR_EXPECTED_EXPRESSION},
      {"ExpectedCloseParen", FM_PARSE_ERROR_EXPECTED_CLOSE_PAREN},
      {"UnbalancedBraces", FM_PARSE_ERROR_UNBALANCED_BRACES},
      {"ArrayRowMismatch", FM_PARSE_ERROR_ARRAY_ROW_MISMATCH},
      {"ExpectedRParenOrComma", FM_PARSE_ERROR_EXPECTED_RPAREN_OR_COMMA},
      {"ExpectedCommaOrSemiInArray", FM_PARSE_ERROR_EXPECTED_COMMA_OR_SEMI_IN_ARRAY},
      {"InvalidReference", FM_PARSE_ERROR_INVALID_REFERENCE},
      {"UnsupportedConstruct", FM_PARSE_ERROR_UNSUPPORTED_CONSTRUCT},
      {"ExpectedOpenParen", FM_PARSE_ERROR_EXPECTED_OPEN_PAREN},
      {"ExpectedComma", FM_PARSE_ERROR_EXPECTED_COMMA},
      {"UnbalancedBrackets", FM_PARSE_ERROR_UNBALANCED_BRACKETS},
      {"InvalidRange", FM_PARSE_ERROR_INVALID_RANGE},
      {"NestedFormulaTooDeep", FM_PARSE_ERROR_NESTED_FORMULA_TOO_DEEP},
      {"TooManyErrors", FM_PARSE_ERROR_TOO_MANY_ERRORS},
      {"LetInvalidName", FM_PARSE_ERROR_LET_INVALID_NAME},
      {"LetWrongArity", FM_PARSE_ERROR_LET_WRONG_ARITY},
      {"LambdaInvalidParam", FM_PARSE_ERROR_LAMBDA_INVALID_PARAM},
      {"LambdaEmpty", FM_PARSE_ERROR_LAMBDA_EMPTY},
      {"LambdaDuplicateParam", FM_PARSE_ERROR_LAMBDA_DUPLICATE_PARAM},
      {"EntryWrongArity", FM_PARSE_ERROR_ENTRY_WRONG_ARITY},
      {"EntryReferenceRequired", FM_PARSE_ERROR_ENTRY_REFERENCE_REQUIRED},
      {"EntryQualifiedBuiltin", FM_PARSE_ERROR_ENTRY_QUALIFIED_BUILTIN},
      {"EntryRefusedWhitespace", FM_PARSE_ERROR_ENTRY_REFUSED_WHITESPACE},
      {"EntryParamShapedName", FM_PARSE_ERROR_ENTRY_PARAM_SHAPED_NAME},
      {"EntryDoubleEquals", FM_PARSE_ERROR_ENTRY_DOUBLE_EQUALS},
      {"EntryBomNotSupported", FM_PARSE_ERROR_ENTRY_BOM_NOT_SUPPORTED},
      {"StructuredRefInvalid", FM_PARSE_ERROR_STRUCTURED_REF_INVALID},
  };
  return kCodes;
}

// The row's formula: `formula`, or the concatenation of `formula_repeat`.
std::string FormulaOf(const Json& row) {
  const Json* repeat = row.Find("formula_repeat");
  if (repeat == nullptr || repeat->kind != Json::Kind::kArray) {
    return row.At("formula").text;
  }
  std::string out;
  for (const Json& segment : repeat->items) {
    const auto count = static_cast<std::size_t>(segment.items[1].number);
    for (std::size_t i = 0; i < count; ++i) {
      out += segment.items[0].text;
    }
  }
  return out;
}

// Calls the entry point `entry` names with `formula`; false when this suite
// does not drive that entry.
bool CallEntry(fm_workbook_t* wb, const std::string& entry, const std::string& formula, fm_status_t* status) {
  const char* text = formula.c_str();
  if (entry == "set_formula") {
    *status = fm_sheet_set_formula(wb, 0, 0, 0, text);
  } else if (entry == "set_array_formula") {
    const fm_cell_rect_t range{0, 0, 1, 1};
    *status = fm_sheet_set_array_formula(wb, 0, &range, text);
  } else if (entry == "evaluate_formula") {
    fm_value_t v{};
    *status = fm_sheet_evaluate_formula(wb, 0, 0, 0, text, &v);
  } else if (entry == "evaluate_formula_array") {
    fm_value_array_t* array = nullptr;
    *status = fm_sheet_evaluate_formula_array(wb, 0, 0, 0, text, &array);
    fm_value_array_destroy(array);
  } else if (entry == "evaluate_cf_formula") {
    fm_value_t v{};
    *status = fm_sheet_evaluate_cf_formula(wb, 0, 0, 0, 0, 0, text, &v);
  } else if (entry == "defined_name") {
    fm_defined_name_input_t input = FM_STRUCT_INIT(fm_defined_name_input_t);
    input.name = "Rate";
    input.formula = text;
    input.local_sheet_id = -1;
    *status = fm_workbook_set_defined_name(wb, &input);
  } else if (entry == "parse_reference") {
    fm_parsed_reference_t out = FM_STRUCT_INIT(fm_parsed_reference_t);
    *status = fm_workbook_parse_reference(wb, text, 0, 0, 0, 0, &out);
  } else {
    return false;
  }
  return true;
}

TEST(FormulonCApiParseError, CaseTable) {
  Json table;
  ASSERT_TRUE(LoadCases(&table));
  std::size_t driven = 0;
  for (const Json& row : table.At("cases").items) {
    const std::string id = row.At("id").text;
    SCOPED_TRACE(id);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    fm_status_t status = 0;
    if (!CallEntry(wb.handle, row.At("entry").text, FormulaOf(row), &status)) {
      continue;
    }
    ++driven;
    EXPECT_EQ(status, static_cast<fm_status_t>(row.At("status").number));
    // The context is read first: the parse record is the one read that leaves it alone.
    const std::string context = fm_get_last_error_context();

    fm_parse_error_t error = FM_STRUCT_INIT(fm_parse_error_t);
    const fm_status_t read = fm_get_last_parse_error(&error);
    EXPECT_EQ(read, 0);
    if (read != 0) {
      continue;
    }
    bool code_listed = false;
    std::string listed;
    for (const Json& name : row.At("codes").items) {
      const auto found = CodeValues().find(name.text);
      ASSERT_NE(found, CodeValues().end()) << name.text;
      code_listed = code_listed || found->second == error.code;
      listed += name.text + " ";
    }
    EXPECT_TRUE(code_listed) << "code " << error.code << " not in " << listed;

    if (const Json* field = row.Find("field"); field != nullptr && field->kind == Json::Kind::kString) {
      EXPECT_NE(context.find("field=" + field->text), std::string::npos) << context;
    }
    if (row.At("start").kind == Json::Kind::kNumber) {
      EXPECT_EQ(error.start, static_cast<uint32_t>(row.At("start").number));
      EXPECT_EQ(error.end, static_cast<uint32_t>(row.At("end").number));
      EXPECT_EQ(error.line, static_cast<uint32_t>(row.At("line").number));
      EXPECT_EQ(error.column, static_cast<uint32_t>(row.At("column").number));
    }
  }
  EXPECT_GT(driven, 0U);
}

TEST(FormulonCApiParseError, RecordIsClearedByTheNextCall) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_NE(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=(1+2"), 0);
  fm_parse_error_t error = FM_STRUCT_INIT(fm_parse_error_t);
  ASSERT_EQ(fm_get_last_parse_error(&error), 0);
  EXPECT_EQ(error.code, FM_PARSE_ERROR_EXPECTED_CLOSE_PAREN);

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=1+2"), 0);
  EXPECT_EQ(fm_get_last_parse_error(&error), FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiParseError, ReadersLeaveEachOtherUntouched) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 3, "=(1+2"), FM_STATUS_PARSER_MISMATCHED_BRACKETS);
  const std::string message = fm_get_last_error_message();
  const std::string context = fm_get_last_error_context();
  EXPECT_EQ(context, "row=2 col=3");

  fm_parse_error_t first = FM_STRUCT_INIT(fm_parse_error_t);
  ASSERT_EQ(fm_get_last_parse_error(&first), 0);
  // A record too small to hold the answer fails without touching any state.
  fm_parse_error_t small = FM_STRUCT_INIT(fm_parse_error_t);
  small.struct_size = 4;
  EXPECT_EQ(fm_get_last_parse_error(&small), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_get_last_parse_error(nullptr), FM_STATUS_BINDING_NULL_POINTER);

  EXPECT_EQ(message, fm_get_last_error_message());
  EXPECT_EQ(context, fm_get_last_error_context());
  fm_parse_error_t second = FM_STRUCT_INIT(fm_parse_error_t);
  ASSERT_EQ(fm_get_last_parse_error(&second), 0);
  EXPECT_EQ(second.code, first.code);
  EXPECT_EQ(second.start, first.start);
  EXPECT_EQ(second.end, first.end);
  EXPECT_EQ(second.struct_size, sizeof(fm_parse_error_t));
}

TEST(FormulonCApiParseError, FailedEntryLeavesTheCellUntouched) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 7.0), 0);
  ASSERT_NE(fm_sheet_set_formula(wb.handle, 0, 0, 0, "==1"), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_EQ(v.u.number, 7.0);
}

}  // namespace
