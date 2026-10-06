//
// Default diagnostic messages for `ParseErrorCode`. Declared in
// `parser/parse_error.h`.

#include "parser/parse_error.h"

namespace formulon {
namespace parser {

const char* default_message(ParseErrorCode code) noexcept {
  switch (code) {
    case ParseErrorCode::LexerInvalidCharacter:
      return "invalid character";
    case ParseErrorCode::LexerUnterminatedString:
      return "unterminated string literal";
    case ParseErrorCode::LexerUnterminatedSheetQuote:
      return "unterminated quoted sheet name";
    case ParseErrorCode::LexerInvalidNumberLiteral:
      return "invalid number literal";
    case ParseErrorCode::LexerInvalidErrorLiteral:
      return "invalid error literal";
    case ParseErrorCode::LexerInvalidEscape:
      return "invalid escape sequence";
    case ParseErrorCode::LexerInvalidReference:
      return "invalid reference";
    case ParseErrorCode::LexerExcessiveLength:
      return "formula exceeds maximum length";
    case ParseErrorCode::UnexpectedToken:
      return "unexpected token";
    case ParseErrorCode::UnexpectedEof:
      return "unexpected end of input";
    case ParseErrorCode::ExpectedExpression:
      return "expected expression";
    case ParseErrorCode::ExpectedCloseParen:
      return "expected ')'";
    case ParseErrorCode::UnbalancedBraces:
      return "unbalanced braces in array literal";
    case ParseErrorCode::ArrayRowMismatch:
      return "array literal rows have inconsistent column counts";
    case ParseErrorCode::ExpectedRParenOrComma:
      return "expected ')' or ','";
    case ParseErrorCode::ExpectedCommaOrSemiInArray:
      return "expected ',' or ';' in array literal";
    case ParseErrorCode::InvalidReference:
      return "invalid reference";
    case ParseErrorCode::UnsupportedConstruct:
      return "construct not yet supported";
    case ParseErrorCode::ExpectedOpenParen:
      return "expected '('";
    case ParseErrorCode::ExpectedComma:
      return "expected ','";
    case ParseErrorCode::UnbalancedBrackets:
      return "unbalanced brackets in structured reference";
    case ParseErrorCode::InvalidRange:
      return "invalid range expression";
    case ParseErrorCode::NestedFormulaTooDeep:
      return "formula nesting depth exceeds the configured limit";
    case ParseErrorCode::TooManyErrors:
      return "too many parse errors; stopping";
    case ParseErrorCode::LetInvalidName:
      return "invalid LET binding name";
    case ParseErrorCode::LetWrongArity:
      return "LET requires an odd number of arguments (name, expr, [name, expr, ...], body)";
    case ParseErrorCode::LambdaInvalidParam:
      return "invalid LAMBDA parameter name";
    case ParseErrorCode::LambdaEmpty:
      return "LAMBDA requires at least a body expression";
    case ParseErrorCode::LambdaDuplicateParam:
      return "LAMBDA parameter names must be unique";
  }
  return "parse error";
}

}  // namespace parser
}  // namespace formulon
