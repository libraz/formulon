//
// Implementation of `format_lambda_value`. See the header for contract.

#include "eval/lambda_format.h"

#include <cstdint>
#include <string>

#include "eval/lambda_value.h"
#include "parser/ast_format.h"

namespace formulon {
namespace eval {

std::string format_lambda_value(const LambdaValue& lv) {
  std::string out;
  out.append("LAMBDA(");
  const std::uint32_t first_optional = lv.param_count - lv.optional_count;
  for (std::uint32_t i = 0; i < lv.param_count; ++i) {
    if (i > 0U) {
      out.push_back(',');
    }
    if (i >= first_optional) {
      out.push_back('[');
      out.append(lv.params[i]);
      out.push_back(']');
    } else {
      out.append(lv.params[i]);
    }
  }
  if (lv.param_count > 0U) {
    out.push_back(',');
  }
  if (lv.body != nullptr) {
    out.append(parser::format_formula(*lv.body));
  }
  out.push_back(')');
  return out;
}

}  // namespace eval
}  // namespace formulon
