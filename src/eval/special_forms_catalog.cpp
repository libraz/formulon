//
// Sole source of truth for the parser-integrated special-form name list.
// See `eval/special_forms_catalog.h` for the rationale.

#include "eval/special_forms_catalog.h"

#include "utils/strings.h"

namespace formulon {
namespace eval {

namespace {

struct SpecialForm {
  const char* name;
  std::uint32_t min_arity;
  std::uint32_t max_arity;
};

// `LET` and `LAMBDA` are both lowered by the parser to dedicated AST shapes
// (`NodeKind::LetBinding` and `NodeKind::Lambda` / `NodeKind::LambdaCall`);
// neither name reaches the `FunctionRegistry` or the tree walker's
// `kLazyDispatch` table. The parser enforces the minimums and LET's odd
// count; the maximums are Excel's documented limits (126 name/value pairs
// plus the body, 253 parameters plus the body).
constexpr SpecialForm kSpecialForms[] = {
    {"LET", 3U, 253U},
    {"LAMBDA", 1U, 254U},
};

}  // namespace

const char* const* parser_special_form_names() {
  // Nullptr-terminated so callers can walk the array without a separate
  // length accessor.
  static constexpr const char* kNames[] = {kSpecialForms[0].name, kSpecialForms[1].name, nullptr};
  return kNames;
}

bool find_special_form_arity(std::string_view name, std::uint32_t* min_arity, std::uint32_t* max_arity) {
  for (const SpecialForm& form : kSpecialForms) {
    if (strings::case_insensitive_eq(name, form.name)) {
      *min_arity = form.min_arity;
      *max_arity = form.max_arity;
      return true;
    }
  }
  return false;
}

}  // namespace eval
}  // namespace formulon
