//
// Name spelling for formula text: sheet-name quoting, and the `_xlpm.` /
// `_xlnm.` prefixes that keep cell-shaped LET / LAMBDA parameter and defined
// names from reading back as cells.

#include <string>
#include <string_view>

#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_format_internal.h"
#include "parser/formula_prefix.h"
#include "parser/reference.h"
#include "utils/strings.h"

namespace formulon {
namespace parser {

void append_sheet_name(std::string_view sheet, bool force_quote, std::string& out) {
  append_sheet_name_quoted_if(sheet, force_quote || local_sheet_needs_quoting_a1(sheet), out);
}

void append_quote_escaped(std::string_view s, std::string& out) {
  for (char c : s) {
    if (c == '\'') {
      out.push_back('\'');
    }
    out.push_back(c);
  }
}

void append_sheet_name_quoted_if(std::string_view sheet, bool quote, std::string& out) {
  if (quote) {
    out.push_back('\'');
    append_quote_escaped(sheet, out);
    out.push_back('\'');
  } else {
    out.append(sheet);
  }
}

namespace detail {
namespace {

// Appends the `Sheet1!` / `'My Sheet'!` qualifier of a sheet-scoped NameRef;
// nothing for an unqualified one.
void AppendNameSheetQualifier(const AstNode& node, std::string& out) {
  const std::string_view sheet = node.as_name_sheet();
  if (sheet.empty()) {
    return;
  }
  append_sheet_name(sheet, node.as_name_sheet_quoted(), out);
  out.push_back('!');
}

}  // namespace

// Case-insensitive: Excel folds ASCII case on name resolution.
bool InScope(const NameScope& scope, std::string_view bare) {
  for (const std::string_view s : scope) {
    if (strings::case_insensitive_eq(s, bare)) {
      return true;
    }
  }
  return false;
}

bool HasParamPrefix(std::string_view name) noexcept {
  return name.size() >= kParamPrefix.size() && formula_prefix_ci_eq(name.substr(0, kParamPrefix.size()), kParamPrefix);
}

// Appends `raw` carrying the parameter prefix, once.
void AppendParamName(std::string_view raw, std::string& out) {
  if (!HasParamPrefix(raw)) {
    out.append(kParamPrefix);
  }
  out.append(raw);
}

// A cell-shaped binding name keeps `_xlpm.` so it cannot read back as a cell;
// `storage` gives every other one the prefix too, once.
void AppendBindingName(std::string_view raw, bool storage, std::string& out) {
  const std::string_view bare = cell_shaped_name_bare(raw);
  if (is_a1_cell_shaped(bare)) {
    out.append(kParamPrefix);
    out.append(bare);
    return;
  }
  if (storage) {
    AppendParamName(raw, out);
    return;
  }
  out.append(raw);
}

namespace {

// Appends the unqualified name `raw` when it is cell-shaped: `_xlpm.` for a
// binding in scope unless it is spelled `_xlnm.`, `_xlnm.` otherwise. False,
// appending nothing, for any other name.
bool AppendCellShapedName(std::string_view raw, const NameScope& scope, std::string& out) {
  const std::string_view bare = cell_shaped_name_bare(raw);
  if (!is_a1_cell_shaped(bare)) {
    return false;
  }
  const bool bound = cell_shaped_name_prefix(raw) != kNamePrefix && InScope(scope, bare);
  out.append(bound ? kParamPrefix : kNamePrefix);
  out.append(bare);
  return true;
}

}  // namespace

// `AppendCellShapedName` for a call name; the built-in `LOG10` is no name.
bool AppendCellShapedCallName(std::string_view raw, const NameScope& scope, std::string& out) {
  return !is_cellref_shaped_function_name(raw) && AppendCellShapedName(raw, scope, out);
}

// A NameRef with its qualifier. A sheet-qualified name is never a binding,
// so a cell-shaped one is always `Sheet1!_xlnm.B1`.
void AppendNameRef(const AstNode& node, const NameScope& scope, std::string& out) {
  const std::string_view name = node.as_name();
  if (node.as_name_sheet().empty()) {
    if (!AppendCellShapedName(name, scope, out)) {
      out.append(name);
    }
    return;
  }
  AppendNameSheetQualifier(node, out);
  if (!AppendCellShapedName(name, NameScope{}, out)) {
    out.append(name);
  }
}

}  // namespace detail
}  // namespace parser
}  // namespace formulon
