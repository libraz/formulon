// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's Web-category built-ins into a FunctionRegistry:
//
//   * ENCODEURL(text)          - real implementation (RFC 3986 percent-encode)
//   * FILTERXML(xml, xpath)    - real implementation (pugixml + XPath 1.0)
//   * WEBSERVICE(url)          - deliberate stub returning #VALUE!
//   * PY(expression)           - deliberate stub returning #NAME?
//
// Formulon is a pure calculation engine: it does not perform network I/O
// (WEBSERVICE) and does not embed a Python runtime (PY). The two stubs are
// present so the catalog stays complete and formulas that reference them
// parse and surface a deterministic Excel-visible error rather than
// `#NAME?` from an unknown function.

#ifndef FORMULON_EVAL_BUILTINS_WEB_H_
#define FORMULON_EVAL_BUILTINS_WEB_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the Web-category built-ins (ENCODEURL, FILTERXML, WEBSERVICE,
/// PY) into `registry`. Intended to be invoked from `register_builtins`.
void register_web_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_WEB_H_
