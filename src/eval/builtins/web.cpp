// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's Web-category built-in functions:
//
//   * ENCODEURL(text)        -- RFC 3986 unreserved-set percent-encoding.
//   * FILTERXML(xml, xpath)  -- pugixml-backed XPath 1.0 evaluation.
//   * WEBSERVICE(url)        -- fixed #VALUE! (no network I/O).
//   * PY(expression)         -- fixed #NAME? (no embedded Python).
//
// ENCODEURL and FILTERXML are real implementations. WEBSERVICE and PY are
// intentional stubs: their semantics require a runtime environment that a
// pure calc engine does not provide. The stub bodies still evaluate all
// their arguments (via the dispatcher's normal eager path) so that an
// error in any argument propagates as usual before the fixed return fires.
//
// pugixml is already a project dependency (used by the OOXML readers) and
// is linked into formulon_core. PUGIXML_NO_EXCEPTIONS is enforced globally
// so parse/xpath failures surface as status codes on the returned objects
// rather than as C++ exceptions.

#include "eval/builtins/web.h"

#include <cstdint>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "pugixml.hpp"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// ENCODEURL
// ---------------------------------------------------------------------------
//
// Percent-encodes `text` per RFC 3986: every byte that is NOT in the
// unreserved set (`A-Z a-z 0-9 - _ . ~`) is emitted as `%XX` with uppercase
// hex digits. The input is treated as opaque UTF-8 bytes, so a multi-byte
// code point becomes multiple `%XX` sequences (e.g. the Japanese character
// `日` -> `%E6%97%A5`).
//
// Coercion rules (via `coerce_to_text`):
//   * Text    -> input verbatim.
//   * Number  -> shortest-roundtrip textual form.
//   * Bool    -> "TRUE" / "FALSE" (Excel's canonical literals).
//   * Blank   -> "".
//   * Error   -> propagated (dispatcher short-circuits before we run).
//
// Uppercase hex matches Mac Excel 365 ja-JP output; RFC 3986 §2.1 actually
// recommends uppercase as well ("For consistency, URI producers and
// normalizers should use uppercase hexadecimal digits").

bool is_unreserved(unsigned char c) noexcept {
  if (c >= 'A' && c <= 'Z') {
    return true;
  }
  if (c >= 'a' && c <= 'z') {
    return true;
  }
  if (c >= '0' && c <= '9') {
    return true;
  }
  return c == '-' || c == '_' || c == '.' || c == '~';
}

Value EncodeUrl(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& s = text.value();
  std::string out;
  // Worst case: every byte expands to 3 characters.
  out.reserve(s.size() * 3);
  constexpr char kHex[] = "0123456789ABCDEF";
  for (char raw : s) {
    const auto c = static_cast<unsigned char>(raw);
    if (is_unreserved(c)) {
      out.push_back(static_cast<char>(c));
      continue;
    }
    out.push_back('%');
    out.push_back(kHex[(c >> 4) & 0x0F]);
    out.push_back(kHex[c & 0x0F]);
  }
  return Value::text(arena.intern(out));
}

// ---------------------------------------------------------------------------
// FILTERXML
// ---------------------------------------------------------------------------
//
// Parses `xml_text` with pugixml, evaluates `xpath_text` against it, and
// returns the textual content of the matched node(s).
//
// Error surface:
//   * XML parse failure   -> #VALUE!
//   * XPath parse failure -> #VALUE!
//   * Empty node set      -> #N/A  (Excel's documented no-match code)
//   * Parse / select OK   -> first node's text content.
//
// TODO(filterxml-spill): Excel 365's dynamic-array engine spills the entire
// node set into a vertical array on the worksheet. Formulon now has
// `Value::Array` (used by SUMPRODUCT operator broadcasting), but a bare
// `=FILTERXML(...)` cell still has no spill plumbing through OOXML
// serialization, the C-API, or downstream consumers. Returning a
// `Value::Array` cell value here without spill would worsen the user
// experience vs. the current "first node's text" fallback. Revisit once
// dynamic-array spill semantics land at the cell level.
//
// pugixml quirks worth noting:
//   * Its XPath engine implements XPath 1.0 only. XPath 2.0+ features (`for`,
//     `let`, regex functions, sequence types) will fail to parse. This is
//     the same subset Mac Excel 365 ja-JP supports, so no divergence is
//     expected for real workbooks.
//   * `select_nodes(expr)` with a malformed expression populates a non-ok
//     xpath_parse_result in the returned xpath_node_set's status. With
//     PUGIXML_NO_EXCEPTIONS set, the malformed expression returns an empty
//     node set rather than throwing -- we detect this via `xpath_query`'s
//     result status so the distinction between "bad xpath" (#VALUE!) and
//     "valid xpath, no match" (#N/A) is preserved.

// Extracts the text content of a pugixml node. For element nodes we
// concatenate the children's text (pugixml's xml_text::as_string handles
// leaf CDATA / PCDATA); for attribute nodes we return the attribute value.
std::string_view node_text(const pugi::xpath_node& n) {
  if (n.node()) {
    // Element / document / text node. xml_text::as_string returns the
    // aggregated text of the node's PCDATA children, which matches
    // Excel's "inner text" semantics for FILTERXML.
    return n.node().text().as_string();
  }
  if (n.attribute()) {
    return n.attribute().value();
  }
  return {};
}

Value FilterXml(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto xml_text = coerce_to_text(args[0]);
  if (!xml_text) {
    return Value::error(xml_text.error());
  }
  auto xpath_text = coerce_to_text(args[1]);
  if (!xpath_text) {
    return Value::error(xpath_text.error());
  }

  pugi::xml_document doc;
  const std::string& xml = xml_text.value();
  pugi::xml_parse_result parse = doc.load_buffer(xml.data(), xml.size());
  if (!parse) {
    return Value::error(ErrorCode::Value);
  }

  // Compile the XPath query explicitly so we can distinguish a malformed
  // expression (parse error -> #VALUE!) from a valid expression that
  // produces an empty node set (-> #N/A). With PUGIXML_NO_EXCEPTIONS in
  // effect the constructor records the compile error on `query.result()`
  // rather than throwing.
  const std::string& xpath = xpath_text.value();
  pugi::xpath_query query(xpath.c_str(), nullptr);
  if (!query.result()) {
    return Value::error(ErrorCode::Value);
  }

  pugi::xpath_node_set nodes = query.evaluate_node_set(doc);
  if (nodes.empty()) {
    return Value::error(ErrorCode::NA);
  }

  // Multi-node spill is deferred until dynamic-array spill semantics land
  // at the cell level (see TODO above); for now return the first node's text.
  return Value::text(arena.intern(node_text(nodes.first())));
}

// ---------------------------------------------------------------------------
// WEBSERVICE(url)
// ---------------------------------------------------------------------------
//
// Always returns #VALUE!. Formulon is a pure calculation engine and does
// not perform network I/O under any circumstances; there is no HTTP client
// in the dependency set (see `backup/plans/00-design-principles.md` and
// the strict 6-library policy in CLAUDE.md). The argument is still
// evaluated by the dispatcher so an error there propagates normally.
Value WebService(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// PY(expression)
// ---------------------------------------------------------------------------
//
// Always returns #NAME?. Excel's PY function dispatches the expression to
// a hosted Python runtime; Formulon does not embed one. Returning #NAME?
// matches Excel's own behaviour when the cloud Python service is
// unavailable (function-unknown surface).
Value Py(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Name);
}

}  // namespace

void register_web_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"ENCODEURL", 1u, 1u, &EncodeUrl});
  registry.register_function(FunctionDef{"FILTERXML", 2u, 2u, &FilterXml});
  // Stubs: eager dispatch still propagates argument errors (default
  // propagate_errors = true) before the fixed return fires.
  registry.register_function(FunctionDef{"WEBSERVICE", 1u, 1u, &WebService});
  registry.register_function(FunctionDef{"PY", 1u, 1u, &Py});
}

}  // namespace eval
}  // namespace formulon
