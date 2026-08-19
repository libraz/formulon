//
// Structural validation for the raw-XML-fragment setters on the C ABI.
//
// Several worksheet features are modelled as verbatim XML the engine
// carries through a save cycle unchanged (`<autoFilter>`, and the whole
// print-settings group). Their setters accept caller-authored bytes, so
// the engine confirms only that what arrived is one complete, well-formed
// element with the expected name — schema validation belongs to Excel and
// other OOXML consumers, and rejecting content the engine merely does not
// model would defeat the point of a passthrough.
//
// The check runs at set time rather than at save time on purpose: a
// fragment that cannot be emitted must not be storable, or the caller
// learns about it from a workbook Excel refuses to open.

#ifndef FORMULON_C_API_PARTS_XML_FRAGMENT_H_
#define FORMULON_C_API_PARTS_XML_FRAGMENT_H_

#include <cstddef>
#include <string>
#include <string_view>

namespace formulon {
namespace c_api {
namespace parts {

/// Outcome of one fragment check. `context` is the machine-friendly
/// `key=value` diagnostic the entry point forwards verbatim into the
/// thread-local error context.
struct FragmentValidation {
  bool valid = false;
  std::string context;
};

/// Validates that `fragment` parses as exactly one top-level element named
/// `expected_name`.
///
/// Rejects: a parse failure, a top-level node that is not an element
/// (declaration, comment, processing instruction, doctype), more than one
/// top-level element, and a root whose name differs from `expected_name`.
///
/// The name comparison is exact, so a prefixed root (`<x:pageSetup/>`) is
/// rejected: pugixml reports `name()` with the prefix attached, and a
/// fragment lifted out of its namespace context would serialise into the
/// worksheet with a prefix that no longer binds.
///
/// An empty `fragment` is the caller's "remove this element" signal and is
/// never passed here — entry points short-circuit before calling.
FragmentValidation validate_single_element_fragment(std::string_view fragment, std::string_view expected_name);

/// Returns `true` when `fragment` is within `limit_bytes`.
///
/// Separate from the structural check so the entry point can map an
/// oversized fragment to `kPreconditionFailed` (a resource bound) rather
/// than `kInvalidArgument` (malformed input), and so the size gate runs
/// before the parser ever sees hostile input.
bool fragment_within_size_limit(std::string_view fragment, std::size_t limit_bytes);

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_XML_FRAGMENT_H_
