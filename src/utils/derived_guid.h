//
// Deterministic GUID derivation. The engine never mints a random GUID, so an
// operation that needs a fresh identity (a sheet copy's table, validation and
// comment uids) derives one from the GUID it replaces.

#ifndef FORMULON_UTILS_DERIVED_GUID_H_
#define FORMULON_UTILS_DERIVED_GUID_H_

#include <string>
#include <string_view>

namespace formulon {

/// Returns a GUID derived from `source_guid` and `salt`: the 128-bit FNV-1a
/// hash of `source_guid`, one NUL byte and `salt`, with the RFC 4122 version
/// (4) and variant (10xx) bits set, printed as uppercase hex in 8-4-4-4-12
/// groups. Wrapped in braces iff `source_guid` is (`{...}`), so a braced
/// input yields a value `is_threaded_comment_guid` accepts.
std::string derived_guid(std::string_view source_guid, std::string_view salt);

/// Replaces the value of the ` <key>="..."` pair in an escaped attribute list
/// with `derived_guid(value, salt)`; a list without the pair is left as is.
void rederive_attr_guid(std::string& attrs, std::string_view key, std::string_view salt);

}  // namespace formulon

#endif  // FORMULON_UTILS_DERIVED_GUID_H_
