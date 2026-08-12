//
// Locale-independent parser for the XSD `double` / `decimal` lexical values
// OOXML uses for numeric payloads — a `<v>` cell body, a pivot-cache `<n v=…>`
// attribute, and anything else Excel writes as a round-trip-friendly decimal
// string. Excel never localises these, so the C locale is the whole contract.
//
// One definition serves the sheet path and the pivot path so the two cannot
// drift into producing different numbers for the same bytes.

#ifndef FORMULON_IO_XSD_DOUBLE_H_
#define FORMULON_IO_XSD_DOUBLE_H_

#include <string_view>

namespace formulon::io {

/// Parses `text` as a locale-independent double.
///
/// Returns false — leaving `*out` unchanged — on empty input, on input that
/// does not start with a number, and on trailing characters other than
/// whitespace. Returns true and writes the value otherwise.
bool parse_xsd_double(std::string_view text, double* out);

}  // namespace formulon::io

#endif  // FORMULON_IO_XSD_DOUBLE_H_
