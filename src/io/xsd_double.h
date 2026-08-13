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
/// Returns true only when `text` lies in the xs:double lexical space —
/// optional sign, digits, optional fraction, optional `e`/`E` exponent —
/// *and* the resulting double is finite, i.e. a value the writer can
/// re-emit so that it reloads identically.
///
/// Returns false, leaving `*out` unchanged, on empty input, on input that
/// does not start with a number, on trailing characters other than
/// whitespace, and on the three forms `std::strtod` accepts outside that
/// space:
///
///   * a hexadecimal literal (`0x10` would otherwise read as 16);
///   * an `inf` / `infinity` / `nan` spelling in any case, which the
///     writer cannot express in a `<v>` body and turns back into `#NUM!`;
///   * an overflow to ±infinity or an underflow to exact zero.
///
/// Callers treat a false return as `kIoSheetCorrupt` rather than
/// substituting a value, so a number the file states but IEEE 754 cannot
/// hold surfaces at load rather than at the next save.
bool parse_xsd_double(std::string_view text, double* out);

}  // namespace formulon::io

#endif  // FORMULON_IO_XSD_DOUBLE_H_
