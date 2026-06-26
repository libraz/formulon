// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Strict ISO 8601 date-time parsing for OOXML payloads that carry a typed
// date value (`<c t="d">` cells in strict OOXML, and `<d>` shared / record
// items in a pivot cache). These payloads use the fixed lexical form
// `YYYY-MM-DD` optionally followed by `Thh:mm:ss[.fff]` and an optional
// trailing time-zone designator. Unlike the locale-aware DATEVALUE parser,
// this grammar is fixed by the spec, so it is decoded here directly against
// the engine's serial-number helpers rather than routed through the
// ja-JP-aware text parser.

#ifndef FORMULON_IO_ISO_DATE_H_
#define FORMULON_IO_ISO_DATE_H_

#include <string_view>

namespace formulon::io {

/// Parses a strict ISO 8601 date / date-time body into an Excel serial
/// number (1900 date system, the convention the rest of the engine uses).
///
/// Accepted shapes:
///   * `YYYY-MM-DD`
///   * `YYYY-MM-DDThh:mm:ss`
///   * `YYYY-MM-DDThh:mm:ss.fff` (fractional seconds)
///   * any of the above with a trailing `Z` or `+hh:mm` / `-hh:mm` offset,
///     which is accepted and ignored (Excel stores a wall-clock serial).
///
/// On success returns true and writes the serial to `*out_serial`. On any
/// lexical violation returns false and leaves `*out_serial` untouched, so
/// the caller can fall back to surfacing the raw text.
bool parse_iso_date_serial(std::string_view text, double* out_serial) noexcept;

}  // namespace formulon::io

#endif  // FORMULON_IO_ISO_DATE_H_
