// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tiny shared helper for OOXML serialisation: escapes the five XML-critical
// characters (`& < > " '`) into named entities. Non-ASCII bytes (UTF-8
// continuation bytes etc.) are emitted verbatim so multi-byte payloads such
// as Japanese sheet names round-trip unchanged.
//
// Lives in its own TU so that both the package-level writer
// (ooxml_writer.cpp) and the cell-level builder (ooxml_writer_cell.cpp) link
// against a single instantiation; previously each TU carried an
// anonymous-namespace copy.

#ifndef FORMULON_IO_XML_ESCAPE_H_
#define FORMULON_IO_XML_ESCAPE_H_

#include <string>
#include <string_view>

namespace formulon {
namespace io {

// Escapes `in` for element text content: `& < > " '` become named
// entities; TAB / LF / CR pass through verbatim (a conforming XML reader
// preserves them unchanged in text content, so no character reference is
// needed there).
void AppendXmlEscaped(std::string& out, std::string_view in);

// Escapes `in` for an attribute value (`name="..."`). In addition to the
// same five XML-critical characters as `AppendXmlEscaped`, TAB / LF / CR
// are emitted as character references (`&#9;` / `&#10;` / `&#13;`).
//
// This differs from element-text escaping because XML attribute-value
// normalisation (a mandatory step every conforming parser performs, e.g.
// pugixml and Excel itself) replaces a *literal* TAB/LF/CR inside an
// attribute value with a single space; only a character reference for
// those code points survives normalisation intact. Without this, a
// multi-line string written into an attribute (e.g. a defined name with
// an embedded newline) silently loses its exact whitespace on reload.
void AppendXmlAttrEscaped(std::string& out, std::string_view in);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XML_ESCAPE_H_
