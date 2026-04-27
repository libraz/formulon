// Copyright 2026 libraz. Licensed under the MIT License.
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

void AppendXmlEscaped(std::string& out, std::string_view in);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XML_ESCAPE_H_
