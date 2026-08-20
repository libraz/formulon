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
// entities. TAB and LF are legal XML characters that survive element text
// unchanged, so they are written literally. CR would be folded into LF by
// XML line-end normalisation, and every other C0 control is illegal in XML
// 1.0 outright, so both take OOXML's `_xHHHH_` notation; a literal sequence
// that resembles that notation is escaped as `_x005F_xHHHH_` so a
// subsequent OOXML reader preserves it literally.
void AppendXmlEscaped(std::string& out, std::string_view in);

// Escapes `in` for an attribute value (`name="..."`). Same five
// XML-critical characters as `AppendXmlEscaped`, but attribute-value
// normalisation turns a literal TAB / LF / CR into a space, so those three
// are written as the `&#9;` / `&#10;` / `&#13;` character references a
// conforming parser restores verbatim.
//
// No OOXML `_xHHHH_` notation is produced. Attribute readers take the
// parser's output directly and never call `AppendOoxmlTextUnescaped`, so an
// escape spelled here would survive into the value on reload and gain
// another six bytes on the next save. Reading an attribute raw is therefore
// the exact inverse of this function: every string an XML 1.0 attribute can
// carry round-trips byte for byte. The bytes it cannot carry -- the C0
// controls other than TAB / LF / CR, and invalid UTF-8 -- are replaced with
// U+FFFD, which makes a second pass a fixed point.
//
// This is also the escaper for `Relationship/@Target`: an `xsd:anyURI` must
// not gain OOXML escapes, and this function emits none.
void AppendXmlAttrEscaped(std::string& out, std::string_view in);

// Appends `in` after decoding OOXML `_xHHHH_` escapes. This is for text
// payloads read from OOXML parts after XML entity decoding; it preserves a
// literal escaped-looking sequence emitted as `_x005F_xHHHH_`.
void AppendOoxmlTextUnescaped(std::string& out, std::string_view in);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XML_ESCAPE_H_
