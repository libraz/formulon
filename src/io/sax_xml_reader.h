//
// Streaming SAX-style scanner for `xl/worksheets/sheet*.xml`.
//
// The default OOXML reader path (`io::read_sheet_data`) fully parses each
// sheet through pugixml, producing a DOM proportional in memory to the
// number of cells. For very large sheets (1M+ cells) the DOM cost
// dominates total memory; this scanner delivers the same `(row, col, t,
// s, formula, value, is_inline_string)` records via callbacks instead so
// the consumer can emit cells directly without retaining the tree.
//
// Recognised vocabulary (deliberately narrow):
//   * `<sheetData>`              - opens/closes the cell stream
//   * `<row r="N" ...>`          - row boundary; `r` decoded to 1-based
//   * `<c r="A1" t="..." s="...">` ... `</c>` (or self-closing)
//   * `<f>...</f>`               - formula text inside a `<c>`
//   * `<v>...</v>`               - cached value inside a `<c>`
//   * `<is><t>...</t></is>`      - inline-string body inside a `<c>`
// Every other element (extLst, mergeCells, hyperlinks, conditionalFormatting,
// dataValidations, dimension, sheetViews, cols, ...) is skipped in
// O(content) time without decoding children.
//
// XML entity references are decoded in semantic PCDATA and attribute values
// (`&amp; &lt; &gt; &quot; &apos;`, decimal character references, and
// lowercase-`x` hexadecimal character references). Unknown / malformed
// references pass through verbatim for compatibility with the historical
// scanner. Literal CR/CRLF in PCDATA becomes LF; literal TAB/LF/CR in
// attributes becomes one space (CRLF is one space), matching pugixml's
// `parse_default` conversion rules. Character-reference whitespace is not
// normalized. Values that need no decoding or normalization remain
// zero-copy views into the input buffer.
//
// Character data may be interrupted by non-element markup wherever XML
// permits it. A comment or processing instruction forms no node under
// `parse_default`, a CDATA section and a child element each form one, and
// an element's text is its first character-data node — so the scanner
// reports the same single run pugixml would expose, with a CDATA run
// end-of-line-normalized but never entity-decoded.
//
// On a malformed stream the scanner returns `kIoXmlParse` with an
// indicative offset / message; it never throws and never reads past the
// supplied buffer. Only input pugixml itself rejects is reported that
// way: the scanner is the alternate reader for the same bytes, so it may
// not turn a document the DOM path loads into a load failure.

#ifndef FORMULON_IO_SAX_XML_READER_H_
#define FORMULON_IO_SAX_XML_READER_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// Decoded record for a single `<c>` element. Text and semantic attribute
/// fields are `string_view`s into either the input buffer (when no entity
/// decoding or XML normalization was needed) or per-scan scratch storage.
/// The views remain valid only for the duration of the `on_cell` callback.
struct CellRecord {
  /// 0-based row index, decoded from the enclosing `<row r="N">`.
  std::uint32_t row = 0;
  /// 0-based column index, decoded from the cell's `r="A1"` attribute.
  std::uint32_t col = 0;
  /// Decoded `t=` attribute value, e.g. `"n"`, `"s"`, `"b"`, `"e"`,
  /// `"str"`, `"inlineStr"`. Empty when the cell omits `t=` (default =
  /// number).
  std::string_view t;
  /// Decoded `s=` attribute value (style index). Empty when absent.
  std::string_view s;
  /// `<f>` body, with leading `=` stripped if present. Empty when no
  /// `<f>` child existed, and also empty for a shared-formula follower
  /// (`<f t="shared" si="N"/>`), whose body lives on the group master.
  std::string_view formula;
  /// Decoded `<f>` element's `t=` attribute (`"shared"`, `"array"`,
  /// `"dataTable"`), or empty when the `<f>` has no `t=` (a plain formula)
  /// or no `<f>` child exists.
  std::string_view f_t;
  /// Decoded `<f>` element's `si=` (shared-formula group index), or empty.
  std::string_view f_si;
  /// Decoded `<f>` element's `ref=` (shared-formula master range), or empty.
  std::string_view f_ref;
  /// Decoded text content of the cell:
  ///   * For `t="inlineStr"`: concatenation of all `<is><t>` and
  ///     `<is><r><t>` text nodes (rich-text formatting dropped). Any
  ///     `<is><rPh><t>` payloads are excluded — they are surfaced
  ///     separately via `phonetic`. The `is_inline_string` flag is
  ///     set in this branch.
  ///   * Otherwise: text content of `<v>` (entity-decoded and PCDATA-
  ///     normalized). Empty when
  ///     no `<v>` was present.
  std::string_view value;
  /// True iff the value came from `<is>...</is>` (i.e. `t="inlineStr"`).
  bool is_inline_string = false;
  /// Concatenated kana from every `<is><rPh><t>` payload, in document
  /// order. Empty when the cell has no inline phonetic annotation. SST-
  /// referenced cells (`t="s"`) carry their phonetic through the SST
  /// resolution pass instead and leave this field empty.
  std::string_view phonetic;
};

/// One `<row>` element's opening attributes, surfaced to `on_row_start`
/// so the streaming path can recover per-row overrides (height / hidden /
/// custom-height / outline) the DOM path reads off `<row>`. Attribute views
/// are decoded and XML-normalized; they alias either the input buffer or
/// distinct per-row scratch storage and are valid only for the callback's
/// duration. Each is empty when the attribute is absent. `row_1based` is the
/// decoded `r="N"` value, and 0 when the attribute is absent or outside the
/// non-negative-integer lexical space — the same "no usable row number"
/// signal the DOM path produces for those inputs.
struct RowRecord {
  std::uint32_t row_1based = 0;
  std::string_view ht;             ///< Decoded `ht=` (row height in points).
  std::string_view hidden;         ///< Decoded `hidden=` (XSD boolean).
  std::string_view custom_height;  ///< Decoded `customHeight=` (XSD boolean).
  std::string_view outline_level;  ///< Decoded `outlineLevel=` (0..7).
  std::string_view custom_format;  ///< Decoded `customFormat=` (XSD boolean).
  std::string_view style;          ///< Decoded `s=` (row style xf index).
};

/// Callback bundle handed to `scan_sheet_data`. Each callback returns
/// `Expected<void, Error>`; a non-OK return aborts the scan and the
/// scanner forwards the error to its caller verbatim.
///
/// The callback shape is a raw function pointer plus a `user_data`
/// pointer rather than `std::function` — this keeps the WASM binary
/// small (`std::function` pulls in extensive template instantiation
/// for every distinct lambda type) and matches the project's
/// "thin-vtable, no heap allocation in dispatch glue" style. Set a
/// pointer to `nullptr` to disable that callback.
struct SheetSaxCallbacks {
  /// User-supplied state. Forwarded as the first argument to every
  /// callback. The scanner does not interpret this pointer.
  void* user_data = nullptr;
  /// Invoked when a `<row>` opens, carrying the row's opening attributes
  /// (index plus any height / hidden / custom-height / outline overrides) so
  /// the consumer can reproduce the per-row layout the DOM path reads.
  Expected<void, Error> (*on_row_start)(void* user_data, const RowRecord& row) = nullptr;
  /// Invoked when the matching `</row>` closes (or when the row was
  /// self-closed).
  Expected<void, Error> (*on_row_end)(void* user_data, std::uint32_t row_1based) = nullptr;
  /// Invoked once per `<c>` element. The record's text views are
  /// valid only for the duration of this callback.
  Expected<void, Error> (*on_cell)(void* user_data, const CellRecord& cell) = nullptr;
};

/// Streams the `<sheetData>` content of a worksheet XML buffer through
/// the supplied callbacks.
///
/// The scanner reads the buffer once, in order, without retaining a
/// DOM. Memory usage is O(longest decoded text node) plus the
/// callbacks' own state.
///
/// Returns `kIoXmlParse` when the input is malformed (unterminated tag,
/// unbalanced `<c>` / `<row>`, oversized cell reference). The error
/// payload's `context` field carries the byte offset where the failure
/// was detected.
Expected<void, Error> scan_sheet_data(ByteSpan xml, const SheetSaxCallbacks& callbacks);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SAX_XML_READER_H_
