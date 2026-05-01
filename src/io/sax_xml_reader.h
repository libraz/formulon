// Copyright 2026 libraz. Licensed under the MIT License.
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
// XML entity references are decoded inside `<v>` and `<t>` text content
// (`&amp; &lt; &gt; &quot; &apos;`); unknown / numeric character
// references pass through verbatim. Attribute-value entity decoding is
// not performed in this slice (Excel never emits entities into `r=` /
// `t=` / `s=`).
//
// On a malformed stream the scanner returns `kIoXmlParse` with an
// indicative offset / message; it never throws and never reads past the
// supplied buffer.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.4 (Reader pipeline)
//   * backup/plans/26-implementation-plan.md (Phase 5.5 SAX)

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

/// Decoded record for a single `<c>` element. Text fields are
/// `string_view`s into either the input buffer (when no entity decoding
/// was needed) or the per-call decode arena (the `SheetSaxCallbacks`
/// implementation owns the lifetime contract: the views remain valid
/// only for the duration of the `on_cell` call).
struct CellRecord {
  /// 0-based row index, decoded from the enclosing `<row r="N">`.
  std::uint32_t row = 0;
  /// 0-based column index, decoded from the cell's `r="A1"` attribute.
  std::uint32_t col = 0;
  /// `t=` attribute value, e.g. `"n"`, `"s"`, `"b"`, `"e"`, `"str"`,
  /// `"inlineStr"`. Empty when the cell omits `t=` (default = number).
  std::string_view t;
  /// `s=` attribute value (style index). Empty when absent.
  std::string_view s;
  /// `<f>` body, with leading `=` stripped if present. Empty when no
  /// `<f>` child existed.
  std::string_view formula;
  /// Decoded text content of the cell:
  ///   * For `t="inlineStr"`: concatenation of all `<is><t>` and
  ///     `<is><r><t>` text nodes (rich-text formatting dropped). Any
  ///     `<is><rPh><t>` payloads are excluded — they are surfaced
  ///     separately via `phonetic`. The `is_inline_string` flag is
  ///     set in this branch.
  ///   * Otherwise: text content of `<v>` (entity-decoded). Empty when
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
  /// Invoked when a `<row>` opens. `row_1based` is the value of the
  /// `r="N"` attribute, or 0 when the attribute was absent (rare).
  Expected<void, Error> (*on_row_start)(void* user_data, std::uint32_t row_1based) = nullptr;
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
