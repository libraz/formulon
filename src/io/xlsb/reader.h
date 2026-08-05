//
// MS-XLSB (.xlsb) package reader.
//
// The XLSB package envelope is identical to .xlsx: a ZIP archive with
// `[Content_Types].xml`, `_rels/.rels`, `xl/_rels/workbook.xml.rels`
// (XML even in XLSB), plus a parallel set of binary parts —
// `xl/workbook.bin`, `xl/worksheets/sheet*.bin`, `xl/sharedStrings.bin`.
// The reader:
//
//   * walks the same XML rels/content-types parts as the OOXML reader;
//   * decodes `xl/workbook.bin` for the sheet-bundle list;
//   * decodes each `xl/worksheets/sheet*.bin` for cell data
//     (BrtCellRk / BrtCellReal / BrtCellBool / BrtCellSt /
//      BrtCellIsst / BrtCellBlank / BrtCellError / BrtFmlaNum /
//      BrtFmlaString / BrtFmlaBool / BrtFmlaError);
//   * decodes `xl/sharedStrings.bin` (BrtSSTItem entries) into the
//     workbook-lifetime `text_storage` deque.
//
// Formulas are decoded through a full `Ptg → AST → Excel-formula-text`
// pipeline (`io/xlsb/ptg_reader.h`'s `decode_ptgs` plus
// `parser::format_formula`; see `DecodeFormulaText` in reader.cpp).
// When a Ptg stream uses a token outside the supported set, the reader
// logs a structured warning and leaves the cell's formula text empty so
// the already-decoded cached value is preserved instead of a fabricated
// formula that would recalc to `#NAME?`.
//
// Default-typed binary parts (images, OLE) are NOT captured here; only
// Override-listed parts the reader does not consume flow into
// `XlsbReadResult::unknown_parts` and `Workbook::passthrough_parts()`,
// matching the OOXML reader's contract.

#ifndef FORMULON_IO_XLSB_READER_H_
#define FORMULON_IO_XLSB_READER_H_

#include <cstdint>
#include <vector>

#include "io/passthrough_part.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Result of `read_xlsb`: a constructed (but not recalc'd) Workbook
/// plus a list of binary parts the reader did not consume. Mirrors
/// `OoxmlReadResult` so callers can treat both flows symmetrically.
///
/// Text payloads (BrtCellSt inline strings, BrtSSTItem entries) are
/// interned into the workbook itself
/// (`Workbook::mutable_text_storage()`) so that `Value::text` views
/// remain valid for the workbook's lifetime — even after the
/// `XlsbReadResult` is destroyed and the workbook is moved out. Earlier
/// slices stored the deque on this struct, which caused a use-after-free
/// when callers wrote
/// `Workbook wb = std::move(read_xlsb(...).value().workbook);` — the
/// result (and its deque) went out of scope at the end of the
/// statement while text views were still aliasing it.
///
/// `cells_read` is an audit counter: every cell record successfully
/// decoded into the workbook (literal or formula) bumps it once.
/// `undecoded_*` counters record lossy recovery from Ptg streams outside
/// the supported vocabulary; their cached values remain available but a
/// caller can now detect the missing formulas without parsing log output.
struct XlsbReadResult {
  Workbook workbook;
  std::vector<PassthroughPart> unknown_parts;
  std::uint32_t cells_read = 0;
  std::uint32_t undecoded_formula_count = 0;
  std::uint32_t undecoded_defined_name_count = 0;
};

/// Reads a `.xlsb` package from in-memory bytes.
///
/// Returns:
///   * `kIoZipCorrupt`         — archive-level miniz failure.
///   * `kIoXmlParse`           — malformed `[Content_Types].xml`,
///                                `_rels/.rels`, or
///                                `xl/_rels/workbook.xml.rels`.
///   * `kIoRelationshipBroken` — required relationship missing.
///   * `kIoContentTypeInvalid` — required content type missing.
///   * `kIoXlsbRecordTruncated`/`kIoXlsbCorrupt` — record-stream
///                                                    decoding failure.
Expected<XlsbReadResult, Error> read_xlsb(ByteSpan bytes);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_READER_H_
