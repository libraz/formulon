//
// Writer for `xl/styles.xml`. Symmetric counterpart of
// `src/io/styles_reader.{h,cpp}`: feeding the bytes produced here back
// into the reader must reproduce the input `StylesTable` (modulo
// canonical-form normalisations such as section ordering and the
// suppression of built-in number-format ids 0..163, which Excel rejects
// inside `<numFmts>`).
//
// Section ordering matches the OOXML schema requirement:
//   numFmts -> fonts -> fills -> borders -> cellStyleXfs -> cellXfs ->
//   cellStyles.
//
// When `cellStyleXfs` is empty, the writer emits a synthesized default
// `cellStyleXfs` record and a built-in `Normal` `cellStyles` entry. This is a
// writer-only normalization: the input table is never modified. Existing
// non-empty named-style tables retain their source ordering. Dangling
// `xfId` references are emitted as `0` against the effective style-xf table;
// this normalization is also writer-only.
//
// Design references:
//   * src/io/styles_reader.h (sister reader; canonical schema)

#ifndef FORMULON_IO_STYLES_WRITER_H_
#define FORMULON_IO_STYLES_WRITER_H_

#include <string>

#include "io/styles_reader.h"

namespace formulon {
namespace io {

/// Emits the OOXML `xl/styles.xml` document for `table`. The output
/// always begins with the canonical XML declaration and the
/// `<styleSheet xmlns="...">` root.
///
/// Number-format records whose `id` falls in the built-in range
/// (0..163) are suppressed from `<numFmts>`: Excel rejects packages
/// that re-declare built-in ids. Custom ids (>= 164) are emitted in
/// declaration order.
///
/// Empty input (default-constructed `StylesTable`) yields a
/// minimal-but-valid styles document with one default font / fill /
/// border / cellXf record, plus the default `Normal` named-style pair. This
/// is the same shape Excel emits for a freshly-created workbook.
/// Explicitly present but empty `<alignment/>` children are retained via
/// `CellXf::has_alignment`, including when all alignment values are defaults.
/// The four `CellXf::has_*` alignment-attribute flags likewise preserve
/// explicit schema defaults such as `horizontal="general"` and `wrapText="0"`.
std::string write_styles(const StylesTable& table);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_STYLES_WRITER_H_
