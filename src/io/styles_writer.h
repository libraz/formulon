// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
// `cellStyleXfs` and `cellStyles` are emitted only when non-empty;
// freshly-created or never-named-styled workbooks omit both.
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
/// border / cellXf record. This is the same shape Excel emits for a
/// freshly-created workbook.
std::string write_styles(const StylesTable& table);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_STYLES_WRITER_H_
