// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Parser AST -> MS-XLSB Ptg stream encoder. The exact inverse of
// `ptg_reader.h`: a post-order walk of the AST emits operand tokens
// first, then the operator / function token that consumes them, so the
// resulting `rgce` byte stream decodes back to a structurally
// equivalent AST.
//
// The encoder is matched byte-for-byte with the decoder in this module
// (the two are the engine's own private round-trip pair); the on-wire
// shapes follow [MS-XLSB] §2.5.97 for the common token set but the
// decoder is the authoritative consumer, so the encoder need only stay
// consistent with it.
//
// Tokens the AST can carry but the encoder cannot lower (defined-name
// refs, structured refs, external refs, lambda / let forms, spilled
// refs, implicit-intersection) return `kIoXlsbUnsupportedPtg` rather
// than silently dropping data; the cell writer surfaces that as a hard
// failure through `write_xlsb`'s `Expected` return.

#ifndef FORMULON_IO_XLSB_PTG_WRITER_H_
#define FORMULON_IO_XLSB_PTG_WRITER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "parser/ast.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Encodes the AST rooted at `node` into an `rgce` Ptg byte stream
/// appended to a freshly returned vector. `sheet_names` maps a sheet
/// display name to its 0-based index (the `ixti` used by 3-D
/// references); names not present resolve to index 0 with a `#REF!`
/// shape is *not* emitted — instead the encode fails so the caller can
/// decide. (In practice the cell writer passes the full workbook sheet
/// list, so any sheet referenced by a live formula resolves.)
///
/// Returns `kIoXlsbUnsupportedPtg` for any node kind outside the
/// supported set (see header banner). The error context names the
/// offending node kind.
Expected<std::vector<std::uint8_t>, Error> encode_ptgs(const parser::AstNode& node,
                                                       const std::vector<std::string>& sheet_names);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_WRITER_H_
