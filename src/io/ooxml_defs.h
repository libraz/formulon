// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared OOXML schema constants used by the reader and writer.
// Only the relationship type URIs that BOTH sides need are hoisted
// here; content-type strings that are needed by exactly one side
// (e.g. the `kCtWorkbook*` family the reader uses for kind detection,
// or the writer-only `kCtPackageRels` / `kCtComments`) stay in the
// consuming `.cpp` so this header tracks the minimum shared surface
// and not the union.
//
// Header-only, `<string_view>`-only, `inline constexpr` — no
// translation unit needed. Lives under `io/` because no module
// outside `io/` consumes the OOXML relationship vocabulary. Symbols
// live directly in `formulon::io` to match the sibling-header style
// used by `workbook_kind.h` and `passthrough_part.h`, which keeps
// reader / writer call sites unqualified.
//
// Schema reference: ECMA-376 Part 1 §11.3 (Relationships) and
// §15 (SpreadsheetML content types).

#ifndef FORMULON_IO_OOXML_DEFS_H_
#define FORMULON_IO_OOXML_DEFS_H_

#include <string_view>

namespace formulon {
namespace io {

// Relationship type URIs used by Excel-produced packages. The values
// are the canonical ECMA-376 strings; comparison is case-sensitive and
// must remain byte-identical to what Excel emits.

inline constexpr std::string_view kRelOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
inline constexpr std::string_view kRelWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
inline constexpr std::string_view kRelSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
inline constexpr std::string_view kRelStyles =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
inline constexpr std::string_view kRelTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
inline constexpr std::string_view kRelPivotCacheDefinition =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
inline constexpr std::string_view kRelPivotCacheRecords =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
inline constexpr std::string_view kRelPivotTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
inline constexpr std::string_view kRelHyperlink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
inline constexpr std::string_view kRelComments =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
inline constexpr std::string_view kRelVmlDrawing =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
inline constexpr std::string_view kRelPrinterSettings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
inline constexpr std::string_view kRelExternalLink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
inline constexpr std::string_view kRelExternalLinkPath =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
inline constexpr std::string_view kRelOleLink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleLink";
inline constexpr std::string_view kRelDdeLink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ddeLink";

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_DEFS_H_
