// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of workbook container-format detection. See
// `io/format_detect.h` for the contract.

#include "io/format_detect.h"

#include <string_view>
#include <vector>

#include "io/zip_reader.h"

namespace formulon {
namespace io {
namespace {

// XLSB workbook content types ([Content_Types].xml Override). Either form
// marks the package as MS-XLSB.
constexpr std::string_view kCtWorkbookXlsb = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
constexpr std::string_view kCtWorkbookXlsbAlt = "application/vnd.ms-excel.sheet.macroEnabled.main";

bool buffer_contains(const std::vector<std::uint8_t>& haystack, std::string_view needle) {
  if (needle.empty() || haystack.size() < needle.size()) {
    return false;
  }
  const std::string_view view(reinterpret_cast<const char*>(haystack.data()), haystack.size());
  return view.find(needle) != std::string_view::npos;
}

}  // namespace

WorkbookFormat detect_workbook_format(ByteSpan bytes) {
  ZipReader zip;
  if (auto open = zip.open(bytes); !open) {
    return WorkbookFormat::Unknown;
  }

  // Primary signal: the binary workbook part. Excel always emits the
  // workbook at `xl/workbook.bin` (xlsb) or `xl/workbook.xml` (xlsx).
  const bool has_bin = zip.has_entry("xl/workbook.bin");
  const bool has_xml = zip.has_entry("xl/workbook.xml");
  if (has_bin && !has_xml) {
    return WorkbookFormat::Xlsb;
  }
  if (has_xml && !has_bin) {
    return WorkbookFormat::Ooxml;
  }

  // Secondary signal: when neither/both standard part names are present
  // (a non-default workbook part location), consult [Content_Types].xml
  // for the xlsb content type. This keeps detection independent of the
  // exact part path, matching the reader's own content-type gate.
  if (zip.has_entry("[Content_Types].xml")) {
    if (auto ct = zip.read_entry("[Content_Types].xml"); ct) {
      if (buffer_contains(ct.value(), kCtWorkbookXlsb) || buffer_contains(ct.value(), kCtWorkbookXlsbAlt)) {
        return WorkbookFormat::Xlsb;
      }
    }
  }

  // A readable ZIP that we could not classify. Prefer OOXML when the XML
  // workbook part exists (covers the has_bin && has_xml ambiguous case),
  // else Unknown so the caller falls back to the OOXML diagnostics path.
  if (has_xml) {
    return WorkbookFormat::Ooxml;
  }
  return WorkbookFormat::Unknown;
}

}  // namespace io
}  // namespace formulon
