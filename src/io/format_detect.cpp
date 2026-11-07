//
// Implementation of workbook container-format detection. See
// `io/format_detect.h` for the contract.

#include "io/format_detect.h"

#include <string_view>

#include "io/ooxml/package_validator.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/part_paths.h"
#include "io/zip_reader.h"

namespace formulon {
namespace io {

WorkbookFormat detect_workbook_format(ByteSpan bytes) {
  ZipReader zip;
  if (auto open = zip.open(bytes); !open) {
    return WorkbookFormat::Unknown;
  }

  // Primary signal: the binary workbook part. Excel always emits the
  // workbook at `xl/workbook.bin` (xlsb) or `xl/workbook.xml` (xlsx).
  const bool has_bin = zip.has_entry(xlsb::kWorkbookBinPath);
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
      // Only parsed Override / Default declarations count, never stray text in the part.
      if (auto declared = ooxml::list_content_type_declarations(ct.value()); declared) {
        for (const ooxml::OverrideEntry& entry : declared.value().overrides) {
          if (is_xlsb_workbook_content_type(entry.content_type)) {
            return WorkbookFormat::Xlsb;
          }
        }
        // An XLSB package may declare the workbook through the `bin` Default instead of an Override.
        for (const DefaultContentType& entry : declared.value().defaults) {
          if (entry.extension == "bin" && is_xlsb_workbook_content_type(entry.content_type)) {
            return WorkbookFormat::Xlsb;
          }
        }
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
