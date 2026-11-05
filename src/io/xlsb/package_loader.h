//
// The OPC package envelope of an XLSB file: `[Content_Types].xml`, the
// package-level and workbook-level relationship parts, and the sheet and
// pivot relationship lookups the reader performs before decoding any
// binary part.

#ifndef FORMULON_IO_XLSB_PACKAGE_LOADER_H_
#define FORMULON_IO_XLSB_PACKAGE_LOADER_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "default_content_type.h"
#include "io/zip_reader.h"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace detail {

/// Walks `[Content_Types].xml` and asserts at least one Override is
/// for the xlsb workbook content type. Also returns the full Override
/// list so callers can compute passthrough parts.
struct ContentTypesView {
  std::vector<std::pair<std::string, std::string>> overrides;  // (part_name, content_type)
  std::vector<DefaultContentType> defaults;
};

/// Parses `[Content_Types].xml`.
Expected<ContentTypesView, Error> LoadContentTypes(const std::vector<std::uint8_t>& ct_bytes);

/// Returns the package-relative part path the package-level rels name as the
/// OfficeDocument.
Expected<std::string, Error> ResolveOfficeDocumentPath(const std::vector<std::uint8_t>& rels_bytes);

struct WorkbookRels {
  std::unordered_map<std::string, std::string> sheet_targets;  // rId -> resolved part path
  std::string sst_path;
  std::string styles_path;
  std::vector<UnknownRelationship> unknown_rels;
};

/// Reads the workbook part's rels.
Expected<WorkbookRels, Error> LoadWorkbookRels(const ZipReader& zip, std::string_view workbook_path);

/// Reads every entry of a worksheet's `.rels` part (see the definition).
Expected<std::vector<UnknownRelationship>, Error> LoadSheetRelationships(const ZipReader& zip,
                                                                         std::string_view rels_path,
                                                                         std::string_view sheet_dir);

/// Returns the package-relative target of the last internal relationship of
/// `part_path` whose `attr` equals `value` (see the definition).
Expected<std::string, Error> FindRelationship(const ZipReader& zip, std::string_view part_path, const char* attr,
                                              std::string_view value, std::string_view label);

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PACKAGE_LOADER_H_
