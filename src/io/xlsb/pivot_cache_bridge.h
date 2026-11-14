//
// Bridge between the XLSB pivot cache parts (definition + records) and the
// in-memory `pivot::PivotCache`. Reading fills the members the xlsx cache
// reader fills for the same cache, spelled the way that reader stores them
// (root attributes as text, `<sharedItems>` hints, `<fieldGroup>`, the
// `<calculatedItem>` formula text, the cache `extLst` as raw XML), so a cache
// read from an .xlsb and from its .xlsx twin compare equal. Writing encodes
// from the model alone.
//
// Members only the xlsx route carries, refused on write when they hold data:
// `cache_source_children_xml`, `trailing_children_xml`, any extLst other than
// the x14 cache definition and dataSourceClean, `PivotCacheField::formula`,
// `extra_attrs`, `trailing_children_xml`, `SharedItemsHints::extra_attrs`,
// discrete grouping, a calculated item's `extra_children_xml` and `@field`.
// An explicit `refreshOnLoad="0"` is written as the clear flag (the record
// has no presence bit). `cache_id` lives in the workbook part.

#ifndef FORMULON_IO_XLSB_PIVOT_CACHE_BRIDGE_H_
#define FORMULON_IO_XLSB_PIVOT_CACHE_BRIDGE_H_

#include <cstdint>
#include <string_view>
#include <vector>

#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// The two parts of one pivot cache.
struct XlsbPivotCacheParts {
  std::vector<std::uint8_t> definition;  ///< `xl/pivotCache/pivotCacheDefinitionN.bin`.
  std::vector<std::uint8_t> records;     ///< `xl/pivotCache/pivotCacheRecordsN.bin`.
};

/// Decodes a cache from its definition and records parts. `date1904` is the
/// workbook's date system. Errors are the codecs' (kIoXlsbRecordCorrupt,
/// kIoXlsbRecordTruncated, kIoXlsbUnknownRecord, kIoXlsbUnsupportedPtg),
/// kIoXlsbUnknownRecord with an "unmeasured" message for a record state with
/// no Excel-saved sample, and kSecurityResourceLimit past
/// `kMaxPivotCacheRecordCells`.
Expected<pivot::PivotCache, Error> pivot_cache_from_xlsb(ByteSpan definition_part, ByteSpan records_part,
                                                         bool date1904);

/// Encodes `cache` as its two parts; the definition names its records part
/// by `records_rel_id`. A model state the records cannot hold, or one with no
/// Excel-saved sample, is kIoFeatureNotWritable (`kind=pivot_cache`).
Expected<XlsbPivotCacheParts, Error> pivot_cache_to_xlsb(const pivot::PivotCache& cache, bool date1904,
                                                         std::string_view records_rel_id);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_CACHE_BRIDGE_H_
