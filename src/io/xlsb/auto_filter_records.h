//
// AutoFilter records ([MS-XLSB] AUTOFILTER: BrtBeginAFilter 161 ..
// BrtEndAFilter 162) to and from the typed `AutoFilter` model.
//
// Covered: the range, per-column button flags, value lists with date groups
// (BrtBeginFilters 165 / BrtFilter 167 / BrtAFilterDateGroupItem 175),
// custom criteria (BrtBeginCustomFilters 172 / BrtCustomFilter 174), top-N
// (BrtTop10Filter 170), dynamic (BrtDynamicFilter 171), colour
// (BrtColorFilter 168) and icon (BrtIconFilter 169) criteria, and the sort
// state Excel nests after the columns (SORTSTATE: BrtBeginSortState 530,
// BrtBeginSortCond 532 / BrtEndSortCond 533, BrtEndSortState 531). The block
// sequence follows the record descriptions and Excel's own files (the
// common-production grammar file the spec refers to is not published
// inline). Anything else -- the x14 rich-filter and sort records, a
// future-record wrapper -- fails the read so the caller can keep the owning
// part verbatim.

#ifndef FORMULON_IO_XLSB_AUTO_FILTER_RECORDS_H_
#define FORMULON_IO_XLSB_AUTO_FILTER_RECORDS_H_

#include <cstdint>
#include <vector>

#include "auto_filter.h"
#include "io/xlsb/record.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Decodes the AUTOFILTER block whose BrtBeginAFilter (payload `begin`) was
/// just read from `cursor`, consuming through BrtEndAFilter. Returns
/// `kIoXlsbUnknownRecord` for a record outside the covered set and
/// `kIoXlsbRecordCorrupt` for a malformed one.
Expected<AutoFilter, Error> read_auto_filter_block(ByteSpan& cursor, ByteSpan begin);

/// Appends the AUTOFILTER block for `filter` and returns how many of its
/// settings the records cannot hold (an opaque filter, a calendar type,
/// a colour filter without a dxf, XML-only extensions, a phonetic sort
/// method).
std::uint32_t emit_auto_filter_block(std::vector<std::uint8_t>& out, const AutoFilter& filter);

/// Decodes the SORTSTATE block whose BrtBeginSortState (payload `begin`) was
/// just read from `cursor`, consuming through BrtEndSortState.
Expected<SortState, Error> read_sort_state_block(ByteSpan& cursor, ByteSpan begin);

/// Appends the SORTSTATE block for `sort`; returns the settings it cannot hold.
std::uint32_t emit_sort_state_block(std::vector<std::uint8_t>& out, const SortState& sort);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_AUTO_FILTER_RECORDS_H_
