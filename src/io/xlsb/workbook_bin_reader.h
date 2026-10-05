//
// Decoding of the global records of `xl/workbook.bin`: the sheet bundle,
// date system and protection, the `BrtName` table, the supporting-book list
// and the `BrtExternSheet` table that `PtgRef3d` / `PtgNameX` resolve
// through. Consumed by the package reader (`reader.cpp`).

#ifndef FORMULON_IO_XLSB_WORKBOOK_BIN_READER_H_
#define FORMULON_IO_XLSB_WORKBOOK_BIN_READER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "io/xlsb/ptg_reader.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

/// One decoded entry of the workbook's sheet bundle.
///
///   * `name`  — display name of the sheet.
///   * `rid`   — workbook-rels relationship id pointing at the sheet
///               binary part.
struct SheetBundleEntry {
  std::string name;
  std::string rid;
  /// The sheet's `hsState`, which shares its numbering with
  /// `SheetVisibility` and so is carried across whole. Very-hidden stays
  /// distinct from hidden: it is the state that keeps a sheet out of
  /// Excel's "Unhide" dialog.
  SheetVisibility visibility = SheetVisibility::kVisible;
};

/// Workbook-global fields needed while constructing the model.
struct WorkbookBinInfo {
  std::vector<SheetBundleEntry> sheets;
  bool date1904 = false;
  /// `<workbookProtection>` rebuilt from BrtBookProtection(Iso); empty
  /// when the book is unprotected or the records were not decodable.
  std::string protection_xml;
};

/// Decodes the supporting-book list from `xl/workbook.bin` into one
/// entry per book, in `BrtExternSheet`'s `iSupBook` order. The value is
/// `0` for this workbook and `N >= 1` for the N-th external book in
/// package order, which is the number the equivalent xlsx formula text
/// spells as `[N]`.
///
/// Layout verified against Excel-365-produced `xl/workbook.bin` files:
/// the books are the records between `BrtBeginExternals` and
/// `BrtEndExternals`, one record each, and the record id alone says
/// which kind it is. Only `BrtSupSelf` is this workbook; `BrtSupAddin`
/// and `BrtSupSame` are counted so later entries keep their index, but
/// neither names a resolvable external package, so both are reported as
/// external and therefore refused downstream.
///
/// A workbook can legitimately have no self entry: one whose only
/// qualified references are cross-workbook carries a single
/// `BrtSupBookSrc`, so `iSupBook == 0` is then an external book. That
/// case is why this list has to be read rather than assumed.
///
/// A `BrtSupBookSrc` also names the relationship its external link part
/// hangs off, which is how a `[N]` is turned back into a package part.
/// That id is retained on the entry; the other kinds leave it empty.
struct XlsbSupBook {
  std::uint32_t external_book = 0;
  std::string rel_id;
};

/// Decodes `xl/workbook.bin` to extract the ordered sheet-bundle list and
/// workbook date system and protection. Other records are skipped.
Expected<WorkbookBinInfo, Error> DecodeWorkbookBin(const std::vector<std::uint8_t>& body);

/// Decodes the workbook-scope `BrtName` table (declaration order).
Expected<std::vector<XlsbName>, Error> DecodeWorkbookNames(const std::vector<std::uint8_t>& body);

/// Decodes the supporting-book list in `BrtExternSheet`'s `iSupBook` order.
std::vector<XlsbSupBook> DecodeSupBooks(const std::vector<std::uint8_t>& body);

/// Decodes the `BrtExternSheet` table into `(itabFirst, itabLast)` ranges.
Expected<std::vector<XlsbSheetRange>, Error> DecodeExternSheet(const std::vector<std::uint8_t>& body,
                                                               const std::vector<XlsbSupBook>& books);

/// Registers every user-visible `BrtName` entry as a workbook defined name.
Expected<void, Error> RegisterDefinedNames(const std::vector<std::uint8_t>& body,
                                           const std::vector<XlsbName>& name_table,
                                           const std::vector<std::string>& sheet_names,
                                           const std::vector<XlsbSheetRange>& sheet_ranges,
                                           const XlsbExternalBooks& external_books, Workbook& wb,
                                           std::uint32_t* undecoded_defined_name_count);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_WORKBOOK_BIN_READER_H_
