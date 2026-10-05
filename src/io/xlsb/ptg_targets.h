//
// Internal primitives shared by the Ptg encoder (`ptg_writer.cpp`) and the
// workbook-wide name / sheet-range collection pass (`ptg_declarations.cpp`),
// which decide which `BrtName` / `BrtExternSheet` entries a formula needs.

#ifndef FORMULON_IO_XLSB_PTG_TARGETS_H_
#define FORMULON_IO_XLSB_PTG_TARGETS_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/future_functions.h"
#include "sheet_name.h"

namespace formulon {
namespace io {
namespace xlsb {

/// `itabFirst` / `itabLast` of the `BrtExternSheet` entry a `PtgNameX`
/// naming one of this workbook's own defined names resolves through: the
/// entry names no sheet, the record's own scope does.
inline constexpr std::int32_t kXtiNoSheet = -2;

/// True when Excel stores a call to `name` under a hidden `_xlfn.*`
/// BrtName rather than a function id. Classified by
/// `io::classify_storage_prefix`, the same enumeration the OOXML writer
/// consults, so the two formats name the callee identically.
inline bool IsFutureFunction(std::string_view name) {
  return classify_storage_prefix(name) != parser::StoragePrefixKind::None;
}

/// True when a call to `name` is encoded through the hidden-name route
/// (`PtgName` + `PtgFuncVar(255)`) rather than a function id: either a
/// `_xlfn.*` future function, or one of the enumerated names Excel
/// spells bare in OOXML yet has no function id for. Both sets are
/// enumerated from observed Excel output — "absent from
/// `func_id_table`" deliberately does not qualify, because that table
/// grows incrementally and a hidden `_xlfn.<NAME>` Excel does not know
/// resolves to `#NAME?`.
inline bool UsesHiddenNameRoute(std::string_view name) {
  return IsFutureFunction(name) || xlsb_uses_hidden_name(name);
}

/// Resolves a sheet name to its 0-based ixti. Returns -1 when absent.
inline int resolve_ixti(const std::vector<std::string>& sheet_names, std::string_view sheet) {
  for (std::size_t i = 0; i < sheet_names.size(); ++i) {
    if (formulon::sheet_names::equal(sheet_names[i], sheet)) {
      return static_cast<int>(i);
    }
  }
  return -1;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_TARGETS_H_
