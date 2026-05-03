// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Reader for `<conditionalFormatting>` blocks inside an OOXML sheet
// (`xl/worksheets/sheet*.xml`). Walks the worksheet DOM and decodes
// every CF block into `cf::ConditionalFormat` records. Does NOT consult
// `<extLst>` for x14 modern overlays — that pass is layered on top in a
// follow-up PR; the legacy DOM contains enough information for the
// majority of authored workbooks.
//
// Design references:
//   * backup/plans/20-conditional-format-deep.md §20.2 / §20.3
//   * src/io/pivot_table_reader.h (sister reader, similar style)

#ifndef FORMULON_IO_CF_READER_H_
#define FORMULON_IO_CF_READER_H_

#include <vector>

#include "cf/cf_types.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {

/// Walks a `<worksheet>` node and returns every `<conditionalFormatting>`
/// block parsed into a `cf::ConditionalFormat`. Returns an empty vector
/// when the sheet has no CF blocks.
///
/// Behaviour:
///   * `<conditionalFormatting sqref="A1:A10 D5:D15" pivot="0">` —
///     `sqref` is split on whitespace and each token decoded into a
///     `CFCellRange`. `pivot="1"` sets `pivot_scope`.
///   * Unparseable / empty `sqref` is `kIoSheetCorrupt`.
///   * `<cfRule type="..." priority="N" stopIfTrue="..." dxfId="...">`
///     attributes feed the corresponding `CFRule` fields. Unknown
///     `type` attributes fold the rule to `Expression` and continue
///     (forward-compat with future Excel additions); the `formula1` is
///     captured if present.
///   * `<formula>` children populate `formula1` / `formula2` in
///     document order; the leading `=` is stripped if present.
///   * `<colorScale>`, `<dataBar>`, `<iconSet>` children populate the
///     respective payload optionals. `<cfvo>` thresholds and `<color
///     rgb="...">` entries are walked in document order.
///   * Visual rule subtree errors fold to defaults (no hard fail) so a
///     malformed colour scale does not reject the entire sheet.
///
/// Errors:
///   * `kIoSheetCorrupt` — sqref missing/empty/unparseable.
Expected<std::vector<cf::ConditionalFormat>, Error> read_conditional_formats(const pugi::xml_node& worksheet);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_READER_H_
