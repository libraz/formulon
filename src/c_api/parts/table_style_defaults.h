//
// C ABI - the table / PivotTable style a new table or pivot starts with.

#pragma once

#include <string_view>

#include "c_api/formulon_c.h"
#include "workbook.h"

namespace formulon::c_api::parts {

/// The workbook's table-style default, else the desktop application's `TableStyleMedium2`.
std::string_view default_table_style_name(const formulon::Workbook& book);

/// The workbook's PivotTable-style default, else the desktop application's `PivotStyleLight16`.
std::string_view default_pivot_style_name(const formulon::Workbook& book);

/// 0 when the workbook's default style for `kind` may be applied to that kind; else
/// `FM_STATUS_PRECONDITION_FAILED` (a loaded file can name a default its own flags exclude).
fm_status_t check_default_style_kind(const formulon::Workbook& book, formulon::StyleKind kind);

}  // namespace formulon::c_api::parts
