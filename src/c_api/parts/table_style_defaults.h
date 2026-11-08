//
// C ABI - the table / PivotTable style a new table or pivot starts with.

#pragma once

#include <string_view>

#include "workbook.h"

namespace formulon::c_api::parts {

/// The workbook's table-style default, else the desktop application's `TableStyleMedium2`.
std::string_view default_table_style_name(const formulon::Workbook& book);

/// The workbook's PivotTable-style default, else the desktop application's `PivotStyleLight16`.
std::string_view default_pivot_style_name(const formulon::Workbook& book);

}  // namespace formulon::c_api::parts
