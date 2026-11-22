// Custom table and PivotTable style bookkeeping: kind checks, user counts and removal.

#include <cstddef>
#include <iterator>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "new_workbook_defaults.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_name.h"
#include "styles.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {

namespace {

/// True when the style name `used` resolves to the custom style at `index`.
bool style_name_resolves_to(const TableStylesSection& section, std::size_t index, std::string_view used) {
  return !used.empty() && sheet_names::resolve_named(section.styles, used) == index;
}

}  // namespace

Expected<void, Error> Workbook::check_style_kind(std::string_view name, StyleKind kind, std::string_view field) const {
  const TableStylesSection& section = styles_.table_styles;
  const std::size_t index = sheet_names::resolve_named(section.styles, name);
  const bool excluded = index == sheet_names::kNotFound
                            ? builtin_style_family(name) == BuiltinStyleFamily::kSlicer
                            : !(kind == StyleKind::kTable ? section.styles[index].table : section.styles[index].pivot);
  if (!excluded) {
    return Expected<void, Error>::Ok();
  }
  const bool table = kind == StyleKind::kTable;
  return make_error(FormulonErrorCode::kInvalidArgument,
                    std::string(field) + (table ? " must name a style that applies to tables"
                                                : " must name a style that applies to PivotTables"),
                    std::string(field) + "=" + std::string(name) + (table ? " kind=table" : " kind=pivot"));
}

std::size_t Workbook::style_users(std::size_t index, StyleKind kind) const {
  const TableStylesSection& section = styles_.table_styles;
  std::size_t users = 0;
  if (kind == StyleKind::kTable) {
    for (const TableMetadata& table : tables_) {
      users += style_name_resolves_to(section, index, table.style_info.name) ? 1U : 0U;
    }
    users += style_name_resolves_to(section, index, section.effective_table_default()) ? 1U : 0U;
  } else {
    for (const Sheet& sheet : sheets_) {
      for (const std::unique_ptr<pivot::PivotTable>& pt : sheet.pivot_tables()) {
        users +=
            pt->style_info().has_value() && style_name_resolves_to(section, index, pt->style_info()->name) ? 1U : 0U;
      }
    }
    users += style_name_resolves_to(section, index, section.effective_pivot_default()) ? 1U : 0U;
  }
  return users;
}

Expected<void, Error> Workbook::remove_table_style(std::string_view name) {
  TableStylesSection& section = styles_.table_styles;
  const std::size_t index = sheet_names::resolve_named(section.styles, name);
  if (index == sheet_names::kNotFound) {
    return make_error(FormulonErrorCode::kNotFound, "name names no custom table style", "name=" + std::string(name));
  }
  std::vector<std::size_t> tables_using;
  for (std::size_t t = 0; t < tables_.size(); ++t) {
    if (style_name_resolves_to(section, index, tables_[t].style_info.name)) {
      tables_using.push_back(t);
    }
  }
  std::vector<pivot::PivotTable*> pivots_using;
  for (Sheet& sheet : sheets_) {
    for (const std::unique_ptr<pivot::PivotTable>& pt : sheet.mutable_pivot_tables()) {
      if (pt->style_info().has_value() && style_name_resolves_to(section, index, pt->style_info()->name)) {
        pivots_using.push_back(pt.get());
      }
    }
  }
  if (style_name_resolves_to(section, index, section.default_table_style)) {
    section.default_table_style = std::string(kBuiltinDefaultTableStyle);
  }
  if (style_name_resolves_to(section, index, section.default_pivot_style)) {
    section.default_pivot_style = std::string(kBuiltinDefaultPivotStyle);
  }
  section.styles.erase(section.styles.begin() + static_cast<std::ptrdiff_t>(index));
  const std::string table_default(section.effective_table_default());
  const std::string pivot_default(section.effective_pivot_default());
  for (const std::size_t t : tables_using) {
    tables_[t].style_info.name = table_default;
  }
  for (pivot::PivotTable* pt : pivots_using) {
    pivot::PivotStyleInfo info = *pt->style_info();
    info.name = pivot_default;
    pt->set_style_info(std::move(info));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
