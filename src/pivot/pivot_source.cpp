// @size-budget: 3 KB
//
// Worksheet-source resolution for a pivot cache refresh. See the header.

#include "pivot/pivot_source.h"

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "auto_filter.h"
#include "defined_name.h"
#include "eval/a1_parse.h"
#include "merge_range.h"
#include "pivot/pivot_cache.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon::pivot {
namespace {

Error unresolved(const char* message) {
  return make_error(FormulonErrorCode::kPreconditionFailed, message);
}

// A qualified or bare A1 range: the sheet it names (empty when bare) and the rectangle.
struct ParsedRange {
  bool valid = false;
  std::string sheet;
  MergeRange rect;
};

ParsedRange parse_range(std::string_view text) {
  if (!text.empty() && text.front() == '=') {
    text.remove_prefix(1);
  }
  ParsedRange out;
  const eval::refs_internal::A1Parse p = eval::refs_internal::parse_a1_ref(text);
  if (!p.valid) {
    return out;
  }
  // The sheet view may point into scratch storage the next parse reuses.
  out.sheet.assign(p.sheet);
  const std::uint32_t row2 = p.is_range ? p.row2 : p.row;
  const std::uint32_t col2 = p.is_range ? p.col2 : p.col;
  out.rect = MergeRange{p.row < row2 ? p.row : row2, p.col < col2 ? p.col : col2, p.row < row2 ? row2 : p.row,
                        p.col < col2 ? col2 : p.col};
  out.valid = true;
  return out;
}

// The defined name `name` in scope `scope` (-1 is the workbook), or null.
const DefinedName* find_name(const std::vector<DefinedName>& names, std::string_view name, std::int32_t scope) {
  const std::size_t index = sheet_names::resolve_by(names.size(), [&](std::size_t i, sheet_names::Fold fold) {
    return names[i].local_sheet_id == scope && sheet_names::same(names[i].name, name, fold);
  });
  return index == sheet_names::kNotFound ? nullptr : &names[index];
}

Expected<WorksheetSourceRange, Error> table_source(const TableMetadata& table) {
  if (!table.header_row) {
    return unresolved("The pivot cache source table has no header row");
  }
  const std::optional<MergeRange> ref = parse_a1_rectangle(table.ref);
  if (!ref.has_value()) {
    return unresolved("The pivot cache source table has no range");
  }
  WorksheetSourceRange out;
  out.sheet = table.sheet_index;
  out.rect = normalized(*ref);
  if (table.totals_row && out.rect.last_row > out.rect.first_row) {
    --out.rect.last_row;
  }
  return out;
}

}  // namespace

Expected<WorksheetSourceRange, Error> resolve_worksheet_source(const Workbook& wb, const PivotCache& cache) {
  const WorksheetSource& source = cache.worksheet_source();
  if (!source.present) {
    return unresolved("The pivot cache has no worksheet source");
  }
  for (const auto& attr : cache.cache_source_attrs()) {
    if (attr.first == "type" && attr.second != "worksheet") {
      return unresolved("The pivot cache source is not a worksheet");
    }
  }
  for (const std::string* text : {&source.ref, &source.sheet, &source.name}) {
    if (text->find('[') != std::string::npos) {
      return unresolved("The pivot cache source is in another workbook");
    }
  }

  if (!source.name.empty()) {
    const std::size_t table = sheet_names::resolve_table(wb.tables(), source.name);
    if (table != sheet_names::kNotFound) {
      return table_source(wb.tables()[table]);
    }
    const DefinedName* name = nullptr;
    if (!source.sheet.empty()) {
      const std::size_t local = wb.sheet_index_by_name(source.sheet);
      if (local != Workbook::kNoSheet) {
        name = find_name(wb.defined_names(), source.name, static_cast<std::int32_t>(local));
      }
    }
    if (name == nullptr) {
      name = find_name(wb.defined_names(), source.name, -1);
    }
    if (name == nullptr) {
      return unresolved("The pivot cache source names no table or defined name");
    }
    const ParsedRange range = parse_range(name->formula);
    if (!range.valid || range.sheet.empty()) {
      return unresolved("The pivot cache source name does not refer to one range on one sheet");
    }
    const std::size_t sheet = wb.sheet_index_by_name(range.sheet);
    if (sheet == Workbook::kNoSheet) {
      return unresolved("The pivot cache source name refers to a sheet that does not exist");
    }
    return WorksheetSourceRange{sheet, range.rect};
  }

  const ParsedRange range = parse_range(source.ref);
  if (!range.valid) {
    return unresolved("The pivot cache source range is not an A1 range");
  }
  const std::string& sheet_name = range.sheet.empty() ? source.sheet : range.sheet;
  const std::size_t sheet = sheet_name.empty() ? Workbook::kNoSheet : wb.sheet_index_by_name(sheet_name);
  if (sheet == Workbook::kNoSheet) {
    return unresolved("The pivot cache source sheet does not exist");
  }
  return WorksheetSourceRange{sheet, range.rect};
}

}  // namespace formulon::pivot
