#include <cstddef>
#include <cstdint>
#include <optional>
#include <ostream>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "cli/c_api_raii.h"
#include "cli/cli.h"
#include "cli/diagnostics.h"
#include "cli/file_io.h"
#include "cli/render.h"
#include "cli/workbook_load.h"

namespace formulon::cli {
namespace {

void print_paginate_usage(std::ostream& out) {
  out << "Usage: formulon paginate [--sheet <name | index>] [--json] [--] <in | ->\n"
      << "\n"
      << "Resolve a worksheet's print area, automatic page breaks, and physical page count.\n"
      << "--sheet takes a sheet name or a zero-based index and defaults to the first sheet. printArea lists A1 ranges "
         "(e.g. B2:D5);\n"
      << "break positions are zero-based. <in> is an xlsx or xlsb file; '-' reads stdin.\n"
      << "  --json       print the same keys as one JSON object {\"format\":1,\"sheets\":[...]}; \"format\" is the "
         "schema version\n"
      << "  --           end options; the next token is the input path\n";
  print_shared_options(out);
}

}  // namespace

int run_paginate(const ArgList& args, std::ostream& out, std::ostream& err) {
  std::string input_path;
  std::optional<std::string> sheet_spec;
  bool input_seen = false;
  bool options_ended = false;
  bool want_json = false;
  SharedOptions shared;
  int exit_code = 0;
  for (std::size_t i = 0; i < args.size(); ++i) {
    const std::string_view arg = args[i];
    switch (
        handle_common_option(args, i, options_ended, shared, print_paginate_usage, "paginate", out, err, exit_code)) {
      case CommonOption::kConsumed:
        continue;
      case CommonOption::kExit:
        return exit_code;
      case CommonOption::kNotCommon:
        break;
    }
    if (!options_ended && arg == "--json") {
      want_json = true;
      continue;
    }
    if (!options_ended && arg == "--sheet") {
      if (i + 1 == args.size()) {
        err << "formulon: paginate: --sheet requires a sheet name or index\n";
        return kExitUsage;
      }
      sheet_spec.emplace(args[i + 1]);
      ++i;
      continue;
    }
    if (!accept_positional(arg, options_ended, input_seen, "paginate", "input path", err, /*allow_dash=*/true)) {
      return kExitUsage;
    }
    input_path.assign(arg);
    input_seen = true;
  }
  if (!input_seen) {
    return missing_positional("paginate", "input path", print_paginate_usage, err);
  }

  WorkbookGuard workbook;
  if (const auto status = load_workbook_file(input_path, shared, workbook, err, "paginate"); status != 0) {
    return status;
  }
  if (auto rc = use_manual_calculation(workbook.handle, err, "paginate"); rc != 0) {
    return rc;
  }
  std::size_t sheet_index = 0;
  if (sheet_spec.has_value()) {
    if (auto rc = resolve_sheet(workbook.handle, *sheet_spec, &sheet_index, err, "paginate"); rc != 0) {
      return rc;
    }
  }
  PaginationGuard pagination;
  if (const auto status = fm_sheet_paginate(workbook.handle, sheet_index, &pagination.handle); status != 0) {
    emit_last_error(err, "paginate");
    return status;
  }

  std::size_t page_count = 0;
  if (const auto status = fm_pagination_get_page_count(pagination.handle, &page_count); status != 0) {
    emit_last_error(err, "paginate");
    return status;
  }
  const char* sheet_name = nullptr;
  if (const auto status = fm_workbook_get_sheet_name(workbook.handle, sheet_index, &sheet_name); status != 0) {
    emit_last_error(err, "paginate");
    return status;
  }
  std::size_t print_area_count = 0;
  if (const auto status = fm_pagination_get_print_area_count(pagination.handle, &print_area_count); status != 0) {
    emit_last_error(err, "paginate");
    return status;
  }
  std::vector<std::string> print_areas;
  for (std::size_t i = 0; i < print_area_count; ++i) {
    fm_cell_rect_t range{};
    if (const auto status = fm_pagination_get_print_area_at(pagination.handle, i, &range); status != 0) {
      emit_last_error(err, "paginate");
      return status;
    }
    print_areas.push_back(format_a1(range.first_row, range.first_col) + ':' +
                          format_a1(range.last_row, range.last_col));
  }
  const auto collect_breaks = [&](std::vector<std::string>& positions, auto record, auto position, auto count_fn,
                                  auto at_fn) -> bool {
    std::size_t count = 0;
    if (const auto status = count_fn(pagination.handle, &count); status != 0) {
      emit_last_error(err, "paginate");
      return false;
    }
    for (std::size_t i = 0; i < count; ++i) {
      auto entry = record;
      if (const auto status = at_fn(pagination.handle, i, &entry); status != 0) {
        emit_last_error(err, "paginate");
        return false;
      }
      positions.push_back(std::to_string(position(entry)));
    }
    return true;
  };
  std::vector<std::string> row_breaks;
  std::vector<std::string> column_breaks;
  if (!collect_breaks(
          row_breaks, FM_STRUCT_INIT(fm_row_break_t), [](const fm_row_break_t& b) { return b.row; },
          fm_pagination_get_row_break_count, fm_pagination_get_row_break_at) ||
      !collect_breaks(
          column_breaks, FM_STRUCT_INIT(fm_column_break_t), [](const fm_column_break_t& b) { return b.col; },
          fm_pagination_get_column_break_count, fm_pagination_get_column_break_at)) {
    return kExitFailure;
  }
  std::vector<std::string> pages;
  for (std::size_t i = 0; i < page_count; ++i) {
    fm_page_layout_info_t page = FM_STRUCT_INIT(fm_page_layout_info_t);
    if (const auto status = fm_pagination_get_page_at(pagination.handle, i, &page); status != 0) {
      emit_last_error(err, "paginate");
      return status;
    }
    pages.push_back(format_a1(page.rect.first_row, page.rect.first_col) + ':' +
                    format_a1(page.rect.last_row, page.rect.last_col));
  }

  const auto join = [](const std::vector<std::string>& items, bool quote) {
    std::string joined;
    for (std::size_t i = 0; i < items.size(); ++i) {
      if (i != 0) {
        joined.push_back(',');
      }
      joined += quote ? '"' + items[i] + '"' : items[i];
    }
    return joined;
  };
  if (want_json) {
    std::string page_objects;
    for (std::size_t i = 0; i < pages.size(); ++i) {
      page_objects += (i != 0 ? ",{\"range\":\"" : "{\"range\":\"") + pages[i] + "\"}";
    }
    out << "{\"format\":1,\"sheets\":[{\"name\":" << render_json_string(view_or_empty(sheet_name))
        << ",\"pageCount\":" << page_count << ",\"printArea\":[" << join(print_areas, true) << "],\"rowBreaks\":["
        << join(row_breaks, false) << "],\"columnBreaks\":[" << join(column_breaks, false) << "],\"pages\":["
        << page_objects << "]}]}\n";
    return flush_output(out, err, "paginate");
  }
  out << "name=" << view_or_empty(sheet_name) << '\n'
      << "pageCount=" << page_count << '\n'
      << "printArea=" << join(print_areas, false) << '\n'
      << "rowBreaks=" << join(row_breaks, false) << '\n'
      << "columnBreaks=" << join(column_breaks, false) << '\n'
      << "pages=" << join(pages, false) << '\n';
  return flush_output(out, err, "paginate");
}

}  // namespace formulon::cli
