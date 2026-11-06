#include <cstddef>
#include <cstdint>
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
  out << "Usage: formulon paginate [--sheet INDEX] [--json] [--] <in | ->\n"
      << "\n"
      << "Resolve a worksheet's print area, automatic page breaks, and physical page count.\n"
      << "INDEX is zero-based and defaults to 0. print_area lists A1 ranges (e.g. B2:D5);\n"
      << "break positions are zero-based. <in> is an xlsx or xlsb file; '-' reads stdin.\n"
      << "  --json       print the sheet's name and one A1 range per physical page as JSON\n"
      << "  --           end options; the next token is the input path\n";
  print_shared_options(out);
}

}  // namespace

int run_paginate(const ArgList& args, std::ostream& out, std::ostream& err) {
  std::string input_path;
  std::size_t sheet_index = 0;
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
      if (i + 1 == args.size() || !parse_unsigned(args[i + 1], &sheet_index)) {
        err << "formulon: paginate: --sheet requires a non-negative integer\n";
        return kExitUsage;
      }
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
  PaginationGuard pagination;
  if (const auto status = fm_sheet_get_pagination(workbook.handle, sheet_index, &pagination.handle); status != 0) {
    emit_last_error(err, "paginate");
    return status;
  }

  std::size_t page_count = 0;
  if (const auto status = fm_pagination_get_page_count(pagination.handle, &page_count); status != 0) {
    emit_last_error(err, "paginate");
    return status;
  }
  if (want_json) {
    const char* sheet_name = nullptr;
    if (const auto status = fm_workbook_get_sheet_name(workbook.handle, sheet_index, &sheet_name); status != 0) {
      emit_last_error(err, "paginate");
      return status;
    }
    std::string doc =
        "{\"format\":1,\"sheets\":[{\"name\":" + render_json_string(view_or_empty(sheet_name)) + ",\"pages\":[";
    for (std::size_t i = 0; i < page_count; ++i) {
      fm_page_layout_t page{};
      if (const auto status = fm_pagination_get_page_at(pagination.handle, i, &page); status != 0) {
        emit_last_error(err, "paginate");
        return status;
      }
      if (i != 0) {
        doc.push_back(',');
      }
      doc += "{\"range\":\"" + format_a1(page.rect.first_row, page.rect.first_col) + ':' +
             format_a1(page.rect.last_row, page.rect.last_col) + "\"}";
    }
    doc += "]}]}\n";
    out << doc;
    return flush_output(out, err, "paginate");
  }
  std::size_t print_area_count = 0;
  if (const auto status = fm_pagination_get_print_area_count(pagination.handle, &print_area_count); status != 0) {
    emit_last_error(err, "paginate");
    return status;
  }
  out << "sheet=" << sheet_index << '\n';
  out << "pages=" << page_count << '\n';
  out << "print_area=";
  for (std::size_t i = 0; i < print_area_count; ++i) {
    fm_cell_rect_t range{};
    if (const auto status = fm_pagination_get_print_area_at(pagination.handle, i, &range); status != 0) {
      emit_last_error(err, "paginate");
      return status;
    }
    if (i != 0) {
      out << ',';
    }
    out << format_a1(range.first_row, range.first_col) << ':' << format_a1(range.last_row, range.last_col);
  }
  out << '\n';
  const auto emit_breaks = [&](std::string_view label, auto count_fn, auto at_fn) -> bool {
    std::size_t count = 0;
    if (const auto status = count_fn(pagination.handle, &count); status != 0) {
      emit_last_error(err, "paginate");
      return false;
    }
    out << label << '=';
    for (std::size_t i = 0; i < count; ++i) {
      uint32_t value = 0;
      if (const auto status = at_fn(pagination.handle, i, &value); status != 0) {
        emit_last_error(err, "paginate");
        return false;
      }
      if (i != 0) {
        out << ',';
      }
      out << value;
    }
    out << '\n';
    return true;
  };
  if (!emit_breaks("horizontal_breaks", fm_pagination_get_row_break_count, fm_pagination_get_row_break_at) ||
      !emit_breaks("vertical_breaks", fm_pagination_get_column_break_count, fm_pagination_get_column_break_at)) {
    return 1;
  }
  return flush_output(out, err, "paginate");
}

}  // namespace formulon::cli
