#include <cerrno>
#include <charconv>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <ostream>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "cli/cli.h"
#include "cli/file_io.h"

namespace formulon::cli {
namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
};

struct PaginationGuard {
  fm_pagination_t* handle = nullptr;
  ~PaginationGuard() { fm_pagination_destroy(handle); }
};

void print_paginate_usage(std::ostream& out) {
  out << "Usage: formulon paginate [--sheet INDEX] <in.xlsx>\n"
      << "\n"
      << "Resolve a worksheet's print area, automatic page breaks, and physical page count.\n"
      << "INDEX is zero-based and defaults to 0. Output coordinates are zero-based inclusive.\n";
}

void emit_last_error(std::ostream& err) {
  err << "formulon: paginate: " << fm_last_error_message();
  const char* context = fm_last_error_context();
  if (context != nullptr && context[0] != '\0') {
    err << " (" << context << ')';
  }
  err << '\n';
}

bool parse_sheet_index(std::string_view text, std::size_t& out) {
  std::size_t value = 0;
  const auto [end, ec] = std::from_chars(text.data(), text.data() + text.size(), value);
  if (ec != std::errc{} || end != text.data() + text.size()) {
    return false;
  }
  out = value;
  return true;
}

}  // namespace

int run_paginate(const ArgList& args, std::ostream& out, std::ostream& err) {
  std::string input_path;
  std::size_t sheet_index = 0;
  bool input_seen = false;
  for (std::size_t i = 0; i < args.size(); ++i) {
    const std::string_view arg = args[i];
    if (arg == "-h" || arg == "--help") {
      print_paginate_usage(out);
      return 0;
    }
    if (arg == "--sheet") {
      if (i + 1 == args.size() || !parse_sheet_index(args[i + 1], sheet_index)) {
        err << "formulon: paginate: --sheet requires a non-negative integer\n";
        return kExitUsage;
      }
      ++i;
      continue;
    }
    if (!arg.empty() && arg[0] == '-') {
      err << "formulon: paginate: unknown flag '" << arg << "'\n";
      return kExitUsage;
    }
    if (input_seen) {
      err << "formulon: paginate: only one input path may be supplied\n";
      return kExitUsage;
    }
    input_path.assign(arg);
    input_seen = true;
  }
  if (!input_seen) {
    err << "formulon: paginate: missing input path\n";
    print_paginate_usage(err);
    return kExitUsage;
  }

  std::vector<std::uint8_t> bytes;
  if (const auto status = read_file(input_path, bytes); status != 0) {
    err << "formulon: paginate: cannot read '" << input_path << "': " << std::strerror(errno) << '\n';
    return exit_code_for_status(status);
  }
  WorkbookGuard workbook;
  if (const auto status = fm_workbook_load(bytes.data(), bytes.size(), &workbook.handle); status != 0) {
    emit_last_error(err);
    return exit_code_for_status(status);
  }
  PaginationGuard pagination;
  if (const auto status = fm_workbook_paginate(workbook.handle, sheet_index, &pagination.handle); status != 0) {
    emit_last_error(err);
    return exit_code_for_status(status);
  }

  out << "sheet=" << sheet_index << '\n';
  out << "pages=" << fm_pagination_page_count(pagination.handle) << '\n';
  out << "print_area=";
  for (std::size_t i = 0; i < fm_pagination_print_area_count(pagination.handle); ++i) {
    fm_print_range_t range{};
    if (const auto status = fm_pagination_print_area_at(pagination.handle, i, &range); status != 0) {
      emit_last_error(err);
      return exit_code_for_status(status);
    }
    if (i != 0) {
      out << ',';
    }
    out << range.first_row << ':' << range.first_col << '-' << range.last_row << ':' << range.last_col;
  }
  out << '\n';
  const auto emit_breaks = [&](std::string_view label, auto count_fn, auto at_fn) -> bool {
    out << label << '=';
    for (std::size_t i = 0; i < count_fn(pagination.handle); ++i) {
      uint32_t value = 0;
      if (const auto status = at_fn(pagination.handle, i, &value); status != 0) {
        emit_last_error(err);
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
  if (!emit_breaks("horizontal_breaks", fm_pagination_horizontal_break_count, fm_pagination_horizontal_break_at) ||
      !emit_breaks("vertical_breaks", fm_pagination_vertical_break_count, fm_pagination_vertical_break_at)) {
    return 1;
  }
  out.flush();
  if (!out) {
    err << "formulon: paginate: failed to write output\n";
    return 1;
  }
  return 0;
}

}  // namespace formulon::cli
