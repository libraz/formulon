// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `formulon dump <in.xlsx> [--formulas|--values|--sheets|--metadata]`.
//
// Renders a stable, diff-friendly snapshot of a workbook so it can be
// fed into `diff(1)` between recalc passes. The four flags are
// mutually exclusive; `--formulas` is the default.
//
// Iteration order is sheet index ASC, then cell `(row, col)` ASC. The
// C ABI's `fm_workbook_cell_at` enumerates cells in a stable order
// matching this ordering, so the CLI does not need to sort here.

#include <cerrno>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <fstream>
#include <ios>
#include <ostream>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "cli/cli.h"
#include "cli/render.h"
#include "utils/error.h"

namespace formulon {
namespace cli {

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
};

enum class DumpMode { kFormulas, kValues, kSheets, kMetadata };

void print_dump_usage(std::ostream& out) {
  out << "Usage: formulon dump [--formulas|--values|--sheets|--metadata] <in.xlsx>\n"
      << "\n"
      << "Print a diff-friendly snapshot of <in.xlsx>:\n"
      << "  --formulas  list every formula cell (default)\n"
      << "  --values    list every non-blank cell with its rendered value\n"
      << "  --sheets    list sheet names in document order\n"
      << "  --metadata  list defined names, tables and passthrough parts\n";
}

void emit_last_error(std::ostream& err, const char* subcommand) {
  err << "formulon: " << subcommand << ": " << fm_last_error_message();
  const char* ctx = fm_last_error_context();
  if (ctx != nullptr && ctx[0] != '\0') {
    err << " (" << ctx << ')';
  }
  err << '\n';
}

fm_status_t read_all(const std::string& path, std::vector<std::uint8_t>& out) {
  std::ifstream in(path, std::ios::binary);
  if (!in) {
    return static_cast<fm_status_t>(FormulonErrorCode::kCliFileNotFound);
  }
  in.seekg(0, std::ios::end);
  const std::streamsize size = in.tellg();
  if (size < 0) {
    return static_cast<fm_status_t>(FormulonErrorCode::kCliFileNotFound);
  }
  in.seekg(0, std::ios::beg);
  out.resize(static_cast<std::size_t>(size));
  if (size > 0) {
    in.read(reinterpret_cast<char*>(out.data()), size);
    if (!in) {
      return static_cast<fm_status_t>(FormulonErrorCode::kCliFileNotFound);
    }
  }
  return 0;
}

// Emits every formula cell on every sheet as `Sheet!A1 =formula`.
fm_status_t dump_formulas(const fm_workbook_t* wb, std::ostream& out) {
  const std::size_t n_sheets = fm_workbook_sheet_count(wb);
  for (std::size_t s = 0; s < n_sheets; ++s) {
    const char* sheet_name = nullptr;
    if (auto rc = fm_workbook_sheet_name(wb, s, &sheet_name); rc != 0) {
      return rc;
    }
    std::size_t cell_count = 0;
    if (auto rc = fm_workbook_cell_count(wb, s, &cell_count); rc != 0) {
      return rc;
    }
    for (std::size_t i = 0; i < cell_count; ++i) {
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      const char* formula = nullptr;
      fm_value_t v{};
      if (auto rc = fm_workbook_cell_at(wb, s, i, &row, &col, &formula, &v); rc != 0) {
        return rc;
      }
      if (formula == nullptr) {
        continue;
      }
      out << sheet_name << '!' << format_a1(row, col) << ' ' << formula << '\n';
    }
  }
  return 0;
}

// Emits every non-blank cell on every sheet as `Sheet!A1 <rendered>`.
fm_status_t dump_values(const fm_workbook_t* wb, std::ostream& out) {
  const std::size_t n_sheets = fm_workbook_sheet_count(wb);
  for (std::size_t s = 0; s < n_sheets; ++s) {
    const char* sheet_name = nullptr;
    if (auto rc = fm_workbook_sheet_name(wb, s, &sheet_name); rc != 0) {
      return rc;
    }
    std::size_t cell_count = 0;
    if (auto rc = fm_workbook_cell_count(wb, s, &cell_count); rc != 0) {
      return rc;
    }
    for (std::size_t i = 0; i < cell_count; ++i) {
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      const char* formula = nullptr;
      fm_value_t v{};
      if (auto rc = fm_workbook_cell_at(wb, s, i, &row, &col, &formula, &v); rc != 0) {
        return rc;
      }
      if (v.kind == FM_VAL_BLANK) {
        continue;
      }
      out << sheet_name << '!' << format_a1(row, col) << ' ' << render_value(v) << '\n';
    }
  }
  return 0;
}

// Emits one sheet name per line in document order.
fm_status_t dump_sheets(const fm_workbook_t* wb, std::ostream& out) {
  const std::size_t n_sheets = fm_workbook_sheet_count(wb);
  for (std::size_t s = 0; s < n_sheets; ++s) {
    const char* sheet_name = nullptr;
    if (auto rc = fm_workbook_sheet_name(wb, s, &sheet_name); rc != 0) {
      return rc;
    }
    out << sheet_name << '\n';
  }
  return 0;
}

// Emits defined names, tables, and passthrough parts under
// section headers. Empty sections still emit their header so callers
// can diff "now empty" vs "had entries".
fm_status_t dump_metadata(const fm_workbook_t* wb, std::ostream& out) {
  out << "[defined_names]\n";
  const std::size_t n_names = fm_workbook_defined_name_count(wb);
  for (std::size_t i = 0; i < n_names; ++i) {
    const char* name = nullptr;
    const char* formula = nullptr;
    int32_t local_sheet_id = -1;
    if (auto rc = fm_workbook_defined_name_at_ex(wb, i, &name, &formula, &local_sheet_id); rc != 0) {
      return rc;
    }
    if (local_sheet_id >= 0) {
      const char* sheet_name = nullptr;
      if (auto rc = fm_workbook_sheet_name(wb, static_cast<std::size_t>(local_sheet_id), &sheet_name); rc != 0) {
        return rc;
      }
      out << sheet_name << '!' << name << ' ' << formula << '\n';
    } else {
      out << name << ' ' << formula << '\n';
    }
  }
  out << "[tables]\n";
  const std::size_t n_tables = fm_workbook_table_count(wb);
  for (std::size_t i = 0; i < n_tables; ++i) {
    const char* name = nullptr;
    const char* display = nullptr;
    const char* ref = nullptr;
    std::size_t sheet_index = 0;
    if (auto rc = fm_workbook_table_at(wb, i, &name, &display, &ref, &sheet_index); rc != 0) {
      return rc;
    }
    out << name << ' ' << display << ' ' << sheet_index << ' ' << ref << '\n';
  }
  out << "[passthrough_parts]\n";
  const std::size_t n_parts = fm_workbook_passthrough_count(wb);
  for (std::size_t i = 0; i < n_parts; ++i) {
    const char* path = nullptr;
    if (auto rc = fm_workbook_passthrough_at(wb, i, &path); rc != 0) {
      return rc;
    }
    out << path << '\n';
  }
  return 0;
}

}  // namespace

int run_dump(const ArgList& args, std::ostream& out, std::ostream& err) {
  DumpMode mode = DumpMode::kFormulas;
  bool mode_set = false;
  std::string input_path;
  bool input_seen = false;

  auto select_mode = [&](DumpMode m, std::string_view flag) -> bool {
    if (mode_set) {
      err << "formulon: dump: " << flag << " conflicts with another mode flag\n";
      return false;
    }
    mode = m;
    mode_set = true;
    return true;
  };

  for (std::size_t i = 0; i < args.size(); ++i) {
    const std::string_view a = args[i];
    if (a == "-h" || a == "--help") {
      print_dump_usage(out);
      return 0;
    }
    if (a == "--formulas") {
      if (!select_mode(DumpMode::kFormulas, a)) {
        return kExitUsage;
      }
      continue;
    }
    if (a == "--values") {
      if (!select_mode(DumpMode::kValues, a)) {
        return kExitUsage;
      }
      continue;
    }
    if (a == "--sheets") {
      if (!select_mode(DumpMode::kSheets, a)) {
        return kExitUsage;
      }
      continue;
    }
    if (a == "--metadata") {
      if (!select_mode(DumpMode::kMetadata, a)) {
        return kExitUsage;
      }
      continue;
    }
    if (!a.empty() && a[0] == '-') {
      err << "formulon: dump: unknown flag '" << a << "'\n";
      return kExitUsage;
    }
    if (input_seen) {
      err << "formulon: dump: only one input path may be supplied\n";
      return kExitUsage;
    }
    input_path.assign(a);
    input_seen = true;
  }

  if (!input_seen) {
    err << "formulon: dump: missing input path\n";
    print_dump_usage(err);
    return kExitUsage;
  }

  std::vector<std::uint8_t> bytes;
  if (auto rc = read_all(input_path, bytes); rc != 0) {
    err << "formulon: dump: cannot read '" << input_path << "': " << std::strerror(errno) << '\n';
    return rc;
  }

  WorkbookGuard wb;
  if (auto rc = fm_workbook_load(bytes.data(), bytes.size(), &wb.handle); rc != 0) {
    emit_last_error(err, "dump");
    return rc;
  }

  // `--values` mode runs a recalc so cached values reflect the
  // post-load state. `--formulas` and the metadata views deliberately
  // skip recalc to keep the dump cheap and side-effect-free.
  if (mode == DumpMode::kValues) {
    if (auto rc = fm_workbook_recalc(wb.handle); rc != 0) {
      emit_last_error(err, "dump");
      return rc;
    }
  }

  fm_status_t rc = 0;
  switch (mode) {
    case DumpMode::kFormulas:
      rc = dump_formulas(wb.handle, out);
      break;
    case DumpMode::kValues:
      rc = dump_values(wb.handle, out);
      break;
    case DumpMode::kSheets:
      rc = dump_sheets(wb.handle, out);
      break;
    case DumpMode::kMetadata:
      rc = dump_metadata(wb.handle, out);
      break;
  }
  if (rc != 0) {
    emit_last_error(err, "dump");
    return rc;
  }
  return 0;
}

}  // namespace cli
}  // namespace formulon
