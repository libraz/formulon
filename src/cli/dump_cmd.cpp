//
// `formulon dump <in> [--formulas|--values|--sheets|--metadata]`.
//
// Renders a stable, diff-friendly snapshot of a workbook so it can be
// fed into `diff(1)` between recalc passes. The four flags are
// mutually exclusive; `--formulas` is the default.
//
// Iteration order is sheet index ASC, then cell `(row, col)` ASC. The
// C ABI's `fm_sheet_get_cell_at` enumerates cells in a stable order
// matching this ordering, so the CLI does not need to sort here.
//
// Anything the load could not represent is warned about on stderr before
// the snapshot is written, so a diff is never taken on a workbook that
// quietly lost content.

#include <cstddef>
#include <cstdint>
#include <ostream>
#include <sstream>
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
#include "utils/error.h"

namespace formulon {
namespace cli {

namespace {

enum class DumpMode { kFormulas, kValues, kSheets, kMetadata };

std::string escape_c_string(const char* text) {
  return escape_single_line(view_or_empty(text));
}

void print_dump_usage(std::ostream& out) {
  out << "Usage: formulon dump [--formulas|--values|--sheets|--metadata] [--json] [--] <in | ->\n"
      << "\n"
      << "Print a diff-friendly snapshot of <in>, an xlsx or xlsb file ('-' reads stdin):\n"
      << "  --formulas  list every formula cell (default)\n"
      << "  --values    recalculate, then list every non-blank cell with its rendered value\n"
      << "  --sheets    list sheet names in document order\n"
      << "  --metadata  list defined names, tables and passthrough parts\n"
      << "  --json      print the selected view as one JSON object: {\"format\":1,...}\n"
      << "  --           end options; the next token is the input path\n"
      << "Plain output escapes only \\n, \\r, \\t and \\\\.\n";
  print_shared_options(out);
}

// Calls `visit(sheet_index, sheet_name)` for every sheet in document order,
// stopping at the first non-zero status from a C call or from `visit`.
template <typename SheetVisitor>
fm_status_t for_each_sheet(const fm_workbook_t* wb, SheetVisitor&& visit) {
  std::size_t n_sheets = 0;
  if (auto rc = fm_workbook_get_sheet_count(wb, &n_sheets); rc != 0) {
    return rc;
  }
  for (std::size_t s = 0; s < n_sheets; ++s) {
    const char* sheet_name = nullptr;
    if (auto rc = fm_workbook_get_sheet_name(wb, s, &sheet_name); rc != 0) {
      return rc;
    }
    if (auto rc = visit(s, sheet_name); rc != 0) {
      return rc;
    }
  }
  return 0;
}

// Calls `visit(row, col, formula, value)` for every cell of sheet `s` in the
// order `fm_sheet_get_cell_at` enumerates them.
template <typename CellVisitor>
fm_status_t for_each_cell(const fm_workbook_t* wb, std::size_t s, CellVisitor&& visit) {
  std::size_t cell_count = 0;
  if (auto rc = fm_sheet_get_cell_count(wb, s, &cell_count); rc != 0) {
    return rc;
  }
  for (std::size_t i = 0; i < cell_count; ++i) {
    std::uint32_t row = 0;
    std::uint32_t col = 0;
    const char* formula = nullptr;
    fm_value_t v{};
    if (auto rc = fm_sheet_get_cell_at(wb, s, i, &row, &col, &formula, &v); rc != 0) {
      return rc;
    }
    visit(row, col, formula, v);
  }
  return 0;
}

// Emits every formula cell (`Sheet!A1 =formula`) or, for `kValues`, every
// non-blank cell (`Sheet!A1 <rendered>`) on every sheet.
fm_status_t dump_cells(const fm_workbook_t* wb, DumpMode mode, std::ostream& out) {
  const bool formulas = mode == DumpMode::kFormulas;
  return for_each_sheet(wb, [&](std::size_t s, const char* sheet_name) -> fm_status_t {
    const std::string escaped_sheet_name = escape_c_string(sheet_name);
    return for_each_cell(wb, s, [&](std::uint32_t row, std::uint32_t col, const char* formula, const fm_value_t& v) {
      if (formulas ? formula == nullptr : v.kind == FM_VAL_BLANK) {
        return;
      }
      out << escaped_sheet_name << '!' << format_a1(row, col) << ' '
          << (formulas ? escape_single_line(formula) : escape_single_line(render_value(v))) << '\n';
    });
  });
}

// Emits `{"format":1,"sheets":[{"name":...,"cells":[...]}]}` as one document.
// A cell is listed when it holds a formula or a non-blank value; `kind` and
// `value` are the `eval --json` rendering, and `formula` appears on formula
// cells only. The document is assembled before anything is written so a
// failure leaves `out` untouched.
fm_status_t dump_cells_json(const fm_workbook_t* wb, std::ostream& out) {
  std::string doc = "{\"format\":1,\"sheets\":[";
  const fm_status_t rc = for_each_sheet(wb, [&](std::size_t s, const char* sheet_name) -> fm_status_t {
    if (s != 0) {
      doc.push_back(',');
    }
    doc += "{\"name\":" + render_json_string(view_or_empty(sheet_name)) + ",\"cells\":[";
    bool first_cell = true;
    const fm_status_t cells_rc =
        for_each_cell(wb, s, [&](std::uint32_t row, std::uint32_t col, const char* formula, const fm_value_t& v) {
          if (formula == nullptr && v.kind == FM_VAL_BLANK) {
            return;
          }
          if (!first_cell) {
            doc.push_back(',');
          }
          first_cell = false;
          doc += "{\"ref\":\"" + format_a1(row, col) + "\",";
          append_value_json_members(doc, v);
          if (formula != nullptr) {
            doc += ",\"formula\":" + render_json_string(formula);
          }
          doc.push_back('}');
        });
    if (cells_rc != 0) {
      return cells_rc;
    }
    doc += "]}";
    return 0;
  });
  if (rc != 0) {
    return rc;
  }
  doc += "]}\n";
  out << doc;
  return 0;
}

// Emits one sheet name per line in document order, or with `json`
// `{"format":1,"sheets":[{"name":...}]}`.
fm_status_t dump_sheets(const fm_workbook_t* wb, bool json, std::ostream& out) {
  std::string doc = "{\"format\":1,\"sheets\":[";
  const fm_status_t rc = for_each_sheet(wb, [&](std::size_t s, const char* sheet_name) -> fm_status_t {
    if (json) {
      doc += (s != 0 ? "," : "") + std::string("{\"name\":") + render_json_string(view_or_empty(sheet_name)) + "}";
    } else {
      out << escape_c_string(sheet_name) << '\n';
    }
    return 0;
  });
  if (rc == 0 && json) {
    out << doc << "]}\n";
  }
  return rc;
}

// Emits defined names, tables, and passthrough parts under
// section headers. Empty sections still emit their header so callers
// can diff "now empty" vs "had entries". With `json` the same three sections
// are the members of one `{"format":1,...}` object; a name's `scope` is its
// sheet name, `#<id>` for a sheet that does not exist, or null for workbook scope.
fm_status_t dump_metadata(const fm_workbook_t* wb, bool json, std::ostream& out) {
  std::ostringstream snapshot;
  std::string doc = "{\"format\":1,\"defined_names\":[";
  std::size_t n_sheets = 0;
  if (auto rc = fm_workbook_get_sheet_count(wb, &n_sheets); rc != 0) {
    return rc;
  }

  snapshot << "[defined_names]\n";
  std::size_t n_names = 0;
  if (auto rc = fm_workbook_get_defined_name_count(wb, &n_names); rc != 0) {
    return rc;
  }
  for (std::size_t i = 0; i < n_names; ++i) {
    fm_defined_name_t entry{};
    if (auto rc = fm_workbook_get_defined_name_at(wb, i, &entry); rc != 0) {
      return rc;
    }
    const int32_t local_sheet_id = entry.local_sheet_id;
    std::string scope;
    if (local_sheet_id >= 0 && static_cast<std::size_t>(local_sheet_id) < n_sheets) {
      const char* sheet_name = nullptr;
      if (auto rc = fm_workbook_get_sheet_name(wb, static_cast<std::size_t>(local_sheet_id), &sheet_name); rc != 0) {
        return rc;
      }
      scope = std::string(view_or_empty(sheet_name));
    } else if (local_sheet_id >= 0) {
      scope = '#' + std::to_string(local_sheet_id);
    }
    if (json) {
      doc += (i != 0 ? "," : "") + std::string("{\"name\":") + render_json_string(view_or_empty(entry.name)) +
             ",\"formula\":" + render_json_string(view_or_empty(entry.formula)) +
             ",\"scope\":" + (local_sheet_id >= 0 ? render_json_string(scope) : std::string("null")) + "}";
    } else {
      if (local_sheet_id >= 0) {
        snapshot << escape_single_line(scope) << '!';
      }
      snapshot << escape_c_string(entry.name) << ' ' << escape_c_string(entry.formula) << '\n';
    }
  }
  snapshot << "[tables]\n";
  doc += "],\"tables\":[";
  std::size_t n_tables = 0;
  if (auto rc = fm_workbook_get_table_count(wb, &n_tables); rc != 0) {
    return rc;
  }
  for (std::size_t i = 0; i < n_tables; ++i) {
    const char* name = nullptr;
    const char* display = nullptr;
    fm_cell_rect_t rect{};
    std::size_t sheet_index = 0;
    if (auto rc = fm_workbook_get_table(wb, i, &name, &display, &rect, &sheet_index); rc != 0) {
      return rc;
    }
    // The ref as the table part stores it: a single cell is one address.
    std::string ref = format_a1(rect.first_row, rect.first_col);
    if (rect.first_row != rect.last_row || rect.first_col != rect.last_col) {
      ref += ':' + format_a1(rect.last_row, rect.last_col);
    }
    if (json) {
      doc += (i != 0 ? "," : "") + std::string("{\"name\":") + render_json_string(view_or_empty(name)) +
             ",\"display_name\":" + render_json_string(view_or_empty(display)) +
             ",\"sheet\":" + std::to_string(sheet_index) + ",\"ref\":" + render_json_string(ref) + "}";
    } else {
      snapshot << escape_c_string(name) << ' ' << escape_c_string(display) << ' ' << sheet_index << ' '
               << escape_c_string(ref.c_str()) << '\n';
    }
  }
  snapshot << "[passthrough_parts]\n";
  doc += "],\"passthrough_parts\":[";
  std::size_t n_parts = 0;
  if (auto rc = fm_workbook_get_passthrough_count(wb, &n_parts); rc != 0) {
    return rc;
  }
  for (std::size_t i = 0; i < n_parts; ++i) {
    const char* path = nullptr;
    if (auto rc = fm_workbook_get_passthrough_at(wb, i, &path); rc != 0) {
      return rc;
    }
    if (json) {
      doc += (i != 0 ? "," : "") + render_json_string(view_or_empty(path));
    } else {
      snapshot << escape_c_string(path) << '\n';
    }
  }
  if (json) {
    out << doc << "]}\n";
  } else {
    out << snapshot.str();
  }
  return 0;
}

struct DumpModeFlag {
  std::string_view flag;
  DumpMode mode;
};

constexpr DumpModeFlag kDumpModeFlags[] = {
    {"--formulas", DumpMode::kFormulas},
    {"--values", DumpMode::kValues},
    {"--sheets", DumpMode::kSheets},
    {"--metadata", DumpMode::kMetadata},
};

}  // namespace

int run_dump(const ArgList& args, std::ostream& out, std::ostream& err) {
  DumpMode mode = DumpMode::kFormulas;
  bool mode_set = false;
  bool want_json = false;
  std::string input_path;
  bool input_seen = false;
  bool options_ended = false;
  SharedOptions shared;
  int exit_code = 0;

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
    switch (handle_common_option(args, i, options_ended, shared, print_dump_usage, "dump", out, err, exit_code)) {
      case CommonOption::kConsumed:
        continue;
      case CommonOption::kExit:
        return exit_code;
      case CommonOption::kNotCommon:
        break;
    }
    if (!options_ended && a == "--json") {
      want_json = true;
      continue;
    }
    const DumpModeFlag* mode_flag = nullptr;
    for (const DumpModeFlag& candidate : kDumpModeFlags) {
      if (!options_ended && a == candidate.flag) {
        mode_flag = &candidate;
      }
    }
    if (mode_flag != nullptr) {
      if (!select_mode(mode_flag->mode, a)) {
        return kExitUsage;
      }
      continue;
    }
    if (!accept_positional(a, options_ended, input_seen, "dump", "input path", err, /*allow_dash=*/true)) {
      return kExitUsage;
    }
    input_path.assign(a);
    input_seen = true;
  }

  if (!input_seen) {
    return missing_positional("dump", "input path", print_dump_usage, err);
  }

  WorkbookGuard wb;
  if (auto rc = load_workbook_file(input_path, shared, wb, err, "dump"); rc != 0) {
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
    case DumpMode::kValues:
      rc = want_json ? dump_cells_json(wb.handle, out) : dump_cells(wb.handle, mode, out);
      break;
    case DumpMode::kSheets:
      rc = dump_sheets(wb.handle, want_json, out);
      break;
    case DumpMode::kMetadata:
      rc = dump_metadata(wb.handle, want_json, out);
      break;
  }
  if (rc != 0) {
    emit_last_error(err, "dump");
    return rc;
  }
  return flush_output(out, err, "dump");
}

}  // namespace cli
}  // namespace formulon
