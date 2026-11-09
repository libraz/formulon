//
// `formulon dump <in> [--view formulas|values|sheets|metadata]`.
//
// Renders a stable, diff-friendly snapshot of a workbook so it can be
// fed into `diff(1)` between recalc passes. `--view` picks one of
// four views; `formulas` is the default.
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
#include <optional>
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

/// Which cells the `formulas` / `values` views list.
struct DumpScope {
  std::optional<std::size_t> sheet;
  std::optional<fm_cell_rect_t> range;

  bool covers_sheet(std::size_t s) const { return !sheet.has_value() || *sheet == s; }
  bool covers_cell(std::uint32_t row, std::uint32_t col) const {
    return !range.has_value() ||
           (row >= range->first_row && row <= range->last_row && col >= range->first_col && col <= range->last_col);
  }
};

std::string escape_c_string(const char* text) {
  return escape_single_line(view_or_empty(text));
}

void print_dump_usage(std::ostream& out) {
  out << "Usage: formulon dump [--view formulas|values|sheets|metadata] [--sheet <name | index>] [--range A1:B2]\n"
      << "                     [--json] [--] <in | ->\n"
      << "\n"
      << "Print a diff-friendly snapshot of <in>, an xlsx or xlsb file ('-' reads stdin):\n"
      << "  --view V    what to print (default formulas):\n"
      << "                formulas  list every formula cell\n"
      << "                values    recalculate, then list every non-blank cell with its rendered value\n"
      << "                sheets    list sheet names in document order\n"
      << "                metadata  list defined names, tables and passthrough parts\n"
      << "  --sheet <name | index>  with the formulas or values view: only that sheet (a name or a 0-based index)\n"
      << "  --range R   with the formulas or values view: only the cells in the A1 range R (with --sheet, on that "
         "sheet)\n"
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

// Calls `visit(row, col, formula, value)` for every cell of sheet `s` inside
// `scope`, in the order `fm_sheet_get_cell_at` enumerates them.
template <typename CellVisitor>
fm_status_t for_each_cell(const fm_workbook_t* wb, std::size_t s, const DumpScope& scope, CellVisitor&& visit) {
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
    if (scope.covers_cell(row, col)) {
      visit(row, col, formula, v);
    }
  }
  return 0;
}

// Emits every formula cell (`Sheet!A1 =formula`) or, for `kValues`, every
// non-blank cell (`Sheet!A1 <rendered>`) on every sheet.
fm_status_t dump_cells(const fm_workbook_t* wb, DumpMode mode, const DumpScope& scope, std::ostream& out) {
  const bool formulas = mode == DumpMode::kFormulas;
  return for_each_sheet(wb, [&](std::size_t s, const char* sheet_name) -> fm_status_t {
    if (!scope.covers_sheet(s)) {
      return 0;
    }
    const std::string escaped_sheet_name = escape_c_string(sheet_name);
    return for_each_cell(
        wb, s, scope, [&](std::uint32_t row, std::uint32_t col, const char* formula, const fm_value_t& v) {
          if (formulas ? formula == nullptr : v.kind == FM_VALUE_KIND_BLANK) {
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
fm_status_t dump_cells_json(const fm_workbook_t* wb, const DumpScope& scope, std::ostream& out) {
  std::string doc = "{\"format\":1,\"sheets\":[";
  bool first_sheet = true;
  const fm_status_t rc = for_each_sheet(wb, [&](std::size_t s, const char* sheet_name) -> fm_status_t {
    if (!scope.covers_sheet(s)) {
      return 0;
    }
    if (!first_sheet) {
      doc.push_back(',');
    }
    first_sheet = false;
    doc += "{\"name\":" + render_json_string(view_or_empty(sheet_name)) + ",\"cells\":[";
    bool first_cell = true;
    const fm_status_t cells_rc = for_each_cell(
        wb, s, scope, [&](std::uint32_t row, std::uint32_t col, const char* formula, const fm_value_t& v) {
          if (formula == nullptr && v.kind == FM_VALUE_KIND_BLANK) {
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
    fm_defined_name_t entry = FM_STRUCT_INIT(fm_defined_name_t);
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
  std::size_t n_sheets_for_tables = 0;
  if (auto rc = fm_workbook_get_sheet_count(wb, &n_sheets_for_tables); rc != 0) {
    return rc;
  }
  bool first_table = true;
  for (std::size_t sheet_index = 0; sheet_index < n_sheets_for_tables; ++sheet_index) {
    std::size_t n_tables = 0;
    if (auto rc = fm_sheet_get_table_count(wb, sheet_index, &n_tables); rc != 0) {
      return rc;
    }
    for (std::size_t i = 0; i < n_tables; ++i) {
      fm_table_info_t info = FM_STRUCT_INIT(fm_table_info_t);
      if (auto rc = fm_sheet_get_table_at(wb, sheet_index, i, &info); rc != 0) {
        return rc;
      }
      const char* name = info.name;
      const fm_cell_rect_t rect = info.range;
      // The ref as the table part stores it: a single cell is one address.
      std::string ref = format_a1(rect.first_row, rect.first_col);
      if (rect.first_row != rect.last_row || rect.first_col != rect.last_col) {
        ref += ':' + format_a1(rect.last_row, rect.last_col);
      }
      if (json) {
        doc += (first_table ? "" : ",") + std::string("{\"name\":") + render_json_string(view_or_empty(name)) +
               ",\"sheet\":" + std::to_string(sheet_index) + ",\"ref\":" + render_json_string(ref) + "}";
      } else {
        snapshot << escape_c_string(name) << ' ' << sheet_index << ' ' << escape_c_string(ref.c_str()) << '\n';
      }
      first_table = false;
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

struct DumpView {
  std::string_view name;
  DumpMode mode;
};

constexpr DumpView kDumpViews[] = {
    {"formulas", DumpMode::kFormulas},
    {"values", DumpMode::kValues},
    {"sheets", DumpMode::kSheets},
    {"metadata", DumpMode::kMetadata},
};

}  // namespace

int run_dump(const ArgList& args, std::ostream& out, std::ostream& err) {
  DumpMode mode = DumpMode::kFormulas;
  bool want_json = false;
  std::optional<std::string> sheet_spec;
  std::optional<std::string> range_spec;
  std::string input_path;
  bool input_seen = false;
  bool options_ended = false;
  SharedOptions shared;
  int exit_code = 0;

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
    if (!options_ended && a == "--view") {
      if (i + 1 >= args.size()) {
        err << "formulon: dump: --view requires formulas, values, sheets or metadata\n";
        return kExitUsage;
      }
      const DumpView* view = nullptr;
      for (const DumpView& candidate : kDumpViews) {
        if (args[i + 1] == candidate.name) {
          view = &candidate;
        }
      }
      if (view == nullptr) {
        err << "formulon: dump: --view must be formulas, values, sheets or metadata, not '" << args[i + 1] << "'\n";
        return kExitUsage;
      }
      mode = view->mode;
      ++i;
      continue;
    }
    if (!options_ended && (a == "--sheet" || a == "--range")) {
      if (i + 1 >= args.size()) {
        err << "formulon: dump: " << a << " requires a value\n";
        return kExitUsage;
      }
      (a == "--sheet" ? sheet_spec : range_spec).emplace(args[i + 1]);
      ++i;
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

  if ((sheet_spec.has_value() || range_spec.has_value()) && mode != DumpMode::kFormulas && mode != DumpMode::kValues) {
    err << "formulon: dump: --sheet and --range apply to the formulas and values views only\n";
    return kExitUsage;
  }

  WorkbookGuard wb;
  if (auto rc = load_workbook_file(input_path, shared, wb, err, "dump"); rc != 0) {
    return rc;
  }
  if (auto rc = use_manual_calculation(wb.handle, err, "dump"); rc != 0) {
    return rc;
  }

  DumpScope scope;
  if (sheet_spec.has_value()) {
    std::size_t sheet = 0;
    if (auto rc = resolve_sheet(wb.handle, *sheet_spec, &sheet, err, "dump"); rc != 0) {
      return rc;
    }
    scope.sheet = sheet;
  }
  if (range_spec.has_value()) {
    fm_parsed_reference_info_t ref = FM_STRUCT_INIT(fm_parsed_reference_info_t);
    const fm_status_t parsed = fm_workbook_parse_reference(wb.handle, range_spec->c_str(), nullptr, &ref);
    if (parsed != 0 || ref.sheet_name != nullptr) {
      err << "formulon: dump: --range requires an A1 range without a sheet name, such as A1:B2\n";
      return kExitUsage;
    }
    scope.range = ref.rect;
  }

  // The values view runs a recalc so cached values reflect the
  // post-load state. The formulas and metadata views deliberately
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
      rc = want_json ? dump_cells_json(wb.handle, scope, out) : dump_cells(wb.handle, mode, scope, out);
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
