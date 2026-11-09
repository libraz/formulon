//
// `formulon recalc <in.xlsx> -o <out.xlsx>` — load, recalc, save.
//
// File I/O happens in the CLI layer (not the C ABI) because the engine
// stays bytes-in / bytes-out for portability with WASM and language
// bindings. The CLI handler:
//
//   1. Reads `in` into memory (binary mode).
//   2. Calls `fm_workbook_load`, then the serial `fm_workbook_recalc` or
//      opt-in `fm_workbook_recalc_parallel`, then
//      `fm_workbook_save_with_diagnostics` with a format derived from
//      `-o`'s extension (`.xlsx`, `.xlsb`) or `--format`.
//   3. Writes the saved buffer back out to `out`.
//
// `--iterative` enables iterative-calc while preserving any iteration cap
// and convergence threshold loaded from the workbook. `--threads N` opts
// into the parallel scheduler; omitting it deliberately preserves the
// serial `recalc` contract. `--quiet` suppresses the `Recalculated ...`
// status line on stderr; otherwise we print one per successful run.

#include <cerrno>
#include <cstddef>
#include <cstdint>
#include <cstring>
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
#include "cli/workbook_load.h"
#include "utils/strings.h"

namespace formulon {
namespace cli {

namespace {

/// Largest worker count `--threads` accepts.
constexpr std::uint32_t kMaxRecalcThreads = 8U;

void print_recalc_usage(std::ostream& out) {
  out << "Usage: formulon recalc [--iterative] [--threads N] [--quiet] [--format xlsx|xlsb] <in.xlsx-or-xlsb | -> -o "
         "<out | ->\n"
      << "       formulon recalc [--iterative] [--threads N] [--quiet] [--format xlsx|xlsb] -o <out> -- "
         "<in.xlsx-or-xlsb>\n"
      << "\n"
      << "Load <in.xlsx-or-xlsb>, drive a full recalc, and write the result to <out.xlsx-or-xlsb>.\n"
      << "The output format is chosen from -o's extension: '.xlsx' writes OOXML, '.xlsb' writes MS-XLSB.\n"
      << "Any other extension needs --format xlsx|xlsb (exit 64 without it); --format must agree with\n"
      << "a .xlsx or .xlsb extension. '-o -' writes .xlsx unless --format says otherwise.\n"
      << "--threads N opts into parallel recalc: 0 selects automatic detection (capped at " << kMaxRecalcThreads
      << "),\n"
      << "1 stays on the caller thread, and 2.." << kMaxRecalcThreads << " select the worker cap. Without --threads,\n"
      << "recalc remains serial.\n"
      << "Options must precede --; after --, exactly one input path is accepted.\n"
      << "An input of '-' reads the workbook from stdin; '-o -' writes it to stdout as .xlsx unless --format says "
         "otherwise\n"
      << "(refused when stdout is a terminal). Use './-' for a file named '-'.\n"
      << "Status: prints \"formulon: recalc: ok, wrote M bytes to 'OUT'\" on stderr unless\n"
      << "--quiet is supplied; --quiet does not suppress load/save loss warnings.\n";
  print_shared_options(out);
}

bool parse_thread_count(std::string_view text, std::uint32_t* out) {
  std::uint32_t parsed = 0U;
  if (!parse_unsigned(text, &parsed) || parsed > kMaxRecalcThreads) {
    return false;
  }
  *out = parsed;
  return true;
}

// Returns the container `path`'s extension names (`.xlsx` or `.xlsb`,
// case-insensitive), or nothing for any other extension.
std::optional<fm_workbook_format_t> format_from_extension(const std::string& path) {
  const std::size_t dot = path.find_last_of('.');
  if (dot == std::string::npos) {
    return std::nullopt;
  }
  const std::string_view extension = std::string_view(path).substr(dot + 1);
  if (strings::case_insensitive_eq(extension, "xlsx")) {
    return FM_WORKBOOK_FORMAT_XLSX;
  }
  if (strings::case_insensitive_eq(extension, "xlsb")) {
    return FM_WORKBOOK_FORMAT_XLSB;
  }
  return std::nullopt;
}

}  // namespace

int run_recalc(const ArgList& args, std::ostream& out, std::ostream& err) {
  std::string input_path;
  std::string output_path;
  bool iterative = false;
  bool quiet = false;
  bool parallel = false;
  std::uint32_t thread_count = 0U;
  std::optional<fm_workbook_format_t> format_option;
  bool input_seen = false;
  bool output_seen = false;
  bool options_ended = false;
  SharedOptions shared;
  int exit_code = 0;

  for (std::size_t i = 0; i < args.size(); ++i) {
    const std::string_view a = args[i];
    switch (handle_common_option(args, i, options_ended, shared, print_recalc_usage, "recalc", out, err, exit_code)) {
      case CommonOption::kConsumed:
        continue;
      case CommonOption::kExit:
        return exit_code;
      case CommonOption::kNotCommon:
        break;
    }
    if (!options_ended && a == "--iterative") {
      iterative = true;
      continue;
    }
    if (!options_ended && a == "--quiet") {
      quiet = true;
      continue;
    }
    if (!options_ended && a == "--threads") {
      if (i + 1 >= args.size()) {
        err << "formulon: recalc: --threads requires an integer from 0 to " << kMaxRecalcThreads << "\n";
        return kExitUsage;
      }
      if (!parse_thread_count(args[i + 1], &thread_count)) {
        err << "formulon: recalc: --threads must be an integer from 0 to " << kMaxRecalcThreads << "\n";
        return kExitUsage;
      }
      parallel = true;
      ++i;
      continue;
    }
    if (!options_ended && a == "--format") {
      if (i + 1 >= args.size()) {
        err << "formulon: recalc: --format requires xlsx or xlsb\n";
        return kExitUsage;
      }
      if (args[i + 1] == "xlsx") {
        format_option = FM_WORKBOOK_FORMAT_XLSX;
      } else if (args[i + 1] == "xlsb") {
        format_option = FM_WORKBOOK_FORMAT_XLSB;
      } else {
        err << "formulon: recalc: --format must be xlsx or xlsb, not '" << args[i + 1] << "'\n";
        return kExitUsage;
      }
      ++i;
      continue;
    }
    if (!options_ended && (a == "-o" || a == "--output")) {
      if (i + 1 >= args.size()) {
        err << "formulon: recalc: " << a << " requires a path argument\n";
        return kExitUsage;
      }
      output_path.assign(args[i + 1]);
      output_seen = true;
      ++i;
      continue;
    }
    if (!accept_positional(a, options_ended, input_seen, "recalc", "input path", err, /*allow_dash=*/true)) {
      return kExitUsage;
    }
    input_path.assign(a);
    input_seen = true;
  }

  if (!input_seen) {
    return missing_positional("recalc", "input path", print_recalc_usage, err);
  }
  // "Never supplied" and "supplied as an empty string" are different
  // defects — the second usually means an unset shell variable expanded
  // into the value slot — so they do not share one sentinel. `--threads`
  // already draws the same distinction.
  if (!output_seen) {
    err << "formulon: recalc: missing -o/--output PATH\n";
    return kExitUsage;
  }
  if (output_path.empty()) {
    err << "formulon: recalc: -o/--output requires a non-empty path\n";
    return kExitUsage;
  }

  const bool to_stdout = output_path == kStdioPath;
  if (to_stdout && stdout_is_terminal()) {
    err << "formulon: recalc: refusing to write a binary workbook to a terminal; redirect stdout or name a file\n";
    return kExitUsage;
  }

  // `-o -` has no extension and defaults to .xlsx, as a stream does elsewhere.
  const std::optional<fm_workbook_format_t> implied = to_stdout ? std::nullopt : format_from_extension(output_path);
  if (format_option.has_value() && implied.has_value() && *format_option != *implied) {
    err << "formulon: recalc: format " << (*format_option == FM_WORKBOOK_FORMAT_XLSB ? "XLSB" : "XLSX")
        << " contradicts the extension of '" << output_path << "'\n";
    return kExitUsage;
  }
  if (!to_stdout && !format_option.has_value() && !implied.has_value()) {
    err << "formulon: recalc: cannot infer the workbook format from '" << output_path
        << "'; use a .xlsx or .xlsb extension or pass --format xlsx|xlsb\n";
    return kExitUsage;
  }
  const fm_workbook_format_t format = format_option.value_or(implied.value_or(FM_WORKBOOK_FORMAT_XLSX));

  WorkbookGuard wb;
  if (auto rc = load_workbook_file(input_path, shared, wb, err, "recalc"); rc != 0) {
    return rc;
  }

  if (iterative) {
    fm_iterative_settings_patch_t enable = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
    enable.enabled_op = FM_PATCH_OP_SET;
    enable.enabled = 1;
    if (auto rc = fm_workbook_update_iterative(wb.handle, &enable); rc != 0) {
      emit_last_error(err, "recalc");
      return rc;
    }
  }

  fm_parallel_recalc_stats_info_t parallel_stats = FM_STRUCT_INIT(fm_parallel_recalc_stats_info_t);
  if (parallel) {
    if (auto rc = fm_workbook_recalc_parallel(wb.handle, thread_count, &parallel_stats); rc != 0) {
      emit_last_error(err, "recalc");
      return rc;
    }
  } else {
    if (auto rc = fm_workbook_recalc(wb.handle); rc != 0) {
      emit_last_error(err, "recalc");
      return rc;
    }
  }

  SaveBuffer buf;
  fm_save_diagnostics_info_t save_diagnostics = FM_STRUCT_INIT(fm_save_diagnostics_info_t);
  if (auto rc = fm_workbook_save_with_diagnostics(wb.handle, format, &buf.data, &buf.len, &save_diagnostics); rc != 0) {
    emit_last_error(err, "recalc");
    return rc;
  }
  emit_write_diagnostics(err, "recalc", format, save_diagnostics);

  if (to_stdout) {
    out.flush();
    if (auto rc = write_stdout(buf.data, buf.len); rc != 0) {
      err << "formulon: recalc: cannot write to stdout: " << std::strerror(errno) << '\n';
      return rc;
    }
  } else if (auto rc = write_file_atomically(output_path, buf.data, buf.len); rc != 0) {
    err << "formulon: recalc: cannot write '" << output_path << "': " << std::strerror(errno) << '\n';
    return rc;
  }

  if (!quiet) {
    if (parallel) {
      err << "formulon: recalc: ok, threads=" << thread_count << " cells_evaluated=" << parallel_stats.cells_evaluated
          << " sccs_processed=" << parallel_stats.sccs_processed << " parallel_steps=" << parallel_stats.parallel_steps
          << " serial_fallback_steps=" << parallel_stats.serial_fallback_steps
          << " cycle_recoveries=" << parallel_stats.cycle_recoveries
          << " worker_threads_started=" << parallel_stats.worker_threads_started
          << " worker_threads_used=" << parallel_stats.worker_threads_used << ", wrote " << buf.len << " bytes to '"
          << output_path << "'\n";
    } else {
      err << "formulon: recalc: ok, wrote " << buf.len << " bytes to '" << output_path << "'\n";
    }
  }
  return 0;
}

}  // namespace cli
}  // namespace formulon
