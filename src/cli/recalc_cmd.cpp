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
//      `fm_workbook_save_ex_with_diagnostics` with a format derived from
//      `-o`'s extension (`.xlsb` -> MS-XLSB, anything else -> `.xlsx`).
//   3. Writes the saved buffer back out to `out`.
//
// `--iterative` enables iterative-calc while preserving any iteration cap
// and convergence threshold loaded from the workbook. `--threads N` opts
// into the parallel scheduler; omitting it deliberately preserves the
// serial `recalc` contract. `--quiet` suppresses the `Recalculated ...`
// status line on stderr; otherwise we print one per successful run.

#include <cctype>
#include <cerrno>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <ostream>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "cli/cli.h"
#include "cli/file_io.h"
#include "utils/error.h"

namespace formulon {
namespace cli {

namespace {

// Same RAII pattern as `eval_cmd.cpp`. Duplicated rather than shared
// because the CLI handlers are deliberately small and a shared helper
// header would obscure their independence.
struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
};

// Heap buffer returned by `fm_workbook_save` paired with its dedicated
// `fm_buffer_free` deallocator.
struct SaveBuffer {
  std::uint8_t* data = nullptr;
  std::size_t len = 0;
  SaveBuffer() = default;
  SaveBuffer(const SaveBuffer&) = delete;
  SaveBuffer& operator=(const SaveBuffer&) = delete;
  ~SaveBuffer() { fm_buffer_free(data); }
};

void print_recalc_usage(std::ostream& out) {
  out << "Usage: formulon recalc [--iterative] [--threads N] [--quiet] <in.xlsx-or-xlsb> -o <out.xlsx-or-xlsb>\n"
      << "       formulon recalc [--iterative] [--threads N] [--quiet] -o <out.xlsx-or-xlsb> -- <in.xlsx-or-xlsb>\n"
      << "\n"
      << "Load <in.xlsx-or-xlsb>, drive a full recalc, and write the result to <out.xlsx-or-xlsb>.\n"
      << "The output format is chosen from -o's extension: '.xlsb' writes MS-XLSB,\n"
      << "any other extension (or none) writes OOXML .xlsx.\n"
      << "--threads N opts into parallel recalc: 0 selects automatic detection (capped at 8),\n"
      << "1 stays on the caller thread, and 2..8 select the worker cap. Without --threads,\n"
      << "recalc remains serial.\n"
      << "Options must precede --; after --, exactly one input path is accepted.\n"
      << "Status: prints \"formulon: recalc: ok, wrote M bytes to 'OUT'\" on stderr unless\n"
      << "--quiet is supplied; --quiet does not suppress XLSB loss warnings.\n";
}

bool parse_thread_count(std::string_view text, std::uint32_t* out) {
  if (text.empty()) {
    return false;
  }
  std::uint32_t parsed = 0U;
  for (const char c : text) {
    if (c < '0' || c > '9') {
      return false;
    }
    const std::uint32_t digit = static_cast<std::uint32_t>(c - '0');
    if (parsed > (std::numeric_limits<std::uint32_t>::max() - digit) / 10U) {
      return false;
    }
    parsed = parsed * 10U + digit;
  }
  if (parsed > 8U) {
    return false;
  }
  *out = parsed;
  return true;
}

void emit_last_error(std::ostream& err, const char* subcommand) {
  err << "formulon: " << subcommand << ": " << fm_last_error_message();
  const char* ctx = fm_last_error_context();
  if (ctx != nullptr && ctx[0] != '\0') {
    err << " (" << ctx << ')';
  }
  err << '\n';
}

// Derives the `fm_workbook_save_ex_with_diagnostics` container format from `path`'s
// extension: `.xlsb` (case-insensitive) selects MS-XLSB; every other
// extension (including none) selects `.xlsx` so existing callers that
// never named `.xlsb` keep writing OOXML, matching `fm_workbook_save`'s
// prior behaviour.
fm_workbook_format_t format_from_extension(const std::string& path) {
  const std::size_t dot = path.find_last_of('.');
  if (dot == std::string::npos) {
    return FM_WORKBOOK_FORMAT_XLSX;
  }
  std::string ext = path.substr(dot + 1);
  for (char& c : ext) {
    c = static_cast<char>(std::tolower(static_cast<unsigned char>(c)));
  }
  return ext == "xlsb" ? FM_WORKBOOK_FORMAT_XLSB : FM_WORKBOOK_FORMAT_XLSX;
}

void emit_read_diagnostics(std::ostream& err, std::size_t undecoded_formula_count,
                           std::size_t undecoded_defined_name_count, std::size_t dropped_part_count) {
  if (undecoded_formula_count == 0 && undecoded_defined_name_count == 0 && dropped_part_count == 0) {
    return;
  }
  err << "formulon: recalc: warning: XLSB read diagnostics";
  if (undecoded_formula_count != 0) {
    err << "; undecoded_formula_count=" << undecoded_formula_count;
  }
  if (undecoded_defined_name_count != 0) {
    err << "; undecoded_defined_name_count=" << undecoded_defined_name_count;
  }
  if (dropped_part_count != 0) {
    err << "; dropped_part_count=" << dropped_part_count;
  }
  err << '\n';
}

void emit_write_diagnostics(std::ostream& err, std::size_t downgraded_formula_count,
                            std::size_t deferred_feature_count) {
  if (downgraded_formula_count == 0 && deferred_feature_count == 0) {
    return;
  }
  err << "formulon: recalc: warning: XLSB write diagnostics";
  if (downgraded_formula_count != 0) {
    err << "; downgraded_formula_count=" << downgraded_formula_count;
  }
  if (deferred_feature_count != 0) {
    err << "; deferred_feature_count=" << deferred_feature_count;
  }
  err << '\n';
}

}  // namespace

int run_recalc(const ArgList& args, std::ostream& out, std::ostream& err) {
  std::string input_path;
  std::string output_path;
  bool iterative = false;
  bool quiet = false;
  bool parallel = false;
  std::uint32_t thread_count = 0U;
  bool input_seen = false;
  bool options_ended = false;

  for (std::size_t i = 0; i < args.size(); ++i) {
    const std::string_view a = args[i];
    if (!options_ended && a == "--") {
      options_ended = true;
      continue;
    }
    if (!options_ended && (a == "-h" || a == "--help")) {
      print_recalc_usage(out);
      return 0;
    }
    if (!options_ended && a == "--version") {
      return print_version(out);
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
        err << "formulon: recalc: --threads requires an integer from 0 to 8\n";
        return kExitUsage;
      }
      if (!parse_thread_count(args[i + 1], &thread_count)) {
        err << "formulon: recalc: --threads must be an integer from 0 to 8\n";
        return kExitUsage;
      }
      parallel = true;
      ++i;
      continue;
    }
    if (!options_ended && (a == "-o" || a == "--output")) {
      if (i + 1 >= args.size()) {
        err << "formulon: recalc: " << a << " requires a path argument\n";
        return kExitUsage;
      }
      output_path.assign(args[i + 1]);
      ++i;
      continue;
    }
    if (!options_ended && !a.empty() && a[0] == '-') {
      err << "formulon: recalc: unknown flag '" << a << "'\n";
      return kExitUsage;
    }
    if (input_seen) {
      err << "formulon: recalc: only one input path may be supplied\n";
      return kExitUsage;
    }
    input_path.assign(a);
    input_seen = true;
  }

  if (!input_seen) {
    err << "formulon: recalc: missing input path\n";
    print_recalc_usage(err);
    return kExitUsage;
  }
  if (output_path.empty()) {
    err << "formulon: recalc: missing -o/--output PATH\n";
    return kExitUsage;
  }

  std::vector<std::uint8_t> bytes;
  if (auto rc = read_file(input_path, bytes); rc != 0) {
    err << "formulon: recalc: cannot read '" << input_path << "': " << std::strerror(errno) << '\n';
    return rc;
  }

  WorkbookGuard wb;
  if (auto rc = fm_workbook_load(bytes.data(), bytes.size(), &wb.handle); rc != 0) {
    emit_last_error(err, "recalc");
    return rc;
  }

  std::size_t undecoded_formula_count = 0;
  std::size_t undecoded_defined_name_count = 0;
  std::size_t dropped_part_count = 0;
  if (auto rc = fm_workbook_xlsb_read_diagnostics_ex(wb.handle, &undecoded_formula_count, &undecoded_defined_name_count,
                                                     &dropped_part_count);
      rc != 0) {
    emit_last_error(err, "recalc");
    return rc;
  }
  emit_read_diagnostics(err, undecoded_formula_count, undecoded_defined_name_count, dropped_part_count);

  if (iterative) {
    if (auto rc = fm_workbook_set_iterative_enabled(wb.handle, 1); rc != 0) {
      emit_last_error(err, "recalc");
      return rc;
    }
  }

  fm_parallel_recalc_stats parallel_stats{};
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
  const fm_workbook_format_t format = format_from_extension(output_path);
  std::size_t downgraded_formula_count = 0;
  std::size_t deferred_feature_count = 0;
  if (auto rc = fm_workbook_save_ex_with_diagnostics(wb.handle, format, &buf.data, &buf.len, &downgraded_formula_count,
                                                     &deferred_feature_count);
      rc != 0) {
    emit_last_error(err, "recalc");
    return rc;
  }
  emit_write_diagnostics(err, downgraded_formula_count, deferred_feature_count);

  if (auto rc = write_file_atomically(output_path, buf.data, buf.len); rc != 0) {
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
  // Suppress unused-stream warnings: the success path of `recalc` is
  // intentionally silent on stdout. The block below is a no-op in
  // release; in tests it lets the harness assert "no stdout chatter".
  (void)out;
  return 0;
}

}  // namespace cli
}  // namespace formulon
