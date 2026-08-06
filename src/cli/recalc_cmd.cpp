//
// `formulon recalc <in.xlsx> -o <out.xlsx>` — load, recalc, save.
//
// File I/O happens in the CLI layer (not the C ABI) because the engine
// stays bytes-in / bytes-out for portability with WASM and language
// bindings. The CLI handler:
//
//   1. Reads `in` into memory (binary mode).
//   2. Calls `fm_workbook_load`, then `fm_workbook_recalc`, then
//      `fm_workbook_save_ex` with a format derived from `-o`'s
//      extension (`.xlsb` -> MS-XLSB, anything else -> `.xlsx`).
//   3. Writes the saved buffer back out to `out`.
//
// `--iterative` enables iterative-calc while preserving any iteration cap
// and convergence threshold loaded from the workbook. `--quiet` suppresses
// the `Recalculated ...` status line on stderr; otherwise we print one per
// successful run.

#include <cctype>
#include <cerrno>
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
  out << "Usage: formulon recalc [--iterative] [--quiet] <in.xlsx> -o <out.xlsx>\n"
      << "\n"
      << "Load <in.xlsx>, drive a full recalc, and write the result to <out.xlsx>.\n"
      << "The output format is chosen from -o's extension: '.xlsb' writes MS-XLSB,\n"
      << "any other extension (or none) writes OOXML .xlsx.\n"
      << "Status: prints \"formulon: recalc: ok, wrote M bytes to 'OUT'\" on stderr unless\n"
      << "--quiet is supplied.\n";
}

void emit_last_error(std::ostream& err, const char* subcommand) {
  err << "formulon: " << subcommand << ": " << fm_last_error_message();
  const char* ctx = fm_last_error_context();
  if (ctx != nullptr && ctx[0] != '\0') {
    err << " (" << ctx << ')';
  }
  err << '\n';
}

// Derives the `fm_workbook_save_ex` container format from `path`'s
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

}  // namespace

int run_recalc(const ArgList& args, std::ostream& out, std::ostream& err) {
  std::string input_path;
  std::string output_path;
  bool iterative = false;
  bool quiet = false;
  bool input_seen = false;

  for (std::size_t i = 0; i < args.size(); ++i) {
    const std::string_view a = args[i];
    if (a == "-h" || a == "--help") {
      print_recalc_usage(out);
      return 0;
    }
    if (a == "--iterative") {
      iterative = true;
      continue;
    }
    if (a == "--quiet") {
      quiet = true;
      continue;
    }
    if (a == "-o" || a == "--output") {
      if (i + 1 >= args.size()) {
        err << "formulon: recalc: " << a << " requires a path argument\n";
        return kExitUsage;
      }
      output_path.assign(args[i + 1]);
      ++i;
      continue;
    }
    if (!a.empty() && a[0] == '-') {
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

  if (iterative) {
    if (auto rc = fm_workbook_set_iterative_enabled(wb.handle, 1); rc != 0) {
      emit_last_error(err, "recalc");
      return rc;
    }
  }

  if (auto rc = fm_workbook_recalc(wb.handle); rc != 0) {
    emit_last_error(err, "recalc");
    return rc;
  }

  SaveBuffer buf;
  const fm_workbook_format_t format = format_from_extension(output_path);
  if (auto rc = fm_workbook_save_ex(wb.handle, format, &buf.data, &buf.len); rc != 0) {
    emit_last_error(err, "recalc");
    return rc;
  }

  if (auto rc = write_file_atomically(output_path, buf.data, buf.len); rc != 0) {
    err << "formulon: recalc: cannot write '" << output_path << "': " << std::strerror(errno) << '\n';
    return rc;
  }

  if (!quiet) {
    // We do not have a per-pass cell count from the C ABI yet, so the
    // status line reports the saved-byte budget alongside the input
    // workbook's sheet count to give an order-of-magnitude signal.
    err << "formulon: recalc: ok, wrote " << buf.len << " bytes to '" << output_path << "'\n";
  }
  // Suppress unused-stream warnings: the success path of `recalc` is
  // intentionally silent on stdout. The block below is a no-op in
  // release; in tests it lets the harness assert "no stdout chatter".
  (void)out;
  return 0;
}

}  // namespace cli
}  // namespace formulon
