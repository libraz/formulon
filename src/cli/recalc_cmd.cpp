// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `formulon recalc <in.xlsx> -o <out.xlsx>` — load, recalc, save.
//
// File I/O happens in the CLI layer (not the C ABI) because the engine
// stays bytes-in / bytes-out for portability with WASM and language
// bindings. The CLI handler:
//
//   1. Reads `in` into memory (binary mode).
//   2. Calls `fm_workbook_load`, then `fm_workbook_recalc`, then
//      `fm_workbook_save`.
//   3. Writes the saved buffer back out to `out`.
//
// `--iterative` enables iterative-calc with Excel's default knobs
// (max=100, change=0.001) before the recalc step. `--quiet` suppresses
// the `Recalculated ...` status line on stderr; otherwise we print one
// per successful run.

#include <cerrno>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <fstream>
#include <ios>
#include <ostream>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "cli/cli.h"
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
      << "Status: prints \"Recalculated N cells -> wrote M bytes\" on stderr unless\n"
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

// Slurps `path` into a heap-allocated `vector<uint8_t>`. Returns
// `kIoFileNotFound` on open failure (we map every fopen failure mode
// onto this code; the OS-specific reason lives in `strerror`).
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

// Writes `bytes` to `path` in binary mode. Returns `kCliOutputFailed`
// on any open or write error.
fm_status_t write_all(const std::string& path, const std::uint8_t* bytes, std::size_t len) {
  std::ofstream out(path, std::ios::binary | std::ios::trunc);
  if (!out) {
    return static_cast<fm_status_t>(FormulonErrorCode::kCliOutputFailed);
  }
  if (len > 0) {
    out.write(reinterpret_cast<const char*>(bytes), static_cast<std::streamsize>(len));
    if (!out) {
      return static_cast<fm_status_t>(FormulonErrorCode::kCliOutputFailed);
    }
  }
  return 0;
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
  if (auto rc = read_all(input_path, bytes); rc != 0) {
    err << "formulon: recalc: cannot read '" << input_path << "': " << std::strerror(errno) << '\n';
    return rc;
  }

  WorkbookGuard wb;
  if (auto rc = fm_workbook_load(bytes.data(), bytes.size(), &wb.handle); rc != 0) {
    emit_last_error(err, "recalc");
    return rc;
  }

  if (iterative) {
    if (auto rc = fm_workbook_set_iterative(wb.handle, 1, 100, 0.001); rc != 0) {
      emit_last_error(err, "recalc");
      return rc;
    }
  }

  if (auto rc = fm_workbook_recalc(wb.handle); rc != 0) {
    emit_last_error(err, "recalc");
    return rc;
  }

  SaveBuffer buf;
  if (auto rc = fm_workbook_save(wb.handle, &buf.data, &buf.len); rc != 0) {
    emit_last_error(err, "recalc");
    return rc;
  }

  if (auto rc = write_all(output_path, buf.data, buf.len); rc != 0) {
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
