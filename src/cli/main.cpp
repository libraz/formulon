// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// `formulon_cli` entry point: tiny argv dispatcher.
//
// Subcommands live in their own translation units (`eval_cmd.cpp`,
// `recalc_cmd.cpp`, `dump_cmd.cpp`); this file just routes `argv[1]`
// to the right handler and translates the returned `fm_status_t` into
// a process exit code. The translation is "low byte clamped to
// [1, 127]", which mirrors the historical sysexits convention and
// survives shell encoding.

#include <cstdint>
#include <cstdlib>
#include <iostream>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "cli/cli.h"

namespace {

// Maps a non-zero `fm_status_t` to a POSIX-friendly exit code. We
// avoid 0 (would mask error) and clamp the high bits because shells
// historically truncate to 7-bit on macOS / Linux. Specific carve-outs
// preserve the well-known `kExitUsage = 64` value.
int status_to_exit(int rc) {
  if (rc == 0) {
    return 0;
  }
  if (rc == formulon::cli::kExitUsage) {
    return rc;
  }
  const int low = rc & 0xff;
  if (low == 0) {
    return 1;
  }
  if (low > 127) {
    return low - 128;
  }
  return low;
}

void print_top_usage(std::ostream& out) {
  out << "Usage: formulon <command> [options]\n"
      << "\n"
      << "Commands:\n"
      << "  eval <formula>          Evaluate a single formula on a fresh empty workbook.\n"
      << "  recalc <in> -o <out>    Load, recalc, and write a workbook.\n"
      << "  dump <in> [--formulas|--values|--sheets|--metadata]\n"
      << "                          Print a diff-friendly snapshot of a workbook.\n"
      << "\n"
      << "Common options:\n"
      << "  -h, --help              Show this help (or per-subcommand help).\n"
      << "  --version               Print the engine version and exit.\n";
}

}  // namespace

namespace formulon {
namespace cli {

int print_usage(std::ostream& out) {
  print_top_usage(out);
  return 0;
}

}  // namespace cli
}  // namespace formulon

int main(int argc, char** argv) {
  if (argc < 2) {
    print_top_usage(std::cerr);
    return formulon::cli::kExitUsage;
  }

  const std::string_view cmd(argv[1]);

  if (cmd == "-h" || cmd == "--help") {
    print_top_usage(std::cout);
    return 0;
  }
  if (cmd == "--version") {
    const char* v = fm_version_string();
    std::cout << (v != nullptr ? v : "") << '\n';
    return 0;
  }

  // Pack the post-subcommand args into a `string_view` vector so the
  // handlers can do their own flag parsing without re-walking argv.
  formulon::cli::ArgList args;
  args.reserve(static_cast<std::size_t>(argc));
  for (int i = 2; i < argc; ++i) {
    args.emplace_back(argv[i]);
  }

  int rc = 0;
  if (cmd == "eval") {
    rc = formulon::cli::run_eval(args, std::cout, std::cerr);
  } else if (cmd == "recalc") {
    rc = formulon::cli::run_recalc(args, std::cout, std::cerr);
  } else if (cmd == "dump") {
    rc = formulon::cli::run_dump(args, std::cout, std::cerr);
  } else {
    std::cerr << "formulon: unknown command '" << cmd << "'\n";
    print_top_usage(std::cerr);
    return formulon::cli::kExitUsage;
  }

  return status_to_exit(rc);
}
