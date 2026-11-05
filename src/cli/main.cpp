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

#include "cli/cli.h"

namespace {

void print_top_usage(std::ostream& out) {
  out << "Usage: formulon <command> [options]\n"
      << "\n"
      << "Commands:\n"
      << "  eval <formula>          Evaluate a single formula on a fresh empty workbook.\n"
      << "  recalc [--iterative] [--threads N] [--quiet] <in> -o <out>\n"
      << "                          Load, recalc, and write a workbook.\n"
      << "  dump [--formulas|--values|--sheets|--metadata] <in>\n"
      << "                          Print a diff-friendly snapshot of a workbook.\n"
      << "                          --values recalculates first; the others do not.\n"
      << "  paginate [--sheet N] <in>\n"
      << "                          Resolve print area (A1 ranges), page breaks, and page count.\n"
      << "\n"
      << "<in> and <out> are xlsx or xlsb files. eval accepts --json and --repeat N;\n"
      << "recalc accepts -o/--output, --iterative, --threads N and --quiet; every\n"
      << "subcommand accepts -- to end options.\n"
      << "Plain output escapes only \\n, \\r, \\t and \\\\.\n"
      << "\n"
      << "Exit codes:\n"
      << "  0   success (an Excel error value such as #DIV/0! is still a result)\n"
      << "  1   engine, I/O or output failure\n"
      << "  64  usage error\n"
      << "\n"
      << "Common options:\n";
  formulon::cli::print_shared_options(out);
  out << "  -h, --help              Show this help (or per-subcommand help).\n"
      << "  --version               Print the engine version and exit.\n";
}

}  // namespace

namespace formulon {
namespace cli {

int print_usage(std::ostream& out, std::ostream& err) {
  print_top_usage(out);
  return flush_output(out, err, "help");
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
    return formulon::cli::exit_code_for_status(formulon::cli::print_usage(std::cout, std::cerr));
  }
  if (cmd == "--version") {
    return formulon::cli::exit_code_for_status(formulon::cli::print_version(std::cout, std::cerr));
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
  } else if (cmd == "paginate") {
    rc = formulon::cli::run_paginate(args, std::cout, std::cerr);
  } else {
    std::cerr << "formulon: unknown command '" << cmd << "'\n";
    print_top_usage(std::cerr);
    return formulon::cli::kExitUsage;
  }

  return formulon::cli::exit_code_for_status(rc);
}
