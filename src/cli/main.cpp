//
// `formulon_cli` entry point: tiny argv dispatcher.
//
// Subcommands live in their own translation units (`eval_cmd.cpp`,
// `recalc_cmd.cpp`, `dump_cmd.cpp`); this file just routes `argv[1]`
// to the right handler and translates the returned status into a process
// exit code, shared by every subcommand: 0 success, 1 engine failure,
// 64 usage, 66 input missing or unreadable, 73 output cannot be created,
// 74 I/O error mid-stream. `--opt=value` is split into `--opt value` here,
// once, before dispatch. SIGPIPE is ignored so a closed stdout surfaces as
// exit 74 instead of killing the process.

#include <csignal>
#include <cstdint>
#include <cstdlib>
#include <iostream>
#include <string>
#include <string_view>
#include <vector>

#include "cli/cli.h"
#include "cli/render.h"

namespace {

void print_top_usage(std::ostream& out) {
  out << "Usage: formulon <command> [options]\n"
      << "\n"
      << "Commands:\n"
      << "  eval <formula>          Evaluate a single formula on a fresh empty workbook.\n"
      << "  recalc [--iterative] [--threads N] [--quiet] <in> -o <out>\n"
      << "                          Load, recalc, and write a workbook ('-o -' writes to stdout).\n"
      << "  dump [--formulas|--values|--sheets|--metadata] <in>\n"
      << "                          Print a diff-friendly snapshot of a workbook.\n"
      << "                          --values recalculates first; the others do not.\n"
      << "  paginate [--sheet N] <in>\n"
      << "                          Resolve print area (A1 ranges), page breaks, and page count.\n"
      << "  profiles [--json]       List the supported Excel profile ids, one per line.\n"
      << "\n"
      << "<in> and <out> are xlsx or xlsb files. eval accepts --json and --repeat N;\n"
      << "dump, paginate and profiles accept --json (output starts with {\"format\":1,...});\n"
      << "recalc accepts -o/--output, --iterative, --threads N and --quiet; every\n"
      << "subcommand accepts -- to end options. Options take their value as\n"
      << "'--opt value' or '--opt=value'.\n"
      << "Plain output escapes only \\n, \\r, \\t and \\\\.\n"
      << "\n"
      << "Standard streams:\n"
      << "  An input of '-' reads the whole workbook from stdin (refused on a terminal).\n"
      << "  recalc -o - writes the .xlsx workbook to stdout (refused on a terminal).\n"
      << "  Use './-' for a file named '-'.\n"
      << "\n"
      << "Exit codes:\n"
      << "  0   success (an Excel error value such as #DIV/0! is still a result)\n"
      << "  1   engine failure, including an empty or corrupt input workbook\n"
      << "  64  usage error\n"
      << "  66  input file missing, unreadable or a directory\n"
      << "  73  output file cannot be created\n"
      << "  74  I/O error while reading or writing, including a closed stdout\n"
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

namespace {

void print_profiles_usage(std::ostream& out) {
  out << "Usage: formulon profiles [--json] [--]\n"
      << "\n"
      << "Print every supported Excel profile id, one per line.\n"
      << "  --json       print {\"format\":1,\"profiles\":[...]} instead\n";
}

}  // namespace

int run_profiles(const ArgList& args, std::ostream& out, std::ostream& err) {
  bool want_json = false;
  bool options_ended = false;
  for (const std::string_view arg : args) {
    if (!options_ended && arg == "--") {
      options_ended = true;
    } else if (!options_ended && (arg == "-h" || arg == "--help")) {
      print_profiles_usage(out);
      return flush_output(out, err, "profiles");
    } else if (!options_ended && arg == "--version") {
      return print_version(out, err);
    } else if (!options_ended && arg == "--json") {
      want_json = true;
    } else {
      err << "formulon: profiles: unexpected argument '" << arg << "'\n";
      return kExitUsage;
    }
  }
  if (want_json) {
    out << "{\"format\":1,\"profiles\":[";
    const char* separator = "";
    for (const detail::ExcelProfileIdEntry& entry : detail::kExcelProfileIds) {
      out << separator << render_json_string(entry.id);
      separator = ",";
    }
    out << "]}\n";
  } else {
    for (const detail::ExcelProfileIdEntry& entry : detail::kExcelProfileIds) {
      out << entry.id << '\n';
    }
  }
  return flush_output(out, err, "profiles");
}

}  // namespace cli
}  // namespace formulon

int main(int argc, char** argv) {
  std::signal(SIGPIPE, SIG_IGN);
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
  formulon::cli::ArgList raw_args;
  raw_args.reserve(static_cast<std::size_t>(argc));
  for (int i = 2; i < argc; ++i) {
    raw_args.emplace_back(argv[i]);
  }
  std::vector<std::string> expanded_storage;
  formulon::cli::ArgList args;
  if (!formulon::cli::split_option_values(raw_args, cmd, expanded_storage, args, std::cerr)) {
    return formulon::cli::kExitUsage;
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
  } else if (cmd == "profiles") {
    rc = formulon::cli::run_profiles(args, std::cout, std::cerr);
  } else {
    std::cerr << "formulon: unknown command '" << cmd << "'\n";
    print_top_usage(std::cerr);
    return formulon::cli::kExitUsage;
  }

  return formulon::cli::exit_code_for_status(rc);
}
