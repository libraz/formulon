//
// `formulon_cli` shared command surface.
//
// The CLI is a single-binary tool that drives the engine through the
// stable C ABI declared in `c_api/formulon_c.h`. Each subcommand lives
// in its own translation unit (`eval_cmd.cpp`, `recalc_cmd.cpp`,
// `dump_cmd.cpp`); `main.cpp` parses `argv[1]`, dispatches to one of
// the handlers below, and translates the returned `fm_status_t` into a
// process exit code.
//
// All handlers accept the post-subcommand argument list (i.e.
// `argv[2..argc]` packaged as `string_view`s) and a pair of output
// streams so tests can inject in-memory streams. Stdout carries the
// command's primary output; stderr carries diagnostics and progress
// chatter.
//
// Error contract:
//   * `0` on success.
//   * Any value mirrored from `fm_status_t`: bound by `set_last_error`
//     in the C API, so the CLI surfaces the diagnostic by reading
//     `fm_last_error_message()` and prefixing with the subcommand name.
//   * `64` on a usage error (mirrors sysexits(3) `EX_USAGE`).
//   * Any other engine / I/O / binding failure maps to the stable
//     generic-failure exit code `1`. The detailed `fm_status_t` remains
//     available in the diagnostic text; it is never encoded in the
//     process status, where it could collide after POSIX truncation.

#ifndef FORMULON_CLI_CLI_H_
#define FORMULON_CLI_CLI_H_

#include <ostream>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "utils/error.h"

namespace formulon {
namespace cli {

/// Argument list as packaged by `main`. The vector lifetime is bounded
/// by the call frame in `main`; handlers must not retain the views
/// across function boundaries.
using ArgList = std::vector<std::string_view>;

/// Generic usage exit code. Mirrors sysexits(3) `EX_USAGE = 64`.
inline constexpr int kExitUsage = 64;

/// Maps an internal command status to the CLI's small, stable exit-code
/// vocabulary. `fm_status_t` values are intentionally not exposed as
/// process statuses: POSIX shells truncate them to eight bits, making
/// unrelated failures indistinguishable (for example 8000 and EX_USAGE).
constexpr int exit_code_for_status(int status) {
  if (status == 0) {
    return 0;
  }
  return status == kExitUsage ? kExitUsage : 1;
}

/// Flushes a command's primary output and turns a stream failure into the
/// stable CLI output error. Commands must call this after emitting their
/// main result so exit status 0 means the complete result reached `out`.
inline int flush_output(std::ostream& out, std::ostream& err, std::string_view subcommand) {
  out.flush();
  if (out) {
    return 0;
  }
  err << "formulon: " << subcommand << ": failed to write output\n";
  return static_cast<int>(FormulonErrorCode::kCliOutputFailed);
}

/// Prints the engine version line to `out` and returns `0`. The usage
/// banner lists `--version` alongside `-h | --help` as a common option, so
/// every handler short-circuits on it through this helper and the flag
/// means the same thing wherever it appears.
inline int print_version(std::ostream& out) {
  const char* version = fm_version_string();
  out << (version != nullptr ? version : "") << '\n';
  return 0;
}

/// `eval` handler: evaluate a single formula on a fresh empty workbook.
///
/// `args` carries the post-`eval` arguments. The first non-flag argument
/// is the formula text (with or without a leading `=`).
///
/// Supported flags: `--json`, `--repeat N` (re-evaluate `N` times and
/// report timing on stderr), `-h | --help`, and `--` to end option
/// parsing before a formula beginning with `-`.
///
/// `--json` has two shapes, selected by the result's dimensions: a single
/// `{"kind": ..., "value": ...}` object when the result is exactly one row
/// by one column, and otherwise a row-major array of arrays of those
/// objects (one inner array per row). A formula that spills therefore
/// produces the nested shape and a scalar one does not.
int run_eval(const ArgList& args, std::ostream& out, std::ostream& err);

/// `recalc` handler: load `.xlsx`, recalc, save to `--output`.
///
/// `args` carries the post-`recalc` arguments. The first non-flag
/// argument is the input path.
///
/// Supported flags: `-o | --output PATH` (required), `--iterative`
/// (enable iterative calc), `--quiet`, `-h | --help`.
int run_recalc(const ArgList& args, std::ostream& out, std::ostream& err);

/// `dump` handler: print workbook contents in a diff-friendly form.
///
/// `args` carries the post-`dump` arguments. The first non-flag argument
/// is the input path.
///
/// Supported flags (mutually exclusive): `--formulas` (default),
/// `--values`, `--sheets`, `--metadata`, `-h | --help`.
int run_dump(const ArgList& args, std::ostream& out, std::ostream& err);

/// `paginate` handler: resolve the print geometry of one worksheet.
///
/// Supported flags: `--sheet INDEX` (0-based, default 0), `-h | --help`.
int run_paginate(const ArgList& args, std::ostream& out, std::ostream& err);

/// Prints the top-level usage banner to `out` and returns `0`.
int print_usage(std::ostream& out);

}  // namespace cli
}  // namespace formulon

#endif  // FORMULON_CLI_CLI_H_
