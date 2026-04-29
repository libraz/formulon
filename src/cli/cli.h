// Copyright 2026 libraz. Licensed under the MIT License.
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
//   * Anything outside the [0, 127] range is clamped to `1` by `main`
//     so it survives shell encoding without misinterpretation.

#ifndef FORMULON_CLI_CLI_H_
#define FORMULON_CLI_CLI_H_

#include <iosfwd>
#include <string_view>
#include <vector>

namespace formulon {
namespace cli {

/// Argument list as packaged by `main`. The vector lifetime is bounded
/// by the call frame in `main`; handlers must not retain the views
/// across function boundaries.
using ArgList = std::vector<std::string_view>;

/// Generic usage exit code. Mirrors sysexits(3) `EX_USAGE = 64`.
inline constexpr int kExitUsage = 64;

/// `eval` handler: evaluate a single formula on a fresh empty workbook.
///
/// `args` carries the post-`eval` arguments. The first non-flag argument
/// is the formula text (with or without a leading `=`).
///
/// Supported flags: `--json` (object output), `--repeat N` (re-evaluate
/// `N` times and report timing on stderr), `-h | --help`.
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

/// Prints the top-level usage banner to `out` and returns `0`.
int print_usage(std::ostream& out);

}  // namespace cli
}  // namespace formulon

#endif  // FORMULON_CLI_CLI_H_
