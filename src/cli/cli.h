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

#include <cstddef>
#include <cstdint>
#include <optional>
#include <ostream>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "excel_profile.h"
#include "utils/date_time.h"
#include "utils/error.h"

namespace formulon {
namespace cli {

/// Argument list as packaged by `main`. The vector lifetime is bounded
/// by the call frame in `main`; handlers must not retain the views
/// across function boundaries.
using ArgList = std::vector<std::string_view>;

/// Writes the supported profile ids, comma-separated, from the engine's table.
inline void print_profile_ids(std::ostream& out) {
  const char* separator = "";
  for (const detail::ExcelProfileIdEntry& entry : detail::kExcelProfileIds) {
    out << separator << entry.id;
    separator = ", ";
  }
}

/// Writes the help lines of the value-taking shared options.
inline void print_shared_options(std::ostream& out) {
  out << "  --profile <id>        Select the Excel formula profile (default: "
      << excel_profile_id(default_excel_profile()) << ").\n"
      << "                        Supported ids: ";
  print_profile_ids(out);
  out << ".\n"
      << "  --now <date[Thh:mm:ss]>\n"
      << "                        Pin the clock NOW, TODAY and relative date filters\n"
      << "                        read, as local wall time (e.g. 2026-10-09T09:30:00;\n"
      << "                        a bare date is midnight). Default: the host clock.\n";
}

/// The value-taking options every subcommand shares.
struct SharedOptions {
  std::optional<std::string_view> profile_id;
  std::optional<fm_civil_time_t> now;
};

/// Parses `YYYY-MM-DD` or `YYYY-MM-DDThh:mm:ss` into a wall-clock reading in
/// the range `fm_workbook_set_pinned_now` accepts.
inline bool parse_pinned_now(std::string_view text, fm_civil_time_t* out) {
  const auto field = [&](std::size_t pos, std::size_t len, std::int32_t* value) {
    if (pos + len > text.size()) {
      return false;
    }
    std::int32_t v = 0;
    for (std::size_t i = pos; i < pos + len; ++i) {
      if (text[i] < '0' || text[i] > '9') {
        return false;
      }
      v = v * 10 + (text[i] - '0');
    }
    *value = v;
    return true;
  };
  fm_civil_time_t t{};
  if (text.size() != 10 && text.size() != 19) {
    return false;
  }
  if (!field(0, 4, &t.year) || text[4] != '-' || !field(5, 2, &t.month) || text[7] != '-' || !field(8, 2, &t.day)) {
    return false;
  }
  if (text.size() == 19 && (text[10] != 'T' || !field(11, 2, &t.hour) || text[13] != ':' || !field(14, 2, &t.minute) ||
                            text[16] != ':' || !field(17, 2, &t.second))) {
    return false;
  }
  if (t.year < 1900 || t.month < 1 || t.month > 12 || t.day < 1 ||
      t.day > static_cast<std::int32_t>(date_time::days_in_month(t.year, static_cast<unsigned>(t.month))) ||
      t.hour > 23 || t.minute > 59 || t.second > 59) {
    return false;
  }
  *out = t;
  return true;
}

/// Applies the shared options to a freshly created or loaded workbook.
inline fm_status_t apply_shared_options(fm_workbook_t* workbook, const SharedOptions& options) {
  if (options.profile_id.has_value()) {
    const std::string id(*options.profile_id);
    if (const fm_status_t rc = fm_workbook_set_excel_profile_id(workbook, id.c_str()); rc != 0) {
      return rc;
    }
  }
  if (options.now.has_value()) {
    return fm_workbook_set_pinned_now(workbook, &*options.now);
  }
  return 0;
}

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

/// Prints the engine version line to `out`. The usage banner lists
/// `--version` alongside `-h | --help` as a common option, so every handler
/// short-circuits on it through this helper and the flag means the same
/// thing wherever it appears.
///
/// The version line is this invocation's primary output, so it goes through
/// `flush_output` like any other: a script reading `$(formulon --version)`
/// must not receive an empty string together with a success status.
inline int print_version(std::ostream& out, std::ostream& err) {
  const char* version = fm_version_string();
  out << (version != nullptr ? version : "") << '\n';
  return flush_output(out, err, "version");
}

/// Outcome of `handle_common_option`.
enum class CommonOption {
  kNotCommon,  ///< The argument is not a shared option; the handler keeps parsing it.
  kConsumed,   ///< The argument was a shared option (`--`, or `--profile` / `--now` and its value).
  kExit,       ///< The invocation is finished; return `exit_code`.
};

/// Handles the options every subcommand shares: `--` (ends option parsing,
/// recorded in `options_ended`), `-h | --help` (prints `print_usage_fn` to
/// `out`), `--version`, `--profile <id>` and `--now <date[Thh:mm:ss]>`. A
/// value-taking option is validated, stored in `shared`, and advances `index`
/// over its value; repeating it replaces the previous value. Once
/// `options_ended` is set nothing is shared.
inline CommonOption handle_common_option(const ArgList& args, std::size_t& index, bool& options_ended,
                                         SharedOptions& shared, void (*print_usage_fn)(std::ostream&),
                                         std::string_view subcommand, std::ostream& out, std::ostream& err,
                                         int& exit_code) {
  const std::string_view arg = args[index];
  if (options_ended) {
    return CommonOption::kNotCommon;
  }
  if (arg == "--") {
    options_ended = true;
    return CommonOption::kConsumed;
  }
  if (arg == "-h" || arg == "--help") {
    print_usage_fn(out);
    exit_code = flush_output(out, err, subcommand);
    return CommonOption::kExit;
  }
  if (arg == "--version") {
    exit_code = print_version(out, err);
    return CommonOption::kExit;
  }
  if (arg == "--profile") {
    if (index + 1 >= args.size()) {
      err << "formulon: " << subcommand << ": --profile requires a value\n";
      exit_code = kExitUsage;
      return CommonOption::kExit;
    }
    const std::string_view value = args[index + 1];
    ExcelProfile parsed{};
    if (!parse_excel_profile_id(value, &parsed)) {
      err << "formulon: " << subcommand << ": invalid Excel profile id '" << value << "'\n"
          << "formulon: " << subcommand << ": supported ids: ";
      print_profile_ids(err);
      err << "\n";
      exit_code = kExitUsage;
      return CommonOption::kExit;
    }
    shared.profile_id = value;
    ++index;
    return CommonOption::kConsumed;
  }
  if (arg == "--now") {
    fm_civil_time_t now{};
    if (index + 1 >= args.size() || !parse_pinned_now(args[index + 1], &now)) {
      err << "formulon: " << subcommand << ": --now requires a date YYYY-MM-DD or YYYY-MM-DDThh:mm:ss"
          << " from 1900-01-01 to 9999-12-31\n";
      exit_code = kExitUsage;
      return CommonOption::kExit;
    }
    shared.now = now;
    ++index;
    return CommonOption::kConsumed;
  }
  return CommonOption::kNotCommon;
}

/// `eval` handler: evaluate a single formula on a fresh empty workbook.
///
/// `args` carries the post-`eval` arguments. The first non-flag argument
/// is the formula text (with or without a leading `=`).
///
/// Supported flags: the common options above, `--json`, `--repeat N` (re-evaluate `N` times and
/// report timing on stderr for every `N` the flag accepts, including 1),
/// `-h | --help`, and `--` to end option parsing before a formula
/// beginning with `-`.
///
/// Malformed syntax is not a handler failure: the evaluator turns parser
/// recovery placeholders into a cell-level `#NAME?`, and `eval` prints
/// that value and exits `0` like every other binding over the same C ABI.
///
/// Plain output is a TAB-separated grid of exactly `rows` lines with
/// `cols` fields each, so cell payloads are escaped the way `dump`
/// escapes them and an embedded newline or TAB cannot forge a record.
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
/// Supported flags: the common options above, `-o | --output PATH` (required), `--iterative`
/// (enable iterative calc), `--threads N` (opt in to parallel recalc with
/// an inclusive 0..8 worker setting), `--quiet`, `-h | --help`, and `--` to
/// end option parsing before the input path. All options, including `-o`,
/// must precede `--`; omitting `--threads` preserves the serial recalc
/// contract.
int run_recalc(const ArgList& args, std::ostream& out, std::ostream& err);

/// `dump` handler: print workbook contents in a diff-friendly form.
///
/// `args` carries the post-`dump` arguments. The first non-flag argument
/// is the input path.
///
/// Supported flags: the common options above and (mutually exclusive): `--formulas` (default),
/// `--values`, `--sheets`, `--metadata`, `-h | --help`, and `--` to end
/// option parsing before an input path.
int run_dump(const ArgList& args, std::ostream& out, std::ostream& err);

/// `paginate` handler: resolve the print geometry of one worksheet.
///
/// Supported flags: the common options above, `--sheet INDEX` (0-based,
/// default 0), and `--` to end option parsing before an input path.
int run_paginate(const ArgList& args, std::ostream& out, std::ostream& err);

/// Prints the top-level usage banner to `out`. Requested help is a
/// primary output too, so the banner is flushed through `flush_output`
/// and a closed stdout fails the invocation instead of exiting `0`.
int print_usage(std::ostream& out, std::ostream& err);

}  // namespace cli
}  // namespace formulon

#endif  // FORMULON_CLI_CLI_H_
