//
// `formulon eval <formula>` — evaluate a single formula on a fresh,
// empty workbook (or at A1 of a sheet of `--in FILE`) and print the result.
//
// The handler evaluates the formula through the C ABI's side-effect-free
// array surface. Cell-level errors (`#NAME?`, `#DIV/0!`, …) are *not*
// process-level failures: the cell value just happens to be an
// `FM_VALUE_KIND_ERROR`, and we print its display name. A formula that does not
// parse is a failure of the input: exit 65 with the position on stderr and
// nothing on stdout.
//
// `--repeat N` re-evaluates the same formula `N` times through the ad-hoc
// evaluator. Timing is reported on stderr without changing the stdout payload.

#include <cerrno>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
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
#include "cli/render.h"
#include "cli/workbook_load.h"

namespace formulon {
namespace cli {

namespace {

void print_eval_usage(std::ostream& out) {
  out << "Usage: formulon eval [--json] [--local] [--repeat N] [--in <file | -> [--sheet <name | index>]] [--] "
         "<formula | ->\n"
      << "\n"
      << "Evaluate <formula> on a fresh empty workbook and print the result. A formula of '-'\n"
      << "is read from stdin (refused on a terminal).\n"
      << "Cell-level Excel errors (#NAME?, #DIV/0!, ...) print to stdout and\n"
      << "do NOT cause a non-zero exit code; only structural failures do.\n"
      << "A formula that does not parse exits 65 and prints 'formulon: eval: <message> at <line>:<column>'\n"
      << "on stderr.\n"
      << "  --json       print {\"format\":1,\"rows\":[[{\"kind\":...,\"value\":...},...],...]}\n"
      << "  --local      the formula is spelled in the --profile locale (separators, function names); otherwise en-US\n"
      << "  --in <file>  evaluate at A1 of a sheet of this xlsx or xlsb workbook ('-' reads it from stdin);\n"
      << "               the workbook is not saved\n"
      << "  --sheet <s>  with --in: the sheet name or 0-based index (default: the first sheet)\n"
      << "  --repeat N   evaluate N times and report the elapsed time on stderr\n"
      << "  --           end options; the next token is the formula\n"
      << "Plain output escapes only \\n, \\r, \\t and \\\\.\n";
  print_shared_options(out);
}

// Parses `--repeat` argument from `args[idx]`. Returns the parsed
// count, or `-1` on parse failure (which the caller turns into a usage
// error). `idx` is advanced past the value on success.
long parse_repeat(const ArgList& args, std::size_t& idx) {
  if (idx >= args.size()) {
    return -1;
  }
  const std::string_view text = args[idx];
  ++idx;
  // Zero / negative counts would skip the `recalc` entirely.
  long n = 0;
  if (!parse_unsigned<long>(text, &n) || n <= 0) {
    return -1;
  }
  return n;
}

int render_eval_result(const fm_value_array_t* result, bool want_json, std::ostream& out, std::ostream& err) {
  uint32_t rows = 0;
  uint32_t cols = 0;
  if (auto rc = fm_value_array_get_dimensions(result, &rows, &cols); rc != 0) {
    emit_last_error(err, "eval");
    return kExitFailure;
  }
  // The JSON document is assembled before anything is written so a failure
  // leaves `out` untouched.
  std::string doc = "{\"format\":1,\"rows\":[";
  for (uint32_t row = 0; row < rows; ++row) {
    if (want_json) {
      doc += (row != 0 ? ",[" : "[");
    }
    for (uint32_t col = 0; col < cols; ++col) {
      fm_value_t value{};
      const size_t index = static_cast<size_t>(row) * cols + col;
      if (auto rc = fm_value_array_get_value_at(result, index, &value); rc != 0) {
        emit_last_error(err, "eval");
        return rc;
      }
      if (want_json) {
        doc += (col != 0 ? "," : "") + render_value_json(value);
        continue;
      }
      if (col != 0) {
        out << '\t';
      }
      // The plain grid is TAB-separated and newline-terminated, so a cell
      // carrying either character would add fields and rows a consumer
      // reads as real ones. `dump` escapes for the same reason.
      out << escape_single_line(render_value(value));
    }
    if (want_json) {
      doc += ']';
    } else {
      out << '\n';
    }
  }
  if (want_json) {
    out << doc << "]}\n";
  }
  return 0;
}

// Reports a failed evaluation. A formula syntax error is exit 65 with its
// position; anything else is the engine's message and the generic failure.
int report_eval_failure(std::ostream& err) {
  fm_parse_error_t position = FM_STRUCT_INIT(fm_parse_error_t);
  if (fm_get_last_parse_error(&position) == 0) {
    err << "formulon: eval: " << fm_get_last_error_message() << " at " << position.line << ':' << position.column
        << '\n';
    return kExitDataErr;
  }
  emit_last_error(err, "eval");
  return kExitFailure;
}

// Reads the formula from stdin, without its trailing line break.
int read_formula_from_stdin(std::string* formula, std::ostream& err) {
  std::vector<std::uint8_t> bytes;
  if (auto rc = read_stdin(bytes); rc != 0) {
    if (rc == kExitUsage) {
      err << "formulon: eval: stdin is a terminal; pipe a formula into '-' or pass it as an argument\n";
    } else {
      err << "formulon: eval: cannot read stdin: " << std::strerror(errno) << '\n';
    }
    return rc;
  }
  formula->assign(bytes.begin(), bytes.end());
  while (!formula->empty() && (formula->back() == '\n' || formula->back() == '\r')) {
    formula->pop_back();
  }
  return 0;
}

}  // namespace

int run_eval(const ArgList& args, std::ostream& out, std::ostream& err) {
  bool want_json = false;
  bool local = false;
  std::optional<std::string> workbook_path;
  std::optional<std::string> sheet_spec;
  // Engaged only when `--repeat` was actually given. The iteration count
  // and the decision to report timing are the same fact, so they live in
  // one value; holding the count in a plain `long` invites keying the
  // report off its magnitude, which reports nothing for an explicit 1.
  std::optional<long> repeat;
  std::string_view formula;
  bool formula_seen = false;
  bool options_ended = false;
  SharedOptions shared;
  int exit_code = 0;

  for (std::size_t i = 0; i < args.size(); ++i) {
    const std::string_view a = args[i];
    switch (handle_common_option(args, i, options_ended, shared, print_eval_usage, "eval", out, err, exit_code)) {
      case CommonOption::kConsumed:
        continue;
      case CommonOption::kExit:
        return exit_code;
      case CommonOption::kNotCommon:
        break;
    }
    if (!options_ended && a == "--json") {
      want_json = true;
      continue;
    }
    if (!options_ended && a == "--local") {
      local = true;
      continue;
    }
    if (!options_ended && (a == "--in" || a == "--sheet")) {
      if (i + 1 >= args.size()) {
        err << "formulon: eval: " << a << " requires a value\n";
        return kExitUsage;
      }
      (a == "--in" ? workbook_path : sheet_spec).emplace(args[i + 1]);
      ++i;
      continue;
    }
    if (!options_ended && a == "--repeat") {
      // `--repeat N` consumes one extra argument. `parse_repeat`
      // increments the index it receives once so it points at the next
      // unconsumed argument; we leave the outer loop's own `++i` to
      // step past `--repeat` itself, which is why we pass `i + 1` and
      // copy the result back rather than &i directly.
      std::size_t value_index = i + 1;
      const long n = parse_repeat(args, value_index);
      if (n < 0) {
        err << "formulon: eval: --repeat requires a positive integer\n";
        return kExitUsage;
      }
      repeat = n;
      // Skip the value: the outer `++i` handles `--repeat`, and we
      // bump `i` once more here for the value.
      i = value_index - 1;
      continue;
    }
    if (!accept_positional(a, options_ended, formula_seen, "eval", "formula", err, /*allow_dash=*/true)) {
      return kExitUsage;
    }
    formula = a;
    formula_seen = true;
  }

  if (!formula_seen) {
    return missing_positional("eval", "formula", print_eval_usage, err);
  }

  if (sheet_spec.has_value() && !workbook_path.has_value()) {
    err << "formulon: eval: --sheet requires --in\n";
    return kExitUsage;
  }
  std::string formula_str(formula);
  if (formula == kStdioPath) {
    if (workbook_path == kStdioPath) {
      err << "formulon: eval: the formula and --in cannot both read stdin\n";
      return kExitUsage;
    }
    if (auto rc = read_formula_from_stdin(&formula_str, err); rc != 0) {
      return rc;
    }
  }

  WorkbookGuard wb;
  std::size_t sheet_index = 0;
  if (workbook_path.has_value()) {
    if (auto rc = load_workbook_file(*workbook_path, shared, wb, err, "eval"); rc != 0) {
      return rc;
    }
    if (auto rc = fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_AUTO); rc != 0) {
      emit_last_error(err, "eval");
      return kExitFailure;
    }
    if (sheet_spec.has_value()) {
      if (auto rc = resolve_sheet(wb.handle, *sheet_spec, &sheet_index, err, "eval"); rc != 0) {
        return rc;
      }
    }
  } else {
    if (auto rc = fm_workbook_create_empty(&wb.handle); rc != 0) {
      emit_last_error(err, "eval");
      return kExitFailure;
    }
    if (auto rc = apply_shared_options(wb.handle, shared); rc != 0) {
      emit_last_error(err, "eval");
      return kExitFailure;
    }
    if (auto rc = fm_workbook_add_sheet(wb.handle, "Sheet1", &sheet_index); rc != 0) {
      emit_last_error(err, "eval");
      return kExitFailure;
    }
  }
  if (local) {
    const char* profile_id = nullptr;
    const char* canonical = nullptr;
    if (auto rc = fm_workbook_get_profile_id(wb.handle, &profile_id); rc != 0) {
      emit_last_error(err, "eval");
      return kExitFailure;
    }
    if (auto rc = fm_canonicalize_formula(formula_str.c_str(), profile_id, &canonical); rc != 0) {
      emit_last_error(err, "eval");
      return kExitFailure;
    }
    formula_str = canonical;
  }
  // Evaluate ad hoc at A1. This keeps evaluation side-effect free: `=A1+1`
  // sees the empty cell instead of creating a circular A1 formula.

  // `--repeat N` invokes the complete ad-hoc evaluator for every pass;
  // without the flag the formula is still evaluated exactly once.
  using Clock = std::chrono::steady_clock;
  const long passes = repeat.value_or(1);
  ValueArrayGuard result;
  const auto t0 = Clock::now();
  for (long i = 0; i < passes; ++i) {
    fm_value_array_destroy(result.handle);
    result.handle = nullptr;
    if (auto rc = fm_sheet_evaluate_formula_array(wb.handle, sheet_index, 0, 0, formula_str.c_str(), &result.handle);
        rc != 0) {
      return report_eval_failure(err);
    }
  }
  const auto elapsed = Clock::now() - t0;

  if (render_eval_result(result.handle, want_json, out, err) != 0) {
    return kExitFailure;
  }

  if (const auto rc = flush_output(out, err, "eval"); rc != 0) {
    return rc;
  }

  // Asking for a measurement is what selects the report, not how large the
  // measurement turned out to be: every count `parse_repeat` accepts gets
  // the same timing line.
  if (repeat.has_value()) {
    const auto micros = std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count();
    err << "formulon: eval: " << *repeat << " iterations in " << micros << "us\n";
  }
  return 0;
}

}  // namespace cli
}  // namespace formulon
