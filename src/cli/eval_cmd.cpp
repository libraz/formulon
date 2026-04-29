// Copyright 2026 libraz. Licensed under the MIT License.
//
// `formulon eval <formula>` — evaluate a single formula on a fresh,
// empty workbook and print the result.
//
// The handler creates an empty workbook, adds a single sheet, parks the
// formula at A1, drives a recalc, and reads the cached value back via
// the C ABI. Cell-level errors (`#NAME?`, `#DIV/0!`, …) are *not*
// process-level failures: the cell value just happens to be an
// `FM_VAL_ERROR`, and we print its display name. Only structural
// failures (NULL handle, parser refusal, save/load) bubble out.
//
// `--repeat N` is a Phase 5 prep knob: it runs the eval `N` times and
// reports timing on stderr without changing the stdout payload.

#include "cli/cli.h"

#include <chrono>
#include <cstdint>
#include <cstdlib>
#include <ostream>
#include <string>
#include <string_view>

#include "c_api/formulon_c.h"
#include "cli/render.h"
#include "utils/error.h"

namespace formulon {
namespace cli {

namespace {

// RAII wrapper so the workbook handle is released even when an early
// return path is taken. Mirrors the helper in `tests/c_api/`.
struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
};

void print_eval_usage(std::ostream& out) {
  out << "Usage: formulon eval [--json] [--repeat N] <formula>\n"
      << "\n"
      << "Evaluate <formula> on a fresh empty workbook and print the result.\n"
      << "Cell-level Excel errors (#NAME?, #DIV/0!, ...) print to stdout and\n"
      << "do NOT cause a non-zero exit code; only structural failures do.\n";
}

// Logs the most recent `fm_*` diagnostic to `err`, prefixed with the
// subcommand name so the output is grep-friendly.
void emit_last_error(std::ostream& err, const char* subcommand) {
  err << "formulon: " << subcommand << ": " << fm_last_error_message();
  const char* ctx = fm_last_error_context();
  if (ctx != nullptr && ctx[0] != '\0') {
    err << " (" << ctx << ')';
  }
  err << '\n';
}

// Parses `--repeat` argument from `args[idx]`. Returns the parsed
// count, or `-1` on parse failure (which the caller turns into a usage
// error). `idx` is advanced past the value on success.
long parse_repeat(const ArgList& args, std::size_t& idx) {
  if (idx >= args.size()) {
    return -1;
  }
  const std::string s(args[idx]);
  ++idx;
  // `std::strtol` accepts trailing junk; treat any non-empty tail as
  // a parse failure. Also reject zero / negative counts because they
  // would skip the `recalc` entirely.
  char* end = nullptr;
  const long n = std::strtol(s.c_str(), &end, 10);
  if (end == nullptr || *end != '\0' || n <= 0) {
    return -1;
  }
  return n;
}

}  // namespace

int run_eval(const ArgList& args, std::ostream& out, std::ostream& err) {
  bool want_json = false;
  long repeat = 1;
  std::string_view formula;
  bool formula_seen = false;

  for (std::size_t i = 0; i < args.size(); ++i) {
    const std::string_view a = args[i];
    if (a == "-h" || a == "--help") {
      print_eval_usage(out);
      return 0;
    }
    if (a == "--json") {
      want_json = true;
      continue;
    }
    if (a == "--repeat") {
      // `--repeat N` consumes one extra argument. `parse_repeat`
      // increments the index it receives once so it points at the next
      // unconsumed argument; we leave the outer loop's own `++i` to
      // step past `--repeat` itself, which is why we pass `i + 1` and
      // copy the result back rather than &i directly.
      std::size_t value_idx = i + 1;
      const long n = parse_repeat(args, value_idx);
      if (n < 0) {
        err << "formulon: eval: --repeat requires a positive integer\n";
        return kExitUsage;
      }
      repeat = n;
      // Skip the value: the outer `++i` handles `--repeat`, and we
      // bump `i` once more here for the value.
      i = value_idx - 1;
      continue;
    }
    if (!a.empty() && a[0] == '-') {
      err << "formulon: eval: unknown flag '" << a << "'\n";
      return kExitUsage;
    }
    if (formula_seen) {
      err << "formulon: eval: only one formula may be supplied\n";
      return kExitUsage;
    }
    formula = a;
    formula_seen = true;
  }

  if (!formula_seen) {
    print_eval_usage(err);
    return kExitUsage;
  }

  WorkbookGuard wb;
  if (auto rc = fm_workbook_create_empty(&wb.handle); rc != 0) {
    emit_last_error(err, "eval");
    return rc;
  }
  if (auto rc = fm_workbook_add_sheet(wb.handle, "Sheet1"); rc != 0) {
    emit_last_error(err, "eval");
    return rc;
  }
  // Strip an optional leading `=` so users can pass either form.
  std::string formula_str(formula);
  if (auto rc = fm_workbook_set_formula(wb.handle, 0, 0, 0, formula_str.c_str()); rc != 0) {
    emit_last_error(err, "eval");
    return rc;
  }

  using Clock = std::chrono::steady_clock;
  const auto t0 = Clock::now();
  for (long i = 0; i < repeat; ++i) {
    if (auto rc = fm_workbook_recalc(wb.handle); rc != 0) {
      emit_last_error(err, "eval");
      return rc;
    }
  }
  const auto elapsed = Clock::now() - t0;

  fm_value_t v{};
  if (auto rc = fm_workbook_get_value(wb.handle, 0, 0, 0, &v); rc != 0) {
    emit_last_error(err, "eval");
    return rc;
  }

  if (want_json) {
    out << render_value_json(v) << '\n';
  } else {
    out << render_value(v) << '\n';
  }

  if (repeat > 1) {
    const auto micros = std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count();
    err << "formulon: eval: " << repeat << " iterations in " << micros << "us\n";
  }
  return 0;
}

}  // namespace cli
}  // namespace formulon
