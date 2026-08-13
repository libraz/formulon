//
// `formulon eval <formula>` — evaluate a single formula on a fresh,
// empty workbook and print the result.
//
// The handler creates an empty workbook, adds a single sheet, and evaluates
// the formula through the C ABI's side-effect-free array surface. Cell-level
// errors (`#NAME?`, `#DIV/0!`, …) are *not*
// process-level failures: the cell value just happens to be an
// `FM_VAL_ERROR`, and we print its display name. Only structural
// failures (NULL handle, parser refusal, save/load) bubble out.
//
// `--repeat N` re-evaluates the same formula `N` times through the ad-hoc
// evaluator. Timing is reported on stderr without changing the stdout payload.

#include <cerrno>
#include <chrono>
#include <cstdint>
#include <cstdlib>
#include <ostream>
#include <string>
#include <string_view>

#include "c_api/formulon_c.h"
#include "cli/cli.h"
#include "cli/render.h"
#include "parser/parser.h"
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
  out << "Usage: formulon eval [--json] [--repeat N] [--] <formula>\n"
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
  errno = 0;
  const long n = std::strtol(s.c_str(), &end, 10);
  if (errno == ERANGE || end == nullptr || *end != '\0' || n <= 0) {
    return -1;
  }
  return n;
}

// The evaluator intentionally turns parser recovery placeholders into
// cell-level `#NAME?` values so a workbook can retain invalid formulas for
// repair. A one-off CLI expression is different: reject malformed syntax
// before evaluation, while still letting a syntactically valid unknown
// function produce Excel's ordinary `#NAME?` value.
bool validate_formula_syntax(std::string_view formula, std::ostream& err) {
  Arena arena;
  parser::Parser parser(formula, arena);
  (void)parser.parse();
  if (parser.errors().empty()) {
    return true;
  }
  err << "formulon: eval: invalid formula syntax: " << parser.errors().front().message << '\n';
  return false;
}

int render_eval_result(const fm_workbook_t* wb, uint32_t rows, uint32_t cols, bool want_json, std::ostream& out,
                       std::ostream& err) {
  if (want_json && rows == 1 && cols == 1) {
    fm_value_t value{};
    if (auto rc = fm_workbook_evaluate_formula_array_cell(wb, 0, &value); rc != 0) {
      emit_last_error(err, "eval");
      return rc;
    }
    out << render_value_json(value) << '\n';
    return 0;
  }
  if (want_json) {
    out << '[';
  }
  for (uint32_t row = 0; row < rows; ++row) {
    if (want_json) {
      if (row != 0) {
        out << ',';
      }
      out << '[';
    }
    for (uint32_t col = 0; col < cols; ++col) {
      fm_value_t value{};
      const size_t index = static_cast<size_t>(row) * cols + col;
      if (auto rc = fm_workbook_evaluate_formula_array_cell(wb, index, &value); rc != 0) {
        emit_last_error(err, "eval");
        return rc;
      }
      if (col != 0) {
        out << (want_json ? ',' : '\t');
      }
      out << (want_json ? render_value_json(value) : render_value(value));
    }
    if (want_json) {
      out << ']';
    } else {
      out << '\n';
    }
  }
  if (want_json) {
    out << "]\n";
  }
  return 0;
}

}  // namespace

int run_eval(const ArgList& args, std::ostream& out, std::ostream& err) {
  bool want_json = false;
  long repeat = 1;
  std::string_view formula;
  bool formula_seen = false;
  bool options_ended = false;

  for (std::size_t i = 0; i < args.size(); ++i) {
    const std::string_view a = args[i];
    if (!options_ended && a == "--") {
      options_ended = true;
      continue;
    }
    if (!options_ended && (a == "-h" || a == "--help")) {
      print_eval_usage(out);
      return 0;
    }
    if (!options_ended && a == "--json") {
      want_json = true;
      continue;
    }
    if (!options_ended && a == "--repeat") {
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
    if (!options_ended && !a.empty() && a[0] == '-') {
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
  // Evaluate ad hoc at A1. This keeps evaluation side-effect free: `=A1+1`
  // sees the empty cell instead of creating a circular A1 formula.
  std::string formula_str(formula);
  if (!validate_formula_syntax(formula_str, err)) {
    return static_cast<int>(FormulonErrorCode::kParserUnexpectedToken);
  }

  // `--repeat N` invokes the complete ad-hoc evaluator for every pass.
  using Clock = std::chrono::steady_clock;
  uint32_t result_rows = 0;
  uint32_t result_cols = 0;
  const auto t0 = Clock::now();
  for (long i = 0; i < repeat; ++i) {
    if (auto rc =
            fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 0, formula_str.c_str(), &result_rows, &result_cols);
        rc != 0) {
      emit_last_error(err, "eval");
      return rc;
    }
  }
  const auto elapsed = Clock::now() - t0;

  if (auto rc = render_eval_result(wb.handle, result_rows, result_cols, want_json, out, err); rc != 0) {
    return rc;
  }

  if (const auto rc = flush_output(out, err, "eval"); rc != 0) {
    return rc;
  }

  if (repeat > 1) {
    const auto micros = std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count();
    err << "formulon: eval: " << repeat << " iterations in " << micros << "us\n";
  }
  return 0;
}

}  // namespace cli
}  // namespace formulon
