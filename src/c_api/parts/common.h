//
// Shared private helpers for the C-ABI implementation split across
// `src/c_api/parts/*.cpp`. These leak no symbols outside the binary
// because the per-part TUs are linked into `formulon_core` only.
//
// Surface intentionally narrow: opaque-handle definition, thread-local
// diagnostic storage, the per-handle text-store + interning helper, the
// `value_to_fm` marshaller, and the validation helpers that every part
// needs.

#ifndef FORMULON_C_API_PARTS_COMMON_H_
#define FORMULON_C_API_PARTS_COMMON_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace c_api {
namespace parts {

// Per-handle UTF-8 storage. `std::deque` is used (matching the OOXML
// reader) for pointer stability: we hand out `c_str()` pointers from
// elements deep inside the container, so a `std::vector` reallocation
// would invalidate every previously surfaced view.
using TextStore = std::deque<std::string>;

// Per-handle stash for the most recent ad-hoc array evaluation
// (`fm_workbook_evaluate_formula_array`). The two-step array surface
// evaluates once, stashes the whole result here, then hands cells back one
// at a time by row-major index. The evaluation arena is discarded when the
// producing call returns, so any `Text` cell is deep-copied into `text_owner`
// and re-pointed at that owned storage: the stashed `Value`s stay valid until
// the next array evaluation (or handle destruction). Non-text cells are
// trivially copied; `value_to_fm` reports Array / Ref / Lambda cells by kind
// only and never dereferences their (now-dangling) arena pointers.
struct AdhocArrayStash {
  std::vector<formulon::Value> cells;  // row-major; size == rows * cols
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  TextStore text_owner;  // owns the bytes behind every Text cell above

  // Drops the previous evaluation's cells and owned text.
  void clear() {
    cells.clear();
    text_owner.clear();
    rows = 0;
    cols = 0;
  }
};

// Resets the thread-local diagnostics. Called at the start of every
// fallible entry point so a successful return surfaces an empty error
// message rather than the previous call's residue.
void clear_last_error();

// Captures `err` into the thread-local buffers and returns its numeric
// status code. The Error code is mirrored bit-for-bit into `fm_status_t`.
fm_status_t set_last_error(const formulon::Error& err);

// Convenience wrapper for the very common "binding misuse" case (NULL
// pointer argument, unknown handle, ...). `context` is appended verbatim;
// callers should keep it short and machine-friendly (key=value).
fm_status_t set_binding_error(formulon::FormulonErrorCode code, const char* message, std::string context = {});

// Returns the thread-local diagnostic message buffer (always non-null).
const char* last_error_message();

// Returns the thread-local diagnostic context buffer (always non-null).
const char* last_error_context();

// Hard cap on the number of cell ranges any single C-API call may carry.
// Validation lists and conditional-format sqrefs are unbounded on the
// wire (`uint32_t`), so a hostile caller can pass `0xFFFFFFFF` and force
// a 4 GiB `reserve()` before the loop body even runs. 16384 is several
// orders of magnitude past Excel's own usage (sheets rarely exceed a
// handful of CF sqrefs) but high enough that no legitimate caller will
// trip the cap.
constexpr std::uint32_t kMaxRangesPerCApiCall = 16384U;

// Returns `false` and sets the thread-local binding error if `n` is past
// the per-call range cap. The caller threads the result back as the
// fail-fast path so the body never starts the per-range loop with a
// bogus reservation.
bool check_range_count(std::uint32_t n, const char* api);

// Translates a `Value` into a C-side `fm_value_t`. For text variants the
// payload is appended to `store` and the returned pointer is a stable,
// NUL-terminated `c_str()` into it. Read-path callers pass the per-handle
// `read_scratch` (cleared each call) so returned pointers stay valid only
// until the next read; intern-path callers pass long-lived storage.
void value_to_fm(const formulon::Value& v, TextStore& store, fm_value_t* out);

// Validates a `(handle, sheet_index)` pair and returns the in-bounds
// status. On failure the diagnostic is already populated; callers should
// just propagate the returned code. `sheet_index` is the `size_t`-typed
// API variant used by most pre-existing entry points.
fm_status_t check_sheet_index(const fm_workbook_t* wb, std::size_t sheet_index, const char* fn);

// Bounds-check for the uint32-typed sheet index used by the sheet UI
// surface. Mirrors `check_sheet_index` but accepts the ABI's uint32.
// Accepts `const` so read-only entry points (`fm_sheet_get_*`) can share
// the same validator as the mutators.
fm_status_t check_sheet_u32(const fm_workbook_t* wb, std::uint32_t sheet, const char* fn);

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

// Opaque handle definition (translation-unit-public so every part can
// access the workbook + per-handle text store without a getter
// indirection). Kept at the top level so the public C ABI's forward
// declaration `typedef struct fm_workbook fm_workbook_t;` matches by
// name.
//
// Holds a single `Workbook`. The workbook itself owns the text-storage
// deque that backs every `Value::text` view, so the load path no longer
// needs to keep an `OoxmlReadResult` alive beyond the call that
// produced it - the workbook is move-constructed out of the result and
// retains the storage.
struct fm_workbook {
  std::optional<formulon::Workbook> wb;

  // Scratch storage for strings handed back to the caller on the read
  // path (`fm_workbook_get_value` / `fm_workbook_cell_at` /
  // `fm_workbook_lambda_text_at`). Each fallible read entry point clears
  // this at its start, so it only ever holds the strings produced by the
  // single most recent read call. This bounds memory: a long-lived handle
  // that loops over reads no longer accumulates one entry per call. The
  // returned `const char*` is therefore valid only until the next read
  // call on the same handle (or any mutation, or handle destruction) —
  // the standard C-ABI scratch contract documented in `formulon_c.h`.
  formulon::c_api::parts::TextStore read_scratch;

  // Scratch buffers for `fm_sheet_cf_get_at` visual payload arrays. The
  // returned pointers follow the same read-scratch lifetime as textual
  // readbacks: valid until the next read/mutation on this handle.
  std::vector<fm_cfvo_t> cfvo_scratch;
  std::vector<fm_cf_color_t> cf_color_scratch;

  // Result of the most recent `fm_workbook_evaluate_formula_array`. Owns its
  // own text storage so cells stay readable via
  // `fm_workbook_evaluate_formula_array_cell` after the producing call's
  // arena is gone; superseded by the next array evaluation on this handle.
  formulon::c_api::parts::AdhocArrayStash adhoc_array;

  formulon::Workbook& workbook() { return *wb; }
  const formulon::Workbook& workbook() const { return *wb; }
};

#endif  // FORMULON_C_API_PARTS_COMMON_H_
