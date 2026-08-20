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
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/borrowed_arena.h"
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

// --- Input validation ------------------------------------------------
//
// Every C-ABI entry point that copies a caller-supplied aggregate into
// the model runs the matching `validate` overload below exactly once,
// before it touches any model state. That is what makes validation
// coverage a property of the argument type instead of a property of the
// call site: an entry point cannot accept `fm_cf_rule_t` without
// admitting the checks that type carries, and adding a field to a
// struct puts its check in the one place every surface picks up.
//
// The primitives below are what the overloads are built from. They are
// also called directly by the scalar entry points, which take the same
// values as loose arguments and have to reject the same domains.
//
// Each returns `0` on success and `kInvalidArgument` otherwise,
// having already populated the thread-local diagnostic; callers just
// propagate the returned code. `api` names the entry point and `field`
// the struct member, so the context string identifies the argument
// without the caller assembling a message.

// Rejects a rectangle that is inverted or reaches outside Excel's grid.
// Structured entry points that copy a caller-supplied rectangle into the
// model run this first, so downstream writers never observe a range they
// cannot encode. Mirrors the per-coordinate check the scalar coordinate
// entry points apply.
fm_status_t check_sheet_rect(std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                             std::uint32_t last_col, const char* api);

// Rejects a column span that is inverted or reaches past Excel's last
// column. Such a span reaches the worksheet part as a `<col min="..."
// max="...">` pair outside the grid, which Excel treats as a damaged
// worksheet rather than as an entry to ignore, so the whole file is
// repaired and the layout the caller authored is lost. Out-of-grid
// columns are refused rather than clamped: clamping would move the
// override onto the last addressable column, which is a column the
// caller never named.
fm_status_t check_column_span(std::uint32_t first, std::uint32_t last, const char* api);

// Rejects a row index past Excel's last row, for the same reason
// `check_column_span` rejects an out-of-grid span: `<row r="...">` past
// the grid is a repair prompt, not a discarded attribute.
fm_status_t check_row_index(std::uint32_t row, const char* api);

// Rejects a measurement that is not a finite, non-negative double. Row
// heights and column widths reach the file as `xsd:double` attributes,
// and the number writer renders a NaN or an infinity as an empty
// string, which is not a lexical `xsd:double` at all; a negative value
// is outside the measurement's own domain. Zero is accepted: it is the
// legitimate metric of a zero-width column or zero-height row.
fm_status_t check_finite_non_negative(double value, const char* api, const char* field);

// Rejects an ordinal outside `[0, max]`. `max` is the last declared
// enumerator of the model enum the field mirrors, so extending that enum
// widens the domain at the same time. A value past it would otherwise be
// `static_cast` into the enum and collapse onto whichever case the
// consuming switch treats as its default, which the caller observes as a
// silently different rule rather than as a rejected one.
fm_status_t check_enum_domain(std::int64_t value, std::int64_t max, const char* api, const char* field);

// Rejects a string that is not GUID-shaped, i.e. not
// `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` with hexadecimal digits and
// the braces present. `text` may be NULL or empty, which passes: an
// absent id is not a malformed one, and the entry point synthesizes one.
// Applied wherever a caller-supplied string reaches an XML attribute
// whose schema type is `ST_Guid` — Excel repairs (and so discards) the
// whole block when such an attribute holds anything else.
fm_status_t check_guid(const char* text, const char* api, const char* field);

// Validates a conditional-formatting rule record: sqref shape and every
// rectangle in it, plus the `type` / `op` / `time_period` /
// `icon_set_name` / `data_bar_axis_position` / `fm_cfvo_t::type` domains
// and the GUID shape of `id`. Checks that need the workbook (the
// `dxf_id` bound) stay with the entry point.
fm_status_t validate(const fm_cf_rule_t& rule, const char* api);

// Validates a data-validation record: range shape and rectangles, plus
// the `type` / `op` / `error_style` domains.
fm_status_t validate(const fm_data_validation& v, const char* api);

// Validates a pivot field spec: required strings and the `axis` domain.
fm_status_t validate(const fm_pivot_field_spec_t& spec, const char* api);

// Validates a pivot data-field spec: required strings and the
// `aggregation` / `show_as` domains.
fm_status_t validate(const fm_pivot_data_field_spec_t& spec, const char* api);

// Validates a pivot filter spec: the `axis` / `type` domains. The
// payload-variant and data-field checks need the pivot table and stay
// with the entry point.
fm_status_t validate(const fm_pivot_filter_spec_t& spec, const char* api);

// Validates a viewport rectangle. A collapsed rectangle is allowed (it
// recalculates nothing, as the header states), so only the sheet index
// is constrained here; the bound against `sheet_count` needs the
// workbook and stays with the entry point.
fm_status_t validate(const fm_viewport& viewport, const char* api);

// Translates a `Value` into a C-side `fm_value_t`. For text variants the
// payload is appended to `store` and the returned pointer is a stable,
// NUL-terminated `c_str()` into it. Read-path callers pass the per-handle
// `read_scratch`; after argument/model validation, each successful
// scratch-backed producer clears and refreshes that store. A
// validation-rejected call leaves it untouched, so returned pointers remain
// valid until the next successful scratch-backed read, mutation, or handle
// destruction. Intern-path callers pass long-lived storage.
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

  // Loss / fallback counters captured while loading this handle's package,
  // for either container format. They stay attached to the handle so callers
  // can inspect what the load cost after `fm_workbook_load` returns
  // successfully. A workbook built in memory leaves every counter zero.
  fm_read_diagnostics_t read_diagnostics{};

  // Scratch storage for strings handed back to the caller on the read path.
  // After argument/model validation, each successful text-producing read
  // clears and refreshes this store; the successful CF getter also refreshes
  // it while serializing threshold value strings. Validation-rejected calls
  // leave it untouched. This bounds memory: a long-lived handle that loops
  // over reads no longer accumulates one entry per call. Returned `const
  // char*` values are valid only until the next successful scratch-backed
  // read on the same handle (or any mutation, or handle destruction) — the
  // standard C-ABI scratch contract documented in `formulon_c.h`.
  formulon::c_api::parts::TextStore read_scratch;

  // Sorted coordinate cache for the `cell_count` / `cell_at` iteration
  // pair. The associated Sheet revision invalidates it after any cell,
  // spill, or structural mutation.
  struct CellEnumerationCache {
    std::size_t sheet_index = std::numeric_limits<std::size_t>::max();
    std::uint64_t revision = std::numeric_limits<std::uint64_t>::max();
    std::vector<std::pair<std::uint32_t, std::uint32_t>> addresses;
  };
  mutable CellEnumerationCache cell_enumeration_cache;

  // Storage behind every borrowed field `fm_sheet_cf_get_at` fills. They are
  // separate from `read_scratch` on purpose: only a successful
  // `fm_sheet_cf_get_at` clears and refreshes them, so neither an unrelated
  // read nor a CF mutation can invalidate a rule the caller is still holding.
  // That is what makes the ABI's only edit path - get, remove, add back -
  // safe to execute as documented. Validation-rejected CF calls leave them
  // untouched.
  //
  // The getter copies the rule out of the model rather than pointing at it:
  // a model-backed view would die with the block the caller is about to
  // remove. Each payload is handed to the arenas as one finished array, whose
  // address does not move as later payloads are adopted, so a rule engaging
  // several visual payloads cannot invalidate the pointers already written
  // for the earlier ones.
  formulon::c_api::BorrowedStringArena cf_text_scratch;
  formulon::c_api::BorrowedArrayArena<fm_cf_cell_range_t> cf_range_scratch;
  formulon::c_api::BorrowedArrayArena<fm_cfvo_t> cfvo_scratch;
  formulon::c_api::BorrowedArrayArena<fm_cf_color_t> cf_color_scratch;

  // Result of the most recent `fm_workbook_evaluate_formula_array`. Owns its
  // own text storage so cells stay readable via
  // `fm_workbook_evaluate_formula_array_cell` after the producing call's
  // arena is gone; superseded by the next array evaluation on this handle.
  formulon::c_api::parts::AdhocArrayStash adhoc_array;

  // Iterative-solver progress callback as the C caller registered it. The
  // engine's own callback type returns `bool`, which the C ABI does not use
  // in any declaration, so the engine is handed a fixed adapter with this
  // handle as its `user_data`; the adapter reads the pair below and narrows
  // the caller's `int32_t` to the engine's `bool`. Cleared by passing a NULL
  // callback to `fm_workbook_set_iterative_progress`.
  fm_iterative_progress_cb iterative_progress_cb = nullptr;
  void* iterative_progress_user_data = nullptr;

  formulon::Workbook& workbook() { return *wb; }
  const formulon::Workbook& workbook() const { return *wb; }
};

#endif  // FORMULON_C_API_PARTS_COMMON_H_
