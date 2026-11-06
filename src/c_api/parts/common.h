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

#include <array>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <initializer_list>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/borrowed_arena.h"
#include "c_api/formulon_c.h"
#include "utils/error.h"
#include "utils/strings.h"
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

// Appends a copy to `store` and returns its stable NUL-terminated view.
const char* store_cstr(TextStore& store, std::string_view text);

// Allocates an owned copy of `bytes` for a C-ABI out parameter. The caller
// remains responsible for validating the out pointers before calling this;
// zero-length input intentionally still uses `new[]` so `fm_buffer_destroy` can
// release the returned pointer exactly like a non-empty buffer.
void copy_owned_bytes(const std::vector<std::uint8_t>& bytes, std::uint8_t** out_bytes, std::size_t* out_len);

// Resets the thread-local diagnostics. Called at the start of every
// fallible entry point so a successful return surfaces an empty error
// message rather than the previous call's residue.
void clear_last_error();

// Captures `err` into the thread-local buffers and returns its numeric
// status code. The Error code is mirrored bit-for-bit into `fm_status_t`.
fm_status_t set_last_error(const formulon::Error& err);

// Maps a void result to a status: 0 on success, else `set_last_error` of its error.
fm_status_t finish_status(const formulon::Expected<void, formulon::Error>& r);

// Convenience wrapper for the very common "binding misuse" case (NULL
// pointer argument, unknown handle, ...). `context` is appended verbatim;
// callers should keep it short and machine-friendly (key=value).
fm_status_t set_binding_error(formulon::FormulonErrorCode code, const char* message, std::string context);
fm_status_t set_binding_error(formulon::FormulonErrorCode code, const char* message);

// Binding error for a NULL required argument: "a required argument is NULL".
fm_status_t null_argument();

// Binding error for a NULL required argument: "<what> is NULL".
fm_status_t null_argument(const char* what);

// kInvalidArgument binding error with the sentence `what`, carrying `context`.
fm_status_t invalid_argument(const char* what, std::string context = std::string());

// True when `error` is an `ErrorCode` ordinal (0 through `ErrorCode::Unknown`).
bool is_valid_error_code(fm_error_code_t error);

// kInvalidArgument binding error "error code out of range" for `error`.
fm_status_t invalid_error_code(fm_error_code_t error);

// kNotFound binding error with the sentence `what`, carrying `context`: a
// natural key (name, id, cell) that names nothing.
fm_status_t not_found(const char* what, std::string context = std::string());

// `text` as an owned string; NULL reads as empty.
inline std::string text_or_empty(const char* text) {
  return text != nullptr ? std::string(text) : std::string();
}

// Rejects `text` that is not well-formed UTF-8 (`sheet_names::valid_utf8`:
// surrogate code points, overlong forms and truncated sequences fail) with
// `kBindingUtf8EncodingError` ("text is not valid UTF-8", `context`).
// The context string is built only on failure; the keyed forms spell it
// `<key>=<value>` / `<key>=<index>`.
fm_status_t check_utf8(std::string_view text, std::string_view context);
fm_status_t check_utf8(std::string_view text, std::string_view key, std::string_view value);
fm_status_t check_utf8(std::string_view text, std::string_view key, std::size_t index);

// `check_utf8` of every non-NULL entry of `texts`, context `text_index=<i>`.
// NULL entries are left to the caller's own NULL rule.
fm_status_t check_utf8(std::initializer_list<const char*> texts);

// Entry form of formula text stored outside a cell (CF, DV, defined name, CF
// threshold): a non-empty `formula` must pass the entry parse check
// (`check_entered_cell_formula` without its bind and R1C1 fallback), and
// `*out` receives the stored spelling a cell would take
// (`Workbook::normalize_formula_text`, `kEntered`) without its one leading
// '=', the spelling files carry and readers produce; a second leading '=' is
// refused. An empty `formula` yields "".
fm_status_t check_stored_formula(std::string_view formula, std::string_view context, std::string* out);

// Stored form of a `formula` that passed `check_stored_formula`, taken as the
// entry commits: `Workbook::ingest_entered_formula` (which binds the books it
// names) without the one leading '='. An empty `formula` yields "".
std::string ingest_entered_formula(fm_workbook_t* wb, std::string_view formula);

// Entry check for formula text entered at (`sheet`, `row`, `col`): refuses
// malformed UTF-8 (`check_utf8`), binds the structured references at that
// cell (`Workbook::bind_entered_structured_refs`: `[@Qty]` and `ROWS(T)` take
// their stored forms), then on the bound text refuses the whitespace Excel
// refuses (`parser::has_refused_whitespace`: a tab, a carriage return, a run
// between a built-in's name and its `(`) and requires
// `eval::parse_formula_entry` to accept its storage normalization. Text that
// fails that, or that `eval::needs_r1c1_entry` sends to R1C1, is taken as
// `eval::translate_r1c1_entry` at the cell when that passes the same checks.
// Typed `_xlpm.` before a cell-shaped name is refused everywhere. Returns 0,
// or sets `kParserUnexpectedToken` ("formula does not parse") with
// `context`. `*bound` receives the bound (or translated) text with one
// leading '=', the spelling the readers store, not yet normalized, for
// `Workbook::set_cell_formula`.
fm_status_t check_entered_cell_formula(const fm_workbook_t* wb, std::size_t sheet, std::uint32_t row, std::uint32_t col,
                                       std::string_view formula, std::string_view context, std::string* bound);

// Display name of the table whose range holds (`sheet`, `row`, `col`), or ""
// outside every table.
std::string_view table_display_name_at(const formulon::Workbook& book, std::size_t sheet, std::uint32_t row,
                                       std::uint32_t col);

// Whether a getter may spell `model` differently from the model: the
// structured-reference display form and the self-book spelling need a `[`,
// a cell-shaped name's display form an `_xl` prefix.
inline bool may_respell_formula(std::string_view model) noexcept {
  return model.find('[') != std::string_view::npos || formulon::strings::case_insensitive_contains(model, "_xl");
}

// `model` as the getters spell it for a formula in `own_table`'s range
// (empty outside every table): `parser::display_structured_refs`, then
// `Workbook::spell_self_book_names`, then `parser::display_cell_shaped_names`.
std::string display_formula_text(const formulon::Workbook& book, std::string_view model, std::string_view own_table);

// Formula text as every getter hands it out: `text` with one leading '=',
// which the stored spelling outside a cell omits. Empty text (no formula)
// stays empty, and text that already starts with '=' is kept as it is.
inline std::string with_formula_prefix(std::string_view text) {
  if (text.empty() || text.front() == '=') {
    return std::string(text);
  }
  std::string out;
  out.reserve(text.size() + 1U);
  out += '=';
  out.append(text);
  return out;
}

// Model formula text as a getter hands it out: `model`'s own bytes when
// `display_formula_text` leaves it unchanged, else that spelling appended to
// `store`.
const char* display_formula(const formulon::Workbook& book, const std::string& model, TextStore& store,
                            std::string_view own_table);

// Returns the thread-local diagnostic message buffer (always non-null).
const char* last_error_message();

// Returns the thread-local diagnostic context buffer (always non-null).
const char* last_error_context();

// Hard cap on the number of cell ranges any single C-API call may carry.
// Validation lists and conditional-format sqrefs are unbounded on the
// wire (`size_t`), so a hostile caller can pass `SIZE_MAX` and force
// a huge `reserve()` before the loop body even runs. 16384 is several
// orders of magnitude past Excel's own usage (sheets rarely exceed a
// handful of CF sqrefs) but high enough that no legitimate caller will
// trip the cap.
constexpr std::uint32_t kMaxRangesPerCApiCall = 16384U;

// Returns `false` and sets the thread-local binding error if `n` is past
// the per-call range cap. The caller threads the result back as the
// fail-fast path so the body never starts the per-range loop with a
// bogus reservation.
bool check_range_count(std::size_t n);

// Validates a required `const fm_cell_rect_t* ranges, size_t range_count`
// list: at least one entry (`kInvalidArgument`), `ranges` non-NULL
// (`kBindingNullPointer`), at most the per-call cap, and every rectangle
// ordered and inside the grid (`kInvalidArgument`).
fm_status_t check_rect_list(const fm_cell_rect_t* ranges, std::size_t range_count);

// Field-for-field conversions between the engine range and the C rectangle.
inline fm_cell_rect_t to_cell_rect(const formulon::MergeRange& r) {
  return fm_cell_rect_t{r.first_row, r.first_col, r.last_row, r.last_col};
}
inline formulon::MergeRange to_merge_range(const fm_cell_rect_t& r) {
  return formulon::MergeRange{r.first_row, r.first_col, r.last_row, r.last_col};
}

// Copies `ranges` into the handle's rectangle scratch (read-scratch-backed
// lifetime) and returns the published array, or NULL when it is empty.
const fm_cell_rect_t* publish_rects(const fm_workbook_t* wb, const std::vector<formulon::MergeRange>& ranges);

// A1 text of an in-grid, ordered `rect`, qualified by `sheet_name` when it is
// non-empty: one endpoint for a single cell, `B:D` for a rectangle spanning
// every row, `2:11` for one spanning every column, `$` anchors when
// `absolute`.
std::string format_rect_a1(std::string_view sheet_name, const fm_cell_rect_t& rect, bool absolute);

// Resolves a non-NULL `profile_id` to its `ExcelProfile`. Returns 0, or sets
// `kInvalidArgument` ("unknown profile", context `profile_id=<id>`).
fm_status_t check_profile_id(const char* profile_id, formulon::ExcelProfile* out);

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
// propagate the returned code. `field` names the struct member, so the
// context string identifies the argument without the caller assembling a
// message.

// Rejects a rectangle that is inverted or reaches outside Excel's grid.
// Structured entry points that copy a caller-supplied rectangle into the
// model run this first, so downstream writers never observe a range they
// cannot encode. Mirrors the per-coordinate check the scalar coordinate
// entry points apply.
fm_status_t check_sheet_rect(std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                             std::uint32_t last_col);
fm_status_t check_sheet_rect(const fm_cell_rect_t& rect);

// The rectangle `rows` x `cols` cells whose top-left cell is `(row, col)`;
// both extents are at least 1.
inline fm_cell_rect_t cell_rect_from_extent(std::uint32_t row, std::uint32_t col, std::uint32_t rows,
                                            std::uint32_t cols) {
  return fm_cell_rect_t{row, col, row + rows - 1U, col + cols - 1U};
}

// Rejects a column span that is inverted or reaches past Excel's last
// column. Such a span reaches the worksheet part as a `<col min="..."
// max="...">` pair outside the grid, which Excel treats as a damaged
// worksheet rather than as an entry to ignore, so the whole file is
// repaired and the layout the caller authored is lost. Out-of-grid
// columns are refused rather than clamped: clamping would move the
// override onto the last addressable column, which is a column the
// caller never named.
fm_status_t check_column_span(std::uint32_t first, std::uint32_t last);

// Rejects a row index past Excel's last row, for the same reason
// `check_column_span` rejects an out-of-grid span: `<row r="...">` past
// the grid is a repair prompt, not a discarded attribute.
fm_status_t check_row_index(std::uint32_t row);

// Row counterpart of `check_column_span`: `first <= last` inside the grid.
fm_status_t check_row_span(std::uint32_t first, std::uint32_t last);

// Rejects a measurement that is not a finite, non-negative double. Row
// heights and column widths reach the file as `xsd:double` attributes,
// and the number writer renders a NaN or an infinity as an empty
// string, which is not a lexical `xsd:double` at all; a negative value
// is outside the measurement's own domain. Zero is accepted: it is the
// legitimate metric of a zero-width column or zero-height row.
fm_status_t check_finite_non_negative(double value, const char* field);

// Rejects a value that is not a finite double. A Number-kind `Value` is
// otherwise the engine's own invariant everywhere else -- every arithmetic
// result that would be NaN/Inf is converted to `#NUM!` before it can be
// observed (`eval/scalar_ops.cpp`) -- and `save()` downgrades a non-finite
// literal to `#NUM!` on write (`io/ooxml_writer_cell.cpp`), so a number
// setter that skipped this check would let ISNUMBER and arithmetic
// disagree about the cell right up until the first save.
fm_status_t check_finite(double value, const char* field);

// Converts a caller-supplied scalar into a `Value`, rejecting a non-finite
// number, a NULL or malformed-UTF-8 text, an out-of-range error code, and the array / reference
// / lambda kinds. A Text value aliases the caller's string, so the result is
// valid only for the duration of the call.
fm_status_t value_from_fm(const fm_value_t& in, formulon::Value* out);

// Rejects an ordinal outside `[min, max]`. `min` is `0`, or `-1` for an enum
// whose `-1` member means "absent". `max` is the last declared enumerator
// of the model enum the field mirrors, so extending that enum widens the
// domain at the same time. A value past it would otherwise be `static_cast`
// into the enum and collapse onto whichever case the consuming switch treats
// as its default, which the caller observes as a silently different rule
// rather than as a rejected one.
fm_status_t check_enum_domain(std::int64_t value, std::int64_t min, std::int64_t max, const char* field);

// Rejects a string that is not GUID-shaped, i.e. not
// `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` with hexadecimal digits and
// the braces present. `text` may be NULL or empty, which passes: an
// absent id is not a malformed one, and the entry point synthesizes one.
// Applied wherever a caller-supplied string reaches an XML attribute
// whose schema type is `ST_Guid` — Excel repairs (and so discards) the
// whole block when such an attribute holds anything else.
fm_status_t check_guid(const char* text, const char* field);

// Returns whether `id` is a built-in number-format id in some locale or one
// registered in `styles.num_fmts`.
bool num_fmt_id_known(const formulon::StylesTable& styles, std::uint16_t id);

// Rejects a `numFmtId` string that is malformed UTF-8 (`check_utf8`) or is not
// decimal digits naming an id
// `num_fmt_id_known` accepts. NULL and the empty string pass: both mean
// "no number format".
fm_status_t check_num_fmt_id_text(const formulon::Workbook& wb, const char* text);

// Validates a conditional-formatting rule record: sqref shape and every
// rectangle in it, plus the `type` / `op` / `time_period` /
// `icon_set_name` / `data_bar_axis_position` / `fm_cfvo_t::type` domains
// the GUID shape of `id` and the UTF-8 of `text`. Checks that need the workbook (the
// `dxf_id` bound) stay with the entry point.
fm_status_t validate(const fm_cf_rule_t& rule);

// Validates a data-validation record: range shape and rectangles, the
// `type` / `op` / `error_style` domains, and the UTF-8 of the four messages.
fm_status_t validate(const fm_data_validation_t& v);

// Validates a pivot field spec: required and UTF-8 strings and the `axis` domain.
fm_status_t validate(const fm_pivot_field_spec_t& spec);

// Validates a pivot data-field spec: UTF-8 strings and the
// `aggregation` / `show_as` domains.
fm_status_t validate(const fm_pivot_data_field_spec_t& spec);

// Validates a pivot filter spec: UTF-8 strings and the `axis` / `type` domains. The
// payload-variant and data-field checks need the pivot table and stay
// with the entry point.
fm_status_t validate(const fm_pivot_filter_spec_t& spec);

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
fm_status_t check_sheet_index(const fm_workbook_t* wb, std::size_t sheet_index);

// Clears the handle's read scratch and returns it. The store is logically
// internal (it never affects observable workbook state), so a const handle
// may refresh it; returned text pointers stay valid until the next read.
TextStore& fresh_read_scratch(const fm_workbook_t* wb);

// Replaces the read scratch with `text` and returns its stable `c_str()`.
const char* publish_scratch(const fm_workbook_t* wb, std::string text);

// Rejects a cell coordinate outside the sheet grid, naming the offending
// row / column in the diagnostic context.
fm_status_t check_cell_coord(std::uint32_t row, std::uint32_t col);

// Bounds-checks a caller-supplied element index against a collection of
// `size` entries. On failure sets `kIndexOutOfRange` with the message
// `"<field> <index> is out of range (<size> items)"` and an
// `<field>=<index> count=<size>` context, and returns the error status.
fm_status_t check_index(std::size_t index, std::size_t size, const char* field);

// Bounds-checks a caller-supplied limit value (freeze counts, outline level)
// against the exclusive `bound` it must stay below. Unlike `check_index` the
// value names no element of a collection, so a failure sets
// `kInvalidArgument` with the message
// `"<field> <value> is out of range (maximum <bound - 1>)"` and a
// `<field>=<value> max=<bound - 1>` context.
fm_status_t check_limit(std::size_t value, std::size_t bound, const char* field);

// Narrows an ABI `size_t` index for an engine call taking `uint32_t`,
// saturating so the engine's own range check refuses an oversized index
// instead of a truncated one naming another element.
inline std::uint32_t saturate_index(std::size_t index) {
  return index > std::numeric_limits<std::uint32_t>::max() ? std::numeric_limits<std::uint32_t>::max()
                                                           : static_cast<std::uint32_t>(index);
}

// `check_sheet_index` followed by `check_row_index`.
fm_status_t check_sheet_row(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row);

// Rejects an `anchor_kind` / `edit_as` ordinal outside the drawing anchor enums.
fm_status_t check_anchor_enums(std::int64_t anchor_kind, std::int64_t edit_as);

// Deepest row / column outline level the file formats encode (`outlineLevel`
// is 0-7 in OOXML and three bits in XLSB).
constexpr std::uint8_t kMaxOutlineLevel = 7U;

// `check_index` of `xf_index` against the workbook's cell-XF table. Every
// setter storing a cell, row or column `xf_index` runs it before writing.
fm_status_t check_cell_xf_index(const fm_workbook_t* wb, std::size_t xf_index);

// Rejects hiding or removing sheet `index` when it is the workbook's only
// visible sheet: Excel repairs a file whose sheets are all hidden. Returns
// `kInvalidArgument` ("workbook needs a visible sheet") or 0.
fm_status_t check_keeps_visible_sheet(const formulon::Workbook& workbook, std::uint32_t index);

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

  // Sorted coordinate caches for the `cell_count` / `cell_at` iteration
  // pair. Slot zero is the most recently used entry; keeping two entries
  // avoids rebuilding when callers alternate between two sheets. The
  // associated Sheet revision invalidates an entry after any cell, spill,
  // or structural mutation.
  struct CellEnumerationCache {
    std::size_t sheet_index = std::numeric_limits<std::size_t>::max();
    std::uint64_t revision = std::numeric_limits<std::uint64_t>::max();
    std::vector<std::pair<std::uint32_t, std::uint32_t>> addresses;
  };
  mutable std::array<CellEnumerationCache, 2> cell_enumeration_cache{};

  void clear_cell_enumeration_cache() noexcept { cell_enumeration_cache = {}; }

  // Storage behind every borrowed field `fm_sheet_get_cf_rule` fills. They are
  // separate from `read_scratch` on purpose: only a successful
  // `fm_sheet_get_cf_rule` clears and refreshes them, so neither an unrelated
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
  formulon::c_api::BorrowedArrayArena<fm_cell_rect_t> cf_range_scratch;
  formulon::c_api::BorrowedArrayArena<fm_cfvo_t> cfvo_scratch;
  formulon::c_api::BorrowedArrayArena<fm_cf_color_t> cf_color_scratch;
  formulon::c_api::BorrowedArrayArena<fm_cf_icon_t> cf_icon_scratch;

  // Arrays behind the borrowed `const T*` fields of the typed AutoFilter,
  // threaded-comment and custom table-style getters. They carry
  // `read_scratch`'s lifetime: each producer clears them together with
  // `read_scratch` before refilling.
  formulon::c_api::BorrowedArrayArena<fm_filter_column_t> filter_column_scratch;
  formulon::c_api::BorrowedArrayArena<const char*> filter_value_scratch;
  formulon::c_api::BorrowedArrayArena<fm_date_group_item_t> date_group_scratch;
  formulon::c_api::BorrowedArrayArena<fm_sort_condition_t> sort_condition_scratch;
  formulon::c_api::BorrowedArrayArena<fm_mention_t> mention_scratch;
  formulon::c_api::BorrowedArrayArena<fm_table_style_element_t> table_style_element_scratch;
  // Stops behind the `gradient_stops` view of `fm_workbook_get_fill` /
  // `fm_workbook_get_dxf`; refreshed by each of those calls.
  formulon::c_api::BorrowedArrayArena<fm_gradient_stop_t> gradient_stop_scratch;
  // Rectangles behind the `const fm_cell_rect_t*` outputs of the print-area,
  // protected-range and table getters (`publish_rects`); each call replaces
  // the previous array.
  formulon::c_api::BorrowedArrayArena<fm_cell_rect_t> rect_scratch;

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
