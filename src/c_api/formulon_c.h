/*
 * Copyright 2026 libraz. Licensed under the MIT License.
 *
 * Formulon stable C ABI.
 *
 * This is a pure C11 header: every declaration sits inside `extern "C"`
 * when included from C++, no `bool` / `nullptr` / templates leak across
 * the boundary, and every type is either an opaque struct, a primitive
 * integer, a `const char*` (NUL-terminated UTF-8), or a small fixed
 * `fm_value_t` POD.
 *
 * The C ABI is the single contract between `formulon_core` and every
 * out-of-process consumer:
 *
 *   * the `formulon_cli` native binary,
 *   * the WASM build (Emscripten embind / direct cwrap users),
 *   * the Python wheel (`formulon` on PyPI, via ctypes),
 *   * any third-party language binding.
 *
 * Status codes mirror `formulon::FormulonErrorCode` (see
 * `src/utils/error.h`); a return value of `0` (`kOk`) always denotes
 * success. Error strings and contexts associated with the most recent
 * call are exposed via `fm_last_error_message()` /
 * `fm_last_error_context()`. Both buffers are thread-local and are
 * overwritten on every API call on the same thread; copy them out
 * before invoking another API entry point if the caller wants to keep
 * the diagnostic.
 *
 * Threading model:
 *   * Each `fm_workbook_t*` is owned by exactly one thread at a time.
 *     The library does not introduce internal locks; concurrent calls
 *     on the same handle from multiple threads are undefined behaviour.
 *   * `fm_last_error_*` storage is thread-local, so two threads driving
 *     two distinct workbooks can read their own diagnostics
 *     independently.
 *   * `fm_status_string()` and `fm_version_string()` are pure functions
 *     and are safe to call concurrently.
 *
 * Memory model:
 *   * Every `out_*` pointer borrowed from a workbook handle is valid
 *     until the next mutation on that handle, or until the handle is
 *     destroyed.
 *   * Buffers returned through `fm_workbook_save` are heap-allocated
 *     by the library with `new[]` and must be released by passing the
 *     same pointer to `fm_buffer_free`. Mixing `malloc`/`free` and
 *     `new[]`/`delete[]` across the boundary is undefined.
 */

#ifndef FORMULON_C_API_H
#define FORMULON_C_API_H

#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

/*
 * Public symbol visibility.
 *
 * The core library is built with `-fvisibility=hidden`, so every
 * exported entry point must opt back in via `__attribute__((visibility
 * ("default")))` (or `__declspec(dllexport)` on Windows). Consumers
 * that link against a static `formulon_static` archive can ignore the
 * macro; consumers that link against a shared library rely on it for
 * symbol export filtering.
 */
#if defined(_WIN32)
#define FM_API __declspec(dllexport)
#else
#define FM_API __attribute__((visibility("default")))
#endif

/**
 * @brief Status / error code returned by every fallible C API entry.
 *
 * Numerically identical to `formulon::FormulonErrorCode`. `0` is
 * `kOk` (success); every other value identifies a module band as
 * documented in `src/utils/error.h`.
 */
typedef int32_t fm_status_t;

/** @brief Opaque workbook handle. */
typedef struct fm_workbook fm_workbook_t;

/**
 * @brief Discriminator tag for `fm_value_t::kind`.
 *
 * Numbering matches `formulon::ValueKind` so bindings can cast
 * directly without a translation table.
 */
typedef enum {
  FM_VAL_BLANK = 0,
  FM_VAL_NUMBER = 1,
  FM_VAL_BOOL = 2,
  FM_VAL_TEXT = 3,
  FM_VAL_ERROR = 4,
  FM_VAL_ARRAY = 5,
  FM_VAL_REF = 6,
  FM_VAL_LAMBDA = 7
} fm_value_kind_t;

/**
 * @brief Cell value POD shared across the C ABI boundary.
 *
 * The active union member is selected by `kind`:
 *   * `FM_VAL_BLANK`   — no payload (other fields undefined).
 *   * `FM_VAL_NUMBER`  — `u.number` holds the IEEE-754 double.
 *   * `FM_VAL_BOOL`    — `u.boolean` is `0` (FALSE) or `1` (TRUE).
 *   * `FM_VAL_ERROR`   — `u.error_code` is an `Excel ErrorCode`
 *                        ordinal (Null=0, Div0=1, Value=2, ...). The
 *                        full enumeration is `formulon::ErrorCode`
 *                        in `src/value.h`.
 *   * `FM_VAL_TEXT`    — `u.text` is a NUL-terminated UTF-8 view
 *                        owned by the originating workbook handle.
 *                        The view is valid until the next mutation of
 *                        the handle or until the handle is destroyed.
 *   * `FM_VAL_ARRAY`   — array passthrough is reserved for a later
 *                        bundle. The current ABI surfaces the kind
 *                        but does not expose the payload.
 *   * `FM_VAL_REF`     — references are reserved.
 *   * `FM_VAL_LAMBDA`  — lambdas are reserved.
 */
typedef struct {
  fm_value_kind_t kind;
  union {
    double number;
    int32_t boolean;
    int32_t error_code;
    const char* text;
  } u;
} fm_value_t;

/* -------------------------------------------------------------------------- */
/* Construction / lifecycle                                                   */
/* -------------------------------------------------------------------------- */

/**
 * @brief Builds a default workbook (a single sheet named `"Sheet1"`).
 *
 * @param out  On success receives a freshly allocated handle. The
 *             caller owns the handle and must release it with
 *             `fm_workbook_destroy`.
 * @return `kOk` on success; `kBindingNullPointer` if `out == NULL`.
 *
 * @thread-safety Not thread-safe with respect to its own `out`
 *                pointer. Two threads must not race to populate the
 *                same `out` slot.
 */
FM_API fm_status_t fm_workbook_create(fm_workbook_t** out);

/**
 * @brief Builds an empty workbook (no sheets).
 *
 * Designed for callers that intend to add sheets immediately via
 * `fm_workbook_add_sheet`. Saving an empty workbook fails just as
 * Excel rejects sheet-less archives.
 *
 * @param out  On success receives a freshly allocated handle.
 * @return `kOk` on success; `kBindingNullPointer` if `out == NULL`.
 */
FM_API fm_status_t fm_workbook_create_empty(fm_workbook_t** out);

/**
 * @brief Loads a workbook from an in-memory `.xlsx` (OOXML) buffer.
 *
 * The handle takes ownership of the parsed read result, including
 * the text storage that backs every `Value::text` view inside the
 * workbook.
 *
 * @param bytes  Pointer to the first byte of the OOXML archive.
 * @param len    Length of the buffer in bytes.
 * @param out    On success receives the freshly allocated handle.
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL` or
 *         `len == 0`;
 *         a `kIo*` / `kCrypto*` code from
 *         `formulon::FormulonErrorCode` on parse / archive failure.
 */
FM_API fm_status_t fm_workbook_load(const uint8_t* bytes, size_t len, fm_workbook_t** out);

/**
 * @brief Releases a workbook handle.
 *
 * `wb == NULL` is a no-op (mirrors `free()` semantics). Every text
 * pointer previously surfaced through `fm_workbook_get_value` /
 * `fm_workbook_sheet_name` is invalidated.
 */
FM_API void fm_workbook_destroy(fm_workbook_t* wb);

/* -------------------------------------------------------------------------- */
/* Save                                                                       */
/* -------------------------------------------------------------------------- */

/**
 * @brief Serialises the workbook to an in-memory `.xlsx` byte stream.
 *
 * On success the caller receives a heap-allocated buffer that MUST be
 * released with `fm_buffer_free` (NOT `free`). Mixing allocators
 * across the boundary is undefined.
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param out_bytes  Receives a pointer to the freshly allocated buffer.
 * @param out_len    Receives the buffer length in bytes.
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         a `kIo*` code on archive / writer failure.
 */
FM_API fm_status_t fm_workbook_save(const fm_workbook_t* wb, uint8_t** out_bytes, size_t* out_len);

/**
 * @brief Releases a buffer returned by `fm_workbook_save`.
 *
 * `bytes == NULL` is a no-op. Internally pairs with `new uint8_t[]`;
 * callers must not pass pointers obtained from `malloc`.
 */
FM_API void fm_buffer_free(uint8_t* bytes);

/* -------------------------------------------------------------------------- */
/* Sheets                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the number of sheets in the workbook.
 *
 * Returns `0` when `wb == NULL`; this is unambiguous because
 * Workbooks never expose zero sheets through `fm_workbook_create`
 * (always at least `"Sheet1"`), and an `fm_workbook_create_empty`
 * caller is expected to consult the return value via the construction
 * status, not by polling sheet count.
 */
FM_API size_t fm_workbook_sheet_count(const fm_workbook_t* wb);

/**
 * @brief Returns the display name of the sheet at `index` (UTF-8).
 *
 * The returned pointer is owned by the workbook and is valid until
 * the next mutation that touches the sheet list (for example
 * `fm_workbook_add_sheet`) or until the handle is destroyed.
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param index      0-based sheet index. Must be `< sheet_count`.
 * @param out_utf8   Receives a NUL-terminated UTF-8 pointer.
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_sheet_name(const fm_workbook_t* wb, size_t index, const char** out_utf8);

/**
 * @brief Appends a new sheet with the given UTF-8 display name.
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param utf8_name  Display name; must be non-NULL.
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_add_sheet(fm_workbook_t* wb, const char* utf8_name);

/**
 * @brief Moves the sheet at `from_index` to `to_index`.
 *
 * `to_index` is interpreted in the *post-removal* sheet list (Excel UI
 * semantics): with three sheets, moving sheet 0 to the end uses
 * `to_index == 2`, not `3`. A move to the same position is a successful
 * no-op.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kSheetIndexOutOfRange` when either index is out of range.
 */
FM_API fm_status_t fm_workbook_move_sheet(fm_workbook_t* wb, uint32_t from_index, uint32_t to_index);

/**
 * @brief Removes the sheet at `index`.
 *
 * Defined names that reference the removed sheet are dropped (matches
 * the simpler semantics chosen for this revision; a future bundle may
 * surface `#REF!` instead). The recalc engine's dep-graph entries for
 * cells on the removed sheet are unregistered.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kSheetIndexOutOfRange` when `index` is out of range;
 *         `kCannotRemoveLastSheet` when the workbook has only one
 *         sheet remaining.
 */
FM_API fm_status_t fm_workbook_remove_sheet(fm_workbook_t* wb, uint32_t index);

/**
 * @brief Renames the sheet at `index` to `new_name` (UTF-8).
 *
 * Updates the sheet's stored name and any workbook-scoped defined-name
 * targets that mention the renamed sheet. Cell formulas inside the
 * renamed sheet (and other sheets) are LEFT UNTOUCHED — the AST-level
 * reference shifter handles those in a follow-up bundle.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kSheetIndexOutOfRange` when `index` is out of range;
 *         `kInvalidSheetName` when `new_name` is empty, longer than 31
 *         characters, contains a forbidden character (`: \ / ? * [ ]`),
 *         or collides case-insensitively with an existing sheet.
 */
FM_API fm_status_t fm_workbook_rename_sheet(fm_workbook_t* wb, uint32_t index, const char* new_name);

/**
 * @brief Sets (or appends, or removes) a workbook-scoped defined name.
 *
 * If a workbook-scoped entry with `name` already exists (case-
 * insensitive match), its formula text is replaced. Otherwise — when
 * `formula` is non-empty — a new entry is appended. Passing an empty
 * `formula` removes the existing entry, or is a no-op when no such
 * entry is present. Sheet-scoped defined names (`local_sheet_id >= 0`)
 * are NOT addressable through this entry point.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `name` is empty.
 */
FM_API fm_status_t fm_workbook_set_defined_name(fm_workbook_t* wb, const char* name, const char* formula);

/* -------------------------------------------------------------------------- */
/* Row / column structural edits                                              */
/* -------------------------------------------------------------------------- */

/**
 * @brief Inserts `count` rows at `row` on `sheet`. Cells at `row` and
 *        beyond shift forward; cells pushed past the sheet bound are
 *        dropped. References across the workbook are rewritten to
 *        track the new positions; references that would land past the
 *        sheet bound collapse to `#REF!`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range, `row` is
 *         past the sheet bound, or `count == 0`.
 */
FM_API fm_status_t fm_workbook_insert_rows(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t count);

/**
 * @brief Deletes `count` rows starting at `row` on `sheet`. The deleted
 *        rows are dropped wholesale; subsequent rows shift back.
 *        References pointing inside the deleted interval collapse to
 *        `#REF!`; references past the deletion shift back.
 *
 * @return Same status codes as `fm_workbook_insert_rows`.
 */
FM_API fm_status_t fm_workbook_delete_rows(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t count);

/**
 * @brief Inserts `count` columns at `col` on `sheet`. Mirrors
 *        `fm_workbook_insert_rows` along the column axis.
 */
FM_API fm_status_t fm_workbook_insert_cols(fm_workbook_t* wb, uint32_t sheet, uint32_t col, uint32_t count);

/**
 * @brief Deletes `count` columns starting at `col` on `sheet`. Mirrors
 *        `fm_workbook_delete_rows` along the column axis.
 */
FM_API fm_status_t fm_workbook_delete_cols(fm_workbook_t* wb, uint32_t sheet, uint32_t col, uint32_t count);

/* -------------------------------------------------------------------------- */
/* Cell mutation                                                              */
/* -------------------------------------------------------------------------- */

/**
 * @brief Stores a numeric literal at `(row, col)` on the given sheet.
 *
 * Mirrors `Workbook::set_cell_value(... Value::number)`. Marks the
 * cell and any existing dependents dirty for the next recalc.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_set_number(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                          double value);

/**
 * @brief Stores a boolean literal. `value` is interpreted as a C
 *        boolean: `0` is FALSE, any non-zero is TRUE.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_set_bool(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        int32_t value);

/**
 * @brief Stores a text literal. The handle copies the UTF-8 contents
 *        into its internal text storage; `utf8` does not need to
 *        outlive the call.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_set_text(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        const char* utf8);

/**
 * @brief Stores a `Blank` literal at `(row, col)`. Equivalent to
 *        clearing a cell.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_set_blank(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col);

/**
 * @brief Stores a formula at `(row, col)` and registers its
 *        dependencies.
 *
 * `formula` is the raw Excel-style formula text (with or without a
 * leading `=`); the parser controls the contract.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_set_formula(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                           const char* formula);

/* -------------------------------------------------------------------------- */
/* Cell read                                                                  */
/* -------------------------------------------------------------------------- */

/**
 * @brief Reads the cached cell value at `(row, col)`.
 *
 * The cell value reflects the most recent recalc. Callers that just
 * mutated cells should invoke `fm_workbook_recalc` first.
 *
 * For `FM_VAL_TEXT`, `out->u.text` borrows a NUL-terminated UTF-8
 * pointer owned by the workbook. The pointer is valid until the next
 * mutation of the handle or until the handle is destroyed. Callers
 * that need to retain the string across mutations must copy it.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_value(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                         fm_value_t* out);

/* -------------------------------------------------------------------------- */
/* Iteration / dump                                                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the number of stored cell slots on `sheet_index`.
 *
 * Counts every populated `Cell` (literal, formula, or implicitly created
 * during row growth) on the sheet's row-sparse / column-dense storage.
 * Phantom cells of a spill region that have no underlying stored slot
 * are NOT counted.
 *
 * The count is suitable as the upper bound for `fm_workbook_cell_at`
 * iteration: indices in `[0, count)` enumerate every stored cell. The
 * iteration order is implementation-defined but stable for a given
 * workbook state (no cell mutation has occurred between calls).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL` or `out_count == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_cell_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `idx`-th cell on `sheet_index`.
 *
 * Iteration order is implementation-defined but stable for a given
 * workbook state (sorted by `(row, col)` ascending).
 *
 * On success the call writes `*out_row`, `*out_col`, `*out_value`, and
 * — if `out_formula != NULL` — points `*out_formula` at the cell's raw
 * formula text (or `NULL` when the cell is a pure literal). The
 * formula pointer borrows from the workbook handle and is valid until
 * the next mutation that touches the sheet's cell store or until the
 * handle is destroyed.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb`, `out_row`, `out_col`, or
 *         `out_value` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range or
 *         `idx >= cell_count`.
 */
FM_API fm_status_t fm_workbook_cell_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, uint32_t* out_row,
                                       uint32_t* out_col, const char** out_formula, fm_value_t* out_value);

/**
 * @brief Returns the number of defined names attached to the workbook.
 *
 * Defined names round-trip through OOXML even when not yet evaluated.
 * Counts every `<definedName>` entry in declaration order.
 *
 * @return `0` when `wb == NULL`; otherwise the size of the workbook's
 *         defined-name list.
 */
FM_API size_t fm_workbook_defined_name_count(const fm_workbook_t* wb);

/**
 * @brief Reads the `idx`-th defined name.
 *
 * On success `*out_name` and `*out_formula` borrow NUL-terminated UTF-8
 * pointers from the workbook handle. Both are valid until the next
 * mutation of the defined-name list or until the handle is destroyed.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_defined_name_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                               const char** out_formula);

/**
 * @brief Returns the number of tables attached to the workbook.
 *
 * @return `0` when `wb == NULL`; otherwise the size of the workbook's
 *         table-metadata list.
 */
FM_API size_t fm_workbook_table_count(const fm_workbook_t* wb);

/**
 * @brief Reads the `idx`-th table's identifying metadata.
 *
 * On success `*out_name`, `*out_display_name`, and `*out_ref` borrow
 * NUL-terminated UTF-8 pointers from the workbook handle. Same lifetime
 * contract as `fm_workbook_defined_name_at`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_table_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                        const char** out_display_name, const char** out_ref, size_t* out_sheet_index);

/**
 * @brief Returns the number of passthrough OOXML parts the reader
 *        carried through unchanged.
 *
 * @return `0` when `wb == NULL`; otherwise the size of the workbook's
 *         passthrough-part list.
 */
FM_API size_t fm_workbook_passthrough_count(const fm_workbook_t* wb);

/**
 * @brief Reads the `idx`-th passthrough part's path.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_passthrough_at(const fm_workbook_t* wb, size_t idx, const char** out_path);

/* -------------------------------------------------------------------------- */
/* Sheet UI features (merges, hyperlinks, comments)                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Merge range descriptor: 0-based inclusive `(first, last)` corners.
 */
typedef struct {
  uint32_t first_row;
  uint32_t first_col;
  uint32_t last_row;
  uint32_t last_col;
} fm_merge_range;

/**
 * @brief Hyperlink descriptor. String fields are NUL-terminated UTF-8
 *        pointers borrowed from the workbook handle; they are valid
 *        until the next mutation that touches the sheet's hyperlink
 *        list or until the handle is destroyed. Newly-added entries
 *        accept NULL for any of `target`, `display`, `tooltip`,
 *        `location` to mean "empty".
 */
typedef struct {
  uint32_t row;
  uint32_t col;
  const char* target;
  const char* location;
  const char* display;
  const char* tooltip;
} fm_hyperlink;

/**
 * @brief Comment descriptor. `author` and `text` are borrowed UTF-8
 *        pointers with the same lifetime contract as `fm_hyperlink`.
 */
typedef struct {
  uint32_t row;
  uint32_t col;
  const char* author;
  const char* text;
} fm_comment;

/**
 * @brief Appends a merge range to the sheet's merge list.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 */
FM_API fm_status_t fm_sheet_add_hyperlink(fm_workbook_t* wb, uint32_t sheet, fm_hyperlink hl);

/**
 * @brief Appends a merge range to the sheet's merge list.
 */
FM_API fm_status_t fm_sheet_add_merge(fm_workbook_t* wb, uint32_t sheet, fm_merge_range merge);

/**
 * @brief Removes every merge range on `sheet` that overlaps `range`
 *        (inclusive rectangle intersection). No-op when nothing
 *        overlaps. `range` corners are normalised internally so the
 *        caller may pass either (first <= last) or transposed corners,
 *        mirroring fm_sheet_add_merge.
 *
 * @return `kOk` on success (including the no-overlap case);
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_merge(fm_workbook_t* wb, uint32_t sheet, fm_merge_range range);

/**
 * @brief Removes the merge at `index` on `sheet`. Mirrors the
 *        index domain of fm_sheet_get_merge_at.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` or `index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_merge_at(fm_workbook_t* wb, uint32_t sheet, uint32_t index);

/**
 * @brief Drops every merge range on `sheet`. No-op when the sheet
 *        already has no merges.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_merges(fm_workbook_t* wb, uint32_t sheet);

/**
 * @brief Reads `(out_row, out_col, out_author, out_text)` for the
 *        comment at `(row, col)` on `sheet`. Returns `kInvalidArgument`
 *        when no comment is anchored there.
 *
 * `out->author` and `out->text` borrow NUL-terminated UTF-8 pointers
 * from the workbook handle. Both are valid until the next mutation
 * that touches the sheet's comment list or until the handle is
 * destroyed.
 */
FM_API fm_status_t fm_sheet_get_comment_at(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                           fm_comment* out);

/**
 * @brief Reads the `index`-th hyperlink on `sheet` into `out`.
 *
 * String fields in `out` are borrowed pointers (see `fm_hyperlink`).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet` or `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_hyperlink_at(fm_workbook_t* wb, uint32_t sheet, uint32_t index, fm_hyperlink* out);

/**
 * @brief Returns the number of hyperlinks attached to `sheet`.
 */
FM_API fm_status_t fm_sheet_get_hyperlink_count(fm_workbook_t* wb, uint32_t sheet, uint32_t* out_count);

/**
 * @brief Reads the `index`-th merge range on `sheet`.
 */
FM_API fm_status_t fm_sheet_get_merge_at(fm_workbook_t* wb, uint32_t sheet, uint32_t index, fm_merge_range* out);

/**
 * @brief Returns the number of merge ranges attached to `sheet`.
 */
FM_API fm_status_t fm_sheet_get_merge_count(fm_workbook_t* wb, uint32_t sheet, uint32_t* out_count);

/**
 * @brief Inserts or replaces the comment at `(row, col)` on `sheet`.
 *        Pass `text == NULL || text[0] == '\0'` to remove an existing
 *        comment. The strings are copied into the workbook handle's
 *        text storage.
 */
FM_API fm_status_t fm_sheet_set_comment(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                        const char* author, const char* text);

/* -------------------------------------------------------------------------- */
/* Recalc                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Drives a full incremental recalc using the default function
 *        registry (`eval::default_registry()`).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         a `kEval*` / `kGraph*` code on engine failure.
 */
FM_API fm_status_t fm_workbook_recalc(fm_workbook_t* wb);

/**
 * @brief Configures iterative-calculation knobs for the workbook's
 *        recalc engine.
 *
 * @param enabled         `0` disables iterative calc (the default,
 *                        cyclic SCCs surface `#REF!`); non-zero
 *                        enables fixed-point resolution up to
 *                        `max_iterations` passes.
 * @param max_iterations  Iteration cap. Values < 1 are treated as 1.
 * @param max_change      Absolute convergence threshold.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`.
 */
FM_API fm_status_t fm_workbook_set_iterative(fm_workbook_t* wb, int32_t enabled, int32_t max_iterations,
                                             double max_change);

/**
 * @brief Workbook-relative viewport rectangle, expressed in 0-based
 *        inclusive coordinates. Used by `fm_workbook_partial_recalc`.
 *
 * `last_row` / `last_col` are inclusive — a single-cell viewport sets
 * the corresponding `first_*` and `last_*` to the same value. An empty
 * viewport (the row or column range collapsed) is allowed and produces
 * a no-op recalc.
 */
typedef struct {
  uint32_t sheet;     /**< 0-based sheet index. */
  uint32_t first_row; /**< First row, 0-based, inclusive. */
  uint32_t last_row;  /**< Last row, 0-based, inclusive. */
  uint32_t first_col; /**< First column, 0-based, inclusive. */
  uint32_t last_col;  /**< Last column, 0-based, inclusive. */
} fm_viewport;

/**
 * @brief Iterative-solver progress callback signature.
 *
 * Invoked once per Gauss-Seidel sweep with the 1-based iteration index,
 * the maximum residual observed during the sweep, and the configured
 * iteration cap. `user_data` is the opaque pointer the caller registered
 * via `fm_workbook_set_iterative_progress`.
 *
 * Return `1` (true) to continue iterating, `0` (false) to abort early.
 * An aborted solve leaves the workbook in its current
 * partially-converged state.
 */
typedef bool (*fm_iterative_progress_cb)(uint32_t iteration, double max_residual, uint32_t max_iterations,
                                         void* user_data);

/**
 * @brief Recalculates the dependency closure required to produce
 *        correct values for the supplied viewport rectangle.
 *
 * Cells outside the closure remain dirty; a subsequent
 * `fm_workbook_recalc` (or a `fm_workbook_partial_recalc` whose closure
 * overlaps them) will visit them.
 *
 * @param wb                     Workbook handle. Must be non-NULL.
 * @param viewport               Viewport rectangle. Must be non-NULL.
 * @param out_recomputed_count   Optional. Receives the number of cells
 *                               actually recomputed during this call.
 *                               May be NULL if the caller does not
 *                               care.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `viewport` is NULL;
 *         a `kEval*` / `kGraph*` code on engine failure.
 */
FM_API fm_status_t fm_workbook_partial_recalc(fm_workbook_t* wb, const fm_viewport* viewport,
                                              uint32_t* out_recomputed_count);

/**
 * @brief Sets the iterative-solver progress callback for subsequent
 *        recalcs on this workbook.
 *
 * Pass `cb == NULL` to clear the callback. `user_data` is forwarded
 * verbatim to every invocation; the workbook does not take ownership
 * of it.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`.
 */
FM_API fm_status_t fm_workbook_set_iterative_progress(fm_workbook_t* wb, fm_iterative_progress_cb cb, void* user_data);

/* -------------------------------------------------------------------------- */
/* Conditional formatting                                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Conditional-format match-kind ordinal.
 *
 * Mirrors `formulon::cf::CFMatchKind` (`src/cf/cf_match.h`). The active
 * payload fields on `fm_cf_match_t` are determined by this enumerator.
 */
/* NOLINTNEXTLINE(performance-enum-size): C11 enums cannot specify an
 * underlying type; matching `fm_value_kind_t`'s shape keeps the C ABI
 * uniform across enums. */
typedef enum {
  FM_CF_DIFFERENTIAL_FORMAT = 0,
  FM_CF_COLOR_SCALE = 1,
  FM_CF_DATA_BAR = 2,
  FM_CF_ICON_SET = 3
} fm_cf_match_kind_t;

/**
 * @brief Plain RGBA colour. Channels are 0-255 (sRGB).
 */
typedef struct {
  uint8_t r;
  uint8_t g;
  uint8_t b;
  uint8_t a;
} fm_cf_color_t;

/**
 * @brief Resolved CF match for one rule on one cell.
 *
 * The active fields depend on `kind`:
 *   * `FM_CF_DIFFERENTIAL_FORMAT` — `dxf_id_engaged` is `1` when the
 *     rule carries a dxf reference, in which case `dxf_id` indexes
 *     `styles.dxfs[]`.
 *   * `FM_CF_COLOR_SCALE` — `color` is the interpolated cell-fill RGBA.
 *   * `FM_CF_DATA_BAR` — `bar_length_pct`, `bar_axis_position_pct`,
 *     `bar_is_negative`, `bar_fill`, `bar_border_engaged` (and
 *     `bar_border` when engaged), and `bar_gradient`.
 *   * `FM_CF_ICON_SET` — `icon_set_name` is the
 *     `formulon::cf::IconSetName` ordinal; `icon_index` is `0..N-1`
 *     after the rule's `reverse` flag has been applied.
 *
 * All non-active fields carry default-zero values. `priority` is the
 * workbook-global priority lifted from `CFRule::priority` (smaller
 * numbers ranked higher); per-cell match lists are returned in
 * priority-ascending order so a UI can fold matches left-to-right.
 */
typedef struct {
  fm_cf_match_kind_t kind;
  int32_t priority;
  int32_t dxf_id_engaged; /* 0/1 */
  uint32_t dxf_id;
  fm_cf_color_t color; /* ColorScale fill */
  double bar_length_pct;
  double bar_axis_position_pct;
  int32_t bar_is_negative; /* 0/1 */
  fm_cf_color_t bar_fill;
  int32_t bar_border_engaged; /* 0/1 */
  fm_cf_color_t bar_border;
  int32_t bar_gradient;  /* 0/1 */
  int32_t icon_set_name; /* formulon::cf::IconSetName ordinal */
  uint8_t icon_index;
  uint8_t _pad[3]; /* padding for alignment determinism */
} fm_cf_match_t;

/**
 * @brief Opaque handle for a CF range evaluation result.
 *
 * Owns the per-cell match lists produced by `fm_workbook_cf_evaluate_range`.
 * Index into it via `fm_cf_results_cell_count` /
 * `fm_cf_results_cell_at` / `fm_cf_results_match_at`. Release with
 * `fm_cf_results_destroy`.
 */
typedef struct fm_cf_results fm_cf_results_t;

/**
 * @brief Evaluates every CF block on `sheet_index` against the cells in
 *        the inclusive range `[(first_row, first_col), (last_row,
 *        last_col)]`.
 *
 * The result is sparse: only cells that produced at least one match
 * appear. Internally constructs an `Arena`, the default function
 * registry, and an `EvalContext` bound to the requested sheet.
 *
 * `today_serial` pins the date reference for `TimePeriod` rules. Pass
 * `NaN` to disable (in which case `TimePeriod` rules will not match).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb` or `out` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 *
 * On success the caller owns `*out` and must free it with
 * `fm_cf_results_destroy`.
 */
FM_API fm_status_t fm_workbook_cf_evaluate_range(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                                 uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                 double today_serial, fm_cf_results_t** out);

/**
 * @brief Releases a results handle. `results == NULL` is a no-op.
 */
FM_API void fm_cf_results_destroy(fm_cf_results_t* results);

/**
 * @brief Returns the number of cells in the result that produced at
 *        least one match. Returns `0` when `results == NULL`.
 */
FM_API size_t fm_cf_results_cell_count(const fm_cf_results_t* results);

/**
 * @brief Reads cell `cell_idx`'s coordinates and match count.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when any pointer argument is `NULL`;
 *         `kInvalidArgument` when `cell_idx >= cell_count`.
 */
FM_API fm_status_t fm_cf_results_cell_at(const fm_cf_results_t* results, size_t cell_idx, uint32_t* out_row,
                                         uint32_t* out_col, size_t* out_match_count);

/**
 * @brief Reads match `match_idx` for cell `cell_idx`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when any pointer argument is `NULL`;
 *         `kInvalidArgument` when either index is out of range.
 */
FM_API fm_status_t fm_cf_results_match_at(const fm_cf_results_t* results, size_t cell_idx, size_t match_idx,
                                          fm_cf_match_t* out);

/* -------------------------------------------------------------------------- */
/* Sheet view / layout (viewport, frozen panes, column / row overrides)       */
/* -------------------------------------------------------------------------- */

/**
 * @brief Per-sheet view state surfaced over the C ABI.
 *
 * Mirrors `formulon::SheetView`. Fields that are at their default value
 * still surface (zoom_scale defaults to `100`, freeze_rows /
 * freeze_cols default to `0`, tab_hidden defaults to `0`).
 */
typedef struct {
  uint32_t zoom_scale;
  uint32_t freeze_rows;
  uint32_t freeze_cols;
  int32_t tab_hidden; /* 0/1 */
} fm_sheet_view_t;

/**
 * @brief Per-column layout override surfaced over the C ABI.
 *
 * Mirrors `formulon::ColumnLayout`. Both endpoints are 0-based and
 * inclusive (matches the engine-side type; OOXML's `min`/`max`
 * 1-based conversion stays inside the writer).
 */
typedef struct {
  uint32_t first;
  uint32_t last;
  double width;
  int32_t hidden; /* 0/1 */
  uint8_t outline_level;
  uint8_t _pad[3];
} fm_column_layout_t;

/**
 * @brief Per-row layout override surfaced over the C ABI.
 *
 * Mirrors `formulon::RowLayout`. The row index is 0-based; height is
 * in points (matches OOXML `<row ht=...>`).
 */
typedef struct {
  uint32_t row;
  double height;
  int32_t hidden; /* 0/1 */
  uint8_t outline_level;
  uint8_t _pad[3];
} fm_row_layout_t;

/**
 * @brief Returns the column-layout-override count for `sheet_index`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_column_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `idx`-th column-layout override.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` or `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_column(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                       fm_column_layout_t* out);

/**
 * @brief Returns the row-layout-override count for `sheet_index`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_override_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `idx`-th row-layout override.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` or `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_override(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                             fm_row_layout_t* out);

/**
 * @brief Reads the sheet-view state (zoom, frozen panes, tab hidden).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_t* out);

/**
 * @brief Sets or replaces the column width override for the inclusive
 *        column span `[first, last]`. Existing overrides whose span
 *        intersects the requested span are merged so the new width
 *        takes precedence on overlapping columns; non-overlapping
 *        portions of pre-existing entries are retained.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range or
 *         `last < first`.
 */
FM_API fm_status_t fm_sheet_set_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                             double width);

/**
 * @brief Sets or replaces the column hidden flag for the inclusive
 *        column span `[first, last]`. See `fm_sheet_set_column_width`
 *        for the merge semantics.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range or
 *         `last < first`.
 */
FM_API fm_status_t fm_sheet_set_column_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                              int32_t hidden);

/**
 * @brief Sets or replaces the column outline level for the inclusive
 *        column span `[first, last]`. See `fm_sheet_set_column_width`
 *        for the merge semantics.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range or
 *         `last < first`.
 */
FM_API fm_status_t fm_sheet_set_column_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                               uint8_t level);

/**
 * @brief Sets or replaces the row height override for `row`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height);

/**
 * @brief Sets or replaces the row hidden flag for `row`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t hidden);

/**
 * @brief Sets or replaces the row outline level for `row`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_row_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint8_t level);

/**
 * @brief Sets the sheet's zoom percentage. Values outside `[10, 400]`
 *        are clamped to the nearest endpoint.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, size_t sheet_index, uint32_t zoom_scale);

/**
 * @brief Sets the sheet's freeze pane in `(rows, cols)`. Either or
 *        both may be `0` to remove that axis.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_freeze(fm_workbook_t* wb, size_t sheet_index, uint32_t freeze_rows,
                                       uint32_t freeze_cols);

/**
 * @brief Sets the sheet's tab-hidden flag. Non-zero hides the sheet
 *        tab; zero shows it.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_tab_hidden(fm_workbook_t* wb, size_t sheet_index, int32_t hidden);

/* -------------------------------------------------------------------------- */
/* Diagnostics                                                                */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the most recent error message produced by an API
 *        call on the current thread.
 *
 * The pointer is valid until the next API call on the same thread.
 * Returns an empty string (never `NULL`) when no error has been
 * recorded yet.
 */
FM_API const char* fm_last_error_message(void);

/**
 * @brief Returns the optional context string associated with the
 *        most recent error on the current thread.
 *
 * Same lifetime / non-NULL guarantees as `fm_last_error_message`.
 */
FM_API const char* fm_last_error_context(void);

/**
 * @brief Returns a static C string describing `status`.
 *
 * For known codes the result matches `formulon::to_cstring` (e.g.
 * `"kInvalidArgument"`); unknown codes yield `"kUnknownError"`.
 * Always non-NULL with program lifetime.
 */
FM_API const char* fm_status_string(fm_status_t status);

/* -------------------------------------------------------------------------- */
/* Styles                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Plain-data projection of a `formulon::io::CellXf` record.
 *
 * Mirrors the underlying record field-for-field. Bindings copy the
 * struct out of the workbook's styles table; subsequent mutations to
 * the workbook do not invalidate the copy.
 */
typedef struct {
  uint32_t font_index;
  uint32_t fill_index;
  uint32_t border_index;
  uint16_t num_fmt_id;
  uint8_t horizontal_align;
  uint8_t vertical_align;
  int32_t wrap_text; /* 0=false, 1=true */
} fm_cell_xf;

/**
 * @brief Plain-data projection of a `formulon::io::FontRecord`.
 *
 * `name` is a NUL-terminated UTF-8 pointer borrowed from the
 * workbook's styles table; it is valid until the next mutation that
 * replaces the styles table or until the handle is destroyed.
 */
typedef struct {
  const char* name; /* NUL-terminated UTF-8, borrowed */
  double size;
  uint32_t color_argb;
  int32_t bold;      /* 0=false, 1=true */
  int32_t italic;    /* 0=false, 1=true */
  int32_t strike;    /* 0=false, 1=true */
  uint8_t underline; /* 0=none, 1=single, 2=double, 3/4=accounting variants */
} fm_font_record;

/**
 * @brief Reads the `xf_index` (style record id) attached to the cell at
 *        `(row, col)` on `sheet`.
 *
 * Returns `0` (the default xf) when the cell is absent — callers that
 * need to distinguish "no cell" from "default-formatted cell" should
 * combine this with `fm_workbook_cell_at` iteration.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_cell_get_xf_index(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                        uint32_t* out_xf_index);

/**
 * @brief Stores `xf_index` on the cell at `(row, col)` on `sheet`.
 *
 * Materialises the cell as a default-blank slot when none exists yet,
 * mirroring the row-vector growth semantics of
 * `fm_workbook_set_blank`. Coexists with literal / formula writes:
 * those calls leave `xf_index` untouched, so the caller can layer a
 * style update on top of either order of operations.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_cell_set_xf_index(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                        uint32_t xf_index);

/**
 * @brief Reads the `xf_index`-th `<xf>` record from the workbook's
 *        styles table.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `xf_index >= cell_xfs.size()`.
 */
FM_API fm_status_t fm_styles_get_cell_xf(fm_workbook_t* wb, uint32_t xf_index, fm_cell_xf* out);

/**
 * @brief Reads the `font_index`-th font record from the workbook's
 *        styles table.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `font_index >= fonts.size()`.
 */
FM_API fm_status_t fm_styles_get_font(fm_workbook_t* wb, uint32_t font_index, fm_font_record* out);

/**
 * @brief Looks up the format string for `num_fmt_id` (built-in 0..163
 *        or custom >= 164).
 *
 * On success `*out` borrows a NUL-terminated UTF-8 pointer. Built-in
 * ids resolve to a static `.rodata` string with program lifetime;
 * custom ids resolve to storage owned by the workbook's styles table
 * (valid until the next styles-replacing mutation or until the handle
 * is destroyed). Returns `kInvalidArgument` when `num_fmt_id` is
 * neither a documented built-in nor a registered custom id.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when the id is unknown.
 */
FM_API fm_status_t fm_styles_get_num_fmt_string(fm_workbook_t* wb, uint16_t num_fmt_id, const char** out);

/* -------------------------------------------------------------------------- */
/* Version                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the Formulon library version as a NUL-terminated
 *        UTF-8 string. Always non-NULL with program lifetime.
 */
FM_API const char* fm_version_string(void);

#ifdef __cplusplus
} /* extern "C" */
#endif

#endif /* FORMULON_C_API_H */
