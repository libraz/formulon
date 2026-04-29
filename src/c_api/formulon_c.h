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
