/*
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
 *   * the Python wheel (`formulon` on PyPI, via Wasmtime/WASM),
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
 *   * Each `fm_workbook_t*` is owned by exactly one caller thread at a time.
 *     Ordinary API calls do not synchronize concurrent calls on the same
 *     handle; callers must not race reads, mutations, or separate recalc
 *     calls on one handle. `fm_workbook_recalc_parallel` is the exception
 *     that may launch and join internal worker threads for the duration of
 *     one call. Those implementation threads do not make cross-call handle
 *     access safe, and the same-handle external-concurrency restriction
 *     remains in force.
 *   * `fm_last_error_*` storage is thread-local, so two threads driving
 *     two distinct workbooks can read their own diagnostics
 *     independently.
 *   * `fm_status_string()` and `fm_version_string()` are pure functions
 *     and are safe to call concurrently.
 *
 * Pointer-view vocabulary:
 *   * A **model-backed view** points into workbook-owned model storage. Its
 *     lifetime ends at the invalidating mutation named by that API, or when
 *     `fm_workbook_destroy` releases the handle. A read on the same handle
 *     does not invalidate a model-backed view.
 *   * A **read-scratch-backed view** points into transient per-handle read
 *     storage. The next successful scratch-backed read on the same handle
 *     may invalidate it; a validation-rejected call does not refresh this
 *     storage. Mutations and handle destruction also invalidate it. Callers
 *     must copy strings or arrays before making another call when they need
 *     to retain the value.
 *   * An **owned snapshot** is returned through a separately owned opaque
 *     result handle and stays valid until that result is destroyed. A
 *     **static view** points at process-static storage and has program
 *     lifetime. Each individual API below names which view it returns.
 *
 * Memory model:
 *   * Workbook-backed views are never allocated for the caller and must not
 *     be freed. Their exact invalidation boundary is part of each API's
 *     contract below; do not infer it from the C pointer type alone.
 *   * Buffers returned through `fm_workbook_save`, `fm_workbook_save_as`
 *     and `fm_workbook_save_with_diagnostics` are heap-allocated by the
 *     library with `new[]` and must be released by passing the same
 *     pointer to `fm_buffer_free`. Mixing `malloc`/`free` and
 *     `new[]`/`delete[]` across the boundary is undefined.
 *   * Every API that returns an owned opaque handle through an `out`
 *     parameter (including workbook, conditional-format-result, and
 *     cell-node handles) initializes `*out` to `NULL` before validation and
 *     leaves it `NULL` on every failure path.
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

/** @brief Inclusive, 0-based print-area rectangle. */
typedef struct {
  uint32_t first_row;
  uint32_t first_col;
  uint32_t last_row;
  uint32_t last_col;
} fm_print_range_t;

/** @brief Opaque result of one sheet pagination operation. */
typedef struct fm_pagination fm_pagination_t;

/**
 * @brief Cell error code payload. Mirrors `formulon::ErrorCode`.
 *
 * Used by cell and PivotCache error setters. The ordinal values are the
 * engine's stable in-memory enum values, not OOXML wire codes.
 */
typedef int32_t fm_error_code_t;

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
 *   * `FM_VAL_TEXT`    — `u.text` is a NUL-terminated UTF-8 view. Its
 *                        backing category is specified by the API that
 *                        writes the `fm_value_t`: cell/evaluation getters
 *                        use a read-scratch-backed view, while an owned
 *                        result handle may provide an owned snapshot.
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
 * The only difference from `fm_workbook_create` is the missing sheet:
 * the style table is seeded with Excel's reserved defaults either way
 * (font 0, fill 0 = `none`, fill 1 = `gray125`, border 0, cell-style xf
 * 0, cell xf 0, the `Normal` cell style). Style records added through
 * `fm_styles_add_*` therefore always land after the reserved slots, so
 * the first index handed back is non-zero.
 *
 * @param out  On success receives a freshly allocated handle.
 * @return `kOk` on success; `kBindingNullPointer` if `out == NULL`.
 */
FM_API fm_status_t fm_workbook_create_empty(fm_workbook_t** out);

/**
 * @brief Loads a workbook from an in-memory `.xlsx` or `.xlsb` buffer.
 *
 * The handle takes ownership of the parsed read result, including
 * the text storage that backs every `Value::text` view inside the
 * workbook.
 *
 * The container is auto-detected from its package contents; the caller does
 * not need to provide a filename or extension.
 *
 * @param bytes  Pointer to the first byte of the OOXML or MS-XLSB archive.
 * @param len    Length of the buffer in bytes.
 * @param out    On success receives the freshly allocated handle.
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL` or
 *         `len == 0`;
 *         a `kIo*` / `kCrypto*` code from
 *         `formulon::FormulonErrorCode` on parse / archive failure.
 */
FM_API fm_status_t fm_workbook_load(const uint8_t* bytes, size_t len, fm_workbook_t** out);

/*
 * Contract shared by `fm_read_diagnostics_t` and `fm_save_diagnostics_t`:
 *
 *   * A field means the same thing regardless of which container was
 *     loaded or saved. Where the underlying event exists in only one
 *     container, the field still counts it and the other container simply
 *     never raises it. A caller never has to know which reader or writer
 *     ran in order to read a number.
 *   * Every field lists the structured-log events that feed it, by event
 *     name. That log ships disabled, so these counters — not the log —
 *     are what a host actually sees.
 *   * Coverage is partial by design. These counters cover part-,
 *     relationship- and feature-level loss in the package readers and
 *     writers. Three engine events stay outside them, each needing its
 *     own semantics rather than a slot in an existing field: an
 *     unrepresentable cached cell value, unsupported worksheet format
 *     metadata, and a failed recalc worker launch. An all-zero result
 *     therefore means "none of the losses enumerated below occurred",
 *     not "nothing at all went wrong".
 *   * The counters are `uint32_t`, not `size_t`, so both structs have the
 *     same layout on native and wasm32 (20 bytes, 4-aligned) and a
 *     binding's hand-written layout cannot silently disagree with the
 *     native one.
 */

/**
 * @brief Loss and fallback counters captured by the load that created a
 *        workbook handle.
 */
typedef struct {
  /** Formula cells whose stored formula could not be decoded; the cached
   *  value survives, the formula does not. XLSB only.
   *  Fed by: `xlsb.formula.not_decoded`. */
  uint32_t undecoded_formula_count;
  /** Defined names skipped for the same reason. XLSB only.
   *  Fed by: `xlsb.defined_name.not_decoded`. */
  uint32_t undecoded_defined_name_count;
  /** Package parts whose content type could not be resolved, so they were
   *  not decoded and are absent from the loaded model. XLSB only: the
   *  OOXML reader carries every unmodelled part through passthrough
   *  instead of dropping it.
   *  Fed by: `xlsb.package.parts_dropped`. */
  uint32_t undecoded_part_count;
  /** Presentation-overlay entries dropped because their reference was
   *  missing or unparseable. One per dropped `<mergeCells>` /
   *  `<hyperlinks>` / `<dataValidations>` entry, and one per dropped
   *  `<conditionalFormatting>` block — a block carries several rules, so
   *  one increment can cost more than one rule. OOXML only.
   *  Fed by: `io.sheet.overlay.skip`, `io.cf.skip`. */
  uint32_t skipped_feature_count;
  /** Workbook parts whose declared content type was unrecognised, so the
   *  read fell back to the plain `.xlsx` kind. At most 1 per load.
   *  OOXML only.
   *  Fed by: `ooxml.reader.unknown_workbook_content_type`. */
  uint32_t unknown_content_type_count;
} fm_read_diagnostics_t;

/**
 * @brief Returns the load diagnostics captured when `wb` was created.
 *
 * `*out_diagnostics` is written to its all-zero state before argument
 * validation and stays all-zero on every non-`kOk` return, so a caller
 * may reuse one struct across calls without observing a previous call's
 * numbers. A workbook built in memory rather than loaded reports every
 * counter as zero.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if either pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_read_diagnostics(const fm_workbook_t* wb, fm_read_diagnostics_t* out_diagnostics);

/**
 * @brief Reports the estimated heap footprint of `wb` in bytes.
 *
 * For hosts whose collector sizes an object by the handle rather than by
 * what the handle owns: a workbook is pointer-sized to a garbage
 * collector and arbitrarily large in reality, so a runtime that wants
 * collection pressure to track the real cost has to be told the figure.
 *
 * The value covers the cell store, the shared-string storage, the
 * passthrough part payloads and the workbook-level metadata; it excludes
 * allocator overhead, the dependency graph and the style tables. It is an
 * estimate for pressure reporting, not an allocation ledger, and it walks
 * every materialised cell — call it at coarse boundaries (after a load or
 * a recalc), not per edit.
 *
 * @param wb         Workbook to measure.
 * @param out_bytes  On success receives the estimate.
 * @return `kOk` on success; `kBindingNullPointer` if either pointer is
 *         `NULL`.
 */
FM_API fm_status_t fm_workbook_memory_usage(const fm_workbook_t* wb, size_t* out_bytes);

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

/*
 * Failure-path contract shared by the whole save family
 * (`fm_workbook_save`, `fm_workbook_save_as`,
 * `fm_workbook_save_with_diagnostics`):
 *
 *   * Every non-NULL out-parameter is written to its zero state
 *     (`*out_bytes = NULL`, `*out_len = 0`, every counter `0`) *before*
 *     argument and model validation, and that zero state survives every
 *     non-`kOk` return. A caller may therefore reuse the same out variables
 *     across calls without observing a previous call's pointer, and the
 *     "initialise to NULL, call, free unconditionally" idiom stays safe
 *     because `fm_buffer_free(NULL)` is a no-op.
 *   * A NULL out-parameter is rejected with `kBindingNullPointer` and never
 *     dereferenced.
 *   * On `kOk`, `*out_bytes` owns exactly `*out_len` bytes allocated with
 *     `new uint8_t[]` and must be released with `fm_buffer_free`.
 */

/**
 * @brief Serialises the workbook to an in-memory `.xlsx` byte stream.
 *
 * On success the caller receives a heap-allocated buffer that MUST be
 * released with `fm_buffer_free` (NOT `free`). This ownership rule applies
 * to every save API returning bytes, including the diagnostics and legacy
 * XLSB-result variants. Mixing allocators across the boundary is undefined.
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
 * @brief Container format selector for `fm_workbook_save_as`.
 *
 * Mirrors `formulon::io::WorkbookFormat`. `FM_WORKBOOK_FORMAT_UNKNOWN`
 * is not a valid save target; passing it to `fm_workbook_save_as`
 * returns `kInvalidArgument`.
 */
typedef enum {
  FM_WORKBOOK_FORMAT_UNKNOWN = 0,
  FM_WORKBOOK_FORMAT_XLSX = 1,
  FM_WORKBOOK_FORMAT_XLSB = 2,
} fm_workbook_format_t;

/**
 * @brief Serialises the workbook to an in-memory byte stream in the
 *        requested container `format`.
 *
 * `FM_WORKBOOK_FORMAT_XLSX` produces the same bytes as
 * `fm_workbook_save`. `FM_WORKBOOK_FORMAT_XLSB` produces an MS-XLSB
 * package via the `.xlsb` writer.
 * `format` is a raw signed 32-bit ordinal so FFI callers can probe unknown
 * values without constructing an out-of-domain C enum; unknown values return
 * `kInvalidArgument`.
 *
 * On success the caller receives a heap-allocated buffer that MUST be
 * released with `fm_buffer_free` (NOT `free`). This applies to every save
 * API returning `out_bytes`: `fm_workbook_save`, `fm_workbook_save_as`,
 * and `fm_workbook_save_with_diagnostics`. Mixing allocators across the
 * boundary is undefined.
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param format     Target container format.
 * @param out_bytes  Receives a pointer to the freshly allocated buffer.
 * @param out_len    Receives the buffer length in bytes.
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` if `format` is `FM_WORKBOOK_FORMAT_UNKNOWN`
 *         or otherwise undocumented;
 *         a `kIo*` code on archive / writer failure.
 */
FM_API fm_status_t fm_workbook_save_as(const fm_workbook_t* wb, int32_t format, uint8_t** out_bytes, size_t* out_len);

/**
 * @brief Loss and fallback counters produced by one save.
 *
 * The save-side counterpart to `fm_read_diagnostics_t`, under the same
 * contract.
 *
 * Summing these fields does NOT give "how many things did I lose":
 * `dropped_part_count` and `dropped_relationship_count` intentionally
 * both rise for one dropped part. See `dropped_relationship_count`.
 */
typedef struct {
  /** Formula cells emitted as their cached literal because the AST could
   *  not be lowered to the container's encoding. XLSB only: the OOXML
   *  writer emits formula text verbatim.
   *  Fed by: `xlsb.writer.formula_downgraded`,
   *  `xlsb.writer.array_formula_downgraded`. */
  uint32_t downgraded_formula_count;
  /** Modelled sheet features the writer could not lower to records at
   *  all. XLSB only: the OOXML writer represents every modelled feature,
   *  and rejects the one case it cannot (a table whose metadata names a
   *  removed sheet) with `kIoWriteFailed` before building any part, so
   *  the caller loses the save rather than the table.
   *  Fed by: `xlsb.writer.deferred`. */
  uint32_t deferred_feature_count;
  /** Passthrough parts dropped, either because their package path
   *  collided with a path the writer generates itself (the generated copy
   *  wins) or because two passthrough entries claimed the same path.
   *  Fed by: `ooxml_writer.passthrough_collision`,
   *  `xlsb.writer.passthrough_collision`. */
  uint32_t dropped_part_count;
  /** Round-tripped relationships not re-emitted because the part they
   *  target is no longer in the package. Emitting them would leave a
   *  dangling relationship and Excel would open the package in repair
   *  mode, so dropping is the correct write.
   *
   *  This observes a loss rather than being an independent one. A part
   *  dropped while a relationship pointed at it raises BOTH this counter
   *  and `dropped_part_count` by one, for that single part; the two are
   *  separate observations of one event and must not be added together.
   *  It is deliberately not suppressed in that case, because a
   *  relationship can equally dangle for a part that was never captured
   *  at all, and a suppressed count would make those indistinguishable.
   *  Fed by: `ooxml_writer.package_rel_skipped`,
   *  `ooxml_writer.workbook_rel_skipped`,
   *  `xlsb.writer.package_rel_skipped`,
   *  `xlsb.writer.workbook_rel_skipped`,
   *  `xlsb.writer.sheet_rel_skipped`. */
  uint32_t dropped_relationship_count;
  /** Parts emitted under a writer-assigned id rather than the id the
   *  model carried. The part still ships; an external reference to the
   *  old id does not survive. OOXML only: the binary writer does not
   *  reassign part ids.
   *  Fed by: `ooxml_writer.table_id_fallback`. */
  uint32_t renumbered_part_count;
} fm_save_diagnostics_t;

/**
 * @brief Serialises `wb` and reports what the save cost.
 *
 * `format` uses the same raw signed 32-bit ordinal contract as
 * `fm_workbook_save_as`. Every output parameter — including every
 * `*out_diagnostics` field — is written to its zero / `NULL` state before
 * argument and model validation, and stays there on every non-`kOk`
 * return. `*out_diagnostics` is therefore always written, success or
 * failure; on failure it reads all-zero.
 *
 * The returned `out_bytes` buffer follows the `fm_buffer_free` ownership
 * rule documented above.
 */
FM_API fm_status_t fm_workbook_save_with_diagnostics(const fm_workbook_t* wb, int32_t format, uint8_t** out_bytes,
                                                     size_t* out_len, fm_save_diagnostics_t* out_diagnostics);

/**
 * @brief Releases a buffer returned by any Formulon save API returning
 *        `out_bytes`: `fm_workbook_save`, `fm_workbook_save_as`, or
 *        `fm_workbook_save_with_diagnostics`.
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
 * The returned pointer is a model-backed view into the workbook. It remains
 * valid until a mutation that touches the sheet list (for example
 * `fm_workbook_add_sheet`) or until the handle is destroyed. Reads on the
 * same handle, including read-scratch-backed reads, do not invalidate it.
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
 * @param utf8_name  Strict UTF-8 display name; must be non-NULL. Worksheet
 *                   identity uses locale-independent Unicode simple case
 *                   folding (C/S mappings only, without normalization).
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidSheetName` when `utf8_name` is empty, malformed UTF-8,
 *         exceeds 31 UTF-16 units, contains a forbidden character
 *         (`: \ / ? * [ ]`), or collides under locale-independent Unicode
 *         simple case folding with an existing sheet;
 *         `kSheetCountLimitExceeded` when the workbook already contains
 *         the maximum supported sheet count.
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
 * References to the removed sheet are rewritten to `#REF!` at the affected
 * subexpression, so surrounding literals and unrelated operands of the same
 * formula survive. This applies to cell formulas, defined-name definitions,
 * conditional-format and data-validation formulas, and table formulas
 * alike. Only names *scoped* to the removed sheet are dropped;
 * workbook-scoped names are retained with their `#REF!` definitions, and
 * names scoped to a later sheet follow their owner as the sheet list closes
 * the gap. Tables owned by the removed sheet are dropped and surviving
 * table sheet indices are remapped; pivot caches whose worksheet source was
 * removed, and the pivot tables bound to them, are dropped. The recalc
 * engine's dep-graph entries for cells on the removed sheet are
 * unregistered.
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
 * Updates the sheet's stored name and rewrites affected references in cell
 * formulas, defined names, conditional-format and data-validation formulas,
 * hyperlinks, table formulas, and pivot sources. Unrelated literals,
 * unresolved references, and external-workbook references remain
 * byte-identical. Worksheet identity matching uses locale-independent
 * Unicode simple case folding, so a casing-only rename updates references
 * to the same sheet.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kSheetIndexOutOfRange` when `index` is out of range;
 *         `kInvalidSheetName` when `new_name` is empty, malformed UTF-8,
 *         exceeds 31 UTF-16 units, contains a forbidden character
 *         (`: \ / ? * [ ]`), or collides under locale-independent Unicode
 *         simple case folding with an existing sheet. Folding is not
 *         normalization, full (expanding) case folding, locale tailoring, or
 *         a claim of full Excel-equivalence.
 */
FM_API fm_status_t fm_workbook_rename_sheet(fm_workbook_t* wb, uint32_t index, const char* new_name);

/**
 * @brief Sets (or appends, or removes) a workbook-scoped defined name.
 *
 * If a workbook-scoped entry with `name` already exists (case-
 * insensitive match), its formula text is replaced. Otherwise — when
 * `formula` is non-empty — a new entry is appended. Passing an empty
 * `formula` removes the existing entry, or is a no-op when no such
 * entry is present. Use `fm_workbook_set_defined_name_scoped` for
 * sheet-scoped names.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `name` is empty.
 */
FM_API fm_status_t fm_workbook_set_defined_name(fm_workbook_t* wb, const char* name, const char* formula);

/**
 * @brief Sets (or appends, or removes) a defined name in a specific scope.
 *
 * `local_sheet_id == -1` selects workbook scope. Values `>= 0` select
 * the corresponding 0-based sheet-local scope. An empty `formula`
 * removes the matching entry in that scope.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `name` is empty or `local_sheet_id`
 *         is outside `[-1, sheet_count)`.
 */
FM_API fm_status_t fm_workbook_set_defined_name_scoped(fm_workbook_t* wb, const char* name, const char* formula,
                                                       int32_t local_sheet_id);

/* -------------------------------------------------------------------------- */
/* Row / column structural edits                                              */
/* -------------------------------------------------------------------------- */

/*
 * The four edits below move every structure the engine models: cells,
 * formulas, merges, conditional-format ranges, data validations,
 * hyperlinks, tables, print areas, manual breaks and the auto-filter
 * range. What they do not move is the worksheet content the engine keeps
 * byte-verbatim because it does not model it -- the worksheet-level
 * `<extLst>` (2010+ extensions: the `x14` conditional-formatting block
 * behind DataBar negative-fill / axis / gradient settings, sparkline
 * groups, slicer anchors) and any unmodelled `<worksheet>` child retained
 * to keep it from being dropped on save.
 *
 * A `ref`, an `sqref` or a formula inside those keeps its pre-edit
 * rectangle. The visible outcome depends on the extension: Excel drops an
 * `x14` conditional-formatting entry whose range no longer matches the
 * legacy rule it extends, so an extended DataBar quietly falls back to its
 * legacy rendering, while a sparkline group keeps drawing and reads its
 * source range from the wrong cells. Neither is reported: no status code,
 * no diagnostic counter.
 *
 * A host that edits rows or columns on a sheet carrying those extensions
 * should re-author them after the edit rather than rely on the round trip.
 */

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
 *         past the sheet bound, or `count == 0`. Delete operations also
 *         reject a `count` that extends past the sheet bound.
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
 * @brief Stores a static Excel error literal at `(row, col)`.
 *
 * This writes a cell whose value is `#DIV/0!`, `#VALUE!`, etc. It is
 * distinct from writing a formula that evaluates to an error.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` or `error` is out of range.
 */
FM_API fm_status_t fm_workbook_set_error(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                         fm_error_code_t error);

/**
 * @brief Stores a text literal. The destination cell copies the UTF-8
 *        contents, so `utf8` does not need to outlive the call.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_set_text(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        const char* utf8);

/**
 * @brief Stores the phonetic guide (OOXML `<rPh>`) for a cell.
 *
 * The guide is independent of the cell's visible text. Passing an empty
 * string clears an existing guide. The destination cell copies `utf8`, so
 * the input does not need to outlive the call.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_set_cell_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
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
 * For `FM_VAL_TEXT`, `out->u.text` is a read-scratch-backed view into the
 * handle's transient read storage. It may be invalidated by the next
 * successful scratch-backed read on this handle; a validation-rejected call
 * does not refresh the storage. It is also invalid after a mutation or
 * handle destruction. The scratch is reset on each successful producer so a
 * long-lived read loop does not accumulate memory. Callers that need to
 * retain the string must copy it before another call.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_value(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                         fm_value_t* out);

/**
 * @brief Reads the cell's phonetic guide (OOXML `<rPh>`), or an empty string
 *        when the cell has no guide.
 *
 * `*out_text` is a read-scratch-backed view into the handle's transient read
 * storage. It may be invalidated by the next successful scratch-backed read
 * on this handle; a validation-rejected call does not refresh the storage.
 * It is also invalid after a mutation or handle destruction. Copy it before
 * another call when it must outlive that window.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out_text` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_cell_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                 uint32_t col, const char** out_text);

/**
 * @brief Renders the lambda closure stored at `(sheet_index, row, col)`
 *        as Excel formula text.
 *
 * Use this to recover a printable form of a `Value::Lambda` whose
 * source is otherwise opaque — for example, a cell whose formula
 * evaluates to a lambda via `=LET(...) -> LAMBDA(...)` rather than
 * carrying a literal `LAMBDA(...)` in the formula text. The rendering
 * is the full surface form `LAMBDA(p1,p2,body)` (no leading `=`),
 * suitable for re-parsing through `fm_workbook_set_formula`.
 *
 * On success `*out_text` is a read-scratch-backed view into the handle's
 * transient read storage. It may be invalidated by the next successful
 * scratch-backed read on this handle; a validation-rejected call does not
 * refresh the storage. It is also invalid after a mutation or handle
 * destruction. The scratch is reset on each successful producer so a
 * long-lived handle does not accumulate memory. Callers that need to retain
 * the string must copy it before another call.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range, the
 *         cell is absent, or the cached value is not a lambda.
 */
FM_API fm_status_t fm_workbook_lambda_text_at(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              const char** out_text);

/* -------------------------------------------------------------------------- */
/* Ad-hoc, side-effect-free formula evaluation                                */
/* -------------------------------------------------------------------------- */

/**
 * @brief Evaluates `formula` as if entered at `(row, col)` on
 *        `sheet_index`, returning a single scalar result — WITHOUT
 *        mutating the workbook.
 *
 * The formula is parsed and evaluated against a read-only view of the
 * workbook: local and qualified/renamed cross-sheet references,
 * workbook-scoped defined names, and `ROW()` / `COLUMN()` all resolve
 * relative to `(row, col)`, with full locale / coercion / 1904 fidelity
 * from the workbook profile. No cell value is written, no dynamic-array
 * spill is committed, and the dependency graph is untouched. The `const`
 * on `wb` is the ABI-level purity contract.
 *
 * An Array / spill result is reduced to its top-left element. This is a
 * pragmatic API-shape choice, NOT Excel implicit intersection (which
 * selects the element sharing the anchor's row / column and yields
 * `#VALUE!` when there is none) and NOT dynamic-array spilling (which
 * returns the whole array). To recover the full multi-cell result instead,
 * use `fm_workbook_evaluate_formula_array` + `_array_cell`.
 *
 * For `FM_VAL_TEXT`, `out->u.text` is a read-scratch-backed view and may be
 * invalidated by the next successful scratch-backed read on this handle;
 * a validation-rejected call does not refresh the storage. Mutation and
 * handle destruction also invalidate it; copy it before another call if you
 * need to retain it.
 *
 * @note Because the ad-hoc formula is never inserted into the dependency
 *       graph, a self-reference such as `=A1` anchored at A1's own address
 *       reads A1's *cached* value rather than raising `#REF!` or engaging
 *       iterative calc. This diverges from true in-cell F9 entry;
 *       self-reference detection is out of scope for this API.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb`, `formula`, or `out` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_evaluate_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                const char* formula, fm_value_t* out);

/**
 * @brief Evaluates `formula` as a conditional-formatting rule predicate
 *        anchored at `(row, col)`, with relative references written
 *        relative to `(anchor_row, anchor_col)` (the CF-applied range's
 *        top-left) — WITHOUT mutating the workbook.
 *
 * Relative references are shifted from the anchor to `(row, col)` before
 * evaluation, mirroring how Excel re-anchors a shared CF rule per cell.
 * The result is coerced with Excel's CF-predicate rules rather than
 * propagated verbatim: an error, blank, text, or numeric-zero result
 * yields `FALSE` (rule does not fire); any non-zero number yields `TRUE`.
 * `*out` is always `FM_VAL_BOOL`. Read-only, identical purity contract to
 * `fm_workbook_evaluate_formula`.
 *
 * @note The same self-reference caveat as `fm_workbook_evaluate_formula`
 *       applies: a rule that reads its own anchored cell sees the cached
 *       value rather than raising `#REF!`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb`, `formula`, or `out` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_evaluate_cf_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                   uint32_t col, uint32_t anchor_row, uint32_t anchor_col,
                                                   const char* formula, fm_value_t* out);

/**
 * @brief Evaluates `formula` as if entered at `(row, col)` and stashes the
 *        WHOLE result — array included — on the handle, returning its
 *        dimensions. Companion to `fm_workbook_evaluate_formula_array_cell`.
 *
 * This is the multi-cell counterpart to `fm_workbook_evaluate_formula`,
 * which reduces an array / spill result to its top-left element. Here the
 * result is preserved in full: a dynamic-array formula such as
 * `=SEQUENCE(2,3)` reports `*out_rows = 2`, `*out_cols = 3`; a scalar
 * result such as `=1+2` reports `*out_rows = *out_cols = 1` (a 1x1 array).
 * Parsing, anchoring, cross-sheet / defined-name resolution, `ROW()` /
 * `COLUMN()`, and the read-only purity contract are identical to
 * `fm_workbook_evaluate_formula` (the `const` on `wb` is the ABI-level
 * purity guarantee; the stash is internal handle state, not observable
 * workbook state, mirroring how the read scratch is populated on the read
 * path).
 *
 * Two-step protocol: call this to evaluate + learn the dimensions, then call
 * `fm_workbook_evaluate_formula_array_cell` for each row-major index in
 * `[0, (*out_rows) * (*out_cols))`. The handle's private array stash remains
 * available until the next `fm_workbook_evaluate_formula_array` call or any
 * mutation on this handle, whichever comes first. Text returned by each
 * `_array_cell` call is a read-scratch-backed view and may be invalidated by
 * the next successful scratch-backed read on this handle; a
 * validation-rejected call does not refresh the storage. Copy it before
 * another call when retaining it.
 *
 * A degenerate empty array result (which producers should never emit)
 * surfaces as a single `#VALUE!` cell with `*out_rows = *out_cols = 1`,
 * matching the scalar reduction's empty-array handling.
 *
 * @note The same self-reference caveat as `fm_workbook_evaluate_formula`
 *       applies: a self-reference reads the target cell's cached value.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb`, `formula`, `out_rows`, or
 *         `out_cols` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_evaluate_formula_array(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                      uint32_t col, const char* formula, uint32_t* out_rows,
                                                      uint32_t* out_cols);

/**
 * @brief Reads the row-major `index`-th cell of the most recent
 *        `fm_workbook_evaluate_formula_array` result into `*out`.
 *
 * Valid `index` values are `[0, rows * cols)` where `rows` / `cols` are the
 * dimensions the producing `fm_workbook_evaluate_formula_array` call
 * returned; cell `(r, c)` is at `index = r * cols + c`.
 *
 * For a `FM_VAL_TEXT` cell, `out->u.text` is a read-scratch-backed view
 * (scratch is refreshed by each successful producer, exactly like
 * `fm_workbook_cell_at`) and may be invalidated by the next successful
 * scratch-backed read on this handle. A validation-rejected call does not
 * refresh the storage; mutation and handle destruction also invalidate it.
 * Copy it before another call if you need to retain it.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb` or `out` is `NULL`;
 *         `kInvalidArgument` when `index` is past the stashed result (or no
 *         array evaluation has run on this handle yet).
 */
FM_API fm_status_t fm_workbook_evaluate_formula_array_cell(const fm_workbook_t* wb, size_t index, fm_value_t* out);

/* -------------------------------------------------------------------------- */
/* Iteration / dump                                                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the number of stored cell slots on `sheet_index`.
 *
 * Counts every populated `Cell` (literal, formula, or implicitly created
 * during row growth) on the sheet's row-sparse / column-dense storage,
 * plus every dynamic-array spill phantom that has no underlying stored
 * slot. Spill phantoms carry an effective value through
 * `fm_workbook_get_value` / `fm_workbook_cell_at` but live only in the
 * spill table; counting them keeps this total aligned with the
 * `fm_workbook_cell_at` index range.
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
 * formula pointer is a model-backed view into the cell's own storage and is
 * valid until the next mutation that touches the sheet's cell store or until
 * the handle is destroyed. Reads on the same handle do not invalidate it.
 *
 * A `FM_VAL_TEXT` payload in `*out_value`, by contrast, is a
 * read-scratch-backed view and may be invalidated by the next successful
 * scratch-backed read on this handle; a validation-rejected call does not
 * refresh the storage. Mutation and handle destruction also invalidate it.
 * The scratch is refreshed by each successful producer so iterating
 * `cell_at` does not accumulate memory.
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
 * @brief Reads the `idx`-th defined name with its scope.
 *
 * On success `*out_name` and `*out_formula` are model-backed views into the
 * workbook handle. Both are valid until the next mutation of the
 * defined-name list or until the handle is destroyed. Reads do not
 * invalidate them.
 *
 * `out_local_sheet_id` is optional: when non-NULL it receives `-1` for
 * workbook scope, or a 0-based sheet index for sheet-local scope.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb`, `out_name`, or `out_formula` is
 *         `NULL`;
 *         `kInvalidArgument` when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_defined_name_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                               const char** out_formula, int32_t* out_local_sheet_id);

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
 * On success `*out_name`, `*out_display_name`, and `*out_ref` are
 * model-backed views into the workbook handle. They have the same lifetime
 * contract as `fm_workbook_defined_name_at`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_table_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                        const char** out_display_name, const char** out_ref, size_t* out_sheet_index);

/** Creates a worksheet table (`xl/tables/tableN.xml`). `column_names` must
 * contain one non-empty header for every column in `ref`; each name must be
 * unique within the table, and `column_count` must equal the width of `ref`.
 * `style_name` may be NULL or empty to omit table styling. The table's filter
 * definition is created automatically. When `header_row` is set, the caller
 * still owns the header cells: Excel expects the first row of `ref` to hold
 * exactly the column names. */
FM_API fm_status_t fm_workbook_table_create(fm_workbook_t* wb, size_t sheet_index, const char* ref, const char* name,
                                            const char* display_name, const char* const* column_names,
                                            size_t column_count, const char* style_name, int32_t header_row,
                                            int32_t totals_row, size_t* out_index);

/** Replaces a table's range and optional visual style. `ref` is required;
 * `style_name == NULL` preserves the existing raw style payload, while an
 * empty style removes `<tableStyleInfo>`. Negative `header_row` and
 * `totals_row` preserve their existing flags; zero and positive values set
 * them false and true respectively. `ref` must keep the width the table's
 * column list already describes. */
FM_API fm_status_t fm_workbook_table_update(fm_workbook_t* wb, size_t index, const char* ref, const char* style_name,
                                            int32_t header_row, int32_t totals_row);

/** Removes the table at `index`. */
FM_API fm_status_t fm_workbook_table_remove(fm_workbook_t* wb, size_t index);

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
 * On success `*out_path` is a model-backed view into the workbook's
 * passthrough-part list. It remains valid until the passthrough list is
 * replaced by a reader/writer round-trip, another invalidating workbook
 * mutation, or handle destruction. Reads do not invalidate it.
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
 *        model-backed views into the workbook handle; they are valid
 *        until the next mutation that touches the sheet's hyperlink
 *        list or until the handle is destroyed. Reads do not invalidate
 *        them. Newly-added entries
 *        accept NULL for any of `target`, `display`, `tooltip`,
 *        `location` to mean "empty".
 */
typedef struct {
  uint32_t row;
  uint32_t col;
  uint32_t last_row;
  uint32_t last_col;
  const char* target;
  const char* location;
  const char* display;
  const char* tooltip;
} fm_hyperlink;

/**
 * @brief Comment descriptor. `author` and `text` are model-backed UTF-8
 *        views with the same lifetime contract as `fm_hyperlink`.
 */
typedef struct {
  uint32_t row;
  uint32_t col;
  const char* author;
  const char* text;
} fm_comment;

/**
 * @brief Appends a hyperlink covering the inclusive rectangle
 *        `(hl.row, hl.col)..(hl.last_row, hl.last_col)` to `sheet`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range or the hyperlink
 *         rectangle is inverted or outside the Excel grid.
 */
FM_API fm_status_t fm_sheet_add_hyperlink(fm_workbook_t* wb, uint32_t sheet, fm_hyperlink hl);

/**
 * @brief Removes every hyperlink anchored at the exact `(row, col)`
 *        cell on `sheet`. No-op when nothing matches.
 *
 * @return `kOk` on success (including the no-match case);
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_hyperlink(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col);

/**
 * @brief Removes the hyperlink at `index` on `sheet`. Mirrors the
 *        index domain of fm_sheet_get_hyperlink_at.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` or `index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_hyperlink_at(fm_workbook_t* wb, uint32_t sheet, uint32_t index);

/**
 * @brief Drops every hyperlink on `sheet`. No-op when the sheet
 *        already has no hyperlinks.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_hyperlinks(fm_workbook_t* wb, uint32_t sheet);

/**
 * @brief Appends a merge range to the sheet's merge list.
 *        Corners are normalised internally. The resulting inclusive
 *        rectangle must lie within the Excel grid.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range or the merge
 *         rectangle is outside the Excel grid.
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
 *        comment at `(row, col)` on `sheet`. Returns `kNotFound` when no
 *        comment is anchored there, and `kInvalidArgument` when `sheet` is
 *        out of range.
 *
 * `out->author` and `out->text` are model-backed views into the workbook
 * handle. Both are valid until the next mutation that touches the sheet's
 * comment list or until the handle is destroyed. Reads do not invalidate
 * them.
 */
FM_API fm_status_t fm_sheet_get_comment_at(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                           fm_comment* out);

/**
 * @brief Returns the number of comments attached to `sheet`, including
 *        comments anchored on cells that carry no value.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 */
FM_API fm_status_t fm_sheet_get_comment_count(fm_workbook_t* wb, uint32_t sheet, uint32_t* out_count);

/**
 * @brief Reads the `index`-th comment on `sheet` into `out`, in storage
 *        order. Unlike `fm_sheet_get_comment_at`, this enumerates every
 *        comment on the sheet regardless of whether the anchor cell holds
 *        a value, so callers can discover comments without already
 *        knowing their `(row, col)`.
 *
 * `out->author` and `out->text` are model-backed views into the workbook
 * handle, with the same lifetime contract as `fm_sheet_get_comment_at`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet` or `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_comment_at_index(fm_workbook_t* wb, uint32_t sheet, uint32_t index, fm_comment* out);

/**
 * @brief Reads the `index`-th hyperlink on `sheet` into `out`.
 *
 * String fields in `out` are model-backed views (see `fm_hyperlink`).
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
 *        comment. The strings are copied into the workbook handle's text
 *        storage. The coordinate must lie within the Excel grid.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range or the coordinate
 *         is outside the Excel grid.
 */
FM_API fm_status_t fm_sheet_set_comment(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                        const char* author, const char* text);

/**
 * @brief Data-validation rule descriptor.
 *
 * Mirrors `formulon::DataValidation`:
 *   * `ranges` / `range_count` — the cell-range list this rule applies to.
 *     For getter results the pointer is a model-backed view into the
 *     workbook handle (see `fm_sheet_get_validation_at` for the lifetime
 *     contract). For `fm_sheet_add_validation` the caller owns the buffer and
 *     must keep
 *     it alive for the duration of the call only; the implementation
 *     deep-copies every range into the sheet's owned storage.
 *   * `type` — 0 none, 1 whole, 2 decimal, 3 list, 4 date, 5 time,
 *     6 textLength, 7 custom.
 *   * `op`   — 0 between, 1 notBetween, 2 equal, 3 notEqual,
 *     4 greaterThan, 5 lessThan, 6 greaterThanOrEqual, 7 lessThanOrEqual.
 *   * `error_style` — 0 stop, 1 warning, 2 information.
 *   * `allow_blank` / `show_input_message` / `show_error_message` /
 *     `show_dropdown` — 0 = false, 1 = true. `show_dropdown` is the
 *     user-facing "is the in-cell dropdown arrow shown" meaning for
 *     `list` validations (default true); this is the inverse of the raw
 *     OOXML `showDropDown` XML attribute, which the reader/writer already
 *     translate.
 *   * `formula1` / `formula2` / `error_title` / `error_message` /
 *     `prompt_title` / `prompt_message` — model-backed views on getter
 *     results, and caller-owned input strings on the add path. On the input
 *     path (`fm_sheet_add_validation`) `NULL`
 *     means "absent" and is treated as the empty string. Getters always
 *     return non-NULL pointers (possibly pointing at an empty string).
 */
typedef struct {
  const fm_merge_range* ranges;
  uint32_t range_count;
  uint8_t type;
  uint8_t op;
  uint8_t error_style;
  int32_t allow_blank;
  int32_t show_input_message;
  int32_t show_error_message;
  int32_t show_dropdown;
  const char* formula1;
  const char* formula2;
  const char* error_title;
  const char* error_message;
  const char* prompt_title;
  const char* prompt_message;
} fm_data_validation;

/**
 * @brief Returns the number of data-validation rules attached to `sheet`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 */
FM_API fm_status_t fm_sheet_get_validation_count(fm_workbook_t* wb, uint32_t sheet, uint32_t* out_count);

/**
 * @brief Reads the `index`-th data-validation rule on `sheet` into `out`.
 *
 * Every pointer in `*out` (`ranges` and the six string fields) is a
 * model-backed view owned by the workbook handle. It remains valid until the
 * next mutation that touches the sheet's validation list (any of
 * `fm_sheet_add_validation`, `fm_sheet_remove_validation_at`,
 * `fm_sheet_clear_validations`, or any reader / writer round-trip), or
 * until the handle is destroyed. Callers that need to retain the
 * payload across mutations must copy it.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet` or `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_validation_at(fm_workbook_t* wb, uint32_t sheet, uint32_t index,
                                              fm_data_validation* out);

/**
 * @brief Appends a data-validation rule to the sheet's validation list.
 *
 * Every string field accepts `NULL`, which is treated as the empty
 * string. `v.ranges` may be `NULL` only when `v.range_count == 0`
 * (a rule with no anchor ranges); otherwise it must point to
 * `v.range_count` consecutive `fm_merge_range` values that the
 * implementation copies into the rule's owned storage. Range corners
 * are not normalised — callers should pass already-normalised
 * `(first <= last)` rectangles. Every rectangle must lie within the
 * Excel grid; all ranges are validated before the rule is stored.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range, when
 *         `v.range_count > 0 && v.ranges == NULL` or `v.range_count`
 *         exceeds the per-call range cap, when any supplied range is
 *         outside the Excel grid, or when `v.type`, `v.op` or
 *         `v.error_style` is outside the domain listed above.
 */
FM_API fm_status_t fm_sheet_add_validation(fm_workbook_t* wb, uint32_t sheet, fm_data_validation v);

/**
 * @brief Removes the validation rule at `index` on `sheet`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` or `index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_validation_at(fm_workbook_t* wb, uint32_t sheet, uint32_t index);

/**
 * @brief Drops every validation rule on `sheet`. No-op when the sheet
 *        already has no rules.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_validations(fm_workbook_t* wb, uint32_t sheet);

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

/** @brief Per-call telemetry returned by `fm_workbook_recalc_parallel`. */
typedef struct {
  uint64_t cells_evaluated;
  uint64_t sccs_processed;
  uint64_t parallel_steps;
  uint64_t serial_fallback_steps;
  uint64_t cycle_recoveries;
  uint32_t worker_threads_started;
  uint32_t worker_threads_used;
} fm_parallel_recalc_stats;

/**
 * @brief Drives a full incremental recalc using the parallel SCC scheduler.
 *
 * `thread_count == 0` selects automatic detection (capped at 8); `1` keeps
 * all evaluation on the caller thread and starts no worker OS threads; values
 * `2..8` select the maximum worker count. The scheduler may start fewer
 * workers when the operating system refuses a launch and then completes the
 * pass through serial fallback. Workers are created once per complete call,
 * including any internal dynamic-array spill retry waves, and are joined
 * before this function returns.
 *
 * `out_stats` is optional. When non-NULL it is zeroed on entry and remains
 * zero on every failure path. On success it reports the work performed by
 * this call; `worker_threads_started` counts successfully launched OS
 * workers and `worker_threads_used` counts distinct launched workers that
 * claimed at least one SCC task. A configured iterative-progress callback
 * is always invoked on the caller thread for cyclic layers.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `thread_count > 8`;
 *         `kGraphRecalcReentrant` when called recursively from a recalc
 *         callback (including a scheduler worker);
 *         a `kEval*` / `kGraph*` code on engine failure.
 */
FM_API fm_status_t fm_workbook_recalc_parallel(fm_workbook_t* wb, uint32_t thread_count,
                                               fm_parallel_recalc_stats* out_stats);

/**
 * @brief Configures iterative-calculation knobs for the workbook's
 *        recalc engine.
 *
 * @param enabled         `0` disables iterative calc (the default,
 *                        cyclic SCCs surface `#REF!`); non-zero
 *                        enables fixed-point resolution up to
 *                        `max_iterations` passes.
 * @param max_iterations  Iteration cap. Values < 1 are treated as 1 and
 *                        values above 32767 are treated as 32767, which
 *                        is Excel's own limit for this setting.
 *                        `fm_workbook_get_iterative` reports the clamped
 *                        value, not the requested one.
 * @param max_change      Absolute convergence threshold.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`.
 */
FM_API fm_status_t fm_workbook_set_iterative(fm_workbook_t* wb, int32_t enabled, int32_t max_iterations,
                                             double max_change);

/**
 * @brief Changes only whether iterative calculation is enabled.
 *
 * Preserves the workbook's existing iteration cap and convergence threshold.
 * This is useful for hosts that expose Excel's enable checkbox separately
 * from its advanced iterative-calculation settings.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`.
 */
FM_API fm_status_t fm_workbook_set_iterative_enabled(fm_workbook_t* wb, int32_t enabled);

/**
 * @brief Reads the iterative-calculation settings currently stored on a workbook.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_iterative(const fm_workbook_t* wb, int32_t* out_enabled,
                                             uint32_t* out_max_iterations, double* out_max_change);

/* -------------------------------------------------------------------------- */
/* Calculation mode (workbook-level `<calcPr calcMode>` policy)               */
/* -------------------------------------------------------------------------- */

/**
 * @brief Workbook-level calculation mode.
 *
 * Mirrors Excel's "Calculation options" workbook setting and the
 * `calcMode` attribute on the `<calcPr>` element. `kAuto` is the
 * default; `kManual` suppresses automatic recalc on input;
 * `kAutoNoTable` recalcs everything except data-table cells.
 *
 * The engine itself does NOT gate evaluation on this enum (every
 * `fm_workbook_recalc` call honours all dirty cells); it is preserved
 * as round-trip metadata and surfaced through this API so the host UI
 * can mirror Excel's user-visible state.
 */
typedef enum {
  FM_CALC_MODE_AUTO = 0,
  FM_CALC_MODE_MANUAL = 1,
  FM_CALC_MODE_AUTO_NO_TABLE = 2,
} fm_calc_mode_t;

/**
 * @brief Returns the workbook's calc mode.
 *
 * @param wb       Workbook handle. Must not be NULL.
 * @param out_mode Populated on success.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out_mode` is NULL.
 */
FM_API fm_status_t fm_workbook_calc_mode(const fm_workbook_t* wb, fm_calc_mode_t* out_mode);

/**
 * @brief Sets the workbook's calc mode.
 *
 * Plain metadata — does not affect evaluation. The value is a raw signed
 * 32-bit enum ordinal so FFI callers can pass an arbitrary value without
 * first constructing an out-of-domain C enum. Unknown values (outside the
 * documented range) return `kInvalidArgument`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` if `mode` is not a documented value.
 */
FM_API fm_status_t fm_workbook_set_calc_mode(fm_workbook_t* wb, int32_t mode);

/* -------------------------------------------------------------------------- */
/* Clock seam (`NOW` / `TODAY` and pivot relative-period filters)             */
/* -------------------------------------------------------------------------- */

/**
 * @brief A wall-clock reading in local civil fields.
 *
 * Local calendar fields are carried rather than a timestamp so a pinned
 * reading has no residual timezone interpretation: the same values
 * reproduce the same results on any host, in any zone. Every field is
 * `int32_t` for ABI stability, matching the wide-POD convention used by
 * the other structs in this header.
 *
 * Ranges accepted by `fm_workbook_set_pinned_now`: `year` in
 * `[1900, 9999]`, `month` in `[1, 12]`, `day` in `[1, N]` where `N` is the
 * length of that month under the Gregorian leap rule, `hour` in `[0, 23]`,
 * `minute` and `second` in `[0, 59]`. The pin is a calendar instant, not a
 * normalising constructor: a month of `13` is rejected rather than rolled
 * into the next year, because silently moving a host's pin would make the
 * results it pinned unexplainable.
 */
typedef struct {
  int32_t year;
  int32_t month;
  int32_t day;
  int32_t hour;
  int32_t minute;
  int32_t second;
} fm_civil_time_t;

/**
 * @brief Reads the workbook's pinned wall-clock reading, if any.
 *
 * @param wb         Workbook handle. Must not be NULL.
 * @param out_now    Populated with the pin when one is set; zero-filled
 *                   otherwise. Must not be NULL.
 * @param out_pinned Set to `1` when a pin is in effect and `0` when the
 *                   workbook follows the host clock. Must not be NULL —
 *                   without it a zero-filled `out_now` is indistinguishable
 *                   from a pin, since no civil field combination is
 *                   reserved as a sentinel.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any argument is NULL.
 */
FM_API fm_status_t fm_workbook_pinned_now(const fm_workbook_t* wb, fm_civil_time_t* out_now, int32_t* out_pinned);

/**
 * @brief Pins every clock-dependent result in the workbook to one instant.
 *
 * `NOW`, `TODAY` and the pivot relative-period filters ("this month",
 * "year to date", ...) each read the clock independently by default, which
 * makes a recalc internally inconsistent across a midnight boundary and
 * makes any such result untestable. Pinning gives the whole workbook one
 * instant to agree on.
 *
 * This is model state, not file state: nothing in the OOXML package records
 * it, so a save drops it and a reload comes back unpinned. Cached formula
 * values are not recomputed here — drive `fm_workbook_recalc` (or an
 * equivalent partial recalc) to see the pin take effect on existing cells.
 *
 * @param wb  Workbook handle. Must not be NULL.
 * @param now The reading to pin. Must not be NULL, and every field must be
 *            in the range documented on `fm_civil_time_t`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `now` is NULL;
 *         `kInvalidArgument` if any field is out of range.
 */
FM_API fm_status_t fm_workbook_set_pinned_now(fm_workbook_t* wb, const fm_civil_time_t* now);

/**
 * @brief Releases the pin so clock-dependent results follow the host clock
 *        again. Clearing an unpinned workbook succeeds and does nothing.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`.
 */
FM_API fm_status_t fm_workbook_clear_pinned_now(fm_workbook_t* wb);

/* -------------------------------------------------------------------------- */
/* Excel formula compatibility profile                                        */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the workbook's active Excel formula profile id.
 *
 * The returned pointer is a static view into Formulon's profile table and
 * remains valid for the process lifetime. Current ids are:
 * `mac-365-ja_JP`, `win-365-ja_JP`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out_profile_id` is NULL.
 */
FM_API fm_status_t fm_workbook_excel_profile_id(const fm_workbook_t* wb, const char** out_profile_id);

/**
 * @brief Sets the workbook's Excel formula profile by id.
 *
 * Existing cached formula values are not recomputed until the caller drives
 * `fm_workbook_recalc` or an equivalent partial recalc.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `profile_id` is NULL`;
 *         `kInvalidArgument` if `profile_id` is not a documented profile.
 */
FM_API fm_status_t fm_workbook_set_excel_profile_id(fm_workbook_t* wb, const char* profile_id);

/* -------------------------------------------------------------------------- */
/* Sheet protection (per-sheet `<sheetProtection>`)                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Wide POD mirror of `formulon::SheetProtection` (ECMA-376
 *        §18.3.1.85 `<sheetProtection>`).
 *
 * Strings are NUL-terminated UTF-8. On the read path
 * (`fm_sheet_get_protection`) the pointers are model-backed views into the
 * workbook's own storage and remain valid until the next mutation of the
 * same sheet's protection record or handle destruction. Reads do not
 * invalidate them. On the write path
 * (`fm_sheet_set_protection`) the strings are copied into the
 * workbook; the caller's buffers may be released after the call. Pass
 * `NULL` for any string field to leave it empty.
 *
 * Boolean flags use `int32_t` for ABI stability: `0` = false,
 * non-zero = true. The flag semantics follow the OOXML attribute
 * names verbatim — ECMA-376 documents which flags grant or restrict a
 * given operation under "Protect Sheet".
 *
 * `enabled` controls whether the `<sheetProtection>` element is
 * emitted at all. Setting `enabled = 0` clears the protection block;
 * the other fields are then preserved in memory but not written.
 */
typedef struct {
  int32_t enabled;
  const char* algorithm_name;
  const char* hash_value;
  const char* salt_value;
  uint32_t spin_count;
  const char* legacy_password;
  int32_t sheet;
  int32_t objects;
  int32_t scenarios;
  int32_t format_cells;
  int32_t format_columns;
  int32_t format_rows;
  int32_t insert_columns;
  int32_t insert_rows;
  int32_t insert_hyperlinks;
  int32_t delete_columns;
  int32_t delete_rows;
  int32_t select_locked_cells;
  int32_t select_unlocked_cells;
  int32_t sort;
  int32_t auto_filter;
  int32_t pivot_tables;
} fm_sheet_protection_t;

/**
 * @brief Reads the sheet's protection state.
 *
 * @param wb           Workbook handle. Must not be NULL.
 * @param sheet_index  Sheet index, 0-based.
 * @param out          Populated on success. String pointers are
 *                     model-backed views into the workbook's storage and
 *                     remain valid until the next protection mutation on
 *                     this sheet or handle destruction.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out` is NULL;
 *         `kInvalidArgument` if `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_protection(const fm_workbook_t* wb, uint32_t sheet_index, fm_sheet_protection_t* out);

/**
 * @brief Sets the sheet's protection state.
 *
 * Replaces the per-sheet `SheetProtection` record wholesale. Strings
 * are deep-copied into the workbook; NULL strings are stored as
 * empty.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `in` is NULL;
 *         `kInvalidArgument` if `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_protection(fm_workbook_t* wb, uint32_t sheet_index, const fm_sheet_protection_t* in);

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
 * Return non-zero to continue iterating, `0` to abort early. An aborted
 * solve leaves the workbook in its current partially-converged state.
 * The return type follows the header-wide wide-POD boolean convention
 * (`int32_t`, `0` = false); it is not a C `bool`.
 */
typedef int32_t (*fm_iterative_progress_cb)(uint32_t iteration, double max_residual, uint32_t max_iterations,
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
 *         `kInvalidArgument` when `viewport->sheet` is past the last
 *           sheet, or exceeds the 16-bit range a sheet id is stored in;
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
 * priority-ascending order, so `matches[0]` is the highest-priority
 * match. Fold first-wins: walk the list in the returned order and apply
 * each property only where nothing has set it yet, so a lower-priority
 * match fills what a higher-priority one left unset and never overwrites
 * it.
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

/* -------------------------------------------------------------------------- */
/* Print pagination                                                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Paginates one worksheet and returns an owned snapshot of its
 *        resolved print area, row/column breaks, and physical page count.
 *
 * `sheet_index` is 0-based. Release a successful result with
 * `fm_pagination_destroy`.
 *
 * Fails with `kPrintPageCountOverflow` (9005) when the page grid the sheet
 * declares exceeds the supported page count, which a file listing a manual
 * break before nearly every row and column can reach. On success the count
 * `fm_pagination_page_count` reports is the true total; it is never a
 * truncated one.
 */
FM_API fm_status_t fm_workbook_paginate(const fm_workbook_t* wb, size_t sheet_index, fm_pagination_t** out);

/** @brief Releases a pagination result. `pagination == NULL` is a no-op. */
FM_API void fm_pagination_destroy(fm_pagination_t* pagination);

/** @brief Returns the total physical page count, or zero for `NULL`. */
FM_API uint32_t fm_pagination_page_count(const fm_pagination_t* pagination);

/** @brief Number of rectangles in the sheet's declared `_xlnm.Print_Area`,
 *         or zero for `NULL`.
 *
 * Zero also means the sheet declares no print area, mirroring Excel's own
 * `PageSetup.PrintArea`; it is not backfilled with the used range. Pagination
 * still falls back to the used range in that case, so
 * `fm_pagination_page_count` can be non-zero while this is zero. */
FM_API size_t fm_pagination_print_area_count(const fm_pagination_t* pagination);

/** @brief Copies a resolved print-area rectangle by index. */
FM_API fm_status_t fm_pagination_print_area_at(const fm_pagination_t* pagination, size_t index, fm_print_range_t* out);

/** @brief Number of horizontal (row) page breaks, or zero for `NULL`. */
FM_API size_t fm_pagination_horizontal_break_count(const fm_pagination_t* pagination);

/** @brief Reads a 0-based row index that a horizontal break precedes. */
FM_API fm_status_t fm_pagination_horizontal_break_at(const fm_pagination_t* pagination, size_t index,
                                                     uint32_t* out_row);

/** @brief Number of vertical (column) page breaks, or zero for `NULL`. */
FM_API size_t fm_pagination_vertical_break_count(const fm_pagination_t* pagination);

/** @brief Reads a 0-based column index that a vertical break precedes. */
FM_API fm_status_t fm_pagination_vertical_break_at(const fm_pagination_t* pagination, size_t index, uint32_t* out_col);

/* -------------------------------------------------------------------------- */
/* Print settings - raw XML                                                   */
/* -------------------------------------------------------------------------- */

/**
 * @brief Reads a worksheet print-settings element as raw XML.
 *
 * The engine stores these five elements verbatim and re-emits them on save,
 * so what comes back is what the file stated (or what the matching setter
 * last stored), not a re-serialisation of a partial model. A sheet that
 * declares no such element yields an empty string, never NULL.
 *
 * The returned pointer follows the standard scratch contract: valid until
 * the next successful scratch-backed read, any mutation, or handle
 * destruction on the same handle.
 */
FM_API fm_status_t fm_sheet_get_page_setup_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/**
 * @brief Replaces a worksheet print-settings element with raw XML.
 *
 * `xml` must be exactly one well-formed element whose root name matches the
 * entry point (`pageSetup` here). A prefixed root (`<x:pageSetup/>`) is
 * rejected: the fragment is spliced into the worksheet unchanged, and a
 * prefix lifted out of its original namespace context would not bind there.
 * Content the engine does not model is preserved as-is.
 *
 * An empty string removes the element and resets the structured views it
 * feeds (orientation / paper size / scale / fit-to-width / fit-to-height)
 * to their ECMA-376 defaults. Otherwise the structured views are re-derived
 * from the new fragment immediately, so a `fm_workbook_paginate` that
 * follows observes the new settings.
 *
 * `fitToPage` is NOT part of `<pageSetup>` - it lives in
 * `<sheetPr><pageSetUpPr>`; use `fm_sheet_set_fit_to_page`.
 *
 * Returns `kInvalidArgument` (2) when the fragment is malformed, carries
 * more than one top-level element, or names the wrong root; and when it
 * declares an `r:id` on a sheet that has no printer-settings part, which
 * would leave a dangling relationship that makes Excel repair the file.
 * Returns `kPreconditionFailed` (9) when the fragment exceeds 4 KiB.
 */
FM_API fm_status_t fm_sheet_set_page_setup_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<pageMargins>` as raw XML. See
 *         `fm_sheet_get_page_setup_xml` for the shared contract. */
FM_API fm_status_t fm_sheet_get_page_margins_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<pageMargins>`. Re-derives the structured margins; an
 *         empty string removes the element and restores the defaults.
 *         Fragment ceiling 4 KiB. */
FM_API fm_status_t fm_sheet_set_page_margins_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<printOptions>` as raw XML. */
FM_API fm_status_t fm_sheet_get_print_options_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<printOptions>`. This element has no structured view;
 *         the raw fragment is the whole model. Fragment ceiling 4 KiB. */
FM_API fm_status_t fm_sheet_set_print_options_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<headerFooter>` as raw XML. */
FM_API fm_status_t fm_sheet_get_header_footer_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<headerFooter>`. No structured view. Fragment ceiling
 *         64 KiB, which is wider than the other elements because six
 *         sections of formatting codes (and header images) live here.
 *
 * Header and footer text uses `&` as its formatting-code introducer
 * (`&P` = page number, `&D` = date). Those codes live in XML element text,
 * so a caller composing this fragment writes them escaped - Excel's own
 * output is `<oddHeader>&amp;C...</oddHeader>`, and a raw `&` would make
 * the fragment ill-formed. `fm_sheet_set_header_footer` takes the decoded
 * text instead and handles the escaping. */
FM_API fm_status_t fm_sheet_set_header_footer_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<sheetPr>` as raw XML.
 *
 * `<sheetPr>` is not print-only: it also carries `tabColor` (sheet tab
 * colour) and `codeName` (VBA binding). Replacing it wholesale to toggle
 * `fitToPage` would drop those; `fm_sheet_set_fit_to_page` exists to avoid
 * exactly that. */
FM_API fm_status_t fm_sheet_get_sheet_pr_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<sheetPr>`. Re-derives `fitToPage` from
 *         `<pageSetUpPr>`; absent means false. Fragment ceiling 8 KiB. */
FM_API fm_status_t fm_sheet_set_sheet_pr_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/**
 * @brief Sets or clears `<sheetPr><pageSetUpPr fitToPage>` without
 *        disturbing any other `<sheetPr>` content.
 *
 * Creates `<sheetPr>` and `<pageSetUpPr>` when absent; `enabled == 0`
 * removes the `fitToPage` attribute but keeps both elements and everything
 * else they carry (`tabColor`, `codeName`, `<outlinePr>`, ...).
 *
 * `fitToPage` alone does not scale anything: it selects the fit-to-page
 * mode, and `<pageSetup fitToWidth= fitToHeight=>` states the target. Set
 * both - "fit onto one page" is `fm_sheet_set_fit_to_page(wb, s, 1)` plus
 * `fit_to_width = 1` and `fit_to_height = 1`.
 */
FM_API fm_status_t fm_sheet_set_fit_to_page(fm_workbook_t* wb, size_t sheet_index, int32_t enabled);

/* -------------------------------------------------------------------------- */
/* Print settings - print area and titles                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Writes the sheet-scoped `_xlnm.Print_Area` defined name.
 *
 * `ranges_a1` accepts one or more comma-separated A1 ranges, with or
 * without `$` anchors: `"A1:F8"`, `"A1:B10,D5:E20"`, `"A:D"`, `"1:50"`.
 * Every corner is absolutised and each area is qualified with the sheet's
 * own name, quoted per Excel's rules, so the stored formula matches what
 * Excel writes. Area order is preserved - Excel prints them in the stated
 * order.
 *
 * An empty string removes the defined name. Returns `kPrintInvalidArea`
 * (9002) when any area fails to parse, and `kInvalidArgument` (2) when
 * the list holds more areas than the per-call cap every range-taking
 * entry point shares.
 */
FM_API fm_status_t fm_sheet_set_print_area(fm_workbook_t* wb, size_t sheet_index, const char* ranges_a1);

/**
 * @brief Reads the print area back as unqualified A1 ranges, or an empty
 *        string when the sheet declares none.
 *
 * Whole-axis areas come back expanded to explicit corners (`"A:D"` reads
 * back as `"A1:D1048576"`): the resolver normalises them to rectangles for
 * the paginator, and this getter reports the resolved form rather than
 * re-deriving the authored shape.
 *
 * Returns `kPrintInvalidArea` (9002) when the defined name exists but does
 * not parse, so a malformed one is never mistaken for "no print area".
 */
FM_API fm_status_t fm_sheet_get_print_area(const fm_workbook_t* wb, size_t sheet_index, const char** out_ranges_a1);

/**
 * @brief Writes the sheet-scoped `_xlnm.Print_Titles` defined name.
 *
 * `repeat_rows` is a row span (`"1:2"`), `repeat_cols` a column span
 * (`"A:A"`); either may be empty or NULL. Both empty removes the defined
 * name. The stored formula lists rows before columns, matching Excel.
 *
 * Returns `kPrintInvalidArea` (9002) when either span is present but not a
 * whole-row / whole-column span of the right shape.
 */
FM_API fm_status_t fm_sheet_set_print_titles(fm_workbook_t* wb, size_t sheet_index, const char* repeat_rows,
                                             const char* repeat_cols);

/**
 * @brief Reads the repeat rows / columns back as `"1:2"` / `"A:A"`, each
 *        an empty string when that axis is unset.
 *
 * Both output pointers stay valid together until the next scratch-backed
 * read on the handle.
 */
FM_API fm_status_t fm_sheet_get_print_titles(const fm_workbook_t* wb, size_t sheet_index, const char** out_repeat_rows,
                                             const char** out_repeat_cols);

/* -------------------------------------------------------------------------- */
/* Print settings - manual page breaks                                        */
/* -------------------------------------------------------------------------- */

/** @brief One manual page break as the worksheet declares it. */
typedef struct {
  uint32_t id;    /* 0-based row / column index the break precedes */
  uint32_t min;   /* span start on the perpendicular axis */
  uint32_t max;   /* span end on the perpendicular axis */
  int32_t manual; /* 0 = automatic (`man` absent), 1 = user break */
} fm_page_break_t;

/**
 * @brief Upserts a row break before `row`.
 *
 * The perpendicular span defaults to the whole sheet (`min = 0`,
 * `max = 16383`), which is what Excel writes for a break inserted with no
 * active selection. Re-adding an existing index replaces its span and
 * `manual` flag rather than duplicating it; the vector stays sorted
 * ascending by index.
 *
 * Returns `kInvalidArgument` (2) for a row past the grid, and
 * `kPreconditionFailed` (9) once the axis holds 1026 breaks. A row break
 * divides the page horizontally, and 1026 is the limit Excel publishes
 * for horizontal page breaks.
 */
FM_API fm_status_t fm_sheet_add_row_break(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t manual);

/**
 * @brief Upserts a column break before `col`.
 *
 * Whole-sheet perpendicular span is `min = 0`, `max = 1048575`.
 * Upsert-replaces-in-place and the ascending-order guarantee match
 * `fm_sheet_add_row_break`.
 *
 * Returns `kInvalidArgument` (2) for a column past the grid, and
 * `kPreconditionFailed` (9) once the axis holds 1026 breaks. A column
 * break divides the page vertically, and Excel publishes no separate
 * figure for vertical page breaks; the horizontal limit is applied here
 * as a conservative bound rather than leaving the axis unbounded, so
 * this ceiling is Formulon's, not a documented Excel one.
 */
FM_API fm_status_t fm_sheet_add_col_break(fm_workbook_t* wb, size_t sheet_index, uint32_t col, int32_t manual);

/** @brief Removes the row break at `row`. Absent is a successful no-op. */
FM_API fm_status_t fm_sheet_remove_row_break(fm_workbook_t* wb, size_t sheet_index, uint32_t row);

/** @brief Removes the column break at `col`. Absent is a successful no-op. */
FM_API fm_status_t fm_sheet_remove_col_break(fm_workbook_t* wb, size_t sheet_index, uint32_t col);

/** @brief Removes every manual break on both axes. */
FM_API fm_status_t fm_sheet_clear_breaks(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Number of row breaks declared on the sheet, or zero on any
 *        argument error.
 *
 * The return type carries no status, so zero is ambiguous: it is both
 * "this axis holds no breaks" and "`wb` is `NULL` or `sheet_index` is
 * past the last sheet". The two are told apart by the thread-local
 * diagnostic, which this call clears on entry and populates only on
 * rejection - `fm_last_error_message()` is the empty string after a
 * successful count and non-empty after a rejected one.
 *
 * A caller enumerating breaks as count-then-`fm_sheet_row_break_at` must
 * check that message after the count. The `_at` loop is what would
 * otherwise surface `kInvalidArgument`, and on a zero count it never
 * runs, so a bad sheet index reads back as an empty break list.
 */
FM_API size_t fm_sheet_row_break_count(const fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Copies the row break at `index` (ascending by row).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `out` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` or `index` is out of
 *           range.
 */
FM_API fm_status_t fm_sheet_row_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                         fm_page_break_t* out);

/**
 * @brief Number of column breaks declared on the sheet, or zero on any
 *        argument error. Same zero-is-ambiguous contract as
 *        `fm_sheet_row_break_count`.
 */
FM_API size_t fm_sheet_col_break_count(const fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Copies the column break at `index` (ascending by column).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `out` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` or `index` is out of
 *           range.
 */
FM_API fm_status_t fm_sheet_col_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                         fm_page_break_t* out);

/* -------------------------------------------------------------------------- */
/* Print settings - typed patch setters                                       */
/* -------------------------------------------------------------------------- */

#define FM_ORIENTATION_DEFAULT 0u
#define FM_ORIENTATION_PORTRAIT 1u
#define FM_ORIENTATION_LANDSCAPE 2u

/**
 * @brief Partial `<pageSetup>` update.
 *
 * Each field is applied only when its `_engaged` companion is non-zero, so
 * a caller states the two or three attributes it cares about and leaves
 * everything else - including attributes the engine does not model, such as
 * `r:id`, `horizontalDpi` and `copies` - exactly as the file had them.
 *
 * On the read path `_engaged` means "the attribute is present in the XML";
 * the value field always carries the effective value, so a caller can tell
 * an explicit setting from an inherited default.
 */
typedef struct {
  int32_t orientation_engaged;
  uint32_t orientation; /* FM_ORIENTATION_* */
  int32_t paper_size_engaged;
  uint32_t paper_size; /* OOXML paperSize code; 9 = A4 */
  int32_t scale_engaged;
  uint32_t scale; /* percent, [10, 400] */
  int32_t fit_to_width_engaged;
  uint32_t fit_to_width;
  int32_t fit_to_height_engaged;
  uint32_t fit_to_height;
  int32_t fit_to_page_engaged;
  int32_t fit_to_page; /* routed to <sheetPr><pageSetUpPr>, not <pageSetup> */
} fm_page_setup_t;

/** @brief Partial `<pageMargins>` update, in inches. Same `_engaged`
 *         semantics as `fm_page_setup_t`. */
typedef struct {
  int32_t left_engaged;
  double left;
  int32_t right_engaged;
  double right;
  int32_t top_engaged;
  double top;
  int32_t bottom_engaged;
  double bottom;
  int32_t header_engaged;
  double header;
  int32_t footer_engaged;
  double footer;
} fm_page_margins_t;

/** @brief Partial `<printOptions>` update. Same `_engaged` semantics. */
typedef struct {
  int32_t grid_lines_engaged;
  int32_t grid_lines;
  int32_t headings_engaged;
  int32_t headings;
  int32_t horizontal_centered_engaged;
  int32_t horizontal_centered;
  int32_t vertical_centered_engaged;
  int32_t vertical_centered;
} fm_print_options_t;

/**
 * @brief Partial `<headerFooter>` update.
 *
 * Each section pointer is tri-state: NULL leaves that section untouched, an
 * empty string clears it, and any other value replaces it.
 *
 * Section text is the decoded string, so Excel's formatting codes are
 * written plainly: `"&C&\"MS Gothic\"Report &P/&N"`. The engine applies XML
 * escaping on the way out, which is what turns the leading `&` of each code
 * into the `&amp;` Excel itself stores.
 */
typedef struct {
  const char* odd_header;
  const char* odd_footer;
  const char* even_header;
  const char* even_footer;
  const char* first_header;
  const char* first_footer;
  int32_t different_odd_even_engaged;
  int32_t different_odd_even;
  int32_t different_first_engaged;
  int32_t different_first;
  int32_t scale_with_doc_engaged;
  int32_t scale_with_doc;
  int32_t align_with_margins_engaged;
  int32_t align_with_margins;
} fm_header_footer_t;

/**
 * @brief Applies a partial `<pageSetup>` update.
 *
 * `scale` outside [10, 400] is rejected with `kInvalidArgument` (2) rather
 * than clamped. `fm_sheet_set_zoom` clamps its own input because an
 * out-of-range screen zoom is harmless, but a mis-stated print scale lands
 * on paper, so silently rounding it would hide the mistake.
 */
FM_API fm_status_t fm_sheet_set_page_setup(fm_workbook_t* wb, size_t sheet_index, const fm_page_setup_t* setup);

/** @brief Applies a partial `<pageMargins>` update. A negative, infinite or
 *         NaN margin is rejected with `kInvalidArgument` (2). */
FM_API fm_status_t fm_sheet_set_page_margins(fm_workbook_t* wb, size_t sheet_index, const fm_page_margins_t* margins);

/** @brief Applies a partial `<printOptions>` update. */
FM_API fm_status_t fm_sheet_set_print_options(fm_workbook_t* wb, size_t sheet_index, const fm_print_options_t* options);

/** @brief Applies a partial `<headerFooter>` update. */
FM_API fm_status_t fm_sheet_set_header_footer(fm_workbook_t* wb, size_t sheet_index, const fm_header_footer_t* hf);

/** @brief Reads the effective page setup. Every `_engaged` flag reports
 *         whether the attribute is stated in the XML. */
FM_API fm_status_t fm_sheet_get_page_setup(const fm_workbook_t* wb, size_t sheet_index, fm_page_setup_t* out);

/** @brief Reads the effective page margins, in inches. */
FM_API fm_status_t fm_sheet_get_page_margins(const fm_workbook_t* wb, size_t sheet_index, fm_page_margins_t* out);

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
/* Conditional formatting — mutation                                          */
/* -------------------------------------------------------------------------- */

/**
 * @brief Inclusive cell range used by the CF mutation API.
 *
 * Matches `formulon::cf::CFCellRange` (top-left + bottom-right corners,
 * both 0-based). Single-cell rules use `first_row == last_row` and
 * `first_col == last_col`.
 */
typedef struct {
  uint32_t first_row;
  uint32_t first_col;
  uint32_t last_row;
  uint32_t last_col;
} fm_cf_cell_range_t;

/**
 * @brief Conditional-format value object used by visual rule authoring.
 *
 * `type` is the `formulon::cf::CfvoType` ordinal:
 * `0=number`, `1=percent`, `2=percentile`, `3=min`, `4=max`,
 * `5=formula`, `6=autoMin`, `7=autoMax`. `value` is optional and is
 * copied by `fm_sheet_cf_add_rule`; pass `NULL` for valueless min/max
 * thresholds. `gte` maps to OOXML `gte` (`1` by default).
 */
typedef struct {
  uint8_t type;
  uint8_t _pad[3];
  int32_t gte; /* 0/1 */
  const char* value;
} fm_cfvo_t;

/**
 * @brief CF rule wire format used by `fm_sheet_cf_*` APIs.
 *
 * Wide POD covering both differential-format rules and visual rules.
 * `fm_sheet_cf_add_rule` deep-copies every pointer-backed payload into
 * the engine model. `fm_sheet_cf_get_at` returns a mixed view for the
 * selected rule: ids, sqref ranges, and ordinary formula/text strings are
 * model-backed; visual threshold arrays and their value strings are
 * read-scratch-backed and may be invalidated by the next successful
 * scratch-backed read on the same handle. A validation-rejected call does
 * not refresh the scratch storage.
 *
 * Active fields by `type`:
 *   - `Expression` (0): `formula1`.
 *   - `CellIs` (1): `op_engaged` + `op` + `formula1` (+ `formula2` for
 *     `Between` / `NotBetween`).
 *   - `Top10` (5): `rank_engaged` + `rank` + `percent` + `bottom`.
 *   - `AboveAverage` (6): `above_average` + `equal_average`
 *     + (optional) `std_dev_engaged` + `std_dev`.
 *   - `ContainsText` / `NotContainsText` / `BeginsWith` / `EndsWith`
 *     (7-10): `text`.
 *   - `ContainsBlanks` / `NotContainsBlanks` / `ContainsErrors` /
 *     `NotContainsErrors` (11-14): no extra payload.
 *   - `TimePeriod` (15): `time_period_engaged` + `time_period`.
 *   - `DuplicateValues` / `UniqueValues` (16-17): no extra payload.
 *   - `ColorScale` (2): `color_scale_thresholds` +
 *     `color_scale_colors` with matching counts of 2 or 3.
 *   - `DataBar` (3): `data_bar_min`, `data_bar_max`,
 *     `data_bar_fill`, optional display flags, and the `x14` extension
 *     block at the end of this struct (axis, gradient, negative and
 *     border colours), each governed by its own `*_engaged` flag.
 *   - `IconSet` (4): `icon_set_name` plus N-1
 *     `icon_set_thresholds`.
 *
 * String fields use C-string convention: `NULL` means "absent",
 * non-`NULL` is a NUL-terminated view. On the input path
 * (`fm_sheet_cf_add_rule`) the caller owns the buffer until the call
 * returns; on the output path (`fm_sheet_cf_get_at`) model-backed fields
 * remain valid until the next CF-list mutation or handle destruction, while
 * read-scratch-backed visual fields may be invalidated by the next successful
 * scratch-backed read on the same handle as well. A validation-rejected call
 * does not refresh the scratch storage.
 */
typedef struct {
  /* Stable rule id, linking the rule to the `<x14:cfRule id="...">`
   * block that carries whatever settings the legacy OOXML schema cannot
   * express. On input, pass `NULL` or `""` to auto-generate one; a
   * supplied id must be GUID-shaped, since that is the schema type of
   * the attribute it becomes. On output, always a model-backed, non-NULL
   * view.
   *
   * Only rules that need such a block put their id in the saved file, so
   * for most rules this is an in-memory handle and nothing more. */
  const char* id;
  uint8_t type;        /* `formulon::cf::RuleType` ordinal */
  uint8_t op;          /* `formulon::cf::CellIsOperator` ordinal */
  uint8_t time_period; /* `formulon::cf::TimePeriod` ordinal */
  uint8_t _pad0;       /* alignment */
  int32_t priority;
  int32_t stop_if_true;   /* 0/1 */
  int32_t dxf_id_engaged; /* 0/1 */
  uint32_t dxf_id;

  /* sqref union — at least one entry. On input, must be non-NULL with
   * range_count >= 1. On output, points to the engine's model-backed vector
   * buffer for the containing block. */
  const fm_cf_cell_range_t* sqref;
  uint32_t sqref_count;

  /* Formula sources (Expression / CellIs / containsText-derived). */
  const char* formula1;
  const char* formula2;
  int32_t op_engaged; /* 0/1 (CellIs) */

  /* Top10 payload */
  int32_t rank_engaged; /* 0/1 */
  int32_t rank;
  int32_t percent; /* 0/1 */
  int32_t bottom;  /* 0/1 */

  /* AboveAverage payload */
  int32_t above_average;   /* 0/1 (default 1) */
  int32_t equal_average;   /* 0/1 */
  int32_t std_dev_engaged; /* 0/1 */
  double std_dev;

  /* ContainsText / BeginsWith / EndsWith / NotContainsText literal */
  const char* text;

  /* TimePeriod */
  int32_t time_period_engaged; /* 0/1 */

  /* ColorScale payload. Counts must match and be 2 or 3 on input. */
  const fm_cfvo_t* color_scale_thresholds;
  const fm_cf_color_t* color_scale_colors;
  uint32_t color_scale_count;

  /* DataBar payload. */
  int32_t data_bar_engaged; /* 0/1 */
  fm_cfvo_t data_bar_min;
  fm_cfvo_t data_bar_max;
  fm_cf_color_t data_bar_fill;
  int32_t data_bar_show_value; /* 0/1, default 1 */
  uint8_t data_bar_min_length_pct;
  uint8_t data_bar_max_length_pct;
  uint8_t _pad1[2];

  /* IconSet payload. */
  int32_t icon_set_engaged; /* 0/1 */
  uint8_t icon_set_name;    /* `formulon::cf::IconSetName` ordinal */
  uint8_t _pad2[3];
  const fm_cfvo_t* icon_set_thresholds;
  uint32_t icon_set_threshold_count;
  int32_t icon_set_reverse;    /* 0/1 */
  int32_t icon_set_show_value; /* 0/1, default 1 */
  /* `percent` attribute on `<iconSet>`, carried round-trip only: it is read,
   * exposed and re-emitted verbatim, and is never consulted during
   * evaluation. Each `fm_cfvo_t` names its own `type`, and that type is the
   * authoritative interpretation of the threshold, so setting this field
   * changes the saved XML but not the icon bucketing. To change the
   * interpretation, change the `fm_cfvo_t::type` values instead. */
  int32_t icon_set_percent; /* 0/1, default 1 */

  /* DataBar extension payload.
   *
   * The legacy `<dataBar>` element carries only min/max/fill/lengths; the
   * axis, gradient and negative colours live in the `x14` extension. They
   * are appended here rather than beside the other `data_bar_*` fields so
   * every field above keeps the offset it already had, and a consumer that
   * reads only the older prefix stays correct.
   *
   * Each `*_engaged` flag distinguishes "caller did not specify" from an
   * explicit value. A zero-initialized `fm_cf_rule_t` therefore reproduces
   * the same rule these fields produced before they existed: gradient on,
   * automatic axis, negative fill equal to the positive fill, no border,
   * and a black axis colour. Set the flag to override any of those.
   *
   * `fm_sheet_cf_get_at` engages all six, so feeding its output straight
   * back into `fm_sheet_cf_add_rule` reproduces the rule exactly. */
  int32_t data_bar_gradient_engaged;      /* 0/1 */
  int32_t data_bar_gradient;              /* 0/1; model default 1 (gradient fill) */
  int32_t data_bar_axis_position_engaged; /* 0/1 */
  uint8_t data_bar_axis_position;         /* 0=automatic (default), 1=middle, 2=none */
  uint8_t _pad3[3];
  int32_t data_bar_negative_fill_engaged; /* 0/1; when 0 the positive fill is reused */
  fm_cf_color_t data_bar_negative_fill;
  int32_t data_bar_border_engaged; /* 0/1; when 0 the bar has no border */
  fm_cf_color_t data_bar_border;
  int32_t data_bar_negative_border_engaged; /* 0/1; when 0 the bar has no negative border */
  fm_cf_color_t data_bar_negative_border;
  int32_t data_bar_axis_color_engaged; /* 0/1; when 0 the axis is black */
  fm_cf_color_t data_bar_axis_color;
} fm_cf_rule_t;

/**
 * @brief Returns the total number of CF rules on `sheet_index`.
 *
 * The count flattens across all `<conditionalFormatting>` blocks: a
 * sheet with two blocks of two rules each reports `4`. Indexing into
 * the flattened sequence is stable until the next mutation.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_cf_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `idx`-th CF rule (in flattened order) into `out`.
 *
 * Every pointer in `*out` - `id`, `sqref`, `formula1`, `formula2`, `text`,
 * the visual threshold and colour arrays, and the threshold value strings -
 * addresses storage owned by `wb`, not the workbook model. That storage is
 * released only by the next successful `fm_sheet_cf_get_at` on the same
 * handle, or by handle destruction. CF mutations
 * (`fm_sheet_cf_add_rule`, `fm_sheet_cf_remove_at`, `fm_sheet_cf_clear`),
 * unrelated reads, and save / load calls all leave the record readable, so
 * the edit procedure this API provides - get a rule, remove it, add the
 * amended record back - is safe to execute without copying any field out
 * first. A validation-rejected call does not refresh the storage, so the
 * previously read rule stays valid. Copy any field that must survive the
 * next get.
 *
 * `out->dxf_id` is reported engaged only when it resolves against the
 * workbook's `<dxfs>` table. A rule whose stored index reaches past that
 * table - which a third-party package can carry - reads back with
 * `dxf_id_engaged == 0` rather than surfacing an index
 * `fm_styles_get_dxf` would reject and `fm_sheet_cf_add_rule` would
 * refuse. The record therefore always feeds straight back into
 * `fm_sheet_cf_add_rule`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` or `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_cf_get_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, fm_cf_rule_t* out);

/**
 * @brief Appends a new single-rule `<conditionalFormatting>` block.
 *
 * The new block carries `rule.sqref` (deep-copied) and a single CF
 * rule constructed from the remaining fields. The auto-generated
 * priority (when `rule.priority <= 0`) is `existing_max + 1`. If
 * `rule.id` is `NULL` or empty, a new id is synthesised from the
 * priority.
 *
 * `*out_index` receives the new rule's position in the sheet's
 * flattened CF rule sequence (the same indexing `fm_sheet_cf_get_at`
 * and `fm_sheet_cf_remove_at` use). Since this call always appends a
 * new block after every existing one, the returned index equals the
 * flattened rule count observed just before the call; it stays valid
 * until a subsequent mutation (`fm_sheet_cf_add_rule`,
 * `fm_sheet_cf_remove_at`, `fm_sheet_cf_clear`) on the same sheet
 * renumbers the sequence.
 *
 * When `rule.dxf_id_engaged` is non-zero, `rule.dxf_id` must already name a
 * registered differential format (`fm_styles_add_dxf`). A rule referencing a
 * missing `<dxf>` is what makes Excel discard every conditional format in
 * the workbook on open, so the reference is validated here rather than at
 * save time.
 *
 * A `DataBar` payload is rejected when either length percentage exceeds
 * 100, when `data_bar_min_length_pct > data_bar_max_length_pct` (an
 * inverted bar Excel cannot render), or when an engaged
 * `data_bar_axis_position` is outside `0..2`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb` or `out_index` is `NULL`, or
 *           `rule.sqref` is `NULL` while `rule.sqref_count > 0`;
 *         `kInvalidArgument` when `sheet_index` is out of range, when
 *           `rule.sqref_count == 0` or exceeds the per-call range cap,
 *           when any `rule.sqref` entry is inverted (`first` past
 *           `last`) or reaches outside the grid (`row >= 1048576`,
 *           `col >= 16384`) and therefore has no A1 spelling, when
 *           `rule.id` is non-empty but not GUID-shaped, when `rule.type`
 *           / `rule.op` / `rule.time_period` / `rule.icon_set_name` /
 *           `rule.data_bar_axis_position` or any engaged `fm_cfvo_t`
 *           `type` names no declared enumerator, when `rule.dxf_id` is
 *           engaged but past the end of the workbook's `<dxfs>` table,
 *           or when a visual payload is missing / malformed.
 */
FM_API fm_status_t fm_sheet_cf_add_rule(fm_workbook_t* wb, size_t sheet_index, fm_cf_rule_t rule, size_t* out_index);

/**
 * @brief Removes the `idx`-th CF rule (flattened order). When the
 *        containing block becomes empty it is removed too.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` or `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_cf_remove_at(fm_workbook_t* wb, size_t sheet_index, size_t idx);

/**
 * @brief Removes every CF block on `sheet_index`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_cf_clear(fm_workbook_t* wb, size_t sheet_index);

/* -------------------------------------------------------------------------- */
/* Trace precedents / dependents (dependency-graph reverse lookup)            */
/* -------------------------------------------------------------------------- */

/**
 * @brief Workbook-wide cell coordinate used by the trace API.
 *
 * Mirrors `formulon::eval::CellNodeId`: `sheet` is the 0-based sheet
 * index, `row` and `col` are 0-based cell coordinates.
 */
typedef struct {
  uint32_t sheet;
  uint32_t row;
  uint32_t col;
} fm_cell_node_t;

/**
 * @brief Opaque handle for a trace result.
 *
 * Owns the cell-node list produced by `fm_workbook_precedents` /
 * `fm_workbook_dependents`. Index into it via `fm_cell_nodes_count` /
 * `fm_cell_nodes_at`. Release with `fm_cell_nodes_destroy`.
 */
typedef struct fm_cell_nodes fm_cell_nodes_t;

/**
 * @brief Returns the cells `(sheet, row, col)` directly reads
 *        (1-step precedents) when `depth <= 1`, or every cell reached
 *        within `depth` BFS steps when `depth >= 2`. `depth` is capped
 *        at 32 to avoid runaway expansion in cyclic graphs.
 *
 * The set excludes the seed cell itself. Order is unspecified but
 * stable for a given graph state.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb` or `out` is `NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 *
 * On success the caller owns `*out` and must free it with
 * `fm_cell_nodes_destroy`.
 */
FM_API fm_status_t fm_workbook_precedents(const fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                          uint32_t depth, fm_cell_nodes_t** out);

/**
 * @brief Returns the cells that read `(sheet, row, col)` directly
 *        (1-step dependents) when `depth <= 1`, or every cell reached
 *        within `depth` BFS steps in the reverse graph when
 *        `depth >= 2`. `depth` is capped at 32.
 *
 * Same semantics as `fm_workbook_precedents` for ownership and seed
 * exclusion.
 */
FM_API fm_status_t fm_workbook_dependents(const fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                          uint32_t depth, fm_cell_nodes_t** out);

/**
 * @brief Releases a trace results handle. `nodes == NULL` is a no-op.
 */
FM_API void fm_cell_nodes_destroy(fm_cell_nodes_t* nodes);

/**
 * @brief Returns the number of cells in the result. Returns `0` when
 *        `nodes == NULL`.
 */
FM_API size_t fm_cell_nodes_count(const fm_cell_nodes_t* nodes);

/**
 * @brief Reads cell `idx` from the result.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when any pointer argument is `NULL`;
 *         `kInvalidArgument` when `idx` is out of range.
 */
FM_API fm_status_t fm_cell_nodes_at(const fm_cell_nodes_t* nodes, size_t idx, fm_cell_node_t* out);

/* -------------------------------------------------------------------------- */
/* PivotTable layout projection                                               */
/* -------------------------------------------------------------------------- */

/**
 * @brief Pivot layout cell discriminator.
 *
 * Numbering mirrors `formulon::pivot::PivotCellKind`.
 */
typedef enum {
  FM_PIVOT_CELL_HEADER = 0,
  FM_PIVOT_CELL_ROW_LABEL = 1,
  FM_PIVOT_CELL_COL_LABEL = 2,
  FM_PIVOT_CELL_DATA = 3,
  FM_PIVOT_CELL_ROW_SUBTOTAL = 4,
  FM_PIVOT_CELL_COL_SUBTOTAL = 5,
  FM_PIVOT_CELL_GRAND_TOTAL = 6,
  FM_PIVOT_CELL_BLANK = 7
} fm_pivot_cell_kind_t;

/**
 * @brief One projected PivotTable cell.
 *
 * `row` / `col` are absolute 0-based sheet coordinates. `field_name` and
 * `number_format` are NUL-terminated UTF-8 strings owned by the containing
 * `fm_pivot_cells_t` handle and remain valid until that handle is destroyed.
 */
typedef struct {
  uint32_t row;
  uint32_t col;
  fm_value_t value;
  fm_pivot_cell_kind_t kind;
  uint32_t depth;
  const char* field_name;
  const char* number_format;
} fm_pivot_cell_t;

/**
 * @brief Opaque handle for a projected PivotTable layout.
 *
 * Owns the projected cell list and every string pointer reachable from
 * `fm_pivot_cell_t`. Index into it via `fm_pivot_cells_count` /
 * `fm_pivot_cells_at`. Release with `fm_pivot_cells_destroy`.
 */
typedef struct fm_pivot_cells fm_pivot_cells_t;

/**
 * @brief Returns the number of pivot tables anchored on `sheet_index`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb` or `out_count` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_pivot_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Evaluates and projects the `pivot_index`-th PivotTable on a sheet.
 *
 * On success the caller owns `*out` and must free it with
 * `fm_pivot_cells_destroy`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `wb` or `out` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` or `pivot_index` is out of range;
 *         a pivot evaluation/layout error when the table references a missing
 *         or invalid cache.
 */
FM_API fm_status_t fm_workbook_pivot_layout(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                            fm_pivot_cells_t** out);

/**
 * @brief Releases a pivot layout handle. `cells == NULL` is a no-op.
 */
FM_API void fm_pivot_cells_destroy(fm_pivot_cells_t* cells);

/**
 * @brief Returns the number of projected cells. Returns `0` when
 *        `cells == NULL`.
 */
FM_API size_t fm_pivot_cells_count(const fm_pivot_cells_t* cells);

/**
 * @brief Returns the projected layout bounds.
 *
 * `top` / `left` are absolute 0-based sheet coordinates; `rows` / `cols`
 * are the rectangular layout span.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_pivot_cells_bounds(const fm_pivot_cells_t* cells, uint32_t* out_top, uint32_t* out_left,
                                         uint32_t* out_rows, uint32_t* out_cols);

/**
 * @brief Reads projected cell `idx`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when any pointer argument is `NULL`;
 *         `kInvalidArgument` when `idx` is out of range.
 */
FM_API fm_status_t fm_pivot_cells_at(const fm_pivot_cells_t* cells, size_t idx, fm_pivot_cell_t* out);

/* -------------------------------------------------------------------------- */
/* PivotCache & PivotTable mutation                                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief PivotTable axis enumeration.
 *
 * Numbering mirrors `formulon::pivot::PivotAxis`.
 */
typedef enum {
  FM_PIVOT_AXIS_ROW = 0,
  FM_PIVOT_AXIS_COL = 1,
  FM_PIVOT_AXIS_VALUE = 2,
  FM_PIVOT_AXIS_PAGE = 3
} fm_pivot_axis_t;

/**
 * @brief Aggregation function for a value-axis field. Numbering mirrors
 *        `formulon::pivot::Aggregation` and `SubtotalFn` (the two enums are
 *        currently 1-to-1).
 */
typedef enum {
  FM_PIVOT_AGG_SUM = 0,
  FM_PIVOT_AGG_COUNT = 1,
  FM_PIVOT_AGG_AVERAGE = 2,
  FM_PIVOT_AGG_MAX = 3,
  FM_PIVOT_AGG_MIN = 4,
  FM_PIVOT_AGG_PRODUCT = 5,
  FM_PIVOT_AGG_COUNT_NUMBERS = 6,
  FM_PIVOT_AGG_STDDEV = 7,
  FM_PIVOT_AGG_STDDEVP = 8,
  FM_PIVOT_AGG_VAR = 9,
  FM_PIVOT_AGG_VARP = 10
} fm_pivot_aggregation_t;

/**
 * @brief Show-values-as derivation applied to a data-field aggregate.
 *        Numbering mirrors `formulon::pivot::ShowValuesAs`.
 */
typedef enum {
  FM_PIVOT_SHOW_AS_NORMAL = 0,
  FM_PIVOT_SHOW_AS_PERCENT_OF_ROW = 1,
  FM_PIVOT_SHOW_AS_PERCENT_OF_COL = 2,
  FM_PIVOT_SHOW_AS_PERCENT_OF_TOTAL = 3,
  FM_PIVOT_SHOW_AS_RUNNING_TOTAL_IN_ROW = 4,
  FM_PIVOT_SHOW_AS_RUNNING_TOTAL_IN_COL = 5,
  FM_PIVOT_SHOW_AS_INDEX = 6,
  FM_PIVOT_SHOW_AS_DIFFERENCE_FROM = 7,
  FM_PIVOT_SHOW_AS_PERCENT_DIFFERENCE_FROM = 8,
  FM_PIVOT_SHOW_AS_PERCENT_OF_PARENT_ROW = 9,
  FM_PIVOT_SHOW_AS_PERCENT_OF_PARENT_COL = 10,
  FM_PIVOT_SHOW_AS_PERCENT_OF_PARENT = 11
} fm_pivot_show_as_t;

/**
 * @brief Filter type for an active (slicer-applied) filter.
 *        Numbering mirrors `formulon::pivot::FilterType`.
 */
typedef enum {
  FM_PIVOT_FILTER_VALUE_TOP_10 = 0,
  FM_PIVOT_FILTER_VALUE_GREATER_THAN = 1,
  FM_PIVOT_FILTER_VALUE_BETWEEN = 2,
  FM_PIVOT_FILTER_LABEL_CONTAINS = 3,
  FM_PIVOT_FILTER_LABEL_BEGINS_WITH = 4,
  FM_PIVOT_FILTER_LABEL_DATE = 5
} fm_pivot_filter_type_t;

/**
 * @brief Date-grouping granularity. Numbering mirrors
 *        `formulon::pivot::DateGrouping`.
 */
typedef enum {
  FM_PIVOT_DATE_DAY = 0,
  FM_PIVOT_DATE_MONTH = 1,
  FM_PIVOT_DATE_QUARTER = 2,
  FM_PIVOT_DATE_YEAR = 3,
  FM_PIVOT_DATE_WEEK = 4,
  FM_PIVOT_DATE_HOUR = 5,
  FM_PIVOT_DATE_MINUTE = 6,
  FM_PIVOT_DATE_SECOND = 7
} fm_pivot_date_grouping_t;

/**
 * @brief Calendar system used by date grouping. Numbering mirrors
 *        `formulon::pivot::CalendarSystem`.
 */
typedef enum { FM_PIVOT_CALENDAR_GREGORIAN = 0, FM_PIVOT_CALENDAR_JAPANESE = 1 } fm_pivot_calendar_t;

/**
 * @brief Pivot report layout form. Numbering mirrors
 *        `formulon::pivot::PivotLayout`.
 */
typedef enum {
  FM_PIVOT_LAYOUT_COMPACT = 0,
  FM_PIVOT_LAYOUT_TABULAR = 1,
  FM_PIVOT_LAYOUT_OUTLINE = 2
} fm_pivot_layout_t;

/**
 * @brief Discriminator for the variant payload carried by a pivot filter
 *        spec. `FM_PIVOT_FILTER_VALUE_NONE` (= -1) means the slot is unset
 *        (only meaningful for the optional upper-bound payload on range
 *        filters); the int / double / text variants mirror
 *        `std::variant<int, double, std::string>`.
 */
typedef enum {
  FM_PIVOT_FILTER_VALUE_NONE = -1,
  FM_PIVOT_FILTER_VALUE_INT = 0,
  FM_PIVOT_FILTER_VALUE_DOUBLE = 1,
  FM_PIVOT_FILTER_VALUE_TEXT = 2
} fm_pivot_filter_value_kind_t;

/**
 * @brief Plain-data spec for `fm_workbook_pivot_field_add`.
 *
 *   * `source_name`   — required; matches a `PivotCacheField::name`.
 *   * `custom_name`   — nullable; pass `NULL` or the empty string for none.
 *   * `axis`          — initial axis for the field.
 *   * `subtotal_top`  — non-zero to render subtotals at the group head.
 *   * `number_format` — nullable; `NULL` means "leave blank".
 */
typedef struct {
  const char* source_name;
  const char* custom_name;
  fm_pivot_axis_t axis;
  int32_t subtotal_top;
  const char* number_format;
} fm_pivot_field_spec_t;

/**
 * @brief Plain-data spec for `fm_workbook_pivot_data_field_add` /
 *        `fm_workbook_pivot_data_field_set`.
 *
 *   * `name`          — required; display name (key for GETPIVOTDATA).
 *   * `field_index`   — index into `PivotTable::fields()` of the source
 *                       pivot field.
 *   * `aggregation`   — aggregation function applied to the source.
 *   * `number_format` — nullable.
 *   * `show_as`       — derivation mode.
 *   * `show_as_base_field` — `-1` means unset; otherwise an index into
 *                       `PivotTable::fields()`.
 *   * `show_as_base_item`  — `-1` means unset; `1048828` means "(previous)";
 *                       `1048829` means "(next)"; any other value is an
 *                       index into the base field's `items[]`.
 */
typedef struct {
  const char* name;
  uint32_t field_index;
  fm_pivot_aggregation_t aggregation;
  const char* number_format;
  fm_pivot_show_as_t show_as;
  int32_t show_as_base_field;
  int32_t show_as_base_item;
} fm_pivot_data_field_spec_t;

/**
 * @brief Plain-data spec for `fm_workbook_pivot_filter_add`.
 *
 *   * `axis`           — pivot axis the filter targets.
 *   * `field_name`     — required; the pivot field's name (or display name).
 *   * `type`           — filter type.
 *   * `value_kind`     — discriminator for the primary `value` payload.
 *                        For non-range filters this must be one of the
 *                        three concrete kinds (the `_NONE` sentinel is
 *                        rejected). For range filters (`VALUE_BETWEEN`,
 *                        `LABEL_DATE`) this is the lower bound.
 *   * `value_int`      — int payload; consulted iff `value_kind == INT`.
 *   * `value_double`   — double payload; consulted iff `value_kind == DOUBLE`.
 *   * `value_text`     — UTF-8 payload; consulted iff `value_kind == TEXT`.
 *                        Must be non-NULL when `value_kind == TEXT`.
 *   * `value_high_kind` — discriminator for the upper-bound payload.
 *                        `FM_PIVOT_FILTER_VALUE_NONE` for non-range filters.
 *                        For range filters, must be `INT` or `DOUBLE`
 *                        (text upper bounds are not modelled).
 *   * `value_high_int` / `value_high_double` — upper bound payload.
 *   * `data_field_index` — which data-field aggregate a value filter scores.
 *                        Validated against the pivot's data fields for the
 *                        value filter types; label/date filters ignore it.
 */
typedef struct {
  fm_pivot_axis_t axis;
  const char* field_name;
  fm_pivot_filter_type_t type;
  fm_pivot_filter_value_kind_t value_kind;
  int32_t value_int;
  double value_double;
  const char* value_text;
  fm_pivot_filter_value_kind_t value_high_kind;
  int32_t value_high_int;
  double value_high_double;
  uint32_t data_field_index;
} fm_pivot_filter_spec_t;

/* --- Pivot caches (workbook-owned) --------------------------------------- */

/**
 * @brief Returns the number of pivot caches owned by the workbook.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_pivot_cache_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Returns the `cache_id` of the pivot cache at flat index `idx`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_pivot_cache_id_at(const fm_workbook_t* wb, size_t idx, uint32_t* out_cache_id);

/**
 * @brief Creates a new empty pivot cache.
 *
 * `requested_id` may be `0` to request auto-assignment; the new id is
 * `max(existing_ids) + 1` (or `1` if no caches exist). When non-zero,
 * `requested_id` must not collide with an existing cache. The assigned
 * id is written to `*out_cache_id`.
 *
 * A new cache has no worksheet source, and a workbook containing one
 * cannot be saved: `fm_workbook_save` fails rather than write a package
 * Excel offers to repair on open. Call
 * `fm_workbook_pivot_cache_set_worksheet_source` before saving. The range
 * only has to name where the records came from; it need not hold data,
 * because the records carry the values themselves.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `requested_id` collides with an
 *         existing cache.
 */
FM_API fm_status_t fm_workbook_pivot_cache_create(fm_workbook_t* wb, uint32_t requested_id, uint32_t* out_cache_id);

/**
 * @brief Removes the pivot cache with id `cache_id`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when no cache matches `cache_id` or when
 *         the cache is still referenced by at least one pivot table.
 */
FM_API fm_status_t fm_workbook_pivot_cache_remove(fm_workbook_t* wb, uint32_t cache_id);

/**
 * @brief Reads the cache's worksheet source metadata.
 *
 * `out_present` is non-zero when a `<worksheetSource>` is present.
 * `out_ref`, `out_sheet`, and `out_name` are model-backed strings owned by
 * the workbook handle and remain valid until the cache is mutated or the
 * handle is destroyed. Reads do not invalidate them.
 */
FM_API fm_status_t fm_workbook_pivot_cache_get_worksheet_source(const fm_workbook_t* wb, uint32_t cache_id,
                                                                int32_t* out_present, const char** out_ref,
                                                                const char** out_sheet, const char** out_name);

/**
 * @brief Sets or clears the cache's worksheet source metadata.
 *
 * Pass `present == 0` to clear `<worksheetSource>`. When present is
 * non-zero, nullable string attributes are copied; `NULL` means absent.
 */
FM_API fm_status_t fm_workbook_pivot_cache_set_worksheet_source(fm_workbook_t* wb, uint32_t cache_id, int32_t present,
                                                                const char* ref, const char* sheet, const char* name);

/** @brief Number of fields on the cache identified by `cache_id`. */
FM_API fm_status_t fm_workbook_pivot_cache_field_count(const fm_workbook_t* wb, uint32_t cache_id, size_t* out_count);

/**
 * @brief Reads the name of the cache field at `field_idx`. The pointer is a
 *        model-backed view into the workbook handle and remains valid until
 *        the cache field list is mutated or the handle is destroyed. Reads
 *        do not invalidate it.
 */
FM_API fm_status_t fm_workbook_pivot_cache_field_name(const fm_workbook_t* wb, uint32_t cache_id, size_t field_idx,
                                                      const char** out_utf8);

/**
 * @brief Appends a new field with the given UTF-8 name to the cache.
 *        `out_field_idx` receives the new field's index. `utf8_name`
 *        must be non-NULL.
 */
FM_API fm_status_t fm_workbook_pivot_cache_field_add(fm_workbook_t* wb, uint32_t cache_id, const char* utf8_name,
                                                     size_t* out_field_idx);

/** @brief Drops every field (and every record) from the cache. */
FM_API fm_status_t fm_workbook_pivot_cache_field_clear(fm_workbook_t* wb, uint32_t cache_id);

/** @brief Number of shared items configured on cache field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_cache_field_shared_item_count(const fm_workbook_t* wb, uint32_t cache_id,
                                                                   size_t field_idx, size_t* out_count);

/** @brief Appends a numeric shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_cache_field_add_shared_item_number(fm_workbook_t* wb, uint32_t cache_id,
                                                                        size_t field_idx, double value);

/** @brief Appends a text shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_cache_field_add_shared_item_text(fm_workbook_t* wb, uint32_t cache_id,
                                                                      size_t field_idx, const char* utf8);

/** @brief Appends a boolean shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_cache_field_add_shared_item_bool(fm_workbook_t* wb, uint32_t cache_id,
                                                                      size_t field_idx, int32_t value);

/** @brief Appends a blank shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_cache_field_add_shared_item_blank(fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_idx);

/** @brief Appends an Excel error shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_cache_field_add_shared_item_error(fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_idx, fm_error_code_t error);

/** @brief Drops every shared item from cache field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_cache_field_clear_shared_items(fm_workbook_t* wb, uint32_t cache_id,
                                                                    size_t field_idx);

/** @brief Returns the number of records (rows) on the cache. */
FM_API fm_status_t fm_workbook_pivot_cache_record_count(const fm_workbook_t* wb, uint32_t cache_id, size_t* out_count);

/**
 * @brief Appends a new empty record to the cache. `out_record_idx`
 *        receives the new record's index. The record's cell vector is
 *        empty; populate it via `fm_workbook_pivot_cache_record_set_*`.
 */
FM_API fm_status_t fm_workbook_pivot_cache_record_add(fm_workbook_t* wb, uint32_t cache_id, size_t* out_record_idx);

/** @brief Drops every record from the cache. */
FM_API fm_status_t fm_workbook_pivot_cache_record_clear(fm_workbook_t* wb, uint32_t cache_id);

/**
 * @brief Sets cell `(record_idx, field_idx)` to a numeric value. The
 *        record's cell vector auto-extends to `field_idx + 1` with
 *        Blank fillers when shorter.
 */
FM_API fm_status_t fm_workbook_pivot_cache_record_set_number(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                             size_t field_idx, double value);

/** @brief Sets cell `(record_idx, field_idx)` to a UTF-8 text value. */
FM_API fm_status_t fm_workbook_pivot_cache_record_set_text(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                           size_t field_idx, const char* utf8);

/** @brief Sets cell `(record_idx, field_idx)` to a boolean value. */
FM_API fm_status_t fm_workbook_pivot_cache_record_set_bool(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                           size_t field_idx, int32_t value);

/** @brief Sets cell `(record_idx, field_idx)` to Blank. */
FM_API fm_status_t fm_workbook_pivot_cache_record_set_blank(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                            size_t field_idx);

/** @brief Sets cell `(record_idx, field_idx)` to an Excel error value. */
FM_API fm_status_t fm_workbook_pivot_cache_record_set_error(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                            size_t field_idx, fm_error_code_t error);

/* --- Pivot tables (sheet-owned) ------------------------------------------ */

/**
 * @brief Creates a new empty pivot table on `sheet_index`. `cache_id`
 *        must reference an existing cache. The new table's flat index
 *        is written to `*out_pivot_index`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb`, `utf8_name`, or `out_pivot_index`
 *         is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range or
 *         `cache_id` does not match any existing cache.
 */
FM_API fm_status_t fm_workbook_pivot_create(fm_workbook_t* wb, size_t sheet_index, const char* utf8_name,
                                            uint32_t cache_id, uint32_t anchor_row, uint32_t anchor_col,
                                            size_t* out_pivot_index);

/** @brief Removes the pivot table at `pivot_index` from `sheet_index`. */
FM_API fm_status_t fm_workbook_pivot_remove(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/** @brief Renames the pivot table. `utf8_name` must be non-NULL. */
FM_API fm_status_t fm_workbook_pivot_set_name(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                              const char* utf8_name);

/** @brief Updates the pivot's anchor cell and span. */
FM_API fm_status_t fm_workbook_pivot_set_anchor(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                uint32_t anchor_row, uint32_t anchor_col, uint32_t span_rows,
                                                uint32_t span_cols);

/** @brief Toggles the row / column grand total bands on the pivot. */
FM_API fm_status_t fm_workbook_pivot_set_grand_totals(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      int32_t rows_enabled, int32_t cols_enabled);

/** @brief Reads the pivot's compact / tabular / outline report layout. */
FM_API fm_status_t fm_workbook_pivot_get_layout(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                fm_pivot_layout_t* out_layout);

/**
 * @brief Sets the pivot's compact / tabular / outline report layout.
 *
 * `layout` is a raw signed 32-bit ordinal. Unknown values return
 * `kInvalidArgument`; callers do not need to construct an out-of-domain
 * `fm_pivot_layout_t` value.
 */
FM_API fm_status_t fm_workbook_pivot_set_layout(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                int32_t layout);

/** @brief Number of fields configured on the pivot. */
FM_API fm_status_t fm_workbook_pivot_field_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 size_t* out_count);

/**
 * @brief Appends a new field to the pivot. `spec->source_name` must be
 *        non-NULL; the other string fields are nullable. `out_field_idx`
 *        receives the new field's index.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb`, `spec`, `out_field_idx` or
 *           `spec->source_name` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` or `pivot_index` is out
 *           of range, or `spec->axis` names no declared enumerator.
 */
FM_API fm_status_t fm_workbook_pivot_field_add(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                               const fm_pivot_field_spec_t* spec, size_t* out_field_idx);

/** @brief Drops every field from the pivot. */
FM_API fm_status_t fm_workbook_pivot_field_clear(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Sets the axis of pivot field `field_idx`.
 *
 * `axis` is a raw signed 32-bit ordinal; unknown values return
 * `kInvalidArgument`.
 */
FM_API fm_status_t fm_workbook_pivot_field_set_axis(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    size_t field_idx, int32_t axis);

/**
 * @brief Sets the sort directive on pivot field `field_idx`. Pass
 *        `by_field == NULL` to clear the by-field key.
 */
FM_API fm_status_t fm_workbook_pivot_field_set_sort(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    size_t field_idx, int32_t ascending, const char* by_field);

/** @brief Sets the `subtotal_top` flag on pivot field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_field_set_subtotal_top(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                            size_t field_idx, int32_t top);

/**
 * @brief Appends an aggregation to pivot field `field_idx`. Only
 *        meaningful for value-axis fields. `agg` is a raw signed 32-bit
 *        ordinal; unknown values return `kInvalidArgument`.
 */
FM_API fm_status_t fm_workbook_pivot_field_add_aggregation(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                           size_t field_idx, int32_t agg);

/** @brief Drops every aggregation from pivot field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_field_clear_aggregations(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                              size_t field_idx);

/**
 * @brief Appends a manual-filter item to pivot field `field_idx`, addressed
 *        by its label.
 *        `utf8_name` must be non-NULL. `visible` is a 32-bit boolean.
 *
 * The item carries no cache binding, so the writer emits its document-order
 * position as `<item x="N">` and the filter engine matches source records by
 * comparing their rendered label against `utf8_name`. That makes an empty
 * `utf8_name` unusable as the blank item: the blank has no label of its own
 * and is identified by what it binds to. Build one with
 * `fm_workbook_pivot_field_add_item_at` instead.
 */
FM_API fm_status_t fm_workbook_pivot_field_add_item(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    size_t field_idx, const char* utf8_name, int32_t visible);

/**
 * @brief Appends a manual-filter item to pivot field `field_idx`, addressed
 *        by its position in the bound cache field's shared items.
 *
 * `cache_index` selects `shared_items[cache_index]` of the cache field the
 * pivot field is bound to — the same index space as OOXML `<item x="N">`,
 * which the writer re-emits verbatim. The item's label is resolved from
 * that shared item, so this is the form that can express the blank item:
 * a shared item that renders to nothing yields an item with no label,
 * which the filter engine then matches by its binding rather than by text.
 *
 * `cache_index` is not validated here because the cache field may be
 * populated after the pivot definition; an item whose index resolves to
 * nothing keeps an empty label and filters nothing.
 *
 * `visible` is a 32-bit boolean.
 */
FM_API fm_status_t fm_workbook_pivot_field_add_item_at(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       size_t field_idx, uint32_t cache_index, int32_t visible);

/** @brief Drops every manual-filter item from pivot field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_field_clear_items(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       size_t field_idx);

/** @brief Toggles the visibility of item `item_idx` on field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_field_set_item_visible(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                            size_t field_idx, size_t item_idx, int32_t visible);

/**
 * @brief Appends a subtotal-fn entry to pivot field `field_idx`.
 *
 * `agg` is a raw signed 32-bit ordinal; unknown values return
 * `kInvalidArgument`.
 */
FM_API fm_status_t fm_workbook_pivot_field_add_subtotal_fn(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                           size_t field_idx, int32_t agg);

/** @brief Drops every subtotal-fn entry from pivot field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_field_clear_subtotal_fns(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                              size_t field_idx);

/**
 * @brief Configures date-grouping on pivot field `field_idx`. Pass
 *        `start_year_or_neg1 == -1` (and likewise `end_year_or_neg1`)
 *        to leave the bound unset. `granularity` and `calendar` are raw
 *        signed 32-bit ordinals; unknown values return `kInvalidArgument`.
 */
FM_API fm_status_t fm_workbook_pivot_field_set_date_group(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                          size_t field_idx, int32_t granularity, int32_t calendar,
                                                          int32_t start_year_or_neg1, int32_t end_year_or_neg1);

/** @brief Removes the date-grouping config from pivot field `field_idx`. */
FM_API fm_status_t fm_workbook_pivot_field_clear_date_group(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                            size_t field_idx);

/**
 * @brief Sets the OOXML number-format string on pivot field `field_idx`.
 *        `utf8` must be non-NULL (use the empty string to clear).
 */
FM_API fm_status_t fm_workbook_pivot_field_set_number_format(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                             size_t field_idx, const char* utf8);

/**
 * @brief Replaces the row-axis field order with `indices[0..count)`.
 *        Each entry must be `< field_count`. Pass `count == 0` to clear.
 */
FM_API fm_status_t fm_workbook_pivot_set_row_field_order(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                         const uint32_t* indices, size_t count);

/** @brief Replaces the column-axis field order. Same contract as rows. */
FM_API fm_status_t fm_workbook_pivot_set_col_field_order(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                         const uint32_t* indices, size_t count);

/** @brief Number of `<dataField>` entries on the pivot. */
FM_API fm_status_t fm_workbook_pivot_data_field_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      size_t* out_count);

/**
 * @brief Appends a new data-field entry. `spec->name` must be non-NULL.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb`, `spec`, `out_idx` or
 *           `spec->name` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` or `pivot_index` is out
 *           of range, or `spec->aggregation` / `spec->show_as` names no
 *           declared enumerator.
 */
FM_API fm_status_t fm_workbook_pivot_data_field_add(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    const fm_pivot_data_field_spec_t* spec, size_t* out_idx);

/** @brief Drops every data-field entry from the pivot. */
FM_API fm_status_t fm_workbook_pivot_data_field_clear(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Replaces the data-field entry at `data_field_idx` in place.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb`, `spec` or `spec->name` is
 *           `NULL`;
 *         `kInvalidArgument` when `sheet_index`, `pivot_index` or
 *           `data_field_idx` is out of range, or `spec->aggregation` /
 *           `spec->show_as` names no declared enumerator.
 */
FM_API fm_status_t fm_workbook_pivot_data_field_set(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    size_t data_field_idx, const fm_pivot_data_field_spec_t* spec);

/*
 * Session-state contract shared by the whole pivot-filter family
 * (`fm_workbook_pivot_filter_count`, `_at`, `_add`, `_clear`,
 * `_remove_at`):
 *
 * This filter list is session state. Entries added here shape evaluation for
 * as long as the workbook handle lives and are not written to the package by
 * `fm_workbook_save*`. Conversely a `<filters>` block authored by Excel is
 * preserved verbatim on save but does not appear here, so
 * `fm_workbook_pivot_filter_count` reports only what this session added. The
 * persisted filter surface is pivot-field item visibility
 * (`fm_workbook_pivot_field_set_item_visible`).
 */

/** @brief Number of active (slicer-applied) filters on the pivot.
 *
 * Counts only filters this session added; see the session-state contract
 * above. */
FM_API fm_status_t fm_workbook_pivot_filter_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                  size_t* out_count);

/**
 * @brief Reads the active filter at `filter_idx` without mutating the pivot.
 *
 * Reaches only filters this session added; see the session-state contract
 * above.
 *
 * `out` reuses the `fm_pivot_filter_spec_t` shape accepted by
 * `fm_workbook_pivot_filter_add`. `field_name` and (for a text primary
 * value) `value_text` are model-backed views; they remain valid until the
 * next mutation of this pivot's filter list, a workbook model replacement,
 * or handle destruction. Reads do not invalidate the views. Copy them when
 * they must outlive those events. For non-text primary values,
 * `value_text` is `NULL`; an unset upper bound has kind
 * `FM_PIVOT_FILTER_VALUE_NONE` and zero numeric payloads.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out` is `NULL`;
 *         `kInvalidArgument` if `sheet_index`, `pivot_index`, or
 *         `filter_idx` is out of range, or the model contains an axis or
 *         filter type that is not representable by this C ABI. On any error,
 *         `*out` is left unchanged.
 */
FM_API fm_status_t fm_workbook_pivot_filter_at(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                               size_t filter_idx, fm_pivot_filter_spec_t* out);

/**
 * @brief Appends an active filter. `spec->field_name` must be non-NULL;
 *        `spec->value_text` must be non-NULL when
 *        `spec->value_kind == FM_PIVOT_FILTER_VALUE_TEXT`. The optional
 *        upper-bound payload is honoured only for range filter types
 *        (`VALUE_BETWEEN`, `LABEL_DATE`).
 *
 * The appended entry is session state and is not saved; see the
 * session-state contract above. To persist a filtered view, drive
 * pivot-field item visibility instead.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb`, `spec`, `spec->field_name` or a
 *         required `spec->value_text` is `NULL`;
 *         `kInvalidArgument` if `axis` or `type` is outside the C ABI enum
 *         ranges, the sheet/pivot indices do not resolve, the value payload
 *         discriminators are unusable, `data_field_index` is out of range
 *         for a value filter, or a label/date filter's `field_name` names no
 *         pivot field.
 */
FM_API fm_status_t fm_workbook_pivot_filter_add(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                const fm_pivot_filter_spec_t* spec);

/** @brief Drops every active filter from the pivot.
 *
 * Affects only filters this session added; a `<filters>` block authored by
 * Excel is untouched and still round-trips. See the session-state contract
 * above. */
FM_API fm_status_t fm_workbook_pivot_filter_clear(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/** @brief Removes the active filter at `filter_idx`.
 *
 * Indexes only filters this session added; see the session-state contract
 * above. */
FM_API fm_status_t fm_workbook_pivot_filter_remove_at(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      size_t filter_idx);

/* -------------------------------------------------------------------------- */
/* Dynamic-array spill payload                                                */
/* -------------------------------------------------------------------------- */

/**
 * @brief Result shape for `fm_workbook_spill_info`.
 *
 *   * `engaged` — `1` when `(row, col)` is part of a registered spill
 *     region, `0` otherwise. The other fields carry meaning only when
 *     `engaged == 1`.
 *   * `anchor_row`, `anchor_col` — the region's anchor cell (the cell
 *     that holds the dynamic-array formula).
 *   * `rows`, `cols` — region dimensions including the anchor.
 *
 * Use `fm_workbook_get_value(sheet, anchor_row + r, anchor_col + c)`
 * to read individual cells; `fm_workbook_get_value` is already
 * spill-aware and returns the per-cell value verbatim.
 */
typedef struct {
  uint32_t anchor_row;
  uint32_t anchor_col;
  uint32_t rows;
  uint32_t cols;
  int32_t engaged; /* 0/1 */
} fm_spill_info_t;

/**
 * @brief Returns the spill-region info for `(sheet, row, col)`.
 *
 * If `(row, col)` is the anchor of a region, returns that region. If
 * it is a phantom of a region, returns the same region. Otherwise sets
 * `out->engaged = 0` and leaves the remaining fields zero.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet` is out of range.
 */
FM_API fm_status_t fm_workbook_spill_info(const fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                          fm_spill_info_t* out);

/* -------------------------------------------------------------------------- */
/* Function catalog metadata                                                  */
/* -------------------------------------------------------------------------- */

/**
 * @brief Coarse availability class for a catalogued function.
 *
 * The value distinguishes "the parser/registry recognizes this function"
 * from "Formulon has a real Excel-compatible implementation." Consumers
 * should surface `FM_FUNCTION_UNAVAILABLE_STUB` clearly in autocomplete
 * and migration reports instead of treating it as implemented parity.
 *
 * `FM_FUNCTION_IMPLEMENTED_UNVERIFIED` is a reserved ordinal that the
 * engine never returns: every catalogued function is classified by the
 * table in `parts/function_catalog.cpp`, which uses only the other three.
 * A consumer switching over this enum still has to accept it so the
 * switch stays total.
 */
typedef enum {
  FM_FUNCTION_IMPLEMENTED = 0,
  FM_FUNCTION_IMPLEMENTED_UNVERIFIED = 1,
  FM_FUNCTION_ENVIRONMENT_BOUND = 2,
  FM_FUNCTION_UNAVAILABLE_STUB = 3
} fm_function_availability_t;

/**
 * @brief Result shape for `fm_function_metadata`.
 *
 * `canonical_name` is always populated when the function is known.
 * `min_arity` / `max_arity` are pulled from `FunctionDef`; the latter
 * is `0xFFFFFFFFu` (i.e. `eval::kVariadic`) for unbounded variadics.
 * Lazy-dispatch forms (e.g. `XLOOKUP`, `SUMIFS`) and parser special
 * forms (`LET`, `LAMBDA`) carry no `FunctionDef`, so their arity is
 * reported as `min_arity = 0` with the `0xFFFFFFFFu` `max_arity`
 * sentinel (unknown / unbounded).
 * `availability` reports whether the function is a real implementation,
 * host/environment-bound, or an intentionally unavailable fixed-error
 * stub (see `fm_function_availability_t`).
 *
 * `description` and `signature_template` are **always `NULL`**. They are
 * host-injected display metadata, not engine-owned data: the engine ships
 * no human-readable function text at all, and expects a host (editor /
 * docs surface) to supply its own document and merge it over this
 * structural result at display time. Merging over a `NULL` is the whole
 * contract — a host never has to handle the engine winning that
 * precedence. The document contract lives in
 * `docs/function-metadata-schema.md`; the native Node and Python bindings
 * ship pure merge helpers (`mergeFunctionMetadata` /
 * `merge_function_metadata`) for it. This
 * metadata is display-only: formula input parsing stays fixed to the
 * English canonical names, so `fm_function_localize` /
 * `fm_function_canonicalize` are unaffected by any injected document and
 * remain canonical-fallback (formula-language input localization is a
 * separate, engine-level concern outside this seam).
 *
 * String storage is process-static (the catalog is initialised at
 * static-init time); callers do not free the returned pointers.
 */
typedef struct {
  const char* canonical_name;
  uint32_t min_arity;
  uint32_t max_arity;
  fm_function_availability_t availability;
  /* Always `NULL` - supplied by the host's provider document. */
  const char* signature_template;
  /* Always `NULL` - supplied by the host's provider document. */
  const char* description;
} fm_function_metadata_t;

/**
 * @brief Locale codes for `fm_function_metadata`.
 *
 *   * `0` — `en-US` (default).
 *   * `1` — `ja-JP`.
 */
typedef enum { FM_LOCALE_EN_US = 0, FM_LOCALE_JA_JP = 1 } fm_locale_t;

/**
 * @brief Returns metadata for the function `name` in `locale`.
 *
 * `name` is matched case-insensitively against the canonical name. On
 * success, `*out` receives a view into the catalog's static storage;
 * the views are valid for the lifetime of the process. `locale` is a raw
 * signed 32-bit ordinal; unknown values return `kInvalidArgument`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `name` or `out` is `NULL`;
 *         `kInvalidArgument` when no function matches `name`.
 */
FM_API fm_status_t fm_function_metadata(const char* name, int32_t locale, fm_function_metadata_t* out);

/**
 * @brief Returns the total number of registered Formulon functions.
 *
 * Drives the canonical-name iterator (`fm_function_name_at`).
 */
FM_API size_t fm_function_count(void);

/**
 * @brief Returns the canonical name of the `idx`-th registered function.
 *
 * Order is sorted ascending so consumers can build deterministic UI
 * lists. `*out_name` is a static view into the catalog's process-static
 * storage.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when `out_name` is `NULL`;
 *         `kInvalidArgument` when `idx` is out of range.
 */
FM_API fm_status_t fm_function_name_at(size_t idx, const char** out_name);

/**
 * @brief Returns the localized display name for `canonical_name` in
 *        `locale`, or `canonical_name` itself when no alias is
 *        registered.
 *
 * `*out_localized` is a static view into process-static storage and must not
 * be freed. No alias table exists for any locale, so this always falls
 * through to `canonical_name` — for both supported locales, and by design
 * rather than pending work. The primary locale is `ja-JP`, whose function
 * names are identical to the English canonical names (the alias would be
 * the identity map); display names for any other locale belong to the
 * host's provider document, not to the engine. `locale` is still
 * validated: it is a raw signed 32-bit ordinal and unknown values return
 * `kInvalidArgument`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when any pointer argument is `NULL`;
 *         `kInvalidArgument` when `canonical_name` does not match a
 *           registered function.
 */
FM_API fm_status_t fm_function_localize(const char* canonical_name, int32_t locale, const char** out_localized);

/**
 * @brief Returns the canonical (English UPPERCASE) name for the
 *        localized function `localized_name` in `locale`.
 *
 * With no alias table in any locale (see `fm_function_localize`), this is
 * a case-insensitive ASCII match against the canonical name list, in
 * every locale. `locale` is still validated: it is a raw signed 32-bit
 * ordinal and unknown values return `kInvalidArgument`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` when any pointer argument is `NULL`;
 *         `kInvalidArgument` when no function matches.
 */
FM_API fm_status_t fm_function_canonicalize(const char* localized_name, int32_t locale, const char** out_canonical);

/* -------------------------------------------------------------------------- */
/* Sheet view / layout (viewport, frozen panes, column / row overrides)       */
/* -------------------------------------------------------------------------- */

/**
 * @brief Sheet tab visibility. Mirrors OOXML `<sheet state>` and the
 *        MS-XLSB `hsState` numbering, so neither container needs a table.
 *
 * `FM_SHEET_VERY_HIDDEN` differs from `FM_SHEET_HIDDEN` in that Excel
 * leaves such a sheet out of its "Unhide" dialog, which is how a workbook
 * keeps a settings or lookup sheet out of a user's reach.
 */
typedef enum {
  FM_SHEET_VISIBLE = 0,
  FM_SHEET_HIDDEN = 1,
  FM_SHEET_VERY_HIDDEN = 2,
} fm_sheet_visibility_t;

/**
 * @brief Per-sheet view state surfaced over the C ABI.
 *
 * Mirrors `formulon::SheetView` in full. Fields that are at their default
 * value still surface (zoom_scale defaults to `100`, freeze_rows /
 * freeze_cols default to `0`, tab_hidden defaults to `0`).
 *
 * `visibility` is the authoritative tab state and `tab_hidden` is the
 * two-state view of it, kept because it is the field every binding already
 * reads: `tab_hidden` is non-zero for both hidden states, so a caller that
 * knows only the bool still sees a very-hidden sheet as hidden rather than
 * as visible. Read `visibility` to tell the two apart. The pair is always
 * consistent on the way out; a struct is never handed back in, so the
 * contradictory combination cannot be constructed across this boundary.
 *
 * `view_mode` is a model-backed NUL-terminated UTF-8 view from the workbook
 * handle. It remains valid until a mutation of this sheet's view record
 * (including `fm_sheet_set_view_mode` or another sheet-view setter) or
 * handle destruction; reads, including scratch-backed reads, do not
 * invalidate it. It is `""` for the OOXML-default "normal" view,
 * `"pageBreakPreview"`, or `"pageLayout"`.
 */
typedef struct {
  uint32_t zoom_scale;
  uint32_t freeze_rows;
  uint32_t freeze_cols;
  int32_t tab_hidden;           /* 0/1; non-zero for either hidden state */
  int32_t visibility;           /* fm_sheet_visibility_t */
  int32_t show_grid_lines;      /* 0/1; OOXML default 1 */
  int32_t show_row_col_headers; /* 0/1; OOXML default 1 */
  int32_t show_zeros;           /* 0/1; OOXML default 1 */
  int32_t right_to_left;        /* 0/1; OOXML default 0 */
  int32_t tab_selected;         /* 0/1; OOXML default 0 */
  const char* view_mode;        /* "", "pageBreakPreview", or "pageLayout" */
} fm_sheet_view_t;

/**
 * @brief Per-column layout override surfaced over the C ABI.
 *
 * Mirrors `formulon::ColumnLayout`. Both endpoints are 0-based and
 * inclusive (matches the engine-side type; OOXML's `min`/`max`
 * 1-based conversion stays inside the writer). `has_width` is logical
 * presence: it is true for the raw presence bit or a legacy non-zero width,
 * while an explicit zero is preserved by the raw bit. `has_style` preserves
 * style attribute presence, including explicit style="0".
 */
typedef struct {
  uint32_t first;
  uint32_t last;
  double width;
  int32_t hidden; /* 0/1 */
  uint8_t outline_level;
  uint8_t _pad[3];
  /* Attribute-presence flags. has_width distinguishes width="0" from an
   * omitted width and is normalized true for legacy non-zero widths;
   * has_style keeps an explicit style="0". */
  int32_t has_width; /* 0/1 */
  int32_t has_style; /* 0/1 */
  uint32_t style_xf;
} fm_column_layout_t;

/**
 * @brief Per-row layout override surfaced over the C ABI.
 *
 * Mirrors `formulon::RowLayout`. The row index is 0-based; height is
 * in points (matches OOXML `<row ht=...>`). `has_style` is set only when
 * OOXML `customFormat="1"` makes the row `s` attribute effective.
 */
typedef struct {
  uint32_t row;
  double height;
  int32_t hidden; /* 0/1 */
  uint8_t outline_level;
  uint8_t _pad[3];
  /* Row s= is effective only when customFormat=1 in OOXML. */
  int32_t has_style; /* 0/1 */
  uint32_t style_xf;
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
 * @brief Reads the sheet-view state (zoom, frozen panes, tab hidden, and
 *        the display / orientation flags).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_t* out);

/**
 * @brief Returns the complete worksheet-level `<autoFilter>` XML fragment.
 *
 * `*out_xml` is a read-scratch-backed view and may be invalidated by the
 * next successful scratch-backed read on the same handle; a
 * validation-rejected call does not refresh the storage. Mutation and
 * handle destruction also invalidate it. It is the empty string when the
 * sheet has no AutoFilter.
 * The fragment includes any filter-column criteria, sort state, and extension
 * payload carried by the OOXML reader so callers can preserve definitions
 * they do not otherwise interpret.
 *
 * @return `kOk` on success; `kBindingNullPointer` if `out_xml` is `NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_auto_filter_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/**
 * @brief Replaces the worksheet-level `<autoFilter>` XML fragment.
 *
 * Pass an empty string to remove the AutoFilter. Non-empty input must be a
 * complete `<autoFilter ...>` fragment; detailed schema validation remains
 * with Excel/OOXML consumers because filter extensions are intentionally
 * preserved verbatim by the engine.
 *
 * @return `kOk` on success; `kBindingNullPointer` if `xml` is `NULL`;
 *         `kInvalidArgument` for a bad sheet index or malformed fragment.
 */
FM_API fm_status_t fm_sheet_set_auto_filter_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/**
 * @brief Sets or replaces the column width override for the inclusive
 *        column span `[first, last]`. Existing overrides whose span
 *        intersects the requested span are merged so the new width
 *        takes precedence on overlapping columns; non-overlapping
 *        portions of pre-existing entries are retained.
 *
 * `width` is in Excel's character-width unit. A rejected call leaves the
 * layout exactly as it was.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range,
 *         `last < first`, `last` is past the last column (16383), or
 *         `width` is negative, NaN, or infinite.
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
 *         `kInvalidArgument` when `sheet_index` is out of range,
 *         `last < first`, or `last` is past the last column (16383).
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
 *         `kInvalidArgument` when `sheet_index` is out of range,
 *         `last < first`, or `last` is past the last column (16383).
 */
FM_API fm_status_t fm_sheet_set_column_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                               uint8_t level);

/**
 * @brief Sets or replaces the row height override for `row`.
 *
 * `height` is in points. A rejected call leaves the layout exactly as it
 * was.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range, `row`
 *         is past the last row (1048575), or `height` is negative, NaN,
 *         or infinite.
 */
FM_API fm_status_t fm_sheet_set_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height);

/**
 * @brief Sets or replaces the row hidden flag for `row`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range or `row`
 *         is past the last row (1048575).
 */
FM_API fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t hidden);

/**
 * @brief Sets or replaces the row outline level for `row`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range or `row`
 *         is past the last row (1048575).
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
 * This is the two-state view of `fm_sheet_set_visibility`, kept for callers
 * that only ever mean "hidden" or "not hidden". Because the bool cannot
 * express `FM_SHEET_VERY_HIDDEN`, the two interact asymmetrically: passing
 * non-zero to an already very-hidden sheet leaves it very-hidden rather
 * than demoting it, since asking for "hidden" says nothing the sheet does
 * not already satisfy, while zero makes the sheet visible from either
 * hidden state. A very-hidden sheet loaded from a file therefore keeps that
 * visibility on save unless this call shows it. To state the stronger
 * visibility, or to demote very-hidden to plain hidden, call
 * `fm_sheet_set_visibility`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_tab_hidden(fm_workbook_t* wb, size_t sheet_index, int32_t hidden);

/**
 * @brief Sets the sheet's tab visibility to one of the three states.
 *
 * The only way to newly state `FM_SHEET_VERY_HIDDEN`, and the only way to
 * demote a very-hidden sheet to plain hidden; `fm_sheet_set_tab_hidden`
 * can express neither. Each call states the whole visibility, so the
 * resulting `fm_sheet_view_t` always reports a `tab_hidden` that agrees
 * with `visibility`.
 *
 * `visibility` is a raw signed 32-bit ordinal, not a validated enum type:
 * a value outside `fm_sheet_visibility_t` returns `kInvalidArgument` and
 * leaves the sheet unchanged.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range or
 *           `visibility` is not an `fm_sheet_visibility_t` value.
 */
FM_API fm_status_t fm_sheet_set_visibility(fm_workbook_t* wb, size_t sheet_index, int32_t visibility);

/**
 * @brief Sets the sheet's gridline-visibility flag (`showGridLines`).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_show_grid_lines(fm_workbook_t* wb, size_t sheet_index, int32_t show);

/**
 * @brief Sets the sheet's row/column header-visibility flag
 *        (`showRowColHeaders`).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_show_row_col_headers(fm_workbook_t* wb, size_t sheet_index, int32_t show);

/**
 * @brief Sets the sheet's zero-value display flag (`showZeros`).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_show_zeros(fm_workbook_t* wb, size_t sheet_index, int32_t show);

/**
 * @brief Sets the sheet's right-to-left display flag (`rightToLeft`).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_right_to_left(fm_workbook_t* wb, size_t sheet_index, int32_t right_to_left);

/**
 * @brief Sets the sheet's tab-selected flag (`tabSelected`). Plain
 *        metadata mirroring the OOXML attribute; setting this on more
 *        than one sheet does not affect evaluation, only what a host
 *        UI reproduces as the active tab.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_tab_selected(fm_workbook_t* wb, size_t sheet_index, int32_t selected);

/**
 * @brief Sets the sheet's view mode (`<sheetView view="...">`).
 *
 * `mode` is stored verbatim (no validation): pass `""` for the OOXML-
 * default "normal" view, `"pageBreakPreview"`, or `"pageLayout"`. Any
 * other value round-trips unchanged, matching the engine's tolerance
 * for future OOXML view-mode values.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL` or `mode == NULL`;
 *         `kInvalidArgument` when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_view_mode(fm_workbook_t* wb, size_t sheet_index, const char* mode);

/* -------------------------------------------------------------------------- */
/* Diagnostics                                                                */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the most recent error message produced by an API
 *        call on the current thread.
 *
 * The pointer is a thread-local diagnostic view and is valid until the next
 * API call on the same thread. Copy it before making another call when it
 * must be retained.
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

/**
 * @brief Returns the Excel display literal for a cell error code.
 *
 * Known values return literals such as `"#DIV/0!"` and `"#REF!"`.
 * Unknown numeric values return `"#UNKNOWN!"`. The returned pointer is
 * never `NULL` and has program lifetime.
 */
FM_API const char* fm_error_display_name(fm_error_code_t error);

/**
 * @brief Minimum severity for the engine's structured log stream.
 *
 * Ordinals match `formulon::StructuredLogLevel`. `FM_LOG_LEVEL_OFF`
 * discards every record and is the default: an embedded library must not
 * write to the host's stderr unless the host asks for it.
 */
typedef enum {
  FM_LOG_LEVEL_DEBUG = 0,
  FM_LOG_LEVEL_INFO = 1,
  FM_LOG_LEVEL_WARN = 2,
  FM_LOG_LEVEL_ERROR = 3,
  FM_LOG_LEVEL_OFF = 4,
} fm_log_level_t;

/**
 * @brief Receives one complete structured-log record.
 *
 * `record` is a single line of JSON terminated by `\n`, `len` bytes long.
 * The view is valid only for the duration of the call; copy it to retain
 * it. `user_data` is the opaque pointer registered with `fm_set_log_sink`.
 * The sink is invoked synchronously on whichever thread produced the
 * record, including engine worker threads during a parallel recalc, so it
 * must be thread-safe. It may reconfigure logging from inside the call.
 */
typedef void (*fm_log_sink_cb)(const char* record, size_t len, void* user_data);

/**
 * @brief Sets the lowest severity the engine emits, process-wide.
 *
 * The default is `FM_LOG_LEVEL_OFF`, so a caller that never invokes this
 * observes no engine output on stderr or through a sink. `level` is a raw
 * signed 32-bit ordinal so FFI callers can probe unknown values without
 * constructing an out-of-domain C enum.
 *
 * This setting is process-wide, not per workbook handle.
 *
 * @return `kOk` on success;
 *         `kInvalidArgument` when `level` is outside `fm_log_level_t`.
 */
FM_API fm_status_t fm_set_log_min_level(int32_t level);

/**
 * @brief Routes structured-log records to `sink` instead of stderr.
 *
 * Passing `sink == NULL` clears the sink and restores the stderr fallback,
 * which stays silent while the threshold is `FM_LOG_LEVEL_OFF`. `user_data`
 * is forwarded verbatim to every invocation; the engine does not take
 * ownership of it and the caller must keep it valid until the sink is
 * replaced or cleared.
 *
 * This setting is process-wide, not per workbook handle.
 *
 * @return `kOk` always.
 */
FM_API fm_status_t fm_set_log_sink(fm_log_sink_cb sink, void* user_data);

/* -------------------------------------------------------------------------- */
/* Styles                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Plain-data projection of a `formulon::io::CellXf` record.
 *
 * This record is a fixed 88-byte C ABI layout carrying every optional
 * alignment attribute represented by Formulon's style model.
 * `has_alignment` is the presence of the `<alignment>` child itself, so an
 * explicit empty child is distinct from an omitted child.
 * Each `has_*` flag below distinguishes an omitted OOXML attribute from an
 * explicit zero / false value. For the four core alignment attributes,
 * values are ignored when their corresponding flag is zero; insertion
 * canonicalizes them to the model defaults (`horizontal_align=0`,
 * `vertical_align=2` (bottom), `wrap_text=0`, and
 * `justify_last_line=0`). Thus a zero-initialized record means an omitted
 * alignment child and is deduplicated with the default XF; getters return
 * those canonical defaults. `relative_indent` is signed, matching the
 * OOXML `relativeIndent` attribute.
 *
 * Bindings copy the struct out of the workbook's styles table; subsequent
 * mutations to the workbook do not invalidate the copy.
 */
typedef struct {
  uint32_t font_index;
  uint32_t fill_index;
  uint32_t border_index;
  uint16_t num_fmt_id;
  uint8_t horizontal_align;
  uint8_t vertical_align;    /* 0=top, 1=center, 2=bottom (default), 3=justify, 4=distributed */
  int32_t wrap_text;         /* 0=false, 1=true */
  int32_t justify_last_line; /* 0=false, 1=true */
  uint32_t xf_id;            /* parent `<cellStyleXfs>` index for named styles */
  int32_t has_alignment;     /* 0= no `<alignment>` child, 1= child present */
  int32_t has_text_rotation;
  uint32_t text_rotation; /* 0..180, or 255 for vertical text */
  int32_t has_indent;
  uint32_t indent; /* 0..255 */
  int32_t has_relative_indent;
  int32_t relative_indent;
  int32_t has_shrink_to_fit;
  int32_t shrink_to_fit; /* 0=false, 1=true */
  int32_t has_reading_order;
  uint32_t reading_order;        /* 0=context, 1=LTR, 2=RTL */
  int32_t has_horizontal_align;  /* 0=omitted, 1=explicit `horizontal` */
  int32_t has_vertical_align;    /* 0=omitted, 1=explicit `vertical` */
  int32_t has_wrap_text;         /* 0=omitted, 1=explicit `wrapText` */
  int32_t has_justify_last_line; /* 0=omitted, 1=explicit `justifyLastLine` */
} fm_cell_xf;

/** Discriminator for `fm_color_spec::kind`, mirroring
 *  `formulon::io::ColorSpec::Kind`. */
typedef enum {
  kFmColorNone = 0,    /**< No selector; the writer emits sibling `*_argb` as `rgb`. */
  kFmColorRgb = 1,     /**< `rgb="AARRGGBB"`. */
  kFmColorTheme = 2,   /**< `theme="N"` with optional `tint`. */
  kFmColorIndexed = 3, /**< `indexed="N"` legacy palette index. */
  kFmColorAuto = 4,    /**< `auto="1"` system foreground / background. */
} fm_color_kind_t;

/**
 * @brief Plain-data projection of a `formulon::io::ColorSpec`.
 *
 * How the OOXML `<color>` element expressed its value. When `kind` is not
 * `kFmColorNone`, this selector is authoritative and the writer re-emits it;
 * the sibling `*_argb` field is not an Excel-rendered colour for theme,
 * indexed, or auto selectors. It is literal RGB for `kFmColorRgb`, or a
 * compatibility fallback otherwise. `kind == kFmColorNone` makes the
 * writer emit the sibling `*_argb` value as `rgb`.
 */
typedef struct {
  double tint;      /* theme tint (-1..1); meaningful when kind == kFmColorTheme */
  uint32_t rgb;     /* AARRGGBB; meaningful when kind == kFmColorRgb */
  uint32_t theme;   /* theme index; meaningful when kind == kFmColorTheme */
  uint32_t indexed; /* palette index; meaningful when kind == kFmColorIndexed */
  uint8_t kind;     /* fm_color_kind_t ordinal */
} fm_color_spec;

/**
 * @brief Plain-data projection of a `formulon::io::FontRecord`.
 *
 * `name` is a NUL-terminated UTF-8 model-backed view into the workbook's
 * styles table; it is valid until the next mutation that replaces the
 * styles table or until the handle is destroyed. Reads do not invalidate it.
 *
 * The `has_*` flags distinguish an absent OOXML element from an explicit
 * `val="0"`. They matter for a differential font, where an absent `<b>`
 * means "leave the source formatting unchanged" while `<b val="0"/>`
 * means "switch bold off"; on a cell font a `has_*` flag without its
 * value emits the explicit-off form.
 */
typedef struct {
  const char* name; /* NUL-terminated UTF-8, model-backed view */
  double size;
  /* Literal RGB for color.kind == kFmColorRgb; compatibility fallback for
   * theme/indexed/auto selectors (not an Excel-rendered colour). */
  uint32_t color_argb;
  int32_t bold;        /* 0=false, 1=true */
  int32_t italic;      /* 0=false, 1=true */
  int32_t strike;      /* 0=false, 1=true */
  int32_t has_bold;    /* 0= no `<b>` element, 1= element present */
  int32_t has_italic;  /* 0= no `<i>` element, 1= element present */
  int32_t has_strike;  /* 0= no `<strike>` element, 1= element present */
  int32_t has_family;  /* 0= no `<family>` element, 1= element present */
  int32_t has_charset; /* 0= no `<charset>` element, 1= element present */
  uint8_t underline;   /* 0=none, 1=single, 2=double, 3/4=accounting variants */
  uint8_t vert_align;  /* 0=baseline, 1=superscript, 2=subscript */
  uint8_t family;      /* `<family>` font-family class (0..5) */
  uint8_t charset;     /* `<charset>` codepage id (e.g. 128 = Shift_JIS) */
  fm_color_spec color;
} fm_font_record;

/**
 * @brief Plain-data projection of a `formulon::io::FillRecord`.
 *
 * `pattern` is the OOXML pattern index: `0=none`, `1=solid`,
 * `2..18=standard pattern set`. `fg_argb` and `bg_argb` are AARRGGBB
 * packed literal RGB or compatibility fallback values; `fg` and `bg` carry
 * the authoritative `<fgColor>` / `<bgColor>` specifications.
 */
typedef struct {
  uint8_t pattern; /* 0=none, 1=solid, 2..18=standard pattern set */
  /* Literal RGB when the matching spec kind is kFmColorRgb; compatibility
   * fallbacks otherwise (not Excel-rendered theme/indexed/auto colours). */
  uint32_t fg_argb;
  uint32_t bg_argb;
  fm_color_spec fg;
  fm_color_spec bg;
} fm_fill_record;

/**
 * @brief Plain-data projection of one side of a
 *        `formulon::io::BorderRecord`.
 *
 * `style` is the OOXML border-style ordinal: `0=none`, `1=thin`,
 * `2=medium`, `3=dashed`, ..., `13=slantDashDot`. `color` carries the
 * original `<color>` specification. `color_argb` is literal RGB for
 * `color.kind == kFmColorRgb`, or a compatibility fallback otherwise; it is
 * not a resolved theme/indexed/auto render colour.
 */
typedef struct {
  uint8_t style;       /* 0=none, 1=thin, ..., 13=slantDashDot */
  uint32_t color_argb; /* literal RGB or compatibility fallback, AARRGGBB */
  fm_color_spec color;
} fm_border_side;

/**
 * @brief Plain-data projection of a `formulon::io::BorderRecord`.
 *
 * Each side mirrors the OOXML schema. `diagonal_up` / `diagonal_down`
 * carry the boolean flags as `0`/`1` int32_t to match the rest of the
 * binding surface.
 */
typedef struct {
  fm_border_side left;
  fm_border_side right;
  fm_border_side top;
  fm_border_side bottom;
  fm_border_side diagonal;
  int32_t diagonal_up;   /* 0=false, 1=true */
  int32_t diagonal_down; /* 0=false, 1=true */
} fm_border_record;

/**
 * @brief One-shot style-table insertion request.
 *
 * Each non-empty input array requires its matching output-index array. The
 * call deduplicates each table in linear time across the existing records
 * and the supplied batch; fonts, fills and borders are installed before xfs
 * so an xf may refer to indices returned by the earlier arrays.
 *
 * `cell_xfs` entries carry the same presence-flag contract as
 * `fm_styles_add_cell_xf`: an alignment attribute is emitted only when its
 * `has_*` flag is set, and the value is ignored otherwise.
 */
typedef struct {
  const fm_font_record* fonts;
  size_t font_count;
  uint32_t* font_indices;
  const fm_fill_record* fills;
  size_t fill_count;
  uint32_t* fill_indices;
  const fm_border_record* borders;
  size_t border_count;
  uint32_t* border_indices;
  const fm_cell_xf* cell_xfs;
  size_t cell_xf_count;
  uint32_t* cell_xf_indices;
  const char* const* num_fmt_codes;
  size_t num_fmt_count;
  uint16_t* num_fmt_ids;
} fm_styles_batch;

/**
 * @brief Plain-data projection of one OOXML `<dxf>` differential format.
 *
 * Each `*_engaged` flag mirrors whether the corresponding child element
 * exists in the source `<dxf>`. `num_fmt_code`, `alignment_xml`, and
 * `protection_xml` are model-backed views, valid until the next styles-table
 * mutation (including `fm_styles_add_dxf`) or until the handle is destroyed;
 * reads do not invalidate them. The XML fields are never `NULL` and are empty
 * when their child is absent. Their semantic content is preserved without a
 * typed projection, but XML parsing may normalize lexical formatting such as
 * whitespace, quotes, and empty-element spelling.
 */
typedef struct {
  int32_t font_engaged; /* 0=false, 1=true */
  fm_font_record font;
  int32_t fill_engaged; /* 0=false, 1=true */
  fm_fill_record fill;
  int32_t border_engaged; /* 0=false, 1=true */
  fm_border_record border;
  int32_t num_fmt_engaged; /* 0=false, 1=true */
  uint16_t num_fmt_id;
  const char* num_fmt_code;   /* UTF-8, NUL-terminated; never NULL */
  const char* alignment_xml;  /* UTF-8 serialized `<alignment/>`; never NULL */
  const char* protection_xml; /* UTF-8 serialized `<protection/>`; never NULL */
} fm_dxf_record;

/** Sentinel for `fm_cell_style_record_t::builtin_id` indicating the
 *  style is custom (no `builtinId` attribute on the OOXML element). */
#define FM_CELL_STYLE_BUILTIN_ID_NONE 0xFFFFFFFFu

/**
 * @brief Plain-data projection of a `formulon::io::CellStyleRecord`
 *        (one OOXML `<cellStyle>` entry).
 *
 * `name` is a NUL-terminated UTF-8 model-backed view into the workbook's
 * styles table; it is valid until the next styles-replacing mutation or
 * until the handle is destroyed. `xf_id` indexes into the parallel
 * `<cellStyleXfs>` table (queryable via
 * `fm_styles_get_cell_style_xf_count` / `fm_styles_get_cell_style_xf`).
 *
 * `builtin_id` carries the OOXML built-in style ordinal (`0..47`); the
 * sentinel `FM_CELL_STYLE_BUILTIN_ID_NONE` indicates the entry is
 * custom and the attribute was absent on the source document. `i_level`
 * is the outline level for built-in heading styles (`0` for everything
 * else). The boolean flags follow the wide-POD convention used
 * elsewhere on this surface.
 */
typedef struct {
  const char* name;       /* UTF-8, NUL-terminated; never NULL */
  uint32_t xf_id;         /* index into cell_style_xfs */
  uint32_t builtin_id;    /* 0..47, or FM_CELL_STYLE_BUILTIN_ID_NONE */
  uint32_t i_level;       /* outline level for heading styles */
  int32_t hidden;         /* 0=false, 1=true */
  int32_t custom_builtin; /* 0=false, 1=true */
} fm_cell_style_record_t;

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
 * @brief Stores `xf_index` on every cell in the inclusive rectangle
 *        `[(first_row, first_col), (last_row, last_col)]`.
 *
 * Cells that do not exist yet are materialised as blanks carrying only the
 * style, which is what makes an empty ruled box render. Swapped corners are
 * normalised.
 *
 * Ruling a report means writing one xf across a rectangle; doing that a
 * cell at a time costs one ABI crossing per cell, which dominates the work
 * on any binding that marshals arguments.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb == NULL`;
 *         `kInvalidArgument` when `sheet` or a corner is out of range;
 *         `kPreconditionFailed` when the rectangle exceeds the per-call
 *         cell ceiling.
 */
FM_API fm_status_t fm_sheet_set_range_xf_index(fm_workbook_t* wb, uint32_t sheet, uint32_t first_row,
                                               uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                               uint32_t xf_index);

/**
 * @brief Reads the `xf_index`-th `<xf>` record from the workbook's
 *        styles table, including optional alignment attributes and their
 *        presence flags.
 *
 * `has_alignment` reports whether the `<alignment>` child existed;
 * explicit zero / false attributes remain present in their value fields.
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
 * @brief Looks up the effective format string for `num_fmt_id`.
 *
 * Resolution searches valid custom `<numFmt>` records in document order;
 * this includes a custom record below 164 that overrides a built-in slot.
 * When no valid custom record matches, a non-empty built-in code in the
 * `0..163` range is used as the fallback. On success `*out` is either a
 * model-backed view into the custom record's format string (valid until the
 * next styles-replacing mutation or until handle destruction), or a static
 * built-in view with program lifetime. The styles writer intentionally emits
 * only custom records with ids >= 164, so a below-164 override is not
 * promised to survive save/load. Returns `kInvalidArgument` when no effective
 * mapping exists for `num_fmt_id`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when the id is unknown.
 */
FM_API fm_status_t fm_styles_get_num_fmt_string(fm_workbook_t* wb, uint16_t num_fmt_id, const char** out);

/**
 * @brief Reads the `fill_index`-th fill record from the workbook's
 *        styles table.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `fill_index >= fills.size()`.
 */
FM_API fm_status_t fm_styles_get_fill(fm_workbook_t* wb, uint32_t fill_index, fm_fill_record* out);

/**
 * @brief Reads the `border_index`-th border record from the workbook's
 *        styles table.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `border_index >= borders.size()`.
 */
FM_API fm_status_t fm_styles_get_border(fm_workbook_t* wb, uint32_t border_index, fm_border_record* out);

/**
 * @brief Returns the number of `<dxf>` differential-format records
 *        available for conditional-format `dxfId` resolution.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_styles_get_dxf_count(fm_workbook_t* wb, uint32_t* out_count);

/**
 * @brief Reads the `dxf_index`-th differential-format record.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `dxf_index >= dxfs.size()`.
 */
FM_API fm_status_t fm_styles_get_dxf(fm_workbook_t* wb, uint32_t dxf_index, fm_dxf_record* out);

/**
 * @brief Adds a `<dxf>` differential-format record to the workbook's
 *        styles table, deduplicating against existing entries.
 *
 * Unlike `<xf>`, a dxf carries inline optional style fragments. Only
 * fields whose `*_engaged` flag is non-zero are copied. `num_fmt_code`
 * is copied when `num_fmt_engaged != 0`; passing `NULL` stores an empty
 * string. `alignment_xml` and `protection_xml` are always deep-copied when
 * non-`NULL`; passing `NULL` stores an empty fragment.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out_dxf_index` is `NULL`.
 */
FM_API fm_status_t fm_styles_add_dxf(fm_workbook_t* wb, fm_dxf_record record, uint32_t* out_dxf_index);

/**
 * @brief Returns the number of font records currently registered in the
 *        workbook's styles table.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_styles_get_font_count(fm_workbook_t* wb, uint32_t* out_count);

/**
 * @brief Returns the number of fill records currently registered in the
 *        workbook's styles table.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_styles_get_fill_count(fm_workbook_t* wb, uint32_t* out_count);

/**
 * @brief Returns the number of border records currently registered in
 *        the workbook's styles table.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_styles_get_border_count(fm_workbook_t* wb, uint32_t* out_count);

/**
 * @brief Returns the number of `<xf>` records currently registered in
 *        the workbook's styles table (i.e. the size of the `cellXfs`
 *        vector).
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_styles_get_cell_xf_count(fm_workbook_t* wb, uint32_t* out_count);

/**
 * @brief Returns the number of named cell styles (`<cellStyle>` entries)
 *        registered in the workbook. Zero for workbooks that do not
 *        declare any named styles.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_styles_get_cell_style_count(fm_workbook_t* wb, uint32_t* out_count);

/**
 * @brief Reads the `index`-th named cell style. The returned `name` is a
 *        model-backed view into storage owned by the workbook; see
 *        `fm_cell_style_record_t` for lifetime rules.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `index >= cell_styles.size()`.
 */
FM_API fm_status_t fm_styles_get_cell_style(fm_workbook_t* wb, uint32_t index, fm_cell_style_record_t* out);

/**
 * @brief Returns the number of `<cellStyleXfs>` records — the named-
 *        style xf table. This table is independent of `cellXfs` and
 *        is referenced by `fm_cell_style_record_t::xf_id`.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_styles_get_cell_style_xf_count(fm_workbook_t* wb, uint32_t* out_count);

/**
 * @brief Reads the `index`-th `<cellStyleXfs>` record (named-style xf
 *        table). Output shape mirrors `fm_styles_get_cell_xf`, including
 *        the optional alignment attributes and their presence flags.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `index >= cell_style_xfs.size()`.
 */
FM_API fm_status_t fm_styles_get_cell_style_xf(fm_workbook_t* wb, uint32_t index, fm_cell_xf* out);

/**
 * @brief Adds a font record to the workbook's styles table, deduplicating
 *        against existing entries.
 *
 * Linear-search dedup: returns the index of the first existing record the
 * styles writer cannot distinguish from `record` — that is, one which
 * serialises to the same `<font>` XML, presence flags and colour
 * specification included. A record that differs only in its `color`
 * specification, or only in a `has_*` flag that changes the emitted
 * element, is therefore a distinct record and gets its own index. When no
 * match exists the record is appended and the new (now-largest) index is
 * returned. The dedup is `O(N)` per call; callers that bulk-build a
 * workbook should batch when possible.
 *
 * `record.name` must be a NUL-terminated UTF-8 pointer; an empty /
 * `NULL` pointer is treated as the empty string. The string is copied
 * into the styles table and may be freed by the caller after the call
 * returns.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out_index` is `NULL`.
 */
FM_API fm_status_t fm_styles_add_font(fm_workbook_t* wb, fm_font_record record, uint32_t* out_index);

/**
 * @brief Adds a fill record to the workbook's styles table, deduplicating
 *        against existing entries.
 *
 * Linear-search dedup on the same writer-visible-state basis as
 * `fm_styles_add_font`: a record that differs only in its `fg` / `bg`
 * colour specification gets its own index. When no match exists the
 * record is appended and the new (now-largest) index is returned. The
 * dedup is `O(N)` per call; callers that bulk-build a workbook should
 * batch when possible.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out_index` is `NULL`.
 */
FM_API fm_status_t fm_styles_add_fill(fm_workbook_t* wb, fm_fill_record record, uint32_t* out_index);

/**
 * @brief Adds a border record to the workbook's styles table, deduplicating
 *        against existing entries.
 *
 * Linear-search dedup on the same writer-visible-state basis as
 * `fm_styles_add_font`: a record whose sides differ only in their colour
 * specification gets its own index. When no match exists the record is
 * appended and the new (now-largest) index is returned. The dedup is
 * `O(N)` per call; callers that bulk-build a workbook should batch when
 * possible.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out_index` is `NULL`.
 */
FM_API fm_status_t fm_styles_add_border(fm_workbook_t* wb, fm_border_record record, uint32_t* out_index);

/**
 * @brief Adds a number-format code to the workbook's styles table,
 *        deduplicating against existing entries. Returns the resolved
 *        OOXML `numFmtId`.
 *
 * Resolution proceeds in three steps. An id's effective mapping is the
 *      first valid custom `<numFmt>` record for that id in document order,
 *      falling back to a non-empty built-in code when no custom record
 *      resolves it.
 *   1. If `format_code` matches an id in the built-in range (`0..163`) after
 *      that effective-resolution step, the id is returned and nothing is
 *      appended to the table. A built-in id is therefore returned only when
 *      its current effective mapping matches `format_code`.
 *   2. Otherwise existing custom entries whose ids effectively resolve to
 *      their code are searched; the first match returns its registered id.
 *   3. Otherwise a new custom entry is appended with id
 *      `max(existing_custom_id, 163) + 1`, and the new id is returned.
 *      When the 16-bit OOXML id space is exhausted, the function returns
 *      `kPreconditionFailed` and leaves both the table and `*out_num_fmt_id`
 *      unchanged.
 *
 * `format_code` must be a NUL-terminated UTF-8 pointer; passing `NULL`
 * is treated as the empty string.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out_num_fmt_id` is `NULL`;
 *         `kPreconditionFailed` when no custom OOXML id remains.
 */
FM_API fm_status_t fm_styles_add_num_fmt(fm_workbook_t* wb, const char* format_code, uint16_t* out_num_fmt_id);

/**
 * @brief Adds an `<xf>` record to the workbook's styles table,
 *        deduplicating against existing entries.
 *
 * Linear-search dedup: returns the index of the first existing record
 * that is field-for-field equal to `record`. When no match exists the
 * record is appended and the new (now-largest) index is returned. The
 * dedup is `O(N)` per call; callers that bulk-build a workbook should
 * batch when possible.
 *
 * `has_alignment` preserves an explicit empty `<alignment/>`; the other
 * presence flags preserve explicit defaults.
 *
 * Validation:
 *   * `record.font_index` must satisfy `< font_count` (no auto-grow).
 *   * `record.fill_index` must satisfy `< fill_count`.
 *   * `record.border_index` must satisfy `< border_count`.
 *   * `record.num_fmt_id` must either be a documented built-in
 *     (`< 164` AND resolves through `builtin_num_fmt`) or already
 *     registered as a custom entry. Unknown ids surface
 *     `kInvalidArgument`.
 *   * Alignment values outside their Excel ranges are rejected when the
 *     matching presence flag is set.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if `wb` or `out_xf_index` is `NULL`;
 *         `kInvalidArgument` for any out-of-range / unregistered field.
 */
FM_API fm_status_t fm_styles_add_cell_xf(fm_workbook_t* wb, fm_cell_xf record, uint32_t* out_xf_index);

/** Adds and deduplicates a named-style xf including optional alignment
 * attributes. `has_alignment` preserves an explicit empty child. Invalid
 * Excel ranges return `kInvalidArgument`. */
FM_API fm_status_t fm_styles_add_cell_style_xf(fm_workbook_t* wb, fm_cell_xf record, uint32_t* out_xf_id);

/** Adds or replaces a named `<cellStyle>` record. `xf_id` must reference an
 * existing named-style xf. Pass `FM_CELL_STYLE_BUILTIN_ID_NONE` for custom styles. */
FM_API fm_status_t fm_styles_set_cell_style(fm_workbook_t* wb, const char* name, uint32_t xf_id, uint32_t builtin_id);

/**
 * @brief Adds and deduplicates a batch of fonts, fills, borders, and cell xfs.
 *
 * Empty arrays may be `NULL`. A non-empty input or output array that is NULL
 * returns `kBindingNullPointer`; invalid xfs return `kInvalidArgument`.
 * The operation is staged and committed as one transaction: any returned
 * error, including `kPreconditionFailed` when custom number-format ids are
 * exhausted, leaves the workbook and every caller-provided output array
 * unchanged. Allocation failures are not converted into a fabricated
 * success or recovery path; callers observe only the documented status
 * returns from validation / precondition checks.
 */
FM_API fm_status_t fm_styles_add_batch(fm_workbook_t* wb, const fm_styles_batch* batch);

/* -------------------------------------------------------------------------- */
/* External links                                                             */
/* -------------------------------------------------------------------------- */

/**
 * @brief Kind discriminator for `fm_external_link_record_t::kind`. Values
 *        mirror `formulon::io::ExternalLinkRecord::Kind`.
 */
#define FM_EXTERNAL_LINK_KIND_UNKNOWN 0u
#define FM_EXTERNAL_LINK_KIND_EXTERNAL_BOOK 1u
#define FM_EXTERNAL_LINK_KIND_OLE 2u
#define FM_EXTERNAL_LINK_KIND_DDE 3u

/**
 * @brief Plain-data projection of `formulon::io::ExternalLinkRecord`.
 *
 * `rel_id`, `part_path`, and `target` are NUL-terminated UTF-8
 * model-backed views into the workbook's external-links table; they remain
 * valid until the workbook is destroyed or the external-links table is
 * replaced (currently only the OOXML reader replaces it; there is no
 * mutator on this surface). `target_external` follows the wide-POD
 * convention used elsewhere on this surface.
 *
 * `index` is the 1-based position in `<externalReferences>` document
 * order. `kind` matches the `FM_EXTERNAL_LINK_KIND_*` constants above.
 * `target` is empty when the per-link rels file was missing or
 * unparseable; callers that need the URL to be present should check
 * `target[0] != '\0'` before using it.
 */
typedef struct {
  uint32_t index;          /* 1-based document order */
  const char* rel_id;      /* workbook.xml.rels rId for this link */
  const char* part_path;   /* resolved package-relative body path */
  const char* target;      /* remote workbook URL, or "" */
  int32_t target_external; /* TargetMode="External" => 1 */
  uint32_t kind;           /* one of FM_EXTERNAL_LINK_KIND_* */
} fm_external_link_record_t;

/**
 * @brief Returns the number of external-link records carried by the
 *        workbook.
 *
 * Always succeeds for a non-NULL workbook. Returns `0` for fresh
 * workbooks and any package whose source archive had no
 * `<externalReferences>` block.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_external_link_count(fm_workbook_t* wb, uint32_t* out_count);

/**
 * @brief Reads the `index`-th external-link record (1-based document
 *        order minus one — i.e. `0` is the first `<externalReference>`).
 *
 * The returned string pointers are model-backed views into workbook-owned
 * storage; see `fm_external_link_record_t` for lifetime rules.
 *
 * @return `kOk` on success;
 *         `kBindingNullPointer` if any pointer argument is `NULL`;
 *         `kInvalidArgument` when `index >= external_link_count`.
 */
FM_API fm_status_t fm_workbook_external_link_at(fm_workbook_t* wb, uint32_t index, fm_external_link_record_t* out);

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
