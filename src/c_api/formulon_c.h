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
 * `src/utils/error.h`); a return value of `0` (`FM_STATUS_OK` (0)) always denotes
 * success. Error strings and contexts associated with the most recent
 * call are exposed via `fm_get_last_error_message()` /
 * `fm_get_last_error_context()`. Both buffers are thread-local and are
 * overwritten on every API call on the same thread; copy them out
 * before invoking another API entry point if the caller wants to keep
 * the diagnostic.
 *
 * Text input: every caller-supplied string (argument or struct field) must be
 * well-formed UTF-8. A lone surrogate, overlong form or truncated sequence
 * returns `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) with the workbook unchanged; bindings
 * holding UTF-16 replace lone surrogates with U+FFFD before the call.
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
 *   * `fm_get_status_name()` and `fm_get_version()` are pure functions
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
 *   * Buffers returned through `fm_workbook_save`, `fm_workbook_save_as`,
 *     `fm_workbook_save_with_diagnostics`, `fm_sheet_snapshot_drawing` and
 *     `fm_sheet_snapshot_range_style` are heap-allocated by the library with
 *     `new[]` and must be released by passing the same pointer to
 *     `fm_buffer_destroy`. Mixing `malloc`/`free` and
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

#include "c_api/formulon_status.h"

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
 * `FM_STATUS_OK` (0) (success); every other value identifies a module band as
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
  int32_t kind; /* fm_value_kind_t */
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
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`.
 *
 * @thread-safety Not thread-safe with respect to its own `out`
 *                pointer. Two threads must not race to populate the
 *                same `out` slot.
 */
FM_API fm_status_t fm_workbook_create_default(fm_workbook_t** out);

/**
 * @brief Builds a workbook with the new-workbook defaults of an Excel profile.
 *
 * Like `fm_workbook_create_default` (one sheet), but the styles table, the first
 * sheet and the workbook's Excel profile follow the profile's locale: font 0
 * (name, charset, `scheme` = minor), the phonetic font (font 1, CJK locales
 * only), the vertical alignment of cell xf 0, the `Normal` cell style name,
 * the sheet name, `phoneticPr`, page margins and, on `mac-*`, the default
 * row height and base column width. No theme part is written; Excel treats
 * a workbook without one as the Office theme. `fm_workbook_create_default` is
 * unchanged and still seeds `Calibri` 11.
 *
 * Only the `mac-*` values come from Excel. The `win-*` ids are estimates: the
 * Mac locale table with an 11 pt body font, and no default row height or
 * base column width (the engine defaults apply).
 *
 * Changing font 0 afterwards (`fm_workbook_set_default_font`) leaves the
 * phonetic font (font 1) as it was.
 *
 * @param profile_id  An id from `fm_workbook_set_excel_profile_id`,
 *                    e.g. `"mac-365-ja_JP"`.
 * @param out         On success receives a freshly allocated handle owned by
 *                    the caller (release with `fm_workbook_destroy`).
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `profile_id` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an unknown profile id.
 */
FM_API fm_status_t fm_workbook_create_for_profile(const char* profile_id, fm_workbook_t** out);

/**
 * @brief Builds an empty workbook (no sheets).
 *
 * Designed for callers that intend to add sheets immediately via
 * `fm_workbook_add_sheet`. Saving an empty workbook fails just as
 * Excel rejects sheet-less archives.
 *
 * The only difference from `fm_workbook_create_default` is the missing sheet:
 * the style table is seeded with Excel's reserved defaults either way
 * (font 0, fill 0 = `none`, fill 1 = `gray125`, border 0, cell-style xf
 * 0, cell xf 0, the `Normal` cell style). A record added through
 * `fm_styles_add_*` that deduplicates against one of those reserved
 * slots -- including slot 0 -- returns that slot's own index rather than
 * a new one past it; see `fm_workbook_add_font`'s own doc for the dedup
 * rule.
 *
 * @param out  On success receives a freshly allocated handle.
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`.
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL` or
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
 * validation and stays all-zero on every non-`FM_STATUS_OK` (0) return, so a caller
 * may reuse one struct across calls without observing a previous call's
 * numbers. A workbook built in memory rather than loaded reports every
 * counter as zero.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if either pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_diagnostics(const fm_workbook_t* wb, fm_read_diagnostics_t* out_diagnostics);

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
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if either pointer is
 *         `NULL`.
 */
FM_API fm_status_t fm_workbook_get_memory_usage(const fm_workbook_t* wb, size_t* out_bytes);

/**
 * @brief Releases a workbook handle.
 *
 * `wb == NULL` is a no-op (mirrors `free()` semantics). Every text
 * pointer previously surfaced through `fm_sheet_get_value` /
 * `fm_workbook_get_sheet_name` is invalidated.
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
 *     non-`FM_STATUS_OK` (0) return. A caller may therefore reuse the same out variables
 *     across calls without observing a previous call's pointer, and the
 *     "initialise to NULL, call, free unconditionally" idiom stays safe
 *     because `fm_buffer_destroy(NULL)` is a no-op.
 *   * A NULL out-parameter is rejected with `FM_STATUS_BINDING_NULL_POINTER` (7001) and never
 *     dereferenced.
 *   * On `FM_STATUS_OK` (0), `*out_bytes` owns exactly `*out_len` bytes allocated with
 *     `new uint8_t[]` and must be released with `fm_buffer_destroy`.
 */

/**
 * @brief Serialises the workbook to an in-memory `.xlsx` byte stream.
 *
 * On success the caller receives a heap-allocated buffer that MUST be
 * released with `fm_buffer_destroy` (NOT `free`). This ownership rule applies
 * to every save API returning bytes -- `fm_workbook_save_as` (XLSX or
 * XLSB) and `fm_workbook_save_with_diagnostics` alike. Mixing allocators
 * across the boundary is undefined.
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param out_bytes  Receives a pointer to the freshly allocated buffer.
 * @param out_len    Receives the buffer length in bytes.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         a `kIo*` code on archive / writer failure.
 */
FM_API fm_status_t fm_workbook_save(const fm_workbook_t* wb, uint8_t** out_bytes, size_t* out_len);

/**
 * @brief Container format selector for `fm_workbook_save_as`.
 *
 * Mirrors `formulon::WorkbookFormat`. `FM_WORKBOOK_FORMAT_UNKNOWN`
 * is not a valid save target; passing it to `fm_workbook_save_as`
 * returns `FM_STATUS_INVALID_ARGUMENT` (2).
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
 * `FM_STATUS_INVALID_ARGUMENT` (2).
 *
 * On success the caller receives a heap-allocated buffer that MUST be
 * released with `fm_buffer_destroy` (NOT `free`). This applies to every save
 * API returning `out_bytes`: `fm_workbook_save`, `fm_workbook_save_as`,
 * and `fm_workbook_save_with_diagnostics`. Mixing allocators across the
 * boundary is undefined.
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param format     Target container format.
 * @param out_bytes  Receives a pointer to the freshly allocated buffer.
 * @param out_len    Receives the buffer length in bytes.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `format` is `FM_WORKBOOK_FORMAT_UNKNOWN`
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
   *  removed sheet) with `FM_STATUS_IO_WRITE_FAILED` (5016) before building any part, so
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
 * argument and model validation, and stays there on every non-`FM_STATUS_OK` (0)
 * return. `*out_diagnostics` is therefore always written, success or
 * failure; on failure it reads all-zero.
 *
 * The returned `out_bytes` buffer follows the `fm_buffer_destroy` ownership
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
FM_API void fm_buffer_destroy(uint8_t* bytes);

/* -------------------------------------------------------------------------- */
/* Sheets                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Writes the number of sheets in the workbook to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is NULL.
 */
FM_API fm_status_t fm_workbook_get_sheet_count(const fm_workbook_t* wb, size_t* out_count);

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_sheet_name(const fm_workbook_t* wb, size_t index, const char** out_utf8);

/**
 * @brief Appends a new sheet with the given UTF-8 display name.
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param utf8_name  Strict UTF-8 display name; must be non-NULL. Worksheet
 *                   identity uses locale-independent Unicode simple case
 *                   folding (C/S mappings only, without normalization).
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `utf8_name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_SHEET_NAME` (5051) when `utf8_name` is empty,
 *         exceeds 31 UTF-16 units, contains a forbidden character
 *         (`: \ / ? * [ ]`), begins or ends with an apostrophe, is
 *         `History` in any case, or collides under locale-independent
 *         Unicode simple case folding with an existing sheet -- the rules
 *         `fm_workbook_rename_sheet` applies;
 *         `FM_STATUS_SHEET_COUNT_LIMIT_EXCEEDED` (5053) when the workbook already contains
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when either index is out of range.
 */
FM_API fm_status_t fm_workbook_move_sheet(fm_workbook_t* wb, size_t from_index, size_t to_index);

/**
 * @brief Copies the worksheet at `source_index` into the workbook.
 *
 * The copy is placed before the sheet now at `insert_before` (`sheet_count`
 * appends) and its index is written to `*out_index`. References qualified
 * with the source sheet are re-pointed at the copy; tables are renamed and
 * rebound; sheet-scoped names, and workbook names referring to the source,
 * gain a copy scoped to the new sheet; pivot tables share the source's cache.
 * The active sheet and the tab selection do not change. See
 * `Workbook::copy_sheet` for the full contract. On any failure the workbook is
 * left unchanged.
 *
 * @param wb            Workbook handle. Must be non-NULL.
 * @param source_index  0-based index of the sheet to copy.
 * @param insert_before 0-based index of the sheet the copy goes before;
 *                      `sheet_count` appends.
 * @param new_name      Strict UTF-8 display name validated like
 *                      `fm_workbook_add_sheet`, or `NULL` to name the copy as
 *                      Excel does (`Src (2)`).
 * @param out_index     Receives the copy's index. Must be non-NULL.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_index` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `new_name` is malformed UTF-8;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `source_index >= sheet_count` or
 *         `insert_before > sheet_count`;
 *         `FM_STATUS_INVALID_SHEET_NAME` (5051) when `new_name` is empty, too long, contains a
 *         forbidden character or collides with an existing sheet;
 *         `FM_STATUS_SHEET_COUNT_LIMIT_EXCEEDED` (5053) when the workbook is full;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the source is very hidden;
 *         `FM_STATUS_NOT_IMPLEMENTED` (3) when the source is not a worksheet or carries
 *         content the copy cannot model (unrecognised relationships, slicers,
 *         shapes, pivot charts, or a drawing retained from an `.xlsb`).
 */
FM_API fm_status_t fm_workbook_copy_sheet(fm_workbook_t* wb, size_t source_index, size_t insert_before,
                                          const char* new_name, size_t* out_index);

/**
 * @brief Reads the active sheet: the first `<workbookView>`'s `activeTab`.
 *
 * A workbook with no `<bookViews>` (an XLSB load included) reports its first
 * `tabSelected` sheet, else `0`; an `activeTab` outside the sheet list
 * reports `0`. Inserting, removing and moving sheets keep the stored index
 * on the sheet it named; removing or hiding the active sheet activates the
 * first visible sheet after it, else the nearest visible one before it, as
 * Excel does.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_index` is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_active_sheet(const fm_workbook_t* wb, size_t* out_index);

/**
 * @brief Marks sheet `index` active so a save records it.
 *
 * Writes `activeTab` on the first `<workbookView>` and lowers its
 * `firstSheet` to `index` when that is larger; every other attribute is kept
 * verbatim. A workbook with no `<bookViews>` gets a minimal
 * `<bookViews><workbookView activeTab="N"/></bookViews>`. Sheet `index`
 * becomes the only `tabSelected` sheet, as clicking its tab does in Excel.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `index` is not a sheet of the workbook;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet is hidden or very hidden.
 */
FM_API fm_status_t fm_workbook_set_active_sheet(fm_workbook_t* wb, size_t index);

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
 * unregistered. Removing the active sheet activates its visible successor
 * (`fm_workbook_get_active_sheet`). A workbook keeps at least one visible
 * sheet, so the only visible sheet cannot be removed while hidden ones
 * remain.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `index` is out of range;
 *         `FM_STATUS_CANNOT_REMOVE_LAST_SHEET` (5052) when the workbook has only one
 *         sheet remaining;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is the only visible sheet
 *         (the workbook is left unchanged).
 */
FM_API fm_status_t fm_workbook_remove_sheet(fm_workbook_t* wb, size_t index);

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `new_name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_SHEET_NAME` (5051) when `new_name` is empty,
 *         exceeds 31 UTF-16 units, contains a forbidden character
 *         (`: \ / ? * [ ]`), begins or ends with an apostrophe, is
 *         `History` in any case, or collides under locale-independent
 *         Unicode simple case folding with an existing sheet. Folding is not
 *         normalization, full (expanding) case folding, locale tailoring, or
 *         a claim of full Excel-equivalence.
 */
FM_API fm_status_t fm_workbook_rename_sheet(fm_workbook_t* wb, size_t index, const char* new_name);

/**
 * @brief Sets (or appends, or removes) a workbook-scoped defined name.
 *
 * If a workbook-scoped entry with `name` already exists (case-
 * insensitive match), its formula text is replaced. Otherwise — when
 * `formula` is non-empty — a new entry is appended. Passing an empty
 * `formula` removes the existing entry, or is a no-op when no such
 * entry is present. Use `fm_workbook_set_defined_name_scoped` for
 * sheet-scoped names. A non-empty `formula` that does not parse is
 * rejected and leaves the name list unchanged. One leading `=` is
 * accepted and dropped: the name stores the formula as files carry it,
 * spelled as `fm_sheet_set_formula` stores a cell formula.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `name` is empty;
 *         `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000) when a non-empty `formula` does
 *         not parse.
 */
FM_API fm_status_t fm_workbook_set_defined_name(fm_workbook_t* wb, const char* name, const char* formula);

/**
 * @brief Sets (or appends, or removes) a defined name in a specific scope.
 *
 * `local_sheet_id == -1` selects workbook scope. Values `>= 0` select
 * the corresponding 0-based sheet-local scope. An empty `formula`
 * removes the matching entry in that scope. A non-empty `formula` that
 * does not parse is rejected and leaves the name list unchanged. One
 * leading `=` is accepted and dropped, as for `fm_workbook_set_defined_name`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `name` is empty or `local_sheet_id`
 *         is outside `[-1, sheet_count)`;
 *         `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000) when a non-empty `formula` does
 *         not parse.
 */
FM_API fm_status_t fm_workbook_set_defined_name_scoped(fm_workbook_t* wb, const char* name, const char* formula,
                                                       int32_t local_sheet_id);

/**
 * @brief Sets the `comment` and `hidden` attributes of an existing defined
 *        name.
 *
 * `local_sheet_id` selects the scope as in
 * `fm_workbook_set_defined_name_scoped`. A NULL `comment` leaves the comment
 * unchanged and an empty one clears it; a negative `hidden` leaves the flag
 * unchanged, otherwise non-zero hides the name from the Name Manager.
 * Replacing a name's formula keeps both attributes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when no such name exists in that scope.
 */
FM_API fm_status_t fm_workbook_set_defined_name_attrs(fm_workbook_t* wb, const char* name, int32_t local_sheet_id,
                                                      const char* comment, int32_t hidden);

/* -------------------------------------------------------------------------- */
/* Row / column structural edits                                              */
/* -------------------------------------------------------------------------- */

/*
 * The four edits below move every structure the engine models: cells,
 * formulas, merges, conditional-format ranges, data validations,
 * hyperlinks, tables, print areas, manual breaks and the auto-filter
 * range. The ranges and formulas inside the worksheet-level `<extLst>`
 * (the `x14` conditional-formatting block, sparkline groups, `x14`
 * validations) move with them, as Excel moves them, and so do the cells
 * named by the other worksheet content kept verbatim for a lossless save:
 * protected ranges, scenarios, a sheet sort state, cell watches,
 * ignored-error ranges and web-publish ranges.
 *
 * What they do not move is the anchor of an embedded OLE object or form
 * control, also kept verbatim: it keeps its pre-edit cell, and nothing
 * reports it. An insert also does not copy the formatting of the row or
 * column before it onto the new one, which Excel does -- a sparkline
 * included.
 */

/**
 * @brief Inserts `count` rows at `row` on `sheet_index`. Cells at `row` and
 *        beyond shift forward; cells pushed past the sheet bound are
 *        dropped. References across the workbook are rewritten to
 *        track the new positions; references that would land past the
 *        sheet bound collapse to `#REF!`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row` is
 *         past the sheet bound, or `count == 0`. Delete operations also
 *         reject a `count` that extends past the sheet bound;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the edit cuts into a data table,
 *         legacy array block or PivotTable, or would push one of them or a
 *         table past the sheet bound; the workbook is unchanged.
 */
FM_API fm_status_t fm_sheet_insert_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t count);

/**
 * @brief Deletes `count` rows starting at `row` on `sheet_index`. The deleted
 *        rows are dropped wholesale; subsequent rows shift back.
 *        References pointing inside the deleted interval collapse to
 *        `#REF!`; references past the deletion shift back.
 *
 * @return Same status codes as `fm_sheet_insert_rows`.
 */
FM_API fm_status_t fm_sheet_delete_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t count);

/**
 * @brief Inserts `count` columns at `col` on `sheet_index`. Mirrors
 *        `fm_sheet_insert_rows` along the column axis.
 */
FM_API fm_status_t fm_sheet_insert_cols(fm_workbook_t* wb, size_t sheet_index, uint32_t col, uint32_t count);

/**
 * @brief Deletes `count` columns starting at `col` on `sheet_index`. Mirrors
 *        `fm_sheet_delete_rows` along the column axis.
 */
FM_API fm_status_t fm_sheet_delete_cols(fm_workbook_t* wb, size_t sheet_index, uint32_t col, uint32_t count);

/* -------------------------------------------------------------------------- */
/* Cell mutation                                                              */
/* -------------------------------------------------------------------------- */

/*
 * The number / bool / error / text / formula setters store an entry the way
 * Excel does: a cell that has no style of its own takes its row's style
 * (`fm_sheet_set_row_style`), else its column's, as its cell XF. A cell that
 * already carries an XF, including an explicit XF 0, keeps it.
 *
 * Every setter here, `fm_sheet_set_blank` included, refuses a non-anchor
 * cell of a legacy array formula (CSE) block: it returns
 * `FM_STATUS_PRECONDITION_FAILED` (9) and leaves the block unchanged. A write to the
 * block's anchor replaces the whole block.
 */

/**
 * @brief Stores a numeric literal at `(row, col)` on the given sheet.
 *
 * Mirrors `Workbook::set_cell_value(... Value::number)`. Marks the
 * cell and any existing dependents dirty for the next recalc.
 *
 * `value` must be finite: a Number-kind cell always holds a finite
 * IEEE-754 double, matching the engine's own invariant that every
 * arithmetic result which would be NaN/Infinity becomes `#NUM!` before
 * it can be observed (see `eval/scalar_ops.cpp`) and that `save()`
 * downgrades a non-finite literal to a `#NUM!` error cell on write.
 * NaN or infinity is rejected rather than silently accepted and later
 * downgraded, so `ISNUMBER`, arithmetic on the cell, and a save/reload
 * round trip never disagree about whether it is a number or an error.
 *
 * A table's header cell takes the number in General format (`2024`) as text instead: the column is renamed
 * to it (`fm_workbook_rename_table_column`), and the cell holds the final
 * name.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `value` is NaN or infinite.
 */
FM_API fm_status_t fm_sheet_set_number(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col, double value);

/**
 * @brief Stores a boolean literal. `value` is interpreted as a C
 *        boolean: `0` is FALSE, any non-zero is TRUE.
 *
 * A table's header cell takes `TRUE` or `FALSE` as text instead: the column is renamed
 * to it (`fm_workbook_rename_table_column`), and the cell holds the final
 * name.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_bool(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col, int32_t value);

/**
 * @brief Stores a static Excel error literal at `(row, col)`.
 *
 * This writes a cell whose value is `#DIV/0!`, `#VALUE!`, etc. It is
 * distinct from writing a formula that evaluates to an error.
 *
 * A table's header cell takes the error's name (`#N/A`) as text instead: the column is renamed
 * to it (`fm_workbook_rename_table_column`), and the cell holds the final
 * name.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `error` is out of range.
 */
FM_API fm_status_t fm_sheet_set_error(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                      fm_error_code_t error);

/**
 * @brief Stores a text literal. The destination cell copies the UTF-8
 *        contents, so `utf8` does not need to outlive the call.
 *
 * A table's header cell takes `utf8` as text instead: the column is renamed
 * to it (`fm_workbook_rename_table_column`), and the cell holds the final
 * name.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_text(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                     const char* utf8);

/**
 * @brief One `<rPh>` block: the kana reading `text` covers the half-open
 *        span `[sb, eb)` of the cell's surface text.
 *
 * Offsets are UTF-16 code units, which is how Excel indexes string
 * positions everywhere else. A cell whose reading was typed in one go
 * carries the single run `{0, <length of the text>, kana}`.
 *
 * The spans are observable rather than presentation detail: `PHONETIC`
 * substitutes each annotated span with its kana and passes the text
 * outside every span through unchanged, so a partially annotated string
 * surfaces a mix of kana and original characters.
 */
typedef struct {
  uint32_t sb;      /* inclusive span start, in UTF-16 code units */
  uint32_t eb;      /* exclusive span end, in UTF-16 code units   */
  const char* text; /* NUL-terminated UTF-8 kana                  */
} fm_phonetic_run_t;

/**
 * @brief Stores a whole-cell phonetic guide (OOXML `<rPh>`).
 *
 * The guide is independent of the cell's visible text. Passing an empty
 * string clears an existing guide. The destination cell copies `utf8`, so
 * the input does not need to outlive the call.
 *
 * This entry point has no span vocabulary, so the reading it stores covers
 * the surface text end to end — one run spanning the whole cell. Feeding
 * `fm_sheet_get_cell_phonetic`'s result back through here therefore
 * flattens a partially annotated cell into a single whole-cell annotation.
 * Use `fm_sheet_set_cell_phonetic_runs` to preserve the spans.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_cell_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              const char* utf8);

/**
 * @brief Stores a cell's phonetic guide as one `<rPh>` block per run,
 *        spans included.
 *
 * `count == 0` clears the annotation (`runs` may be `NULL` in that case).
 * Each run's `text` is copied, so the array and its strings do not need to
 * outlive the call.
 *
 * The runs must describe a partition of the surface text in reading order:
 * every run needs `sb <= eb`, and each run must start at or after the
 * previous run's `eb`. Overlapping or out-of-order runs are rejected rather
 * than normalised, because `PHONETIC` walks them in order and a run whose
 * span has already been passed would silently move its kana onto characters
 * it does not read.
 *
 * `eb` is not checked against the cell's current text length. The surface
 * text may legitimately be written after the annotation, and every
 * value-mutating setter clears the annotation anyway, so a span that
 * overshoots is treated as covering the remainder of the text rather than
 * as an error.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, if `runs` is `NULL` with
 *         a non-zero `count`, or if any run's `text` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the cell coordinate is out of range, or when
 *         the runs are not an ordered partition.
 */
FM_API fm_status_t fm_sheet_set_cell_phonetic_runs(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                   const fm_phonetic_run_t* runs, size_t count);

/** `type` values for `fm_sheet_set_cell_phonetic_properties`. */
#define FM_PHONETIC_TYPE_HALFWIDTH_KATAKANA 0u
#define FM_PHONETIC_TYPE_FULLWIDTH_KATAKANA 1u
#define FM_PHONETIC_TYPE_HIRAGANA 2u
#define FM_PHONETIC_TYPE_NO_CONVERSION 3u

/** `alignment` values for `fm_sheet_set_cell_phonetic_properties`. */
#define FM_PHONETIC_ALIGNMENT_NO_CONTROL 0u
#define FM_PHONETIC_ALIGNMENT_LEFT 1u
#define FM_PHONETIC_ALIGNMENT_CENTER 2u
#define FM_PHONETIC_ALIGNMENT_DISTRIBUTED 3u

/**
 * @brief Sets the cell's "Show Phonetic" display flag (OOXML `<c ph="1">`,
 *        XLSB `fPhShow`).
 *
 * The flag only says whether the guide is displayed. It is independent of the
 * readings and survives value writes, so re-entering a cell's value does not
 * clear it (the readings, by contrast, are cleared). A blank cell carrying
 * only the flag is kept and saved. A formula cell's flag is kept in memory,
 * so it returns once the cell holds a constant again, but it is neither saved
 * nor read from a file: Excel keeps it on constants and blanks only. Row (`ph`) and column
 * (`phonetic`) flags read from OOXML are preserved on save but have no setter
 * here.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the cell coordinate is out of range.
 */
FM_API fm_status_t fm_sheet_set_cell_show_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                   int32_t show);

/**
 * @brief Stores how a cell's phonetic guide renders (OOXML `<phoneticPr>`).
 *
 * `font_id` indexes the workbook's font table for the ruby text; `type` is
 * the kana form Excel generates readings in (`FM_PHONETIC_TYPE_*`) and
 * `alignment` how the kana is distributed over the characters it covers
 * (`FM_PHONETIC_ALIGNMENT_*`).
 *
 * Deliberately separate from `fm_sheet_set_cell_phonetic_runs` in both
 * directions: editing the readings does not reset the rendering, and
 * setting the rendering does not touch the readings. The two are stored
 * together and travel together, but a host that only has one of them should
 * not have to supply the other.
 *
 * Only observable on a cell that has runs — the writer emits no element for
 * an unannotated cell — and any value-mutating setter clears the runs and
 * these properties together, so set them after the cell's text.
 *
 * The all-zero triple is what Excel infers for a guide that arrived with no
 * `<phoneticPr>` at all, so a host that does not care can leave it alone.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the cell coordinate is out of range, when
 *         `font_id` exceeds 65535, or when `type` or `alignment` is not one of the
 *         four values above.
 */
FM_API fm_status_t fm_sheet_set_cell_phonetic_properties(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                         uint32_t col, uint32_t font_id, uint32_t type,
                                                         uint32_t alignment);

/**
 * @brief Stores a `Blank` literal at `(row, col)`. Equivalent to
 *        clearing a cell.
 *
 * A table's header cell is not cleared: the column takes the profile's
 * default name (`Column1`, `列1`; `fm_workbook_rename_table_column` with an
 * empty name), and the cell holds that name.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_blank(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col);

/**
 * @brief Stores a formula at `(row, col)` and registers its
 *        dependencies.
 *
 * `formula` is the raw Excel-style formula text (with or without a
 * leading `=`). Text the formula parser cannot parse is rejected without
 * changing the workbook: a syntax error, or a built-in called through a
 * sheet / workbook qualifier.
 * Every formula accepted here can be read back through
 * `fm_sheet_get_formula_r1c1`. The text is stored the way Excel stores
 * entered text: number literals respelled from their values (`1.50` reads
 * back as `1.5`), a unary sign touching one absorbed (`+5` as `5`), and
 * whitespace Excel drops (trailing, before `,`, around `:`, inside array
 * constants) removed.
 *
 * Structured references are accepted in the formula-bar form the getters
 * return and stored in Excel's long form: in a cell of table `T`, `[Qty]`
 * and `[@Qty]` store as `T[Qty]` and `T[[#This Row],[Qty]]`; anywhere,
 * `T[@Qty]` stores as `T[[#This Row],[Qty]]` and a bare table name `T` not
 * hidden by a defined name as `T[]`. An unqualified `[Qty]` outside every
 * table does not parse.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000) when `formula` does not parse.
 */
FM_API fm_status_t fm_sheet_set_formula(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row` or `col` is out of range.
 */
FM_API fm_status_t fm_sheet_get_value(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                      fm_value_t* out);

/**
 * @brief Reads the cell's phonetic guide (OOXML `<rPh>`), or an empty string
 *        when the cell has no guide.
 *
 * Every run's kana concatenated in reading order — the shape a furigana
 * field wants. Which characters each run covers is dropped; read the runs
 * individually when the spans matter, and note that writing this string
 * back through `fm_sheet_set_cell_phonetic` collapses a partially
 * annotated cell to a single whole-cell annotation.
 *
 * `*out_text` is a read-scratch-backed view into the handle's transient read
 * storage. It may be invalidated by the next successful scratch-backed read
 * on this handle; a validation-rejected call does not refresh the storage.
 * It is also invalid after a mutation or handle destruction. Copy it before
 * another call when it must outlive that window.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_text` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cell_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              const char** out_text);

/**
 * @brief Counts the `<rPh>` blocks attached to a cell.
 *
 * Zero for a cell with no annotation, including one that does not exist.
 * Pair with `fm_sheet_get_cell_phonetic_run` to walk the spans.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cell_phonetic_run_count(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                        uint32_t col, size_t* out_count);

/**
 * @brief Reads the `run_index`-th `<rPh>` block of a cell.
 *
 * `out->text` is a read-scratch-backed view with the same lifetime rules as
 * `fm_sheet_get_cell_phonetic`'s: the next successful scratch-backed read
 * on this handle may invalidate it, so a caller collecting every run must
 * copy each one before asking for the next.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `run_index` is past the cell's run count.
 */
FM_API fm_status_t fm_sheet_get_cell_phonetic_run(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                  uint32_t col, size_t run_index, fm_phonetic_run_t* out);

/**
 * @brief Reads how a cell's phonetic guide renders (OOXML `<phoneticPr>`).
 *
 * Reports the all-zero triple for a cell with no annotation, including one
 * that does not exist — the same state Excel infers for a guide written
 * without a `<phoneticPr>` element.
 *
 * Any out-parameter may be `NULL` to skip that field.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cell_phonetic_properties(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                         uint32_t col, uint32_t* out_font_id, uint32_t* out_type,
                                                         uint32_t* out_alignment);

/**
 * @brief Reports whether the cell's `<phoneticPr>` triple is explicit.
 *
 * `*out_present` is `1` when the triple came from a `<phoneticPr>` element
 * or `fm_sheet_set_cell_phonetic_properties` (an all-zero triple included),
 * and `0` when it is the state an absent element resolves to, which is also
 * what a cell with no annotation reports. Like the triple, it is cleared by
 * any value-mutating setter. A save spells the element out beside any runs,
 * and XLSB always stores the triple, so a reloaded or XLSB-loaded cell with
 * runs reports `1`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_present` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_has_cell_phonetic_properties(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                         uint32_t col, int32_t* out_present);

/**
 * @brief Reads the cell's "Show Phonetic" flag into `*out_show` (`0` or `1`).
 *        An absent cell reports `0`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_show` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cell_show_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                   uint32_t col, int32_t* out_show);

/**
 * @brief Renders the lambda closure stored at `(sheet_index, row, col)`
 *        as Excel formula text.
 *
 * Use this to recover a printable form of a `Value::Lambda` whose
 * source is otherwise opaque — for example, a cell whose formula
 * evaluates to a lambda via `=LET(...) -> LAMBDA(...)` rather than
 * carrying a literal `LAMBDA(...)` in the formula text. The rendering
 * is the full surface form `LAMBDA(p1,p2,body)` (no leading `=`),
 * suitable for re-parsing through `fm_sheet_set_formula`.
 *
 * On success `*out_text` is a read-scratch-backed view into the handle's
 * transient read storage. It may be invalidated by the next successful
 * scratch-backed read on this handle; a validation-rejected call does not
 * refresh the storage. It is also invalid after a mutation or handle
 * destruction. The scratch is reset on each successful producer so a
 * long-lived handle does not accumulate memory. Callers that need to retain
 * the string must copy it before another call.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         cell is absent, or the cached value is not a lambda.
 */
FM_API fm_status_t fm_sheet_get_lambda_text(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
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
 * on `wb` is the ABI-level purity contract. Structured references are bound
 * at `(row, col)` as `fm_sheet_set_formula` binds them (`[@Qty]` in a
 * table cell).
 *
 * An Array / spill result is reduced to its top-left element. This is a
 * pragmatic API-shape choice, NOT Excel implicit intersection (which
 * selects the element sharing the anchor's row / column and yields
 * `#VALUE!` when there is none) and NOT dynamic-array spilling (which
 * returns the whole array). To recover the full multi-cell result instead,
 * use `fm_sheet_evaluate_formula_array` + `_array_cell`.
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `formula`, or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row` or `col` is out of range;
 *         `FM_STATUS_OUT_OF_MEMORY` (4) when the evaluation arena is exhausted.
 */
FM_API fm_status_t fm_sheet_evaluate_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
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
 * `fm_sheet_evaluate_formula`.
 *
 * @note The same self-reference caveat as `fm_sheet_evaluate_formula`
 *       applies: a rule that reads its own anchored cell sees the cached
 *       value rather than raising `#REF!`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `formula`, or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the target cell or anchor cell is out of range;
 *         `FM_STATUS_OUT_OF_MEMORY` (4) when the evaluation arena is exhausted.
 */
FM_API fm_status_t fm_sheet_evaluate_cf_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                uint32_t anchor_row, uint32_t anchor_col, const char* formula,
                                                fm_value_t* out);

/**
 * @brief Evaluates `formula` as if entered at `(row, col)` and stashes the
 *        WHOLE result — array included — on the handle, returning its
 *        dimensions. Companion to `fm_workbook_evaluate_formula_array_cell`.
 *
 * This is the multi-cell counterpart to `fm_sheet_evaluate_formula`,
 * which reduces an array / spill result to its top-left element. Here the
 * result is preserved in full: a dynamic-array formula such as
 * `=SEQUENCE(2,3)` reports `*out_rows = 2`, `*out_cols = 3`; a scalar
 * result such as `=1+2` reports `*out_rows = *out_cols = 1` (a 1x1 array).
 * Parsing, anchoring, cross-sheet / defined-name resolution, `ROW()` /
 * `COLUMN()`, and the read-only purity contract are identical to
 * `fm_sheet_evaluate_formula` (the `const` on `wb` is the ABI-level
 * purity guarantee; the stash is internal handle state, not observable
 * workbook state, mirroring how the read scratch is populated on the read
 * path).
 *
 * Two-step protocol: call this to evaluate + learn the dimensions, then call
 * `fm_workbook_evaluate_formula_array_cell` for each row-major index in
 * `[0, (*out_rows) * (*out_cols))`. The handle's private array stash remains
 * available until the next `fm_sheet_evaluate_formula_array` call or any
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
 * @note The same self-reference caveat as `fm_sheet_evaluate_formula`
 *       applies: a self-reference reads the target cell's cached value.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `formula`, `out_rows`, or
 *         `out_cols` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row` or `col` is out of range;
 *         `FM_STATUS_OUT_OF_MEMORY` (4) when the evaluation arena is exhausted (the
 *         handle's array stash is left untouched, per the "a rejected call
 *         does not refresh storage" rule above).
 */
FM_API fm_status_t fm_sheet_evaluate_formula_array(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                   uint32_t col, const char* formula, uint32_t* out_rows,
                                                   uint32_t* out_cols);

/**
 * @brief Reads the row-major `index`-th cell of the most recent
 *        `fm_sheet_evaluate_formula_array` result into `*out`.
 *
 * Valid `index` values are `[0, rows * cols)` where `rows` / `cols` are the
 * dimensions the producing `fm_sheet_evaluate_formula_array` call
 * returned; cell `(r, c)` is at `index = r * cols + c`.
 *
 * For a `FM_VAL_TEXT` cell, `out->u.text` is a read-scratch-backed view
 * (scratch is refreshed by each successful producer, exactly like
 * `fm_sheet_get_cell_at`) and may be invalidated by the next successful
 * scratch-backed read on this handle. A validation-rejected call does not
 * refresh the storage; mutation and handle destruction also invalidate it.
 * Copy it before another call if you need to retain it.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is past the stashed result (or no
 *         array evaluation has run on this handle yet).
 */
FM_API fm_status_t fm_workbook_evaluate_formula_array_cell(const fm_workbook_t* wb, size_t index, fm_value_t* out);

/* -------------------------------------------------------------------------- */
/* Reference text                                                             */
/* -------------------------------------------------------------------------- */

/**
 * @brief A parsed reference: a 0-based, inclusive rectangle with an optional
 *        sheet qualifier.
 *
 * `valid` is 0 when the text is not a reference, and every other field is
 * then zero / `NULL`. `sheet_name` is `NULL` when `has_sheet` is 0, else the
 * unquoted UTF-8 name as a read-scratch-backed view. A whole column or row
 * is expanded to the grid edge, and `row0 <= row1`, `col0 <= col1`.
 */
typedef struct {
  int32_t valid;
  int32_t has_sheet;
  const char* sheet_name;
  uint32_t row0;
  uint32_t col0;
  uint32_t row1;
  uint32_t col1;
} fm_parsed_reference_t;

/**
 * @brief Parses `text` as `INDIRECT` reads a reference under the workbook's
 *        profile: A1 when `r1c1` is 0, R1C1 (letters and brackets of the
 *        profile's locale) when 1.
 *
 * Accepts an optional local sheet qualifier (quoted with `''` escaping, or
 * bare), `$` markers (dropped), a `:` range in either order, and whole
 * columns / rows. 3D, external, named, structured and multi-area
 * references, surrounding whitespace, and a sheet name longer than 31
 * UTF-16 units read as `valid = 0`. With `has_base` 1, R1C1 relative axes
 * are measured from `(base_row, base_col)` and wrap at the grid edge; with
 * `has_base` 0 a relative axis reads as `valid = 0`.
 *
 * @return `FM_STATUS_OK` (0) on success, including text that is not a reference;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `text` or `out` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `text` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `r1c1` or `has_base` is not 0/1, or
 *         `has_base` is 1 with the base outside the grid.
 */
FM_API fm_status_t fm_workbook_parse_reference(const fm_workbook_t* wb, const char* text, int32_t r1c1,
                                               int32_t has_base, uint32_t base_row, uint32_t base_col,
                                               fm_parsed_reference_t* out);

/**
 * @brief Formats the inclusive rectangle `(row0, col0)..(row1, col1)` as A1
 *        reference text, the inverse of `fm_workbook_parse_reference`.
 *
 * `sheet_name` `NULL` or empty writes no qualifier; otherwise the name is
 * quoted as a stored formula quotes it (and is not validated as a sheet
 * name). A single cell writes one endpoint (`A1`), a rectangle spanning
 * every row writes columns (`B:D`), one spanning every column writes rows
 * (`2:11`). `absolute` 1 prefixes each component with `$`. `*out_text` is
 * a read-scratch-backed view.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out_text` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `sheet_name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a coordinate is outside the grid,
 *         `row0 > row1` or `col0 > col1`, or `absolute` is not 0/1.
 */
FM_API fm_status_t fm_workbook_format_reference(const fm_workbook_t* wb, const char* sheet_name, uint32_t row0,
                                                uint32_t col0, uint32_t row1, uint32_t col1, int32_t absolute,
                                                const char** out_text);

/* -------------------------------------------------------------------------- */
/* Iteration / dump                                                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the number of written coordinates on `sheet_index`.
 *
 * Counts every coordinate that a value, formula or format setter,
 * `fm_sheet_set_blank`, or a spill anchor commit or rejection has
 * written. Clearing a cell keeps its slot, so it stays counted. Columns that
 * were never written are not counted. Also counts every dynamic-array spill
 * phantom that has no stored slot: phantoms carry an effective value through
 * `fm_sheet_get_value` / `fm_sheet_get_cell_at` but live only in the spill
 * table.
 *
 * The count is the upper bound for `fm_sheet_get_cell_at` iteration: indices
 * in `[0, count)` are dense and enumerate every counted coordinate in
 * `(row, col)` ascending order. The enumeration is stable for a given
 * workbook state (no cell mutation has occurred between calls).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL` or `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cell_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `idx`-th cell on `sheet_index`.
 *
 * Enumerates the coordinates counted by `fm_sheet_get_cell_count`, sorted by
 * `(row, col)` ascending; unwritten columns are not enumerated.
 *
 * On success the call writes `*out_row`, `*out_col`, `*out_value`, and
 * — if `out_formula != NULL` — points `*out_formula` at the cell's
 * formula text, spelled as `fm_sheet_get_formula` returns it (or `NULL`
 * when the cell is a pure literal). The formula pointer has the
 * read-scratch lifetime described below for `*out_value`: valid until the
 * next successful scratch-backed read on this handle, the next mutation, or
 * handle destruction.
 *
 * A `FM_VAL_TEXT` payload in `*out_value`, by contrast, is a
 * read-scratch-backed view and may be invalidated by the next successful
 * scratch-backed read on this handle; a validation-rejected call does not
 * refresh the storage. Mutation and handle destruction also invalidate it.
 * The scratch is refreshed by each successful producer so iterating
 * `cell_at` does not accumulate memory.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_row`, `out_col`, or
 *         `out_value` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx >= cell_count`.
 */
FM_API fm_status_t fm_sheet_get_cell_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, uint32_t* out_row,
                                        uint32_t* out_col, const char** out_formula, fm_value_t* out_value);

/**
 * @brief Reads the formula of the cell at `(row, col)` in A1 notation.
 *
 * The text carries the leading `=`, as `fm_sheet_get_cell_at` reports it. A
 * cell that holds no formula (a literal, a spill phantom, or no cell at
 * all) yields `""`. A data table's result
 * cell yields the table's formula, `=TABLE(row_input,col_input)` with an
 * absent input left empty (`=TABLE(,B3)`) and a deleted one as `#REF!`; the
 * formula bar shows it in braces, as for a legacy array formula.
 *
 * The text is the formula-bar spelling. External references name the book
 * (`[Book.xlsx]Sheet1!A1`), never its `[N]` index. Structured references
 * take the display form for this cell: `[#This Row]` is spelled `@`, inside
 * the referenced table's own range a single-column reference with no
 * specifier or `@` alone drops the table name (`[@Qty]`, `SUM([Qty])`;
 * outside it `tblSales[@Qty]`), and `T[]` reads `T`.
 * `fm_sheet_set_formula` accepts this text back.
 *
 * `*out_formula` is a read-scratch-backed view: valid until the next
 * successful scratch-backed read, any mutation, or handle destruction.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         coordinate is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        const char** out_formula);

/**
 * @brief Reads the formula of the cell at `(row, col)` in R1C1 notation,
 *        relative to that cell.
 *
 * The text carries no leading `=`, and follows the A1 formatter's spacing
 * and case rules. Structured references take the display form
 * `fm_sheet_get_formula` gives them at this cell. Sheet and external qualifiers are quoted as a whole
 * (`'[Book.xlsx]My Sheet'!R1C1`) only when the path, book or sheet name needs
 * it; a 3-D external reference with no path stays unquoted. A cell that holds no formula yields `""`. `*out_formula` is
 * a read-scratch-backed view, as for `fm_sheet_get_formula`.
 *
 * A formula loaded from a file that does not parse is kept verbatim: it
 * evaluates to `#NAME?` and fails here with `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000).
 * `fm_sheet_set_formula` never produces that state.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         coordinate is outside the grid;
 *         `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000) when the stored formula does not
 *         parse.
 */
FM_API fm_status_t fm_sheet_get_formula_r1c1(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             const char** out_formula);

/** @brief Opaque owned snapshot of one page of cells inside a rectangle. */
typedef struct fm_cell_range fm_cell_range_t;

/**
 * @brief Enumerates the populated cells of an inclusive rectangle, one page
 *        at a time.
 *
 * A cell is populated when it holds a formula or a non-blank value; spill
 * phantoms are included with their spilled value. Cells come in row-major
 * order. `cursor` is the position to resume from, encoded as
 * `row * 16384 + col`: pass `0` for the first page and the value
 * `fm_cell_range_next_cursor` reported for each following one. A page holds
 * at most `limit` cells; `0` means the maximum, 65,536, and larger values
 * are capped to it.
 *
 * On success `*out` is an owned handle the caller releases with
 * `fm_cell_range_destroy`; it is a snapshot, so later mutations do not
 * change it.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         rectangle is inverted or outside the grid, or `cursor` is past
 *         the last cell of the grid.
 */
FM_API fm_status_t fm_sheet_get_cells_in_range(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                               uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                               uint64_t cursor, uint32_t limit, fm_cell_range_t** out);

/**
 * @brief Number of cells in the page.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_cell_range_get_cell_count(const fm_cell_range_t* range, size_t* out_count);

/**
 * @brief Reads cell `idx` of the page.
 *
 * `*out_formula` is the stored A1 formula (as `fm_sheet_get_formula`
 * reports it), or `NULL` when the cell holds no formula; `out_formula` may
 * be `NULL` to skip it. Both the formula and a `FM_VAL_TEXT` payload are an
 * owned snapshot held by the handle and stay valid until `fm_cell_range_destroy`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `range`, `out_row`, `out_col`, or
 *         `out_value` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_cell_range_get_cell_at(const fm_cell_range_t* range, size_t idx, uint32_t* out_row,
                                             uint32_t* out_col, const char** out_formula, fm_value_t* out_value);

/**
 * @brief Reads the cursor that resumes after this page, or `UINT64_MAX` when
 *        the rectangle holds no further populated cell.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_cell_range_next_cursor(const fm_cell_range_t* range, uint64_t* out_cursor);

/** @brief Releases a page handle. `range == NULL` is a no-op. */
FM_API void fm_cell_range_destroy(fm_cell_range_t* range);

/* -------------------------------------------------------------------------- */
/* Display text                                                               */
/* -------------------------------------------------------------------------- */

/**
 * @brief Outcome of a display-text rendering.
 *
 * `FM_DISPLAY_OVERFLOW`: the value cannot be shown in the format (a date
 * serial outside the calendar, for one) and the text is `"########"`.
 * `FM_DISPLAY_INVALID_FORMAT`: the format code is malformed and the text is
 * the General rendering.
 */
typedef enum {
  FM_DISPLAY_OK = 0,
  FM_DISPLAY_OVERFLOW = 1,
  FM_DISPLAY_INVALID_FORMAT = 2,
} fm_display_status_t;

/**
 * @brief Reads the text Excel displays for the cell at `(row, col)`.
 *
 * Rendered from the resolved value and effective cell/row/column/default
 * XF number format under the workbook's date system, independent of column
 * width: a `*x` repeat fill contributes no characters and an overflow is
 * reported through `*out_status` rather than as a width-dependent `#` run.
 * A coordinate with no value or spill result is `""`. `*out_text` is a
 * read-scratch-backed view.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         coordinate is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_display_text(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             const char** out_text, int32_t* out_status);

/**
 * @brief Renders `value` under the stored-form number format `format_code`
 *        (`General`, `[Red]`, ...), using the workbook's date system.
 *
 * An empty `format_code` means General. `value` must be blank, a number,
 * a boolean, text, or an error. `*out_text` is a read-scratch-backed view.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`, or a
 *         `FM_VAL_TEXT` value carries a `NULL` text;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the value kind is not one of the above,
 *         a number is NaN or infinite, or an error code is out of range.
 */
FM_API fm_status_t fm_workbook_format_value(const fm_workbook_t* wb, const fm_value_t* value, const char* format_code,
                                            const char** out_text, int32_t* out_status);

/**
 * @brief Writes the number of defined names attached to the workbook to
 *        `*out_count`.
 *
 * Defined names round-trip through OOXML even when not yet evaluated.
 * Counts every `<definedName>` entry in declaration order.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is NULL.
 */
FM_API fm_status_t fm_workbook_get_defined_name_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `idx`-th defined name with its scope.
 *
 * On success `*out_name` is a model-backed view into the workbook handle,
 * valid until the next mutation of the defined-name list --
 * `fm_workbook_set_defined_name`/`_scoped`, or any workbook edit that
 * rewrites defined names as a side effect (a structural row/column
 * insert/delete, or a sheet rename/removal) -- or until the handle is
 * destroyed. `*out_formula` is spelled as `fm_sheet_get_formula` spells
 * formulas and is valid until the next successful scratch-backed read on
 * this handle, the next mutation, or handle destruction.
 *
 * `out_local_sheet_id` is optional: when non-NULL it receives `-1` for
 * workbook scope, or a 0-based sheet index for sheet-local scope.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_name`, or `out_formula` is
 *         `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_get_defined_name_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                                   const char** out_formula, int32_t* out_local_sheet_id);

/**
 * @brief Reads the `comment` and `hidden` attributes of the `idx`-th defined
 *        name (the entry `fm_workbook_get_defined_name_at` returns for `idx`).
 *
 * `*out_comment` is a model-backed view with the lifetime of
 * `fm_workbook_get_defined_name_at`'s strings, `""` when the name has no comment.
 * `*out_hidden` is 0/1.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_get_defined_name_attrs_at(const fm_workbook_t* wb, size_t idx, const char** out_comment,
                                                         int32_t* out_hidden);

/**
 * @brief Writes the number of tables attached to the workbook to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is NULL.
 */
FM_API fm_status_t fm_workbook_get_table_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `idx`-th table's identifying metadata.
 *
 * On success `*out_name`, `*out_display_name`, and `*out_ref` are
 * model-backed views into the workbook handle. They are valid until the
 * next mutation of the table list -- `fm_workbook_add_table`/`_update`/
 * `_remove`, or any workbook edit that rewrites table metadata as a side
 * effect (a structural row/column insert/delete, or a sheet rename/removal)
 * -- or until the handle is destroyed. Reads do not invalidate them.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_get_table(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                         const char** out_display_name, const char** out_ref, size_t* out_sheet_index);

/** Creates a worksheet table (`xl/tables/tableN.xml`). `column_names` must
 * contain one non-empty header for every column in `ref`; each name must be
 * unique within the table, and `column_count` must equal the width of `ref`.
 * `style_name` may be NULL or empty to omit table styling. The table's filter
 * definition is created automatically. When `header_row` is set, the caller
 * still owns the header cells: Excel expects the first row of `ref` to hold
 * exactly the column names. */
FM_API fm_status_t fm_workbook_add_table(fm_workbook_t* wb, size_t sheet_index, const char* ref, const char* name,
                                         const char* display_name, const char* const* column_names, size_t column_count,
                                         const char* style_name, int32_t header_row, int32_t totals_row,
                                         size_t* out_index);

/** Replaces a table's range and optional visual style. `ref` is required;
 * `style_name == NULL` preserves the existing style, a non-empty name
 * replaces the style name and keeps the four `fm_table_info_t` toggles
 * (a table without `<tableStyleInfo>` gets the creation defaults), and an
 * empty style removes `<tableStyleInfo>`. Negative `header_row` and
 * `totals_row` preserve their existing flags; zero and positive values set
 * them false and true respectively. `ref` must keep the width the table's
 * column list already describes. Turning the totals row on writes each
 * column's totals cell (`fm_workbook_set_table_column_totals`) on the last
 * row of `ref`. */
FM_API fm_status_t fm_workbook_update_table(fm_workbook_t* wb, size_t index, const char* ref, const char* style_name,
                                            int32_t header_row, int32_t totals_row);

/** Removes the table at `index`. */
FM_API fm_status_t fm_workbook_remove_table(fm_workbook_t* wb, size_t index);

/**
 * @brief Table metadata beyond its identity (`fm_workbook_get_table_info`).
 *
 *   * `style_name`    — model-backed view, empty when the table has no style
 *                       element or the element names no style. It may name a
 *                       built-in style or a custom one (`fm_table_style_t`).
 *   * `style_present` — non-zero when the table has `<tableStyleInfo>`.
 *   * the `show_*` toggles are `0` or `1`, and `0` without the element.
 *   * `header_row_count` / `totals_row_count` are `0` or `1`.
 *   * `column_count` bounds `fm_workbook_get_table_column`'s `column`.
 *
 * Views share the lifetime of `fm_workbook_get_table`'s.
 */
typedef struct {
  const char* style_name;
  int32_t style_present;
  int32_t show_first_column;
  int32_t show_last_column;
  int32_t show_row_stripes;
  int32_t show_column_stripes;
  uint32_t header_row_count;
  uint32_t totals_row_count;
  uint32_t column_count;
} fm_table_info_t;

/**
 * @brief One table column (`fm_workbook_get_table_column`).
 *
 *   * `totals_row_function` — the `ST_TotalsRowFunction` token: `none`, `sum`,
 *                             `min`, `max`, `average`, `count`, `countNums`,
 *                             `stdDev`, `var` or `custom`.
 *   * `totals_row_label`    — the totals-row label, empty when absent.
 *   * `totals_row_formula`  — the `custom` function's formula without a
 *                             leading `=`, empty otherwise, spelled as
 *                             `fm_sheet_get_formula` spells formulas.
 *
 * `totals_row_formula` is spelled as at the column's totals cell, inside
 * the table (`SUBTOTAL(109,[Qty])`).
 *
 * Views share the lifetime of `fm_workbook_get_table`'s; none is `NULL`. A
 * `totals_row_formula` respelled from the model's (see
 * `fm_workbook_get_file_name`) is read-scratch-backed instead.
 */
typedef struct {
  const char* name;
  const char* totals_row_function;
  const char* totals_row_label;
  const char* totals_row_formula;
} fm_table_column_info_t;

/**
 * @brief Reads the style, row counts and column count of table `index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_table_info(const fm_workbook_t* wb, size_t index, fm_table_info_t* out);

/**
 * @brief Reads column `column` (0-based, left to right) of table `index`.
 *
 * @return As `fm_workbook_get_table_info`, with `FM_STATUS_INVALID_ARGUMENT` (2) also for an
 *         out-of-range `column`.
 */
FM_API fm_status_t fm_workbook_get_table_column(const fm_workbook_t* wb, size_t index, size_t column,
                                                fm_table_column_info_t* out);

/**
 * @brief Sets the table's style name and toggles, making `<tableStyleInfo>`
 *        present. A `NULL` or empty `style_name` writes the element without a
 *        style name (no style, toggles kept); the toggles are `0` or non-zero.
 *        `fm_workbook_add_table` starts a styled table with row stripes on
 *        and the other three off.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_set_table_style_info(fm_workbook_t* wb, size_t index, const char* style_name,
                                                    int32_t show_first_column, int32_t show_last_column,
                                                    int32_t show_row_stripes, int32_t show_column_stripes);

/**
 * @brief Renames table `index`. `display_name == NULL` uses `name` for both.
 *
 * Both names follow `fm_workbook_add_table`'s rules: non-empty, and unique
 * case-insensitively against every other table's name and display name.
 * Structured references naming the old name or display name are rewritten to
 * the new display name wherever formulas live: cells, defined names,
 * conditional formats, data validations, calculated-column and totals-row
 * formulas, and chart series. A pivot cache sourced from the table follows.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range or a name is empty
 *         or taken.
 */
FM_API fm_status_t fm_workbook_rename_table(fm_workbook_t* wb, size_t index, const char* name,
                                            const char* display_name);

/**
 * @brief Renames column `column` (0-based) of table `index`, as typing over
 *        its header cell does.
 *
 * The final name is `name` kept verbatim (spaces and a case-only change
 * included), except that an empty `name` takes the profile's default
 * `<word>N` (`Column1`, `列1`) with the lowest unused N >= 1, and a name
 * another column of the table already has (ASCII case-insensitively)
 * takes `<name>N` with the lowest unused N >= 2 (`Amount2` becomes
 * `Amount22`). When it differs from the current name, every structured
 * reference to the column through this table is rewritten: cell formulas
 * on every sheet, totals-row formulas, defined names and chart series. A
 * table with a header row has the final name written to the header cell,
 * also when the name is unchanged. External-book references, `INDIRECT`
 * text and pivot caches are not rewritten.
 *
 * `*out_name` receives the final name, read-scratch-backed: valid until the
 * next successful scratch-backed read, any mutation, or handle destruction.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `name` or `out_name` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` or `column` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is not valid UTF-8.
 */
FM_API fm_status_t fm_workbook_rename_table_column(fm_workbook_t* wb, size_t index, size_t column, const char* name,
                                                   const char** out_name);

/**
 * @brief Sets column `column`'s totals-row function and label.
 *
 * `function` is an `ST_TotalsRowFunction` token (see `fm_table_column_info_t`);
 * `NULL` means `none`. `formula` is required for `custom` and ignored
 * otherwise; it is stored without a leading `=`, its structured references
 * bound as `fm_sheet_set_formula` binds them at the column's totals cell
 * (`SUBTOTAL(109,[Qty])` stores `SUBTOTAL(109,Table[Qty])`). `label` may be
 * `NULL` or empty for none.
 *
 * While the totals row is shown its cell on the last row of the table's
 * range follows, as the desktop application writes it: a built-in function
 * becomes `SUBTOTAL(<101..110>,Table[Column])`, `custom` becomes `formula`,
 * otherwise the label is written as text, or the cell is cleared.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` or `column` is out of range, the
 *         token is unknown, or `custom` has no formula;
 *         `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000) when `formula` does not parse.
 */
FM_API fm_status_t fm_workbook_set_table_column_totals(fm_workbook_t* wb, size_t index, size_t column,
                                                       const char* function, const char* label, const char* formula);

/**
 * @brief Writes the number of passthrough OOXML parts the reader carried
 *        through unchanged to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is NULL.
 */
FM_API fm_status_t fm_workbook_get_passthrough_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `idx`-th passthrough part's path.
 *
 * On success `*out_path` is a model-backed view into the workbook's
 * passthrough-part list. It remains valid until the passthrough list is
 * replaced by a reader/writer round-trip, another invalidating workbook
 * mutation, or handle destruction. Reads do not invalidate it.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_get_passthrough_at(const fm_workbook_t* wb, size_t idx, const char** out_path);

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
} fm_hyperlink_t;

/**
 * @brief Comment descriptor. `author` and `text` are model-backed UTF-8
 *        views with the same lifetime contract as `fm_hyperlink_t`.
 */
typedef struct {
  uint32_t row;
  uint32_t col;
  const char* author;
  const char* text;
} fm_comment_t;

/**
 * @brief Appends a hyperlink covering the inclusive rectangle
 *        `(hl.row, hl.col)..(hl.last_row, hl.last_col)` to `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the hyperlink
 *         rectangle is inverted or outside the Excel grid.
 */
FM_API fm_status_t fm_sheet_add_hyperlink(fm_workbook_t* wb, size_t sheet_index, fm_hyperlink_t hl);

/**
 * @brief Removes every hyperlink anchored at the exact `(row, col)`
 *        cell on `sheet_index`. No-op when nothing matches.
 *
 * @return `FM_STATUS_OK` (0) on success (including the no-match case);
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_hyperlink(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col);

/**
 * @brief Removes the hyperlink at `index` on `sheet_index`. Mirrors the
 *        index domain of fm_sheet_get_hyperlink_at.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_hyperlink_at(fm_workbook_t* wb, size_t sheet_index, size_t index);

/**
 * @brief Drops every hyperlink on `sheet_index`. No-op when the sheet
 *        already has no hyperlinks.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_hyperlinks(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Appends a merge range to the sheet's merge list.
 *        Corners are normalised internally. The resulting inclusive
 *        rectangle must lie within the Excel grid.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the merge
 *         rectangle is outside the Excel grid.
 */
FM_API fm_status_t fm_sheet_add_merge(fm_workbook_t* wb, size_t sheet_index, fm_merge_range merge);

/**
 * @brief Removes every merge range on `sheet_index` that overlaps `range`
 *        (inclusive rectangle intersection). No-op when nothing
 *        overlaps. `range` corners are normalised internally so the
 *        caller may pass either (first <= last) or transposed corners,
 *        mirroring fm_sheet_add_merge.
 *
 * @return `FM_STATUS_OK` (0) on success (including the no-overlap case);
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_merge(fm_workbook_t* wb, size_t sheet_index, fm_merge_range range);

/**
 * @brief Removes the merge at `index` on `sheet_index`. Mirrors the
 *        index domain of fm_sheet_get_merge_at.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_merge_at(fm_workbook_t* wb, size_t sheet_index, size_t index);

/**
 * @brief Drops every merge range on `sheet_index`. No-op when the sheet
 *        already has no merges.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_merges(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Reads `(out_row, out_col, out_author, out_text)` for the
 *        comment at `(row, col)` on `sheet_index`. Returns `FM_STATUS_NOT_FOUND` (6) when no
 *        comment is anchored there, and `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when
 *        `sheet_index` is out of range.
 *
 * `out->author` and `out->text` are model-backed views into the workbook
 * handle. Both are valid until the next mutation that touches the sheet's
 * comment list or until the handle is destroyed. Reads do not invalidate
 * them.
 */
FM_API fm_status_t fm_sheet_get_comment(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        fm_comment_t* out);

/**
 * @brief Returns the number of comments attached to `sheet_index`, including
 *        comments anchored on cells that carry no value.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_comment_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `index`-th comment on `sheet_index` into `out`, in storage
 *        order. Unlike `fm_sheet_get_comment`, this enumerates every
 *        comment on the sheet regardless of whether the anchor cell holds
 *        a value, so callers can discover comments without already
 *        knowing their `(row, col)`.
 *
 * `out->author` and `out->text` are model-backed views into the workbook
 * handle, with the same lifetime contract as `fm_sheet_get_comment`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_comment_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                           fm_comment_t* out);

/**
 * @brief Reads the `index`-th hyperlink on `sheet_index` into `out`.
 *
 * String fields in `out` are model-backed views (see `fm_hyperlink_t`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_hyperlink_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                             fm_hyperlink_t* out);

/**
 * @brief Returns the number of hyperlinks attached to `sheet_index`.
 */
FM_API fm_status_t fm_sheet_get_hyperlink_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `index`-th merge range on `sheet_index`.
 */
FM_API fm_status_t fm_sheet_get_merge_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                         fm_merge_range* out);

/**
 * @brief Returns the number of merge ranges attached to `sheet_index`.
 */
FM_API fm_status_t fm_sheet_get_merge_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Copies every merge range on `sheet_index` that intersects `range` into
 *        the caller's buffer, in storage order.
 *
 * `*out_count` always receives the total number of intersecting merges; at
 * most `capacity` of them are written to `out`. Call with `capacity == 0`
 * (and `out` may then be `NULL`) to size the buffer first. `range` corners
 * are normalised, mirroring `fm_sheet_remove_merge`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is `NULL`, or `out`
 *         is `NULL` while `capacity > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_merges_in_range(const fm_workbook_t* wb, size_t sheet_index, fm_merge_range range,
                                                fm_merge_range* out, uint32_t capacity, uint32_t* out_count);

/**
 * @brief Inserts or replaces the comment at `(row, col)` on `sheet_index`.
 *        Pass `text == NULL || text[0] == '\0'` to remove an existing
 *        comment. The strings are copied into the workbook handle's text
 *        storage. The coordinate must lie within the Excel grid.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the coordinate
 *         is outside the Excel grid;
 *         `FM_STATUS_THREADED_COMMENT_INVALID` (9006) when the cell already holds a threaded
 *         comment (a cell carries a note or a thread, never both).
 *         Replacing a note keeps its `xr:uid`.
 */
FM_API fm_status_t fm_sheet_set_comment(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        const char* author, const char* text);

/**
 * @brief Reads whether the note at `(row, col)` is shown without hovering
 *        (`visibility:visible` in the note's VML shape).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_visible` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no note is anchored at the cell.
 */
FM_API fm_status_t fm_sheet_get_comment_visible(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                int32_t* out_visible);

/**
 * @brief Shows (`visible != 0`) or hides the note at `(row, col)`. The note's
 *        box geometry is kept.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no note is anchored at the cell.
 */
FM_API fm_status_t fm_sheet_set_comment_visible(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                int32_t visible);

/* -------------------------------------------------------------------------- */
/* Protection records                                                         */
/* -------------------------------------------------------------------------- */

/**
 * @brief `<workbookProtection>`. The host computes every hash; the engine
 *        stores the strings as given. On input a `NULL` string and an empty
 *        string both mean absent, and a zero spin count means absent. On
 *        output the strings are views valid until the next successful
 *        scratch-backed read on the handle.
 */
typedef struct {
  int32_t lock_structure;
  int32_t lock_windows;
  int32_t lock_revision;
  uint32_t workbook_spin_count;
  uint32_t revisions_spin_count;
  /** Legacy 16-bit password hash, as the hex text Excel writes. */
  const char* workbook_password;
  const char* workbook_algorithm_name;
  const char* workbook_hash_value;
  const char* workbook_salt_value;
  const char* revisions_password;
  const char* revisions_algorithm_name;
  const char* revisions_hash_value;
  const char* revisions_salt_value;
} fm_workbook_protection_t;

/**
 * @brief Reads the workbook's `<workbookProtection>`; every field is zero or
 *        empty when the workbook has none.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_workbook_protection(const fm_workbook_t* wb, fm_workbook_protection_t* out);

/**
 * @brief Replaces the modelled attributes of `<workbookProtection>`. A record
 *        with every field zero or empty removes the element. Attributes the
 *        record does not model (the password character sets) are kept.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `protection` is `NULL`.
 */
FM_API fm_status_t fm_workbook_set_workbook_protection(fm_workbook_t* wb, const fm_workbook_protection_t* protection);

/**
 * @brief One `<protectedRange>`: a named range users may edit on a protected
 *        sheet. `sqref` is the A1 range list (`B2:C3 E5`). Input handling and
 *        output lifetime follow `fm_workbook_protection_t`.
 */
typedef struct {
  uint32_t spin_count;
  const char* name;
  const char* sqref;
  const char* password;
  const char* algorithm_name;
  const char* hash_value;
  const char* salt_value;
  const char* security_descriptor;
} fm_protected_range_t;

/**
 * @brief Number of protected ranges on `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_protected_range_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads protected range `index` in file order.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_protected_range(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                fm_protected_range_t* out);

/**
 * @brief Appends a protected range. The range follows row and column edits
 *        like any other sheet range and disappears when an edit deletes all
 *        of it.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `range` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `name` is empty or
 *         already names another range of the sheet, or `sqref` does not parse
 *         as an A1 range list.
 */
FM_API fm_status_t fm_sheet_add_protected_range(fm_workbook_t* wb, size_t sheet_index,
                                                const fm_protected_range_t* range);

/**
 * @brief Replaces protected range `index`; the argument checks of
 *        `fm_sheet_add_protected_range` apply, with `index` also validated.
 */
FM_API fm_status_t fm_sheet_update_protected_range(fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                   const fm_protected_range_t* range);

/**
 * @brief Removes protected range `index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_protected_range(fm_workbook_t* wb, size_t sheet_index, size_t index);

/** @brief Data-validation error style (`fm_data_validation_t::error_style`, `fm_validation_outcome_t::error_style`). */
typedef enum {
  FM_VALIDATION_ERROR_STYLE_STOP = 0,
  FM_VALIDATION_ERROR_STYLE_WARNING = 1,
  FM_VALIDATION_ERROR_STYLE_INFORMATION = 2,
} fm_validation_error_style_t;

/**
 * @brief Data-validation rule descriptor.
 *
 * Mirrors `formulon::DataValidation`:
 *   * `ranges` / `range_count` — the cell-range list this rule applies to.
 *     For getter results the pointer is a model-backed view into the
 *     workbook handle (see `fm_sheet_get_validation` for the lifetime
 *     contract). For `fm_sheet_add_validation` the caller owns the buffer and
 *     must keep
 *     it alive for the duration of the call only; the implementation
 *     deep-copies every range into the sheet's owned storage.
 *   * `type` — 0 none, 1 whole, 2 decimal, 3 list, 4 date, 5 time,
 *     6 textLength, 7 custom.
 *   * `op`   — 0 between, 1 notBetween, 2 equal, 3 notEqual,
 *     4 greaterThan, 5 lessThan, 6 greaterThanOrEqual, 7 lessThanOrEqual.
 *   * `error_style` — an `fm_validation_error_style_t`.
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
} fm_data_validation_t;

/**
 * @brief Returns the number of data-validation rules attached to `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_validation_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `index`-th data-validation rule on `sheet_index` into `out`.
 *
 * Every pointer in `*out` (`ranges` and the six string fields) is a
 * model-backed view owned by the workbook handle. It remains valid until the
 * next mutation that touches the sheet's validation list (any of
 * `fm_sheet_add_validation`, `fm_sheet_remove_validation`,
 * `fm_sheet_clear_validations`, or any reader / writer round-trip), or
 * until the handle is destroyed. Callers that need to retain the
 * payload across mutations must copy it. `formula1` / `formula2` are
 * spelled as `fm_sheet_get_formula` spells formulas; when that spelling
 * differs from the model's (see `fm_workbook_get_file_name`), both are
 * read-scratch-backed instead, valid until the next scratch-backed read.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_validation(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                           fm_data_validation_t* out);

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
 * `v.formula1` / `v.formula2` may carry one leading `=`, which is dropped:
 * the rule stores them as files carry them, spelled as
 * `fm_sheet_set_formula` stores a cell formula.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `v.range_count > 0 && v.ranges == NULL` or `v.range_count`
 *         exceeds the per-call range cap, when any supplied range is
 *         outside the Excel grid, or when `v.type`, `v.op` or
 *         `v.error_style` is outside the domain listed above;
 *         `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000) when a non-empty `v.formula1` or
 *         `v.formula2` does not parse (the rule is not stored).
 */
FM_API fm_status_t fm_sheet_add_validation(fm_workbook_t* wb, size_t sheet_index, fm_data_validation_t v);

/**
 * @brief Removes the validation rule at `index` on `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_validation(fm_workbook_t* wb, size_t sheet_index, size_t index);

/**
 * @brief Drops every validation rule on `sheet_index`. No-op when the sheet
 *        already has no rules.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_validations(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Result of checking a value against the data-validation rule that
 *        covers a cell.
 *
 * `has_rule` is 0 when no rule covers the cell, and `valid` is then 1.
 * Otherwise `rule_index` is the `fm_sheet_get_validation` index of the
 * applied rule and `error_style` its error style (an
 * `fm_validation_error_style_t`).
 */
typedef struct {
  int32_t has_rule;
  int32_t valid;
  uint32_t rule_index;
  uint8_t error_style;
  uint8_t _pad[3];
} fm_validation_outcome_t;

/**
 * @brief Checks `proposed` against the data-validation rule covering
 *        `(row, col)`, without storing it.
 *
 * Excel applies one rule per cell: when sqrefs overlap, the first rule in
 * `fm_sheet_get_validation` order wins. `proposed` is taken as the
 * already-parsed value (text "5" is text, not a number). The rule's
 * formulas, including a `custom` formula that refers to the cell itself,
 * read the workbook as stored.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`, or a
 *         `FM_VAL_TEXT` value carries a `NULL` text;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the cell is
 *         outside the grid, the number is not finite, the error code is out
 *         of range, or the value kind is an array, reference or lambda.
 */
FM_API fm_status_t fm_sheet_validate_value(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                           const fm_value_t* proposed, fm_validation_outcome_t* out);

/**
 * @brief Enumerates, one page at a time, the cells whose current value fails
 *        the data-validation rule that covers them.
 *
 * Only populated cells are considered, with the meaning
 * `fm_sheet_get_cells_in_range` gives the word, in row-major order. `cursor`,
 * `limit` and the returned handle follow `fm_sheet_get_cells_in_range`: pass `0`
 * for the first page and `fm_cell_range_next_cursor`'s value for each
 * following one, which is `UINT64_MAX` once no further invalid cell exists.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cursor`
 *         is past the last cell of the grid.
 */
FM_API fm_status_t fm_sheet_get_invalid_cells(const fm_workbook_t* wb, size_t sheet_index, uint64_t cursor,
                                              uint32_t limit, fm_cell_range_t** out);

/* -------------------------------------------------------------------------- */
/* Threaded comments and persons                                              */
/* -------------------------------------------------------------------------- */

/**
 * @brief One `@person` mention inside a threaded comment's text.
 *
 * `start` and `length` are in UTF-16 code units of the text.
 */
typedef struct {
  const char* person_id;
  const char* mention_id;
  uint32_t start;
  uint32_t length;
} fm_mention_t;

/**
 * @brief One threaded comment: a thread's opening comment, or a reply when
 *        `parent_id` is non-empty.
 *
 * Ids are brace-wrapped uppercase GUIDs and `created` is
 * `YYYY-MM-DDTHH:MM:SS.ss`; the caller supplies both, the engine never mints
 * them. `done` is the thread's resolved state and is meaningful on the
 * opening comment only. On input every `const char*` may be `NULL`, meaning
 * empty, and `mentions` may be `NULL` only with a zero `mention_count`.
 */
typedef struct {
  const char* id;
  uint32_t row;
  uint32_t col;
  const char* person_id;
  const char* created;
  const char* text;
  const char* parent_id;
  int32_t done;
  const fm_mention_t* mentions;
  uint32_t mention_count;
} fm_threaded_comment_t;

/** @brief One entry of the workbook's person list. */
typedef struct {
  const char* id;
  const char* display_name;
  const char* user_id;
  const char* provider_id;
} fm_person_t;

/**
 * @brief Number of threaded comments on the sheet, replies included.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_threaded_comment_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads comment `idx` of the sheet's flat list, in file order: each
 *        thread's opening comment followed by its replies.
 *
 * Every string and the `mentions` array are read-scratch-backed views, with
 * the lifetime `fm_sheet_get_auto_filter` describes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_threaded_comment_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                    fm_threaded_comment_t* out);

/**
 * @brief Adds a threaded comment. An empty `parent_id` opens a new thread at
 *        `row`/`col`, which must hold neither a note nor a thread; a
 *        non-empty one appends a reply to that thread, taking its anchor and
 *        ignoring `row`, `col` and `done`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `comment` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a new
 *         thread's cell is outside the grid, or `mentions` is `NULL` with a
 *         non-zero count;
 *         `FM_STATUS_THREADED_COMMENT_INVALID` (9006) for a malformed id or timestamp,
 *         a duplicate id, an unknown person or parent, a cell already
 *         carrying a note or thread, or a mention outside the text. The
 *         sheet is then unchanged.
 */
FM_API fm_status_t fm_sheet_add_threaded_comment(fm_workbook_t* wb, size_t sheet_index,
                                                 const fm_threaded_comment_t* comment);

/**
 * @brief Replaces the text and mentions of the comment with `id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `mentions` is `NULL` with a non-zero count;
 *         `FM_STATUS_THREADED_COMMENT_INVALID` (9006) when no comment has `id` or a
 *         mention is invalid.
 */
FM_API fm_status_t fm_sheet_update_threaded_comment(fm_workbook_t* wb, size_t sheet_index, const char* id,
                                                    const char* text, const fm_mention_t* mentions,
                                                    uint32_t mention_count);

/**
 * @brief Sets the resolved state of the thread opened by `thread_id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `thread_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_THREADED_COMMENT_INVALID` (9006) when `thread_id` names no opening
 *         comment.
 */
FM_API fm_status_t fm_sheet_set_thread_resolved(fm_workbook_t* wb, size_t sheet_index, const char* thread_id,
                                                int32_t done);

/**
 * @brief Removes the comment with `id`; removing an opening comment removes
 *        its whole thread.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_THREADED_COMMENT_INVALID` (9006) when no comment has `id`.
 */
FM_API fm_status_t fm_sheet_remove_threaded_comment(fm_workbook_t* wb, size_t sheet_index, const char* id);

/**
 * @brief Number of entries in the workbook's person list.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_person_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads person `idx`, in file order. The strings are
 *        read-scratch-backed views.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_get_person_at(const fm_workbook_t* wb, size_t idx, fm_person_t* out);

/**
 * @brief Appends a person to the workbook's person list.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_THREADED_COMMENT_INVALID` (9006) when the id is malformed or
 *         already listed, or the display name is empty.
 */
FM_API fm_status_t fm_workbook_add_person(fm_workbook_t* wb, const fm_person_t* person);

/**
 * @brief Removes the person with `id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_THREADED_COMMENT_INVALID` (9006) when no person has `id`, or a
 *         threaded comment or mention still names them.
 */
FM_API fm_status_t fm_workbook_remove_person(fm_workbook_t* wb, const char* id);

/* -------------------------------------------------------------------------- */
/* Recalc                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Drives a full incremental recalc using the default function
 *        registry (`eval::default_registry()`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
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
} fm_parallel_recalc_stats_t;

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `thread_count > 8`;
 *         `FM_STATUS_GRAPH_RECALC_REENTRANT` (4005) when called recursively from a recalc
 *         callback (including a scheduler worker);
 *         a `kEval*` / `kGraph*` code on engine failure.
 */
FM_API fm_status_t fm_workbook_recalc_parallel(fm_workbook_t* wb, uint32_t thread_count,
                                               fm_parallel_recalc_stats_t* out_stats);

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`.
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`.
 */
FM_API fm_status_t fm_workbook_set_iterative_enabled(fm_workbook_t* wb, int32_t enabled);

/**
 * @brief Reads the iterative-calculation settings currently stored on a workbook.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_mode` is NULL.
 */
FM_API fm_status_t fm_workbook_get_calc_mode(const fm_workbook_t* wb, int32_t* out_mode);

/**
 * @brief Sets the workbook's calc mode.
 *
 * Plain metadata — does not affect evaluation. The value is a raw signed
 * 32-bit enum ordinal so FFI callers can pass an arbitrary value without
 * first constructing an out-of-domain C enum. Unknown values (outside the
 * documented range) return `FM_STATUS_INVALID_ARGUMENT` (2).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `mode` is not a documented value.
 */
FM_API fm_status_t fm_workbook_set_calc_mode(fm_workbook_t* wb, int32_t mode);

/** @brief `<calcPr refMode>`: the reference notation the host shows formulas in. */
typedef enum {
  FM_REF_MODE_A1 = 0,
  FM_REF_MODE_R1C1 = 1,
} fm_ref_mode_t;

/**
 * @brief Reads `<calcPr fullPrecision>`: 1 for full precision (the default),
 *        0 for "set precision as displayed".
 *
 * Plain metadata like the calc mode: the engine evaluates at full precision
 * whatever this holds.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is NULL.
 */
FM_API fm_status_t fm_workbook_get_full_precision(const fm_workbook_t* wb, int32_t* out);

/** @brief Sets `<calcPr fullPrecision>` (non-zero = full precision). @return `FM_STATUS_OK` (0), or
 * `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is NULL. */
FM_API fm_status_t fm_workbook_set_full_precision(fm_workbook_t* wb, int32_t full_precision);

/** @brief Reads `<calcPr calcOnSave>` (default 1). @return `FM_STATUS_OK` (0), or `FM_STATUS_BINDING_NULL_POINTER`
 * (7001) if `wb` or `out` is NULL. */
FM_API fm_status_t fm_workbook_get_calc_on_save(const fm_workbook_t* wb, int32_t* out);

/** @brief Sets `<calcPr calcOnSave>` (non-zero = recalculate before saving). @return `FM_STATUS_OK` (0), or
 * `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is NULL. */
FM_API fm_status_t fm_workbook_set_calc_on_save(fm_workbook_t* wb, int32_t calc_on_save);

/** @brief Reads `<calcPr refMode>` as an `fm_ref_mode_t`. @return `FM_STATUS_OK` (0), or
 * `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is NULL. */
FM_API fm_status_t fm_workbook_get_ref_mode(const fm_workbook_t* wb, int32_t* out);

/**
 * @brief Sets `<calcPr refMode>` from a raw `fm_ref_mode_t` ordinal.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `mode` is not a documented value.
 */
FM_API fm_status_t fm_workbook_set_ref_mode(fm_workbook_t* wb, int32_t mode);

/** @brief Reads the workbook's 1904 date system flag (`<workbookPr date1904>`). @return `FM_STATUS_OK` (0), or
 * `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is NULL. */
FM_API fm_status_t fm_workbook_get_date1904(const fm_workbook_t* wb, int32_t* out);

/**
 * @brief Sets the 1904 date system flag.
 *
 * Existing date serials keep their numbers, so every date shifts by 1462
 * days; cached results depending on the date system are invalidated.
 *
 * @return `FM_STATUS_OK` (0), or `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is NULL.
 */
FM_API fm_status_t fm_workbook_set_date1904(fm_workbook_t* wb, int32_t date1904);

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is NULL.
 */
FM_API fm_status_t fm_workbook_get_pinned_now(const fm_workbook_t* wb, fm_civil_time_t* out_now, int32_t* out_pinned);

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `now` is NULL;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if any field is out of range.
 */
FM_API fm_status_t fm_workbook_set_pinned_now(fm_workbook_t* wb, const fm_civil_time_t* now);

/**
 * @brief Releases the pin so clock-dependent results follow the host clock
 *        again. Clearing an unpinned workbook succeeds and does nothing.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`.
 */
FM_API fm_status_t fm_workbook_clear_pinned_now(fm_workbook_t* wb);

/* -------------------------------------------------------------------------- */
/* Excel formula compatibility profile                                        */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the workbook's active Excel formula profile id.
 *
 * The returned pointer is a static view into Formulon's profile table and
 * remains valid for the process lifetime. New workbooks default to
 * `win-365-en_US`. The id is `<host>-365-<locale>` with host `mac` or `win`
 * and locale `ja_JP`, `en_US`, `de_DE`, `fr_FR`, `zh_CN`, `ko_KR`, `th_TH`,
 * `es_ES`, `es_MX`, `pt_BR`, `ru_RU`, `zh_TW`, `it_IT` or `nl_NL` (28 ids); the corresponding setter accepts exactly
 * that set.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_profile_id` is NULL.
 */
FM_API fm_status_t fm_workbook_get_excel_profile_id(const fm_workbook_t* wb, const char** out_profile_id);

/**
 * @brief Sets the workbook's Excel formula profile by id.
 *
 * New workbooks use `win-365-en_US` by default. Supported ids are
 * `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,
 * pt_BR,ru_RU,zh_TW,it_IT,nl_NL}`. The `mac-*`
 * ids and `win-365-ja_JP` follow Excel; the other `win-*` ids
 * are estimated from the Mac values plus the Windows host rules.
 *
 * Existing cached formula values are not recomputed until the caller drives
 * `fm_workbook_recalc` or an equivalent partial recalc.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `profile_id` is NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `profile_id` is not a documented profile.
 */
FM_API fm_status_t fm_workbook_set_excel_profile_id(fm_workbook_t* wb, const char* profile_id);

/**
 * @brief Reads the file name the workbook is shown under, extension included
 *        and without a directory; `""` when unset.
 *
 * While a name is set, every getter that returns formula text spells a
 * reference to the workbook's own defined name (stored `[0]!Name`) with it,
 * as Excel's formula bar does: `M5.xlsx!Name`, quoted when the name needs it
 * (`'M5 b-1.xlsx'!Name`). Unset, they return `[0]!Name`. Loading from bytes
 * leaves the name unset. `*out_name` is read-scratch-backed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_name` is NULL.
 */
FM_API fm_status_t fm_workbook_get_file_name(const fm_workbook_t* wb, const char** out_name);

/**
 * @brief Sets the file name `fm_workbook_get_file_name` reads; `""` unsets it.
 *
 * While a name is set, formula text entered through the C ABI that names it
 * refers to this workbook and creates no external link: `<name>!Name` and
 * `'<name>'!Name` are stored as `[0]!Name`, and `[<name>]Sheet1!A1` as
 * `Sheet1!A1`. The book is matched under ASCII case folding. The stored form
 * and the saved files do not depend on the name.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is NULL;
 *         `FM_STATUS_INVALID_ARGUMENT` (2), changing nothing, when `name` does not end in a
 *           workbook extension (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`, `.xltx`,
 *           `.xltm`, `.xlt`, `.xlam`, `.xla`), holds one of `[ ] : / \ ? *`,
 *           is not valid UTF-8, is longer than 218 bytes, or equals the book
 *           name of an external link.
 */
FM_API fm_status_t fm_workbook_set_file_name(fm_workbook_t* wb, const char* name);

/* -------------------------------------------------------------------------- */
/* Sheet protection (per-sheet `<sheetProtection>`)                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Wide POD mirror of `formulon::SheetProtection` (ECMA-376
 *        §18.3.1.85 `<sheetProtection>`).
 *
 * Strings are NUL-terminated UTF-8. On the read path
 * (`fm_sheet_get_sheet_protection`) the pointers are model-backed views into the
 * workbook's own storage and remain valid until the next mutation of the
 * same sheet's protection record or handle destruction. Reads do not
 * invalidate them. On the write path
 * (`fm_sheet_set_sheet_protection`) the strings are copied into the
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is NULL;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_sheet_protection(const fm_workbook_t* wb, size_t sheet_index,
                                                 fm_sheet_protection_t* out);

/**
 * @brief Sets the sheet's protection state.
 *
 * Replaces the per-sheet `SheetProtection` record wholesale. Strings
 * are deep-copied into the workbook; NULL strings are stored as
 * empty.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `in` is NULL;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_sheet_protection(fm_workbook_t* wb, size_t sheet_index,
                                                 const fm_sheet_protection_t* in);

/**
 * @brief Viewport rectangle on one sheet, expressed in 0-based
 *        inclusive coordinates. Used by `fm_sheet_recalc_viewport`.
 *
 * `last_row` / `last_col` are inclusive — a single-cell viewport sets
 * the corresponding `first_*` and `last_*` to the same value. An empty
 * viewport (the row or column range collapsed) is allowed and produces
 * a no-op recalc.
 */
typedef struct {
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
 * `fm_workbook_recalc` (or a `fm_sheet_recalc_viewport` whose closure
 * overlaps them) will visit them.
 *
 * @param wb                     Workbook handle. Must be non-NULL.
 * @param sheet_index            0-based sheet the viewport lies on.
 * @param viewport               Viewport rectangle. Must be non-NULL.
 * @param out_recomputed_count   Optional. Receives the number of cells
 *                               actually recomputed during this call.
 *                               May be NULL if the caller does not
 *                               care.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `viewport` is NULL;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is past the last sheet;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `sheet_index` exceeds the 16-bit range a sheet id is stored in;
 *         a `kEval*` / `kGraph*` code on engine failure.
 */
FM_API fm_status_t fm_sheet_recalc_viewport(fm_workbook_t* wb, size_t sheet_index, const fm_viewport* viewport,
                                            uint32_t* out_recomputed_count);

/**
 * @brief Sets the iterative-solver progress callback for subsequent
 *        recalcs on this workbook.
 *
 * Pass `cb == NULL` to clear the callback. `user_data` is forwarded
 * verbatim to every invocation; the workbook does not take ownership
 * of it.
 *
 * The callback runs inside the recalc and must not call any `fm_*`
 * function on the same workbook: a nested recalc returns
 * `FM_STATUS_GRAPH_RECALC_REENTRANT` (4005), and any other call may block or observe a
 * half-computed sweep.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`.
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
 * @brief The icon drawn for one level of a custom icon set (x14 `cfIcon`).
 *
 * `icon_set` is the `ST_IconSetType` spelling of the set the icon comes
 * from: a built-in name such as `"3Arrows"`, `"3Triangles"`, `"3Stars"`,
 * `"5Boxes"`, or `"NoIcons"` for a level that draws nothing (`icon_id` 0).
 * `icon_id` is the zero-based icon within that set.
 */
typedef struct {
  const char* icon_set;
  int32_t icon_id;
} fm_cf_icon_t;

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
 *     `bar_border` when engaged), `bar_gradient`, and `bar_direction`
 *     (the rule's `data_bar_direction`; the geometry is stated from the
 *     left edge, and a host mirrors it for a right-to-left bar).
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
  int32_t kind; /* fm_cf_match_kind_t */
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
  uint8_t bar_direction; /* 0=context, 1=leftToRight, 2=rightToLeft */
  uint8_t _pad[2];       /* padding for alignment determinism */
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
 * Fails with `FM_STATUS_PRINT_PAGE_COUNT_OVERFLOW` (9005) when the page grid the sheet
 * declares exceeds the supported page count, which a file listing a manual
 * break before nearly every row and column can reach. On success the count
 * `fm_pagination_get_page_count` reports is the true total; it is never a
 * truncated one.
 */
FM_API fm_status_t fm_sheet_get_pagination(const fm_workbook_t* wb, size_t sheet_index, fm_pagination_t** out);

/** @brief Releases a pagination result. `pagination == NULL` is a no-op. */
FM_API void fm_pagination_destroy(fm_pagination_t* pagination);

/** @brief Writes the total physical page count to `*out_count`;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) for a NULL argument. */
FM_API fm_status_t fm_pagination_get_page_count(const fm_pagination_t* pagination, size_t* out_count);

/** @brief Writes the number of rectangles in the sheet's declared
 *         `_xlnm.Print_Area` to `*out_count`;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) for a NULL argument.
 *
 * Zero also means the sheet declares no print area, mirroring Excel's own
 * `PageSetup.PrintArea`; it is not backfilled with the used range. Pagination
 * still falls back to the used range in that case, so
 * `fm_pagination_get_page_count` can be non-zero while this is zero. */
FM_API fm_status_t fm_pagination_get_print_area_count(const fm_pagination_t* pagination, size_t* out_count);

/** @brief Copies a resolved print-area rectangle by index. */
FM_API fm_status_t fm_pagination_get_print_area_at(const fm_pagination_t* pagination, size_t index,
                                                   fm_print_range_t* out);

/** @brief Writes the number of horizontal (row) page breaks to `*out_count`;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) for a NULL argument. */
FM_API fm_status_t fm_pagination_get_row_break_count(const fm_pagination_t* pagination, size_t* out_count);

/** @brief Reads a 0-based row index that a horizontal break precedes. */
FM_API fm_status_t fm_pagination_get_row_break_at(const fm_pagination_t* pagination, size_t index, uint32_t* out_row);

/** @brief Writes the number of vertical (column) page breaks to `*out_count`;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) for a NULL argument. */
FM_API fm_status_t fm_pagination_get_col_break_count(const fm_pagination_t* pagination, size_t* out_count);

/** @brief Reads a 0-based column index that a vertical break precedes. */
FM_API fm_status_t fm_pagination_get_col_break_at(const fm_pagination_t* pagination, size_t index, uint32_t* out_col);

/** @brief A rectangle in points. Sheet geometry measures from the top-left
 *         of cell A1; pagination measures from the page's top-left corner. */
typedef struct {
  double x;
  double y;
  double width;
  double height;
} fm_rect_pt_t;

/** @brief The paper a sheet prints on, after the orientation swap. */
typedef struct {
  double width_pt;
  double height_pt;
  int32_t landscape; /* 0/1 */
  int32_t known;     /* 0 when `paperSize` was unrecognised and A4 substituted */
} fm_paper_info_t;

/** @brief Page margins in points; `header` / `footer` are measured from the
 *         page edge to the header / footer band. */
typedef struct {
  double left;
  double right;
  double top;
  double bottom;
  double header;
  double footer;
} fm_margins_pt_t;

/** @brief The repeat-rows / repeat-columns in effect (0-based, inclusive).
 *         A span is meaningful only when its `has_*` flag is non-zero. */
typedef struct {
  int32_t has_rows; /* 0/1 */
  uint32_t first_row;
  uint32_t last_row;
  int32_t has_cols; /* 0/1 */
  uint32_t first_col;
  uint32_t last_col;
} fm_print_titles_t;

/**
 * @brief One physical page: the cell block it carries and where it sits.
 *
 * `area_index` indexes `fm_pagination_get_print_area_at` (0 for the used-range
 * fallback). `width_pt` / `height_pt` are the scaled extent of the page's
 * own cell block, repeated titles excluded; `origin_*_pt` is that block's
 * top-left corner on the paper.
 */
typedef struct {
  uint32_t area_index;
  uint32_t first_row;
  uint32_t last_row;
  uint32_t first_col;
  uint32_t last_col;
  double origin_x_pt;
  double origin_y_pt;
  double width_pt;
  double height_pt;
} fm_page_layout_t;

/**
 * @brief Reads the paper size and orientation pagination laid out on.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_paper(const fm_pagination_t* pagination, fm_paper_info_t* out);

/**
 * @brief Reads the page margins in points.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_margins(const fm_pagination_t* pagination, fm_margins_pt_t* out);

/**
 * @brief Reads the cell-body rectangle before print-title reservation, in
 *        points from the page's top-left corner.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_printable(const fm_pagination_t* pagination, fm_rect_pt_t* out);

/**
 * @brief Reads the effective scale as a factor (1.0 is 100%): the `scale`
 *        percentage, or the fit-to-page factor.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_scale(const fm_pagination_t* pagination, double* out_scale);

/**
 * @brief Reads the page order: `0` down-then-over (the OOXML default), `1`
 *        over-then-down.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_page_order(const fm_pagination_t* pagination, int32_t* out_order);

/**
 * @brief Reads the print-title rows / columns repeated on every page.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_print_titles(const fm_pagination_t* pagination, fm_print_titles_t* out);

/**
 * @brief Reads physical page `index`, in print order.
 *
 * The index domain is `[0, fm_pagination_get_page_count)`: one entry per
 * physical page, ordered by the page order and, across several print
 * areas, by declaration order.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_pagination_get_page_at(const fm_pagination_t* pagination, size_t index, fm_page_layout_t* out);

/**
 * @brief Reports whether horizontal break `index` (the domain of
 *        `fm_pagination_get_row_break_at`) is a manual break (`1`) rather
 *        than one the page walk computed (`0`).
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_pagination_is_row_break_manual(const fm_pagination_t* pagination, size_t index,
                                                     int32_t* out_manual);

/**
 * @brief Vertical-break counterpart of
 *        `fm_pagination_is_row_break_manual`.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_pagination_is_col_break_manual(const fm_pagination_t* pagination, size_t index,
                                                     int32_t* out_manual);

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
 * `*out_xml` is a read-scratch-backed view: valid until the next successful
 * scratch-backed read, any mutation, or handle destruction on the same
 * handle.
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
 * from the new fragment immediately, so a `fm_sheet_get_pagination` that
 * follows observes the new settings.
 *
 * `fitToPage` is NOT part of `<pageSetup>` - it lives in
 * `<sheetPr><pageSetUpPr>`; use `fm_sheet_set_fit_to_page`.
 *
 * Returns `FM_STATUS_INVALID_ARGUMENT` (2) when the fragment is malformed, carries
 * more than one top-level element, or names the wrong root; and when it
 * declares an `r:id` on a sheet that has no printer-settings part, which
 * would leave a dangling relationship that makes Excel repair the file.
 * Returns `FM_STATUS_PRECONDITION_FAILED` (9) when the fragment exceeds 4 KiB.
 */
FM_API fm_status_t fm_sheet_set_page_setup_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<pageMargins>` as raw XML. `*out_xml` is a read-scratch-backed
 *         view; see `fm_sheet_get_page_setup_xml` for the shared contract. */
FM_API fm_status_t fm_sheet_get_page_margins_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<pageMargins>`. Re-derives the structured margins; an
 *         empty string removes the element and restores the defaults.
 *         Fragment ceiling 4 KiB. */
FM_API fm_status_t fm_sheet_set_page_margins_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<printOptions>` as raw XML. `*out_xml` is a read-scratch-backed
 *         view; see `fm_sheet_get_page_setup_xml` for the shared contract. */
FM_API fm_status_t fm_sheet_get_print_options_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<printOptions>`. This element has no structured view;
 *         the raw fragment is the whole model. Fragment ceiling 4 KiB. */
FM_API fm_status_t fm_sheet_set_print_options_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<headerFooter>` as raw XML. `*out_xml` is a read-scratch-backed
 *         view; see `fm_sheet_get_page_setup_xml` for the shared contract. */
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

/** @brief Reads `<sheetPr>` as raw XML. `*out_xml` is a read-scratch-backed
 *         view; see `fm_sheet_get_page_setup_xml` for the shared contract.
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
 * An empty string removes the defined name. Returns `FM_STATUS_PRINT_INVALID_AREA` (9002)
 * (9002) when any area fails to parse, and `FM_STATUS_INVALID_ARGUMENT` (2) when
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
 * `*out_ranges_a1` is a read-scratch-backed view.
 *
 * Returns `FM_STATUS_PRINT_INVALID_AREA` (9002) when the defined name exists but does
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
 * Returns `FM_STATUS_PRINT_INVALID_AREA` (9002) when either span is present but not a
 * whole-row / whole-column span of the right shape.
 */
FM_API fm_status_t fm_sheet_set_print_titles(fm_workbook_t* wb, size_t sheet_index, const char* repeat_rows,
                                             const char* repeat_cols);

/**
 * @brief Reads the repeat rows / columns back as `"1:2"` / `"A:A"`, each
 *        an empty string when that axis is unset.
 *
 * Both outputs are read-scratch-backed views and stay valid together until
 * the next scratch-backed read on the handle.
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
 * Returns `FM_STATUS_INVALID_ARGUMENT` (2) for a row past the grid, and
 * `FM_STATUS_PRECONDITION_FAILED` (9) once the axis holds 1026 breaks. A row break
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
 * Returns `FM_STATUS_INVALID_ARGUMENT` (2) for a column past the grid, and
 * `FM_STATUS_PRECONDITION_FAILED` (9) once the axis holds 1026 breaks. A column
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
 * @brief Writes the number of row breaks declared on the sheet to
 *        `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is NULL;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_break_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Copies the row break at `index` (ascending by row).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                             fm_page_break_t* out);

/**
 * @brief Writes the number of column breaks declared on the sheet to
 *        `*out_count`. Same status contract as `fm_sheet_get_row_break_count`.
 */
FM_API fm_status_t fm_sheet_get_col_break_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Copies the column break at `index` (ascending by column).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_col_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
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
 * `scale` outside [10, 400] is rejected with `FM_STATUS_INVALID_ARGUMENT` (2) rather
 * than clamped. `fm_sheet_set_zoom` clamps its own input because an
 * out-of-range screen zoom is harmless, but a mis-stated print scale lands
 * on paper, so silently rounding it would hide the mistake.
 */
FM_API fm_status_t fm_sheet_set_page_setup(fm_workbook_t* wb, size_t sheet_index, const fm_page_setup_t* setup);

/** @brief Applies a partial `<pageMargins>` update. A negative, infinite or
 *         NaN margin is rejected with `FM_STATUS_INVALID_ARGUMENT` (2). */
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
 * Owns the per-cell match lists produced by `fm_sheet_evaluate_cf_range`.
 * Index into it via `fm_cf_results_get_cell_count` /
 * `fm_cf_results_get_cell_at` / `fm_cf_results_get_match_at`. Release with
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 *
 * On success the caller owns `*out` and must free it with
 * `fm_cf_results_destroy`.
 */
FM_API fm_status_t fm_sheet_evaluate_cf_range(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                              uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                              double today_serial, fm_cf_results_t** out);

/**
 * @brief Releases a results handle. `results == NULL` is a no-op.
 */
FM_API void fm_cf_results_destroy(fm_cf_results_t* results);

/**
 * @brief Writes the number of cells in the result that produced at least
 *        one match to `*out_count`. `FM_STATUS_BINDING_NULL_POINTER` (7001)
 *        when `results` or `out_count` is NULL.
 */
FM_API fm_status_t fm_cf_results_get_cell_count(const fm_cf_results_t* results, size_t* out_count);

/**
 * @brief Reads cell `cell_idx`'s coordinates and match count.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cell_idx >= cell_count`.
 */
FM_API fm_status_t fm_cf_results_get_cell_at(const fm_cf_results_t* results, size_t cell_idx, uint32_t* out_row,
                                             uint32_t* out_col, size_t* out_match_count);

/**
 * @brief Reads match `match_idx` for cell `cell_idx`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when either index is out of range.
 */
FM_API fm_status_t fm_cf_results_get_match_at(const fm_cf_results_t* results, size_t cell_idx, size_t match_idx,
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
 * copied by `fm_sheet_add_cf_rule`; pass `NULL` for valueless min/max
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
 * `fm_sheet_add_cf_rule` deep-copies every pointer-backed payload into
 * the engine model. `fm_sheet_get_cf_rule` returns every pointer in the
 * selected rule -- id, sqref, formula/text strings, and the visual
 * threshold arrays and their value strings alike -- as a view into
 * handle-owned storage, not the workbook model; see that function's own
 * doc for the full lifetime contract (valid until the next successful
 * `fm_sheet_get_cf_rule` on the same handle, or handle destruction; CF
 * mutations, unrelated reads, and a validation-rejected call do not
 * invalidate it).
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
 *     `icon_set_thresholds`, and the `icon_set_floor` at the end of this
 *     struct.
 *
 * String fields use C-string convention: `NULL` means "absent",
 * non-`NULL` is a NUL-terminated view. On the input path
 * (`fm_sheet_add_cf_rule`) the caller owns the buffer until the call
 * returns; on the output path (`fm_sheet_get_cf_rule`) every field is a
 * view into handle-owned storage per that function's lifetime contract
 * (see its doc), not the workbook model.
 */
typedef struct {
  /* Stable rule id, linking the rule to the `<x14:cfRule id="...">`
   * block that carries whatever settings the legacy OOXML schema cannot
   * express. On input, pass `NULL` or `""` to auto-generate one; a
   * supplied id must be GUID-shaped, since that is the schema type of
   * the attribute it becomes. On output, always a non-NULL view into
   * `fm_sheet_get_cf_rule`'s handle-owned storage (see that function's doc).
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
   * sqref_count >= 1. On output, points to `fm_sheet_get_cf_rule`'s
   * handle-owned storage (see that function's doc), not the model. */
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
   * `fm_sheet_get_cf_rule` engages all six, so feeding its output straight
   * back into `fm_sheet_add_cf_rule` reproduces the rule exactly. */
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

  /* Appended after the extension block for the same reason: every earlier
   * field keeps its offset.
   *
   * `icon_set_floor` is the icon set's first threshold, the lower bound of
   * the lowest icon's bucket: a cell below it gets no icon at all. When
   * `icon_set_floor_engaged` is 0 the floor is Excel's default, `percent 0`,
   * which every value clears. `fm_sheet_get_cf_rule` always engages it.
   *
   * `data_bar_direction` is the edge the bar grows from: 0=context (follows
   * the sheet's reading direction; the default), 1=leftToRight,
   * 2=rightToLeft. */
  int32_t icon_set_floor_engaged; /* 0/1 */
  fm_cfvo_t icon_set_floor;
  uint8_t data_bar_direction;
  uint8_t _pad4[3];

  /* Custom icons of the icon set (x14 `cfIcon`), appended for the same
   * reason: one entry per icon of the set, or none when the set draws its
   * own icons. `icon_set_icons` may be NULL only with a zero count. On
   * output it points into `fm_sheet_get_cf_rule`'s handle-owned storage. */
  const fm_cf_icon_t* icon_set_icons;
  uint32_t icon_set_icon_count;
} fm_cf_rule_t;

/**
 * @brief Returns the total number of CF rules on `sheet_index`.
 *
 * The count flattens across all `<conditionalFormatting>` blocks: a
 * sheet with two blocks of two rules each reports `4`. Indexing into
 * the flattened sequence is stable until the next mutation.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cf_rule_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `idx`-th CF rule (in flattened order) into `out`.
 *
 * Every pointer in `*out` - `id`, `sqref`, `formula1`, `formula2`, `text`,
 * the visual threshold and colour arrays, and the threshold value strings -
 * addresses storage owned by `wb`, not the workbook model. That storage is
 * released only by the next successful `fm_sheet_get_cf_rule` on the same
 * handle, or by handle destruction. CF mutations
 * (`fm_sheet_add_cf_rule`, `fm_sheet_remove_cf_rule`, `fm_sheet_clear_cf_rules`),
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
 * `fm_workbook_get_dxf` would reject and `fm_sheet_add_cf_rule` would
 * refuse. The record therefore always feeds straight back into
 * `fm_sheet_add_cf_rule`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cf_rule(const fm_workbook_t* wb, size_t sheet_index, size_t idx, fm_cf_rule_t* out);

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
 * flattened CF rule sequence (the same indexing `fm_sheet_get_cf_rule`
 * and `fm_sheet_remove_cf_rule` use). Since this call always appends a
 * new block after every existing one, the returned index equals the
 * flattened rule count observed just before the call; it stays valid
 * until a subsequent mutation (`fm_sheet_add_cf_rule`,
 * `fm_sheet_remove_cf_rule`, `fm_sheet_clear_cf_rules`) on the same sheet
 * renumbers the sequence.
 *
 * When `rule.dxf_id_engaged` is non-zero, `rule.dxf_id` must already name a
 * registered differential format (`fm_workbook_add_dxf`). A rule referencing a
 * missing `<dxf>` is what makes Excel discard every conditional format in
 * the workbook on open, so the reference is validated here rather than at
 * save time.
 *
 * A `DataBar` payload is rejected when either length percentage exceeds
 * 100, when `data_bar_min_length_pct > data_bar_max_length_pct` (an
 * inverted bar Excel cannot render), or when an engaged
 * `data_bar_axis_position` is outside `0..2`. An `IconSet` payload is
 * rejected unless it carries exactly N-1 `icon_set_thresholds` for its
 * N-icon set (the floor is `icon_set_floor`, not a threshold).
 *
 * `rule.formula1` / `rule.formula2` and formula-typed `fm_cfvo_t` values
 * may carry one leading `=`, which is dropped: the rule stores them as
 * files carry them, spelled as `fm_sheet_set_formula` stores a cell
 * formula. For the text (7-10), blanks / errors (11-14) and
 * `TimePeriod` (15) types, a `NULL` or empty `rule.formula1` is filled
 * with the formula Excel stores for the rule, relative to the top-left
 * corner of the bounding box of `rule.sqref` (e.g. `NOT(ISERROR(SEARCH("ab",B1)))`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out_index` is `NULL`, or
 *           `rule.sqref` is `NULL` while `rule.sqref_count > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `rule.sqref_count == 0` or exceeds the per-call range cap,
 *           when any `rule.sqref` entry is inverted (`first` past
 *           `last`) or reaches outside the grid (`row >= 1048576`,
 *           `col >= 16384`) and therefore has no A1 spelling, when
 *           `rule.id` is non-empty but not GUID-shaped, when `rule.type`
 *           / `rule.op` / `rule.time_period` / `rule.icon_set_name` /
 *           `rule.data_bar_axis_position` or any engaged `fm_cfvo_t`
 *           `type` names no declared enumerator, when `rule.dxf_id` is
 *           engaged but past the end of the workbook's `<dxfs>` table,
 *           when a visual payload is missing / malformed, or when
 *           `rule.priority <= 0` and an existing rule already holds
 *           `INT32_MAX`, so the auto-assigned priority would overflow;
 *         `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000) when a non-empty `rule.formula1`,
 *           `rule.formula2` or formula-typed `fm_cfvo_t` value does not
 *           parse (the rule is not stored).
 */
FM_API fm_status_t fm_sheet_add_cf_rule(fm_workbook_t* wb, size_t sheet_index, fm_cf_rule_t rule, size_t* out_index);

/**
 * @brief Removes the `idx`-th CF rule (flattened order). When the
 *        containing block becomes empty it is removed too.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_cf_rule(fm_workbook_t* wb, size_t sheet_index, size_t idx);

/**
 * @brief Removes every CF block on `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_cf_rules(fm_workbook_t* wb, size_t sheet_index);

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
 * Owns the cell-node list produced by `fm_sheet_get_precedents` /
 * `fm_sheet_get_dependents`. Index into it via `fm_cell_nodes_get_node_count` /
 * `fm_cell_nodes_get_node_at`. Release with `fm_cell_nodes_destroy`.
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 *
 * On success the caller owns `*out` and must free it with
 * `fm_cell_nodes_destroy`.
 */
FM_API fm_status_t fm_sheet_get_precedents(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                           uint32_t depth, fm_cell_nodes_t** out);

/**
 * @brief Returns the cells that read `(sheet, row, col)` directly
 *        (1-step dependents) when `depth <= 1`, or every cell reached
 *        within `depth` BFS steps in the reverse graph when
 *        `depth >= 2`. `depth` is capped at 32.
 *
 * Same semantics as `fm_sheet_get_precedents` for ownership and seed
 * exclusion.
 */
FM_API fm_status_t fm_sheet_get_dependents(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                           uint32_t depth, fm_cell_nodes_t** out);

/**
 * @brief Releases a trace results handle. `nodes == NULL` is a no-op.
 */
FM_API void fm_cell_nodes_destroy(fm_cell_nodes_t* nodes);

/**
 * @brief Writes the number of cells in the result to `*out_count`.
 *        `FM_STATUS_BINDING_NULL_POINTER` (7001) when `nodes` or
 *        `out_count` is NULL.
 */
FM_API fm_status_t fm_cell_nodes_get_node_count(const fm_cell_nodes_t* nodes, size_t* out_count);

/**
 * @brief Reads cell `idx` from the result.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_cell_nodes_get_node_at(const fm_cell_nodes_t* nodes, size_t idx, fm_cell_node_t* out);

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
  int32_t kind; /* fm_pivot_cell_kind_t */
  uint32_t depth;
  const char* field_name;
  const char* number_format;
} fm_pivot_cell_t;

/**
 * @brief Opaque handle for a projected PivotTable layout.
 *
 * Owns the projected cell list and every string pointer reachable from
 * `fm_pivot_cell_t`. Index into it via `fm_pivot_cells_get_cell_count` /
 * `fm_pivot_cells_get_cell_at`. Release with `fm_pivot_cells_destroy`.
 */
typedef struct fm_pivot_cells fm_pivot_cells_t;

/**
 * @brief Returns the number of pivot tables anchored on `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out_count` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Evaluates and projects the `pivot_index`-th PivotTable on a sheet.
 *
 * On success the caller owns `*out` and must free it with
 * `fm_pivot_cells_destroy`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `pivot_index` is out of range;
 *         a pivot evaluation/layout error when the table references a missing
 *         or invalid cache.
 */
FM_API fm_status_t fm_sheet_get_pivot_layout(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                             fm_pivot_cells_t** out);

/**
 * @brief Releases a pivot layout handle. `cells == NULL` is a no-op.
 */
FM_API void fm_pivot_cells_destroy(fm_pivot_cells_t* cells);

/**
 * @brief Writes the number of projected cells to `*out_count`.
 *        `FM_STATUS_BINDING_NULL_POINTER` (7001) when `cells` or
 *        `out_count` is NULL.
 */
FM_API fm_status_t fm_pivot_cells_get_cell_count(const fm_pivot_cells_t* cells, size_t* out_count);

/**
 * @brief Returns the projected layout bounds.
 *
 * `top` / `left` are absolute 0-based sheet coordinates; `rows` / `cols`
 * are the rectangular layout span.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_pivot_cells_get_bounds(const fm_pivot_cells_t* cells, uint32_t* out_top, uint32_t* out_left,
                                             uint32_t* out_rows, uint32_t* out_cols);

/**
 * @brief Reads projected cell `idx`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_pivot_cells_get_cell_at(const fm_pivot_cells_t* cells, size_t idx, fm_pivot_cell_t* out);

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
 * @brief Sibling item order of a pivot field.
 *
 * Mirrors OOXML `sortType` (`descending` / `ascending` / `manual`).
 */
typedef enum {
  FM_PIVOT_SORT_DESCENDING = 0,
  FM_PIVOT_SORT_ASCENDING = 1,
  FM_PIVOT_SORT_MANUAL = 2
} fm_pivot_sort_order_t;

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
  FM_PIVOT_SHOW_VALUES_AS_NORMAL = 0,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_ROW = 1,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_COL = 2,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_TOTAL = 3,
  FM_PIVOT_SHOW_VALUES_AS_RUNNING_TOTAL_IN_ROW = 4,
  FM_PIVOT_SHOW_VALUES_AS_RUNNING_TOTAL_IN_COL = 5,
  FM_PIVOT_SHOW_VALUES_AS_INDEX = 6,
  FM_PIVOT_SHOW_VALUES_AS_DIFFERENCE_FROM = 7,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_DIFFERENCE_FROM = 8,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_PARENT_ROW = 9,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_PARENT_COL = 10,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_PARENT = 11,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF = 12,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_RUNNING_TOTAL = 13,
  FM_PIVOT_SHOW_VALUES_AS_RANK_ASCENDING = 14,
  FM_PIVOT_SHOW_VALUES_AS_RANK_DESCENDING = 15
} fm_pivot_show_values_as_t;

/**
 * @brief Filter type of a pivot label / value / date filter: every
 *        `ST_PivotFilterType` member but `unknown`. Numbering mirrors
 *        `formulon::pivot::FilterType`.
 *
 *   * Value types (0-2, 6-11, 24-28) take a number and score
 *     `data_field_index`. The top / bottom types keep items until the count,
 *     the percent of the field total, or the running sum is reached.
 *   * Label types (3, 4, 12-23) take text as entered; `*` and `?` are
 *     wildcards and `~` escapes them. `LABEL_BETWEEN` / `LABEL_NOT_BETWEEN`
 *     take their upper bound in `value_high_text`.
 *   * Date types (5, 29-35) take Excel date serials; `LABEL_DATE` is the
 *     inclusive date range.
 *   * Relative periods (36-51) and calendar positions (52-67) take no
 *     operand.
 */
typedef enum {
  FM_PIVOT_FILTER_VALUE_TOP_10 = 0,
  FM_PIVOT_FILTER_VALUE_GREATER_THAN = 1,
  FM_PIVOT_FILTER_VALUE_BETWEEN = 2,
  FM_PIVOT_FILTER_LABEL_CONTAINS = 3,
  FM_PIVOT_FILTER_LABEL_BEGINS_WITH = 4,
  FM_PIVOT_FILTER_LABEL_DATE = 5,
  FM_PIVOT_FILTER_VALUE_EQUAL = 6,
  FM_PIVOT_FILTER_VALUE_NOT_EQUAL = 7,
  FM_PIVOT_FILTER_VALUE_GREATER_THAN_OR_EQUAL = 8,
  FM_PIVOT_FILTER_VALUE_LESS_THAN = 9,
  FM_PIVOT_FILTER_VALUE_LESS_THAN_OR_EQUAL = 10,
  FM_PIVOT_FILTER_VALUE_NOT_BETWEEN = 11,
  FM_PIVOT_FILTER_LABEL_EQUAL = 12,
  FM_PIVOT_FILTER_LABEL_NOT_EQUAL = 13,
  FM_PIVOT_FILTER_LABEL_NOT_BEGINS_WITH = 14,
  FM_PIVOT_FILTER_LABEL_ENDS_WITH = 15,
  FM_PIVOT_FILTER_LABEL_NOT_ENDS_WITH = 16,
  FM_PIVOT_FILTER_LABEL_NOT_CONTAINS = 17,
  FM_PIVOT_FILTER_LABEL_GREATER_THAN = 18,
  FM_PIVOT_FILTER_LABEL_GREATER_THAN_OR_EQUAL = 19,
  FM_PIVOT_FILTER_LABEL_LESS_THAN = 20,
  FM_PIVOT_FILTER_LABEL_LESS_THAN_OR_EQUAL = 21,
  FM_PIVOT_FILTER_LABEL_BETWEEN = 22,
  FM_PIVOT_FILTER_LABEL_NOT_BETWEEN = 23,
  FM_PIVOT_FILTER_VALUE_BOTTOM_COUNT = 24,
  FM_PIVOT_FILTER_VALUE_TOP_PERCENT = 25,
  FM_PIVOT_FILTER_VALUE_BOTTOM_PERCENT = 26,
  FM_PIVOT_FILTER_VALUE_TOP_SUM = 27,
  FM_PIVOT_FILTER_VALUE_BOTTOM_SUM = 28,
  FM_PIVOT_FILTER_DATE_EQUAL = 29,
  FM_PIVOT_FILTER_DATE_NOT_EQUAL = 30,
  FM_PIVOT_FILTER_DATE_OLDER_THAN = 31,
  FM_PIVOT_FILTER_DATE_OLDER_THAN_OR_EQUAL = 32,
  FM_PIVOT_FILTER_DATE_NEWER_THAN = 33,
  FM_PIVOT_FILTER_DATE_NEWER_THAN_OR_EQUAL = 34,
  FM_PIVOT_FILTER_DATE_NOT_BETWEEN = 35,
  FM_PIVOT_FILTER_DATE_TOMORROW = 36,
  FM_PIVOT_FILTER_DATE_TODAY = 37,
  FM_PIVOT_FILTER_DATE_YESTERDAY = 38,
  FM_PIVOT_FILTER_DATE_NEXT_WEEK = 39,
  FM_PIVOT_FILTER_DATE_THIS_WEEK = 40,
  FM_PIVOT_FILTER_DATE_LAST_WEEK = 41,
  FM_PIVOT_FILTER_DATE_NEXT_MONTH = 42,
  FM_PIVOT_FILTER_DATE_THIS_MONTH = 43,
  FM_PIVOT_FILTER_DATE_LAST_MONTH = 44,
  FM_PIVOT_FILTER_DATE_NEXT_QUARTER = 45,
  FM_PIVOT_FILTER_DATE_THIS_QUARTER = 46,
  FM_PIVOT_FILTER_DATE_LAST_QUARTER = 47,
  FM_PIVOT_FILTER_DATE_NEXT_YEAR = 48,
  FM_PIVOT_FILTER_DATE_THIS_YEAR = 49,
  FM_PIVOT_FILTER_DATE_LAST_YEAR = 50,
  FM_PIVOT_FILTER_DATE_YEAR_TO_DATE = 51,
  FM_PIVOT_FILTER_DATE_Q1 = 52,
  FM_PIVOT_FILTER_DATE_Q2 = 53,
  FM_PIVOT_FILTER_DATE_Q3 = 54,
  FM_PIVOT_FILTER_DATE_Q4 = 55,
  FM_PIVOT_FILTER_DATE_M1 = 56,
  FM_PIVOT_FILTER_DATE_M2 = 57,
  FM_PIVOT_FILTER_DATE_M3 = 58,
  FM_PIVOT_FILTER_DATE_M4 = 59,
  FM_PIVOT_FILTER_DATE_M5 = 60,
  FM_PIVOT_FILTER_DATE_M6 = 61,
  FM_PIVOT_FILTER_DATE_M7 = 62,
  FM_PIVOT_FILTER_DATE_M8 = 63,
  FM_PIVOT_FILTER_DATE_M9 = 64,
  FM_PIVOT_FILTER_DATE_M10 = 65,
  FM_PIVOT_FILTER_DATE_M11 = 66,
  FM_PIVOT_FILTER_DATE_M12 = 67
} fm_pivot_filter_type_t;

/**
 * @brief The Values pseudo-field (the "Σ Values" level) in a row / column
 *        field order, as `fm_sheet_get_pivot_row_field_order` returns it
 *        and `fm_sheet_set_pivot_row_field_order` accepts it. Its bits
 *        match the OOXML `x="-2"` field reference.
 */
#define FM_PIVOT_VALUES_FIELD 0xFFFFFFFEu

/**
 * @brief Calendar periods of a date group, as bits of the `periods_mask` of
 *        `fm_sheet_set_pivot_field_date_group` and the `date_periods` of
 *        `fm_pivot_field_info_t`.
 */
#define FM_PIVOT_PERIOD_SECONDS 1u
#define FM_PIVOT_PERIOD_MINUTES 2u
#define FM_PIVOT_PERIOD_HOURS 4u
#define FM_PIVOT_PERIOD_DAYS 8u
#define FM_PIVOT_PERIOD_MONTHS 16u
#define FM_PIVOT_PERIOD_QUARTERS 32u
#define FM_PIVOT_PERIOD_YEARS 64u

/**
 * @brief Calendar a date group's years follow: Gregorian years, or Japanese
 *        era years. Numbering mirrors `formulon::pivot::CalendarSystem`.
 */
typedef enum { FM_PIVOT_CALENDAR_GREGORIAN = 0, FM_PIVOT_CALENDAR_JAPANESE = 1 } fm_pivot_calendar_t;

/**
 * @brief Pivot report layout form. Numbering mirrors
 *        `formulon::pivot::PivotLayout`.
 */
typedef enum {
  FM_PIVOT_REPORT_LAYOUT_COMPACT = 0,
  FM_PIVOT_REPORT_LAYOUT_TABULAR = 1,
  FM_PIVOT_REPORT_LAYOUT_OUTLINE = 2
} fm_pivot_report_layout_t;

/**
 * @brief Discriminator for the variant payload carried by a pivot filter
 *        spec. `FM_PIVOT_FILTER_VALUE_NONE` (= -1) means the slot is unset:
 *        the payload of a relative-period or calendar-position type, or an
 *        absent upper bound. The kinds mirror
 *        `std::variant<std::monostate, int, double, std::string>`.
 */
typedef enum {
  FM_PIVOT_FILTER_VALUE_NONE = -1,
  FM_PIVOT_FILTER_VALUE_INT = 0,
  FM_PIVOT_FILTER_VALUE_DOUBLE = 1,
  FM_PIVOT_FILTER_VALUE_TEXT = 2
} fm_pivot_filter_value_kind_t;

/**
 * @brief Plain-data spec for `fm_sheet_add_pivot_field`.
 *
 *   * `source_name`   — required; matches a `PivotCacheField::name`.
 *   * `custom_name`   — nullable; pass `NULL` or the empty string for none.
 *   * `axis`          — initial axis for the field.
 *   * `subtotal_top`  — non-zero to render subtotals at the group head. Excel
 *                       and OOXML (`subtotalTop`) default to 1; a
 *                       zero-initialised spec asks for subtotals at the
 *                       bottom, so set it explicitly.
 *   * `number_format` — nullable; the field's `numFmtId` as a decimal string,
 *                       e.g. an id returned by `fm_workbook_add_num_fmt` or a
 *                       built-in id. `NULL` or `""` means none; any other
 *                       string that is not a built-in or registered id is
 *                       `FM_STATUS_INVALID_ARGUMENT` (2).
 */
typedef struct {
  const char* source_name;
  const char* custom_name;
  int32_t axis; /* fm_pivot_axis_t */
  int32_t subtotal_top;
  const char* number_format;
} fm_pivot_field_spec_t;

/**
 * @brief Plain-data spec for `fm_sheet_add_pivot_data_field` /
 *        `fm_sheet_set_pivot_data_field`.
 *
 *   * `name`          — display name (key for GETPIVOTDATA). `NULL` or the
 *                       empty string asks the engine to name the data field
 *                       from the workbook profile's pivot wording, e.g.
 *                       "Sum of Sales" (en-US) or "合計 / Sales" (ja-JP):
 *                       the aggregation label, the locale's separator and
 *                       the source pivot field's display name, followed by
 *                       2, 3, ... while that repeats another data field's
 *                       name or a field's display name under the profile's
 *                       collation ("Sum of Amt2").
 *   * `field_index`   — index into `PivotTable::fields()` of the source
 *                       pivot field; must be in range.
 *   * `aggregation`   — aggregation function applied to the source.
 *   * `number_format` — nullable; the data field's `numFmtId` as a decimal
 *                       string, under the same rules as
 *                       `fm_pivot_field_spec_t::number_format`.
 *   * `show_as`       — derivation mode.
 *   * `show_as_base_field` — `-1` means unset; otherwise an index into
 *                       `PivotTable::fields()`, which must be in range.
 *   * `show_as_base_item`  — `-1` means unset; `1048828` means "(previous)";
 *                       `1048829` means "(next)"; any other value is an
 *                       index into the base field's `items[]`.
 */
typedef struct {
  const char* name;
  uint32_t field_index;
  int32_t aggregation; /* fm_pivot_aggregation_t */
  const char* number_format;
  int32_t show_as; /* fm_pivot_show_values_as_t */
  int32_t show_as_base_field;
  int32_t show_as_base_item;
} fm_pivot_data_field_spec_t;

/**
 * @brief Read-back of a pivot field definition (`fm_sheet_get_pivot_field`).
 *
 * The first five members carry the same meaning as `fm_pivot_field_spec_t`,
 * except that `axis` is `-1` for a field that sits on no axis. Strings are
 * model-backed views, never `NULL`.
 *
 *   * `cache_field_index` — index of the bound pivot-cache field, `-1` when
 *                           the cache does not cover this field. Pivot field
 *                           `i` binds to cache field `i`.
 *   * `sort_order`        — `fm_pivot_sort_order_t`, as written by
 *                           `fm_sheet_set_pivot_field_sort`.
 *   * `sort_by_field`     — the by-field key; empty for label order.
 *   * `date_periods`      — for a date-grouped field, the `FM_PIVOT_PERIOD_*`
 *                           bits of its group fields; `0` otherwise.
 *   * `date_calendar`     — `fm_pivot_calendar_t` of that grouping;
 *                           `FM_PIVOT_CALENDAR_GREGORIAN` when ungrouped.
 *   * `group_base_index`  — for a group field, the index of the field it
 *                           groups; `-1` otherwise.
 *   * `default_subtotal`  — non-zero when the field shows the automatic
 *                           subtotal (OOXML `defaultSubtotal`). Custom
 *                           functions replace it, so it reads `0` whenever
 *                           `subtotal_fns` is non-zero.
 *   * `subtotal_fns`      — bit `1u << agg` is set for each
 *                           `fm_pivot_aggregation_t` selected through
 *                           `fm_sheet_add_pivot_field_subtotal_fn` or read
 *                           from the `*Subtotal` attributes; `0` for none.
 *   * `insert_blank_row`  — non-zero when a blank row follows each of the
 *                           field's item blocks (OOXML `insertBlankRow`).
 *   * `fill_down_labels`  — non-zero when the field's label repeats on every
 *                           row of its block (x14 `fillDownLabels`).
 *   * `show_all`          — non-zero when items without data are listed
 *                           (OOXML `showAll`; absent in a file reads `1`, a
 *                           field created through the API starts at `0`).
 */
typedef struct {
  const char* source_name;
  const char* custom_name;
  int32_t axis;
  int32_t subtotal_top;
  const char* number_format;
  int32_t cache_field_index;
  int32_t sort_order;
  const char* sort_by_field;
  uint32_t date_periods;
  int32_t date_calendar;
  int32_t group_base_index;
  int32_t default_subtotal;
  uint32_t subtotal_fns;
  int32_t insert_blank_row;
  int32_t fill_down_labels;
  int32_t show_all;
} fm_pivot_field_info_t;

/**
 * @brief Read-back of one pivot-field item (`fm_sheet_get_pivot_field_item`).
 *
 *   * `name`        — the item's own label, a model-backed view that is
 *                     empty for an item bound only by cache index.
 *   * `cache_index` — index into the bound cache field's shared items (the
 *                     OOXML `<item x="N">`), `-1` for an item bound by label
 *                     only.
 *   * `visible`     — non-zero when the item is shown.
 *   * `caption`     — the caption shown instead of the item's value, a
 *                     model-backed view; `NULL` when the item has none.
 *   * `show_detail` — non-zero when the item is expanded, `0` when it is
 *                     collapsed.
 */
typedef struct {
  const char* name;
  int32_t cache_index;
  int32_t visible;
  const char* caption;
  int32_t show_detail;
} fm_pivot_item_info_t;

/**
 * @brief Read-back of a pivot table's `<pivotTableStyleInfo>`
 *        (`fm_sheet_get_pivot_style_info`).
 *
 *   * `name`    — the table style name, a model-backed view that is empty
 *                 when the element carries no name or is absent.
 *   * `present` — non-zero when the definition has a style element at all.
 *   * the `show_*` flags are `0` or `1`, and `0` when `present` is `0`.
 */
typedef struct {
  const char* name;
  int32_t present;
  int32_t show_row_headers;
  int32_t show_col_headers;
  int32_t show_row_stripes;
  int32_t show_col_stripes;
  int32_t show_last_column;
} fm_pivot_style_info_t;

/**
 * @brief Plain-data spec for `fm_sheet_add_pivot_filter`.
 *
 *   * `axis`           — pivot axis the filter targets.
 *   * `field_name`     — required; the pivot field's name (or display name).
 *   * `type`           — filter type.
 *   * `value_kind`     — discriminator for the primary `value` payload, the
 *                        lower bound of a range type. `_NONE` for the
 *                        relative-period and calendar-position types and
 *                        rejected for every other type; `TEXT` for a label
 *                        type, `INT` or `DOUBLE` for a value or date type.
 *   * `value_int`      — int payload; consulted iff `value_kind == INT`.
 *   * `value_double`   — double payload; consulted iff `value_kind == DOUBLE`.
 *   * `value_text`     — UTF-8 payload; consulted iff `value_kind == TEXT`.
 *                        Must be non-NULL when `value_kind == TEXT`.
 *   * `value_high_kind` — discriminator for the upper-bound payload.
 *                        `FM_PIVOT_FILTER_VALUE_NONE` for non-range filters.
 *                        `TEXT` for `LABEL_BETWEEN` / `LABEL_NOT_BETWEEN`;
 *                        `INT` or `DOUBLE` for the other range types.
 *   * `value_high_int` / `value_high_double` — upper bound payload.
 *   * `data_field_index` — which data-field aggregate a value filter scores.
 *                        Validated against the pivot's data fields for the
 *                        value filter types; label/date filters ignore it.
 *   * `value_high_text` — UTF-8 upper bound of `LABEL_BETWEEN` /
 *                        `LABEL_NOT_BETWEEN`; must be non-NULL for them and
 *                        NULL for every other type.
 */
typedef struct {
  int32_t axis; /* fm_pivot_axis_t */
  const char* field_name;
  int32_t type;       /* fm_pivot_filter_type_t */
  int32_t value_kind; /* fm_pivot_filter_value_kind_t */
  int32_t value_int;
  double value_double;
  const char* value_text;
  int32_t value_high_kind; /* fm_pivot_filter_value_kind_t */
  int32_t value_high_int;
  double value_high_double;
  uint32_t data_field_index;
  const char* value_high_text;
} fm_pivot_filter_spec_t;

/**
 * @brief The PivotTable Options of one pivot table
 *        (`fm_sheet_get_pivot_options` / `fm_sheet_set_pivot_options`).
 *
 * `fm_sheet_set_pivot_options` writes every member, so read the current
 * options, change what is needed and set them back. Flags read `0` or `1`
 * and any non-zero value sets them. A string reads `NULL` when the option is
 * unset; setting `NULL` or an empty string unsets it. Strings read back are
 * read-scratch-backed views.
 *
 *   * `show_error` / `error_caption`     — error cells show `error_caption`
 *                                          (`showError`, `errorCaption`). The
 *                                          caption is kept while `show_error`
 *                                          is `0`.
 *   * `show_missing` / `missing_caption` — empty value cells show
 *                                          `missing_caption`; a caption of
 *                                          exactly `-` is written as the
 *                                          number 0 (`showMissing`,
 *                                          `missingCaption`).
 *   * `show_headers`         — field captions and filter drop-downs
 *                              (`showHeaders`).
 *   * `show_drill`           — expand / collapse buttons (`showDrill`).
 *   * `show_data_tips`       — contextual tooltips (`showDataTips`).
 *   * `grid_drop_zones`      — classic layout with grid drop zones
 *                              (`gridDropZones`).
 *   * `show_empty_row` / `show_empty_col` — `showEmptyRow` / `showEmptyCol`;
 *                              the layout follows each field's `show_all`.
 *   * `show_data_drop_down`  — `showDataDropDown`.
 *   * `print_drill`          — print expand / collapse buttons (`printDrill`).
 *   * `field_print_titles`   — repeat field titles on each printed page
 *                              (`fieldPrintTitles`).
 *   * `item_print_titles`    — repeat row labels on each printed page
 *                              (`itemPrintTitles`).
 *   * `merge_item`           — merge and center cells with labels
 *                              (`mergeItem`).
 *   * `indent`               — compact-form indent per level, 0 to 126
 *                              (`indent`; the dialog's characters minus one).
 *   * `page_over_then_down`  — report filter fields run across, then down
 *                              (`pageOverThenDown`).
 *   * `page_wrap`            — report filter fields per column or row, 0 to
 *                              255; 0 does not wrap (`pageWrap`).
 *   * `multiple_field_filters` — allow several filters per field
 *                              (`multipleFieldFilters`).
 *   * `preserve_formatting`  — keep cell formatting on update
 *                              (`preserveFormatting`).
 *   * `custom_list_sort`     — sort by custom lists (`customListSort`).
 *   * `alt_text` / `alt_text_summary` — alternative text title and
 *                              description (x14 `altText`, `altTextSummary`).
 *   * `grand_total_caption`  — the grand total label in place of the
 *                              locale's word (`grandTotalCaption`).
 *   * `auto_formatting`      — autofit column widths on update
 *                              (`useAutoFormatting`).
 *   * `enable_drill`         — Show Details is allowed (`enableDrill`).
 *
 * A pivot created through the API starts with `item_print_titles` on,
 * `indent` 0 and `multiple_field_filters` off, as Excel creates one.
 */
typedef struct {
  int32_t show_error;
  const char* error_caption;
  int32_t show_missing;
  const char* missing_caption;
  int32_t show_headers;
  int32_t show_drill;
  int32_t show_data_tips;
  int32_t grid_drop_zones;
  int32_t show_empty_row;
  int32_t show_empty_col;
  int32_t show_data_drop_down;
  int32_t print_drill;
  int32_t field_print_titles;
  int32_t item_print_titles;
  int32_t merge_item;
  uint32_t indent;
  int32_t page_over_then_down;
  uint32_t page_wrap;
  int32_t multiple_field_filters;
  int32_t preserve_formatting;
  int32_t custom_list_sort;
  const char* alt_text;
  const char* alt_text_summary;
  const char* grand_total_caption;
  int32_t auto_formatting;
  int32_t enable_drill;
} fm_pivot_options_t;

/**
 * @brief The options of one pivot cache, shared by every pivot table drawing
 *        from it (`fm_workbook_get_pivot_cache_options` /
 *        `fm_workbook_set_pivot_cache_options`).
 *
 *   * `save_data`           — save the source records with the file
 *                             (`saveData`). Off, the records part is left out
 *                             and the reports keep their sheet cells.
 *   * `missing_items_limit` — items kept per field after they leave the
 *                             source (`missingItemsLimit`): `-1` automatic
 *                             (no attribute), `0` none, up to `1048576`.
 *                             Excel's "Max" writes 32500.
 *   * `refresh_on_load`     — refresh the data when the file opens
 *                             (`refreshOnLoad`).
 */
typedef struct {
  int32_t save_data;
  int32_t missing_items_limit;
  int32_t refresh_on_load;
} fm_pivot_cache_options_t;

/**
 * @brief Read-back of one calculated item
 *        (`fm_workbook_get_pivot_cache_calculated_item`).
 *
 *   * `field_index` — the cache field the item belongs to.
 *   * `name`        — the item's name, a read-scratch-backed view.
 *   * `formula`     — its formula without a leading `=`, a
 *                     read-scratch-backed view.
 *   * `scoped`      — non-zero for an entry read from a file that applies the
 *                     formula only where the item meets items of other
 *                     fields.
 */
typedef struct {
  uint32_t field_index;
  const char* name;
  const char* formula;
  int32_t scoped;
} fm_pivot_calculated_item_t;

/* --- Pivot caches (workbook-owned) --------------------------------------- */

/**
 * @brief Returns the number of pivot caches owned by the workbook.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Returns the `cache_id` of the pivot cache at flat index `idx`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_id_at(const fm_workbook_t* wb, size_t idx, uint32_t* out_cache_id);

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
 * `fm_workbook_set_pivot_cache_worksheet_source` before saving. The range
 * only has to name where the records came from; it need not hold data,
 * because the records carry the values themselves.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `requested_id` collides with an
 *         existing cache or auto-assignment would overflow.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache(fm_workbook_t* wb, uint32_t requested_id, uint32_t* out_cache_id);

/**
 * @brief Removes the pivot cache with id `cache_id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when no cache matches `cache_id` or when
 *         the cache is still referenced by at least one pivot table.
 */
FM_API fm_status_t fm_workbook_remove_pivot_cache(fm_workbook_t* wb, uint32_t cache_id);

/**
 * @brief Reads the cache's worksheet source metadata.
 *
 * `out_present` is non-zero when a `<worksheetSource>` is present.
 * `out_ref`, `out_sheet`, and `out_name` are model-backed strings owned by
 * the workbook handle and remain valid until the cache is mutated or the
 * handle is destroyed. Reads do not invalidate them.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_worksheet_source(const fm_workbook_t* wb, uint32_t cache_id,
                                                                int32_t* out_present, const char** out_ref,
                                                                const char** out_sheet, const char** out_name);

/**
 * @brief Sets or clears the cache's worksheet source metadata.
 *
 * Pass `present == 0` to clear `<worksheetSource>`. When present is
 * non-zero, nullable string attributes are copied; `NULL` means absent, and
 * the cache becomes a worksheet-type source: any other `<cacheSource>` type,
 * attribute or child captured on load is dropped.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_worksheet_source(fm_workbook_t* wb, uint32_t cache_id, int32_t present,
                                                                const char* ref, const char* sheet, const char* name);

/** @brief Number of fields on the cache identified by `cache_id`. */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_count(const fm_workbook_t* wb, uint32_t cache_id,
                                                           size_t* out_count);

/**
 * @brief Reads the name of the cache field at `field_idx`. The pointer is a
 *        model-backed view into the workbook handle and remains valid until
 *        the cache field list is mutated or the handle is destroyed. Reads
 *        do not invalidate it.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_name(const fm_workbook_t* wb, uint32_t cache_id, size_t field_idx,
                                                          const char** out_utf8);

/**
 * @brief Appends a new field with the given UTF-8 name to the cache.
 *        `out_field_idx` receives the new field's index. `utf8_name`
 *        must be non-NULL. Every pivot table drawing from the cache gains an
 *        unplaced field for it.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_field(fm_workbook_t* wb, uint32_t cache_id, const char* utf8_name,
                                                     size_t* out_field_idx);

/** @brief Drops every field (and every record) from the cache. */
FM_API fm_status_t fm_workbook_clear_pivot_cache_fields(fm_workbook_t* wb, uint32_t cache_id);

/** @brief Number of shared items configured on cache field `field_idx`. */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_shared_item_count(const fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_idx, size_t* out_count);

/**
 * @brief Reads shared item `item_idx` of cache field `field_idx` into `out`.
 *        Date items are stored as serial numbers and read back as
 *        `FM_VAL_NUMBER`; query `fm_workbook_is_pivot_cache_field_shared_item_date`
 *        to tell them from plain numbers. Text is read-scratch backed: the pointer remains
 *        valid until the next scratch-backed read, any mutation, or handle
 *        destruction. `item_idx >= count` is `FM_STATUS_INVALID_ARGUMENT` (2).
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_shared_item_value(const fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_idx, size_t item_idx,
                                                                       fm_value_t* out);

/** @brief Appends a numeric shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_number(fm_workbook_t* wb, uint32_t cache_id,
                                                                        size_t field_idx, double value);

/**
 * @brief Appends a date shared item to cache field `field_idx`. `serial` is a
 *        date serial in the workbook's date system, as a cell holds it; it must
 *        be finite, non-negative and not past 9999-12-31, else
 *        `FM_STATUS_INVALID_ARGUMENT` (2). Saved with seconds precision.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_date(fm_workbook_t* wb, uint32_t cache_id,
                                                                      size_t field_idx, double serial);

/**
 * @brief Reports whether shared item `item_idx` of cache field `field_idx` is
 *        a date (`*out_is_date` is 1) or not (0). `item_idx >= count` is
 *        `FM_STATUS_INVALID_ARGUMENT` (2).
 */
FM_API fm_status_t fm_workbook_is_pivot_cache_field_shared_item_date(const fm_workbook_t* wb, uint32_t cache_id,
                                                                     size_t field_idx, size_t item_idx,
                                                                     int32_t* out_is_date);

/** @brief Appends a text shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_text(fm_workbook_t* wb, uint32_t cache_id,
                                                                      size_t field_idx, const char* utf8);

/** @brief Appends a boolean shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_bool(fm_workbook_t* wb, uint32_t cache_id,
                                                                      size_t field_idx, int32_t value);

/** @brief Appends a blank shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_blank(fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_idx);

/** @brief Appends an Excel error shared item to cache field `field_idx`. */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_error(fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_idx, fm_error_code_t error);

/** @brief Drops every shared item from cache field `field_idx`. */
FM_API fm_status_t fm_workbook_clear_pivot_cache_field_shared_items(fm_workbook_t* wb, uint32_t cache_id,
                                                                    size_t field_idx);

/** @brief Returns the number of records (rows) on the cache. */
FM_API fm_status_t fm_workbook_get_pivot_cache_record_count(const fm_workbook_t* wb, uint32_t cache_id,
                                                            size_t* out_count);

/**
 * @brief Appends a new empty record to the cache. `out_record_idx`
 *        receives the new record's index. The record's cell vector is
 *        empty; populate it via `fm_workbook_pivot_cache_record_set_*`.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_record(fm_workbook_t* wb, uint32_t cache_id, size_t* out_record_idx);

/** @brief Drops every record from the cache. */
FM_API fm_status_t fm_workbook_clear_pivot_cache_records(fm_workbook_t* wb, uint32_t cache_id);

/**
 * @brief Sets cell `(record_idx, field_idx)` to a numeric value. The
 *        record's cell vector auto-extends to `field_idx + 1` with
 *        Blank fillers when shorter.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_number(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                             size_t field_idx, double value);

/**
 * @brief Sets cell `(record_idx, field_idx)` to an inline date. `serial` is a
 *        date serial in the workbook's date system and is validated as in
 *        `fm_workbook_add_pivot_cache_field_shared_item_date`. The cell is
 *        written as a date value even on a field with shared items; saved
 *        with seconds precision.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_date(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                           size_t field_idx, double serial);

/** @brief Sets cell `(record_idx, field_idx)` to a UTF-8 text value. */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_text(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                           size_t field_idx, const char* utf8);

/** @brief Sets cell `(record_idx, field_idx)` to a boolean value. */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_bool(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                           size_t field_idx, int32_t value);

/** @brief Sets cell `(record_idx, field_idx)` to Blank. */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_blank(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                            size_t field_idx);

/** @brief Sets cell `(record_idx, field_idx)` to an Excel error value. */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_error(fm_workbook_t* wb, uint32_t cache_id, size_t record_idx,
                                                            size_t field_idx, fm_error_code_t error);

/*
 * Every successful cache mutation above and below (sources, fields, shared
 * items, records, options, calculated fields and items, date grouping) leaves
 * each pivot table drawing from the cache projection-pending: its sheet cells
 * are not rewritten until that table's next edit, `fm_sheet_project_pivot`
 * or a save (see `fm_sheet_is_pivot_projection_pending`).
 */

/**
 * @brief Reads the cache's options into `out`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cache_id` does not match any cache.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_options(const fm_workbook_t* wb, uint32_t cache_id,
                                                       fm_pivot_cache_options_t* out);

/**
 * @brief Writes every member of `options` to the cache. A
 *        `missing_items_limit` equal to the value read from the file is kept
 *        even when outside the settable range.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `options` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cache_id` does not match any cache or
 *         `missing_items_limit` is outside `-1` to `1048576`.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_options(fm_workbook_t* wb, uint32_t cache_id,
                                                       const fm_pivot_cache_options_t* options);

/**
 * @brief Adds calculated field `name` with `formula` to the cache, as the last
 *        cache field; `*out_field_index` receives its index. Every pivot
 *        table drawing from the cache gains an unplaced field for it; place
 *        it with `fm_sheet_add_pivot_data_field`.
 *
 * The formula is in English canonical form and names cache fields only,
 * quoting a name with spaces or other non-identifier characters as
 * `'Amt Net'` (`''` escapes a quote). A leading `=` is optional and is not
 * stored. A cell shows the formula applied to the Sums of the named fields,
 * whatever the data field's function.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `name`, `formula` or
 *         `out_field_index` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) for malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cache_id` does not match any cache, the
 *         name is empty, longer than 255 UTF-16 code units, `All` or `Blank`,
 *         contains `[` or `]`, or equals an existing field name under the
 *         profile's collation, or the formula is empty, longer than 8192
 *         UTF-16 code units, does not parse, names something other than a
 *         cache field, refers to a cell, range, table, external workbook or
 *         defined name, holds an array constant or a function a calculated
 *         field cannot use, or reaches back to itself.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_calculated_field(fm_workbook_t* wb, uint32_t cache_id, const char* name,
                                                                const char* formula, size_t* out_field_index);

/**
 * @brief Replaces the formula of calculated field `field_index`.
 *
 * @return As `fm_workbook_add_pivot_cache_calculated_field` for the formula;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `field_index` is not a calculated field.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_calculated_field_formula(fm_workbook_t* wb, uint32_t cache_id,
                                                                        size_t field_index, const char* formula);

/**
 * @brief Removes calculated field `field_index` and its data fields in every
 *        pivot table; the fields after it move down by one.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cache_id` does not match any cache or
 *         `field_index` is not a calculated field;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) while another calculated field's formula
 *         names it.
 */
FM_API fm_status_t fm_workbook_remove_pivot_cache_calculated_field(fm_workbook_t* wb, uint32_t cache_id,
                                                                   size_t field_index);

/**
 * @brief Reads the formula of cache field `field_idx`, without a leading `=`,
 *        into `*out_utf8`: an empty string when the field is not calculated.
 *        A model-backed view valid until the field's formula or the cache
 *        field list changes; a formula respelled from the model's (see
 *        `fm_workbook_get_file_name`) is read-scratch-backed instead.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_formula(const fm_workbook_t* wb, uint32_t cache_id,
                                                             size_t field_idx, const char** out_utf8);

/**
 * @brief Adds calculated item `name` with `formula` to cache field
 *        `field_index`. The item follows the field's items in every pivot
 *        table and is last in solve order: where items meet, the formula of
 *        the item later in that order wins.
 *
 * The formula names items of the same field, as `North`, `'North East'`,
 * `Region[North]`, `Region[2]` (position) or `Region[-1]` (offset). A cell
 * shows the formula applied to the named items' values in the same context,
 * and the field's totals include the item.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `name` or `formula` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) for malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cache_id` or `field_index` is out of
 *         range, for a name as `fm_workbook_add_pivot_cache_calculated_field`
 *         refuses it or equal to an item of the field, or for a formula it
 *         refuses or that names another field;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the field is grouped, calculated or on
 *         no pivot table's axis, or a data field drawing from the cache uses
 *         Average, StdDev, StdDevP, Var or VarP.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_calculated_item(fm_workbook_t* wb, uint32_t cache_id, size_t field_index,
                                                               const char* name, const char* formula);

/** @brief Number of calculated items in the cache, in solve order. */
FM_API fm_status_t fm_workbook_get_pivot_cache_calculated_item_count(const fm_workbook_t* wb, uint32_t cache_id,
                                                                     size_t* out_count);

/**
 * @brief Reads calculated item `idx` (in solve order) into `out`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cache_id` or `idx` is out of range.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_calculated_item(const fm_workbook_t* wb, uint32_t cache_id, size_t idx,
                                                               fm_pivot_calculated_item_t* out);

/**
 * @brief Replaces the formula of calculated item `idx`.
 *
 * @return As `fm_workbook_add_pivot_cache_calculated_item` for the formula;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range or the formula refers
 *         back to the item.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_calculated_item_formula(fm_workbook_t* wb, uint32_t cache_id, size_t idx,
                                                                       const char* formula);

/**
 * @brief Moves calculated item `idx` to position `new_idx` in solve order.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cache_id`, `idx` or `new_idx` is out of
 *         range.
 */
FM_API fm_status_t fm_workbook_move_pivot_cache_calculated_item(fm_workbook_t* wb, uint32_t cache_id, size_t idx,
                                                                size_t new_idx);

/**
 * @brief Removes calculated item `idx`, its item in the cache field and in
 *        every pivot table.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cache_id` or `idx` is out of range;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) while another calculated item names it.
 */
FM_API fm_status_t fm_workbook_remove_pivot_cache_calculated_item(fm_workbook_t* wb, uint32_t cache_id, size_t idx);

/* --- Pivot tables (sheet-owned) ------------------------------------------ */

/*
 * A pivot table's report is kept in its sheet's cells, as Excel keeps it.
 * Every successful pivot-table mutation below writes the report into the
 * cells: those the previous report covered and the new range are cleared,
 * then each report cell is written as a value with Excel's format for its
 * role, and formulas reading the range recalculate. A mutation whose report
 * would overlap another pivot table, a table, a data table, an array formula
 * or a merge it does not own returns `FM_STATUS_PRECONDITION_FAILED` (9) and leaves the
 * pivot table and the sheet unchanged. A pivot table read from a file keeps
 * the cells read with it until it is edited.
 */

/**
 * @brief Creates a new empty pivot table on `sheet_index`, with the options
 *        Excel gives a new pivot. `cache_id` must reference an existing
 *        cache. The new table's flat index is written to `*out_pivot_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `utf8_name`, or `out_pivot_index`
 *         is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `cache_id` does not match any existing cache;
 *         `FM_STATUS_PRECONDITION_FAILED` (9), creating nothing, when the report range
 *         would overlap another structure.
 */
FM_API fm_status_t fm_sheet_add_pivot(fm_workbook_t* wb, size_t sheet_index, const char* utf8_name, uint32_t cache_id,
                                      uint32_t anchor_row, uint32_t anchor_col, size_t* out_pivot_index);

/**
 * @brief Removes the pivot table at `pivot_index` from `sheet_index`, clearing
 *        the cells and merges its report was last written to.
 */
FM_API fm_status_t fm_sheet_remove_pivot(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/** @brief Renames the pivot table. `utf8_name` must be non-NULL. */
FM_API fm_status_t fm_sheet_set_pivot_name(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                           const char* utf8_name);

/** @brief Updates the pivot's anchor cell and span. */
FM_API fm_status_t fm_sheet_set_pivot_anchor(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                             uint32_t anchor_row, uint32_t anchor_col, uint32_t span_rows,
                                             uint32_t span_cols);

/** @brief Toggles the row / column grand total bands on the pivot. */
FM_API fm_status_t fm_sheet_set_pivot_grand_totals(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                   int32_t rows_enabled, int32_t cols_enabled);

/**
 * @brief Sets the pivot's table style and toggles, making the style element
 *        present. A `NULL` or empty `utf8_style_name` writes the element
 *        without a style name (no table style); the flags are `0` or non-zero.
 *        A pivot built by `fm_sheet_add_pivot` starts as
 *        `PivotStyleLight16` with row / column headers and last column on,
 *        banded rows / columns off.
 */
FM_API fm_status_t fm_sheet_set_pivot_style_info(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 const char* utf8_style_name, int32_t show_row_headers,
                                                 int32_t show_col_headers, int32_t show_row_stripes,
                                                 int32_t show_col_stripes, int32_t show_last_column);

/** @brief Reads the pivot's compact / tabular / outline report layout. */
FM_API fm_status_t fm_sheet_get_pivot_report_layout(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    int32_t* out_layout);

/**
 * @brief Sets the pivot's compact / tabular / outline report layout.
 *
 * `layout` is a raw signed 32-bit ordinal. Unknown values return
 * `FM_STATUS_INVALID_ARGUMENT` (2); callers do not need to construct an out-of-domain
 * `fm_pivot_report_layout_t` value.
 */
FM_API fm_status_t fm_sheet_set_pivot_report_layout(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    int32_t layout);

/** @brief Number of fields configured on the pivot. */
FM_API fm_status_t fm_sheet_get_pivot_field_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                  size_t* out_count);

/**
 * @brief Places a field on the pivot. `spec->source_name` must be non-NULL;
 *        the other string fields are nullable.
 *
 * A pivot's field list is positional over its cache's fields: pivot field
 * `i` is cache field `i`. When `source_name` names a field of the bound
 * cache, the pivot first gets one bare field (no axis) per cache field, and
 * the named one is configured in place; `*out_field_idx` is that cache
 * field's position, and adding a field that is already there updates it
 * rather than adding a second. A name the cache does not carry (or a pivot
 * whose cache has no fields) appends a new field at the end instead.
 * A field placed on the row or column axis is appended to that axis's field
 * order.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `spec`, `out_field_idx` or
 *           `spec->source_name` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `pivot_index` is out of range, or
 *           `spec->axis` names no declared enumerator.
 */
FM_API fm_status_t fm_sheet_add_pivot_field(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                            const fm_pivot_field_spec_t* spec, size_t* out_field_idx);

/** @brief Drops every field from the pivot, emptying the row and column field orders. */
FM_API fm_status_t fm_sheet_clear_pivot_fields(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Sets the axis of pivot field `field_idx`.
 *
 * `axis` is a raw signed 32-bit ordinal; unknown values return
 * `FM_STATUS_INVALID_ARGUMENT` (2). Moving a field onto the row or column axis appends it
 * to that axis's field order (a field already there keeps its place);
 * moving it off removes it from the order it was in.
 *
 * `-1` takes the field off every axis, as unchecking it in a field list
 * does. The data fields aggregating it are removed (later data-field
 * indices shift down), as are the filters naming the field and the
 * value filters scoring a removed data field; the field's item visibility
 * is kept.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_axis(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 size_t field_idx, int32_t axis);

/**
 * @brief Sets the sort directive on pivot field `field_idx`.
 *
 * `order` is a raw signed 32-bit `fm_pivot_sort_order_t` ordinal; unknown
 * values return `FM_STATUS_INVALID_ARGUMENT` (2). `FM_PIVOT_SORT_MANUAL` keeps sibling
 * items in the field's item document order (`fm_sheet_add_pivot_field_item`
 * order) and takes no `by_field` key: a non-NULL, non-empty `by_field` with it
 * returns `FM_STATUS_INVALID_ARGUMENT` (2). For the other orders, `by_field == NULL` (or
 * empty) sorts by label and clears the by-field key.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_sort(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 size_t field_idx, int32_t order, const char* by_field);

/** @brief Sets the `subtotal_top` flag on pivot field `field_idx`. */
FM_API fm_status_t fm_sheet_set_pivot_field_subtotal_top(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                         size_t field_idx, int32_t top);

/**
 * @brief Turns the automatic subtotal of pivot field `field_idx` on or off
 *        (`0` off, any other value on). Custom subtotal functions replace the
 *        automatic subtotal whatever this flag says.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_default_subtotal(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                             size_t field_idx, int32_t on);

/**
 * @brief Sets whether a blank row follows each item block of pivot field
 *        `field_idx` (`0` off, any other value on).
 */
FM_API fm_status_t fm_sheet_set_pivot_field_insert_blank_row(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                             size_t field_idx, int32_t on);

/**
 * @brief Sets whether pivot field `field_idx` repeats its item labels on
 *        every row of a block (`0` off, any other value on).
 */
FM_API fm_status_t fm_sheet_set_pivot_field_fill_down_labels(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                             size_t field_idx, int32_t on);

/**
 * @brief Renames pivot field `field_idx` in the report; `NULL` or an empty
 *        `name` restores its source name. Data field names are not changed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) for malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if an index is out of range, `name` is longer
 *           than 255 UTF-16 code units, or it equals under the profile's
 *           collation another field's source or custom name or a data
 *           field's name (the field's own source name is allowed);
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the report would overlap another
 *           structure.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_custom_name(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                        size_t field_idx, const char* name);

/**
 * @brief Sets whether pivot field `field_idx` lists items without data
 *        (`0` off, any other value on).
 */
FM_API fm_status_t fm_sheet_set_pivot_field_show_all(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                     size_t field_idx, int32_t on);

/**
 * @brief Expands (non-zero) or collapses (`0`) every item of pivot field
 *        `field_idx`, together with the field's subtotal row.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_show_detail(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                        size_t field_idx, int32_t on);

/**
 * @brief Expands (non-zero) or collapses (`0`) item `item_idx` of pivot field
 *        `field_idx`.
 *
 * @return `FM_STATUS_INVALID_ARGUMENT` (2) when an index is out of range, otherwise as
 *         `fm_sheet_set_pivot_field_show_detail`.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_item_show_detail(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                             size_t field_idx, size_t item_idx, int32_t on);

/**
 * @brief Shows `caption` instead of the value of item `item_idx` of pivot
 *        field `field_idx`.
 *
 * `NULL` or an empty string clears the caption, as does a caption equal to
 * the item's own value. The item's value, not the caption, still decides
 * hidden-item matching and ordering.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an index is out of range, `caption` is not
 *         valid UTF-8 or is longer than 255 UTF-16 code units, or it matches,
 *         under the workbook profile's text collation, the label another item
 *         of the field shows.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_item_caption(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                         size_t field_idx, size_t item_idx, const char* caption);

/**
 * @brief Appends a manual-filter item to pivot field `field_idx`, addressed
 *        by its label.
 *        `utf8_name` must be non-NULL. `visible` is a 32-bit boolean.
 *
 * When a shared item of the bound cache field renders as `utf8_name`, the
 * item binds to the first such shared item and the writer emits that index as
 * `<item x="N">`; otherwise it stays unbound and the writer emits its
 * document-order position. The filter engine matches source records by
 * comparing their rendered label against `utf8_name`. That makes an empty
 * `utf8_name` unusable as the blank item: the blank has no label of its own
 * and is identified by what it binds to. Build one with
 * `fm_sheet_add_pivot_field_shared_item` instead.
 */
FM_API fm_status_t fm_sheet_add_pivot_field_item(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
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
FM_API fm_status_t fm_sheet_add_pivot_field_shared_item(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                        size_t field_idx, size_t cache_index, int32_t visible);

/** @brief Drops every manual-filter item from pivot field `field_idx`. */
FM_API fm_status_t fm_sheet_clear_pivot_field_items(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    size_t field_idx);

/** @brief Toggles the visibility of item `item_idx` on field `field_idx`. */
FM_API fm_status_t fm_sheet_set_pivot_field_item_visible(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                         size_t field_idx, size_t item_idx, int32_t visible);

/**
 * @brief Appends a subtotal-fn entry to pivot field `field_idx`.
 *
 * `agg` is a raw signed 32-bit ordinal; unknown values return
 * `FM_STATUS_INVALID_ARGUMENT` (2). The selection reads back through
 * `fm_pivot_field_info_t::subtotal_fns`.
 */
FM_API fm_status_t fm_sheet_add_pivot_field_subtotal_fn(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                        size_t field_idx, int32_t agg);

/** @brief Drops every subtotal-fn entry from pivot field `field_idx`. */
FM_API fm_status_t fm_sheet_clear_pivot_field_subtotal_fns(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                           size_t field_idx);

/**
 * @brief Groups the date field `field_idx` (a database field, or one of its
 *        group fields) by the `FM_PIVOT_PERIOD_*` bits of `periods_mask`, as
 *        Excel's Group command does, replacing any grouping it has.
 *
 * Each period becomes a group field appended to the pivot cache, finest
 * first, named and labelled in the workbook profile's locale. Every pivot on
 * the cache gains them: this pivot shows them in the field's place on its
 * axis, coarsest outermost; another pivot showing the field shows them
 * before it. `interval_days` above 1 groups by runs of that many days and
 * needs `periods_mask == FM_PIVOT_PERIOD_DAYS`.
 * `FM_PIVOT_CALENDAR_JAPANESE` groups the years period by era year.
 * `start_or_nan` / `end_or_nan` are date serials; NaN takes the field's
 * first value and its last (the midnight after it for a date without a time).
 *
 * @return `FM_STATUS_INVALID_ARGUMENT` (2) for an index out of range, a `periods_mask` of 0
 *         or with an unknown bit, an unknown `calendar`, `interval_days == 0`,
 *         `interval_days > 1` with any period but days, an infinite or
 *         otherwise invalid date bound, a start after the end, or a field
 *         holding a blank, a value that is not a date, or no value;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the field is grouped and its group fields
 *         are not the last fields of the cache.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_date_group(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       size_t field_idx, uint32_t periods_mask, int32_t calendar,
                                                       uint32_t interval_days, double start_or_nan, double end_or_nan);

/**
 * @brief Removes the date grouping of field `field_idx` (a database field, or
 *        one of its group fields) from every pivot on its cache; the field
 *        returns to the axis position its group fields held. A field that is
 *        not grouped is left as it is.
 *
 * @return `FM_STATUS_INVALID_ARGUMENT` (2) for an index out of range; `FM_STATUS_PRECONDITION_FAILED` (9)
 *         when the group fields are not the last fields of the cache.
 */
FM_API fm_status_t fm_sheet_clear_pivot_field_date_group(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                         size_t field_idx);

/**
 * @brief Sets the `numFmtId` of pivot field `field_idx`, as a decimal string
 *        naming a built-in id or one returned by `fm_workbook_add_num_fmt`.
 *        `utf8` must be non-NULL (use the empty string to clear); any other
 *        string that is not a known id returns `FM_STATUS_INVALID_ARGUMENT` (2).
 */
FM_API fm_status_t fm_sheet_set_pivot_field_number_format(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                          size_t field_idx, const char* utf8);

/**
 * @brief Replaces the row-axis field order with `indices[0..count)`.
 *        Each entry must be `< field_count` or `FM_PIVOT_VALUES_FIELD`, and
 *        appear once, else `FM_STATUS_INVALID_ARGUMENT` (2). The list is the axis's whole
 *        membership: listed fields move onto the row axis (leaving the column
 *        order), and fields previously on it but not listed come off every
 *        axis. Pass `count == 0` to clear. Adding or moving a field already
 *        maintains the order, so this is needed only to reorder.
 *
 * `FM_PIVOT_VALUES_FIELD` places the Values level at that position and takes
 * it off the other axis. It needs at least two data fields
 * (`FM_STATUS_INVALID_ARGUMENT` (2) otherwise). With two or more data fields the Values
 * level is always on one axis, so a list for the axis that holds it must
 * include it (`FM_STATUS_INVALID_ARGUMENT` (2) otherwise); the column axis holds it, at the
 * end, until it is moved.
 */
FM_API fm_status_t fm_sheet_set_pivot_row_field_order(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      const uint32_t* indices, size_t count);

/** @brief Replaces the column-axis field order. Same contract as rows. */
FM_API fm_status_t fm_sheet_set_pivot_col_field_order(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      const uint32_t* indices, size_t count);

/** @brief Number of `<dataField>` entries on the pivot. */
FM_API fm_status_t fm_sheet_get_pivot_data_field_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       size_t* out_count);

/**
 * @brief Appends a new data-field entry. A `NULL` or empty `spec->name` is
 *        replaced by the engine-derived name described on
 *        `fm_pivot_data_field_spec_t`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `spec` or `out_idx` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `pivot_index` is out of range,
 *           `spec->aggregation` / `spec->show_as` names no
 *           declared enumerator, or `spec->field_index` or a set
 *           `spec->show_as_base_field` is not an index into the pivot's
 *           fields (the pivot is left unchanged);
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the cache has calculated items and
 *           `spec->aggregation` is Average, StdDev, StdDevP, Var or VarP.
 */
FM_API fm_status_t fm_sheet_add_pivot_data_field(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 const fm_pivot_data_field_spec_t* spec, size_t* out_idx);

/** @brief Drops every data-field entry from the pivot. */
FM_API fm_status_t fm_sheet_clear_pivot_data_fields(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Replaces the data-field entry at `data_field_idx` in place. A
 *        `NULL` or empty `spec->name` is derived as in
 *        `fm_sheet_add_pivot_data_field`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `spec` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `pivot_index` or
 *           `data_field_idx` is out of range, `spec->aggregation` /
 *           `spec->show_as` names no declared enumerator, or
 *           `spec->field_index` or a set `spec->show_as_base_field` is not
 *           an index into the pivot's fields (the pivot is left unchanged);
 *         `FM_STATUS_PRECONDITION_FAILED` (9) as for `fm_sheet_add_pivot_data_field`.
 */
FM_API fm_status_t fm_sheet_set_pivot_data_field(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 size_t data_field_idx, const fm_pivot_data_field_spec_t* spec);

/**
 * @brief Reads into `*out` the name a data field over pivot field
 *        `field_index` with `aggregation` would get from an empty
 *        `spec->name`, without changing anything. `exclude_data_field` is the
 *        entry `fm_sheet_set_pivot_data_field` would replace, whose name is
 *        not counted as taken; `SIZE_MAX` for none (the add case). The string
 *        is a read-scratch-backed view.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `pivot_index`,
 *           `field_index` or a `exclude_data_field` other than `SIZE_MAX`
 *           is out of range, or `aggregation` names no
 *           `fm_pivot_aggregation_t`.
 */
FM_API fm_status_t fm_sheet_get_pivot_data_field_default_name(const fm_workbook_t* wb, size_t sheet_index,
                                                              size_t pivot_index, size_t field_index,
                                                              int32_t aggregation, size_t exclude_data_field,
                                                              const char** out);

/*
 * Definition read-back. Every getter below is read-only and works the same
 * for a pivot created in this session and one loaded from a package. String
 * and index-list outputs are model-backed views: they remain valid until the
 * next mutation of the pivot (any `fm_workbook_pivot_*` setter on it), a
 * workbook model replacement, or handle destruction. Reads do not invalidate
 * them. Copy them when they must outlive those events. On any error the
 * output arguments are left unchanged.
 */

/**
 * @brief Reads the `cache_id` of the pivot cache the pivot draws from.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_cache_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_cache_id(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                               uint32_t* out_cache_id);

/**
 * @brief Reads the pivot's name. `*out_name` is a NUL-terminated UTF-8
 *        model-backed view, valid until the next mutation of the pivot or its
 *        sheet, or until the handle is destroyed.
 *
 * @return As `fm_sheet_get_pivot_cache_id`, with `out_name` in place of
 *         `out_cache_id`.
 */
FM_API fm_status_t fm_sheet_get_pivot_name(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                           const char** out_name);

/**
 * @brief Reads the pivot's options into `out`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_options(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                              fm_pivot_options_t* out);

/**
 * @brief Writes every member of `options` to the pivot. An `indent` or
 *        `page_wrap` equal to the pivot's current value is kept even when
 *        outside the settable range, so a definition read from a file
 *        survives a get / set round trip.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `options` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) for a malformed string;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if an index is out of range, `error_caption`,
 *           `missing_caption`, `alt_text` or `grand_total_caption` is longer
 *           than 255 UTF-16 code units, `indent` exceeds 126 or `page_wrap`
 *           exceeds 255;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the report would overlap another
 *           structure (the pivot is left unchanged).
 */
FM_API fm_status_t fm_sheet_set_pivot_options(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                              const fm_pivot_options_t* options);

/**
 * @brief Writes the pivot's report into its sheet cells now, as a pivot-table
 *        mutation does, and clears its pending flag. A pivot over a cache
 *        loaded without its records writes nothing.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `pivot_index` is out of range;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the report would overlap another
 *           structure; the pivot stays pending.
 */
FM_API fm_status_t fm_sheet_project_pivot(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Reads into `*out` whether the pivot's sheet cells are stale after a
 *        change to its cache (`1`) or current (`0`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_is_pivot_projection_pending(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                        int32_t* out);

/**
 * @brief Writes the cells behind the data or total cell (`row`, `col`) of the
 *        pivot onto a new sheet, as Excel's Show Details does.
 *
 * The new sheet is inserted immediately before the pivot's sheet, gets the
 * locale's default sheet name and holds one table of the matching source
 * records, with a header row. It becomes the active sheet and its index is
 * stored in `*out_sheet_index`. Sheets at and after the pivot's sheet index
 * shift by one, so indices held for them must be refreshed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_sheet_index` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if an index is out of range or the cell is not
 *         a data, subtotal or grand total cell of the pivot;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the pivot disables drill-down.
 */
FM_API fm_status_t fm_sheet_show_pivot_details(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index, uint32_t row,
                                               uint32_t col, size_t* out_sheet_index);

/**
 * @brief Reads the row / column grand total flags (`0` or `1`), the values
 *        `fm_sheet_set_pivot_grand_totals` writes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_rows` or `out_cols` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_grand_totals(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                   int32_t* out_rows, int32_t* out_cols);

/**
 * @brief Reads the pivot's table style into `out`, the values
 *        `fm_sheet_set_pivot_style_info` writes.
 *
 * @return As `fm_sheet_get_pivot_grand_totals`, with `out` in place of
 *         `out_rows` / `out_cols`.
 */
FM_API fm_status_t fm_sheet_get_pivot_style_info(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 fm_pivot_style_info_t* out);

/**
 * @brief Reads the definition of pivot field `field_idx` into `out`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `pivot_index` or `field_idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_field(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                            size_t field_idx, fm_pivot_field_info_t* out);

/**
 * @brief Number of items in pivot field `field_idx`'s item list, the list
 *        `fm_sheet_add_pivot_field_item` / `_add_item_at` append to.
 *
 * @return As `fm_sheet_get_pivot_field`, with `out_count` in place of `out`.
 */
FM_API fm_status_t fm_sheet_get_pivot_field_item_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       size_t field_idx, size_t* out_count);

/**
 * @brief Reads item `item_idx` of pivot field `field_idx` into `out`. The
 *        index is the one `fm_sheet_set_pivot_field_item_visible` takes.
 *
 * @return As `fm_sheet_get_pivot_field`, plus `FM_STATUS_INVALID_ARGUMENT` (2) when
 *         `item_idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_field_item(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 size_t field_idx, size_t item_idx, fm_pivot_item_info_t* out);

/**
 * @brief Reads data field `data_field_idx` into the shape
 *        `fm_sheet_add_pivot_data_field` accepts. `name` and
 *        `number_format` are views and never `NULL` (`number_format` is `""`
 *        for none); an unset `show_as_base_field` / `show_as_base_item` is `-1`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `pivot_index` or `data_field_idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_data_field(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 size_t data_field_idx, fm_pivot_data_field_spec_t* out);

/**
 * @brief Reads the row-axis field order: `*out_indices` points at
 *        `*out_count` indices into the pivot's fields, the list
 *        `fm_sheet_set_pivot_row_field_order` takes. `*out_indices` may
 *        be `NULL` when `*out_count` is `0`. With two or more data fields the
 *        axis holding the Values level lists `FM_PIVOT_VALUES_FIELD` at its
 *        position; when it is listed the view is read-scratch-backed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_indices` or `out_count` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_row_field_order(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      const uint32_t** out_indices, size_t* out_count);

/** @brief Reads the column-axis field order. Same contract as rows. */
FM_API fm_status_t fm_sheet_get_pivot_col_field_order(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      const uint32_t** out_indices, size_t* out_count);

/*
 * Persistence contract shared by the whole pivot-filter family
 * (`fm_sheet_get_pivot_filter_count`, `_at`, `_add`, `_clear`,
 * `_remove_at`):
 *
 * The list is the pivot's label / value / date filters and is saved as the
 * definition's `<filters>` block. Every entry of a loaded block whose type
 * this ABI enumerates appears here, and `_clear` / `_remove_at` remove it.
 * Only entries the model cannot express (an unknown type, a member-property
 * or OLAP measure filter, criteria that do not decode) are not listed; they
 * filter nothing and are re-emitted verbatim on save in their original
 * position. A loaded entry is saved byte-for-byte while it still applies to
 * the field and data field it was read with. An entry whose field no longer
 * resolves filters nothing and is not saved.
 */

/** @brief Number of filters on the pivot; see the persistence contract
 * above. */
FM_API fm_status_t fm_sheet_get_pivot_filter_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                   size_t* out_count);

/**
 * @brief Reads the filter at `filter_idx` without mutating the pivot.
 *
 * See the persistence contract above. A loaded entry names its field by
 * source name and reads back the axis its field sits on.
 *
 * `out` reuses the `fm_pivot_filter_spec_t` shape accepted by
 * `fm_sheet_add_pivot_filter`. `field_name` and (for a text primary
 * value) `value_text` are model-backed views; they remain valid until the
 * next mutation of this pivot's filter list, a workbook model replacement,
 * or handle destruction. Reads do not invalidate the views. Copy them when
 * they must outlive those events. For non-text primary values,
 * `value_text` is `NULL`; an unset upper bound has kind
 * `FM_PIVOT_FILTER_VALUE_NONE` and zero numeric payloads.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `pivot_index` or `filter_idx` is out of range, or
 *         the model contains an axis or filter type that is not representable by
 *         this C ABI. On any error,
 *         `*out` is left unchanged.
 */
FM_API fm_status_t fm_sheet_get_pivot_filter(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                             size_t filter_idx, fm_pivot_filter_spec_t* out);

/**
 * @brief Appends an active filter. `spec->field_name` must be non-NULL;
 *        `spec->value_text` must be non-NULL when
 *        `spec->value_kind == FM_PIVOT_FILTER_VALUE_TEXT`. The upper-bound
 *        payload is honoured only for the range types (`*_BETWEEN`,
 *        `*_NOT_BETWEEN`, `LABEL_DATE`).
 *
 * The entry is saved with the pivot; see the persistence contract above.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `spec`, `spec->field_name` or a
 *         required `spec->value_text` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `axis` or `type` is outside the C ABI enum
 *         ranges, the sheet/pivot indices do not resolve, the value payload
 *         discriminators are unusable, `data_field_index` is out of range
 *         for a value filter, or a label/date filter's `field_name` names no
 *         pivot field.
 */
FM_API fm_status_t fm_sheet_add_pivot_filter(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                             const fm_pivot_filter_spec_t* spec);

/** @brief Drops every listed filter from the pivot. Loaded entries outside
 * the list are kept; see the persistence contract above. */
FM_API fm_status_t fm_sheet_clear_pivot_filters(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/** @brief Removes the filter at `filter_idx`; see the persistence contract
 * above. */
FM_API fm_status_t fm_sheet_remove_pivot_filter(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                size_t filter_idx);

/* -------------------------------------------------------------------------- */
/* Dynamic-array spill payload                                                */
/* -------------------------------------------------------------------------- */

/**
 * @brief Result shape for `fm_sheet_get_spill_info`.
 *
 *   * `engaged` — `1` when `(row, col)` is part of a registered spill
 *     region, `0` otherwise.
 *   * `blocked` — `1` when `engaged == 0` and `(row, col)` is the anchor
 *     of a spill refused with `#SPILL!` whose attempted rectangle fits
 *     inside the grid, `0` otherwise. Only the anchor reports it; cells
 *     inside the attempted rectangle read `engaged == 0`, `blocked == 0`.
 *   * `anchor_row`, `anchor_col` — the region's anchor cell (the cell
 *     that holds the dynamic-array formula).
 *   * `rows`, `cols` — region dimensions including the anchor; for a
 *     blocked anchor, the rectangle the spill attempted.
 *
 * `anchor_*`, `rows` and `cols` carry meaning only when `engaged` or
 * `blocked` is `1`.
 *
 * Use `fm_sheet_get_value(sheet, anchor_row + r, anchor_col + c)`
 * to read individual cells; `fm_sheet_get_value` is already
 * spill-aware and returns the per-cell value verbatim.
 */
typedef struct {
  uint32_t anchor_row;
  uint32_t anchor_col;
  uint32_t rows;
  uint32_t cols;
  int32_t engaged; /* 0/1 */
  int32_t blocked; /* 0/1 */
} fm_spill_info_t;

/**
 * @brief Returns the spill-region info for `(sheet, row, col)`.
 *
 * If `(row, col)` is the anchor of a region, returns that region. If
 * it is a phantom of a region, returns the same region. If it is a
 * blocked anchor (see `blocked`), returns the attempted rectangle with
 * `out->blocked = 1`. Otherwise every field is zero.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_spill_info(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                           fm_spill_info_t* out);

/* -------------------------------------------------------------------------- */
/* Legacy array (CSE) formulas                                                */
/* -------------------------------------------------------------------------- */

/**
 * @brief Result shape for `fm_sheet_get_array_formula`.
 *
 *   * `engaged` — `1` when `(row, col)` belongs to a legacy array formula
 *     (`<f t="array" ref=...>`), `0` otherwise. The other fields carry
 *     meaning only when `engaged == 1`; `formula` is `NULL` otherwise.
 *   * `anchor_row`, `anchor_col` — the block's top-left cell, which holds
 *     the formula.
 *   * `rows`, `cols` — block dimensions including the anchor.
 *   * `formula` — the anchor's stored formula text, spelled as
 *     `fm_sheet_get_formula` returns it. Scratch-backed: valid until the
 *     next call on `wb` that publishes scratch text.
 *
 * Every cell of the block reports the same values.
 */
typedef struct {
  uint32_t anchor_row;
  uint32_t anchor_col;
  uint32_t rows;
  uint32_t cols;
  const char* formula;
  int32_t engaged; /* 0/1 */
} fm_array_formula_t;

/**
 * @brief Returns the legacy array formula covering `(sheet, row, col)`.
 *
 * A dynamic-array spill is not a legacy array formula; use
 * `fm_sheet_get_spill_info` for that. When the cell is outside any block,
 * sets `out->engaged = 0` and leaves the remaining fields zero / `NULL`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the coordinate
 *         is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_array_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              fm_array_formula_t* out);

/**
 * @brief Enters `formula` as one legacy array formula over the `rows` x
 *        `cols` rectangle whose top-left cell is `(row, col)`.
 *
 * Every cell of the rectangle loses its value and formula (styles are
 * kept); a recalc distributes the result, filling `#N/A` where a smaller
 * result cannot reach. A block wholly inside the rectangle is replaced.
 * Writing a value, blank or formula to the anchor of a block replaces the
 * whole block, which is how a block is cleared; a write to any other cell
 * of a block is refused. Structured references are bound at `(row, col)`
 * as `fm_sheet_set_formula` binds them.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `formula` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `rows` or `cols`
 *         is zero, or the rectangle leaves the grid or exceeds the
 *         dynamic-array cell ceiling;
 *         a parser code when `formula` does not parse;
 *         `FM_STATUS_PRECONDITION_FAILED` (9), changing nothing, when the rectangle
 *         partially overlaps another array block, intersects a merged
 *         area, a data table or a PivotTable, or (for more than one cell) a
 *         table.
 */
FM_API fm_status_t fm_sheet_set_array_formula(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              uint32_t rows, uint32_t cols, const char* formula);

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
 * @brief Result shape for `fm_get_function_metadata`.
 *
 * `canonical_name` is always populated when the function is known.
 * `min_arity` / `max_arity` bound the written argument count Excel
 * accepts at entry, the bounds formula entry enforces; `max_arity` is
 * `0xFFFFFFFFu` (i.e. `eval::kVariadic`) for an open argument list.
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
 * `docs/function-metadata-schema.md`; the WASM, native Node and Python
 * bindings ship pure merge helpers (`mergeFunctionMetadata` /
 * `merge_function_metadata`) for it. This
 * metadata is display-only: formula input parsing stays fixed to the
 * English canonical names, so `fm_localize_function_name` /
 * `fm_canonicalize_function_name` are unaffected by any injected document
 * (formula input parsing never accepts localized names).
 *
 * String storage is process-static (the catalog is initialised at
 * static-init time); callers do not free the returned pointers.
 */
typedef struct {
  const char* canonical_name;
  uint32_t min_arity;
  uint32_t max_arity;
  int32_t availability; /* fm_function_availability_t */
  /* Always `NULL` - supplied by the host's provider document. */
  const char* signature_template;
  /* Always `NULL` - supplied by the host's provider document. */
  const char* description;
} fm_function_metadata_t;

/**
 * @brief Returns metadata for the function `name`.
 *
 * `name` is matched case-insensitively against the canonical name. On
 * success, `*out` receives a view into the catalog's static storage;
 * the views are valid for the lifetime of the process. The metadata is
 * locale-invariant.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `name` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when no function matches `name`.
 */
FM_API fm_status_t fm_get_function_metadata(const char* name, fm_function_metadata_t* out);

/**
 * @brief Writes the total number of registered Formulon functions to
 *        `*out_count`; `FM_STATUS_BINDING_NULL_POINTER` (7001) when
 *        `out_count` is NULL.
 *
 * Drives the canonical-name iterator (`fm_get_function_name_at`).
 */
FM_API fm_status_t fm_get_function_count(size_t* out_count);

/**
 * @brief Returns the canonical name of the `idx`-th registered function.
 *
 * Order is sorted ascending so consumers can build deterministic UI
 * lists. `*out_name` is a static view into the catalog's process-static
 * storage.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out_name` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_get_function_name_at(size_t idx, const char** out_name);

/**
 * @brief Returns the localized display name for `canonical_name` under the
 *        profile `profile_id`, or `canonical_name` itself when the locale
 *        has no alias.
 *
 * `*out_localized` is a static view into process-static storage and must not
 * be freed. `profile_id` is one of `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,
 * zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}`. `de-DE`,
 * `fr-FR`, `es-ES`, `es-MX`, `pt-BR`, `ru-RU`, `it-IT` and `nl-NL` return the
 * function names Excel shows in that locale; `en-US`, `zh-CN`, `ko-KR` and
 * `th-TH` show the canonical names and return `canonical_name` unchanged.
 * `ja-JP` renames only `DOLLAR` (`YEN`), `USDOLLAR` (`DOLLAR`) and `DBCS`
 * (`JIS`); `zh-TW` renames only `DBCS` (`BIG5`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`) or `canonical_name` does not match
 *           a registered function.
 */
FM_API fm_status_t fm_localize_function_name(const char* canonical_name, const char* profile_id,
                                             const char** out_localized);

/**
 * @brief Returns the canonical (English UPPERCASE) name for the
 *        localized function `localized_name` under the profile
 *        `profile_id`.
 *
 * Inverts the per-locale name table of `fm_localize_function_name`
 * (ASCII case-insensitive), accepting the canonical spelling as a fallback; a
 * locale without renames reduces to that canonical match. `*out_canonical` is a
 * static view into process-static storage and must not be freed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`) or no function matches.
 */
FM_API fm_status_t fm_canonicalize_function_name(const char* localized_name, const char* profile_id,
                                                 const char** out_canonical);

/**
 * @brief Rewrites stored (en-invariant) formula text into `profile_id`'s
 *        spelling.
 *
 * Function names, list / array separators, decimal points, booleans and
 * error names take the locale form; strings, quoted sheet names, external
 * qualifiers and structured-reference brackets are kept verbatim. A number
 * literal with an exponent is regenerated (`1.5E-3` becomes `0.0015`). A
 * leading `=` is preserved. This is a pure text rewrite: the grammar is not
 * validated.
 *
 * `*out_formula` is a thread-local view, valid until the next
 * `fm_localize_formula` / `fm_canonicalize_formula` call on the same
 * thread. Copy it before making another call when it must be retained.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`).
 */
FM_API fm_status_t fm_localize_formula(const char* formula, const char* profile_id, const char** out_formula);

/**
 * @brief Rewrites `profile_id`-spelled formula text into the en-invariant
 *        form; the inverse of `fm_localize_formula`.
 *
 * `*out_formula` is the same thread-local view as `fm_localize_formula`
 * returns; the rewrite scope and status codes are also the same.
 */
FM_API fm_status_t fm_canonicalize_formula(const char* formula, const char* profile_id, const char** out_formula);

/** @brief Order in which a locale writes day, month and year. */
typedef enum { FM_DATE_ORDER_MDY = 0, FM_DATE_ORDER_YMD = 1, FM_DATE_ORDER_DMY = 2 } fm_date_order_t;

/**
 * @brief Scalar locale facts of one profile (`fm_get_locale_facts`).
 *
 * Every string is a NUL-terminated UTF-8 view into process-static storage.
 * `measured` is `1` for the profiles backed by Excel (every `mac-*`
 * profile and `win-365-ja_JP`) and `0` for the estimated `win-*` ones; the
 * other facts hold in every Mac locale. Collation backing is
 * separate: `collation_measured` is `1` only for `mac-*` profiles, so
 * `win-365-ja_JP` is Excel-backed for facts but its collation order is estimated.
 */
typedef struct {
  const char* decimal_separator;
  const char* group_separator;
  const char* list_separator;
  const char* array_column_separator;
  const char* array_row_separator;
  const char* true_name;
  const char* false_name;
  int32_t date_order; /* fm_date_order_t */
  const char* currency_symbol;
  int32_t currency_suffix; /* 1: the symbol follows the number */
  int32_t currency_space;  /* 1: one space between symbol and number */
  int32_t currency_default_decimals;
  int32_t measured;           /* 1: Excel-backed profile; 0: estimated */
  int32_t collation_measured; /* 1: collation order Excel-backed (mac-*); 0: estimated (win-*) */
} fm_locale_facts_t;

/**
 * @brief Fills `*out` with the locale facts of `profile_id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `profile_id` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`).
 */
FM_API fm_status_t fm_get_locale_facts(const char* profile_id, fm_locale_facts_t* out);

/**
 * @brief Which Excel text fold a collation call applies.
 *
 * Selects the text fold only; type boundaries (number, text, boolean, error,
 * blank) never depend on it.
 */
typedef enum {
  FM_COLLATION_COMPARE = 0,      /* the fold of =, criteria, SEARCH and approximate lookup */
  FM_COLLATION_DYNAMIC = 1,      /* the fold of SORT, UNIQUE, XLOOKUP, GROUPBY and pivots */
  FM_COLLATION_CLASSIC_EXACT = 2 /* the fold of MATCH 0 and VLOOKUP FALSE */
} fm_collation_strength_t;

/**
 * @brief Compares two texts under `profile_id`'s collation.
 *
 * `*out_order` is the sign (-1, 0, 1) of comparing `a` with `b`; it is a text
 * order, not `=` on two cells. Both strings are NUL-terminated UTF-8.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`) or `strength` is out of range.
 */
FM_API fm_status_t fm_compare_text(const char* a, const char* b, const char* profile_id, int32_t strength,
                                   int32_t* out_order);

/**
 * @brief Classifies `count` values into equivalence classes.
 *
 * Writes one class id per value to `out_class_ids` (`count` entries) and the
 * number of distinct classes to `*out_class_count`. Ids are dense from 0 in
 * ascending SORT order (number < text < boolean < error by code < blank).
 * Two values share an id only when they have the same kind and an equal
 * value: text by the `strength` fold of `profile_id`, numbers by value
 * (`-0` equals `0`), errors by code; blanks are always one class.
 *
 * Usage: deduplicate by keeping the first index per id; sort by a stable
 * sort on the id, keeping blanks last when descending; re-apply a saved
 * selection by classifying the saved values and the column values in ONE
 * call and picking the matching ids. Ids are meaningful only within one
 * call.
 *
 * When `count` is `0`, `values` and `out_class_ids` may be `NULL`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `profile_id` or `out_class_count` is
 *           `NULL`, or `values` / `out_class_ids` is `NULL` with `count > 0`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile,
 *           `strength` is out of range, `count` exceeds `UINT32_MAX`, or a
 *           value is a non-finite number, text with a `NULL` pointer, an
 *           out-of-range error code, or an array / reference / lambda kind.
 */
FM_API fm_status_t fm_get_collation_classes(const fm_value_t* values, size_t count, const char* profile_id,
                                            int32_t strength, uint32_t* out_class_ids, uint32_t* out_class_count);

/**
 * @brief Writes the number of error values, the bound of `index` in
 *        `fm_get_locale_error_name_at`, to `*out_count`;
 *        `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out_count` is NULL.
 *
 * `index` uses the `fm_error_code_t` numbering (`4` is `#NAME?`).
 */
FM_API fm_status_t fm_get_locale_error_name_count(size_t* out_count);

/**
 * @brief Returns the canonical and the `profile_id`-localized spelling of
 *        the error at `index`.
 *
 * Both strings are static views into process-static storage. `*out_measured` is `1` when the
 * localized spelling comes from Excel under an Excel-backed profile and
 * `0` when it is assumed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           or `index` is out of range.
 */
FM_API fm_status_t fm_get_locale_error_name_at(const char* profile_id, size_t index, const char** out_canonical,
                                               const char** out_localized, int32_t* out_measured);

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
 * `top_left_*` is the scroll origin (`<sheetView topLeftCell>`, A1 when
 * absent). `pane_top_left_*` is `<pane topLeftCell>` of a frozen view: the
 * first visible cell of the unfrozen area, which is the first cell past the
 * frozen rows and columns until a host scrolls it; `0, 0` when nothing is
 * frozen.
 *
 * `split_x` / `split_y` are an unfrozen split's position in twips (`<pane
 * xSplit ySplit>`); `0` for no split on that axis, and always `0` while the
 * view is frozen.
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
  uint32_t top_left_row;        /* scroll origin of the sheet view (`topLeftCell`), 0-based */
  uint32_t top_left_col;
  uint32_t pane_top_left_row; /* first visible cell of a frozen or split view's lower-right pane */
  uint32_t pane_top_left_col;
  double split_x; /* unfrozen split position in twips; 0 = none on that axis */
  double split_y;
  int32_t show_formulas;        /* 0/1; OOXML default 0 */
  int32_t show_outline_symbols; /* 0/1; OOXML default 1 */
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
  uint8_t collapsed; /* 0/1; `collapsed` attribute of the span */
  uint8_t _pad[2];
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
 *
 * `has_height` is the presence of a stored height; `height` is meaningful
 * only when it is set. `custom_height` separates an explicit override
 * (`customHeight="1"`, the state `fm_sheet_set_row_height` produces) from an
 * auto height that merely caches `ht`, which Excel recomputes on its own.
 */
typedef struct {
  uint32_t row;
  double height;
  int32_t hidden; /* 0/1 */
  uint8_t outline_level;
  uint8_t collapsed; /* 0/1; `collapsed` attribute of the row */
  uint8_t _pad[2];
  /* Row s= is effective only when customFormat=1 in OOXML. */
  int32_t has_style; /* 0/1 */
  uint32_t style_xf;
  int32_t has_height;    /* 0/1 */
  int32_t custom_height; /* 0/1 */
} fm_row_layout_t;

/**
 * @brief Returns the column-layout-override count for `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_col_override_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `idx`-th column-layout override.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_col_override_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                fm_column_layout_t* out);

/**
 * @brief Returns the row-layout-override count for `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_override_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `idx`-th row-layout override.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_override_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                fm_row_layout_t* out);

/**
 * @brief Reads the sheet-view state (zoom, frozen panes, tab hidden, and
 *        the display / orientation flags).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_t* out);

/**
 * @brief Sets the sheet view's scroll origin (`<sheetView topLeftCell>`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the cell is out of range.
 */
FM_API fm_status_t fm_sheet_set_top_left(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col);

/**
 * @brief Sets the first visible cell of a frozen or split view's lower-right
 *        pane (`<pane topLeftCell>`).
 *
 * In a frozen view the cell must lie in the unfrozen area: `row >=
 * freeze_rows` and `col >= freeze_cols`. Changing the freeze with `fm_sheet_set_freeze`
 * resets this to the first unfrozen cell.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet
 *         has neither frozen nor split panes, or the cell is not in the
 *         unfrozen area.
 */
FM_API fm_status_t fm_sheet_set_pane_top_left(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col);

/**
 * @brief Splits the sheet view at `x_twips` / `y_twips` (`<pane xSplit
 *        ySplit>` without a freeze state); `0` for no split on that axis.
 *
 * A split replaces any freeze, as `fm_sheet_set_freeze` replaces a split;
 * `(0, 0)` removes it. A change resets the pane origin and active pane and
 * keeps only the active selection, re-homed to the new active pane. Setting
 * the current split is a no-op.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a value is
 *         negative or not finite (the view is left unchanged).
 */
FM_API fm_status_t fm_sheet_set_split(fm_workbook_t* wb, size_t sheet_index, double x_twips, double y_twips);

/**
 * @brief Reads the active pane's selection: the active cell and its ranges.
 *
 * `*out_total` receives the number of selected ranges; at most `capacity` of
 * them are written to `out_ranges`, which may be `NULL` when `capacity` is
 * `0` (size the buffer first). The selection reads as it is saved: a pane
 * that stores none reports `A1` and one range, `A1` (the pane's top-left
 * cell for a frozen or split view's active pane); a selection without an
 * explicit `sqref` is the active cell alone; an active cell outside every
 * range moves to the first range's top-left. Whole-row and whole-column
 * references read as full-width / full-height ranges.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_active_row`, `out_active_col`
 *         or `out_total` is `NULL`, or `out_ranges` is `NULL` while
 *         `capacity > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_selection(const fm_workbook_t* wb, size_t sheet_index, uint32_t* out_active_row,
                                          uint32_t* out_active_col, fm_merge_range* out_ranges, size_t capacity,
                                          size_t* out_total);

/**
 * @brief Replaces the active pane's selection (`<selection>`): the active
 *        cell and the selected ranges (`sqref`). `count == 0` selects the
 *        active cell alone.
 *
 * The selection of a frozen or split view belongs to the pane it is active
 * in (`bottomRight`, `topRight` or `bottomLeft`, see `fm_sheet_set_freeze`);
 * other panes' selections are kept. An active cell outside every range is
 * saved as the first range's top-left, as Excel does.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, or `ranges` is `NULL` with
 *           a non-zero `count`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the active cell or a range is out of range, or a
 *         range has `first > last`.
 */
FM_API fm_status_t fm_sheet_set_selection(fm_workbook_t* wb, size_t sheet_index, uint32_t active_row,
                                          uint32_t active_col, const fm_merge_range* ranges, size_t count);

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
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
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
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if `xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a malformed fragment.
 */
FM_API fm_status_t fm_sheet_set_auto_filter_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/* -------------------------------------------------------------------------- */
/* Typed AutoFilter                                                           */
/* -------------------------------------------------------------------------- */

/** @brief `<dateGroupItem>` granularity (`fm_date_group_item_t::grouping`); mirrors `ST_DateTimeGrouping`. */
typedef enum {
  FM_DATE_TIME_GROUPING_YEAR = 0,
  FM_DATE_TIME_GROUPING_MONTH = 1,
  FM_DATE_TIME_GROUPING_DAY = 2,
  FM_DATE_TIME_GROUPING_HOUR = 3,
  FM_DATE_TIME_GROUPING_MINUTE = 4,
  FM_DATE_TIME_GROUPING_SECOND = 5,
} fm_date_time_grouping_t;

/** @brief Criterion selector (`fm_filter_column_t::kind`). */
typedef enum {
  FM_FILTER_KIND_NONE = 0,
  FM_FILTER_KIND_VALUES = 1,
  FM_FILTER_KIND_CUSTOM = 2,
  FM_FILTER_KIND_TOP10 = 3,
  FM_FILTER_KIND_DYNAMIC = 4,
  FM_FILTER_KIND_COLOR = 5,
  FM_FILTER_KIND_ICON = 6,
} fm_filter_kind_t;

/** @brief Custom-filter comparison operator (`fm_filter_column_t::op1` / `op2`); mirrors `ST_FilterOperator`. */
typedef enum {
  FM_FILTER_OPERATOR_EQUAL = 0,
  FM_FILTER_OPERATOR_LESS_THAN = 1,
  FM_FILTER_OPERATOR_LESS_THAN_OR_EQUAL = 2,
  FM_FILTER_OPERATOR_NOT_EQUAL = 3,
  FM_FILTER_OPERATOR_GREATER_THAN_OR_EQUAL = 4,
  FM_FILTER_OPERATOR_GREATER_THAN = 5,
} fm_filter_operator_t;

/** @brief Dynamic-filter type (`fm_filter_column_t::dynamic_type`); `ST_DynamicFilterType` in schema order. */
typedef enum {
  FM_DYNAMIC_FILTER_NULL = 0,
  FM_DYNAMIC_FILTER_ABOVE_AVERAGE = 1,
  FM_DYNAMIC_FILTER_BELOW_AVERAGE = 2,
  FM_DYNAMIC_FILTER_TOMORROW = 3,
  FM_DYNAMIC_FILTER_TODAY = 4,
  FM_DYNAMIC_FILTER_YESTERDAY = 5,
  FM_DYNAMIC_FILTER_NEXT_WEEK = 6,
  FM_DYNAMIC_FILTER_THIS_WEEK = 7,
  FM_DYNAMIC_FILTER_LAST_WEEK = 8,
  FM_DYNAMIC_FILTER_NEXT_MONTH = 9,
  FM_DYNAMIC_FILTER_THIS_MONTH = 10,
  FM_DYNAMIC_FILTER_LAST_MONTH = 11,
  FM_DYNAMIC_FILTER_NEXT_QUARTER = 12,
  FM_DYNAMIC_FILTER_THIS_QUARTER = 13,
  FM_DYNAMIC_FILTER_LAST_QUARTER = 14,
  FM_DYNAMIC_FILTER_NEXT_YEAR = 15,
  FM_DYNAMIC_FILTER_THIS_YEAR = 16,
  FM_DYNAMIC_FILTER_LAST_YEAR = 17,
  FM_DYNAMIC_FILTER_YEAR_TO_DATE = 18,
  FM_DYNAMIC_FILTER_Q1 = 19,
  FM_DYNAMIC_FILTER_Q2 = 20,
  FM_DYNAMIC_FILTER_Q3 = 21,
  FM_DYNAMIC_FILTER_Q4 = 22,
  FM_DYNAMIC_FILTER_M1 = 23,
  FM_DYNAMIC_FILTER_M2 = 24,
  FM_DYNAMIC_FILTER_M3 = 25,
  FM_DYNAMIC_FILTER_M4 = 26,
  FM_DYNAMIC_FILTER_M5 = 27,
  FM_DYNAMIC_FILTER_M6 = 28,
  FM_DYNAMIC_FILTER_M7 = 29,
  FM_DYNAMIC_FILTER_M8 = 30,
  FM_DYNAMIC_FILTER_M9 = 31,
  FM_DYNAMIC_FILTER_M10 = 32,
  FM_DYNAMIC_FILTER_M11 = 33,
  FM_DYNAMIC_FILTER_M12 = 34,
} fm_dynamic_filter_type_t;

/** @brief Sort key source (`fm_sort_condition_t::sort_by`); mirrors `ST_SortBy`. */
typedef enum {
  FM_SORT_BY_VALUE = 0,
  FM_SORT_BY_CELL_COLOR = 1,
  FM_SORT_BY_FONT_COLOR = 2,
  FM_SORT_BY_ICON = 3,
} fm_sort_by_t;

/** @brief Sort method (`fm_auto_filter_t::sort_method`); mirrors `ST_SortMethod`. */
typedef enum {
  FM_SORT_METHOD_NONE = 0,
  FM_SORT_METHOD_PIN_YIN = 1,
  FM_SORT_METHOD_STROKE = 2,
} fm_sort_method_t;

/**
 * @brief One `<dateGroupItem>` of a values filter.
 *
 * `grouping` is the granularity the item matches at, an
 * `fm_date_time_grouping_t`. Fields finer than `grouping` are zero.
 */
typedef struct {
  uint16_t year;
  uint8_t month;
  uint8_t day;
  uint8_t hour;
  uint8_t minute;
  uint8_t second;
  uint8_t grouping;
} fm_date_group_item_t;

/**
 * @brief One `<filterColumn>` of an AutoFilter.
 *
 * `col_id` is the 0-based offset from the AutoFilter range's first column.
 * `kind` selects the criterion, and only that criterion's fields are read on
 * input; on output the fields of every other criterion are zero.
 *   * `kind` -- an `fm_filter_kind_t`.
 *   * values -- `filter_blank` matches blank cells; `values` lists displayed
 *     texts and `date_groups` date-group items.
 *   * custom -- `custom_count` (1 or 2) conditions `op1`/`val1` and
 *     `op2`/`val2`, joined by AND when `custom_and` is set. Operators
 *     are `fm_filter_operator_t`.
 *   * top10 -- `top` (else bottom), `percent`, the rank `top_val`, and the
 *     `filter_val` threshold Excel last stored when `has_filter_val`.
 *   * dynamic -- `dynamic_type` is an `fm_dynamic_filter_type_t`; `dyn_val` / `dyn_max_val` are the
 *     stored bounds when their `has_` flag is set, and `val_iso` /
 *     `max_val_iso` their ISO spellings.
 *   * color -- `dxf_id` names the differential format whose fill
 *     (`cell_color` set) or font colour matches; `UINT32_MAX` means absent.
 *   * icon -- `icon_set` is the `formulon::cf::IconSetName` ordinal and
 *     `icon_id` the icon index when `has_icon_id` (else "no icon").
 *
 * On input every `const char*` may be `NULL`, meaning empty, and an array
 * pointer may be `NULL` only with a zero count.
 */
typedef struct {
  uint32_t col_id;
  int32_t hidden_button;
  int32_t show_button;
  int32_t kind;
  int32_t filter_blank;
  const char* const* values;
  uint32_t value_count;
  const fm_date_group_item_t* date_groups;
  uint32_t date_group_count;
  int32_t custom_and;
  int32_t custom_count;
  int32_t op1;
  const char* val1;
  int32_t op2;
  const char* val2;
  int32_t top;
  int32_t percent;
  int32_t has_filter_val;
  double top_val;
  double filter_val;
  int32_t dynamic_type;
  int32_t has_dyn_val;
  int32_t has_dyn_max_val;
  double dyn_val;
  double dyn_max_val;
  const char* val_iso;
  const char* max_val_iso;
  uint32_t dxf_id;
  int32_t cell_color;
  int32_t icon_set;
  int32_t icon_id;
  int32_t has_icon_id;
} fm_filter_column_t;

/**
 * @brief One `<sortCondition>` of an AutoFilter's sort state.
 *
 * `sort_by` is an `fm_sort_by_t`. `dxf_id` is read
 * only when `has_dxf_id`, `icon_id` only when `has_icon_id`; `icon_set` is
 * the `formulon::cf::IconSetName` ordinal.
 */
typedef struct {
  fm_merge_range ref;
  int32_t descending;
  int32_t sort_by;
  const char* custom_list;
  uint32_t dxf_id;
  int32_t has_dxf_id;
  int32_t icon_set;
  int32_t icon_id;
  int32_t has_icon_id;
} fm_sort_condition_t;

/**
 * @brief A sheet or table AutoFilter.
 *
 * `range` includes the header row. `columns` must be in strictly ascending
 * `col_id` order inside the range. The sort state is present when
 * `has_sort`; `sort_method` is an `fm_sort_method_t`.
 *
 * Attributes and extensions the record does not carry (`xr:uid`, unknown
 * `filterColumn` children, `extLst` payloads, a values filter's
 * `calendarType`) are kept across a set: they are carried over from the
 * AutoFilter being replaced, per column by `col_id`, so a get followed by a
 * set of the unchanged record leaves the stored element as it was.
 */
typedef struct {
  fm_merge_range range;
  const fm_filter_column_t* columns;
  uint32_t column_count;
  int32_t has_sort;
  fm_merge_range sort_ref;
  int32_t column_sort;
  int32_t case_sensitive;
  int32_t sort_method;
  const fm_sort_condition_t* conditions;
  uint32_t condition_count;
} fm_auto_filter_t;

/**
 * @brief Reads the sheet AutoFilter.
 *
 * `*out_present` is 0 and `*out` all-zero when the sheet has no AutoFilter.
 * Every string and array reachable from `*out` is a read-scratch-backed view
 * and may be invalidated by the next successful scratch-backed read on the
 * same handle, by a mutation, or by handle destruction.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when the AutoFilter did not parse into
 *         the typed model (only `fm_sheet_get_auto_filter_xml` reads it).
 */
FM_API fm_status_t fm_sheet_get_auto_filter(const fm_workbook_t* wb, size_t sheet_index, fm_auto_filter_t* out,
                                            int32_t* out_present);

/**
 * @brief Replaces the sheet AutoFilter and maintains the sheet's hidden
 *        `_xlnm._FilterDatabase` name over its range.
 *
 * Row visibility is not changed; call `fm_sheet_apply_auto_filter` to hide
 * the rows the criteria exclude.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `filter` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an
 *         enumeration field is outside the domain listed on the record, or an
 *         array pointer is `NULL` with a non-zero count;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when the range is outside the grid,
 *         column ids are outside the range or not ascending, a custom filter
 *         has other than one or two conditions, or a date group item does
 *         not fit its grouping.
 */
FM_API fm_status_t fm_sheet_set_auto_filter(fm_workbook_t* wb, size_t sheet_index, const fm_auto_filter_t* filter);

/**
 * @brief Removes the sheet AutoFilter and its `_xlnm._FilterDatabase` name.
 *        Row visibility is not changed. No-op when the sheet has none.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_auto_filter(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Evaluates the sheet AutoFilter and sets every body row's hidden flag
 *        to the negation of its match, as Excel's Reapply does; a matching
 *        row hidden by hand is shown again, and rows outside the body are
 *        untouched.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when the sheet has no AutoFilter;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when it did not parse into the model.
 */
FM_API fm_status_t fm_sheet_apply_auto_filter(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Drops every column criterion of the sheet AutoFilter and shows all
 *        of its body rows. The range, the drop-down buttons and the sort
 *        state stay.
 *
 * @return as `fm_sheet_apply_auto_filter`.
 */
FM_API fm_status_t fm_sheet_clear_auto_filter(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Evaluates the sheet AutoFilter without changing the sheet.
 *
 * Body row `*out_first_row + i` matches the criteria when `buf[i]` is
 * 1. `*out_len` always receives the number of body rows; at most `cap`
 * flags are written. Call with `cap == 0` (`buf` may then be `NULL`)
 * to size the buffer first.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_len` or `out_first_row` is
 *         `NULL`, or `buf` is `NULL` while `cap > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when the sheet has no AutoFilter;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when it did not parse into the model.
 */
FM_API fm_status_t fm_sheet_evaluate_auto_filter(const fm_workbook_t* wb, size_t sheet_index, uint8_t* buf, size_t cap,
                                                 size_t* out_len, uint32_t* out_first_row);

/**
 * @brief Table counterparts of the sheet AutoFilter entry points.
 *
 * `table_index` addresses the workbook table list as `fm_workbook_get_table`
 * does, and the table's sheet is the one whose rows are evaluated and
 * hidden. Tables carry no `_FilterDatabase` name. The contracts are those
 * of the sheet functions, with `FM_STATUS_INVALID_ARGUMENT` (2) for an out-of-range
 * `table_index`.
 */
FM_API fm_status_t fm_workbook_get_table_auto_filter(const fm_workbook_t* wb, size_t table_index, fm_auto_filter_t* out,
                                                     int32_t* out_present);
FM_API fm_status_t fm_workbook_set_table_auto_filter(fm_workbook_t* wb, size_t table_index,
                                                     const fm_auto_filter_t* filter);
FM_API fm_status_t fm_workbook_remove_table_auto_filter(fm_workbook_t* wb, size_t table_index);
FM_API fm_status_t fm_workbook_apply_table_auto_filter(fm_workbook_t* wb, size_t table_index);
FM_API fm_status_t fm_workbook_clear_table_auto_filter(fm_workbook_t* wb, size_t table_index);
FM_API fm_status_t fm_workbook_evaluate_table_auto_filter(const fm_workbook_t* wb, size_t table_index, uint8_t* buf,
                                                          size_t cap, size_t* out_len, uint32_t* out_first_row);

/**
 * @brief Sets or replaces the column width override for the inclusive
 *        column span `[first, last]`. Existing overrides whose span
 *        intersects the requested span are merged so the new width
 *        takes precedence on overlapping columns; non-overlapping
 *        portions of pre-existing entries are retained.
 *
 * `width` is a stored column width: Excel's character-width unit with the
 * column padding included, as `<col width>` holds it (Calibri 11 stores the
 * 8.43-character column as 9.140625). A rejected call leaves the layout
 * exactly as it was.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, `last` is past the last column (16383), or
 *         `width` is negative, NaN, or infinite.
 */
FM_API fm_status_t fm_sheet_set_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                             double width);

/**
 * @brief Removes the width override of the inclusive column span
 *        `[first, last]`, returning those columns to the sheet default
 *        width.
 *
 * Clears only the stored width; hidden state, outline level and column
 * style are kept. Overlapping spans are split at the boundaries, and a span
 * left with no state (no width, visible, outline level 0, no style) is
 * dropped. Columns with no width override are unaffected. A rejected call
 * leaves the layout exactly as it was.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last column (16383).
 */
FM_API fm_status_t fm_sheet_clear_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/**
 * @brief Sets or replaces the column hidden flag for the inclusive
 *        column span `[first, last]`. See `fm_sheet_set_column_width`
 *        for the merge semantics.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last column (16383).
 */
FM_API fm_status_t fm_sheet_set_column_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                              int32_t hidden);

/**
 * @brief Sets or replaces the column outline level for the inclusive
 *        column span `[first, last]`. See `fm_sheet_set_column_width`
 *        for the merge semantics.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, `last` is past the last column (16383), or
 *         `level` is above 7, the deepest outline level Excel stores.
 */
FM_API fm_status_t fm_sheet_set_column_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                               uint8_t level);

/**
 * @brief Sets or clears the `collapsed` flag (the outline group summarised by
 *        the span is collapsed) for the inclusive column span `[first, last]`.
 *        See `fm_sheet_set_column_width` for the merge semantics.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last column (16383).
 */
FM_API fm_status_t fm_sheet_set_column_collapsed(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                 int32_t collapsed);

/**
 * @brief Sets or replaces the cell-style XF for the inclusive column span
 *        `[first, last]`.
 *
 * The style applies to cells in the span that do not carry a cell-level XF.
 * Existing width, hidden and outline attributes are retained, and overlapping
 * column spans are split at the requested boundaries. `xf_index` addresses
 * the workbook's cell-XF table; index 0 is the default cell style.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the span or `xf_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_column_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                             size_t xf_index);

/**
 * @brief Removes the cell-style XF from the inclusive column span
 *        `[first, last]`.
 *
 * Width, hidden and outline attributes are retained. Empty spans left with no
 * observable state are removed.
 */
FM_API fm_status_t fm_sheet_clear_column_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/**
 * @brief Sets or replaces the row height override for `row`.
 *
 * `height` is in points and is stored as a manual height
 * (`customHeight="1"`), which Excel keeps when the row's content or font
 * changes. A rejected call leaves the layout exactly as it was.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row`
 *         is past the last row (1048575), or `height` is negative, NaN,
 *         or infinite.
 */
FM_API fm_status_t fm_sheet_set_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height);

/**
 * @brief Stores `height` as the cached value of `row`'s automatic height.
 *
 * Same contract as `fm_sheet_set_row_height`, except the row stays
 * automatic (`custom_height == 0`): the height is saved as `ht` without
 * `customHeight`, so the row still follows its content and font. Row
 * style, hidden state, outline level and collapsed flag are kept. Use it
 * to refresh a derived height; `fm_sheet_clear_row_height` drops it.
 *
 * @return As `fm_sheet_set_row_height`.
 */
FM_API fm_status_t fm_sheet_set_row_auto_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height);

/**
 * @brief Sets or replaces the row hidden flag for `row`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row`
 *         is past the last row (1048575).
 */
FM_API fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t hidden);

/**
 * @brief Sets or replaces the row outline level for `row`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row`
 *         is past the last row (1048575), or `level` is above 7, the
 *         deepest outline level Excel stores.
 */
FM_API fm_status_t fm_sheet_set_row_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint8_t level);

/**
 * @brief Sets or clears the `collapsed` flag (the outline group summarised by
 *        the row is collapsed) for `row`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row`
 *         is past the last row (1048575).
 */
FM_API fm_status_t fm_sheet_set_row_collapsed(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t collapsed);

/**
 * @brief Sets or replaces the cell-style XF for the inclusive row span
 *        `[first, last]`.
 *
 * A row style is effective only when the row has custom-format state, which
 * this function sets. Existing height, custom-height, hidden and outline
 * attributes are retained. `xf_index` addresses the workbook's cell-XF table;
 * index 0 is the default cell style. Nothing changes on failure.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the span or `xf_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_row_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                          size_t xf_index);

/**
 * @brief Removes the cell-style XF from the inclusive row span
 *        `[first, last]`.
 *
 * Existing height, custom-height, hidden and outline attributes are retained;
 * an otherwise empty row override is removed.
 */
FM_API fm_status_t fm_sheet_clear_row_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/*
 * Attribute-only formatting of a selection. The rect
 * `[(first_row, first_col), (last_row, last_col)]` (corners in either order)
 * is a selection: spanning every row makes it a column selection, every
 * column a row selection, both the whole sheet. To change some attributes
 * (say the font size) and keep the rest, read the XFs the selection resolves
 * to with `fm_sheet_get_range_style_bases`, build each one's patched XF with
 * `fm_workbook_add_font` / `fm_workbook_add_cell_xf`, and apply the pairs with
 * `fm_sheet_remap_range_style`. Every cell then resolves to the patched form
 * of the XF it resolved to before, as Excel applies such a change: covered
 * column and row styles are rewritten, existing cells are remapped, and a
 * cell is materialised only where the row-over-column precedence would
 * otherwise resolve to something else (for a plain range, every cell). Blank
 * cells of a covered row or column are not enumerated. Only style state is
 * touched. `fm_sheet_snapshot_range_style` captures what such an edit can
 * change, for undo.
 */

/**
 * @brief Lists every cell-XF index some cell of the selection resolves to,
 *        ascending: the XFs of cells that carry their own, then the row,
 *        column or default (0) style of the rest.
 *
 * `*out_len` always receives the number of indices; at most `cap` are
 * written. Call with `cap == 0` (`out_xfs` may then be `NULL`) to size the
 * buffer first.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_len` is `NULL`, or `out_xfs`
 *         is `NULL` while `cap > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the rect is out of range.
 */
FM_API fm_status_t fm_sheet_get_range_style_bases(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                                  uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                  uint32_t* out_xfs, size_t cap, size_t* out_len);

/**
 * @brief `fm_sheet_get_range_style_bases` over the selection less the cells
 *        `(excluded_rows[i], excluded_cols[i])`.
 *
 * An excluded cell contributes no XF: neither its own nor the row, column or
 * default style beneath it. Use it when the host resolves those cells itself
 * (say, a pending edit overrides their style). Excluded cells outside the
 * selection and repeats are ignored; excluding every cell yields no indices.
 * Values, styles and history are not touched.
 *
 * @return As `fm_sheet_get_range_style_bases`; also `FM_STATUS_BINDING_NULL_POINTER` (7001) when
 *         either excluded array is `NULL` while `excluded_count > 0`.
 */
FM_API fm_status_t fm_sheet_get_range_style_bases_excluding(const fm_workbook_t* wb, size_t sheet_index,
                                                            uint32_t first_row, uint32_t first_col, uint32_t last_row,
                                                            uint32_t last_col, const uint32_t* excluded_rows,
                                                            const uint32_t* excluded_cols, size_t excluded_count,
                                                            uint32_t* out_xfs, size_t cap, size_t* out_len);

/**
 * @brief Remaps the XFs of a selection: each cell ends at `to_xfs[i]` where
 *        it resolved to `from_xfs[i]` before; an unlisted XF is kept.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, or either array is `NULL`
 *         while `count > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the rect is out of range,
 *         a `from_xfs` entry repeats, or a `to_xfs` entry is past the
 *         cell-XF table;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when more than 1,048,576 cells would be
 *         written. Nothing changes on failure.
 */
FM_API fm_status_t fm_sheet_remap_range_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                              uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                              const uint32_t* from_xfs, const uint32_t* to_xfs, size_t count);

/**
 * @brief Captures the style state of a selection as an opaque, versioned
 *        byte string: the covered column styles (column selection), row
 *        styles (row selection), and every stored cell's XF and
 *        explicit-assignment mark.
 *
 * On success `*out_bytes` is a heap-allocated buffer of `*out_len` bytes
 * that MUST be released with `fm_buffer_destroy`. Both are zeroed before
 * validation and stay zero on failure.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the rect is out of range.
 */
FM_API fm_status_t fm_sheet_snapshot_range_style(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                                 uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                 uint8_t** out_bytes, size_t* out_len);

/**
 * @brief Puts a `fm_sheet_snapshot_range_style` capture back on sheet
 *        `sheet_index`.
 *
 * Captured axis styles and cell XFs are restored exactly; a stored cell of
 * the captured rect that was not captured loses its style. Values, formulas,
 * phonetic runs and comments are not touched.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, or `bytes` is `NULL`
 *         while `len > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the bytes
 *         are not a valid capture, or an XF is past the cell-XF table.
 *         Nothing changes on failure.
 */
FM_API fm_status_t fm_sheet_restore_range_style(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes,
                                                size_t len);

/**
 * @brief Removes the height override of `row`, returning it to the sheet
 *        default height.
 *
 * Clears the stored height and its custom flag. A row override left with
 * no other state (visible, outline level 0, no row style) is dropped,
 * except on a zeroHeight sheet, where a visible override is the row's
 * exception to the hidden default and is kept. A row with no height
 * override is a no-op.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row`
 *         is past the last row (1048575).
 */
FM_API fm_status_t fm_sheet_clear_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row);

/**
 * @brief Sheet default column / row metrics (`<sheetFormatPr>`).
 *
 * `default_col_width` is a stored column width (character units with the
 * column padding included, as OOXML writes it) and `default_row_height` is
 * in points; each is meaningful only when its `has_*` flag is non-zero,
 * since an absent attribute and an explicit `0` are different states.
 * `base_col_width` is in characters and always present (OOXML default 8).
 *
 * `effective_default_col_width` (stored width) and
 * `effective_default_row_height` (points) are the values the sheet actually
 * lays out with: the stated default when its `has_*` flag is set, else the
 * column Excel derives from `base_col_width` and the Normal font (Calibri
 * 11, base 8: 9.140625, shown as 8.43) and 102/7 pt; the row height is 0 when
 * the sheet's zeroHeight hides rows by default. They are filled by
 * `fm_sheet_get_format_defaults` and ignored by
 * `fm_sheet_set_format_defaults`.
 *
 * `custom_height` is `<sheetFormatPr customHeight>`: the stated
 * `default_row_height` is an explicit override that stays as written, rather
 * than a cached figure Excel recomputes from the Normal font. `zero_height`
 * is `<sheetFormatPr zeroHeight>`: rows are hidden unless they say otherwise.
 */
typedef struct {
  double default_col_width;
  double default_row_height;
  double base_col_width;
  int32_t has_default_col_width;  /* 0/1 */
  int32_t has_default_row_height; /* 0/1 */
  int32_t custom_height;          /* 0/1; `customHeight` */
  int32_t zero_height;            /* 0/1; `zeroHeight` */
  double effective_default_col_width;
  double effective_default_row_height;
} fm_sheet_format_defaults_t;

/**
 * @brief Reads the sheet's default column / row metrics.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_format_defaults(const fm_workbook_t* wb, size_t sheet_index,
                                                fm_sheet_format_defaults_t* out);

/**
 * @brief Replaces the sheet's default column / row metrics.
 *
 * A value whose `has_*` flag is zero is ignored and stored as `0`; the
 * `effective_*` fields are read-only and ignored. `custom_height` and
 * `zero_height` replace the stored flags, so read the current record first to
 * keep them. A rejected call leaves the sheet unchanged.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a
 *         stated value is negative, NaN, or infinite.
 */
FM_API fm_status_t fm_sheet_set_format_defaults(fm_workbook_t* wb, size_t sheet_index,
                                                const fm_sheet_format_defaults_t* defaults);

/**
 * @brief Which Windows column-width figure a geometry call uses.
 *
 * `FM_GEOMETRY_DISPLAY` is the screen figure (`Range.Width` at 96 DPI);
 * `FM_GEOMETRY_PRINT` is the DPI-stable figure pagination breaks on. Row
 * heights are the same in both modes. Only Windows is calibrated.
 */
typedef enum {
  FM_GEOMETRY_DISPLAY = 0,
  FM_GEOMETRY_PRINT = 1,
} fm_geometry_mode_t;

/**
 * @brief The character-unit to point conversion for the workbook's Normal
 *        font in one mode.
 *
 * `points_per_char` and `padding_pt` describe a column in displayed
 * characters (`Range.ColumnWidth`): `width_pt = chars * points_per_char +
 * padding_pt`. Stored widths go through `fm_sheet_convert_column_chars_to_pt`,
 * which also snaps to whole screen pixels.
 *
 * `calibrated` is `0` when the Normal font is not a tabulated calibration
 * point and the Calibri 11 figure was extrapolated. `normal_font_name` is a
 * read-scratch-backed view; `platform` is a static view (`"win"`).
 */
typedef struct {
  double points_per_char;
  double padding_pt;
  double normal_font_size;
  int32_t calibrated; /* 0/1 */
  int32_t _pad;
  const char* normal_font_name;
  const char* platform;
} fm_width_model_t;

/**
 * @brief Reads the rectangle covering the inclusive cell range, in points
 *        from the top-left of cell A1.
 *
 * Hidden rows and columns contribute zero extent. `mode` is an
 * `fm_geometry_mode_t` ordinal.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the range
 *         is inverted or outside the grid, or `mode` is unknown.
 */
FM_API fm_status_t fm_sheet_get_cell_rect_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                             uint32_t first_col, uint32_t last_row, uint32_t last_col, int32_t mode,
                                             fm_rect_pt_t* out);

/**
 * @brief Reads the effective width of column `col` in points; zero when the
 *        column is hidden.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `col` is out of range or
 *         `mode` is unknown.
 */
FM_API fm_status_t fm_sheet_get_column_width_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t col, int32_t mode,
                                                double* out_pt);

/**
 * @brief Reads the effective height of row `row` in points; zero when the
 *        row is hidden.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_height_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                              double* out_pt);

/**
 * @brief Reads the column-width conversion model in `mode`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `mode`
 *         is unknown.
 */
FM_API fm_status_t fm_sheet_get_width_model(const fm_workbook_t* wb, size_t sheet_index, int32_t mode,
                                            fm_width_model_t* out);

/**
 * @brief Converts a stored column width (character units with the column
 *        padding included, as `<col width>` holds it) to points. On screen
 *        the width snaps to whole pixels of the Normal font's maximum digit
 *        width (ECMA-376 §18.3.1.13). Zero is exactly zero points.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `mode` is
 *         unknown, or `chars` is negative, NaN, or infinite.
 */
FM_API fm_status_t fm_sheet_convert_column_chars_to_pt(const fm_workbook_t* wb, size_t sheet_index, int32_t mode,
                                                       double chars, double* out_pt);

/**
 * @brief Inverse of `fm_sheet_convert_column_chars_to_pt`: the stored width Excel
 *        writes for a column of `pt`, rounded to whole screen pixels; zero
 *        points maps to zero.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `mode` is
 *         unknown, or `pt` is negative, NaN, or infinite.
 */
FM_API fm_status_t fm_sheet_convert_column_pt_to_chars(const fm_workbook_t* wb, size_t sheet_index, int32_t mode,
                                                       double pt, double* out_chars);

/**
 * @brief Sets the sheet's zoom percentage. Values outside `[10, 400]`
 *        are clamped to the nearest endpoint.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, size_t sheet_index, uint32_t zoom_scale);

/**
 * @brief Sets the sheet's freeze pane in `(rows, cols)`. Either or
 *        both may be `0` to remove that axis.
 *
 * The first unfrozen cell `(freeze_rows, freeze_cols)` must lie inside the
 * grid, so `freeze_rows` is at most 1048575 and `freeze_cols` at most 16383.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `freeze_rows` / `freeze_cols` reaches past the grid (the view
 *           is left unchanged).
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
 * `fm_sheet_set_visibility`. Hiding the active sheet activates its visible
 * successor (`fm_workbook_get_active_sheet`). A workbook keeps at least one
 * visible sheet, so the only visible sheet cannot be hidden.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `hidden` is non-zero and the sheet is the only visible one
 *           (the workbook is left unchanged).
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
 * a value outside `fm_sheet_visibility_t` returns `FM_STATUS_INVALID_ARGUMENT` (2) and
 * leaves the sheet unchanged. Hiding the active sheet activates its visible
 * successor (`fm_workbook_get_active_sheet`). A workbook keeps at least one
 * visible sheet, so the only visible sheet cannot be hidden.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `visibility` is not an `fm_sheet_visibility_t` value, or it
 *           hides the only visible sheet (the workbook is left unchanged).
 */
FM_API fm_status_t fm_sheet_set_visibility(fm_workbook_t* wb, size_t sheet_index, int32_t visibility);

/**
 * @brief Sets the sheet's gridline-visibility flag (`showGridLines`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_show_grid_lines(fm_workbook_t* wb, size_t sheet_index, int32_t show);

/**
 * @brief Sets the sheet's row/column header-visibility flag
 *        (`showRowColHeaders`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_show_row_col_headers(fm_workbook_t* wb, size_t sheet_index, int32_t show);

/**
 * @brief Sets the sheet's zero-value display flag (`showZeros`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_show_zeros(fm_workbook_t* wb, size_t sheet_index, int32_t show);

/**
 * @brief Sets the sheet's formula-text display flag (`showFormulas`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_show_formulas(fm_workbook_t* wb, size_t sheet_index, int32_t show);

/**
 * @brief Sets the sheet's outline-symbol display flag (`showOutlineSymbols`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_show_outline_symbols(fm_workbook_t* wb, size_t sheet_index, int32_t show);

/**
 * @brief Sets the sheet's right-to-left display flag (`rightToLeft`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_set_right_to_left(fm_workbook_t* wb, size_t sheet_index, int32_t right_to_left);

/**
 * @brief Sets the sheet's tab-selected flag (`tabSelected`). Plain
 *        metadata mirroring the OOXML attribute; setting this on more
 *        than one sheet does not affect evaluation, only what a host
 *        UI reproduces as the active tab.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL` or `mode == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
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
FM_API const char* fm_get_last_error_message(void);

/**
 * @brief Returns the optional context string associated with the
 *        most recent error on the current thread.
 *
 * Same lifetime / non-NULL guarantees as `fm_get_last_error_message`.
 */
FM_API const char* fm_get_last_error_context(void);

/**
 * @brief Returns a static C string describing `status`.
 *
 * For known codes the result is the status name without the `FM_STATUS_`
 * prefix (e.g. `"INVALID_ARGUMENT"` for `FM_STATUS_INVALID_ARGUMENT`);
 * unknown codes yield `"UNRECOGNIZED_STATUS"`.
 * Always non-NULL with program lifetime.
 */
FM_API const char* fm_get_status_name(fm_status_t status);

/**
 * @brief Returns the Excel display literal for a cell error code.
 *
 * Known values return literals such as `"#DIV/0!"` and `"#REF!"`.
 * Unknown numeric values return `"#UNKNOWN!"`. The returned pointer is
 * never `NULL` and has program lifetime.
 */
FM_API const char* fm_get_error_display_name(fm_error_code_t error);

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `level` is outside `fm_log_level_t`.
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
 * @return `FM_STATUS_OK` (0) always.
 */
FM_API fm_status_t fm_set_log_sink(fm_log_sink_cb sink, void* user_data);

/* -------------------------------------------------------------------------- */
/* Styles                                                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Plain-data projection of a `formulon::io::CellXf` record.
 *
 * This record is a fixed 128-byte C ABI layout carrying every optional
 * alignment attribute represented by Formulon's style model, the six
 * `apply*` flags, `quotePrefix`, and the `<protection>` child.
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
 * `apply_*` mirror the `<xf apply*>` attributes, which say whether the
 * record's own value for that group overrides its parent named style.
 * `has_protection` is the presence of the `<protection>` child; when it is
 * zero `locked` / `hidden` are ignored on insertion and getters report the
 * model defaults (`locked=1`, `hidden=0`). All of these take part in XF
 * deduplication, so two records differing only in protection stay distinct.
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
  int32_t apply_number_format;   /* 0/1 */
  int32_t apply_font;            /* 0/1 */
  int32_t apply_fill;            /* 0/1 */
  int32_t apply_border;          /* 0/1 */
  int32_t apply_alignment;       /* 0/1 */
  int32_t apply_protection;      /* 0/1 */
  int32_t quote_prefix;          /* 0/1 */
  int32_t has_protection;        /* 0= no `<protection>` child, 1= child present */
  int32_t locked;                /* 0/1; ignored unless has_protection */
  int32_t hidden;                /* 0/1; ignored unless has_protection */
} fm_cell_xf_t;

/** Discriminator for `fm_color_spec_t::kind`, mirroring
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
} fm_color_spec_t;

/**
 * @brief Plain-data projection of a `formulon::io::FontRecord`.
 *
 * `name` is a NUL-terminated UTF-8 model-backed view into the workbook's
 * styles table; it is valid until the next mutation of the font list
 * (`fm_workbook_add_font`, `fm_workbook_set_font`) or until the handle is
 * destroyed. Reads do not invalidate it.
 *
 * The `has_*` flags distinguish an absent OOXML element from an explicit
 * `val="0"`. They matter for a differential font, where an absent `<b>`
 * means "leave the source formatting unchanged" while `<b val="0"/>`
 * means "switch bold off"; on a cell font a `has_*` flag without its
 * value emits the explicit-off form.
 *
 * `outline`, `shadow`, `condense` and `extend` are the matching toggle
 * elements and follow the same `has_*` rule as `bold`.
 *
 * `has_size` is honoured by `fm_workbook_add_dxf` alone: a differential font
 * with `has_size == 0` leaves the size unchanged, while every cell-font
 * entry point writes `size` regardless.
 *
 * `scheme` is the `<scheme>` link to the workbook theme's major (heading)
 * or minor (body) typeface. It needs no presence flag because OOXML has no
 * "absent value" spelling for it: `0` is exactly "no `<scheme>` element".
 * Set it only on a workbook that carries a theme part — a link into a
 * theme that is not there resolves to nothing.
 */
typedef struct {
  const char* name; /* NUL-terminated UTF-8, model-backed view */
  double size;
  /* Literal RGB for color.kind == kFmColorRgb; compatibility fallback for
   * theme/indexed/auto selectors (not an Excel-rendered colour). */
  uint32_t color_argb;
  int32_t bold;         /* 0=false, 1=true */
  int32_t italic;       /* 0=false, 1=true */
  int32_t strike;       /* 0=false, 1=true */
  int32_t has_bold;     /* 0= no `<b>` element, 1= element present */
  int32_t has_italic;   /* 0= no `<i>` element, 1= element present */
  int32_t has_strike;   /* 0= no `<strike>` element, 1= element present */
  int32_t has_family;   /* 0= no `<family>` element, 1= element present */
  int32_t has_charset;  /* 0= no `<charset>` element, 1= element present */
  int32_t has_size;     /* 0= no `<sz>` element, 1= element present */
  int32_t outline;      /* 0=false, 1=true */
  int32_t shadow;       /* 0=false, 1=true */
  int32_t condense;     /* 0=false, 1=true */
  int32_t extend;       /* 0=false, 1=true */
  int32_t has_outline;  /* 0= no `<outline>` element, 1= element present */
  int32_t has_shadow;   /* 0= no `<shadow>` element, 1= element present */
  int32_t has_condense; /* 0= no `<condense>` element, 1= element present */
  int32_t has_extend;   /* 0= no `<extend>` element, 1= element present */
  uint8_t underline;    /* 0=none, 1=single, 2=double, 3/4=accounting variants */
  uint8_t vert_align;   /* 0=baseline, 1=superscript, 2=subscript */
  uint8_t family;       /* `<family>` font-family class (0..5) */
  uint8_t charset;      /* `<charset>` codepage id (e.g. 128 = Shift_JIS) */
  uint8_t scheme;       /* `<scheme>` theme link: 0=absent, 1=major, 2=minor */
  fm_color_spec_t color;
} fm_font_record_t;

/**
 * @brief One `<stop>` of a gradient fill: a colour at `position` (0..1).
 *
 * `color_argb` is literal RGB when `color.kind` is kFmColorRgb and a
 * compatibility fallback otherwise.
 */
typedef struct {
  double position;
  uint32_t color_argb;
  fm_color_spec_t color;
} fm_gradient_stop_t;

/** `fm_fill_record_t::gradient_type` values. */
#define FM_GRADIENT_LINEAR 0
#define FM_GRADIENT_PATH 1

/**
 * @brief Plain-data projection of a `formulon::io::FillRecord`.
 *
 * `pattern` is the OOXML pattern index: `0=none`, `1=solid`,
 * `2..18=standard pattern set`. `fg_argb` and `bg_argb` are AARRGGBB
 * packed literal RGB or compatibility fallback values; `fg` and `bg` carry
 * the authoritative `<fgColor>` / `<bgColor>` specifications.
 *
 * With `has_gradient` set the record is a `<gradientFill>` and the pattern
 * fields are ignored: `gradient_type` is FM_GRADIENT_LINEAR or
 * FM_GRADIENT_PATH, `gradient_degree` is a linear gradient's angle,
 * `gradient_left/right/top/bottom` (0..1) locate a path gradient's focus.
 * `gradient_stops` points at `gradient_stop_count` stops; on a read it is
 * valid until the next `fm_workbook_get_fill` / `fm_workbook_get_dxf` call, a
 * mutation or handle destruction, on a write it is borrowed for the call. A non-finite number
 * is stored as 0 and at most 1024 stops are kept.
 */
typedef struct {
  uint8_t pattern; /* 0=none, 1=solid, 2..18=standard pattern set */
  /* Literal RGB when the matching spec kind is kFmColorRgb; compatibility
   * fallbacks otherwise (not Excel-rendered theme/indexed/auto colours). */
  uint32_t fg_argb;
  uint32_t bg_argb;
  fm_color_spec_t fg;
  fm_color_spec_t bg;
  int32_t has_gradient;  /* 0=pattern fill, 1=`<gradientFill>` */
  int32_t gradient_type; /* FM_GRADIENT_LINEAR / FM_GRADIENT_PATH */
  double gradient_degree;
  double gradient_left;
  double gradient_right;
  double gradient_top;
  double gradient_bottom;
  const fm_gradient_stop_t* gradient_stops;
  size_t gradient_stop_count;
} fm_fill_record_t;

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
  fm_color_spec_t color;
} fm_border_side_t;

/**
 * @brief Plain-data projection of a `formulon::io::BorderRecord`.
 *
 * Each side mirrors the OOXML schema. `diagonal_up` / `diagonal_down`
 * carry the boolean flags as `0`/`1` int32_t to match the rest of the
 * binding surface.
 */
typedef struct {
  fm_border_side_t left;
  fm_border_side_t right;
  fm_border_side_t top;
  fm_border_side_t bottom;
  fm_border_side_t diagonal;
  int32_t diagonal_up;   /* 0=false, 1=true */
  int32_t diagonal_down; /* 0=false, 1=true */
} fm_border_record_t;

/**
 * @brief One-shot style-table insertion request.
 *
 * Each non-empty input array requires its matching output-index array. The
 * call deduplicates each table in linear time across the existing records
 * and the supplied batch; fonts, fills and borders are installed before xfs
 * so an xf may refer to indices returned by the earlier arrays.
 *
 * `cell_xfs` entries carry the same presence-flag contract as
 * `fm_workbook_add_cell_xf`: an alignment attribute is emitted only when its
 * `has_*` flag is set, and the value is ignored otherwise.
 */
typedef struct {
  const fm_font_record_t* fonts;
  size_t font_count;
  uint32_t* font_indices;
  const fm_fill_record_t* fills;
  size_t fill_count;
  uint32_t* fill_indices;
  const fm_border_record_t* borders;
  size_t border_count;
  uint32_t* border_indices;
  const fm_cell_xf_t* cell_xfs;
  size_t cell_xf_count;
  uint32_t* cell_xf_indices;
  const char* const* num_fmt_codes;
  size_t num_fmt_count;
  uint16_t* num_fmt_ids;
} fm_styles_batch_t;

/**
 * @brief Plain-data projection of one OOXML `<dxf>` differential format.
 *
 * Each `*_engaged` flag mirrors whether the corresponding child element
 * exists in the source `<dxf>`. `num_fmt_code`, `alignment_xml`, and
 * `protection_xml` are model-backed views, valid until the next styles-table
 * mutation (including `fm_workbook_add_dxf`) or until the handle is destroyed;
 * reads do not invalidate them. The XML fields are never `NULL` and are empty
 * when their child is absent. Their semantic content is preserved without a
 * typed projection, but XML parsing may normalize lexical formatting such as
 * whitespace, quotes, and empty-element spelling.
 */
typedef struct {
  int32_t font_engaged; /* 0=false, 1=true */
  fm_font_record_t font;
  int32_t fill_engaged; /* 0=false, 1=true */
  fm_fill_record_t fill;
  int32_t border_engaged; /* 0=false, 1=true */
  fm_border_record_t border;
  int32_t num_fmt_engaged; /* 0=false, 1=true */
  uint16_t num_fmt_id;
  const char* num_fmt_code;   /* UTF-8, NUL-terminated; never NULL */
  const char* alignment_xml;  /* UTF-8 serialized `<alignment/>`; never NULL */
  const char* protection_xml; /* UTF-8 serialized `<protection/>`; never NULL */
} fm_dxf_record_t;

/** Sentinel for `fm_cell_style_record_t::builtin_id` indicating the
 *  style is custom (no `builtinId` attribute on the OOXML element). */
#define FM_CELL_STYLE_BUILTIN_ID_NONE 0xFFFFFFFFu

/**
 * @brief Plain-data projection of a `formulon::io::CellStyleRecord`
 *        (one OOXML `<cellStyle>` entry).
 *
 * `name` is a NUL-terminated UTF-8 model-backed view into the workbook's
 * styles table; it is valid until the next mutation of the cell-style
 * list (`fm_workbook_set_cell_style`, `fm_workbook_remove_cell_style`) or
 * until the handle is destroyed. `xf_id` indexes into the parallel
 * `<cellStyleXfs>` table (queryable via
 * `fm_workbook_get_cell_style_xf_count` / `fm_workbook_get_cell_style_xf`).
 *
 * `builtin_id` carries the OOXML built-in style ordinal (`0..53`); the
 * sentinel `FM_CELL_STYLE_BUILTIN_ID_NONE` indicates the entry is
 * custom and the attribute was absent on the source document. `i_level`
 * is the outline level of the `RowLevel_n` / `ColLevel_n` styles
 * (`builtin_id` 1 / 2, `0..6`) and `0` for everything else. The boolean
 * flags follow the wide-POD convention used elsewhere on this surface.
 */
typedef struct {
  const char* name;       /* UTF-8, NUL-terminated; never NULL */
  uint32_t xf_id;         /* index into cell_style_xfs */
  uint32_t builtin_id;    /* 0..53, or FM_CELL_STYLE_BUILTIN_ID_NONE */
  uint32_t i_level;       /* outline level for heading styles */
  int32_t hidden;         /* 0=false, 1=true */
  int32_t custom_builtin; /* 0=false, 1=true */
} fm_cell_style_record_t;

/**
 * @brief Reads the `xf_index` (style record id) attached to the cell at
 *        `(row, col)` on `sheet_index`.
 *
 * Returns `0` (the default xf) when the cell is absent — callers that
 * need to distinguish "no cell" from "default-formatted cell" should
 * combine this with `fm_sheet_get_cell_at` iteration.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cell_xf_index(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              size_t* out_xf_index);

/**
 * @brief Stores `xf_index` on the cell at `(row, col)` on `sheet_index`.
 *
 * Materialises the cell as a default-blank slot when none exists yet,
 * mirroring the row-vector growth semantics of
 * `fm_sheet_set_blank`. Coexists with literal / formula writes:
 * those calls leave `xf_index` untouched, so the caller can layer a
 * style update on top of either order of operations.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `xf_index` is not below `fm_workbook_get_cell_xf_count` (no cell
 *           is materialised).
 */
FM_API fm_status_t fm_sheet_set_cell_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              size_t xf_index);

/**
 * @brief Removes the direct style assignment from the cell at `(row, col)`
 *        on `sheet_index`, so a blank cell inherits its row / column style again.
 *
 * Resets the cell's `xf_index` to `0` and drops the explicit-assignment
 * mark; the value, formula, phonetic runs and comments are kept. An
 * absent cell is left absent (no-op). A cell holding a value or formula
 * still owns its style, which becomes the default xf; only a blank cell
 * falls back to row / column / default inheritance. Tell an unassigned
 * cell from one assigned the default xf through the `source` of
 * `fm_sheet_get_effective_style`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the address is out of range.
 */
FM_API fm_status_t fm_sheet_clear_cell_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col);

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a corner is out of range, or
 *           `xf_index` is not below `fm_workbook_get_cell_xf_count` (checked
 *           before any cell is written);
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the rectangle exceeds the per-call
 *         cell ceiling.
 */
FM_API fm_status_t fm_sheet_set_range_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                               uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                               size_t xf_index);

/**
 * @brief Reads the `xf_index`-th `<xf>` record from the workbook's
 *        styles table, including optional alignment attributes and their
 *        presence flags.
 *
 * `has_alignment` reports whether the `<alignment>` child existed;
 * explicit zero / false attributes remain present in their value fields.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `xf_index >= cell_xfs.size()`.
 */
FM_API fm_status_t fm_workbook_get_cell_xf(const fm_workbook_t* wb, size_t xf_index, fm_cell_xf_t* out);

/**
 * @brief Reads the `font_index`-th font record from the workbook's
 *        styles table.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `font_index >= fonts.size()`.
 */
FM_API fm_status_t fm_workbook_get_font(const fm_workbook_t* wb, size_t font_index, fm_font_record_t* out);

/**
 * @brief Looks up the effective format string for `num_fmt_id`.
 *
 * Resolution searches valid custom `<numFmt>` records in document order;
 * this includes a custom record below 164 that overrides a built-in slot.
 * When no valid custom record matches, a non-empty built-in code in the
 * `0..163` range is used as the fallback. A built-in id with no ECMA-376 code
 * (the East Asian date ids, the Thai-digit ids, the currency ids 5-8)
 * resolves to the code the workbook's Excel profile renders it with, so the
 * result for those ids follows `fm_workbook_set_excel_profile_id`. On success `*out` is either a
 * model-backed view into the custom record's format string (valid until the
 * next mutation of the numFmt list, `fm_workbook_add_num_fmt`, or until
 * handle destruction), or a static built-in view with program lifetime.
 * The styles writer intentionally emits
 * only custom records with ids >= 164, so a below-164 override is not
 * promised to survive save/load. Returns `FM_STATUS_INVALID_ARGUMENT` (2) when no effective
 * mapping exists for `num_fmt_id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the id is unknown.
 */
FM_API fm_status_t fm_workbook_get_num_fmt_string(const fm_workbook_t* wb, uint32_t num_fmt_id, const char** out);

/**
 * @brief Reads the `fill_index`-th fill record from the workbook's
 *        styles table.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `fill_index >= fills.size()`.
 */
FM_API fm_status_t fm_workbook_get_fill(const fm_workbook_t* wb, size_t fill_index, fm_fill_record_t* out);

/**
 * @brief Reads the `border_index`-th border record from the workbook's
 *        styles table.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `border_index >= borders.size()`.
 */
FM_API fm_status_t fm_workbook_get_border(const fm_workbook_t* wb, size_t border_index, fm_border_record_t* out);

/**
 * @brief Returns the number of `<dxf>` differential-format records
 *        available for conditional-format `dxfId` resolution.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_dxf_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `dxf_index`-th differential-format record.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `dxf_index >= dxfs.size()`.
 */
FM_API fm_status_t fm_workbook_get_dxf(const fm_workbook_t* wb, size_t dxf_index, fm_dxf_record_t* out);

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_dxf_index` is `NULL`.
 */
FM_API fm_status_t fm_workbook_add_dxf(fm_workbook_t* wb, fm_dxf_record_t record, size_t* out_dxf_index);

/**
 * @brief One region of a custom table style (`fm_table_style_t`).
 *
 *   * `type`   — the `ST_TableStyleType` token: `wholeTable`, `headerRow`,
 *                `totalRow`, `firstColumn`, `lastColumn`, `firstRowStripe`,
 *                `secondRowStripe`, `firstColumnStripe`, `secondColumnStripe`,
 *                `firstHeaderCell`, `lastHeaderCell`, `firstTotalCell`,
 *                `lastTotalCell`, `firstSubtotalColumn`,
 *                `secondSubtotalColumn`, `thirdSubtotalColumn`,
 *                `firstSubtotalRow`, `secondSubtotalRow`, `thirdSubtotalRow`,
 *                `blankRow`, `firstColumnSubheading`,
 *                `secondColumnSubheading`, `thirdColumnSubheading`,
 *                `firstRowSubheading`, `secondRowSubheading`,
 *                `thirdRowSubheading`, `pageFieldLabels`, `pageFieldValues`.
 *   * `size`   — stripe band size, at least 1 (1 for non-stripe regions).
 *   * `dxf_id` — read only when `has_dxf_id`; indexes `<dxfs>`
 *                (`fm_workbook_add_dxf`).
 */
typedef struct {
  const char* type;
  uint32_t size;
  int32_t has_dxf_id;
  uint32_t dxf_id;
} fm_table_style_element_t;

/**
 * @brief A custom table / PivotTable style from `<tableStyles>`.
 *
 * `pivot` / `table` are non-zero when PivotTables / tables may use the style.
 * On read, `name` and the element `type`s are model-backed views valid until
 * the next styles mutation, and `elements` is a read-scratch-backed array.
 */
typedef struct {
  const char* name;
  int32_t pivot;
  int32_t table;
  const fm_table_style_element_t* elements;
  uint32_t element_count;
} fm_table_style_t;

/**
 * @brief Returns the number of custom table / PivotTable styles.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_table_style_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `index`-th custom table style.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_table_style_at(const fm_workbook_t* wb, size_t index, fm_table_style_t* out);

/**
 * @brief Adds a custom table style, or replaces the one with the same name
 *        (case-insensitive) in place.
 *
 * The name must be non-empty and must not be a built-in style name
 * (`TableStyleLight1`, `PivotStyleMedium2`, ...). Each element type may
 * appear once. A table's style name or a pivot's style name may then name it.
 * The first style added to a workbook without `<tableStyles>` sets the
 * defaults to `TableStyleMedium2` / `PivotStyleLight16`, as the desktop
 * application writes them.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `style` or `style->name` is `NULL`,
 *         or `elements` is `NULL` with a non-zero count;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an invalid name, an unknown or repeated
 *         element type, a zero `size`, or a `dxf_id` past `<dxfs>`.
 */
FM_API fm_status_t fm_workbook_set_table_style(fm_workbook_t* wb, const fm_table_style_t* style);

/**
 * @brief Removes the custom table style `name` (case-insensitive).
 *
 * A default naming it reverts to the built-in default. Tables and pivots
 * naming it keep the name, which the desktop application then renders
 * unstyled.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when no custom style has that name.
 */
FM_API fm_status_t fm_workbook_remove_table_style(fm_workbook_t* wb, const char* name);

/**
 * @brief Reads the default table and PivotTable style names (model-backed
 *        views, empty when unset).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_table_style_defaults(const fm_workbook_t* wb, const char** out_table_style,
                                                        const char** out_pivot_style);

/**
 * @brief Sets the default table and PivotTable style names. `NULL` keeps a
 *        default, an empty string removes it.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`.
 */
FM_API fm_status_t fm_workbook_set_table_style_defaults(fm_workbook_t* wb, const char* table_style,
                                                        const char* pivot_style);

/**
 * @brief Returns the number of font records currently registered in the
 *        workbook's styles table.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_font_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Returns the number of fill records currently registered in the
 *        workbook's styles table.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_fill_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Returns the number of border records currently registered in
 *        the workbook's styles table.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_border_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Returns the number of `<xf>` records currently registered in
 *        the workbook's styles table (i.e. the size of the `cellXfs`
 *        vector).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_cell_xf_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Returns the number of named cell styles (`<cellStyle>` entries)
 *        registered in the workbook. Zero for workbooks that do not
 *        declare any named styles.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_cell_style_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `index`-th named cell style. The returned `name` is a
 *        model-backed view into storage owned by the workbook; see
 *        `fm_cell_style_record_t` for lifetime rules.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index >= cell_styles.size()`.
 */
FM_API fm_status_t fm_workbook_get_cell_style_at(const fm_workbook_t* wb, size_t index, fm_cell_style_record_t* out);

/**
 * @brief Returns the number of `<cellStyleXfs>` records — the named-
 *        style xf table. This table is independent of `cellXfs` and
 *        is referenced by `fm_cell_style_record_t::xf_id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_cell_style_xf_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `index`-th `<cellStyleXfs>` record (named-style xf
 *        table). Output shape mirrors `fm_workbook_get_cell_xf`, including
 *        the optional alignment attributes and their presence flags.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index >= cell_style_xfs.size()`.
 */
FM_API fm_status_t fm_workbook_get_cell_style_xf(const fm_workbook_t* wb, size_t index, fm_cell_xf_t* out);

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
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_index` is `NULL`.
 */
FM_API fm_status_t fm_workbook_add_font(fm_workbook_t* wb, fm_font_record_t record, size_t* out_index);

/**
 * @brief Overwrites the `font_index`-th font record in place.
 *
 * The counterpart to `fm_workbook_add_font` for a slot that already exists.
 * Every `<xf>` whose `font_index` names this slot renders with the new
 * record, so this is a bulk restyle rather than a local edit: replacing a
 * font that several cell formats share changes all of them at once. Use
 * `fm_workbook_add_font` when the intent is to introduce a new appearance and
 * repoint selected `<xf>` records at it.
 *
 * No dedup and no auto-grow: the index must already be in range, and a
 * record equal to another slot's is allowed to sit at two indices (the
 * styles writer emits both, which Excel accepts).
 *
 * `record.name` follows `fm_workbook_add_font`: NUL-terminated UTF-8, copied
 * into the styles table, `NULL` treated as the empty string.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `font_index >= fonts.size()`.
 */
FM_API fm_status_t fm_workbook_set_font(fm_workbook_t* wb, size_t font_index, fm_font_record_t record);

/**
 * @brief Declares the workbook's default font.
 *
 * Font 0 is the record every unformatted cell resolves to: `fm_workbook_create_default`
 * seeds it with Excel's `Calibri` 11, and cell xf 0 — the format of a cell
 * that was never styled — names it. Because the seeded table already owns
 * index 0, `fm_workbook_add_font` can only ever append beside it, so this is
 * the only way to change what an unstyled cell is saved as. A ja-JP host
 * declaring `游ゴシック` calls this once after creating the workbook.
 *
 * Equivalent to `fm_workbook_set_font(wb, 0, record)`, except that it also
 * succeeds on a workbook whose styles table is empty (a table replaced
 * wholesale by `fm_workbook_load` on a package with no `xl/styles.xml`):
 * the reserved roots are seeded first, then index 0 is overwritten.
 *
 * Reading the current default is `fm_workbook_get_font(wb, 0, out)`.
 *
 * The record is a font, not a theme. A workbook without a theme part is
 * treated by Excel as having the Office theme, and Excel's own new-workbook
 * font 0 carries `scheme = 2` (minor), which `fm_workbook_create_for_profile`
 * reproduces. Pass `scheme = 0` when the name must resolve literally rather
 * than through the theme's body font. In a workbook loaded from a package
 * that brought its own theme, reading font 0, changing the name and writing
 * it back with `scheme` preserved keeps Excel showing the font as the
 * theme's body font.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`.
 */
FM_API fm_status_t fm_workbook_set_default_font(fm_workbook_t* wb, fm_font_record_t record);

/**
 * @brief Adds a fill record to the workbook's styles table, deduplicating
 *        against existing entries.
 *
 * Linear-search dedup on the same writer-visible-state basis as
 * `fm_workbook_add_font`: a record that differs only in its `fg` / `bg`
 * colour specification gets its own index. When no match exists the
 * record is appended and the new (now-largest) index is returned. The
 * dedup is `O(N)` per call; callers that bulk-build a workbook should
 * batch when possible.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_index` is `NULL`.
 */
FM_API fm_status_t fm_workbook_add_fill(fm_workbook_t* wb, fm_fill_record_t record, size_t* out_index);

/**
 * @brief Adds a border record to the workbook's styles table, deduplicating
 *        against existing entries.
 *
 * Linear-search dedup on the same writer-visible-state basis as
 * `fm_workbook_add_font`: a record whose sides differ only in their colour
 * specification gets its own index. When no match exists the record is
 * appended and the new (now-largest) index is returned. The dedup is
 * `O(N)` per call; callers that bulk-build a workbook should batch when
 * possible.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_index` is `NULL`.
 */
FM_API fm_status_t fm_workbook_add_border(fm_workbook_t* wb, fm_border_record_t record, size_t* out_index);

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
 *      `FM_STATUS_PRECONDITION_FAILED` (9) and leaves both the table and `*out_num_fmt_id`
 *      unchanged.
 *
 * `format_code` must be a NUL-terminated UTF-8 pointer; passing `NULL`
 * is treated as the empty string.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_num_fmt_id` is `NULL`;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when no custom OOXML id remains.
 */
FM_API fm_status_t fm_workbook_add_num_fmt(fm_workbook_t* wb, const char* format_code, uint32_t* out_num_fmt_id);

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
 *   * `record.num_fmt_id` must either be a built-in id with a format
 *     code in at least one locale (`5`-`8`, `41`-`44` included) or already
 *     registered as a custom entry. Unknown ids surface
 *     `FM_STATUS_INVALID_ARGUMENT` (2).
 *   * Alignment values outside their Excel ranges are rejected when the
 *     matching presence flag is set.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_xf_index` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for any out-of-range / unregistered field.
 */
FM_API fm_status_t fm_workbook_add_cell_xf(fm_workbook_t* wb, fm_cell_xf_t record, size_t* out_xf_index);

/** Adds and deduplicates a named-style xf including optional alignment
 * attributes. `has_alignment` preserves an explicit empty child. Invalid
 * Excel ranges return `FM_STATUS_INVALID_ARGUMENT` (2). */
FM_API fm_status_t fm_workbook_add_cell_style_xf(fm_workbook_t* wb, fm_cell_xf_t record, uint32_t* out_xf_id);

/**
 * @brief Adds or replaces the named `<cellStyle>` record `record->name`.
 *
 * An existing style of that name is overwritten field for field; otherwise
 * the record is appended. `record->name` is copied. `xf_id` must reference
 * an existing named-style xf; `builtin_id` is `0..53` or
 * `FM_CELL_STYLE_BUILTIN_ID_NONE` for a custom style; `i_level` is `0..6`
 * for `builtin_id` 1 / 2 and must be `0` otherwise.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `record` or `record->name` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an empty name or an out-of-range field.
 */
FM_API fm_status_t fm_workbook_set_cell_style(fm_workbook_t* wb, const fm_cell_style_record_t* record);

/**
 * @brief Removes the named `<cellStyle>` record `name`.
 *
 * When no other named style shares the removed record's `xf_id`, every
 * cell xf that inherits from it is repointed at Normal (`xf_id` 0), and each
 * attribute group whose `apply_*` flag is zero takes Normal's value; groups
 * whose flag is set keep their own. Neither `<cellXfs>` nor
 * `<cellStyleXfs>` is compacted, so every xf index stays valid.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when no style has that name or it is the
 *         Normal style (`builtin_id` 0).
 */
FM_API fm_status_t fm_workbook_remove_cell_style(fm_workbook_t* wb, const char* name);

/**
 * @brief Adds and deduplicates a batch of fonts, fills, borders, and cell xfs.
 *
 * Empty arrays may be `NULL`. A non-empty input or output array that is NULL
 * returns `FM_STATUS_BINDING_NULL_POINTER` (7001); invalid xfs return `FM_STATUS_INVALID_ARGUMENT` (2).
 * The operation is staged and committed as one transaction: any returned
 * error, including `FM_STATUS_PRECONDITION_FAILED` (9) when custom number-format ids are
 * exhausted, leaves the workbook and every caller-provided output array
 * unchanged. Allocation failures are not converted into a fabricated
 * success or recovery path; callers observe only the documented status
 * returns from validation / precondition checks.
 */
FM_API fm_status_t fm_workbook_add_batch(fm_workbook_t* wb, const fm_styles_batch_t* batch);

/* -------------------------------------------------------------------------- */
/* Theme, colour resolution and effective style                               */
/* -------------------------------------------------------------------------- */

/**
 * @brief The 12 theme colours as 0xFFRRGGBB, in `a:clrScheme` order:
 *        dk1, lt1, dk2, lt2, accent1..accent6, hlink, folHlink.
 *
 * This is not the order of a spreadsheet colour's `theme` index
 * (`fm_color_spec_t::theme`), which swaps the first two pairs
 * (0=lt1, 1=dk1, 2=lt2, 3=dk2).
 */
typedef struct {
  uint32_t argb[12];
} fm_theme_colors_t;

/**
 * @brief The theme's major (heading) and minor (body) typefaces.
 *
 * `*_east_asian` is the `a:ea` face when it names one, otherwise the
 * `a:font script="Jpan"` face. Getter strings are read-scratch-backed views.
 */
typedef struct {
  const char* major_latin;
  const char* major_east_asian;
  const char* minor_latin;
  const char* minor_east_asian;
} fm_theme_fonts_t;

/** @brief Where a colour is used; decides what an automatic colour means. */
typedef enum {
  FM_COLOR_CONTEXT_FONT = 0,
  FM_COLOR_CONTEXT_FILL_FG = 1,
  FM_COLOR_CONTEXT_FILL_BG = 2,
  FM_COLOR_CONTEXT_BORDER = 3,
} fm_color_context_t;

/**
 * @brief How a colour was resolved.
 *
 * `FM_COLOR_RESOLUTION_EXACT`: literal RGB, or a theme / palette colour
 * resolved from the workbook's own data. `DEFAULT_THEME` /
 * `THEME_UNPARSEABLE`: a theme colour resolved against the
 * current Excel Office (Aptos) theme because the theme part is absent / unparseable.
 * `INDEX_OUT_OF_RANGE`: a theme or palette index beyond its table; the
 * colour is black. `AUTO_CONTEXT`: an automatic or system colour (or no
 * colour at all) chosen by context: black for font and border, white for
 * fills.
 */
typedef enum {
  FM_COLOR_RESOLUTION_EXACT = 0,
  FM_COLOR_RESOLUTION_DEFAULT_THEME = 1,
  FM_COLOR_RESOLUTION_INDEX_OUT_OF_RANGE = 2,
  FM_COLOR_RESOLUTION_THEME_UNPARSEABLE = 3,
  FM_COLOR_RESOLUTION_AUTO_CONTEXT = 4,
} fm_color_resolution_t;

/**
 * @brief The formatting a cell shows, read from the selected `<cellXfs>`
 *        record directly.
 *
 * `source` names the level that supplied `xf_index`: 0 the cell, 1 its row
 * (a row style takes effect only with `customFormat`), 2 its column, 3 the
 * default xf 0. An xf, font, fill or border index beyond its table reads as
 * 0. Colours are resolved against the theme and `<indexedColors>`; each
 * `*_resolution` is an `fm_color_resolution_t`. `border_*[5]` are left,
 * right, top, bottom, diagonal. Named-style inheritance is not recomputed
 * and conditional formatting is not applied. `num_fmt_code` is a
 * read-scratch-backed view.
 */
typedef struct {
  uint32_t xf_index;
  int32_t source; /* 0=cell, 1=row, 2=column, 3=default */
  uint32_t font_index;
  uint32_t fill_index;
  uint32_t border_index;
  uint32_t font_argb;
  int32_t font_resolution;
  uint32_t fill_fg_argb;
  int32_t fill_fg_resolution;
  uint32_t fill_bg_argb;
  int32_t fill_bg_resolution;
  uint32_t border_argb[5];
  int32_t border_resolution[5];
  int32_t locked; /* 0/1 */
  int32_t hidden; /* 0/1 */
  const char* num_fmt_code;
} fm_effective_style_t;

/**
 * @brief Reads the theme's 12 scheme colours.
 *
 * `*out_source` is 0 when they come from the workbook's theme part, 1 when
 * the workbook has no theme part and 2 when its theme part cannot be
 * parsed; in both latter cases the default theme's colours are reported.
 * The theme is parsed on every call.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_theme_colors(const fm_workbook_t* wb, fm_theme_colors_t* out, int32_t* out_source);

/**
 * @brief Writes the theme's 12 scheme colours. The alpha byte is ignored.
 *
 * Only `a:clrScheme` of the existing theme part is rewritten. A workbook
 * without a theme part first gets a minimal valid one carrying the default
 * theme, registered with its content type and workbook relationship.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_IO_XML_PARSE` (5007) when the existing theme part cannot be parsed (the
 *         part is left untouched).
 */
FM_API fm_status_t fm_workbook_set_theme_colors(fm_workbook_t* wb, const fm_theme_colors_t* colors);

/**
 * @brief Reads the theme's major and minor typefaces. The default theme's
 *        faces are reported when the theme part is absent or unparseable
 *        (see `fm_workbook_get_theme_colors` for telling the cases apart).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_theme_fonts(const fm_workbook_t* wb, fm_theme_fonts_t* out);

/**
 * @brief Writes the theme's major and minor typefaces. The strings are
 *        copied. Part generation and failure rules match
 *        `fm_workbook_set_theme_colors`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument or any face is `NULL`;
 *         `FM_STATUS_IO_XML_PARSE` (5007) when the existing theme part cannot be parsed.
 */
FM_API fm_status_t fm_workbook_set_theme_fonts(fm_workbook_t* wb, const fm_theme_fonts_t* fonts);

/**
 * @brief Counts the `a:font script=... typeface=...` entries of the theme's
 *        major (`major` nonzero) or minor (`major == 0`) font set.
 *
 * The default Office theme's 47 scripts are reported when the theme part is
 * absent or unparseable (the same rule as `fm_workbook_get_theme_fonts`).
 * The theme is parsed on every call.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_theme_script_font_count(const fm_workbook_t* wb, int32_t major, size_t* out_count);

/**
 * @brief Reads one script entry (in the theme part's file order) of the
 *        major or minor font set.
 *
 * `*out_script` is an ISO 15924 code such as `Jpan`; both strings are
 * read-scratch-backed views, valid until the next call that uses the
 * workbook's read scratch.
 *
 * To pick the face for text whose font has `scheme` major or minor, choose
 * the script from the font's `charset`: 128 -> `Jpan`, 129 -> `Hang`,
 * 134 -> `Hans`, 136 -> `Hant`, 222 -> `Thai`, anything else -> the latin
 * face (`fm_workbook_get_theme_fonts`; charsets such as 204 have no script
 * entry). Use the same table when choosing by the script of the characters
 * in the text.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is not below the count.
 */
FM_API fm_status_t fm_workbook_get_theme_script_font_at(const fm_workbook_t* wb, int32_t major, size_t idx,
                                                        const char** out_script, const char** out_typeface);

/**
 * @brief Sets the typeface of one script in the major (`major` nonzero) or
 *        minor font set. The strings are copied.
 *
 * Other scripts are unchanged; a script not yet in the set is appended.
 * Setting `Jpan` also changes the east-Asian face reported by
 * `fm_workbook_get_theme_fonts`. A workbook without a theme part first gets
 * one carrying the full default set, so the other 46 scripts keep their
 * default values. Part generation and failure rules match
 * `fm_workbook_set_theme_colors`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `script` is not exactly four ASCII letters
 *         or `typeface` is empty;
 *         `FM_STATUS_IO_XML_PARSE` (5007) when the existing theme part cannot be parsed.
 */
FM_API fm_status_t fm_workbook_set_theme_script_font(fm_workbook_t* wb, int32_t major, const char* script,
                                                     const char* typeface);

/**
 * @brief Returns the workbook to the state of having no theme part, so
 *        `fm_workbook_get_theme_colors` reports source 1 and the default
 *        theme's colours and faces.
 *
 * The theme part, its relationships and anything only it referenced are
 * removed, together with its content type and the workbook relationship; a
 * theme part that cannot be parsed is removed the same way. A workbook
 * without a theme part is left unchanged. The default theme is not Excel's
 * current default look, and Excel adds a theme part again when it saves the
 * file.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`.
 */
FM_API fm_status_t fm_workbook_clear_theme(fm_workbook_t* wb);

/**
 * @brief Resolves a colour specification to the 0xAARRGGBB Excel renders.
 *
 * Theme colours take `spec.tint`; indexed colours read the workbook's
 * `<indexedColors>` over the built-in palette; `kFmColorAuto`,
 * `kFmColorNone` and the system indices 64 / 65 depend on `context`
 * (an `fm_color_context_t`). `*out_resolution` is an
 * `fm_color_resolution_t`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an unknown `context` or `spec.kind`, or a
 *         non-finite theme tint.
 */
FM_API fm_status_t fm_workbook_resolve_color(const fm_workbook_t* wb, fm_color_spec_t spec, int32_t context,
                                             uint32_t* out_argb, int32_t* out_resolution);

/**
 * @brief Reads the effective style of the cell at `(row, col)`; see
 *        `fm_effective_style_t`. The cell need not exist.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         coordinate is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_effective_style(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                fm_effective_style_t* out);

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
 * replaced (the OOXML reader replaces it; entering a formula that names a
 * new book appends to it). `target_external` follows the wide-POD
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
 * `<externalReferences>` block. A formula, defined name, conditional-format
 * rule or data validation entered through this API that names an
 * unregistered book (bracket or path spelling, e.g. `[Book.xlsx]Sheet1!A1`)
 * adds a record at the end without renumbering the others; a book already
 * present, in any letter case, adds none.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_external_link_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `index`-th external-link record (1-based document
 *        order minus one — i.e. `0` is the first `<externalReference>`).
 *
 * The returned string pointers are model-backed views into workbook-owned
 * storage; see `fm_external_link_record_t` for lifetime rules.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `index >= external_link_count`.
 */
FM_API fm_status_t fm_workbook_get_external_link(const fm_workbook_t* wb, size_t index, fm_external_link_record_t* out);

/* -------------------------------------------------------------------------- */
/* Drawing objects                                                            */
/* -------------------------------------------------------------------------- */

/** @brief Raster format of an image; `UNKNOWN` marks "no image bytes found". */
typedef enum {
  FM_IMAGE_FORMAT_UNKNOWN = 0,
  FM_IMAGE_FORMAT_PNG = 1,
  FM_IMAGE_FORMAT_JPEG = 2,
  FM_IMAGE_FORMAT_GIF = 3,
  FM_IMAGE_FORMAT_BMP = 4,
} fm_image_format_t;

/** @brief What a drawing anchor holds. */
typedef enum {
  FM_DRAWING_OBJECT_PICTURE = 0,
  FM_DRAWING_OBJECT_SHAPE = 1,
  FM_DRAWING_OBJECT_CHART = 2,
  FM_DRAWING_OBJECT_GROUP = 3,
  FM_DRAWING_OBJECT_CONNECTOR = 4,
  FM_DRAWING_OBJECT_GRAPHIC_FRAME = 5,
  FM_DRAWING_OBJECT_OTHER = 6,
} fm_drawing_object_kind_t;

/** @brief The anchor element type of a drawing object. */
typedef enum {
  FM_ANCHOR_KIND_ONE_CELL = 0,
  FM_ANCHOR_KIND_TWO_CELL = 1,
  FM_ANCHOR_KIND_ABSOLUTE = 2,
} fm_anchor_kind_t;

/** @brief How a drawing object follows row and column edits. */
typedef enum {
  FM_ANCHOR_EDIT_AS_TWO_CELL = 0,
  FM_ANCHOR_EDIT_AS_ONE_CELL = 1,
  FM_ANCHOR_EDIT_AS_ABSOLUTE = 2,
} fm_anchor_edit_as_t;

/**
 * @brief Format and pixel size of an image. `format` is an
 *        `fm_image_format_t`.
 */
typedef struct {
  int32_t format;
  uint32_t px_width;
  uint32_t px_height;
} fm_image_info_t;

/**
 * @brief One anchored object of a sheet's drawing.
 *
 * `kind` is an `fm_drawing_object_kind_t`, `anchor_kind` an
 * `fm_anchor_kind_t` and `edit_as` an `fm_anchor_edit_as_t`. `from_*` is the
 * top-left marker (an absolute anchor reports its position from A1 as
 * `from_col_off` = x and `from_row_off` = y); `to_*` is the bottom-right
 * marker of a two-cell anchor and zero otherwise. `cx` and `cy` are the size
 * in EMU. `image_format` is an `fm_image_format_t`, non-zero for a picture
 * whose media part is present. `name`, `descr` and `media_path` are
 * read-scratch-backed views (the lifetime `fm_sheet_get_auto_filter`
 * describes); `media_path` is empty for a non-picture.
 */
typedef struct {
  uint32_t object_id;
  int32_t kind;
  int32_t anchor_kind;
  int32_t edit_as;
  uint32_t from_row;
  uint32_t from_col;
  int64_t from_row_off;
  int64_t from_col_off;
  uint32_t to_row;
  uint32_t to_col;
  int64_t to_row_off;
  int64_t to_col_off;
  int64_t cx;
  int64_t cy;
  int32_t image_format;
  const char* name;
  const char* descr;
  const char* media_path;
} fm_drawing_object_t;

/**
 * @brief Placement of an inserted image.
 *
 * `anchor_kind` is `FM_ANCHOR_KIND_ONE_CELL` or `FM_ANCHOR_KIND_TWO_CELL`;
 * `edit_as` (an `fm_anchor_edit_as_t`) is honoured for a two-cell anchor
 * only. `row`/`col` is the 0-based top-left cell and `*_off_emu` the offset
 * inside it. A `width_emu` or `height_emu` of `0` means the image's pixel
 * size times 9525. `name` and `descr` may be `NULL`, meaning empty; an empty
 * name becomes `Picture <id>`.
 */
typedef struct {
  const char* name;
  const char* descr;
  int32_t anchor_kind;
  int32_t edit_as;
  uint32_t row;
  uint32_t col;
  int64_t row_off_emu;
  int64_t col_off_emu;
  int64_t width_emu;
  int64_t height_emu;
} fm_image_insert_t;

/**
 * @brief New placement of an existing picture or chart, in the quantities
 *        `fm_sheet_get_drawing_object_at` reports.
 *
 * `anchor_kind` is `FM_ANCHOR_KIND_ONE_CELL` or `FM_ANCHOR_KIND_TWO_CELL`;
 * `edit_as` (an `fm_anchor_edit_as_t`) is honoured for a two-cell anchor
 * only. `row`/`col`/`*_off_emu` are the top-left marker of the anchor box
 * (`from_*`). `width_emu`/`height_emu` are the object's own unrotated size
 * (`cx`/`cy`); `0` keeps the current value.
 */
typedef struct {
  int32_t anchor_kind;
  int32_t edit_as;
  uint32_t row;
  uint32_t col;
  int64_t row_off_emu;
  int64_t col_off_emu;
  int64_t width_emu;
  int64_t height_emu;
} fm_drawing_anchor_t;

/** @brief `fm_sheet_restore_drawing` flag: restore as a copy with a fresh id, added on top. */
#define FM_DRAWING_RESTORE_NEW_ID 1u

/**
 * @brief Sniffs the format and pixel size of PNG, JPEG, GIF or BMP bytes
 *        from the header alone, without touching the workbook.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_IO_IMAGE_UNSUPPORTED` (5023) when the bytes are not a recognised
 *         image, the header is truncated, or a dimension is zero.
 */
FM_API fm_status_t fm_workbook_get_image_info(const fm_workbook_t* wb, const uint8_t* bytes, size_t len,
                                              fm_image_info_t* out);

/**
 * @brief Number of anchored objects in the sheet's drawing; `0` for a sheet
 *        without a drawing.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_drawing_object_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads object `idx` of the sheet's drawing, in document order.
 *
 * Document order is the stacking order: item 0 is at the back and the last
 * item is on top.
 *
 * The three strings are read-scratch-backed views.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_drawing_object_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                  fm_drawing_object_t* out);

/**
 * @brief Returns the media bytes and header info of the picture `object_id`.
 *
 * `*out_bytes` points into the workbook's own storage and is not
 * NUL-terminated: it stays valid until the next call that mutates the
 * workbook (insert, remove, any edit) or the handle is destroyed; unrelated
 * reads do not invalidate it. Copy it to keep it longer.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet
 *         has no picture `object_id` with a present media part;
 *         `FM_STATUS_IO_IMAGE_UNSUPPORTED` (5023) when the media bytes are not a
 *         recognised image;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_image(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                      const uint8_t** out_bytes, size_t* out_len, fm_image_info_t* out_info);

/**
 * @brief Inserts `bytes` as a picture and returns its object id, unique
 *        within the sheet's drawing.
 *
 * Existing charts, shapes and every other drawing content are kept. A sheet
 * without a drawing gets a new one.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an enum
 *         field is outside its domain, the anchor kind is absolute, an
 *         offset or size is negative or 2^31 EMU or more, the anchor lies
 *         outside the sheet, the image exceeds 32 MiB, or the package would
 *         exceed the reader's limits;
 *         `FM_STATUS_IO_IMAGE_UNSUPPORTED` (5023) when `bytes` is not a recognised image;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package. The workbook is then
 *         unchanged.
 */
FM_API fm_status_t fm_sheet_add_image(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes, size_t len,
                                      const fm_image_insert_t* opts, uint32_t* out_object_id);

/**
 * @brief Removes the picture or chart `object_id`, the relationships it
 *        names, and the parts it reaches (media, or a chart part and its
 *        style, colors and embeddings) once nothing else references them.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet
 *         has no picture or chart with that id;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package.
 */
FM_API fm_status_t fm_sheet_remove_drawing(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id);

/**
 * @brief Moves and resizes the picture or chart `object_id` (the first one
 *        when ids repeat), keeping its id, name, crop, rotation, flips and
 *        every other element.
 *
 * Only the anchor markers, `editAs` when its value changes, and the
 * object's `a:off`/`a:ext` (a chart's `xdr:xfrm`) are rewritten. A picture
 * rotated into [45, 135) or [225, 315) degrees gets an anchor box with
 * width and height swapped about the same centre. A placement equal to the
 * current one writes nothing. An absolute object can be converted by naming
 * a one-cell or two-cell kind.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an enum
 *         field is outside its domain, the anchor kind is absolute, an
 *         offset or size is negative or 2^31 EMU or more, the anchor lies
 *         outside the sheet, or the sheet has no picture or chart with that
 *         id;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `anchor == NULL`;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package.
 */
FM_API fm_status_t fm_sheet_set_drawing_anchor(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                               const fm_drawing_anchor_t* anchor);

/**
 * @brief Renames the picture or chart `object_id` (the first one when ids
 *        repeat). Duplicate names are allowed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `name` is
 *         empty or not valid UTF-8, or the sheet has no picture or chart with
 *         that id;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package.
 */
FM_API fm_status_t fm_sheet_set_drawing_name(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                             const char* name);

/**
 * @brief Moves the picture or chart `object_id` so it becomes item `index`
 *        of the drawing's object list (0 is the back).
 *
 * The unit moved is the object's top-level element, a whole
 * `mc:AlternateContent` included; other objects keep their relative order
 * and ids. An unchanged order writes nothing.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet
 *         has no picture or chart with that id, or `index` is not less than
 *         the object count;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package.
 */
FM_API fm_status_t fm_sheet_set_drawing_z_order(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                size_t index);

/**
 * @brief Captures the picture or chart `object_id` as an opaque, versioned
 *        byte string: its list position, its element, the relationships it
 *        names with their parts, and every part reachable from those parts'
 *        relationships (a chart's style, colors, user shapes, embeddings).
 *
 * On success `*out_bytes` is a heap-allocated buffer of `*out_len` bytes
 * that MUST be released with `fm_buffer_destroy`. Both are zeroed before
 * validation and stay zero on failure. Hand the bytes to
 * `fm_sheet_restore_drawing` to put the object back.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet
 *         has no picture or chart with that id;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package.
 */
FM_API fm_status_t fm_sheet_snapshot_drawing(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                             uint8_t** out_bytes, size_t* out_len);

/**
 * @brief Puts a `fm_sheet_snapshot_drawing` capture (this version or an
 *        earlier one) into sheet `sheet_index` and returns the object's id.
 *
 * An object with the same id is replaced, the parts it reaches are
 * rewritten in place, and the capture returns to its recorded list position
 * (capped at the object count); an id not present is added at that
 * position. With `FM_DRAWING_RESTORE_NEW_ID` the capture is added on top as
 * a copy with a fresh id and no `a16:creationId`, its chart parts copied to
 * fresh paths. A part whose path is free is written there, a media part
 * with identical bytes is shared, and any other occupied path takes a fresh
 * one with the relationships naming it re-pointed. A sheet without a
 * drawing gets a new one.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the bytes
 *         are not a valid capture, `flags` has an unknown bit, the id is
 *         held by another kind of object, or the package would exceed the
 *         reader's limits;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package. The workbook is then
 *         unchanged.
 */
FM_API fm_status_t fm_sheet_restore_drawing(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes, size_t len,
                                            uint32_t flags, uint32_t* out_object_id);

/* -------------------------------------------------------------------------- */
/* Drawing shapes                                                             */
/* -------------------------------------------------------------------------- */

/** @brief How a shape's fill or outline is specified (`fm_shape_t::fill_kind`, `line_kind`). */
typedef enum {
  FM_SHAPE_PAINT_UNSET = 0, /**< No fill or outline element. */
  FM_SHAPE_PAINT_SOLID = 1, /**< Solid `srgbClr`; the colour is `fill_rgb` / `line_rgb`. */
  FM_SHAPE_PAINT_NONE = 2,  /**< `noFill`. */
  FM_SHAPE_PAINT_OTHER = 3, /**< Gradient, pattern, picture or theme colour; read only. */
} fm_shape_paint_t;

/** @brief `fm_shape_t::fields` bits: the fields an insert or update applies. */
#define FM_SHAPE_FIELD_NAME 0x1u
#define FM_SHAPE_FIELD_PRST 0x2u
#define FM_SHAPE_FIELD_ANCHOR 0x4u /* anchor_kind, edit_as, from_*, to_*, cx, cy */
#define FM_SHAPE_FIELD_TEXT 0x8u
#define FM_SHAPE_FIELD_TEXT_BOX 0x10u
#define FM_SHAPE_FIELD_FILL 0x20u /* fill_kind (SOLID or NONE) and fill_rgb */
#define FM_SHAPE_FIELD_LINE 0x40u /* line_kind (SOLID or NONE) and line_rgb */
#define FM_SHAPE_FIELD_LINE_WIDTH 0x80u
#define FM_SHAPE_FIELD_ROTATION 0x100u
#define FM_SHAPE_FIELD_HIDDEN 0x200u

/**
 * @brief A shape, connector or group of a sheet's drawing; the record
 *        `fm_sheet_get_shape_at` reads and the insert / update calls take.
 *
 * The anchor fields read as `fm_drawing_object_t` does: `kind` is an
 * `fm_drawing_object_kind_t` (shape, connector or group), an absolute anchor
 * reports its position as `from_col_off` = x and `from_row_off` = y, and
 * `cx`/`cy` are the size in EMU. `fill_kind` / `line_kind` are
 * `fm_shape_paint_t`; `fill_rgb` / `line_rgb` are 0xRRGGBB when the kind is
 * solid. `rotation` is in 60000ths of a degree and `line_width_emu` is `0`
 * without a width. `prst` is the preset geometry, empty for custom geometry
 * or a group; `text` joins paragraphs with `\n`.
 *
 * `fields` is zero on read. On insert and update it is a mask of
 * `FM_SHAPE_FIELD_*`: only the named fields are applied (an insert
 * defaults `prst` to `rect` and `name` to `Shape <id>`); `object_id`, `kind`
 * and the other fields are ignored. A two-cell anchor's size is its box and
 * its `cx`/`cy` are ignored; an update cannot change the anchor kind. The
 * strings are read-scratch-backed views on read and may be `NULL` (empty) on
 * input.
 */
typedef struct {
  uint32_t fields;
  uint32_t object_id;
  int32_t kind;
  int32_t anchor_kind;
  int32_t edit_as;
  uint32_t from_row;
  uint32_t from_col;
  uint32_t to_row;
  uint32_t to_col;
  int32_t fill_kind;
  uint32_t fill_rgb;
  int32_t line_kind;
  uint32_t line_rgb;
  int32_t rotation;
  int32_t hidden;
  int32_t text_box;
  int64_t from_row_off;
  int64_t from_col_off;
  int64_t to_row_off;
  int64_t to_col_off;
  int64_t cx;
  int64_t cy;
  int64_t line_width_emu;
  const char* name;
  const char* prst;
  const char* text;
} fm_shape_t;

/**
 * @brief Number of shapes, connectors and groups in the sheet's drawing; `0`
 *        for a sheet without a drawing. Pictures, charts and other frames are
 *        not counted; children of a group are not listed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_shape_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads shape `idx` of the sheet's drawing, in z-order (0 is the back).
 *
 * Fields the model does not express read as `FM_SHAPE_PAINT_OTHER` or empty
 * and pass through a save untouched. The three strings are read-scratch-backed
 * views.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_shape_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, fm_shape_t* out);

/**
 * @brief Appends a shape on top of the sheet's drawing and returns its id.
 *
 * `FM_SHAPE_FIELD_ANCHOR` is required. A `prst` naming a line or connector
 * geometry (`line`, `straightConnector1`, `bentConnector*`,
 * `curvedConnector*`) inserts a connector, which holds no text. A sheet
 * without a drawing gets one; existing drawing content is kept.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `spec` or `out_object_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a bad enumeration, a missing anchor,
 *         a `prst` that is not an alphanumeric token, an offset or size
 *         outside [0, 2^31) EMU, a cell outside the sheet, a two-cell `to`
 *         before `from`, a fill or outline that is neither solid nor none,
 *         or text with control characters;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package.
 */
FM_API fm_status_t fm_sheet_add_shape(fm_workbook_t* wb, size_t sheet_index, const fm_shape_t* spec,
                                      uint32_t* out_object_id);

/**
 * @brief Applies the fields `spec->fields` names to the shape, connector or
 *        group `object_id`, keeping the rest of its XML.
 *
 * Text applies to a shape only and keeps the first paragraph's properties and
 * first run's formatting. Fails like `fm_sheet_add_shape`, and with
 * `FM_STATUS_INVALID_ARGUMENT` (2) for an unknown id or a changed anchor kind.
 */
FM_API fm_status_t fm_sheet_update_shape(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                         const fm_shape_t* spec);

/**
 * @brief Removes the shape, connector or group `object_id` with its children.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet has no such shape;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_remove_shape(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id);

/* -------------------------------------------------------------------------- */
/* Table slicers                                                              */
/* -------------------------------------------------------------------------- */

/** @brief `fm_slicer_t::sort_order`. */
typedef enum {
  FM_SLICER_SORT_ORDER_ASCENDING = 0,
  FM_SLICER_SORT_ORDER_DESCENDING = 1,
} fm_slicer_sort_order_t;

/** @brief `fm_slicer_t::cross_filter`. */
typedef enum {
  FM_SLICER_CROSS_SHOW_ITEMS_WITH_DATA_AT_TOP = 0,
  FM_SLICER_CROSS_SHOW_ITEMS_WITH_NO_DATA = 1,
  FM_SLICER_CROSS_NONE = 2,
} fm_slicer_cross_filter_t;

/** @brief `fm_slicer_t::fields` bits: the fields an add or update applies. */
#define FM_SLICER_FIELD_NAME 0x1u /* add only */
#define FM_SLICER_FIELD_CAPTION 0x2u
#define FM_SLICER_FIELD_STYLE 0x4u /* an empty style removes it */
#define FM_SLICER_FIELD_COLUMN_COUNT 0x8u
#define FM_SLICER_FIELD_ROW_HEIGHT 0x10u
#define FM_SLICER_FIELD_ANCHOR 0x20u /* from_*, to_*; required by add */
#define FM_SLICER_FIELD_SORT_ORDER 0x40u
#define FM_SLICER_FIELD_CROSS_FILTER 0x80u

/**
 * @brief A table slicer; the record `fm_sheet_get_slicer_at` reads and the add /
 *        update calls take.
 *
 * `object_id` is the slicer's frame in the sheet's drawing (`0` when the
 * drawing holds none). `anchor_kind` is an `fm_anchor_kind_t`; `from_*` and
 * `to_*` are the frame's markers (`to_*` is zero unless the anchor is
 * two-cell) and an add or update always writes a two-cell anchor. `row_height`
 * is in EMU. `sort_order` is an `fm_slicer_sort_order_t` and `cross_filter` an
 * `fm_slicer_cross_filter_t`. `cache_name`, `table_name` and `column_name`
 * identify the slicer cache and the table column it filters; the selection
 * itself lives in the table's AutoFilter, not here. `style` is empty without
 * a style.
 *
 * `fields` is zero on read. On add and update it is a mask of
 * `FM_SLICER_FIELD_*`: only the named fields are applied. An add defaults
 * `name` and `caption` to the column name (the name is made unique), `row_height`
 * to 241300 and `column_count` to 1. The strings are read-scratch-backed
 * views on read and may be `NULL` (empty) on input; `object_id`,
 * `anchor_kind`, `cache_name`, `table_name` and `column_name` are ignored on
 * input.
 */
typedef struct {
  uint32_t fields;
  uint32_t object_id;
  uint32_t column_count;
  int32_t sort_order;
  int32_t cross_filter;
  int32_t anchor_kind;
  uint32_t from_row;
  uint32_t from_col;
  uint32_t to_row;
  uint32_t to_col;
  int64_t from_row_off;
  int64_t from_col_off;
  int64_t to_row_off;
  int64_t to_col_off;
  int64_t row_height;
  const char* name;
  const char* caption;
  const char* cache_name;
  const char* table_name;
  const char* column_name;
  const char* style;
} fm_slicer_t;

/**
 * @brief Number of table slicers hosted by the sheet; `0` for a sheet without
 *        any. Slicers over a pivot cache are not counted.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the sheet's drawing does not parse.
 */
FM_API fm_status_t fm_sheet_get_slicer_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads slicer `idx` of the sheet. The strings are read-scratch-backed
 *        views.
 *
 * @return As `fm_sheet_get_slicer_count`, with `FM_STATUS_INVALID_ARGUMENT` (2) also for an
 *         out-of-range `idx`.
 */
FM_API fm_status_t fm_sheet_get_slicer_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, fm_slicer_t* out);

/**
 * @brief Adds a slicer on the sheet over column `column_index` (0-based) of
 *        table `table_index` and returns its frame id.
 *
 * Creates the slicer and cache parts, their relationships and content types,
 * the `Slicer_*` defined name and the drawing frame (with its fallback shape),
 * or extends the sheet's existing slicer part. `FM_SLICER_FIELD_ANCHOR` is
 * required. The table may live on another sheet.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `spec` or `out_object_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a bad table, column, field bit,
 *         anchor, column count, row height or text with control characters;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package.
 */
FM_API fm_status_t fm_sheet_add_slicer(fm_workbook_t* wb, size_t sheet_index, size_t table_index, size_t column_index,
                                       const fm_slicer_t* spec, uint32_t* out_object_id);

/**
 * @brief Applies the fields `spec->fields` names (all but the name) to slicer
 *        `idx`, keeping the rest of its XML.
 *
 * @return As `fm_sheet_add_slicer`, with `FM_STATUS_INVALID_ARGUMENT` (2) also for an
 *         out-of-range `idx`, a name field, or an anchor update of a slicer
 *         whose frame is missing or not two-cell.
 */
FM_API fm_status_t fm_sheet_update_slicer(fm_workbook_t* wb, size_t sheet_index, size_t idx, const fm_slicer_t* spec);

/**
 * @brief Removes slicer `idx` with its drawing frame, and its cache when no
 *        other slicer uses it, and the sheet's drawing when it held nothing
 *        else. Removing a table or a sheet, or deleting the rows or columns
 *        under a slicer's frame, removes the slicers bound to it the same way.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the sheet's drawing does not parse.
 */
FM_API fm_status_t fm_sheet_remove_slicer(fm_workbook_t* wb, size_t sheet_index, size_t idx);

/* -------------------------------------------------------------------------- */
/* Sparkline groups                                                           */
/* -------------------------------------------------------------------------- */

/** @brief `fm_sparkline_group_t::type`. */
typedef enum {
  FM_SPARKLINE_LINE = 0,
  FM_SPARKLINE_COLUMN = 1,
  FM_SPARKLINE_STACKED = 2,
} fm_sparkline_type_t;

/** @brief `fm_sparkline_group_t::min_axis_type` and `max_axis_type`. */
typedef enum {
  FM_SPARKLINE_AXIS_INDIVIDUAL = 0,
  FM_SPARKLINE_AXIS_GROUP = 1,
  FM_SPARKLINE_AXIS_CUSTOM = 2,
} fm_sparkline_axis_type_t;

/** @brief `fm_sparkline_group_t::empty_cells_as`. */
typedef enum {
  FM_SPARKLINE_EMPTY_GAP = 0,
  FM_SPARKLINE_EMPTY_ZERO = 1,
  FM_SPARKLINE_EMPTY_SPAN = 2,
} fm_sparkline_empty_cells_t;

/**
 * @brief One sparkline: the data `formula` (`xm:f`, a sheet-qualified range
 *        without `=`) and the single-cell `location` (`xm:sqref`, A1 text)
 *        that shows it. Read-scratch-backed views on read.
 */
typedef struct {
  const char* formula;
  const char* location;
} fm_sparkline_t;

/**
 * @brief A sparkline group (`x14:sparklineGroup`) of the sheet's worksheet
 *        extension list.
 *
 * Each colour has `kind == kFmColorNone` when the element is absent.
 * `has_manual_min` / `has_manual_max` say whether `manualMin` / `manualMax`
 * are present. `date_axis_ref` is the group's own `xm:f` date range, empty
 * when absent. `line_weight` is in points (default 0.75). `sparkline_count`
 * is filled on read and ignored on input. Enumerations are the
 * `fm_sparkline_*_t` types; flags are 0/1.
 */
typedef struct {
  fm_color_spec_t color_series;
  fm_color_spec_t color_negative;
  fm_color_spec_t color_axis;
  fm_color_spec_t color_markers;
  fm_color_spec_t color_first;
  fm_color_spec_t color_last;
  fm_color_spec_t color_high;
  fm_color_spec_t color_low;
  double manual_min;
  double manual_max;
  double line_weight;
  int32_t type;
  int32_t markers;
  int32_t high;
  int32_t low;
  int32_t first;
  int32_t last;
  int32_t negative;
  int32_t display_x_axis;
  int32_t display_hidden;
  int32_t min_axis_type;
  int32_t max_axis_type;
  int32_t has_manual_min;
  int32_t has_manual_max;
  int32_t date_axis;
  int32_t empty_cells_as;
  const char* date_axis_ref;
  size_t sparkline_count;
} fm_sparkline_group_t;

/**
 * @brief Number of sparkline groups in the sheet's worksheet extension list.
 *
 * Groups are addressed by their position in document order; adding or
 * removing one shifts the later indices. Groups of a workbook loaded from
 * `.xlsb` are listed and editable the same way; a group the reader cannot
 * decode stays a binary record and is not listed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_sparkline_group_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads sparkline group `idx`; `date_axis_ref` is a read-scratch-backed view.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_sparkline_group(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                fm_sparkline_group_t* out);

/**
 * @brief Reads sparkline `idx` of group `group_idx`; both strings are
 *        read-scratch-backed views.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `group_idx` or `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_get_sparkline(const fm_workbook_t* wb, size_t sheet_index, size_t group_idx, size_t idx,
                                          fm_sparkline_t* out);

/**
 * @brief Appends a sparkline group with `count` sparklines.
 *
 * Attributes and children outside the model are not written. A sheet without
 * a worksheet extension list gets one; other extensions are kept.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `group` or `items` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for `count == 0`, a sparkline whose
 *         location is not one A1 cell or whose formula is empty, a
 *         non-finite number, a negative line weight, or an enumeration out of
 *         range.
 */
FM_API fm_status_t fm_sheet_add_sparkline_group(fm_workbook_t* wb, size_t sheet_index,
                                                const fm_sparkline_group_t* group, const fm_sparkline_t* items,
                                                size_t count);

/**
 * @brief Rewrites the modelled attributes, colours and date-axis reference of
 *        group `idx` from `group`; its sparklines and every attribute or
 *        child outside the model are kept.
 *
 * @return `FM_STATUS_OK` (0) on success; failures as `fm_sheet_add_sparkline_group`, and
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_update_sparkline_group(fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                   const fm_sparkline_group_t* group);

/**
 * @brief Replaces the sparklines of group `idx` with `count` new ones.
 *
 * @return `FM_STATUS_OK` (0) on success; failures as `fm_sheet_update_sparkline_group`.
 */
FM_API fm_status_t fm_sheet_set_sparkline_group_sparklines(fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                           const fm_sparkline_t* items, size_t count);

/**
 * @brief Removes sparkline group `idx`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_sparkline_group(fm_workbook_t* wb, size_t sheet_index, size_t idx);

/* -------------------------------------------------------------------------- */
/* Charts                                                                     */
/* -------------------------------------------------------------------------- */

/*
 * A chart is addressed as `(sheet_index, object_id)`, the `cNvPr` id
 * `fm_sheet_get_drawing_object_at` reports with `FM_DRAWING_OBJECT_CHART`; the
 * drawing operations above move, reorder, capture, restore and remove it.
 *
 * Reads go through an owned handle (`fm_chart_t`, `fm_chart_data_t`) that
 * snapshots the chart when opened; every string it hands out is UTF-8 owned
 * by the handle and stays valid until the handle is destroyed. On the read
 * side an optional integer has a `has_*` flag, an optional double is NaN
 * when absent, and an optional string is `NULL` when absent.
 *
 * Patch structs carry one `fm_chart_patch_op_t` per field: a zeroed struct
 * changes nothing, `FM_CHART_PATCH_CLEAR` removes the element, and
 * `FM_CHART_PATCH_SET` writes the value beside it. A field that cannot be
 * removed rejects `CLEAR` with `FM_STATUS_INVALID_ARGUMENT` (2). Colors are ARGB
 * (`0xAARRGGBB`); transparency is a percentage (0..100).
 *
 * Edits patch the chart part and write nothing when the result is
 * unchanged. Common failures, besides those each entry lists:
 * `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) for an out-of-range sheet;
 * `FM_STATUS_INVALID_ARGUMENT` (2) for an unknown id, plot, series or
 * axis, or an enum value outside its domain; `FM_STATUS_IO_CHART_UNPARSEABLE` (5025)
 * when the chart part does not parse; `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for
 * an edit to a chartex chart; `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) for an edit to
 * a chart on a drawing retained from an `.xlsb` package. A failed edit
 * leaves the workbook unchanged.
 */

/** @brief Plot group type. Mirrors `formulon::ChartPlotType`. */
typedef enum {
  FM_CHART_PLOT_BAR = 0,
  FM_CHART_PLOT_BAR_3D = 1,
  FM_CHART_PLOT_LINE = 2,
  FM_CHART_PLOT_LINE_3D = 3,
  FM_CHART_PLOT_AREA = 4,
  FM_CHART_PLOT_AREA_3D = 5,
  FM_CHART_PLOT_PIE = 6,
  FM_CHART_PLOT_PIE_3D = 7,
  FM_CHART_PLOT_DOUGHNUT = 8,
  FM_CHART_PLOT_OF_PIE = 9,
  FM_CHART_PLOT_SCATTER = 10,
  FM_CHART_PLOT_BUBBLE = 11,
  FM_CHART_PLOT_RADAR = 12,
  FM_CHART_PLOT_STOCK = 13,
  FM_CHART_PLOT_SURFACE = 14,
  FM_CHART_PLOT_SURFACE_3D = 15,
  FM_CHART_PLOT_CHARTEX = 16,
} fm_chart_plot_type_t;

/** @brief Kinds `fm_sheet_add_chart` and `fm_sheet_set_chart_type` produce. */
typedef enum {
  FM_CHART_INSERT_COLUMN = 0,
  FM_CHART_INSERT_BAR = 1,
  FM_CHART_INSERT_LINE = 2,
  FM_CHART_INSERT_AREA = 3,
  FM_CHART_INSERT_PIE = 4,
  FM_CHART_INSERT_SCATTER = 5,
  FM_CHART_INSERT_DOUGHNUT = 6,
} fm_chart_insert_kind_t;

typedef enum {
  FM_CHART_GROUPING_CLUSTERED = 0,
  FM_CHART_GROUPING_STACKED = 1,
  FM_CHART_GROUPING_PERCENT_STACKED = 2,
  FM_CHART_GROUPING_STANDARD = 3,
} fm_chart_grouping_t;

/** @brief `c:dispBlanksAs`; `ABSENT` means the element is not written. */
typedef enum {
  FM_CHART_BLANKS_ABSENT = 0,
  FM_CHART_BLANKS_GAP = 1,
  FM_CHART_BLANKS_ZERO = 2,
  FM_CHART_BLANKS_SPAN = 3,
} fm_chart_blanks_as_t;

/** @brief How a resolved point is drawn. */
typedef enum {
  FM_CHART_POINT_VALUE = 0,
  FM_CHART_POINT_GAP = 1,
  FM_CHART_POINT_CONNECT = 2,
} fm_chart_point_state_t;

typedef enum {
  FM_CHART_AXIS_CATEGORY = 0,
  FM_CHART_AXIS_VALUE = 1,
  FM_CHART_AXIS_DATE = 2,
  FM_CHART_AXIS_SERIES = 3,
} fm_chart_axis_kind_t;

typedef enum {
  FM_CHART_AXIS_POS_BOTTOM = 0,
  FM_CHART_AXIS_POS_TOP = 1,
  FM_CHART_AXIS_POS_LEFT = 2,
  FM_CHART_AXIS_POS_RIGHT = 3,
} fm_chart_axis_position_t;

/** @brief Where an axis crosses its cross axis; `AT` uses `crosses_at`. */
typedef enum {
  FM_CHART_CROSSES_ABSENT = 0,
  FM_CHART_CROSSES_AUTO_ZERO = 1,
  FM_CHART_CROSSES_MIN = 2,
  FM_CHART_CROSSES_MAX = 3,
  FM_CHART_CROSSES_AT = 4,
} fm_chart_axis_crosses_t;

/** @brief `AUTO` (title element without text) lets the viewer supply its default. */
typedef enum {
  FM_CHART_TITLE_NONE = 0,
  FM_CHART_TITLE_TEXT = 1,
  FM_CHART_TITLE_REF = 2,
  FM_CHART_TITLE_AUTO = 3,
} fm_chart_title_kind_t;

typedef enum {
  FM_CHART_DIRECTION_AUTO = 0,
  FM_CHART_DIRECTION_COLUMNS = 1,
  FM_CHART_DIRECTION_ROWS = 2,
} fm_chart_series_direction_t;

/** @brief Whether a source range's first row / column is used as headings. */
typedef enum {
  FM_CHART_HEADER_AUTO = 0,
  FM_CHART_HEADER_YES = 1,
  FM_CHART_HEADER_NO = 2,
} fm_chart_header_mode_t;

typedef enum {
  FM_CHART_LEGEND_NONE = 0,
  FM_CHART_LEGEND_BOTTOM = 1,
  FM_CHART_LEGEND_TOP = 2,
  FM_CHART_LEGEND_LEFT = 3,
  FM_CHART_LEGEND_RIGHT = 4,
  FM_CHART_LEGEND_TOP_RIGHT = 5,
} fm_chart_legend_position_t;

typedef enum {
  FM_CHART_DASH_SOLID = 0,
  FM_CHART_DASH_DASH = 1,
  FM_CHART_DASH_DOT = 2,
  FM_CHART_DASH_DASH_DOT = 3,
  FM_CHART_DASH_OTHER = 4,
} fm_chart_line_dash_t;

typedef enum {
  FM_CHART_MARKER_NONE = 0,
  FM_CHART_MARKER_AUTO = 1,
  FM_CHART_MARKER_CIRCLE = 2,
  FM_CHART_MARKER_SQUARE = 3,
  FM_CHART_MARKER_DIAMOND = 4,
  FM_CHART_MARKER_TRIANGLE = 5,
  FM_CHART_MARKER_X = 6,
  FM_CHART_MARKER_STAR = 7,
  FM_CHART_MARKER_PLUS = 8,
  FM_CHART_MARKER_DOT = 9,
  FM_CHART_MARKER_DASH = 10,
  FM_CHART_MARKER_PICTURE = 11,
} fm_chart_marker_symbol_t;

typedef enum {
  FM_CHART_LABEL_BEST_FIT = 0,
  FM_CHART_LABEL_CENTER = 1,
  FM_CHART_LABEL_INSIDE_BASE = 2,
  FM_CHART_LABEL_INSIDE_END = 3,
  FM_CHART_LABEL_OUTSIDE_END = 4,
  FM_CHART_LABEL_ABOVE = 5,
  FM_CHART_LABEL_BELOW = 6,
  FM_CHART_LABEL_LEFT = 7,
  FM_CHART_LABEL_RIGHT = 8,
} fm_chart_label_position_t;

/** @brief `c:barDir` of a bar plot. */
typedef enum {
  FM_CHART_BAR_DIR_ABSENT = 0,
  FM_CHART_BAR_DIR_COL = 1,
  FM_CHART_BAR_DIR_BAR = 2,
} fm_chart_bar_direction_t;

/** @brief A boolean element whose absence differs from either value. */
typedef enum {
  FM_CHART_TRI_ABSENT = 0,
  FM_CHART_TRI_FALSE = 1,
  FM_CHART_TRI_TRUE = 2,
} fm_chart_tri_t;

/** @brief Which input of a series a data source or resolved range belongs to. */
typedef enum {
  FM_CHART_ROLE_NAME = 0,
  FM_CHART_ROLE_CATEGORIES = 1,
  FM_CHART_ROLE_VALUES = 2,
  FM_CHART_ROLE_X = 3,
  FM_CHART_ROLE_Y = 4,
  FM_CHART_ROLE_SIZE = 5,
} fm_chart_role_t;

/** @brief Patch field operation; the zero value leaves the element alone. */
typedef enum {
  FM_CHART_PATCH_KEEP = 0,
  FM_CHART_PATCH_CLEAR = 1,
  FM_CHART_PATCH_SET = 2,
} fm_chart_patch_op_t;

/* ---- Chart model reads --------------------------------------------------- */

/** @brief Opaque snapshot of a chart's typed view. Release with `fm_chart_destroy`. */
typedef struct fm_chart fm_chart_t;

/**
 * @brief A chart or axis title. `kind` is an `fm_chart_title_kind_t`;
 *        `text` is the flattened rich text (`TEXT`) and `formula` the
 *        reference (`REF`), each empty otherwise. As input, `NULL` reads as
 *        empty.
 */
typedef struct {
  int32_t kind;
  const char* text;
  const char* formula;
} fm_chart_title_t;

/** @brief Solid fill; `transparency_pct` is NaN when the color has no alpha. */
typedef struct {
  int32_t has_color;
  uint32_t color;
  double transparency_pct;
} fm_chart_fill_t;

/**
 * @brief Outline. `width_pt` and `transparency_pct` are NaN when absent;
 *        `dash` (an `fm_chart_line_dash_t`) is meaningful when `has_dash`;
 *        `none` is non-zero for an explicit no-line.
 */
typedef struct {
  int32_t has_color;
  uint32_t color;
  double transparency_pct;
  double width_pt;
  int32_t has_dash;
  int32_t dash;
  int32_t none;
} fm_chart_line_t;

/** @brief Marker; `symbol` is an `fm_chart_marker_symbol_t`. */
typedef struct {
  int32_t symbol;
  int32_t has_size;
  uint32_t size;
  fm_chart_fill_t fill;
  fm_chart_line_t line;
} fm_chart_marker_t;

/**
 * @brief Data labels of a series. `separator` and `number_format` are
 *        `NULL` when absent; `position` (an `fm_chart_label_position_t`) is
 *        meaningful when `has_position`; `font_size_pt` is NaN when absent.
 */
typedef struct {
  int32_t show_value;
  int32_t show_category;
  int32_t show_series;
  int32_t show_percent;
  int32_t show_legend_key;
  const char* separator;
  int32_t has_position;
  int32_t position;
  const char* number_format;
  int32_t source_linked;
  int32_t has_font_color;
  uint32_t font_color;
  double font_size_pt;
  int32_t deleted;
} fm_chart_labels_t;

/**
 * @brief Top-level fields of a chart.
 *
 * The anchor fields read as `fm_drawing_object_t` does. `legend_position` is
 * `FM_CHART_LEGEND_NONE` when the chart has no legend. `auto_title_deleted`,
 * `plot_vis_only` and `na_as_blank` (`c16r3:dispNaAsBlank`, whether `#N/A`
 * points draw as blank cells) are `fm_chart_tri_t`; `blanks_as` an
 * `fm_chart_blanks_as_t`. `source_range` is the sheet-qualified rectangle
 * every series was split from, `NULL` (with the other `source_*` fields
 * zero) when the series share no such rectangle.
 */
typedef struct {
  uint32_t object_id;
  const char* name;
  const char* descr;
  int32_t is_chartex;
  int32_t is_pivot_chart;
  int32_t on_chartsheet;
  int32_t anchor_kind;
  int32_t edit_as;
  uint32_t from_row;
  uint32_t from_col;
  int64_t from_row_off;
  int64_t from_col_off;
  uint32_t to_row;
  uint32_t to_col;
  int64_t to_row_off;
  int64_t to_col_off;
  int64_t cx;
  int64_t cy;
  fm_chart_title_t title;
  int32_t auto_title_deleted;
  int32_t legend_position;
  int32_t legend_overlay;
  size_t legend_hidden_count;
  int32_t blanks_as;
  int32_t plot_vis_only;
  int32_t date1904;
  int32_t has_style;
  int32_t style;
  size_t plot_count;
  size_t axis_count;
  size_t hidden_category_count;
  const char* source_range;
  int32_t source_direction;
  int32_t source_header_row;
  int32_t source_header_column;
  int32_t na_as_blank;
} fm_chart_info_t;

/**
 * @brief One plot group. `type` is an `fm_chart_plot_type_t`, `bar_dir` an
 *        `fm_chart_bar_direction_t`, `vary_colors` an `fm_chart_tri_t`.
 *        `chartex_layout` is empty unless the plot is chartex.
 */
typedef struct {
  int32_t type;
  const char* chartex_layout;
  int32_t bar_dir;
  int32_t has_grouping;
  int32_t grouping;
  int32_t vary_colors;
  int32_t has_gap_width;
  uint32_t gap_width;
  int32_t has_overlap;
  int32_t overlap;
  int32_t has_hole_size;
  uint32_t hole_size;
  int32_t has_first_slice_angle;
  uint32_t first_slice_angle;
  size_t axis_id_count;
  size_t series_count;
} fm_chart_plot_t;

/** @brief Kind of a series input. */
typedef enum {
  FM_CHART_SOURCE_NONE = 0,
  FM_CHART_SOURCE_REF = 1,
  FM_CHART_SOURCE_LITERAL = 2,
} fm_chart_source_kind_t;

/**
 * @brief One series input. `kind` is an `fm_chart_source_kind_t`;
 *        `formula` is the reference text (empty unless `REF`);
 *        `literal_count` counts the points `fm_chart_get_series_literal_at`
 *        reads (`LITERAL` only).
 */
typedef struct {
  int32_t kind;
  const char* formula;
  size_t literal_count;
} fm_chart_data_source_t;

/**
 * @brief One series of a plot group. `smooth` is an `fm_chart_tri_t`;
 *        `marker` and `labels` are meaningful when their `has_*` flag is
 *        set. `point_count` counts the formatted points
 *        `fm_chart_get_series_point_at` reads.
 */
typedef struct {
  uint32_t idx;
  uint32_t order;
  int32_t filtered;
  fm_chart_data_source_t name;
  fm_chart_data_source_t categories;
  fm_chart_data_source_t values;
  fm_chart_data_source_t x;
  fm_chart_data_source_t y;
  fm_chart_data_source_t size;
  fm_chart_fill_t fill;
  fm_chart_line_t line;
  int32_t has_marker;
  fm_chart_marker_t marker;
  int32_t smooth;
  size_t point_count;
  int32_t has_labels;
  fm_chart_labels_t labels;
} fm_chart_series_t;

/** @brief A literal point: a number, or `text` when `is_text`. */
typedef struct {
  uint32_t idx;
  int32_t is_text;
  double number;
  const char* text;
} fm_chart_literal_t;

/** @brief One formatted data point (`c:dPt`). */
typedef struct {
  uint32_t idx;
  fm_chart_fill_t fill;
  fm_chart_line_t line;
  int32_t has_marker;
  fm_chart_marker_t marker;
} fm_chart_point_t;

/**
 * @brief One axis. `kind` is an `fm_chart_axis_kind_t`, `position` an
 *        `fm_chart_axis_position_t`, `crosses` an `fm_chart_axis_crosses_t`.
 *        `crosses_at`, `log_base`, `min`, `max`, `major_unit` and
 *        `minor_unit` are NaN when absent; `number_format` and
 *        `tick_label_pos` are `NULL` when absent.
 */
typedef struct {
  uint32_t id;
  int32_t kind;
  int32_t position;
  int32_t deleted;
  uint32_t cross_axis;
  int32_t crosses;
  double crosses_at;
  int32_t reversed;
  double log_base;
  double min;
  double max;
  double major_unit;
  double minor_unit;
  const char* number_format;
  int32_t source_linked;
  fm_chart_title_t title;
  int32_t major_gridlines;
  int32_t minor_gridlines;
  const char* tick_label_pos;
} fm_chart_axis_t;

/**
 * @brief Reads the chart `object_id` into an owned snapshot.
 *
 * `*out` is set to `NULL` before validation; on success the caller owns it
 * and releases it with `fm_chart_destroy`. Later workbook edits do not
 * change it.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet
 *         has no chart with that id;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse;
 *         `FM_STATUS_IO_CHART_UNPARSEABLE` (5025) when the chart part does not parse.
 */
FM_API fm_status_t fm_sheet_get_chart(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                      fm_chart_t** out);

/** @brief Releases a chart handle. `chart == NULL` is a no-op. */
FM_API void fm_chart_destroy(fm_chart_t* chart);

/**
 * @brief Reads the chart's top-level fields.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_chart_get_info(const fm_chart_t* chart, fm_chart_info_t* out);

/**
 * @brief Reads the `c:idx` of legend entry `idx` hidden from the legend
 *        (`idx < legend_hidden_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_chart_get_legend_hidden_at(const fm_chart_t* chart, size_t idx, uint32_t* out);

/**
 * @brief Reads the position of hidden category `idx`
 *        (`idx < hidden_category_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_chart_get_hidden_category_at(const fm_chart_t* chart, size_t idx, uint32_t* out);

/**
 * @brief Reads plot group `plot` (`plot < plot_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `plot` is out of range.
 */
FM_API fm_status_t fm_chart_get_plot_at(const fm_chart_t* chart, size_t plot, fm_chart_plot_t* out);

/**
 * @brief Reads axis id `idx` of plot group `plot` (`idx < axis_id_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `plot` or `idx` is out of range.
 */
FM_API fm_status_t fm_chart_get_plot_axis_id_at(const fm_chart_t* chart, size_t plot, size_t idx, uint32_t* out);

/**
 * @brief Reads series `series` of plot group `plot`, in document order
 *        (`series < series_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `plot` or `series` is out of range.
 */
FM_API fm_status_t fm_chart_get_series_at(const fm_chart_t* chart, size_t plot, size_t series, fm_chart_series_t* out);

/**
 * @brief Reads literal point `idx` of input `role` (an `fm_chart_role_t`)
 *        of a series (`idx < literal_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `plot`, `series`, `role` or `idx` is out
 *         of range.
 */
FM_API fm_status_t fm_chart_get_series_literal_at(const fm_chart_t* chart, size_t plot, size_t series, int32_t role,
                                                  size_t idx, fm_chart_literal_t* out);

/**
 * @brief Reads formatted point `idx` of a series (`idx < point_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `plot`, `series` or `idx` is out of range.
 */
FM_API fm_status_t fm_chart_get_series_point_at(const fm_chart_t* chart, size_t plot, size_t series, size_t idx,
                                                fm_chart_point_t* out);

/**
 * @brief Reads axis `idx` (`idx < axis_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_chart_get_axis_at(const fm_chart_t* chart, size_t idx, fm_chart_axis_t* out);

/* ---- Resolved chart data ------------------------------------------------- */

/** @brief Opaque resolved series data. Release with `fm_chart_data_destroy`. */
typedef struct fm_chart_data fm_chart_data_t;

/**
 * @brief Overrides `fm_sheet_get_chart_preview` applies to a copy of the
 *        chart. `source` is `NULL` to keep the chart's own range; each other
 *        field applies when its `_op` is `FM_CHART_PATCH_SET` (`CLEAR` is
 *        rejected). `direction` is an `fm_chart_series_direction_t`,
 *        `blanks_as` an `fm_chart_blanks_as_t`; the rest are 0/1.
 *        `na_as_blank` overrides whether `#N/A` points draw as blank cells.
 */
typedef struct {
  const char* source;
  int32_t direction_op;
  int32_t direction;
  int32_t header_row_op;
  int32_t header_row;
  int32_t header_column_op;
  int32_t header_column;
  int32_t blanks_as_op;
  int32_t blanks_as;
  int32_t plot_vis_only_op;
  int32_t plot_vis_only;
  int32_t na_as_blank_op;
  int32_t na_as_blank;
} fm_chart_preview_overrides_t;

/**
 * @brief Summary of resolved chart data.
 *
 * `is_xy` is non-zero for the XY domain (scatter, bubble), where points
 * carry an x value. `category_value_count` is non-zero only for numeric
 * (date) categories and then equals `category_label_count`.
 * `categories_generated` is set when no category reference exists and the
 * labels are 1..n. `title_text` is `NULL` when the chart has no resolved
 * title text.
 */
typedef struct {
  int32_t is_xy;
  size_t point_count;
  size_t series_count;
  size_t category_label_count;
  size_t category_value_count;
  size_t category_level_count;
  const char* category_format_code;
  int32_t categories_generated;
  const char* title_text;
  size_t axis_title_count;
  size_t unresolved_count;
} fm_chart_data_info_t;

/**
 * @brief One resolved series. `name` is `NULL` when the series has no name
 *        reference. `point_count` is the length of the arrays
 *        `fm_chart_data_get_series_points` returns. `range_count[role]` counts
 *        the resolved rectangles of each `fm_chart_role_t` input.
 *        `format_code` is the number format of the values (the y values in
 *        the XY domain); `x_format_code` is that of the x values, empty
 *        outside the XY domain.
 */
typedef struct {
  size_t plot_index;
  uint32_t idx;
  const char* name;
  const char* format_code;
  int32_t filtered;
  size_t point_count;
  int32_t has_sizes;
  size_t range_count[6];
  const char* x_format_code;
} fm_chart_data_series_t;

/**
 * @brief Parallel point arrays of one series, owned by the handle.
 *
 * `values` (the y value in the XY domain) and `states` (each an
 * `fm_chart_point_state_t`) have `count` entries. `x` is `NULL` outside the
 * XY domain; `sizes` is `NULL` unless the series is a bubble series.
 */
typedef struct {
  size_t count;
  const double* values;
  const uint8_t* states;
  const double* x;
  const double* sizes;
} fm_chart_data_points_t;

/** @brief A rectangle of cells on one sheet (0-based, inclusive). */
typedef struct {
  size_t sheet;
  uint32_t top;
  uint32_t left;
  uint32_t bottom;
  uint32_t right;
} fm_chart_sheet_range_t;

/** @brief Resolved title text of one axis. */
typedef struct {
  uint32_t axis_id;
  const char* text;
} fm_chart_axis_title_text_t;

/**
 * @brief An input that did not resolve; its points read as gaps. `source`
 *        is `name`, `cat`, `val`, `x`, `y`, `size`, `title` or
 *        `axis_title`; `series_idx` is the series' `c:idx` (0 for a title)
 *        and `axis_id` the axis of an `axis_title`.
 */
typedef struct {
  uint32_t series_idx;
  const char* source;
  uint32_t axis_id;
} fm_chart_unresolved_t;

/**
 * @brief Resolves the chart's series against the current cell values, by
 *        Excel's rules (hidden cells, blanks, errors, filters).
 *
 * `*out` is set to `NULL` before validation; on success the caller owns it
 * and releases it with `fm_chart_data_destroy`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet
 *         has no chart with that id;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse;
 *         `FM_STATUS_IO_CHART_UNPARSEABLE` (5025) when the chart part does not parse.
 */
FM_API fm_status_t fm_sheet_get_chart_data(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                           fm_chart_data_t** out);

/**
 * @brief `fm_sheet_get_chart_data` for a copy of the chart with
 *        `overrides` applied; the workbook is not changed.
 *
 * A source, direction or header override re-splits the series as
 * `fm_sheet_set_chart_source` does: over the overridden range (an unset
 * direction or header is auto), else over the chart's derived range (unset
 * ones as derived).
 *
 * @return as `fm_sheet_get_chart_data`, and
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `overrides` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a `CLEAR` op or an out-of-domain value;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when `source` does not resolve;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a chartex chart.
 */
FM_API fm_status_t fm_sheet_get_chart_preview(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                              const fm_chart_preview_overrides_t* overrides, fm_chart_data_t** out);

/** @brief Releases a chart data handle. `data == NULL` is a no-op. */
FM_API void fm_chart_data_destroy(fm_chart_data_t* data);

/**
 * @brief Reads the data summary.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_chart_data_get_info(const fm_chart_data_t* data, fm_chart_data_info_t* out);

/**
 * @brief Reads category label `idx` (`idx < category_label_count`) and, for
 *        numeric categories, its value (NaN otherwise).
 *
 * `*out_label` is an owned snapshot held by `data`, valid until
 * `fm_chart_data_destroy`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_category_at(const fm_chart_data_t* data, size_t idx, const char** out_label,
                                                 double* out_value);

/**
 * @brief Reads the label count of upper category level `level`
 *        (`level < category_level_count`, outermost first).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `level` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_category_level_size(const fm_chart_data_t* data, size_t level, size_t* out_count);

/**
 * @brief Reads label `idx` of upper category level `level`.
 *
 * `*out_label` is an owned snapshot held by `data`, valid until
 * `fm_chart_data_destroy`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `level` or `idx` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_category_level_at(const fm_chart_data_t* data, size_t level, size_t idx,
                                                       const char** out_label);

/**
 * @brief Reads resolved series `series` (`series < series_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `series` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_series_at(const fm_chart_data_t* data, size_t series, fm_chart_data_series_t* out);

/**
 * @brief Returns the point arrays of series `series`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `series` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_series_points(const fm_chart_data_t* data, size_t series,
                                                   fm_chart_data_points_t* out);

/**
 * @brief Reads resolved rectangle `idx` of input `role` (an
 *        `fm_chart_role_t`) of series `series` (`idx < range_count[role]`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `series`, `role` or `idx` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_range_at(const fm_chart_data_t* data, size_t series, int32_t role, size_t idx,
                                              fm_chart_sheet_range_t* out);

/**
 * @brief Reads axis title `idx` (`idx < axis_title_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_axis_title_at(const fm_chart_data_t* data, size_t idx,
                                                   fm_chart_axis_title_text_t* out);

/**
 * @brief Reads unresolved input `idx` (`idx < unresolved_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `idx` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_unresolved_at(const fm_chart_data_t* data, size_t idx, fm_chart_unresolved_t* out);

/* ---- Creation and edits -------------------------------------------------- */

/**
 * @brief Options of `fm_sheet_add_chart`.
 *
 * `kind` is an `fm_chart_insert_kind_t`; `grouping` (an
 * `fm_chart_grouping_t`) applies to column, bar, line and area. `source` is
 * one A1 rectangle, sheet-qualified or on the target sheet. `direction` is
 * an `fm_chart_series_direction_t`, `header_row`/`header_column`
 * `fm_chart_header_mode_t`s (`AUTO` follows Excel's detection). `title` is
 * `NULL` for Excel's automatic title, empty for none, else the text.
 * `legend` is an `fm_chart_legend_position_t`. `placement` reads as for
 * `fm_sheet_add_image`, its `name` and `descr` naming the chart (an
 * empty name becomes `Chart <id>`), except that a zero width or height
 * means Excel's default 360 x 216 pt frame. Excel inserts charts with a
 * two-cell anchor.
 */
typedef struct {
  int32_t kind;
  int32_t grouping;
  const char* source;
  int32_t direction;
  int32_t header_row;
  int32_t header_column;
  const char* title;
  int32_t legend;
  fm_image_insert_t placement;
} fm_chart_insert_t;

/** @brief Solid fill patch; `transparency_pct` needs a solid color. */
typedef struct {
  int32_t color_op;
  uint32_t color;
  int32_t transparency_op;
  double transparency_pct;
} fm_chart_fill_patch_t;

/** @brief Outline patch; `dash` is an `fm_chart_line_dash_t`, `none` 0/1 (`CLEAR` rejected). */
typedef struct {
  int32_t color_op;
  uint32_t color;
  int32_t transparency_op;
  double transparency_pct;
  int32_t width_op;
  double width_pt;
  int32_t dash_op;
  int32_t dash;
  int32_t none_op;
  int32_t none;
} fm_chart_line_patch_t;

/** @brief Marker patch; `symbol` is an `fm_chart_marker_symbol_t` other than `FM_CHART_MARKER_PICTURE` (`CLEAR`
 * rejected). */
typedef struct {
  int32_t symbol_op;
  int32_t symbol;
  int32_t size_op;
  double size;
  fm_chart_fill_patch_t fill;
  fm_chart_line_patch_t line;
} fm_chart_marker_patch_t;

/**
 * @brief Data-label patch. The `show_*` flags are 0/1 (`CLEAR` rejected);
 *        `position` is an `fm_chart_label_position_t`.
 */
typedef struct {
  int32_t show_value_op;
  int32_t show_value;
  int32_t show_category_op;
  int32_t show_category;
  int32_t show_series_op;
  int32_t show_series;
  int32_t show_percent_op;
  int32_t show_percent;
  int32_t show_legend_key_op;
  int32_t show_legend_key;
  int32_t separator_op;
  const char* separator;
  int32_t position_op;
  int32_t position;
  int32_t number_format_op;
  const char* number_format;
  int32_t font_color_op;
  uint32_t font_color;
  int32_t font_size_op;
  double font_size_pt;
} fm_chart_labels_patch_t;

/**
 * @brief Series patch.
 *
 * Each reference is `NULL` to keep it; otherwise the new reference text, an
 * empty string removing `name`, `categories` or `x`. An empty `values`, `y`
 * or `size` is `FM_STATUS_INVALID_ARGUMENT` (2). `marker_op`/`labels_op` `CLEAR` removes
 * the marker or the data labels; `smooth` is 0/1 (`CLEAR` rejected).
 */
typedef struct {
  const char* name;
  const char* categories;
  const char* values;
  const char* x;
  const char* y;
  const char* size;
  fm_chart_fill_patch_t fill;
  fm_chart_line_patch_t line;
  int32_t marker_op;
  fm_chart_marker_patch_t marker;
  int32_t smooth_op;
  int32_t smooth;
  int32_t labels_op;
  fm_chart_labels_patch_t labels;
} fm_chart_series_patch_t;

/** @brief Data-point patch; a patch that changes nothing removes the point's formatting. */
typedef struct {
  fm_chart_fill_patch_t fill;
  fm_chart_line_patch_t line;
  int32_t marker_op;
  fm_chart_marker_patch_t marker;
} fm_chart_point_patch_t;

/**
 * @brief Plot-group patch (`CLEAR` rejected on every field). A field the
 *        plot type does not have is `FM_STATUS_INVALID_ARGUMENT` (2); `hole_size` is
 *        10..90 and doughnut only.
 */
typedef struct {
  int32_t gap_width_op;
  uint32_t gap_width;
  int32_t overlap_op;
  int32_t overlap;
  int32_t first_slice_angle_op;
  uint32_t first_slice_angle;
  int32_t vary_colors_op;
  int32_t vary_colors;
  int32_t hole_size_op;
  uint32_t hole_size;
} fm_chart_plot_patch_t;

/**
 * @brief Axis patch.
 *
 * `visible`, `major_gridlines` and `reversed` are 0/1 and `crosses` an
 * `fm_chart_axis_crosses_t` (`CLEAR` rejected on these four); `crosses_at`
 * is read when `crosses` is `FM_CHART_CROSSES_AT`. `title` takes the
 * kinds of `fm_sheet_set_chart_title` (`NONE` removes the axis title, `AUTO`
 * leaves the text to the viewer). `CLEAR` on `title` removes the axis
 * title, on `log_base` makes the axis linear, and on `tick_label_pos`
 * (`nextTo`, `low`, `high`, `none`) restores the default.
 */
typedef struct {
  int32_t visible_op;
  int32_t visible;
  int32_t title_op;
  fm_chart_title_t title;
  int32_t min_op;
  double min;
  int32_t max_op;
  double max;
  int32_t major_unit_op;
  double major_unit;
  int32_t number_format_op;
  const char* number_format;
  int32_t major_gridlines_op;
  int32_t major_gridlines;
  int32_t reversed_op;
  int32_t reversed;
  int32_t log_base_op;
  double log_base;
  int32_t crosses_op;
  int32_t crosses;
  double crosses_at;
  int32_t tick_label_pos_op;
  const char* tick_label_pos;
} fm_chart_axis_patch_t;

/**
 * @brief Legend state. `position` is an `fm_chart_legend_position_t`;
 *        `FM_CHART_LEGEND_NONE` removes the legend. With `hidden_series_op`
 *        `SET` the `hidden_series_count` series `c:idx` values replace the
 *        hidden legend entries (`hidden_series` may be `NULL` when the count
 *        is 0); `KEEP` leaves them alone.
 */
typedef struct {
  int32_t position;
  int32_t overlay;
  int32_t hidden_series_op;
  const uint32_t* hidden_series;
  size_t hidden_series_count;
} fm_chart_legend_patch_t;

/**
 * @brief Display-option patch. `blanks_as` (`c:dispBlanksAs`) is an
 *        `fm_chart_blanks_as_t` of `GAP`, `ZERO` or `SPAN`; `plot_vis_only`
 *        (`c:plotVisOnly`) and `na_as_blank` (`c16r3:dispNaAsBlank`, `#N/A`
 *        points drawn as blank cells) are 0/1. `CLEAR` removes the element;
 *        `na_as_blank` set to 0 removes it too, as Excel writes it.
 */
typedef struct {
  int32_t blanks_as_op;
  int32_t blanks_as;
  int32_t plot_vis_only_op;
  int32_t plot_vis_only;
  int32_t na_as_blank_op;
  int32_t na_as_blank;
} fm_chart_display_patch_t;

/**
 * @brief Creates a chart from `opts` with filled caches and returns its
 *        object id.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `opts`, `opts->source` or
 *         `out_object_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a bad enum value or placement (as
 *         `fm_sheet_add_image`), or a source that does not yield a series;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when `source` does not resolve to
 *         one rectangle;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or was retained from an `.xlsb` package. The workbook is then
 *         unchanged.
 */
FM_API fm_status_t fm_sheet_add_chart(fm_workbook_t* wb, size_t sheet_index, const fm_chart_insert_t* opts,
                                      uint32_t* out_object_id);

/**
 * @brief Replaces plot group `plot` with `kind` (an
 *        `fm_chart_insert_kind_t`) and `grouping` (an `fm_chart_grouping_t`),
 *        keeping each series' references, idx, order and explicit colors.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a pivot chart, or a plot that
 *         shares its axes with another plot group.
 */
FM_API fm_status_t fm_sheet_set_chart_type(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                           int32_t kind, int32_t grouping);

/**
 * @brief Sets the chart title: `NONE` removes it, `TEXT` writes `text`,
 *        `REF` writes the reference `formula`, `AUTO` leaves the text to
 *        the viewer.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `title` is `NULL`;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when `formula` does not resolve.
 */
FM_API fm_status_t fm_sheet_set_chart_title(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                            const fm_chart_title_t* title);

/**
 * @brief Sets the legend position, overlay and hidden entries.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `legend` is `NULL`, or
 *         `hidden_series` is `NULL` with a non-zero count.
 */
FM_API fm_status_t fm_sheet_set_chart_legend(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                             const fm_chart_legend_patch_t* legend);

/**
 * @brief Patches axis `axis_id`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`.
 */
FM_API fm_status_t fm_sheet_set_chart_axis(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, uint32_t axis_id,
                                           const fm_chart_axis_patch_t* patch);

/**
 * @brief Patches the chart's display options.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a `SET` value is outside its domain
 *         (`blanks_as` `ABSENT` included; `CLEAR` removes the element).
 */
FM_API fm_status_t fm_sheet_set_chart_display(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                              const fm_chart_display_patch_t* patch);

/**
 * @brief Patches the series with `c:idx` `series_idx` of plot group `plot`.
 *        A changed reference refreshes its cache.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when a reference does not
 *         resolve;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a reference edit on a pivot
 *         chart.
 */
FM_API fm_status_t fm_sheet_set_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                             size_t series_idx, const fm_chart_series_patch_t* patch);

/**
 * @brief Patches point `point_idx` of series `series_idx`; a patch that
 *        changes nothing removes the point's formatting.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`.
 */
FM_API fm_status_t fm_sheet_set_chart_point(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                            size_t series_idx, size_t point_idx, const fm_chart_point_patch_t* patch);

/**
 * @brief Patches plot group `plot`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`.
 */
FM_API fm_status_t fm_sheet_set_chart_plot(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                           const fm_chart_plot_patch_t* patch);

/**
 * @brief Replaces the chart filter: the listed series (by `c:idx`) and
 *        categories (positions in the full category range) are hidden.
 *        Either array may be `NULL` when its count is 0.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when an array is `NULL` with a non-zero count;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a pivot chart.
 */
FM_API fm_status_t fm_sheet_set_chart_filter(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                             const uint32_t* hidden_series, size_t hidden_series_count,
                                             const uint32_t* hidden_categories, size_t hidden_category_count);

/**
 * @brief Appends a series to plot group `plot` from `patch` (which must set
 *        `values`, or `y` for a scatter plot) and returns its `c:idx`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` or `out_series_idx` is `NULL`;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when a reference does not
 *         resolve;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a pivot chart.
 */
FM_API fm_status_t fm_sheet_add_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                             const fm_chart_series_patch_t* patch, size_t* out_series_idx);

/**
 * @brief Removes the series with `c:idx` `series_idx` from plot group `plot`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a pivot chart.
 */
FM_API fm_status_t fm_sheet_remove_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                                size_t series_idx);

/**
 * @brief Rebuilds every series from `source` by the insert splitting rules,
 *        keeping the chart kind and each series' formatting by idx, and
 *        clears the filter. `direction` is an `fm_chart_series_direction_t`,
 *        the header modes `fm_chart_header_mode_t`s.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `source` is `NULL`;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when `source` does not resolve to
 *         one rectangle;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a pivot chart.
 */
FM_API fm_status_t fm_sheet_set_chart_source(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                             const char* source, int32_t direction, int32_t header_row,
                                             int32_t header_column);

/* -------------------------------------------------------------------------- */
/* What-If data tables                                                        */
/* -------------------------------------------------------------------------- */

/**
 * @brief A What-If data table (`TABLE()`): the record `fm_sheet_get_data_table`
 *        reads and `fm_sheet_add_data_table` takes.
 *
 * `first_row`..`last_col` is the result range (0-based, inclusive); the
 * formulas and substitution values sit in the row above and the column left
 * of it. The row input cell receives the values of the top row, the column
 * input cell those of the left column: a one-variable table has one of them
 * (`row_oriented` says which, on read), a two-variable table both
 * (`two_dimensional`). A flag is `0` or `1`.
 *
 * On add only the range and the `has_*` flags and inputs are read; the
 * result cells are set to 0 and filled by the next recalc. On read
 * `row_input_deleted` / `col_input_deleted` mark an input cell removed by a
 * row or column delete; the table then evaluates to `#REF!`. An input that is
 * absent reads as row and column 0.
 */
typedef struct {
  uint32_t first_row;
  uint32_t first_col;
  uint32_t last_row;
  uint32_t last_col;
  uint32_t row_input_row;
  uint32_t row_input_col;
  uint32_t col_input_row;
  uint32_t col_input_col;
  int32_t has_row_input;
  int32_t has_col_input;
  int32_t two_dimensional;
  int32_t row_oriented;
  int32_t row_input_deleted;
  int32_t col_input_deleted;
} fm_data_table_t;

/**
 * @brief Creates a data table over the result range of `spec`. Cells in the
 *        range may no longer be written, except by removing the table.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `spec == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a
 *         coordinate is outside the grid;
 *         `FM_STATUS_DATA_TABLE_INVALID` (8005) when neither input is given, the range
 *         has no row above or column left of it, an input lies inside the
 *         range, or the range overlaps another data table.
 */
FM_API fm_status_t fm_sheet_add_data_table(fm_workbook_t* wb, size_t sheet_index, const fm_data_table_t* spec);

/**
 * @brief Reads the data table whose result range contains `(row, col)`.
 *        `*out_present` is `1` and `*out` filled when there is one, else `0`
 *        and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_present` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_data_table(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                           int32_t* out_present, fm_data_table_t* out);

/**
 * @brief Removes the data table whose result range contains `(row, col)` and
 *        blanks its result cells.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (5050) when `sheet_index` is out of range;
 *         `FM_STATUS_DATA_TABLE_INVALID` (8005) when no data table covers the cell.
 */
FM_API fm_status_t fm_sheet_remove_data_table(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col);

/**
 * @brief Recomputes every data table now, whatever the calc mode. A recalc
 *        does this itself after a pass that had a dirty cell or a volatile
 *        formula, except under
 *        `FM_CALC_MODE_AUTO_NO_TABLE`, which leaves it to this call.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         a `kEval*` / `kGraph*` code on engine failure.
 */
FM_API fm_status_t fm_workbook_recalc_data_tables(fm_workbook_t* wb);

/* -------------------------------------------------------------------------- */
/* Version                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the Formulon library version as a NUL-terminated
 *        UTF-8 string. Always non-NULL with program lifetime.
 */
FM_API const char* fm_get_version(void);

#ifdef __cplusplus
} /* extern "C" */
#endif

#endif /* FORMULON_C_API_H */
