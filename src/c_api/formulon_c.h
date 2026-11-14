/*
 * Formulon stable C ABI.
 *
 * This is a pure C11 header: every declaration sits inside `extern "C"`
 * when included from C++, no `bool` / `nullptr` / templates leak across
 * the boundary (the C++-only `FM_STRUCT_INIT` helper is inline and exports
 * nothing), and every type is either an opaque struct, a primitive
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
 * Quick start:
 *
 *     fm_workbook_t* wb = NULL;
 *     if (fm_workbook_create_default(&wb) != FM_STATUS_OK) { ... }
 *     fm_sheet_set_number(wb, 0, 0, 0, 42.0);                 // Sheet1!A1
 *     fm_sheet_set_formula(wb, 0, 1, 0, "=A1*2");             // Sheet1!A2
 *     fm_workbook_recalc(wb);
 *     fm_value_t v;
 *     fm_sheet_get_value(wb, 0, 1, 0, &v);                    // v.u.number == 84
 *     fm_workbook_destroy(wb);
 *
 * Every fallible function returns an `fm_status_t`; results come back through
 * `out` parameters. Sheets are addressed by 0-based index, cells by 0-based
 * `(row, col)`.
 *
 * Status codes mirror `formulon::FormulonErrorCode` (see
 * `src/utils/error.h`) and are named `FM_STATUS_*` in `formulon_status.h`; a
 * return value of `0` (`FM_STATUS_OK`) always denotes success.
 *
 * Status rules every entry point follows; an entry's `@return` restates the
 * ones it can hit, or points at its section's "Common failures" block:
 *   * An index that addresses an element of a counted collection and is not
 *     below the count returns `FM_STATUS_INDEX_OUT_OF_RANGE` (13); a sheet
 *     index returns `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) instead.
 *   * A value or shape outside its domain returns `FM_STATUS_INVALID_ARGUMENT`
 *     (2): an enum ordinal outside its enumeration, a number outside a
 *     documented limit, an empty name, an inverted rectangle. A cell
 *     coordinate, row index or column index outside the sheet grid returns it
 *     on every entry point that takes one.
 *   * A key (name, id or cell) that matches nothing returns
 *     `FM_STATUS_NOT_FOUND` (6) on a remove, a mutation and a lookup that
 *     acquires an element by id; a keyed read succeeds instead and reports
 *     the absence (below).
 *   * A name or id another element already holds returns
 *     `FM_STATUS_ALREADY_EXISTS` (7).
 *   * A refusal caused by the workbook's state or a dependent object -- a
 *     hidden sheet, the last sheet or last visible sheet, an object still in
 *     use, an edit that cuts into a data table, array block or PivotTable -- returns
 *     `FM_STATUS_PRECONDITION_FAILED` (9); the same arguments succeed in
 *     another state.
 *   * A `NULL` required pointer returns `FM_STATUS_BINDING_NULL_POINTER`
 *     (7001); malformed UTF-8 returns `FM_STATUS_BINDING_UTF8ENCODING_ERROR`
 *     (7002) (Text input, below).
 *
 * Absence: a read keyed by a name, id or cell takes `int32_t* out_present`
 * before its result and succeeds with `*out_present == 0` when nothing
 * matches; an optional string reads `NULL`, with `out_present` beside it
 * when the string is the function's only result; an optional number inside a
 * record has an `int32_t has_<field>` flag before it; an optional enumeration
 * reads its `FM_<TYPE>_ABSENT` member (-1). No field uses NaN to mean absent.
 *
 * Patches: `update_<x>` takes a `_patch_t` record. Each patchable field
 * `<field>` has an `fm_patch_op_t <field>_op` before it: `FM_PATCH_OP_KEEP`
 * (0) leaves the field, `FM_PATCH_OP_SET` (2) writes the value beside the op,
 * and `FM_PATCH_OP_CLEAR` (1) removes the value, on the fields whose doc says
 * they can be cleared; `CLEAR` on any other field, an op outside 0..2, and
 * `SET` with a `NULL` string return `FM_STATUS_INVALID_ARGUMENT` (2). An array
 * field is `const T* <field>`, `<field>_count` and `<field>_op`, and `SET`
 * replaces the whole array. An embedded sub-patch carries its own ops and has
 * none in the parent. The whole patch is checked before anything is written,
 * so a failed update changes nothing, and a patch whose ops are all `KEEP`
 * (`FM_STRUCT_INIT`) is a successful no-op.
 *
 * Verbs: `remove` deletes one element and returns `FM_STATUS_NOT_FOUND` when
 * it is absent; keyed by a range (`fm_sheet_remove_merge`,
 * `fm_sheet_remove_hyperlink`) it deletes every element that intersects the
 * range. `clear` empties a collection and is a no-op when it is
 * already empty; `reset` returns one property to its default. A singleton
 * part (a sheet's protection or AutoFilter, a chart legend, the pinned clock)
 * is replaced by `set_<x>` with a record and removed by `set_<x>` with
 * `NULL`, which is a no-op when the part is absent; `update_<x>` merges a
 * `_patch_t` into it (Patches, above). A per-key singleton (a page break at a row or
 * column) takes the key beside the record: `NULL` removes the element at
 * that key, a no-op when there is none.
 *
 * Last error: `fm_get_last_error_message()` and `fm_get_last_error_context()`
 * describe the most recent failing call on the calling thread. Both are
 * thread-local views, valid until the next API call on that thread; copy them
 * out to keep the diagnostic. A call that succeeds resets the message to the
 * empty string. A formula syntax failure also records its position, read with
 * `fm_get_last_parse_error()`; every other call clears that record. The
 * functions that read or release error state leave it untouched:
 * `fm_get_last_error_message`, `fm_get_last_error_context`,
 * `fm_get_last_parse_error`, `fm_get_status_name`, `fm_get_version`, `fm_get_version_number`,
 * `fm_get_error_display_name` and every `*_destroy`.
 *
 * Message text is for display only: it is an English sentence meant for a
 * person, not a contract. Branch on the status code; read the context for the
 * offending argument values (`key=value` pairs). An argument error's message
 * leads with the parameter or field (`zoom_pct must be in 10..400`).
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
 *   * A value read can calculate: under `FM_CALC_MODE_AUTO` /
 *     `FM_CALC_MODE_AUTO_NO_TABLE` it first recalculates dirty cells, so the
 *     owner-thread contract covers reads exactly as it covers mutations. Such a
 *     read can invoke the caller's progress callback and can return any
 *     `FM_STATUS_GRAPH_*` recalc status (4000-4006); called from inside a
 *     progress callback it returns `FM_STATUS_GRAPH_RECALC_REENTRANT` (4005).
 *     A failed calculation is returned as is and leaves the cells it did not
 *     finish dirty, so the next read tries again.
 *   * The last-error storage is thread-local, so two threads driving
 *     two distinct workbooks can read their own diagnostics
 *     independently.
 *   * `fm_get_status_name()`, `fm_get_version()` and
 *     `fm_get_version_number()` are pure functions and are safe to call
 *     concurrently.
 *
 * Pointer lifetimes: every pointer a function hands back belongs to one of
 * five classes, and each API names its class. They are the model-backed view,
 * the read-scratch-backed view, the owned snapshot, the static view and the
 * thread-local view:
 *   * A **model-backed view** points into workbook-owned model storage. Its
 *     lifetime ends at the invalidating mutation named by that API, or when
 *     `fm_workbook_destroy` releases the handle.
 *   * A **read-scratch-backed view** points into transient per-handle read
 *     storage. The next successful scratch-backed read on the same handle
 *     may invalidate it; a validation-rejected call does not refresh this
 *     storage. Mutations and handle destruction also invalidate it. Callers
 *     must copy strings or arrays before making another call when they need
 *     to retain the value.
 *   * An **owned snapshot** is returned through a separately owned opaque
 *     result handle and stays valid until that result is destroyed.
 *   * A **static view** points at process-static storage and has program
 *     lifetime.
 *   * A **thread-local view** points into storage owned by the calling
 *     thread and stays valid until the next API call on that thread.
 *
 * Memory model:
 *   * Workbook-backed views are never allocated for the caller and must not
 *     be freed. Their exact invalidation boundary is part of each API's
 *     contract below; do not infer it from the C pointer type alone.
 *   * Buffers returned through `fm_workbook_save`,
 *     `fm_workbook_save_with_diagnostics`, `fm_sheet_snapshot_drawing` and
 *     `fm_sheet_snapshot_range_xfs` are heap-allocated by the library with
 *     `new[]` and must be released by passing the same pointer to
 *     `fm_buffer_destroy` (a NULL pointer is a no-op). Mixing `malloc`/`free` and
 *     `new[]`/`delete[]` across the boundary is undefined.
 *   * Every API that returns an owned opaque handle through an `out`
 *     parameter (including workbook, conditional-format-result, and
 *     cell-node handles) initializes `*out` to `NULL` before validation and
 *     leaves it `NULL` on every failure path.
 *
 * Versioning and stability: `FM_VERSION_MAJOR`, `FM_VERSION_MINOR` and
 * `FM_VERSION_PATCH` give the library version at compile time, and
 * `FM_VERSION_NUMBER` packs them into one integer;
 * `fm_get_version()` and `fm_get_version_number()` return it at run time;
 * `FM_CHECK_VERSION()` compares the two. From 1.0 on, no function is removed
 * or changed in meaning; otherwise new data arrives through new functions.
 *
 * Records and value types: a struct whose first member is
 * `uint32_t struct_size` is a record. From 1.0 on, a record may gain fields
 * only at its end, and `struct_size` tells them apart. The caller sets
 * `struct_size` to `sizeof` the record it was compiled against
 * (`FM_STRUCT_INIT(T)`), on input and output records alike; a value below
 * this library's first layout of the record is `FM_STATUS_INVALID_ARGUMENT`
 * (2), a larger one is read and written only up to the fields this library
 * knows, and the library never rewrites an output's `struct_size`. Every
 * other struct is a value type (coordinates, values, colours, array elements
 * and structs embedded by value) and is frozen: it never gains, loses or
 * reorders a field. A record's suffix names its direction: `_spec_t` is only
 * taken, `_info_t` only filled, `_patch_t` is a partial update, and a plain
 * `_t` record is both read and written back.
 */

#ifndef FORMULON_C_API_H
#define FORMULON_C_API_H

#include <stddef.h>
#include <stdint.h>

#include "formulon_status.h"

/** @brief Library version at compile time; equals `fm_get_version()`. */
#define FM_VERSION_MAJOR 0
#define FM_VERSION_MINOR 19
#define FM_VERSION_PATCH 0

/** @brief Numeric library version, `MAJOR * 10000 + MINOR * 100 + PATCH`; equals `fm_get_version_number()`. */
#define FM_VERSION_NUMBER (FM_VERSION_MAJOR * 10000 + FM_VERSION_MINOR * 100 + FM_VERSION_PATCH)

/**
 * @brief Initialiser for a record: `struct_size` set to `sizeof(T)`, every
 *        other field zero. `fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);`
 *
 * C uses a designated initialiser and C++ a value-initialising helper, so
 * neither trips `-Wmissing-field-initializers`.
 */
#ifdef __cplusplus
template <typename T>
inline T fm_struct_init() {
  T record{};
  record.struct_size = static_cast<decltype(record.struct_size)>(sizeof(T));
  return record;
}
#define FM_STRUCT_INIT(T) (fm_struct_init<T>())
#else
#define FM_STRUCT_INIT(T) {.struct_size = sizeof(T)}
#endif

/**
 * @brief Nonzero when the linked library has the same major and minor version
 *        as this header.
 */
#define FM_CHECK_VERSION() (fm_get_version_number() / 100 == FM_VERSION_NUMBER / 100)

#ifdef __cplusplus
extern "C" {
#endif

/*
 * Public symbol visibility.
 *
 * The core library is built with `-fvisibility=hidden`, so every
 * exported entry point must opt back in via `__attribute__((visibility
 * ("default")))`. On Windows the macro is `__declspec(dllexport)` while the
 * library itself is compiled (`FORMULON_BUILDING_DLL`), `__declspec(dllimport)`
 * in a consumer of the shared library (`FORMULON_SHARED`) and empty otherwise.
 * Consumers that link the static `formulon_static` archive define neither.
 */
#if defined(_WIN32)
#if defined(FORMULON_BUILDING_DLL)
#define FM_API __declspec(dllexport)
#elif defined(FORMULON_SHARED)
#define FM_API __declspec(dllimport)
#else
#define FM_API
#endif
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

/**
 * @brief A boolean whose zero value means "Excel's default".
 *
 * Every optional boolean of a `_spec_t` input, and every boolean of a
 * read-and-write record whose Excel default is true, takes this type (held in
 * an `int32_t` field), so a record zeroed by `FM_STRUCT_INIT` asks for Excel's
 * defaults. A read of such a record field returns `FM_TRISTATE_FALSE` or
 * `FM_TRISTATE_TRUE`, never `FM_TRISTATE_DEFAULT`. A chart field documented as
 * an `fm_tristate_t` reads `FM_TRISTATE_DEFAULT` when its element is absent.
 * A value outside the enumeration is `FM_STATUS_INVALID_ARGUMENT` (2).
 */
typedef enum {
  FM_TRISTATE_DEFAULT = 0,
  FM_TRISTATE_FALSE = 1,
  FM_TRISTATE_TRUE = 2,
} fm_tristate_t;

/**
 * @brief Operation of one `_patch_t` field (`<field>_op`); the zero value
 *        leaves the field alone. The rules are under "Patches" above.
 */
typedef enum {
  FM_PATCH_OP_KEEP = 0,
  FM_PATCH_OP_CLEAR = 1,
  FM_PATCH_OP_SET = 2,
} fm_patch_op_t;

/**
 * @brief Cell rectangle: 0-based, inclusive corners `(first_row, first_col)`
 *        through `(last_row, last_col)`.
 *
 * Every rectangle the API takes or returns uses this type. A single cell has
 * `first_* == last_*`.
 */
typedef struct {
  /** The 0-based first row. */
  uint32_t first_row;
  /** The 0-based first column. */
  uint32_t first_col;
  /** The 0-based last row, inclusive. */
  uint32_t last_row;
  /** The 0-based last column, inclusive. */
  uint32_t last_col;
} fm_cell_rect_t;

/** @brief A cell address: 0-based `row` and `col`. */
typedef struct {
  /** The 0-based row. */
  uint32_t row;
  /** The 0-based column. */
  uint32_t col;
} fm_cell_address_t;

/** @brief A cell rectangle on sheet `sheet_index`. */
typedef struct {
  /** The 0-based index of the sheet. */
  uint32_t sheet_index;
  /** The cell rectangle on that sheet. */
  fm_cell_rect_t rect;
} fm_sheet_rect_t;

/** @brief Discriminator for `fm_color_t::kind`, mirroring `formulon::ColorSpec::Kind`. */
typedef enum {
  FM_COLOR_KIND_NONE = 0,    /**< No colour. */
  FM_COLOR_KIND_RGB = 1,     /**< `rgb="AARRGGBB"`. */
  FM_COLOR_KIND_THEME = 2,   /**< `theme="N"` with optional `tint`. */
  FM_COLOR_KIND_INDEXED = 3, /**< `indexed="N"` legacy palette index. */
  FM_COLOR_KIND_AUTO = 4,    /**< `auto="1"` system foreground / background. */
} fm_color_kind_t;

/**
 * @brief A colour as the workbook specifies it; every unresolved colour
 *        field of the API takes this type; a resolved colour (for example in
 *        `fm_effective_style_info_t`) is a plain `0xAARRGGBB` integer.
 *
 * A theme, indexed or auto colour is a selector, not a rendered colour: resolve it
 * with `fm_workbook_resolve_color`. A field documented as "None or Rgb"
 * (CF evaluation colours) never carries a selector, and its `argb` is the
 * colour as drawn. A chart or shape colour is "None, Rgb or Theme": Theme
 * (tint quantised to 1e-5) reads for a palette scheme colour with only a
 * luminance tint, Rgb as drawn for any other colour.
 */
typedef struct {
  /** The theme tint from -1 (darker) to 1 (lighter); meaningful when `kind` is theme. */
  double tint;
  /** The colour as 0xAARRGGBB; meaningful when `kind` is RGB. */
  uint32_t argb;
  /** The theme colour index; meaningful when `kind` is theme. */
  uint32_t theme;
  /** The legacy palette index; meaningful when `kind` is indexed. */
  uint32_t indexed;
  /** Which selector the colour carries; only the matching field below is meaningful. */
  int32_t kind;
} fm_color_t;

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
  FM_VALUE_KIND_BLANK = 0,
  FM_VALUE_KIND_NUMBER = 1,
  FM_VALUE_KIND_BOOL = 2,
  FM_VALUE_KIND_TEXT = 3,
  FM_VALUE_KIND_ERROR = 4,
  FM_VALUE_KIND_ARRAY = 5,
  FM_VALUE_KIND_REF = 6,
  FM_VALUE_KIND_LAMBDA = 7
} fm_value_kind_t;

/**
 * @brief Cell value POD shared across the C ABI boundary.
 *
 * The active union member is selected by `kind`. A blank value has no payload, an array value surfaces
 * only its kind, and reference and lambda kinds are reserved.
 */
typedef struct {
  /** The `ValueKind` of the value; it selects which payload field is populated. */
  int32_t kind;
  /** The payload, a union whose active member is selected by `kind`; blank, array, reference and lambda values carry
   *  none. */
  union {
    /** The IEEE-754 double of a number value. */
    double number;
    /** 0 (FALSE) or 1 (TRUE) of a bool value. */
    int32_t boolean;
    /** The Excel `ErrorCode` ordinal of an error value (Null=0, Div0=1, Value=2, ...); the enumeration is
     *  `formulon::ErrorCode` in `src/value.h`. */
    int32_t error_code;
    /** A NUL-terminated UTF-8 view of a text value; cell and evaluation getters back it with read scratch, while an
     *  owned result handle may provide an owned snapshot. */
    const char* text;
  } u;
} fm_value_t;

/* -------------------------------------------------------------------------- */
/* Construction / lifecycle                                                   */
/* -------------------------------------------------------------------------- */

/**
 * @brief Builds a default workbook (a single sheet named `"Sheet1"`).
 *
 * Lifetime: owned (until fm_workbook_destroy).
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
 * Lifetime: owned (until fm_workbook_destroy).
 *
 * @param profile_id  An id from `fm_workbook_set_profile_id`,
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
 * `fm_workbook_add_font` / `_fill` / `_border` / `_cell_xf` that
 * deduplicates against one of those reserved
 * slots -- including slot 0 -- returns that slot's own index rather than
 * a new one past it; see `fm_workbook_add_font`'s own doc for the dedup
 * rule.
 *
 * Lifetime: owned (until fm_workbook_destroy).
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
 * Names in the file resolve by the identity fold (see the name identity notes
 * above `fm_workbook_get_sheet_name`): an exact match, then a case-only match,
 * then a unique wide fold match. Two names of one collection that collide
 * only under the wide fold load and stay as stored; the load neither rejects
 * nor renames them. A threaded comment's `dT` is normalised: a trailing `Z`
 * is dropped, and a numeric offset or any other malformed stamp reads as no
 * `created` (the attribute is omitted again on save).
 *
 * Lifetime: owned (until fm_workbook_destroy).
 *
 * @param bytes  Pointer to the first byte of the OOXML or MS-XLSB archive.
 * @param len    Length of the buffer in bytes.
 * @param out    On success receives the freshly allocated handle.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `bytes` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `len == 0`;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the file exceeds an engine
 *         ceiling (the context's `part=` names it, e.g. part=zip_parts,
 *         part=gradient_stops or part=array_elements);
 *         another `FM_STATUS_IO_*` code on parse / archive failure.
 */
FM_API fm_status_t fm_workbook_load(const uint8_t* bytes, size_t len, fm_workbook_t** out);

/*
 * Contract shared by `fm_read_diagnostics_info_t` and `fm_save_diagnostics_info_t`:
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
 *     same layout on native and wasm32 (24 bytes, 4-aligned) and a
 *     binding's hand-written layout cannot silently disagree with the
 *     native one.
 */

/**
 * @brief Loss and fallback counters captured by the load that created a
 *        workbook handle.
 */
typedef struct {
  uint32_t struct_size;
  /** The number of formula cells whose stored formula could not be decoded. XLSB only. */
  uint32_t undecoded_formula_count;
  /** The number of defined names skipped because their stored formula could not be decoded. XLSB only. */
  uint32_t undecoded_defined_name_count;
  /** The number of package parts whose content type could not be resolved, so they were not decoded. XLSB only; the
   *  XLSX reader keeps unmodelled parts as passthrough. */
  uint32_t undecoded_part_count;
  /** The number of entries dropped on load: one per merge, hyperlink or data validation and one per
   *  conditional-format block with an unusable reference (XLSX only), and one per table column whose header dxf or
   *  header cell style was dropped, which Excel refuses to open (XLSX and XLSB), and one per XLSB table part dropped
   *  whole because its header slots could not be cleared. */
  uint32_t skipped_feature_count;
  /** The number of workbook parts whose content type was unrecognised; at most 1. XLSX only. */
  uint32_t unknown_content_type_count;
} fm_read_diagnostics_info_t;

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
FM_API fm_status_t fm_workbook_get_diagnostics(const fm_workbook_t* wb, fm_read_diagnostics_info_t* out_diagnostics);

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
 * every stored cell — call it at coarse boundaries (after a load or
 * a recalc), not per edit.
 *
 * @param wb         Workbook to measure.
 * @param out_byte_count  On success receives the estimate, in bytes.
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if either pointer is
 *         `NULL`.
 */
FM_API fm_status_t fm_workbook_get_memory_usage(const fm_workbook_t* wb, size_t* out_byte_count);

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
 * (`fm_workbook_save`, `fm_workbook_save_with_diagnostics`):
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
 * @brief Container format a save writes. The engine's own format enum keeps
 *        its own ordinals; these are the C values.
 */
typedef enum {
  FM_WORKBOOK_FORMAT_XLSX = 0,
  FM_WORKBOOK_FORMAT_XLSB = 1,
} fm_workbook_format_t;

/**
 * @brief Serialises the workbook to an in-memory byte stream in the
 *        container `format` (an `fm_workbook_format_t`).
 *
 * `FM_WORKBOOK_FORMAT_XLSX` (0) produces an `.xlsx` package and
 * `FM_WORKBOOK_FORMAT_XLSB` (1) an MS-XLSB package via the `.xlsb` writer.
 * `format` is a raw signed 32-bit ordinal so FFI callers can probe unknown
 * values without constructing an out-of-domain C enum.
 *
 * On success the caller receives a heap-allocated buffer that MUST be
 * released with `fm_buffer_destroy` (NOT `free`). This ownership rule applies
 * to every save API returning bytes -- `fm_workbook_save` and
 * `fm_workbook_save_with_diagnostics` alike. Mixing allocators across the
 * boundary is undefined.
 *
 * Lifetime: owned (until fm_buffer_destroy).
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param format     Target container format.
 * @param out_bytes  Receives a pointer to the freshly allocated buffer.
 * @param out_len    Receives the buffer length in bytes.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `format` is outside `fm_workbook_format_t`;
 *         `FM_STATUS_IO_FEATURE_NOT_WRITABLE` (5028) when the workbook holds a
 *         feature this writer cannot write to xlsb, with nothing written: the
 *         error context's `kind=` names the feature and, for a sheet-scoped
 *         one, `sheet_index=` the sheet;
 *         another `FM_STATUS_IO_*` code on archive / writer failure.
 */
FM_API fm_status_t fm_workbook_save(const fm_workbook_t* wb, int32_t format, uint8_t** out_bytes, size_t* out_len);

/**
 * @brief Loss and fallback counters produced by one save.
 *
 * The save-side counterpart to `fm_read_diagnostics_info_t`, under the same
 * contract.
 *
 * Summing these fields does NOT give "how many things did I lose":
 * `dropped_part_count` and `dropped_relationship_count` intentionally
 * both rise for one dropped part. See `dropped_relationship_count`.
 */
typedef struct {
  uint32_t struct_size;
  /** The number of formula cells emitted as cached literals. Always 0 for XLSX. */
  uint32_t downgraded_formula_count;
  /** The number of sheet features not lowered to records. Always 0 for XLSX. */
  uint32_t deferred_feature_count;
  /** The number of passthrough parts dropped for colliding with a generated path, or for two entries claiming one path.
   *  Applies to both containers. */
  uint32_t dropped_part_count;
  /** The number of relationships dropped because their target part is gone, at package, workbook or (XLSB) sheet scope.
   *  Rises with `dropped_part_count` for one dropped part, so do not add them. */
  uint32_t dropped_relationship_count;
  /** The number of tables emitted under a writer-assigned id. Always 0 for XLSB. */
  uint32_t renumbered_part_count;
} fm_save_diagnostics_info_t;

/**
 * @brief Serialises `wb` and reports what the save cost.
 *
 * `format` uses the same raw signed 32-bit ordinal contract as
 * `fm_workbook_save`. Every output parameter — including every
 * `*out_diagnostics` field — is written to its zero / `NULL` state before
 * argument and model validation, and stays there on every non-`FM_STATUS_OK` (0)
 * return. `*out_diagnostics` is therefore always written, success or
 * failure; on failure it reads all-zero.
 *
 * The returned `out_bytes` buffer follows the `fm_buffer_destroy` ownership
 * rule documented above.
 *
 * Lifetime: owned (until fm_buffer_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `format` is outside `fm_workbook_format_t`;
 *         `FM_STATUS_IO_FEATURE_NOT_WRITABLE` (5028) when the workbook holds a
 *         feature this writer cannot write to xlsb, with nothing written: the
 *         error context's `kind=` names the feature and, for a sheet-scoped
 *         one, `sheet_index=` the sheet;
 *         another `FM_STATUS_IO_*` code on archive / writer failure.
 */
FM_API fm_status_t fm_workbook_save_with_diagnostics(const fm_workbook_t* wb, int32_t format, uint8_t** out_bytes,
                                                     size_t* out_len, fm_save_diagnostics_info_t* out_diagnostics);

/**
 * @brief Releases a buffer returned by any Formulon save API returning
 *        `out_bytes`: `fm_workbook_save` or
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

/*
 * Name identity. Every named collection (sheets, tables, table columns,
 * defined names, pivots, protected ranges, cell styles, custom table styles)
 * shares one identity fold with two strengths:
 *
 *   * The narrow fold is locale-independent Unicode simple case folding
 *     (C/S mappings only, without normalization). It is the hashing key and
 *     the duplicate check a load applies.
 *   * The wide fold is the workbook collation, which also ignores full/half
 *     width, hiragana/katakana and control characters.
 *
 * A write (add, rename, copy, insert) is refused with
 * `FM_STATUS_ALREADY_EXISTS` when another member holds the name under the
 * wide fold. A lookup by name tries an exact match, then a case-only match,
 * then the one member matching under the wide fold, if exactly one does. A
 * built-in cell style also answers to its name in the workbook profile's
 * locale, and among exact or case-only matches the first record wins. A
 * default sheet name is free under the wide fold as well. Folding is not
 * normalization, full (expanding) case folding or locale tailoring.
 */

/**
 * @brief Returns the display name of the sheet at `sheet_index` (UTF-8).
 *
 * The returned pointer is a model-backed view into the workbook. It remains
 * valid until a mutation that touches the sheet list (for example
 * `fm_workbook_add_sheet`) or until the handle is destroyed. Reads on the
 * same handle, including read-scratch-backed reads, do not invalidate it.
 *
 * Lifetime: transient.
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param sheet_index      0-based sheet index. Must be `< sheet_count`.
 * @param out_name   Receives a NUL-terminated UTF-8 pointer.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_sheet_name(const fm_workbook_t* wb, size_t sheet_index, const char** out_name);

/**
 * @brief Finds the sheet `name` resolves to under the identity fold.
 *
 * Resolution order: an exact match, a case-only match, then the one sheet
 * matching under the wide fold (a unique wide fold match; none when two
 * sheets of a loaded file collide there). A name that matches no sheet is
 * not an error: the call succeeds with `*out_present == 0`.
 *
 * @param wb          Workbook handle. Must be non-NULL.
 * @param name        Sheet name (UTF-8). Must be non-NULL.
 * @param out_present Receives `1` when a sheet matches, `0` otherwise.
 * @param out_sheet_index   Receives the 0-based sheet index when present; `0` otherwise.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `name`, `out_present` or `out_sheet_index` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) if `name` is not well-formed UTF-8.
 */
FM_API fm_status_t fm_workbook_get_sheet_index(const fm_workbook_t* wb, const char* name, int32_t* out_present,
                                               size_t* out_sheet_index);

/**
 * @brief Adds a new worksheet before the sheet now at `insert_before`.
 *
 * A `NULL` name takes the profile's default sheet name `<base><N>` (`Sheet2`,
 * `Tabelle2`, ...): `<base>` is the profile's first sheet name without its
 * trailing `1`, and `N` the smallest number whose name no sheet holds under
 * the wide fold of the identity fold.
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param name       Strict UTF-8 display name, or `NULL` for the default
 *                   name. Worksheet identity is the identity fold (see the
 *                   name identity notes above).
 * @param insert_before 0-based index of the sheet the new one goes before;
 *                   `sheet_count` appends.
 * @param out_sheet_index  Receives the new sheet's index; must be non-NULL.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_sheet_index` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `insert_before > sheet_count`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_SHEET_NAME` (11) when `name` is empty,
 *         exceeds 31 UTF-16 units, contains a forbidden character
 *         (`: \ / ? * [ ]`), begins or ends with an apostrophe, or is
 *         `History` in any case -- the rules `fm_workbook_rename_sheet` applies;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when `name` equals an existing sheet's
 *         name under the wide fold;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the workbook already contains
 *         the maximum supported sheet count (context `part=sheets requested=<n> limit=65535`).
 */
FM_API fm_status_t fm_workbook_add_sheet(fm_workbook_t* wb, const char* name, size_t insert_before,
                                         size_t* out_sheet_index);

/**
 * @brief Moves the sheet at `from_index` before the sheet now at `insert_before`.
 *
 * `insert_before` names a position in the current sheet list, as
 * `fm_workbook_add_sheet` and `fm_workbook_copy_sheet` do: with three sheets,
 * `insert_before == 3` moves sheet 0 to the end. Moving a sheet before itself
 * or before the sheet after it is a successful no-op.
 *
 * @param wb            Workbook handle. Must be non-NULL.
 * @param from_index    0-based index of the sheet to move.
 * @param insert_before 0-based index of the sheet it goes before;
 *                      `sheet_count` appends.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `from_index >= sheet_count` or
 *         `insert_before > sheet_count`.
 */
FM_API fm_status_t fm_workbook_move_sheet(fm_workbook_t* wb, size_t from_index, size_t insert_before);

/**
 * @brief Copies the worksheet at `source_index` into the workbook.
 *
 * The copy is placed before the sheet now at `insert_before` (`sheet_count`
 * appends) and its index is written to `*out_sheet_index`. References qualified
 * with the source sheet are re-pointed at the copy; tables are renamed and
 * rebound; sheet-scoped names, and workbook names referring to the source,
 * gain a copy scoped to the new sheet; pivot tables share the source's cache.
 * The active sheet and the tab selection do not change. See
 * the status list below for the failure cases. On any failure the workbook is
 * left unchanged.
 *
 * @param wb            Workbook handle. Must be non-NULL.
 * @param source_index  0-based index of the sheet to copy.
 * @param insert_before 0-based index of the sheet the copy goes before;
 *                      `sheet_count` appends.
 * @param new_name      Strict UTF-8 display name validated like
 *                      `fm_workbook_add_sheet`, or `NULL` to name the copy as
 *                      Excel does (`Src (2)`).
 * @param out_sheet_index     Receives the copy's index. Must be non-NULL.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_sheet_index` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `new_name` is malformed UTF-8;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `source_index >= sheet_count` or
 *         `insert_before > sheet_count`;
 *         `FM_STATUS_INVALID_SHEET_NAME` (11) when `new_name` is empty, too long or contains a
 *         forbidden character;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when `new_name` is an existing sheet's name;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the workbook is full;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the source is very hidden;
 *         `FM_STATUS_NOT_IMPLEMENTED` (3) when the source is not a worksheet or carries
 *         content the copy cannot model (unrecognised relationships, slicers,
 *         shapes, pivot charts, or an `.xlsb` drawing the reader could not lift).
 */
FM_API fm_status_t fm_workbook_copy_sheet(fm_workbook_t* wb, size_t source_index, size_t insert_before,
                                          const char* new_name, size_t* out_sheet_index);

/**
 * @brief Reads the active sheet: the first `<workbookView>`'s `activeTab`.
 *
 * `0` is reported when the workbook has no `<bookViews>` or `activeTab` is
 * outside the sheet list; `tabSelected` is not consulted. A load follows
 * Excel's open rule: a hidden active sheet is unhidden (only that one), a very
 * hidden one passes activation to the first visible sheet, and an active sheet
 * that is not selected becomes the only `tabSelected` one. Inserting, removing and moving sheets keep the stored index
 * on the sheet it named; removing or hiding the active sheet activates the
 * first visible sheet after it, else the nearest visible one before it, as
 * Excel does.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_sheet_index` is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_active_sheet(const fm_workbook_t* wb, size_t* out_sheet_index);

/**
 * @brief Marks sheet `sheet_index` active so a save records it.
 *
 * Writes `activeTab` on the first `<workbookView>` and lowers its
 * `firstSheet` to `sheet_index` when that is larger; every other attribute is kept
 * verbatim. A workbook with no `<bookViews>` gets a minimal
 * `<bookViews><workbookView activeTab="N"/></bookViews>`. Sheet `sheet_index`
 * becomes the only `tabSelected` sheet, as clicking its tab does in Excel.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is not a sheet of the workbook;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the sheet is hidden or very hidden.
 */
FM_API fm_status_t fm_workbook_set_active_sheet(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Removes the sheet at `sheet_index`.
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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the workbook has only one sheet, or
 *         `sheet_index` is the only visible sheet (the workbook is left unchanged).
 */
FM_API fm_status_t fm_workbook_remove_sheet(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Renames the sheet at `sheet_index` to `new_name` (UTF-8).
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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `new_name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_SHEET_NAME` (11) when `new_name` is empty,
 *         exceeds 31 UTF-16 units, contains a forbidden character
 *         (`: \ / ? * [ ]`), begins or ends with an apostrophe, or is
 *         `History` in any case;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when `new_name` equals another sheet's
 *         name under the wide fold of the identity fold, which is not
 *         normalization, full (expanding) case folding, locale tailoring, or a
 *         claim of full Excel-equivalence.
 */
FM_API fm_status_t fm_workbook_rename_sheet(fm_workbook_t* wb, size_t sheet_index, const char* new_name);

/**
 * @brief The `local_sheet_id` of a workbook-scoped defined name.
 */
#define FM_DEFINED_NAME_SCOPE_WORKBOOK (-1)

/**
 * @brief One defined name, as `fm_workbook_set_defined_name` takes it and
 *        `fm_workbook_get_defined_name` / `fm_workbook_get_defined_name_at`
 *        read it.
 *
 * On read `name` and `comment` are model-backed views and `formula` is a
 * read-scratch-backed view; see `fm_workbook_get_defined_name_at` for the
 * lifetimes.
 */
typedef struct {
  uint32_t struct_size;
  /** The name, non-empty UTF-8; on read, the stored spelling as a model-backed view. */
  const char* name;
  /** The formula text, non-empty on input with at most one leading `=`; on read, spelled as the sheet formula getter
   *  spells it, with the `=`. */
  const char* formula;
  /** The 0-based index of the sheet the name is local to, or `FM_DEFINED_NAME_SCOPE_WORKBOOK` for workbook scope. */
  int32_t local_sheet_id;
  /** The Name Manager note; `NULL` or empty on input means none, and a read returns `NULL` when there is none. */
  const char* comment;
  /** 1 hides the name from the Name Manager; 0 leaves it visible. */
  int32_t hidden;
} fm_defined_name_t;

/**
 * @brief Sets a defined name, appending it or replacing the entry with the
 *        same `name` (case-insensitive) in the scope `local_sheet_id`.
 *
 * The record replaces the whole entry: a replaced name takes the given
 * formula, comment and hidden flag, so a `NULL` `comment` or a zero `hidden`
 * resets that attribute. A formula that does not parse is rejected and
 * leaves the name list unchanged. One leading `=` is accepted and dropped:
 * the name stores the formula as files carry it, spelled as
 * `fm_sheet_set_formula` stores a cell formula, and the getters read it back
 * with the `=`. `fm_workbook_remove_defined_name`
 * removes a name. A new name follows `fm_sheet_add_table`'s display-name
 * syntax, refuses `\` followed by one character, and may not be a table's
 * name; replacing an existing entry does not re-check it. A new name that is,
 * or starts with, a reference in the profile locale's R1C1 letters (de
 * `Z1S1`) is refused; redefining an existing name stays allowed. The check
 * uses the profile at the time of the call, so a loaded name is never
 * re-checked.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `name`, `name->name` or
 *         `name->formula` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a string is malformed UTF-8;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `local_sheet_id` is neither
 *         `FM_DEFINED_NAME_SCOPE_WORKBOOK` nor a sheet of the workbook;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `name` or `formula` is empty, or a new `name`
 *         breaks the name syntax or opens with an R1C1 reference in the profile locale's letters;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when a new `name` is a table's name;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position) when `formula` does not
 * parse.
 */
FM_API fm_status_t fm_workbook_set_defined_name(fm_workbook_t* wb, const fm_defined_name_t* name);

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
 * @brief Inserts the inclusive row span `[first, last]` on `sheet_index`:
 *        `last - first + 1` blank rows open at `first`. Cells at `first` and
 *        beyond shift forward. An insert that would push a value, formula,
 *        note, fill, border or merge past the sheet bound is refused; a font
 *        or number format, hyperlink, validation or conditional format pushed
 *        off is dropped, as are row and column formats. References across the
 *        workbook are rewritten to track the new positions; references that
 *        would land past the sheet bound collapse to `#REF!`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last row (1048575);
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the edit cuts into a data table,
 *         legacy array block or PivotTable, or would push one of them, a
 *         table, or a value or formula cell past the sheet bound; the
 *         workbook is unchanged.
 */
FM_API fm_status_t fm_sheet_insert_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/**
 * @brief Deletes the inclusive row span `[first, last]` on `sheet_index`. The
 *        deleted rows are dropped wholesale; subsequent rows shift back.
 *        References pointing inside the deleted interval collapse to
 *        `#REF!`; references past the deletion shift back.
 *
 *        A span holding a table's header row, without covering the whole
 *        table, fails with `FM_STATUS_PRECONDITION_FAILED`. A span covering
 *        every data row keeps the first data row, emptied; one covering the
 *        totals row turns the table's totals off and clears every column's
 *        totals settings. A span covering the whole table removes it.
 *
 * @return Same status codes as `fm_sheet_insert_rows`.
 */
FM_API fm_status_t fm_sheet_delete_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/**
 * @brief Inserts the inclusive column span `[first, last]` on `sheet_index`.
 *        Mirrors `fm_sheet_insert_rows` along the column axis.
 *
 * @return Same status codes as `fm_sheet_insert_rows`, with the last column (16383) as the bound.
 */
FM_API fm_status_t fm_sheet_insert_columns(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/**
 * @brief Deletes the inclusive column span `[first, last]` on `sheet_index`.
 *        Mirrors `fm_sheet_delete_rows` along the column axis.
 *
 * @return Same status codes as `fm_sheet_insert_rows`, with the last column (16383) as the bound.
 */
FM_API fm_status_t fm_sheet_delete_columns(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/**
 * @brief Direction the cells of the band move in `fm_sheet_insert_cells` /
 *        `fm_sheet_delete_cells`.
 *
 * An insert takes `DOWN` or `RIGHT`, a delete `UP` or `LEFT`. The shift is a
 * required scalar argument, not a record field, so its zero is not a default:
 * `FM_CELL_SHIFT_DOWN` (0) is valid for an insert and
 * `FM_STATUS_INVALID_ARGUMENT` (2) for a delete.
 */
typedef enum {
  FM_CELL_SHIFT_DOWN = 0,
  FM_CELL_SHIFT_RIGHT = 1,
  FM_CELL_SHIFT_UP = 2,
  FM_CELL_SHIFT_LEFT = 3,
} fm_cell_shift_t;

/*
 * Insert / Delete Cells. The band is the range's columns for `DOWN` / `UP`
 * and its rows for `RIGHT` / `LEFT`; only the band moves, by the range's
 * height or width. The moving area is the band from the range's leading edge
 * to the grid's end, the deleted rectangle included for a delete. A range
 * covering the whole perpendicular extent (every column for `DOWN` / `UP`,
 * every row for `RIGHT` / `LEFT`) behaves as `fm_sheet_insert_rows` /
 * `fm_sheet_insert_columns` and their delete counterparts.
 *
 *   * References (cells, ranges, defined names, hyperlink ranges) are
 *     rewritten only when their perpendicular extent lies inside the band;
 *     one straddling the band, a whole column or a whole row keeps its
 *     address. An insert moves or grows them; a delete turns a wholly
 *     deleted one into `#REF!` and shrinks a partly deleted one.
 *   * A conditional-format or validation range straddling the moving area
 *     splits into its moving and fixed parts. A merge straddling it, or
 *     partly intersecting the deleted rectangle, is unmerged first.
 *   * Notes and threaded comments move with their cell and are removed with
 *     a deleted one. A table, data table or PivotTable whose perpendicular
 *     extent the band covers moves as a whole, and one the deleted rectangle
 *     contains is removed.
 *   * Row and column layout, manual breaks, drawing anchors and the sheet
 *     view do not move. An inserted cell takes the XF of the cell directly
 *     above (`DOWN`) or to its left (`RIGHT`), if that cell has one.
 *
 * Every check runs before anything changes; on any failure the workbook is
 * unchanged. `FM_STATUS_PRECONDITION_FAILED` (9) covers every refusal:
 *
 *   * a table, multi-cell array formula or data table intersecting the
 *     moving area whose perpendicular extent the band does not cover;
 *   * a PivotTable intersecting the moving area, unless the deleted rectangle
 *     contains it, or the band covers its whole perpendicular extent and the
 *     range does not lie inside it;
 *   * an insert that would push a value, formula, note, fill, border or merge
 *     past the grid (a font or number format, hyperlink, validation or
 *     conditional format pushed off is dropped);
 *   * an edit Excel accepts without changing anything (a delete over part of
 *     a table, an insert inside a PivotTable).
 */

/**
 * @brief Inserts blank cells over `range` on `sheet_index`, shifting the band
 *        `DOWN` or `RIGHT` (Insert / Delete Cells, above).
 *
 * `shift` is an `fm_cell_shift_t`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `range` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `range` is inverted or outside the grid,
 *         `shift` is outside `fm_cell_shift_t` or is `UP` / `LEFT`, or `range`
 *         spans every row with `DOWN` or every column with `RIGHT`;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) for a refusal listed above.
 */
FM_API fm_status_t fm_sheet_insert_cells(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                         int32_t shift);

/**
 * @brief Deletes the cells of `range` on `sheet_index`, closing the gap by
 *        shifting the band `UP` or `LEFT` (Insert / Delete Cells, above).
 *
 * `shift` is an `fm_cell_shift_t`; `FM_CELL_SHIFT_DOWN` (0) is not a delete
 * direction.
 *
 * A band covering a table's columns follows the table rules of
 * `fm_sheet_delete_rows` (header row refused, all data rows keep the first,
 * totals row delete turns totals off).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `range` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `range` is inverted or outside the grid,
 *         `shift` is outside `fm_cell_shift_t` or is `DOWN` / `RIGHT`, or `range`
 *         spans every row with `UP` or every column with `LEFT`;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) for a refusal listed above.
 */
FM_API fm_status_t fm_sheet_delete_cells(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                         int32_t shift);

/* -------------------------------------------------------------------------- */
/* Cell mutation                                                              */
/* -------------------------------------------------------------------------- */

/*
 * The number / bool / error / text / formula setters store an entry the way
 * Excel does: a cell that has no style of its own takes its row's style
 * (`fm_sheet_set_row_xf_index`), else its column's, as its cell XF. A cell that
 * already carries an XF, including an explicit XF 0, keeps it.
 *
 * Every setter here, `fm_sheet_set_blank` included, refuses a non-anchor
 * cell of a legacy array formula (CSE) block: it returns
 * `FM_STATUS_PRECONDITION_FAILED` (9) and leaves the block unchanged. A write to the
 * block's anchor replaces the whole block.
 *
 * A cell holds at most 32,767 UTF-16 code units of text, Excel's limit; a
 * longer string (counted in UTF-16 units of the UTF-8 input) is refused by
 * `fm_sheet_set_text`, `fm_sheet_set_cell_phonetic` and each run of
 * `fm_sheet_set_cell_phonetic_runs`.
 *
 * Common failures, besides those each entry lists:
 * `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) for an out-of-range sheet;
 * `FM_STATUS_INVALID_ARGUMENT` (2) for a `row` or `col` outside the grid and
 * for text past the limit above; `FM_STATUS_PRECONDITION_FAILED` (9) for a
 * non-anchor cell of a CSE block.
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
 * to it (`fm_sheet_rename_table_column`), and the cell holds the final
 * name.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `value` is NaN or infinite.
 */
FM_API fm_status_t fm_sheet_set_number(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col, double value);

/**
 * @brief Stores a boolean literal. `value` is interpreted as a C
 *        boolean: `0` is FALSE, any non-zero is TRUE.
 *
 * A table's header cell takes `TRUE` or `FALSE` as text instead: the column is renamed
 * to it (`fm_sheet_rename_table_column`), and the cell holds the final
 * name.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`.
 */
FM_API fm_status_t fm_sheet_set_bool(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col, int32_t value);

/**
 * @brief Stores a static Excel error literal at `(row, col)`.
 *
 * This writes a cell whose value is `#DIV/0!`, `#VALUE!`, etc. It is
 * distinct from writing a formula that evaluates to an error.
 *
 * A table's header cell takes the error's name (`#N/A`) as text instead: the column is renamed
 * to it (`fm_sheet_rename_table_column`), and the cell holds the final
 * name.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `error` is out of range.
 */
FM_API fm_status_t fm_sheet_set_error(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                      fm_error_code_t error);

/**
 * @brief Stores a text literal. The destination cell copies the UTF-8
 *        contents, so `text` does not need to outlive the call.
 *
 * A table's header cell takes `text` as its column name instead: the column is renamed
 * to it (`fm_sheet_rename_table_column`), and the cell holds the final
 * name.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `text` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `text` exceeds 32,767 UTF-16 units.
 */
FM_API fm_status_t fm_sheet_set_text(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                     const char* text);

/**
 * @brief Stores the literal `*value` at `(row, col)`, dispatching on its kind.
 *
 * `FM_VALUE_KIND_NUMBER`, `_BOOL`, `_ERROR` and `_TEXT` behave as
 * `fm_sheet_set_number`, `fm_sheet_set_bool`, `fm_sheet_set_error` and
 * `fm_sheet_set_text`; `FM_VALUE_KIND_BLANK` as `fm_sheet_set_blank` (it removes
 * the content and keeps the cell's own non-default format, or an explicit
 * default where it overrides a row or column format). A value
 * read with `fm_sheet_get_value` can be written back unchanged.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure, or what the
 *         dispatched setter returns;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `value` is `NULL`, or `value->u.text` is
 *         `NULL` for a text value;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `value->kind` is `FM_VALUE_KIND_ARRAY`, `_REF`, `_LAMBDA` or
 *         not a `fm_value_kind_t` member.
 */
FM_API fm_status_t fm_sheet_set_value(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                      const fm_value_t* value);

/**
 * @brief What a text offset counts: UTF-16 code units, which is how Excel
 *        indexes string positions, or Unicode code points. A character outside
 *        the Basic Multilingual Plane is 2 UTF-16 units and 1 code point.
 */
typedef enum { FM_TEXT_OFFSET_UNIT_UTF16 = 0, FM_TEXT_OFFSET_UNIT_CODE_POINT = 1 } fm_text_offset_unit_t;

/**
 * @brief One `<rPh>` block: the kana reading `text` covers the half-open
 *        span `[start, end)` of the cell's surface text.
 *
 * A read fills both offset pairs; a write reads the pair its
 * `fm_text_offset_unit_t` argument names and ignores the other.
 * A cell whose reading was typed in one go carries the single run
 * `{0, <length of the text>, kana}`.
 *
 * The spans are observable rather than presentation detail: `PHONETIC`
 * substitutes each annotated span with its kana and passes the text
 * outside every span through unchanged, so a partially annotated string
 * surfaces a mix of kana and original characters.
 *
 * Frozen value type: no `struct_size`.
 */
typedef struct {
  /** The inclusive start of the span, in UTF-16 code units. */
  uint32_t start;
  /** The exclusive end of the span, in UTF-16 code units. */
  uint32_t end;
  /** The inclusive start of the span, in code points. */
  uint32_t start_cp;
  /** The exclusive end of the span, in code points. */
  uint32_t end_cp;
  /** The kana reading of the span. */
  const char* text;
} fm_phonetic_run_t;

/**
 * @brief Stores a whole-cell phonetic guide (OOXML `<rPh>`).
 *
 * The guide is independent of the cell's visible text. Passing an empty
 * string clears an existing guide. The destination cell copies `text`, so
 * the input does not need to outlive the call.
 *
 * This entry point has no span vocabulary, so the reading it stores covers
 * the surface text end to end — one run spanning the whole cell. Feeding
 * `fm_sheet_get_cell_phonetic`'s result back through here therefore
 * flattens a partially annotated cell into a single whole-cell annotation.
 * Use `fm_sheet_set_cell_phonetic_runs` to preserve the spans.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `text` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `text` exceeds 32,767 UTF-16 units.
 */
FM_API fm_status_t fm_sheet_set_cell_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              const char* text);

/**
 * @brief Stores a cell's phonetic guide as one `<rPh>` block per run,
 *        spans included.
 *
 * `count == 0` clears the annotation (`runs` may be `NULL` in that case).
 * `offset_unit` names the pair of offsets each run supplies (`start`/`end` or
 * `start_cp`/`end_cp`); code-point offsets are counted against the cell's
 * current surface text, and a position past its end counts one unit per
 * code point. Each run's `text` is copied, so the array and its strings do not need to
 * outlive the call.
 *
 * The runs must describe a partition of the surface text in reading order:
 * every run needs `start <= end`, and each run must start at or after the
 * previous run's `end`. Overlapping or out-of-order runs are rejected rather
 * than normalised, because `PHONETIC` walks them in order and a run whose
 * span has already been passed would silently move its kana onto characters
 * it does not read.
 *
 * `end` is not checked against the cell's current text length. The surface
 * text may legitimately be written after the annotation, and every
 * value-mutating setter clears the annotation anyway, so a span that
 * overshoots is treated as covering the remainder of the text rather than
 * as an error.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, if `runs` is `NULL` with
 *         a non-zero `count`, or if any run's `text` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a run's `text` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `offset_unit` is not a `fm_text_offset_unit_t`
 *         member, the runs are not an ordered partition, or a run's `text` exceeds 32,767 UTF-16 units.
 */
FM_API fm_status_t fm_sheet_set_cell_phonetic_runs(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                   int32_t offset_unit, const fm_phonetic_run_t* runs, size_t count);

/**
 * @brief Kana form a phonetic guide generates readings in (OOXML
 *        `<phoneticPr type>`). Zero is the state of an absent
 *        `<phoneticPr>`, not the schema default.
 */
typedef enum {
  FM_PHONETIC_TYPE_HALFWIDTH_KATAKANA = 0,
  FM_PHONETIC_TYPE_FULLWIDTH_KATAKANA = 1,
  FM_PHONETIC_TYPE_HIRAGANA = 2,
  FM_PHONETIC_TYPE_NO_CONVERSION = 3,
} fm_phonetic_type_t;

/**
 * @brief How a phonetic guide's kana is distributed over the characters it
 *        covers (OOXML `<phoneticPr alignment>`). Zero is the state of an
 *        absent `<phoneticPr>`, not the schema default.
 */
typedef enum {
  FM_PHONETIC_ALIGNMENT_NO_CONTROL = 0,
  FM_PHONETIC_ALIGNMENT_LEFT = 1,
  FM_PHONETIC_ALIGNMENT_CENTER = 2,
  FM_PHONETIC_ALIGNMENT_DISTRIBUTED = 3,
} fm_phonetic_alignment_t;

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
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`.
 */
FM_API fm_status_t fm_sheet_set_cell_show_phonetic(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                   int32_t show);

/**
 * @brief How a cell's phonetic guide renders (OOXML `<phoneticPr>`), read by
 *        `fm_sheet_get_cell_phonetic_properties` and written back by
 *        `fm_sheet_set_cell_phonetic_properties`.
 *
 * The all-zero triple is what Excel infers for a guide that arrived with no
 * `<phoneticPr>` at all.
 */
typedef struct {
  uint32_t struct_size;
  /** The index into the workbook's font table of the ruby text. */
  uint32_t font_index;
  /** The kana form Excel generates readings in, an `fm_phonetic_type_t`. */
  int32_t type;
  /** How the kana is distributed over the characters it covers, an `fm_phonetic_alignment_t`. */
  int32_t alignment;
  /** `1` when the properties are explicit: a phoneticPr element or a set call, an all-zero triple included; `0` for the
   *  state an absent element resolves to. Ignored on input. */
  int32_t stated;
} fm_phonetic_properties_t;

/**
 * @brief Stores how a cell's phonetic guide renders (`fm_phonetic_properties_t`),
 *        or resets it to the unstated all-zero triple when `properties` is
 *        `NULL`.
 *
 * Deliberately separate from `fm_sheet_set_cell_phonetic_runs` in both
 * directions: editing the readings does not reset the rendering, and
 * setting the rendering does not touch the readings. The two are stored
 * together and travel together, but a host that only has one of them should
 * not have to supply the other.
 *
 * Only observable on a cell that has runs — the writer emits no element for
 * an unannotated cell — and any value-mutating setter clears the runs and
 * these properties together, so set them after the cell's text. A host that
 * does not care can leave them alone.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `font_index` exceeds 65535, or when `type` or
 *         `alignment` is outside its enumeration.
 */
FM_API fm_status_t fm_sheet_set_cell_phonetic_properties(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                         uint32_t col, const fm_phonetic_properties_t* properties);

/**
 * @brief Stores a `Blank` literal at `(row, col)`. Equivalent to
 *        clearing a cell.
 *
 * Removes the content and keeps the cell's own non-default format, or an
 * explicit default format where that overrides a row or column format. A
 * cell with neither stops existing: it is no longer counted or enumerated and
 * inherits its row or column format. Clearing a blank that has only a format
 * changes nothing.
 *
 * A table's header cell is not cleared: the column takes the profile's
 * default name (`Column1`, `列1`; `fm_sheet_rename_table_column` with an
 * empty name), and the cell holds that name.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`.
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
 * Text that fails as A1, or names an unbound R1C1-shaped reference
 * (`RC[-1]`, `R1C1`), is read as R1C1 relative to `(row, col)` as Excel
 * does, so `=RC[-1]` stores as an A1 reference and `=LET(B1,5,B1+R1C1)`
 * keeps `B1` as a name.
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
 * table does not parse. A typed structured reference to a table whose name
 * is, or starts with, a reference in the profile locale's R1C1 letters is
 * `FM_STATUS_PARSER_INVALID_REFERENCE`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-mutation failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position) when `formula` does not
 * parse.
 */
FM_API fm_status_t fm_sheet_set_formula(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        const char* formula);

/* -------------------------------------------------------------------------- */
/* Cell read                                                                  */
/* -------------------------------------------------------------------------- */

/*
 * Common failures, besides those each entry lists:
 * `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) for an out-of-range sheet and
 * `FM_STATUS_INVALID_ARGUMENT` (2) for a `row` or `col` outside the grid.
 */

/**
 * @brief Reads the cached cell value at `(row, col)`.
 *
 * Read contract: the value follows the workbook's calc mode. Under
 * `FM_CALC_MODE_AUTO` the read first computes the dirty cells (those a change
 * left uncomputed); under `FM_CALC_MODE_MANUAL` it returns the values as of
 * the last recalc, and only `fm_workbook_recalc` brings them up to date.
 * The other value getters cite this contract.
 *
 * For `FM_VALUE_KIND_TEXT`, `out->u.text` is a read-scratch-backed view into the
 * handle's transient read storage. It may be invalidated by the next
 * successful scratch-backed read on this handle; a validation-rejected call
 * does not refresh the storage. It is also invalid after a mutation or
 * handle destruction. The scratch is reset on each successful producer so a
 * long-lived read loop does not accumulate memory. Callers that need to
 * retain the string must copy it before another call.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-read failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_value(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                      fm_value_t* out);

/**
 * @brief Reads the cell's phonetic guide (OOXML `<rPh>`). `*out_present` is
 *        `0` and `*out_text` `NULL` (still `FM_STATUS_OK` (0)) when the cell
 *        has no guide.
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
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-read failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_present` or `out_text` is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_cell_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              int32_t* out_present, const char** out_text);

/**
 * @brief Counts the `<rPh>` blocks attached to a cell.
 *
 * Zero for a cell with no annotation, including one that does not exist.
 * Pair with `fm_sheet_get_cell_phonetic_run_at` to walk the spans.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-read failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is `NULL`.
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
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-read failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `run_index` is past the cell's run count.
 */
FM_API fm_status_t fm_sheet_get_cell_phonetic_run_at(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                     uint32_t col, size_t run_index, fm_phonetic_run_t* out);

/**
 * @brief Reads how a cell's phonetic guide renders into `*out`
 *        (`fm_phonetic_properties_t`).
 *
 * Reports the all-zero triple with `stated == 0` for a cell with no
 * annotation, including one that does not exist — the same state Excel
 * infers for a guide written without a `<phoneticPr>` element. Like the
 * triple, `stated` is cleared by any value-mutating setter. A save spells the
 * element out beside any runs, and XLSB always stores the triple, so a
 * reloaded or XLSB-loaded cell with runs reports `stated == 1`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-read failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_cell_phonetic_properties(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                         uint32_t col, fm_phonetic_properties_t* out);

/**
 * @brief Reads the cell's "Show Phonetic" flag into `*out_show` (`0` or `1`).
 *        An absent cell reports `0`. To enumerate the flagged cells, use
 *        `fm_sheet_get_show_phonetic_cells`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-read failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_show` is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_cell_show_phonetic(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                   uint32_t col, int32_t* out_show);

/**
 * @brief Lists every stored cell of `range` whose "Show Phonetic" flag is set
 *        (the flag `fm_sheet_get_cell_show_phonetic` reads), in row-major
 *        order (row, then column, ascending).
 *
 * A blank cell that carries only the flag is listed; cells outside the used
 * range are included, so pass the whole grid to list every stored cell of the
 * sheet. `*out_count` always receives the number of cells; at most `capacity` are
 * written. Call with `capacity == 0` (`out` may then be `NULL`) to size the
 * buffer first.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `range` or `out_count` is `NULL`, or
 *         `out` is `NULL` while `capacity > 0` (`*out_count` is then `0` when `out_count` is non-`NULL`);
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the rect is out of range.
 */
FM_API fm_status_t fm_sheet_get_show_phonetic_cells(const fm_workbook_t* wb, size_t sheet_index,
                                                    const fm_cell_rect_t* range, fm_cell_address_t* out,
                                                    size_t capacity, size_t* out_count);

/**
 * @brief Renders the lambda closure stored at `(sheet_index, row, col)`
 *        as Excel formula text.
 *
 * Use this to recover a printable form of a `Value::Lambda` whose
 * source is otherwise opaque — for example, a cell whose formula
 * evaluates to a lambda via `=LET(...) -> LAMBDA(...)` rather than
 * carrying a literal `LAMBDA(...)` in the formula text. The rendering
 * is the full surface form `=LAMBDA(p1,p2,body)`, with the leading `=` every
 * formula getter carries, suitable for re-parsing through
 * `fm_sheet_set_formula`.
 * `*out_present` is `0` and `*out_text` `NULL` (still `FM_STATUS_OK` (0))
 * when the cell is absent or its cached value is not a lambda.
 *
 * When present, `*out_text` is a read-scratch-backed view into the handle's
 * transient read storage. It may be invalidated by the next successful
 * scratch-backed read on this handle; a validation-rejected call does not
 * refresh the storage. It is also invalid after a mutation or handle
 * destruction. The scratch is reset on each successful producer so a
 * long-lived handle does not accumulate memory. Callers that need to retain
 * the string must copy it before another call.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-read failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_lambda_text(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                            int32_t* out_present, const char** out_text);

/* -------------------------------------------------------------------------- */
/* Ad-hoc, side-effect-free formula evaluation                                */
/* -------------------------------------------------------------------------- */

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
 * `*out` is always `FM_VALUE_KIND_BOOL`. Read-only, identical purity contract to
 * `fm_sheet_evaluate_formula_array`.
 *
 * @note The same self-reference caveat as `fm_sheet_evaluate_formula_array`
 *       applies: a rule that reads its own anchored cell sees the cached
 *       value rather than raising `#REF!`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `formula`, or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the target cell or anchor cell is out of range;
 *         `FM_STATUS_OUT_OF_MEMORY` (4) when the evaluation arena is exhausted.
 */
FM_API fm_status_t fm_sheet_evaluate_cf_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                uint32_t anchor_row, uint32_t anchor_col, const char* formula,
                                                fm_value_t* out);

/**
 * @brief Opaque handle for the whole result of an ad-hoc array evaluation.
 *
 * Owns the row-major cells `fm_sheet_evaluate_formula_array` produced. Read
 * the shape with `fm_value_array_get_dimensions` and the cells with
 * `fm_value_array_get_value_count` / `fm_value_array_get_value_at`. Release
 * with `fm_value_array_destroy`. The handle is an owned snapshot: it does not
 * follow later workbook edits and stays valid after the workbook is
 * destroyed.
 */
typedef struct fm_value_array fm_value_array_t;

/**
 * @brief Evaluates `formula` as if entered at `(row, col)` on
 *        `sheet_index` and returns the WHOLE result — array included — as an
 *        owned handle, WITHOUT mutating the workbook.
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
 * The evaluation computes every precedent it reads, even under
 * `FM_CALC_MODE_MANUAL`, without storing the results; `fm_sheet_get_value`
 * follows the calc mode.
 *
 * The result is preserved in full: a dynamic-array formula such as
 * `=SEQUENCE(2,3)` yields a 2 x 3 array; a scalar result such as `=1+2`
 * yields a 1 x 1 array. The bindings' `evaluateFormula` / `evaluate_formula`
 * use this entry point and fold a 1 x 1 result to a scalar.
 *
 * A blank cell the formula reads evaluates as 0, as it does in a cell (`=B9`
 * on an empty B9 is 0); `fm_sheet_get_value` reads the blank as
 * `FM_VALUE_KIND_BLANK`.
 *
 * A degenerate empty array result (which producers should never emit)
 * surfaces as a single `#VALUE!` cell.
 *
 * @note Because the ad-hoc formula is never inserted into the dependency
 *       graph, a self-reference such as `=A1` anchored at A1's own address
 *       reads A1's *cached* value rather than raising `#REF!` or engaging
 *       iterative calc. This diverges from true in-cell F9 entry;
 *       self-reference detection is out of scope for this API.
 *
 * Lifetime: owned (until fm_value_array_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `formula`, or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row` or `col` is out of range;
 *         `FM_STATUS_OUT_OF_MEMORY` (4) when the evaluation arena is exhausted.
 *
 * On success the caller owns `*out` and must free it with
 * `fm_value_array_destroy`.
 */
FM_API fm_status_t fm_sheet_evaluate_formula_array(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                   uint32_t col, const char* formula, fm_value_array_t** out);

/** @brief Frees a handle returned by `fm_sheet_evaluate_formula_array`. `NULL` is a no-op. */
FM_API void fm_value_array_destroy(fm_value_array_t* array);

/**
 * @brief Writes the array's row and column counts, both at least 1.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_value_array_get_dimensions(const fm_value_array_t* array, uint32_t* out_rows,
                                                 uint32_t* out_columns);

/**
 * @brief Writes the number of cells, `rows * cols`, to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_value_array_get_value_count(const fm_value_array_t* array, size_t* out_count);

/**
 * @brief Reads the row-major `index`-th cell into `*out`; cell `(r, c)` is at
 *        `index = r * cols + c`.
 *
 * For a `FM_VALUE_KIND_TEXT` cell, `out->u.text` is an owned-snapshot view: it stays
 * valid until `fm_value_array_destroy`.
 *
 * Lifetime: owned (until fm_value_array_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `array` or `out` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is not below the cell count.
 */
FM_API fm_status_t fm_value_array_get_value_at(const fm_value_array_t* array, size_t index, fm_value_t* out);

/* -------------------------------------------------------------------------- */
/* Reference text                                                             */
/* -------------------------------------------------------------------------- */

/**
 * @brief A parsed reference: a 0-based, inclusive rectangle with an optional
 *        sheet qualifier.
 *
 * Text that is not a reference is a parser status, never a record.
 */
typedef struct {
  uint32_t struct_size;
  /** The unquoted UTF-8 sheet qualifier as a read-scratch-backed view; `NULL` when the text has none. */
  const char* sheet_name;
  /** The 0-based inclusive rectangle, with a whole column or row expanded to the grid edge. */
  fm_cell_rect_t rect;
} fm_parsed_reference_info_t;

/**
 * @brief The reference notation of `fm_workbook_parse_reference`,
 *        `fm_localize_formula` and `fm_canonicalize_formula`.
 */
typedef struct {
  uint32_t struct_size;
  /** The notation, an `fm_tristate_t`: default and false use A1, true uses R1C1 with the letters and brackets of
   *  the profile's locale. */
  int32_t r1c1; /* fm_tristate_t */
  /** 1 when `base_row` and `base_col` are set; 0 means no base. A1 does not use the base, but one outside the grid
   *  is refused in either notation. */
  int32_t has_base;
  /** The 0-based row that R1C1 relative axes are measured from. */
  uint32_t base_row;
  /** The 0-based column that R1C1 relative axes are measured from. */
  uint32_t base_col;
} fm_reference_notation_spec_t;

/**
 * @brief Parses `text` as `INDIRECT` reads a reference under the workbook's
 *        profile: A1 unless `spec->r1c1` is `FM_TRISTATE_TRUE`, then R1C1
 *        (letters and brackets of the profile's locale). `spec == NULL` is a
 *        plain A1 parse with no base.
 *
 * Accepts an optional local sheet qualifier (quoted with `''` escaping, or
 * bare), `$` markers (dropped), a `:` range in either order, and whole
 * columns / rows. 3D, external, named, structured and multi-area
 * references, surrounding whitespace, a sheet name longer than 31 UTF-16
 * units, and (with `has_base` 0) an R1C1 relative axis are not references:
 * the call fails, and `fm_get_last_parse_error` reports
 * `FM_PARSE_ERROR_CODE_INVALID_REFERENCE` over the whole text. With `has_base` 1,
 * R1C1 relative axes are measured from `(base_row, base_col)` and wrap at
 * the grid edge.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_PARSER_INVALID_REFERENCE` (1004) when `text` is not a reference;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `text` or `out` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `text` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `r1c1` is outside `fm_tristate_t`,
 *         `has_base` is not 0/1, or `has_base` is 1 with the base outside the grid.
 */
FM_API fm_status_t fm_workbook_parse_reference(const fm_workbook_t* wb, const char* text,
                                               const fm_reference_notation_spec_t* spec,
                                               fm_parsed_reference_info_t* out);

/**
 * @brief Formats the rectangle `range` as A1 reference text, the inverse of
 *        `fm_workbook_parse_reference`.
 *
 * `sheet_name` `NULL` or empty writes no qualifier; otherwise the name is
 * quoted as a stored formula quotes it (and is not validated as a sheet
 * name). A single cell writes one endpoint (`A1`), a rectangle spanning
 * every row writes columns (`B:D`), one spanning every column writes rows
 * (`2:11`). `absolute` 1 prefixes each component with `$`. `*out_text` is
 * a read-scratch-backed view.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `range` or `out_text` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `sheet_name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a coordinate is outside the grid,
 *         the rectangle is inverted, or `absolute` is not 0/1.
 */
FM_API fm_status_t fm_workbook_format_reference(const fm_workbook_t* wb, const char* sheet_name,
                                                const fm_cell_rect_t* range, int32_t absolute, const char** out_text);

/* -------------------------------------------------------------------------- */
/* Iteration / dump                                                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the number of written coordinates on `sheet_index`.
 *
 * Counts every coordinate that holds content (a value, a formula, phonetic
 * text, an array role or a spill anchor commit or rejection) and every blank
 * coordinate with its own format: a non-default format, or an explicit default
 * format where it overrides a row or column format. A cell that was cleared
 * without leaving such a format is not counted; neither are coordinates that
 * only inherit a row or column format. Also counts every dynamic-array spill
 * phantom that has no stored slot: phantoms carry an effective value through
 * `fm_sheet_get_value` / `fm_sheet_get_cell_at` but live only in the spill
 * table. The count is cheap to compute, but a sheet formatted over a large
 * rectangle counts every coordinate of it.
 *
 * The count is the upper bound for `fm_sheet_get_cell_at` iteration: indices
 * in `[0, count)` are dense and enumerate every counted coordinate in
 * `(row, col)` ascending order. The enumeration is stable for a given
 * workbook state (no cell mutation has occurred between calls).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL` or `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cell_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `index`-th cell on `sheet_index`.
 *
 * Enumerates the coordinates counted by `fm_sheet_get_cell_count`, sorted by
 * `(row, col)` ascending. A coordinate that has only a format is enumerated as
 * a blank value with a `NULL` formula. The values follow the calc-mode read
 * contract of `fm_sheet_get_value`. The enumeration can be as large as the
 * rectangle that was formatted; read a large sheet in ranges
 * (`fm_sheet_get_cells_in_range`, `fm_sheet_get_cell_xf_indices`) rather than
 * by index.
 *
 * On success the call writes `*out_row`, `*out_col`, `*out_value`, and
 * — if `out_formula != NULL` — points `*out_formula` at the cell's
 * formula text, spelled as `fm_sheet_get_formula` returns it (or `NULL`
 * when the cell is a pure literal). The formula pointer has the
 * read-scratch lifetime described below for `*out_value`: valid until the
 * next successful scratch-backed read on this handle, the next mutation, or
 * handle destruction.
 *
 * A `FM_VALUE_KIND_TEXT` payload in `*out_value`, by contrast, is a
 * read-scratch-backed view and may be invalidated by the next successful
 * scratch-backed read on this handle; a validation-rejected call does not
 * refresh the storage. Mutation and handle destruction also invalidate it.
 * The scratch is refreshed by each successful producer so iterating
 * `cell_at` does not accumulate memory.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_row`, `out_col`, or
 *         `out_value` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index >= cell_count`.
 */
FM_API fm_status_t fm_sheet_get_cell_at(const fm_workbook_t* wb, size_t sheet_index, size_t index, uint32_t* out_row,
                                        uint32_t* out_col, const char** out_formula, fm_value_t* out_value);

/**
 * @brief Reads the formula of the cell at `(row, col)` in A1 notation.
 *
 * The text carries the leading `=`, as `fm_sheet_get_cell_at` reports it. A
 * cell that holds no formula (a literal, a spill phantom, or no cell at
 * all) yields `*out_present == 0` and `*out_formula == NULL` (still
 * `FM_STATUS_OK` (0)). A data table's result
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
 * When present, `*out_formula` is a read-scratch-backed view: valid until the
 * next successful scratch-backed read, any mutation, or handle destruction.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         coordinate is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        int32_t* out_present, const char** out_formula);

/**
 * @brief Reads the formula of the cell at `(row, col)` in R1C1 notation,
 *        relative to that cell.
 *
 * The text carries the leading `=`, as `fm_sheet_get_formula` does, and
 * follows the A1 formatter's spacing and case rules. Structured references take the display form
 * `fm_sheet_get_formula` gives them at this cell. Sheet and external qualifiers are quoted as a whole
 * (`'[Book.xlsx]My Sheet'!R1C1`) only when the path, book or sheet name needs
 * it; a 3-D external reference with no path stays unquoted. A cell that holds no formula yields
 * `*out_present == 0` and `*out_formula == NULL`, as for `fm_sheet_get_formula`. A present
 * `*out_formula` is a read-scratch-backed view, as for `fm_sheet_get_formula`.
 *
 * A formula loaded from a file that does not parse is kept verbatim: it
 * evaluates to `#NAME?` and fails here with `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000).
 * `fm_sheet_set_formula` never produces that state.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         coordinate is outside the grid;
 *         `FM_STATUS_PARSER_UNEXPECTED_TOKEN` (1000) when the stored formula does not
 *         parse.
 */
FM_API fm_status_t fm_sheet_get_formula_r1c1(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             int32_t* out_present, const char** out_formula);

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
 * Lifetime: owned (until fm_cell_range_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `range` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         rectangle is inverted or outside the grid, or `cursor` is past
 *         the last cell of the grid.
 */
FM_API fm_status_t fm_sheet_get_cells_in_range(const fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                               uint64_t cursor, uint32_t limit, fm_cell_range_t** out);

/**
 * @brief Number of cells in the page.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_cell_range_get_cell_count(const fm_cell_range_t* range, size_t* out_count);

/**
 * @brief Reads cell `index` of the page.
 *
 * `*out_formula` is the stored A1 formula (as `fm_sheet_get_formula`
 * reports it), or `NULL` when the cell holds no formula; `out_formula` may
 * be `NULL` to skip it. Both the formula and a `FM_VALUE_KIND_TEXT` payload are an
 * owned snapshot held by the handle and stay valid until `fm_cell_range_destroy`.
 *
 * Lifetime: owned (until fm_cell_range_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `range`, `out_row`, `out_col`, or
 *         `out_value` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_cell_range_get_cell_at(const fm_cell_range_t* range, size_t index, uint32_t* out_row,
                                             uint32_t* out_col, const char** out_formula, fm_value_t* out_value);

/**
 * @brief Reads the cursor that resumes after this page, or `UINT64_MAX` when
 *        the rectangle holds no further populated cell.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_cell_range_next_cursor(const fm_cell_range_t* range, uint64_t* out_cursor);

/** @brief Releases a cell-range handle. `range == NULL` is a no-op. */
FM_API void fm_cell_range_destroy(fm_cell_range_t* range);

/* -------------------------------------------------------------------------- */
/* Display text                                                               */
/* -------------------------------------------------------------------------- */

/**
 * @brief Outcome of a display-text rendering.
 *
 * `FM_DISPLAY_STATUS_OVERFLOW`: the value cannot be shown in the format (a date
 * serial outside the calendar, for one) and the text is `"########"`.
 * `FM_DISPLAY_STATUS_INVALID_FORMAT`: the format code is malformed and the text is
 * the General rendering.
 */
typedef enum {
  FM_DISPLAY_STATUS_OK = 0,
  FM_DISPLAY_STATUS_OVERFLOW = 1,
  FM_DISPLAY_STATUS_INVALID_FORMAT = 2,
} fm_display_status_t;

/**
 * @brief Reads the text Excel displays for the cell at `(row, col)`.
 *
 * Rendered from the resolved value and effective cell/row/column/default
 * XF number format under the workbook's date system, independent of column
 * width: a `*x` repeat fill contributes no characters and an overflow is
 * reported through `*out_status` rather than as a width-dependent `#` run.
 * A coordinate with no value or spill result is `""`. `*out_text` is a
 * read-scratch-backed view. The text follows the calc-mode read contract of
 * `fm_sheet_get_value`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         coordinate is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_display_text(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             const char** out_text, int32_t* out_status);

/**
 * @brief Renders `value` under the stored-form number format `number_format`
 *        (`General`, `[Red]`, ...), using the workbook's date system.
 *
 * An empty `number_format` means General. `value` must be blank, a number,
 * a boolean, text, or an error. `*out_text` is a read-scratch-backed view.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`, or a
 *         `FM_VALUE_KIND_TEXT` value carries a `NULL` text;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the value kind is not one of the above,
 *         a number is NaN or infinite, or an error code is out of range.
 */
FM_API fm_status_t fm_workbook_format_value(const fm_workbook_t* wb, const fm_value_t* value, const char* number_format,
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
 * @brief Reads the `index`-th defined name with its scope and attributes.
 *
 * On success `out->name` and `out->comment` are model-backed views into the
 * workbook handle, valid until the next mutation of the defined-name list --
 * `fm_workbook_set_defined_name`, `fm_workbook_remove_defined_name`, or any
 * workbook edit that rewrites defined names as a side effect (a structural
 * row/column insert/delete, or a sheet rename/removal) -- or until the
 * handle is destroyed. `out->formula` is valid until the next successful
 * scratch-backed read on this handle, the next mutation, or handle
 * destruction.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_defined_name_at(const fm_workbook_t* wb, size_t index, fm_defined_name_t* out);

/**
 * @brief Reads the defined name `name` (Unicode simple case folding) in scope
 *        `local_sheet_id` (`FM_DEFINED_NAME_SCOPE_WORKBOOK` for workbook scope,
 *        else a 0-based sheet index). `*out_present` is `1` and `*out` filled the way
 *        `fm_workbook_get_defined_name_at` fills it when there is one, else
 *        `0`, the strings of `*out` `NULL` and `hidden` `0` (still
 *        `FM_STATUS_OK` (0)).
 *
 * Lifetime: transient; the views in `*out` follow
 * `fm_workbook_get_defined_name_at`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is not valid UTF-8;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `local_sheet_id` is neither
 *         `FM_DEFINED_NAME_SCOPE_WORKBOOK` nor a sheet of the workbook.
 */
FM_API fm_status_t fm_workbook_get_defined_name(const fm_workbook_t* wb, const char* name, int32_t local_sheet_id,
                                                int32_t* out_present, fm_defined_name_t* out);

/**
 * @brief Removes the defined name `name` (Unicode simple case folding) in scope
 *        `local_sheet_id` (`FM_DEFINED_NAME_SCOPE_WORKBOOK` for workbook scope,
 *        else a 0-based sheet index), refreshing the formulas that reference it.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is not valid UTF-8;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `local_sheet_id` is neither
 *         `FM_DEFINED_NAME_SCOPE_WORKBOOK` nor a sheet of the workbook;
 *         `FM_STATUS_NOT_FOUND` (6) when no such name exists in that scope.
 */
FM_API fm_status_t fm_workbook_remove_defined_name(fm_workbook_t* wb, const char* name, int32_t local_sheet_id);

/* -------------------------------------------------------------------------- */
/* Tables (sheet-owned)                                                       */
/* -------------------------------------------------------------------------- */

/*
 * A table belongs to a sheet and is addressed as `(sheet_index, table_index)`,
 * `table_index` being its 0-based position among that sheet's tables. Table
 * names are unique across the workbook (formulas use them), so
 * `fm_workbook_find_table` is the one workbook-wide lookup by name.
 *
 * Common failures, besides those each entry lists:
 * `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) for an out-of-range sheet and
 * `FM_STATUS_INDEX_OUT_OF_RANGE` (13) for an out-of-range table or column.
 */

/**
 * @brief Writes the number of tables on sheet `sheet_index` to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_table_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief One table as `fm_sheet_get_table_at` reads it.
 *
 * `column_count` bounds `fm_sheet_get_table_column_at`'s `column_index`.
 * `name` and `style_name` are model-backed views into the workbook handle,
 * valid until the next mutation of the table list -- `fm_sheet_add_table` /
 * `fm_sheet_update_table` / `fm_sheet_remove_table` / `fm_sheet_rename_table`,
 * or any workbook edit that rewrites table metadata as a side effect (a
 * structural row/column insert/delete, or a sheet rename/removal) -- or until
 * the handle is destroyed. Reads do not invalidate them.
 */
typedef struct {
  uint32_t struct_size;
  /** The table name formulas use. */
  const char* name;
  /** The table's range, header and totals rows included. */
  fm_cell_rect_t range;
  /** The name of the table's style, a built-in or custom one; `NULL` when the table has no style element or the element
   *  names no style. */
  const char* style_name;
  /** 1 when the first column is emphasized; 0 also when the style element is absent. */
  int32_t show_first_column;
  /** 1 when the last column is emphasized; 0 also when the style element is absent. */
  int32_t show_last_column;
  /** 1 when odd rows are banded; 0 also when the style element is absent. */
  int32_t show_row_stripes;
  /** 1 when odd columns are banded; 0 also when the style element is absent. */
  int32_t show_column_stripes;
  /** 1 when the table has a header row, else 0. */
  uint32_t header_row_count;
  /** 1 when the table has a totals row, else 0. */
  uint32_t totals_row_count;
  /** The number of columns in the table. */
  uint32_t column_count;
} fm_table_info_t;

/**
 * @brief Reads table `table_index` on `sheet_index`: its name, range, style
 *        and row and column counts.
 *
 * Lifetime: transient; the views in `*out` follow `fm_table_info_t`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a loaded table's `ref` is not an A1 range.
 */
FM_API fm_status_t fm_sheet_get_table_at(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                         fm_table_info_t* out);

/**
 * @brief Finds the table named `name` (case-insensitive) anywhere in the
 *        workbook. `*out_present` is `1` and the two indices are written when
 *        there is one; otherwise `*out_present` is `0` and the indices are
 *        left unchanged (still `FM_STATUS_OK` (0)).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is not valid UTF-8.
 */
FM_API fm_status_t fm_workbook_find_table(const fm_workbook_t* wb, const char* name, int32_t* out_present,
                                          size_t* out_sheet_index, size_t* out_table_index);

/**
 * @brief Writes the runtime id of table `table_index` on `sheet_index` to `*out`.
 *
 * A runtime id names one table for the life of the handle: it is assigned on
 * the first request (non-zero, never reused, shared numbering with pivot
 * runtime ids), follows the table across every edit that moves it within the
 * sheet's table list (adding, removing or reordering others, structural edits,
 * removal of another sheet), and is not saved. A copied sheet's tables get ids
 * of their own.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_table_runtime_id(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                 uint32_t* out);

/**
 * @brief Writes the sheet and the sheet-local index of the table whose runtime
 *        id is `id` to `*out_sheet_index` and `*out_table_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_NOT_FOUND` (6) when no table has `id` (the table was
 *         removed, or `id` is `0` or was never handed out).
 */
FM_API fm_status_t fm_workbook_resolve_table_runtime_id(const fm_workbook_t* wb, uint32_t id, size_t* out_sheet_index,
                                                        size_t* out_table_index);

/** @brief A worksheet table to create (`fm_sheet_add_table`). */
typedef struct {
  uint32_t struct_size;
  /** The cells the table covers, header row included, on the receiving sheet. */
  fm_cell_rect_t range;
  /** The new table's name; it must not be one another table already uses. */
  const char* name;
  /** One name per column of the range, as for `columns`; `NULL` takes the names from the header cells of the sheet. */
  const char* const* column_names;
  /** The number of entries in `column_names`. */
  uint32_t column_count;
  /** The built-in or custom table style to apply; `NULL` is the workbook's table-style default, else
   *  `TableStyleMedium2`; an empty string is invalid. */
  const char* style_name;
  /** Tristate: default and true create a styled table, false creates the table without a style element. */
  int32_t styled; /* fm_tristate_t */
  /** Tristate: whether the table has a header row; default is on. */
  int32_t header_row; /* fm_tristate_t */
  /** Tristate: whether the table has a totals row; default is off. */
  int32_t totals_row; /* fm_tristate_t */
  /** Tristate: whether the first column is emphasized; default takes the creation default, which is off. */
  int32_t show_first_column; /* fm_tristate_t */
  /** Tristate: whether the last column is emphasized; default takes the creation default, which is off. */
  int32_t show_last_column; /* fm_tristate_t */
  /** Tristate: whether odd rows are banded; default takes the creation default, which is on for a styled table. */
  int32_t show_row_stripes; /* fm_tristate_t */
  /** Tristate: whether odd columns are banded; default takes the creation default, which is off. */
  int32_t show_column_stripes; /* fm_tristate_t */
} fm_table_spec_t;

/**
 * @brief Creates a worksheet table (`xl/tables/tableN.xml`) over `spec->range`
 *        on `sheet_index`, with its style toggles, and writes its sheet-local
 *        index to `*out_table_index`.
 *
 * `column_names` `NULL` takes the headers from the first row's displayed
 * text, naming a blank cell with the locale's column word and the first free
 * number and writing that name into the cell. Otherwise it holds one
 * non-empty header for every column of the range (`column_count` equals the
 * range's width), each must equal its non-blank header cell, and a blank
 * header cell receives its name. Names must be unique within the table.
 *
 * `name` is the one table name, the name formulas use; the file's `name`
 * attribute follows it. It follows Excel's name syntax (a letter, `_` or `\`
 * first; letters, digits, `_`, `.`, `?` and `\` after; at most 255 UTF-16
 * units; not TRUE / FALSE, an A1 cell such as `T1`, or a name opening with an
 * R1C1 reference in the profile locale's R1C1 letters (`R2D2` in en, `Z1S1` in
 * de)) and may not be used by another table or by a
 * defined name of any scope (case-insensitive). The range may not overlap
 * another table, a PivotTable, a data table, an array formula or a merged
 * range on the sheet.
 *
 * `style_name` `NULL` applies the workbook's table-style default
 * (`fm_workbook_get_table_style_defaults`), else `TableStyleMedium2`; a styled
 * table starts with row stripes on and the other three toggles off. `styled`
 * `FM_TRISTATE_FALSE` creates the table without `<tableStyleInfo>` and
 * requires a `NULL` `style_name`. `styled`, `header_row`,
 * `totals_row` and the four toggles are `fm_tristate_t`; a toggle set to
 * `FM_TRISTATE_TRUE` or `FM_TRISTATE_FALSE` is written in the same call, and
 * a failure creates nothing. The table's filter definition is created
 * automatically. `range` is the whole table, so with `totals_row` on its
 * last row is the totals row.
 *
 * A one-row `range` with a header row creates the header plus one empty data
 * row below it: the cells of the table's columns in that row shift down when
 * occupied, and the table ends one row below `range`. That fails with
 * `FM_STATUS_PRECONDITION_FAILED` when `range` is the last grid row, and
 * `header_row` with `totals_row` on a one-row `range` is
 * `FM_STATUS_INVALID_ARGUMENT`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `spec`, `spec->name` or `out_table_index` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `range` is inverted or outside the grid, a
 *         name breaks the rules above, `style_name` is empty, given with `styled`
 *         `FALSE`, names a custom style that does not apply to tables or names a built-in slicer style
 *         (`SlicerStyleLight1`, ...), the column names do not match the range, or a flag is outside
 *         `fm_tristate_t`;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when `name` is taken;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the range overlaps one of the objects above, or the
 *         workbook's default table style, used when `style_name` is `NULL`, does not apply to tables;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a string is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_add_table(fm_workbook_t* wb, size_t sheet_index, const fm_table_spec_t* spec,
                                      size_t* out_table_index);

/**
 * @brief A partial update of one table column (an entry of
 *        `fm_table_patch_t::columns`), with the fields of `fm_table_column_info_t`
 *        a column's settings are read through.
 *
 * Frozen value type: no `struct_size`. Each field has an
 * `fm_patch_op_t` companion; `CLEAR` removes a setting and `KEEP` leaves it.
 * The result is checked as a whole: a custom totals function needs a non-empty
 * formula, and any other function drops the formula. A `SET` of a label or
 * formula that is `NULL` or empty is `FM_STATUS_INVALID_ARGUMENT` (2).
 */
typedef struct {
  /** The operation applied to the totals-row function: keep it, set `totals_row_function`, or clear it to
   *  `FM_TABLE_TOTALS_ROW_FUNCTION_NONE`. */
  fm_patch_op_t totals_row_function_op;
  /** The totals-row function to set, an `fm_table_totals_row_function_t`. */
  int32_t totals_row_function;
  /** The operation applied to the totals-row label. */
  fm_patch_op_t totals_row_label_op;
  /** The totals-row label to set. */
  const char* totals_row_label;
  /** The operation applied to the formula of the custom totals-row function. */
  fm_patch_op_t totals_row_formula_op;
  /** The formula to set, with or without one leading `=`; structured references bind as `fm_sheet_set_formula` binds
   *  them at the column's totals cell (`SUBTOTAL(109,[Qty])` stores `SUBTOTAL(109,Table[Qty])`). Applies only while
   *  the function is `FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM`. */
  const char* totals_row_formula;
  /** The operation applied to the calculated-column formula: keep it, set `calculated_formula`, or clear the mark
   *  (the cells keep what they hold). */
  fm_patch_op_t calculated_formula_op;
  /** The calculated-column formula to set, with or without one leading `=`, as entered in the column's first data row
   *  (`[@Qty]` binds to this table). The mark is the state: a formula that differs from the column's mark (after
   *  binding and normalizing) marks the column and fills every data cell, constants included, with it moved to its
   *  row; the same formula changes nothing and refills nothing. A table with no data row only takes the mark. */
  const char* calculated_formula;
} fm_table_column_patch_t;

/**
 * @brief Partial update of a table's range, header / totals rows and columns
 *        (`fm_sheet_update_table`).
 *
 * Each field has an `fm_patch_op_t` companion (Patches, in the header
 * preamble); no field can be cleared. The style is patched through
 * `fm_sheet_update_table_style_options`.
 */
typedef struct {
  uint32_t struct_size;
  /** The operation applied to `range`: keep it, or set it; it cannot be cleared. */
  fm_patch_op_t range_op;
  /** The table's new range; it is checked as for `fm_sheet_add_table`, resizes the table, keeps the width the column
   *  list describes and may not overlap the objects `fm_sheet_add_table` lists. */
  fm_cell_rect_t range;
  /** The operation applied to `header_row`: keep it, or set it; it cannot be cleared. */
  fm_patch_op_t header_row_op;
  /** Turns the header row on (1) or off (0); it moves cells (see `fm_sheet_update_table`). */
  int32_t header_row;
  /** The operation applied to `totals_row`: keep it, or set it; it cannot be cleared. */
  fm_patch_op_t totals_row_op;
  /** Turns the totals row on (1) or off (0), growing or shrinking the table by one row (see
   *  `fm_sheet_update_table`). */
  int32_t totals_row;
  /** The operation applied to `columns`: keep the columns as they are, or apply the entries of `columns`; it cannot be
   *  cleared. */
  fm_patch_op_t columns_op;
  /** The `columns_count` column patches `columns_op` `SET` applies: entry i patches column i, so the array may be
   *  shorter than the table's column count and an entry with every op `KEEP` changes nothing; `NULL` with a zero
   *  count. */
  const fm_table_column_patch_t* columns;
  /** The number of entries of `columns`; more than the table has columns is `FM_STATUS_INDEX_OUT_OF_RANGE` (13). */
  uint32_t columns_count;
} fm_table_patch_t;

/**
 * @brief Merges `patch` into table `table_index`. The whole patch is checked
 *        before anything changes.
 *
 * Header row off: the header cells are deleted (value and style), the range
 * shrinks by its top row, nothing shifts, the table's AutoFilter is dropped with
 * its criteria, the rows those criteria hid are shown again, and the column
 * names are kept; while it is off a `[#Headers]` reference
 * evaluates to `#REF!`. Header row on: when the row above the data is empty in
 * the table's columns, the names are written there and the range grows up by
 * one row; when it is occupied (or is the sheet's first row), the table's
 * columns shift down one row from the first data row, as `fm_sheet_insert_cells`
 * with `FM_CELL_SHIFT_DOWN`, and the names are written into the freed row.
 * Turning it off on a table without data rows fails.
 *
 * With `range` and a toggle in one patch the toggle (header, then totals)
 * applies first and `range` is checked against the toggled shape; a `range`
 * whose first row is not the toggled shape's first row is
 * `FM_STATUS_INVALID_ARGUMENT`, and any refusal leaves the whole patch
 * unapplied.
 *
 * Totals row on: one row is inserted below the table in
 * the table's columns, shifting the cells beneath down (as
 * `fm_sheet_insert_cells` with `FM_CELL_SHIFT_DOWN`), and each column's
 * stored totals settings are written there; the data rows are kept. If a merged range, another table,
 * a PivotTable, an array formula or a data table straddles the table's
 * columns below it, nothing shifts and the row directly below is used when
 * it is free (no value, formula, spill output, merged range or table cell;
 * format-only cells count as free).
 *
 * Totals row off: every refusal check runs first. If a cell formula or a
 * defined name references the totals row (a range meeting it, or a
 * `[#Totals]` / `[#All]` structured reference to the table) the call is
 * refused. The totals row then leaves the table's range and the cells
 * beneath shift up; if something straddles the table's columns below, the old
 * totals row is only emptied and nothing shifts. The columns' stored totals
 * labels and functions are kept for the next time it is turned on, and turning
 * it on restores them. Fresh defaults are written instead only when the table
 * has never shown its totals row or lost it to a row delete: the first column
 * takes the profile language's label, the last column a sum when every data
 * cell is a non-date number and a count otherwise, and a one-column table
 * gets the function only.
 *
 * A toggle that changes nothing is a no-op.
 *
 * Column patches apply after the range, header and totals rows, so a column's
 * totals settings take effect on the totals row the patch leaves. They are
 * checked against that final shape before anything changes. A resize first gives each unmarked column the
 * calculated-column mark its cells imply and fills the blank cells of a marked
 * one in the added rows. An empty `columns` changes nothing.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is out of range or `CLEAR` (a column patch's ops may clear),
 *         `range` is inverted, outside the grid, of another width or does not start where the toggled shape
 *         starts, a row flag is not 0 / 1, `columns` is `NULL` with
 *         a `columns_count`, a totals function is outside its enumeration, a label or formula to set is `NULL` or
 *         empty, the totals function that results is custom with no formula, or a calculated-column formula is an
 *         array or dynamic-array formula or moves a reference off the grid;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `columns_count` exceeds the table's column count;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position) when a totals or
 *         calculated-column formula does not parse;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when `range` overlaps another table, a
 *         PivotTable, a data table, an array formula or a merged range,
 *         or when turning the header row on or off cannot shift or has no data row to keep,
 *         or when turning the totals row on finds the row below occupied (a
 *         non-blank cell in the table's columns, or a merged range, table,
 *         PivotTable, array formula, data table or spill over those cells,
 *         while something straddles the table's columns below so that the
 *         insert cannot shift) or the table ends on the last grid row, or
 *         when turning the totals row off finds a formula or defined name
 *         referencing it; nothing changes then.
 */
FM_API fm_status_t fm_sheet_update_table(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                         const fm_table_patch_t* patch);

/**
 * @brief Removes table `table_index` from `sheet_index`; the tables after it
 *        move down one index.
 *
 * Formulas that name the table keep their structured references and
 * evaluate to `#NAME?`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`.
 */
FM_API fm_status_t fm_sheet_remove_table(fm_workbook_t* wb, size_t sheet_index, size_t table_index);

/**
 * @brief A table column's totals-row function (OOXML `totalsRowFunction`), in
 *        `ST_TotalsRowFunction` order; zero is the schema default `none`.
 */
typedef enum {
  FM_TABLE_TOTALS_ROW_FUNCTION_NONE = 0,
  FM_TABLE_TOTALS_ROW_FUNCTION_SUM = 1,
  FM_TABLE_TOTALS_ROW_FUNCTION_MIN = 2,
  FM_TABLE_TOTALS_ROW_FUNCTION_MAX = 3,
  FM_TABLE_TOTALS_ROW_FUNCTION_AVERAGE = 4,
  FM_TABLE_TOTALS_ROW_FUNCTION_COUNT = 5,
  FM_TABLE_TOTALS_ROW_FUNCTION_COUNT_NUMBERS = 6,
  FM_TABLE_TOTALS_ROW_FUNCTION_STD_DEV = 7,
  FM_TABLE_TOTALS_ROW_FUNCTION_VAR = 8,
  FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM = 9,
} fm_table_totals_row_function_t;

/**
 * @brief One table column (`fm_sheet_get_table_column_at`).
 *
 * `name` and `totals_row_label` share the lifetime of `fm_table_info_t`'s
 * views; a present `totals_row_formula` or `calculated_formula` is
 * read-scratch-backed instead.
 */
typedef struct {
  uint32_t struct_size;
  /** The column name; never `NULL`. */
  const char* name;
  /** The totals-row function of the column, an `fm_table_totals_row_function_t`. */
  int32_t totals_row_function;
  /** The totals-row label; `NULL` when absent. */
  const char* totals_row_label;
  /** The formula of the custom totals-row function with its leading `=`, spelled as at the column's totals cell inside
   *  the table (such as `=SUBTOTAL(109,[Qty])`); `NULL` for any other function. */
  const char* totals_row_formula;
  /** The calculated-column formula with its leading `=`, spelled as in the column's first data row inside the table
   *  (such as `=[@Qty]*10`); `NULL` when the column has none. Read-scratch-backed. */
  const char* calculated_formula;
} fm_table_column_info_t;

/**
 * @brief Reads column `column_index` (0-based, left to right) of table `table_index`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_table_column_at(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                size_t column_index, fm_table_column_info_t* out);

/**
 * @brief A table's applied style (`<tableStyleInfo>`), the counterpart of
 *        `fm_pivot_style_options_t` (`fm_sheet_get_table_style_options`).
 *
 * `style_name` is a model-backed view sharing `fm_table_info_t`'s lifetime.
 */
typedef struct {
  uint32_t struct_size;
  /** The name of the applied table style; `NULL` for an element without a named style. It may name a custom style. */
  const char* style_name;
  /** Whether the first column is formatted specially. */
  int32_t show_first_column;
  /** Whether the last column is formatted specially. */
  int32_t show_last_column;
  /** Whether odd and even rows are banded. */
  int32_t show_row_stripes;
  /** Whether odd and even columns are banded. */
  int32_t show_column_stripes;
} fm_table_style_options_t;

/**
 * @brief Partial update of a table's style (`fm_sheet_update_table_style_options`).
 *
 * Each field has an `fm_patch_op_t` companion. A patch on a table without
 * `<tableStyleInfo>` creates the element, the fields it leaves at `KEEP` taking
 * the creation defaults of `fm_sheet_add_table`. A patch whose ops are all
 * `KEEP` changes nothing.
 */
typedef struct {
  uint32_t struct_size;
  /** The update operation for `style_name`: keep it, clear it, or set it to the value beside it. */
  fm_patch_op_t style_name_op;
  /** The new style name when the operation is set, non-empty; clearing removes the name and keeps the element. */
  const char* style_name;
  /** The update operation for `show_first_column`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t show_first_column_op;
  /** Whether the first column is formatted specially. */
  int32_t show_first_column;
  /** The update operation for `show_last_column`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t show_last_column_op;
  /** Whether the last column is formatted specially. */
  int32_t show_last_column;
  /** The update operation for `show_row_stripes`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t show_row_stripes_op;
  /** Whether odd and even rows are banded. */
  int32_t show_row_stripes;
  /** The update operation for `show_column_stripes`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t show_column_stripes_op;
  /** Whether odd and even columns are banded. */
  int32_t show_column_stripes;
} fm_table_style_options_patch_t;

/**
 * @brief Reads table `table_index`'s applied style. `*out_present` is `1`
 *        and `*out` filled when the table has `<tableStyleInfo>`; otherwise
 *        `0` and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_present` or `out` is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_table_style_options(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                    int32_t* out_present, fm_table_style_options_t* out);

/**
 * @brief Merges `patch` into table `table_index`'s applied style. The whole
 *        patch is checked before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is out of range, `CLEAR` hits a
 *         toggle, `SET` carries a `NULL` or empty `style_name`, a `SET` `style_name` names a custom
 *         style that does not apply to tables or a built-in slicer style, or a toggle is not 0 / 1;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `style_name` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_update_table_style_options(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                       const fm_table_style_options_patch_t* patch);

/**
 * @brief Renames table `table_index` to `name`.
 *
 * `name` follows `fm_sheet_add_table`'s rules (including the profile locale's
 * R1C1 letters, `R2D2` in en and `Z1S1` in de; renaming to the table's own
 * name in another case stays allowed), unique case-insensitively
 * against every other table and every defined name. Structured references
 * naming the old name are rewritten to the new one wherever formulas live:
 * cells, defined names, conditional formats, data validations,
 * calculated-column and totals-row formulas, and chart series. A pivot cache
 * sourced from the table follows.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `name` is empty or invalid;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when `name` is taken.
 */
FM_API fm_status_t fm_sheet_rename_table(fm_workbook_t* wb, size_t sheet_index, size_t table_index, const char* name);

/**
 * @brief Renames column `column_index` (0-based) of table `table_index`, as typing
 *        over its header cell does.
 *
 * The final name is `name` kept verbatim (spaces and a case-only change
 * included), except that an empty `name` takes the profile's default
 * `<word>N` (`Column1`, `列1`) with the lowest unused N >= 1, and a name
 * another column of the table already has (under Unicode simple case
 * folding) takes `<name>N` with the lowest unused N >= 2 (`Amount2` becomes
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
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `name` or `out_name` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_rename_table_column(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                size_t column_index, const char* name, const char** out_name);

/**
 * @brief Inserts the data rows `[first, last]` (0-based, inclusive) into table
 *        `table_index`, as ListRows.Add does.
 *
 * The cells of the table's columns from the first new row downwards move down
 * by `last - first + 1`, the cells of other columns stay, and the table grows
 * by that many rows. `first` equal to the table's data-row count appends: the
 * rows go above the totals row when the table has one. Every calculated
 * column's formula is written into the new rows with its relative references
 * moved along; the other columns stay blank. Structured references and the
 * table's AutoFilter follow the new extent.
 *
 * @param wb          The workbook handle.
 * @param sheet_index The sheet holding the table.
 * @param table_index The table's index among the tables of `sheet_index`.
 * @param first       The first new data row, 0 to the current data-row count.
 * @param last        The last new data row, not below `first`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `first` is past the data-row
 *         count;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last` is below `first` or the
 *         span reaches past the last row;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when a PivotTable lies in the
 *         shift path, a multi-cell array formula or data table is cut by the
 *         band, or a value, formula, note, fill, border, merge or the table
 *         itself would be pushed past the sheet bound; the workbook is
 *         unchanged.
 */
FM_API fm_status_t fm_sheet_insert_table_rows(fm_workbook_t* wb, size_t sheet_index, size_t table_index, uint32_t first,
                                              uint32_t last);

/**
 * @brief Inserts the columns `[first, last]` (0-based, inclusive) into table
 *        `table_index`, as ListColumns.Add does.
 *
 * The cells of the table's rows (header through totals row) from the first new
 * column rightwards move right by `last - first + 1`, the cells of other rows
 * stay, and the table grows by that many columns. `first` equal to the table's
 * column count appends after the last column. Each new column is named from
 * its left neighbour (`Q2` gives `Q3`; a name without trailing digits, or no
 * left neighbour, takes the profile's column word plus the lowest free number,
 * `Column1` / `列1`); the name is written to the header cell when the table
 * has a header row, and the data and totals cells stay blank.
 *
 * @param wb          The workbook handle.
 * @param sheet_index The sheet holding the table.
 * @param table_index The table's index among the tables of `sheet_index`.
 * @param first       The first new column, 0 to the current column count.
 * @param last        The last new column, not below `first`.
 *
 * @return Same status codes as `fm_sheet_insert_table_rows`, with the last
 *         column (16383) as the bound.
 */
FM_API fm_status_t fm_sheet_insert_table_columns(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                 uint32_t first, uint32_t last);

/**
 * @brief Writes the number of passthrough OOXML parts the reader carried
 *        through unchanged to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is NULL.
 */
FM_API fm_status_t fm_workbook_get_passthrough_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `index`-th passthrough part's path.
 *
 * On success `*out_path` is a model-backed view into the workbook's
 * passthrough-part list. It remains valid until the passthrough list is
 * replaced by a reader/writer round-trip, another invalidating workbook
 * mutation, or handle destruction. Reads do not invalidate it.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_passthrough_at(const fm_workbook_t* wb, size_t index, const char** out_path);

/**
 * @brief Reports whether the reader carried a passthrough part at `path`
 *        (matched exactly, without a leading `/`). `*out_present` is `1` when
 *        there is one, else `0` (still `FM_STATUS_OK` (0)).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `path` is not valid UTF-8.
 */
FM_API fm_status_t fm_workbook_has_passthrough(const fm_workbook_t* wb, const char* path, int32_t* out_present);

/* -------------------------------------------------------------------------- */
/* Sheet UI features (merges, hyperlinks, comments)                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Hyperlink descriptor. String fields are NUL-terminated UTF-8
 *        model-backed views into the workbook handle; they are valid
 *        until the next mutation that touches the sheet's hyperlink
 *        list or until the handle is destroyed. Reads do not invalidate
 *        them. `target`, `location`, `display` and `tooltip` are each
 *        optional: on input `NULL` and `""` both mean absent.
 *
 * A sheet may hold several hyperlinks over the same cells; they are kept in
 * the order they were added (file order after a load), duplicates of one
 * rectangle from a file included.
 */
typedef struct {
  uint32_t struct_size;
  /** The cells the hyperlink covers; the key `fm_sheet_set_hyperlink` replaces on. */
  fm_cell_rect_t rect;
  /** The absolute or relative target (URL, email address, file); `NULL` for an in-workbook link or when the hyperlink
   *  carries none. */
  const char* target;
  /** The in-workbook destination; `NULL` for an external target. */
  const char* location;
  /** The display text override; `NULL` when the hyperlink has none. */
  const char* display;
  /** The tooltip text; `NULL` when the hyperlink has none. */
  const char* tooltip;
} fm_hyperlink_t;

/**
 * @brief A note (legacy comment), read by `fm_sheet_get_comment` /
 *        `fm_sheet_get_comment_at` and written back by
 *        `fm_sheet_set_comment`.
 *
 * `author` and `text` are model-backed UTF-8 views with the same lifetime
 * contract as `fm_hyperlink_t`; on input a `NULL` `author` or `text` means empty.
 */
typedef struct {
  uint32_t struct_size;
  /** The 0-based row of the cell the comment is anchored to; ignored on input, where the call's cell is the key. */
  uint32_t row;
  /** The 0-based column of the cell the comment is anchored to; ignored on input, where the call's cell is the key. */
  uint32_t col;
  /** The comment's author; `NULL` when none. */
  const char* author;
  /** The comment's text. */
  const char* text;
} fm_comment_t;

/**
 * @brief Sets the hyperlink covering exactly `hl->rect` on `sheet_index`.
 *
 * A hyperlink on that same rectangle is replaced in place (a repeated one
 * from a loaded file is dropped); otherwise `hl` is appended. Hyperlinks on
 * any other rectangle are kept, overlapping ones included.
 *
 * Every cell of `hl->rect` takes the built-in Hyperlink cell style: the font
 * is replaced by the Normal font underlined in theme colour 10; fill, number
 * format, border, alignment, protection and quote prefix are kept. Blank and
 * merge-covered cells are formatted; a rectangle spanning every row (or
 * column) formats the column (or row). An empty top-left cell receives
 * `display`, else the normalised address (`<address> - <fragment>` when a
 * fragment was split off), else `location`, as text; a cell that holds a
 * value keeps it.
 *
 * An http, https or ftp target without userinfo is stored the way Excel stores it: scheme and host lowercased, an
 * empty path becomes `/`, percent-escapes of unreserved characters decoded, and a `#fragment` split off into
 * `location` (a given `location` is kept). A mailto: target has only its scheme lowercased. A local path is held
 * absolute: an absolute POSIX path becomes a `file:///` URL, a drive path `file:///C:/...` and a UNC path
 * `file:////server/share/...`, each with a `#fragment` split off into `location` as for a URL; a POSIX path holding
 * `%` stays a plain path with only `%20` decoded to a space; a relative path is resolved against the folder of
 * `fm_workbook_get_location` when one is set (a `..` above the root keeps it as written). Any other target is stored
 * as written. With no `display`, a top-left cell that holds other text gets the stored target as `display`, and a
 * location-only link gets its location. A saved file writes a local target relative to the location's folder when
 * the two share at least one folder (compared case-sensitively), with that relative text as the saved `display`
 * (followed by ` - <location>` when there is one) unless the link has its own; a stray `%` in a plain path (one
 * not starting a `%XX` escape) is written `%25`. A space in an external target is written as `%20` in a saved file.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `hl` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `struct_size` is below the v1 size, `hl->rect` is
 *         inverted or outside the grid, or the text for the top-left cell is over 32,767 UTF-16 units;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a string is not well-formed UTF-8;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the rectangle covers more than 134,217,728 cells and is
 *         not a whole column or row (the engine's own ceiling, not an Excel limit). Nothing changes on
 *         failure.
 */
FM_API fm_status_t fm_sheet_set_hyperlink(fm_workbook_t* wb, size_t sheet_index, const fm_hyperlink_t* hl);

/**
 * @brief Inserts `hl` at list position `index` of the hyperlinks on
 *        `sheet_index`, replacing nothing (a hyperlink on the same rectangle
 *        is kept).
 *
 * `index` is a position in the order `fm_sheet_get_hyperlink_at` enumerates;
 * `index == count` appends. Edits the list only: no cell values or formats
 * are written or reset, so a caller undoing `fm_sheet_remove_hyperlink`
 * restores the cell formats itself. The target, location and display
 * are stored verbatim, with no normalisation. Pairs with
 * `fm_sheet_remove_hyperlink_at` for a lossless undo of the list.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `hl` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `struct_size` is below the v1 size, or the rectangle is
 *         inverted or outside the grid;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is above the count (nothing changes);
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a string is not well-formed UTF-8.
 */
FM_API fm_status_t fm_sheet_insert_hyperlink_at(fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                const fm_hyperlink_t* hl);

/**
 * @brief Removes every hyperlink on `sheet_index` whose rectangle intersects
 *        `range` (inclusive), each one whole: a hyperlink is not split or
 *        trimmed, and one only partly inside `range` goes entirely.
 *
 * Resets every cell of each removed hyperlink's rectangle to the default
 * format (cell format 0), including cells another remaining hyperlink still
 * covers; values stay. Formatting added before or after the hyperlink is
 * lost.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `range` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `range` is inverted or outside the grid;
 *         `FM_STATUS_NOT_FOUND` (6) when no hyperlink intersects `range`;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005), nothing removed, when a removed hyperlink covers more
 *         than 134,217,728 cells and is not a whole column or row; `fm_sheet_remove_hyperlink_at`
 *         removes such an entry without touching cell formats.
 */
FM_API fm_status_t fm_sheet_remove_hyperlink(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range);

/**
 * @brief Removes the hyperlink at `index` on `sheet_index`. `index` is a
 *        position in the order `fm_sheet_get_hyperlink_at` enumerates.
 *
 * Edits the list only: no cell values or formats are written or reset
 * (nothing else changes).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_hyperlink_at(fm_workbook_t* wb, size_t sheet_index, size_t index);

/**
 * @brief Drops every hyperlink on `sheet_index`. No-op when the sheet
 *        already has no hyperlinks.
 *
 * Resets every cell of each removed hyperlink's rectangle to the default
 * format (cell format 0), as `fm_sheet_remove_hyperlink` does; values stay.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005), nothing removed, when a hyperlink covers more than
 *         134,217,728 cells and is not a whole column or row; `fm_sheet_remove_hyperlink_at`
 *         removes such an entry without touching cell formats.
 */
FM_API fm_status_t fm_sheet_clear_hyperlinks(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Appends a merge range to the sheet's merge list and writes the
 *        stored rectangle to `*out_merge`.
 *
 * Covered cells keep their values and stay writable; clear them first to
 * match Excel's Merge. `out_merge` may be `NULL`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `merge` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the merge
 *         rectangle is inverted or outside the Excel grid;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the rectangle intersects a merge
 *         already on the sheet (an identical one included);
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when it intersects a table, a
 *         PivotTable, an array formula or a data table.
 */
FM_API fm_status_t fm_sheet_add_merge(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* merge,
                                      fm_cell_rect_t* out_merge);

/**
 * @brief Removes every merge range on `sheet_index` that overlaps `range`
 *        (inclusive rectangle intersection).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `range` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `range` is inverted or outside the grid;
 *         `FM_STATUS_NOT_FOUND` (6) when no merge overlaps `range`.
 */
FM_API fm_status_t fm_sheet_remove_merge(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range);

/**
 * @brief Removes the merge at `index` on `sheet_index`. Mirrors the
 *        index domain of fm_sheet_get_merge_at.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_remove_merge_at(fm_workbook_t* wb, size_t sheet_index, size_t index);

/**
 * @brief Drops every merge range on `sheet_index`. No-op when the sheet
 *        already has no merges.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_merges(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Reads the comment at `(row, col)` on `sheet_index`. `*out_present`
 *        is `1` and `*out` filled when a comment is anchored there, else `0`
 *        and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * `out->author` and `out->text` are model-backed views into the workbook
 * handle. Both are valid until the next mutation that touches the sheet's
 * comment list or until the handle is destroyed. Reads do not invalidate
 * them.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `(row, col)` is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_comment(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        int32_t* out_present, fm_comment_t* out);

/**
 * @brief Returns the number of comments attached to `sheet_index`, including
 *        comments anchored on cells that carry no value.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
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
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_comment_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                           fm_comment_t* out);

/**
 * @brief Reads the `index`-th hyperlink on `sheet_index` into `out`, in the
 *        order the hyperlinks were added (file order after a load).
 *
 * String fields in `out` are model-backed views (see `fm_hyperlink_t`).
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_hyperlink_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                             fm_hyperlink_t* out);

/**
 * @brief Reads the earliest-added hyperlink whose rectangle contains
 *        `(row, col)` on `sheet_index`. `*out_present` is `1` and `*out`
 *        filled when there is one, else `0` and `*out` zeroed (still
 *        `FM_STATUS_OK` (0)). "Earliest" is list order: the lowest position
 *        `fm_sheet_get_hyperlink_at` enumerates.
 *
 * String fields in `out` are model-backed views (see `fm_hyperlink_t`).
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `(row, col)` is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_hyperlink(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                          int32_t* out_present, fm_hyperlink_t* out);

/**
 * @brief Returns the number of hyperlinks attached to `sheet_index`, the
 *        positions `fm_sheet_get_hyperlink_at` enumerates.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_hyperlink_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `index`-th merge range on `sheet_index`, in storage order.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_merge_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                         fm_cell_rect_t* out);

/**
 * @brief Returns the number of merge ranges attached to `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_merge_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Copies every merge range on `sheet_index` that intersects `range` into
 *        the caller's buffer, in storage order.
 *
 * `*out_count` always receives the total number of intersecting merges; at
 * most `capacity` of them are written to `out`. Call with `capacity == 0`
 * (and `out` may then be `NULL`) to size the buffer first.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `range` or `out_count` is `NULL`, or `out`
 *         is `NULL` while `capacity > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `range` is inverted or outside the grid.
 */
FM_API fm_status_t fm_sheet_get_merges_in_range(const fm_workbook_t* wb, size_t sheet_index,
                                                const fm_cell_rect_t* range, fm_cell_rect_t* out, size_t capacity,
                                                size_t* out_count);

/**
 * @brief Sets or removes the comment at `(row, col)` on `sheet_index`.
 *        This family is Excel 365's Notes (legacy comments); Excel's
 *        comments are the `fm_sheet_*_threaded_comment` family.
 *
 * A non-`NULL` `comment` inserts or replaces the note; its `row` / `col` are
 * ignored, and a `NULL` or empty `text` writes a note with no text, as Excel
 * keeps an empty note. The strings are copied into the workbook handle's text
 * storage. `comment == NULL` removes the note and is a no-op when there is
 * none. Replacing a note keeps its `xr:uid`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the coordinate
 *         is outside the Excel grid;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `author` or `text` is malformed UTF-8;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the cell already holds a threaded
 *         comment (a cell carries a note or a thread, never both).
 */
FM_API fm_status_t fm_sheet_set_comment(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                        const fm_comment_t* comment);

/**
 * @brief Reads whether the note at `(row, col)` is shown without hovering
 *        (`visibility:visible` in the note's VML shape). `*out_present` is `1`
 *        and `*out_visible` set when a note is anchored there, else both are
 *        `0` (still `FM_STATUS_OK` (0)).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `(row, col)` is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_comment_visible(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                int32_t* out_present, int32_t* out_visible);

/**
 * @brief Shows (`visible != 0`) or hides the note at `(row, col)`. The note's
 *        box geometry is kept.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
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
  uint32_t struct_size;
  /** `1` when the workbook structure is locked, so sheets cannot be added, deleted, renamed, moved or hidden. */
  int32_t lock_structure;
  /** `1` when the workbook windows are locked, so they cannot be moved or resized. */
  int32_t lock_windows;
  /** `1` when the workbook is locked for revisions (shared-workbook tracking cannot be switched off). */
  int32_t lock_revision;
  /** The number of hash iterations of the structure-lock password; 0 when absent. */
  uint32_t workbook_spin_count;
  /** The number of hash iterations of the revision-lock password; 0 when absent. */
  uint32_t revisions_spin_count;
  /** The legacy 16-bit hash of the structure-lock password, as the hex text Excel writes; `NULL` when absent. */
  const char* workbook_legacy_password_hash;
  /** The name of the hash algorithm of the structure-lock password; `NULL` when absent. */
  const char* workbook_algorithm_name;
  /** The hash of the structure-lock password, stored as given; `NULL` when absent. */
  const char* workbook_hash_value;
  /** The salt of the structure-lock password hash, stored as given; `NULL` when absent. */
  const char* workbook_salt_value;
  /** The legacy 16-bit hash of the revision-lock password, as the hex text Excel writes; `NULL` when absent. */
  const char* revisions_legacy_password_hash;
  /** The name of the hash algorithm of the revision-lock password; `NULL` when absent. */
  const char* revisions_algorithm_name;
  /** The hash of the revision-lock password, stored as given; `NULL` when absent. */
  const char* revisions_hash_value;
  /** The salt of the revision-lock password hash, stored as given; `NULL` when absent. */
  const char* revisions_salt_value;
} fm_workbook_protection_t;

/**
 * @brief Reads the workbook's `<workbookProtection>`. `*out_present` is `1`
 *        and `*out` filled when the workbook carries the element (one with
 *        every lock off included), else `0` and `*out` zeroed (still
 *        `FM_STATUS_OK` (0)).
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_workbook_protection(const fm_workbook_t* wb, int32_t* out_present,
                                                       fm_workbook_protection_t* out);

/**
 * @brief Replaces the modelled attributes of `<workbookProtection>`, writing
 *        the element when the workbook has none. `protection == NULL` removes
 *        the element and is a no-op when there is none. Attributes the record
 *        does not model (the password character sets) are kept on a replace.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`.
 */
FM_API fm_status_t fm_workbook_set_workbook_protection(fm_workbook_t* wb, const fm_workbook_protection_t* protection);

/**
 * @brief One `<protectedRange>`: a named range users may edit on a protected
 *        sheet. Input handling, the `NULL` of an absent string on output and
 *        output lifetime follow `fm_workbook_protection_t`; the `ranges` of a
 *        getter result are read-scratch-backed.
 */
typedef struct {
  uint32_t struct_size;
  /** The number of hash iterations; 0 when absent. */
  uint32_t spin_count;
  /** The range's name, 1..255 UTF-16 units, unique among the sheet's protected ranges. */
  const char* name;
  /** The rectangles the protected range covers; a stored range that does not parse is left out. */
  const fm_cell_rect_t* ranges;
  /** The number of entries in `ranges`. */
  uint32_t range_count;
  /** The legacy 16-bit password hash, as the hex text Excel writes; `NULL` when absent. */
  const char* legacy_password_hash;
  /** The name of the hash algorithm that produced `hash_value`; `NULL` when absent. */
  const char* algorithm_name;
  /** The password hash, stored as given; `NULL` when absent. */
  const char* hash_value;
  /** The salt the password hash was computed with, stored as given; `NULL` when absent. */
  const char* salt_value;
  /**
   * The security descriptor (SDDL text) naming who may edit the range without a password; `NULL` when absent.
   * An `.xlsb` save of a range holding one fails with `FM_STATUS_IO_FEATURE_NOT_WRITABLE` (5028, context
   * `kind=protected_range_descriptor`).
   */
  const char* security_descriptor;
} fm_protected_range_t;

/**
 * @brief Partial update of one protected range
 *        (`fm_sheet_update_protected_range`), following the header's patch
 *        rules (Patches, above).
 *
 * Each field of `fm_protected_range_t` has its `<field>_op`. A `SET` string
 * follows the input rule of `fm_protected_range_t` (an empty string writes
 * the attribute absent).
 */
typedef struct {
  uint32_t struct_size;
  /** The operation applied to `spin_count`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t spin_count_op;
  /** The new number of hash iterations. */
  uint32_t spin_count;
  /** The operation applied to `name`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t name_op;
  /** The new name: 1..255 UTF-16 units, unused by the sheet's other protected ranges. */
  const char* name;
  /** The operation applied to `ranges`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t ranges_op;
  /** The new rectangles, replacing every stored one; at least one. Keep leaves the stored rectangles, including any
   *  token the reader could not parse. */
  const fm_cell_rect_t* ranges;
  /** The number of entries in `ranges`. */
  uint32_t range_count;
  /** The operation applied to `legacy_password_hash`: keep it, set it, or clear it. */
  fm_patch_op_t legacy_password_hash_op;
  /** The new legacy 16-bit password hash, as the hex text Excel writes. */
  const char* legacy_password_hash;
  /** The operation applied to `algorithm_name`: keep it, set it, or clear it. */
  fm_patch_op_t algorithm_name_op;
  /** The new name of the hash algorithm that produced `hash_value`. */
  const char* algorithm_name;
  /** The operation applied to `hash_value`: keep it, set it, or clear it. */
  fm_patch_op_t hash_value_op;
  /** The new password hash, stored as given. */
  const char* hash_value;
  /** The operation applied to `salt_value`: keep it, set it, or clear it. */
  fm_patch_op_t salt_value_op;
  /** The new salt the password hash was computed with. */
  const char* salt_value;
  /** The operation applied to `security_descriptor`: keep it, set it, or clear it. */
  fm_patch_op_t security_descriptor_op;
  /** The new security descriptor (SDDL text). */
  const char* security_descriptor;
} fm_protected_range_patch_t;

/**
 * @brief Number of protected ranges on `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_protected_range_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads protected range `protected_range_index` in file order.
 *
 * A range loaded from a file can carry a name longer than 255 UTF-16 units or more than 8191 rectangles; it reads
 * back as stored, `fm_sheet_add_protected_range` and `fm_sheet_update_protected_range` refuse those values, and an
 * `.xlsb` save of it fails with `FM_STATUS_IO_FEATURE_NOT_WRITABLE` (5028).
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `protected_range_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_protected_range_at(const fm_workbook_t* wb, size_t sheet_index,
                                                   size_t protected_range_index, fm_protected_range_t* out);

/**
 * @brief Reads the protected range named `name` (Unicode simple case folding)
 *        on `sheet_index`. `*out_present` is `1` and `*out` filled the way
 *        `fm_sheet_get_protected_range_at` fills it when there is one, else
 *        `0` and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * Lifetime: transient; the views in `*out` follow
 * `fm_sheet_get_protected_range_at`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_get_protected_range(const fm_workbook_t* wb, size_t sheet_index, const char* name,
                                                int32_t* out_present, fm_protected_range_t* out);

/**
 * @brief Appends a protected range and writes its stored `name`, the key the
 *        other calls take, to `*out_name`. The range follows row and column
 *        edits like any other sheet range and disappears when an edit deletes
 *        all of it.
 *
 * `out_name` may be `NULL`. `*out_name` is a model-backed view, valid until
 * the next mutation of the workbook.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `range` is `NULL`, or
 *           `range.ranges` is `NULL` while `range.range_count > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `name` is empty or longer than 255 UTF-16 units (context
 *         `part=name requested=<n> limit=255`), or when `range.range_count` is 0 or past 8,191 (context
 *         `part=ranges requested=<n> limit=8191`; up to the per-call cap below), or a rectangle is inverted
 *         or outside the grid;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when `range.range_count` is past
 *         the per-call range cap of 16,384 (context `part=ranges requested=<n> limit=16384`);
 *         `FM_STATUS_ALREADY_EXISTS` (7) when `name` already names another range
 *         of the sheet.
 */
FM_API fm_status_t fm_sheet_add_protected_range(fm_workbook_t* wb, size_t sheet_index,
                                                const fm_protected_range_t* range, const char** out_name);

/**
 * @brief Merges `patch` into the protected range named `name`. The whole
 *        patch is checked before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `name` or `patch` is `NULL`, or a `SET`
 *           `ranges` is `NULL` while `range_count > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits a
 *           field other than `legacy_password_hash`, `SET` carries a `NULL` string or
 *           a name that is empty or longer than 255 UTF-16 units (context
 *           `part=name requested=<n> limit=255`), or a `SET` `ranges` is empty,
 *           past 8,191 rectangles (context `part=ranges requested=<n> limit=8191`; up to the
 *           per-call cap below), or holds an inverted or out-of-grid rectangle;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when a `SET` `ranges` is past the
 *           per-call range cap of 16,384 (context `part=ranges requested=<n> limit=16384`);
 *         `FM_STATUS_NOT_FOUND` (6) when no range of the sheet has `name`;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when a `SET` name already names another
 *           range of the sheet;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` or a `SET` string is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_update_protected_range(fm_workbook_t* wb, size_t sheet_index, const char* name,
                                                   const fm_protected_range_patch_t* patch);

/**
 * @brief Removes the protected range named `name` on `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no range of the sheet has `name`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_remove_protected_range(fm_workbook_t* wb, size_t sheet_index, const char* name);

/** @brief Data-validation error style (`fm_data_validation_t::error_style`,
 * `fm_validation_outcome_info_t::error_style`). */
typedef enum {
  FM_VALIDATION_ERROR_STYLE_STOP = 0,
  FM_VALIDATION_ERROR_STYLE_WARNING = 1,
  FM_VALIDATION_ERROR_STYLE_INFORMATION = 2,
} fm_validation_error_style_t;

/** @brief Data-validation rule type (`fm_data_validation_t::type`), OOXML `ST_DataValidationType` order. */
typedef enum {
  FM_VALIDATION_TYPE_NONE = 0,
  FM_VALIDATION_TYPE_WHOLE = 1,
  FM_VALIDATION_TYPE_DECIMAL = 2,
  FM_VALIDATION_TYPE_LIST = 3,
  FM_VALIDATION_TYPE_DATE = 4,
  FM_VALIDATION_TYPE_TIME = 5,
  FM_VALIDATION_TYPE_TEXT_LENGTH = 6,
  FM_VALIDATION_TYPE_CUSTOM = 7,
} fm_validation_type_t;

/**
 * @brief Relational comparison shared by data validation
 *        (`fm_data_validation_t::op`), conditional formatting `cellIs`
 *        (`fm_cf_rule_t::op`) and AutoFilter custom conditions
 *        (`fm_filter_condition_t::op`).
 *
 * OOXML `ST_DataValidationOperator` order, the one schema order covering all
 * eight; zero is its default `between`. An AutoFilter condition always states
 * its operator and takes only the six non-range members.
 */
typedef enum {
  FM_COMPARISON_OPERATOR_BETWEEN = 0,
  FM_COMPARISON_OPERATOR_NOT_BETWEEN = 1,
  FM_COMPARISON_OPERATOR_EQUAL = 2,
  FM_COMPARISON_OPERATOR_NOT_EQUAL = 3,
  FM_COMPARISON_OPERATOR_LESS_THAN = 4,
  FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL = 5,
  FM_COMPARISON_OPERATOR_GREATER_THAN = 6,
  FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL = 7,
} fm_comparison_operator_t;

/**
 * @brief Data-validation rule descriptor.
 *
 * Mirrors `formulon::DataValidation`. On getter results `formula1`, `formula2`,
 * `error_title`, `error_message`, `prompt_title` and `prompt_message` are views
 * (see `fm_sheet_get_validation_at` for their lifetime); on the add path
 * (`fm_sheet_add_validation`) they are caller-owned input strings, where `NULL`
 * and the empty string both mean absent. A formula reads back with a leading
 * `=`; input takes it with or without one. The tristate fields read back as
 * `FM_TRISTATE_FALSE` or `FM_TRISTATE_TRUE`.
 */
typedef struct {
  uint32_t struct_size;
  /** The rule's engine id: an opaque string, unique on its sheet, stable while the workbook is open across structural
   *  edits and never reused within the session; assigned again on load or sheet copy and not saved. A
   *  read-scratch-backed view on getter results; ignored on the add path, which assigns a new id. */
  const char* id;
  /** The cell ranges the rule applies to. On getter results a model-backed view (lifetime: see
   *  `fm_sheet_get_validation_at`); on `fm_sheet_add_validation` the caller owns the buffer, which need live only for
   *  the call, and every range is deep-copied. */
  const fm_cell_rect_t* ranges;
  /** The number of entries in `ranges`. */
  uint32_t range_count;
  /** The kind of value the rule allows, an `fm_validation_type_t`. */
  int32_t type;
  /** The comparison the rule applies between the entered value and its formulas, an `fm_comparison_operator_t`. */
  int32_t op;
  /** How an invalid entry is reported: stop, warning or information, an `fm_validation_error_style_t`. */
  int32_t error_style;
  /** Tristate: whether a blank cell passes the rule; default is yes. */
  int32_t allow_blank; /* fm_tristate_t */
  /** Tristate: whether the prompt (`prompt_title` and `prompt_message`) shows when the cell is selected; default is
   *  yes. */
  int32_t show_input_message; /* fm_tristate_t */
  /** Tristate: whether the error alert (`error_title` and `error_message`) shows on an invalid entry; default is yes.
   */
  int32_t show_error_message; /* fm_tristate_t */
  /** Tristate: whether the in-cell dropdown arrow is shown (`list` rules only); default is yes. The inverse of the raw
   *  OOXML `showDropDown` attribute, which the reader and writer translate. */
  int32_t show_in_cell_dropdown; /* fm_tristate_t */
  /** The first formula or value of the rule, leading `=` included; `NULL` when the rule states none. */
  const char* formula1;
  /** The second formula or value of the rule (the upper bound of a between rule), leading `=` included; `NULL` when the
   *  rule states none. */
  const char* formula2;
  /** The title of the error alert; `NULL` when the rule has none. */
  const char* error_title;
  /** The text of the error alert; `NULL` when the rule has none. */
  const char* error_message;
  /** The title of the input prompt; `NULL` when the rule has none. */
  const char* prompt_title;
  /** The text of the input prompt; `NULL` when the rule has none. */
  const char* prompt_message;
} fm_data_validation_t;

/**
 * @brief Returns the number of data-validation rules attached to `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_validation_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `validation_index`-th data-validation rule on `sheet_index` into `out`.
 *
 * `ranges` and the four message strings in `*out` are model-backed views
 * owned by the workbook handle. They remain valid until the next mutation
 * that touches the sheet's validation list (any of
 * `fm_sheet_add_validation`, `fm_sheet_update_validation`,
 * `fm_sheet_remove_validation`, `fm_sheet_clear_validations`, or any
 * reader / writer round-trip), or
 * until the handle is destroyed. Callers that need to retain the
 * payload across mutations must copy it. `formula1` / `formula2` are
 * spelled as `fm_sheet_get_formula` spells formulas, leading `=` included
 * (an absent formula is `NULL`); a present one and `id` are
 * read-scratch-backed, valid until the next scratch-backed read.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `validation_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_validation_at(const fm_workbook_t* wb, size_t sheet_index, size_t validation_index,
                                              fm_data_validation_t* out);

/**
 * @brief Reads the data-validation rule whose id is `id` on `sheet_index`.
 *        `*out_present` is `1` and `*out` filled when there is one, else `0`
 *        and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * The views in `out` have the lifetime `fm_sheet_get_validation_at`
 * describes.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `id` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_get_validation(const fm_workbook_t* wb, size_t sheet_index, const char* id,
                                           int32_t* out_present, fm_data_validation_t* out);

/**
 * @brief Appends a data-validation rule to the sheet's validation list and
 *        writes its new id to `*out_validation_id`.
 *
 * `*out_validation_id` is a read-scratch-backed view, valid until the next
 * scratch-backed read. `v.id` is ignored. Every other string field accepts
 * `NULL`, which is treated as the empty string. `v.ranges` must point to
 * `v.range_count >= 1` consecutive
 * `fm_cell_rect_t` values that the implementation copies into the
 * rule's owned storage. Every rectangle must be normalised
 * (`first <= last`) and lie within the Excel grid; an inverted or
 * out-of-grid one is `FM_STATUS_INVALID_ARGUMENT`. All ranges are
 * validated before the rule is stored.
 * `v.formula1` / `v.formula2` may carry one leading `=`, which is dropped:
 * the rule stores them as files carry them, spelled as
 * `fm_sheet_set_formula` stores a cell formula.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `v` or `out_validation_id` is `NULL`, or `v.ranges` is
 *           `NULL` while `v.range_count > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `v.range_count` is 0, when
 *         any supplied range is inverted or outside the Excel grid, or when
 *         `v.type`, `v.op`, `v.error_style`, `v.allow_blank` or
 *         `v.show_in_cell_dropdown` is outside the domain listed above;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when `v.range_count` exceeds the
 *         per-call range cap of 16,384 (context `part=ranges requested=<n> limit=16384`);
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position) when a non-empty
 * `v.formula1` or `v.formula2` does not parse (the rule is not stored).
 */
FM_API fm_status_t fm_sheet_add_validation(fm_workbook_t* wb, size_t sheet_index, const fm_data_validation_t* v,
                                           const char** out_validation_id);

/**
 * @brief Partial update of one data-validation rule
 *        (`fm_sheet_update_validation`), following the header's patch rules
 *        (Patches, above).
 *
 * Each field of `fm_data_validation_t` but `id` has its `<field>_op`, with the
 * domains `fm_sheet_add_validation` checks. A `SET` string follows the add
 * path's input rule.
 */
typedef struct {
  uint32_t struct_size;
  /** The operation applied to `ranges`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t ranges_op;
  /** The new cell ranges, replacing every stored one; at least one. */
  const fm_cell_rect_t* ranges;
  /** The number of entries in `ranges`. */
  uint32_t range_count;
  /** The operation applied to `type`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t type_op;
  /** The new kind of value the rule allows. */
  int32_t type;
  /** The operation applied to `op`: keep it, or set it; it cannot be cleared. */
  fm_patch_op_t op_op;
  /** The new comparison between the entered value and the formulas. */
  int32_t op;
  /** The operation applied to `error_style`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t error_style_op;
  /** The new error style: stop, warning or information. */
  int32_t error_style;
  /** The operation applied to `allow_blank`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t allow_blank_op;
  /** Tristate: whether a blank cell passes the rule. */
  int32_t allow_blank; /* fm_tristate_t */
  /** The operation applied to `show_input_message`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t show_input_message_op;
  /** Tristate: whether the prompt shows when the cell is selected. */
  int32_t show_input_message; /* fm_tristate_t */
  /** The operation applied to `show_error_message`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t show_error_message_op;
  /** Tristate: whether the error alert shows on an invalid entry. */
  int32_t show_error_message; /* fm_tristate_t */
  /** The operation applied to `show_in_cell_dropdown`: keep it, set it; it cannot be cleared. */
  fm_patch_op_t show_in_cell_dropdown_op;
  /** Tristate: whether the in-cell dropdown arrow is shown. */
  int32_t show_in_cell_dropdown; /* fm_tristate_t */
  /** The operation applied to `formula1`: keep it, set it, or clear it. */
  fm_patch_op_t formula1_op;
  /** The new first formula or value, with or without a leading `=`. */
  const char* formula1;
  /** The operation applied to `formula2`: keep it, set it, or clear it. */
  fm_patch_op_t formula2_op;
  /** The new second formula or value, with or without a leading `=`. */
  const char* formula2;
  /** The operation applied to `error_title`: keep it, set it, or clear it. */
  fm_patch_op_t error_title_op;
  /** The new title of the error alert. */
  const char* error_title;
  /** The operation applied to `error_message`: keep it, set it, or clear it. */
  fm_patch_op_t error_message_op;
  /** The new text of the error alert. */
  const char* error_message;
  /** The operation applied to `prompt_title`: keep it, set it, or clear it. */
  fm_patch_op_t prompt_title_op;
  /** The new title of the input prompt. */
  const char* prompt_title;
  /** The operation applied to `prompt_message`: keep it, set it, or clear it. */
  fm_patch_op_t prompt_message_op;
  /** The new text of the input prompt. */
  const char* prompt_message;
} fm_data_validation_patch_t;

/**
 * @brief Merges `patch` into the validation rule with `id`. The whole patch is
 *        checked before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `id` or `patch` is `NULL`, or a `SET` `ranges` is
 *           `NULL` while `range_count > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits a field
 *           that cannot be cleared, `SET` carries a `NULL` string, a `SET` enumeration is outside
 *           its domain, or a `SET` `ranges` is empty or holds an inverted or out-of-grid
 *           rectangle;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when a `SET` `ranges` is past the
 *           per-call range cap of 16,384 (context `part=ranges requested=<n> limit=16384`);
 *         `FM_STATUS_NOT_FOUND` (6) when no rule has `id`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `id` or a `SET` string is not valid UTF-8;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position) when a
 *           `SET` formula does not parse.
 */
FM_API fm_status_t fm_sheet_update_validation(fm_workbook_t* wb, size_t sheet_index, const char* id,
                                              const fm_data_validation_patch_t* patch);

/**
 * @brief Removes the validation rule with `id` on `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no rule has `id`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `id` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_remove_validation(fm_workbook_t* wb, size_t sheet_index, const char* id);

/**
 * @brief Drops every validation rule on `sheet_index`. No-op when the sheet
 *        already has no rules.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_validations(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Result of checking a value against the data-validation rule that
 *        covers a cell.
 */
typedef struct {
  uint32_t struct_size;
  /** 1 when a rule covers the cell; 0 means none, and `valid` is then 1. */
  int32_t has_rule;
  /** 1 when the value passes the rule, or when no rule covers the cell. */
  int32_t valid;
  /** The applied rule's error style, an `fm_validation_error_style_t`; 0 (stop) when no rule covers the cell. */
  int32_t error_style;
  /** The id of the rule that was applied (`fm_data_validation_t::id`), a read-scratch-backed view; `NULL` when no rule
   *  covers the cell. */
  const char* rule_id;
} fm_validation_outcome_info_t;

/**
 * @brief Checks `proposed` against the data-validation rule covering
 *        `(row, col)`, without storing it.
 *
 * Excel applies one rule per cell: when sqrefs overlap, the first rule in
 * `fm_sheet_get_validation_at` order wins. `proposed` is taken as the
 * already-parsed value (text "5" is text, not a number). The rule's
 * formulas, including a `custom` formula that refers to the cell itself,
 * read the workbook as stored. `out->rule_id` is a read-scratch-backed view,
 * valid until the next scratch-backed read.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`, or a
 *         `FM_VALUE_KIND_TEXT` value carries a `NULL` text;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the cell is
 *         outside the grid, the number is not finite, the error code is out
 *         of range, or the value kind is an array, reference or lambda.
 */
FM_API fm_status_t fm_sheet_validate_value(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                           const fm_value_t* proposed, fm_validation_outcome_info_t* out);

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
 * Lifetime: owned (until fm_cell_range_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
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
 * A read fills both offset pairs; a write reads the pair its
 * `fm_text_offset_unit_t` argument names and ignores the other.
 * `end < start` is `FM_STATUS_INVALID_ARGUMENT` (2). Frozen value type: no
 * `struct_size`.
 */
typedef struct {
  /** The id of the mentioned person. */
  const char* person_id;
  /** The mention's own id, a brace-wrapped GUID; on input `NULL` or empty takes an engine-generated GUID, as
   *  `fm_threaded_comment_t.id` does. */
  const char* mention_id;
  /** The start of the mentioned span in the comment text, as a UTF-16 code unit offset (inclusive). */
  uint32_t start;
  /** The end of the mentioned span in the comment text, as a UTF-16 code unit offset (exclusive). */
  uint32_t end;
  /** The start of the mentioned span in the comment text, as a code point offset (inclusive). */
  uint32_t start_cp;
  /** The end of the mentioned span in the comment text, as a code point offset (exclusive). */
  uint32_t end_cp;
} fm_mention_t;

/**
 * @brief One threaded comment: a thread's opening comment, or a reply when
 *        `parent_id` names its thread.
 *
 * Ids are brace-wrapped GUIDs, taken in either letter case, matched across
 * it, and stored uppercase as Excel writes them. On input every other
 * `const char*` may be `NULL`, meaning empty.
 */
typedef struct {
  uint32_t struct_size;
  /** The comment's id, a brace-wrapped GUID stored in upper case; on add a `NULL` or empty id takes an engine-generated
   *  GUID (deterministic with a pinned clock). */
  const char* id;
  /** The 0-based row of the cell the thread is anchored to. */
  uint32_t row;
  /** The 0-based column of the cell the thread is anchored to. */
  uint32_t col;
  /** The id of the person who wrote the comment. */
  const char* person_id;
  /** The creation time as UTC text `YYYY-MM-DDTHH:MM:SS.ss` with no zone designator; `NULL` when a loaded comment
   *  states none. On add a trailing `Z` is accepted and dropped, a numeric offset (`+09:00`) or a three-digit fraction
   *  is `FM_STATUS_INVALID_ARGUMENT`, and `NULL` or empty takes the current UTC time, or the pinned time written as-is
   *  when a clock is pinned (`fm_workbook_set_pinned_now`). */
  const char* created;
  /** The comment's text. */
  const char* text;
  /** The id of the thread's opening comment this comment replies to; `NULL` for an opening comment; on input `NULL`
   * opens a thread and an empty string is `FM_STATUS_INVALID_ARGUMENT` (2). */
  const char* parent_id;
  /** `1` when the thread is resolved; meaningful on an opening comment only. */
  int32_t done;
  /** The people mentioned in the text; `NULL` only with a zero `mentions_count`. */
  const fm_mention_t* mentions;
  /** The number of entries in `mentions`. */
  uint32_t mentions_count;
} fm_threaded_comment_t;

/**
 * @brief One entry of the workbook's person list. On input `user_id` and
 *        `provider_id` take `NULL` and `""` alike as absent.
 */
typedef struct {
  uint32_t struct_size;
  /** The person's id; on add a `NULL` or empty id takes an engine-generated GUID, as `fm_threaded_comment_t.id` does
   *  (deterministic with the clock pinned); a read always carries it. */
  const char* id;
  /** The person's display name; required. */
  const char* display_name;
  /** The user id; `NULL` when the person has none. */
  const char* user_id;
  /** The identity provider id; `NULL` when the person has none. */
  const char* provider_id;
} fm_person_t;

/**
 * @brief Number of threaded comments on the sheet, replies included.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_threaded_comment_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads comment `index` of the sheet's flat list, in file order: each
 *        thread's opening comment followed by its replies.
 *
 * Every string and the `mentions` array are read-scratch-backed views, with
 * the lifetime `fm_sheet_get_auto_filter` describes.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_threaded_comment_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                    fm_threaded_comment_t* out);

/**
 * @brief Reads the threaded comment whose id is `id` on `sheet_index`.
 *        `*out_present` is `1` and `*out` filled when there is one, else `0`
 *        and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * The views in `out` have the lifetime `fm_sheet_get_threaded_comment_at`
 * describes.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `id` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_get_threaded_comment(const fm_workbook_t* wb, size_t sheet_index, const char* id,
                                                 int32_t* out_present, fm_threaded_comment_t* out);

/**
 * @brief Adds a threaded comment. A `NULL` `parent_id` opens a new
 *        thread at `row`/`col`, which must hold neither a note nor a thread;
 *        any other one appends a reply to that thread, taking its anchor and
 *        ignoring `row`, `col` and `done`. The comment's id, given or
 *        generated (`fm_threaded_comment_t`), is written to `*out_id`, a
 *        read-scratch-backed view valid until the next scratch-backed read.
 *        On any failure the sheet is unchanged. `offset_unit` names the pair
 *        of offsets each mention supplies, counted against `comment->text`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `comment` or `out_id` is `NULL`, or
 *         `comment->mentions` is `NULL` with a non-zero count;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `offset_unit` is not a `fm_text_offset_unit_t`
 *         member, a new thread's cell is outside the
 *         grid, the id or a mention id is not a GUID, `created` is malformed
 *         (including a numeric offset), a
 *         mention ends before it starts or lies outside the text, `parent_id`
 *         is empty, or `parent_id` names a reply;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the id is already used on the
 *         sheet, or a new thread's cell already carries a note or thread;
 *         `FM_STATUS_NOT_FOUND` (6) when `comment->person_id` names no person of the workbook
 *         or `parent_id` no comment of the sheet.
 */
FM_API fm_status_t fm_sheet_add_threaded_comment(fm_workbook_t* wb, size_t sheet_index, int32_t offset_unit,
                                                 const fm_threaded_comment_t* comment, const char** out_id);

/**
 * @brief Partial update of one threaded comment
 *        (`fm_sheet_update_threaded_comment`), following the header's patch
 *        rules (Patches, above). No field can be cleared.
 */
typedef struct {
  uint32_t struct_size;
  /** The operation applied to `text`: keep it, or set it; it cannot be cleared. */
  fm_patch_op_t text_op;
  /** The new comment text; omitted leaves it unchanged, and an empty string empties it. */
  const char* text;
  /** The operation applied to `mentions`: keep it, or set it; it cannot be cleared. */
  fm_patch_op_t mentions_op;
  /** The new mention list, replacing the whole list (`NULL` with a zero count empties it; a `NULL` or empty
   *  `mention_id` is generated); omitted leaves it unchanged. */
  const fm_mention_t* mentions;
  /** The number of entries in `mentions`. */
  uint32_t mentions_count;
  /** The operation applied to `done`: keep it, or set it; it cannot be cleared. */
  fm_patch_op_t done_op;
  /** Resolves (`1`) or reopens (`0`) the thread; accepted only on an opening comment. */
  int32_t done;
} fm_threaded_comment_patch_t;

/**
 * @brief Merges `patch` into the comment with `id`. The kept text and
 *        mentions are validated together with the replaced ones; on any
 *        failure the sheet is unchanged. `offset_unit` names the pair of
 *        offsets each replaced mention supplies, counted against the new text,
 *        or the kept text when the patch does not set one.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `id` or `patch` is `NULL`, or a `SET`
 *           `mentions` is `NULL` with a non-zero count;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `offset_unit` is not a `fm_text_offset_unit_t`
 *           member, an op is outside 0..2 or `CLEAR`,
 *           `SET` carries a `NULL` `text`, a `SET` `done` is not `0` or `1`, a
 *           mention id is not a GUID, a mention ends before it starts or lies
 *           outside the text, or `done` is set on a reply;
 *         `FM_STATUS_NOT_FOUND` (6) when no comment has `id`.
 */
FM_API fm_status_t fm_sheet_update_threaded_comment(fm_workbook_t* wb, size_t sheet_index, const char* id,
                                                    int32_t offset_unit, const fm_threaded_comment_patch_t* patch);

/**
 * @brief Removes the comment with `id`; removing an opening comment removes
 *        its whole thread.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no comment has `id`.
 */
FM_API fm_status_t fm_sheet_remove_threaded_comment(fm_workbook_t* wb, size_t sheet_index, const char* id);

/**
 * @brief Number of entries in the workbook's person list.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_person_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads person `index`, in file order. The strings are
 *        read-scratch-backed views.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_person_at(const fm_workbook_t* wb, size_t index, fm_person_t* out);

/**
 * @brief Reads the person whose id is `id`. `*out_present` is `1` and `*out`
 *        filled when there is one, else `0` and `*out` zeroed (still
 *        `FM_STATUS_OK` (0)). The strings are read-scratch-backed views.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `id` is not valid UTF-8.
 */
FM_API fm_status_t fm_workbook_get_person(const fm_workbook_t* wb, const char* id, int32_t* out_present,
                                          fm_person_t* out);

/**
 * @brief Appends a person to the workbook's person list and writes the id as
 *        stored to `*out_person_id`. A given id is a brace-wrapped GUID in
 *        either letter case, stored uppercase; a `NULL` or empty one is
 *        generated (`fm_person_t`). Person ids match across letter case.
 *
 * `*out_person_id` is a read-scratch-backed view, valid until the next
 * scratch-backed read.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a given id is not a GUID or the
 *         display name is `NULL` or empty;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the id is already listed.
 */
FM_API fm_status_t fm_workbook_add_person(fm_workbook_t* wb, const fm_person_t* person, const char** out_person_id);

/**
 * @brief Removes the person with `id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_NOT_FOUND` (6) when no person has `id`;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when a threaded comment or mention
 *         still names them.
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
 *         an `FM_STATUS_EVAL_*` or `FM_STATUS_GRAPH_*` code on engine failure.
 */
FM_API fm_status_t fm_workbook_recalc(fm_workbook_t* wb);

/**
 * @brief Writes `1` to `*out_has_dirty` when the workbook holds a change that
 *        no full recalc (`fm_workbook_recalc` / `fm_workbook_recalc_parallel`)
 *        has applied yet, else `0`.
 *
 * Cell reads report the values of the last full recalc, so a `1` means some
 * of them may be stale. A cell is dirty while a change has left its value
 * uncomputed. A change is a cell write, a defined-name edit that a
 * formula reads, a row / column insert or delete, a sheet removal, a file
 * name change, or a profile change on a workbook with formulas; setting the
 * current profile or file name to its current value is not. Volatile cells (`NOW()`, `RAND()`) do not count: they
 * re-evaluate on every recalc, and a workbook holding only them reads `0`
 * after a recalc. A freshly created workbook reads `0`. Under manual
 * calculation a loaded workbook holds dirty only the formulas its file marks
 * for recalculation and those without a cached value; under automatic
 * calculation, every formula until the first recalc.
 * `fm_sheet_recalc_viewport` leaves the cells it skipped dirty, so the
 * answer stays `1` after it; a failed full recalc keeps every dirty cell.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_GRAPH_RECALC_REENTRANT` (4005) when called from inside a recalc on
 *         the same thread (an iterative-progress callback).
 */
FM_API fm_status_t fm_workbook_has_dirty_cells(const fm_workbook_t* wb, int32_t* out_has_dirty);

/**
 * @brief Returns the number of cells on `sheet_index` that sit in an unresolved
 *        circular reference, the positions `fm_sheet_get_circular_reference_at`
 *        enumerates.
 *
 * A cell counts when it is a member of a cycle that reads back into itself and
 * that the last evaluation of the cycle left unresolved; a cell that only reads
 * a cycle does not. A cycle across sheets counts on each sheet for that sheet's
 * members. While iterative calculation is enabled the count is 0, as Excel then
 * reports no circular reference. Under automatic calculation the read first
 * computes dirty cells, as value reads do; under manual calculation it reports
 * the last recalc. Read it when a host shows the list, not on every render.
 *
 * Excel's `Worksheet.CircularReference` names one member per sheet, and which
 * one is not determined by the workbook, so every member is reported.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_GRAPH_RECALC_REENTRANT` (4005) when called from inside a recalc on
 *         the same thread (an iterative-progress callback).
 */
FM_API fm_status_t fm_sheet_get_circular_reference_count(const fm_workbook_t* wb, size_t sheet_index,
                                                         size_t* out_count);

/**
 * @brief Reads the `index`-th cell of `fm_sheet_get_circular_reference_count`,
 *        in row-major order (row, then column).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range;
 *         `FM_STATUS_GRAPH_RECALC_REENTRANT` (4005) when called from inside a recalc on
 *         the same thread (an iterative-progress callback).
 */
FM_API fm_status_t fm_sheet_get_circular_reference_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                      fm_cell_address_t* out);

/**
 * @brief Writes `1` to `*out_stale` when the cell at (`row`, `col`) holds a
 *        formula whose value has not been computed since its inputs changed,
 *        under manual or automatic-except-tables calculation, else `0`.
 *
 * The cell's value is the one last computed; this reports only that it is out
 * of date. Under automatic calculation it is always `0`. A cell without a
 * formula, an empty cell and a member of a spill reads `0`; query the spill's
 * anchor. A cell is out of date while it, or a cell it reads directly or
 * through other cells, carries a change no recalc has applied: an edit, a
 * row / column insert or delete, a defined-name edit. A loaded workbook
 * reports the formulas its file marks for recalculation (ca) as stale. While
 * any such change is pending,
 * every volatile cell (NOW, RAND, OFFSET, INDIRECT and the like) and every
 * cell that reads one is out of date too, as Excel marks them. The read never
 * recalculates. Hosts can use it to draw stale-value formatting.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row` or `col` is outside the grid;
 *         `FM_STATUS_GRAPH_RECALC_REENTRANT` (4005) when called from inside a recalc on
 *         the same thread (an iterative-progress callback).
 */
FM_API fm_status_t fm_sheet_is_cell_stale(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                          int32_t* out_stale);

/** @brief Per-call telemetry returned by `fm_workbook_recalc_parallel`. */
typedef struct {
  uint32_t struct_size;
  /** The number of cells whose formula was executed. */
  uint64_t cells_evaluated;
  /** The number of dirty strongly connected components the scheduler processed. */
  uint64_t sccs_processed;
  /** The number of layers dispatched to the worker pool. */
  uint64_t parallel_steps;
  /** The number of serial batches processed on the calling thread, including the dynamic-reference tail of a pooled
   *  layer. */
  uint64_t serial_fallback_steps;
  /** The number of cyclic components resolved by the iterative solver, not counting those that fell back to `#REF!`. */
  uint64_t cycle_recoveries;
  /** The number of worker threads successfully started for this call; 0 when everything ran on the calling thread. */
  uint32_t worker_threads_started;
  /** The number of started worker threads that claimed at least one task. */
  uint32_t worker_threads_used;
} fm_parallel_recalc_stats_info_t;

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
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when `thread_count > 8` (context
 *           `part=threads requested=<n> limit=8`);
 *         `FM_STATUS_GRAPH_RECALC_REENTRANT` (4005) when called recursively from a recalc
 *         callback (including a scheduler worker);
 *         an `FM_STATUS_EVAL_*` or `FM_STATUS_GRAPH_*` code on engine failure.
 */
FM_API fm_status_t fm_workbook_recalc_parallel(fm_workbook_t* wb, uint32_t thread_count,
                                               fm_parallel_recalc_stats_info_t* out_stats);

/**
 * @brief The iterative-calculation settings stored on a workbook
 *        (`fm_workbook_get_iterative`).
 */
typedef struct {
  uint32_t struct_size;
  /** `1` when iterative calculation is on; off by default, when cyclic components surface `#REF!`. */
  int32_t enabled;
  /** The iteration cap, 1 to 32767; meaningful even while `enabled` is `0`. */
  uint32_t max_iterations;
  /** The absolute convergence threshold, finite and non-negative; meaningful even while `enabled` is `0`. */
  double max_change;
} fm_iterative_settings_info_t;

/**
 * @brief A partial update of the iterative-calculation settings
 *        (`fm_workbook_update_iterative`), following the header's patch rules
 *        (Patches, above). No field can be cleared.
 */
typedef struct {
  uint32_t struct_size;
  /** What to do with `enabled`: keep, or set (clear is not supported). */
  fm_patch_op_t enabled_op;
  /** With a set operation, 0 disables iterative calculation and any other value enables it. */
  int32_t enabled;
  /** What to do with `max_iterations`: keep, or set (clear is not supported). */
  fm_patch_op_t max_iterations_op;
  /** With a set operation, the iteration cap, 1 to 32767; 0 or above 32767 is an invalid argument. */
  uint32_t max_iterations;
  /** What to do with `max_change`: keep, or set (clear is not supported). */
  fm_patch_op_t max_change_op;
  /** With a set operation, the absolute convergence threshold, finite and non-negative. */
  double max_change;
} fm_iterative_settings_patch_t;

/**
 * @brief Merges `patch` into the iterative-calculation settings of the
 *        workbook's recalc engine. The whole patch is checked before anything
 *        changes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2 or `CLEAR`, a `SET`
 *           `max_iterations` is 0 or above 32767, or a `SET` `max_change` is negative, NaN or infinite.
 */
FM_API fm_status_t fm_workbook_update_iterative(fm_workbook_t* wb, const fm_iterative_settings_patch_t* patch);

/**
 * @brief Reads the iterative-calculation settings currently stored on a workbook.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_iterative(const fm_workbook_t* wb, fm_iterative_settings_info_t* out);

/* -------------------------------------------------------------------------- */
/* Calculation mode (workbook-level `<calcPr calcMode>` policy)               */
/* -------------------------------------------------------------------------- */

/**
 * @brief Workbook-level calculation mode.
 *
 * Mirrors Excel's "Calculation options" workbook setting and the
 * `calcMode` attribute on the `<calcPr>` element.
 *   * `FM_CALC_MODE_AUTO` (the default): dirty cells are calculated before
 *     a value is read and before the workbook is saved.
 *   * `FM_CALC_MODE_MANUAL`: nothing is calculated until `fm_workbook_recalc`.
 *   * `FM_CALC_MODE_AUTO_NO_TABLE`: as Auto, but data tables are skipped.
 *
 * Volatile functions (`NOW()`, `RAND()`) are re-evaluated when an edit or a
 * `fm_workbook_recalc` recalculates them, not on every read.
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
 * Decides whether value reads and saves calculate dirty cells first (see
 * `fm_calc_mode_t`). The value is a raw signed
 * 32-bit enum ordinal so FFI callers can pass an arbitrary value without
 * first constructing an out-of-domain C enum. Unknown values (outside the
 * documented range) return `FM_STATUS_INVALID_ARGUMENT` (2).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `mode` is not a documented value.
 */
FM_API fm_status_t fm_workbook_set_calc_mode(fm_workbook_t* wb, int32_t mode);

/**
 * @brief Excel's Compatibility Version of a workbook.
 *
 * Version 1 makes LEN, MID, FIND, SEARCH and REPLACE count UTF-16 units;
 * version 2 counts code points (a supplementary-plane character is one
 * position); version 3 adds nested arrays, which Formulon does not implement.
 * The ordinals are the values stored in the file, so they start at 1.
 */
typedef enum {
  FM_COMPATIBILITY_VERSION_1 = 1,
  FM_COMPATIBILITY_VERSION_2 = 2,
  FM_COMPATIBILITY_VERSION_3 = 3,
} fm_compatibility_version_t;

/**
 * @brief Returns the Compatibility Version the workbook evaluates under.
 *
 * The workbook's stored value. A new workbook reads 2, and a loaded workbook
 * whose file carries none reads 1, as Excel computes it. 3 evaluates the five
 * functions as 2. A stored value above 3, written by a later Excel, reads as
 * stored (saturating at INT32_MAX), evaluates as `FM_COMPATIBILITY_VERSION_2`
 * and is written back as read.
 *
 * Under version 1, Excel's MID and REPLACE can return a lone surrogate when a
 * cut falls inside a supplementary-plane character; Formulon takes the whole
 * character instead (text is UTF-8).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_version` is NULL.
 */
FM_API fm_status_t fm_workbook_get_compatibility_version(const fm_workbook_t* wb, int32_t* out_version);

/**
 * @brief Sets the workbook's Compatibility Version.
 *
 * Setting the current value changes nothing. Another value is stored and
 * written on save; 3 is stored and written as given and evaluates as 2. A
 * change marks every formula cell dirty only when it moves between 1 and 2 or
 * 3, which changes how the five functions count: in Auto mode the next read
 * recomputes, in Manual mode `fm_workbook_recalc` does. A change between 2 and
 * 3 marks nothing.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is NULL;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `version` is not 1, 2 or 3
 *         (a larger value read back from the getter included); the workbook
 *         is left unchanged.
 */
FM_API fm_status_t fm_workbook_set_compatibility_version(fm_workbook_t* wb, int32_t version);

/** @brief `<calcPr refMode>`: the reference notation the host shows formulas in. */
typedef enum {
  FM_REF_MODE_A1 = 0,
  FM_REF_MODE_R1C1 = 1,
} fm_ref_mode_t;

/**
 * @brief Reads `<calcPr fullPrecision>`: 1 for full precision (the default),
 *        0 for "set precision as displayed".
 *
 * Plain metadata: the engine evaluates at full precision whatever this
 * holds.
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
 * The ranges are those `fm_workbook_set_pinned_now` accepts. The pin is a
 * calendar instant, not a normalising constructor: a month of `13` is rejected
 * rather than rolled into the next year, because silently moving a host's pin
 * would make the results it pinned unexplainable.
 */
typedef struct {
  /** The calendar year, 1900 to 9999 (1904 to 9999 in a 1904 date-system conversion). */
  int32_t year;
  /** The month, 1 to 12. */
  int32_t month;
  /** The day of the month, 1 to the length of that month under the Gregorian leap rule. */
  int32_t day;
  /** The hour, 0 to 23. */
  int32_t hour;
  /** The minute, 0 to 59. */
  int32_t minute;
  /** The second, 0 to 59. */
  int32_t second;
} fm_civil_time_t;

/**
 * @brief Reads the workbook's pinned wall-clock reading, if any.
 *
 * @param wb          Workbook handle. Must not be NULL.
 * @param out_present Set to `1` when a pin is in effect and `0` when the
 *                    workbook follows the host clock. Must not be NULL.
 * @param out_now     Populated with the pin when one is set; zero-filled
 *                    otherwise. Must not be NULL.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is NULL.
 */
FM_API fm_status_t fm_workbook_get_pinned_now(const fm_workbook_t* wb, int32_t* out_present, fm_civil_time_t* out_now);

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
 * it, so a save drops it and a reload comes back unpinned. Setting or
 * clearing the pin marks every formula cell dirty, as a profile or iterative
 * change does; in Auto mode the next read recomputes, in Manual mode
 * `fm_workbook_recalc` does.
 *
 * @param wb  Workbook handle. Must not be NULL.
 * @param now The reading to pin, every field in the range documented on
 *            `fm_civil_time_t`; `NULL` releases the pin so clock-dependent
 *            results follow the host clock again (a no-op when unpinned).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is NULL;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if any field is out of range.
 */
FM_API fm_status_t fm_workbook_set_pinned_now(fm_workbook_t* wb, const fm_civil_time_t* now);

/* -------------------------------------------------------------------------- */
/* Date serial conversion                                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Converts a civil date and time to an Excel date serial.
 *
 * The result equals `=DATE(year, month, day) + TIME(hour, minute, second)`
 * evaluated in a workbook with the same date system. Fields are not
 * normalised: every field must lie in the ranges documented on
 * `fm_civil_time_t`, with one exception for the 1900 date system, where
 * `{1900, 2, 29}` is accepted and maps to serial 60 (the day Excel's 1900
 * system counts). The 1904 system rejects 1900-02-29 and every date before
 * 1904-01-01.
 *
 * @param civil      The civil reading. Must not be NULL.
 * @param date1904   Non-zero for the 1904 date system, `0` for 1900.
 * @param out_serial Receives the serial. Must not be NULL; untouched on failure.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `civil` or `out_serial` is NULL;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if a field is out of range or the
 *         date precedes the date system's first day.
 */
FM_API fm_status_t fm_convert_civil_to_serial(const fm_civil_time_t* civil, int32_t date1904, double* out_serial);

/**
 * @brief Converts an Excel date serial to a civil date and time.
 *
 * The fractional part is rounded to the nearest whole second, as `SECOND()`
 * does; a rounding that reaches 86400 seconds carries into the next day. The
 * result matches `YEAR()` through `SECOND()` of the same serial. In the 1900
 * system the serials in `[60, 61)` give 1900-02-29.
 *
 * A serial is valid when it is finite, at least `1` (1900 system) or `0`
 * (1904 system), and below the serial of 9999-12-31 plus one day. A carry
 * past 9999-12-31 23:59:59 is out of range.
 *
 * @param serial    The serial to convert.
 * @param date1904  Non-zero for the 1904 date system, `0` for 1900.
 * @param out_civil Receives the civil fields. Must not be NULL; untouched on failure.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_civil` is NULL;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if the serial is out of range.
 */
FM_API fm_status_t fm_convert_serial_to_civil(double serial, int32_t date1904, fm_civil_time_t* out_civil);

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
 * Lifetime: static.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_profile_id` is NULL.
 */
FM_API fm_status_t fm_workbook_get_profile_id(const fm_workbook_t* wb, const char** out_profile_id);

/**
 * @brief Sets the workbook's Excel formula profile by id.
 *
 * New workbooks use `win-365-en_US` by default. Supported ids are
 * `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,
 * pt_BR,ru_RU,zh_TW,it_IT,nl_NL}`. The `mac-*`
 * ids and `win-365-ja_JP` follow Excel; the other `win-*` ids
 * are estimated from the Mac values plus the Windows host rules.
 *
 * The change marks every formula cell dirty: in Auto mode the next read
 * recomputes, in Manual mode `fm_workbook_recalc` does.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `profile_id` is NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `profile_id` is not a documented profile.
 */
FM_API fm_status_t fm_workbook_set_profile_id(fm_workbook_t* wb, const char* profile_id);

/**
 * @brief Writes the number of supported Excel formula profile ids (28) to
 *        `*out_count`.
 *
 * Drives the profile id iterator (`fm_get_profile_id_at`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out_count` is `NULL`.
 */
FM_API fm_status_t fm_get_profile_id_count(size_t* out_count);

/**
 * @brief Returns the `index`-th supported Excel formula profile id, the set
 *        `fm_workbook_set_profile_id` accepts.
 *
 * The order is fixed across calls and releases of the same table.
 * `*out_profile_id` is a static view into Formulon's profile table.
 *
 * Lifetime: static.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out_profile_id` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_get_profile_id_at(size_t index, const char** out_profile_id);

/**
 * @brief Where the workbook is saved: its folder and its file name.
 *
 * The location is metadata only: it is never read from or written to a file
 * system. It drives `CELL("filename")`, `CELL("address")` for other sheets,
 * the `[0]!Name` spelling, the saved absolute folder path, and the relative
 * form of local hyperlink targets in a saved file. `INFO("directory")` is the application's
 * current directory and does not follow it. Set the location to the
 * destination before saving to a new place, as Excel's Save As does; relative
 * targets in the saved file are computed against it.
 *
 * `directory` is absolute in one of three forms: POSIX (`/a/b`, a backslash
 * being a literal character, `/` alone the root), drive (`C:\a`, `C:/a`,
 * `c:\a\`; stored `C:/a`, the drive root `C:/`) or UNC (`\\server\share\a`,
 * `//server/share/a`; stored `//server/share/a`). The trailing separator is
 * stripped except at a root. The getter returns this canonical form, so
 * setting what it returns changes nothing.
 */
typedef struct {
  uint32_t struct_size;
  /** The folder, in canonical form; `NULL` when the book has a name but no folder. */
  const char* directory;
  /** The file name, extension included; required. */
  const char* file_name;
} fm_workbook_location_t;

/**
 * @brief Reads the workbook's location. `*out_present` is `0` when it is
 *        unset, as for a book that has not been saved, and `1` with `*out`
 *        filled otherwise.
 *
 * Loading from bytes leaves the location unset. While a file name is set,
 * every getter that returns formula text spells a reference to the workbook's
 * own defined name (stored `[0]!Name`) with it, as Excel's formula bar does:
 * `M5.xlsx!Name`, quoted when the name needs it (`'M5 b-1.xlsx'!Name`).
 * The strings are read-scratch-backed.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is NULL;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `out->struct_size` is below v1.
 */
FM_API fm_status_t fm_workbook_get_location(const fm_workbook_t* wb, int32_t* out_present, fm_workbook_location_t* out);

/**
 * @brief Sets the location `fm_workbook_get_location` reads; `location ==
 *        NULL` unsets it. The whole record is validated before anything
 *        changes.
 *
 * While a file name is set, formula text entered through the C ABI that names
 * it refers to this workbook and creates no external link: `<name>!Name` and
 * `'<name>'!Name` are stored as `[0]!Name`, and `[<name>]Sheet1!A1` as
 * `Sheet1!A1`. The book is matched under ASCII case folding. When a folder
 * first appears, every relative local hyperlink target is resolved against it.
 * Nothing is marked dirty: values that read the location (`CELL`,
 * `FORMULATEXT`) refresh at the next recalculation, as Excel's do after Save
 * As, and a save writes them as computed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is NULL, or
 *           `location->file_name` is NULL;
 *         `FM_STATUS_INVALID_ARGUMENT` (2), changing nothing, when
 *           `location->struct_size` is below v1; when `file_name` is empty,
 *           does not end in a workbook extension (`.xlsx`, `.xlsm`, `.xlsb`,
 *           `.xls`, `.xltx`, `.xltm`, `.xlt`, `.xlam`, `.xla`), holds one of
 *           `[ ] : / \ ? *`, is not valid UTF-8, is longer than 218 bytes,
 *           or equals the book name of an external link; when `directory` is
 *           `""` (use NULL), is not absolute in one of the three forms
 *           (`C:`, `C:rel`, `\\server`, `//server` and `//a` are refused),
 *           holds `[` or `]`, is not valid UTF-8, or has a `.`, `..` or empty
 *           inner segment;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005), changing nothing, when
 *           `directory` + `/` + `file_name` is longer than 1,024 bytes after
 *           canonicalisation.
 */
FM_API fm_status_t fm_workbook_set_location(fm_workbook_t* wb, const fm_workbook_location_t* location);

/* -------------------------------------------------------------------------- */
/* Sheet protection (per-sheet `<sheetProtection>`)                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Wide POD mirror of `formulon::SheetProtection` (ECMA-376
 *        §18.3.1.85 `<sheetProtection>`).
 *
 * A sheet is protected exactly when it has this record:
 * `fm_sheet_set_sheet_protection` with a record protects it, with `NULL`
 * unprotects it, and `fm_sheet_get_sheet_protection` reports the record's
 * presence in `*out_present`.
 *
 * Strings are NUL-terminated UTF-8. On the read path the pointers are
 * model-backed views into the workbook's own storage and remain valid until
 * the next mutation of the same sheet's protection record or handle
 * destruction. Reads do not invalidate them. On the write path the strings are
 * copied into the workbook; the caller's buffers may be released after the
 * call. Pass `NULL` for any string field to leave it absent.
 *
 * Each `allow_*` flag is the inverse of the OOXML attribute of the same action
 * (`allow_format_cells` is `formatCells="0"`). A getter returns
 * `FM_TRISTATE_FALSE` or `FM_TRISTATE_TRUE` for the two tristate select flags.
 * A record zeroed by `FM_STRUCT_INIT` is Excel's default protection without a
 * password.
 */
typedef struct {
  uint32_t struct_size;
  /** The name of the hash algorithm that produced `hash_value`; `NULL` when the sheet states none. */
  const char* algorithm_name;
  /** The password hash, stored as given; `NULL` when the sheet states none. */
  const char* hash_value;
  /** The salt the password hash was computed with, stored as given; `NULL` when the sheet states none. */
  const char* salt_value;
  /** The number of hash iterations; 0 when the sheet states none. */
  uint32_t spin_count;
  /** The legacy 16-bit password hash, as the hex text Excel writes; `NULL` when the sheet states none. */
  const char* legacy_password_hash;
  /** 1 when users may edit objects (shapes, charts, pictures) on the protected sheet; Excel's default 0. */
  int32_t allow_edit_objects;
  /** 1 when users may edit scenarios on the protected sheet; Excel's default 0. */
  int32_t allow_edit_scenarios;
  /** 1 when users may format cells on the protected sheet; Excel's default 0. */
  int32_t allow_format_cells;
  /** 1 when users may format columns on the protected sheet; Excel's default 0. */
  int32_t allow_format_columns;
  /** 1 when users may format rows on the protected sheet; Excel's default 0. */
  int32_t allow_format_rows;
  /** 1 when users may insert columns on the protected sheet; Excel's default 0. */
  int32_t allow_insert_columns;
  /** 1 when users may insert rows on the protected sheet; Excel's default 0. */
  int32_t allow_insert_rows;
  /** 1 when users may insert hyperlinks on the protected sheet; Excel's default 0. */
  int32_t allow_insert_hyperlinks;
  /** 1 when users may delete columns on the protected sheet; Excel's default 0. */
  int32_t allow_delete_columns;
  /** 1 when users may delete rows on the protected sheet; Excel's default 0. */
  int32_t allow_delete_rows;
  /** Tristate: whether users may select locked cells on the protected sheet; default allows. */
  int32_t allow_select_locked_cells; /* fm_tristate_t */
  /** Tristate: whether users may select unlocked cells on the protected sheet; default allows. */
  int32_t allow_select_unlocked_cells; /* fm_tristate_t */
  /** 1 when users may sort on the protected sheet; Excel's default 0. */
  int32_t allow_sort;
  /** 1 when users may use AutoFilter on the protected sheet; Excel's default 0. */
  int32_t allow_auto_filter;
  /** 1 when users may use PivotTables on the protected sheet; Excel's default 0. */
  int32_t allow_pivot_tables;
} fm_sheet_protection_t;

/**
 * @brief Partial update of a sheet's protection
 *        (`fm_sheet_update_sheet_protection`), following the header's patch
 *        rules (Patches, above).
 *
 * Each field of `fm_sheet_protection_t` has its `<field>_op`.
 */
typedef struct {
  uint32_t struct_size;
  /** The update operation for `algorithm_name`: keep it, clear it, or set it to the value beside it. */
  fm_patch_op_t algorithm_name_op;
  /** The password hash algorithm name (for example SHA-512). */
  const char* algorithm_name;
  /** The update operation for `hash_value`: keep it, clear it, or set it to the value beside it. */
  fm_patch_op_t hash_value_op;
  /** The base64 password hash. */
  const char* hash_value;
  /** The update operation for `salt_value`: keep it, clear it, or set it to the value beside it. */
  fm_patch_op_t salt_value_op;
  /** The base64 salt the password hash was computed with. */
  const char* salt_value;
  /** The update operation for `spin_count`: keep it, clear it, or set it to the value beside it. */
  fm_patch_op_t spin_count_op;
  /** The number of hash iterations. */
  uint32_t spin_count;
  /** The update operation for `legacy_password_hash`: keep it, clear it, or set it to the value beside it. */
  fm_patch_op_t legacy_password_hash_op;
  /** The legacy 16-bit password hash, as the hex text Excel writes. */
  const char* legacy_password_hash;
  /** The update operation for `allow_edit_objects`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_edit_objects_op;
  /** Whether the protected sheet allows editing objects. */
  int32_t allow_edit_objects;
  /** The update operation for `allow_edit_scenarios`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_edit_scenarios_op;
  /** Whether the protected sheet allows editing scenarios. */
  int32_t allow_edit_scenarios;
  /** The update operation for `allow_format_cells`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_format_cells_op;
  /** Whether the protected sheet allows formatting cells. */
  int32_t allow_format_cells;
  /** The update operation for `allow_format_columns`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_format_columns_op;
  /** Whether the protected sheet allows formatting columns. */
  int32_t allow_format_columns;
  /** The update operation for `allow_format_rows`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_format_rows_op;
  /** Whether the protected sheet allows formatting rows. */
  int32_t allow_format_rows;
  /** The update operation for `allow_insert_columns`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_insert_columns_op;
  /** Whether the protected sheet allows inserting columns. */
  int32_t allow_insert_columns;
  /** The update operation for `allow_insert_rows`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_insert_rows_op;
  /** Whether the protected sheet allows inserting rows. */
  int32_t allow_insert_rows;
  /** The update operation for `allow_insert_hyperlinks`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_insert_hyperlinks_op;
  /** Whether the protected sheet allows inserting hyperlinks. */
  int32_t allow_insert_hyperlinks;
  /** The update operation for `allow_delete_columns`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_delete_columns_op;
  /** Whether the protected sheet allows deleting columns. */
  int32_t allow_delete_columns;
  /** The update operation for `allow_delete_rows`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_delete_rows_op;
  /** Whether the protected sheet allows deleting rows. */
  int32_t allow_delete_rows;
  /** The update operation for `allow_select_locked_cells`: keep it, clear it, or set it to the value beside it. Clear
   *  is rejected. */
  fm_patch_op_t allow_select_locked_cells_op;
  /** Whether the protected sheet allows selecting locked cells. */
  int32_t allow_select_locked_cells;
  /** The update operation for `allow_select_unlocked_cells`: keep it, clear it, or set it to the value beside it. Clear
   *  is rejected. */
  fm_patch_op_t allow_select_unlocked_cells_op;
  /** Whether the protected sheet allows selecting unlocked cells. */
  int32_t allow_select_unlocked_cells;
  /** The update operation for `allow_sort`: keep it, clear it, or set it to the value beside it. Clear is rejected. */
  fm_patch_op_t allow_sort_op;
  /** Whether the protected sheet allows sorting. */
  int32_t allow_sort;
  /** The update operation for `allow_auto_filter`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_auto_filter_op;
  /** Whether the protected sheet allows using the AutoFilter. */
  int32_t allow_auto_filter;
  /** The update operation for `allow_pivot_tables`: keep it, clear it, or set it to the value beside it. Clear is
   *  rejected. */
  fm_patch_op_t allow_pivot_tables_op;
  /** Whether the protected sheet allows using PivotTables. */
  int32_t allow_pivot_tables;
} fm_sheet_protection_patch_t;

/**
 * @brief Reads the sheet's protection.
 *
 * `*out_present` is `1` when the sheet is protected and `*out` then holds the
 * record; otherwise `*out_present` is `0` and `*out` is zero-filled except
 * `struct_size` (still `FM_STATUS_OK` (0)).
 *
 * Lifetime: transient.
 *
 * @param wb           Workbook handle. Must not be NULL.
 * @param sheet_index  Sheet index, 0-based.
 * @param out_present  Set to `1` when the sheet is protected, else `0`.
 * @param out          Populated on success. String pointers are
 *                     model-backed views into the workbook's storage and
 *                     remain valid until the next protection mutation on
 *                     this sheet or handle destruction.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_present` or `out` is NULL;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_sheet_protection(const fm_workbook_t* wb, size_t sheet_index, int32_t* out_present,
                                                 fm_sheet_protection_t* out);

/**
 * @brief Protects the sheet with record `in`, replacing any protection it
 *        has, or unprotects it when `in` is `NULL` (a no-op on an unprotected
 *        sheet). Strings are deep-copied into the workbook.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is NULL;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a select flag is outside
 *         `fm_tristate_t` or another flag is not 0 / 1.
 */
FM_API fm_status_t fm_sheet_set_sheet_protection(fm_workbook_t* wb, size_t sheet_index,
                                                 const fm_sheet_protection_t* in);

/**
 * @brief Merges `patch` into the sheet's protection. The whole patch is
 *        checked before anything changes. On an unprotected sheet the patch
 *        applies to Excel's default protection (`FM_STRUCT_INIT`) and protects
 *        the sheet, unless every op is `KEEP`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is NULL;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits
 *         a flag, `SET` carries a `NULL` string, or a flag is not 0 / 1;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a `SET` string is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_update_sheet_protection(fm_workbook_t* wb, size_t sheet_index,
                                                    const fm_sheet_protection_patch_t* patch);

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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is past the last sheet;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `viewport` is inverted or outside the grid, or
 *         `sheet_index` is above 65535;
 *         an `FM_STATUS_EVAL_*` or `FM_STATUS_GRAPH_*` code on engine failure.
 */
FM_API fm_status_t fm_sheet_recalc_viewport(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* viewport,
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
 * payload fields on `fm_cf_match_info_t` are determined by this enumerator.
 */
/* NOLINTNEXTLINE(performance-enum-size): C11 enums cannot specify an
 * underlying type; matching `fm_value_kind_t`'s shape keeps the C ABI
 * uniform across enums. */
typedef enum {
  FM_CF_MATCH_KIND_DIFFERENTIAL_FORMAT = 0,
  FM_CF_MATCH_KIND_COLOR_SCALE = 1,
  FM_CF_MATCH_KIND_DATA_BAR = 2,
  FM_CF_MATCH_KIND_ICON_SET = 3
} fm_cf_match_kind_t;

/**
 * @brief Icon set (`fm_cf_rule_t::icon_set`, `fm_cf_match_info_t::icon_set`,
 *        `fm_cf_icon_t::icon_set`), in OOXML `ST_IconSetType` order.
 *
 * The first 17 members are the built-in sets a rule draws and mirror
 * `formulon::cf::IconSetName`. `FM_ICON_SET_NAME_THREE_STARS` ..
 * `FM_ICON_SET_NAME_NO_ICONS` are the x14-only sets a custom icon level
 * (`fm_cf_icon_t`) can take its icon from; a rule's own `icon_set`
 * rejects them with `FM_STATUS_INVALID_ARGUMENT` (2).
 */
typedef enum {
  FM_ICON_SET_NAME_THREE_ARROWS = 0,
  FM_ICON_SET_NAME_THREE_ARROWS_GRAY = 1,
  FM_ICON_SET_NAME_THREE_FLAGS = 2,
  FM_ICON_SET_NAME_THREE_TRAFFIC_LIGHTS_1 = 3,
  FM_ICON_SET_NAME_THREE_TRAFFIC_LIGHTS_2 = 4,
  FM_ICON_SET_NAME_THREE_SIGNS = 5,
  FM_ICON_SET_NAME_THREE_SYMBOLS = 6,
  FM_ICON_SET_NAME_THREE_SYMBOLS_2 = 7,
  FM_ICON_SET_NAME_FOUR_ARROWS = 8,
  FM_ICON_SET_NAME_FOUR_ARROWS_GRAY = 9,
  FM_ICON_SET_NAME_FOUR_RED_TO_BLACK = 10,
  FM_ICON_SET_NAME_FOUR_RATING = 11,
  FM_ICON_SET_NAME_FOUR_TRAFFIC_LIGHTS = 12,
  FM_ICON_SET_NAME_FIVE_ARROWS = 13,
  FM_ICON_SET_NAME_FIVE_ARROWS_GRAY = 14,
  FM_ICON_SET_NAME_FIVE_RATING = 15,
  FM_ICON_SET_NAME_FIVE_QUARTERS = 16,
  FM_ICON_SET_NAME_THREE_STARS = 17,
  FM_ICON_SET_NAME_THREE_TRIANGLES = 18,
  FM_ICON_SET_NAME_FIVE_BOXES = 19,
  FM_ICON_SET_NAME_NO_ICONS = 20,
} fm_icon_set_name_t;

/**
 * @brief The icon drawn for one level of a custom icon set (x14 `cfIcon`).
 */
typedef struct {
  /** The icon set the icon comes from, an `fm_icon_set_name_t`; `FM_ICON_SET_NAME_NO_ICONS` for a level that draws
   *  nothing, -1 for a file spelling outside the enumeration. */
  int32_t icon_set;
  /** The 0-based index of the icon within `icon_set`; 0 for a level that draws nothing. */
  int32_t icon_index;
} fm_cf_icon_t;

/** @brief `<x14:dataBar direction>`: the edge a data bar grows from. */
typedef enum {
  FM_CF_DATA_BAR_DIRECTION_CONTEXT = 0, /**< Follows the sheet's reading direction (the default). */
  FM_CF_DATA_BAR_DIRECTION_LEFT_TO_RIGHT = 1,
  FM_CF_DATA_BAR_DIRECTION_RIGHT_TO_LEFT = 2,
} fm_cf_data_bar_direction_t;

/** @brief `<x14:dataBar axisPosition>`. */
typedef enum {
  FM_CF_DATA_BAR_AXIS_POSITION_AUTOMATIC = 0, /**< The default. */
  FM_CF_DATA_BAR_AXIS_POSITION_MIDDLE = 1,
  FM_CF_DATA_BAR_AXIS_POSITION_NONE = 2,
} fm_cf_data_bar_axis_position_t;

/**
 * @brief Resolved CF match for one rule on one cell.
 *
 * The active fields depend on `kind`; all non-active fields carry default-zero
 * values. The colours are resolved against the workbook theme and palette, so
 * their kind is `FM_COLOR_KIND_RGB` (`FM_COLOR_KIND_NONE` when inactive).
 * Per-cell match lists are returned in priority-ascending order, so
 * `matches[0]` is the highest-priority match. Fold first-wins: walk the list
 * in the returned order and apply each property only where nothing has set it
 * yet, so a lower-priority match fills what a higher-priority one left unset
 * and never overwrites it.
 */
typedef struct {
  uint32_t struct_size;
  /** The kind of match; selects which of the other members carry data. */
  int32_t kind;
  /** The workbook-global priority of the matching rule; a smaller number ranks higher, and matches of a cell come in
   *  ascending order. */
  int32_t priority;
  /** 1 when `dxf_index` is set; 0 means the rule applies no differential format. */
  int32_t has_dxf_index;
  /** The 0-based index of the differential format the rule applies; read only when `has_dxf_index` is 1. Active for a
   *  differential-format match; otherwise 0. */
  uint32_t dxf_index;
  /** The interpolated fill colour of a colour-scale match, resolved to RGB; kind `NONE` for other kinds. */
  fm_color_t color;
  /** Length of the data bar as a percentage (0 to 100) of the cell width. Active for a data-bar match; otherwise 0. */
  double bar_length_pct;
  /** Position of the data-bar axis as a percentage of the cell width from the left edge: 0 at the left edge, 50 when
   *  centred. Active for a data-bar match; otherwise 0. */
  double bar_axis_position_pct;
  /** 1 when the data bar draws in the negative direction; 0 otherwise. */
  int32_t bar_is_negative;
  /** The fill colour of the data bar, resolved to RGB; kind `NONE` for other kinds. */
  fm_color_t bar_fill;
  /** The border colour of the data bar, resolved to RGB; kind `NONE` when the bar draws no border. */
  fm_color_t bar_border;
  /** 1 when the data bar is filled with a gradient; 0 for a solid fill. */
  int32_t bar_gradient;
  /** The icon set an icon-set match draws from, an `fm_icon_set_name_t`. */
  int32_t icon_set;
  /** The 0-based index of the icon within the set, after the rule's reverse setting has been applied. Active for an
   *  icon-set match; otherwise 0. */
  uint8_t icon_index;
  /** The direction of the data-bar rule, an `fm_cf_data_bar_direction_t`; the geometry is stated from the left edge and
   *  a host mirrors a right-to-left bar. */
  int32_t bar_direction;
} fm_cf_match_info_t;

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
 * On success the count `fm_pagination_get_page_count` reports is the true
 * total; it is never a truncated one.
 *
 * Lifetime: owned (until fm_pagination_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_PRINT_INVALID_AREA` (9002) when the sheet's `_xlnm.Print_Area` exists but does not parse;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the page grid the sheet declares exceeds the
 *           supported page count, which a file listing a manual break before nearly every row and column
 *           can reach (context `part=print_pages requested=<n> limit=16777216`).
 */
FM_API fm_status_t fm_sheet_paginate(const fm_workbook_t* wb, size_t sheet_index, fm_pagination_t** out);

/** @brief Releases a pagination result. `pagination == NULL` is a no-op. */
FM_API void fm_pagination_destroy(fm_pagination_t* pagination);

/**
 * @brief Writes the total physical page count to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_page_count(const fm_pagination_t* pagination, size_t* out_count);

/**
 * @brief Writes the number of rectangles in the sheet's declared
 *        `_xlnm.Print_Area` to `*out_count`.
 *
 * Zero also means the sheet declares no print area, mirroring Excel's own
 * `PageSetup.PrintArea`; it is not backfilled with the used range. Pagination
 * still falls back to the used range in that case, so
 * `fm_pagination_get_page_count` can be non-zero while this is zero.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_print_area_count(const fm_pagination_t* pagination, size_t* out_count);

/**
 * @brief Copies resolved print-area rectangle `index` (the domain `[0,
 *        fm_pagination_get_print_area_count)`) into `*out`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `pagination` or `out` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_pagination_get_print_area_at(const fm_pagination_t* pagination, size_t index,
                                                   fm_cell_rect_t* out);

/**
 * @brief Writes the number of horizontal (row) page breaks to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_row_break_count(const fm_pagination_t* pagination, size_t* out_count);

/** @brief One row break as the worksheet declares it (`<rowBreaks><brk>`). */
typedef struct {
  uint32_t struct_size;
  /** 0-based row the break sits before. */
  uint32_t row;
  /** The first column the break covers (OOXML `min`). */
  uint32_t first;
  /** The last column the break covers (OOXML `max`). */
  uint32_t last;
  /** An `fm_tristate_t`: default or true is a user-placed break, false an automatic one; a read returns true or false.
   * `man` is absent for an automatic break. */
  int32_t manual; /* fm_tristate_t */
} fm_row_break_t;

/** @brief One column break (`<colBreaks><brk>`), the counterpart of `fm_row_break_t`. */
typedef struct {
  uint32_t struct_size;
  /** 0-based column the break sits before. */
  uint32_t col;
  /** The first row the break covers (OOXML `min`). */
  uint32_t first;
  /** The last row the break covers (OOXML `max`). */
  uint32_t last;
  /** An `fm_tristate_t`: default or true is a user-placed break, false an automatic one; a read returns true or false.
   */
  int32_t manual; /* fm_tristate_t */
} fm_column_break_t;

/**
 * @brief Reads horizontal break `index` (the domain `[0,
 *        fm_pagination_get_row_break_count)`) into `*out`.
 *
 * `out->row` is the 0-based row the break precedes. For a manual break
 * `first` / `last` are the span declared on the sheet and `manual` is
 * `FM_TRISTATE_TRUE`; for a break the page walk computed they are `0` / `0`
 * (the whole-sheet span, the convention of `fm_sheet_set_row_break`) and
 * `manual` is `FM_TRISTATE_FALSE`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `pagination` or `out` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_pagination_get_row_break_at(const fm_pagination_t* pagination, size_t index, fm_row_break_t* out);

/**
 * @brief Writes the number of vertical (column) page breaks to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_column_break_count(const fm_pagination_t* pagination, size_t* out_count);

/**
 * @brief Reads vertical break `index` (the domain `[0,
 *        fm_pagination_get_column_break_count)`) into `*out`; the column
 *        counterpart of `fm_pagination_get_row_break_at`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `pagination` or `out` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_pagination_get_column_break_at(const fm_pagination_t* pagination, size_t index,
                                                     fm_column_break_t* out);

/** @brief A rectangle in points. Sheet geometry measures from the top-left
 *         of cell A1; pagination measures from the page's top-left corner. */
typedef struct {
  uint32_t struct_size;
  /** The distance in points from the rectangle's origin to its left edge. */
  double x;
  /** The distance in points from the rectangle's origin to its top edge. */
  double y;
  /** The width in points. */
  double width_pt;
  /** The height in points. */
  double height_pt;
} fm_rect_pt_info_t;

/** @brief The paper a sheet prints on. */
typedef struct {
  uint32_t struct_size;
  /** The paper width in points, after the orientation swap. */
  double width_pt;
  /** The paper height in points, after the orientation swap. */
  double height_pt;
  /** `1` when the paper is laid out landscape. */
  int32_t landscape;
  /** `0` when the stored paper size was unrecognised and A4 was substituted. */
  int32_t known;
} fm_paper_info_t;

/** @brief Page margins in inches, as `<pageMargins>` states them. */
typedef struct {
  uint32_t struct_size;
  /** The left page margin in inches. */
  double left_in;
  /** The right page margin in inches. */
  double right_in;
  /** The top page margin in inches. */
  double top_in;
  /** The bottom page margin in inches. */
  double bottom_in;
  /** The distance in inches from the page top edge to the header band. */
  double header_in;
  /** The distance in inches from the page bottom edge to the footer band. */
  double footer_in;
} fm_margins_pt_info_t;

/** @brief Repeat rows / repeat columns (0-based, inclusive), as
 *         `fm_sheet_set_print_titles` takes them and `fm_sheet_get_print_titles`
 *         and `fm_pagination_get_print_titles` report them. */
typedef struct {
  uint32_t struct_size;
  /** 1 when rows repeat; 0 means no rows repeat. */
  int32_t has_rows;
  /** 0-based index of the first repeated row; meaningful only when `has_rows` is 1. */
  uint32_t first_row;
  /** 0-based index of the last repeated row, inclusive; meaningful only when `has_rows` is 1. */
  uint32_t last_row;
  /** 1 when columns repeat; 0 means no columns repeat. */
  int32_t has_columns;
  /** 0-based index of the first repeated column; meaningful only when `has_columns` is 1. */
  uint32_t first_col;
  /** 0-based index of the last repeated column, inclusive; meaningful only when `has_columns` is 1. */
  uint32_t last_col;
} fm_print_titles_t;

/** @brief One physical page: the cell block it carries and where it sits. */
typedef struct {
  uint32_t struct_size;
  /** 0-based index of the print area this page belongs to, as `fm_pagination_get_print_area_at` indexes them (0 for the
   *  used-range fallback). */
  uint32_t area_index;
  /** The cell block the page carries. */
  fm_cell_rect_t rect;
  /** Horizontal position in points of the cell block's left edge on the paper. */
  double origin_x_pt;
  /** Vertical position in points of the cell block's top edge on the paper. */
  double origin_y_pt;
  /** Scaled width in points of the page's cell block, repeated title columns excluded. */
  double width_pt;
  /** Scaled height in points of the page's cell block, repeated title rows excluded. */
  double height_pt;
} fm_page_layout_info_t;

/**
 * @brief Reads the paper size and orientation pagination laid out on.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_paper(const fm_pagination_t* pagination, fm_paper_info_t* out);

/**
 * @brief Reads the page margins in inches.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_margins(const fm_pagination_t* pagination, fm_margins_pt_info_t* out);

/**
 * @brief Reads the cell-body rectangle before print-title reservation, in
 *        points from the page's top-left corner.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_printable(const fm_pagination_t* pagination, fm_rect_pt_info_t* out);

/**
 * @brief Reads the effective scale as a percentage (100 is 100%): the
 *        `scale_pct` of the page setup, or the fit-to-page scale.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`.
 */
FM_API fm_status_t fm_pagination_get_scale(const fm_pagination_t* pagination, double* out_scale_pct);

/** @brief Print page order (`<pageSetup pageOrder>`); zero is the OOXML default `downThenOver`. */
typedef enum {
  FM_PAGE_ORDER_DOWN_THEN_OVER = 0,
  FM_PAGE_ORDER_OVER_THEN_DOWN = 1,
} fm_page_order_t;

/**
 * @brief Reads the page order into `*out_order`, an `fm_page_order_t`.
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
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_pagination_get_page_at(const fm_pagination_t* pagination, size_t index,
                                             fm_page_layout_info_t* out);

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
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
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
 * from the new fragment immediately, so a `fm_sheet_paginate` that
 * follows observes the new settings.
 *
 * `fitToPage` is NOT part of `<pageSetup>` - it lives in
 * `<sheetPr><pageSetUpPr>`; set it through `fm_page_setup_patch_t::fit_to_page`
 * on `fm_sheet_update_page_setup`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `xml` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the fragment is malformed, carries more than one
 *           top-level element, or names the wrong root; and when it declares an `r:id` on a sheet that
 *           has no printer-settings part, which would leave a dangling relationship that makes Excel
 *           repair the file;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the fragment exceeds 4 KiB.
 */
FM_API fm_status_t fm_sheet_set_page_setup_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<pageMargins>` as raw XML. `*out_xml` is a read-scratch-backed
 *         view; see `fm_sheet_get_page_setup_xml` for the shared contract.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range. */
FM_API fm_status_t fm_sheet_get_page_margins_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<pageMargins>`. Re-derives the structured margins; an
 *         empty string removes the element and restores the defaults.
 *         Fragment ceiling 4 KiB.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `xml` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the fragment is malformed or names the wrong root;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the fragment exceeds 4 KiB. */
FM_API fm_status_t fm_sheet_set_page_margins_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<printOptions>` as raw XML. `*out_xml` is a read-scratch-backed
 *         view; see `fm_sheet_get_page_setup_xml` for the shared contract.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range. */
FM_API fm_status_t fm_sheet_get_print_options_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<printOptions>`. This element has no structured view;
 *         the raw fragment is the whole model. Fragment ceiling 4 KiB.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `xml` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the fragment is malformed or names the wrong root;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the fragment exceeds 4 KiB. */
FM_API fm_status_t fm_sheet_set_print_options_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<headerFooter>` as raw XML. `*out_xml` is a read-scratch-backed
 *         view; see `fm_sheet_get_page_setup_xml` for the shared contract.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range. */
FM_API fm_status_t fm_sheet_get_header_footer_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<headerFooter>`. No structured view. Fragment ceiling
 *         64 KiB, which is wider than the other elements because six
 *         sections of formatting codes (and header images) live here.
 *
 * Header and footer text uses `&` as its formatting-code introducer
 * (`&P` = page number, `&D` = date). Those codes live in XML element text,
 * so a caller composing this fragment writes them escaped - Excel's own
 * output is `<oddHeader>&amp;C...</oddHeader>`, and a raw `&` would make
 * the fragment ill-formed. `fm_sheet_update_header_footer` takes the decoded
 * text instead and handles the escaping.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `xml` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the fragment is malformed or names the wrong root;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the fragment exceeds 64 KiB. */
FM_API fm_status_t fm_sheet_set_header_footer_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/** @brief Reads `<sheetPr>` as raw XML. `*out_xml` is a read-scratch-backed
 *         view; see `fm_sheet_get_page_setup_xml` for the shared contract.
 *
 * `<sheetPr>` is not print-only: it also carries `tabColor` (sheet tab
 * colour) and `codeName` (VBA binding). Replacing it wholesale to toggle
 * `fitToPage` would drop those; `fm_sheet_update_page_setup` with
 * `fit_to_page_op` and `fm_sheet_set_tab_color` edit one child each.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range. */
FM_API fm_status_t fm_sheet_get_sheet_pr_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml);

/** @brief Replaces `<sheetPr>`. Re-derives `fitToPage` from
 *         `<pageSetUpPr>`; absent means false. Fragment ceiling 8 KiB.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `xml` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the fragment is malformed or names the wrong root;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the fragment exceeds 8 KiB. */
FM_API fm_status_t fm_sheet_set_sheet_pr_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml);

/* -------------------------------------------------------------------------- */
/* Print settings - print area and titles                                     */
/* -------------------------------------------------------------------------- */

/**
 * @brief Writes the sheet-scoped `_xlnm.Print_Area` defined name from the
 *        `range_count` rectangles of `ranges`.
 *
 * Each area is stored absolute and qualified with the sheet's own name,
 * quoted per Excel's rules, so the stored formula matches what Excel writes:
 * a rectangle spanning every row is stored as columns (`$A:$D`), one
 * spanning every column as rows (`$1:$50`). Area order is preserved - Excel
 * prints them in the stated order. `range_count` 0 removes the defined name.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, or `ranges` is `NULL` with a
 *           non-zero `range_count`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a rectangle is inverted or outside the grid;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when `range_count` is past the per-call range
 *           cap of 16,384 every range-taking entry point shares (context
 *           `part=ranges requested=<n> limit=16384`).
 */
FM_API fm_status_t fm_sheet_set_print_area(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* ranges,
                                           size_t range_count);

/**
 * @brief Copies the print area's rectangles into the caller's buffer, in
 *        the stated order.
 *
 * Whole-axis areas come back expanded to the grid edge (`A:D` reads back as
 * `A1:D1048576`): the resolver normalises them to rectangles for the
 * paginator, and this getter reports the resolved form. `*out_count` always
 * receives the number of rectangles, 0 when the sheet declares none; at most
 * `capacity` of them are written to `out`. Call with `capacity == 0` (and
 * `out` may then be `NULL`) to size the buffer first.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is `NULL`, or `out` is `NULL`
 *           while `capacity > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_PRINT_INVALID_AREA` (9002) when the defined name exists but does not
 *           parse, so a malformed one is never mistaken for "no print area".
 */
FM_API fm_status_t fm_sheet_get_print_area(const fm_workbook_t* wb, size_t sheet_index, fm_cell_rect_t* out,
                                           size_t capacity, size_t* out_count);

/**
 * @brief Writes the sheet-scoped `_xlnm.Print_Titles` defined name from the
 *        axes of `titles` whose `has_*` flag is non-zero.
 *
 * `titles == NULL`, or both flags zero, removes the defined name. The stored
 * formula lists rows before columns and anchors every endpoint, matching
 * Excel (`Sheet1!$1:$2,Sheet1!$A:$A`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a flagged span is inverted or outside the grid.
 */
FM_API fm_status_t fm_sheet_set_print_titles(fm_workbook_t* wb, size_t sheet_index, const fm_print_titles_t* titles);

/**
 * @brief Reads the repeat rows / columns into `*out`; an axis the sheet does
 *        not repeat has its `has_*` flag 0 and its span zeroed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_PRINT_INVALID_AREA` (9002) when the defined name exists but does not parse.
 */
FM_API fm_status_t fm_sheet_get_print_titles(const fm_workbook_t* wb, size_t sheet_index, fm_print_titles_t* out);

/* -------------------------------------------------------------------------- */
/* Print settings - manual page breaks                                        */
/* -------------------------------------------------------------------------- */

/**
 * @brief Sets or removes the row break before `row`.
 *
 * A non-`NULL` `brk` upserts the break: an existing break at `row` has its
 * span and `manual` flag replaced rather than duplicated, and the list stays
 * sorted ascending by row. `brk->row` is ignored; the break sits at `row`.
 * `first == last == 0` takes the whole-sheet span (`0..16383`), which is what
 * Excel writes for a break inserted with no active selection. `brk == NULL`
 * removes the break at `row` and is a no-op when there is none.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a row past the grid, an inverted
 *         span, a span past the last column, a `manual` outside `fm_tristate_t`,
 *         or adding a break to an axis that holds 1026 breaks (context
 *         `part=manual_breaks requested=1027 limit=1026`). A row break divides
 *         the page horizontally, and 1026 is the limit Excel publishes for
 *         horizontal page breaks.
 */
FM_API fm_status_t fm_sheet_set_row_break(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                          const fm_row_break_t* brk);

/**
 * @brief Sets or removes the column break before `col`.
 *
 * Upsert, removal and ordering match `fm_sheet_set_row_break`; `brk->col` is
 * ignored, and `first == last == 0` takes the whole-sheet span
 * (`0..1048575`).
 *
 * @return As `fm_sheet_set_row_break`, for a column past the grid or a span
 *         past the last row. Excel publishes no separate limit for vertical
 *         page breaks; the horizontal limit of 1026 is applied here as a
 *         conservative bound rather than leaving the axis unbounded, so this
 *         ceiling is Formulon's, not a documented Excel one.
 */
FM_API fm_status_t fm_sheet_set_column_break(fm_workbook_t* wb, size_t sheet_index, uint32_t col,
                                             const fm_column_break_t* brk);

/**
 * @brief Removes every manual break on both axes; a no-op when there is none.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_breaks(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Writes the number of row breaks declared on the sheet to
 *        `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is NULL;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_break_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Copies the row break at `index` (ascending by row).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                             fm_row_break_t* out);

/**
 * @brief Writes the number of column breaks declared on the sheet to
 *        `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is NULL;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_column_break_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Copies the column break at `index` (ascending by column).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_column_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                fm_column_break_t* out);

/* -------------------------------------------------------------------------- */
/* Print settings - typed patch setters                                       */
/* -------------------------------------------------------------------------- */

/** @brief `<pageSetup orientation>` (`fm_page_setup_t::orientation`). */
typedef enum { FM_PAGE_ORIENTATION_PORTRAIT = 0, FM_PAGE_ORIENTATION_LANDSCAPE = 1 } fm_page_orientation_t;

/**
 * @brief The `<pageSetup>` settings in force (`fm_sheet_get_page_setup`).
 *
 * Every value field carries the effective value, and its `has_`
 * companion is non-zero when the attribute is stated in the XML, so a caller
 * can tell an explicit setting from an inherited default.
 */
typedef struct {
  uint32_t struct_size;
  /** 1 when the XML states `orientation`; 0 means no orientation is stated. */
  int32_t has_orientation;
  /** The page orientation, an `fm_page_orientation_t`; meaningful only when `has_orientation` is 1, otherwise the sheet
   *  prints portrait, as does the file value `default`. */
  int32_t orientation;
  /** 1 when the XML states `paper_size`; 0 means it reads a default. */
  int32_t has_paper_size;
  /** The OOXML paperSize code in force; an open set, so unknown codes round-trip. The engine lays out the eight it
   *  knows: 1 Letter, 5 Legal, 8 A3, 9 A4, 11 A5, 12 B4 (JIS), 13 B5 (JIS), 70 A6. */
  uint32_t paper_size;
  /** `1` when the XML states `scale_pct`; `0` means it reads a default. */
  int32_t has_scale_pct;
  /** The print scale as a percentage, 10 to 400; 100 when the attribute is absent. */
  uint32_t scale_pct;
  /** 1 when the XML states `fit_to_width`; 0 means it reads a default. */
  int32_t has_fit_to_width;
  /** The number of pages across the fit-to-page mode targets. */
  uint32_t fit_to_width;
  /** 1 when the XML states `fit_to_height`; 0 means it reads a default. */
  int32_t has_fit_to_height;
  /** The number of pages down the fit-to-page mode targets. */
  uint32_t fit_to_height;
  /** 1 when the XML states `fit_to_page`; 0 means it reads a default. */
  int32_t has_fit_to_page;
  /** `1` when the fit-to-page mode is selected; read from `<sheetPr><pageSetUpPr>`, not `<pageSetup>`. */
  int32_t fit_to_page;
} fm_page_setup_t;

/**
 * @brief Partial `<pageSetup>` update (`fm_sheet_update_page_setup`),
 *        following the header's patch rules (Patches, above).
 *
 * An attribute the patch keeps stays exactly as the file had it, and so do
 * the attributes the engine does not model (`r:id`, `horizontalDpi`,
 * `copies`).
 */
typedef struct {
  uint32_t struct_size;
  /** Operation on `orientation`: keep leaves it, set writes it, clear removes the attribute and the sheet prints
   * portrait. */
  fm_patch_op_t orientation_op;
  /** The page orientation to set, an `fm_page_orientation_t`; read when `orientation_op` is set. */
  int32_t orientation;
  /** Operation on `paper_size`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t paper_size_op;
  /** The OOXML paperSize code to set; any code is stored. */
  uint32_t paper_size;
  /** Operation on `scale_pct`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t scale_pct_op;
  /** The print scale percentage to set, in [10, 400]. */
  uint32_t scale_pct;
  /** Operation on `fit_to_width`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t fit_to_width_op;
  /** The number of pages across the fit-to-page mode targets. */
  uint32_t fit_to_width;
  /** Operation on `fit_to_height`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t fit_to_height_op;
  /** The number of pages down the fit-to-page mode targets. */
  uint32_t fit_to_height;
  /** Operation on `fit_to_page`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t fit_to_page_op;
  /** Non-zero selects the fit-to-page mode, stored on `<sheetPr><pageSetUpPr>`. */
  int32_t fit_to_page;
} fm_page_setup_patch_t;

/** @brief The `<pageMargins>` in force, in inches (`fm_sheet_get_page_margins`).
 *         Each `has_` flag reports whether the XML states the attribute,
 *         as in `fm_page_setup_t`. */
typedef struct {
  uint32_t struct_size;
  /** `1` when the XML states `left_in`; `0` means it reads a default. */
  int32_t has_left_in;
  /** The left margin in inches. */
  double left_in;
  /** `1` when the XML states `right_in`; `0` means it reads a default. */
  int32_t has_right_in;
  /** The right margin in inches. */
  double right_in;
  /** `1` when the XML states `top_in`; `0` means it reads a default. */
  int32_t has_top_in;
  /** The top margin in inches. */
  double top_in;
  /** `1` when the XML states `bottom_in`; `0` means it reads a default. */
  int32_t has_bottom_in;
  /** The bottom margin in inches. */
  double bottom_in;
  /** `1` when the XML states `header_in`; `0` means it reads a default. */
  int32_t has_header_in;
  /** The distance in inches from the top edge of the paper to the header. */
  double header_in;
  /** `1` when the XML states `footer_in`; `0` means it reads a default. */
  int32_t has_footer_in;
  /** The distance in inches from the bottom edge of the paper to the footer. */
  double footer_in;
} fm_page_margins_t;

/**
 * @brief Partial `<pageMargins>` update in inches
 *        (`fm_sheet_update_page_margins`), following the header's patch rules
 *        (Patches, above). No field can be cleared.
 */
typedef struct {
  uint32_t struct_size;
  /** Operation on `left_in`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t left_in_op;
  /** The left margin in inches; finite and non-negative. */
  double left_in;
  /** Operation on `right_in`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t right_in_op;
  /** The right margin in inches; finite and non-negative. */
  double right_in;
  /** Operation on `top_in`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t top_in_op;
  /** The top margin in inches; finite and non-negative. */
  double top_in;
  /** Operation on `bottom_in`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t bottom_in_op;
  /** The bottom margin in inches; finite and non-negative. */
  double bottom_in;
  /** Operation on `header_in`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t header_in_op;
  /** The distance in inches from the top edge of the paper to the header; finite and non-negative. */
  double header_in;
  /** Operation on `footer_in`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t footer_in_op;
  /** The distance in inches from the bottom edge of the paper to the footer; finite and non-negative. */
  double footer_in;
} fm_page_margins_patch_t;

/** @brief The `<printOptions>` flags (`fm_sheet_get_print_options`). Each
 *         `has_` flag reports whether the XML states the attribute. */
typedef struct {
  uint32_t struct_size;
  /** 1 when the XML states `grid_lines`; 0 means it reads 0. */
  int32_t has_grid_lines;
  /** `1` when the cell gridlines are printed. */
  int32_t grid_lines;
  /** 1 when the XML states `headings`; 0 means it reads 0. */
  int32_t has_headings;
  /** `1` when the row and column headings are printed. */
  int32_t headings;
  /** 1 when the XML states `horizontal_centered`; 0 means it reads 0. */
  int32_t has_horizontal_centered;
  /** `1` when the print area is centered horizontally on the page. */
  int32_t horizontal_centered;
  /** 1 when the XML states `vertical_centered`; 0 means it reads 0. */
  int32_t has_vertical_centered;
  /** `1` when the print area is centered vertically on the page. */
  int32_t vertical_centered;
} fm_print_options_t;

/** @brief Partial `<printOptions>` update (`fm_sheet_update_print_options`),
 *         following the header's patch rules (Patches, above). A `SET` flag
 *         is 0 or non-zero; no field can be cleared. */
typedef struct {
  uint32_t struct_size;
  /** Operation on `grid_lines`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t grid_lines_op;
  /** Non-zero prints the cell gridlines. */
  int32_t grid_lines;
  /** Operation on `headings`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t headings_op;
  /** Non-zero prints the row and column headings. */
  int32_t headings;
  /** Operation on `horizontal_centered`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t horizontal_centered_op;
  /** Non-zero centers the print area horizontally on the page. */
  int32_t horizontal_centered;
  /** Operation on `vertical_centered`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t vertical_centered_op;
  /** Non-zero centers the print area vertically on the page. */
  int32_t vertical_centered;
} fm_print_options_patch_t;

/**
 * @brief The `<headerFooter>` part (`fm_sheet_get_header_footer`).
 *
 * A present section is its decoded text, as `fm_header_footer_patch_t` takes
 * it. Each `has_` flag reports whether the XML states the attribute.
 */
typedef struct {
  uint32_t struct_size;
  /** The decoded text of the header of odd pages (every page unless `different_odd_even` or `different_first` apply);
   *  `NULL` when the sheet has no such section. */
  const char* odd_header;
  /** The decoded text of the footer of odd pages (every page unless `different_odd_even` or `different_first` apply);
   *  `NULL` when the sheet has no such section. */
  const char* odd_footer;
  /** The decoded text of the header of even pages, used when `different_odd_even` is set; `NULL` when the sheet has no
   *  such section. */
  const char* even_header;
  /** The decoded text of the footer of even pages, used when `different_odd_even` is set; `NULL` when the sheet has no
   *  such section. */
  const char* even_footer;
  /** The decoded text of the header of the first page, used when `different_first` is set; `NULL` when the sheet has no
   *  such section. */
  const char* first_header;
  /** The decoded text of the footer of the first page, used when `different_first` is set; `NULL` when the sheet has no
   *  such section. */
  const char* first_footer;
  /** 1 when the XML states `different_odd_even`; 0 means it reads its schema default. */
  int32_t has_different_odd_even;
  /** `1` when odd and even pages use different headers and footers; reads `0` when unstated. */
  int32_t different_odd_even;
  /** 1 when the XML states `different_first`; 0 means it reads its schema default. */
  int32_t has_different_first;
  /** `1` when the first page uses its own header and footer; reads `0` when unstated. */
  int32_t different_first;
  /** 1 when the XML states `scale_with_doc`; 0 means it reads its schema default. */
  int32_t has_scale_with_doc;
  /** `1` when the header and footer scale with the document's print scale; reads `1` when unstated. */
  int32_t scale_with_doc;
  /** 1 when the XML states `align_with_margins`; 0 means it reads its schema default. */
  int32_t has_align_with_margins;
  /** `1` when the header and footer align with the page margins; reads `1` when unstated. */
  int32_t align_with_margins;
} fm_header_footer_t;

/**
 * @brief Partial `<headerFooter>` update (`fm_sheet_update_header_footer`),
 *        following the header's patch rules (Patches, above).
 *
 * Section text is the decoded string, so Excel's formatting codes are
 * written plainly: `"&C&\"MS Gothic\"Report &P/&N"`. The engine applies XML
 * escaping on the way out, which is what turns the leading `&` of each code
 * into the `&amp;` Excel itself stores.
 */
typedef struct {
  uint32_t struct_size;
  /** Operation on `odd_header`: keep leaves it, set writes it, clear removes it. */
  fm_patch_op_t odd_header_op;
  /** The decoded text to set for the header of odd pages (every page unless `different_odd_even` or `different_first`
   *  apply); an empty string also removes the section. */
  const char* odd_header;
  /** Operation on `odd_footer`: keep leaves it, set writes it, clear removes it. */
  fm_patch_op_t odd_footer_op;
  /** The decoded text to set for the footer of odd pages (every page unless `different_odd_even` or `different_first`
   *  apply); an empty string also removes the section. */
  const char* odd_footer;
  /** Operation on `even_header`: keep leaves it, set writes it, clear removes it. */
  fm_patch_op_t even_header_op;
  /** The decoded text to set for the header of even pages; an empty string also removes the section. */
  const char* even_header;
  /** Operation on `even_footer`: keep leaves it, set writes it, clear removes it. */
  fm_patch_op_t even_footer_op;
  /** The decoded text to set for the footer of even pages; an empty string also removes the section. */
  const char* even_footer;
  /** Operation on `first_header`: keep leaves it, set writes it, clear removes it. */
  fm_patch_op_t first_header_op;
  /** The decoded text to set for the header of the first page; an empty string also removes the section. */
  const char* first_header;
  /** Operation on `first_footer`: keep leaves it, set writes it, clear removes it. */
  fm_patch_op_t first_footer_op;
  /** The decoded text to set for the footer of the first page; an empty string also removes the section. */
  const char* first_footer;
  /** Operation on `different_odd_even`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t different_odd_even_op;
  /** Non-zero uses different headers and footers for odd and even pages. */
  int32_t different_odd_even;
  /** Operation on `different_first`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t different_first_op;
  /** Non-zero gives the first page its own header and footer. */
  int32_t different_first;
  /** Operation on `scale_with_doc`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t scale_with_doc_op;
  /** Non-zero scales the header and footer with the document's print scale. */
  int32_t scale_with_doc;
  /** Operation on `align_with_margins`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t align_with_margins_op;
  /** Non-zero aligns the header and footer with the page margins. */
  int32_t align_with_margins;
} fm_header_footer_patch_t;

/**
 * @brief Merges `patch` into the sheet's `<pageSetup>`. The whole patch is
 *        checked before anything changes.
 *
 * `scale_pct` outside [10, 400] is rejected rather than clamped: a
 * mis-stated print scale lands on paper, so silently rounding it would hide
 * the mistake.
 *
 * A `SET` `fit_to_page` sets or clears `<sheetPr><pageSetUpPr fitToPage>`
 * without disturbing any other `<sheetPr>` content: a non-zero value creates
 * `<sheetPr>` and `<pageSetUpPr>` when absent, and zero removes the
 * `fitToPage` attribute but keeps both elements and everything else they
 * carry (`tabColor`, `codeName`, `<outlinePr>`, ...). `fitToPage` alone does
 * not scale anything: it selects the fit-to-page mode, and
 * `<pageSetup fitToWidth= fitToHeight=>` states the target, so "fit onto one
 * page" is `fit_to_page = 1` with `fit_to_width = 1` and `fit_to_height = 1`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits a field other than
 *           `orientation`, a `SET` `orientation` names no `fm_page_orientation_t`, or a `SET`
 *           `scale_pct` is outside [10, 400].
 */
FM_API fm_status_t fm_sheet_update_page_setup(fm_workbook_t* wb, size_t sheet_index,
                                              const fm_page_setup_patch_t* patch);

/**
 * @brief Merges `patch` into the sheet's `<pageMargins>`. The whole patch is
 *        checked before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2 or `CLEAR`, or a `SET` margin is
 *           negative, infinite or NaN.
 */
FM_API fm_status_t fm_sheet_update_page_margins(fm_workbook_t* wb, size_t sheet_index,
                                                const fm_page_margins_patch_t* patch);

/**
 * @brief Merges `patch` into the sheet's `<printOptions>`. The whole patch is
 *        checked before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2 or `CLEAR`.
 */
FM_API fm_status_t fm_sheet_update_print_options(fm_workbook_t* wb, size_t sheet_index,
                                                 const fm_print_options_patch_t* patch);

/**
 * @brief Merges `patch` into the sheet's `<headerFooter>`. The whole patch is
 *        checked before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits one of the four
 *           flags, or a `SET` section is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a `SET` section is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_update_header_footer(fm_workbook_t* wb, size_t sheet_index,
                                                 const fm_header_footer_patch_t* patch);

/**
 * @brief Reads the effective page setup into `*out`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_page_setup(const fm_workbook_t* wb, size_t sheet_index, fm_page_setup_t* out);

/**
 * @brief Reads the effective page margins, in inches, into `*out`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_page_margins(const fm_workbook_t* wb, size_t sheet_index, fm_page_margins_t* out);

/**
 * @brief Reads `<printOptions>` into `*out`; an absent attribute reads `0`
 *        (the schema default).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_print_options(const fm_workbook_t* wb, size_t sheet_index, fm_print_options_t* out);

/**
 * @brief Reads `<headerFooter>` into `*out`. The section strings are
 *        read-scratch-backed views.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_header_footer(const fm_workbook_t* wb, size_t sheet_index, fm_header_footer_t* out);

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
 * @brief Options of `fm_sheet_evaluate_cf_range`; the zero record is the
 *        default.
 */
typedef struct {
  uint32_t struct_size;
  /** An `fm_tristate_t`: `DEFAULT` or `TRUE` evaluates time-period rules, `FALSE` makes none of them match. */
  int32_t time_period_rules; /* fm_tristate_t */
  /** 1 when `today_serial` is set; 0 means time-period rules measure from the pinned clock when one is set, else from
   *  today. */
  int32_t has_today_serial;
  /** The date serial that time-period rules measure from. */
  double today_serial;
} fm_cf_range_evaluation_spec_t;

/**
 * @brief Evaluates every CF block on `sheet_index` against the cells of
 *        `range`.
 *
 * The result is sparse: only cells that produced at least one match
 * appear. Internally constructs an `Arena`, the default function
 * registry, and an `EvalContext` bound to the requested sheet.
 * `options == NULL` is the zero `fm_cf_range_evaluation_spec_t`.
 *
 * Lifetime: owned (until fm_cf_results_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `range` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `range` is inverted or outside the grid,
 *           `options->time_period_rules` is outside `fm_tristate_t`, or
 *           `options->has_today_serial` is non-zero and `options->today_serial`
 *           is NaN or infinite.
 *
 * On success the caller owns `*out` and must free it with
 * `fm_cf_results_destroy`.
 */
FM_API fm_status_t fm_sheet_evaluate_cf_range(const fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                              const fm_cf_range_evaluation_spec_t* options, fm_cf_results_t** out);

/**
 * @brief Releases a results handle. `results == NULL` is a no-op.
 */
FM_API void fm_cf_results_destroy(fm_cf_results_t* results);

/**
 * @brief Writes the number of cells in the result that produced at least
 *        one match to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `results` or `out_count` is `NULL`.
 */
FM_API fm_status_t fm_cf_results_get_cell_count(const fm_cf_results_t* results, size_t* out_count);

/**
 * @brief Reads cell `cell_index`'s coordinates and match count.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `cell_index >= cell_count`.
 */
FM_API fm_status_t fm_cf_results_get_cell_at(const fm_cf_results_t* results, size_t cell_index, uint32_t* out_row,
                                             uint32_t* out_col, size_t* out_match_count);

/**
 * @brief Reads match `match_index` for cell `cell_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when either index is out of range.
 */
FM_API fm_status_t fm_cf_results_get_match_at(const fm_cf_results_t* results, size_t cell_index, size_t match_index,
                                              fm_cf_match_info_t* out);

/* -------------------------------------------------------------------------- */
/* Conditional formatting — mutation                                          */
/* -------------------------------------------------------------------------- */

/** @brief CF rule type (`fm_cf_rule_t::type`). Mirrors `formulon::cf::RuleType`. */
typedef enum {
  FM_CF_RULE_TYPE_EXPRESSION = 0,
  FM_CF_RULE_TYPE_CELL_IS = 1,
  FM_CF_RULE_TYPE_COLOR_SCALE = 2,
  FM_CF_RULE_TYPE_DATA_BAR = 3,
  FM_CF_RULE_TYPE_ICON_SET = 4,
  FM_CF_RULE_TYPE_TOP10 = 5,
  FM_CF_RULE_TYPE_ABOVE_AVERAGE = 6,
  FM_CF_RULE_TYPE_CONTAINS_TEXT = 7,
  FM_CF_RULE_TYPE_NOT_CONTAINS_TEXT = 8,
  FM_CF_RULE_TYPE_BEGINS_WITH = 9,
  FM_CF_RULE_TYPE_ENDS_WITH = 10,
  FM_CF_RULE_TYPE_CONTAINS_BLANKS = 11,
  FM_CF_RULE_TYPE_NOT_CONTAINS_BLANKS = 12,
  FM_CF_RULE_TYPE_CONTAINS_ERRORS = 13,
  FM_CF_RULE_TYPE_NOT_CONTAINS_ERRORS = 14,
  FM_CF_RULE_TYPE_TIME_PERIOD = 15,
  FM_CF_RULE_TYPE_DUPLICATE_VALUES = 16,
  FM_CF_RULE_TYPE_UNIQUE_VALUES = 17,
} fm_cf_rule_type_t;

/**
 * @brief Relative date period (`fm_cf_rule_t::time_period`), in OOXML
 *        `ST_TimePeriod` order. `fm_dynamic_filter_type_t` and
 *        `fm_pivot_filter_kind_t` spell every period they share with it the
 *        same way.
 */
typedef enum {
  FM_DATE_PERIOD_TODAY = 0,
  FM_DATE_PERIOD_YESTERDAY = 1,
  FM_DATE_PERIOD_TOMORROW = 2,
  FM_DATE_PERIOD_LAST_7_DAYS = 3,
  FM_DATE_PERIOD_THIS_MONTH = 4,
  FM_DATE_PERIOD_LAST_MONTH = 5,
  FM_DATE_PERIOD_NEXT_MONTH = 6,
  FM_DATE_PERIOD_THIS_WEEK = 7,
  FM_DATE_PERIOD_LAST_WEEK = 8,
  FM_DATE_PERIOD_NEXT_WEEK = 9,
} fm_date_period_t;

/** @brief `<cfvo type>` (`fm_cfvo_t::type`). Mirrors `formulon::cf::CfvoType`. */
typedef enum {
  FM_CF_VALUE_OBJECT_TYPE_NUMBER = 0,
  FM_CF_VALUE_OBJECT_TYPE_PERCENT = 1,
  FM_CF_VALUE_OBJECT_TYPE_PERCENTILE = 2,
  FM_CF_VALUE_OBJECT_TYPE_MIN = 3,
  FM_CF_VALUE_OBJECT_TYPE_MAX = 4,
  FM_CF_VALUE_OBJECT_TYPE_FORMULA = 5,
  FM_CF_VALUE_OBJECT_TYPE_AUTO_MIN = 6,
  FM_CF_VALUE_OBJECT_TYPE_AUTO_MAX = 7,
} fm_cf_value_object_type_t;

/**
 * @brief Conditional-format value object used by visual rule authoring.
 */
typedef struct {
  /** The kind of threshold, which decides how `value` is read (a number, a percentage, a percentile, a formula, or the
   *  minimum or maximum of the data). */
  int32_t type;
  /** An `fm_tristate_t`: `DEFAULT` or `TRUE` is an inclusive threshold, `FALSE` an exclusive one; a read returns `TRUE`
   *  or `FALSE`. */
  int32_t gte; /* fm_tristate_t */
  /** The number, percentage or formula the `type` reads, copied on add; `NULL` for a type that reads none (minimum,
   *  maximum). A formula reads back with a leading `=` and may be given with or without it. */
  const char* value;
} fm_cfvo_t;

/**
 * @brief CF rule wire format used by `fm_sheet_add_cf_rule`, `fm_sheet_get_cf_rule_at`
 *        and `fm_sheet_get_cf_rule`.
 *
 * Wide POD covering both differential-format rules and visual rules.
 * `fm_sheet_add_cf_rule` deep-copies every pointer-backed payload into
 * the engine model. `fm_sheet_get_cf_rule_at` / `fm_sheet_get_cf_rule` return every pointer in the
 * selected rule -- id, ranges, formula/text strings, and the visual
 * threshold arrays and their value strings alike -- as a view into
 * handle-owned storage, not the workbook model; see that function's own
 * doc for the full lifetime contract (valid until the next successful
 * `fm_sheet_get_cf_rule_at` / `fm_sheet_get_cf_rule` on the same handle, or handle destruction; CF
 * mutations, unrelated reads, and a validation-rejected call do not
 * invalidate it).
 *
 * Active fields by `type`:
 *   - `Expression` (0): `formula1`.
 *   - `CellIs` (1): `has_op` + `op` + `formula1` (+ `formula2` for
 *     `Between` / `NotBetween`).
 *   - `Top10` (5): `has_rank` + `rank` + `percent` + `bottom`.
 *   - `AboveAverage` (6): `above_average` + `equal_average`
 *     + (optional) `has_std_dev` + `std_dev`.
 *   - `ContainsText` / `NotContainsText` / `BeginsWith` / `EndsWith`
 *     (7-10): `text`.
 *   - `ContainsBlanks` / `NotContainsBlanks` / `ContainsErrors` /
 *     `NotContainsErrors` (11-14): no extra payload.
 *   - `TimePeriod` (15): `has_time_period` + `time_period`.
 *   - `DuplicateValues` / `UniqueValues` (16-17): no extra payload.
 *   - `ColorScale` (2): `color_scale_thresholds` +
 *     `color_scale_colors` with matching counts of 2 or 3.
 *   - `DataBar` (3): `data_bar_min`, `data_bar_max`,
 *     `data_bar_fill`, optional display flags, and the `x14` extension
 *     block at the end of this struct (axis, gradient, negative and
 *     border colours), each governed by its own `has_*` flag.
 *   - `IconSet` (4): `icon_set` plus N-1
 *     `icon_set_thresholds`, and the `icon_set_floor` at the end of this
 *     struct.
 *
 * String fields use C-string convention: `NULL` means "absent",
 * non-`NULL` is a NUL-terminated view. On the input path
 * (`fm_sheet_add_cf_rule`) the caller owns the buffer until the call
 * returns; on the output path (`fm_sheet_get_cf_rule_at`) every field is a
 * view into handle-owned storage per that function's lifetime contract
 * (see its doc), not the workbook model.
 */
typedef struct {
  uint32_t struct_size;
  /* Stable rule id, linking the rule to the `<x14:cfRule id="...">`
   * block that carries whatever settings the legacy OOXML schema cannot
   * express, and the key `fm_sheet_get_cf_rule`, `fm_sheet_update_cf_rule`
   * and `fm_sheet_remove_cf_rule` take. On input it is a brace-wrapped GUID
   * (`{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`, hex digits in either case);
   * `NULL`, `""` or an id already used on the sheet makes the engine
   * generate an upper-case one. On output, always a non-NULL view into
   * `fm_sheet_get_cf_rule_at`'s handle-owned storage (see that function's doc).
   *
   * Only rules that need such a block put their id in the saved file, so
   * for most rules this is an in-memory handle and nothing more. */
  /** The rule's identifier, a GUID unique within the workbook; empty or `NULL` on add lets the engine assign one. */
  const char* id;
  /** The kind of rule, an `fm_cf_rule_type_t`; selects which of the other fields apply. */
  int32_t type;
  /** The comparison operator of a cell-is rule, an `fm_comparison_operator_t`; read only when `has_op` is 1. */
  int32_t op;
  /** The date period of a time-period rule, an `fm_date_period_t`; read only when `has_time_period` is 1. */
  int32_t time_period;
  /** The workbook-global priority of the rule; a smaller number ranks higher. On add, 0 or less lets the engine assign
   *  one after the existing rules. */
  int32_t priority;
  /** 1 when lower-priority rules are not applied to a cell this rule matches; 0 otherwise. */
  int32_t stop_if_true;
  /** 1 when `dxf_index` is set; 0 means the rule applies no differential format. */
  int32_t has_dxf_index;
  /** The 0-based index of the differential format the rule applies; read only when `has_dxf_index` is 1. */
  uint32_t dxf_index;

  /* sqref union — at least one entry. On input, must be non-NULL with
   * range_count >= 1. On output, points to `fm_sheet_get_cf_rule_at`'s
   * handle-owned storage (see that function's doc), not the model. */
  /** The cell ranges the rule applies to; `range_count` entries. */
  const fm_cell_rect_t* ranges;
  /** The number of entries in `ranges`. */
  uint32_t range_count;

  /* Formula sources (Expression / CellIs / containsText-derived). On
   * output each non-empty one carries a leading `=`; input takes it with or
   * without one. */
  /** The first formula of the rule; leading `=` included on read, optional on input. `NULL` when the rule has none. */
  const char* formula1;
  /** The second formula, for the between and not-between operators; `NULL` otherwise. */
  const char* formula2;
  /** 1 when `op` is set; 0 means the rule has no comparison operator. */
  int32_t has_op;

  /* Top10 payload */
  /** 1 when `rank` is set; 0 means the rule has no rank. */
  int32_t has_rank;
  /** The number of values (or percent of values with `percent`) a top/bottom rule selects; read only when `has_rank` is
   *  1. */
  int32_t rank;
  /** 1 when `rank` is a percentage of the values rather than a count; 0 otherwise. */
  int32_t percent;
  /** 1 when a top/bottom rule selects the lowest values; 0 selects the highest. */
  int32_t bottom;

  /* AboveAverage payload */
  /** 1 for values above the average, 0 for values below it; defaults to 1. */
  int32_t above_average;
  /** 1 when values equal to the average also match; 0 otherwise. */
  int32_t equal_average;
  /** 1 when `std_dev` is set; 0 means a plain comparison with the average. */
  int32_t has_std_dev;
  /** The number of standard deviations from the average a value must differ by; read only when `has_std_dev` is 1. */
  double std_dev;

  /* ContainsText / BeginsWith / EndsWith / NotContainsText literal */
  /** The text a text rule tests for; `NULL` when the rule has none. */
  const char* text;

  /* TimePeriod */
  /** 1 when `time_period` is set; 0 means the rule has no date period. */
  int32_t has_time_period;

  /* ColorScale payload. Counts must match and be 2 or 3 on input. */
  /** The two or three thresholds of a colour-scale rule, from lowest to highest; `color_scale_count` entries. */
  const fm_cfvo_t* color_scale_thresholds;
  /** The colour at each colour-scale threshold, in the same order; `color_scale_count` entries. */
  const fm_color_t* color_scale_colors;
  /** The number of entries in `color_scale_thresholds` and in `color_scale_colors`, 2 or 3. */
  uint32_t color_scale_count;

  /* DataBar payload. */
  /** 1 when the data-bar fields are set; 0 means the rule has no data-bar payload. */
  int32_t has_data_bar;
  /** The threshold at which the data bar is shortest. */
  fm_cfvo_t data_bar_min;
  /** The threshold at which the data bar is longest. */
  fm_cfvo_t data_bar_max;
  /** The fill colour of the data bar. */
  fm_color_t data_bar_fill;
  /** 1 (the default) shows the cell value beside the data bar; 0 hides it. */
  int32_t data_bar_show_value;
  /** The shortest data-bar length as a percentage of the cell width; defaults to 10. */
  uint8_t data_bar_min_length_pct;
  /** The longest data-bar length as a percentage of the cell width; defaults to 90. */
  uint8_t data_bar_max_length_pct;
  /** Padding; always 0. */
  uint8_t _pad1[2];

  /* IconSet payload. */
  /** 1 when the icon-set fields are set; 0 means the rule has no icon-set payload. */
  int32_t has_icon_set;
  /** The icon set the rule draws from, an `fm_icon_set_name_t`. */
  int32_t icon_set;
  /** The thresholds between the icons, N-1 for an icon set of N icons; `icon_set_threshold_count` entries. */
  const fm_cfvo_t* icon_set_thresholds;
  /** The number of entries in `icon_set_thresholds`. */
  uint32_t icon_set_threshold_count;
  /** 1 draws the icons in reverse order; 0 otherwise. */
  int32_t icon_set_reverse;
  /** 1 (the default) shows the cell value beside the icon; 0 hides it. */
  int32_t icon_set_show_value;
  /* `percent` attribute on `<iconSet>`, carried round-trip only: it is read,
   * exposed and re-emitted verbatim, and is never consulted during
   * evaluation. Each `fm_cfvo_t` names its own `type`, and that type is the
   * authoritative interpretation of the threshold, so setting this field
   * changes the saved XML but not the icon bucketing. To change the
   * interpretation, change the `fm_cfvo_t::type` values instead. */
  /** Round-trip only: preserved across load and save but never consulted during evaluation. Defaults to 1. */
  int32_t icon_set_percent;

  /* DataBar extension payload.
   *
   * The legacy `<dataBar>` element carries only min/max/fill/lengths; the
   * axis, gradient and negative colours live in the `x14` extension. They
   * are appended here rather than beside the other `data_bar_*` fields so
   * every field above keeps the offset it already had, and a consumer that
   * reads only the older prefix stays correct.
   *
   * The gradient and axis-position fields carry a `has_*` flag that
   * distinguishes "caller did not specify" from an explicit value. Each of
   * the four colours reads `FM_COLOR_KIND_NONE` when absent; on input NONE
   * means "not specified": the negative fill reuses the positive fill, and
   * there is no border, no negative border and a black axis colour. A
   * zero-initialized `fm_cf_rule_t` therefore gives gradient on, automatic
   * axis and those colour defaults.
   *
   * `fm_sheet_get_cf_rule_at` engages both flags and fills every colour that
   * is set, so feeding its output straight back into `fm_sheet_add_cf_rule`
   * reproduces the rule exactly. */
  /** 1 when `data_bar_gradient` is set; 0 means the rule states none, which means a gradient fill. */
  int32_t has_data_bar_gradient;
  /** 1 fills the data bar with a gradient, 0 with a solid colour; read only when `has_data_bar_gradient` is 1. */
  int32_t data_bar_gradient;
  /** 1 when `data_bar_axis_position` is set; 0 means the rule states none, which means automatic. */
  int32_t has_data_bar_axis_position;
  /** Where the axis between negative and positive bars sits, an `fm_cf_data_bar_axis_position_t`; read only when
   *  `has_data_bar_axis_position` is 1. */
  int32_t data_bar_axis_position;
  /** The fill colour of negative bars; kind `NONE` reuses the positive fill. */
  fm_color_t data_bar_negative_fill;
  /** The border colour of positive bars; kind `NONE` draws no border. */
  fm_color_t data_bar_border;
  /** The border colour of negative bars; kind `NONE` draws no negative border. */
  fm_color_t data_bar_negative_border;
  /** The colour of the axis line; kind `NONE` draws a black axis. */
  fm_color_t data_bar_axis_color;

  /* Appended after the extension block for the same reason: every earlier
   * field keeps its offset.
   *
   * `icon_set_floor` is the icon set's first threshold, the lower bound of
   * the lowest icon's bucket: a cell below it gets no icon at all. When
   * `has_icon_set_floor` is 0 the floor is Excel's default, `percent 0`,
   * which every value clears. `fm_sheet_get_cf_rule_at` always engages it.
   *
   * `data_bar_direction` is the edge the bar grows from. */
  /** 1 when `icon_set_floor` is set; 0 means the rule states none, which means Excel's default, a percent threshold of
   *  0. */
  int32_t has_icon_set_floor;
  /** The lower bound of the lowest icon's bucket; a cell below it gets no icon. Read only when `has_icon_set_floor` is
   *  1. */
  fm_cfvo_t icon_set_floor;
  /** The edge the data bar grows from, an `fm_cf_data_bar_direction_t`. */
  int32_t data_bar_direction;

  /* Custom icons of the icon set (x14 `cfIcon`), appended for the same
   * reason: one entry per icon of the set, or none when the set draws its
   * own icons. `icon_set_icons` may be NULL only with a zero count. On
   * output it points into `fm_sheet_get_cf_rule_at`'s handle-owned storage. */
  /** For a custom icon set, the icon drawn for each level; `icon_set_icon_count` entries, `NULL` when the set draws its
   *  own icons. */
  const fm_cf_icon_t* icon_set_icons;
  /** The number of entries in `icon_set_icons`. */
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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cf_rule_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `cf_rule_index`-th CF rule (in flattened order) into `out`.
 *
 * Every pointer in `*out` - `id`, `ranges`, `formula1`, `formula2`, `text`,
 * the visual threshold and colour arrays, and the threshold value strings -
 * addresses storage owned by `wb`, not the workbook model. That storage is
 * released only by the next successful `fm_sheet_get_cf_rule_at` on the same
 * handle, or by handle destruction. CF mutations
 * (`fm_sheet_add_cf_rule`, `fm_sheet_remove_cf_rule`, `fm_sheet_clear_cf_rules`),
 * unrelated reads, and save / load calls all leave the record readable, so
 * the edit procedure this API provides - get a rule, remove it, add the
 * amended record back - is safe to execute without copying any field out
 * first. A validation-rejected call does not refresh the storage, so the
 * previously read rule stays valid. Copy any field that must survive the
 * next get.
 *
 * `out->dxf_index` is reported present only when it resolves against the
 * workbook's `<dxfs>` table. A rule whose stored index reaches past that
 * table - which a third-party package can carry - reads back with
 * `has_dxf_index == 0` rather than surfacing an index
 * `fm_workbook_get_dxf_at` would reject and `fm_sheet_add_cf_rule` would
 * refuse. The record therefore always feeds straight back into
 * `fm_sheet_add_cf_rule`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `cf_rule_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_cf_rule_at(const fm_workbook_t* wb, size_t sheet_index, size_t cf_rule_index,
                                           fm_cf_rule_t* out);

/**
 * @brief Reads the CF rule whose id is `id` (hex digits in either case) on
 *        `sheet_index`. `*out_present` is `1` and `*out` filled the way
 *        `fm_sheet_get_cf_rule_at` fills it when there is one, else `0` and
 *        `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * The pointers in `*out` share the handle-owned storage of
 * `fm_sheet_get_cf_rule_at`: either getter's next successful call releases
 * them.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `id` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_get_cf_rule(const fm_workbook_t* wb, size_t sheet_index, const char* id,
                                        int32_t* out_present, fm_cf_rule_t* out);

/**
 * @brief Appends a new single-rule `<conditionalFormatting>` block and
 *        writes the new rule's id to `*out_cf_rule_id`.
 *
 * The new block carries `rule.ranges` (deep-copied) and a single CF
 * rule constructed from the remaining fields. The auto-generated
 * priority (when `rule.priority <= 0`) is `existing_max + 1`. `rule.id`
 * follows the input rule of `fm_cf_rule_t::id`.
 *
 * `*out_cf_rule_id` is the id actually stored (the caller's, or the one the
 * engine generated), a read-scratch-backed view valid until the next
 * scratch-backed read. Since this call always appends a new block after
 * every existing one, the rule's flattened position equals the flattened
 * rule count observed just before the call.
 *
 * When `rule.has_dxf_index` is non-zero, `rule.dxf_index` must already name a
 * registered differential format (`fm_workbook_add_dxf`). A rule referencing a
 * missing `<dxf>` is what makes Excel discard every conditional format in
 * the workbook on open, so the reference is validated here rather than at
 * save time.
 *
 * A `DataBar` payload is rejected when either length percentage exceeds
 * 100, when `data_bar_min_length_pct > data_bar_max_length_pct` (an
 * inverted bar Excel cannot render), or when a present
 * `data_bar_axis_position` is outside `0..2`. An `IconSet` payload is
 * rejected unless it carries exactly N-1 `icon_set_thresholds` for its
 * N-icon set (the floor is `icon_set_floor`, not a threshold).
 *
 * `rule.formula1` / `rule.formula2` and formula-typed `fm_cfvo_t` values
 * may carry one leading `=`, which is dropped: the rule stores them as
 * files carry them, spelled as `fm_sheet_set_formula` stores a cell
 * formula. `fm_sheet_get_cf_rule_at` reads `formula1` / `formula2` back with
 * the `=`, so a read record re-adds the same formulas. For the text (7-10), blanks / errors (11-14) and
 * `TimePeriod` (15) types, a `NULL` or empty `rule.formula1` is filled
 * with the formula Excel stores for the rule, relative to the top-left
 * corner of the bounding box of `rule.ranges` (e.g. `NOT(ISERROR(SEARCH("ab",B1)))`).
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `rule` or `out_cf_rule_id` is `NULL`, or
 *           `rule.ranges` is `NULL` while `rule.range_count > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when `rule.range_count` exceeds the per-call range
 *           cap of 16,384 (context `part=ranges requested=<n> limit=16384`);
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `rule.range_count == 0`,
 *           when any `rule.ranges` entry is inverted (`first` past
 *           `last`) or reaches outside the grid (`row >= 1048576`,
 *           `col >= 16384`) and therefore has no A1 spelling, when
 *           `rule.id` is non-empty but not a brace-wrapped GUID, when `rule.type`
 *           / `rule.op` / `rule.time_period` / `rule.icon_set` /
 *           `rule.data_bar_axis_position` or any present `fm_cfvo_t`
 *           `type` names no declared enumerator, when `rule.dxf_index` is
 *           present but past the end of the workbook's `<dxfs>` table,
 *           when a visual payload is missing / malformed, or when
 *           `rule.priority <= 0` and an existing rule already holds
 *           `INT32_MAX`, so the auto-assigned priority would overflow;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position) when a non-empty
 * `rule.formula1`, `rule.formula2` or formula-typed `fm_cfvo_t` value does not parse (the rule is not stored).
 */
FM_API fm_status_t fm_sheet_add_cf_rule(fm_workbook_t* wb, size_t sheet_index, const fm_cf_rule_t* rule,
                                        const char** out_cf_rule_id);

/**
 * @brief Partial update of one CF rule (`fm_sheet_update_cf_rule`),
 *        following the header's patch rules (Patches, above).
 *
 * The fields are those of `fm_cf_rule_t` except `id` and `type`: a rule
 * keeps its type, and `SET` on a field its type does not use (the
 * "Active fields by `type`" list of `fm_cf_rule_t`) is refused. A `SET` value
 * follows the input rule of `fm_sheet_add_cf_rule`, and the merged rule must
 * pass its checks: the colour-scale thresholds and colours end with the same
 * count, 2 or 3. A setting the record flags with `has_<field>` is `SET` here
 * with an explicit value: `data_bar_gradient` `1`, `data_bar_axis_position`
 * automatic, an `FM_COLOR_KIND_NONE` colour and a `percent 0`
 * `icon_set_floor` restore the defaults `fm_sheet_add_cf_rule` applies.
 */
typedef struct {
  uint32_t struct_size;
  /** How an update treats `priority`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t priority_op;
  /** The workbook-global priority of the rule; a smaller number ranks higher. */
  int32_t priority;
  /** How an update treats `stop_if_true`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t stop_if_true_op;
  /** 1 when lower-priority rules are not applied to a cell this rule matches; 0 otherwise. */
  int32_t stop_if_true;
  /** How an update treats `dxf_index`: `KEEP` leaves it, `SET` writes the value beside it, `CLEAR` removes it. */
  fm_patch_op_t dxf_index_op;
  /** The 0-based index of the differential format the rule applies. */
  uint32_t dxf_index;
  /** How an update treats `ranges`: `KEEP` leaves it, `SET` replaces the whole array; `CLEAR` is refused. */
  fm_patch_op_t ranges_op;
  /** The cell ranges the rule applies to; `range_count` entries. */
  const fm_cell_rect_t* ranges;
  /** The number of entries in `ranges`. */
  uint32_t range_count;
  /** How an update treats `formula1`: `KEEP` leaves it, `SET` writes the value beside it, `CLEAR` removes it. */
  fm_patch_op_t formula1_op;
  /** The first formula of the rule; leading `=` optional. */
  const char* formula1;
  /** How an update treats `formula2`: `KEEP` leaves it, `SET` writes the value beside it, `CLEAR` removes it. */
  fm_patch_op_t formula2_op;
  /** The second formula of the rule, for the between and not-between operators; leading `=` optional. */
  const char* formula2;
  /** How an update treats `op`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t op_op;
  /** The comparison operator of a cell-is rule, an `fm_comparison_operator_t`. */
  int32_t op;
  /** How an update treats `rank`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t rank_op;
  /** The number of values (or percent of values with `percent`) a top/bottom rule selects. */
  int32_t rank;
  /** How an update treats `percent`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t percent_op;
  /** 1 when `rank` is a percentage of the values rather than a count; 0 otherwise. */
  int32_t percent;
  /** How an update treats `bottom`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t bottom_op;
  /** 1 when a top/bottom rule selects the lowest values; 0 selects the highest. */
  int32_t bottom;
  /** How an update treats `above_average`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t above_average_op;
  /** 1 for values above the average, 0 for values below it. */
  int32_t above_average;
  /** How an update treats `equal_average`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t equal_average_op;
  /** 1 when values equal to the average also match; 0 otherwise. */
  int32_t equal_average;
  /** How an update treats `std_dev`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t std_dev_op;
  /** The number of standard deviations from the average a value must differ by. */
  double std_dev;
  /** How an update treats `text`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t text_op;
  /** The text a text rule tests for. */
  const char* text;
  /** How an update treats `time_period`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t time_period_op;
  /** The date period of a time-period rule, an `fm_date_period_t`. */
  int32_t time_period;

  /** How an update treats `color_scale_thresholds`: `KEEP` leaves it, `SET` replaces the whole array; `CLEAR` is
   *  refused. */
  fm_patch_op_t color_scale_thresholds_op;
  /** The two or three thresholds of a colour-scale rule; `color_scale_thresholds_count` entries. */
  const fm_cfvo_t* color_scale_thresholds;
  /** The number of entries in `color_scale_thresholds`. */
  uint32_t color_scale_thresholds_count;
  /** How an update treats `color_scale_colors`: `KEEP` leaves it, `SET` replaces the whole array; `CLEAR` is refused.
   */
  fm_patch_op_t color_scale_colors_op;
  /** The colour at each colour-scale threshold; `color_scale_colors_count` entries. */
  const fm_color_t* color_scale_colors;
  /** The number of entries in `color_scale_colors`. */
  uint32_t color_scale_colors_count;

  /** How an update treats `data_bar_min`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t data_bar_min_op;
  /** The threshold at which the data bar is shortest. */
  fm_cfvo_t data_bar_min;
  /** How an update treats `data_bar_max`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t data_bar_max_op;
  /** The threshold at which the data bar is longest. */
  fm_cfvo_t data_bar_max;
  /** How an update treats `data_bar_fill`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t data_bar_fill_op;
  /** The fill colour of the data bar. */
  fm_color_t data_bar_fill;
  /** How an update treats `data_bar_show_value`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is
   *  refused. */
  fm_patch_op_t data_bar_show_value_op;
  /** 1 shows the cell value beside the data bar; 0 hides it. */
  int32_t data_bar_show_value;
  /** How an update treats `data_bar_min_length_pct`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is
   *  refused. */
  fm_patch_op_t data_bar_min_length_pct_op;
  /** The shortest data-bar length as a percentage of the cell width. */
  uint32_t data_bar_min_length_pct;
  /** How an update treats `data_bar_max_length_pct`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is
   *  refused. */
  fm_patch_op_t data_bar_max_length_pct_op;
  /** The longest data-bar length as a percentage of the cell width. */
  uint32_t data_bar_max_length_pct;
  /** How an update treats `data_bar_gradient`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused.
   */
  fm_patch_op_t data_bar_gradient_op;
  /** 1 fills the data bar with a gradient, 0 with a solid colour. */
  int32_t data_bar_gradient;
  /** How an update treats `data_bar_axis_position`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is
   *  refused. */
  fm_patch_op_t data_bar_axis_position_op;
  /** Where the axis between negative and positive bars sits, an `fm_cf_data_bar_axis_position_t`. */
  int32_t data_bar_axis_position;
  /** How an update treats `data_bar_negative_fill`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is
   *  refused. */
  fm_patch_op_t data_bar_negative_fill_op;
  /** The fill colour of negative bars; kind `NONE` reuses the positive fill. */
  fm_color_t data_bar_negative_fill;
  /** How an update treats `data_bar_border`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t data_bar_border_op;
  /** The border colour of positive bars; kind `NONE` draws no border. */
  fm_color_t data_bar_border;
  /** How an update treats `data_bar_negative_border`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is
   *  refused. */
  fm_patch_op_t data_bar_negative_border_op;
  /** The border colour of negative bars; kind `NONE` draws no negative border. */
  fm_color_t data_bar_negative_border;
  /** How an update treats `data_bar_axis_color`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is
   *  refused. */
  fm_patch_op_t data_bar_axis_color_op;
  /** The colour of the axis line; kind `NONE` draws a black axis. */
  fm_color_t data_bar_axis_color;
  /** How an update treats `data_bar_direction`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused.
   */
  fm_patch_op_t data_bar_direction_op;
  /** The edge the data bar grows from, an `fm_cf_data_bar_direction_t`. */
  int32_t data_bar_direction;

  /** How an update treats `icon_set`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t icon_set_op;
  /** The icon set the rule draws from, an `fm_icon_set_name_t`. */
  int32_t icon_set;
  /** How an update treats `icon_set_thresholds`: `KEEP` leaves it, `SET` replaces the whole array; `CLEAR` is refused.
   */
  fm_patch_op_t icon_set_thresholds_op;
  /** The thresholds between the icons, N-1 for an icon set of N icons; `icon_set_thresholds_count` entries. */
  const fm_cfvo_t* icon_set_thresholds;
  /** The number of entries in `icon_set_thresholds`. */
  uint32_t icon_set_thresholds_count;
  /** How an update treats `icon_set_reverse`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused.
   */
  fm_patch_op_t icon_set_reverse_op;
  /** 1 draws the icons in reverse order; 0 otherwise. */
  int32_t icon_set_reverse;
  /** How an update treats `icon_set_show_value`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is
   *  refused. */
  fm_patch_op_t icon_set_show_value_op;
  /** 1 shows the cell value beside the icon; 0 hides it. */
  int32_t icon_set_show_value;
  /** How an update treats `icon_set_percent`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused.
   */
  fm_patch_op_t icon_set_percent_op;
  /** Round-trip only: preserved across load and save but never consulted during evaluation. */
  int32_t icon_set_percent;
  /** How an update treats `icon_set_floor`: `KEEP` leaves it, `SET` writes the value beside it; `CLEAR` is refused. */
  fm_patch_op_t icon_set_floor_op;
  /** The lower bound of the lowest icon's bucket; a percent threshold of 0 restores Excel's default. */
  fm_cfvo_t icon_set_floor;
  /** How an update treats `icon_set_icons`: `KEEP` leaves it, `SET` replaces the whole array; `CLEAR` is refused. */
  fm_patch_op_t icon_set_icons_op;
  /** For a custom icon set, the icon drawn for each level; `icon_set_icons_count` entries. */
  const fm_cf_icon_t* icon_set_icons;
  /** The number of entries in `icon_set_icons`. */
  uint32_t icon_set_icons_count;
} fm_cf_rule_patch_t;

/**
 * @brief Merges `patch` into the CF rule whose id is `id` (hex digits in
 *        either case). The whole patch is checked before anything changes;
 *        the rule keeps its id, type and position.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `id` or `patch` is `NULL`, or a `SET` array is
 *           `NULL` while its count is non-zero;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no rule of the sheet has `id`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits a field that cannot be
 *           cleared, `SET` carries a `NULL` string, a `SET` field is not used by the rule's type, or a
 *           `SET` value fails a check `fm_sheet_add_cf_rule` applies;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `id` or a `SET` string is not valid UTF-8;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position) when a `SET`
 *           formula does not parse.
 */
FM_API fm_status_t fm_sheet_update_cf_rule(fm_workbook_t* wb, size_t sheet_index, const char* id,
                                           const fm_cf_rule_patch_t* patch);

/**
 * @brief Removes the CF rule whose id is `id` (hex digits in either case).
 *        When the containing block becomes empty it is removed too.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no rule of the sheet has `id`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `id` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_remove_cf_rule(fm_workbook_t* wb, size_t sheet_index, const char* id);

/**
 * @brief Removes every CF block on `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_clear_cf_rules(fm_workbook_t* wb, size_t sheet_index);

/* -------------------------------------------------------------------------- */
/* Trace precedents / dependents (dependency-graph reverse lookup)            */
/* -------------------------------------------------------------------------- */

/**
 * @brief Workbook-wide cell coordinate used by the trace API.
 *
 * Mirrors `formulon::eval::CellNodeId`.
 */
typedef struct {
  uint32_t struct_size;
  /** The 0-based index of the sheet. */
  uint32_t sheet_index;
  /** The 0-based row. */
  uint32_t row;
  /** The 0-based column. */
  uint32_t col;
} fm_cell_node_info_t;

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
 * Lifetime: owned (until fm_cell_nodes_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `(row, col)` is outside the grid.
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
 *
 * Lifetime: owned (until fm_cell_nodes_destroy).
 *
 * @return As `fm_sheet_get_precedents`.
 */
FM_API fm_status_t fm_sheet_get_dependents(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                           uint32_t depth, fm_cell_nodes_t** out);

/**
 * @brief Releases a trace results handle. `nodes == NULL` is a no-op.
 */
FM_API void fm_cell_nodes_destroy(fm_cell_nodes_t* nodes);

/**
 * @brief Writes the number of cells in the result to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `nodes` or `out_count` is `NULL`.
 */
FM_API fm_status_t fm_cell_nodes_get_node_count(const fm_cell_nodes_t* nodes, size_t* out_count);

/**
 * @brief Reads cell `index` from the result.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_cell_nodes_get_node_at(const fm_cell_nodes_t* nodes, size_t index, fm_cell_node_info_t* out);

/* -------------------------------------------------------------------------- */
/* PivotTable layout projection                                               */
/* -------------------------------------------------------------------------- */

/**
 * @brief Pivot layout cell discriminator.
 *
 * Numbering mirrors `formulon::pivot::PivotCellKind`.
 */
typedef enum {
  FM_PIVOT_CELL_KIND_HEADER = 0,
  FM_PIVOT_CELL_KIND_ROW_LABEL = 1,
  FM_PIVOT_CELL_KIND_COLUMN_LABEL = 2,
  FM_PIVOT_CELL_KIND_DATA = 3,
  FM_PIVOT_CELL_KIND_ROW_SUBTOTAL = 4,
  FM_PIVOT_CELL_KIND_COLUMN_SUBTOTAL = 5,
  FM_PIVOT_CELL_KIND_GRAND_TOTAL = 6,
  FM_PIVOT_CELL_KIND_BLANK = 7
} fm_pivot_cell_kind_t;

/**
 * @brief One projected PivotTable cell.
 *
 * `field_name` and `number_format` are NUL-terminated UTF-8 strings owned by
 * the containing `fm_pivot_report_t` handle and remain valid until that handle
 * is destroyed.
 */
typedef struct {
  uint32_t struct_size;
  /** The absolute 0-based sheet row of the cell. */
  uint32_t row;
  /** The absolute 0-based sheet column of the cell. */
  uint32_t col;
  /** The cell's value; a date item label reads as its serial number, and `number_format` tells it is a date. */
  fm_value_t value;
  /** What the cell is in the layout, an `fm_pivot_cell_kind_t`: a header, a row or column label, a data value, a
   *  subtotal, a grand total or a blank. */
  int32_t kind;
  /** The nesting depth of a header cell; 0 for data cells. */
  uint32_t depth;
  /** The name of the source field the cell belongs to; `NULL` when the cell has no source field. */
  const char* field_name;
  /** The data field's number format code of the cell; `NULL` when it has none. */
  const char* number_format;
} fm_pivot_cell_info_t;

/**
 * @brief Opaque handle for a projected PivotTable layout.
 *
 * Owns the projected cell list and every string pointer reachable from
 * `fm_pivot_cell_info_t`. Index into it via `fm_pivot_report_get_cell_count` /
 * `fm_pivot_report_get_cell_at`. Release with `fm_pivot_report_destroy`.
 */
typedef struct fm_pivot_cells fm_pivot_report_t;

/**
 * @brief Returns the number of pivot tables anchored on `sheet_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out_count` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Evaluates and projects the `pivot_index`-th PivotTable on a sheet.
 *
 * A blocked report returns the layout it would draw.
 *
 * On success the caller owns `*out` and must free it with
 * `fm_pivot_report_destroy`.
 *
 * Lifetime: owned (until fm_pivot_report_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `pivot_index` is out of range;
 *         a pivot evaluation/layout error when the table references a missing
 *         or invalid cache.
 */
FM_API fm_status_t fm_sheet_get_pivot_layout(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                             fm_pivot_report_t** out);

/**
 * @brief Releases a pivot report handle. `report == NULL` is a no-op.
 */
FM_API void fm_pivot_report_destroy(fm_pivot_report_t* report);

/**
 * @brief Writes the number of projected cells to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `report` or `out_count` is `NULL`.
 */
FM_API fm_status_t fm_pivot_report_get_cell_count(const fm_pivot_report_t* report, size_t* out_count);

/**
 * @brief Returns the rectangle the projected layout covers, in absolute
 *        sheet coordinates.
 *
 * `*out_present` is `1` and `*out_rect` the covered rectangle when the
 * layout spans at least one row and one column; otherwise `*out_present` is
 * `0` and every field of `*out_rect` is zero.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_pivot_report_get_bounds(const fm_pivot_report_t* report, int32_t* out_present,
                                              fm_cell_rect_t* out_rect);

/**
 * @brief Reads projected cell `index`.
 *
 * Lifetime: owned (until fm_pivot_report_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_pivot_report_get_cell_at(const fm_pivot_report_t* report, size_t index,
                                               fm_pivot_cell_info_t* out);

/* -------------------------------------------------------------------------- */
/* PivotCache & PivotTable mutation                                           */
/* -------------------------------------------------------------------------- */

/*
 * Common failures, besides those each entry lists:
 * `FM_STATUS_BINDING_NULL_POINTER` (7001) for a `NULL` `wb`, required string
 * or out pointer; `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) for malformed
 * UTF-8; `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) for an out-of-range
 * sheet; `FM_STATUS_NOT_FOUND` (6) for a `cache_id` no cache has;
 * `FM_STATUS_INDEX_OUT_OF_RANGE` (13) for a pivot, field, item, record,
 * data-field, filter or calculated-item index at or past its count. A
 * pivot-table edit whose report would grow over something else succeeds and
 * leaves the report blocked (Pivot tables, below).
 */

/**
 * @brief PivotTable axis enumeration.
 *
 * Numbering mirrors `formulon::pivot::PivotAxis`. `ABSENT` is a field that
 * sits on no axis.
 */
typedef enum {
  FM_PIVOT_AXIS_ABSENT = -1,
  FM_PIVOT_AXIS_ROW = 0,
  FM_PIVOT_AXIS_COLUMN = 1,
  FM_PIVOT_AXIS_VALUE = 2,
  FM_PIVOT_AXIS_PAGE = 3
} fm_pivot_axis_t;

/**
 * @brief Sibling item order of a pivot field.
 *
 * OOXML `sortType` in `ST_FieldSortType` order; zero is its default `manual`.
 */
typedef enum {
  FM_PIVOT_SORT_ORDER_MANUAL = 0,
  FM_PIVOT_SORT_ORDER_ASCENDING = 1,
  FM_PIVOT_SORT_ORDER_DESCENDING = 2
} fm_pivot_sort_order_t;

/**
 * @brief Aggregation function for a value-axis field. Numbering mirrors
 *        `formulon::pivot::Aggregation` and `SubtotalFn` (the two enums are
 *        currently 1-to-1).
 */
typedef enum {
  FM_PIVOT_AGGREGATION_SUM = 0,
  FM_PIVOT_AGGREGATION_COUNT = 1,
  FM_PIVOT_AGGREGATION_AVERAGE = 2,
  FM_PIVOT_AGGREGATION_MAX = 3,
  FM_PIVOT_AGGREGATION_MIN = 4,
  FM_PIVOT_AGGREGATION_PRODUCT = 5,
  FM_PIVOT_AGGREGATION_COUNT_NUMBERS = 6,
  FM_PIVOT_AGGREGATION_STDDEV = 7,
  FM_PIVOT_AGGREGATION_STDDEVP = 8,
  FM_PIVOT_AGGREGATION_VAR = 9,
  FM_PIVOT_AGGREGATION_VARP = 10
} fm_pivot_aggregation_t;

/**
 * @brief Show-values-as derivation applied to a data-field aggregate.
 *        Numbering mirrors `formulon::pivot::ShowValuesAs`.
 */
typedef enum {
  FM_PIVOT_SHOW_VALUES_AS_NORMAL = 0,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_ROW = 1,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_COLUMN = 2,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_TOTAL = 3,
  FM_PIVOT_SHOW_VALUES_AS_RUNNING_TOTAL_IN_ROW = 4,
  FM_PIVOT_SHOW_VALUES_AS_RUNNING_TOTAL_IN_COLUMN = 5,
  FM_PIVOT_SHOW_VALUES_AS_INDEX = 6,
  FM_PIVOT_SHOW_VALUES_AS_DIFFERENCE_FROM = 7,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_DIFFERENCE_FROM = 8,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_PARENT_ROW = 9,
  FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_PARENT_COLUMN = 10,
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
 *   * Date types (5, 29-35) take Excel date serials; `DATE_BETWEEN` is the
 *     inclusive date range.
 *   * Relative periods (36-51) and calendar positions (52-67) take no
 *     operand.
 */
typedef enum {
  FM_PIVOT_FILTER_KIND_VALUE_TOP_COUNT = 0,
  FM_PIVOT_FILTER_KIND_VALUE_GREATER_THAN = 1,
  FM_PIVOT_FILTER_KIND_VALUE_BETWEEN = 2,
  FM_PIVOT_FILTER_KIND_LABEL_CONTAINS = 3,
  FM_PIVOT_FILTER_KIND_LABEL_BEGINS_WITH = 4,
  FM_PIVOT_FILTER_KIND_DATE_BETWEEN = 5,
  FM_PIVOT_FILTER_KIND_VALUE_EQUAL = 6,
  FM_PIVOT_FILTER_KIND_VALUE_NOT_EQUAL = 7,
  FM_PIVOT_FILTER_KIND_VALUE_GREATER_THAN_OR_EQUAL = 8,
  FM_PIVOT_FILTER_KIND_VALUE_LESS_THAN = 9,
  FM_PIVOT_FILTER_KIND_VALUE_LESS_THAN_OR_EQUAL = 10,
  FM_PIVOT_FILTER_KIND_VALUE_NOT_BETWEEN = 11,
  FM_PIVOT_FILTER_KIND_LABEL_EQUAL = 12,
  FM_PIVOT_FILTER_KIND_LABEL_NOT_EQUAL = 13,
  FM_PIVOT_FILTER_KIND_LABEL_NOT_BEGINS_WITH = 14,
  FM_PIVOT_FILTER_KIND_LABEL_ENDS_WITH = 15,
  FM_PIVOT_FILTER_KIND_LABEL_NOT_ENDS_WITH = 16,
  FM_PIVOT_FILTER_KIND_LABEL_NOT_CONTAINS = 17,
  FM_PIVOT_FILTER_KIND_LABEL_GREATER_THAN = 18,
  FM_PIVOT_FILTER_KIND_LABEL_GREATER_THAN_OR_EQUAL = 19,
  FM_PIVOT_FILTER_KIND_LABEL_LESS_THAN = 20,
  FM_PIVOT_FILTER_KIND_LABEL_LESS_THAN_OR_EQUAL = 21,
  FM_PIVOT_FILTER_KIND_LABEL_BETWEEN = 22,
  FM_PIVOT_FILTER_KIND_LABEL_NOT_BETWEEN = 23,
  FM_PIVOT_FILTER_KIND_VALUE_BOTTOM_COUNT = 24,
  FM_PIVOT_FILTER_KIND_VALUE_TOP_PERCENT = 25,
  FM_PIVOT_FILTER_KIND_VALUE_BOTTOM_PERCENT = 26,
  FM_PIVOT_FILTER_KIND_VALUE_TOP_SUM = 27,
  FM_PIVOT_FILTER_KIND_VALUE_BOTTOM_SUM = 28,
  FM_PIVOT_FILTER_KIND_DATE_EQUAL = 29,
  FM_PIVOT_FILTER_KIND_DATE_NOT_EQUAL = 30,
  FM_PIVOT_FILTER_KIND_DATE_OLDER_THAN = 31,
  FM_PIVOT_FILTER_KIND_DATE_OLDER_THAN_OR_EQUAL = 32,
  FM_PIVOT_FILTER_KIND_DATE_NEWER_THAN = 33,
  FM_PIVOT_FILTER_KIND_DATE_NEWER_THAN_OR_EQUAL = 34,
  FM_PIVOT_FILTER_KIND_DATE_NOT_BETWEEN = 35,
  FM_PIVOT_FILTER_KIND_TOMORROW = 36,
  FM_PIVOT_FILTER_KIND_TODAY = 37,
  FM_PIVOT_FILTER_KIND_YESTERDAY = 38,
  FM_PIVOT_FILTER_KIND_NEXT_WEEK = 39,
  FM_PIVOT_FILTER_KIND_THIS_WEEK = 40,
  FM_PIVOT_FILTER_KIND_LAST_WEEK = 41,
  FM_PIVOT_FILTER_KIND_NEXT_MONTH = 42,
  FM_PIVOT_FILTER_KIND_THIS_MONTH = 43,
  FM_PIVOT_FILTER_KIND_LAST_MONTH = 44,
  FM_PIVOT_FILTER_KIND_NEXT_QUARTER = 45,
  FM_PIVOT_FILTER_KIND_THIS_QUARTER = 46,
  FM_PIVOT_FILTER_KIND_LAST_QUARTER = 47,
  FM_PIVOT_FILTER_KIND_NEXT_YEAR = 48,
  FM_PIVOT_FILTER_KIND_THIS_YEAR = 49,
  FM_PIVOT_FILTER_KIND_LAST_YEAR = 50,
  FM_PIVOT_FILTER_KIND_YEAR_TO_DATE = 51,
  FM_PIVOT_FILTER_KIND_Q1 = 52,
  FM_PIVOT_FILTER_KIND_Q2 = 53,
  FM_PIVOT_FILTER_KIND_Q3 = 54,
  FM_PIVOT_FILTER_KIND_Q4 = 55,
  FM_PIVOT_FILTER_KIND_M1 = 56,
  FM_PIVOT_FILTER_KIND_M2 = 57,
  FM_PIVOT_FILTER_KIND_M3 = 58,
  FM_PIVOT_FILTER_KIND_M4 = 59,
  FM_PIVOT_FILTER_KIND_M5 = 60,
  FM_PIVOT_FILTER_KIND_M6 = 61,
  FM_PIVOT_FILTER_KIND_M7 = 62,
  FM_PIVOT_FILTER_KIND_M8 = 63,
  FM_PIVOT_FILTER_KIND_M9 = 64,
  FM_PIVOT_FILTER_KIND_M10 = 65,
  FM_PIVOT_FILTER_KIND_M11 = 66,
  FM_PIVOT_FILTER_KIND_M12 = 67
} fm_pivot_filter_kind_t;

/**
 * @brief The Values pseudo-field (the "Σ Values" level) in a row / column
 *        field order, as `fm_sheet_get_pivot_row_field_order` returns it
 *        and `fm_sheet_set_pivot_row_field_order` accepts it. Its bits
 *        match the OOXML `x="-2"` field reference.
 */
#define FM_PIVOT_VALUES_FIELD 0xFFFFFFFEu

/**
 * @brief The "(previous)" and "(next)" base items of a show-values-as
 *        derivation, as `fm_pivot_data_field_t::show_as_base_item` holds
 *        them beside an item index.
 */
#define FM_PIVOT_SHOW_AS_BASE_PREVIOUS 1048828u
#define FM_PIVOT_SHOW_AS_BASE_NEXT 1048829u

/**
 * @brief "No data field to exclude" for the `exclude_data_field` argument of
 *        `fm_sheet_get_pivot_data_field_default_name`.
 */
#define FM_PIVOT_DATA_FIELD_NONE SIZE_MAX

/**
 * @brief Calendar periods of a date group, as bits of the `periods` of
 *        `fm_pivot_date_group_t`.
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
 *        spec. `FM_PIVOT_FILTER_OPERAND_KIND_ABSENT` (= -1) means the slot is unset:
 *        the payload of a relative-period or calendar-position type, or an
 *        absent upper bound. The kinds mirror
 *        `std::variant<std::monostate, int, double, std::string>`.
 */
typedef enum {
  FM_PIVOT_FILTER_OPERAND_KIND_ABSENT = -1,
  FM_PIVOT_FILTER_OPERAND_KIND_INT = 0,
  FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE = 1,
  FM_PIVOT_FILTER_OPERAND_KIND_TEXT = 2
} fm_pivot_filter_operand_kind_t;

/**
 * @brief Plain-data spec for `fm_sheet_add_pivot_data_field` /
 *        `fm_sheet_set_pivot_data_field`.
 */
typedef struct {
  uint32_t struct_size;
  /** The display name of the data field, the key GETPIVOTDATA looks it up by. When `NULL` or empty the engine names it
   *  from the workbook profile's pivot wording, such as `Sum of Sales` (en-US) or `合計 / Sales` (ja-JP), numbered 2,
   * 3,
   *  ... while the name repeats another data field's or a field's display name; that needs `field_index` to be in
   *  range. */
  const char* name;
  /** The index of the source pivot field the data field aggregates; it must be in range. */
  uint32_t field_index;
  /** The aggregation function applied to the source field. */
  int32_t aggregation;
  /** The number format code of the data field's values, under the rules of `fm_pivot_field_patch_t.number_format`;
   *  `NULL` or empty for none. */
  const char* number_format;
  /** How the value is derived from the data: as is, or as a difference, percentage, running total or rank against a
   *  base. */
  int32_t show_as;
  /** 1 when `show_as_base_field` is set; 0 means no base field. */
  int32_t has_show_as_base_field;
  /** The index of the base pivot field a relative show-as mode compares against; read only when
   *  `has_show_as_base_field` is 1. It must be in range. */
  uint32_t show_as_base_field;
  /** 1 when `show_as_base_item` is set; 0 means no base item. */
  int32_t has_show_as_base_item;
  /** The base item of a relative show-as mode: `FM_PIVOT_SHOW_AS_BASE_PREVIOUS`, `FM_PIVOT_SHOW_AS_BASE_NEXT` or an
   *  index into the base field's items; read only when `has_show_as_base_item` is 1. */
  uint32_t show_as_base_item;
} fm_pivot_data_field_t;

/**
 * @brief Read-back of a pivot field definition (`fm_sheet_get_pivot_field_at`).
 *
 * A pivot carries one field per field of its cache, at the same index, from
 * the moment it is created. Strings are model-backed views, valid until the
 * next mutation.
 */
typedef struct {
  uint32_t struct_size;
  /** The name of the bound cache field; never `NULL`. */
  const char* source_name;
  /** The report name of the field; `NULL` when the field has no report name. */
  const char* custom_name;
  /** The axis the field sits on, an `fm_pivot_axis_t`; `FM_PIVOT_AXIS_ABSENT` when it sits on no axis. */
  int32_t axis;
  /** Where subtotals render, an `fm_tristate_t`: true at the group head, false at the bottom. */
  int32_t subtotal_top; /* fm_tristate_t */
  /** The number format code the field's numFmtId resolves to, as `fm_workbook_get_number_format` resolves it in the
   *  workbook profile's vocabulary; `NULL` for none or an unregistered id. Writing it back selects the same id. */
  const char* number_format;
  /** 1 when `cache_field_index` is set; 0 means the cache does not cover this field. */
  int32_t has_cache_field_index;
  /** The index of the bound pivot-cache field; read only when `has_cache_field_index` is 1. Pivot field i binds to
   *  cache field i. */
  uint32_t cache_field_index;
  /** How the field's items are ordered, an `fm_pivot_sort_order_t`. */
  int32_t sort_order;
  /** The key of the data field the items sort by; `NULL` sorts by label. */
  const char* sort_by_field;
  /** 1 when `group_base_index` is set; 0 means the field is not a group field. */
  int32_t has_group_base_index;
  /** For a group field, the index of the field it groups; read only when `has_group_base_index` is 1. */
  uint32_t group_base_index;
  /** Whether the automatic subtotal shows; 0 whenever `subtotal_functions` is non-zero, since custom functions replace
   *  it. */
  int32_t default_subtotal;
  /** A bit mask of the custom subtotal functions: bit `1u << aggregation` is set for each `fm_pivot_aggregation_t`
   *  selected; 0 for none. */
  uint32_t subtotal_functions;
  /** Whether a blank row follows each item block of the field. */
  int32_t insert_blank_row;
  /** Whether the field's item labels repeat on every row of its block. */
  int32_t fill_down_labels;
  /** Whether the field lists items without data. */
  int32_t show_all;
  /** Whether every item and the default marker of the field are expanded; true for a field without items. Writing
   *  applies to all of them. */
  int32_t show_detail;
} fm_pivot_field_info_t;

/**
 * @brief Read-back of one pivot-field item (`fm_sheet_get_pivot_field_item_at`).
 */
typedef struct {
  uint32_t struct_size;
  /** The item's label: the one it was added with, else its shared item's value; `NULL` for the blank item or an item
   *  bound to nothing. A loaded item reads the same. */
  const char* name;
  /** 1 when `cache_index` is set; 0 means the item is bound by label only. */
  int32_t has_cache_index;
  /** The index into the bound cache field's shared items; read only when `has_cache_index` is 1. */
  uint32_t cache_index;
  /** Whether the item is shown. */
  int32_t visible;
  /** The caption shown instead of the item's value; `NULL` when it has none. */
  const char* caption;
  /** Whether the item is expanded; false for a collapsed item. */
  int32_t show_detail;
} fm_pivot_item_info_t;

/**
 * @brief Read-back of a pivot table's `<pivotTableStyleInfo>`
 *        (`fm_sheet_get_pivot_style_options`).
 */
typedef struct {
  uint32_t struct_size;
  /** The name of the table style, a model-backed view; `NULL` when the style element carries no name. */
  const char* style_name;
  /** Whether the style formats the row headers. */
  int32_t show_row_headers;
  /** Whether the style formats the column headers. */
  int32_t show_column_headers;
  /** Whether the style bands the rows. */
  int32_t show_row_stripes;
  /** Whether the style bands the columns. */
  int32_t show_column_stripes;
  /** Whether the style formats the last column. */
  int32_t show_last_column;
} fm_pivot_style_options_t;

/**
 * @brief Plain-data spec for `fm_sheet_add_pivot_filter`.
 */
typedef struct {
  uint32_t struct_size;
  /** The index of the pivot field the filter applies to; the filter takes that field's axis. */
  uint32_t field_index;
  /** The kind of the filter: a value, label, date or relative-period filter, and its comparison. */
  int32_t kind;
  /** Which payload carries the operand: `FM_PIVOT_FILTER_OPERAND_KIND_INT`, `FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE`,
   *  `FM_PIVOT_FILTER_OPERAND_KIND_TEXT`, or `FM_PIVOT_FILTER_OPERAND_KIND_ABSENT` for a relative-period or
   *  calendar-position kind. It is the lower bound of a range kind. */
  int32_t value_kind;
  /** The integer operand; read only when `value_kind` is `FM_PIVOT_FILTER_OPERAND_KIND_INT`. */
  int32_t value_int;
  /** The floating-point operand; read only when `value_kind` is `FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE`. */
  double value_double;
  /** The UTF-8 text operand; read only when `value_kind` is `FM_PIVOT_FILTER_OPERAND_KIND_TEXT`, and then non-`NULL`.
   */
  const char* value_text;
  /** Which payload carries the upper bound: `FM_PIVOT_FILTER_OPERAND_KIND_TEXT` for the label range kinds, INT or
   *  DOUBLE for the other range kinds, and `FM_PIVOT_FILTER_OPERAND_KIND_ABSENT` for a kind that is not a range. */
  int32_t value_high_kind;
  /** The integer upper bound; read only when `value_high_kind` is `FM_PIVOT_FILTER_OPERAND_KIND_INT`. */
  int32_t value_high_int;
  /** The floating-point upper bound; read only when `value_high_kind` is `FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE`. */
  double value_high_double;
  /** The index of the data field a value filter scores, validated against the pivot's data fields for value filter
   *  kinds; label and date filters ignore it. */
  uint32_t data_field_index;
  /** The UTF-8 text upper bound of a label range kind, non-`NULL` for those and `NULL` for every other kind. */
  const char* value_high_text;
} fm_pivot_filter_t;

/**
 * @brief The PivotTable Options of one pivot table
 *        (`fm_sheet_get_pivot_options`; `fm_pivot_options_patch_t` changes
 *        them).
 *
 * Flags read `0` or `1`. Strings read back are read-scratch-backed views.
 *
 * A pivot created through the API starts with `item_print_titles` on,
 * `indent` 0 and `multiple_field_filters` off, as Excel creates one.
 */
typedef struct {
  uint32_t struct_size;
  /** Whether error cells show `error_caption`; the caption is kept while this is off. */
  int32_t show_error;
  /** The text error cells show while `show_error` is on; `NULL` when unset. */
  const char* error_caption;
  /** Whether empty value cells show `missing_caption`. */
  int32_t show_missing;
  /** The text empty value cells show while `show_missing` is on; a caption of exactly `-` is written as the number 0.
   *  `NULL` when unset. */
  const char* missing_caption;
  /** Whether field captions and filter drop-downs show. */
  int32_t show_headers;
  /** Whether the expand / collapse buttons show. */
  int32_t show_drill;
  /** Whether contextual tooltips show. */
  int32_t show_data_tips;
  /** Whether the classic layout with drop zones drawn on the grid is used. */
  int32_t grid_drop_zones;
  /** Whether items without data show on rows (`showEmptyRow`); the layout follows each field's show-all setting. */
  int32_t show_empty_row;
  /** Whether items without data show on columns (`showEmptyCol`); the layout follows each field's show-all setting. */
  int32_t show_empty_column;
  /** Whether the field drop-down buttons show. */
  int32_t show_data_drop_down;
  /** Whether the expand / collapse buttons are printed. */
  int32_t print_drill;
  /** Whether field captions repeat as print titles on each printed page. */
  int32_t field_print_titles;
  /** Whether row labels repeat on each printed page. */
  int32_t item_print_titles;
  /** Whether item labels are merged and centered. */
  int32_t merge_item;
  /** The indent per level of compact-form labels, 0 to 126 (the dialog shows this plus one). */
  uint32_t indent;
  /** Whether report filter fields fill across before they go down. */
  int32_t page_over_then_down;
  /** The number of report filter fields per column or row before wrapping, 0 to 255; 0 does not wrap. */
  uint32_t page_wrap;
  /** Whether a field may carry more than one filter type. */
  int32_t multiple_field_filters;
  /** Whether cell formatting is kept when the report updates. */
  int32_t preserve_formatting;
  /** Whether custom lists take part in sorting. */
  int32_t custom_list_sort;
  /** The alternative-text title of the PivotTable; `NULL` when unset. */
  const char* alt_text;
  /** The alternative-text description of the PivotTable; `NULL` when unset. */
  const char* alt_text_summary;
  /** The grand total label shown in place of the locale's word; `NULL` when unset. */
  const char* grand_total_caption;
  /** Whether column widths autofit when the report updates. */
  int32_t auto_formatting;
  /** Whether Show Details is allowed. */
  int32_t enable_drill;
} fm_pivot_options_t;

/**
 * @brief Partial update of a pivot table's options
 *        (`fm_sheet_update_pivot_options`), following the header's patch
 *        rules (Patches, above); the members mean what they mean in
 *        `fm_pivot_options_t`.
 */
typedef struct {
  uint32_t struct_size;
  /** How `show_error` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t show_error_op;
  /** Whether error cells show `error_caption`; the caption is kept while this is off (0 or non-zero when written). Read
   *  when `show_error_op` is `FM_PATCH_OP_SET`. */
  int32_t show_error;
  /** How `error_caption` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it (an empty string unsets
   *  it), `FM_PATCH_OP_CLEAR` unsets it. */
  fm_patch_op_t error_caption_op;
  /** The text to write to the option. A set value is at most 255 UTF-16 code units. Read when `error_caption_op` is
   *  `FM_PATCH_OP_SET`. */
  const char* error_caption;
  /** How `show_missing` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t show_missing_op;
  /** Whether empty value cells show `missing_caption` (0 or non-zero when written). Read when `show_missing_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t show_missing;
  /** How `missing_caption` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it (an empty string
   *  unsets it), `FM_PATCH_OP_CLEAR` unsets it. */
  fm_patch_op_t missing_caption_op;
  /** The text to write to the option. A set value is at most 255 UTF-16 code units. Read when `missing_caption_op` is
   *  `FM_PATCH_OP_SET`. */
  const char* missing_caption;
  /** How `show_headers` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t show_headers_op;
  /** Whether field captions and filter drop-downs show (0 or non-zero when written). Read when `show_headers_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t show_headers;
  /** How `show_drill` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t show_drill_op;
  /** Whether the expand / collapse buttons show (0 or non-zero when written). Read when `show_drill_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t show_drill;
  /** How `show_data_tips` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t show_data_tips_op;
  /** Whether contextual tooltips show (0 or non-zero when written). Read when `show_data_tips_op` is `FM_PATCH_OP_SET`.
   */
  int32_t show_data_tips;
  /** How `grid_drop_zones` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t grid_drop_zones_op;
  /** Whether the classic layout with drop zones drawn on the grid is used (0 or non-zero when written). Read when
   *  `grid_drop_zones_op` is `FM_PATCH_OP_SET`. */
  int32_t grid_drop_zones;
  /** How `show_empty_row` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t show_empty_row_op;
  /** Whether items without data show on rows; the layout follows each field's own show-all setting (0 or non-zero when
   *  written). Read when `show_empty_row_op` is `FM_PATCH_OP_SET`. */
  int32_t show_empty_row;
  /** How `show_empty_column` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t show_empty_column_op;
  /** Whether items without data show on columns; the layout follows each field's own show-all setting (0 or non-zero
   *  when written). Read when `show_empty_column_op` is `FM_PATCH_OP_SET`. */
  int32_t show_empty_column;
  /** How `show_data_drop_down` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it;
   *  `FM_PATCH_OP_CLEAR` is rejected. */
  fm_patch_op_t show_data_drop_down_op;
  /** Whether the field drop-down buttons show (0 or non-zero when written). Read when `show_data_drop_down_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t show_data_drop_down;
  /** How `print_drill` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t print_drill_op;
  /** Whether the expand / collapse buttons are printed (0 or non-zero when written). Read when `print_drill_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t print_drill;
  /** How `field_print_titles` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it;
   *  `FM_PATCH_OP_CLEAR` is rejected. */
  fm_patch_op_t field_print_titles_op;
  /** Whether field captions repeat as print titles on each printed page (0 or non-zero when written). Read when
   *  `field_print_titles_op` is `FM_PATCH_OP_SET`. */
  int32_t field_print_titles;
  /** How `item_print_titles` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t item_print_titles_op;
  /** Whether row labels repeat on each printed page (0 or non-zero when written). Read when `item_print_titles_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t item_print_titles;
  /** How `merge_item` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t merge_item_op;
  /** Whether item labels are merged and centered (0 or non-zero when written). Read when `merge_item_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t merge_item;
  /** How `indent` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t indent_op;
  /** The indent per level to write, at most 126 unless it equals the current value. Read when `indent_op` is
   *  `FM_PATCH_OP_SET`. */
  uint32_t indent;
  /** How `page_over_then_down` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it;
   *  `FM_PATCH_OP_CLEAR` is rejected. */
  fm_patch_op_t page_over_then_down_op;
  /** Whether report filter fields fill across before they go down (0 or non-zero when written). Read when
   *  `page_over_then_down_op` is `FM_PATCH_OP_SET`. */
  int32_t page_over_then_down;
  /** How `page_wrap` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t page_wrap_op;
  /** The number of report filter fields per column or row to write, at most 255 unless it equals the current value; 0
   *  does not wrap. Read when `page_wrap_op` is `FM_PATCH_OP_SET`. */
  uint32_t page_wrap;
  /** How `multiple_field_filters` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it;
   *  `FM_PATCH_OP_CLEAR` is rejected. */
  fm_patch_op_t multiple_field_filters_op;
  /** Whether a field may carry more than one filter type (0 or non-zero when written). Read when
   *  `multiple_field_filters_op` is `FM_PATCH_OP_SET`. */
  int32_t multiple_field_filters;
  /** How `preserve_formatting` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it;
   *  `FM_PATCH_OP_CLEAR` is rejected. */
  fm_patch_op_t preserve_formatting_op;
  /** Whether cell formatting is kept when the report updates (0 or non-zero when written). Read when
   *  `preserve_formatting_op` is `FM_PATCH_OP_SET`. */
  int32_t preserve_formatting;
  /** How `custom_list_sort` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t custom_list_sort_op;
  /** Whether custom lists take part in sorting (0 or non-zero when written). Read when `custom_list_sort_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t custom_list_sort;
  /** How `alt_text` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it (an empty string unsets it),
   *  `FM_PATCH_OP_CLEAR` unsets it. */
  fm_patch_op_t alt_text_op;
  /** The text to write to the option. A set value is at most 255 UTF-16 code units. Read when `alt_text_op` is
   *  `FM_PATCH_OP_SET`. */
  const char* alt_text;
  /** How `alt_text_summary` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it (an empty string
   *  unsets it), `FM_PATCH_OP_CLEAR` unsets it. */
  fm_patch_op_t alt_text_summary_op;
  /** The text to write to the option. Read when `alt_text_summary_op` is `FM_PATCH_OP_SET`. */
  const char* alt_text_summary;
  /** How `grand_total_caption` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it (an empty string
   *  unsets it), `FM_PATCH_OP_CLEAR` unsets it. */
  fm_patch_op_t grand_total_caption_op;
  /** The text to write to the option. A set value is at most 255 UTF-16 code units. Read when `grand_total_caption_op`
   *  is `FM_PATCH_OP_SET`. */
  const char* grand_total_caption;
  /** How `auto_formatting` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t auto_formatting_op;
  /** Whether column widths autofit when the report updates (0 or non-zero when written). Read when `auto_formatting_op`
   *  is `FM_PATCH_OP_SET`. */
  int32_t auto_formatting;
  /** How `enable_drill` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t enable_drill_op;
  /** Whether Show Details is allowed (0 or non-zero when written). Read when `enable_drill_op` is `FM_PATCH_OP_SET`. */
  int32_t enable_drill;
} fm_pivot_options_patch_t;

/**
 * @brief The options of one pivot cache, shared by every pivot table drawing
 *        from it (`fm_workbook_get_pivot_cache_options`;
 *        `fm_pivot_cache_options_patch_t` changes them).
 */
typedef struct {
  uint32_t struct_size;
  /** Whether the source records are saved with the file; off, the records part is left out and the reports keep their
   *  sheet cells. */
  int32_t save_data;
  /** 1 when `missing_items_limit` is set; 0 means automatic. */
  int32_t has_missing_items_limit;
  /** The number of items kept per field after they leave the source, 0 to 1048576; read only when
   *  `has_missing_items_limit` is 1. */
  uint32_t missing_items_limit;
  /** Whether the data refreshes when the file opens. */
  int32_t refresh_on_load;
} fm_pivot_cache_options_t;

/**
 * @brief Partial update of a pivot cache's options
 *        (`fm_workbook_update_pivot_cache_options`), following the header's
 *        patch rules (Patches, above); the members mean what they mean in
 *        `fm_pivot_cache_options_t`.
 */
typedef struct {
  uint32_t struct_size;
  /** How `save_data` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t save_data_op;
  /** Whether the source records are saved with the file; 0 or non-zero. Read when `save_data_op` is `FM_PATCH_OP_SET`.
   */
  int32_t save_data;
  /** How `missing_items_limit` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it,
   *  `FM_PATCH_OP_CLEAR` makes the limit automatic. */
  fm_patch_op_t missing_items_limit_op;
  /** The number of items kept per field after they leave the source, 0 to 1048576 or the limit the cache already holds.
   *  Read when `missing_items_limit_op` is `FM_PATCH_OP_SET`. Clearing makes the limit automatic
   *  (`has_missing_items_limit` is 0 on the read record). */
  uint32_t missing_items_limit;
  /** How `refresh_on_load` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t refresh_on_load_op;
  /** Whether the data refreshes when the file opens; 0 or non-zero. Read when `refresh_on_load_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t refresh_on_load;
} fm_pivot_cache_options_patch_t;

/**
 * @brief Read-back of one calculated item
 *        (`fm_workbook_get_pivot_cache_calculated_item_at`).
 */
typedef struct {
  uint32_t struct_size;
  /** The index of the cache field the item belongs to. */
  uint32_t field_index;
  /** The item's name; `NULL` for an entry read from a file whose item is not a text item of the field. */
  const char* name;
  /** The formula text, leading `=` included. */
  const char* formula;
  /** Whether the entry was read from a file and applies its formula only where the item meets items of other fields. */
  int32_t scoped;
} fm_pivot_calculated_item_info_t;

/* --- Pivot caches (workbook-owned) --------------------------------------- */

/**
 * @brief Returns the number of pivot caches owned by the workbook.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Returns the `cache_id` of the pivot cache at flat index `index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_id_at(const fm_workbook_t* wb, size_t index, uint32_t* out_cache_id);

/**
 * @brief Creates a new empty pivot cache.
 *
 * A cache id is never `0`; `requested_id` `0` asks the engine to assign one,
 * `max(existing_ids) + 1` (or `1` if no caches exist). When non-zero,
 * `requested_id` must not collide with an existing cache. The assigned
 * id is written to `*out_cache_id`.
 *
 * A new cache has no worksheet source. A cache no pivot table draws from is
 * not written, so `fm_workbook_save` succeeds and the saved package carries
 * no trace of it. Once a pivot table draws from a cache without a source, an
 * xlsx save returns `FM_STATUS_IO_WRITE_FAILED` (5016) rather than write a
 * package Excel offers to repair on open; name one with
 * `fm_workbook_set_pivot_cache_worksheet_source`. The range only has to name
 * where the records came from; it need not hold data, because the records
 * carry the values themselves.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when another cache holds `requested_id`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when auto-assignment would overflow.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache(fm_workbook_t* wb, uint32_t requested_id, uint32_t* out_cache_id);

/**
 * @brief Removes the pivot cache with id `cache_id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_NOT_FOUND` (6) when no cache matches `cache_id`;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the cache is still referenced by
 *         at least one pivot table.
 */
FM_API fm_status_t fm_workbook_remove_pivot_cache(fm_workbook_t* wb, uint32_t cache_id);

/**
 * @brief A cache's `<worksheetSource>`, as `fm_workbook_get_pivot_cache_worksheet_source`
 *        reads it and `fm_workbook_set_pivot_cache_worksheet_source` writes it.
 *
 * `ref`, `sheet` and `name` are the element's attributes as text; each may
 * be `NULL` (absent) on input. On read the strings are model-backed views
 * owned by the workbook handle.
 */
typedef struct {
  uint32_t struct_size;
  /** The source range as A1 text; `NULL` when absent. */
  const char* ref;
  /** The name of the source sheet, kept as written (it may name a sheet the workbook does not have); `NULL` when
   *  absent. */
  const char* sheet;
  /** The defined name or table the source reads; `NULL` when absent. */
  const char* name;
} fm_pivot_worksheet_source_t;

/**
 * @brief Reads the cache's worksheet source into `*out`.
 *
 * `*out_present` is non-zero when a `<worksheetSource>` is present; when it
 * is `0`, `*out` is zeroed apart from `struct_size`. The views in `*out`
 * remain valid until the cache is mutated or the handle is destroyed. Reads
 * do not invalidate them.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_present` or `out` is `NULL`;
 *         `FM_STATUS_NOT_FOUND` (6) when no cache matches `cache_id`.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_worksheet_source(const fm_workbook_t* wb, uint32_t cache_id,
                                                                int32_t* out_present, fm_pivot_worksheet_source_t* out);

/**
 * @brief Sets or clears the cache's worksheet source.
 *
 * `source == NULL` clears `<worksheetSource>`. Otherwise the strings are
 * copied, `NULL` means absent, and the cache becomes a worksheet-type
 * source: any other `<cacheSource>` type, attribute or child captured on
 * load is dropped.
 *
 * Records come from `fm_workbook_add_pivot_cache_field` /
 * `fm_workbook_add_pivot_cache_record` / `fm_workbook_set_pivot_cache_record_*`;
 * setting the source does not re-read the range, and neither
 * `fm_sheet_project_pivot` nor a cell edit does. `fm_workbook_refresh_pivot_cache`
 * rebuilds the records from the source.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_NOT_FOUND` (6) when no cache matches `cache_id`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a string is not valid UTF-8.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_worksheet_source(fm_workbook_t* wb, uint32_t cache_id,
                                                                const fm_pivot_worksheet_source_t* source);

/**
 * @brief Reads the range the cache's worksheet source resolves to now.
 *
 * Follows the source's name or table to its current range, or takes a range source as written, so it is the
 * rectangle `fm_workbook_refresh_pivot_cache` would read. A table resolves to its header row through its last data
 * row, without the totals row. `*out_present` is `0`, and `*out` zeroed, when the source does not resolve: no
 * source, a source of another kind or in another workbook, a table or name that is missing or does not name one
 * rectangle on one sheet, or a missing sheet. A source whose table or name was deleted stays as the dangling name
 * and reads absent here.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_present` or `out` is `NULL`;
 *         `FM_STATUS_NOT_FOUND` (6) when no cache matches `cache_id`.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_worksheet_source_range(const fm_workbook_t* wb, uint32_t cache_id,
                                                                      int32_t* out_present, fm_sheet_rect_t* out);

/**
 * @brief Rebuilds the cache from its worksheet source and re-projects every
 *        PivotTable drawing from it, as Excel's Refresh does.
 *
 * The source is the cache's `<worksheetSource>`: a range, a defined name or a
 * table of this workbook. Pending cells are calculated first and the source
 * values are read as they stand. The cache's fields, shared items, records
 * and `<sharedItems>` attributes, the definitions of the dependent pivots and
 * their report cells are replaced together; a failed call changes none of
 * them. A refreshed report whose new range is occupied or runs past the sheet
 * edge is blocked (Pivot tables, below); when two refreshed reports meet, the
 * one growing into the other's current report is blocked, and of two whose new
 * ranges alone overlap, the one created first renders.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_NOT_FOUND` (6) when no cache matches `cache_id`;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the source does not resolve
 *         (no source, an external or non-worksheet source, a missing sheet,
 *         name or table, a `#REF!` or dynamic name);
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when (rows - 1) x columns
 *         of the source exceeds the pivot cache record cell limit;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) also when a numeric group of the cache would
 *         hold more items than a field does (context `part=pivot_field_items requested=<n> limit=1048576`);
 *         `FM_STATUS_NOT_IMPLEMENTED` (3) when a numeric group of the cache
 *         has a start, end or interval that is not an integer.
 */
FM_API fm_status_t fm_workbook_refresh_pivot_cache(fm_workbook_t* wb, uint32_t cache_id);

/**
 * @brief Number of fields on the cache identified by `cache_id`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_count(const fm_workbook_t* wb, uint32_t cache_id,
                                                           size_t* out_count);

/**
 * @brief Reads the name of the cache field at `field_index`. The pointer is a
 *        model-backed view into the workbook handle and remains valid until
 *        the cache field list is mutated or the handle is destroyed. Reads
 *        do not invalidate it.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_name(const fm_workbook_t* wb, uint32_t cache_id,
                                                          size_t field_index, const char** out_name);

/**
 * @brief The `<sharedItems>` content flags of one cache field
 *        (`fm_workbook_get_pivot_cache_field_types`). Each is `0` or `1`.
 *
 * They are effective values: a flag the file omits reads as its schema default
 * (`contains_string`, `contains_semi_mixed_types` and `contains_non_date` are
 * `1`, the rest `0`), and a field with no flags of its own reads as the flags
 * derived from the values it holds.
 */
typedef struct {
  uint32_t struct_size;
  /** Whether the field holds a number. */
  int32_t contains_number;
  /** Whether every number the field holds is an integer. */
  int32_t contains_integer;
  /** Whether the field holds text, a boolean or an error. */
  int32_t contains_string;
  /** Whether the field holds a blank. */
  int32_t contains_blank;
  /** Whether the field holds a date. */
  int32_t contains_date;
  /** Whether the field mixes two or more of text, numbers and dates. */
  int32_t contains_mixed_types;
  /** Whether the field holds text or a blank. */
  int32_t contains_semi_mixed_types;
  /** Whether the field holds text or a number, something that is not a date. */
  int32_t contains_non_date;
} fm_pivot_cache_field_types_info_t;

/**
 * @brief Reads the content flags of the cache field at `field_index` into `*out`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_NOT_FOUND` (6) when no cache matches `cache_id`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `field_index` is not below the field count.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_types(const fm_workbook_t* wb, uint32_t cache_id,
                                                           size_t field_index, fm_pivot_cache_field_types_info_t* out);

/**
 * @brief Appends a new field with the given UTF-8 name to the cache.
 *        `out_field_index` receives the new field's index. `name`
 *        must be non-NULL. Every pivot table drawing from the cache gains an
 *        unplaced field for it.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_field(fm_workbook_t* wb, uint32_t cache_id, const char* name,
                                                     size_t* out_field_index);

/**
 * @brief Drops every field (and every record) from the cache.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_clear_pivot_cache_fields(fm_workbook_t* wb, uint32_t cache_id);

/**
 * @brief Number of shared items configured on cache field `field_index`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_shared_item_count(const fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_index, size_t* out_count);

/**
 * @brief Reads shared item `item_index` of cache field `field_index` into `out`.
 *        Date items are stored as serial numbers and read back as
 *        `FM_VALUE_KIND_NUMBER`; query `fm_workbook_is_pivot_cache_field_shared_item_date`
 *        to tell them from plain numbers. Text is read-scratch backed: the pointer remains
 *        valid until the next scratch-backed read, any mutation, or handle
 *        destruction.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_shared_item_value(const fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_index, size_t item_index,
                                                                       fm_value_t* out);

/**
 * @brief Appends a numeric shared item to cache field `field_index`
 *        and writes its index to `*out_shared_item_index`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `value` is NaN or infinite;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the field already holds the value.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_number(fm_workbook_t* wb, uint32_t cache_id,
                                                                        size_t field_index, double value,
                                                                        size_t* out_shared_item_index);

/**
 * @brief Appends a date shared item to cache field `field_index`. and writes
 *        its index to `*out_shared_item_index`. `serial` is a
 *        date serial in the workbook's date system, as a cell holds it.
 *        Saved with seconds precision. A date and the number of its serial
 *        are different items.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `serial` is not finite, is
 *         negative or is past 9999-12-31;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the field already holds the value.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_date(fm_workbook_t* wb, uint32_t cache_id,
                                                                      size_t field_index, double serial,
                                                                      size_t* out_shared_item_index);

/**
 * @brief Reports whether shared item `item_index` of cache field `field_index` is
 *        a date (`*out_is_date` is 1) or not (0).
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_is_pivot_cache_field_shared_item_date(const fm_workbook_t* wb, uint32_t cache_id,
                                                                     size_t field_index, size_t item_index,
                                                                     int32_t* out_is_date);

/**
 * @brief Appends a text shared item to cache field `field_index`
 *        and writes its index to `*out_shared_item_index`. Text that differs
 *        from an item only in case is that item.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the field already holds the value.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_text(fm_workbook_t* wb, uint32_t cache_id,
                                                                      size_t field_index, const char* text,
                                                                      size_t* out_shared_item_index);

/**
 * @brief Appends a boolean shared item to cache field `field_index`
 *        and writes its index to `*out_shared_item_index`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the field already holds the value.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_bool(fm_workbook_t* wb, uint32_t cache_id,
                                                                      size_t field_index, int32_t value,
                                                                      size_t* out_shared_item_index);

/**
 * @brief Appends a blank shared item to cache field `field_index`
 *        and writes its index to `*out_shared_item_index`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the field already holds the value.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_blank(fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_index,
                                                                       size_t* out_shared_item_index);

/**
 * @brief Appends an Excel error shared item to cache field `field_index`
 *        and writes its index to `*out_shared_item_index`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `error` names no `fm_error_code_t`;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the field already holds the value.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_field_shared_item_error(fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t field_index, fm_error_code_t error,
                                                                       size_t* out_shared_item_index);

/**
 * @brief Drops every shared item from cache field `field_index`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_clear_pivot_cache_field_shared_items(fm_workbook_t* wb, uint32_t cache_id,
                                                                    size_t field_index);

/**
 * @brief Returns the number of records (rows) on the cache.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_record_count(const fm_workbook_t* wb, uint32_t cache_id,
                                                            size_t* out_count);

/**
 * @brief Reads cell `(record_index, field_index)` of the cache into `out`.
 *
 * A cell that indexes a shared item reads the item's value. A cell the
 * record does not reach (its cell vector is shorter) reads Blank. A date
 * reads as its serial number (`FM_VALUE_KIND_NUMBER`), as shared items do;
 * `fm_workbook_is_pivot_cache_record_value_date` tells it apart. Text is
 * read-scratch backed: the pointer remains valid until the next
 * scratch-backed read, any mutation, or handle destruction.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `record_index` is not below the record count or
 *           `field_index` is not below the cache's field count.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_record_value(const fm_workbook_t* wb, uint32_t cache_id,
                                                            size_t record_index, size_t field_index, fm_value_t* out);

/**
 * @brief Tells whether cell `(record_index, field_index)` of the cache holds a
 *        date: `*out_is_date` is `1` for an inline date or a cell indexing a
 *        date shared item, else `0` (a blank cell or one past the record's
 *        end reads `0`).
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `record_index` is not below the record count or
 *           `field_index` is not below the cache's field count.
 */
FM_API fm_status_t fm_workbook_is_pivot_cache_record_value_date(const fm_workbook_t* wb, uint32_t cache_id,
                                                                size_t record_index, size_t field_index,
                                                                int32_t* out_is_date);

/**
 * @brief Appends a new empty record to the cache. `out_record_index`
 *        receives the new record's index. The record's cell vector is
 *        empty; populate it via `fm_workbook_set_pivot_cache_record_*`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_record(fm_workbook_t* wb, uint32_t cache_id, size_t* out_record_index);

/**
 * @brief Drops every record from the cache.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_clear_pivot_cache_records(fm_workbook_t* wb, uint32_t cache_id);

/*
 * The record setters below write cell `(record_index, field_index)`: the
 * record's cell vector auto-extends to `field_index + 1` with Blank fillers
 * when shorter, and `field_index` must be below the cache's field count.
 */

/**
 * @brief Sets cell `(record_index, field_index)` to a numeric value.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `value` is NaN or infinite.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_number(fm_workbook_t* wb, uint32_t cache_id, size_t record_index,
                                                             size_t field_index, double value);

/**
 * @brief Sets cell `(record_index, field_index)` to an inline date. `serial` is a
 *        date serial in the workbook's date system. The cell is written as a
 *        date value even on a field with shared items; saved with seconds
 *        precision.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `serial` is refused as
 *         `fm_workbook_add_pivot_cache_field_shared_item_date` refuses it.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_date(fm_workbook_t* wb, uint32_t cache_id, size_t record_index,
                                                           size_t field_index, double serial);

/**
 * @brief Sets cell `(record_index, field_index)` to a UTF-8 text value.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_text(fm_workbook_t* wb, uint32_t cache_id, size_t record_index,
                                                           size_t field_index, const char* text);

/**
 * @brief Sets cell `(record_index, field_index)` to a boolean value.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_bool(fm_workbook_t* wb, uint32_t cache_id, size_t record_index,
                                                           size_t field_index, int32_t value);

/**
 * @brief Sets cell `(record_index, field_index)` to Blank.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_blank(fm_workbook_t* wb, uint32_t cache_id, size_t record_index,
                                                            size_t field_index);

/**
 * @brief Sets cell `(record_index, field_index)` to an Excel error value.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `error` names no `fm_error_code_t`.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_record_error(fm_workbook_t* wb, uint32_t cache_id, size_t record_index,
                                                            size_t field_index, fm_error_code_t error);

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
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_options(const fm_workbook_t* wb, uint32_t cache_id,
                                                       fm_pivot_cache_options_t* out);

/**
 * @brief Merges `patch` into the cache's options. The whole patch is checked
 *        before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits a field other
 *           than `missing_items_limit`, or a `SET` `missing_items_limit` is outside the range
 *           `fm_pivot_cache_options_patch_t` states.
 */
FM_API fm_status_t fm_workbook_update_pivot_cache_options(fm_workbook_t* wb, uint32_t cache_id,
                                                          const fm_pivot_cache_options_patch_t* patch);

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
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the name equals an existing field name
 *           under the profile's collation;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position)
 *           when the formula does not parse;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the name is empty, longer than 255 UTF-16 code
 *           units, `All` or `Blank`, or contains `[` or `]`, or the formula is empty, longer than
 *           8192 UTF-16 code units, names something other than a cache field, refers to a cell,
 *           range, table, external workbook or defined name, holds an array constant or a
 *           function a calculated field cannot use, or reaches back to itself.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_calculated_field(fm_workbook_t* wb, uint32_t cache_id, const char* name,
                                                                const char* formula, size_t* out_field_index);

/**
 * @brief Replaces the formula of calculated field `field_index`.
 *
 * @return As `fm_workbook_add_pivot_cache_calculated_field` for the formula, or a common pivot
 *         failure; `FM_STATUS_INVALID_ARGUMENT` (2) when `field_index` is not a calculated field.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_calculated_field_formula(fm_workbook_t* wb, uint32_t cache_id,
                                                                        size_t field_index, const char* formula);

/**
 * @brief Removes calculated field `field_index` and its data fields in every
 *        pivot table; the fields after it move down by one.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `field_index` is not a calculated field;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) while another calculated field's formula
 *           names it.
 */
FM_API fm_status_t fm_workbook_remove_pivot_cache_calculated_field(fm_workbook_t* wb, uint32_t cache_id,
                                                                   size_t field_index);

/**
 * @brief Reads the formula of cache field `field_index`, with its leading `=`.
 *
 * `*out_present` is `1` and `*out_formula` the formula for a calculated
 * field; for any other field `*out_present` is `0` and `*out_formula` is
 * `NULL`. The formula is a read-scratch-backed view, spelled as
 * `fm_sheet_get_formula` spells formulas.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_field_formula(const fm_workbook_t* wb, uint32_t cache_id,
                                                             size_t field_index, int32_t* out_present,
                                                             const char** out_formula);

/**
 * @brief Adds calculated item `name` with `formula` to cache field
 *        `field_index`. The item follows the field's items in every pivot
 *        table and is last in solve order: where items meet, the formula of
 *        the item later in that order wins.
 *
 * The formula names items of the same field, as `North`, `'North East'`,
 * `Region[North]`, `Region[2]` (position) or `Region[-1]` (offset); a
 * leading `=` is optional. A cell
 * shows the formula applied to the named items' values in the same context,
 * and the field's totals include the item. `*out_item_index` receives the new
 * item's position in solve order (`fm_workbook_get_pivot_cache_calculated_item_at`).
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the name equals an item of the field;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position)
 *           when the formula does not parse;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a name as `fm_workbook_add_pivot_cache_calculated_field`
 *           refuses it, or for a formula it refuses or that names another field;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the field is grouped, calculated or on
 *           no pivot table's axis, or a data field drawing from the cache uses
 *           Average, StdDev, StdDevP, Var or VarP.
 */
FM_API fm_status_t fm_workbook_add_pivot_cache_calculated_item(fm_workbook_t* wb, uint32_t cache_id, size_t field_index,
                                                               const char* name, const char* formula,
                                                               size_t* out_item_index);

/**
 * @brief Number of calculated items in the cache, in solve order.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_calculated_item_count(const fm_workbook_t* wb, uint32_t cache_id,
                                                                     size_t* out_count);

/**
 * @brief Reads calculated item `item_index` (in solve order) into `out`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_get_pivot_cache_calculated_item_at(const fm_workbook_t* wb, uint32_t cache_id,
                                                                  size_t item_index,
                                                                  fm_pivot_calculated_item_info_t* out);

/**
 * @brief Replaces the formula of calculated item `item_index`.
 *
 * @return As `fm_workbook_add_pivot_cache_calculated_item` for the formula, or a common pivot
 *         failure; `FM_STATUS_INVALID_ARGUMENT` (2) when the formula refers back to the item.
 */
FM_API fm_status_t fm_workbook_set_pivot_cache_calculated_item_formula(fm_workbook_t* wb, uint32_t cache_id,
                                                                       size_t item_index, const char* formula);

/**
 * @brief Moves calculated item `item_index` to position `new_index` in solve
 *        order; `new_index` is below the calculated-item count.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_workbook_move_pivot_cache_calculated_item(fm_workbook_t* wb, uint32_t cache_id, size_t item_index,
                                                                size_t new_index);

/**
 * @brief Removes calculated item `item_index`, its item in the cache field and in
 *        every pivot table.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) while another calculated item names it.
 */
FM_API fm_status_t fm_workbook_remove_pivot_cache_calculated_item(fm_workbook_t* wb, uint32_t cache_id,
                                                                  size_t item_index);

/* --- Pivot tables (sheet-owned) ------------------------------------------ */

/*
 * A pivot table's report is kept in its sheet's cells, as Excel keeps it.
 * Every successful pivot-table mutation below writes the report into the
 * cells: those the previous report covered are cleared, then each report cell
 * is written as a value with Excel's format for its role (replacing any
 * format a cell had), and formulas reading the range recalculate. A pivot
 * table with no field on any axis draws a bordered 3 x 18 placeholder.
 *
 * The new range is occupied when it covers, outside the report's own cells,
 * a non-empty cell (a `=""` formula included), a spill range cell, any cell
 * of a merge, a legacy array block, a data table's result cells, a table or
 * another pivot table's report; a cell holding only a format is not an
 * occupant. As in Excel, an edit, a refresh or `fm_sheet_project_pivot` whose
 * range is occupied or runs past the sheet edge is not refused: the change
 * is applied, the old report's cells are cleared and the anchor holds
 * `#SPILL!` with its own format kept. The report is then blocked
 * (`fm_sheet_get_pivot_spill`) and the occupant is left as it is. The next
 * recalc after the range comes free renders it (in manual calculation, the
 * next explicit recalc); a format-only change such as an unmerge does not
 * start one. GETPIVOTDATA over a blocked report returns `#REF!`. Of two
 * reports that need the same free cells, the one created first renders.
 *
 * Creating a pivot (`fm_sheet_add_pivot`) over an occupied range is refused
 * instead, and moving one (`fm_sheet_set_pivot_anchor`) works as a cut and
 * paste. A pivot table read from a file keeps the cells read with it until
 * it is edited.
 */

/**
 * @brief A pivot table to create (`fm_sheet_add_pivot`): its name, the id of
 *        the cache it reads and the cell its report starts at.
 */
typedef struct {
  uint32_t struct_size;
  /** The new pivot's name. */
  const char* name;
  /** The id of the pivot cache the pivot draws from. */
  uint32_t cache_id;
  /** The 0-based row of the top-left cell of the report. */
  uint32_t anchor_row;
  /** The 0-based column of the top-left cell of the report. */
  uint32_t anchor_col;
} fm_pivot_spec_t;

/**
 * @brief Creates a new pivot table on `sheet_index`, with the options
 *        Excel gives a new pivot. `spec->cache_id` must reference an existing
 *        cache. The new table's flat index is written to `*out_pivot_index`.
 *
 * The pivot starts with one field per cache field and none on an axis. Place
 * them with `fm_sheet_set_pivot_field_axis` (or the row / column field order
 * setters) and add values with `fm_sheet_add_pivot_data_field`.
 *
 * The pivot takes the workbook's PivotTable style default
 * (`fm_workbook_get_table_style_defaults`), else `PivotStyleLight16`, with
 * row / column headers and last column on, banded rows / columns off.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `spec`, `spec->name`, or `out_pivot_index`
 *         is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when `spec->cache_id` does not match any existing cache;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when `spec->name` equals another pivot's name on
 *         the sheet under Unicode simple case folding;
 *         `FM_STATUS_PRECONDITION_FAILED` (9), creating nothing, when the report's range
 *         (the 3 x 18 placeholder of a pivot without fields) is occupied or runs past the
 *         sheet edge, or when the workbook's default PivotTable style does not apply to
 *         PivotTables. Creation is not an edit of an existing report, so it never blocks.
 */
FM_API fm_status_t fm_sheet_add_pivot(fm_workbook_t* wb, size_t sheet_index, const fm_pivot_spec_t* spec,
                                      size_t* out_pivot_index);

/**
 * @brief Removes the pivot table at `pivot_index` from `sheet_index`, clearing
 *        the cells and merges its report was last written to.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_remove_pivot(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Renames the pivot table to `name`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `pivot_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `name` is empty;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when `name` equals another pivot's name on the
 *         sheet under Unicode simple case folding.
 */
FM_API fm_status_t fm_sheet_rename_pivot(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index, const char* name);

/**
 * @brief Moves the pivot so its report starts at `(anchor_row, anchor_col)`,
 *        the cell `fm_sheet_add_pivot` takes; the report is rewritten there.
 *
 * As Excel's cut and paste of a report does, the move overwrites the
 * destination (the report's current rectangle moved by the same offset):
 * values, formulas, merges, spill range cells and legacy array blocks or data
 * tables lying wholly inside are cleared, and a table meeting it is deleted.
 * The report then renders there. Only an occupant outside that rectangle,
 * where a report that grows on the move can reach, blocks it as an edit does.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `pivot_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the anchor is outside the grid;
 *         `FM_STATUS_PRECONDITION_FAILED` (9), changing nothing, when the destination
 *           runs past the sheet edge, covers another pivot table, or meets a legacy array
 *           block or a data table's result range that crosses its edge.
 */
FM_API fm_status_t fm_sheet_set_pivot_anchor(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                             uint32_t anchor_row, uint32_t anchor_col);

/**
 * @brief Partial update of the row / column grand total bands
 *        (`fm_sheet_update_pivot_grand_totals`), following the header's patch
 *        rules (Patches, above). Neither field can be cleared.
 */
typedef struct {
  uint32_t struct_size;
  /** How `rows` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is rejected.
   */
  fm_patch_op_t rows_op;
  /** Whether the grand total band for the rows shows; 0 or non-zero. Read when `rows_op` is `FM_PATCH_OP_SET`. */
  int32_t rows;
  /** How `columns` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t columns_op;
  /** Whether the grand total band for the columns shows; 0 or non-zero. Read when `columns_op` is `FM_PATCH_OP_SET`. */
  int32_t columns;
} fm_pivot_grand_totals_patch_t;

/**
 * @brief Merges `patch` into the pivot's row / column grand total bands, the
 *        values `fm_sheet_get_pivot_grand_totals` reads.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2 or `CLEAR`.
 */
FM_API fm_status_t fm_sheet_update_pivot_grand_totals(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      const fm_pivot_grand_totals_patch_t* patch);

/**
 * @brief Partial update of a pivot table's style (`fm_sheet_update_pivot_style_options`).
 *
 * Each field has an `fm_patch_op_t` companion. A patch on a pivot without the
 * style element creates it, the fields it leaves at `KEEP` taking the
 * creation defaults of `fm_sheet_add_pivot`. A patch whose ops are all
 * `KEEP` changes nothing.
 */
typedef struct {
  uint32_t struct_size;
  /** How `style_name` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it, `FM_PATCH_OP_CLEAR` writes
   *  the element without a style name (no table style). */
  fm_patch_op_t style_name_op;
  /** The name of the table style to write; it must be non-empty. Read when `style_name_op` is `FM_PATCH_OP_SET`. */
  const char* style_name;
  /** How `show_row_headers` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t show_row_headers_op;
  /** Whether the style formats the row headers; 0 or non-zero. Read when `show_row_headers_op` is `FM_PATCH_OP_SET`. */
  int32_t show_row_headers;
  /** How `show_column_headers` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it;
   *  `FM_PATCH_OP_CLEAR` is rejected. */
  fm_patch_op_t show_column_headers_op;
  /** Whether the style formats the column headers; 0 or non-zero. Read when `show_column_headers_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t show_column_headers;
  /** How `show_row_stripes` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t show_row_stripes_op;
  /** Whether the style bands the rows; 0 or non-zero. Read when `show_row_stripes_op` is `FM_PATCH_OP_SET`. */
  int32_t show_row_stripes;
  /** How `show_column_stripes` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it;
   *  `FM_PATCH_OP_CLEAR` is rejected. */
  fm_patch_op_t show_column_stripes_op;
  /** Whether the style bands the columns; 0 or non-zero. Read when `show_column_stripes_op` is `FM_PATCH_OP_SET`. */
  int32_t show_column_stripes;
  /** How `show_last_column` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t show_last_column_op;
  /** Whether the style formats the last column; 0 or non-zero. Read when `show_last_column_op` is `FM_PATCH_OP_SET`. */
  int32_t show_last_column;
} fm_pivot_style_options_patch_t;

/**
 * @brief Merges `patch` into the pivot's table style and toggles. The whole
 *        patch is checked before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `pivot_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is out of range, `CLEAR` hits a
 *         toggle, `SET` carries a `NULL` or empty `style_name`, a `SET` `style_name` names a custom
 *         style that does not apply to PivotTables or a built-in slicer style, or a toggle is not 0 / 1;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `style_name` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_update_pivot_style_options(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       const fm_pivot_style_options_patch_t* patch);

/**
 * @brief Reads the pivot's compact / tabular / outline report layout, an
 *        `fm_pivot_report_layout_t`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_get_pivot_report_layout(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    int32_t* out_layout);

/**
 * @brief Sets the pivot's compact / tabular / outline report layout.
 *
 * `layout` is a raw signed 32-bit ordinal; callers do not need to construct
 * an out-of-domain `fm_pivot_report_layout_t` value.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `layout` names no `fm_pivot_report_layout_t`.
 */
FM_API fm_status_t fm_sheet_set_pivot_report_layout(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    int32_t layout);

/**
 * @brief Number of fields configured on the pivot.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_get_pivot_field_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                  size_t* out_count);

/**
 * @brief Moves every field off every axis, empties the values area and drops
 *        the active pivot filters.
 *
 * The fields themselves stay, one per cache field, each reading
 * `FM_PIVOT_AXIS_ABSENT` afterwards; `fm_sheet_get_pivot_field_count` is
 * unchanged, and a field is placed again with `fm_sheet_set_pivot_field_axis`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_clear_pivot_fields(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Sets the axis of pivot field `field_index`.
 *
 * `axis` is a raw signed 32-bit ordinal. Moving a field onto the row or column axis appends it
 * to that axis's field order (a field already there keeps its place);
 * moving it off removes it from the order it was in.
 *
 * `FM_PIVOT_AXIS_ABSENT` takes the field off every axis, as unchecking it in
 * a field list does. The data fields aggregating it are removed (later data-field
 * indices shift down), as are the filters naming the field and the
 * value filters scoring a removed data field; the field's item visibility
 * is kept.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `axis` names no `fm_pivot_axis_t` member.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_axis(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 size_t field_index, int32_t axis);

/**
 * @brief Partial update of a pivot field (`fm_sheet_update_pivot_field`).
 *
 * Each field has an `fm_patch_op_t` companion: `KEEP` (0) leaves it,
 * `SET` writes the value beside it, `CLEAR` is accepted only where stated.
 */
typedef struct {
  uint32_t struct_size;
  /** How `custom_name` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it, `FM_PATCH_OP_CLEAR`
   *  clears it. */
  fm_patch_op_t custom_name_op;
  /** The report name to write; it is non-`NULL`, at most 255 UTF-16 code units, and must not equal another field's
   *  source or custom name or a data field's name under the profile's collation (the field's own source name is
   *  allowed). `FM_PATCH_OP_CLEAR` restores the source name. Read when `custom_name_op` is `FM_PATCH_OP_SET`. */
  const char* custom_name;
  /** How `number_format` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it, `FM_PATCH_OP_CLEAR`
   *  clears it. */
  fm_patch_op_t number_format_op;
  /** The number format code to write, in the workbook profile's vocabulary; a string of digits is a format code, not an
   *  id. `FM_PATCH_OP_CLEAR` removes the format. Read when `number_format_op` is `FM_PATCH_OP_SET`. */
  const char* number_format;
  /** How `default_subtotal` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t default_subtotal_op;
  /** Whether the automatic subtotal shows; 0 or non-zero. Custom subtotal functions replace it whatever this says. Read
   *  when `default_subtotal_op` is `FM_PATCH_OP_SET`. */
  int32_t default_subtotal;
  /** How `insert_blank_row` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t insert_blank_row_op;
  /** Whether a blank row follows each item block of the field; 0 or non-zero. Read when `insert_blank_row_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t insert_blank_row;
  /** How `fill_down_labels` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR`
   *  is rejected. */
  fm_patch_op_t fill_down_labels_op;
  /** Whether the field's item labels repeat on every row of its block; 0 or non-zero. Read when `fill_down_labels_op`
   *  is `FM_PATCH_OP_SET`. */
  int32_t fill_down_labels;
  /** How `show_all` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t show_all_op;
  /** Whether the field lists items without data; 0 or non-zero. Read when `show_all_op` is `FM_PATCH_OP_SET`. */
  int32_t show_all;
  /** How `subtotal_top` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t subtotal_top_op;
  /** Whether subtotals show on the group's own heading row; 0 or non-zero. Read when `subtotal_top_op` is
   *  `FM_PATCH_OP_SET`. */
  int32_t subtotal_top;
  /** How `sort_order` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t sort_order_op;
  /** How the field's items are ordered, an `fm_pivot_sort_order_t`. Manual keeps sibling items in the field's item
   *  order and is rejected together with a non-empty `sort_by_field`. Read when `sort_order_op` is `FM_PATCH_OP_SET`.
   */
  int32_t sort_order;
  /** How `sort_by_field` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it, `FM_PATCH_OP_CLEAR`
   *  clears it. */
  fm_patch_op_t sort_by_field_op;
  /** The key of the data field the items sort by. `FM_PATCH_OP_CLEAR` sorts by label. Read when `sort_by_field_op` is
   *  `FM_PATCH_OP_SET`. */
  const char* sort_by_field;
  /** How `show_detail` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t show_detail_op;
  /** Whether every item of the field is expanded together with its subtotal row; 0 or non-zero. Read when
   *  `show_detail_op` is `FM_PATCH_OP_SET`. */
  int32_t show_detail;
} fm_pivot_field_patch_t;

/**
 * @brief Merges `patch` into pivot field `field_index`. Every field of the
 *        patch is checked first, then all are applied in one report rewrite:
 *        on any failure nothing changes. A patch whose ops are all `KEEP`
 *        (including a zero-initialised one) is a no-op.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) for malformed UTF-8;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `pivot_index` or `field_index` is out
 *           of range;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when a `SET` `custom_name` equals another field's
 *           name or a data field's name (rules above);
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is out of range, `CLEAR` hits a field that does not accept it, a
 *           flag is not 0 / 1, `sort_order` is not an `fm_pivot_sort_order_t`, `MANUAL`
 *           meets a by-field key, `SET` carries a `NULL` string, or `custom_name` is longer
 *           than 255 UTF-16 code units.
 */
FM_API fm_status_t fm_sheet_update_pivot_field(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                               size_t field_index, const fm_pivot_field_patch_t* patch);

/**
 * @brief Partial update of a pivot field item (`fm_sheet_update_pivot_field_item`).
 *
 * Ops as for `fm_pivot_field_patch_t`.
 */
typedef struct {
  uint32_t struct_size;
  /** How `visible` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t visible_op;
  /** Whether the item is shown; 0 or non-zero. Read when `visible_op` is `FM_PATCH_OP_SET`. */
  int32_t visible;
  /** How `show_detail` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t show_detail_op;
  /** Whether the item is expanded; 0 or non-zero. Read when `show_detail_op` is `FM_PATCH_OP_SET`. */
  int32_t show_detail;
  /** How `caption` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it, `FM_PATCH_OP_CLEAR` removes
   *  the caption. */
  fm_patch_op_t caption_op;
  /** The caption to show instead of the item's value, at most 255 UTF-16 code units and distinct from the label another
   *  item of the field shows under the profile's text collation; a caption equal to the item's own value removes it.
   *  Read when `caption_op` is `FM_PATCH_OP_SET`. */
  const char* caption;
} fm_pivot_item_patch_t;

/**
 * @brief Merges `patch` into item `item_index` of pivot field `field_index`,
 *        with the validate-then-apply and empty-patch rules of
 *        `fm_sheet_update_pivot_field`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when a `SET` caption matches the label another
 *           item of the field shows;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is out of range, `CLEAR` hits a
 *           field other than `caption`, a flag is not 0 / 1, `SET` carries a `NULL`
 *           `caption`, or the caption is longer than 255 UTF-16 code units.
 */
FM_API fm_status_t fm_sheet_update_pivot_field_item(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    size_t field_index, size_t item_index,
                                                    const fm_pivot_item_patch_t* patch);

/**
 * @brief Appends a manual-filter item to pivot field `field_index`, addressed
 *        by its label.
 *        `name` must be non-NULL. `visible` is a 32-bit boolean.
 *
 * When a shared item of the bound cache field renders as `name`, the
 * item binds to the first such shared item and the writer emits that index as
 * `<item x="N">`; otherwise it stays unbound and the writer emits its
 * document-order position. The filter engine matches source records by
 * comparing their rendered label against `name`. That makes an empty
 * `name` unusable as the blank item: the blank has no label of its own
 * and is identified by what it binds to. Build one with
 * `fm_sheet_add_pivot_field_item_by_cache_index` instead. `*out_item_index` receives the
 * new item's position in the field's item list.
 *
 * A field placed on the row, column or page axis already lists one item per
 * shared item, in display order. To set a custom item order, call
 * `fm_sheet_clear_pivot_field_items` and add the items in that order.
 *
 * @return `FM_STATUS_OK` (0) on success, `FM_STATUS_ALREADY_EXISTS` (7) when
 *         `name` binds a shared item the field already lists (nothing
 *         changes; a label binding no shared item still appends), or a common
 *         pivot failure.
 */
FM_API fm_status_t fm_sheet_add_pivot_field_item(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 size_t field_index, const char* name, int32_t visible,
                                                 size_t* out_item_index);

/**
 * @brief Appends a manual-filter item to pivot field `field_index`, addressed
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
 * nothing keeps an empty label, filters nothing and is left out of a save.
 *
 * `visible` is a 32-bit boolean. `*out_item_index` receives the new item's
 * position in the field's item list.
 *
 * @return `FM_STATUS_OK` (0) on success, `FM_STATUS_ALREADY_EXISTS` (7) when
 *         the field already lists an item for `cache_index` (nothing
 *         changes), or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_add_pivot_field_item_by_cache_index(fm_workbook_t* wb, size_t sheet_index,
                                                                size_t pivot_index, size_t field_index,
                                                                size_t cache_index, int32_t visible,
                                                                size_t* out_item_index);

/**
 * @brief Drops every manual-filter item from pivot field `field_index`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_clear_pivot_field_items(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    size_t field_index);

/**
 * @brief Replaces the subtotal functions of pivot field `field_index` with
 *        the `count` entries of `aggregations`.
 *
 * Each entry is an `fm_pivot_aggregation_t` ordinal; the order of the list
 * does not matter, and a `count` of 0 clears the set. The selection reads
 * back through `fm_pivot_field_info_t::subtotal_functions`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure (a `NULL`
 *           `aggregations` is one only while `count > 0`);
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an entry names no `fm_pivot_aggregation_t`, or
 *           the list holds one function twice. Nothing is written on failure.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_subtotal_functions(fm_workbook_t* wb, size_t sheet_index,
                                                               size_t pivot_index, size_t field_index,
                                                               const int32_t* aggregations, size_t count);

/**
 * @brief A date grouping, as `fm_sheet_set_pivot_field_date_group` takes it and
 *        `fm_sheet_get_pivot_field_date_group` returns it.
 */
typedef struct {
  uint32_t struct_size;
  /** The date periods to group by, an OR of `FM_PIVOT_PERIOD_*` bits. */
  uint32_t periods;
  /** The calendar the periods follow, an `fm_pivot_calendar_t`. */
  int32_t calendar;
  /** The length in days of each group when grouping by days alone; 1 groups by single days. */
  uint32_t interval_days;
  /** An `fm_tristate_t`: default and true make the first bound follow the field's first value (`start` is then not
   *  read on input); false takes `start`. A read returns false or true. */
  int32_t auto_start; /* fm_tristate_t */
  /** The first bound as a date serial: read when `auto_start` is false; on a read always the stored bound. */
  double start;
  /** As `auto_start`, for the last bound: it follows the field's last value (the midnight after it for a date without
   *  a time). */
  int32_t auto_end; /* fm_tristate_t */
  /** As `start`, for the last bound. */
  double end;
} fm_pivot_date_group_t;

/**
 * @brief Groups the date field `field_index` (a source field, or one of its
 *        group fields) by `group->periods`, as Excel's Group command does,
 *        replacing any grouping it has; a `NULL` `group` removes the grouping.
 *
 * Each period becomes a group field appended to the pivot cache, finest
 * first, named and labelled in the workbook profile's locale. Every pivot on
 * the cache gains them: this pivot shows them in the field's place on its
 * axis, coarsest outermost; another pivot showing the field shows them
 * before it. `interval_days` above 1 groups by runs of that many days and
 * needs `periods == FM_PIVOT_PERIOD_DAYS`.
 * `FM_PIVOT_CALENDAR_JAPANESE` groups the years period by era year.
 *
 * The field must hold only dates (values from date-formatted cells, or record values stored as
 * dates). A plain number is refused with `FM_STATUS_PRECONDITION_FAILED` even when it is a valid
 * serial; Excel's object model silently groups such a field by 1 numerically instead, which is
 * `fm_sheet_set_pivot_field_numeric_group` with an interval of 1.
 *
 * With `group == NULL` the grouping is removed from every pivot on the
 * cache and the field returns to the axis position its group fields held;
 * a field that is not grouped is left as it is.
 *
 * @return `FM_STATUS_OK` (0) on success, also when the grown report cannot be placed (an
 *           occupied range, the sheet edge): the grouping is applied and the report is withheld
 *           with `#SPILL!` at its anchor; every refusal below is returned before anything changes;
 *         a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for `periods` of 0
 *         or with an unknown bit, an unknown `calendar`, `interval_days == 0`,
 *         `interval_days > 1` with any period but days, `auto_start` / `auto_end`
 *         outside `fm_tristate_t`, a NaN or infinite bound with its `auto_*` false, a bound that is
 *         not a valid date serial, or a start after the end;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the field holds a blank, a value that is not a
 *         date (a plain number included), or no value, is a calculated field, or is grouped and its group fields are
 * not the last fields of the cache.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_date_group(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       size_t field_index, const fm_pivot_date_group_t* group);

/**
 * @brief Reads the date grouping over the field `field_index` (a source field,
 *        or one of its group fields) into `*out`.
 *
 * `*out_present` is `0` when the field is not date-grouped, and `*out` is then
 * left zeroed apart from `struct_size`. Otherwise `*out` holds what
 * `fm_sheet_set_pivot_field_date_group` takes, so the record can be written
 * back unchanged: `auto_start` / `auto_end` read `FM_TRISTATE_TRUE` when the
 * bound follows the data and `FM_TRISTATE_FALSE` otherwise, and `start` /
 * `end` hold the stored bound either way.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_present` or `out` is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_pivot_field_date_group(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       size_t field_index, int32_t* out_present,
                                                       fm_pivot_date_group_t* out);

/**
 * @brief A numeric grouping, as `fm_sheet_set_pivot_field_numeric_group` takes
 *        it and `fm_sheet_get_pivot_field_numeric_group` returns it.
 */
typedef struct {
  uint32_t struct_size;
  /** An `fm_tristate_t`: default and true make the first bound follow the field's smallest value (`start` is then
   *  not read on input); false takes `start`. A read returns false or true. */
  int32_t auto_start; /* fm_tristate_t */
  /** The first bound: read when `auto_start` is false; on a read always the stored bound. */
  double start;
  /** As `auto_start`, for the last bound, which follows the field's largest value. */
  int32_t auto_end; /* fm_tristate_t */
  /** As `start`, for the last bound. */
  double end;
  /** The width of each group. */
  double interval;
} fm_pivot_numeric_group_t;

/**
 * @brief Groups the number field `field_index` (a source field, or its group
 *        field) into ranges of `group->interval`, replacing any grouping it
 *        has; a `NULL` `group` removes the grouping.
 *
 * The group field's items, as a refresh regenerates them, are `<start`, one
 * range per step from `start` up to the first edge at or past `end` (at least
 * one), then `>` the upper edge of the last range. When `start`, `end` and
 * `interval` are all whole numbers, a range whose upper edge `a+interval` is
 * at most 1e15 reads `a-(a+interval-1)`, the last one `a-end` when `end` is
 * its upper edge; every other range reads `a-(a+interval)`. Each number is
 * spelled with 15 significant digits, in full while the text of its magnitude
 * has at most 20 characters and in exponent form otherwise (`1.3E+20`,
 * `2.43E-17`), with the workbook locale's decimal separator (`0,3-0,8`). A
 * value equal to `end` falls in the last range, one above `end` in the `>`
 * item, and one below `start` in the `<` item. Every pivot on the cache gains
 * the group field and is left pending. With `group == NULL` the grouping is
 * removed from every pivot on the cache; a field that is not grouped is left
 * as it is.
 *
 * @return `FM_STATUS_OK` (0) on success, also when the grown report cannot be placed (an
 *           occupied range, the sheet edge): the grouping is applied and the report is withheld
 *           with `#SPILL!` at its anchor; every refusal below is returned before anything changes;
 *         a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a non-positive or non-finite `interval`,
 *         `auto_start` / `auto_end` outside `fm_tristate_t`, a NaN or infinite bound with its
 *         `auto_*` false, a start after the end, or more group items than a field holds (Excel's
 *         limit of 1,048,576 unique items; context `part=pivot_field_items requested=<n> limit=1048576`);
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the field holds a blank or a value that is
 *         not a number, holds no number to take a following bound from, is a calculated field, or
 *         is grouped and its group fields are not the last fields of the cache.
 */
FM_API fm_status_t fm_sheet_set_pivot_field_numeric_group(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                          size_t field_index, const fm_pivot_numeric_group_t* group);

/**
 * @brief Reads the numeric grouping over the field `field_index` (a source
 *        field, or its group field) into `*out`.
 *
 * `*out_present` is `0` when the field is not numerically grouped, and `*out`
 * is then left zeroed apart from `struct_size`. Otherwise the record can be
 * written back unchanged: `auto_start` / `auto_end` read `FM_TRISTATE_TRUE`
 * when the bound follows the data and `FM_TRISTATE_FALSE` otherwise, and
 * `start` / `end` hold the stored bound either way.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_present` or `out` is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_pivot_field_numeric_group(const fm_workbook_t* wb, size_t sheet_index,
                                                          size_t pivot_index, size_t field_index, int32_t* out_present,
                                                          fm_pivot_numeric_group_t* out);

/**
 * @brief Replaces the row-axis field order with `indices[0..count)`.
 *        Each entry must be `< field_count` (else
 *        `FM_STATUS_INDEX_OUT_OF_RANGE` (13)) or `FM_PIVOT_VALUES_FIELD`, and
 *        appear once (else `FM_STATUS_INVALID_ARGUMENT` (2)). The list is the axis's whole
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
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure; `FM_STATUS_INVALID_ARGUMENT`
 *         (2) as above.
 */
FM_API fm_status_t fm_sheet_set_pivot_row_field_order(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      const uint32_t* indices, size_t count);

/**
 * @brief Replaces the column-axis field order. Same contract as rows.
 *
 * @return As `fm_sheet_set_pivot_row_field_order`.
 */
FM_API fm_status_t fm_sheet_set_pivot_column_field_order(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                         const uint32_t* indices, size_t count);

/**
 * @brief Replaces the page-axis (report filter) field order. Same contract as
 *        rows, except that `FM_PIVOT_VALUES_FIELD` is
 *        `FM_STATUS_INVALID_ARGUMENT` (2): the Values level is never a page
 *        field. The saved `<pageFields>` follow this order.
 *
 * @return As `fm_sheet_set_pivot_row_field_order`.
 */
FM_API fm_status_t fm_sheet_set_pivot_page_field_order(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       const uint32_t* indices, size_t count);

/**
 * @brief Number of `<dataField>` entries on the pivot.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_get_pivot_data_field_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       size_t* out_count);

/**
 * @brief Appends a new data-field entry. A `NULL` or empty `spec->name` is
 *        replaced by the engine-derived name described on
 *        `fm_pivot_data_field_t`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `spec->field_index` or a set
 *           `spec->show_as_base_field` is not an index into the pivot's fields
 *           (the pivot is left unchanged);
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `spec->aggregation` / `spec->show_as`
 *           names no declared enumerator;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the cache has calculated items and
 *           `spec->aggregation` is Average, StdDev, StdDevP, Var or VarP.
 */
FM_API fm_status_t fm_sheet_add_pivot_data_field(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 const fm_pivot_data_field_t* spec, size_t* out_data_field_index);

/**
 * @brief Drops every data-field entry from the pivot.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_clear_pivot_data_fields(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Replaces the data-field entry at `data_field_index` in place. A
 *        `NULL` or empty `spec->name` is derived as in
 *        `fm_sheet_add_pivot_data_field`.
 *
 * @return As `fm_sheet_add_pivot_data_field`.
 */
FM_API fm_status_t fm_sheet_set_pivot_data_field(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 size_t data_field_index, const fm_pivot_data_field_t* spec);

/**
 * @brief Reads into `*out` the name a data field over pivot field
 *        `field_index` with `aggregation` would get from an empty
 *        `spec->name`, without changing anything. `exclude_data_field` is the
 *        entry `fm_sheet_set_pivot_data_field` would replace, whose name is
 *        not counted as taken; `FM_PIVOT_DATA_FIELD_NONE` for none (the add case). The string
 *        is a read-scratch-backed view.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure (an
 *           `exclude_data_field` other than `FM_PIVOT_DATA_FIELD_NONE` is an index);
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `aggregation` names no
 *           `fm_pivot_aggregation_t`.
 */
FM_API fm_status_t fm_sheet_get_pivot_data_field_default_name(const fm_workbook_t* wb, size_t sheet_index,
                                                              size_t pivot_index, size_t field_index,
                                                              int32_t aggregation, size_t exclude_data_field,
                                                              const char** out);

/*
 * Definition read-back. Every getter below is read-only and works the same
 * for a pivot created in this session and one loaded from a package. String
 * outputs are model-backed views: they remain valid until the
 * next mutation of the pivot (any `fm_sheet_set_pivot_*` setter on it), a
 * workbook model replacement, or handle destruction. Reads do not invalidate
 * them. Copy them when they must outlive those events. On any error the
 * output arguments are left unchanged.
 */

/**
 * @brief Reads the `cache_id` of the pivot cache the pivot draws from.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_cache_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) if `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_cache_id(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                               uint32_t* out_cache_id);

/**
 * @brief Reads the pivot's name. `*out_name` is a NUL-terminated UTF-8
 *        model-backed view, valid until the next mutation of the pivot or its
 *        sheet, or until the handle is destroyed.
 *
 * Lifetime: transient.
 *
 * @return As `fm_sheet_get_pivot_cache_id`, with `out_name` in place of
 *         `out_cache_id`.
 */
FM_API fm_status_t fm_sheet_get_pivot_name(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                           const char** out_name);

/**
 * @brief Writes the runtime id of pivot `pivot_index` on `sheet_index` to
 *        `*out`, with the contract of `fm_sheet_get_table_runtime_id`: the
 *        id survives edits that move the pivot within its sheet's list and a
 *        pivot edit that is rolled back, and a copied sheet's pivots get ids
 *        of their own.
 *
 * @return As `fm_sheet_get_pivot_cache_id`, with `out` in place of
 *         `out_cache_id`.
 */
FM_API fm_status_t fm_sheet_get_pivot_runtime_id(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 uint32_t* out);

/**
 * @brief Writes the current index on `sheet_index` of the pivot whose runtime
 *        id is `id` to `*out_pivot_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_pivot_index` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no pivot on the sheet has `id` (the
 *         pivot was removed, or `id` is `0` or was never handed out).
 */
FM_API fm_status_t fm_sheet_resolve_pivot_runtime_id(const fm_workbook_t* wb, size_t sheet_index, uint32_t id,
                                                     size_t* out_pivot_index);

/**
 * @brief Reads the pivot's options into `out`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) if `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_options(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                              fm_pivot_options_t* out);

/**
 * @brief Merges `patch` into the pivot's options. The whole patch is checked
 *        before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits a field that
 *           cannot be cleared, `SET` carries a `NULL` string, or a `SET` value breaks the
 *           limits `fm_pivot_options_patch_t` states.
 */
FM_API fm_status_t fm_sheet_update_pivot_options(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                 const fm_pivot_options_patch_t* patch);

/**
 * @brief Writes the pivot's report into its sheet cells now, as a pivot-table
 *        mutation does, and clears its pending flag. A pivot over a cache
 *        loaded without its records writes nothing. A report whose range is
 *        occupied or runs past the sheet edge is blocked (Pivot tables, above).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) if `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_project_pivot(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Reads into `*out` whether the pivot's sheet cells are stale after a
 *        change to its cache (`1`) or current (`0`).
 *
 * A blocked report is not pending; whether it is drawn is
 * `fm_sheet_get_pivot_spill`'s answer.
 *
 * The worksheet source is metadata only: records come from
 * `fm_workbook_add_pivot_cache_field` / `fm_workbook_add_pivot_cache_record` /
 * `fm_workbook_set_pivot_cache_record_*`; neither `fm_sheet_project_pivot` nor a
 * cell edit re-reads the range.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) if `pivot_index` is out of range.
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
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `(row, col)` is outside the grid or is not
 *         a data, subtotal or grand total cell of the pivot;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the pivot disables drill-down.
 */
FM_API fm_status_t fm_sheet_show_pivot_details(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index, uint32_t row,
                                               uint32_t col, size_t* out_sheet_index);

/**
 * @brief The row / column grand total bands of a pivot
 *        (`fm_sheet_get_pivot_grand_totals`): each `0` or `1`, the values
 *        `fm_pivot_grand_totals_patch_t` writes.
 */
typedef struct {
  uint32_t struct_size;
  /** Whether the grand total band for the rows shows. */
  int32_t rows;
  /** Whether the grand total band for the columns shows. */
  int32_t columns;
} fm_pivot_grand_totals_t;

/**
 * @brief Reads the row / column grand total bands into `*out`, the values
 *        `fm_sheet_update_pivot_grand_totals` writes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) if `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_grand_totals(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                   fm_pivot_grand_totals_t* out);

/**
 * @brief Reads the pivot's table style, the values
 *        `fm_sheet_update_pivot_style_options` writes. `*out_present` is `1`
 *        and `*out` filled when the definition has a style element;
 *        otherwise `0` and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * Lifetime: transient.
 *
 * @return As `fm_sheet_get_pivot_grand_totals`, with `out_present` / `out` in
 *         place of `out`.
 */
FM_API fm_status_t fm_sheet_get_pivot_style_options(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    int32_t* out_present, fm_pivot_style_options_t* out);

/**
 * @brief Reads the definition of pivot field `field_index` into `out`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) if `pivot_index` or `field_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_field_at(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                               size_t field_index, fm_pivot_field_info_t* out);

/**
 * @brief Number of items in pivot field `field_index`'s item list, the list
 *        `fm_sheet_add_pivot_field_item` / `fm_sheet_add_pivot_field_item_by_cache_index` append to.
 *
 * @return As `fm_sheet_get_pivot_field_at`, with `out_count` in place of `out`.
 */
FM_API fm_status_t fm_sheet_get_pivot_field_item_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       size_t field_index, size_t* out_count);

/**
 * @brief Reads item `item_index` of pivot field `field_index` into `out`. The
 *        index is the one `fm_sheet_update_pivot_field_item` takes.
 *
 * Lifetime: transient.
 *
 * @return As `fm_sheet_get_pivot_field_at`, plus `FM_STATUS_INDEX_OUT_OF_RANGE` (13)
 *         when `item_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_field_item_at(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    size_t field_index, size_t item_index, fm_pivot_item_info_t* out);

/**
 * @brief Reads data field `data_field_index` into the shape
 *        `fm_sheet_add_pivot_data_field` accepts. `name` and
 *        `number_format` are model-backed views, valid until the next
 *        mutation; `name` is never `NULL`, and `number_format` is the format
 *        code as `fm_pivot_field_info_t::number_format` resolves it (`NULL`
 *        for none).
 *        An unset `show_as_base_field` / `show_as_base_item` reads its `has_`
 *        flag `0` and the value `0`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) if `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) if `pivot_index` or `data_field_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_data_field_at(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                    size_t data_field_index, fm_pivot_data_field_t* out);

/**
 * @brief Copies the row-axis field order into the caller's buffer: indices
 *        into the pivot's fields, the list `fm_sheet_set_pivot_row_field_order`
 *        takes. With two or more data fields the axis holding the Values
 *        level lists `FM_PIVOT_VALUES_FIELD` at its position.
 *
 * `*out_count` always receives the number of entries, 0 for an empty axis;
 * at most `capacity` of them are written to `out`. Call with
 * `capacity == 0` (and `out` may then be `NULL`) to size the buffer first.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure (a `NULL`
 *           `out` is one only while `capacity > 0`).
 */
FM_API fm_status_t fm_sheet_get_pivot_row_field_order(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                      uint32_t* out, size_t capacity, size_t* out_count);

/**
 * @brief Copies the column-axis field order into the caller's buffer. Same
 *        contract as rows.
 *
 * @return As `fm_sheet_get_pivot_row_field_order`.
 */
FM_API fm_status_t fm_sheet_get_pivot_column_field_order(const fm_workbook_t* wb, size_t sheet_index,
                                                         size_t pivot_index, uint32_t* out, size_t capacity,
                                                         size_t* out_count);

/**
 * @brief Copies the page-axis (report filter) field order into the caller's
 *        buffer. Same contract as rows; the list never holds
 *        `FM_PIVOT_VALUES_FIELD`.
 *
 * @return As `fm_sheet_get_pivot_row_field_order`.
 */
FM_API fm_status_t fm_sheet_get_pivot_page_field_order(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                       uint32_t* out, size_t capacity, size_t* out_count);

/*
 * Persistence contract shared by the whole pivot-filter family
 * (`fm_sheet_get_pivot_filter_count`, `fm_sheet_get_pivot_filter_at`,
 * `fm_sheet_add_pivot_filter`, `fm_sheet_clear_pivot_filters`,
 * `fm_sheet_remove_pivot_filter`):
 *
 * The list is the pivot's label / value / date filters and is saved as the
 * definition's `<filters>` block. Every entry of a loaded block whose type
 * this ABI enumerates appears here, and clear / remove take it away.
 * Only entries the model cannot express (an unknown type, a member-property
 * or OLAP measure filter, criteria that do not decode) are not listed; they
 * filter nothing and are re-emitted verbatim on save in their original
 * position. A loaded entry is saved byte-for-byte while it still applies to
 * the field and data field it was read with. An entry whose field no longer
 * resolves filters nothing and is not saved.
 */

/**
 * @brief Number of filters on the pivot; see the persistence contract above.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_get_pivot_filter_count(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                   size_t* out_count);

/**
 * @brief Reads the filter at `filter_index` without mutating the pivot.
 *
 * See the persistence contract above. A loaded entry reads back the index
 * of the field it names.
 *
 * `out` reuses the `fm_pivot_filter_t` shape accepted by
 * `fm_sheet_add_pivot_filter`. A text primary value's `value_text` is a
 * model-backed view; it remains valid until the
 * next mutation of this pivot's filter list, a workbook model replacement,
 * or handle destruction. Reads do not invalidate the view. Copy it when
 * it must outlive those events. For non-text primary values,
 * `value_text` is `NULL`; an unset upper bound has kind
 * `FM_PIVOT_FILTER_OPERAND_KIND_ABSENT` and zero numeric payloads.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if the model contains a filter kind
 *         that is not representable by this C ABI. On any error, `*out` is left
 *         unchanged.
 */
FM_API fm_status_t fm_sheet_get_pivot_filter_at(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                size_t filter_index, fm_pivot_filter_t* out);

/**
 * @brief Appends an active filter on field `spec->field_index`, which must
 *        sit on a row, column or page axis of the pivot.
 *        `spec->value_text` must be non-NULL when
 *        `spec->value_kind == FM_PIVOT_FILTER_OPERAND_KIND_TEXT`. The upper-bound
 *        payload is honoured only for the range types (`*_BETWEEN`,
 *        `*_NOT_BETWEEN`).
 *
 * The entry is saved with the pivot; see the persistence contract above.
 * `*out_filter_index` receives its position among the active filters.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure (a `NULL`
 *           `spec->value_text` where it is required is a `NULL` required pointer);
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) if `field_index` is not below the pivot's field count, or
 *           `data_field_index` is out of range for a value filter;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) if `kind` is outside the C ABI enum range, the value
 *           payload discriminators are unusable, or field `field_index` sits on no axis.
 */
FM_API fm_status_t fm_sheet_add_pivot_filter(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                             const fm_pivot_filter_t* spec, size_t* out_filter_index);

/**
 * @brief Drops every listed filter from the pivot. Loaded entries outside the
 *        list are kept; see the persistence contract above.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_clear_pivot_filters(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index);

/**
 * @brief Removes the filter at `filter_index`; see the persistence contract
 *        above.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common pivot failure.
 */
FM_API fm_status_t fm_sheet_remove_pivot_filter(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                                size_t filter_index);

/* -------------------------------------------------------------------------- */
/* Dynamic-array spill payload                                                */
/* -------------------------------------------------------------------------- */

/**
 * @brief What keeps a spill or a PivotTable report from being drawn, the
 *        `cause` of an `fm_spill_info_t`.
 *
 * `OCCUPIED` covers every obstruction other than a merge: a value, a formula,
 * another spill, a table, a PivotTable, a legacy array block or a data table.
 * Of several obstructions the first in row-major order decides, and a range
 * running past the sheet edge is `BEYOND_EDGE` whatever lies inside the grid.
 */
typedef enum {
  FM_SPILL_CAUSE_ABSENT = -1,
  FM_SPILL_CAUSE_OCCUPIED = 0,
  FM_SPILL_CAUSE_MERGED = 1,
  FM_SPILL_CAUSE_BEYOND_EDGE = 2,
} fm_spill_cause_t;

/**
 * @brief Result shape for `fm_sheet_get_spill` and `fm_sheet_get_pivot_spill`.
 *
 * Use `fm_sheet_get_value(wb, sheet_index, rect.first_row + r, rect.first_col + c, &v)`
 * to read individual cells; `fm_sheet_get_value` is already
 * spill-aware and returns the per-cell value verbatim.
 */
typedef struct {
  uint32_t struct_size;
  /** The cell holding the dynamic-array formula, or the PivotTable's anchor cell (below the page-field rows that `rect`
   *  includes). */
  fm_cell_address_t anchor;
  /** The region including its anchor, whose top-left cell holds the dynamic-array formula or is a PivotTable's anchor;
   *  a PivotTable report's region also takes the page-field rows above its anchor. A blocked region running past the
   *  sheet edge is clipped to the grid. */
  fm_cell_rect_t rect;
  /** `1` when the cell is the anchor of a spill refused with `#SPILL!`, or of a blocked PivotTable report; the region
   *  is then the rectangle it attempted. `0` when the cell is part of a registered spill region or is a drawn report's
   *  anchor. Only the anchor reports a blocked spill; cells inside the attempted rectangle read
   *  `*out_present == 0`. */
  int32_t blocked;
  /** What blocks the region, an `fm_spill_cause_t`; `FM_SPILL_CAUSE_ABSENT` when `blocked` is `0` and when nothing is
   *  reported. */
  int32_t cause;
} fm_spill_info_t;

/**
 * @brief Reads the spill region at `(sheet, row, col)`. `*out_present` is `1`
 *        and `*out` filled when the cell is the anchor or a phantom of a
 *        region, a blocked anchor (see `blocked`), or a PivotTable report's
 *        anchor, which reports what `fm_sheet_get_pivot_spill` does; otherwise
 *        `0` and `*out` holds no region (every field zero except `cause`, which
 *        is `FM_SPILL_CAUSE_ABSENT`; still `FM_STATUS_OK` (0)). The other cells of
 *        a drawn report are ordinary cells and read `0`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_spill(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                      int32_t* out_present, fm_spill_info_t* out);

/**
 * @brief Reads the range of a PivotTable report (Pivot tables, above).
 *        `*out_present` is `1` and `*out` filled for a drawn report, with the
 *        range it was last drawn over (also while a cache change has left it
 *        pending) and `blocked` `0`, and for a blocked report, with the range
 *        it needs clipped to the grid, `blocked` `1` and the `cause`. A drawn
 *        report read from a file and not laid out since reports the location
 *        it was read with. Otherwise `0` and `*out` holds no region (every
 *        field zero except `cause`, which is `FM_SPILL_CAUSE_ABSENT`; still
 *        `FM_STATUS_OK` (0)), for a report neither drawn nor read from a file.
 *        A blocked report loaded from a file reports the range the file
 *        recorded, the anchor alone when it records none, until a recalc lays
 *        it out again; one whose cache has no records keeps it, and one whose
 *        cache changed since it was blocked reports the last range it needed.
 *        Pending cells are calculated first, which renders a report whose
 *        range came free.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `pivot_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_pivot_spill(const fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,
                                            int32_t* out_present, fm_spill_info_t* out);

/* -------------------------------------------------------------------------- */
/* Legacy array (CSE) formulas                                                */
/* -------------------------------------------------------------------------- */

/**
 * @brief Result shape for `fm_sheet_get_array_formula`.
 *
 * Every cell of the block reports the same values.
 */
typedef struct {
  uint32_t struct_size;
  /** The block the formula fills (`<f t="array" ref=...>`); its top-left cell holds the formula. */
  fm_cell_rect_t rect;
  /** The anchor's stored formula text, spelled as `fm_sheet_get_formula` returns it, with a leading `=` (optional on
   *  input). Scratch-backed: valid until the next call on `wb` that publishes scratch text. */
  const char* formula;
} fm_array_formula_info_t;

/**
 * @brief Reads the legacy array formula covering `(sheet, row, col)`.
 *        `*out_present` is `1` and `*out` filled when the cell belongs to a
 *        block; otherwise `0` and `*out` zeroed, `formula` `NULL` (still
 *        `FM_STATUS_OK` (0)).
 *
 * A dynamic-array spill is not a legacy array formula; use
 * `fm_sheet_get_spill` for that.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the coordinate
 *         is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_array_formula(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              int32_t* out_present, fm_array_formula_info_t* out);

/**
 * @brief Enters `formula` as one legacy array formula over the rectangle
 *        `range`, whose top-left cell anchors it.
 *
 * Every cell of the rectangle loses its value and formula (styles are
 * kept); a recalc distributes the result, filling `#N/A` where a smaller
 * result cannot reach. A block wholly inside the rectangle is replaced.
 * Writing a value, blank or formula to the anchor of a block replaces the
 * whole block, which is how a block is cleared; a write to any other cell
 * of a block is refused. Structured references are bound at the
 * top-left cell as `fm_sheet_set_formula` binds them.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `wb`, `range` or `formula` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the rectangle
 *         is inverted or leaves the grid;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when it holds more than 10,000,000 cells
 *         (context `part=array_cells requested=<n> limit=10000000`);
 *         a parser code when `formula` does not parse;
 *         `FM_STATUS_PRECONDITION_FAILED` (9), changing nothing, when the rectangle
 *         partially overlaps another array block, intersects a merged
 *         area, a data table or a PivotTable, or (for more than one cell) a
 *         table.
 */
FM_API fm_status_t fm_sheet_set_array_formula(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                              const char* formula);

/* -------------------------------------------------------------------------- */
/* Function catalog metadata                                                  */
/* -------------------------------------------------------------------------- */

/**
 * @brief Coarse availability class for a catalogued function.
 *
 * The value distinguishes "the parser/registry recognizes this function"
 * from "Formulon has a real Excel-compatible implementation." Consumers
 * should surface `FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB` clearly in autocomplete
 * and migration reports instead of treating it as implemented parity.
 *
 * `FM_FUNCTION_AVAILABILITY_IMPLEMENTED_UNVERIFIED` is a reserved ordinal that the
 * engine never returns: every catalogued function is classified by the
 * table in `parts/function_catalog.cpp`, which uses only the other three.
 * A consumer switching over this enum still has to accept it so the
 * switch stays total.
 */
typedef enum {
  FM_FUNCTION_AVAILABILITY_IMPLEMENTED = 0,
  FM_FUNCTION_AVAILABILITY_IMPLEMENTED_UNVERIFIED = 1,
  FM_FUNCTION_AVAILABILITY_ENVIRONMENT_BOUND = 2,
  FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB = 3
} fm_function_availability_t;

/**
 * @brief Result shape for `fm_get_function_metadata`.
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
  uint32_t struct_size;
  /** The canonical function name; always populated when the function is known. */
  const char* canonical_name;
  /** The minimum number of arguments Excel accepts at entry, the bound formula entry enforces. */
  uint32_t min_arity;
  /** 1 when `max_arity` bounds the argument count; 0 for an open argument list. */
  int32_t has_max_arity;
  /** The maximum number of arguments Excel accepts at entry, the bound formula entry enforces; meaningful only when
   *  `has_max_arity` is 1. */
  uint32_t max_arity;
  /** How far the engine implements the function: a real implementation, host-bound, or an intentionally unavailable
   *  fixed-error stub, an `fm_function_availability_t`. */
  int32_t availability;
  /* Always `NULL` - supplied by the host's provider document. */
  /** Always `NULL`: display text is host-injected, not engine-owned. */
  const char* signature_template;
  /* Always `NULL` - supplied by the host's provider document. */
  /** The human-readable description; `NULL` until a host metadata document fills it in. */
  const char* description;
} fm_function_metadata_info_t;

/**
 * @brief Returns metadata for the function `name`.
 *
 * `name` is matched case-insensitively against the canonical name.
 * `*out_present` is `1` when a function matches, and `*out` then receives a
 * view into the catalog's static storage, valid for the lifetime of the
 * process; otherwise `*out_present` is `0` and `*out` zeroed (still
 * `FM_STATUS_OK` (0)). The metadata is locale-invariant.
 *
 * Lifetime: static.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `name`, `out_present` or `out` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is malformed UTF-8.
 */
FM_API fm_status_t fm_get_function_metadata(const char* name, int32_t* out_present, fm_function_metadata_info_t* out);

/**
 * @brief Writes the total number of registered Formulon functions to
 *        `*out_count`.
 *
 * Drives the canonical-name iterator (`fm_get_function_name_at`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out_count` is `NULL`.
 */
FM_API fm_status_t fm_get_function_count(size_t* out_count);

/**
 * @brief Returns the canonical name of the `index`-th registered function.
 *
 * Order is sorted ascending so consumers can build deterministic UI
 * lists. `*out_name` is a static view into the catalog's process-static
 * storage.
 *
 * Lifetime: static.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out_name` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_get_function_name_at(size_t index, const char** out_name);

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
 * Lifetime: static.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `canonical_name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`);
 *         `FM_STATUS_NOT_FOUND` (6) when `canonical_name` does not match a
 *           registered function.
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
 * Lifetime: static.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `localized_name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`);
 *         `FM_STATUS_NOT_FOUND` (6) when no function matches.
 */
FM_API fm_status_t fm_canonicalize_function_name(const char* localized_name, const char* profile_id,
                                                 const char** out_canonical);

/**
 * @brief Rewrites stored (en-invariant) formula text into `profile_id`'s
 *        spelling.
 *
 * Function names, list / array separators, decimal points, booleans and
 * error names take the locale form; sheet qualifiers are requoted by the
 * profile's display rule; strings, external qualifiers and
 * structured-reference brackets are kept verbatim. A number literal with an
 * exponent is regenerated (`1.5E-3` becomes `0.0015`). A leading `=` is
 * preserved. `formula` must parse: a syntax error fails with a parser
 * status, and `fm_get_last_parse_error` places it by UTF-16 offsets into
 * `formula`.
 *
 * `spec == NULL` keeps references in A1. With `spec->r1c1` set to
 * `FM_TRISTATE_TRUE` every reference is written in R1C1 with the profile's
 * row / column letters and brackets (de `Z(-1)S(1)`, es `F[-1]C[1]`):
 * absolute axes as the 1-based index, relative axes as the plain difference
 * from `(base_row, base_col)`, a whole column or row as one axis. A relative
 * reference with `has_base` 0 fails with
 * `FM_STATUS_PARSER_INVALID_REFERENCE` (1004) over that reference.
 *
 * `*out_formula` is a thread-local view, valid until the next
 * `fm_localize_formula` / `fm_canonicalize_formula` call on the same
 * thread. Copy it before making another call when it must be retained.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `formula`, `profile_id` or `out_formula` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `formula` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`), `spec->struct_size` is too small, `r1c1` is outside
 *           `fm_tristate_t`, `has_base` is not 0/1, or `has_base` is 1 with the base outside the grid;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position) when
 *           `formula` does not parse, or when an R1C1 relative reference has no base.
 */
FM_API fm_status_t fm_localize_formula(const char* formula, const char* profile_id,
                                       const fm_reference_notation_spec_t* spec, const char** out_formula);

/**
 * @brief Rewrites `profile_id`-spelled formula text into the en-invariant
 *        form; the inverse of `fm_localize_formula`.
 *
 * `*out_formula` is the same thread-local view as `fm_localize_formula`
 * returns, with the same rewrite scope and the same `spec`: with
 * `spec->r1c1` set to `FM_TRISTATE_TRUE` the profile's R1C1 references are
 * read back into A1, relative axes measured from `(base_row, base_col)` and
 * wrapping at the grid edge. A malformed R1C1 reference, or a relative one
 * with `has_base` 0, fails with `FM_STATUS_PARSER_INVALID_REFERENCE` (1004)
 * over that reference. The rewritten text must parse: a syntax error fails with a
 * parser status, and `fm_get_last_parse_error` places it by UTF-16 offsets
 * into `formula`, the localized text passed in.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `formula`, `profile_id` or `out_formula` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `formula` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`), or `spec` is invalid as for `fm_localize_formula`;
 *         a parser status (1000-1014; `fm_get_last_parse_error` gives the code and position) when
 *           `formula` does not parse.
 */
FM_API fm_status_t fm_canonicalize_formula(const char* formula, const char* profile_id,
                                           const fm_reference_notation_spec_t* spec, const char** out_formula);

/** @brief Order in which a locale writes day, month and year. */
typedef enum { FM_DATE_ORDER_MDY = 0, FM_DATE_ORDER_YMD = 1, FM_DATE_ORDER_DMY = 2 } fm_date_order_t;

/**
 * @brief Scalar locale facts of one profile (`fm_get_locale_facts`).
 *
 * Every string is a NUL-terminated UTF-8 view into process-static storage.
 */
typedef struct {
  uint32_t struct_size;
  /** The decimal separator. */
  const char* decimal_separator;
  /** The digit group (thousands) separator. */
  const char* group_separator;
  /** The separator between function arguments in formulas. */
  const char* list_separator;
  /** The separator between columns in an array constant. */
  const char* array_column_separator;
  /** The separator between rows in an array constant. */
  const char* array_row_separator;
  /** The spelling of the logical value TRUE. */
  const char* true_name;
  /** The spelling of the logical value FALSE. */
  const char* false_name;
  /** The order of year, month and day in short dates, a `LocaleDateOrder` value. */
  int32_t date_order;
  /** The currency symbol. */
  const char* currency_symbol;
  /** `1` when the currency symbol follows the number. */
  int32_t currency_suffix;
  /** `1` when one space separates the currency symbol and the number. */
  int32_t currency_space;
  /** The number of decimal places the currency format shows by default. */
  int32_t currency_default_decimals;
  /** `1` for an Excel-backed profile (every mac profile and win-365-ja_JP); `0` for an estimated one. */
  int32_t measured;
  /** `1` when the text collation of the profile was captured from Excel (every mac profile); `0` for the estimated win
   *  profiles, including win-365-ja_JP, whose `measured` is `1` for the other facts. */
  int32_t collation_measured;
} fm_locale_facts_info_t;

/**
 * @brief Fills `*out` with the locale facts of `profile_id`.
 *
 * Lifetime: static.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `profile_id` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile
 *           (context `profile_id=<id>`).
 */
FM_API fm_status_t fm_get_locale_facts(const char* profile_id, fm_locale_facts_info_t* out);

/**
 * @brief Which Excel text fold a collation call applies.
 *
 * Selects the text fold only; type boundaries (number, text, boolean, error,
 * blank) never depend on it.
 */
typedef enum {
  FM_COLLATION_STRENGTH_COMPARE = 0,      /* the fold of =, criteria, SEARCH and approximate lookup */
  FM_COLLATION_STRENGTH_DYNAMIC = 1,      /* the fold of SORT, UNIQUE, XLOOKUP, GROUPBY and pivots */
  FM_COLLATION_STRENGTH_CLASSIC_EXACT = 2 /* the fold of MATCH 0 and VLOOKUP FALSE */
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
FM_API fm_status_t fm_compare_text(const char* a, const char* b, int32_t strength, const char* profile_id,
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
FM_API fm_status_t fm_get_collation_classes(const fm_value_t* values, size_t count, int32_t strength,
                                            const char* profile_id, uint32_t* out_class_ids, uint32_t* out_class_count);

/**
 * @brief Writes the number of error values, the bound of `index` in
 *        `fm_get_locale_error_name_at`, to `*out_count`.
 *
 * `index` uses the `fm_error_code_t` numbering (`4` is `#NAME?`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out_count` is `NULL`.
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
 * Lifetime: static.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `profile_id` is not a documented profile;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
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
 * `FM_SHEET_VISIBILITY_VERY_HIDDEN` differs from `FM_SHEET_VISIBILITY_HIDDEN` in that Excel
 * leaves such a sheet out of its "Unhide" dialog, which is how a workbook
 * keeps a settings or lookup sheet out of a user's reach.
 */
typedef enum {
  FM_SHEET_VISIBILITY_VISIBLE = 0,
  FM_SHEET_VISIBILITY_HIDDEN = 1,
  FM_SHEET_VISIBILITY_VERY_HIDDEN = 2,
} fm_sheet_visibility_t;

/** @brief `<sheetView view>` (`fm_sheet_view_info_t::view_mode`). */
typedef enum {
  FM_SHEET_VIEW_MODE_NORMAL = 0, /**< The default; the attribute is absent. */
  FM_SHEET_VIEW_MODE_PAGE_BREAK_PREVIEW = 1,
  FM_SHEET_VIEW_MODE_PAGE_LAYOUT = 2,
} fm_sheet_view_mode_t;

/**
 * @brief Per-sheet view state surfaced over the C ABI.
 *
 * Mirrors `formulon::SheetView` in full. Fields that are at their default
 * value still surface.
 */
typedef struct {
  uint32_t struct_size;
  /** The zoom as a percentage, 10 to 400; 100 is 100%. */
  uint32_t zoom_pct;
  /** The number of frozen rows at the top; 0 means no frozen rows. */
  uint32_t freeze_rows;
  /** The number of frozen columns at the left; 0 means no frozen columns. */
  uint32_t freeze_cols;
  /** `1` when the cell gridlines are shown; the default is `1`. */
  int32_t show_grid_lines;
  /** `1` when the row and column headers are shown; the default is `1`. */
  int32_t show_row_column_headers;
  /** `1` when cells holding zero show the 0; the default is `1`. */
  int32_t show_zeros;
  /** `1` when the sheet is laid out right to left; the default is `0`. */
  int32_t right_to_left;
  /** `1` when the sheet's tab is selected: `0` when a loaded file leaves it unset and on an added sheet, `1` on the
   *  first sheet of a new workbook. Plain metadata: it does not by itself make the sheet the active one. */
  int32_t tab_selected;
  /** The view mode, an `fm_sheet_view_mode_t`; the normal view when the attribute is absent. A `view` value outside the
   *  OOXML set reads as the normal view and is kept verbatim on save. */
  int32_t view_mode;
  /** 0-based row of the scroll origin (`topLeftCell`); 0 by default. */
  uint32_t top_left_row;
  /** 0-based column of the scroll origin (`topLeftCell`); 0 by default. */
  uint32_t top_left_col;
  /** 0-based row of the first visible cell of a frozen or split view's lower-right pane (`<pane topLeftCell>`); in a
   *  frozen view the first row past the frozen rows until a host scrolls it. 0 when nothing is frozen or split. */
  uint32_t pane_top_left_row;
  /** 0-based column of the first visible cell of a frozen or split view's lower-right pane (`<pane topLeftCell>`);
   *  in a frozen view the first column past the frozen columns until a host scrolls it. 0 when nothing is frozen or
   *  split. */
  uint32_t pane_top_left_col;
  /** The horizontal position of an unfrozen split in twips; 0 means no split on that axis, and it is always 0 while the
   *  view is frozen. */
  double split_x_twips;
  /** The vertical position of an unfrozen split in twips; 0 means no split on that axis, and it is always 0 while the
   *  view is frozen. */
  double split_y_twips;
  /** `1` when cells show their formula text instead of the value; the default is `0`. */
  int32_t show_formulas;
  /** `1` when the outline gutter's group symbols are shown; the default is `1`. */
  int32_t show_outline_symbols;
} fm_sheet_view_info_t;

/**
 * @brief Partial sheet-view update (`fm_sheet_update_view`), following the
 *        header's patch rules (Patches, above). Fields are those of
 *        `fm_sheet_view_info_t`; no field can be cleared.
 *
 * A freeze replaces a split and a split replaces a freeze, so one patch may
 * not state a non-zero freeze and a non-zero split together; a split of
 * `(0, 0)` removes the split. A view without panes accepts only the default
 * pane origin `(0, 0)`. Changing the panes resets the pane origin and active
 * pane and keeps only the active selection, re-homed to the new active pane; a
 * `pane_top_left_*` in the same patch applies after that reset. Setting the
 * current freeze or split is a no-op.
 */
typedef struct {
  uint32_t struct_size;
  /** Operation on `zoom_pct`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t zoom_pct_op;
  /** The zoom percentage to set, in [10, 400]; a value outside it is refused, not clamped. */
  uint32_t zoom_pct;
  /** Operation on `freeze_rows`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t freeze_rows_op;
  /** The number of frozen rows to set; 0 removes the frozen rows. The first unfrozen cell (`freeze_rows`,
   *  `freeze_cols`) must lie inside the grid: `freeze_rows` is at most 1048575. */
  uint32_t freeze_rows;
  /** Operation on `freeze_cols`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t freeze_cols_op;
  /** The number of frozen columns to set; 0 removes the frozen columns. `freeze_cols` is at most 16383. */
  uint32_t freeze_cols;
  /** Operation on `show_grid_lines`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t show_grid_lines_op;
  /** Non-zero shows the cell gridlines. */
  int32_t show_grid_lines;
  /** Operation on `show_row_column_headers`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t show_row_column_headers_op;
  /** Non-zero shows the row and column headers. */
  int32_t show_row_column_headers;
  /** Operation on `show_zeros`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t show_zeros_op;
  /** Non-zero shows zero values. */
  int32_t show_zeros;
  /** Operation on `right_to_left`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t right_to_left_op;
  /** Non-zero lays the sheet out right to left. */
  int32_t right_to_left;
  /** Operation on `tab_selected`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t tab_selected_op;
  /** Non-zero selects the sheet's tab; plain metadata that does not by itself make the sheet active; setting it on
   *  more than one sheet affects only what a host reproduces as the active tab. */
  int32_t tab_selected;
  /** Operation on `view_mode`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t view_mode_op;
  /** The view mode to set, an `fm_sheet_view_mode_t`; the normal view removes the `view` attribute. */
  int32_t view_mode;
  /** Operation on `top_left_row`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t top_left_row_op;
  /** 0-based row of the scroll origin to set; the cell must lie inside the grid. */
  uint32_t top_left_row;
  /** Operation on `top_left_col`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t top_left_col_op;
  /** 0-based column of the scroll origin to set; the cell must lie inside the grid. */
  uint32_t top_left_col;
  /** Operation on `pane_top_left_row`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t pane_top_left_row_op;
  /** 0-based row of the pane's first visible cell; the view must be frozen or split, and in a frozen view the row is at
   *  least `freeze_rows`. */
  uint32_t pane_top_left_row;
  /** Operation on `pane_top_left_col`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t pane_top_left_col_op;
  /** 0-based column of the pane's first visible cell; the view must be frozen or split, and in a frozen view the column
   *  is at least `freeze_cols`. */
  uint32_t pane_top_left_col;
  /** Operation on `split_x_twips`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t split_x_twips_op;
  /** The horizontal position in twips of an unfrozen split; finite and non-negative, 0 for no split on that axis. */
  double split_x_twips;
  /** Operation on `split_y_twips`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t split_y_twips_op;
  /** The vertical position in twips of an unfrozen split; finite and non-negative, 0 for no split on that axis. */
  double split_y_twips;
  /** Operation on `show_formulas`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t show_formulas_op;
  /** Non-zero shows the formula text instead of the value. */
  int32_t show_formulas;
  /** Operation on `show_outline_symbols`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t show_outline_symbols_op;
  /** Non-zero shows the outline gutter's group symbols. */
  int32_t show_outline_symbols;
} fm_sheet_view_patch_t;

/**
 * @brief Per-column layout override surfaced over the C ABI.
 *
 * The OOXML `min` / `max` 1-based conversion stays inside the writer.
 */
typedef struct {
  uint32_t struct_size;
  /** 0-based index of the first column of the span. */
  uint32_t first;
  /** 0-based index of the last column of the span, inclusive. */
  uint32_t last;
  /** The explicit width in characters of the Normal font; meaningful only when `has_width_chars` is 1. */
  double width_chars;
  /** `1` when the columns of the span are hidden. */
  int32_t hidden;
  /** The outline level of the span; 0 means not grouped. */
  uint8_t outline_level;
  /** Padding; always 0. */
  uint8_t _pad[3];
  /** `1` when the outline group this span summarises is collapsed (the `collapsed` attribute). */
  int32_t collapsed;
  /* Attribute-presence flags. has_width_chars distinguishes width="0" from an
   * omitted width and is normalized true for legacy non-zero widths;
   * has_xf_index keeps an explicit style="0". */
  /** 1 when `width_chars` is set; 0 means the span states no width. True for the raw presence bit or a legacy non-zero
   *  width; an explicit zero width is preserved by the raw bit. */
  int32_t has_width_chars;
  /** 1 when the span states a column style (including an explicit style 0); 0 means it states none. */
  int32_t has_xf_index;
  /** The cell XF index the columns of the span apply; meaningful only when `has_xf_index` is 1. */
  uint32_t xf_index;
} fm_column_layout_info_t;

/**
 * @brief Per-row layout override surfaced over the C ABI.
 *
 * Mirrors `formulon::RowLayout`.
 */
typedef struct {
  uint32_t struct_size;
  /** 0-based index of the row. */
  uint32_t row;
  /** The stored height in points; meaningful only when `has_height_pt` is 1. */
  double height_pt;
  /** `1` when the row is hidden. */
  int32_t hidden;
  /** The outline level of the row; 0 means not grouped. */
  uint8_t outline_level;
  /** Padding; always 0. */
  uint8_t _pad[3];
  /** `1` when the outline group this row summarises is collapsed (the `collapsed` attribute). */
  int32_t collapsed;
  /* Row s= is effective only when customFormat=1 in OOXML. */
  /** 1 when the row's style applies (`customFormat="1"`); 0 means the row has no style of its own.
   *  After a save that writes this row as a row format of XF 0, this reports no format until the workbook is
   *  reloaded, although the saved file records XF 0. */
  int32_t has_xf_index;
  /** The cell XF index the row applies; meaningful only when `has_xf_index` is 1. */
  uint32_t xf_index;
  /** 1 when the row stores a height; 0 means none is stored. */
  int32_t has_height_pt;
  /** `1` when the height is an explicit override (`customHeight="1"`) rather than an auto height that Excel recomputes
   *  on its own; the state `fm_sheet_set_row_height` produces. */
  int32_t custom_height;
} fm_row_layout_info_t;

/**
 * @brief Writes the number of column layouts (`fm_column_layout_info_t` spans) on `sheet_index` to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_column_layout_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `index`-th column layout.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_column_layout_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                 fm_column_layout_info_t* out);

/**
 * @brief Writes the number of row layouts (`fm_row_layout_info_t` rows) on `sheet_index` to `*out_count`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_layout_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads the `index`-th row layout.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_layout_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                              fm_row_layout_info_t* out);

/**
 * @brief Reads the sheet-view state (zoom, frozen panes, and the display /
 *        orientation flags). The tab's visibility is
 *        `fm_sheet_get_visibility`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_info_t* out);

/**
 * @brief Merges `patch` into the sheet view (zoom, panes, scroll and pane
 *        origins, display flags, view mode); the field rules are on
 *        `fm_sheet_view_patch_t`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, any op is `CLEAR`, a `SET`
 *           value is outside the domain its field states, a non-zero freeze and a non-zero
 *           split are set together, or the pane origin does not fit the view's panes.
 *         Nothing is written on failure.
 */
FM_API fm_status_t fm_sheet_update_view(fm_workbook_t* wb, size_t sheet_index, const fm_sheet_view_patch_t* patch);

/**
 * @brief Reads the sheet tab's visibility (`<sheet state>`) as an
 *        `fm_sheet_visibility_t`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_visibility` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_visibility(const fm_workbook_t* wb, size_t sheet_index, int32_t* out_visibility);

/**
 * @brief Sets the sheet tab's visibility to `visibility`, an
 *        `fm_sheet_visibility_t` stating the whole tab state: `VISIBLE` shows
 *        the sheet from either hidden state, `HIDDEN` demotes a very-hidden
 *        one. Hiding the active sheet activates its visible successor
 *        (`fm_workbook_get_active_sheet`). Setting the current state is a
 *        no-op.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `visibility` is not an `fm_sheet_visibility_t` member;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when it hides the only visible sheet.
 */
FM_API fm_status_t fm_sheet_set_visibility(fm_workbook_t* wb, size_t sheet_index, int32_t visibility);

/**
 * @brief Reads the active pane's selection: the active cell and its ranges.
 *
 * `*out_count` receives the number of selected ranges; at most `capacity` of
 * them are written to `out`, which may be `NULL` when `capacity` is
 * `0` (size the buffer first). The selection reads as it is saved: a pane
 * that stores none reports `A1` and one range, `A1` (the pane's top-left
 * cell for a frozen or split view's active pane); a selection without an
 * explicit `sqref` is the active cell alone; an active cell outside every
 * range moves to the first range's top-left. Whole-row and whole-column
 * references read as full-width / full-height ranges.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_active_row`, `out_active_col`
 *         or `out_count` is `NULL`, or `out` is `NULL` while
 *         `capacity > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_selection(const fm_workbook_t* wb, size_t sheet_index, uint32_t* out_active_row,
                                          uint32_t* out_active_col, fm_cell_rect_t* out, size_t capacity,
                                          size_t* out_count);

/**
 * @brief Replaces the active pane's selection (`<selection>`): the active
 *        cell and the selected ranges (`sqref`). `range_count == 0` selects the
 *        active cell alone.
 *
 * The selection of a frozen or split view belongs to the pane it is active
 * in (`bottomRight`, `topRight` or `bottomLeft`, see `fm_sheet_update_view`);
 * other panes' selections are kept. An active cell outside every range is
 * saved as the first range's top-left, as Excel does.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, or `ranges` is `NULL` with
 *           a non-zero `range_count`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the active cell or a range is out of range, or a
 *         range has `first > last`.
 */
FM_API fm_status_t fm_sheet_set_selection(fm_workbook_t* wb, size_t sheet_index, uint32_t active_row,
                                          uint32_t active_col, const fm_cell_rect_t* ranges, size_t range_count);

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
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_xml` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
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

/** @brief Dynamic-filter type (`fm_filter_column_t::dynamic_type`); `ST_DynamicFilterType` in schema order. */
typedef enum {
  FM_DYNAMIC_FILTER_TYPE_NULL = 0,
  FM_DYNAMIC_FILTER_TYPE_ABOVE_AVERAGE = 1,
  FM_DYNAMIC_FILTER_TYPE_BELOW_AVERAGE = 2,
  FM_DYNAMIC_FILTER_TYPE_TOMORROW = 3,
  FM_DYNAMIC_FILTER_TYPE_TODAY = 4,
  FM_DYNAMIC_FILTER_TYPE_YESTERDAY = 5,
  FM_DYNAMIC_FILTER_TYPE_NEXT_WEEK = 6,
  FM_DYNAMIC_FILTER_TYPE_THIS_WEEK = 7,
  FM_DYNAMIC_FILTER_TYPE_LAST_WEEK = 8,
  FM_DYNAMIC_FILTER_TYPE_NEXT_MONTH = 9,
  FM_DYNAMIC_FILTER_TYPE_THIS_MONTH = 10,
  FM_DYNAMIC_FILTER_TYPE_LAST_MONTH = 11,
  FM_DYNAMIC_FILTER_TYPE_NEXT_QUARTER = 12,
  FM_DYNAMIC_FILTER_TYPE_THIS_QUARTER = 13,
  FM_DYNAMIC_FILTER_TYPE_LAST_QUARTER = 14,
  FM_DYNAMIC_FILTER_TYPE_NEXT_YEAR = 15,
  FM_DYNAMIC_FILTER_TYPE_THIS_YEAR = 16,
  FM_DYNAMIC_FILTER_TYPE_LAST_YEAR = 17,
  FM_DYNAMIC_FILTER_TYPE_YEAR_TO_DATE = 18,
  FM_DYNAMIC_FILTER_TYPE_Q1 = 19,
  FM_DYNAMIC_FILTER_TYPE_Q2 = 20,
  FM_DYNAMIC_FILTER_TYPE_Q3 = 21,
  FM_DYNAMIC_FILTER_TYPE_Q4 = 22,
  FM_DYNAMIC_FILTER_TYPE_M1 = 23,
  FM_DYNAMIC_FILTER_TYPE_M2 = 24,
  FM_DYNAMIC_FILTER_TYPE_M3 = 25,
  FM_DYNAMIC_FILTER_TYPE_M4 = 26,
  FM_DYNAMIC_FILTER_TYPE_M5 = 27,
  FM_DYNAMIC_FILTER_TYPE_M6 = 28,
  FM_DYNAMIC_FILTER_TYPE_M7 = 29,
  FM_DYNAMIC_FILTER_TYPE_M8 = 30,
  FM_DYNAMIC_FILTER_TYPE_M9 = 31,
  FM_DYNAMIC_FILTER_TYPE_M10 = 32,
  FM_DYNAMIC_FILTER_TYPE_M11 = 33,
  FM_DYNAMIC_FILTER_TYPE_M12 = 34,
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

/** @brief One `<dateGroupItem>` of a values filter. */
typedef struct {
  /** Calendar year. */
  uint16_t year;
  /** Month of the year, 1-12; 0 when the item is coarser than month. */
  uint8_t month;
  /** Day of the month, 1-31; 0 when the item is coarser than day. */
  uint8_t day;
  /** Hour of the day, 0-23; 0 when the item is coarser than hour. */
  uint8_t hour;
  /** Minute of the hour, 0-59; 0 when the item is coarser than minute. */
  uint8_t minute;
  /** Second of the minute, 0-59; 0 when the item is coarser than second. */
  uint8_t second;
  /** The granularity the item matches at; fields finer than it are zero. */
  int32_t grouping;
} fm_date_group_item_t;

/** @brief One custom condition of an AutoFilter column (`<customFilter>`). */
typedef struct {
  /** The comparison applied to the cell, an `fm_comparison_operator_t` other than between and not-between. */
  int32_t op;
  /** The text the cell is compared with (`customFilter val`); `NULL` means empty on input. */
  const char* value;
} fm_filter_condition_t;

/**
 * @brief One `<filterColumn>` of an AutoFilter.
 *
 * `kind` selects the criterion, and only that criterion's fields are read on
 * input; on output the fields of every other criterion are zero. On input
 * every `const char*` may be `NULL`, meaning empty, and an array pointer may
 * be `NULL` only with a zero count.
 */
typedef struct {
  /** The column's 0-based offset from the filter range's first column. */
  uint32_t column_offset;
  /** 1 when the column's filter button is hidden. */
  int32_t hidden_button;
  /** 1 when the column's filter button is shown. */
  int32_t show_button;
  /** Which criterion the column filters on, an `fm_filter_kind_t`; only that criterion's fields are meaningful. */
  int32_t kind;
  /** 1 when the values filter also matches blank cells. */
  int32_t filter_blank;
  /** The displayed texts the values filter keeps. */
  const char* const* values;
  /** The number of entries in `values`. */
  uint32_t value_count;
  /** The date-group items the values filter keeps. */
  const fm_date_group_item_t* date_groups;
  /** The number of entries in `date_groups`. */
  uint32_t date_group_count;
  /** 1 when the two custom conditions are joined by AND, 0 when by OR. */
  int32_t custom_and;
  /** The custom conditions; empty unless the kind is custom. The read array shares the lifetime of the AutoFilter
   *  getter. The setter rejects a between or not-between operator with `FM_STATUS_INVALID_ARGUMENT` (2). */
  const fm_filter_condition_t* conditions;
  /** The number of entries in `conditions`: 1 or 2 for a custom kind, else 0. The setter rejects any other count for a
   *  custom kind with `FM_STATUS_INVALID_ARGUMENT` (2) and ignores it for other kinds. */
  uint32_t condition_count;
  /** 1 when the top-10 filter keeps the top items, 0 for the bottom ones. */
  int32_t top;
  /** 1 when `top_val` is a percentage rather than an item count. */
  int32_t percent;
  /** 1 when `filter_val` is set; 0 means Excel stored no top-10 threshold. */
  int32_t has_filter_val;
  /** The top-10 filter's rank: how many items, or what percentage. */
  double top_val;
  /** The cell value Excel last stored as the top-10 threshold. */
  double filter_val;
  /** Which dynamic criterion the column filters on, such as above average or this month, an
   *  `fm_dynamic_filter_type_t`. */
  int32_t dynamic_type;
  /** 1 when `dyn_val` is set; 0 means Excel stored no lower bound. */
  int32_t has_dyn_val;
  /** 1 when `dyn_max_val` is set; 0 means Excel stored no upper bound. */
  int32_t has_dyn_max_val;
  /** The lower bound Excel last stored for the dynamic filter. */
  double dyn_val;
  /** The upper bound Excel last stored for the dynamic filter. */
  double dyn_max_val;
  /** The ISO date text of the dynamic filter's lower bound, kept as written; `NULL` when none. */
  const char* val_iso;
  /** The ISO date text of the dynamic filter's upper bound, kept as written; `NULL` when none. */
  const char* max_val_iso;
  /** 1 when `dxf_index` is set; 0 means the color filter names no differential format. */
  int32_t has_dxf_index;
  /** The index of the differential format whose fill or font colour the color filter matches. A stored index past the
   *  workbook's dxf count reads as absent. */
  uint32_t dxf_index;
  /** 1 when the color filter matches the cell fill, 0 when it matches the font colour. */
  int32_t cell_color;
  /** The icon set the icon filter's icon comes from, an `fm_icon_set_name_t`. */
  int32_t icon_set;
  /** The index of the icon the icon filter keeps. */
  int32_t icon_index;
  /** 1 when `icon_index` is set; 0 means the icon filter names no icon. */
  int32_t has_icon_index;
} fm_filter_column_t;

/** @brief One `<sortCondition>` of an AutoFilter's sort state. */
typedef struct {
  /** The cells the condition sorts by, the key column of the sorted range. */
  fm_cell_rect_t ref;
  /** 1 sorts in descending order. */
  int32_t descending;
  /** What the condition sorts by: the value, the cell colour, the font colour or the icon, an `fm_sort_by_t`. */
  int32_t sort_by;
  /** The custom sort list, as comma-separated items; `NULL` when the condition uses none. */
  const char* custom_list;
  /** The index of the differential format whose colour the condition sorts by. A stored index past the workbook's dxf
   *  count reads as absent. */
  uint32_t dxf_index;
  /** 1 when `dxf_index` is set; 0 means the condition names no differential format. */
  int32_t has_dxf_index;
  /** The icon set the condition's icon comes from, an `fm_icon_set_name_t`. */
  int32_t icon_set;
  /** The index of the icon the condition sorts by. */
  int32_t icon_index;
  /** 1 when `icon_index` is set; 0 means the condition names no icon. */
  int32_t has_icon_index;
} fm_sort_condition_t;

/**
 * @brief A sheet or table AutoFilter.
 *
 * Attributes and extensions the record does not carry (`xr:uid`, unknown
 * `filterColumn` children, `extLst` payloads, a values filter's
 * `calendarType`) are kept across a set: they are carried over from the
 * AutoFilter being replaced, per column by `column_offset`, so a get followed by a
 * set of the unchanged record leaves the stored element as it was.
 */
typedef struct {
  uint32_t struct_size;
  /** The filtered range, header row included. */
  fm_cell_rect_t range;
  /** The filter columns, in strictly ascending column-offset order. */
  const fm_filter_column_t* columns;
  /** The number of entries in `columns`. */
  uint32_t column_count;
  /** 1 when the AutoFilter has a sort state, described by the sort fields; 0 means it has none. */
  int32_t has_sort;
  /** The cells the sort state applies to. */
  fm_cell_rect_t sort_ref;
  /** 1 when the sort is by columns (left to right) rather than by rows. */
  int32_t column_sort;
  /** 1 when text comparison in the sort is case-sensitive. */
  int32_t case_sensitive;
  /** How East Asian text is ordered: by default order, by pinyin or by stroke count, an `fm_sort_method_t`. */
  int32_t sort_method;
  /** The sort conditions, in priority order. */
  const fm_sort_condition_t* conditions;
  /** The number of entries in `conditions`. */
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
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when the AutoFilter did not parse into
 *         the typed model (only `fm_sheet_get_auto_filter_xml` reads it).
 */
FM_API fm_status_t fm_sheet_get_auto_filter(const fm_workbook_t* wb, size_t sheet_index, int32_t* out_present,
                                            fm_auto_filter_t* out);

/**
 * @brief Replaces the sheet AutoFilter with `filter`, or removes it when
 *        `filter` is `NULL`, and maintains the sheet's hidden
 *        `_xlnm._FilterDatabase` name over its range.
 *
 * Row visibility is not changed; call `fm_sheet_apply_auto_filter` to hide
 * the rows the criteria exclude. Removing drops the `_FilterDatabase` name
 * too and is a no-op when the sheet has no AutoFilter.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, or an array
 *         pointer in `filter` is `NULL` with a non-zero count;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an
 *         enumeration field is outside the domain listed on the record (a custom
 *         condition's `FM_COMPARISON_OPERATOR_BETWEEN` / `NOT_BETWEEN` included);
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when the range is outside the grid,
 *         column offsets are outside the range or not ascending, a custom filter
 *         has other than one or two conditions, or a date group item does
 *         not fit its grouping;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when a present color filter or sort condition `dxf_index` is not below
 *         the workbook's dxf count (`fm_workbook_get_dxf_count`); nothing changes.
 */
FM_API fm_status_t fm_sheet_set_auto_filter(fm_workbook_t* wb, size_t sheet_index, const fm_auto_filter_t* filter);

/**
 * @brief Evaluates the sheet AutoFilter and sets every body row's hidden flag
 *        to the negation of its match, as Excel's Reapply does; a matching
 *        row hidden by hand is shown again, and rows outside the body are
 *        untouched.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when the sheet has no AutoFilter;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when it did not parse into the model.
 */
FM_API fm_status_t fm_sheet_apply_auto_filter(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Drops every column criterion of the sheet AutoFilter and shows all
 *        of its body rows. The range, the drop-down buttons and the sort
 *        state stay; the filter itself is removed with
 *        `fm_sheet_set_auto_filter(wb, sheet_index, NULL)`.
 *
 * A filter without criteria is left as it is (no-op).
 *
 * @return as `fm_sheet_apply_auto_filter`.
 */
FM_API fm_status_t fm_sheet_clear_auto_filter_criteria(fm_workbook_t* wb, size_t sheet_index);

/**
 * @brief Evaluates the sheet AutoFilter without changing the sheet.
 *
 * Body row `*out_first_row + i` matches the criteria when `out[i]` is
 * 1. `*out_count` always receives the number of body rows; at most `capacity`
 * flags are written. Call with `capacity == 0` (`out` may then be `NULL`)
 * to size the buffer first.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_count` or `out_first_row` is
 *         `NULL`, or `out` is `NULL` while `capacity > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when the sheet has no AutoFilter;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when it did not parse into the model.
 */
FM_API fm_status_t fm_sheet_evaluate_auto_filter(const fm_workbook_t* wb, size_t sheet_index, uint8_t* out,
                                                 size_t capacity, size_t* out_count, uint32_t* out_first_row);

/*
 * Table counterparts of the sheet AutoFilter entry points. A table is
 * addressed as `(sheet_index, table_index)`, as `fm_sheet_get_table_at` does,
 * and its sheet is the one whose rows are evaluated and hidden. Tables carry
 * no `_FilterDatabase` name. Each entry also returns the common table
 * failures: `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) for an out-of-range
 * sheet and `FM_STATUS_INDEX_OUT_OF_RANGE` (13) for an out-of-range table.
 */

/**
 * @brief Reads the AutoFilter of table `table_index`, as
 *        `fm_sheet_get_auto_filter` reads a sheet's: `*out_present` is 0 and
 *        `*out` all-zero when the table has none.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when the AutoFilter did not parse into
 *         the typed model.
 */
FM_API fm_status_t fm_sheet_get_table_auto_filter(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                  int32_t* out_present, fm_auto_filter_t* out);

/**
 * @brief Replaces the AutoFilter of table `table_index`, as
 *        `fm_sheet_set_auto_filter` replaces a sheet's, or removes it when
 *        `filter` is `NULL` (row visibility unchanged; a no-op when the table
 *        has none).
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2), `FM_STATUS_AUTO_FILTER_INVALID` (8006) and
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) as for `fm_sheet_set_auto_filter`.
 */
FM_API fm_status_t fm_sheet_set_table_auto_filter(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                  const fm_auto_filter_t* filter);

/**
 * @brief Evaluates the AutoFilter of table `table_index` and sets its body
 *        rows' hidden flags, as `fm_sheet_apply_auto_filter` does for a sheet.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_NOT_FOUND` (6) when the table has no AutoFilter;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when it did not parse into the model.
 */
FM_API fm_status_t fm_sheet_apply_table_auto_filter(fm_workbook_t* wb, size_t sheet_index, size_t table_index);

/**
 * @brief Drops every column criterion of the AutoFilter of table
 *        `table_index` and shows all of its body rows, as
 *        `fm_sheet_clear_auto_filter_criteria` does for a sheet (a no-op on a
 *        filter without criteria).
 *
 * @return as `fm_sheet_apply_table_auto_filter`.
 */
FM_API fm_status_t fm_sheet_clear_table_auto_filter_criteria(fm_workbook_t* wb, size_t sheet_index, size_t table_index);

/**
 * @brief Evaluates the AutoFilter of table `table_index` without changing the
 *        sheet; the buffer contract is `fm_sheet_evaluate_auto_filter`'s.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common table failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `out_count` or `out_first_row` is
 *         `NULL`, or `out` is `NULL` while `capacity > 0`;
 *         `FM_STATUS_NOT_FOUND` (6) when the table has no AutoFilter;
 *         `FM_STATUS_AUTO_FILTER_INVALID` (8006) when it did not parse into the model.
 */
FM_API fm_status_t fm_sheet_evaluate_table_auto_filter(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                       uint8_t* out, size_t capacity, size_t* out_count,
                                                       uint32_t* out_first_row);

/**
 * @brief Sets or replaces the column width override for the inclusive
 *        column span `[first, last]`. Existing overrides whose span
 *        intersects the requested span are merged so the new width
 *        takes precedence on overlapping columns; non-overlapping
 *        portions of pre-existing entries are retained.
 *
 * `width_chars` is a stored column width: Excel's character-width unit with the
 * column padding included, as `<col width>` holds it (Calibri 11 stores the
 * 8.43-character column as 9.140625), at most 255, Excel's widest column.
 * A rejected call leaves the layout exactly as it was.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, `last` is past the last column (16383), or
 *         `width_chars` is negative, above 255, NaN, or infinite.
 */
FM_API fm_status_t fm_sheet_set_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                             double width_chars);

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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last column (16383).
 */
FM_API fm_status_t fm_sheet_reset_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/**
 * @brief Sets or replaces the column hidden flag for the inclusive
 *        column span `[first, last]`. See `fm_sheet_set_column_width`
 *        for the merge semantics.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, `last` is past the last column (16383), or
 *         `level` is above 7, the deepest outline level Excel stores.
 */
FM_API fm_status_t fm_sheet_set_column_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                               uint32_t level);

/**
 * @brief Sets or clears the `collapsed` flag (the outline group summarised by
 *        the span is collapsed) for the inclusive column span `[first, last]`.
 *        See `fm_sheet_set_column_width` for the merge semantics.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last column (16383);
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `xf_index` is not below
 *           `fm_workbook_get_cell_xf_count`.
 */
FM_API fm_status_t fm_sheet_set_column_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                uint32_t xf_index);

/**
 * @brief Resets the inclusive column span `[first, last]` to the default
 *        XF, removing the column style.
 *
 * Width, hidden and outline attributes are retained. Empty spans left with no
 * observable state are removed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last column (16383).
 */
FM_API fm_status_t fm_sheet_reset_column_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/**
 * @brief Sets or replaces the row height override for every row of the
 *        inclusive span `[first, last]`.
 *
 * `height_pt` is in points, at most 409 (Excel's tallest row), and is stored
 * as a manual height (`customHeight="1"`), which Excel keeps when the row's
 * content or font changes. Every other row attribute is kept. A rejected call
 * leaves the layout exactly as it was.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, `last` is past the last row (1048575), or
 *         `height_pt` is negative, above 409, NaN, or infinite.
 */
FM_API fm_status_t fm_sheet_set_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                           double height_pt);

/**
 * @brief Stores `height_pt` as the cached value of the automatic height of every
 *        row of the inclusive span `[first, last]`.
 *
 * Same contract as `fm_sheet_set_row_height`, except the rows stay
 * automatic (`custom_height == 0`): the height is saved as `ht` without
 * `customHeight`, so each row still follows its content and font. Row
 * style, hidden state, outline level and collapsed flag are kept. Use it
 * to refresh a derived height; `fm_sheet_reset_row_height` drops it.
 *
 * @return As `fm_sheet_set_row_height`.
 */
FM_API fm_status_t fm_sheet_set_row_auto_height(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                double height_pt);

/**
 * @brief Sets or replaces the row hidden flag for the inclusive row span
 *        `[first, last]`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last row (1048575).
 */
FM_API fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                           int32_t hidden);

/**
 * @brief Sets or replaces the row outline level for the inclusive row span
 *        `[first, last]`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, `last` is past the last row (1048575), or
 *         `level` is above 7, the deepest outline level Excel stores.
 */
FM_API fm_status_t fm_sheet_set_row_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                            uint32_t level);

/**
 * @brief Sets or clears the `collapsed` flag (the outline group summarised by
 *        the row is collapsed) for the inclusive row span `[first, last]`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last row (1048575).
 */
FM_API fm_status_t fm_sheet_set_row_collapsed(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                              int32_t collapsed);

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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last row (1048575);
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `xf_index` is not below
 *           `fm_workbook_get_cell_xf_count`.
 */
FM_API fm_status_t fm_sheet_set_row_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                             uint32_t xf_index);

/**
 * @brief Resets the inclusive row span `[first, last]` to the default XF,
 *        removing the row style.
 *
 * Existing height, custom-height, hidden and outline attributes are retained;
 * an otherwise empty row override is removed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last row (1048575).
 */
FM_API fm_status_t fm_sheet_reset_row_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/*
 * Attribute-only formatting of a selection. The `range` rectangle is a
 * selection: spanning every row makes it a column selection, every
 * column a row selection, both the whole sheet. To change some attributes
 * (say the font size) and keep the rest, read the XFs the selection resolves
 * to with `fm_sheet_get_range_xf_bases`, build each one's patched XF with
 * `fm_workbook_add_font` / `fm_workbook_add_cell_xf`, and apply the pairs with
 * `fm_sheet_remap_range_xfs`. Every cell then resolves to the patched form
 * of the XF it resolved to before, as Excel applies such a change: covered
 * column and row styles are rewritten, existing cells are remapped, and a
 * blank cell gets its own format only where the row-over-column precedence
 * would otherwise resolve to something else (for a plain range, every cell).
 * A selection that spans whole rows or whole columns is written as row or
 * column formats and adds no cells. A blank cell that ends at the default
 * format keeps a record only where that overrides a row or column format. Only
 * style state is touched. `fm_sheet_snapshot_range_xfs` captures what such an
 * edit can change, for undo.
 *
 * A rectangle that is not a whole row or column selection and covers more
 * than 134,217,728 cells is refused with `FM_STATUS_SECURITY_RESOURCE_LIMIT`:
 * a save writes one cell per address. A rectangle within the bound costs
 * memory in proportion to the rows it covers, not its cells. Results that are
 * sized by the area (`fm_sheet_get_cell_xf_indices`, enumeration by
 * `fm_sheet_get_cell_at`) can be as large as the rectangle; read them in
 * ranges.
 */

/**
 * @brief Lists every cell-XF index some cell of the selection resolves to,
 *        ascending: the XFs of cells that carry their own, then the row,
 *        column or default (0) style of the rest.
 *
 * `*out_count` always receives the number of indices; at most `capacity` are
 * written. Call with `capacity == 0` (`out` may then be `NULL`) to size the
 * buffer first.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `range` or `out_count` is `NULL`, or
 *         `out` is `NULL` while `capacity > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the rect is out of range.
 */
FM_API fm_status_t fm_sheet_get_range_xf_bases(const fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                               uint32_t* out, size_t capacity, size_t* out_count);

/**
 * @brief `fm_sheet_get_range_xf_bases` over the selection less the
 *        `excluded_count` cells of `excluded`.
 *
 * An excluded cell contributes no XF: neither its own nor the row, column or
 * default style beneath it. Use it when the host resolves those cells itself
 * (say, a pending edit overrides their style). Excluded cells outside the
 * selection and repeats are ignored; excluding every cell yields no indices.
 * Values, styles and history are not touched.
 *
 * @return As `fm_sheet_get_range_xf_bases`; also `FM_STATUS_BINDING_NULL_POINTER` (7001) when
 *         `excluded` is `NULL` while `excluded_count > 0`.
 */
FM_API fm_status_t fm_sheet_get_range_xf_bases_excluding(const fm_workbook_t* wb, size_t sheet_index,
                                                         const fm_cell_rect_t* range, const fm_cell_address_t* excluded,
                                                         size_t excluded_count, uint32_t* out, size_t capacity,
                                                         size_t* out_count);

/** @brief One cell and its own non-default cell-XF index, as
 *         `fm_sheet_get_cell_xf_indices` writes it.
 *
 * Frozen value type: no `struct_size`. */
typedef struct {
  /** The 0-based row. */
  uint32_t row;
  /** The 0-based column. */
  uint32_t col;
  /** The cell's own cell-XF index; never 0. */
  uint32_t xf_index;
} fm_xf_cell_t;

/**
 * @brief Lists every cell of `range` whose own cell-XF index is not 0, in
 *        row-major order (row, then column, ascending).
 *
 * Styled blank cells and cells with content are both listed; a cell holding
 * the default XF 0 is not, even when it was assigned explicitly, and a spill
 * cell with no stored record is not either. Blank cells that only inherit a
 * row or column format are not listed. Format-only cells outside the used
 * range are included, so pass the whole grid to list every such cell of the
 * sheet. `*out_count` always receives the number of cells; at most `capacity` are
 * written. Call with `capacity == 0` (`out` may then be `NULL`) to size the
 * buffer first. The result is as large as the number of formatted cells, which
 * a large rectangle makes area-sized; there is no cell ceiling on this call, so
 * read a large sheet in ranges.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `range` or `out_count` is `NULL`, or
 *         `out` is `NULL` while `capacity > 0` (`*out_count` is then `0` when `out_count` is non-`NULL`);
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the rect is out of range.
 */
FM_API fm_status_t fm_sheet_get_cell_xf_indices(const fm_workbook_t* wb, size_t sheet_index,
                                                const fm_cell_rect_t* range, fm_xf_cell_t* out, size_t capacity,
                                                size_t* out_count);

/** @brief One pair of `fm_sheet_remap_range_xfs`: a cell resolving to
 *         `from_xf` ends at `to_xf`. */
typedef struct {
  /** The cell-XF index a cell resolves to before the remap. */
  uint32_t from_xf;
  /** The cell-XF index the cell resolves to after the remap. */
  uint32_t to_xf;
} fm_xf_remap_t;

/**
 * @brief Remaps the XFs of a selection by the `count` pairs of `pairs`; an
 *        unlisted XF is kept.
 *
 * The XF indices are validated before the area. Rows, columns, blank cells
 * and the covered cells are handled as the block comment above describes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `range` is `NULL`, or `pairs`
 *         is `NULL` while `count > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the rect is out of range or a
 *         `from_xf` repeats;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when a `to_xf` is not below
 *         `fm_workbook_get_cell_xf_count`;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the rectangle covers more than 134,217,728
 *         cells and is not a whole row or column selection (the engine's own ceiling, not an Excel
 *         limit). Nothing changes on failure.
 */
FM_API fm_status_t fm_sheet_remap_range_xfs(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                            const fm_xf_remap_t* pairs, size_t count);

/**
 * @brief Captures the style state of a selection as an opaque, versioned
 *        byte string: the covered column styles (column selection), row
 *        styles (row selection), and the XF and explicit-assignment mark of
 *        every cell with its own format.
 *
 * The capture is versioned (currently version 2); a capture of an earlier
 * version is still accepted by `fm_sheet_restore_range_xfs`. Neither this call
 * nor the restore has a cell ceiling.
 *
 * On success `*out_bytes` is a heap-allocated buffer of `*out_len` bytes
 * that MUST be released with `fm_buffer_destroy`. Both are zeroed before
 * validation and stay zero on failure.
 *
 * Lifetime: owned (until fm_buffer_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the rect is out of range.
 */
FM_API fm_status_t fm_sheet_snapshot_range_xfs(const fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                               uint8_t** out_bytes, size_t* out_len);

/**
 * @brief Puts a `fm_sheet_snapshot_range_xfs` capture back on sheet
 *        `sheet_index`.
 *
 * Captured axis styles and cell XFs are restored exactly; a cell of the
 * captured rect that was not captured loses its own format. A captured default
 * format on a blank cell comes back only where it overrides a row or column
 * format after the restore. Values, formulas, phonetic runs and comments are
 * not touched.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, or `bytes` is `NULL`
 *         while `len > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the bytes are not a valid capture: an unknown
 *         version, a row outside the captured rectangle or out of order, entries out of order or
 *         overlapping, a column outside the rectangle, a span whose first column is past its last,
 *         a flag or tag with an unknown value, or an XF past the cell-XF table.
 *         Nothing changes on failure.
 */
FM_API fm_status_t fm_sheet_restore_range_xfs(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes, size_t len);

/**
 * @brief Removes the height override of the inclusive row span
 *        `[first, last]`, returning those rows to the sheet default height.
 *
 * Clears the stored height and its custom flag. A row override left with
 * no other state (visible, outline level 0, no row style) is dropped,
 * except on a zeroHeight sheet, where a visible override is the row's
 * exception to the hidden default and is kept. Rows with no height
 * override are unaffected.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `last < first`, or `last` is past the last row (1048575).
 */
FM_API fm_status_t fm_sheet_reset_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);

/**
 * @brief Sheet default column / row metrics (`<sheetFormatPr>`).
 *
 * Each stored default is meaningful only when its `has_*` flag is non-zero,
 * since an absent attribute and an explicit `0` are different states. A sheet
 * loaded from a file Excel itself refuses can read a width or base above 255;
 * `fm_sheet_update_format_defaults` refuses such a value, and an `.xlsb` save of
 * it fails with `FM_STATUS_IO_FEATURE_NOT_WRITABLE`.
 */
typedef struct {
  uint32_t struct_size;
  /** The stored default column width in characters, column padding included; meaningful only when
   *  `has_default_column_width_chars` is 1. */
  double default_column_width_chars;
  /** The stored default row height in points; meaningful only when `has_default_row_height_pt` is 1. */
  double default_row_height_pt;
  /** The base column width in characters, from which Excel derives the default column width when none is stored; always
   *  present, the OOXML default is 8. */
  double base_column_width_chars;
  /** 1 when `default_column_width_chars` is stored; 0 means the sheet stores none. */
  int32_t has_default_column_width_chars;
  /** 1 when `default_row_height_pt` is stored; 0 means the sheet stores none. */
  int32_t has_default_row_height_pt;
  /** `1` when `default_row_height_pt` is an explicit override that stays as written, rather than a cached figure Excel
   *  recomputes from the Normal font. */
  int32_t custom_height;
  /** `1` when rows are hidden by default unless they say otherwise. */
  int32_t zero_height;
  /** Stored width in characters the sheet lays columns out with: the stated default column width when set, else the
   *  width Excel derives from the base column width and the Normal font (Windows, Calibri 11, base 8: 9.140625, shown
   *  as 8.43; Mac, Calibri 11, base 10: 10.83203125, shown as 10). Under a Mac profile a sheet that stores no default
   *  row height (no `<sheetFormatPr>`) derives it from base 10, although `base_column_width_chars` still reads 8.
   *  Read-only. */
  double effective_default_column_width_chars;
  /** Points the sheet lays rows out with: the stated default row height when set, else 102/7 (about 14.57); 0 when the
   *  sheet hides rows by default. Read-only. */
  double effective_default_row_height_pt;
} fm_sheet_format_defaults_t;

/**
 * @brief Reads the sheet's default column / row metrics.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_format_defaults(const fm_workbook_t* wb, size_t sheet_index,
                                                fm_sheet_format_defaults_t* out);

/**
 * @brief Partial `<sheetFormatPr>` update (`fm_sheet_update_format_defaults`),
 *        following the header's patch rules (Patches, above). The fields are
 *        the writable ones of `fm_sheet_format_defaults_t`, in its units.
 *
 * `default_column_width_chars` and `default_row_height_pt` can be cleared,
 * which removes the attribute (the sheet then lays out with the derived
 * default).
 */
typedef struct {
  uint32_t struct_size;
  /** Operation on `default_column_width_chars`: keep leaves it, set writes it, clear removes it. */
  fm_patch_op_t default_column_width_chars_op;
  /** The default column width in characters to set, 0 to 255. */
  double default_column_width_chars;
  /** Operation on `default_row_height_pt`: keep leaves it, set writes it, clear removes it. */
  fm_patch_op_t default_row_height_pt_op;
  /** The default row height in points to set, 0 to 409.6; Excel reads an `.xlsb` default above 409.5 as 409.5. */
  double default_row_height_pt;
  /** Operation on `base_column_width_chars`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t base_column_width_chars_op;
  /** The base column width in characters to set, a whole number from 0 to 255. */
  double base_column_width_chars;
  /** Operation on `custom_height`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t custom_height_op;
  /** Non-zero marks the default row height as an explicit override. */
  int32_t custom_height;
  /** Operation on `zero_height`: keep leaves it, set writes it; it cannot be cleared. */
  fm_patch_op_t zero_height_op;
  /** Non-zero hides rows by default. */
  int32_t zero_height;
} fm_sheet_format_defaults_patch_t;

/**
 * @brief Merges `patch` into the sheet's default column / row metrics.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits a field other
 *           than the two defaults, a `SET` number is negative, NaN, infinite or above its maximum,
 *           or the base column width is not a whole number.
 *         Nothing is written on failure.
 */
FM_API fm_status_t fm_sheet_update_format_defaults(fm_workbook_t* wb, size_t sheet_index,
                                                   const fm_sheet_format_defaults_patch_t* patch);

/**
 * @brief Which column-width figure a geometry call uses.
 *
 * `FM_GEOMETRY_MODE_DISPLAY` is the screen figure of the workbook profile's host:
 * Windows `Range.Width` at 96 DPI, or Mac Excel's whole-point layout.
 * `FM_GEOMETRY_MODE_PRINT` is the DPI-stable Windows figure pagination breaks on,
 * on every host. Row heights are the same in both modes.
 */
typedef enum {
  FM_GEOMETRY_MODE_DISPLAY = 0,
  FM_GEOMETRY_MODE_PRINT = 1,
} fm_geometry_mode_t;

/**
 * @brief Platform whose column-width metrics a width model uses
 *        (`fm_width_model_info_t::platform`). Zero is Windows, the platform of the
 *        runtime default profile.
 */
typedef enum {
  FM_WIDTH_PLATFORM_WIN = 0,
  FM_WIDTH_PLATFORM_MAC = 1,
} fm_width_platform_t;

/**
 * @brief The character-unit to point conversion for the workbook's Normal
 *        font in one mode.
 *
 * `points_per_char` and `padding_pt` describe a column in displayed
 * characters (`Range.ColumnWidth`): `width_pt = chars * points_per_char +
 * padding_pt`. Stored widths go through `fm_sheet_convert_column_chars_to_pt`,
 * which also snaps to whole screen pixels (whole points under a Mac profile, where a pixel is a point).
 */
typedef struct {
  uint32_t struct_size;
  /** The points added per displayed character. */
  double points_per_char;
  /** The constant padding in points. */
  double padding_pt;
  /** The point size of the Normal font. */
  double normal_font_size_pt;
  /** `0` when the Normal font is not a measured column-width calibration point of the platform and that platform's
   *  Calibri 11 figure was extrapolated. */
  int32_t calibrated;
  /** The `fm_width_platform_t` whose metrics the model uses. Print mode always reports Windows. */
  int32_t platform;
  /** The name of the Normal font the metrics were looked up for, a read-scratch-backed view. Under a Mac profile a
   *  Normal font that follows the theme reports the theme's font. */
  const char* normal_font_name;
} fm_width_model_info_t;

/**
 * @brief Reads the rectangle covering the inclusive cell range, in points
 *        from the top-left of cell A1.
 *
 * Hidden rows and columns contribute zero extent. `mode` is an
 * `fm_geometry_mode_t` ordinal.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the range
 *         is inverted or outside the grid, or `mode` is unknown.
 */
FM_API fm_status_t fm_sheet_get_cell_rect_pt(const fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                             int32_t mode, fm_rect_pt_info_t* out);

/**
 * @brief Reads the effective width of column `col` in points; zero when the
 *        column is hidden.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `row` is out of range.
 */
FM_API fm_status_t fm_sheet_get_row_height_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                              double* out_pt);

/**
 * @brief Reads the column-width conversion model in `mode`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `mode`
 *         is unknown.
 */
FM_API fm_status_t fm_sheet_get_width_model(const fm_workbook_t* wb, size_t sheet_index, int32_t mode,
                                            fm_width_model_info_t* out);

/**
 * @brief Converts a stored column width (character units with the column
 *        padding included, as `<col width>` holds it) to points. On screen
 *        the width snaps to whole pixels of the Normal font's maximum digit
 *        width (ECMA-376 §18.3.1.13; whole points under a Mac profile). Zero is
 *        exactly zero points.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `mode` is
 *         unknown, or `chars` is negative, NaN, or infinite.
 */
FM_API fm_status_t fm_sheet_convert_column_chars_to_pt(const fm_workbook_t* wb, size_t sheet_index, double chars,
                                                       int32_t mode, double* out_pt);

/**
 * @brief Inverse of `fm_sheet_convert_column_chars_to_pt`: the stored width Excel
 *        writes for a column of `pt`, rounded to whole screen pixels; zero
 *        points maps to zero.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `mode` is
 *         unknown, or `pt` is negative, NaN, or infinite.
 */
FM_API fm_status_t fm_sheet_convert_column_pt_to_chars(const fm_workbook_t* wb, size_t sheet_index, double pt,
                                                       int32_t mode, double* out_chars);

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
 *
 * Lifetime: transient.
 *
 * @return The message, or an empty string (never `NULL`) when no error has
 *         been recorded yet.
 */
FM_API const char* fm_get_last_error_message(void);

/**
 * @brief Returns the optional context string associated with the
 *        most recent error on the current thread.
 *
 * Same lifetime as `fm_get_last_error_message`.
 *
 * Lifetime: transient.
 *
 * @return The context (`key=value` pairs), or an empty string (never `NULL`)
 *         when the error recorded none.
 */
FM_API const char* fm_get_last_error_context(void);

/**
 * @brief Why a formula failed to parse; the values of
 *        `fm_parse_error_info_t::code`.
 *
 * The values are fixed: new codes are only appended. A status of 1000-1014
 * from a formula entry point carries one of these; 1002, 1008, 1011 and 1012
 * are never returned for a syntax failure.
 */
typedef enum {
  FM_PARSE_ERROR_CODE_LEXER_INVALID_CHARACTER = 0,
  FM_PARSE_ERROR_CODE_LEXER_UNTERMINATED_STRING = 1,
  FM_PARSE_ERROR_CODE_LEXER_UNTERMINATED_SHEET_QUOTE = 2,
  FM_PARSE_ERROR_CODE_LEXER_INVALID_NUMBER_LITERAL = 3,
  FM_PARSE_ERROR_CODE_LEXER_INVALID_ERROR_LITERAL = 4,
  FM_PARSE_ERROR_CODE_LEXER_INVALID_ESCAPE = 5,
  FM_PARSE_ERROR_CODE_LEXER_INVALID_REFERENCE = 6,
  FM_PARSE_ERROR_CODE_LEXER_EXCESSIVE_LENGTH = 7,
  FM_PARSE_ERROR_CODE_UNEXPECTED_TOKEN = 8,
  FM_PARSE_ERROR_CODE_UNEXPECTED_EOF = 9,
  FM_PARSE_ERROR_CODE_EXPECTED_EXPRESSION = 10,
  FM_PARSE_ERROR_CODE_EXPECTED_CLOSE_PAREN = 11,
  FM_PARSE_ERROR_CODE_UNBALANCED_BRACES = 12,
  FM_PARSE_ERROR_CODE_ARRAY_ROW_MISMATCH = 13,
  FM_PARSE_ERROR_CODE_EXPECTED_RPAREN_OR_COMMA = 14,
  FM_PARSE_ERROR_CODE_EXPECTED_COMMA_OR_SEMI_IN_ARRAY = 15,
  FM_PARSE_ERROR_CODE_INVALID_REFERENCE = 16,
  FM_PARSE_ERROR_CODE_UNSUPPORTED_CONSTRUCT = 17,
  FM_PARSE_ERROR_CODE_EXPECTED_OPEN_PAREN = 18,
  FM_PARSE_ERROR_CODE_EXPECTED_COMMA = 19,
  FM_PARSE_ERROR_CODE_UNBALANCED_BRACKETS = 20,
  FM_PARSE_ERROR_CODE_INVALID_RANGE = 21,
  FM_PARSE_ERROR_CODE_NESTED_FORMULA_TOO_DEEP = 22,
  FM_PARSE_ERROR_CODE_TOO_MANY_ERRORS = 23,
  FM_PARSE_ERROR_CODE_LET_INVALID_NAME = 24,
  FM_PARSE_ERROR_CODE_LET_WRONG_ARITY = 25,
  FM_PARSE_ERROR_CODE_LAMBDA_INVALID_PARAM = 26,
  FM_PARSE_ERROR_CODE_LAMBDA_EMPTY = 27,
  FM_PARSE_ERROR_CODE_LAMBDA_DUPLICATE_PARAM = 28,
  FM_PARSE_ERROR_CODE_ENTRY_WRONG_ARITY = 29,        /* a built-in called with an argument count it refuses */
  FM_PARSE_ERROR_CODE_ENTRY_REFERENCE_REQUIRED = 30, /* a reference-only parameter given a non-reference */
  FM_PARSE_ERROR_CODE_ENTRY_QUALIFIED_BUILTIN = 31,  /* a built-in name typed with a sheet or book qualifier */
  FM_PARSE_ERROR_CODE_ENTRY_REFUSED_WHITESPACE = 32, /* a tab, a carriage return, or a gap before a call's `(` */
  FM_PARSE_ERROR_CODE_ENTRY_PARAM_SHAPED_NAME = 33,  /* typed `_xlpm.` before a cell-shaped name */
  FM_PARSE_ERROR_CODE_ENTRY_DOUBLE_EQUALS = 34,      /* a second leading `=` */
  FM_PARSE_ERROR_CODE_ENTRY_BOM_NOT_SUPPORTED = 35,  /* a leading byte-order mark */
  FM_PARSE_ERROR_CODE_STRUCTURED_REF_INVALID = 36    /* a malformed structured reference */
} fm_parse_error_code_t;

/**
 * @brief Where and why the most recent formula syntax failure happened.
 *
 * When the failure cannot be placed in the caller's text
 * the range covers all of it.
 */
typedef struct {
  uint32_t struct_size;
  /** The parse error code identifying the kind of failure. */
  int32_t code;
  /** Offset of the first UTF-16 code unit of the offending span in the caller's text. */
  uint32_t start;
  /** Offset one past the last UTF-16 code unit of the offending span. */
  uint32_t end;
  /** The 1-based line of the failure. */
  uint32_t line;
  /** The 1-based column of the failure within its line, in UTF-16 code units. */
  uint32_t column;
  /** The same position as `start`, in code points. */
  uint32_t start_cp;
  /** The same position as `end`, in code points. */
  uint32_t end_cp;
  /** The same position as `column`, in code points (1-based). */
  uint32_t column_cp;
} fm_parse_error_info_t;

/**
 * @brief Reads the position of the most recent formula syntax failure on the
 *        calling thread.
 *
 * A failing formula entry point (cell formulas, ad-hoc evaluation, defined
 * names, conditional-format and data-validation formulas, table column
 * formulas, pivot calculated fields and items, `fm_workbook_parse_reference`,
 * `fm_localize_formula`, `fm_canonicalize_formula`) records
 * the position beside its status and message; for a record argument the
 * context names the field (`field=formula2`). The record is thread-local and
 * cleared by every other call, so read it right after the failing call. This
 * function, `fm_get_last_error_message` and `fm_get_last_error_context`
 * leave each other's state untouched, in any order. `out->struct_size` must be
 * set by the caller. `*out_present` is `1` and `*out` filled when the last
 * call recorded a syntax failure, else `0` with every field after
 * `struct_size` zeroed (still `FM_STATUS_OK` (0)).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_present` or `out` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `out->struct_size` is below the size of
 *           this version's layout.
 */
FM_API fm_status_t fm_get_last_parse_error(int32_t* out_present, fm_parse_error_info_t* out);

/**
 * @brief Returns a static C string describing `status`.
 *
 * Lifetime: static.
 *
 * @return The status name without the `FM_STATUS_` prefix for a known code
 *         (`"INVALID_ARGUMENT"` for `FM_STATUS_INVALID_ARGUMENT`), else
 *         `"UNRECOGNIZED_STATUS"`; never `NULL`.
 */
FM_API const char* fm_get_status_name(fm_status_t status);

/**
 * @brief Returns the Excel display literal for a cell error code.
 *
 * Lifetime: static.
 *
 * @return The literal for a known code (`"#DIV/0!"`, `"#REF!"`), else
 *         `"#UNKNOWN!"`; never `NULL`.
 */
FM_API const char* fm_get_error_display_name(fm_error_code_t error);

/**
 * @brief Minimum severity for the engine's structured log stream.
 *
 * Ordinals ascend with severity. `FM_LOG_LEVEL_OFF`
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

/** @brief `fm_cell_xf_t::horizontal` values (OOXML `ST_HorizontalAlignment` order). */
typedef enum {
  FM_HORIZONTAL_ALIGN_GENERAL = 0,
  FM_HORIZONTAL_ALIGN_LEFT = 1,
  FM_HORIZONTAL_ALIGN_CENTER = 2,
  FM_HORIZONTAL_ALIGN_RIGHT = 3,
  FM_HORIZONTAL_ALIGN_FILL = 4,
  FM_HORIZONTAL_ALIGN_JUSTIFY = 5,
  FM_HORIZONTAL_ALIGN_CENTER_CONTINUOUS = 6,
  FM_HORIZONTAL_ALIGN_DISTRIBUTED = 7
} fm_horizontal_align_t;

/** @brief `fm_cell_xf_t::vertical` values (OOXML `ST_VerticalAlignment` order). */
typedef enum {
  FM_VERTICAL_ALIGN_TOP = 0,
  FM_VERTICAL_ALIGN_CENTER = 1,
  FM_VERTICAL_ALIGN_BOTTOM = 2,
  FM_VERTICAL_ALIGN_JUSTIFY = 3,
  FM_VERTICAL_ALIGN_DISTRIBUTED = 4
} fm_vertical_align_t;

/** @brief `fm_cell_xf_t::reading_order` values. */
typedef enum {
  FM_READING_ORDER_CONTEXT = 0,
  FM_READING_ORDER_LEFT_TO_RIGHT = 1,
  FM_READING_ORDER_RIGHT_TO_LEFT = 2
} fm_reading_order_t;

/**
 * @brief Plain-data projection of a `formulon::io::CellXf` record.
 *
 * Frozen value type (an element of `fm_styles_batch_spec_t`): its layout never
 * changes, and extension goes through new functions.
 *
 * This record is a fixed 136-byte C ABI layout carrying every optional
 * alignment attribute represented by Formulon's style model, the six
 * `apply*` flags, `quotePrefix`, and the `<protection>` child.
 * Each `has_*` flag below distinguishes an omitted OOXML attribute from an
 * explicit zero / false value. For the four core alignment attributes,
 * values are ignored when their corresponding flag is zero; insertion
 * canonicalizes them to the model defaults (`horizontal=0`,
 * `vertical=2` (bottom), `wrap_text=0`, and
 * `justify_last_line=0`). Thus a zero-initialized record means an omitted
 * alignment child and is deduplicated with the default XF; getters return
 * those canonical defaults.
 *
 * When `has_protection` is zero `locked` / `hidden` are ignored on insertion
 * and getters report the model defaults (`locked=1`, `hidden=0`). The
 * protection fields take part in XF deduplication, so two records differing
 * only in protection stay distinct.
 *
 * Bindings copy the struct out of the workbook's styles table; subsequent
 * mutations to the workbook do not invalidate the copy.
 */
typedef struct {
  /** Index of the font in the workbook font table. */
  uint32_t font_index;
  /** Index of the fill in the workbook fill table. */
  uint32_t fill_index;
  /** Index of the border in the workbook border table. */
  uint32_t border_index;
  /** The number format id (0 to 65535): a built-in id or one declared in `<numFmts>`. */
  uint32_t num_fmt_id;
  /** The horizontal alignment. Read only when `has_horizontal` is 1. */
  int32_t horizontal;
  /** The vertical alignment; bottom when the attribute is absent. Read only when `has_vertical` is 1. */
  int32_t vertical;
  /** Whether text wraps within the cell. Read only when `has_wrap_text` is 1. */
  int32_t wrap_text;
  /** Whether the last line is justified; effective with distributed horizontal alignment. Read only when
   *  `has_justify_last_line` is 1. */
  int32_t justify_last_line;
  /** Index into `<cellStyleXfs>` of the named style this format inherits from. */
  uint32_t style_xf_index;
  /** 1 when the record has an `<alignment>` child, even an empty one; 0 means the child is omitted. */
  int32_t has_alignment;
  /** 1 when the text rotation is stated, so `text_rotation` is read; 0 means the attribute is omitted. */
  int32_t has_text_rotation;
  /** The text rotation in degrees, 0 to 180, or 255 for vertical text. Read only when `has_text_rotation` is 1. */
  uint32_t text_rotation;
  /** 1 when the indent is stated, so `indent` is read; 0 means the attribute is omitted. */
  int32_t has_indent;
  /** The indent level, 0 to 255. Read only when `has_indent` is 1. */
  uint32_t indent;
  /** 1 when the relative indent is stated, so `relative_indent` is read; 0 means the attribute is omitted. */
  int32_t has_relative_indent;
  /** The signed indent relative to the parent style. Read only when `has_relative_indent` is 1. */
  int32_t relative_indent;
  /** 1 when the shrink-to-fit flag is stated, so `shrink_to_fit` is read; 0 means the attribute is omitted. */
  int32_t has_shrink_to_fit;
  /** Whether text shrinks to fit the cell width. Read only when `has_shrink_to_fit` is 1. */
  int32_t shrink_to_fit;
  /** 1 when the reading order is stated, so `reading_order` is read; 0 means the attribute is omitted. */
  int32_t has_reading_order;
  /** The text reading order. Read only when `has_reading_order` is 1. */
  int32_t reading_order;
  /** 1 when the horizontal alignment is stated, so `horizontal` is read; 0 means the attribute is omitted. */
  int32_t has_horizontal;
  /** 1 when the vertical alignment is stated, so `vertical` is read; 0 means the attribute is omitted. */
  int32_t has_vertical;
  /** 1 when the wrap-text flag is stated, so `wrap_text` is read; 0 means the attribute is omitted. */
  int32_t has_wrap_text;
  /** 1 when the justify-last-line flag is stated, so `justify_last_line` is read; 0 means the attribute is omitted. */
  int32_t has_justify_last_line;
  /** Whether this record's number format overrides the one of its parent named style. */
  int32_t apply_number_format;
  /** Whether this record's font overrides the one of its parent named style. */
  int32_t apply_font;
  /** Whether this record's fill overrides the one of its parent named style. */
  int32_t apply_fill;
  /** Whether this record's border overrides the one of its parent named style. */
  int32_t apply_border;
  /** Whether this record's alignment overrides the one of its parent named style. */
  int32_t apply_alignment;
  /** Whether this record's protection overrides the one of its parent named style. */
  int32_t apply_protection;
  /** Whether the cell text carries a leading apostrophe that marks it as text. */
  int32_t quote_prefix;
  /** 1 when the record has an `<protection>` child, even an empty one; 0 means the child is omitted. */
  int32_t has_protection;
  /** Whether cells are locked once the sheet is protected. Read only when `has_protection` is 1; otherwise a cell is
   *  locked. */
  int32_t locked;
  /** Whether the formula is hidden once the sheet is protected. Read only when `has_protection` is 1; otherwise it is
   *  not hidden. */
  int32_t hidden;
} fm_cell_xf_t;

/** @brief `fm_font_t::underline` values (OOXML `ST_UnderlineValues` order). */
typedef enum {
  FM_UNDERLINE_STYLE_NONE = 0,
  FM_UNDERLINE_STYLE_SINGLE = 1,
  FM_UNDERLINE_STYLE_DOUBLE = 2,
  FM_UNDERLINE_STYLE_SINGLE_ACCOUNTING = 3,
  FM_UNDERLINE_STYLE_DOUBLE_ACCOUNTING = 4
} fm_underline_style_t;

/** @brief `fm_font_t::vert_align` values. */
typedef enum {
  FM_FONT_VERTICAL_ALIGN_BASELINE = 0,
  FM_FONT_VERTICAL_ALIGN_SUPERSCRIPT = 1,
  FM_FONT_VERTICAL_ALIGN_SUBSCRIPT = 2
} fm_font_vertical_align_t;

/** @brief `fm_font_t::scheme` values. */
typedef enum { FM_FONT_SCHEME_NONE = 0, FM_FONT_SCHEME_MAJOR = 1, FM_FONT_SCHEME_MINOR = 2 } fm_font_scheme_t;

/** @brief `fm_font_t::family` values (OOXML `<family>` classes). */
typedef enum {
  FM_FONT_FAMILY_NOT_APPLICABLE = 0,
  FM_FONT_FAMILY_ROMAN = 1,
  FM_FONT_FAMILY_SWISS = 2,
  FM_FONT_FAMILY_MODERN = 3,
  FM_FONT_FAMILY_SCRIPT = 4,
  FM_FONT_FAMILY_DECORATIVE = 5
} fm_font_family_t;

/**
 * @brief Plain-data projection of a `formulon::io::FontRecord`.
 *
 * Frozen value type (an element of `fm_styles_batch_spec_t`): its layout never
 * changes, and extension goes through new functions.
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
 * A required cell-font attribute with no absent state inherits from the Normal
 * font when cleared: `has_size_pt == 0` for the size, `name == NULL` for the
 * name (`fm_workbook_set_default_font` keeps the current one). An attribute
 * with an absent state goes absent: `FM_COLOR_KIND_NONE` writes no `<color>`.
 */
typedef struct {
  /** The typeface name as UTF-8; NULL when the font names none, which on a cell font takes the Normal style's name. */
  const char* name;
  /** The font size in points, 1 to 409. Read only when `has_size_pt` is 1; a cell font without one takes the Normal
   *  style's size. */
  double size_pt;
  /** Whether the font is bold. Read only when `has_bold` is 1. */
  int32_t bold;
  /** Whether the font is italic. Read only when `has_italic` is 1. */
  int32_t italic;
  /** Whether the text is struck through. Read only when `has_strike` is 1. */
  int32_t strike;
  /** 1 when the font has a `<b>` element, so `bold` is read; 0 means the element is absent. */
  int32_t has_bold;
  /** 1 when the font has a `<i>` element, so `italic` is read; 0 means the element is absent. */
  int32_t has_italic;
  /** 1 when the font has a `<strike>` element, so `strike` is read; 0 means the element is absent. */
  int32_t has_strike;
  /** 1 when the font has a `<family>` element, so `family` is read; 0 means the element is absent. */
  int32_t has_family;
  /** 1 when the font has a `<charset>` element, so `charset` is read; 0 means the element is absent. */
  int32_t has_charset;
  /** 1 when the font has a `<sz>` element, so `size_pt` is read; 0 means the element is absent. */
  int32_t has_size_pt;
  /** Whether the glyphs are drawn as outlines. Read only when `has_outline` is 1. */
  int32_t outline;
  /** Whether the glyphs carry a shadow. Read only when `has_shadow` is 1. */
  int32_t shadow;
  /** Whether the glyphs are condensed. Read only when `has_condense` is 1. */
  int32_t condense;
  /** Whether the glyphs are extended. Read only when `has_extend` is 1. */
  int32_t extend;
  /** 1 when the font has a `<outline>` element, so `outline` is read; 0 means the element is absent. */
  int32_t has_outline;
  /** 1 when the font has a `<shadow>` element, so `shadow` is read; 0 means the element is absent. */
  int32_t has_shadow;
  /** 1 when the font has a `<condense>` element, so `condense` is read; 0 means the element is absent. */
  int32_t has_condense;
  /** 1 when the font has a `<extend>` element, so `extend` is read; 0 means the element is absent. */
  int32_t has_extend;
  /** The underline style. */
  int32_t underline;
  /** The vertical alignment of the text run (baseline, superscript or subscript). */
  int32_t vert_align;
  /** The font-family class. Read only when `has_family` is 1. */
  int32_t family;
  /** The charset codepage id (for example 128 for Shift_JIS). Read only when `has_charset` is 1. */
  uint8_t charset;
  /** The `<scheme>` link to the workbook theme's major (heading) or minor (body) typeface; `0` is exactly "no
   *  `<scheme>` element", so it needs no presence flag. Set it only on a workbook that carries a theme part: a link
   *  into a theme that is not there resolves to nothing. */
  int32_t scheme;
  /** The `<color>` element; the none kind without one, which renders as the automatic colour. */
  fm_color_t color;
} fm_font_t;

/** @brief One `<stop>` of a gradient fill: a colour at `position` (0..1). */
typedef struct {
  /** The position of the stop along the gradient, 0..1. */
  double position;
  /** The colour at the stop. */
  fm_color_t color;
} fm_gradient_stop_t;

/** @brief `fm_fill_t::gradient_type` values. */
typedef enum { FM_GRADIENT_TYPE_LINEAR = 0, FM_GRADIENT_TYPE_PATH = 1 } fm_gradient_type_t;

/** @brief `fm_fill_t::pattern` values (OOXML `ST_PatternType` order). */
typedef enum {
  FM_FILL_PATTERN_NONE = 0,
  FM_FILL_PATTERN_SOLID = 1,
  FM_FILL_PATTERN_MEDIUM_GRAY = 2,
  FM_FILL_PATTERN_DARK_GRAY = 3,
  FM_FILL_PATTERN_LIGHT_GRAY = 4,
  FM_FILL_PATTERN_DARK_HORIZONTAL = 5,
  FM_FILL_PATTERN_DARK_VERTICAL = 6,
  FM_FILL_PATTERN_DARK_DOWN = 7,
  FM_FILL_PATTERN_DARK_UP = 8,
  FM_FILL_PATTERN_DARK_GRID = 9,
  FM_FILL_PATTERN_DARK_TRELLIS = 10,
  FM_FILL_PATTERN_LIGHT_HORIZONTAL = 11,
  FM_FILL_PATTERN_LIGHT_VERTICAL = 12,
  FM_FILL_PATTERN_LIGHT_DOWN = 13,
  FM_FILL_PATTERN_LIGHT_UP = 14,
  FM_FILL_PATTERN_LIGHT_GRID = 15,
  FM_FILL_PATTERN_LIGHT_TRELLIS = 16,
  FM_FILL_PATTERN_GRAY125 = 17,
  FM_FILL_PATTERN_GRAY0625 = 18
} fm_fill_pattern_t;

/**
 * @brief Plain-data projection of a `formulon::io::FillRecord`.
 *
 * Frozen value type (an element of `fm_styles_batch_spec_t`): its layout never
 * changes, and extension goes through new functions.
 */
typedef struct {
  /** The fill pattern (none, solid or one of the standard patterns). */
  int32_t pattern;
  /** The pattern's `<fgColor>`; the none kind when the element is absent. */
  fm_color_t foreground;
  /** The pattern's `<bgColor>`; the none kind when the element is absent. */
  fm_color_t background;
  /** 1 when the record is a `<gradientFill>`, in which case the pattern fields are ignored; 0 means a pattern fill. */
  int32_t has_gradient;
  /** Whether the gradient is linear or path; read only when `has_gradient` is 1. */
  int32_t gradient_type;
  /** The angle of a linear gradient in degrees. */
  double gradient_angle_deg;
  /** The left edge (0..1) of a path gradient's focus. */
  double gradient_left;
  /** The right edge (0..1) of a path gradient's focus. */
  double gradient_right;
  /** The top edge (0..1) of a path gradient's focus. */
  double gradient_top;
  /** The bottom edge (0..1) of a path gradient's focus. */
  double gradient_bottom;
  /** The gradient's colour stops; a non-finite number is stored as 0, and a write with more than 1024 stops fails with
   *  `FM_STATUS_SECURITY_RESOURCE_LIMIT` (context `part=gradient_stops requested=<n> limit=1024`). On a read the array
   * is valid until the next `fm_workbook_get_fill_at` / `fm_workbook_get_dxf_at` call, a mutation or handle
   * destruction; on a write it is borrowed for the call. */
  const fm_gradient_stop_t* gradient_stops;
  /** The number of entries in `gradient_stops`. */
  uint32_t gradient_stop_count;
} fm_fill_t;

/** @brief `fm_border_side_t::style` values (OOXML `ST_BorderStyle` order). */
typedef enum {
  FM_BORDER_STYLE_NONE = 0,
  FM_BORDER_STYLE_THIN = 1,
  FM_BORDER_STYLE_MEDIUM = 2,
  FM_BORDER_STYLE_DASHED = 3,
  FM_BORDER_STYLE_DOTTED = 4,
  FM_BORDER_STYLE_THICK = 5,
  FM_BORDER_STYLE_DOUBLE = 6,
  FM_BORDER_STYLE_HAIR = 7,
  FM_BORDER_STYLE_MEDIUM_DASHED = 8,
  FM_BORDER_STYLE_DASH_DOT = 9,
  FM_BORDER_STYLE_MEDIUM_DASH_DOT = 10,
  FM_BORDER_STYLE_DASH_DOT_DOT = 11,
  FM_BORDER_STYLE_MEDIUM_DASH_DOT_DOT = 12,
  FM_BORDER_STYLE_SLANT_DASH_DOT = 13
} fm_border_style_t;

/**
 * @brief Plain-data projection of one side of a
 *        `formulon::io::BorderRecord`.
 */
typedef struct {
  /** The line style of the side (none, thin, medium, dashed, ...). */
  int32_t style;
  /** The side's `<color>`; the none kind without one. */
  fm_color_t color;
} fm_border_side_t;

/**
 * @brief Plain-data projection of a `formulon::io::BorderRecord`.
 *
 * Frozen value type (an element of `fm_styles_batch_spec_t`): its layout never
 * changes, and extension goes through new functions.
 *
 * `diagonal_up` / `diagonal_down` carry the boolean flags as `0`/`1` int32_t to
 * match the rest of the binding surface.
 */
typedef struct {
  /** The left side. */
  fm_border_side_t left;
  /** The right side. */
  fm_border_side_t right;
  /** The top side. */
  fm_border_side_t top;
  /** The bottom side. */
  fm_border_side_t bottom;
  /** The line drawn diagonally; `diagonal_up` and `diagonal_down` choose its directions. */
  fm_border_side_t diagonal;
  /** Whether the diagonal runs from bottom left to top right. */
  int32_t diagonal_up;
  /** Whether the diagonal runs from top left to bottom right. */
  int32_t diagonal_down;
} fm_border_t;

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
  uint32_t struct_size;
  /** The fonts to insert. */
  const fm_font_t* fonts;
  /** The number of entries in `fonts`. */
  uint32_t font_count;
  /** The output array that receives the index of each font, one per entry of `fonts`. */
  uint32_t* font_indices;
  /** The fills to insert. */
  const fm_fill_t* fills;
  /** The number of entries in `fills`. */
  uint32_t fill_count;
  /** The output array that receives the index of each fill, one per entry of `fills`. */
  uint32_t* fill_indices;
  /** The borders to insert. */
  const fm_border_t* borders;
  /** The number of entries in `borders`. */
  uint32_t border_count;
  /** The output array that receives the index of each border, one per entry of `borders`. */
  uint32_t* border_indices;
  /** The cell formats to insert. */
  const fm_cell_xf_t* cell_xfs;
  /** The number of entries in `cell_xfs`. */
  uint32_t cell_xf_count;
  /** The output array that receives the index of each cell format, one per entry of `cell_xfs`. */
  uint32_t* cell_xf_indices;
  /** The number format codes to insert, as UTF-8 strings. */
  const char* const* num_fmt_codes;
  /** The number of entries in `num_fmt_codes`. */
  uint32_t num_fmt_count;
  /** The output array that receives the id of each number format, one per entry of `num_fmt_codes`. */
  uint32_t* num_fmt_ids;
} fm_styles_batch_spec_t;

/**
 * @brief Plain-data projection of one OOXML `<dxf>` differential format.
 *
 * Frozen value type: its layout never
 * changes, and extension goes through new functions.
 *
 * `number_format`, `alignment_xml`, and `protection_xml` are model-backed
 * views, valid until the next styles-table mutation (including
 * `fm_workbook_add_dxf`) or until the handle is destroyed; reads do not
 * invalidate them. The semantic content of the XML fields is preserved without
 * a typed projection, but XML parsing may normalize lexical formatting such as
 * whitespace, quotes, and empty-element spelling.
 */
typedef struct {
  /** 1 when the format sets a font; 0 means it does not. */
  int32_t has_font;
  /** The font part; read only when `has_font` is 1. */
  fm_font_t font;
  /** 1 when the format sets a fill; 0 means it does not. */
  int32_t has_fill;
  /** The fill part; read only when `has_fill` is 1. */
  fm_fill_t fill;
  /** 1 when the format sets a border; 0 means it does not. */
  int32_t has_border;
  /** The border part; read only when `has_border` is 1. */
  fm_border_t border;
  /** The number format code as UTF-8; NULL when the format sets none. The format's id is assigned on save. */
  const char* number_format;
  /** The serialized `<alignment/>` element as UTF-8; `NULL` when the format sets no alignment. */
  const char* alignment_xml;
  /** The serialized `<protection/>` element as UTF-8; `NULL` when the format sets no protection. */
  const char* protection_xml;
} fm_dxf_t;

/**
 * @brief Plain-data projection of a `formulon::io::CellStyleRecord`
 *        (one OOXML `<cellStyle>` entry).
 *
 * Frozen value type: its layout never
 * changes, and extension goes through new functions.
 *
 * `name` is a NUL-terminated UTF-8 model-backed view into the workbook's
 * styles table; it is valid until the next mutation of the cell-style
 * list (`fm_workbook_set_cell_style`, `fm_workbook_remove_cell_style`) or
 * until the handle is destroyed.
 *
 * The boolean flags follow the
 * wide-POD convention used elsewhere on this surface.
 */
typedef struct {
  /** The display name of the style as UTF-8; never NULL. */
  const char* name;
  /** Index into the `<cellStyleXfs>` table of the format the style applies, queryable via
   *  `fm_workbook_get_cell_style_xf_count` / `fm_workbook_get_cell_style_xf_at`. */
  uint32_t style_xf_index;
  /** 1 when the style is a built-in one, so `builtin_id` is read; 0 means a custom style. */
  int32_t has_builtin_id;
  /** The built-in style ordinal (0..53). Read only when `has_builtin_id` is 1. */
  uint32_t builtin_id;
  /** The outline level (0..6) of the RowLevel_n / ColLevel_n styles; 0 for other styles. */
  uint32_t outline_level;
  /** Whether the style is hidden from the style gallery. */
  int32_t hidden;
  /** Whether this is a built-in style the workbook has changed; `fm_workbook_set_cell_style_xf`
   *  sets it on every built-in style naming the edited slot. */
  int32_t custom_builtin;
} fm_cell_style_t;

/**
 * @brief Reads the `xf_index` (style record id) attached to the cell at
 *        `(row, col)` on `sheet_index`. `*out_present` is `1` and
 *        `*out_xf_index` set when a cell with content or its own format exists
 *        there (a cell with an explicit `0` reads `0`); otherwise `0` and
 *        `*out_xf_index` `0` (still `FM_STATUS_OK` (0)). A blank coordinate
 *        that only inherits a row or column format is absent.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common cell-read failure (Cell read);
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_sheet_get_cell_xf_index(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              int32_t* out_present, uint32_t* out_xf_index);

/**
 * @brief Stores `xf_index` on the cell at `(row, col)` on `sheet_index`.
 *
 * A cell with no content takes the format as its own, so a blank cell can
 * render a ruled box. Setting the default format (`0`) on a blank cell leaves
 * a cell only where it overrides a row or column format; elsewhere nothing is
 * stored. Coexists with literal / formula writes: those calls leave
 * `xf_index` untouched, so the caller can layer a style update on top of
 * either order of operations.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `xf_index` is not below `fm_workbook_get_cell_xf_count` (no cell
 *           is changed).
 */
FM_API fm_status_t fm_sheet_set_cell_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                              uint32_t xf_index);

/**
 * @brief Removes the direct style assignment from the cell at `(row, col)`
 *        on `sheet_index`, so a blank cell inherits its row / column style again.
 *
 * Resets the cell's `xf_index` to `0` and drops the explicit-assignment
 * mark; the value, formula, phonetic runs and comments are kept. An
 * absent cell is left absent (no-op). A cell holding a value or formula
 * still owns its style, which becomes the default xf; a blank cell with its
 * own format stops existing (it is no longer counted) and falls back to
 * row / column / default inheritance. Tell an unassigned
 * cell from one assigned the default xf through the `source` of
 * `fm_sheet_get_effective_style`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the address is out of range.
 */
FM_API fm_status_t fm_sheet_reset_cell_xf_index(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col);

/**
 * @brief Stores `xf_index` on every cell of the rectangle `range`.
 *
 * Blank cells take the format as their own, which is what makes an empty
 * ruled box render; cells with content keep it. A rectangle that spans whole
 * rows or whole columns is written as row or column formats and adds no
 * cells. Setting the default format (`0`) leaves a blank cell only where it
 * overrides a row or column format. Writing a large rectangle costs memory in
 * proportion to its rows, not its cells.
 *
 * Ruling a report means writing one xf across a rectangle; doing that a
 * cell at a time costs one ABI crossing per cell, which dominates the work
 * on any binding that marshals arguments.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `range` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a corner is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `xf_index` is not below
 *           `fm_workbook_get_cell_xf_count` (checked before any cell is written);
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the rectangle covers more than
 *           134,217,728 cells and is not a whole row or column selection (the engine's own
 *           ceiling, not an Excel limit; `xf_index` is checked first, nothing changes).
 */
FM_API fm_status_t fm_sheet_set_range_xf_index(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                               uint32_t xf_index);

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
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `xf_index >= cell_xfs.size()`.
 */
FM_API fm_status_t fm_workbook_get_cell_xf_at(const fm_workbook_t* wb, uint32_t xf_index, fm_cell_xf_t* out);

/**
 * @brief Reads the `font_index`-th font record from the workbook's
 *        styles table.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `font_index >= fonts.size()`.
 */
FM_API fm_status_t fm_workbook_get_font_at(const fm_workbook_t* wb, uint32_t font_index, fm_font_t* out);

/**
 * @brief Looks up the effective format string for `num_fmt_id`.
 *
 * Codes use the vocabulary of the workbook's Excel profile
 * (`fm_workbook_set_profile_id`), as Excel stores them: built-in id 14 is
 * `m/d/yy` under en-US and `yyyy/m/d` under ja-JP. Resolution searches valid
 * custom `<numFmt>` records in document order; this includes a custom record
 * below 164 that overrides a built-in slot. When no valid custom record
 * matches, the code the profile renders for a built-in id in the `0..163`
 * range is used. When present, `*out` is either a
 * model-backed view into the custom record's format string (valid until the
 * next mutation of the numFmt list, `fm_workbook_add_number_format`, or until
 * handle destruction), or a static built-in view with program lifetime.
 * The styles writer intentionally emits
 * only custom records with ids >= 164, so a below-164 override is not
 * promised to survive save/load. `*out_present` is `0` and `*out` `NULL`
 * (still `FM_STATUS_OK` (0)) when no effective mapping exists for `num_fmt_id`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_number_format(const fm_workbook_t* wb, uint32_t num_fmt_id, int32_t* out_present,
                                                 const char** out);

/**
 * @brief One declared number format (`fm_workbook_get_number_format_at`).
 *
 * `number_format` is a model-backed view, valid until the next styles
 * mutation or until the handle is destroyed.
 */
typedef struct {
  uint32_t struct_size;
  /** The number format id. */
  uint32_t num_fmt_id;
  /** The number format code. */
  const char* number_format;
} fm_number_format_info_t;

/**
 * @brief Returns the number of number formats the workbook declares: its
 *        `<numFmts>` entries, custom ids and stored overrides of built-in ids.
 *        Built-in ids without a stored code are not counted.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_number_format_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `index`-th declared number format, in ascending
 *        `num_fmt_id` order.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is not below
 *           `fm_workbook_get_number_format_count`.
 */
FM_API fm_status_t fm_workbook_get_number_format_at(const fm_workbook_t* wb, size_t index,
                                                    fm_number_format_info_t* out);

/**
 * @brief Reads the `fill_index`-th fill record from the workbook's
 *        styles table.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `fill_index >= fills.size()`.
 */
FM_API fm_status_t fm_workbook_get_fill_at(const fm_workbook_t* wb, uint32_t fill_index, fm_fill_t* out);

/**
 * @brief Reads the `border_index`-th border record from the workbook's
 *        styles table.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `border_index >= borders.size()`.
 */
FM_API fm_status_t fm_workbook_get_border_at(const fm_workbook_t* wb, uint32_t border_index, fm_border_t* out);

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
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `dxf_index >= dxfs.size()`.
 */
FM_API fm_status_t fm_workbook_get_dxf_at(const fm_workbook_t* wb, uint32_t dxf_index, fm_dxf_t* out);

/**
 * @brief Adds a `<dxf>` differential-format record to the workbook's
 *        styles table, deduplicating against existing entries.
 *
 * Unlike `<xf>`, a dxf carries inline optional style fragments. Only
 * fields whose `has_*` flag is non-zero are copied. A non-`NULL`
 * `number_format` is registered as `fm_workbook_add_number_format` does and
 * becomes the dxf's `<numFmt>`; `NULL` stores none. `alignment_xml` and `protection_xml` are always deep-copied when
 * non-`NULL`; passing `NULL` stores an empty fragment.
 *
 * A dxf no reference uses is not saved; indices are stable only within the live workbook.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `record` or `out_dxf_index` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when the present font's `name`,
 *           `number_format`, `alignment_xml` or `protection_xml` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the present font states a `size_pt` outside
 *           `1..409` or not finite;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the present fill has more than 1024 gradient stops.
 */
FM_API fm_status_t fm_workbook_add_dxf(fm_workbook_t* wb, const fm_dxf_t* record, uint32_t* out_dxf_index);

/**
 * @brief Overwrites the `dxf_index`-th `<dxf>` record in place, with the
 *        no-dedup semantics of `fm_workbook_set_font`: every conditional
 *        format, table style element and filter naming this slot renders with
 *        the new record. Fields copy as `fm_workbook_add_dxf` copies them.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `record` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when the present font's `name`,
 *           `number_format`, `alignment_xml` or `protection_xml` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the present font states a `size_pt` outside
 *           `1..409` or not finite;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the present fill has more than 1024 gradient stops;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `dxf_index` is not below the dxf count.
 */
FM_API fm_status_t fm_workbook_set_dxf(fm_workbook_t* wb, uint32_t dxf_index, const fm_dxf_t* record);

/**
 * @brief The region a custom table style element formats
 *        (`fm_table_style_element_t::type`), in OOXML `ST_TableStyleType`
 *        order; the schema has no default, so zero is its first member.
 */
typedef enum {
  FM_TABLE_STYLE_ELEMENT_TYPE_WHOLE_TABLE = 0,
  FM_TABLE_STYLE_ELEMENT_TYPE_HEADER_ROW = 1,
  FM_TABLE_STYLE_ELEMENT_TYPE_TOTAL_ROW = 2,
  FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_COLUMN = 3,
  FM_TABLE_STYLE_ELEMENT_TYPE_LAST_COLUMN = 4,
  FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_ROW_STRIPE = 5,
  FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_ROW_STRIPE = 6,
  FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_COLUMN_STRIPE = 7,
  FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_COLUMN_STRIPE = 8,
  FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_HEADER_CELL = 9,
  FM_TABLE_STYLE_ELEMENT_TYPE_LAST_HEADER_CELL = 10,
  FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_TOTAL_CELL = 11,
  FM_TABLE_STYLE_ELEMENT_TYPE_LAST_TOTAL_CELL = 12,
  FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_SUBTOTAL_COLUMN = 13,
  FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_SUBTOTAL_COLUMN = 14,
  FM_TABLE_STYLE_ELEMENT_TYPE_THIRD_SUBTOTAL_COLUMN = 15,
  FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_SUBTOTAL_ROW = 16,
  FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_SUBTOTAL_ROW = 17,
  FM_TABLE_STYLE_ELEMENT_TYPE_THIRD_SUBTOTAL_ROW = 18,
  FM_TABLE_STYLE_ELEMENT_TYPE_BLANK_ROW = 19,
  FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_COLUMN_SUBHEADING = 20,
  FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_COLUMN_SUBHEADING = 21,
  FM_TABLE_STYLE_ELEMENT_TYPE_THIRD_COLUMN_SUBHEADING = 22,
  FM_TABLE_STYLE_ELEMENT_TYPE_FIRST_ROW_SUBHEADING = 23,
  FM_TABLE_STYLE_ELEMENT_TYPE_SECOND_ROW_SUBHEADING = 24,
  FM_TABLE_STYLE_ELEMENT_TYPE_THIRD_ROW_SUBHEADING = 25,
  FM_TABLE_STYLE_ELEMENT_TYPE_PAGE_FIELD_LABELS = 26,
  FM_TABLE_STYLE_ELEMENT_TYPE_PAGE_FIELD_VALUES = 27,
} fm_table_style_element_type_t;

/**
 * @brief One region of a custom table style (`fm_table_style_t`).
 */
typedef struct {
  /** The table region this element formats (an `fm_table_style_element_type_t`); a file spelling outside the
   *  enumeration reads as -1. */
  int32_t type;
  /** The stripe band size in rows or columns, at least 1; 1 for a region that is not a stripe. */
  uint32_t size;
  /** 1 when `dxf_index` is set; 0 means the region is listed without formatting. */
  int32_t has_dxf_index;
  /** Index of the differential format applied to the region; read only when `has_dxf_index` is 1. A stored index past
   *  the workbook's dxf count reads as absent. */
  uint32_t dxf_index;
} fm_table_style_element_t;

/**
 * @brief A custom table / PivotTable style from `<tableStyles>`.
 *
 * On read, `name` is a model-backed view valid until the next styles
 * mutation, and `elements` is a read-scratch-backed array.
 */
typedef struct {
  uint32_t struct_size;
  /** The style name. */
  const char* name;
  /** Whether PivotTables may use the style (an `fm_tristate_t`): default or true is usable, false is not and is
   *  refused at entry for a PivotTable; a read returns true or false. */
  int32_t pivot; /* fm_tristate_t */
  /** Whether tables may use the style (an `fm_tristate_t`): default or true is usable, false is not and is refused
   *  at entry for a table; a read returns true or false. */
  int32_t table; /* fm_tristate_t */
  /** The formatted regions of the style. */
  const fm_table_style_element_t* elements;
  /** The number of entries in `elements`. */
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
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_workbook_get_table_style_at(const fm_workbook_t* wb, size_t index, fm_table_style_t* out);

/**
 * @brief Reads the custom table style named `name` (Unicode simple case
 *        folding, as `fm_workbook_remove_table_style` matches). `*out_present` is `1`
 *        and `*out` filled when there is one, else `0` and `*out` zeroed
 *        (still `FM_STATUS_OK` (0)).
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is malformed UTF-8.
 */
FM_API fm_status_t fm_workbook_get_table_style(const fm_workbook_t* wb, const char* name, int32_t* out_present,
                                               fm_table_style_t* out);

/**
 * @brief Adds a custom table style, or replaces the one with the same name
 *        (Unicode simple case folding) in place.
 *
 * The name must be non-empty and must not be a built-in style name
 * (`TableStyleLight1`, `PivotStyleMedium2`, `SlicerStyleLight1`, ...). A loaded file's custom
 * style that has a built-in name is dropped, as the desktop application does; the tables and
 * PivotTables naming it keep the name. Each element type may
 * appear once. A table's style name or a pivot's style name may then name it,
 * except that a kind whose flag is false is refused at entry. A name that is
 * already in use cannot take flags that exclude a kind of its users.
 * The first style added to a workbook without `<tableStyles>` sets the
 * defaults to `TableStyleMedium2` / `PivotStyleLight16`, as the desktop
 * application writes them.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `style` or `style->name` is `NULL`,
 *         or `elements` is `NULL` with a non-zero count;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an invalid name, an element type outside its
 *         enumeration or repeated, or a zero `size`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when an element's `dxf_index` is not below the dxf count;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when a replacement's flags exclude a kind that a
 *         table, a PivotTable or the default style of that kind still uses; nothing changes.
 */
FM_API fm_status_t fm_workbook_set_table_style(fm_workbook_t* wb, const fm_table_style_t* style);

/**
 * @brief Removes the custom table style `name` (Unicode simple case folding).
 *
 * Every table using it takes the default table style and every PivotTable
 * using it takes the default PivotTable style. A default naming it first
 * reverts to the built-in default (`TableStyleMedium2` / `PivotStyleLight16`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) if `name` is not well-formed UTF-8;
 *         `FM_STATUS_NOT_FOUND` (6) when no custom style has that name.
 */
FM_API fm_status_t fm_workbook_remove_table_style(fm_workbook_t* wb, const char* name);

/**
 * @brief The default table and PivotTable style names
 *        (`fm_workbook_get_table_style_defaults`). Each is a model-backed
 *        view valid until the next styles mutation, `NULL` when unset.
 */
typedef struct {
  uint32_t struct_size;
  /** The style a new table takes when none is given; `NULL` when unset. */
  const char* table_style_name;
  /** The style a new PivotTable takes; `NULL` when unset. */
  const char* pivot_style_name;
} fm_table_style_defaults_t;

/**
 * @brief Reads the default table and PivotTable style names into `*out`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_table_style_defaults(const fm_workbook_t* wb, fm_table_style_defaults_t* out);

/**
 * @brief Partial update of the default table and PivotTable style names
 *        (`fm_workbook_update_table_style_defaults`), following the header's
 *        patch rules (Patches, above). Both names can be cleared, which unsets
 *        the default.
 */
typedef struct {
  uint32_t struct_size;
  /** The update operation for `table_style_name`: keep it, clear it, or set it to the value beside it. */
  fm_patch_op_t table_style_name_op;
  /** The new default table style name when the operation is set; non-empty. */
  const char* table_style_name;
  /** The update operation for `pivot_style_name`: keep it, clear it, or set it to the value beside it. */
  fm_patch_op_t pivot_style_name_op;
  /** The new default PivotTable style name when the operation is set; non-empty. */
  const char* pivot_style_name;
} fm_table_style_defaults_patch_t;

/**
 * @brief Merges `patch` into the default table and PivotTable style names.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a `SET` name is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, a `SET` name is `NULL`
 *         or empty, or `table_style_name` / `pivot_style_name` names a custom style that does not
 *         apply to tables / PivotTables or a built-in slicer style. Nothing is written on failure.
 */
FM_API fm_status_t fm_workbook_update_table_style_defaults(fm_workbook_t* wb,
                                                           const fm_table_style_defaults_patch_t* patch);

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
 *        `fm_cell_style_t` for lifetime rules.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index >= cell_styles.size()`.
 */
FM_API fm_status_t fm_workbook_get_cell_style_at(const fm_workbook_t* wb, size_t index, fm_cell_style_t* out);

/**
 * @brief Reads the named cell style `name`, resolved as the name identity
 *        notes above describe. `*out_present` is `1` and
 *        `*out` filled when there is one, else `0` and `*out` zeroed (still
 *        `FM_STATUS_OK` (0)). Lifetimes as `fm_workbook_get_cell_style_at`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is malformed UTF-8.
 */
FM_API fm_status_t fm_workbook_get_cell_style(const fm_workbook_t* wb, const char* name, int32_t* out_present,
                                              fm_cell_style_t* out);

/**
 * @brief Returns the number of `<cellStyleXfs>` records — the named-
 *        style xf table. This table is independent of `cellXfs` and
 *        is referenced by `fm_cell_style_t::style_xf_index`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_workbook_get_cell_style_xf_count(const fm_workbook_t* wb, size_t* out_count);

/**
 * @brief Reads the `cell_style_xf_index`-th `<cellStyleXfs>` record (named-style xf
 *        table). Output shape mirrors `fm_workbook_get_cell_xf_at`, including
 *        the optional alignment attributes and their presence flags.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `cell_style_xf_index >= cell_style_xfs.size()`.
 */
FM_API fm_status_t fm_workbook_get_cell_style_xf_at(const fm_workbook_t* wb, uint32_t cell_style_xf_index,
                                                    fm_cell_xf_t* out);

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
 * `record.name` is a NUL-terminated UTF-8 pointer, or `NULL` for the Normal
 * font's name. The string is copied into the styles table and may be freed by
 * the caller after the call returns.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `record` or `out_font_index` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `record->name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a stated `size_pt` is outside `1..409` or not finite.
 */
FM_API fm_status_t fm_workbook_add_font(fm_workbook_t* wb, const fm_font_t* record, uint32_t* out_font_index);

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
 * into the styles table, `NULL` for the Normal font's name.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `record` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `record->name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a stated `size_pt` is outside `1..409` or not finite;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `font_index >= fonts.size()`.
 */
FM_API fm_status_t fm_workbook_set_font(fm_workbook_t* wb, uint32_t font_index, const fm_font_t* record);

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
 * Reading the current default is `fm_workbook_get_font_at(wb, 0, out)`.
 *
 * Every field but a NULL `name` and an unset `size_pt` is written as the
 * record states; `scheme` and `charset` are never carried over from the
 * font being replaced.
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
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `record` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `record->name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a stated `size_pt` is outside `1..409` or not finite.
 */
FM_API fm_status_t fm_workbook_set_default_font(fm_workbook_t* wb, const fm_font_t* record);

/**
 * @brief Adds a fill record to the workbook's styles table, deduplicating
 *        against existing entries.
 *
 * Linear-search dedup on the same writer-visible-state basis as
 * `fm_workbook_add_font`: a record that differs only in its `foreground` / `background`
 * colour specification gets its own index. When no match exists the
 * record is appended and the new (now-largest) index is returned. The
 * dedup is `O(N)` per call; callers that bulk-build a workbook should
 * batch when possible.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `record` or `out_fill_index` is `NULL`;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when `record->gradient_stop_count` is above 1024.
 */
FM_API fm_status_t fm_workbook_add_fill(fm_workbook_t* wb, const fm_fill_t* record, uint32_t* out_fill_index);

/**
 * @brief Overwrites the `fill_index`-th fill record in place, with the
 *        bulk-restyle and no-dedup semantics of `fm_workbook_set_font`:
 *        every `<xf>` naming this slot renders with the new record.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `record` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a value `fm_workbook_add_fill` refuses is set;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when `record->gradient_stop_count` is above 1024;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `fill_index` is not below the fill count.
 */
FM_API fm_status_t fm_workbook_set_fill(fm_workbook_t* wb, uint32_t fill_index, const fm_fill_t* record);

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
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `record` or `out_border_index` is `NULL`.
 */
FM_API fm_status_t fm_workbook_add_border(fm_workbook_t* wb, const fm_border_t* record, uint32_t* out_border_index);

/**
 * @brief Overwrites the `border_index`-th border record in place, with the
 *        bulk-restyle and no-dedup semantics of `fm_workbook_set_font`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `record` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a value `fm_workbook_add_border` refuses is set;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `border_index` is not below the border count.
 */
FM_API fm_status_t fm_workbook_set_border(fm_workbook_t* wb, uint32_t border_index, const fm_border_t* record);

/**
 * @brief Adds a number-format code to the workbook's styles table,
 *        deduplicating against existing entries. Returns the resolved
 *        OOXML `numFmtId`.
 *
 * `number_format` is in the vocabulary of the workbook's Excel profile, so
 * `fm_workbook_get_number_format(add(x))` returns `x`. An id's effective
 * mapping is the first valid custom `<numFmt>` record for that id in document
 * order, falling back to the code the profile renders for the built-in id.
 * The steps are:
 *   1. If `number_format` matches an id in the built-in range (`0..163`) under
 *      that mapping (lowest id first), the id is returned and nothing is
 *      appended to the table. `m/d/yy` is id 14 under en-US; the ECMA-376
 *      spelling `mm-dd-yy` is not, and becomes a custom format.
 *   2. Otherwise existing custom entries whose ids effectively resolve to
 *      their code are searched; the first match returns its registered id.
 *   3. Otherwise a new custom entry is appended with id
 *      `max(existing_custom_id, 163) + 1`, and the new id is returned.
 *      When the 16-bit OOXML id space is exhausted, the function returns
 *      `FM_STATUS_PRECONDITION_FAILED` (9) and leaves both the table and `*out_num_fmt_id`
 *      unchanged.
 *
 * `number_format` must be a NUL-terminated UTF-8 pointer; passing `NULL`
 * is treated as the empty string.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_num_fmt_id` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `number_format` is malformed UTF-8;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when no custom OOXML id remains.
 */
FM_API fm_status_t fm_workbook_add_number_format(fm_workbook_t* wb, const char* number_format,
                                                 uint32_t* out_num_fmt_id);

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
 *   * A non-zero `record.style_xf_index` must satisfy `< cell_style_xf_count`.
 *   * `record.num_fmt_id` must either be a built-in id with a format
 *     code in at least one locale (`5`-`8`, `41`-`44` included) or already
 *     registered as a custom entry. Unknown ids, and any id above 65535,
 *     surface `FM_STATUS_INVALID_ARGUMENT` (2).
 *   * Alignment values outside their Excel ranges are rejected when the
 *     matching presence flag is set.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `record` or `out_xf_index` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `font_index`, `fill_index`, `border_index` or
 *         `style_xf_index` is past its table;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an unregistered `num_fmt_id` or an alignment value
 *         outside its range.
 */
FM_API fm_status_t fm_workbook_add_cell_xf(fm_workbook_t* wb, const fm_cell_xf_t* record, uint32_t* out_xf_index);

/**
 * @brief Overwrites the `xf_index`-th `<xf>` record in place, with the
 *        bulk-restyle and no-dedup semantics of `fm_workbook_set_font`:
 *        every cell, row and column using this cell format renders with the
 *        new record. The record is validated as `fm_workbook_add_cell_xf`
 *        validates it.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `record` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `xf_index` is not below the cell-xf count, or
 *         `font_index`, `fill_index`, `border_index` or `style_xf_index` is past its table;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an unregistered `num_fmt_id` or an alignment value
 *         outside its range.
 */
FM_API fm_status_t fm_workbook_set_cell_xf(fm_workbook_t* wb, uint32_t xf_index, const fm_cell_xf_t* record);

/**
 * @brief Adds a named-style xf (`<cellStyleXfs>`) including optional
 *        alignment attributes, deduplicating against existing entries.
 *
 * `*out_xf_index` receives the index of the new or equal existing record.
 * `has_alignment` preserves an explicit empty child.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `record` or `out_xf_index` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `font_index`, `fill_index` or
 *           `border_index` is past its table;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an unregistered `num_fmt_id` or an alignment value
 *           outside its range.
 */
FM_API fm_status_t fm_workbook_add_cell_style_xf(fm_workbook_t* wb, const fm_cell_xf_t* record, uint32_t* out_xf_index);

/**
 * @brief Overwrites the `cell_style_xf_index`-th named-style xf
 *        (`<cellStyleXfs>`) in place, with the bulk-restyle and no-dedup
 *        semantics of `fm_workbook_set_font`: every named style naming this
 *        slot takes the new record, and when the record changes the slot each
 *        built-in style naming it is marked changed
 *        (`fm_cell_style_t.custom_builtin`). The record is validated as
 *        `fm_workbook_add_cell_style_xf` validates it.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `record` is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `cell_style_xf_index` is not below the
 *           named-style xf count, or `font_index`, `fill_index` or `border_index` is past its
 *           table;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an unregistered `num_fmt_id` or an alignment value
 *           outside its range.
 */
FM_API fm_status_t fm_workbook_set_cell_style_xf(fm_workbook_t* wb, uint32_t cell_style_xf_index,
                                                 const fm_cell_xf_t* record);

/**
 * @brief Adds or replaces the named `<cellStyle>` record `record->name`.
 *
 * The name resolves as the name identity notes above describe. A custom style
 * it resolves to is overwritten field for field; a built-in one keeps its
 * `builtin_id` and `outline_level`, even when `has_builtin_id` is 0, and takes
 * the record's name, `style_xf_index`, `hidden` and `custom_builtin`. Otherwise
 * the record is appended.
 * `record->name` is copied. `style_xf_index` must reference
 * an existing named-style xf; with `has_builtin_id` set `builtin_id` is
 * `0..53`, and `has_builtin_id == 0` makes a custom style; `outline_level` is
 * `0..6` for `builtin_id` 1 / 2 and must be `0` otherwise.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `record` or `record->name` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `record->name` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an empty name or an out-of-range field;
 *         `FM_STATUS_ALREADY_EXISTS` (7) when the name resolves to a built-in
 *           style and `record` carries a different `builtin_id`, or when it collides with another
 *           style under the wide fold.
 */
FM_API fm_status_t fm_workbook_set_cell_style(fm_workbook_t* wb, const fm_cell_style_t* record);

/**
 * @brief Removes the named `<cellStyle>` record `name`, resolved as the name
 *        identity notes above describe.
 *
 * When no other named style shares the removed record's `style_xf_index`, every
 * cell xf that inherits from it is repointed at Normal (`style_xf_index` 0), and each
 * attribute group whose `apply_*` flag is zero takes Normal's value; groups
 * whose flag is set keep their own. Neither `<cellXfs>` nor
 * `<cellStyleXfs>` is compacted, so every xf index stays valid.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `name` is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `name` is malformed UTF-8;
 *         `FM_STATUS_NOT_FOUND` (6) when no style has that name, or two styles match it only under
 *           the wide fold;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when it is the Normal style (`builtin_id` 0).
 */
FM_API fm_status_t fm_workbook_remove_cell_style(fm_workbook_t* wb, const char* name);

/**
 * @brief Adds and deduplicates a batch of fonts, fills, borders, and cell xfs.
 *
 * Empty arrays may be `NULL`. Each element is validated as its single-record
 * counterpart (`fm_workbook_add_font`, `fm_workbook_add_fill`,
 * `fm_workbook_add_border`, `fm_workbook_add_number_format`,
 * `fm_workbook_add_cell_xf`) validates it. The operation is staged and
 * committed as one transaction: any returned error leaves the workbook and
 * every caller-provided output array unchanged.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `batch` is `NULL`, or a non-empty
 *           input or output array is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a font name or format code is
 *           malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `batch->struct_size` is wrong, a font size is
 *           outside `1..409`, or an xf is invalid;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when an xf's font, fill, border or style xf index
 *           is past its table (counting the batch's own records);
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when a fill has more than 1024 gradient stops;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when custom number-format ids are exhausted.
 */
FM_API fm_status_t fm_workbook_add_batch(fm_workbook_t* wb, const fm_styles_batch_spec_t* batch);

/* -------------------------------------------------------------------------- */
/* Theme, colour resolution and effective style                               */
/* -------------------------------------------------------------------------- */

/** @brief The theme's 12 colours (`a:clrScheme`). */
typedef struct {
  uint32_t struct_size;
  /** The 12 scheme colours as 0xFFRRGGBB in `a:clrScheme` order: dk1, lt1, dk2, lt2, accent1 to accent6, hlink,
   *  folHlink. This is not the order of a colour's theme index, which swaps the first two pairs (0 = lt1, 1 = dk1, 2 =
   *  lt2, 3 = dk2). */
  uint32_t argb[12];
} fm_theme_colors_t;

/**
 * @brief The theme's major (heading) and minor (body) typefaces.
 *
 * Getter strings are read-scratch-backed views.
 * An empty face writes `typeface=""`, as the Office theme's East Asian
 * entries are.
 */
typedef struct {
  uint32_t struct_size;
  /** The Latin typeface of the major (heading) font set. */
  const char* major_latin;
  /** The East Asian typeface of the major (heading) font set: the `a:ea` face, else the Jpan script face. */
  const char* major_east_asian;
  /** The Latin typeface of the minor (body) font set. */
  const char* minor_latin;
  /** The East Asian typeface of the minor (body) font set: the `a:ea` face, else the Jpan script face. */
  const char* minor_east_asian;
} fm_theme_fonts_t;

/** @brief Where a colour is used; decides what an automatic colour means. */
typedef enum {
  FM_COLOR_CONTEXT_FONT = 0,
  FM_COLOR_CONTEXT_FILL_FOREGROUND = 1,
  FM_COLOR_CONTEXT_FILL_BACKGROUND = 2,
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

/** @brief The level that supplied an effective style's xf (`fm_effective_style_info_t::source`). */
typedef enum {
  FM_EFFECTIVE_STYLE_SOURCE_CELL = 0,
  FM_EFFECTIVE_STYLE_SOURCE_ROW = 1,
  FM_EFFECTIVE_STYLE_SOURCE_COLUMN = 2,
  FM_EFFECTIVE_STYLE_SOURCE_DEFAULT = 3,
} fm_effective_style_source_t;

/**
 * @brief The formatting a cell shows, read from the selected `<cellXfs>`
 *        record directly.
 *
 * An xf, font, fill or border index beyond its table reads as 0. Colours are
 * resolved against the theme and `<indexedColors>`. Named-style inheritance is
 * not recomputed and conditional formatting is not applied.
 */
typedef struct {
  uint32_t struct_size;
  /** The index of the selected cell format record. */
  uint32_t xf_index;
  /** The level that supplied `xf_index`, an `fm_effective_style_source_t`: the cell, its row (a row style takes effect
   *  only with `customFormat`), its column, or the default xf 0. */
  int32_t source;
  /** The index into the font table. */
  uint32_t font_index;
  /** The index into the fill table. */
  uint32_t fill_index;
  /** The index into the border table. */
  uint32_t border_index;
  /** The resolved font color as ARGB. */
  uint32_t font_argb;
  /** How the font color was resolved, an `fm_color_resolution_t`. */
  int32_t font_resolution;
  /** The resolved fill foreground color as ARGB. */
  uint32_t fill_foreground_argb;
  /** How the fill foreground color was resolved, an `fm_color_resolution_t`. */
  int32_t fill_foreground_resolution;
  /** The resolved fill background color as ARGB. */
  uint32_t fill_background_argb;
  /** How the fill background color was resolved, an `fm_color_resolution_t`. */
  int32_t fill_background_resolution;
  /** The resolved border colors as ARGB: left, right, top, bottom, diagonal. */
  uint32_t border_argb[5];
  /** How each border color was resolved, `fm_color_resolution_t` values in the order left, right, top, bottom,
   *  diagonal. */
  int32_t border_resolution[5];
  /** 1 when the cell is locked; 0 when it is not. */
  int32_t locked;
  /** 1 when the cell formula is hidden; 0 when it is not. */
  int32_t hidden;
  /** The number format code, a read-scratch-backed view. */
  const char* number_format;
} fm_effective_style_info_t;

/** @brief Where theme values were read from (`fm_workbook_get_theme_colors`'s `*out_source`). */
typedef enum {
  FM_THEME_SOURCE_PART = 0,        /**< The workbook's own theme part. */
  FM_THEME_SOURCE_DEFAULT = 1,     /**< The workbook has no theme part; the default theme. */
  FM_THEME_SOURCE_UNPARSEABLE = 2, /**< The theme part cannot be parsed; the default theme. */
} fm_theme_source_t;

/**
 * @brief Reads the theme's 12 scheme colours.
 *
 * `*out_source` is an `fm_theme_source_t`: `FM_THEME_SOURCE_PART` when they
 * come from the workbook's theme part, otherwise the default theme's
 * colours are reported.
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
 * Lifetime: transient.
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
 *        major or minor font set.
 *
 * `scheme` is an `fm_font_scheme_t`: `FM_FONT_SCHEME_MAJOR` or
 * `FM_FONT_SCHEME_MINOR` selects the set, and the same holds for every
 * theme script-font call. The default Office theme's 47 scripts are reported
 * when the theme part is absent or unparseable (the same rule as
 * `fm_workbook_get_theme_fonts`). The theme is parsed on every call.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out_count` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `scheme` is `FM_FONT_SCHEME_NONE` or any other
 *         value that is not `FM_FONT_SCHEME_MAJOR` / `FM_FONT_SCHEME_MINOR`.
 */
FM_API fm_status_t fm_workbook_get_theme_script_font_count(const fm_workbook_t* wb, int32_t scheme, size_t* out_count);

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
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `scheme` is not `FM_FONT_SCHEME_MAJOR` /
 *         `FM_FONT_SCHEME_MINOR`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is not below the count.
 */
FM_API fm_status_t fm_workbook_get_theme_script_font_at(const fm_workbook_t* wb, int32_t scheme, size_t index,
                                                        const char** out_script, const char** out_typeface);

/**
 * @brief Reads the typeface of script `script` (an ISO 15924 code such as
 *        `Jpan`, matched exactly) in the `scheme` font set (major or minor). `*out_present` is `1` and `*out_typeface`
 * set when the set lists the script, else `0` and `*out_typeface` `NULL` (still `FM_STATUS_OK` (0)).
 *
 * `*out_typeface` is a read-scratch-backed view with the lifetime
 * `fm_workbook_get_theme_script_font_at` describes.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `script` is malformed UTF-8;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `scheme` is not `FM_FONT_SCHEME_MAJOR` /
 *         `FM_FONT_SCHEME_MINOR`.
 */
FM_API fm_status_t fm_workbook_get_theme_script_font(const fm_workbook_t* wb, int32_t scheme, const char* script,
                                                     int32_t* out_present, const char** out_typeface);

/**
 * @brief Sets the typeface of one script in the `scheme` font set (major or
 *        minor). The strings are copied.
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
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `scheme` is not `FM_FONT_SCHEME_MAJOR` /
 *         `FM_FONT_SCHEME_MINOR`, `script` is not exactly four ASCII letters, or `typeface`
 *         is empty;
 *         `FM_STATUS_IO_XML_PARSE` (5007) when the existing theme part cannot be parsed.
 */
FM_API fm_status_t fm_workbook_set_theme_script_font(fm_workbook_t* wb, int32_t scheme, const char* script,
                                                     const char* typeface);

/**
 * @brief Returns the workbook to the state of having no theme part, so
 *        `fm_workbook_get_theme_colors` reports `FM_THEME_SOURCE_DEFAULT` and the default
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
FM_API fm_status_t fm_workbook_reset_theme(fm_workbook_t* wb);

/**
 * @brief Applies a theme file (`.thmx`) to the workbook, as Excel's Page Layout >
 *        Themes > Browse does.
 *
 * The theme part is replaced in place (a workbook without one gets a theme part
 * and its relationship), together with the images the theme uses; the previous
 * theme's relationships and anything only they referenced are removed. A
 * variant in the file is ignored: the file's default theme is applied. Fonts
 * that follow the theme switch to the new theme's faces, and every formula is
 * recalculated when the theme's font faces change (`CELL("width")` reads
 * them). The call is atomic: on any failure the workbook is unchanged. Windows
 * Excel's behaviour is unmeasured.
 *
 * @param wb     Workbook handle. Must be non-NULL.
 * @param bytes  The `.thmx` package. Must be non-NULL.
 * @param len    Length of `bytes`. Must be non-zero and the archive at most 64 MiB uncompressed.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `bytes` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `len == 0`;
 *         `FM_STATUS_IO_ZIP_CORRUPT` (5003) when `bytes` is not a zip archive;
 *         `FM_STATUS_IO_RELATIONSHIP_BROKEN` (5010) when the package has no office-document or theme
 *         relationship, or its target is missing;
 *         `FM_STATUS_IO_XML_PARSE` (5007) when the theme cannot be parsed;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the archive exceeds the size budget.
 */
FM_API fm_status_t fm_workbook_apply_theme(fm_workbook_t* wb, const uint8_t* bytes, size_t len);

/**
 * @brief Writes the workbook's current theme as a theme file (`.thmx`).
 *
 * The package carries the theme part (the default Office theme when the
 * workbook has none) and the images it uses, and no thumbnail. The result
 * round-trips through `fm_workbook_apply_theme`.
 *
 * Lifetime: owned (until fm_buffer_destroy).
 *
 * @param wb         Workbook handle. Must be non-NULL.
 * @param out_bytes  Receives a pointer to the freshly allocated buffer, released with
 *                   `fm_buffer_destroy` (NOT `free`).
 * @param out_len    Receives the buffer length in bytes.
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_IO_XML_PARSE` (5007) when the current theme part cannot be parsed.
 *         On failure `*out_bytes` is `NULL` and `*out_len` is 0.
 */
FM_API fm_status_t fm_workbook_save_theme(const fm_workbook_t* wb, uint8_t** out_bytes, size_t* out_len);

/**
 * @brief Resolves a colour specification to the 0xAARRGGBB Excel renders.
 *
 * Theme colours take `spec.tint`; indexed colours read the workbook's
 * `<indexedColors>` over the built-in palette; `FM_COLOR_KIND_AUTO`,
 * `FM_COLOR_KIND_NONE` and the system indices 64 / 65 depend on `context`
 * (an `fm_color_context_t`). `*out_resolution` is an
 * `fm_color_resolution_t`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an unknown `context` or `spec.kind`, or a
 *         non-finite theme tint.
 */
FM_API fm_status_t fm_workbook_resolve_color(const fm_workbook_t* wb, const fm_color_t* spec, int32_t context,
                                             uint32_t* out_argb, int32_t* out_resolution);

/**
 * @brief Reads the effective style of the cell at `(row, col)`; see
 *        `fm_effective_style_info_t`. The cell need not exist.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the
 *         coordinate is outside the grid.
 */
FM_API fm_status_t fm_sheet_get_effective_style(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                                fm_effective_style_info_t* out);

/* -------------------------------------------------------------------------- */
/* Sheet tab colour                                                           */
/* -------------------------------------------------------------------------- */

/**
 * @brief Reads the sheet's tab colour (`<sheetPr><tabColor>`).
 *
 * `*out` is the colour as stored; `kind` is `FM_COLOR_KIND_NONE` when the sheet has
 * no tab colour.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_tab_color(const fm_workbook_t* wb, size_t sheet_index, fm_color_t* out);

/**
 * @brief Sets or clears the sheet's tab colour (`<sheetPr><tabColor>`).
 *
 * `color == NULL` or `kind == FM_COLOR_KIND_NONE` removes the tab colour. Only the
 * `<tabColor>` child is rewritten; the rest of `<sheetPr>` is kept.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an unknown `kind` or a non-finite theme tint.
 */
FM_API fm_status_t fm_sheet_set_tab_color(fm_workbook_t* wb, size_t sheet_index, const fm_color_t* color);

/* -------------------------------------------------------------------------- */
/* External links                                                             */
/* -------------------------------------------------------------------------- */

/**
 * @brief Kind discriminator for `fm_external_link_info_t::kind`. Values
 *        mirror `formulon::io::ExternalLinkRecord::Kind`.
 */
typedef enum {
  FM_EXTERNAL_LINK_KIND_UNKNOWN = 0,
  FM_EXTERNAL_LINK_KIND_EXTERNAL_BOOK = 1,
  FM_EXTERNAL_LINK_KIND_OLE = 2,
  FM_EXTERNAL_LINK_KIND_DDE = 3,
} fm_external_link_kind_t;

/**
 * @brief Plain-data projection of `formulon::io::ExternalLinkRecord`.
 *
 * `rel_id`, `part_path`, and `target` are NUL-terminated UTF-8
 * model-backed views into the workbook's external-links table; they remain
 * valid until the workbook is destroyed or the external-links table is
 * replaced (the OOXML reader replaces it; entering a formula that names a
 * new book appends to it). `target_external` follows the wide-POD
 * convention used elsewhere on this surface.
 */
typedef struct {
  uint32_t struct_size;
  /** The 1-based document order in the externalReferences list: the `[n]` of formula text, as in `=[1]Sheet1!A1`. */
  uint32_t book_number;
  /** The id of the link in the workbook relationships, such as "rId3". */
  const char* rel_id;
  /** The resolved package-relative path of the body part, such as `xl/externalLinks/externalLink1.xml`. */
  const char* part_path;
  /** The URL or path of the remote workbook; `NULL` when the link's relationship names no target. */
  const char* target;
  /** `1` when the relationship's TargetMode is External. */
  int32_t target_external;
  /** The kind of link, an `fm_external_link_kind_t`. */
  int32_t kind;
} fm_external_link_info_t;

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
 * @brief Reads the `external_link_index`-th external-link record (1-based document
 *        order minus one — i.e. `0` is the first `<externalReference>`).
 *
 * The returned string pointers are model-backed views into workbook-owned
 * storage; see `fm_external_link_info_t` for lifetime rules.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `external_link_index >= external_link_count`.
 */
FM_API fm_status_t fm_workbook_get_external_link_at(const fm_workbook_t* wb, size_t external_link_index,
                                                    fm_external_link_info_t* out);

/* -------------------------------------------------------------------------- */
/* Drawing objects                                                            */
/* -------------------------------------------------------------------------- */

/*
 * An object of a sheet's drawing is addressed by its frame id (`cNvPr`).
 * Common failures of the drawing, shape and slicer entries, besides those each
 * entry lists: `FM_STATUS_NOT_FOUND` (6) when no object of the sheet's drawing
 * holds the id; `FM_STATUS_INVALID_ARGUMENT` (2) when the id belongs to an
 * object of a kind the entry does not take (a slicer passed to a picture
 * operation, a picture to a shape operation); `FM_STATUS_PRECONDITION_FAILED`
 * (9) from an insert once the drawing's id space is exhausted.
 */

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
  FM_DRAWING_OBJECT_KIND_PICTURE = 0,
  FM_DRAWING_OBJECT_KIND_SHAPE = 1,
  FM_DRAWING_OBJECT_KIND_CHART = 2,
  FM_DRAWING_OBJECT_KIND_GROUP = 3,
  FM_DRAWING_OBJECT_KIND_CONNECTOR = 4,
  FM_DRAWING_OBJECT_KIND_GRAPHIC_FRAME = 5,
  FM_DRAWING_OBJECT_KIND_OTHER = 6,
} fm_drawing_object_kind_t;

/**
 * @brief The anchor element type of a drawing object.
 *
 * `DEFAULT` is input only and asks for the operation's default: one-cell for
 * a picture, two-cell for a chart, shape or slicer. A read never returns it.
 */
typedef enum {
  FM_ANCHOR_KIND_DEFAULT = 0,
  FM_ANCHOR_KIND_ONE_CELL = 1,
  FM_ANCHOR_KIND_TWO_CELL = 2,
  FM_ANCHOR_KIND_ABSOLUTE = 3,
} fm_anchor_kind_t;

/**
 * @brief How a drawing object follows row and column edits.
 *
 * `DEFAULT` is input only and asks for the operation's default, the
 * anchor's own kind (`TWO_CELL` for a two-cell anchor). A read never
 * returns it.
 */
typedef enum {
  FM_ANCHOR_EDIT_AS_DEFAULT = 0,
  FM_ANCHOR_EDIT_AS_ONE_CELL = 1,
  FM_ANCHOR_EDIT_AS_TWO_CELL = 2,
  FM_ANCHOR_EDIT_AS_ABSOLUTE = 3,
} fm_anchor_edit_as_t;

/** @brief Format and pixel size of an image. */
typedef struct {
  uint32_t struct_size;
  /** The image's file format, an `fm_image_format_t`. */
  int32_t format;
  /** Image width in pixels. */
  uint32_t width_px;
  /** Image height in pixels. */
  uint32_t height_px;
} fm_image_info_t;

/**
 * @brief The anchor of a drawing object: read back by every drawing record
 *        and taken by every add (`fm_drawing_placement_spec_t::anchor`).
 *
 * On add, a two-cell anchor takes its box from `to_*` and ignores the size,
 * except that an all-zero `to_*` marker sizes the box from `width_emu` /
 * `height_emu`; a one-cell anchor takes the size and ignores `to_*`. A zero
 * size means the image's pixel size times 9525 for a picture and Excel's
 * 360 x 216 pt frame for a chart; a shape or slicer needs a non-zero width
 * or height (a straight connector may be zero in one direction).
 * An absolute kind is not added.
 */
typedef struct {
  /** The anchor kind: absolute, one-cell or two-cell, an `fm_anchor_kind_t`. */
  int32_t kind;
  /** How the object follows row and column edits (it moves and resizes, moves only, or stays put), an
   *  `fm_anchor_edit_as_t`; written for a two-cell anchor only. */
  int32_t edit_as;
  /** 0-based row of the top-left marker. */
  uint32_t from_row;
  /** 0-based column of the top-left marker. */
  uint32_t from_col;
  /** Vertical offset of the top-left corner within `from_row`, in EMU; for an absolute anchor, the y position from A1.
   */
  int64_t from_row_offset_emu;
  /** Horizontal offset of the top-left corner within `from_col`, in EMU; for an absolute anchor, the x position from
   *  A1. */
  int64_t from_col_offset_emu;
  /** 0-based row of the bottom-right marker of a two-cell anchor. 0 unless the anchor is two-cell. */
  uint32_t to_row;
  /** 0-based column of the bottom-right marker of a two-cell anchor. 0 unless the anchor is two-cell. */
  uint32_t to_col;
  /** Vertical offset of the bottom-right corner within `to_row`, in EMU. 0 unless the anchor is two-cell. */
  int64_t to_row_offset_emu;
  /** Horizontal offset of the bottom-right corner within `to_col`, in EMU. 0 unless the anchor is two-cell. */
  int64_t to_col_offset_emu;
  /** The object's own unrotated width in EMU; a two-cell anchor's width is its box. */
  int64_t width_emu;
  /** The object's own unrotated height in EMU; a two-cell anchor's height is its box. */
  int64_t height_emu;
} fm_drawing_anchor_t;

/**
 * @brief One anchored object of a sheet's drawing.
 *
 * `name`, `description` and `media_path` are read-scratch-backed views (the
 * lifetime `fm_sheet_get_auto_filter` describes). A picture's format is read
 * with `fm_sheet_get_picture`.
 */
typedef struct {
  uint32_t struct_size;
  /** The object's id in the sheet's drawing. */
  uint32_t object_id;
  /** The object's kind: picture, shape, chart, group, connector, other graphic frame or other, an
   *  `fm_drawing_object_kind_t`. */
  int32_t kind;
  /** `1` when the object is hidden. */
  int32_t hidden;
  /** Where the object sits on the sheet. */
  fm_drawing_anchor_t anchor;
  /** The object's name. */
  const char* name;
  /** The object's alternative text; `NULL` when it has none. */
  const char* description;
  /** Package path the picture's image resolves to; `NULL` when there is none, as for an object that is not a picture.
   */
  const char* media_path;
} fm_drawing_object_info_t;

/**
 * @brief The common placement of a new drawing object: a picture
 *        (`fm_sheet_add_picture`), and embedded as `placement` in
 *        `fm_chart_spec_t`, `fm_shape_spec_t` and `fm_slicer_spec_t`.
 *        An embedded copy sets its own `struct_size` too.
 */
typedef struct {
  uint32_t struct_size;
  /** The object's name; `NULL`, empty or omitted takes the kind's default (`Picture <n>`, n being the object id minus
   *  one, for a picture). */
  const char* name;
  /** The object's alternative text; `NULL`, empty or omitted means none. */
  const char* description;
  /** Whether the object is hidden; unset means shown. */
  int32_t hidden; /* fm_tristate_t */
  /** Where the object is placed, following the add rules of `fm_drawing_anchor_t`; unset takes the operation's default
   *  anchor kind at A1. */
  fm_drawing_anchor_t anchor;
} fm_drawing_placement_spec_t;

/**
 * @brief Partial update of a drawing anchor, embedded as `anchor` in
 *        `fm_drawing_patch_t`. Each field of `fm_drawing_anchor_t` has its
 *        `<field>_op`; none can be cleared.
 *
 * The patch merges into the object's current anchor: a field left `KEEP`
 * keeps its value, and `kind` left `KEEP` keeps the object's kind. A `SET`
 * `kind` (one-cell or two-cell) converts the anchor: a two-cell anchor made
 * from a one-cell or absolute one takes its `to_*` marker from the current
 * size unless the patch sets it, and a one-cell anchor made from a two-cell
 * one keeps the box's size. As on add, an anchor that ends up two-cell takes
 * its box from a `SET` `to_*` field and ignores a `SET` size, and one that
 * ends up one-cell takes a `SET` size and ignores `to_*`, so a read anchor
 * feeds back. A `SET` `to_*` field without a `SET` size on an anchor that
 * ends up one-cell, or a `SET` size without a `SET` `to_*` field on one that
 * ends up two-cell, is `FM_STATUS_INVALID_ARGUMENT` (2).
 */
typedef struct {
  /** Patch operation for `kind`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t kind_op;
  /** The anchor kind: absolute, one-cell or two-cell; an absolute kind is not added. */
  int32_t kind;
  /** Patch operation for `edit_as`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t edit_as_op;
  /** How the object follows row and column edits (it moves and resizes, moves only, or stays put); written for a
   *  two-cell anchor only. */
  int32_t edit_as;
  /** Patch operation for `from_row`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t from_row_op;
  /** 0-based row of the top-left marker. */
  uint32_t from_row;
  /** Patch operation for `from_col`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t from_col_op;
  /** 0-based column of the top-left marker. */
  uint32_t from_col;
  /** Patch operation for `from_row_offset_emu`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t from_row_offset_emu_op;
  /** Vertical offset of the top-left corner within `from_row`, in EMU; for an absolute anchor, the y position from A1.
   */
  int64_t from_row_offset_emu;
  /** Patch operation for `from_col_offset_emu`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t from_col_offset_emu_op;
  /** Horizontal offset of the top-left corner within `from_col`, in EMU; for an absolute anchor, the x position from
   *  A1. */
  int64_t from_col_offset_emu;
  /** Patch operation for `to_row`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t to_row_op;
  /** 0-based row of the bottom-right marker of a two-cell anchor. Not accepted on a one-cell anchor. */
  uint32_t to_row;
  /** Patch operation for `to_col`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t to_col_op;
  /** 0-based column of the bottom-right marker of a two-cell anchor. Not accepted on a one-cell anchor. */
  uint32_t to_col;
  /** Patch operation for `to_row_offset_emu`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t to_row_offset_emu_op;
  /** Vertical offset of the bottom-right corner within `to_row`, in EMU. Not accepted on a one-cell anchor. */
  int64_t to_row_offset_emu;
  /** Patch operation for `to_col_offset_emu`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t to_col_offset_emu_op;
  /** Horizontal offset of the bottom-right corner within `to_col`, in EMU. Not accepted on a one-cell anchor. */
  int64_t to_col_offset_emu;
  /** Patch operation for `width_emu`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t width_emu_op;
  /** The object's own unrotated width in EMU. Not accepted on a two-cell anchor, except as the size of an all-zero
   *  bottom-right marker box on add. */
  int64_t width_emu;
  /** Patch operation for `height_emu`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t height_emu_op;
  /** The object's own unrotated height in EMU. Not accepted on a two-cell anchor, except as the size of an all-zero
   *  bottom-right marker box on add. */
  int64_t height_emu;
} fm_drawing_anchor_patch_t;

/**
 * @brief Partial update of the fields every drawing object has
 *        (`fm_sheet_update_drawing`, and the `common` half of
 *        `fm_sheet_update_shape` / `fm_sheet_update_slicer`), following the
 *        header's patch rules (Patches, above).
 */
typedef struct {
  uint32_t struct_size;
  /** Patch operation for `name`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t name_op;
  /** The new name; it cannot be empty, though duplicate names are allowed, and a slicer's name cannot be set (it ties
   *  the frame to the slicer part). */
  const char* name;
  /** Patch operation for `description`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t description_op;
  /** The new alternative text; clearing it, or an empty string, removes the description. */
  const char* description;
  /** Patch operation for `hidden`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t hidden_op;
  /** Whether the object is hidden. */
  int32_t hidden;
  /** Update merged into the object's current anchor, following `fm_drawing_anchor_patch_t`. */
  fm_drawing_anchor_patch_t anchor;
} fm_drawing_patch_t;

/** @brief `fm_sheet_restore_drawing` flag: restore as a copy with a fresh id, added on top. */
#define FM_DRAWING_RESTORE_NEW_ID 1u

/**
 * @brief Sniffs the format and pixel size of PNG, JPEG, GIF or BMP bytes
 *        from the header alone. Needs no workbook.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_IO_IMAGE_UNSUPPORTED` (5023) when the bytes are not a recognised
 *         image, the header is truncated, or a dimension is zero.
 */
FM_API fm_status_t fm_get_image_info(const uint8_t* bytes, size_t len, fm_image_info_t* out);

/**
 * @brief Number of anchored objects in the sheet's drawing; `0` for a sheet
 *        without a drawing.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_drawing_object_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads object `index` of the sheet's drawing, in document order.
 *
 * Document order is the stacking order: item 0 is at the back and the last
 * item is on top.
 *
 * The three strings are read-scratch-backed views.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_drawing_object_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                  fm_drawing_object_info_t* out);

/**
 * @brief Reads the object of the sheet's drawing whose frame id is
 *        `object_id`. `*out_present` is `1` and `*out` filled when there is
 *        one, else `0` and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * The three strings are read-scratch-backed views.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_drawing_object(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                               int32_t* out_present, fm_drawing_object_info_t* out);

/**
 * @brief Returns the media bytes and header info of the picture `object_id`.
 *        `*out_present` is `1` and the outputs filled when the sheet has that
 *        picture with a present media part, else `0` with the outputs zeroed
 *        (still `FM_STATUS_OK` (0)).
 *
 * `*out_bytes` points into the workbook's own storage and is not
 * NUL-terminated: it stays valid until the next call that mutates the
 * workbook (insert, remove, any edit) or the handle is destroyed; unrelated
 * reads do not invalidate it. Copy it to keep it longer.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_IMAGE_UNSUPPORTED` (5023) when the media bytes are not a
 *         recognised image;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_picture(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                        int32_t* out_present, const uint8_t** out_bytes, size_t* out_len,
                                        fm_image_info_t* out_info);

/**
 * @brief A picture placed in a cell (Excel's Place in Cell): the image file as stored and its alt text.
 *
 * On a read, `bytes`, `len` and `alt_text` are transient views (see `fm_sheet_get_cell_picture`) and `format`,
 * `width_px` and `height_px` describe the image. On input only `bytes`, `len` and `alt_text` are read; `format`,
 * `width_px` and `height_px` are ignored.
 */
typedef struct {
  uint32_t struct_size;
  /** The image file as stored, PNG or JPEG; not NUL-terminated. */
  const uint8_t* bytes;
  /** The length of `bytes` in bytes. */
  uint64_t len;
  /** The picture's alt text as NUL-terminated UTF-8; `NULL` when it has none. On input `NULL` is none and `""` is
   *  `FM_STATUS_INVALID_ARGUMENT`. */
  const char* alt_text;
  /** The image's file format, an `fm_image_format_t`; read-only. */
  int32_t format;
  /** Image width in pixels; read-only. */
  uint32_t width_px;
  /** Image height in pixels; read-only. */
  uint32_t height_px;
} fm_cell_picture_t;

/**
 * @brief Reads the picture the cell at `(row, col)` of `sheet_index` holds. `*out_present` is `1` and `*out` filled
 *        when the cell holds one, else `0` and `*out` zeroed (still `FM_STATUS_OK` (0)). A picture a formula produced
 *        is reported too.
 *
 * `out->bytes` and `out->alt_text` point into storage the handle keeps for this call: they stay valid until the next
 * call of this function that finds a picture, the next call that mutates the workbook, or until the handle is
 * destroyed; unrelated reads do not invalidate them. Copy them to keep them longer. A picture cell evaluates as
 * `#VALUE!` in formulas.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `struct_size` is below the v1 size or `(row, col)` is outside the
 *         grid.
 */
FM_API fm_status_t fm_sheet_get_cell_picture(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             int32_t* out_present, fm_cell_picture_t* out);

/**
 * @brief Places the PNG or JPEG image `picture->bytes[0..len)` in the cell at `(row, col)` of `sheet_index`, with
 *        `picture->alt_text` when given. `NULL` for `picture` removes the cell's picture and empties the cell; a cell
 *        without one is left as it is.
 *
 * The image is stored as given. Placing is a write of `#VALUE!` to the cell: it replaces a formula, is refused on an
 * array-formula block or data table as a value write is, blocks a spill, and recalculates dependents. A picture cell
 * evaluates as `#VALUE!` in formulas. Nothing changes on failure.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`, or `picture->bytes` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `struct_size` is below the v1 size, `(row, col)` is outside the grid,
 *         or `alt_text` is `""`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `alt_text` is not well-formed UTF-8;
 *         `FM_STATUS_IO_IMAGE_UNSUPPORTED` (5023) when `len` is 0 or the bytes are not a PNG or JPEG image;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when `len` is over 32 MiB (the engine's own ceiling, not an
 *         Excel limit); plus the refusals of a cell value write.
 */
FM_API fm_status_t fm_sheet_set_cell_picture(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                             const fm_cell_picture_t* picture);

/**
 * @brief Inserts `bytes` as a picture and returns its object id, unique
 *        within the sheet's drawing.
 *
 * Existing charts, shapes and every other drawing content are kept. A sheet
 * without a drawing gets a new one.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `opts->struct_size` is below the
 *         first layout of `fm_drawing_placement_spec_t`, an enum
 *         field is outside its domain, the anchor kind is absolute, an
 *         offset or size is negative or 2^31 EMU or more, the anchor lies
 *         outside the sheet, or a two-cell `to` marker lies before `from`;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the image exceeds 32 MiB, or the
 *         package would exceed the reader's limits (context `part=zip_parts` or
 *         `part=zip_total_bytes`, then `requested=` and `limit=`);
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the drawing's id space is exhausted;
 *         `FM_STATUS_IO_IMAGE_UNSUPPORTED` (5023) when `bytes` is not a recognised image;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift. The workbook is then
 *         unchanged.
 */
FM_API fm_status_t fm_sheet_add_picture(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes, size_t len,
                                        const fm_drawing_placement_spec_t* opts, uint32_t* out_object_id);

/**
 * @brief Removes the drawing object `object_id` of any kind with what it
 *        reaches: a picture or chart with the relationships it names and the
 *        parts it reaches (media, or a chart part and its style, colors and
 *        embeddings) once nothing else references them; a shape, connector or
 *        group with its children; a table slicer with its frame, its cache
 *        when no other slicer uses it, and the sheet's drawing when it held
 *        nothing else.
 *
 * Removing a table or a sheet, or deleting the rows or columns under a
 * slicer's frame, removes the slicers bound to it the same way.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no object of the sheet's drawing holds
 *         that id;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the sheet is a chartsheet,
 *         whose drawing is not editable;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift.
 */
FM_API fm_status_t fm_sheet_remove_drawing(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id);

/**
 * @brief Merges `patch` into the drawing object `object_id` of any kind (the
 *        first one when ids repeat), keeping its id, crop, rotation, flips
 *        and every other element. The whole patch is checked before anything
 *        changes.
 *
 * An anchor change rewrites only the anchor markers, `editAs` when its value
 * changes, and the object's `a:off`/`a:ext` (a chart's `xdr:xfrm`). A
 * picture rotated into [45, 135) or [225, 315) degrees gets an anchor box
 * with width and height swapped about the same centre. A patch equal to the
 * current state writes nothing.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `patch` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no object of the sheet's drawing holds
 *         that id;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR`
 *         hits a field other than `description`, `SET` carries a `NULL` or
 *         empty `name` or a `NULL` description, an enum field is outside its
 *         domain, a `SET` anchor kind is not one-cell or two-cell, a field
 *         the resulting anchor kind does not hold is `SET` without one it
 *         holds (`fm_drawing_anchor_patch_t`), an offset or size
 *         is negative or 2^31 EMU or more, the anchor lies outside the sheet,
 *         a two-cell `to` marker lies before `from`, a slicer's anchor is not
 *         two-cell, or the sheet is a chartsheet;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when `name` is `SET` on a slicer;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a `SET` string is not valid UTF-8;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift.
 */
FM_API fm_status_t fm_sheet_update_drawing(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                           const fm_drawing_patch_t* patch);

/**
 * @brief Moves the picture, chart, shape or slicer `object_id` so it becomes
 *        item `z_order_index` of the drawing's object list (0 is the back).
 *
 * The unit moved is the object's top-level element, a whole
 * `mc:AlternateContent` included; other objects keep their relative order
 * and ids. An unchanged order writes nothing.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no object of the sheet's drawing holds
 *         that id;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the id belongs to an object that
 *         cannot be reordered;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `z_order_index` is not less
 *         than the object count;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift.
 */
FM_API fm_status_t fm_sheet_set_drawing_z_order(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                size_t z_order_index);

/**
 * @brief Captures the picture, chart or shape `object_id` as an opaque, versioned
 *        byte string: its list position, its element, the relationships it
 *        names with their parts, and every part reachable from those parts'
 *        relationships (a chart's style, colors, user shapes, embeddings).
 *
 * On success `*out_bytes` is a heap-allocated buffer of `*out_len` bytes
 * that MUST be released with `fm_buffer_destroy`. Both are zeroed before
 * validation and stay zero on failure. Hand the bytes to
 * `fm_sheet_restore_drawing` to put the object back.
 *
 * Lifetime: owned (until fm_buffer_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no object of the sheet's drawing holds
 *         that id;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the id belongs to a slicer (its
 *         parts live outside the drawing, so it cannot be captured);
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift.
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
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the bytes
 *         are not a valid capture, `flags` has an unknown bit, the id is
 *         held by another kind of object;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the package would exceed the
 *         reader's limits (context `part=zip_parts` or `part=zip_total_bytes`);
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when a fresh id is needed and the
 *         drawing's id space is exhausted;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift. The workbook is then
 *         unchanged.
 */
FM_API fm_status_t fm_sheet_restore_drawing(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes, size_t len,
                                            uint32_t flags, uint32_t* out_object_id);

/* -------------------------------------------------------------------------- */
/* Drawing shapes                                                             */
/* -------------------------------------------------------------------------- */

/**
 * @brief Fill of a shape's or chart element's `spPr`, as the shape and
 *        chart reads report it and `fm_shape_spec_t` takes it.
 *
 * An all-zero record states no fill of its own: a shape then draws its style's fill.
 */
typedef struct {
  /** The solid fill colour, None, Rgb or Theme; kind NONE when there is no solid colour. */
  fm_color_t color;
  /** 1 when `transparency_pct` is set; 0 means the colour has no alpha. */
  int32_t has_transparency_pct;
  /** The fill transparency as a percentage (0..100); meaningful when `has_transparency_pct` is set. */
  double transparency_pct;
  /** `1` for an explicit no-fill. */
  int32_t none;
  /** 1 on a read when the stored paint cannot be stated exactly: a gradient, pattern or picture fill (the other fields
   *  are then zero), or a solid colour form the record cannot hold (`color` then carries the nearest rendered colour).
   *  It passes through a save untouched and must be 0 on input. */
  int32_t other;
} fm_chart_fill_t;

/**
 * @brief Outline (`a:ln`) of a shape's or chart element's `spPr`.
 *
 * An all-zero record states no outline of its own: a shape then draws its
 * style's outline.
 */
typedef struct {
  /** The outline colour, None, Rgb or Theme; kind NONE when there is no solid colour. */
  fm_color_t color;
  /** 1 when `transparency_pct` is set; 0 means the colour has no alpha. */
  int32_t has_transparency_pct;
  /** The outline transparency as a percentage (0..100); meaningful when `has_transparency_pct` is set. */
  double transparency_pct;
  /** 1 when `width_pt` is set; 0 means the outline states no width. */
  int32_t has_width_pt;
  /** The outline width in points; meaningful when `has_width_pt` is set. */
  double width_pt;
  /** 1 when `dash` is set; 0 means the outline states no dash style. */
  int32_t has_dash;
  /** The dash style of the outline, an `fm_chart_line_dash_t`; meaningful when `has_dash` is set. */
  int32_t dash;
  /** `1` for an explicit no-line. */
  int32_t none;
  /** 1 on a read when the outline fill cannot be stated exactly (a gradient, pattern or a colour form the record cannot
   *  hold, with `color` as the nearest rendering); it
   *  must be 0 on input. */
  int32_t other;
} fm_chart_line_t;

/**
 * @brief Fill patch. A fill whose members are all `CLEAR` states no fill of
 *        its own (a shape draws its style's).
 */
typedef struct {
  /** Patch operation for `color`: keep it, set it from `color`, or clear it. */
  fm_patch_op_t color_op;
  /** The solid colour to set, kind None, Rgb or Theme (`theme` below 12, `tint` in [-1, 1]); needs `color_op` set. */
  fm_color_t color;
  /** Patch operation for `transparency_pct`: keep it, set it from `transparency_pct`, or clear it. */
  fm_patch_op_t transparency_pct_op;
  /** The transparency percentage to set (0..100); needs a solid colour. */
  double transparency_pct;
  /** Patch operation for `none`: keep it, set it from `none`, or clear it. */
  fm_patch_op_t none_op;
  /** 1 to write an explicit no-fill (`a:noFill`), 0 to remove it; `CLEAR` resets it to 0. A set colour with `none` 1 is
   *  `FM_STATUS_INVALID_ARGUMENT` (2). */
  int32_t none;
} fm_chart_fill_patch_t;

/**
 * @brief Outline patch. An outline whose members are all `CLEAR` states no
 *        outline of its own (a shape draws its style's).
 */
typedef struct {
  /** Patch operation for `color`: keep it, set it from `color`, or clear it. */
  fm_patch_op_t color_op;
  /** The outline colour to set, kind None, Rgb or Theme (`theme` below 12, `tint` within [-1, 1]); needs `color_op`
   *  set. */
  fm_color_t color;
  /** Patch operation for `transparency_pct`: keep it, set it from `transparency_pct`, or clear it. */
  fm_patch_op_t transparency_pct_op;
  /** The transparency percentage to set (0..100); needs a solid colour. */
  double transparency_pct;
  /** Patch operation for `width_pt`: keep it, set it from `width_pt`, or clear it. */
  fm_patch_op_t width_pt_op;
  /** The outline width to set, in points. */
  double width_pt;
  /** Patch operation for `dash`: keep it, set it from `dash`, or clear it. */
  fm_patch_op_t dash_op;
  /** The dash style to set, an `fm_chart_line_dash_t`. */
  int32_t dash;
  /** Patch operation for `none`: keep it, set it from `none`, or clear it. */
  fm_patch_op_t none_op;
  /** 1 to write an explicit no-line, 0 to remove it; `CLEAR` resets it to 0. */
  int32_t none;
} fm_chart_line_patch_t;

/**
 * @brief The shape-specific fields of a shape, connector or group of a
 *        sheet's drawing; the record `fm_sheet_get_shape` reads. The anchor,
 *        name, description and visibility are read with
 *        `fm_sheet_get_drawing_object`.
 *
 * The strings are read-scratch-backed views.
 */
typedef struct {
  uint32_t struct_size;
  /** The object's id in the sheet's drawing. */
  uint32_t object_id;
  /** The object's kind: shape, connector or group, an `fm_drawing_object_kind_t`. */
  int32_t kind;
  /** The shape's own fill; all zero when the shape has no element of its own and draws its style's paint, `other`
   *  set for a paint the record cannot state. Its colour is None, Rgb or Theme (Indexed and Auto are rejected on
   *  write). */
  fm_chart_fill_t fill;
  /** The shape's own outline; all zero when the shape has no element of its own and draws its style's paint, `other`
   *  set for a paint the record cannot state. Its colour is None, Rgb or Theme (Indexed and Auto are rejected on
   *  write); `width_pt` is at most 1584 points (a larger, negative or non-finite value is
   *  `FM_STATUS_INVALID_ARGUMENT` (2)). */
  fm_chart_line_t line;
  /** Clockwise rotation in degrees; an input in [-3600, 3600] is stored normalised to [0, 360), and a value outside
   *  that range or not finite is `FM_STATUS_INVALID_ARGUMENT` (2). */
  double rotation_deg;
  /** `1` when the shape is a text box. */
  int32_t text_box;
  /** The `prst` name of the preset geometry, an ECMA-376 `ST_ShapeType` name; `NULL` for custom geometry or a group. */
  const char* preset_geometry;
  /** The shape's text, paragraphs joined with a newline. */
  const char* text;
} fm_shape_info_t;

/**
 * @brief A shape to add (`fm_sheet_add_shape`): the common `placement` and
 *        the shape fields of `fm_shape_info_t`, in its units and domains.
 *
 * Every field left zero takes Excel's default for a new shape.
 */
typedef struct {
  uint32_t struct_size;
  /** Name, description, visibility and anchor of the new shape; the default is a two-cell anchor and the name
   *  `Shape <id>`. */
  fm_drawing_placement_spec_t placement;
  /** The shape's own fill; all zero draws the shape style's. `other` set is `FM_STATUS_INVALID_ARGUMENT` (2). */
  fm_chart_fill_t fill;
  /** The shape's own outline; all zero draws the shape style's. `other` set is `FM_STATUS_INVALID_ARGUMENT` (2). */
  fm_chart_line_t line;
  /** Clockwise rotation in degrees; 0 for none. */
  double rotation_deg;
  /** 1 to make the shape a text box; 0 for a plain shape. */
  int32_t text_box;
  /** Name of the preset geometry; `NULL` or empty means rect. A name that is not an ECMA-376 `ST_ShapeType` name is
   *  `FM_STATUS_INVALID_ARGUMENT` (2). */
  const char* preset_geometry;
  /** The shape's text, paragraphs joined with a newline; `NULL` for none. */
  const char* text;
} fm_shape_spec_t;

/**
 * @brief Partial update of the shape fields of a shape, connector or group
 *        (`fm_sheet_update_shape`), following the header's patch rules
 *        (Patches, above). The fields are the writable ones of
 *        `fm_shape_info_t`, in its units and domains; the anchor, name,
 *        description and visibility are the `common` half of
 *        `fm_sheet_update_shape`.
 *
 * Apart from `fill` and `line`, no field can be cleared.
 */
typedef struct {
  uint32_t struct_size;
  /** Update merged into the shape's own fill; clearing every member removes it so the shape draws its style's. */
  fm_chart_fill_patch_t fill;
  /** Update merged into the shape's own outline; clearing every member removes it so the shape draws its style's. */
  fm_chart_line_patch_t line;
  /** Patch operation for `rotation_deg`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t rotation_deg_op;
  /** The new clockwise rotation in degrees. */
  double rotation_deg;
  /** Patch operation for `text_box`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t text_box_op;
  /** Whether the shape is a text box. */
  int32_t text_box;
  /** Patch operation for `preset_geometry`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t preset_geometry_op;
  /** The new preset geometry name. */
  const char* preset_geometry;
  /** Patch operation for `text`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t text_op;
  /** The new text, paragraphs joined with a newline; applies to a shape only and keeps the first paragraph's properties
   *  and the first run's formatting. */
  const char* text;
} fm_shape_patch_t;

/**
 * @brief Reads the shape, connector or group whose frame id is `object_id`.
 *        `*out_present` is `1` and `*out` filled when there is one, else `0`
 *        and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * Fields the model does not express read as `other` or empty and pass
 * through a save untouched. Shapes are enumerated with
 * `fm_sheet_get_drawing_object_at`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse.
 */
FM_API fm_status_t fm_sheet_get_shape(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                      int32_t* out_present, fm_shape_info_t* out);

/**
 * @brief Appends a shape on top of the sheet's drawing from `spec` (see
 *        `fm_shape_spec_t` for the defaults) and returns its id.
 *
 * A `preset_geometry` naming a line or connector geometry (`line`,
 * `straightConnector1`, `bentConnector*`, `curvedConnector*`) inserts a
 * connector, which holds no text. A sheet without a drawing gets one; existing
 * drawing content is kept.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `spec` or `out_object_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a `struct_size` below the first layout of
 *         `fm_shape_spec_t` or its `placement`, a bad enumeration, an absolute anchor kind,
 *         a `preset_geometry` that is not an ECMA-376 `ST_ShapeType` name, an offset or size
 *         outside [0, 2^31) EMU, a zero size where the anchor needs one, a cell outside
 *         the sheet, a two-cell `to` before `from`, `other` set on `fill` or `line`,
 *         an Indexed or Auto colour, a colour given with `none`, a transparency
 *         without a solid colour, a line width or dash out of range, or text with
 *         control characters;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a string is not valid UTF-8;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the drawing's id space is exhausted;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift.
 */
FM_API fm_status_t fm_sheet_add_shape(fm_workbook_t* wb, size_t sheet_index, const fm_shape_spec_t* spec,
                                      uint32_t* out_object_id);

/**
 * @brief Merges `common` (the fields every drawing object has, as
 *        `fm_sheet_update_drawing` applies them) and `shape` (the shape
 *        fields) into the shape, connector or group `object_id` as one edit,
 *        keeping the rest of its XML.
 *
 * Either patch may be `NULL`; both `NULL` changes nothing. Both halves are
 * checked before anything is written, so nothing changes on failure.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no object of the sheet's drawing holds that id;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when the id belongs to a picture, chart or
 *         slicer, an op is outside 0..2, a field other than a `fill` / `line` member or
 *         `common`'s `description` is `CLEAR`, `SET` carries a `NULL` string, a `SET`
 *         value fails a rule of `fm_sheet_add_shape`, `text` or `text_box` is `SET` on a
 *         connector or group, `preset_geometry` is `SET` on a group, or `common` fails a
 *         rule of `fm_sheet_update_drawing`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a `SET` string is not valid UTF-8;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift.
 */
FM_API fm_status_t fm_sheet_update_shape(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                         const fm_drawing_patch_t* common, const fm_shape_patch_t* shape);

/* -------------------------------------------------------------------------- */
/* Table slicers                                                              */
/* -------------------------------------------------------------------------- */

/** @brief `fm_slicer_info_t::sort_order`; zero is the OOXML default `ascending`. */
typedef enum {
  FM_SLICER_SORT_ORDER_ASCENDING = 0,
  FM_SLICER_SORT_ORDER_DESCENDING = 1,
} fm_slicer_sort_order_t;

/** @brief `fm_slicer_info_t::cross_filter`. */
typedef enum {
  FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_DATA_AT_TOP = 0,
  FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_NO_DATA = 1,
  FM_SLICER_CROSS_FILTER_NONE = 2,
} fm_slicer_cross_filter_t;

/**
 * @brief The slicer-specific fields of a table slicer; the record
 *        `fm_sheet_get_slicer` reads. The frame's anchor, name, description
 *        and visibility are read with `fm_sheet_get_drawing_object`.
 *
 * The selection itself lives in the table's AutoFilter, not here. The strings
 * are read-scratch-backed views.
 */
typedef struct {
  uint32_t struct_size;
  /** The id of the slicer's frame in the sheet's drawing. */
  uint32_t object_id;
  /** The number of columns the slicer lays its item buttons out in. */
  uint32_t column_count;
  /** The order the slicer lists its items in, an `fm_slicer_sort_order_t`. */
  int32_t sort_order;
  /** How the slicer cross-filters its items (items with data listed first, items without data shown in place, or no
   *  cross-filtering), an `fm_slicer_cross_filter_t`: shown first, shown as unavailable, or not cross-filtered. */
  int32_t cross_filter;
  /** The height of each item row in points (the file's EMU / 12700). */
  double row_height_pt;
  /** The caption shown in the slicer header. */
  const char* caption;
  /** The name of the slicer cache the slicer belongs to. */
  const char* cache_name;
  /** The name of the table the slicer filters. */
  const char* table_name;
  /** The name of the table column the slicer filters. */
  const char* column_name;
  /** The name of the slicer style; `NULL` without a style. */
  const char* style_name;
} fm_slicer_info_t;

/**
 * @brief A table slicer to add (`fm_sheet_add_slicer`): the common
 *        `placement` and the slicer fields of `fm_slicer_info_t`, in its
 *        units and domains.
 *
 * The frame is always two-cell (`placement.anchor.kind` `DEFAULT` or
 * `TWO_CELL`). Every field left zero takes Excel's default for a new slicer.
 */
typedef struct {
  uint32_t struct_size;
  /** The frame placement and common drawing properties of the slicer: its anchor, name, description and visibility;
   *  the name defaults to the column name (made unique). */
  fm_drawing_placement_spec_t placement;
  /** The number of columns the slicer lays its item buttons out in; 0 takes 1. */
  uint32_t column_count;
  /** The order the slicer lists its items in, an `fm_slicer_sort_order_t`; 0 is ascending. */
  int32_t sort_order;
  /** How the slicer cross-filters its items (items with data listed first, items without data shown in place, or no
   *  cross-filtering), an `fm_slicer_cross_filter_t`; 0 is the default cross filter. */
  int32_t cross_filter;
  /** The height of each item row in points; 0 takes Excel's default, and a negative or non-finite height is rejected.
   */
  double row_height_pt;
  /** The caption shown in the slicer header; `NULL` takes the column name. */
  const char* caption;
  /** The name of the slicer style; `NULL` for no style and an empty string is rejected. */
  const char* style_name;
} fm_slicer_spec_t;

/**
 * @brief Partial update of the slicer fields of a table slicer
 *        (`fm_sheet_update_slicer`), following the header's patch rules
 *        (Patches, above). The fields are the writable ones of
 *        `fm_slicer_info_t`, in its units and domains; the frame's anchor,
 *        description and visibility are the `common` half of
 *        `fm_sheet_update_slicer`.
 */
typedef struct {
  uint32_t struct_size;
  /** How `column_count` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t column_count_op;
  /** The number of columns the slicer lays its item buttons out in. Read when `column_count_op` is `FM_PATCH_OP_SET`.
   */
  uint32_t column_count;
  /** How `sort_order` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t sort_order_op;
  /** The order the slicer lists its items in. Read when `sort_order_op` is `FM_PATCH_OP_SET`. */
  int32_t sort_order;
  /** How `cross_filter` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t cross_filter_op;
  /** How the slicer cross-filters its items (items with data listed first, items without data shown in place, or no
   *  cross-filtering). Read when `cross_filter_op` is `FM_PATCH_OP_SET`. */
  int32_t cross_filter;
  /** How `row_height_pt` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t row_height_pt_op;
  /** The height of each item row in points, finite and not negative. Read when `row_height_pt_op` is `FM_PATCH_OP_SET`.
   */
  double row_height_pt;
  /** How `caption` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it; `FM_PATCH_OP_CLEAR` is
   *  rejected. */
  fm_patch_op_t caption_op;
  /** The caption shown in the slicer header. Read when `caption_op` is `FM_PATCH_OP_SET`. */
  const char* caption;
  /** How `style_name` is patched: `FM_PATCH_OP_KEEP` leaves it, `FM_PATCH_OP_SET` writes it, `FM_PATCH_OP_CLEAR`
   *  removes the style. */
  fm_patch_op_t style_name_op;
  /** The name of the slicer style to write, non-empty. Read when `style_name_op` is `FM_PATCH_OP_SET`. */
  const char* style_name;
} fm_slicer_patch_t;

/**
 * @brief Reads the slicer whose frame id is `object_id`, the id
 *        `fm_sheet_add_slicer` returned. `*out_present` is `1` and `*out`
 *        filled when there is one, else `0` and `*out` zeroed (still
 *        `FM_STATUS_OK` (0)). A slicer without a frame has no id and is not
 *        found. Slicers are enumerated with `fm_sheet_get_drawing_object_at`.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the sheet's drawing does not parse.
 */
FM_API fm_status_t fm_sheet_get_slicer(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                       int32_t* out_present, fm_slicer_info_t* out);

/**
 * @brief Adds a slicer on the sheet over column `column_index` (0-based) of
 *        table `table_index` and returns its frame id.
 *
 * Creates the slicer and cache parts, their relationships and content types,
 * the `Slicer_*` defined name and the drawing frame (with its fallback shape),
 * or extends the sheet's existing slicer part, from `spec` (see
 * `fm_slicer_spec_t` for the defaults). The table is `table_index` of sheet `table_sheet_index` (a sheet-local
 * index, as every table function takes); it may live on another sheet than the slicer.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `spec` or `out_object_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` or `table_sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `table_index` is past that sheet's tables
 *         or `column_index` past the table's columns;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a `struct_size` below the first layout of
 *         `fm_slicer_spec_t` or its `placement`, a column whose name yields no slicer
 *         cache name, an anchor kind other than
 *         default or two-cell, an anchor that lies outside the sheet or has no size, a bad
 *         enumeration, column count, row height or text with control characters;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a string is not valid UTF-8;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when the drawing's id space is exhausted;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift.
 */
FM_API fm_status_t fm_sheet_add_slicer(fm_workbook_t* wb, size_t sheet_index, size_t table_sheet_index,
                                       size_t table_index, size_t column_index, const fm_slicer_spec_t* spec,
                                       uint32_t* out_object_id);

/**
 * @brief Merges `common` (the frame's fields, as `fm_sheet_update_drawing`
 *        applies them to a slicer) and `slicer` (the slicer fields) into the
 *        slicer whose frame id is `object_id` as one edit, keeping the rest
 *        of its XML.
 *
 * Either patch may be `NULL`; both `NULL` changes nothing. Both halves are
 * checked before anything is written, so nothing changes on failure.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no slicer has that id;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits a
 *         field other than `style_name` or `common`'s `description`, `SET` carries a
 *         `NULL` string or an empty `style_name`, a `SET` value fails a rule of
 *         `fm_sheet_add_slicer`, or `common` fails a rule of `fm_sheet_update_drawing`;
 *         `FM_STATUS_PRECONDITION_FAILED` (9) when `common` sets `name`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when a `SET` string is not valid UTF-8;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift.
 */
FM_API fm_status_t fm_sheet_update_slicer(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                          const fm_drawing_patch_t* common, const fm_slicer_patch_t* slicer);

/* -------------------------------------------------------------------------- */
/* Sparkline groups                                                           */
/* -------------------------------------------------------------------------- */

/** @brief `fm_sparkline_group_t::type`. */
typedef enum {
  FM_SPARKLINE_TYPE_LINE = 0,
  FM_SPARKLINE_TYPE_COLUMN = 1,
  FM_SPARKLINE_TYPE_STACKED = 2,
} fm_sparkline_type_t;

/** @brief `fm_sparkline_group_t::min_axis_type` and `max_axis_type`. */
typedef enum {
  FM_SPARKLINE_AXIS_TYPE_INDIVIDUAL = 0,
  FM_SPARKLINE_AXIS_TYPE_GROUP = 1,
  FM_SPARKLINE_AXIS_TYPE_CUSTOM = 2,
} fm_sparkline_axis_type_t;

/**
 * @brief How blank cells are drawn: a chart's `c:dispBlanksAs`
 *        (`fm_chart_data_t::blanks_as`) and a sparkline group's
 *        `displayEmptyCellsAs` (`fm_sparkline_group_t::empty_cells_as`).
 *
 * Zero is the OOXML default `zero`. `ABSENT` is a chart read whose element
 * is not written; a sparkline group always states the value and rejects it.
 */
typedef enum {
  FM_BLANKS_AS_ABSENT = -1,
  FM_BLANKS_AS_ZERO = 0,
  FM_BLANKS_AS_GAP = 1,
  FM_BLANKS_AS_SPAN = 2,
} fm_blanks_as_t;

/**
 * @brief One sparkline: the data `formula` (`xm:f`, a sheet-qualified range)
 *        and the single-cell `location` (`xm:sqref`, A1 text) that shows it.
 *        Read-scratch-backed views on read.
 */
typedef struct {
  /** The data range as sheet-qualified text, read with a leading `=`; input takes it with or without one. */
  const char* formula;
  /** The single cell that shows the sparkline. */
  const char* location;
} fm_sparkline_t;

/** @brief A sparkline group (`x14:sparklineGroup`) of the sheet's worksheet extension list. */
typedef struct {
  uint32_t struct_size;
  /** The group's engine id: an opaque string, unique on its sheet, stable while the workbook is open across structural
   *  edits and never reused within the session; assigned again on load or sheet copy and not saved. A
   *  read-scratch-backed view on read; ignored on the add path, which assigns a new id. */
  const char* id;
  /** Colour of the line or columns; kind `NONE` when the group states none. */
  fm_color_t color_series;
  /** Colour of negative points; kind `NONE` when the group states none. */
  fm_color_t color_negative;
  /** Colour of the horizontal axis; kind `NONE` when the group states none. */
  fm_color_t color_axis;
  /** Colour of the markers; kind `NONE` when the group states none. */
  fm_color_t color_markers;
  /** Colour of the first point; kind `NONE` when the group states none. */
  fm_color_t color_first;
  /** Colour of the last point; kind `NONE` when the group states none. */
  fm_color_t color_last;
  /** Colour of the highest point; kind `NONE` when the group states none. */
  fm_color_t color_high;
  /** Colour of the lowest point; kind `NONE` when the group states none. */
  fm_color_t color_low;
  /** The custom vertical axis minimum; read only when `has_manual_min` is 1. */
  double manual_min;
  /** The custom vertical axis maximum; read only when `has_manual_max` is 1. */
  double manual_max;
  /** Line width in points; 0.75 for a new group. */
  double line_width_pt;
  /** The sparkline type: line, column or stacked (win/loss), an `fm_sparkline_type_t`. */
  int32_t type;
  /** `1` to show a marker at every point of a line sparkline. */
  int32_t markers;
  /** `1` to highlight the highest point. */
  int32_t high;
  /** `1` to highlight the lowest point. */
  int32_t low;
  /** `1` to highlight the first point. */
  int32_t first;
  /** `1` to highlight the last point. */
  int32_t last;
  /** `1` to highlight negative points. */
  int32_t negative;
  /** `1` to show the horizontal axis. */
  int32_t display_x_axis;
  /** `1` to plot data in hidden rows and columns. */
  int32_t display_hidden;
  /** `1` to plot the group right to left; the default is `0`. */
  int32_t right_to_left;
  /** How the vertical axis minimum is chosen: per sparkline, shared by the group, or the custom value, an
   *  `fm_sparkline_axis_type_t`. */
  int32_t min_axis_type;
  /** How the vertical axis maximum is chosen: per sparkline, shared by the group, or the custom value, an
   *  `fm_sparkline_axis_type_t`. */
  int32_t max_axis_type;
  /** 1 when `manual_min` is set; 0 means the axis minimum is not set. */
  int32_t has_manual_min;
  /** 1 when `manual_max` is set; 0 means the axis maximum is not set. */
  int32_t has_manual_max;
  /** `1` to space points by a date axis. */
  int32_t date_axis;
  /** How empty cells are plotted: as a gap, as zero, or by connecting the neighbours, an `fm_blanks_as_t` other than
   *  `FM_BLANKS_AS_ABSENT`. */
  int32_t empty_cells_as;
  /** The group's date range as sheet-qualified text, read with a leading `=`; `NULL` when absent. Input takes it with
   *  or without one, and `NULL` or empty writes none. */
  const char* date_axis_ref;
  /** The number of sparklines in the group; filled on read and ignored on input. */
  uint32_t sparkline_count;
} fm_sparkline_group_t;

/**
 * @brief Partial update of a sparkline group
 *        (`fm_sheet_update_sparkline_group`), following the header's patch
 *        rules (Patches, above). The fields are the writable ones of
 *        `fm_sparkline_group_t` but `id`, in its units and domains, and the
 *        group's sparklines.
 *
 * The eight colours, `manual_min`, `manual_max` and `date_axis_ref` can be
 * cleared, which removes the colour (it reads back `FM_COLOR_KIND_NONE`),
 * the attribute or the date range. `items_op` `SET` replaces the
 * group's sparklines with the `items_count` entries of `items` (at least
 * one, under the rules of `fm_sheet_add_sparkline_group`); `KEEP` leaves
 * them. Every attribute or child outside the model is kept.
 */
typedef struct {
  uint32_t struct_size;
  /** Patch operation for `color_series`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t color_series_op;
  /** Colour of the line or columns. Applies when `color_series_op` sets it. */
  fm_color_t color_series;
  /** Patch operation for `color_negative`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t color_negative_op;
  /** Colour of negative points. Applies when `color_negative_op` sets it. */
  fm_color_t color_negative;
  /** Patch operation for `color_axis`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t color_axis_op;
  /** Colour of the horizontal axis. Applies when `color_axis_op` sets it. */
  fm_color_t color_axis;
  /** Patch operation for `color_markers`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t color_markers_op;
  /** Colour of the markers. Applies when `color_markers_op` sets it. */
  fm_color_t color_markers;
  /** Patch operation for `color_first`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t color_first_op;
  /** Colour of the first point. Applies when `color_first_op` sets it. */
  fm_color_t color_first;
  /** Patch operation for `color_last`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t color_last_op;
  /** Colour of the last point. Applies when `color_last_op` sets it. */
  fm_color_t color_last;
  /** Patch operation for `color_high`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t color_high_op;
  /** Colour of the highest point. Applies when `color_high_op` sets it. */
  fm_color_t color_high;
  /** Patch operation for `color_low`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t color_low_op;
  /** Colour of the lowest point. Applies when `color_low_op` sets it. */
  fm_color_t color_low;
  /** Patch operation for `manual_min`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t manual_min_op;
  /** The custom vertical axis minimum. Applies when `manual_min_op` sets it. */
  double manual_min;
  /** Patch operation for `manual_max`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t manual_max_op;
  /** The custom vertical axis maximum. Applies when `manual_max_op` sets it. */
  double manual_max;
  /** Patch operation for `line_width_pt`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t line_width_pt_op;
  /** Line width in points. Applies when `line_width_pt_op` sets it. */
  double line_width_pt;
  /** Patch operation for `type`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t type_op;
  /** The sparkline type: line, column or stacked (win/loss). Applies when `type_op` sets it. */
  int32_t type;
  /** Patch operation for `markers`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t markers_op;
  /** `1` to show a marker at every point of a line sparkline. Applies when `markers_op` sets it. */
  int32_t markers;
  /** Patch operation for `high`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t high_op;
  /** `1` to highlight the highest point. Applies when `high_op` sets it. */
  int32_t high;
  /** Patch operation for `low`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t low_op;
  /** `1` to highlight the lowest point. Applies when `low_op` sets it. */
  int32_t low;
  /** Patch operation for `first`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t first_op;
  /** `1` to highlight the first point. Applies when `first_op` sets it. */
  int32_t first;
  /** Patch operation for `last`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t last_op;
  /** `1` to highlight the last point. Applies when `last_op` sets it. */
  int32_t last;
  /** Patch operation for `negative`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t negative_op;
  /** `1` to highlight negative points. Applies when `negative_op` sets it. */
  int32_t negative;
  /** Patch operation for `display_x_axis`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t display_x_axis_op;
  /** `1` to show the horizontal axis. Applies when `display_x_axis_op` sets it. */
  int32_t display_x_axis;
  /** Patch operation for `display_hidden`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t display_hidden_op;
  /** `1` to plot data in hidden rows and columns. Applies when `display_hidden_op` sets it. */
  int32_t display_hidden;
  /** Patch operation for `right_to_left`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t right_to_left_op;
  /** `1` to plot the group right to left. Applies when `right_to_left_op` sets it. */
  int32_t right_to_left;
  /** Patch operation for `min_axis_type`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t min_axis_type_op;
  /** How the vertical axis minimum is chosen: per sparkline, shared by the group, or the custom value. Applies when
   *  `min_axis_type_op` sets it. */
  int32_t min_axis_type;
  /** Patch operation for `max_axis_type`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t max_axis_type_op;
  /** How the vertical axis maximum is chosen: per sparkline, shared by the group, or the custom value. Applies when
   *  `max_axis_type_op` sets it. */
  int32_t max_axis_type;
  /** Patch operation for `date_axis`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t date_axis_op;
  /** `1` to space points by a date axis. Applies when `date_axis_op` sets it. */
  int32_t date_axis;
  /** Patch operation for `empty_cells_as`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t empty_cells_as_op;
  /** How empty cells are plotted: as a gap, as zero, or by connecting the neighbours. Applies when `empty_cells_as_op`
   *  sets it. */
  int32_t empty_cells_as;
  /** Patch operation for `date_axis_ref`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t date_axis_ref_op;
  /** The new date range as sheet-qualified text, with or without a leading `=`. Applies when `date_axis_ref_op` sets
   *  it. */
  const char* date_axis_ref;
  /** Patch operation for `items`: leave it unchanged, set it, or (where it can be cleared) remove it. */
  fm_patch_op_t items_op;
  /** The new sparklines of the group (at least one). */
  const fm_sparkline_t* items;
  /** The number of entries in `items`. */
  uint32_t items_count;
} fm_sparkline_group_patch_t;

/**
 * @brief Number of sparkline groups in the sheet's worksheet extension list.
 *
 * A group is addressed by its id (`fm_sparkline_group_t::id`); positions, in
 * document order, serve enumeration only, and adding or removing a group
 * shifts the later ones. Groups of a workbook loaded from `.xlsb` are listed
 * and editable the same way; a group the reader cannot decode stays a binary
 * record and is not listed.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_count == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_sparkline_group_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);

/**
 * @brief Reads sparkline group `group_index`; `id` and `date_axis_ref` are
 *        read-scratch-backed views (`date_axis_ref` `NULL` when absent).
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `group_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_sparkline_group_at(const fm_workbook_t* wb, size_t sheet_index, size_t group_index,
                                                   fm_sparkline_group_t* out);

/**
 * @brief Reads the sparkline group whose id is `group_id`. `*out_present` is
 *        `1` and `*out` filled when there is one, else `0` and `*out` zeroed
 *        (still `FM_STATUS_OK` (0)). The views in `out` have the lifetime
 *        `fm_sheet_get_sparkline_group_at` describes.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if any pointer argument is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `group_id` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_get_sparkline_group(const fm_workbook_t* wb, size_t sheet_index, const char* group_id,
                                                int32_t* out_present, fm_sparkline_group_t* out);

/**
 * @brief Reads sparkline `sparkline_index` of group `group_index`; both strings are
 *        read-scratch-backed views.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `group_index` or `sparkline_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_sparkline_at(const fm_workbook_t* wb, size_t sheet_index, size_t group_index,
                                             size_t sparkline_index, fm_sparkline_t* out);

/**
 * @brief Appends a sparkline group with `count` sparklines and writes its new
 *        id to `*out_group_id`, a read-scratch-backed view valid until the
 *        next scratch-backed read. `group->id` is ignored.
 *
 * Attributes and children outside the model are not written. A sheet without
 * a worksheet extension list gets one; other extensions are kept. A sparkline
 * formula or `date_axis_ref` names one cell or rectangle (on any sheet, a 3-D
 * span or another workbook) or a defined name; `#REF!` is stored without a
 * formula and reads back empty.
 *
 * Lifetime: transient.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `group`, `items` or `out_group_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for `count == 0`, a sparkline whose
 *         location is not one A1 cell, a formula or
 *         `date_axis_ref` that is not one reference or defined name (a union,
 *         a structured reference, an expression, a sheet-qualified `#REF!`;
 *         context `formula=<text>` or `date_axis_ref=<text>`), a non-finite
 *         number, a negative line weight, or an enumeration out of range. An empty (or `NULL`) formula is not
 *         refused: it is stored as a deleted (`#REF!`) source;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when `count` is above 65536
 *         (context `part=sparkline_items requested=<n> limit=65536`).
 */
FM_API fm_status_t fm_sheet_add_sparkline_group(fm_workbook_t* wb, size_t sheet_index,
                                                const fm_sparkline_group_t* group, const fm_sparkline_t* items,
                                                size_t count, const char** out_group_id);

/**
 * @brief Merges `patch` into the sparkline group whose id is `group_id`. The
 *        whole patch is checked before anything changes.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb`, `group_id` or `patch` is `NULL`, or a `SET`
 *           `items` is `NULL` while `items_count > 0`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no group has `group_id`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when an op is outside 0..2, `CLEAR` hits a
 *           field other than the eight colours, `manual_min`, `manual_max` and `date_axis_ref`,
 *           `SET` carries a `NULL` string, a `SET` `items` is empty, or a `SET` value fails
 *           a rule of `fm_sheet_add_sparkline_group`;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when a `SET` `items` holds more than 65536 sparklines;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `group_id` or a `SET` string is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_update_sparkline_group(fm_workbook_t* wb, size_t sheet_index, const char* group_id,
                                                   const fm_sparkline_group_patch_t* patch);

/**
 * @brief Removes the sparkline group whose id is `group_id`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `wb` or `group_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no group has `group_id`;
 *         `FM_STATUS_BINDING_UTF8ENCODING_ERROR` (7002) when `group_id` is not valid UTF-8.
 */
FM_API fm_status_t fm_sheet_remove_sparkline_group(fm_workbook_t* wb, size_t sheet_index, const char* group_id);

/* -------------------------------------------------------------------------- */
/* Charts                                                                     */
/* -------------------------------------------------------------------------- */

/*
 * A chart is addressed as `(sheet_index, object_id)`, the `cNvPr` id
 * `fm_sheet_get_drawing_object_at` reports with `FM_DRAWING_OBJECT_KIND_CHART`; the
 * drawing operations above move, reorder, capture, restore and remove it.
 *
 * Reads go through an owned handle (`fm_chart_t`, `fm_chart_data_t`) that
 * snapshots the chart when opened; every string it hands out is UTF-8 owned
 * by the handle and stays valid until the handle is destroyed. Absent values
 * follow the header's absence rules (Absence, above): an optional number has
 * a `has_*` flag, an optional string is `NULL`.
 *
 * Patch structs follow the header's patch rules (Patches, above); `CLEAR`
 * removes the element. Colors are `fm_color_t`; transparency is a percentage
 * (0..100).
 *
 * Edits patch the chart part and write nothing when the result is
 * unchanged. Common failures, besides those each entry lists:
 * `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) for an out-of-range sheet;
 * `FM_STATUS_NOT_FOUND` (6) for an unknown chart id, series `c:idx` or axis
 * id; `FM_STATUS_INDEX_OUT_OF_RANGE` (13) for a `plot_index` past the last plot
 * group; `FM_STATUS_INVALID_ARGUMENT` (2) for a value outside its domain (an
 * enum ordinal, an op outside 0..2, `CLEAR` on a field that cannot be
 * removed); `FM_STATUS_IO_CHART_UNPARSEABLE` (5025)
 * when the chart part does not parse; `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for
 * an edit to an extended chart; `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) for an edit to
 * a chart on an `.xlsb` drawing the reader could not lift. A failed edit
 * leaves the workbook unchanged.
 */

/** @brief Plot group type. Mirrors `formulon::ChartPlotType`. */
typedef enum {
  FM_CHART_PLOT_TYPE_BAR = 0,
  FM_CHART_PLOT_TYPE_BAR_3D = 1,
  FM_CHART_PLOT_TYPE_LINE = 2,
  FM_CHART_PLOT_TYPE_LINE_3D = 3,
  FM_CHART_PLOT_TYPE_AREA = 4,
  FM_CHART_PLOT_TYPE_AREA_3D = 5,
  FM_CHART_PLOT_TYPE_PIE = 6,
  FM_CHART_PLOT_TYPE_PIE_3D = 7,
  FM_CHART_PLOT_TYPE_DOUGHNUT = 8,
  FM_CHART_PLOT_TYPE_OF_PIE = 9,
  FM_CHART_PLOT_TYPE_SCATTER = 10,
  FM_CHART_PLOT_TYPE_BUBBLE = 11,
  FM_CHART_PLOT_TYPE_RADAR = 12,
  FM_CHART_PLOT_TYPE_STOCK = 13,
  FM_CHART_PLOT_TYPE_SURFACE = 14,
  FM_CHART_PLOT_TYPE_SURFACE_3D = 15,
  FM_CHART_PLOT_TYPE_EXTENDED = 16,
} fm_chart_plot_type_t;

/**
 * @brief Layout of an extended chart (`FM_CHART_PLOT_TYPE_EXTENDED`, the
 *        `cx:` chart part); `fm_chart_plot_info_t::extended_layout` reads
 *        `ABSENT` for any other plot.
 */
typedef enum {
  FM_CHART_EXTENDED_LAYOUT_ABSENT = -1,
  FM_CHART_EXTENDED_LAYOUT_BOX_WHISKER = 0,
  FM_CHART_EXTENDED_LAYOUT_CLUSTERED_COLUMN = 1,
  FM_CHART_EXTENDED_LAYOUT_FUNNEL = 2,
  FM_CHART_EXTENDED_LAYOUT_PARETO_LINE = 3,
  FM_CHART_EXTENDED_LAYOUT_REGION_MAP = 4,
  FM_CHART_EXTENDED_LAYOUT_SUNBURST = 5,
  FM_CHART_EXTENDED_LAYOUT_TREEMAP = 6,
  FM_CHART_EXTENDED_LAYOUT_WATERFALL = 7,
} fm_chart_extended_layout_t;

typedef enum {
  FM_CHART_GROUPING_CLUSTERED = 0,
  FM_CHART_GROUPING_STACKED = 1,
  FM_CHART_GROUPING_PERCENT_STACKED = 2,
  FM_CHART_GROUPING_STANDARD = 3,
} fm_chart_grouping_t;

/** @brief How a resolved point is drawn. */
typedef enum {
  FM_CHART_POINT_STATE_VALUE = 0,
  FM_CHART_POINT_STATE_GAP = 1,
  FM_CHART_POINT_STATE_CONNECT = 2,
} fm_chart_point_state_t;

typedef enum {
  FM_CHART_AXIS_KIND_CATEGORY = 0,
  FM_CHART_AXIS_KIND_VALUE = 1,
  FM_CHART_AXIS_KIND_DATE = 2,
  FM_CHART_AXIS_KIND_SERIES = 3,
} fm_chart_axis_kind_t;

typedef enum {
  FM_CHART_AXIS_POSITION_BOTTOM = 0,
  FM_CHART_AXIS_POSITION_TOP = 1,
  FM_CHART_AXIS_POSITION_LEFT = 2,
  FM_CHART_AXIS_POSITION_RIGHT = 3,
} fm_chart_axis_position_t;

/**
 * @brief Where an axis draws its tick labels (`c:tickLblPos`). Zero is the
 *        OOXML default `nextTo`; the others follow `ST_TickLblPos`. `NONE`
 *        draws no labels; `ABSENT` is a read whose element is not written.
 */
typedef enum {
  FM_CHART_TICK_LABEL_POSITION_ABSENT = -1,
  FM_CHART_TICK_LABEL_POSITION_NEXT_TO = 0,
  FM_CHART_TICK_LABEL_POSITION_HIGH = 1,
  FM_CHART_TICK_LABEL_POSITION_LOW = 2,
  FM_CHART_TICK_LABEL_POSITION_NONE = 3,
} fm_chart_tick_label_position_t;

/** @brief Where an axis crosses its cross axis; `AT` uses `crosses_at`. */
typedef enum {
  FM_CHART_AXIS_CROSSES_ABSENT = -1,
  FM_CHART_AXIS_CROSSES_AUTO_ZERO = 0,
  FM_CHART_AXIS_CROSSES_MIN = 1,
  FM_CHART_AXIS_CROSSES_MAX = 2,
  FM_CHART_AXIS_CROSSES_AT = 3,
} fm_chart_axis_crosses_t;

/** @brief `AUTO` (title element without text) lets the viewer supply its default. */
typedef enum {
  FM_CHART_TITLE_KIND_NONE = 0,
  FM_CHART_TITLE_KIND_TEXT = 1,
  FM_CHART_TITLE_KIND_REF = 2,
  FM_CHART_TITLE_KIND_AUTO = 3,
} fm_chart_title_kind_t;

typedef enum {
  FM_CHART_SERIES_DIRECTION_AUTO = 0,
  FM_CHART_SERIES_DIRECTION_COLUMNS = 1,
  FM_CHART_SERIES_DIRECTION_ROWS = 2,
} fm_chart_series_direction_t;

/**
 * @brief Chart legend position (`c:legendPos`). Zero is the OOXML default
 *        `r`; the others follow `ST_LegendPos`. `NONE` is a chart without a
 *        legend.
 */
typedef enum {
  FM_CHART_LEGEND_POSITION_ABSENT = -1,
  FM_CHART_LEGEND_POSITION_RIGHT = 0,
  FM_CHART_LEGEND_POSITION_BOTTOM = 1,
  FM_CHART_LEGEND_POSITION_TOP_RIGHT = 2,
  FM_CHART_LEGEND_POSITION_LEFT = 3,
  FM_CHART_LEGEND_POSITION_TOP = 4,
} fm_chart_legend_position_t;

typedef enum {
  FM_CHART_LINE_DASH_SOLID = 0,
  FM_CHART_LINE_DASH_DASH = 1,
  FM_CHART_LINE_DASH_DOT = 2,
  FM_CHART_LINE_DASH_DASH_DOT = 3,
  FM_CHART_LINE_DASH_OTHER = 4,
} fm_chart_line_dash_t;

typedef enum {
  FM_CHART_MARKER_SYMBOL_NONE = 0,
  FM_CHART_MARKER_SYMBOL_AUTO = 1,
  FM_CHART_MARKER_SYMBOL_CIRCLE = 2,
  FM_CHART_MARKER_SYMBOL_SQUARE = 3,
  FM_CHART_MARKER_SYMBOL_DIAMOND = 4,
  FM_CHART_MARKER_SYMBOL_TRIANGLE = 5,
  FM_CHART_MARKER_SYMBOL_X = 6,
  FM_CHART_MARKER_SYMBOL_STAR = 7,
  FM_CHART_MARKER_SYMBOL_PLUS = 8,
  FM_CHART_MARKER_SYMBOL_DOT = 9,
  FM_CHART_MARKER_SYMBOL_DASH = 10,
  FM_CHART_MARKER_SYMBOL_PICTURE = 11,
} fm_chart_marker_symbol_t;

typedef enum {
  FM_CHART_LABEL_POSITION_BEST_FIT = 0,
  FM_CHART_LABEL_POSITION_CENTER = 1,
  FM_CHART_LABEL_POSITION_INSIDE_BASE = 2,
  FM_CHART_LABEL_POSITION_INSIDE_END = 3,
  FM_CHART_LABEL_POSITION_OUTSIDE_END = 4,
  FM_CHART_LABEL_POSITION_ABOVE = 5,
  FM_CHART_LABEL_POSITION_BELOW = 6,
  FM_CHART_LABEL_POSITION_LEFT = 7,
  FM_CHART_LABEL_POSITION_RIGHT = 8,
} fm_chart_label_position_t;

/** @brief `c:barDir` of a bar plot. */
typedef enum {
  FM_CHART_BAR_DIRECTION_ABSENT = -1,
  FM_CHART_BAR_DIRECTION_COLUMN = 0,
  FM_CHART_BAR_DIRECTION_BAR = 1,
} fm_chart_bar_direction_t;

/**
 * @brief Which input of a chart a data source, resolved range or unresolved
 *        input belongs to. The six series roles (`NAME` .. `SIZE`) index the
 *        per-series arrays such as `range_count`; `TITLE` and `AXIS_TITLE` occur
 *        only in `fm_chart_unresolved_info_t.role`.
 */
typedef enum {
  FM_CHART_ROLE_NAME = 0,
  FM_CHART_ROLE_CATEGORIES = 1,
  FM_CHART_ROLE_VALUES = 2,
  FM_CHART_ROLE_X = 3,
  FM_CHART_ROLE_Y = 4,
  FM_CHART_ROLE_SIZE = 5,
  FM_CHART_ROLE_TITLE = 6,
  FM_CHART_ROLE_AXIS_TITLE = 7,
} fm_chart_role_t;

/* ---- Chart model reads --------------------------------------------------- */

/** @brief Opaque snapshot of a chart's typed view. Release with `fm_chart_destroy`. */
typedef struct fm_chart fm_chart_t;

/** @brief A chart or axis title. As input, `NULL` and `""` both read as empty. */
typedef struct {
  /** Whether the title is absent, automatic, literal text or a cell reference, an `fm_chart_title_kind_t`. */
  int32_t kind;
  /** The flattened rich text of a text title; `NULL` for another kind. */
  const char* text;
  /** The reference of a reference title; `NULL` for another kind. */
  const char* formula;
} fm_chart_title_t;

/** @brief Marker; `symbol` is an `fm_chart_marker_symbol_t`, `size_pt` the size in points. */
typedef struct {
  /** The marker symbol. */
  int32_t symbol;
  /** 1 when `size_pt` is set; 0 means the marker states no size. */
  int32_t has_size_pt;
  /** The marker size in points; meaningful when `has_size_pt` is set. */
  uint32_t size_pt;
  /** The marker fill. */
  fm_chart_fill_t fill;
  /** The marker outline. */
  fm_chart_line_t line;
} fm_chart_marker_t;

/** @brief Data labels of a series. */
typedef struct {
  /** 1 when the label shows the point value. */
  int32_t show_value;
  /** 1 when the label shows the category name. */
  int32_t show_category;
  /** 1 when the label shows the series name. */
  int32_t show_series;
  /** 1 when the label shows the percentage of the whole (pie-type plots). */
  int32_t show_percent;
  /** 1 when the label shows the legend key. */
  int32_t show_legend_key;
  /** The text placed between the label parts; `NULL` when absent. */
  const char* separator;
  /** 1 when `position` is set; 0 means the chart states no label position. */
  int32_t has_position;
  /** Where the label sits relative to its point, an `fm_chart_label_position_t`; meaningful when `has_position` is
   *  set. */
  int32_t position;
  /** The number format code of the label; `NULL` when absent. */
  const char* number_format;
  /** 1 when the label number format follows the source cells. */
  int32_t source_linked;
  /** The label font colour, None, Rgb or Theme; kind NONE when absent. */
  fm_color_t font_color;
  /** 1 when `font_size_pt` is set; 0 means the chart states no font size. */
  int32_t has_font_size_pt;
  /** The label font size in points; meaningful when `has_font_size_pt` is set. */
  double font_size_pt;
  /** 1 when the labels are deleted (c:delete) for the series. */
  int32_t deleted;
} fm_chart_labels_t;

/**
 * @brief Top-level fields of a chart.
 *
 * The anchor, name and description are read with
 * `fm_sheet_get_drawing_object`.
 */
typedef struct {
  uint32_t struct_size;
  /** The id of the chart's drawing object on its sheet. */
  uint32_t object_id;
  /** `1` for an extended (`cx:`) chart. */
  int32_t is_extended;
  /** `1` when the chart is bound to a pivot table. */
  int32_t is_pivot_chart;
  /** `1` when the chart lives on a chartsheet rather than as an embedded object. */
  int32_t on_chartsheet;
  /** The chart title; kind NONE when the chart has none. */
  fm_chart_title_t title;
  /** Whether the automatic title is deleted, a tri-state value (default when the chart states none). */
  int32_t auto_title_deleted; /* fm_tristate_t */
  /** The legend position; `FM_CHART_LEGEND_POSITION_ABSENT` when the chart has no legend. */
  int32_t legend_position;
  /** 1 when the legend overlays the plot area; 0 when it does not or there is no legend. */
  int32_t legend_overlay;
  /** The number of series whose legend entry is hidden. */
  uint32_t legend_hidden_count;
  /** How empty cells are plotted (gap, zero or span), an `fm_blanks_as_t`; the absent value when the file omits it. */
  int32_t blanks_as;
  /** Whether only visible cells are plotted, a tri-state value (default when the chart states none). */
  int32_t plot_visible_only; /* fm_tristate_t */
  /** `1` when the chart uses the 1904 date system. */
  int32_t date1904;
  /** 1 when `style` is set; 0 means the chart states no style. */
  int32_t has_style;
  /** The chart style number (1..48); meaningful when `has_style` is set. */
  int32_t style;
  /** The number of plot groups of the chart. */
  uint32_t plot_count;
  /** The number of axes of the chart. */
  uint32_t axis_count;
  /** The number of categories the chart filter hides. */
  uint32_t hidden_category_count;
  /** The sheet-qualified rectangle every series was split from; `NULL` (with the other `source_*` fields zero) when the
   *  series share none. */
  const char* source_range;
  /** Whether the series run along rows or columns of `source_range`. */
  int32_t source_direction;
  /** 1 when the first row of `source_range` holds names rather than data. */
  int32_t source_header_row;
  /** 1 when the first column of `source_range` holds names rather than data. */
  int32_t source_header_column;
  /** Whether `#N/A` points draw as blank cells (`c16r3:dispNaAsBlank`), a tri-state value (default when the chart
   *  carries no setting). */
  int32_t na_as_blank; /* fm_tristate_t */
} fm_chart_info_t;

/** @brief One plot group. */
typedef struct {
  uint32_t struct_size;
  /** The plot type of the group, an `fm_chart_plot_type_t`. */
  int32_t type;
  /** The series layout of an extended plot, an `fm_chart_extended_layout_t`; `FM_CHART_EXTENDED_LAYOUT_ABSENT` for any
   *  other plot type. */
  int32_t extended_layout;
  /** Whether bars run vertically (column) or horizontally, an `fm_chart_bar_direction_t`; 0 when the plot is not in the
   *  bar family. */
  int32_t bar_direction;
  /** 1 when `grouping` is set; 0 means the plot type has no grouping. */
  int32_t has_grouping;
  /** How the series are stacked or grouped; meaningful when `has_grouping` is set. */
  int32_t grouping;
  /** Whether each point of a single-series plot takes its own colour, a tri-state value. */
  int32_t vary_colors; /* fm_tristate_t */
  /** 1 when `gap_width_pct` is set; 0 means the plot type has none. */
  int32_t has_gap_width_pct;
  /** The gap between bar clusters as a percentage of the bar width; meaningful when `has_gap_width_pct` is set. */
  uint32_t gap_width_pct;
  /** 1 when `overlap_pct` is set; 0 means the plot type has none. */
  int32_t has_overlap_pct;
  /** The overlap between bars of a cluster as a percentage of the bar width; meaningful when `has_overlap_pct` is set.
   */
  int32_t overlap_pct;
  /** 1 when `hole_size_pct` is set; 0 means the plot type has none. */
  int32_t has_hole_size_pct;
  /** The doughnut hole size as a percentage of the doughnut diameter; meaningful when `has_hole_size_pct` is set. */
  uint32_t hole_size_pct;
  /** 1 when `first_slice_angle_deg` is set; 0 means the plot type has none. */
  int32_t has_first_slice_angle_deg;
  /** The angle in degrees of the first pie or doughnut slice; meaningful when `has_first_slice_angle_deg` is set. */
  uint32_t first_slice_angle_deg;
  /** The number of axis ids of the plot group. */
  uint32_t axis_id_count;
  /** The number of series of the plot group. */
  uint32_t series_count;
} fm_chart_plot_info_t;

/** @brief Kind of a series input. */
typedef enum {
  FM_CHART_SOURCE_KIND_NONE = 0,
  FM_CHART_SOURCE_KIND_REF = 1,
  FM_CHART_SOURCE_KIND_LITERAL = 2,
} fm_chart_source_kind_t;

/** @brief One series input. */
typedef struct {
  /** Whether the input is empty, a range reference or literal points, an `fm_chart_source_kind_t`. */
  int32_t kind;
  /** The reference text of a range-reference input; empty for another kind. */
  const char* formula;
  /** The number of literal points of a literal input, as `fm_chart_get_series_literal_at` reads them; 0 for another
   *  kind. */
  uint32_t literal_count;
} fm_chart_data_source_t;

/** @brief One series of a plot group. */
typedef struct {
  uint32_t struct_size;
  /** The series' `c:idx`, the key the series functions take. */
  uint32_t series_id;
  /** The plotting order of the series within the chart (`c:order`). */
  uint32_t order;
  /** 1 when the chart filter hides the series. */
  int32_t filtered;
  /** The input that supplies the series name. */
  fm_chart_data_source_t name;
  /** The input that supplies the category labels. */
  fm_chart_data_source_t categories;
  /** The input that supplies the values of a category plot. */
  fm_chart_data_source_t values;
  /** The input that supplies the x values of a scatter or bubble series. */
  fm_chart_data_source_t x;
  /** The input that supplies the y values of a scatter or bubble series. */
  fm_chart_data_source_t y;
  /** The input that supplies the bubble sizes of a bubble series. */
  fm_chart_data_source_t size;
  /** The series fill. */
  fm_chart_fill_t fill;
  /** The series outline. */
  fm_chart_line_t line;
  /** 1 when `marker` is set; 0 means the series has no marker. */
  int32_t has_marker;
  /** The series marker; meaningful when `has_marker` is set. */
  fm_chart_marker_t marker;
  /** Whether a line or scatter series draws a smoothed line, a tri-state value. */
  int32_t smooth; /* fm_tristate_t */
  /** The number of formatted points of the series, as `fm_chart_get_series_point_at` reads them. */
  uint32_t point_count;
  /** 1 when `labels` is set; 0 means the series has no data labels. */
  int32_t has_labels;
  /** The data labels of the series; meaningful when `has_labels` is set. */
  fm_chart_labels_t labels;
} fm_chart_series_info_t;

/** @brief A literal point: a number, or `text` when `is_text`. */
typedef struct {
  uint32_t struct_size;
  /** The 0-based index of the point within the input. */
  uint32_t idx;
  /** 1 when the point is `text`; 0 when it is `number`. */
  int32_t is_text;
  /** The numeric value of the point; meaningful when `is_text` is 0. */
  double number;
  /** The text of the point; meaningful when `is_text` is 1. */
  const char* text;
} fm_chart_literal_info_t;

/**
 * @brief One series input to write (`fm_chart_series_patch_t`).
 *
 * `REF` writes the reference text `formula`; `LITERAL` writes the
 * `literal_count` points of `literal` (each element carries its
 * `struct_size`): a series name is one point written as `<c:tx><c:v>`
 * (`CT_SerTx` has no `c:strLit`), `values` / `y` / `size` hold numbers
 * (`c:numLit`), and `categories` / `x` are `c:numLit` when every point is a
 * number, else `c:strLit`; `NONE` writes no input.
 */
typedef struct {
  /** Whether the input is empty, a range reference or literal points, an `fm_chart_source_kind_t`. */
  int32_t kind;
  /** The reference text to write; read for a range-reference input only. */
  const char* formula;
  /** The literal points to write; read for a literal input only. */
  const fm_chart_literal_info_t* literal;
  /** The number of entries in `literal`. */
  uint32_t literal_count;
} fm_chart_data_source_spec_t;

/** @brief One formatted data point (`c:dPt`). */
typedef struct {
  uint32_t struct_size;
  /** The 0-based index of the data point within its series. */
  uint32_t idx;
  /** The point fill. */
  fm_chart_fill_t fill;
  /** The point outline. */
  fm_chart_line_t line;
  /** 1 when `marker` is set; 0 means the point states no marker. */
  int32_t has_marker;
  /** The point marker; meaningful when `has_marker` is set. */
  fm_chart_marker_t marker;
} fm_chart_point_info_t;

/** @brief One axis. */
typedef struct {
  uint32_t struct_size;
  /** The axis id (c:axId) that `cross_axis` and a plot group's axis ids refer to. */
  uint32_t id;
  /** The kind of the axis (category, value, date or series), an `fm_chart_axis_kind_t`. */
  int32_t kind;
  /** The side of the plot area the axis sits on, an `fm_chart_axis_position_t`. */
  int32_t position;
  /** 1 when the axis is shown; 0 for a deleted axis (c:delete). */
  int32_t visible;
  /** The id of the axis this axis crosses. */
  uint32_t cross_axis;
  /** Where this axis crosses its cross axis (automatic, minimum, maximum or at `crosses_at`), an
   *  `fm_chart_axis_crosses_t`. */
  int32_t crosses;
  /** 1 when `crosses_at` is set; 0 means the axis states no crossing value. */
  int32_t has_crosses_at;
  /** The value on the cross axis where this axis crosses; meaningful when `has_crosses_at` is set. */
  double crosses_at;
  /** 1 when the axis runs in reverse order. */
  int32_t reversed;
  /** 1 when `log_base` is set; 0 means a linear axis. */
  int32_t has_log_base;
  /** The base of a logarithmic scale; meaningful when `has_log_base` is set. */
  double log_base;
  /** 1 when `min` is set; 0 means the minimum is automatic. */
  int32_t has_min;
  /** The fixed minimum of the axis scale; meaningful when `has_min` is set. */
  double min;
  /** 1 when `max` is set; 0 means the maximum is automatic. */
  int32_t has_max;
  /** The fixed maximum of the axis scale; meaningful when `has_max` is set. */
  double max;
  /** 1 when `major_unit` is set; 0 means the major unit is automatic. */
  int32_t has_major_unit;
  /** The fixed distance between major ticks; meaningful when `has_major_unit` is set. */
  double major_unit;
  /** 1 when `minor_unit` is set; 0 means the minor unit is automatic. */
  int32_t has_minor_unit;
  /** The fixed distance between minor ticks; meaningful when `has_minor_unit` is set. */
  double minor_unit;
  /** The number format code of the tick labels; `NULL` when absent. */
  const char* number_format;
  /** 1 when the tick label number format follows the source cells. */
  int32_t source_linked;
  /** The axis title; kind NONE when the axis has none. */
  fm_chart_title_t title;
  /** 1 when major gridlines are drawn. */
  int32_t major_gridlines;
  /** 1 when minor gridlines are drawn. */
  int32_t minor_gridlines;
  /** Where the tick labels sit relative to the axis, an `fm_chart_tick_label_position_t`;
   *  `FM_CHART_TICK_LABEL_POSITION_ABSENT` when the axis element is absent. */
  int32_t tick_label_position;
} fm_chart_axis_info_t;

/**
 * @brief Reads the chart `object_id` into an owned snapshot.
 *
 * `*out` is set to `NULL` before validation. `*out_present` is `1` when the
 * sheet has a chart with that id; the caller then owns `*out` and releases it
 * with `fm_chart_destroy`. Otherwise `*out_present` is `0` and `*out` stays
 * `NULL` (still `FM_STATUS_OK` (0)). Later workbook edits do not change the
 * snapshot.
 *
 * Lifetime: owned (until fm_chart_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out_present` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse;
 *         `FM_STATUS_IO_CHART_UNPARSEABLE` (5025) when the chart part does not parse.
 */
FM_API fm_status_t fm_sheet_get_chart(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                      int32_t* out_present, fm_chart_t** out);

/** @brief Releases a chart handle. `chart == NULL` is a no-op. */
FM_API void fm_chart_destroy(fm_chart_t* chart);

/**
 * @brief Reads the chart's top-level fields.
 *
 * Lifetime: owned (until fm_chart_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_chart_get_info(const fm_chart_t* chart, fm_chart_info_t* out);

/**
 * @brief Reads the `c:idx` of legend entry `index` hidden from the legend
 *        (`index < legend_hidden_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_chart_get_legend_hidden_at(const fm_chart_t* chart, size_t index, uint32_t* out);

/**
 * @brief Reads the position of hidden category `index`
 *        (`index < hidden_category_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_chart_get_hidden_category_at(const fm_chart_t* chart, size_t index, uint32_t* out);

/**
 * @brief Reads plot group `plot_index` (`plot_index < plot_count`).
 *
 * Lifetime: owned (until fm_chart_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `plot_index` is out of range.
 */
FM_API fm_status_t fm_chart_get_plot_at(const fm_chart_t* chart, size_t plot_index, fm_chart_plot_info_t* out);

/**
 * @brief Reads axis id `index` of plot group `plot_index` (`index < axis_id_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `plot_index` or `index` is out of range.
 */
FM_API fm_status_t fm_chart_get_plot_axis_id_at(const fm_chart_t* chart, size_t plot_index, size_t index,
                                                uint32_t* out);

/**
 * @brief Reads series `series` of plot group `plot_index`, in document order
 *        (`series < series_count`).
 *
 * Lifetime: owned (until fm_chart_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `plot_index` or `series` is out of range.
 */
FM_API fm_status_t fm_chart_get_series_at(const fm_chart_t* chart, size_t plot_index, size_t series,
                                          fm_chart_series_info_t* out);

/**
 * @brief Reads literal point `index` of input `role` (an `fm_chart_role_t`)
 *        of a series (`index < literal_count`).
 *
 * Lifetime: owned (until fm_chart_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `plot_index`, `series` or `index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `role` is not a series role.
 */
FM_API fm_status_t fm_chart_get_series_literal_at(const fm_chart_t* chart, size_t plot_index, size_t series,
                                                  int32_t role, size_t index, fm_chart_literal_info_t* out);

/**
 * @brief Reads formatted point `index` of a series (`index < point_count`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `plot_index`, `series` or `index` is out of range.
 */
FM_API fm_status_t fm_chart_get_series_point_at(const fm_chart_t* chart, size_t plot_index, size_t series, size_t index,
                                                fm_chart_point_info_t* out);

/**
 * @brief Reads axis `index` (`index < axis_count`).
 *
 * Lifetime: owned (until fm_chart_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_chart_get_axis_at(const fm_chart_t* chart, size_t index, fm_chart_axis_info_t* out);

/* ---- Resolved chart data ------------------------------------------------- */

/** @brief Opaque resolved series data. Release with `fm_chart_data_destroy`. */
typedef struct fm_chart_data fm_chart_data_t;

/**
 * @brief Overrides `fm_sheet_get_chart_preview` applies to a copy of the
 *        chart. Each field applies when its `_op` is `FM_PATCH_OP_SET`
 *        (`KEEP` keeps the chart's own value; `CLEAR` is rejected).
 */
typedef struct {
  uint32_t struct_size;
  /** Patch operation for `source`: keep it, set it from `source`; clear is rejected. */
  fm_patch_op_t source_op;
  /** The A1 range of the source, applied when `source_op` is set. */
  const char* source;
  /** Patch operation for `direction`: keep it, set it from `direction`; clear is rejected. */
  fm_patch_op_t direction_op;
  /** Whether the series run along rows or columns, an `fm_chart_series_direction_t`, applied when `direction_op` is
   *  set. */
  int32_t direction;
  /** Patch operation for `header_row`: keep it, set it from `header_row`; clear is rejected. */
  fm_patch_op_t header_row_op;
  /** 1 when the first row holds names, applied when `header_row_op` is set. */
  int32_t header_row;
  /** Patch operation for `header_column`: keep it, set it from `header_column`; clear is rejected. */
  fm_patch_op_t header_column_op;
  /** 1 when the first column holds names, applied when `header_column_op` is set. */
  int32_t header_column;
  /** Patch operation for `blanks_as`: keep it, set it from `blanks_as`; clear is rejected. */
  fm_patch_op_t blanks_as_op;
  /** How empty cells are plotted, an `fm_blanks_as_t`, applied when `blanks_as_op` is set. */
  int32_t blanks_as;
  /** Patch operation for `plot_visible_only`: keep it, set it from `plot_visible_only`; clear is rejected. */
  fm_patch_op_t plot_visible_only_op;
  /** 1 to plot only visible cells, applied when `plot_visible_only_op` is set. */
  int32_t plot_visible_only;
  /** Patch operation for `na_as_blank`: keep it, set it from `na_as_blank`; clear is rejected. */
  fm_patch_op_t na_as_blank_op;
  /** 1 to draw `#N/A` points as blank cells, applied when `na_as_blank_op` is set. */
  int32_t na_as_blank;
} fm_chart_preview_overrides_spec_t;

/** @brief Summary of resolved chart data. */
typedef struct {
  uint32_t struct_size;
  /** 1 for the XY domain (scatter, bubble), where points carry an x value; 0 for the category domain. */
  int32_t is_xy;
  /** The number of points per series. */
  uint32_t point_count;
  /** The number of resolved series. */
  uint32_t series_count;
  /** The number of category labels. */
  uint32_t category_label_count;
  /** The number of numeric (date) category values; 0 unless the categories are numeric, then equal to
   *  `category_label_count`. */
  uint32_t category_value_count;
  /** The number of upper levels of a multi-level category axis (the lowest level is the category labels); 0 for a
   *  single-level axis. */
  uint32_t category_level_count;
  /** The number format code of numeric category values; `NULL` when not numeric. */
  const char* category_number_format;
  /** 1 when no category reference exists and the labels are generated as 1..n. */
  int32_t categories_generated;
  /** The resolved chart title text; `NULL` when the chart has none. */
  const char* title_text;
  /** The number of resolved axis titles. */
  uint32_t axis_title_count;
  /** The number of inputs that did not resolve. */
  uint32_t unresolved_count;
} fm_chart_data_info_t;

/** @brief One resolved series. */
typedef struct {
  uint32_t struct_size;
  /** The 0-based position of the series' plot group. */
  uint32_t plot_index;
  /** The series' `c:idx`. */
  uint32_t series_id;
  /** The resolved series name; `NULL` when the series has no name reference. */
  const char* name;
  /** The number format code of the values (the y values in the XY domain). */
  const char* number_format;
  /** `1` when the chart filter hides the series. */
  int32_t filtered;
  /** The number of points of the series: the length of the arrays `fm_chart_data_get_series_points` returns. */
  uint32_t point_count;
  /** 1 when the series is a bubble series with sizes; 0 otherwise. */
  int32_t has_sizes;
  /** The number of resolved rectangles of each series input, indexed by `fm_chart_role_t`: 0 name, 1 categories, 2
   *  values, 3 x, 4 y, 5 size. */
  uint32_t range_count[6];
  /** The number format code of the x values; `NULL` outside the XY domain. */
  const char* x_number_format;
} fm_chart_data_series_info_t;

/** @brief Parallel point arrays of one series, owned by the handle. */
typedef struct {
  uint32_t struct_size;
  /** The number of entries in each array. */
  uint32_t count;
  /** The resolved values (the y values in the XY domain), `count` entries. */
  const double* values;
  /** The state of each point, an `fm_chart_point_state_t`, `count` entries parallel to `values`. */
  const uint8_t* states;
  /** The resolved x values, `count` entries; `NULL` outside the XY domain. */
  const double* x;
  /** The resolved bubble sizes, `count` entries; `NULL` unless the series is a bubble series. */
  const double* sizes;
} fm_chart_data_points_info_t;

/** @brief Resolved title text of one axis. */
typedef struct {
  uint32_t struct_size;
  /** The id of the axis the title belongs to. */
  uint32_t axis_id;
  /** The resolved title text. */
  const char* text;
} fm_chart_axis_title_text_info_t;

/** @brief An input that did not resolve; its points read as gaps. */
typedef struct {
  uint32_t struct_size;
  /** The `c:idx` of the series the input belongs to; 0 for a title. */
  uint32_t series_id;
  /** The role of the input that did not resolve, an `fm_chart_role_t`. */
  int32_t role;
  /** The id of the axis when the role is an axis title. */
  uint32_t axis_id;
} fm_chart_unresolved_info_t;

/**
 * @brief Resolves the chart's series against the current cell values, by
 *        Excel's rules (hidden cells, blanks, errors, filters).
 *
 * `*out` is set to `NULL` before validation. `*out_present` is `1` when the
 * sheet has a chart with that id; the caller then owns `*out` and releases it
 * with `fm_chart_data_destroy`. Otherwise `*out_present` is `0` and `*out`
 * stays `NULL` (still `FM_STATUS_OK` (0)).
 *
 * Lifetime: owned (until fm_chart_data_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `out_present` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse;
 *         `FM_STATUS_IO_CHART_UNPARSEABLE` (5025) when the chart part does not parse.
 */
FM_API fm_status_t fm_sheet_get_chart_data(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                           int32_t* out_present, fm_chart_data_t** out);

/**
 * @brief `fm_sheet_get_chart_data` for a copy of the chart with
 *        `overrides` applied; the workbook is not changed.
 *
 * A source, direction or header override re-splits the series as
 * `fm_sheet_set_chart_source` does: over the overridden range (an unset
 * direction or header is auto), else over the chart's derived range (unset
 * ones as derived).
 *
 * Lifetime: owned (until fm_chart_data_destroy).
 *
 * @return as `fm_sheet_get_chart_data` (an absent id reads `*out_present == 0`
 *         whatever `overrides` holds), and
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `overrides` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for an op outside 0..2, a `CLEAR` op,
 *         `source_op` `SET` with a `NULL` `source`, or an out-of-domain value;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when `source` does not resolve;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for an extended chart.
 */
FM_API fm_status_t fm_sheet_get_chart_preview(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                              const fm_chart_preview_overrides_spec_t* overrides, int32_t* out_present,
                                              fm_chart_data_t** out);

/** @brief Releases a chart data handle. `data == NULL` is a no-op. */
FM_API void fm_chart_data_destroy(fm_chart_data_t* data);

/**
 * @brief Reads the data summary.
 *
 * Lifetime: owned (until fm_chart_data_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success; `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`.
 */
FM_API fm_status_t fm_chart_data_get_info(const fm_chart_data_t* data, fm_chart_data_info_t* out);

/**
 * @brief Reads category label `index` (`index < category_label_count`) and, for
 *        numeric categories, its value: `*out_has_value` is `1` with `*out_value`
 *        set, else `0` with `*out_value` zero.
 *
 * `*out_label` is an owned snapshot held by `data`, valid until
 * `fm_chart_data_destroy`.
 *
 * Lifetime: owned (until fm_chart_data_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_category_at(const fm_chart_data_t* data, size_t index, const char** out_label,
                                                 int32_t* out_has_value, double* out_value);

/**
 * @brief Reads the label count of upper category level `level`
 *        (`level < category_level_count`, outermost first).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `level` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_category_level_count(const fm_chart_data_t* data, size_t level, size_t* out_count);

/**
 * @brief Reads label `index` of upper category level `level`.
 *
 * `*out_label` is an owned snapshot held by `data`, valid until
 * `fm_chart_data_destroy`.
 *
 * Lifetime: owned (until fm_chart_data_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `level` or `index` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_category_level_at(const fm_chart_data_t* data, size_t level, size_t index,
                                                       const char** out_label);

/**
 * @brief Reads resolved series `series` (`series < series_count`).
 *
 * Lifetime: owned (until fm_chart_data_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `series` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_series_at(const fm_chart_data_t* data, size_t series,
                                               fm_chart_data_series_info_t* out);

/**
 * @brief Returns the point arrays of series `series`.
 *
 * Lifetime: owned (until fm_chart_data_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `series` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_series_points(const fm_chart_data_t* data, size_t series,
                                                   fm_chart_data_points_info_t* out);

/**
 * @brief Reads resolved rectangle `index` of input `role` (an
 *        `fm_chart_role_t`) of series `series` (`index < range_count[role]`).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `series` or `index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `role` is not a series role.
 */
FM_API fm_status_t fm_chart_data_get_range_at(const fm_chart_data_t* data, size_t series, int32_t role, size_t index,
                                              fm_sheet_rect_t* out);

/**
 * @brief Reads axis title `index` (`index < axis_title_count`).
 *
 * Lifetime: owned (until fm_chart_data_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_axis_title_at(const fm_chart_data_t* data, size_t index,
                                                   fm_chart_axis_title_text_info_t* out);

/**
 * @brief Reads unresolved input `index` (`index < unresolved_count`).
 *
 * Lifetime: owned (until fm_chart_data_destroy).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when any pointer argument is `NULL`;
 *         `FM_STATUS_INDEX_OUT_OF_RANGE` (13) when `index` is out of range.
 */
FM_API fm_status_t fm_chart_data_get_unresolved_at(const fm_chart_data_t* data, size_t index,
                                                   fm_chart_unresolved_info_t* out);

/* ---- Creation and edits -------------------------------------------------- */

/** @brief Options of `fm_sheet_add_chart`. */
typedef struct {
  uint32_t struct_size;
  /** The plot type of the new chart, an `fm_chart_plot_type_t`: `BAR`, `LINE`, `AREA`, `PIE`, `DOUGHNUT` or
   *  `SCATTER` (any other is `FM_STATUS_INVALID_ARGUMENT` (2)). */
  int32_t type;
  /** Whether bars run vertically (column) or horizontally, an `fm_chart_bar_direction_t`; read for `BAR` only, 0 is
   *  column, and `FM_CHART_BAR_DIRECTION_BAR` with any other type is `FM_STATUS_INVALID_ARGUMENT` (2). */
  int32_t bar_direction;
  /** How the series are stacked or grouped, an `fm_chart_grouping_t`; applies to bar, line and area plots. */
  int32_t grouping;
  /** One A1 rectangle, sheet-qualified or on the target sheet. */
  const char* source;
  /** Whether the series run along rows or columns of the source range, an `fm_chart_series_direction_t`; automatic when
   *  omitted. */
  int32_t direction;
  /** Whether the first row of the source holds names, a tri-state value (default follows Excel's detection). */
  int32_t header_row; /* fm_tristate_t */
  /** Whether the first column of the source holds names, a tri-state value (default follows Excel's detection). */
  int32_t header_column; /* fm_tristate_t */
  /** 1 when `title` is written; 0 gives Excel's automatic title. */
  int32_t has_title;
  /** The chart title, written when `has_title` is 1 as `fm_sheet_set_chart_title` does (`NONE` for no title; a `TEXT`
   *  title with empty text is `FM_STATUS_INVALID_ARGUMENT` (2)). A `REF` title is resolved before the insert, so a bad
   *  reference fails with no chart added. */
  fm_chart_title_t title;
  /** The legend position, an `fm_chart_legend_position_t`; 0 is the right (the OOXML default) while Excel's insert
   *  default is the bottom, which the bindings pass for an omitted option. */
  int32_t legend;
  /** The placement, name and description of the chart frame, read as for `fm_sheet_add_picture` (an empty name becomes
   *  `Chart <n>`, n being the object id minus one). A `DEFAULT` anchor kind is two-cell, as Excel inserts charts, and a
   *  zero size means Excel's default 360 x 216 pt frame. */
  fm_drawing_placement_spec_t placement;
} fm_chart_spec_t;

/** @brief Marker patch; `symbol` is an `fm_chart_marker_symbol_t` other than `FM_CHART_MARKER_SYMBOL_PICTURE` (`CLEAR`
 * rejected). */
typedef struct {
  /** Patch operation for `symbol`: keep it, set it from `symbol`; clear is rejected. */
  fm_patch_op_t symbol_op;
  /** The marker symbol to set; the picture symbol is rejected. */
  int32_t symbol;
  /** Patch operation for `size_pt`: keep it, set it from `size_pt`, or clear it. */
  fm_patch_op_t size_pt_op;
  /** The marker size to set, in points. */
  double size_pt;
  /** The patch applied to the marker fill. */
  fm_chart_fill_patch_t fill;
  /** The patch applied to the marker outline. */
  fm_chart_line_patch_t line;
} fm_chart_marker_patch_t;

/** @brief Data-label patch. A `SET` `font_color` is kind None, Rgb or Theme. */
typedef struct {
  /** Patch operation for `show_value`: keep it, set it from `show_value`; clear is rejected. */
  fm_patch_op_t show_value_op;
  /** 1 to show the point value, 0 to hide it. */
  int32_t show_value;
  /** Patch operation for `show_category`: keep it, set it from `show_category`; clear is rejected. */
  fm_patch_op_t show_category_op;
  /** 1 to show the category name, 0 to hide it. */
  int32_t show_category;
  /** Patch operation for `show_series`: keep it, set it from `show_series`; clear is rejected. */
  fm_patch_op_t show_series_op;
  /** 1 to show the series name, 0 to hide it. */
  int32_t show_series;
  /** Patch operation for `show_percent`: keep it, set it from `show_percent`; clear is rejected. */
  fm_patch_op_t show_percent_op;
  /** 1 to show the percentage of the whole, 0 to hide it. */
  int32_t show_percent;
  /** Patch operation for `show_legend_key`: keep it, set it from `show_legend_key`; clear is rejected. */
  fm_patch_op_t show_legend_key_op;
  /** 1 to show the legend key, 0 to hide it. */
  int32_t show_legend_key;
  /** Patch operation for `separator`: keep it, set it from `separator`, or clear it. */
  fm_patch_op_t separator_op;
  /** The text placed between the label parts. */
  const char* separator;
  /** Patch operation for `position`: keep it, set it from `position`, or clear it. */
  fm_patch_op_t position_op;
  /** Where the label sits relative to its point, an `fm_chart_label_position_t`. */
  int32_t position;
  /** Patch operation for `number_format`: keep it, set it from `number_format`, or clear it. */
  fm_patch_op_t number_format_op;
  /** The number format code of the label. */
  const char* number_format;
  /** Patch operation for `font_color`: keep it, set it from `font_color`, or clear it. */
  fm_patch_op_t font_color_op;
  /** The label font colour to set, kind None, RGB or theme. */
  fm_color_t font_color;
  /** Patch operation for `font_size_pt`: keep it, set it from `font_size_pt`, or clear it. */
  fm_patch_op_t font_size_pt_op;
  /** The label font size to set, in points. */
  double font_size_pt;
  /** Patch operation for `source_linked`: keep it, set it from `source_linked`; clear is rejected. */
  fm_patch_op_t source_linked_op;
  /** 1 when the label number format follows the source cells. */
  int32_t source_linked;
  /** Patch operation for `deleted`: keep it, set it from `deleted`; clear is rejected. */
  fm_patch_op_t deleted_op;
  /** 1 to delete the labels, 0 to keep them. */
  int32_t deleted;
} fm_chart_labels_patch_t;

/**
 * @brief Series patch.
 *
 * Each input (`name`, `categories`, `values`, `x`, `y`, `size`) has an
 * `fm_patch_op_t` companion and an `fm_chart_data_source_spec_t`. `KEEP`
 * leaves it. `SET` with `REF` writes `formula` (`NULL` or empty is
 * `FM_STATUS_INVALID_ARGUMENT` (2)); with `LITERAL` writes the points
 * (`literal_count` 0 is `FM_STATUS_INVALID_ARGUMENT` (2)); with `NONE`, like
 * `CLEAR`, removes the input of `name`, `categories` or `x` and is
 * `FM_STATUS_INVALID_ARGUMENT` (2) on `values`, `y` and `size`. `smooth` is
 * 0/1 (`CLEAR` rejected).
 */
typedef struct {
  uint32_t struct_size;
  /** Patch operation for `name`: keep it, set it from `name`, or clear it. */
  fm_patch_op_t name_op;
  /** The series name input to write. */
  fm_chart_data_source_spec_t name;
  /** Patch operation for `categories`: keep it, set it from `categories`, or clear it. */
  fm_patch_op_t categories_op;
  /** The category labels input to write. */
  fm_chart_data_source_spec_t categories;
  /** Patch operation for `values`: keep it, set it from `values`, or clear it. */
  fm_patch_op_t values_op;
  /** The values input to write. */
  fm_chart_data_source_spec_t values;
  /** Patch operation for `x`: keep it, set it from `x`, or clear it. */
  fm_patch_op_t x_op;
  /** The x values input to write. */
  fm_chart_data_source_spec_t x;
  /** Patch operation for `y`: keep it, set it from `y`, or clear it. */
  fm_patch_op_t y_op;
  /** The y values input to write. */
  fm_chart_data_source_spec_t y;
  /** Patch operation for `size`: keep it, set it from `size`, or clear it. */
  fm_patch_op_t size_op;
  /** The bubble sizes input to write. */
  fm_chart_data_source_spec_t size;
  /** The patch applied to the series fill. */
  fm_chart_fill_patch_t fill;
  /** The patch applied to the series outline. */
  fm_chart_line_patch_t line;
  /** Patch operation for `marker`: keep it, set it from `marker`, or clear it. */
  fm_patch_op_t marker_op;
  /** The patch applied to the series marker; clearing it with `marker_op` removes the marker. */
  fm_chart_marker_patch_t marker;
  /** Patch operation for `smooth`: keep it, set it from `smooth`; clear is rejected. */
  fm_patch_op_t smooth_op;
  /** 1 to draw a smoothed line, 0 for straight segments. */
  int32_t smooth;
  /** Patch operation for `labels`: keep it, set it from `labels`, or clear it. */
  fm_patch_op_t labels_op;
  /** The patch applied to the data labels; clearing it with `labels_op` removes them. */
  fm_chart_labels_patch_t labels;
} fm_chart_series_patch_t;

/**
 * @brief Data-point patch (`fm_sheet_update_chart_point`).
 *
 * A patch whose ops are all `KEEP` changes nothing. A point whose override
 * ends empty after the patch (the fill and line members all cleared and
 * `marker_op` `CLEAR`) loses its `c:dPt`; clearing every member is how a
 * point override is removed, and removing a point that has none does nothing.
 */
typedef struct {
  uint32_t struct_size;
  /** The patch applied to the point fill. */
  fm_chart_fill_patch_t fill;
  /** The patch applied to the point outline. */
  fm_chart_line_patch_t line;
  /** Patch operation for `marker`: keep it, set it from `marker`, or clear it. */
  fm_patch_op_t marker_op;
  /** The patch applied to the point marker. */
  fm_chart_marker_patch_t marker;
} fm_chart_point_patch_t;

/**
 * @brief Plot-group patch (`CLEAR` rejected on every field). A field the
 *        plot type does not have is `FM_STATUS_INVALID_ARGUMENT` (2).
 */
typedef struct {
  uint32_t struct_size;
  /** Patch operation for `gap_width_pct`: keep it, set it from `gap_width_pct`; clear is rejected. */
  fm_patch_op_t gap_width_pct_op;
  /** The gap between bar clusters as a percentage of the bar width. */
  uint32_t gap_width_pct;
  /** Patch operation for `overlap_pct`: keep it, set it from `overlap_pct`; clear is rejected. */
  fm_patch_op_t overlap_pct_op;
  /** The overlap between bars of a cluster as a percentage of the bar width. */
  int32_t overlap_pct;
  /** Patch operation for `first_slice_angle_deg`: keep it, set it from `first_slice_angle_deg`; clear is rejected. */
  fm_patch_op_t first_slice_angle_deg_op;
  /** The angle in degrees of the first pie or doughnut slice. */
  uint32_t first_slice_angle_deg;
  /** Patch operation for `vary_colors`: keep it, set it from `vary_colors`; clear is rejected. */
  fm_patch_op_t vary_colors_op;
  /** 1 when each point of a single-series plot takes its own colour. */
  int32_t vary_colors;
  /** Patch operation for `hole_size_pct`: keep it, set it from `hole_size_pct`; clear is rejected. */
  fm_patch_op_t hole_size_pct_op;
  /** The doughnut hole size as a percentage of the diameter (10..90); doughnut plots only. */
  uint32_t hole_size_pct;
} fm_chart_plot_patch_t;

/**
 * @brief Axis patch.
 *
 * `CLEAR` is rejected on `visible`, `major_gridlines`, `reversed`, `crosses`,
 * `crosses_at`, `minor_gridlines`, `source_linked` and `position`; on `title`
 * it removes the axis title, on `log_base` it makes the axis linear, on
 * `tick_label_position` it restores the default, and on `minor_unit` it makes
 * the unit automatic.
 */
typedef struct {
  uint32_t struct_size;
  /** Patch operation for `visible`: keep it, set it from `visible`; clear is rejected. */
  fm_patch_op_t visible_op;
  /** 1 to show the axis, 0 to delete it. */
  int32_t visible;
  /** Patch operation for `title`: keep it, set it from `title`, or clear it. */
  fm_patch_op_t title_op;
  /** The axis title to set, taking the kinds of `fm_sheet_set_chart_title`; none removes the title and automatic leaves
   *  the text to the viewer. */
  fm_chart_title_t title;
  /** Patch operation for `min`: keep it, set it from `min`, or clear it. */
  fm_patch_op_t min_op;
  /** The fixed minimum of the axis scale. */
  double min;
  /** Patch operation for `max`: keep it, set it from `max`, or clear it. */
  fm_patch_op_t max_op;
  /** The fixed maximum of the axis scale. */
  double max;
  /** Patch operation for `major_unit`: keep it, set it from `major_unit`, or clear it. */
  fm_patch_op_t major_unit_op;
  /** The fixed distance between major ticks. */
  double major_unit;
  /** Patch operation for `number_format`: keep it, set it from `number_format`, or clear it. */
  fm_patch_op_t number_format_op;
  /** The number format code of the tick labels. */
  const char* number_format;
  /** Patch operation for `major_gridlines`: keep it, set it from `major_gridlines`; clear is rejected. */
  fm_patch_op_t major_gridlines_op;
  /** 1 to draw major gridlines, 0 to remove them. */
  int32_t major_gridlines;
  /** Patch operation for `reversed`: keep it, set it from `reversed`; clear is rejected. */
  fm_patch_op_t reversed_op;
  /** 1 to run the axis in reverse order. */
  int32_t reversed;
  /** Patch operation for `log_base`: keep it, set it from `log_base`, or clear it. */
  fm_patch_op_t log_base_op;
  /** The base of a logarithmic scale; clearing it makes the axis linear. */
  double log_base;
  /** Patch operation for `crosses`: keep it, set it from `crosses`; clear is rejected. */
  fm_patch_op_t crosses_op;
  /** Where this axis crosses its cross axis (automatic, minimum, maximum or at `crosses_at`), an
   *  `fm_chart_axis_crosses_t`. */
  int32_t crosses;
  /** Patch operation for `crosses_at`: keep it, set it from `crosses_at`; clear is rejected. */
  fm_patch_op_t crosses_at_op;
  /** The value on the cross axis where this axis crosses; read when `crosses` is set to `FM_CHART_AXIS_CROSSES_AT`. */
  double crosses_at;
  /** Patch operation for `tick_label_position`: keep it, set it from `tick_label_position`, or clear it. */
  fm_patch_op_t tick_label_position_op;
  /** Where the tick labels sit relative to the axis, an `fm_chart_tick_label_position_t` other than `ABSENT`. */
  int32_t tick_label_position;
  /** Patch operation for `minor_unit`: keep it, set it from `minor_unit`, or clear it. */
  fm_patch_op_t minor_unit_op;
  /** The fixed distance between minor ticks; clearing it makes the unit automatic. */
  double minor_unit;
  /** Patch operation for `minor_gridlines`: keep it, set it from `minor_gridlines`; clear is rejected. */
  fm_patch_op_t minor_gridlines_op;
  /** 1 to draw minor gridlines, 0 to remove them. */
  int32_t minor_gridlines;
  /** Patch operation for `position`: keep it, set it from `position`; clear is rejected. */
  fm_patch_op_t position_op;
  /** The side of the plot area the axis sits on, an `fm_chart_axis_position_t`; an axis always has `c:axPos`. */
  int32_t position;
  /** Patch operation for `source_linked`: keep it, set it from `source_linked`; clear is rejected. */
  fm_patch_op_t source_linked_op;
  /** 1 when the tick label number format follows the source cells. */
  int32_t source_linked;
} fm_chart_axis_patch_t;

/**
 * @brief A chart legend as a whole (`fm_sheet_set_chart_legend`).
 *
 * The read side is `fm_chart_info_t`'s `legend_*` fields and
 * `fm_chart_get_legend_hidden_at`.
 */
typedef struct {
  uint32_t struct_size;
  /** Where the legend sits relative to the plot, an `fm_chart_legend_position_t` other than
   *  `FM_CHART_LEGEND_POSITION_ABSENT` (a chart without a legend takes `NULL` for the whole record instead). */
  int32_t position;
  /** `1` when the legend overlays the plot area. */
  int32_t overlay;
  /** The `c:idx` values of the series whose legend entry is hidden; may be `NULL` when `hidden_entries_count` is 0. */
  const uint32_t* hidden_entries;
  /** The number of entries in `hidden_entries`. */
  uint32_t hidden_entries_count;
} fm_chart_legend_t;

/**
 * @brief Legend merge (`fm_sheet_update_chart_legend`);
 *        `fm_sheet_set_chart_legend(..., NULL)` removes the legend.
 *
 * No field can be cleared. On a chart without a legend the update creates one,
 * and a `KEEP` field takes Excel's default: `FM_CHART_LEGEND_POSITION_RIGHT`,
 * no overlay, no hidden entry.
 */
typedef struct {
  uint32_t struct_size;
  /** Patch operation for `position`: keep it or set it from `position`; clear is rejected. */
  fm_patch_op_t position_op;
  /** Where the legend sits relative to the plot, an `fm_chart_legend_position_t`; the absent value is rejected. */
  int32_t position;
  /** Patch operation for `overlay`: keep it or set it from `overlay`; clear is rejected. */
  fm_patch_op_t overlay_op;
  /** 1 when the legend overlays the plot area, 0 otherwise. */
  int32_t overlay;
  /** Patch operation for `hidden_entries`: keep it or set it from `hidden_entries`; clear is rejected. */
  fm_patch_op_t hidden_entries_op;
  /** The series `c:idx` values whose legend entry is hidden; `NULL` when `hidden_entries_count` is 0, which shows
   *  every entry. */
  const uint32_t* hidden_entries;
  /** The number of entries in `hidden_entries`. */
  uint32_t hidden_entries_count;
} fm_chart_legend_patch_t;

/** @brief Display-option patch. `CLEAR` removes the element. */
typedef struct {
  uint32_t struct_size;
  /** Patch operation for `blanks_as`: keep it, set it from `blanks_as`, or clear (remove) it. */
  fm_patch_op_t blanks_as_op;
  /** How empty cells are plotted (`c:dispBlanksAs`): an `fm_blanks_as_t` of `GAP`, `ZERO` or `SPAN`. */
  int32_t blanks_as;
  /** Patch operation for `plot_visible_only`: keep it, set it from `plot_visible_only`, or clear (remove) it. */
  fm_patch_op_t plot_visible_only_op;
  /** 1 to plot only visible cells (`c:plotVisOnly`), 0 to plot all. */
  int32_t plot_visible_only;
  /** Patch operation for `na_as_blank`: keep it, set it from `na_as_blank`, or clear (remove) it. */
  fm_patch_op_t na_as_blank_op;
  /** 1 to draw `#N/A` points as blank cells (`c16r3:dispNaAsBlank`); 0 removes the setting, as Excel writes it. */
  int32_t na_as_blank;
  /** Patch operation for `style`: keep it, set it from `style`, or clear (remove) it. */
  fm_patch_op_t style_op;
  /** The chart style number (`c:style`), 1..48 (`ST_Style`). */
  int32_t style;
} fm_chart_display_patch_t;

/**
 * @brief Creates a chart from `spec` with filled caches and returns its
 *        object id.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `spec`, `spec->source` or
 *         `out_object_id` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) for a bad enum value or placement (as
 *         `fm_sheet_add_picture`), a source that does not yield a series, or a
 *         title `fm_sheet_set_chart_title` refuses;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when `source` does not resolve to
 *         one rectangle, or a `REF` title's `formula` does not resolve;
 *         `FM_STATUS_IO_DRAWING_UNPARSEABLE` (5024) when the drawing part does not parse
 *         or is an `.xlsb` drawing the reader could not lift;
 *         `FM_STATUS_SECURITY_RESOURCE_LIMIT` (6005) when the package would exceed the
 *         reader's part or size limits (part=zip_parts or part=zip_total_bytes). The
 *         workbook is then unchanged.
 */
FM_API fm_status_t fm_sheet_add_chart(fm_workbook_t* wb, size_t sheet_index, const fm_chart_spec_t* spec,
                                      uint32_t* out_object_id);

/**
 * @brief Replaces plot group `plot_index` with `type` (an `fm_chart_plot_type_t`
 *        of `BAR`, `LINE`, `AREA`, `PIE`, `DOUGHNUT` or `SCATTER`),
 *        `bar_direction` (an `fm_chart_bar_direction_t`, read for `BAR` only
 *        as for `fm_chart_spec_t`) and `grouping` (an
 *        `fm_chart_grouping_t`), keeping each series' references, idx, order
 *        and explicit colors.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `type` is not a writable type or
 *         `bar_direction` is `FM_CHART_BAR_DIRECTION_BAR` for another type;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a pivot chart, or a plot that
 *         shares its axes with another plot group.
 */
FM_API fm_status_t fm_sheet_set_chart_type(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot_index,
                                           int32_t type, int32_t bar_direction, int32_t grouping);

/**
 * @brief Sets the chart title: `NONE` removes it, `TEXT` writes `text`,
 *        `REF` writes the reference `formula`, `AUTO` leaves the text to
 *        the viewer.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `title` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `kind` is outside
 *         `fm_chart_title_kind_t`, `TEXT` has empty text or `REF` no formula;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when `formula` does not resolve.
 */
FM_API fm_status_t fm_sheet_set_chart_title(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                            const fm_chart_title_t* title);

/**
 * @brief Replaces the chart's legend with `legend`, or removes it when
 *        `legend` is `NULL` (a no-op on a chart without one).
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `hidden_entries` is `NULL`
 *         with a non-zero count;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when `position` is
 *         `FM_CHART_LEGEND_POSITION_ABSENT` or outside the enumeration, or
 *         `overlay` is not 0 / 1.
 */
FM_API fm_status_t fm_sheet_set_chart_legend(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                             const fm_chart_legend_t* legend);

/**
 * @brief Merges `patch` into the chart's legend (see `fm_chart_legend_patch_t`).
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`, or
 *         `hidden_entries` is `NULL` with a non-zero count.
 */
FM_API fm_status_t fm_sheet_update_chart_legend(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                const fm_chart_legend_patch_t* patch);

/**
 * @brief Patches axis `axis_id`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`.
 */
FM_API fm_status_t fm_sheet_update_chart_axis(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                              uint32_t axis_id, const fm_chart_axis_patch_t* patch);

/**
 * @brief Patches the chart's display options.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a `SET` value is outside its domain
 *         (`blanks_as` `ABSENT` and a `style` outside 1..48 included; `CLEAR`
 *         removes the element).
 */
FM_API fm_status_t fm_sheet_update_chart_display(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 const fm_chart_display_patch_t* patch);

/**
 * @brief Patches the series whose id (`c:idx`) is `series_id` in plot group `plot_index`.
 *        A changed reference refreshes its cache.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when a reference does not
 *         resolve;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a reference edit on a pivot
 *         chart.
 */
FM_API fm_status_t fm_sheet_update_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                size_t plot_index, uint32_t series_id,
                                                const fm_chart_series_patch_t* patch);

/**
 * @brief Patches point `point_index` of the series whose id is `series_id`. An all-`KEEP`
 *        patch changes nothing; a patch that leaves the point's override
 *        empty removes it (see `fm_chart_point_patch_t`).
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`.
 */
FM_API fm_status_t fm_sheet_update_chart_point(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                               size_t plot_index, uint32_t series_id, size_t point_index,
                                               const fm_chart_point_patch_t* patch);

/**
 * @brief Patches plot group `plot_index`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` is `NULL`.
 */
FM_API fm_status_t fm_sheet_update_chart_plot(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                              size_t plot_index, const fm_chart_plot_patch_t* patch);

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
 * @brief Appends a series to plot group `plot_index` from `patch` (which must set
 *        `values`, or `y` for a scatter plot) and returns its id (`c:idx`).
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) when `patch` or `out_series_id` is `NULL`;
 *         `FM_STATUS_IO_CHART_INVALID_REFERENCE` (5026) when a reference does not
 *         resolve;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a pivot chart.
 */
FM_API fm_status_t fm_sheet_add_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                             size_t plot_index, const fm_chart_series_patch_t* patch,
                                             uint32_t* out_series_id);

/**
 * @brief Removes the series whose id is `series_id` from plot group `plot_index`.
 *
 * @return `FM_STATUS_OK` (0) on success, or a common chart failure;
 *         `FM_STATUS_IO_CHART_UNSUPPORTED_KIND` (5027) for a pivot chart.
 */
FM_API fm_status_t fm_sheet_remove_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                size_t plot_index, uint32_t series_id);

/**
 * @brief Rebuilds every series from `source` by the insert splitting rules,
 *        keeping the chart kind and each series' formatting by idx, and
 *        clears the filter. `direction` is an `fm_chart_series_direction_t`,
 *        `header_row` / `header_column` `fm_tristate_t`s as in
 *        `fm_chart_spec_t`.
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
 * The formulas and substitution values sit in the row above and the column left
 * of `rect`. The row input cell receives the values of the top row, the column
 * input cell those of the left column: a one-variable table has one of them
 * (`row_oriented` says which, on read), a two-variable table both
 * (`two_dimensional`). A flag is `0` or `1`.
 *
 * On add only the range and the `has_*` flags and inputs are read; the
 * result cells are set to 0 and filled by the next recalc.
 */
typedef struct {
  uint32_t struct_size;
  /** The result range of the table, 0-based and inclusive. */
  fm_cell_rect_t rect;
  /** The 0-based row of the row input cell; 0 when `has_row_input` is 0. */
  uint32_t row_input_row;
  /** The 0-based column of the row input cell; 0 when `has_row_input` is 0. */
  uint32_t row_input_col;
  /** The 0-based row of the column input cell; 0 when `has_column_input` is 0. */
  uint32_t column_input_row;
  /** The 0-based column of the column input cell; 0 when `has_column_input` is 0. */
  uint32_t column_input_col;
  /** 1 when the table has a row input cell; 0 means it has none. */
  int32_t has_row_input;
  /** 1 when the table has a column input cell; 0 means it has none. */
  int32_t has_column_input;
  /** 1 for a two-variable table, which has both a row and a column input cell. */
  int32_t two_dimensional;
  /** For a one-variable table, 1 when its input is the row input cell and 0 when it is the column input cell. */
  int32_t row_oriented;
  /** 1 when a row or column delete removed the row input cell; the table then evaluates to `#REF!`. Ignored on input.
   */
  int32_t row_input_deleted;
  /** 1 when a row or column delete removed the column input cell; the table then evaluates to `#REF!`. Ignored on
   *  input. */
  int32_t column_input_deleted;
} fm_data_table_t;

/**
 * @brief Creates a data table over the result range of `spec` and writes the
 *        range's top-left cell, the key `fm_sheet_get_data_table` takes, to
 *        `*out_cell`. Cells in the range may no longer be written, except by
 *        removing the table.
 *
 * `out_cell` may be `NULL`.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `spec == NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_INVALID_ARGUMENT` (2) when a
 *         coordinate is outside the grid;
 *         `FM_STATUS_DATA_TABLE_INVALID` (8005) when neither input is given, the range
 *         has no row above or column left of it, an input lies inside the
 *         range, or the range overlaps another data table.
 */
FM_API fm_status_t fm_sheet_add_data_table(fm_workbook_t* wb, size_t sheet_index, const fm_data_table_t* spec,
                                           fm_cell_address_t* out_cell);

/**
 * @brief Reads the data table whose result range contains `(row, col)`.
 *        `*out_present` is `1` and `*out` filled when there is one, else `0`
 *        and `*out` zeroed (still `FM_STATUS_OK` (0)).
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_BINDING_NULL_POINTER` (7001) if `out_present` or `out` is `NULL`;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range.
 */
FM_API fm_status_t fm_sheet_get_data_table(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                           int32_t* out_present, fm_data_table_t* out);

/**
 * @brief Removes the data table whose result range contains `(row, col)` and
 *        blanks its result cells.
 *
 * @return `FM_STATUS_OK` (0) on success;
 *         `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) when `sheet_index` is out of range;
 *         `FM_STATUS_NOT_FOUND` (6) when no data table covers the cell.
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
 *         an `FM_STATUS_EVAL_*` or `FM_STATUS_GRAPH_*` code on engine failure.
 */
FM_API fm_status_t fm_workbook_recalc_data_tables(fm_workbook_t* wb);

/* -------------------------------------------------------------------------- */
/* Version                                                                    */
/* -------------------------------------------------------------------------- */

/**
 * @brief Returns the Formulon library version as a NUL-terminated
 *        UTF-8 string. Always non-NULL with program lifetime.
 *
 * Lifetime: static.
 *
 * @return The version string; this function cannot fail.
 */
FM_API const char* fm_get_version(void);

/**
 * @brief Returns the library version as `MAJOR * 10000 + MINOR * 100 + PATCH`.
 *
 * A process constant: equal to `FM_VERSION_NUMBER` when the header and the
 * library come from the same release. It does not touch the last-error state.
 *
 * @return The packed version number; this function cannot fail.
 */
FM_API uint32_t fm_get_version_number(void);

#ifdef __cplusplus
} /* extern "C" */
#endif

#endif /* FORMULON_C_API_H */
