// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Cell-level value types used by the worksheet storage layer. A worksheet
// owns a row-sparse, column-dense cell store keyed by 0-based row index
// (see `sheet.h`); each cell carries either a literal value or a formula
// string with an associated cached result.
//
// `Cell` is intentionally a plain aggregate: the storage layer in `Sheet`
// owns lifetime, and downstream layers (parser, evaluator, OOXML writer)
// observe cells through pointers handed out by `Sheet::cell_at`.

#ifndef FORMULON_CELL_H_
#define FORMULON_CELL_H_

#include <cstdint>
#include <memory>
#include <string>

#include "value.h"

namespace formulon {

/// 0-based cell coordinate within a single sheet.
///
/// `row` is in `[0, 1048576)` and `col` is in `[0, 16384)`, matching the
/// Excel 365 sheet dimensions (1,048,576 rows by 16,384 columns).
struct CellAddress {
  std::uint32_t row;
  std::uint32_t col;

  friend bool operator==(CellAddress a, CellAddress b) noexcept { return a.row == b.row && a.col == b.col; }
  friend bool operator!=(CellAddress a, CellAddress b) noexcept { return !(a == b); }
};

/// A single cell's persisted state.
///
/// `formula_text` is the raw formula string starting with `=` (empty when
/// the cell holds a literal). `cached_value` is the cell's effective Value:
/// for a literal cell, this *is* the value; for a formula cell, it is the
/// most recently computed result populated by the evaluator.
///
/// Lifetime contract for Text payloads:
///
///   * Values committed via `Sheet::set_cell_cached_value` (the recalc
///     engine's post-evaluation write path) are deep-copied into
///     `cached_text_owned` so the cell's `cached_value.as_text()` view
///     survives across the per-evaluation arena resets the recalc engine
///     performs between cells. Without this re-interning, a Text scalar
///     produced by one cell's formula would dangle the moment the engine
///     reset its bump arena before evaluating the next cell.
///
///   * Values committed via `Sheet::set_cell_value` keep the caller-owns
///     contract: the OOXML reader, for example, points the `string_view` at
///     the workbook-scoped shared string table and that storage outlives
///     the cell. `cached_text_owned` is left untouched on this path.
///
/// `cached_text_owned` is a `unique_ptr<std::string>` rather than a bare
/// `std::string` so the bytes live at a heap-stable address that survives
/// Cell moves. `Sheet`'s row-sparse cell store routinely moves Cells when
/// row vectors grow or the row map rehashes; a bare `std::string` would
/// relocate its inline (SSO) bytes on every such move and dangle the
/// `cached_value`'s internal `string_view`.
struct Cell {
  std::string formula_text;
  Value cached_value = Value::blank();
  /// Backing storage for `cached_value` when it is a Text re-interned by
  /// `Sheet::set_cell_cached_value`. Null when the cached value was never
  /// re-interned (e.g. a Text supplied through `Sheet::set_cell_value`
  /// where the caller-owns contract applies, or any non-Text value). The
  /// pointer is heap-stable so the `cached_value.as_text()` view does not
  /// dangle when the owning Cell is moved by the `Sheet` storage layer.
  std::unique_ptr<std::string> cached_text_owned;
  /// Kana annotation associated with a Text cell, populated from the OOXML
  /// `<rPh>` markers attached to the cell's source `<si>` (SST entry) or
  /// `<is>` (inline string). Empty when the cell carries no annotation.
  /// `PHONETIC` reads this field directly via the lazy dispatch path; the
  /// writer emits it back as `<rPh sb="0" eb="N">` inside the `<is>` block
  /// on save. The annotation is stored as a single concatenated kana string
  /// (multi-block `<rPh>` runs collapse to one block on round-trip).
  std::string phonetic_text;
  /// Index into the workbook's `StylesTable::cell_xfs` describing this
  /// cell's visual format (number-format id, font/fill/border indices,
  /// alignment). `0` is the default xf and is the value Excel writes when
  /// no `s=` attribute is present on the `<c>` element. The OOXML reader
  /// populates this field from the `s=` attribute; the writer emits it
  /// back when non-zero. Carried as `std::uint32_t` for parity with the
  /// `cellXfs` index width (Excel allows up to ~65,000 entries; we leave
  /// headroom for future widening).
  std::uint32_t xf_index = 0;
};

}  // namespace formulon

#endif  // FORMULON_CELL_H_
