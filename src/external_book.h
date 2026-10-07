//
// The parsed body of an external link: the supporting workbook's sheet
// names, its defined names, and the values Excel cached the last time it
// could read that workbook.
//
// This is what makes a cross-workbook reference evaluable without
// opening the other file. Excel itself works from this cache whenever
// the source is closed, and it caches exactly the cells the consuming
// workbook references -- so a reference that resolves here is the same
// reference Excel resolves.
//
// The model is shared by both container formats. `externalLink<N>.xml`
// and `externalLink<N>.bin` carry the same information in different
// encodings, and both readers populate this struct.
//
// Design references:
//   * ECMA-376 SS18.14 (externalLink, externalBook, sheetDataSet)

#ifndef FORMULON_EXTERNAL_BOOK_H_
#define FORMULON_EXTERNAL_BOOK_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "value.h"

namespace formulon {

/// One cached cell of a supporting workbook.
///
/// Text is held as an owned `std::string` rather than inside `value`,
/// because `Value::text` aliases its argument instead of owning it and
/// this struct is copied along with the workbook that holds it. Read the
/// cell through `resolved()`, never through `value` directly.
struct ExternalCell {
  /// Carries the kind for every cached cell, and the payload for every
  /// kind except Text.
  Value value = Value::blank();
  /// The bytes when `value.kind()` is Text; empty otherwise.
  std::string text;

  /// The cell as a `Value` whose Text payload points into this struct.
  /// The returned value borrows `text`, so it must not outlive the cell.
  Value resolved() const noexcept { return value.is_text() ? Value::text(text) : value; }
};

/// One `<definedName>` of a supporting workbook, already resolved to the
/// rectangle it names.
///
/// The stored `refersTo` is a formula in the *supporting* workbook's
/// coordinate space, so `sheet` indexes `ExternalBook::sheet_names` and
/// never this workbook's sheets.
struct ExternalBookName {
  std::string name;
  /// 0-based index into `ExternalBook::sheet_names` of the sheet the name is
  /// local to (`<definedName sheetId>`), or `ExternalBook::kNoSheet` for a
  /// book-scope name.
  std::uint32_t scope_sheet = static_cast<std::uint32_t>(-1);
  /// 0-based index into `ExternalBook::sheet_names`.
  std::uint32_t sheet = 0;
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  /// Bottom-right corner; equal to `row` / `col` unless `is_range`.
  std::uint32_t row_end = 0;
  std::uint32_t col_end = 0;
  bool is_range = false;
  /// False when the name's `refersTo` was something other than a plain
  /// cell or rectangle -- a constant, a computed expression, a
  /// cross-book chain. Excel allows all of those; resolving them would
  /// mean evaluating a formula in another workbook's namespace, which
  /// this cache cannot do. Such a name reads as `#REF!` rather than
  /// being silently resolved against the wrong coordinates.
  bool resolvable = false;
};

/// The cached state of one supporting workbook.
struct ExternalBook {
  /// Sheet display names in the supporting workbook's own order. A
  /// reference's sheet qualifier is matched against these, not against
  /// this workbook's sheets.
  std::vector<std::string> sheet_names;
  std::vector<ExternalBookName> names;
  /// Parallel to `sheet_names`: whether the part carries a cached
  /// `<sheetData>` for that sheet. A missing entry reads as false.
  std::vector<bool> sheet_data;

  /// Cached cell values keyed by `cell_key`. Absent means Excel never
  /// cached that cell, which is not an error: see `cached_cell`.
  std::unordered_map<std::uint64_t, ExternalCell> cells;

  /// Packs a `(sheet, row, col)` address into a `cells` key. Sheet
  /// indices are 16-bit, rows 21-bit (1,048,576) and columns 14-bit
  /// (16,384), so the three fit one 64-bit word without collision.
  static std::uint64_t cell_key(std::uint32_t sheet, std::uint32_t row, std::uint32_t col) noexcept {
    return (static_cast<std::uint64_t>(sheet) << kSheetShift) | (static_cast<std::uint64_t>(row) << kRowShift) |
           static_cast<std::uint64_t>(col);
  }

  /// Every `cells` key in ascending order, which is sheet, then row, then
  /// column; the writers emit cached cells in this order.
  std::vector<std::uint64_t> sorted_cell_keys() const;

  /// Returns the 0-based index of `sheet` in `sheet_names` under ASCII
  /// case folding, as Excel matches sheet names, or `kNoSheet` when absent.
  std::uint32_t sheet_index(std::string_view sheet) const noexcept;

  /// Whether `sheet` has a cached `<sheetData>`. A sheet without one -- a
  /// link created in this session, or one Excel could not refresh -- reads
  /// as `#REF!` rather than as zeros.
  bool sheet_has_data(std::uint32_t sheet) const noexcept { return sheet < sheet_data.size() && sheet_data[sheet]; }

  /// Returns the entry named `name` in `scope_sheet` (`kNoSheet` for book
  /// scope), or `nullptr` when the supporting workbook declares no such
  /// name there. Matched under ASCII case folding, the same comparison the
  /// engine's own defined-name lookup uses.
  const ExternalBookName* find_name(std::string_view name, std::uint32_t scope_sheet = kNoSheet) const noexcept;

  /// The last row and column holding a cached cell on `sheet`; false when
  /// the sheet caches no cell.
  bool cached_extent(std::uint32_t sheet, std::uint32_t* last_row, std::uint32_t* last_col) const noexcept;

  /// Returns the cached value at `(sheet, row, col)`, borrowing this
  /// book's storage for a Text result.
  ///
  /// An address Excel never cached reads as numeric zero rather than as
  /// blank or `#REF!`. Excel caches only the cells this workbook
  /// actually references, and shows `0` for a reference into a
  /// supporting workbook whose value it does not hold.
  Value cached_cell(std::uint32_t sheet, std::uint32_t row, std::uint32_t col) const noexcept;

  static constexpr std::uint32_t kNoSheet = static_cast<std::uint32_t>(-1);

 private:
  static constexpr unsigned kRowShift = 14U;
  static constexpr unsigned kSheetShift = 35U;
  static constexpr std::uint64_t kRowMask = (std::uint64_t{1} << (kSheetShift - kRowShift)) - 1U;
  static constexpr std::uint64_t kColMask = (std::uint64_t{1} << kRowShift) - 1U;
};

}  // namespace formulon

#endif  // FORMULON_EXTERNAL_BOOK_H_
