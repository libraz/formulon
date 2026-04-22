// Copyright 2026 libraz. Licensed under the MIT License.
//
// Top-level workbook model. At M1 a `Workbook` owns a vector of `Sheet`
// instances and exposes a `save()` method that serialises the workbook to
// a `.xlsx` byte stream via the OOXML writer slice. Later milestones add
// defined names, styles, shared strings, tables, pivots and the full cell
// store (see backup/plans/04-xlsx-io.md).

#ifndef FORMULON_WORKBOOK_H_
#define FORMULON_WORKBOOK_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

/// Move-only container representing a spreadsheet workbook.
///
/// Instances are constructed via the `create()` factory, which returns the
/// M1-default workbook of a single `"Sheet1"`. Additional sheets will be
/// exposed through explicit mutation APIs in later milestones; callers
/// currently only mutate the one sheet's display name.
class Workbook {
 public:
  /// Factory for the default M1 workbook: a single sheet named `"Sheet1"`.
  static Workbook create();

  Workbook(const Workbook&) = delete;
  Workbook& operator=(const Workbook&) = delete;
  Workbook(Workbook&&) noexcept = default;
  Workbook& operator=(Workbook&&) noexcept = default;
  ~Workbook() = default;

  /// Number of sheets in the workbook. Always at least 1 for
  /// `create()`-constructed instances.
  std::size_t sheet_count() const noexcept { return sheets_.size(); }

  /// Immutable access to the sheet at `index`. `index` must be `<
  /// sheet_count()`.
  const Sheet& sheet(std::size_t index) const { return sheets_[index]; }

  /// Mutable access to the sheet at `index`. `index` must be `<
  /// sheet_count()`.
  Sheet& sheet(std::size_t index) { return sheets_[index]; }

  /// Serialises the workbook to an in-memory `.xlsx` byte stream. Delegates
  /// to `io::write_ooxml`; see that function's documentation for the exact
  /// set of OOXML parts emitted in the M1 slice.
  Expected<std::vector<std::uint8_t>, Error> save() const;

 private:
  Workbook() = default;

  std::vector<Sheet> sheets_;
};

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_H_
