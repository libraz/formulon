// Copyright 2026 libraz. Licensed under the MIT License.
//
// Workbook sheet model. Currently a `Sheet` is a naked named container: no
// cells, no rows, no dimensions. This class will grow in place as the cell
// store, merged-range table, conditional formatting and data validation
// layers come online (see backup/plans/04-xlsx-io.md §4.3).

#ifndef FORMULON_SHEET_H_
#define FORMULON_SHEET_H_

#include <string>
#include <utility>

namespace formulon {

/// A single worksheet inside a `Workbook`.
///
/// Currently the only observable state is the display name. All other
/// worksheet features (cells, merges, CF, DV, drawings, etc.) arrive in
/// follow-up work; until then a `Sheet` behaves as an empty-but-named tab.
class Sheet {
 public:
  /// Builds a sheet with the given display name. The name is adopted
  /// verbatim; callers are expected to supply a valid Excel sheet name
  /// (name validation will live in the workbook layer once it is wired up).
  explicit Sheet(std::string name) : name_(std::move(name)) {}

  /// Current display name of the sheet.
  const std::string& name() const noexcept { return name_; }

  /// Replaces the display name.
  void set_name(std::string name) { name_ = std::move(name); }

 private:
  std::string name_;
};

}  // namespace formulon

#endif  // FORMULON_SHEET_H_
