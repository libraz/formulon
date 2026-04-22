// Copyright 2026 libraz. Licensed under the MIT License.
//
// Workbook sheet model. At M1 a `Sheet` is a naked named container: no cells,
// no rows, no dimensions. Later milestones will grow this class in place as
// the cell store, merged-range table, conditional formatting and data
// validation layers come online (see backup/plans/04-xlsx-io.md §4.3).

#ifndef FORMULON_SHEET_H_
#define FORMULON_SHEET_H_

#include <string>
#include <utility>

namespace formulon {

/// A single worksheet inside a `Workbook`.
///
/// At M1 the only observable state is the display name. All other worksheet
/// features (cells, merges, CF, DV, drawings, etc.) arrive in later
/// milestones; until then a `Sheet` behaves as an empty-but-named tab.
class Sheet {
 public:
  /// Builds a sheet with the given display name. The name is adopted
  /// verbatim; callers are expected to supply a valid Excel sheet name
  /// (validation lives in the workbook layer once M6 lands).
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
