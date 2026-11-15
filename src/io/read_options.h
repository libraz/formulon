#ifndef FORMULON_IO_READ_OPTIONS_H_
#define FORMULON_IO_READ_OPTIONS_H_

#include <optional>

#include "excel_profile.h"

namespace formulon {
namespace io {

/// Options shared by the package readers (`read_ooxml`, `read_xlsb`).
struct ReadOptions {
  /// Formula-behaviour profile the loaded workbook is assigned. Unset keeps the
  /// default profile. It is applied after parsing, so sheet names are validated
  /// under the load's default host, and it does not dirty any formula: a
  /// manual-calculation book keeps the stale set it has under the default profile.
  std::optional<ExcelProfile> profile;
};

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_READ_OPTIONS_H_
