//
// Workbook and range protection records. Both are views over XML the model
// keeps verbatim (`Workbook::workbook_protection_xml()` and the sheet's
// retained `<protectedRanges>` child); `io/protection_xml.h` converts.

#ifndef FORMULON_PROTECTION_H_
#define FORMULON_PROTECTION_H_

#include <cstdint>
#include <string>

namespace formulon {

/// `<workbookProtection>`. Hash fields are stored as the host computed them;
/// the engine never hashes. A zero spin count and an empty string are absent.
struct WorkbookProtection {
  bool lock_structure = false;
  bool lock_windows = false;
  bool lock_revision = false;
  /// Legacy 16-bit password hash, as the hex text Excel writes.
  std::string workbook_password;
  std::string workbook_algorithm_name;
  std::string workbook_hash_value;
  std::string workbook_salt_value;
  std::uint32_t workbook_spin_count = 0;
  std::string revisions_password;
  std::string revisions_algorithm_name;
  std::string revisions_hash_value;
  std::string revisions_salt_value;
  std::uint32_t revisions_spin_count = 0;
};

/// One `<protectedRange>`: a named range users may edit without unlocking the
/// sheet. `sqref` is the A1 range list as written.
struct ProtectedRange {
  std::string name;
  std::string sqref;
  std::string password;
  std::string algorithm_name;
  std::string hash_value;
  std::string salt_value;
  std::uint32_t spin_count = 0;
  std::string security_descriptor;
};

}  // namespace formulon

#endif  // FORMULON_PROTECTION_H_
