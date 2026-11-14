//
// Office theme packages (`.thmx`): apply one to a workbook's theme part, or
// save the workbook's current theme as one. Bytes in, bytes out; theme
// variants are ignored and no thumbnail is written.

#ifndef FORMULON_WORKBOOK_THEME_PACKAGE_H_
#define FORMULON_WORKBOOK_THEME_PACKAGE_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace io {

/// Implements `Workbook::apply_theme`. Replaces the workbook's theme part (and
/// the media its relationships carry) with the package's default theme.
///
/// Atomic: the package is fully read and validated before the workbook is
/// touched. Fails with `kBindingNullPointer` for NULL `data`,
/// `kInvalidArgument` for `size == 0`, the ZIP reader's code for a corrupt or
/// over-budget archive (`kIoZipCorrupt`, `kSecurityResourceLimit`),
/// `kIoRelationshipBroken` for a missing officeDocument / theme relationship or
/// target, and `kIoXmlParse` for an unparseable theme. Marks every formula
/// dirty when the major or minor font faces change.
Expected<void, Error> apply_theme_package(Workbook& wb, const std::uint8_t* data, std::size_t size);

/// Implements `Workbook::save_theme`. Packages the current theme part (the
/// default theme when the workbook has none) with its rels and media.
Expected<std::vector<std::uint8_t>, Error> save_theme_package(const Workbook& wb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_WORKBOOK_THEME_PACKAGE_H_
