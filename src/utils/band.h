// Perpendicular extent of a band-limited structural edit.

#ifndef FORMULON_UTILS_BAND_H_
#define FORMULON_UTILS_BAND_H_

#include <cstdint>

namespace formulon {

/// Closed interval of rows (for a vertical shift) or columns (for a horizontal
/// one) that an Insert Cells / Delete Cells edit is confined to.
struct Band {
  std::uint32_t first;
  std::uint32_t last;
};

}  // namespace formulon

#endif  // FORMULON_UTILS_BAND_H_
