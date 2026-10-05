#include "theme.h"

namespace formulon {

const Theme& default_theme() {
  static const Theme kTheme = {{0xFF000000U, 0xFFFFFFFFU, 0xFF44546AU, 0xFFE7E6E6U, 0xFF4472C4U, 0xFFED7D31U,
                                0xFFA5A5A5U, 0xFFFFC000U, 0xFF5B9BD5U, 0xFF70AD47U, 0xFF0563C1U, 0xFF954F72U},
                               {"Calibri Light", "\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF Light",
                                "Calibri", "\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF"}};
  return kTheme;
}

}  // namespace formulon
