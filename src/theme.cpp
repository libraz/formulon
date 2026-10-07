#include "theme.h"

namespace formulon {

const Theme& default_theme() {
  static const Theme kTheme = {{0xFF000000U, 0xFFFFFFFFU, 0xFF0E2841U, 0xFFE8E8E8U, 0xFF156082U, 0xFFE97132U,
                                0xFF196B24U, 0xFF0F9ED5U, 0xFFA02B93U, 0xFF4EA72EU, 0xFF467886U, 0xFF96607DU},
                               {"Aptos Display", "\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF Light",
                                "Aptos Narrow", "\xE6\xB8\xB8\xE3\x82\xB4\xE3\x82\xB7\xE3\x83\x83\xE3\x82\xAF"}};
  return kTheme;
}

}  // namespace formulon
