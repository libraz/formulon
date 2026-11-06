//
// C ABI - library version string.

#include "version.h"

#include "c_api/formulon_c.h"

static_assert(FM_VERSION_MAJOR == FORMULON_VERSION_MAJOR && FM_VERSION_MINOR == FORMULON_VERSION_MINOR &&
                  FM_VERSION_PATCH == FORMULON_VERSION_PATCH,
              "FM_VERSION_* in formulon_c.h must match src/version.h");

extern "C" const char* fm_get_version(void) {
  return FORMULON_VERSION_STRING;
}

extern "C" uint32_t fm_get_version_number(void) {
  return static_cast<uint32_t>(FM_VERSION_NUMBER);
}
