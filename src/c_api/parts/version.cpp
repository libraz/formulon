//
// C ABI - library version string.

#include "version.h"

#include "c_api/formulon_c.h"

extern "C" const char* fm_get_version(void) {
  return FORMULON_VERSION_STRING;
}
