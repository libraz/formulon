// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// C ABI - library version string.

#include "version.h"

#include "c_api/formulon_c.h"

extern "C" const char* fm_version_string(void) {
  return FORMULON_VERSION_STRING;
}
