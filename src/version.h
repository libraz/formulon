#ifndef FORMULON_VERSION_H_
#define FORMULON_VERSION_H_

// Single source of truth for the Formulon version. Keep in lockstep with
// `CMakeLists.txt` (`project(formulon VERSION ...)`),
// `packages/npm/package.json`, `packages/npm-native/package.json`, and
// `packages/python/pyproject.toml`. The release skill bumps all of them
// together.

#define FORMULON_VERSION_MAJOR 0
#define FORMULON_VERSION_MINOR 9
#define FORMULON_VERSION_PATCH 1

#define FORMULON_VERSION_STRINGIFY_(x) #x
#define FORMULON_VERSION_STRINGIFY(x) FORMULON_VERSION_STRINGIFY_(x)

#define FORMULON_VERSION_STRING                      \
  FORMULON_VERSION_STRINGIFY(FORMULON_VERSION_MAJOR) \
  "." FORMULON_VERSION_STRINGIFY(FORMULON_VERSION_MINOR) "." FORMULON_VERSION_STRINGIFY(FORMULON_VERSION_PATCH)

#endif  // FORMULON_VERSION_H_
