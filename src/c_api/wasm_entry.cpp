// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Stub entry point for the capi WASM variant.
//
// emcc requires an executable target to drive the link graph. The capi
// build emits a reactor-style WASM (--no-entry + -sSTANDALONE_WASM=1),
// so this TU exists only to give the executable a single C++ TU at the
// build root; the actual fm_* exports are pulled in via the static
// archive based on the -sEXPORTED_FUNCTIONS list.
//
// Nothing here is called at runtime. The empty TU avoids a separate
// "header-only target" cmake construct.
