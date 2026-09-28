// Public ESM surface for @libraz/formulon: the single-threaded build, which
// loads without cross-origin isolation. The generated Emscripten module
// remains an implementation detail so this shim can expose the value
// constants declared in formulon.d.ts alongside its default factory.

import createFormulon from './formulon_core.js';

export default createFormulon;
export * from './common.js';
