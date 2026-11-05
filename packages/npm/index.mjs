// Public ESM surface for @libraz/formulon: the single-threaded build, which
// loads without cross-origin isolation. The generated Emscripten module
// remains an implementation detail so this shim can expose the value
// constants declared in formulon.d.ts both as named exports and on the
// module object the factory resolves to.

import * as common from './common.js';
import createModule from './formulon_core.js';

export default async function createFormulon(opts) {
  return Object.assign(await createModule(opts), common);
}
export * from './common.js';
