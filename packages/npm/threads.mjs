// Public ESM surface for @libraz/formulon/threads: the pthread build, whose
// `recalcParallel` runs on Web Workers. It needs SharedArrayBuffer, so a
// browser page must be cross-origin isolated (COOP/COEP). The pthread
// workers load formulon_threads_core.js directly, not this shim (see
// scripts/stage.mjs).

import * as common from './common.js';
import createModule from './formulon_threads_core.js';

export default async function createFormulon(opts) {
  return Object.assign(await createModule(opts), common);
}
export * from './common.js';
